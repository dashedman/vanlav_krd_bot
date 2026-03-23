import re
from collections import defaultdict
from typing import Any

from aiogram import Dispatcher, Router, F, Bot
from aiogram.enums import ChatAction
from aiogram.filters import CommandStart, Command, ExceptionTypeFilter
from aiogram.fsm.context import FSMContext
from aiogram.types import Message, ReplyKeyboardRemove, ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, \
    InlineKeyboardButton, CallbackQuery, ErrorEvent
from aiogram.utils.formatting import as_list, Code, Text, Pre, as_line, as_marked_list, Bold, as_marked_section
from aiogram.utils.keyboard import InlineKeyboardBuilder
from anyio import Path

from config import settings
from files import clear_report_file, prepare_xlsx, set_template, get_template, get_xlsx, get_templates_columns, \
    apply_worker_to_template, get_users_tg_mapping, update_tg_user, EmptyValueInBlockError
from fsm import ReportForm, TemplateForm, NameForm

dp = Dispatcher()


form_router = Router()
admin_filter = F.from_user.id.in_(settings.ADMINS)


start_report_kb = KeyboardButton(text='Рассылка отчетов')
view_template_kb = KeyboardButton(text='Показать шаблон отчета')
xlsx_rules_kb = KeyboardButton(text='Ожидаемый формат файла')
start_keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [start_report_kb, view_template_kb, xlsx_rules_kb]
    ],
    is_persistent=True,
)

cancel_kb = InlineKeyboardButton(text='Отменить', callback_data='cancel_report')
cancel_keyboard = InlineKeyboardMarkup(inline_keyboard=[[cancel_kb]])

change_template_kb = InlineKeyboardButton(text='Изменить шаблон', callback_data='change_template')
change_template_keyboard = InlineKeyboardMarkup(inline_keyboard=[[change_template_kb]])

resend_file_kb = InlineKeyboardButton(text='Отправить файл заного', callback_data='resend_file')
change_template_or_resend_file_keyboard = InlineKeyboardMarkup(inline_keyboard=[[change_template_kb, resend_file_kb]])

yes_kb = InlineKeyboardButton(text='Да', callback_data='yes_answer')
no_kb = InlineKeyboardButton(text='Нет', callback_data='no_answer')
yes_no_keyboard = InlineKeyboardMarkup(inline_keyboard=[[yes_kb, no_kb]])

yes_keyboard = InlineKeyboardMarkup(inline_keyboard=[[yes_kb]])

forget_kb = InlineKeyboardButton(text='Забыть меня', callback_data='forget_me')
forget_keyboard = InlineKeyboardMarkup(inline_keyboard=[[forget_kb]])

empty_keyboard = InlineKeyboardMarkup(inline_keyboard=[])

SUPPORTED_EXTENSIONS = ('.xlsx','.xlsm','.xltx','.xltm')
SUPPORTED_EXTENSIONS_AS_CODE = tuple(Code(se) for se in SUPPORTED_EXTENSIONS)
SUPPORTED_EXTENSIONS_AS_LINE = as_line(*SUPPORTED_EXTENSIONS_AS_CODE, sep=', ')
SUPPORTED_EXTENSIONS_AS_LIST = as_list(*SUPPORTED_EXTENSIONS_AS_CODE)


@form_router.message(CommandStart(), admin_filter)
async def start_admin(message: Message):
    await message.answer(
        'Здравствуйте',
        reply_markup=start_keyboard,
    )


@form_router.message(CommandStart())
async def start(message: Message, state: FSMContext):
    tg_mapping = await get_users_tg_mapping()
    for name, tg_id in tg_mapping.items():
        if tg_id == message.from_user.id:
            await message.reply(f'Привет {name}', reply_markup=forget_keyboard)
            return

    await state.set_state(NameForm.wait_for_name)
    await message.reply(f'Привет. Я тебя не знаю. Пришли мне пожалуйста свои имя и фамилию')


@form_router.message(NameForm.wait_for_name)
async def get_name(message: Message, state: FSMContext):
    tg_mapping = await get_users_tg_mapping()
    name = message.text.strip().title()
    if name in tg_mapping:
        await message.reply(
            **as_list(
                'Похоже что такое имя уже есть в базе.',
                'Поговорите об этом с человеком ответственным за рассылку отчетов.'
            ).as_kwargs()
        )
        await state.clear()
        await message.answer('Диалог завершен')
        return

    await state.set_data({'name': name})
    await message.reply(f'Твое имя "{name}"?', reply_markup=yes_keyboard)


@form_router.callback_query(NameForm.wait_for_name, F.data == yes_kb.callback_data)
async def approve_name(query: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    name = data['name']
    await update_tg_user(name, query.from_user.id)
    await query.message.answer('Имя записано!', reply_markup=forget_keyboard)
    await state.clear()


@form_router.callback_query(F.data == forget_kb.callback_data)
async def forget_user(query: CallbackQuery):
    tg_mapping = await get_users_tg_mapping()
    for name, tg_id in tg_mapping.items():
        if tg_id == query.from_user.id:
            await update_tg_user(name, None)
            await query.message.answer('Ты забыт.')
            return
    await query.message.answer('Я не знаю кто ты.')


@form_router.message(Command('report'), admin_filter)
@form_router.message(F.text == start_report_kb.text, admin_filter)
async def report_session(message: Message, state: FSMContext):
    await state.set_state(ReportForm.wait_for_file)
    await message.answer(
        **Text(
            "Пришлите пожалуйста отчет в виде файла документа в одном из форматов: ",
            SUPPORTED_EXTENSIONS_AS_LINE
        ).as_kwargs(),
        reply_markup=ReplyKeyboardRemove(),
    )


@form_router.message(ReportForm.wait_for_file, F.document.is_(None))
async def no_link_file(message: Message):
    await message.answer('Документ не присоединен. Отправьте снова или отмените.', reply_markup=cancel_keyboard)


@form_router.message(ReportForm.wait_for_file)
async def link_file(message: Message, bot: Bot, state: FSMContext):
    file_path = await bot.get_file(message.document.file_id)
    file_pathlib = Path(file_path.file_path)
    if file_pathlib.suffix not in SUPPORTED_EXTENSIONS:
        await message.answer(
            **as_list(
                'Расширение документа не:',
                SUPPORTED_EXTENSIONS_AS_LINE,
                'Отправьте документ в необходимом формате либо отмените.'
            ).as_kwargs(),
            reply_markup=cancel_keyboard
        )
        return
    await bot.send_chat_action(message.chat.id, action=ChatAction.TYPING)
    binary_io = await bot.download_file(file_path.file_path)
    await prepare_xlsx(binary_io)
    await state.set_state(ReportForm.wait_for_sheet)
    # scan lists
    workbook = get_xlsx()
    sheetnames_keyboard = InlineKeyboardBuilder()
    for sheet_name in workbook.sheetnames:
        sheetnames_keyboard.button(text=sheet_name, callback_data=sheet_name)
    sheetnames_keyboard.adjust(3, repeat=True)

    await message.answer('Выберите лист для работы', reply_markup=sheetnames_keyboard.as_markup())

@form_router.callback_query(ReportForm.wait_for_sheet)
async def scan_file(query: CallbackQuery, state: FSMContext):
    sheet_name = query.data
    await state.set_data(data={'sheet_name': sheet_name})
    await read_sheet(sheet_name, state, query.message)

async def read_sheet(sheet_name: str, state: FSMContext, message: Message):
    columns_from_template = await get_templates_columns()

    wookbook = get_xlsx()
    worksheet = wookbook[sheet_name]

    # read columns
    rows_iter = worksheet.iter_rows()
    first_row = next(rows_iter)
    column_name_to_index = {
        re.sub(r'\s+', ' ', cell.value): idx
        for idx, cell in enumerate(first_row)
    }
    column_name_to_index['Группа'] = len(column_name_to_index)
    columns_from_table = set(column_name_to_index)

    template_table_diff = columns_from_template - columns_from_table
    table_template_diff = columns_from_table - columns_from_template

    if template_table_diff:
        template = await get_template()
        warning_msg = as_list(
            'Найдены избыточные столбцы в шаблоне (не достает в таблице):',
            as_marked_list(
                *( Code(extra_column) for extra_column in template_table_diff)
            ),
            'Также, свободные от шаблона столбцы в таблице:',
            as_marked_list(
                *( Code(extra_column) for extra_column in table_template_diff)
            ),
            'Текущий шаблон:',
            Pre(template),
        )
        await state.set_state(ReportForm.wait_for_fixes)
        await message.answer(**warning_msg.as_kwargs(), reply_markup=change_template_or_resend_file_keyboard)
        return

    # read groups
    # skip second row
    _ = next(rows_iter)
    groups = defaultdict[str, list[tuple[tuple[Any, str | None], ...]]](list)
    while True:
        try:
            group_header_row = next(rows_iter)
        except StopIteration:
            break
        group_name = group_header_row[0].value
        if group_name is None:
            if group_header_row[1].value and 'Расчет премий по часам' in group_header_row[1].value:
                # премии
                break
            # not group, skip
            continue

        while True:
            group_row = next(rows_iter)
            if group_row[0].value is None:
                # detected empty row, exit group
                break
            groups[group_name].append(tuple(
                (cell.value, cell.comment and cell.comment.text)
                for cell in group_row
            ) + ((group_name, None),))

    # validate tg group
    tg_mapping = await get_users_tg_mapping()
    error_was = False
    for group_name, workers in groups.items():
        workers_without_tg = []
        for worker in workers:
            worker_name: str = worker[0][0]
            if worker_name not in tg_mapping:
                workers_without_tg.append(worker_name)
        if workers_without_tg:
            error_was = True
            await message.answer(
                **as_marked_section(
                    as_list(
                        Bold(f'Группа: {group_name}'),
                        'Следующие сотрудники не зарегистрированны в боте.',
                        Text('Им следует написать боту ', Code('/start'), ' и ввести имя и фамилию как это указанно в таблице.'),
                    ),
                    *workers_without_tg
                ).as_kwargs()
            )

    # two_workers = [(w, c)  for w, c in workers_counter.items() if c > 1]
    # if two_workers:
    #     await message.answer(
    #         **as_marked_section(
    #             Text(
    #                 Bold('Предупреждение!'),
    #                 'Найдены сотрудники с одинаковыми именами. Рассылка будет не состоятельна. Поговорите с разработчиком'
    #             ),
    #             *(as_key_value(w, c) for w, c in two_workers)
    #         ).as_kwargs()
    #     )
    #     await cancel_report_cb(message, state)
    #     return

    if error_was:
        await state.set_state(ReportForm.wait_for_answer_for_lack_tg_users)
        await state.set_data({'groups': groups, 'column_map': column_name_to_index})
        await message.answer('Продолжить процесс рассылки?', reply_markup=yes_no_keyboard)
        return

    await prepare_reports(message, groups, column_name_to_index, state)


async def prepare_reports(
        message: Message,
        groups: dict[str, list[tuple[tuple[Any, str | None], ...]]],
        column_name_to_index: dict[str, int],
        state: FSMContext
):
    tg_mapping = await get_users_tg_mapping()
    template = await get_template()
    reports = defaultdict(list)

    with_error = False
    for group, workers in groups.items():
        for worker in workers:
            worker_name = worker[0][0]
            worker_id = tg_mapping.get(worker_name)
            try:
                result_for_worker = apply_worker_to_template(template, worker, column_name_to_index)
            except EmptyValueInBlockError as e:
                await message.answer(
                    **as_list(
                        Bold("Ошибка в процессе формирования отчетов!"),
                        Text(*e.args),
                        Text('Группа: ', Bold(group)),
                        Text('Сотрудник: ', Bold(worker_name)),
                    ).as_kwargs()
                )
                with_error = True
            else:
                reports[group].append((worker[0][0], worker_id, result_for_worker))

    if with_error:
        await cancel_report_cb(message, state)
        return

    await state.set_data({'reports': reports})
    await state.set_state(ReportForm.checking_reports)

    kb = InlineKeyboardBuilder()
    for group in reports:
        kb.button(text=group, callback_data=group)
    kb.button(text='Подтвердить отправку', callback_data='approve')
    kb.button(text='Отменинть отправку', callback_data='cancel')
    kb.adjust(1)
    await message.answer(
        'Сформированны отчеты для сотрудников. Желаете посмотреть?',
        reply_markup=kb.as_markup()
    )

@form_router.callback_query(ReportForm.checking_reports, F.data == 'approve')
async def broadcast(query: CallbackQuery, state: FSMContext, bot: Bot):
    data = await state.get_data()
    # set empty keyboard
    await query.message.edit_reply_markup(reply_markup=empty_keyboard)
    reports = data['reports']
    for group, workers in reports.items():
        await query.message.answer(f'Начата рассылка для группы {group}')
        for w_name, w_tg_id, w_report in workers:
            if w_tg_id is None:
                await query.message.answer(f'Пропускаю {w_name}')
                continue
            await bot.send_message(w_tg_id, text=w_report)
    await query.message.answer('Рассылка завершена')


@form_router.callback_query(ReportForm.checking_reports, F.data == 'cancel')
async def cancel_broadcast(query: CallbackQuery, state: FSMContext):
    await query.message.edit_reply_markup(reply_markup=empty_keyboard)
    await query.message.answer('Отправка отменена')
    await cancel_report_cb(query, state)


@form_router.callback_query(ReportForm.checking_reports)
async def checking_reports(query: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    reports = data['reports']

    if query.data == 'back':
        kb = InlineKeyboardBuilder()
        for group in reports:
            kb.button(text=group, callback_data=group)
        kb.button(text='Подтвердить отправку', callback_data='approve')
        kb.button(text='Отменить отправку', callback_data='cancel')
    elif query.data in reports:
        # data is group
        group = reports[query.data]
        # send workers kb
        kb = InlineKeyboardBuilder()
        for worker_name, worker_id, _ in group:
            kb.button(text=worker_name + ('' if worker_id else ' (no tg!)'), callback_data=worker_name)
        kb.button(text='Назад', callback_data='back')
    else:
        # data is worker id
        # search for it
        worker_text = next(wt for g in reports.values() for wn, wtg, wt in g if wn == query.data)
        await query.message.answer(worker_text)
        return
    # edit keyboard
    kb.adjust(1)
    await query.message.edit_reply_markup(
        reply_markup=kb.as_markup()
    )


@form_router.callback_query(ReportForm.wait_for_answer_for_lack_tg_users, F.data == yes_kb.callback_data)
async def yes_for_report_tg_lack(query: CallbackQuery, state: FSMContext):
    await query.message.answer('Хорошо. Продолжаю рассылку для имеющихся в моей базе сотрудников.')
    state_data = await state.get_data()
    groups = state_data['groups']
    column_name_to_index = state_data['column_map']
    await prepare_reports(query.message, groups, column_name_to_index, state)


@form_router.callback_query(ReportForm.wait_for_answer_for_lack_tg_users, F.data == no_kb.callback_data)
async def no_for_report_tg_lack(query: CallbackQuery, state: FSMContext):
    await query.message.answer('Хорошо. Повторите пожалуйста попытку когда данные о сотрудниках будут готовы.')
    await cancel_report_cb(query, state)


@form_router.callback_query(ReportForm.wait_for_fixes, F.data == resend_file_kb.callback_data)
async def resend_file(query: CallbackQuery, state: FSMContext):
    await cancel_report_cb(query, state=state)
    await report_session(message=query.message, state=state)


@form_router.message(Command('cancel_report'), admin_filter)
@form_router.callback_query(F.data == cancel_kb.callback_data)
async def cancel_report_cb(query_or_message: CallbackQuery | Message, state: FSMContext):
    await state.clear()
    await clear_report_file()

    if isinstance(query_or_message, CallbackQuery):
        message = query_or_message.message
    else:
        message = query_or_message
    await message.answer('Контекс забыт.', reply_markup=start_keyboard)


@form_router.message(Command('view_template'), admin_filter)
@form_router.message(F.text == view_template_kb.text, admin_filter)
async def view_template(message: Message):
    curr_template = await get_template()
    if curr_template is None:
        await message.answer('Шаблон пуст.', reply_markup=change_template_keyboard)
        return
    await message.answer(
        **Pre(curr_template).as_kwargs(),
        reply_markup=change_template_keyboard
    )


@form_router.message(Command('change_report'), admin_filter)
@form_router.callback_query(F.data == change_template_kb.callback_data, admin_filter)
async def change_template(query_or_message: CallbackQuery | Message, state: FSMContext):
    prev_state = await state.get_state()
    if prev_state == ReportForm.wait_for_fixes.state:
        await state.set_state(TemplateForm.wait_for_template_in_report_state)
    else:
        await state.set_state(TemplateForm.wait_for_template)

    if isinstance(query_or_message, CallbackQuery):
        message = query_or_message.message
    else:
        message = query_or_message
    await message.answer('Пришлите новый шаблон')


@form_router.message(TemplateForm.wait_for_template)
@form_router.message(TemplateForm.wait_for_template_in_report_state)
async def set_template_handler(message: Message, state: FSMContext):
    await set_template(message.text)
    await message.answer('Шаблон установлен.')
    prev_state = await state.get_state()
    if prev_state == TemplateForm.wait_for_template_in_report_state.state:
        data = await state.get_data()
        if 'sheet_name' in data:
            await read_sheet(sheet_name=data.get('sheet_name'), state=state, message=message)
            return
    await state.clear()
    await view_template(message)


@form_router.message(Command('show_rules'), admin_filter)
@form_router.message(F.text == xlsx_rules_kb.text, admin_filter)
async def show_rules(message: Message):
    await message.answer(
        **as_list(
            Text('Ожидается файл разрешения ', SUPPORTED_EXTENSIONS_AS_LINE),
            'Первая строка - заголовки столбцов',
            'Вторая строка - итого',
            'С третьей группы:',
            ' - первая строка группы - её заголовок',
            ' - последующие строки группы - её сотрудники',
            ' - последняя строка группы - пустая, важный маркер о конце группы и начале следующей',
            'Чтение заканчивается либо в конце таблицы либо при встречи записи "Расчет премий по часам" в "B" столбце',
        ).as_kwargs()
    )


# @form_router.error()
# async def error_handler(event: ErrorEvent):
#     logger.critical("Critical error caused by %s", event.exception, exc_info=True)
#     # do something with error
#     ...


dp.include_router(form_router)