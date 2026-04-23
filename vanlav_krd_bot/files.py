import re
from typing import BinaryIO, Any

import orjson
from aiogram.utils.formatting import Code
from anyio import Path
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook

template_file_path = Path('../data/stored_template.txt')
report_file_path = Path('../data/current_report.xlsx')
users_registry_file_path = Path('../data/users_registry.json')

current_xlsx: Workbook | None = None


ALLOWED_OPERATORS = {
    '>', '<', '>=', "<=", '==', '!=',
    '+', '-', '*', '/', '%',
}


class EmptyValueInBlockError(Exception):
    pass


async def get_template():
    if await template_file_path.is_file():
        return await template_file_path.read_text(encoding='utf-8')
    return None


async def set_template(template_to_write: str):
    await template_file_path.write_text(template_to_write, encoding='utf-8')


async def get_templates_columns() -> set[str]:
    template = await get_template()
    columns = set(re.findall(r'\{\s*?(\w*)\s*?}', template))
    return columns

def apply_worker_to_template(
        template: str,
        worker: tuple[tuple[Any, str | None], ...],
        column_map: dict[str, int],
):
    compiled_lines = []
    for line in template.split('\n'):
        try:
            compiled_line = _apply_worker_to_template_line(line, worker, column_map)
        except EmptyValueInBlockError:
            continue
        compiled_lines.append(compiled_line)
    return '\n'.join(compiled_lines)


def _apply_worker_to_template_line(
        template_line: str,
        worker: tuple[tuple[Any, str | None], ...],
        column_map: dict[str, int],
):
    def optional_blocks_dfs(curr_idx: int = 0):
        copy_from = curr_idx
        result_string = ''
        while curr_idx < len(template_line):
            symbol = template_line[curr_idx]
            if symbol == '(':
                result_string += template_line[copy_from:curr_idx]

                result, temp_end = optional_blocks_dfs(curr_idx + 1)
                result_string += result

                curr_idx = copy_from = temp_end + 1
                continue
            elif symbol == ')':
                block_end_idx = curr_idx
                result_string += template_line[copy_from:block_end_idx]
                # process block
                try:
                    compiled_result_string: str = compile_values(result_string)
                except EmptyValueInBlockError:
                    return '', block_end_idx
                else:
                    wrapped_to_braces = '(' + compiled_result_string.strip(' ') + ')'
                    return wrapped_to_braces, block_end_idx
            curr_idx += 1
        result_string += template_line[copy_from:curr_idx]
        return compile_values(result_string).strip()

    def compile_values(string_to_apply_values: str):
        for match in re.finditer(r'(\[(.+?)])', string_to_apply_values):
            # check computed fields
            external_match, internal_match = match.groups()
            terms: list[str] = [term for term in internal_match.split() if term]
            expression = ''
            terms_iter = iter(terms)
            while True:
                try:
                    term = next(terms_iter)
                except StopIteration:
                    break

                if term.startswith('{'):
                    full_term = term
                    while True:
                        try:
                            subterm = next(terms_iter)
                        except StopIteration:
                            break
                        full_term += ' ' + subterm
                        if subterm.endswith('}'):
                            break

                    column_name = re.sub(r'\s+', ' ', full_term).strip('{} \t\n')
                    column_index = column_map.get(column_name)
                    cell_value, _ = worker[column_index]
                    if cell_value is None:
                        # empty value
                        raise EmptyValueInBlockError('Не найдено необходимое значение для вычислений в колонке ', Code(column_name))

                    expression += orjson.dumps(cell_value).decode()
                    continue
                elif term in ALLOWED_OPERATORS:
                    # check operands
                    expression += f' {term} '
                elif re.fullmatch(r'[\w"\']+', term):
                    expression += term
            try:
                computation_result = eval(expression)
            except Exception:
                computation_result = None

            if computation_result is None:
                raise EmptyValueInBlockError('Вычисление выражения ', Code(external_match), ' не вернуло ничего.')
            elif isinstance(computation_result, bool):
                if not computation_result:
                    raise EmptyValueInBlockError('Вычисление выражения ', Code(external_match), ' ложно.')
            else:
                string_to_apply_values = string_to_apply_values.replace(external_match, str(computation_result))

        for match in re.finditer(r'(\{\s*?(.+?)\s*?})', string_to_apply_values):
            external_match, internal_match = match.groups()
            column = re.sub(r'\s+', ' ', internal_match).strip(' \t\n')
            column_index = column_map.get(column)
            cell_value, _ = worker[column_index]
            if cell_value is None:
                # empty value
                raise EmptyValueInBlockError('Не найдено необходимое значение в колонке ', Code(column))

            formated_value: str
            if isinstance(cell_value, str):
                formated_value = cell_value.strip()
            elif isinstance(cell_value, int):
                formated_value = str(cell_value)
            elif isinstance(cell_value, float):
                formated_value = f'{cell_value:,.2f}'
            else:
                formated_value = str(cell_value)
            string_to_apply_values = string_to_apply_values.replace(external_match, formated_value)

        for match in re.finditer(r'(<\s*?(.+?)\s*?>)', string_to_apply_values):
            external_match, internal_match = match.groups()
            column = internal_match.strip(' \t\n')
            column_index = column_map.get(column)
            _, comment = worker[column_index]
            if comment is None:
                # empty value
                raise EmptyValueInBlockError('Не найден необходимый коментарий в колонке ', Code(column))
            string_to_apply_values = string_to_apply_values.replace(external_match, comment)
        return string_to_apply_values

    return optional_blocks_dfs()


async def clear_report_file():
    global current_xlsx
    if await report_file_path.is_file():
        await report_file_path.unlink()
    current_xlsx = None

async def prepare_xlsx(raw_file: BinaryIO):
    global current_xlsx
    await report_file_path.write_bytes(raw_file.read())
    current_xlsx = load_workbook(str(report_file_path), data_only=True)

def get_xlsx():
    if current_xlsx is None:
        raise TypeError('Xslx is not prepared')
    return current_xlsx


async def update_tg_user(name: str, tg_id: int | None):
    users = await get_users_tg_mapping()
    if tg_id is None:
        del users[name]
    else:
        users[name] = tg_id
    await users_registry_file_path.write_bytes(orjson.dumps(users))


async def get_users_tg_mapping() -> dict[str, int]:
    if await users_registry_file_path.is_file():
        raw = await users_registry_file_path.read_bytes()
        return orjson.loads(raw)
    return {}
