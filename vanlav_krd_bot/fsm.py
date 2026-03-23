from aiogram.fsm.state import StatesGroup, State


class ReportForm(StatesGroup):
    wait_for_file = State()
    wait_for_sheet = State()
    wait_for_fixes = State()
    wait_for_answer_for_lack_tg_users = State()
    checking_reports = State()
    broadcasted_texts = State()

class TemplateForm(StatesGroup):
    wait_for_template = State()
    wait_for_template_in_report_state = State()

class NameForm(StatesGroup):
    wait_for_name = State()
    wait_for_approve = State()
