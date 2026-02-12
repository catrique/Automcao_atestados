from services.sheets_service import SheetsService
from config.loaders import get_config

sheets_name = get_config("google_sheets", "planilha")
sheets = SheetsService(sheets_name)