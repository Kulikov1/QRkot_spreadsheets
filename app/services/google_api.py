from datetime import datetime

from aiogoogle import Aiogoogle

from app.constants import CHARITY_SHEET, FORMAT
from app.core.config import settings


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheets_body = {
        'properties': {'title': f'Отчет на {now_date_time}',
                       'locale': 'ru_RU'},
        'sheets': [CHARITY_SHEET]
    }
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheets_body)
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
    spreadsheetid: str,
    wrapper_services: Aiogoogle,
) -> None:
    permission_body = {'type': 'user',
                       'role': 'writer',
                       'emailAddress': settings.email}
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permission_body,
            fields='id',
        )
    )


async def spreadsheets_update_value(
    spreadsheetid: str,
    charity_projects: list,
    wrapper_services: Aiogoogle,
) -> None:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = [
        ['Отчет от', now_date_time],
        ['Топ проектов по скорости закрытия'],
        ['Название проекта', 'Время сбора', 'Описание']
    ]
    for proj in charity_projects:
        complete_rate = proj.close_date - proj.create_date
        new_row = [str(proj.name), str(complete_rate), str(proj.description)]
        table_values.append(new_row)
    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values,
    }
    await wrapper_services.as_service_account(
        service.spreadsheets.values.append(
            spreadsheetId=spreadsheetid,
            range=f'A1:{len(table_values)}',
            valueInputOption='USER_ENTERED',
            json=update_body,
        )
    )
