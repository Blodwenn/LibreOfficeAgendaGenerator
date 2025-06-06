from AgendaGenerator import generate_calendar, generate_monthly_agenda, generate_all, generate_daily_agenda, configure_page_for_rmk, generate_title_page

if __name__ == '__main__':
    # For debugging, launch libroffice with:
    # "D:\Program Files\LibreOffice\program\soffice.exe" --calc --accept="socket,host=localhost,port=2002;urp;"
    # "D:\Program Files\LibreOffice\program\soffice.exe" --writer --accept="socket,host=localhost,port=2002;urp;"
    # first, then debug the program as normal (running the program from Pycharm)

    # Calendar icon https://www.flaticon.com/free-icon/calendar_55281?term=calendar&page=1&position=6&origin=tag&related_id=55281

    # GenerateAll()
    # ConfigurePageForRMK()
    # GenerateTitlePage(2025)
    # GenerateCalendar(2025)
    # GenerateMonthlyAgenda(2025)
    generate_daily_agenda(year=2025, test=True)