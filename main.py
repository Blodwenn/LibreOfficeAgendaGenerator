from AgendaGenerator import GenerateCalendar, GenerateMonthlyAgenda, GenerateAll, GenerateDailyAgenda, ConfigurePageForRMK, GenerateTitlePage

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
    GenerateDailyAgenda(year=2025, test=True)