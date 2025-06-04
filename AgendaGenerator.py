import calendar
import datetime
from string import ascii_uppercase

import uno

# from com.sun.star.awt import Rectangle, WindowDescriptor
# from com.sun.star.awt.WindowClass import TOP
# from com.sun.star.awt.WindowAttribute import BORDER, MOVEABLE, SIZEABLE,CLOSEABLE

from com.sun.star.awt import FontWeight
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.beans import PropertyValue
from com.sun.star.style.BreakType import PAGE_AFTER, PAGE_BEFORE
from com.sun.star.style.ParagraphAdjust import CENTER as HOR_CENTER
from com.sun.star.table import BorderLine2
from com.sun.star.table.BorderLineStyle import SOLID
from com.sun.star.text import ControlCharacter
from com.sun.star.text.VertOrientation import CENTER as VER_CENTER


# def _askYear(smgr):
#     toolkit = smgr.createInstance("com.sun.star.awt.Toolkit")
#     descriptor = WindowDescriptor()
#     descriptor.Type = TOP
#     descriptor.WindowServiceName = "dialog"
#     descriptor.ParentIndex = -1
#     descriptor.Parent = None
#     descriptor.Bounds = Rectangle(200, 200, 500, 300)
#     descriptor.WindowAttributes = BORDER | MOVEABLE | SIZEABLE | CLOSEABLE
#     window = toolkit.createWindow(descriptor)
#     frame = smgr.createInstance("com.sun.star.frame.Frame")
#     frame.initialize(window)
#     treeroot = smgr.createInstance("com.sun.star.frame.Desktop")
#     childcontainer = treeroot.getFrames()
#     childcontainer.append(frame)
#     # window.setBackground(0x98A9BA)
#     window.setTitle("Introduce year")
#     window.setVisible(True)

def addAwtModel(oDM, srv, sName, dProps):
    '''Insert UnoControl<srv>Model into given DialogControlModel oDM by given sName and properties dProps'''
    oCM = oDM.createInstance("com.sun.star.awt.UnoControl" + srv + "Model")
    while dProps:
        prp = dProps.popitem()
        uno.invoke(oCM, "setPropertyValue", (prp[0], prp[1]))
        # works with awt.UnoControlDialogElement only:
        oCM.Name = sName
    oDM.insertByName(sName, oCM)


def getYear(ctx, smgr):
    # from https://forum.openoffice.org/en/forum/viewtopic.php?t=56397
    oDM = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    oDM.Title = 'Introduce year'
    addAwtModel(oDM, 'Edit', 'year', {})
    addAwtModel(oDM, 'Button', 'btnOK',
                {
                    'Label': 'OK',
                    'DefaultButton': True,
                    'PushButtonType': 1,
                }
                )

    oDialog = smgr.createInstance("com.sun.star.awt.UnoControlDialog")
    oDialog.setModel(oDM)
    year = oDialog.getControl('year')
    # txtP = oDialog.getControl('txtPWD')
    h = 25
    y = 10
    for c in oDialog.getControls():
        c.setPosSize(10, y, 200, h, POSSIZE)
        y += h
    oDialog.setPosSize(300, 300, 300, y + h, POSSIZE)
    oDialog.setVisible(True)
    x = oDialog.execute()
    if x:
        return (year.getText())
    else:
        return False


def getTemplate(ctx, smgr):
    # from https://forum.openoffice.org/en/forum/viewtopic.php?t=56397
    oDM = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    oDM.Title = 'Introduce the path of the template'
    addAwtModel(oDM, 'Edit', 'template', {})
    addAwtModel(oDM, 'Button', 'btnOK',
                {
                    'Label': 'OK',
                    'DefaultButton': True,
                    'PushButtonType': 1,
                }
                )

    oDialog = smgr.createInstance("com.sun.star.awt.UnoControlDialog")
    oDialog.setModel(oDM)
    template = oDialog.getControl('template')
    # txtP = oDialog.getControl('txtPWD')
    h = 25
    y = 10
    for c in oDialog.getControls():
        c.setPosSize(10, y, 200, h, POSSIZE)
        y += h
    oDialog.setPosSize(300, 300, 300, y + h, POSSIZE)
    oDialog.setVisible(True)
    x = oDialog.execute()
    if x:
        return (template.getText())
    else:
        return False


def _formatWholeTable(table, row_count, column_count, orientation=VER_CENTER, border=False):
    # line = BorderLine()

    for column_name in list(ascii_uppercase)[:column_count]:
        for row_idx in range(1, row_count+1):
            cell = table.getCellByName(f"{column_name}{row_idx}")
            cell.VertOrient = orientation

            # if not border:
            #     cell.LeftBorder = line
            #     cell.RightBorder = line
            #     cell.TopBorder = line
            #     cell.BottomBorder = line


def _rgb_to_long(rgb_color):
    """Convert an RGB color tuple to libreoffice long integer format
    https://help.libreoffice.org/latest/en-US/text/sbasic/shared/03010306.html
    return => red*256*256 + green*256 + blue

    Args:
        rgb_color (tuple): (255,0,0)

    Returns:
        int: LO long int
    """
    r, g, b = rgb_color
    #return int(f'{r:02x}{g:02x}{b:02x}', 16)
    return r * 65536 + g * 256 + b


def _daterange(start_date, end_date):
    days = int((end_date - start_date).days)
    for n in range(days):
        yield start_date + datetime.timedelta(n)


def GenerateAll():
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    s_year = getYear(ctx, smgr)
    # s_year = "2025"
    year = int(s_year)

    template = getTemplate(ctx, smgr)
    # template = r"C:\Users\Leire\Google Drive\agenda\remarkable\template_rmk_2.odt"

    ConfigurePageForRMK()
    GenerateTitlePage(year)
    GenerateCalendar(year)
    GenerateMonthlyAgenda(year)
    GenerateDailyAgenda(year, template_file_path=template)


def ConfigurePageForRMK():
    # Get the doc from the scripting context which is made available to all scripts.
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    view_cursor = model.CurrentController.getViewCursor()
    page_style_name = view_cursor.PageStyleName
    style = model.StyleFamilies.getByName("PageStyles").getByName(page_style_name)

    style.BottomMargin = 400
    style.LeftMargin = 1300
    style.RightMargin = 400
    style.TopMargin = 400

    style.Width = 15770
    style.Height = 21030

    # Edit hyperlink style
    link_style = model.StyleFamilies.CharacterStyles.getByName('Internet link')
    link_style.CharColor = "6776679"
    link_style.CharUnderline = False


def GenerateTitlePage(year=None):
    # Get the doc from the scripting context which is made available to all scripts.
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    # get the XText interface
    text = model.Text

    if not year:
        s_year = getYear(ctx, smgr)
        year = int(s_year)
    else:
        s_year = str(year)

    text.End.CharHeight = 70
    text.End.CharFontName = 'Open Sans'
    text.End.CharWeight = FontWeight.BOLD
    text.End.CharColor = "6776679"
    text.End.ParaAdjust = HOR_CENTER
    text.End.String = f"\n{s_year}"

    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.BreakType = PAGE_AFTER
    text.insertControlCharacter(cursor.End, ControlCharacter.PARAGRAPH_BREAK, False)

    # dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    # frame = model.CurrentController.Frame
    # dispatcher.executeDispatch(frame, ".uno:uno:InsertPagebreak", "", 0, [])


def GenerateCalendar(year=None):
    # Get the doc from the scripting context which is made available to all scripts.
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    # get the XText interface
    text = model.Text

    if not year:
        s_year = getYear(ctx, smgr)
        year = int(s_year)
    else:
        s_year = str(year)

    # create an XTextRange at the end of the document
    # tRange = text.End


    # Calendar header
    # text.End.CharHeight = 1
    # text.End.String = '\n'

    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    # cursor.BreakType = PAGE_BOTH

    text.End.CharHeight = 12
    text.End.CharFontName = 'Open Sans'
    text.End.CharWeight = FontWeight.NORMAL
    text.End.CharColor = "6776679"
    text.End.ParaAdjust = HOR_CENTER
    text.End.String = f"CALENDAR {s_year}"

    # cursor = text.createTextCursor()
    # cursor.gotoEnd(False)
    # cursor.BreakType = PAGE_BOTH

    # Get table size
    col_month = [0] * 7 + [None] + [1] * 7 + [None] + [2] * 7
    row_month = [None, None] + [0] * 6 + [None] + [None, None] + [1] * 6 + [None] + [None, None] + [2] * 6 + [None] + [None, None] + [3] * 6 + [None]
    calendar_rows_count = 36 # weeks in a month (6) * number of months/3 (4) + 3 separators for each month row (12)
    calendar_column_count = 23 # days of week (7) * number of months in row (3) + 2 separators
    week_day_header = calendar.weekheader(1).split(" ")
    split_chunks = lambda lst, sz: [lst[i:i + sz] for i in range(0, len(lst), sz)]
    month_header = split_chunks(calendar.month_name[1:], 3)

    # Create the table
    calendar_table_name = "YearlyCalendarTable"

    # Remove the calendar table if it already exists
    # text_tables = model.TextTables
    # if text_tables.hasByName(calendar_table_name):
    #     calendar_table = text_tables.getByName(calendar_table_name)
    #     text.removeTextContent(calendar_table)

    calendar_table = model.createInstance("com.sun.star.text.TextTable")
    calendar_table.initialize(calendar_rows_count, calendar_column_count)

    # calendar_bookmark_name = "InsertCalendarHere"
    # if model.getBookmarks().hasByName(calendar_bookmark_name):
    #     insert_point = model.getBookmarks().getByName(calendar_bookmark_name).getAnchor()
    # else:
    #     insert_point = text.End
    insert_point = text.End
    insert_point.getText().insertTextContent(insert_point, calendar_table, False)

    # Format the table
    # calendar_table.getCellRangeByPosition(0, 0, 23, 36)
    cursor = calendar_table.createCursorByCellName("A1")
    cursor.goRight(22, True)
    cursor.goDown(35, True)
    # cursor.setPropertyValue("CharHeight", 7.0)
    cursor.CharHeight = 7.0
    cursor.CharFontName = 'Open Sans'
    cursor.CharWeight = FontWeight.NORMAL
    cursor.CharColor = "6776679"
    cursor.ParaAdjust = HOR_CENTER
    # cursor.VertOrient = VER_CENTER
    # cursor.ParaVertAlignment = VER_CENTER
    _formatWholeTable(calendar_table, calendar_rows_count, calendar_column_count, VER_CENTER)

    no_line = BorderLine2()
    table_border = calendar_table.TableBorder
    table_border.LeftLine = no_line
    table_border.RightLine = no_line
    table_border.TopLine = no_line
    table_border.BottomLine = no_line
    table_border.HorizontalLine = no_line
    table_border.VerticalLine = no_line
    calendar_table.TableBorder = table_border

    # Add data for the table
    for row_i in range(calendar_rows_count):
        row_type_index = row_i % 9
        month_row_idx = row_month[row_i]  # the index of the month in that column

        if row_type_index == 0:  # Name of the month row
            month_row_i = int(row_i/9)

            cursor = calendar_table.createCursorByCellName(f"A{row_i+1}")
            cursor.goRight(6, True)
            cursor.mergeRange()
            calendar_table.getCellByPosition(0, row_i).setString(f"{month_header[month_row_i][0]}")

            cursor = calendar_table.createCursorByCellName(f"C{row_i+1}")
            cursor.goRight(6, True)
            cursor.mergeRange()
            calendar_table.getCellByPosition(2, row_i).setString(f"{month_header[month_row_i][1]}")

            cursor = calendar_table.createCursorByCellName(f"E{row_i+1}")
            cursor.goRight(6, True)
            cursor.mergeRange()
            calendar_table.getCellByPosition(4, row_i).setString(f"{month_header[month_row_i][2]}")
        else:
            for col_i in range(calendar_column_count):
                day_of_week_idx = col_i % 8
                month_col_idx = col_month[col_i]  # the index of the month in that row

                if month_col_idx is None:  # is a separator column
                    continue

                if month_row_idx is not None:  # is a cell for a day of the month
                    month = 3 * month_row_idx + month_col_idx + 1  # get the month
                    week_number = row_type_index - 2
                    month_calendar_list = list(calendar.monthcalendar(year, month))
                    try:
                        day = month_calendar_list[week_number][day_of_week_idx]
                        if day != 0:  # do not add days that do not belong to the month
                            # DayTable244|table
                            datetime_day = datetime.datetime(year, month, day)
                            day_of_year = datetime_day.timetuple().tm_yday
                            calendar_table.getCellByPosition(col_i, row_i).setString(f"{day}")
                            cell_cursor = calendar_table.getCellByPosition(col_i, row_i).createTextCursor()
                            cell_cursor.gotoStart(False)
                            cell_cursor.gotoEnd(True)
                            cell_cursor.HyperLinkURL = f'#DayTable{day_of_year}|table'
                    except IndexError:
                        pass
                elif row_type_index == 1:  # is a cell for the day of the week (M, T, W, T, F, S, S)
                    calendar_table.getCellByPosition(col_i, row_i).setString(week_day_header[day_of_week_idx])

    calendar_table.TableName = calendar_table_name
    return None


def GenerateMonthlyAgenda(year=None):
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()
    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    # get the XText interface
    text = model.Text

    if not year:
        s_year = getYear(ctx, smgr)
        year = int(s_year)
    else:
        s_year = str(year)

    text.End.CharHeight = 12
    text.End.CharFontName = 'Open Sans'
    text.End.CharWeight = FontWeight.NORMAL
    text.End.CharColor = "6776679"
    text.End.ParaAdjust = HOR_CENTER

    no_line = BorderLine2()

    bottom_line = BorderLine2()
    bottom_line.Color = "6776679"
    bottom_line.InnerLineWidth = 10
    bottom_line.LineDistance = 0
    bottom_line.LineWidth = 5
    bottom_line.OuterLineWidth = 0
    bottom_line.LineStyle = SOLID

    week_day_header = calendar.weekheader(1).split(" ")
    for month_num, month in enumerate(calendar.month_name[1:], 1):
    # for month_num, month in enumerate(["January"], 1):
        # Insert Page break
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.BreakType = PAGE_BEFORE

        # Fill table
        text.End.String = month
        _, days_count = calendar.monthrange(year, month_num)

        month_table = model.createInstance("com.sun.star.text.TextTable")
        month_table.initialize(days_count, 3)

        insert_point = text.End
        insert_point.getText().insertTextContent(insert_point, month_table, False)

        cursor = month_table.createCursorByCellName("A1")
        cursor.goRight(2, True)
        cursor.goDown(days_count-1, True)
        cursor.CharHeight = 8.0
        cursor.CharFontName = 'Open Sans'
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharColor = "6776679"
        cursor.ParaAdjust = HOR_CENTER
        _formatWholeTable(month_table, days_count, 3, orientation=VER_CENTER)

        table_border = month_table.TableBorder
        table_border.LeftLine = no_line
        table_border.RightLine = no_line
        table_border.TopLine = no_line
        table_border.BottomLine = bottom_line
        table_border.HorizontalLine = bottom_line
        table_border.VerticalLine = no_line
        month_table.TableBorder = table_border

        sep = month_table.TableColumnSeparators
        sep[0].Position = 400
        sep[1].Position = 800
        month_table.TableColumnSeparators = sep

        for day_idx in range(0, days_count):
            day = day_idx + 1
            datetime_day = datetime.datetime(year, month_num, day)
            day_of_year = datetime_day.timetuple().tm_yday

            week_day = datetime.datetime(year, month_num, day).weekday()
            month_table.getCellByPosition(0, day_idx).setString(day)
            month_table.getCellByPosition(1, day_idx).setString(week_day_header[week_day])

            cell_cursor = month_table.getCellByPosition(0, day_idx).createTextCursor()
            cell_cursor.gotoStart(False)
            cell_cursor.gotoEnd(True)
            cell_cursor.HyperLinkURL = f'#DayTable{day_of_year}|table'

        month_table.TableName = f"MontlyAgenda{month}Table"

    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.BreakType = PAGE_BEFORE

    text.insertControlCharacter(cursor.End, ControlCharacter.PARAGRAPH_BREAK, False)


def GenerateDailyAgenda(year=None, months=None, template_file_path=None, test=False):
    """
    :param year: int with the year
    :param months: tuple with the starting and end month. For example (2, 5) will generate February to May. If None, it
     will generate the whole year.
    :param template_file_path: full path to the template of a day
    :param test: if True it will only generate a few days
    :return:
    """
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()
        doc = XSCRIPTCONTEXT.getDocument()
    except NameError:
        # get the uno component context from the PyUNO runtime
        localContext = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         localContext)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    model = desktop.getCurrentComponent()

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, ())

    # get the XText interface
    text = model.Text

    if not year:
        s_year = getYear(ctx, smgr)
        year = int(s_year)
    else:
        s_year = str(year)

    if not template_file_path:
        template_file_path = getTemplate(ctx, smgr)

    # template_file_path = r"C:\Users\Leire\Google Drive\filofax\template_rmk_2.odt"
    file_url = uno.systemPathToFileUrl(template_file_path)
    template_doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    template_table = template_doc.getTextTables().getByName("DayTable")

    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    frame = template_doc.CurrentController.Frame
    dispatcher.executeDispatch(frame, ".uno:ClearClipboard", "", 0, [])
    template_doc.CurrentController.select(template_table)

    cursor = template_doc.getCurrentController().getViewCursor()
    cursor.goRight(5, True)
    cursor.goRight(5, True)
    cursor.goDown(1, True)

    # Move the cursor to the end of the target document
    cursor = model.getCurrentController().getViewCursor()
    cursor.gotoEnd(False)

    # Create the montly calendar for all months
    month_calendars = [None] * 13
    for month_num in range(1, 13):
        month_calendar = list(calendar.monthcalendar(year, month_num))
        month_calendar.insert(0, calendar.weekheader(1).split(" "))

        month_calendar = [list(map(lambda x: x if x != 0 else '', i)) for i in month_calendar]

        while len(month_calendar) < 7:
            month_calendar.append([''] * 7)

        month_calendars[month_num] = month_calendar

    if test:
        first_day = datetime.datetime(year, 8, 31)
        last_day = datetime.datetime(year, 9, 3)
    elif months:
        first_day = datetime.datetime(year, months[0], 1)
        last_day = datetime.datetime(year + 1, months[1]+1, 1)
    else:
        first_day = datetime.datetime(year, 1,1)
        last_day = datetime.datetime(year+1, 1, 1)

    for day in _daterange(first_day, last_day):
        day_num = day.day
        month_num = day.month
        month = calendar.month_name[month_num]
        week_day = calendar.day_name[day.weekday()]
        week_number = day.isocalendar().week
        day_of_year = day.timetuple().tm_yday  # returns 1 for January 1st
        month_calendar = month_calendars[month_num]

        if day_num == 1:
            # Insert Page break
            text.End.CharHeight = 3
            text.End.String = ' '

            # cursor = text.createTextCursor()
            # cursor.gotoEnd(False)
            # cursor.BreakType = PAGE_BOTH

            text.End.CharHeight = 70
            text.End.CharFontName = 'Open Sans'
            text.End.CharWeight = FontWeight.BOLD
            text.End.CharColor = "6776679"
            text.End.ParaAdjust = HOR_CENTER
            text.End.String = month

            # cursor = text.createTextCursor()
            # cursor.gotoEnd(False)
            # cursor.BreakType = PAGE_BOTH
            #
            # text.End.CharHeight = 3
            # text.End.String = "test"

        # Create cursor before inserting table
        # cursor_before = text.createTextCursor()
        # cursor_before.gotoEnd(False)

        # Paste the copied table
        model.getCurrentController().insertTransferable(frame.Controller.getTransferable())

        text_tables = model.TextTables
        day_table = text_tables.getByName("DayTable")
        day_table.TableName = f"DayTable{day_of_year}"

        search_cursor = day_table.getCellByName("A1").createTextCursor()

        # Insert hyperlinks
        calendar_icon = model.GraphicObjects.getByName("CalendarIcon")
        calendar_icon.HyperLinkURL = '#YearlyCalendarTable|table'
        calendar_icon.setName(f"DailyCalendarIcon{day_of_year}")

        for cal_month_num, cal_month in enumerate(calendar.month_name[1:], 1):
            search = model.createSearchDescriptor()
            # search.SearchBackwards = True
            # search.SearchAll = False
            search.setSearchString(calendar.month_abbr[cal_month_num].upper())
            found = model.findNext(search_cursor, search)
            if found:
                found.HyperLinkURL = f'#MontlyAgenda{cal_month}Table|table'
            else:
                print("E")

        replace = model.createReplaceDescriptor()
        replace.setSearchString("<d")
        replace.setReplaceString(str(day_num))
        model.replaceAll(replace)

        replace = model.createReplaceDescriptor()
        replace.setSearchString("<MONTH>")
        replace.setReplaceString(month)
        model.replaceAll(replace)

        replace = model.createReplaceDescriptor()
        replace.setSearchString("<WEEKDAY>")
        replace.setReplaceString(week_day)
        model.replaceAll(replace)

        replace = model.createReplaceDescriptor()
        replace.setSearchString("<WEEKNUMBER>")
        replace.setReplaceString(week_number)
        model.replaceAll(replace)

        cursor = model.getCurrentController().getViewCursor()
        cursor.gotoEnd(False)

        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.BreakType = PAGE_AFTER

        text.End.String = " "

        if text_tables.hasByName("CalendarTable"):
            calendar_table = text_tables.getByName("CalendarTable")
            calendar_table.setDataArray(month_calendar)
            calendar_table.TableName = f"DailyCalendarTable{day_of_year}"

            for row_idx, row in enumerate(month_calendar):
                for column_idx, value in enumerate(row):
                    try:
                        int(value)
                    except ValueError:
                        pass
                    else:
                        if value != "0":
                            if value == day_num:
                                cursor = calendar_table.createCursorByCellName(f"{list(ascii_uppercase)[column_idx]}{row_idx+1}")
                                cursor.CharWeight = FontWeight.BOLD

                            calendar_day_datetime = datetime.datetime(year, month_num, int(value))
                            calendar_day_of_year = calendar_day_datetime.timetuple().tm_yday

                            cell_cursor = calendar_table.getCellByPosition(column_idx, row_idx).createTextCursor()
                            cell_cursor.gotoStart(False)
                            cell_cursor.gotoEnd(True)
                            cell_cursor.HyperLinkURL = f'#DayTable{calendar_day_of_year}|table'


    template_doc.close(True)
