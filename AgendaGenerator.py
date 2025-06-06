import calendar
import datetime
from string import ascii_uppercase
import uno

from com.sun.star.awt import FontWeight
from com.sun.star.awt.PosSize import POSSIZE
from com.sun.star.beans import PropertyValue
from com.sun.star.style.BreakType import PAGE_AFTER, PAGE_BEFORE
from com.sun.star.style.ParagraphAdjust import CENTER as HOR_CENTER
from com.sun.star.table import BorderLine2
from com.sun.star.table.BorderLineStyle import SOLID
from com.sun.star.text import ControlCharacter
from com.sun.star.text.VertOrientation import CENTER as VER_CENTER


# variables for page configuration (reMarkable)
# Margins for the page ('top', 'bottom', 'left', 'right').
PAGE_MARGINS = {'top': 400, 'bottom': 400, 'left': 1300, 'right': 400}
PAGE_SIZE = {'width': 15770, 'height': 21030}  # Page size ('width', 'height').
LINK_COLOR = "6776679"  # Hyperlink color
LINK_UNDERLINE = False  # Whether hyperlinks are underlined.

# variables for default text style and border styles for month headers in the monthly agenda
MONTH_HEADER_CHAR_HEIGHT = 12  # Font size for month headers
MONTH_HEADER_FONT_NAME = 'Open Sans'  # Font name for month headers
MONTH_HEADER_FONT_WEIGHT = FontWeight.NORMAL  # Font weight for month headers
MONTH_HEADER_COLOR = "6776679"  # Font color for month headers
MONTH_HEADER_ALIGN = HOR_CENTER  # Paragraph alignment for month headers

# Border style variables
MONTH_TABLE_BORDER_COLOR = "6776679"  # Border color for month tables
MONTH_TABLE_BOTTOM_INNER_WIDTH = 10  # Inner line width for bottom border
MONTH_TABLE_BOTTOM_LINE_DISTANCE = 0  # Line distance for bottom border
MONTH_TABLE_BOTTOM_LINE_WIDTH = 5  # Line width for bottom border
MONTH_TABLE_BOTTOM_OUTER_WIDTH = 0  # Outer line width for bottom border
MONTH_TABLE_LINE_STYLE = SOLID  # Line style for borders


def _get_office_context():
    """
        Obtain the UNO component context, desktop, service manager, and the current Writer document model for 
        LibreOffice.

        Tries to use the XSCRIPTCONTEXT if available (when running as a macro inside LibreOffice).
        If not available, connects to a running LibreOffice instance via socket.

        Returns:
            tuple: (ctx, desktop, smgr, model)
                ctx: The UNO component context.
                desktop: The central desktop object.
                smgr: The UNO service manager.
                model: The current Writer document model.
        """
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        ctx = XSCRIPTCONTEXT.getComponentContext()
        smgr = ctx.getServiceManager()

    except NameError:
        # get the uno component context from the PyUNO runtime
        local_context = uno.getComponentContext()

        # create the UnoUrlResolver
        resolver = local_context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",
                                                                         local_context)

        # connect to the running office
        ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        smgr = ctx.ServiceManager

        # get the central desktop object
        desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    # Check whether there's already an opened document.
    # Otherwise, create a new one
    model = desktop.getCurrentComponent()
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())

    return ctx, desktop, smgr, model


def _add_awt_model(dialog_model, srv, control_name, control_prop):
    """
    Inserts a UnoControl<srv>Model into the given DialogControlModel (dialog_model) with the specified name
    (control_name) and properties (control_prop).

    Args:
        dialog_model: The dialog control model to which the control will be added.
        srv: The type of UnoControl (e.g., 'Edit', 'Button').
        control_name: The name to assign to the control.
        control_prop: A dictionary of property names and values to set on the control.

    Returns:
        None
    """
    control_model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + srv + "Model")
    while control_prop:
        prp = control_prop.popitem()
        uno.invoke(control_model, "setPropertyValue", (prp[0], prp[1]))
        # works with awt.UnoControlDialogElement only:
        control_model.Name = control_name
    dialog_model.insertByName(control_name, control_model)


def _show_input_dialog(smgr, title, field_name):
    """
    Displays a simple input dialog in LibreOffice using UNO API.

    Args:
        smgr: The UNO service manager.
        title (str): The title of the dialog window.
        field_name (str): The name of the input field.

    Returns:
        str or bool: The text entered by the user if OK is pressed, otherwise False.
    """
    dialog_model = smgr.createInstance("com.sun.star.awt.UnoControlDialogModel")
    dialog_model.Title = title
    _add_awt_model(dialog_model, 'Edit', field_name, {})
    _add_awt_model(dialog_model, 'Button', 'btnOK', {
        'Label': 'OK',
        'DefaultButton': True,
        'PushButtonType': 1,
    })

    dialog = smgr.createInstance("com.sun.star.awt.UnoControlDialog")
    dialog.setModel(dialog_model)
    field = dialog.getControl(field_name)
    h = 25
    y = 10
    for c in dialog.getControls():
        c.setPosSize(10, y, 200, h, POSSIZE)
        y += h
    dialog.setPosSize(300, 300, 300, y + h, POSSIZE)
    dialog.setVisible(True)
    x = dialog.execute()
    if x:
        return field.getText()
    else:
        return False


def _get_year(smgr):
    return _show_input_dialog(smgr, 'Introduce year', 'year')


def _get_template(smgr):
    return _show_input_dialog(smgr, 'Introduce the path of the template', 'template')


def _format_whole_table(table, row_count: int, column_count: int, orientation=VER_CENTER):
    """
    Formats the entire table by setting the vertical orientation of each cell.

    Args:
        table: LibreOffice table object.
        row_count (int): Number of rows in the table.
        column_count (int): Number of columns in the table.
        orientation: Vertical orientation for the cells (default is VER_CENTER).
    """
    for column_name in list(ascii_uppercase)[:column_count]:
        for row_idx in range(1, row_count+1):
            cell = table.getCellByName(f"{column_name}{row_idx}")
            cell.VertOrient = orientation

def _rgb_to_long(rgb_color: tuple[int, int, int]) -> int:
    """
    Converts an RGB color tuple to LibreOffice long integer format.

    Args:
        rgb_color (tuple[int, int, int]): A tuple representing the RGB color, e.g., (255, 0, 0).

    Returns:
        int: The color as a LibreOffice long integer (red\*256\*256 + green\*256 + blue).
    """
    r, g, b = rgb_color
    # return int(f'{r:02x}{g:02x}{b:02x}', 16)
    return r * 65536 + g * 256 + b

def _daterange(start_date: datetime.date, end_date: datetime.date):
    """
    Generator that yields each date from start_date up to, but not including, end_date.

    Args:
        start_date (datetime.date or datetime.datetime): The start date (inclusive).
        end_date (datetime.date or datetime.datetime): The end date (not included).

    Yields:
        datetime.date or datetime.datetime: Each date in the range.
    """
    days = int((end_date - start_date).days)
    for n in range(days):
        yield start_date + datetime.timedelta(n)


def generate_all() -> None:
    """
    Generates the complete agenda document in LibreOffice Writer. Configures the page for the reMarkable and
    generates the title page, yearly calendar, monthly agenda, and daily agenda.

    Returns:
        None
    """
    ctx, desktop, smgr, model = _get_office_context()

    s_year = _get_year(smgr)
    # s_year = "2025"
    year = int(s_year)

    template = _get_template(smgr)
    # template = r"C:\Users\Leire\Google Drive\agenda\remarkable\template_rmk_2.odt"

    configure_page_for_rmk()
    generate_title_page(year)
    generate_calendar(year)
    generate_monthly_agenda(year)
    generate_daily_agenda(year, template_path=template)


def configure_page_for_rmk() -> None:
    """
    Configures the page layout in LibreOffice Writer for reMarkable devices using global variables.
    Sets custom page margins, page size, and hyperlink style based on the global configuration.

    Uses:
        RMK_MARGINS (dict): Margins for the page ('top', 'bottom', 'left', 'right').
        RMK_SIZE (dict): Page size ('width', 'height').
        RMK_LINK_COLOR (str): Hyperlink color.
        RMK_LINK_UNDERLINE (bool): Whether hyperlinks are underlined.
    """
    ctx, desktop, smgr, model = _get_office_context()

    view_cursor = model.CurrentController.getViewCursor()
    page_style_name = view_cursor.PageStyleName
    style = model.StyleFamilies.getByName("PageStyles").getByName(page_style_name)

    style.TopMargin = PAGE_MARGINS['top']
    style.BottomMargin = PAGE_MARGINS['bottom']
    style.LeftMargin = PAGE_MARGINS['left']
    style.RightMargin = PAGE_MARGINS['right']

    style.Width = PAGE_SIZE['width']
    style.Height = PAGE_SIZE['height']

    # Edit hyperlink style
    link_style = model.StyleFamilies.CharacterStyles.getByName('Internet link')
    link_style.CharColor = LINK_COLOR
    link_style.CharUnderline = LINK_UNDERLINE


def generate_title_page(year: int = None) -> None:
    """
    Generates the title page for the agenda document in LibreOffice Writer.

    If no year is provided, prompts the user to input the year. Sets the title page
    with a large font size, specific font, bold weight, and color. Inserts a page break
    after the title page.

    Args:
        year (int, optional): The year to display on the title page. If None, the user is prompted.

    Returns:
        None
    """
    # Get the doc from the scripting context which is made available to all scripts.
    ctx, desktop, smgr, model = _get_office_context()

    # get the XText interface
    text = model.Text

    if not year:
        s_year = _get_year(smgr)
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


def generate_calendar(year: int = None) -> None:
    """
    Generates a yearly calendar table in LibreOffice Writer.

    Args:
        year (int, optional): The year for the calendar. If None, prompts the user.

    Returns:
        None
    """
    # Get the document context
    ctx, desktop, smgr, model = _get_office_context()
    text = model.Text

    # Get the year as string and int
    if not year:
        s_year = _get_year(smgr)
        year = int(s_year)
    else:
        s_year = str(year)

    # Insert calendar header
    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    text.End.CharHeight = 12
    text.End.CharFontName = 'Open Sans'
    text.End.CharWeight = FontWeight.NORMAL
    text.End.CharColor = "6776679"
    text.End.ParaAdjust = HOR_CENTER
    text.End.String = f"CALENDAR {s_year}"

    # Table configuration
    col_month = [0] * 7 + [None] + [1] * 7 + [None] + [2] * 7
    row_month = (
        [None, None] + [0] * 6 + [None] +
        [None, None] + [1] * 6 + [None] +
        [None, None] + [2] * 6 + [None] +
        [None, None] + [3] * 6 + [None]
    )
    calendar_rows_count = 36  # weeks in a month (6) * number of months/3 (4) + 3 separators for each month row (12)
    calendar_column_count = 23  # days of week (7) * number of months in row (3) + 2 separators
    week_day_header = calendar.weekheader(1).split(" ")
    split_chunks = lambda lst, sz: [lst[i:i + sz] for i in range(0, len(lst), sz)]
    month_header = split_chunks(calendar.month_name[1:], 3)
    calendar_table_name = "YearlyCalendarTable"

    # Create the table
    calendar_table = model.createInstance("com.sun.star.text.TextTable")
    calendar_table.initialize(calendar_rows_count, calendar_column_count)
    insert_point = text.End
    insert_point.getText().insertTextContent(insert_point, calendar_table, False)

    # Format the table
    cursor = calendar_table.createCursorByCellName("A1")
    cursor.goRight(calendar_column_count - 1, True)
    cursor.goDown(calendar_rows_count - 1, True)
    cursor.CharHeight = 7.0
    cursor.CharFontName = 'Open Sans'
    cursor.CharWeight = FontWeight.NORMAL
    cursor.CharColor = "6776679"
    cursor.ParaAdjust = HOR_CENTER
    _format_whole_table(calendar_table, calendar_rows_count, calendar_column_count, VER_CENTER)

    # Remove all borders
    no_line = BorderLine2()
    table_border = calendar_table.TableBorder
    table_border.LeftLine = no_line
    table_border.RightLine = no_line
    table_border.TopLine = no_line
    table_border.BottomLine = no_line
    table_border.HorizontalLine = no_line
    table_border.VerticalLine = no_line
    calendar_table.TableBorder = table_border

    # Fill the table with month names, weekdays, and days
    for row_i in range(calendar_rows_count):
        row_type_index = row_i % 9
        month_row_idx = row_month[row_i]  # Index of the month in this row

        if row_type_index == 0:  # Month name row
            month_row_i = int(row_i / 9)
            # Merge and set month names for each of the 3 months in the row
            for col, col_name in zip([0, 2, 4], range(3)):
                cursor = calendar_table.createCursorByCellName(f"{chr(65 + col)}{row_i + 1}")
                cursor.goRight(6, True)
                cursor.mergeRange()
                calendar_table.getCellByPosition(col, row_i).setString(f"{month_header[month_row_i][col_name]}")
        else:
            for col_i in range(calendar_column_count):
                day_of_week_idx = col_i % 8
                month_col_idx = col_month[col_i]  # Index of the month in this column

                if month_col_idx is None:
                    continue  # Separator column

                if month_row_idx is not None:
                    # Day cell: fill with day number and hyperlink
                    month = 3 * month_row_idx + month_col_idx + 1  # get the month
                    week_number = row_type_index - 2
                    month_calendar_list = list(calendar.monthcalendar(year, month))
                    try:
                        day = month_calendar_list[week_number][day_of_week_idx]
                        if day != 0:  # do not add days that do not belong to the month
                            datetime_day = datetime.datetime(year, month, day)
                            day_of_year = datetime_day.timetuple().tm_yday
                            calendar_table.getCellByPosition(col_i, row_i).setString(f"{day}")
                            cell_cursor = calendar_table.getCellByPosition(col_i, row_i).createTextCursor()
                            cell_cursor.gotoStart(False)
                            cell_cursor.gotoEnd(True)
                            cell_cursor.HyperLinkURL = f'#DayTable{day_of_year}|table'
                    except IndexError:
                        pass
                elif row_type_index == 1:
                    # Weekday header row (M, T, W, T, F, S, S)
                    calendar_table.getCellByPosition(col_i, row_i).setString(week_day_header[day_of_week_idx])

    calendar_table.TableName = calendar_table_name
    return None


def generate_monthly_agenda(year: int = None) -> None:
    """
    Generates a monthly agenda for each month of the given year in LibreOffice Writer.

    Args:
        year (int, optional): The year for which to generate the agenda. If None, prompts the user.

    Returns:
        None
    """
    ctx, desktop, smgr, model = _get_office_context()
    text = model.Text

    # Get year as string and int
    if not year:
        s_year = _get_year(smgr)
        year = int(s_year)


    # Set default text style for month headers
    text.End.CharHeight = MONTH_HEADER_CHAR_HEIGHT
    text.End.CharFontName = MONTH_HEADER_FONT_NAME
    text.End.CharWeight = MONTH_HEADER_FONT_WEIGHT
    text.End.CharColor = MONTH_HEADER_COLOR
    text.End.ParaAdjust = MONTH_HEADER_ALIGN

    # Define border styles
    no_line = BorderLine2()

    bottom_line = BorderLine2()
    bottom_line.Color = MONTH_TABLE_BORDER_COLOR
    bottom_line.InnerLineWidth = MONTH_TABLE_BOTTOM_INNER_WIDTH
    bottom_line.LineDistance = MONTH_TABLE_BOTTOM_LINE_DISTANCE
    bottom_line.LineWidth = MONTH_TABLE_BOTTOM_LINE_WIDTH
    bottom_line.OuterLineWidth = MONTH_TABLE_BOTTOM_OUTER_WIDTH
    bottom_line.LineStyle = MONTH_TABLE_LINE_STYLE

    week_day_header = calendar.weekheader(1).split(" ")

    for month_num, month in enumerate(calendar.month_name[1:], 1):
        # Insert a page break before each month
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.BreakType = PAGE_BEFORE

        # Insert month name as header
        text.End.String = month
        _, days_count = calendar.monthrange(year, month_num)

        # Create and insert the table for the month
        month_table = model.createInstance("com.sun.star.text.TextTable")
        month_table.initialize(days_count, 3)
        insert_point = text.End
        insert_point.getText().insertTextContent(insert_point, month_table, False)

        # Format the table
        cursor = month_table.createCursorByCellName("A1")
        cursor.goRight(2, True)
        cursor.goDown(days_count - 1, True)
        cursor.CharHeight = 8.0
        cursor.CharFontName = 'Open Sans'
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharColor = "6776679"
        cursor.ParaAdjust = HOR_CENTER
        _format_whole_table(month_table, days_count, 3, orientation=VER_CENTER)

        # Set table borders
        table_border = month_table.TableBorder
        table_border.LeftLine = no_line
        table_border.RightLine = no_line
        table_border.TopLine = no_line
        table_border.BottomLine = bottom_line
        table_border.HorizontalLine = bottom_line
        table_border.VerticalLine = no_line
        month_table.TableBorder = table_border

        # Set column separators
        sep = month_table.TableColumnSeparators
        sep[0].Position = 400
        sep[1].Position = 800
        month_table.TableColumnSeparators = sep

        # Fill table with days and hyperlinks
        for day_idx in range(days_count):
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

    # Insert a page break at the end
    cursor = text.createTextCursor()
    cursor.gotoEnd(False)
    cursor.BreakType = PAGE_BEFORE

    text.insertControlCharacter(cursor.End, ControlCharacter.PARAGRAPH_BREAK, False)


def generate_daily_agenda(year: int = None, months: tuple[int, int] = None, template_path: str = None,
                          test: bool = False) -> None:
    """
    Generates the daily agenda pages for a given year in LibreOffice Writer.

    Args:
        year (int, optional): The year for which to generate the agenda. If None, prompts the user.
        months (tuple[int, int], optional): Tuple with the starting and ending month (e.g., (2, 5) for February to May).
                                            If None, generates for the whole year.
        template_path (str, optional): Full path to the daily template file.
        test (bool, optional): If True, generates only a few days for testing.

    Returns:
        None
    """
    ctx, desktop, smgr, model = _get_office_context()
    text = model.Text

    # Get year as string and int
    if not year:
        s_year = _get_year(smgr)
        year = int(s_year)

    # Get template file path if not provided
    if not template_path:
        template_path = _get_template(smgr)

    # Load the template document and get the day table
    file_url = uno.systemPathToFileUrl(template_path)
    template_doc = desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    template_table = template_doc.getTextTables().getByName("DayTable")

    # Prepare dispatcher and select the table in the template
    dispatcher = smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    frame = template_doc.CurrentController.Frame
    dispatcher.executeDispatch(frame, ".uno:ClearClipboard", "", 0, [])
    template_doc.CurrentController.select(template_table)

    # Move the cursor in the template (for copy-paste)
    cursor = template_doc.getCurrentController().getViewCursor()
    cursor.goRight(5, True)
    cursor.goRight(5, True)
    cursor.goDown(1, True)

    # Move the cursor to the end of the target document
    cursor = model.getCurrentController().getViewCursor()
    cursor.gotoEnd(False)

    # Prepare monthly calendars for all months
    month_calendars = [None] * 13
    for month_num in range(1, 13):
        month_calendar = list(calendar.monthcalendar(year, month_num))
        month_calendar.insert(0, calendar.weekheader(1).split(" "))
        month_calendar = [list(map(lambda x: x if x != 0 else '', i)) for i in month_calendar]
        while len(month_calendar) < 7:
            month_calendar.append([''] * 7)
        month_calendars[month_num] = month_calendar

    # Determine the date range to generate
    if test:
        first_day = datetime.datetime(year, 8, 31)
        last_day = datetime.datetime(year, 9, 3)
    elif months:
        first_day = datetime.datetime(year, months[0], 1)
        last_day = datetime.datetime(year + 1, months[1] + 1, 1)
    else:
        first_day = datetime.datetime(year, 1, 1)
        last_day = datetime.datetime(year + 1, 1, 1)

    for day in _daterange(first_day, last_day):
        day_num = day.day
        month_num = day.month
        month = calendar.month_name[month_num]
        week_day = calendar.day_name[day.weekday()]
        week_number = day.isocalendar().week
        day_of_year = day.timetuple().tm_yday
        month_calendar = month_calendars[month_num]

        # Insert month header and page break at the start of each month
        if day_num == 1:
            text.End.CharHeight = 3
            text.End.String = ' '
            text.End.CharHeight = 70
            text.End.CharFontName = 'Open Sans'
            text.End.CharWeight = FontWeight.BOLD
            text.End.CharColor = "6776679"
            text.End.ParaAdjust = HOR_CENTER
            text.End.String = month

        # Paste the copied table from the template
        model.getCurrentController().insertTransferable(frame.Controller.getTransferable())

        text_tables = model.TextTables
        day_table = text_tables.getByName("DayTable")
        day_table.TableName = f"DayTable{day_of_year}"

        search_cursor = day_table.getCellByName("A1").createTextCursor()

        # Insert hyperlinks for navigation
        calendar_icon = model.GraphicObjects.getByName("CalendarIcon")
        calendar_icon.HyperLinkURL = '#YearlyCalendarTable|table'
        calendar_icon.setName(f"DailyCalendarIcon{day_of_year}")

        for cal_month_num, cal_month in enumerate(calendar.month_name[1:], 1):
            search = model.createSearchDescriptor()
            search.setSearchString(calendar.month_abbr[cal_month_num].upper())
            found = model.findNext(search_cursor, search)
            if found:
                found.HyperLinkURL = f'#MontlyAgenda{cal_month}Table|table'
            else:
                print("E")

        # Replace placeholders in the template
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

        # Insert a page break after each day
        cursor = model.getCurrentController().getViewCursor()
        cursor.gotoEnd(False)
        cursor = text.createTextCursor()
        cursor.gotoEnd(False)
        cursor.BreakType = PAGE_AFTER
        text.End.String = " "

        # Update the calendar table if present
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
                                cell_cursor = calendar_table.createCursorByCellName(
                                    f"{list(ascii_uppercase)[column_idx]}{row_idx+1}"
                                )
                                cell_cursor.CharWeight = FontWeight.BOLD

                            calendar_day_datetime = datetime.datetime(year, month_num, int(value))
                            calendar_day_of_year = calendar_day_datetime.timetuple().tm_yday

                            cell_cursor = calendar_table.getCellByPosition(column_idx, row_idx).createTextCursor()
                            cell_cursor.gotoStart(False)
                            cell_cursor.gotoEnd(True)
                            cell_cursor.HyperLinkURL = f'#DayTable{calendar_day_of_year}|table'

    template_doc.close(True)
