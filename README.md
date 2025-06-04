# LibreOffice Agenda Generator

A LibreOffice automation script written in Python using the UNO API. This tool generates a fully hyperlinked yearly agenda — including a title page, yearly calendar, monthly overviews, and daily pages based on a customizable template.

Each generated daily page includes links back to the corresponding **monthly** and **yearly** views, making it easy to navigate across the planner. Ideal for reMarkable tablet users, printable planners, or anyone wanting a personalized, interactive agenda.

## Installation

1. **Requirements**:
   - LibreOffice with built-in Python support (standard with most installations).
   - The Python file should be placed in your LibreOffice user macros folder. For most installations on Windows, this is:

     ```
     C:\Users\<YourUsername>\AppData\Roaming\LibreOffice\4\user\Scripts\python
     ```

     Replace `<YourUsername>` with your Windows username.

2. **Included Files**:
   - A ready-to-use `.odt` daily template is included in the repository. You can use it directly or customize it to suit your needs.

## Usage

1. **Open LibreOffice normally** and create a new blank Writer document.
2. Navigate to:  
   `Tools > Macros > Run Macro...`
3. In the dialog, select:  
   `My Macros > LibreOfficeAgendaGenerator > GenerateAll`  
   and click **Run**.
4. The script will prompt you to:
   - Enter the **year** for which to generate the agenda.
   - Provide the full **path to the daily template** file.

Once you confirm, the script will automatically generate:
- A title page.
- A full-year calendar with links to each day.
- Monthly overview pages.
- Hyperlinked daily pages, with each day linked to its monthly and yearly view.

## Template Customization

The agenda uses a `.odt` template for daily pages, included in the repository. The template uses placeholder fields that will be automatically replaced by the script. These fields are:

| Field          | Description                               |
|----------------|-------------------------------------------|
| `<d`          | Day of the month (1–31)                   |
| `<MONTH>`      | Full name of the month (e.g., March)      |
| `<WEEKDAY>`    | Name of the weekday (e.g., Monday)        |
| `<WEEKNUMBER>` | ISO week number of the date               |

The script also insert several helpful hyperlinks:

- **Daily Pages**:
  - A hyperlink is added to an image named `"CalendarIcon"` that links back to the **yearly calendar**.
  - Hyperlinks are added to the abbreviated month names (`"JAN"`, `"FEB"`, etc.) that link to their corresponding **monthly overview pages**.
  - If a table named `"CalendarTable"` is present in the template, it is replaced with an auto-generated calendar for the month. Each day in this table includes a hyperlink to the corresponding **daily page**.

You can customize the appearance, layout, or add/remove sections in the template as long as these fields are present where needed.

## License

[GNU GPLv3](https://choosealicense.com/licenses/gpl-3.0/)

## Attributions

[Calendar icon created by Freepik - Flaticon](https://www.flaticon.com/free-icons/calendar)
