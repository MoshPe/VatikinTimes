# Vatikin Times

Vatikin Times is a Windows desktop utility for generating a weekly Hebrew Excel sheet for בית הכנסת משמרת -- ותיקין. It pulls zmanim, Hebrew dates, parasha data, Daf Yomi, holidays, Rosh Chodesh details, molad information, Yom Tov candle-lighting, and haftara references, then formats everything into a printable right-to-left `.xlsx` schedule.

## Features

- City selection for קריות, חיפה, ירושלים, תל אביב, and באר שבע.
- Israel parasha schedule via Hebcal, including correct Israel-only parasha splitting.
- Weekly zmanim for:
  - עלות השחר
  - זמן טלית ותפילין
  - הנץ החמה (הנראה)
  - שחרית ותיקין
- Chai Tables netz lookup with Hebcal fallback behavior.
- Shabbat and Yom Tov Vatikin offsets.
- Daf Yomi row.
- ימים מיוחדים row with:
  - holidays
  - Rosh Chodesh
  - molad text
  - כניסת החג for Erev Yom Tov, using the selected city
- Haftara line in Hebrew, including reference and opening words where available.
- Printable Excel layout with right-to-left sheet direction, A4 landscape page setup, and one-page print scaling.
- Free-form comments shown near the bottom of the sheet.

## Data Sources

The app uses these external sources at runtime:

- Hebcal calendar API for parashiyot, Hebrew dates, Daf Yomi, holidays, Rosh Chodesh, molad, and candle-lighting.
- Hebcal zmanim API for standard zmanim.
- Hebcal leyning API for haftara references.
- Sefaria API for Hebrew haftara opening words.
- Chai Tables for visible netz values, where available.

Network access is required while generating a schedule.

## Requirements

- Python 3.11+ recommended.
- Windows is the primary target because the app uses PyQt6 and opens the generated Excel file through the OS.
- Excel or another `.xlsx` viewer for reviewing/printing the output.

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

If your default `python` does not have the required packages, use the Python installation that has PyQt6 and openpyxl installed.

## Running

```powershell
python main.py
```

Then:

1. Select a city.
2. Select a parasha.
3. Optionally enter comments.
4. Click `עדכן והכן להדפסה`.

The app writes an `.xlsx` file in the project directory and tries to open it automatically.

## Building an Executable

The repository includes an `install` command file with the PyInstaller command:

```powershell
pyinstaller --onefile --noconsole --hidden-import openpyxl.cell._writer main.py
```

Generated Excel files are ignored by git.

## Excel Output Notes

The workbook is configured as:

- right-to-left worksheet
- A4 landscape
- print area limited to the generated schedule columns
- fit to one printed page
- Hebrew title/date/day headers styled with David where needed
- wrapped rows for long Hebrew content such as molad and כניסת החג

If Excel still shows clipping in normal sheet view, use print preview or manually widen/expand as needed. The export is optimized for printing the generated area on a single page.

## Testing

Run the unit tests:

```powershell
python -m unittest test_main.py
```

The tests cover:

- weekly event alignment
- Hebrew year extraction
- Yom Tov detection
- selected-city candle-lighting request parameters
- Israel parasha API parameters
- haftara formatting and opening words
- Excel layout and print settings

## Project Structure

- `main.py` - PyQt6 app, API integration, schedule generation, and Excel export.
- `test_main.py` - unit tests for data formatting and workbook output.
- `requirements.txt` - Python package dependencies.
- `install` - PyInstaller build command.
- `main_backup.py` - older backup copy retained in the repository.

## Troubleshooting

If parashiyot appear wrong, verify that the app is using the Israel schedule. The parasha list and weekly data both use Hebcal's `i=on` parameter.

If `כניסת החג` is missing, confirm that the selected week includes an Erev Yom Tov item and that Hebcal candle-lighting data is available for the selected city.

If the app fails to start, confirm PyQt6 is installed in the Python environment used to run `main.py`.

If tests fail because PyQt6 is missing, `test_main.py` provides a lightweight import stub for GUI classes, but `openpyxl` is still required.
