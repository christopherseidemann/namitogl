# namitogl

Turns a member export from NaMi, the membership database of the German scout
association DPSG, into one Excel list per age group for the group leaders.
I wrote it for my local scout group, where I manage the membership data.

## What it does

1. Reads the NaMi export (`.xls` as downloaded from NaMi, or `.xlsx`) and keeps active
   members and trial members (`Mitglied`, `Schnuppermitglied`).
2. Writes one file per age group (`Wölfling.xlsx`, `Jungpfadfinder.xlsx`,
   `Pfadfinder.xlsx`, `Rover.xlsx`). The lists only contain first and last name; for
   trial members also the date they joined. Leaders do not need addresses or bank
   details, so those are left out.
3. Marks trial members in red if they joined more than six weeks ago, so the leader
   knows to ask whether they want to become full members.
4. Sorts names the way a German phone book does (DIN 5007: Ä like A, case ignored).
5. Writes `Fehler.xlsx` for entries that need fixing in NaMi, such as a missing age group.

## Usage

On Windows, double-clicking `starten.bat` sets up a virtual environment on the first
run and opens a small Streamlit app in the browser. There you upload the export,
check the lists and download them as a ZIP file.

Without the app:

```bash
pip install -r requirements.txt
python namitogl.py nami_export.xls [output_folder]
```

The age groups and the six-week limit are set at the top of `namitogl.py`.
A step-by-step guide in German is in [Anleitung.md](Anleitung.md).

## Tests

```bash
pip install -r requirements-dev.txt
pytest
```

The tests only use invented members. Tested with Python 3.13.

## Data protection

Everything runs locally. The app only listens on `localhost`, and Streamlit's usage
statistics are switched off in `.streamlit/config.toml`. `.gitignore` excludes all
Excel files so that member data cannot be committed by accident.

## License

MIT, see [LICENSE](LICENSE).
