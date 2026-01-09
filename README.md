# Rolodex
A digital rolodex python app for processing images/PDFs of business cards. Useful on its own or for preparing a CSV to import contacts to another application (e.g. Outlook).

Executable: Located at \PyInstaller\dist\rolodex\rolodex_v1.0.exe

Future improvements:
- Eliminate warnings related to font size initialization
- Improve heuristic parsing of business card images/text
- Reduce launch time (currently ~10 to 20 seconds)
- Make sure the Editor window for adding new contacts is put on top of the main window in order. After using PyInstaller to create an executable the new window is showing up behind the main window.