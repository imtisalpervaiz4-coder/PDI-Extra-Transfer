PDI EXTRA TRANSFER - WINDOWS STANDALONE PACKAGE

GOAL
The final EXE can run on users' Windows computers without Python installed.
Only the one-time build process needs Python on a build machine, OR use the included GitHub Actions workflow to build online.

FILES
- PDI_Extra_Transfer_Assistant.py : dashboard source
- requirements.txt               : build dependencies
- build_windows.bat              : one-click local Windows build
- .github/workflows/build-windows-exe.yml : online Windows EXE build workflow

OPTION A - BUILD ON ONE WINDOWS COMPUTER
1. Install Python on the one computer used to create the EXE.
2. Put all package files in one folder.
3. Double-click build_windows.bat.
4. The finished EXE will be in the dist folder.
5. Share ONLY PDI_Extra_Transfer_Assistant.exe with other users.
   Recipients do not need Python or these source files.

OPTION B - BUILD ONLINE USING GITHUB ACTIONS
1. Create a GitHub repository and upload the contents of this package.
2. Open the repository's Actions tab.
3. Select 'Build Windows EXE' and click 'Run workflow'.
4. When it finishes, open that workflow run and download the artifact named
   'PDI-Extra-Transfer-Windows'.
5. Extract the downloaded artifact. It contains the EXE to share.

USAGE
1. Run the EXE on Windows.
2. Browse to the transfer workbook (.xlsx or .xlsm) and load it.
3. In PDI, click the first empty Number cell.
4. Press F9 to start/resume, F10 to pause, F11 to stop.
5. Review PDI before saving/posting.

IMPORTANT
- This automation uses desktop keyboard and clipboard control. Keep PDI visible and focused.
- Do not type or move to another application during transfer.
- Test with a small workbook first and verify the entered rows before processing a full transfer.
- The app's emergency PyAutoGUI failsafe is active: move the mouse to the top-left screen corner.
- Windows SmartScreen/antivirus may show a warning for a newly built unsigned EXE. Only run an EXE you built yourself or trust.
