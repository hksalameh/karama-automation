KafalaCompareApp - build notes

Final files:
- dist/KafalaCompareApp.exe: portable one-file program.
- Output/KafalaCompareApp_setup.exe: installer version.
- KafalaCompareApp_build.zip: zip that contains the portable EXE.

Runtime requirements:
- No Python or Node.js installation is needed on the user's computer.
- Python, node.exe, and the required Node modules are bundled into the EXE.
- Automatic browser updates use the installed system Chrome or Edge browser.
- Runtime logs and reports are written to:
  %LOCALAPPDATA%\KafalaCompareApp

Build requirements on the developer machine:
- Python 3 with PyInstaller and the Python dependencies used by the app.
- KafalaCompareApp_build/node.exe
- KafalaCompareApp_build/node_modules
- Inno Setup 6, only if you want to create Output/KafalaCompareApp_setup.exe.

Recommended end-of-month test:
1. Open KafalaCompareApp.exe.
2. Choose the care-program Excel file.
3. Choose the exported site Excel file.
4. Run comparison.
5. Review the generated result.
6. Keep manual confirmation enabled for the first real batch.
7. Enable temporary save every 15 edits only when wanted.
8. Enable amount update to 0 only after reviewing zeroing cases.
9. Run automatic update.
10. Log in to the site and open the disbursement screen.
11. Continue from the app's automatic-update screen.

Safety behavior:
- The script verifies national ID and name before changing an amount.
- The script does not click the final save button.
- Temporary save is disabled by default.
- Updates to 0 run only when the zero-update option is enabled.
