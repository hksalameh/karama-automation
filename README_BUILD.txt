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

Current monthly workflow:
1. Open KafalaCompareApp.exe.
2. Choose the care-program Excel file.
3. Choose the exported Karama-site Excel file.
4. Run comparison and review the generated result.
5. Open the automatic-update screen.
6. Enter the Karama username and password for the current session.
7. Select year, month, and category: ايتام / اسر / طلاب علم.
8. Confirm start.
9. The program opens Karama, logs in, opens شاشة الصرفية, selects the requested year/month/category, and clicks عرض.
10. Records are processed strictly one at a time. The next record is not processed until the current record has been verified and changed successfully.
11. The script verifies national ID, site name, and current amount before changing the amount.
12. If any critical mismatch occurs, processing stops and temporary save is NOT clicked.
13. If all records finish successfully, the program clicks حفظ مؤقت exactly once.
14. The program confirms the message تم حفظ القيم بنجاح.
15. The browser stays open for the user to review the total.
16. حفظ نهائي is always manual and must never be clicked by the automation.

Business rules:
- Program الرعاية is the authoritative monthly balance source.
- If a person exists in Karama but not in the care-program file, the expected amount is 0.
- If a person exists in the care-program file but not in Karama, the comparison report marks the person as needing to be added to Karama; automatic amount editing does not add new people.
- The three categories are processed separately: ايتام, اسر, طلاب علم.

Safety behavior:
- Duplicate national IDs in the automatic-update file stop the operation.
- The script verifies the old site amount before replacing it.
- A record that already contains the target amount is treated as already correct.
- Any critical error stops the run before temporary save.
- Temporary save is performed only after all requested records are processed successfully.
- The code explicitly excludes any button containing حفظ نهائي from automatic clicking.
- The user's password is passed only to the running browser process for the current session and is not written to the app's report files.
