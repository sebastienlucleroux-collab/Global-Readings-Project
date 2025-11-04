@echo off
REM ---------------------------------------------
REM Interactive workflow for VS Code & Apps Script
REM ---------------------------------------------

REM 1. Go to your project folder (change path if needed)
cd "C:\Users\Seb le Roux\Global Readings Project"

REM 2. Pull latest changes from Apps Script
echo Pulling latest changes from Apps Script...
clasp pull

REM 3. Check if there are merge conflicts
git status > temp_status.txt
findstr /i "both modified" temp_status.txt >nul
if %errorlevel%==0 (
    echo Merge conflicts detected! Opening VS Code to resolve...
    code .
    echo Resolve conflicts in VS Code, then save and close the files.
    pause
) else (
    echo No merge conflicts detected.
)

del temp_status.txt

REM 4. Stage all changes
git add .

REM 5. Commit changes (asks for message)
set /p commitmsg="Enter commit message: "
git commit -m "%commitmsg%"

REM 6. Push changes back to Apps Script
echo Pushing changes to Apps Script...
clasp push

echo.
echo Done! Your VS Code edits and Apps Script edits are now synced.
pause