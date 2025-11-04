@echo off
cd /d "%~dp0"

echo Adding all changes...
git add -A

echo Committing changes...
git commit -m "Update"

echo Pushing to GitHub...
git push origin main

echo Pushing to Google Apps Script...
clasp push

echo Done!
pause