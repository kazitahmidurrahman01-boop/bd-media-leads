@echo off
echo ========================================================
echo   Pushing 100 Media Leads Dashboard to GitHub Pages
echo ========================================================
echo.
set /p REPO_NAME="Enter your GitHub repo name (default: bd-media-leads): "
if "%REPO_NAME%"=="" set REPO_NAME=bd-media-leads

echo Target Repo: https://github.com/kazitahmidurrahman01-boop/%REPO_NAME%.git
git remote remove origin >nul 2>&1
git remote add origin https://github.com/kazitahmidurrahman01-boop/%REPO_NAME%.git
git branch -M main
git add .
git commit -m "Deploy 100 Bangladesh Media Leads and Real-Time Team Sync Dashboard" >nul 2>&1

echo Pushing to GitHub...
git push -u origin main

echo.
echo ========================================================
echo Done! Now go to GitHub:
echo https://github.com/kazitahmidurrahman01-boop/%REPO_NAME%/settings/pages
echo Set Branch: 'main', Folder: '/ (root)' and click Save.
echo Your universal link will be:
echo https://kazitahmidurrahman01-boop.github.io/%REPO_NAME%/
echo ========================================================
pause
