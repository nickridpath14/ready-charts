@echo off
echo Building the project...
call npm run build

echo Navigating to build folder...
cd dist

echo Creating CNAME file for custom domain...
echo readycharts.nickridpath.com > CNAME

echo Initializing temporary git repo...
git init
git checkout -b gh-pages

echo Adding files...
git add -A
git commit -m "Deploying to GitHub Pages"

echo Pushing to GitHub...
REM --- IMPORTANT: This pushes to the gh-pages branch ---
git push -f https://github.com/nickridpath14/ready-charts.git gh-pages

echo Cleaning up...
cd ..
rmdir /s /q dist\.git

echo ---------------------------------------------------------
echo Deployment Complete! 
echo Your site should be live in a few minutes at:
echo https://readycharts.nickridpath.com/
echo ---------------------------------------------------------
pause