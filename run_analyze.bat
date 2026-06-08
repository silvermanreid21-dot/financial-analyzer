@echo off
echo Running Portfolio Analyzer...
echo.
python analyze.py sample_portfolio.csv --cash 4900 --age 22
echo.
echo Copying to OneDrive dashboard...
copy /Y "sample_portfolio_report.html" "C:\Users\silve\OneDrive\Portfolio Dashboard\portfolio-dashboard.html"
echo.
echo Done! Refresh your browser to see updated data.
pause
