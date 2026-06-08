@echo off
echo Starting Portfolio Dashboard...
start "" http://localhost:8080/sample_portfolio_report.html
python -m http.server 8080
