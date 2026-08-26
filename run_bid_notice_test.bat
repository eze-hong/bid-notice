@echo off

cd /d C:\bid_notice

if not exist logs mkdir logs
if not exist output mkdir output

py "bid_notice_20260825.py" >> logs\run.log 2>&1

echo errorlevel=%errorlevel%
pause