@echo off

Echo What is the name of the database? (Note: Parenthesis not allowed)
set /p DB_Name=""

Echo Please enter a database description:
set /p DB_Description=""

Echo Please enter the database server:
set /p DB_Server=""

Echo Please enter the database name exactly as it shows on the server:
set /p DB_Name2=""


Echo %WINDIR%\system32\odbcconf.exe CONFIGSYSDSN "SQL Server" "DSN=%DB_Name%|Description=%DB_Description%|SERVER=%DB_Server%|Trusted_Connection=Yes|Database=%DB_Name2%" > C:\Add_"%DB_Name%".bat

# Created 12.6.19 by Jake Ouellette #

