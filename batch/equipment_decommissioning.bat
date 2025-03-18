@echo off
title Equipment decommissioning
echo Run Python script
echo Logged time = %time% %date%>> "D:\Leltar\log.txt"
start "C:\Python312\python.exe" "D:\Leltar\decommissioning.exe"