# Get_FPL_Player_Statistics.py

## Function

Retrieves player totals, gameweek points and fixture difficulties; sums double-gameweek points and formats the workbook. Optionally adds FotMob statistics. Refreshes the selected season tab while retaining other tabs.

## Overview

Create or refresh an Excel worksheet of Fantasy Premier League player statistics.

## Context and contribution

I wrote this script for a personal project to automate a practical task.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Install requests and openpyxl. Set FILE_PATH and SEASON to match the live FPL season; set ENABLE_FOTMOB as desired. Close the workbook in Excel and run `python Get_FPL_Player_Statistics.py` with an internet connection. The selected season sheet is replaced on each update; use a hyphen rather than a slash in SEASON.
