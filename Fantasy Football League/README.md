# Fantasy Premier League tools

## Function

Retrieve player statistics, gameweek points and fixture difficulties, and organise them into a formatted Excel worksheet. Optionally include additional FotMob statistics.

## Overview

A Python tool for creating and refreshing an Excel workbook of Fantasy Premier League data, bringing player and fixture information together for comparison.

## Contents

Get_FPL_Player_Statistics - Create or refresh an Excel worksheet of Fantasy Premier League player statistics.

## Context and contribution

I developed this tool as a personal project to automate the collection and organisation of football statistics.

Developed with assistance from ChatGPT. I defined the requirements, adapted and tested the code, checked outputs for errors, and debugged issues.

## How to use

Open the Get_FPL_Player_Statistics folder and follow its README. Install requests and openpyxl, set FILE_PATH and SEASON, and configure ENABLE_FOTMOB as required. Close the workbook in Excel before running the script with an internet connection. Each update replaces the selected season worksheet while retaining the other tabs. Data retrieval depends on the external services used by the script.
