# CriarPlanilhaExcel
Script perl to create an Excel spreadsheet for each day of the year.

## Parameters
You can pass no parameters (script will create sheets from january to december) or you can pass one or both parameters below:
1. --MesInicial is numeric, between 1 and 12, meaning the first month to be created
2. --MesFinal is numeric, between 1 and 12, meaning the last month to be created

## Dependencies
- use Excel::Writer::XLSX;
- use Getopt::Long;
