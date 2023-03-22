Import-Module .\module\CustomExcelModule.psd1 -Force

MakeGood -ExcelFile 'c:\tmp\example.xlsx' -CSVFiles @('c:\tmp\example.csv', 'C:\tmp\addresses.csv')
Exit-Appication