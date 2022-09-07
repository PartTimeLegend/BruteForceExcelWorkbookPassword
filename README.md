# BruteForceExcelWorkbookPassword

![Powershell](https://github.com/PartTimeLegend/BruteForceExcelWorkbookPassword/workflows/Powershell/badge.svg) [![Codacy Security Scan](https://github.com/PartTimeLegend/BruteForceExcelWorkbookPassword/actions/workflows/codacy-analysis.yml/badge.svg)](https://github.com/PartTimeLegend/BruteForceExcelWorkbookPassword/actions/workflows/codacy-analysis.yml)

Ever forgot your password for an Excel Workbook? Me neither! However some people do.

Put this script in the same dir as your workbook and run it.

You can pass in the workbook name as a link to a wordlist to use however one is defaulted to.

```powershell
./BruteForceExcelWorkbookPassword.ps1 "book1.xlsx" "https://raw.githubusercontent.com/openethereum/wordlist/master/res/wordlist.txt"
```

### Disclaimer
This may take a long time, and may also never actually find your password if it is not in the wordlist.

### Password for example book1.xslx
"password" at least I think it is. Try it and see.
