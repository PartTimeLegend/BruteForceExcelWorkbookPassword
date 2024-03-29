<#
.SYNOPSIS
  Brute Force an Excel Workbook that is password protected
.DESCRIPTION
  My mother forgot the password for an Excel Spreadsheet. This is how I fixed that.

  This may take a long time, and may also never actually find your password if it is not in the wordlist.

  This PowerShell script is released under the MIT license http://www.opensource.org/licenses/MIT
  .PARAMETER $workbookname the filename of the workbook. This is relative to the directory the script is running in.
  .PARAMETER $worklistURL a URL to a wordlist as a text file. This has a default as a suggested wordlist, bigger and arguably better exist.
.INPUTS
  None
.OUTPUTS
  You should be told when the password has been found.
.NOTES
  Version:        1.0
  Author:         Antony Bailey <hi@antonybailey.net>
  Creation Date:  2021/02/04
  Purpose/Change: Initial script development

.EXAMPLE
  ./BruteForceExcelWorkbookPassword.ps1 "book1.xlsx" "https://raw.githubusercontent.com/openethereum/wordlist/master/res/wordlist.txt"
#>
param (
  [string] $workbookname,
  $wordlistUrl = 'https://raw.githubusercontent.com/openethereum/wordlist/master/res/wordlist.txt'
  )
$FilePath = Get-Location
$fullFilePath = Join-Path $FilePath $workbookname
Invoke-WebRequest $wordlistUrl -OutFile .\wordlist.txt
foreach($password in Get-Content .\wordlist.txt)
{
  try
  {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Output "Attempting to open $fullFilePath with $password"
    $excel.Workbooks.Open($FilePath, [Type]::Missing, [Type]::Missing, [Type]::Missing, $password)
    Write-Output "The password for $fullFilePath is $password"
    Exit 0
  }
  catch
  {
    Write-Output "It wasn't $password"
  }
  finally
  {
    try
    {
      $excel.Quit()
    }
    catch
    {
      # Empty in case Excel fails.
    }
  }
}
