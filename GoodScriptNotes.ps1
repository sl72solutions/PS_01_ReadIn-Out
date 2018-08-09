#$getpcname = Foreach-Object {
#(Get-WmiObject -Class Win32_ComputerSystem -Property Name).Name
#}  | Out-File "$env:userprofile\desktop\pcnames.txt"

#get-content -Path "..........csv" | Select-Object -property ................

=========================== parameters ==================================

[string][Parameter(mandatory = $False, Position = 2)] $PetesDoc = "$env:USERPROFILE\Desktop\PetesDoc-$(Get-Date -Format ddMMyyyy).csv",


=========================== import-csv ==================================

import-csv $..... -Header <names names names>


============================  get powershell & .NET versions ===========

$PSVersionTable.PSVersion

[environment]::Version

$PSVersionTable.CLRVersion

gci 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -recurse | gp -name Version,Release -EA 0 |
     where { $_.PSChildName -match '^(?!S)\p{L}'} | select PSChildName, Version, Release



================================================================



$SomeVariableName = import-Csv -Delimiter "," -Header <Header Name1>,<Header Name1>,<Header Name1>,<Header Name1>,  -Path $SomeVariableName 


Compare-Object -ReferenceObject $SomeVariableName -DifferenceObject $SomeVariableName  | select -ExpandProperty xxxxxx | export-csv $SomeVariableName -NoTypeInformation -Encoding UTF8

(Get-Content $SomeVariableName).replace('"', '') | Set-Content $SomeVariableName    #  alternatively - (Get-Content $SomeVariableName) | foreach {$_ -replace '"'} | Set-Content $SomeVariableName


Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property *user*

Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property RegisteredUser

$env:USERNAME

$NextHostName = $xxxxxxxxxx  | format-hex -Raw 


#=========  how many days back is the date?   ==========================

$daysback = 9

$NewDate = (Get-Date).AddDays(-$daysback).ToString("yyyy-MM-dd") 

#=========  network adapators    ==========================

Get-NetAdapter | select Name,MacAddress

Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled='True'" -ComputerName CACESV2NQE | 
Select-Object -Property Address


#=========  ComputerName   ==========================    plus validation on format

$ComputerName  = Read-Host "Enter target computer name" 

$ComputerName  = $ComputerName.ToUpper();

if ((!($ComputerName)) -or ($ComputerName.Length -ne 9) -or ($ComputerName -match '^[1-9]?$"')) {
write-host -ForegroundColor red "YourBad"
return -1
}

#=========  connect to database   ==========================

Invoke-Sqlcmd  -ServerInstance 'SL2WindowsVM' -Query "select count(*) from dbo.YourPCInfo" -Database JeanettesTestDB




$wshell = New-Object -com WScript.Shell
#$wshell.Run("iexplore.exe $url")
$wshell.Run("https://github.com/sl72solutions/PS_01_ReadIn-Out $url")
Start-Sleep 0

cd $env:USERPROFILE

$DownloadDir = Get-Location

$LSTTrackerDoc = Get-ChildItem -Path $DownloadDir -Filter 'PhilLynott *' | Sort-Object LastAccessTime -Descending | Select-Object -First 1


gci env:*

$SLsTrackingFile = Get-Item $LSTTrackerDoc

$SLsTrackingFileDir = (Get-Item .\).fullname

$shell = New-Object -ComObject "Shell.Application"
$shell.minimizeall()

$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $true

$ExcelWorkBookName = ($SLsTrackingFile.Name) # -replace ".xlsx", ".csv")
$Excel.AutomationSecurity.Equals(3)
$Workbook = $Excel.Workbooks.open($SLsTrackingFile.FullName)
#$AltSheet = $Excel.Workbooks.add()
#$NewSheet = $AltSheet.worksheets.add()
$Excel.DisplayFullScreen = 'True'

$date = '09/07/2018'

#$cellReference = “F3”
#$Excel.Cells  = “F3”

$Range = $_.Range("F3,F3")
$Range.paste($date)
$excel.Run("Sheet2.RefreshData")

$excel | Get-Member -Name 'Refresh Data'


# copy and paste to new sheet, save off locally

$rangeAc = $WorkSheet.Range(“B2:B26”)
$rangeAc.Copy() | out-null
#Select sheet 2
$Worksheet2 = $Workbook2.Worksheets.item(“Worklog”)
$worksheet2.activate()
#SO Range Paste
$rangeAp = $Worksheet2.Range(“K6:K30”)


$shell.minimizeall()
[System.Windows.MessageBox]::Show('Please be patient whilst we....','Please be patient','OK')

# Stopping a Program 

Get-Process -Name VIPUIManager | Foreach-Object { $_.CloseMainWindow() | Out-Null } | stop-process –force


      
      Add-Type -AssemblyName System.Windows.Forms
      $Prompt22 = [System.Windows.Forms.MessageBox]::Show('file not found, please check xxx folder', "..not downloaded", 4)
         # If ($Prompt22 -ne 'Yes')   {
                  Return -1
         # } 


$Age = Read-Host "Please enter your age"

$pwd_secure_string = Read-Host "Enter a Password" -AsSecureString

$securedValue = Read-Host -AsSecureString
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedValue)
$value = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)

