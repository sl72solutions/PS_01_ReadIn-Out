#requires -version 2
<#
.SYNOPSIS
  <Overview of script>
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  <Outputs if any, otherwise state None - example: Log file stored in C:\Windows\Temp\<name>.log>
.NOTES
  Version:        2.1
  Author:         <Name>
  Creation Date:  <Date>
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

[CmdletBinding()]
Param (

)

<#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Set Error Action to Silently Continue
#$ErrorActionPreference = "SilentlyContinue"
#Dot Source required Function Libraries
#"C:\Scripts\Functions\Logging_Functions.ps1"
#----------------------------------------------------------[Declarations]---------------------------------------------------------- #>

#Script Version
$sScriptVersion = "2.0"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

 Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [String]$Message,
        [String]$LogPath = "$env:userprofile\desktop\$(Get-Date -Format ddMMyyyy).log",
        [ValidateSet("Error", "Warning", "Information")]
        [String]$Type = "Information" 
    )
    Begin {
        $OldVerbosePreference = $VerbosePreference
        $VerbosePreference = 'Continue'
    }
    Process {
        if (!(Test-Path $LogPath)) {
            Write-Verbose "Creating $LogPath."
            New-Item $LogPath -Force -ItemType File -ErrorVariable x -ErrorAction SilentlyContinue
               
            if ($x) {
                Write-Verbose "Failed to create log file: $x"
            }          
        }
          
        $FormattedDate = Get-Date -Format "dd-MM-yyyy HH:mm:ss"

        switch ($Type) {
            'Error' {
                Write-Host "ERROR: $Message"
                $TypeText = 'ERROR: '
            }
            'Warning' {
                Write-Host "WARNING: $Message"
                $TypeText = 'WARNING:'
            }
            'Information' {
                Write-Host "INFO: $Message"
                $TypeText = 'INFO: '
            }
        }
        "[$FormattedDate][$CurrentToolVersion][$env:Username] $TypeText $Message" | Out-File -FilePath $LogPath -Append -ErrorAction SilentlyContinue
    }
    End {
        $VerbosePreference = $OldVerbosePreference
    }
}   

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$webRequest = Invoke-WebRequest -Uri https://github.com/sl72solutions/PS_01_ReadIn-Out/blob/master/laptop_detail.csv
#$webRequest = Invoke-WebRequest https://github.com/sl72solutions/PS_01_ReadIn-Out/blob/master/laptop_detail.csv
#[string][Parameter(mandatory = $False, Position = 2)] $PetesDoc = "$env:USERPROFILE\Desktop\PetesDoc-$(Get-Date -Format ddMMyyyy).csv", 
$sReadFile = ConvertFrom-StringData -StringData $webRequest.Content

  if (!$sReadFile)     {
                [System.Windows.Forms.MessageBox]::Show('Laptop detail file not found, please check folder','Is file there?','OK')
                Break   
                } 

 


Function GetFile_2_ExportXLS{
        Begin{
            Write-host "Lets Go.." -foreground Yellow
            $file = Import-Csv  $sReadFile 
        }
      
Process{

    Write-Log -Type Information -Message "Removing any previous versions of Output document"
    Get-ChildItem -Path $sReadFile | ForEach-Object{
        Try{
            remove-item $_
            }
                Catch{
                    Write-log –Message “No previous files found” –Type Information
                }
        }

        Try{
            foreach ($line in $file)
                { $line.NAME | out-host;    
                $line.CREATION_DATE | out-host;   
                $line.NI_ID | out-host;    
                $line.SITE_ID | out-host;   
                $line.DOB | out-host;    }
           }

         catch{
                $error = $_.Exception.Message
                Write-host $error -foreground Red
                Break
              }
          }                   
     
      End{
         If($?){
                   #CLS
                   write-host "Completed Successfully." -foreground Yellow
                   write-host " "
                   write-host " "
                   write-host " "
                   write-host " "
                   write-host "Log file in: " $sLogFile

                   $file | Export-Csv -append $sOutputFile #-NoTypeInformation
                
         }
      }
  }  
  
 
  
#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Log-Start $sLogPath $sLogName 

#Script Execution goes here

Write-Log
GetFile_2_ExportXLS

#Log-Finish -LogPath $sLogFile 