<#

New-ModuleManifest `
-Path "C:\Users\pa1re\OneDrive\WindowsPowerShell\Modules\PoShHrdUtils\PoShHrdUtils.psd1" `
-RootModule "C:\Users\pa1re\OneDrive\WindowsPowerShell\Modules\PoShHrdUtils\PoShHrdUtils.psm1" `
-Author 'Reginald Baalbergen, The PA1REG' `
-CompanyName 'Radio Amateur' `
-Copyright '(c)2017 Reginald Baalbergen (PA1REG)' `
-Description 'Ham Radio Deluxe Utilities, Download and Update silent' `
-ModuleVersion 0.1 `
-PowerShellVersion 5.0 `
-FunctionsToExport 'Update-HamRadioDeluxe', 'Install-HamRadioDeluxe' `
-AliasesToExport 'UH', 'IH' `
-ProjectUri 'https://github.com/PA1REG/PoShHrdUtils' `
-HelpInfoUri 'https://github.com/PA1REG/PoShHrdUtils/blob/master/readme.md' `
-ReleaseNotes 'Initial Release.'
 
  
https://dfinke.github.io/2016/Quickly-Install-PowerShell-Modules-from-GitHub/
get-command -Module InstallModuleFromGitHub

Install-Module -Name InstallModuleFromGitHub -RequiredVersion 0.3
Install-ModuleFromGitHub -GitHubRepo /PA1REG/PoShHrdUtils


#>


function RemoveLineCarriage($lineIn)
# Remove all CR or LF, etc form string
{
  $ReturnValue = [System.String] $lineIn;
  $ReturnValue = $ReturnValue -replace "`t","";
  $ReturnValue = $ReturnValue -replace "`n","";
  $ReturnValue = $ReturnValue -replace "`r","";
  $ReturnValue = $ReturnValue -replace " ;",";";
  $ReturnValue = $ReturnValue -replace "; ",";";
  $ReturnValue = $ReturnValue -replace [Environment]::NewLine, "";
  return $ReturnValue;
}


function Get-AllowedChars ($lineIn) 
# An radio amateur callsign must contain only alfabet and a number with optional an /, this is checked
# 
{
  $ResultChars = $false
  $ResultNumbers = $false
  $ResultPattern = $true
  $ResultLength = $false
  $checkPattern = "^[/a-zA-Z0-9\s]+$" #'^[a-z0-9]+$')
  $checkPatternChars = "[a-zA-Z]"
  $checkPatternNumbers = "[0-9]"
  
  # Check if the string contains the given pattern
  if ($lineIn -notmatch $checkPattern) { $ResultPattern = $false }
  # there are no callsigns less than 3 positions
  if ($lineIn.Length -ge 3) { $ResultLength = $true }
  # Check if the callsign contains a letter and a number
  foreach ($Char In $lineIn)
  {
     if ($Char -match $checkPatternChars ) { $ResultChars = $true }
     if ($Char -match $checkPatternNumbers ) { $ResultNumbers = $true }
  }  
  # check the results and decide
  if($ResultChars -and $ResultNumbers -and $ResultPattern -and $ResultLength)
  {
    $ReturnValue = $true
  } else
  {
    $ReturnValue = $false
  }
  return $ReturnValue 
}


function Is-Call ($lineIn) 
# Check if it is really a callsign
{
  Try 
  {
    $ReturnValue = $true
    switch -wildcard ($lineIn) 
    { 
        "AF-*"  {$ReturnValue = $false} 
        "AS-*"  {$ReturnValue = $false} 
        "OC-*"  {$ReturnValue = $false} 
        "SA-*"  {$ReturnValue = $false} 
        "NA-*"  {$ReturnValue = $false} 
        "EU-*"  {$ReturnValue = $false} 
        "HF+*"  {$ReturnValue = $false} 
        "EU-*"  {$ReturnValue = $false} 
        "ETC.*" {$ReturnValue = $false} 
        ""      {$ReturnValue = $false} 
    }
    
    switch ($lineIn) 
    { 
        "0"         {$ReturnValue = $false} 
        "TEST"      {$ReturnValue = $false} 
        "ALL"       {$ReturnValue = $false} 
        "ASIAN"     {$ReturnValue = $false} 
        "CW"        {$ReturnValue = $false} 
        "ETC."      {$ReturnValue = $false} 
        "5W"        {$ReturnValue = $false} 
        "10W"       {$ReturnValue = $false} 
        "100W"      {$ReturnValue = $false} 
        "200W"      {$ReturnValue = $false} 
        "500W"      {$ReturnValue = $false} 
        "1000W"     {$ReturnValue = $false} 
        "1KW"       {$ReturnValue = $false} 
        "PSK31"     {$ReturnValue = $false} 
        "PSK63"     {$ReturnValue = $false} 
        "JT65"      {$ReturnValue = $false} 
        "JT9"      {$ReturnValue = $false} 
        "FT8"       {$ReturnValue = $false} 
        "160M"      {$ReturnValue = $false} 
        "80M"       {$ReturnValue = $false} 
        "60M"       {$ReturnValue = $false} 
        "40M"       {$ReturnValue = $false} 
        "30M"       {$ReturnValue = $false} 
        "24M"       {$ReturnValue = $false} 
        "20M"       {$ReturnValue = $false} 
        "17M"       {$ReturnValue = $false} 
        "15M"       {$ReturnValue = $false} 
        "12M"       {$ReturnValue = $false} 
        "10M"       {$ReturnValue = $false} 
        "6M"        {$ReturnValue = $false} 
        "0600Z"     {$ReturnValue = $false} 
        "24X7"      {$ReturnValue = $false} 
        "24HRS/DAY" {$ReturnValue = $false} 
    }
  
    $validCall = Get-AllowedChars($lineIn)
    if($validCall -and $ReturnValue)
    {
       $ReturnValue = $true
    } else
    {
       $ReturnValue = $false
    }

   } Catch 
  {
    $ReturnValue = $false
  }
  
  return $ReturnValue 
}
 

function Get-MonthNumber ($lineIn) 
# Return the number of a given month
{
   switch -wildcard ($lineIn) 
    { 
        "JAN" {$ReturnValue = "01"} 
        "FEB" {$ReturnValue = "02"} 
        "MAR" {$ReturnValue = "03"} 
        "APR" {$ReturnValue = "04"} 
        "MAY" {$ReturnValue = "05"} 
        "JUN" {$ReturnValue = "06"} 
        "JUL" {$ReturnValue = "07"} 
        "AUG" {$ReturnValue = "08"} 
        "SEP" {$ReturnValue = "09"} 
        "OCT" {$ReturnValue = "10"} 
        "NOV" {$ReturnValue = "11"} 
        "DEC" {$ReturnValue = "12"} 
        "MAA" {$ReturnValue = "03"} 
        "MEI" {$ReturnValue = "05"} 
        "OKT" {$ReturnValue = "10"} 
    }
 
  return $ReturnValue 
#Get-Culture
}


function Get-DateFromString ($lineIn) 
# Return the date from a string
{
# Get-Culture
  Try 
  {
    $Year  = $lineIn.substring(0,4)
    $Month = $lineIn.substring(5,3)
    $Day   = $lineIn.substring(8,2)
          
    $Month = Get-MonthNumber ($Month)
    $date = "$day/$month/$year"
    $EDate = [datetime]::parseexact($date, 'dd/MM/yyyy', $null)
    return $EDate
  } Catch 
  {
    return $null
  }
}


function Get-CallsCleanedUp
# Remove double entries and sort the output.
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [Object] $DxCallsArray
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      Write-Verbose "Sorting and remove duplicates"
      $DxCallsArray = $DxCallsArray | Sort-Object -Unique
      Return $DxCallsArray
      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
   }
}


function Add-ConcatDxCalls
# Add 2 array with calls together
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [Object] $DxCallsArray1,
        [Parameter(Position=1, Mandatory=$true)]
        [Object] $DxCallsArray2
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      Write-Verbose "Concatenate 2 array objects"
      $ConcatDxCallsArray = $DxCallsArray1 + $DxCallsArray2
      Return $ConcatDxCallsArray
      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
   }
}
 

function New-WriteAlarmFile
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [String] $AlarmFile,
        [Parameter(Position=1, Mandatory=$true)]
        [Object] $DxCallsArray
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      $arrCalls =@()

      if (! (Test-Path $AlarmFile)) 
      {
        Throw "Unable to locate file $AlarmFile from $URL ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "Empty file $AlarmFile)"
        Clear-Content $AlarmFile 
      }

      Try 
      {
        foreach ($Call in $DxCallsArray) 
        {
          $objResults = New-Object PSObject
          $objResults | Add-Member -membertype NoteProperty -name "Call" -Value $Call 
          $objResults | Add-Member -membertype NoteProperty -name "Settings" -Value "Y1101111111000001101111111000001101011111000001"
          $arrCalls += $objResults 
        }
  
        $CsvContent = $arrCalls | ConvertTo-Csv -NoTypeInformation
        Write-Verbose "Write all calls to $AlarmFile)"
        $CsvContent[1..$CsvContent.Length] | Add-Content $AlarmFile
      } Catch 
      {
        Write-Verbose "Unable to write all calls to $AlarmFile)"
      }

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}
 

function New-WriteHRDAlarmFile
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $DXClusterAlarmsFile,
        [Parameter(Position=1, Mandatory=$false)]
        [Object] $DxCallsArray
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

      if (! (Test-Path $DXClusterAlarmsFile)) 
      {
        Throw "Unable to locate file $DXClusterAlarmsFile from $URL ($ErrorMessage, $FailedItem)" 
      }

      Try 
      {
        # Building the xml string for calls
        foreach ($Call in $DxCallsArray) 
        {
          $DXClusterCalls += "$($Call)|"
        }
        # Delete the last character which contains a |-sign
        $DXClusterCalls = $DXClusterCalls.Substring(0,$DXClusterCalls.Length-1)
        Write-Verbose "Builded line is now : $$DXClusterCalls"
        # Naming elements
        $Time = (Get-Date -f g)
        $DXTitle = 'Special Event Callsigns'
        $DXComment = "Created by PS_DxExpeditions.ps1, PA1REG, last update : $Time"

        Write-Verbose "Reading file $DXClusterAlarmsFile"
        $XMLContent = [xml](Get-Content $DXClusterAlarmsFile)
        $Alarms = $XMLContent.SelectNodes("/HRD/Alarm")
        $SpecialAlarm = $Alarms | Where-Object {$_.Title -eq $DXTitle}
        if ($SpecialAlarm.Count -ne 0)
        {
          Write-Verbose "Updating XML file"
          $SpecialAlarm.Callsign = $DXClusterCalls
          $SpecialAlarm.SetAttribute('Comment',$DXComment)
        } else
        {
          Write-Verbose "Adding XML file"
          $NewAlarm = $XMLContent.CreateElement("Alarm")
          $NewAlarm.InnerText = ''
          $NewAlarm.SetAttribute('Callsign',$DXClusterCalls)
          $NewAlarm.SetAttribute('Comment',$DXComment)
          $NewAlarm.SetAttribute('Filter','ALL')
          $NewAlarm.SetAttribute('Title',$DXTitle)
          $NewAlarm.SetAttribute('Enable','1')
          $NewAlarm.SetAttribute('Interval','5')
          $NewAlarm.SetAttribute('Options','16')
          $Response = $XMLContent.DocumentElement.AppendChild($NewAlarm)
        }
      
        Write-Verbose "Write all calls to $DXClusterAlarmsFile"
        $XMLContent.Save($DXClusterAlarmsFile)
      } Catch 
      {
        Write-Verbose "Unable to write all calls to $DXClusterAlarmsFile)"
      }

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Stop-Application
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [String] $ProcessName,
        [Parameter(Position=1, Mandatory=$true)]
        [String] $Activity
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

       $Delaytime   = 1
       $TimeElaped  = 0
       $TimeOut     = 120
	   $Id          = 1


   	   Write-Progress -Id $Id -Activity $Activity  -PercentComplete ($TimeElaped / $TimeOut * 100)
       While ( (Get-Process -Name $($ProcessName) -ErrorAction SilentlyContinue))
       {
     	 Write-Progress -Id $Id -Activity $Activity  -PercentComplete ($TimeElaped / $TimeOut * 100)
         Stop-Process -Name $ProcessName
         Start-Sleep -Seconds $Delaytime
         $TimeElaped =+ $TimeElaped + $Delaytime
         if ($TimeElaped -ge $TimeOut)
         {
     	    Write-Verbose "Timeout : Stopping process takes too long ( > $TimeOut sec.)" -ForegroundColor Red
            Throw "Timeout : Starting process takes too long ( > $TimeOut sec.)" 
         }
       }

       if ($TimeElaped -lt $TimeOut)
       {
	      #Write-Host "---------" -ForegroundColor Red
       } else
       {
   	     Write-Host "Stopping process due timeout" -ForegroundColor Red
       }
  
        Write-Verbose "Eind Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}
  

function Start-Application
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [String] $ProcessName,
        [Parameter(Position=1, Mandatory=$true)]
        [String] $ProcessToStart,
        [Parameter(Position=2, Mandatory=$true)]
        [String] $Activity
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

       $Delaytime   = 1
       $TimeElaped  = 0
       $TimeOut     = 120
	   $Id          = 1


   	   Write-Progress -Id $Id -Activity $Activity  -PercentComplete ($TimeElaped / $TimeOut * 100)
       While ( -not (Get-Process -Name $($ProcessName) -ErrorAction SilentlyContinue))
       {
     	 Write-Progress -Id $Id -Activity $Activity  -PercentComplete ($TimeElaped / $TimeOut * 100)
         Start-Process $strExe 
         Start-Sleep -Seconds $Delaytime
         $TimeElaped =+ $TimeElaped + $Delaytime
         if ($TimeElaped -ge $TimeOut)
         {
     	    Write-Verbose "Timeout : Starting process takes too long ( > $TimeOut sec.)"
            Throw "Timeout : Starting process takes too long ( > $TimeOut sec.)" 
         }
       }

       if ($TimeElaped -lt $TimeOut)
       {
	      #Write-Host "---------" -ForegroundColor Red
       } else
       {
   	     Write-Host "Start process due timeout" -ForegroundColor Red
       }
  
        Write-Verbose "Eind Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Write-DxCallsToCSV
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [Object] $objDxCalls,
        [Parameter(Position=1, Mandatory=$true)]
        [String] $CSVFile
    )
    begin
    {
      Write-Verbose "Start Function : [$($MyInvocation.MyCommand)] *************************************"

      Try 
      {
        Write-Verbose "Start Writing CSV file $CSVFile" 
        $stuff = @()
        foreach($row in $objDxCalls) 
        {
         $obj = new-object PSObject
         $obj | add-member -membertype NoteProperty -name "DxCall" -value $row
         $stuff += $obj
        }
        $stuff | Export-Csv -Path $CSVFile -notypeinformation

      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to write file $CSVFile ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Get-ContentNg3K
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [Uri] $URL
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      $arrDxCalls =@()

  
      Try 
      {
        Write-Verbose "Start Reading from $URL" 
        $result = Invoke-WebRequest $URL
      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Verbose "Unable to open website $URL ($ErrorMessage, $FailedItem)"
        Throw "Unable to open website $URL ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Reading content from site $URL)"

      Try 
      {
        $elementStartdate = 0
        $elementEnddate = 1
        $elementDXCCentity = 2
        $elementCall = 3
        $elementQSLvia = 4
        $elementReportedBy= 5
        $elementInfo= 6 

        $StartDate = Get-DateFromString($EventDate)
        $CurrentDate = Get-Date 

        $result.ParsedHtml.getElementsByTagName('tr') | ForEach-Object {
          $tr =  $_;
          $startdate = Get-DateFromString ($tr.childNodes.item($elementStartdate).outerText)
          if ($startdate -ne $null)
          {
            $EndDate = Get-DateFromString ($tr.childNodes.item($elementEnddate).outerText)
            $DXCCentity = $tr.childNodes.item($elementDXCCentity).outerText
            $Call = $tr.childNodes.item($elementCall).outerText
            $QSLvia = $tr.childNodes.item($elementQSLvia).outerText
            $ReportedBy = $tr.childNodes.item($elementReportedBy).outerText
            $Info = $tr.childNodes.item($elementInfo).outerText

            if ( ($EndDate.Month -eq $CurrentDate.Month -or $StartDate.Month -le $CurrentDate.Month) -and ($EndDate.Year -eq $CurrentDate.Year) )
            {
              $call = RemoveLineCarriage($call)
              $startBracket = $Call.IndexOf("[")
              if ($startBracket -gt 0)
              {
                $call = $Call.substring(0,$startBracket).ToUpper()
                If (Is-Call ($call)) 
                {
                  $arrDxCalls += $call
                 #Write-host "YES $Remark " -ForegroundColor Green
                } 
              }
              $Remarks = RemoveLineCarriage($Remarks)
              $Remarks = $Info.Split(' ')
              foreach($Remark In $Remarks)
              { 
                $Remark = $Remark.Replace(")","")
                $Remark = $Remark.Replace("(","")
                $Remark = $Remark.Replace(",","")
                $Remark = $Remark.Replace(";","").Trim().ToUpper()
                If (Is-Call ($Remark)) 
                {
                  $arrDxCalls += $Remark
                  #Write-host "YES $Remark " -ForegroundColor Green
                } 
              }

            #Write-host "YES : 0=$startdate 1=$EndDate  3=$Call "
            # Write-host "YES : 0=$startdate 1=$EndDate 2=$time2 3=$time3 4=$time4 5=$time5 6=$time6 7=$time7 8=$time8"
          } else 
          {
            #Write-host "NO  : 0=$startdate 1=$EndDate 2=$time2 3=$time3 4=$time4 5=$time5 6=$time6 7=$time7 8=$time8" -ForegroundColor Gray
          }
        }
      }
      Return $arrDxCalls | Sort-Object -Unique
      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Error reading website $URL ($ErrorMessage, $FailedItem)" 
      }
      
      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"

    }
}



function Get-Calendar
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [Uri] $URL
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      $arrDxCalls =@()

  
      Try 
      {
        Write-Verbose "Start Reading from $URL" 
        $result = Invoke-WebRequest $url
      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to open website $URL ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Reading content from site $URL)"
      Try 
      {
        $CalendarFileContent = $result.Content -split "`r`n" | Where-Object { $_ -like "SUMMARY*"}

        foreach ($Line in $CalendarFileContent) 
        {
          $Line = $Line.ToUpper().Trim() 
          If ($Line.StartsWith("SUMMARY")) 
          {
            $LineSplit = $Line.Split(':')
            $DxCalls = $LineSplit[1].Trim().Split(' ')
            #Write-Host "$LineSplit  $DxCallsLine  $DxCalls" -ForegroundColor Yellow

            foreach($DxCall in $DxCalls)
            { 
              $DxCall = $DxCall.Replace(")","")
              $DxCall = $DxCall.Replace("(","").Trim()
              If (Is-Call ($DxCall)) 
              {
                $arrDxCalls += $DxCall
              } 
            }
        }
       }
      Return $arrDxCalls | Sort-Object -Unique
      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to read website $URL ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
   }
}


$AllDxCalls = @()

$DisplayColor           = 'Green'
$LogFile                = "$env:USERPROFILE\downloads\PS_DxExpeditions.csv"
# https://dx-world.net/
[uri]$CalendarURL       = 'https://www.google.com/calendar/ical/hamradioweb2007%40gmail.com/public/basic.ics'
# Announced DX Operations
[uri]$Ng3kURL           = "http://www.ng3k.com/misc/adxo.html"
#$vCalendarFile          = "$env:USERPROFILE\downloads\DXpeditions.acs"
#$Ng3KWebContentFile     = "$env:USERPROFILE\downloads\Ng3KWebContent.txt"
$ve7ccAlarmFile         = "$env:USERPROFILE\downloads\alarm.dat"
$ve7ccExe               = 'D:\ve7cc\ve7cc.exe'
$HRDDXClusterAlarmsFile = "$env:APPDATA\HRDLLC\HRD Logbook\DXClusterAlarms.xml"

<#
$strComputer = $env:computername
$strProcessVE7CC = Get-Process -Name "ve7cc" -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
If ($strProcessVE7CC.Count -ne 0) 
{
  $ve7ccAlarmFile = "D:\ve7cc\alarm.dat"
  $ve7ccProcessName = 've7cc'
  $strExe = $ve7ccExe
} Else 
{
  $ve7ccAlarmFile = "$env:USERPROFILE\downloads\alarm.dat"
  $ve7ccProcessName = 'notepad'
  $strExe = 'notepad'
}

if (! (Test-Path $HRDDXClusterAlarmsFile)) 
{
  $HRDDXClusterAlarmsFile = "$env:USERPROFILE\downloads\DXClusterAlarms.xml"
}


#Write-Host "VE7CC Alarm File : $ve7ccAlarmFile" -ForegroundColor $DisplayColor
#Write-Host "Ham Radio DeLuxe Alarm File : $HRDDXClusterAlarmsFile" -ForegroundColor $DisplayColor



Write-Host "Write VE7CC Alarm File : $ve7ccAlarmFile" -ForegroundColor $DisplayColor
New-WriteAlarmFile -AlarmFile $ve7ccAlarmFile -DxCallsArray $AllDxCalls

Write-Host "Write Ham Radio DeLuxe Alarm File : $HRDDXClusterAlarmsFile" -ForegroundColor $DisplayColor
New-WriteHRDAlarmFile -DXClusterAlarmsFile $HRDDXClusterAlarmsFile -DxCallsArray $AllDxCalls

$ve7ccActivity = "VE7CC application, stopping and starting"
Write-Host "$ve7ccActivity" -ForegroundColor $DisplayColor
Stop-Application -ProcessName $ve7ccProcessName -Activity $ve7ccActivity
Start-Application -ProcessName $ve7ccProcessName -ProcessToStart $strExe -Activity $ve7ccActivity
#>

$VerbosePreference = "SilentlyContinue"


function Get-CallsFromWebsites
{
    begin
    {
      Write-Verbose "Start Function : [$($MyInvocation.MyCommand)] *************************************"
      
      $CalendarCalls = @()
      $Ng3KCalls = @()
      $objDxCalls= @()
      $DisplayColor           = 'Green'
      # https://dx-world.net/
      [uri]$CalendarURL       = 'https://www.google.com/calendar/ical/hamradioweb2007%40gmail.com/public/basic.ics'
      # Announced DX Operations
      [uri]$Ng3kURL           = "http://www.ng3k.com/misc/adxo.html"

      Write-Host "Reading website : $CalendarURL" -ForegroundColor $DisplayColor
      $CalendarCalls = Get-Calendar -URL $CalendarURL
      Write-Verbose "Calls Readed : $CalendarCalls"
      Write-Host "Calls Loaded : $($CalendarCalls.Count)" -ForegroundColor $DisplayColor
      #Write-DxCallsToCSV -objDxCalls $CalendarCalls -CSVFile $env:USERPROFILE\downloads\Calls_Agenda.csv

      Write-Host "Reading website : $Ng3kURL" -ForegroundColor $DisplayColor
      $Ng3KCalls = Get-ContentNg3K -URL $Ng3kURL
      Write-Verbose "Calls Readed : $Ng3KCalls"
      Write-Host "Calls Loaded : $($Ng3KCalls.Count)" -ForegroundColor $DisplayColor
      #Write-DxCallsToCSV -objDxCalls $Ng3KCalls -CSVFile $env:USERPROFILE\downloads\Calls_Ng3K.csv

      $objDxCalls = Add-ConcatDxCalls -DxCallsArray1 $CalendarCalls -DxCallsArray2 $Ng3KCalls
      $NumberInputCalls = $($objDxCalls.Count)
      $objDxCalls = Get-CallsCleanedUp -DxCallsArray $objDxCalls
      Write-Host "Remove duplicate callsigns, total $NumberInputCalls, returned $($objDxCalls.Count)" -ForegroundColor $DisplayColor
      return $objDxCalls
      
      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function New-HrdAlarmFile
{
    <#
        .SYNOPSIS
            Function is to create a Dx Cluster Alarm fiel for use In Ham Radio DeLuxe
        .DESCRIPTION
            This function downloads callsign from DX Expeditions from websites and place It In the Alarm file from Ham Radio DeLuxe
            If there is no entry In the "DX Cluster Alarm Defenitions" It creates one, existing will be updated.
        .PARAMETER $AlarmFile
            Optional : Specify this parameter if the DXClusterAlarms file is on an other location.
            Notes: 
                * Default path =  C:\Users\<user>\AppData\Roaming\HRDLLC\HRD Logbook\DXClusterAlarms.xml
        .EXAMPLE
            PS> New-HrdAlarmFile -AlarmFile ""

            This example creates the alarmfile on default location.
        .EXAMPLE
            PS> New-HrdAlarmFile -DownloadPath "$env:USERPROFILE\downloads\test.txt"

            This example creates a alarmfile "test.txt" on location "C:\Users\<user>\downloads\"
        .EXAMPLE
            PS> Install-HamRadioDeluxe -Verbose

            This example gives extra information about the internal steps.
        .EXAMPLE
        .INPUTS
        .OUTPUTS
            $null
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $AlarmFile
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
 
      if ($AlarmFile)
      {
        $selectedAlarmFile = $AlarmFile
      } else
      {
        $selectedAlarmFile = "$env:APPDATA\HRDLLC\HRD Logbook\DXClusterAlarms.xml"
      }
      
      if (! (Test-Path $selectedAlarmFile)) 
      {
        Throw "Unable to locate file $selectedAlarmFile ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "File found : $selectedAlarmFile"
      }
      
      $AllDxCalls = Get-CallsFromWebsites    
      
      Write-Host "Write Ham Radio DeLuxe Alarm File : $selectedAlarmFile" -ForegroundColor $DisplayColor
      New-WriteHRDAlarmFile -DXClusterAlarmsFile $selectedAlarmFile -DxCallsArray $AllDxCalls

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function New-Ve7ccAlarmFile
{
    <#
        .SYNOPSIS
            Function is to create a Dx Cluster Alarm fiel for use In Ham Radio DeLuxe
        .DESCRIPTION
            This function downloads callsign from DX Expeditions from websites and place It In the Alarm file from Ham Radio DeLuxe
            If there is no entry In the "DX Cluster Alarm Defenitions" It creates one, existing will be updated.
        .PARAMETER $AlarmFile
            Optional : Specify this parameter if the DXClusterAlarms file is on an other location.
            Notes: 
                * Default path =  C:\Users\<user>\AppData\Roaming\HRDLLC\HRD Logbook\DXClusterAlarms.xml
        .EXAMPLE
            PS> New-HrdAlarmFile -AlarmFile ""

            This example creates the alarmfile on default location.
        .EXAMPLE
            PS> New-HrdAlarmFile -DownloadPath "$env:USERPROFILE\downloads\test.txt"

            This example creates a alarmfile "test.txt" on location "C:\Users\<user>\downloads\"
        .EXAMPLE
            PS> Install-HamRadioDeluxe -Verbose

            This example gives extra information about the internal steps.
        .EXAMPLE
        .INPUTS
        .OUTPUTS
            $null
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [String] $AlarmFile
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
 
      if ($AlarmFile)
      {
        $selectedAlarmFile = $AlarmFile
      } else
      {
        $selectedAlarmFile = "$env:APPDATA\HRDLLC\HRD Logbook\DXClusterAlarms.xml"
      }
      
      if (! (Test-Path $selectedAlarmFile)) 
      {
        Throw "Unable to locate file $selectedAlarmFile ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "File found : $selectedAlarmFile"
      }
      
      $AllDxCalls = Get-CallsFromWebsites    
      
      Write-Host "Write Ham Radio DeLuxe Alarm File : $selectedAlarmFile" -ForegroundColor $DisplayColor
      New-WriteHRDAlarmFile -DXClusterAlarmsFile $selectedAlarmFile -DxCallsArray $AllDxCalls

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


New-Alias -Name UH -Value New-HrdAlarmFile
Export-ModuleMember -function New-HrdAlarmFile -alias NHRD

 
