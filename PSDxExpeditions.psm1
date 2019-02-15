$DisplayColor           = 'Green'
$AllDxCalls = @()


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


function Get-CheckAllowedCharsInLine ($lineIn) 
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


function Get-ValidCall ($lineIn) 
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
        "JT9"       {$ReturnValue = $false} 
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
        "160/80M"   {$ReturnValue = $false} 
        "2300M"     {$ReturnValue = $false} 
        "50W"       {$ReturnValue = $false} 
        "100W"      {$ReturnValue = $false} 
        "600W"      {$ReturnValue = $false} 
        "0600Z"     {$ReturnValue = $false} 
        "24X7"      {$ReturnValue = $false} 
        "24HRS/DAY" {$ReturnValue = $false} 
    }
  
    $validCall = Get-CheckAllowedCharsInLine($lineIn)
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
      Write-Verbose "End Module : [$($MyInvocation.MyCommand)] *************************************"
   }
}


function Add-ConcatDxCall
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
      Write-Verbose "End Module : [$($MyInvocation.MyCommand)] *************************************"
   }
}
 

function Update-WriteVE7CCAlarmFile
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
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
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
 

function Update-WriteHRDAlarmFile
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
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
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
        Write-Verbose "Builded line is now : $DXClusterCalls"
        # Naming elements
        $Time = (Get-Date -f g)
        $DXTitle = 'Special Event Callsigns'
        $DXComment = "[$Time] Created by PS_DxExpeditions.ps1 by PA1REG"
        Write-Verbose "Comment is now : $DXComment"

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
          Write-Verbose "Adding XML response $Response"
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
        [String] $ProcessActivity
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

       $Delaytime   = 1
       $TimeElaped  = 0
       $TimeOut     = 120
	   $Id          = 1


   	   Write-Progress -Id $Id -Activity $ProcessActivity  -PercentComplete ($TimeElaped / $TimeOut * 100)
       While ( (Get-Process -Name $($ProcessName) -ErrorAction SilentlyContinue))
       {
     	 Write-Progress -Id $Id -Activity $ProcessActivity  -PercentComplete ($TimeElaped / $TimeOut * 100)
         Stop-Process -Name $ProcessName
         Start-Sleep -Seconds $Delaytime
         $TimeElaped =+ $TimeElaped + $Delaytime
         if ($TimeElaped -ge $TimeOut)
         {
     	    Write-Verbose "Timeout : Stopping process takes too long ( > $TimeOut sec.)"
            Throw "Timeout : Starting process takes too long ( > $TimeOut sec.)" 
         }
       }

       if ($TimeElaped -lt $TimeOut)
       {
	      #Write-Host "---------" -ForegroundColor Red
       } else
       {
   	     Write-Host "Stopping process due timeout" -ForegroundColor $DisplayColor -ForegroundColor $DisplayColor
       }
  
        Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
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
        [String] $ProcessActivity
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

       $Delaytime   = 1
       $TimeElaped  = 0
       $TimeOut     = 120
	   $Id          = 1


   	   Write-Progress -Id $Id -Activity $ProcessActivity  -PercentComplete ($TimeElaped / $TimeOut * 100)
       While ( -not (Get-Process -Name $($ProcessName) -ErrorAction SilentlyContinue))
       {
   	     Write-Verbose "Start process $($ProcessName)"
     	 Write-Progress -Id $Id -Activity $ProcessActivity  -PercentComplete ($TimeElaped / $TimeOut * 100)
         Start-Process $ProcessToStart 
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
	      #Write-Output "---------" -ForegroundColor Red
       } else
       {
   	     Write-Output "Start process stopped due timeout"
       }
  
        Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
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
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"

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

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
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
        #$elementDXCCentity = 2
        $elementCall = 3
        #$elementQSLvia = 4
        #$elementReportedBy= 5
        $elementInfo= 6 

        $StartDate = Get-DateFromString($EventDate)
        $CurrentDate = Get-Date 

        $result.ParsedHtml.getElementsByTagName('tr') | ForEach-Object {
          $tr =  $_;
          $StartDate = Get-DateFromString ($tr.childNodes.item($elementStartdate).outerText)
          if ($null -ne $startdate)
          {
            $EndDate = Get-DateFromString ($tr.childNodes.item($elementEnddate).outerText)
            #$DXCCentity = $tr.childNodes.item($elementDXCCentity).outerText
            $Call = $tr.childNodes.item($elementCall).outerText
            #$QSLvia = $tr.childNodes.item($elementQSLvia).outerText
            #$ReportedBy = $tr.childNodes.item($elementReportedBy).outerText
            $Info = $tr.childNodes.item($elementInfo).outerText

            $startOfMonth = Get-Date $StartDate -day 1 -hour 0 -minute 0 -second 0
            $endOfMonth = (($EndDate).AddMonths(1).AddSeconds(-1))
#            if ( ($EndDate.Month -eq $CurrentDate.Month -or $StartDate.Month -le $CurrentDate.Month) -and ($EndDate.Year -eq $CurrentDate.Year) )
#            if ( ($EndDate.Month -ge $CurrentDate.Month -and $StartDate.Month -le $CurrentDate.Month) -and ($EndDate.Year -eq $CurrentDate.Year) )
            if ( ($startOfMonth -le $CurrentDate -and $endOfMonth -ge $CurrentDate) )
            {
              $call = RemoveLineCarriage($call)
              $startBracket = $Call.IndexOf("[")
              if ($startBracket -gt 0)
              {
                $call = $Call.substring(0,$startBracket).ToUpper()
                If (Get-ValidCall ($call)) 
                {
                  $arrDxCalls += $call
                 #Write-Host "YES $Remark " -ForegroundColor Green
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
                Write-Verbose "Call stripped : $Remark"
                If (Get-ValidCall ($Remark)) 
                {
#                  Write-Verbose "Call accepted : $Remark"
                  Write-Host "Call accepted : $Remark" -ForegroundColor Blue
                  $arrDxCalls += $Remark
                  #Write-Host "YES $Remark " -ForegroundColor Green
                } 
              }

            #Write-Host "YES : 0=$startdate 1=$EndDate  3=$Call "
            # Write-Host "YES : 0=$startdate 1=$EndDate 2=$time2 3=$time3 4=$time4 5=$time5 6=$time6 7=$time7 8=$time8"
          } else 
          {
            #Write-Host "NO  : 0=$startdate 1=$EndDate 2=$time2 3=$time3 4=$time4 5=$time5 6=$time6 7=$time7 8=$time8" -ForegroundColor Gray
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
      
      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"

    }
}


function Get-ContentDxWorld
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
        Throw "Unable to open website $URL ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Reading content from site $URL)"
      Try 
      {
        $CalendarFileContent = $result.Content -split "`r`n" | Where-Object { $_ -like "SUMMARY*"}

        foreach ($Line in $CalendarFileContent) 
        {
          $Line = $Line.ToUpper().Trim() 
#          Write-Host "Line --> $Line" -ForegroundColor Yellow
          If ($Line.StartsWith("SUMMARY")) 
          {
            $LineSplit = $Line.Split(':')
            $DxCalls = $LineSplit[1].Trim().Split(' ')
#            Write-Host "$LineSplit  $DxCallsLine  $DxCalls" -ForegroundColor Yellow

            foreach($DxCall In $DxCalls)
            { 
              $DxCall = $DxCall.Replace(")","")
              $DxCall = $DxCall.Replace("(","").Trim()
              Write-Verbose "Call stripped : $DxCall"
              If (Get-ValidCall ($DxCall)) 
              {
#                Write-Verbose "Call Accepted : $DxCall"
                Write-Host "Call accepted : $DxCall" -ForegroundColor Blue
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

      Write-Verbose "End Module : [$($MyInvocation.MyCommand)] *************************************"
   }
}


function Get-CallsFromFile
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$true)]
        [String] $CallsFile
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      $arrDxCalls =@()
  
      if (! (Test-Path $CallsFile)) 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to open file $CallsFile ($ErrorMessage, $FailedItem)" 
      }

      Try 
      {
        Write-Verbose "Reading extra callsign file $CallsFile)"
        $CallsFileContent = Get-Content "$CallsFile"  
      } Catch 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to read file $CallsFile  ($ErrorMessage, $FailedItem)" 
      }

      Write-Verbose "Reading content from file $CallsFile)"
      foreach ($Line in $CallsFileContent) 
      {
         $DxCall = $Line.ToUpper().Trim() 
         If (Get-ValidCall ($DxCall)) 
         {
            $arrDxCalls += $DxCall
         } 
      }
      
      Return $arrDxCalls | Sort-Object -Unique

      Write-Verbose "Eind Function  : [$($MyInvocation.MyCommand)] *************************************"
   }
}


function Get-CallsFromWebsite
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $CallsFile
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      
      $CalendarCalls = @()
      $Ng3KCalls = @()
      $objDxCalls= @()
      # https://dx-world.net/
      [uri]$CalendarURL       = 'https://www.google.com/calendar/ical/hamradioweb2007%40gmail.com/public/basic.ics'
      # Announced DX Operations
      [uri]$Ng3kURL           = "http://www.ng3k.com/misc/adxo.html"

      Write-Host "Reading website : $CalendarURL"  -ForegroundColor $DisplayColor
      $CalendarCalls = Get-ContentDxWorld -URL $CalendarURL
      Write-Verbose "Calls Readed : $CalendarCalls"
      Write-Host "Calls Loaded : $($CalendarCalls.Count)"  -ForegroundColor $DisplayColor
      #Write-DxCallsToCSV -objDxCalls $CalendarCalls -CSVFile $env:USERPROFILE\downloads\Calls_Agenda.csv

      Write-Host "Reading website : $Ng3kURL"  -ForegroundColor $DisplayColor
      $Ng3KCalls = Get-ContentNg3K -URL $Ng3kURL
      Write-Verbose "Calls Readed : $Ng3KCalls"
      Write-Host "Calls Loaded : $($Ng3KCalls.Count)" -ForegroundColor $DisplayColor
      #Write-DxCallsToCSV -objDxCalls $Ng3KCalls -CSVFile $env:USERPROFILE\downloads\Calls_Ng3K.csv

      $objDxCalls = Add-ConcatDxCall -DxCallsArray1 $CalendarCalls -DxCallsArray2 $Ng3KCalls
      if ($CallsFile)
      {
        Write-Host "Reading calls file  : $CallsFile"  -ForegroundColor $DisplayColor
        $ExtraCalls = Get-CallsFromFile -CallsFile $CallsFile
        Write-Verbose "Calls Readed : $ExtraCalls"
        Write-Host "Calls Loaded : $($ExtraCalls.Count)" -ForegroundColor $DisplayColor
        $objDxCalls = Add-ConcatDxCall -DxCallsArray1 $objDxCalls -DxCallsArray2 $ExtraCalls
      }
      
      $NumberInputCalls = $($objDxCalls.Count)
      $objDxCalls = Get-CallsCleanedUp -DxCallsArray $objDxCalls
      Write-Host "Remove duplicate callsigns, total $NumberInputCalls, returned $($objDxCalls.Count)"  -ForegroundColor $DisplayColor
      Write-Host "Calls readed : $($objDxCalls)"  -ForegroundColor $DisplayColor
      return $objDxCalls
      
      Write-Verbose "End Module : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Update-DxHrdAlarmFile
{
    <#
        .SYNOPSIS
            Function is to create a Dx Cluster Alarm file for use In Ham Radio DeLuxe
        .DESCRIPTION
            This function downloads callsign from DX Expeditions from websites and place It In the Alarm file from Ham Radio DeLuxe
            If there is no entry In the "DX Cluster Alarm Defenitions" It creates one, existing will be updated.
            WARNING : The file DXClusterAlarms.xml must exists!
        .PARAMETER $HRDLogbookDirectory
            Optional : Specify this parameter if the DXClusterAlarms.xml file is on an other location.
            Notes: 
                * Default path =  C:\Users\<user>\AppData\Roaming\HRDLLC\HRD Logbook
        .EXAMPLE
            PS> Update-DxHrdAlarmFile $HRDLogbookDirectory ""

            This example creates the alarmfile on default location.
        .EXAMPLE
            PS> Update-DxHrdAlarmFile $HRDLogbookDirectory "$env:USERPROFILE\downloads"

            This example creates a alarmfile on location "C:\Users\<user>\downloads\"
        .EXAMPLE
            PS> Update-DxHrdAlarmFile -Verbose

            This example gives extra information about the internal steps.
        .INPUTS
        .OUTPUTS
        $null
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $HRDLogbookDirectory,
        [Parameter(Position=1, Mandatory=$false)]
        [String] $ExtraCallsFile
    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      $Module = "PSDxExpeditions"
      $InstalledVersion  = (Get-Module -Name $Module).Version
      $currentVersion = "$($InstalledVersion.Major).$($InstalledVersion.Minor).$($InstalledVersion.Build).$($InstalledVersion.Revision)"
      Write-Host "Version module $($Module): $currentVersion"  -ForegroundColor White

      if ($HRDLogbookDirectory)
      {
        $AlarmFile = "$HRDLogbookDirectory\DXClusterAlarms.xml"
      } else
      {
        $AlarmFile = "$env:APPDATA\HRDLLC\HRD Logbook\DXClusterAlarms.xml"
      }
      
      if (! (Test-Path $AlarmFile)) 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to locate file $AlarmFile ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "File found : $AlarmFile"
      }
      
      $AllDxCalls = Get-CallsFromWebsite -CallsFile $ExtraCallsFile    
      
      Write-Host "Write Ham Radio DeLuxe Alarm File : $AlarmFile"  -ForegroundColor $DisplayColor
      Update-WriteHRDAlarmFile -DXClusterAlarmsFile $AlarmFile -DxCallsArray $AllDxCalls

      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Update-DxVe7ccAlarmFile
{
    <#
        .SYNOPSIS
            Function is to create a VE7CC Cluster Alarm file for use VE7CC Cluster software
        .DESCRIPTION
            This function downloads callsign from DX Expeditions from websites and place It In the Alarm file from VE7CC cluster
            WARNING : The file alarm.dat must exists!
        .PARAMETER VE7CCDirectory
            Mandatory : Specify this parameter to point to location of the directory where the alarm.dat and ve7cc.exe are installed.

        .PARAMETER RestartVE7CC
            Optional : Specify this parameter restart the VE7CC application.

        .EXAMPLE
            PS> Update-DxVe7ccAlarmFile -VE7CCDirectory "D:\VE7CC"

            This example change the alarmfile alarm.dat in this directory.
        .EXAMPLE
            PS> Update-DxVe7ccAlarmFile -VE7CCDirectory "D:\VE7CC" -RestartVE7CC

            This example change the alarmfile alarm.dat in this directory and restart VE7CC application, this will read the new alarm.dat.
        .EXAMPLE
            PS> Update-DxVe7ccAlarmFile -VE7CCDirectory "D:\VE7CC" -RestartVE7CC

            This example change the alarmfile alarm.dat in this directory and restart VE7CC application, this will read the new alarm.dat.
            This example gives extra information about the internal steps.
        .INPUTS
        .OUTPUTS
        $null
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $VE7CCDirectory,
        [Parameter(Position=1, Mandatory=$false)]
        [String] $ExtraCallsFile,
        [Parameter(Position=2, Mandatory=$false)]
        [Switch] $RestartVE7CC 

    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
 
      if ($VE7CCDirectory)
      {
        $AlarmFile = "$VE7CCDirectory\alarm.dat"
        $ve7ccProcessName = 've7cc'
        $ve7ccExe         = "$VE7CCDirectory\ve7cc.exe"
        $ve7ccActivity    = "VE7CC application, stopping and starting"
      }
      
      if ($DebugPreference -eq "Inquire")
      {
        $AlarmFile = "$env:USERPROFILE\downloads\alarm.dat"
        $ve7ccProcessName = 'notepad'
        $ve7ccExe         = "$env:WINDIR\notepad.exe"
        $ve7ccActivity    = "VE7CC application, stopping and starting"
        Write-Debug "Alarm file changend to $AlarmFile"
        Write-Debug ".exe changend to $ve7ccExe"
        Write-Host ".exe changend to $ve7ccExe" -ForegroundColor $DisplayColor
      }
      
      if (! (Test-Path $AlarmFile)) 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to locate file $AlarmFile ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "File found : $AlarmFile"
      }
      
      $AllDxCalls = Get-CallsFromWebsite -CallsFile $ExtraCallsFile    
      
      Write-Host "Write VE7CC Alarm File : $AlarmFile"  -ForegroundColor $DisplayColor
      Update-WriteVE7CCAlarmFile -AlarmFile $alarmFile -DxCallsArray $AllDxCalls

      if($RestartVE7CC)
      {
        Write-Host "$ve7ccActivity"  -ForegroundColor $DisplayColor
        if (! (Test-Path $ve7ccExe)) 
        {
          $ErrorMessage = $_.Exception.Message
          $FailedItem = $_.Exception.ItemName
          Throw "Unable to locate file $($ve7ccExe) ($ErrorMessage, $FailedItem)" 
        } else
        {
          Stop-Application -ProcessName $ve7ccProcessName -ProcessActivity $ve7ccActivity
          Start-Application -ProcessName $ve7ccProcessName -ProcessToStart $ve7ccExe -ProcessActivity $ve7ccActivity
        }
      }
      
      Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}


function Get-DxCheckApplicationRunning
{
    <#
        .SYNOPSIS
            Function is to check if VE7CC cluster is running, if not It's start the cluster application.
        .DESCRIPTION
            This function is to check if VE7CC cluster is running, if not It's start the cluster application.
            It looks for ve7cc.exe.
        .PARAMETER VE7CCDirectory
            Mandatory : Specify this parameter to point to location of the directory where the ve7cc.exe are installed.

        .EXAMPLE
            PS> Get-DxCheckApplicationRunning -VE7CCDirectory "D:\VE7CC"

            This example checks the process and if not running try to start.
        .INPUTS
        .OUTPUTS
        $null
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0, Mandatory=$false)]
        [String] $VE7CCDirectory

    )
    begin
    {
      Write-Verbose "Start Module : [$($MyInvocation.MyCommand)] *************************************"
      if ($VE7CCDirectory)
      {
        $ve7ccProcessName = 've7cc'
        $ve7ccExe         = "$VE7CCDirectory\ve7cc.exe"
        $ve7ccActivity    = "VE7CC application, stopping and starting"
      }
      
      if ($DebugPreference -eq "Inquire")
      {
        $AlarmFile        = "$env:USERPROFILE\downloads\alarm.dat"
        $ve7ccProcessName = 'notepad'
        $ve7ccExe         = "$env:WINDIR\notepad.exe"
        $ve7ccActivity    = "VE7CC application, stopping and starting"
        Write-Debug "Alarm file changend to $AlarmFile"
        Write-Debug ".exe changend to $ve7ccExe"
        Write-Host ".exe changend to $ve7ccExe" -ForegroundColor $DisplayColor
      }

      if (! (Test-Path $ve7ccExe)) 
      {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Throw "Unable to locate file $ve7ccExe ($ErrorMessage, $FailedItem)" 
      } else
      {
        Write-Verbose "File found : $ve7ccExe"
      }

      Write-Verbose "Checking if process $($ve7ccProcessName) is running or not."
      If (Get-Process -Name $($ve7ccProcessName) -ErrorAction SilentlyContinue)
      {
    	Write-Verbose "Process $($ve7ccProcessName) is running."
      } else 
      {
    	Write-Verbose "Process $($ve7ccProcessName) is NOT running, trying to start now."
        Write-Host "Process $($ve7ccProcessName) is NOT running, trying to start now." -ForegroundColor $DisplayColor
        Start-Application -ProcessName $ve7ccProcessName -ProcessToStart $ve7ccExe -ProcessActivity $ve7ccActivity
      }

  
        Write-Verbose "End Module  : [$($MyInvocation.MyCommand)] *************************************"
    }
}



New-Alias -Name UdxH -Value Update-DxHrdAlarmFile
New-Alias -Name UdxV -Value Update-DxVe7ccAlarmFile
New-Alias -Name GdxC -Value Get-DxCheckApplicationRunning
Export-ModuleMember -function Update-DxHrdAlarmFile -alias UdxH
Export-ModuleMember -function Update-DxVe7ccAlarmFile -alias UdxV
Export-ModuleMember -function Get-DxCheckApplicationRunning -alias GdxC

# SIG # Begin signature block
# MIIFqQYJKoZIhvcNAQcCoIIFmjCCBZYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUx+T9C1A+JAd7Z9LgsDvBASvA
# vQKgggMwMIIDLDCCAhSgAwIBAgIQY8/oI4lwML5P5YTr9cdxsTANBgkqhkiG9w0B
# AQUFADAuMSwwKgYDVQQDDCNQQTFSRUcgUFNEeEV4cGVkaXRpb25zIENvZGUgU2ln
# bmluZzAeFw0xNzExMjMxMDM3MjNaFw0yNzExMjMxMDQ3MjRaMC4xLDAqBgNVBAMM
# I1BBMVJFRyBQU0R4RXhwZWRpdGlvbnMgQ29kZSBTaWduaW5nMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuzl4mychYkRaOuPeFX3M//hAKhhmpEnRnloB
# W0T0X0N+olbEcYRcQzlelOYlLfp6nvyVjrv5AyCEoa0lVkRcY+OvZjRVmhN5HNys
# OeQCvrrYU1YPZ6KQH5fJjNNOfpr248mwuRAWFWYOiWnjzBOFmUUbig6REYB6nkxE
# M3ruCTOdcEztPV26pZqlpWRR7i0wFlh9qDrbuyKe3tHvYMnYncFoBuzNKo8uq7Fr
# TNv6RV6KO/LH9T2trnZqpm+gg6D1xcUDEEEZIn0hWZokbIj9uylWiRM0dZwFZZyB
# aX5rjtnlcz9xoBmsXtybTPHVqR4hSc4jX+QoykvP71Ef179EsQIDAQABo0YwRDAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFJBh
# zzqM97xW8cc7dhhA/PV7FPL9MA0GCSqGSIb3DQEBBQUAA4IBAQB3aFNc3pjD2nEY
# qn1WeM5B+WzR9mLH9T40v4ECRugfyQNa9oV9m90g26t4+0eE6eGC5NQj9aLNECpc
# Qjnonof4t9u70cDQouwT9tPNm+dTXvCkIOaIeo5ispyPYydpleEqk7iaw3SnukVN
# tOoQuQ+KKBDjiHy7nHIpIS4t8mJSoATc4jt+O9wz8EVh5g+Oi4avwdbsjkeYLfMt
# BKFE0cpZShsFoeYu9Lmp4UpAf4DcsxRfarT4Gg53Xr6hAUP0INwWmrBL94vIRktM
# bKlGW6HcF7vsHbU0UE7n81t7wdvH7v2xXllfX0FYjkUTAUgs7wYY8Hrn3UGHBgFr
# dq5Mgjx6MYIB4zCCAd8CAQEwQjAuMSwwKgYDVQQDDCNQQTFSRUcgUFNEeEV4cGVk
# aXRpb25zIENvZGUgU2lnbmluZwIQY8/oI4lwML5P5YTr9cdxsTAJBgUrDgMCGgUA
# oHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYB
# BAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0B
# CQQxFgQUf3DtHHX3eC48ZJRwrztlyKaafk4wDQYJKoZIhvcNAQEBBQAEggEAtMr9
# ScEIqmB0jX/VVRPga15Nl7xoUbE0uo1WxrUG0qYzvDZZQEsVYHW4Azvy7pKtrmpt
# p1cWBkpWJIfcM9mRj8m/YUprwAPlZpSF8CVsytoiu8dc6PwKev8IjOqwgUtqgwMI
# Gt/sV4Yh5U2N+zeBPgNfOLK19t+PCigy68EwHTIkoY9mH76DIcvcz3ZTIuHK8TEt
# hobmHT6Ug8+cFLHvKP8YgZDyABUPyigLWo4b6wxq8C4dHm8drNeSgULNn2XRp+22
# 3YSsKhENXr6qq6u1WolQdQ9YHI8Mv9+kF85GAzqSfxx3IicQgq3XJ3l4UtpkX43B
# IJcHdsgQsHIkiFBpmQ==
# SIG # End signature block
