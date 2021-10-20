$DebugTest = $false
$DiscordMessage = $true

#region Library Files
###############Library Files###############

. "$PSScriptroot\Library\DiscordMessageFunctions.ps1"
. "$PSScriptroot\Library\OperationalFunctions.ps1"

#############End Library Files#############
Add-Type -Assembly "Microsoft.Office.Interop.Outlook"

If ($DebugTest -eq $false) {
  #Prod Channel
  $WebhookUrl = "ProdWebHookURI"
}
Else {
  #Dev Channel
  $WebHookURL = "DevWebHookURI"
}

$DayOfWeek = (get-date).dayofweek.value__
If (($DayOfWeek -eq 6 -or $DayOfWeek -eq 0) -and $debugTest -eq $false) {
  Exit
}
$MainPath = "C:\Code\TradeReporter"
$ErrorActionPreference = "Continue"
$Test = Test-Path -Path "$MainPath\Logs\Output.txt"
If ($Test) {
  Remove-Item -Path "$MainPath\Logs\Output.txt" -Force
}
Start-Transcript -path "$MainPath\Logs\Output.txt" -append

$TestPath = Test-Path -Path $MainPath
If (!$TestPath) {
  New-Item -Path $MainPath -ItemType "directory"
}

#$Folders = Get-ChildItem $mainpath | Where-Object { ($_.psIsContainer -eq $true) -AND ($_.Name -ne "Archived") }
<#
ForEach ($Folder in $Folders) {
  #  $FolderDate = $Folder.LastWriteTime
    $FolderDate = $Folder.Name
    $Split = $FolderDate.split("-")
    [int]$FolderMonth = $Split[0]
    $FolderDate = (Get-Date -Month $FolderMonth -f MM)
    If ($FolderDate -lt (Get-Date).AddMonths(-1)) {
  #  If ($Folder.LastWriteTime.Month -lt $Month) {
        $FolderDate = (Get-Date -Month $FolderDate -f MM-yyyy)
        $Destination = "$($MainPath)\Archived\$($FolderDate)"
        $Test = Test-Path -Path "$($Destination).zip"
        If (!$Test) {
            Write-Host -ForegroundColor Green "Creating archive file $($Destination)"
            Compress-Archive -Path "$($MainPath)\$($Folder)" -DestinationPath $Destination -Verbose
        } Else {
            Write-Host -ForegroundColor Green "Updating archive file $($Destination)"
            Compress-Archive -Path "$($MainPath)\$($Folder)" -Update -DestinationPath $Destination -Verbose
        }
        Write-Host -ForegroundColor Red "Removing $($MainPath)\$($Folder)"
        Remove-Item -Recurse -Force "$($MainPath)\$($Folder)"
    }
}
#>

$TotalDailyTrades = $null
$TotalDailyTrades = @()
$ClosedTrades = $null
$ClosedTrades = @()
$TemporaryDailyTrades = $null
$TemporaryDailyTrades = @()
$TradeLog = $null
$TradeLog = @()

#Load any existing trades (in case a crash occurs)
If (Test-path "$MainPath\Logs\TotalDailyTrades.csv") {
  $TemporaryDailyTrades += Import-Csv -Path "$MainPath\Logs\TotalDailyTrades.csv"  | Select-Object Order, Type, Strike, Expiration, Security, @{Name = "AvgPrice"; Expression = { [decimal]$_.AvgPrice } }, @{Name = "Quantity"; Expression = { [int]$_.Quantity } }, `
  @{Name = "Total"; Expression = { [decimal]$_.Total } }, @{Name = "RunningTotal"; Expression = { [decimal]$_.RunningTotal } }, @{Name = "RunningSold"; Expression = { [decimal]$_.RunningSold } }, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }, `
  @{Name = "TimeOpened"; Expression = { [datetime]$_.TimeOpened } }, @{Name = "TimeClosed"; Expression = { [datetime]$_.TimeClosed } }, @{Name = "TimeUpdated"; Expression = { [datetime]$_.TimeUpdated } }, @{Name = "TimeSpan"; Expression = { [datetime]$_.TimeSpan } }, Status
} 
if (Test-path "$MainPath\OpenPositions.csv") {
  $TemporaryDailyTrades += Import-Csv "$MainPath\Logs\OpenPositions.csv" | Select-Object Order, Type, Strike, Expiration, Security, @{Name = "AvgPrice"; Expression = { [decimal]$_.AvgPrice } }, @{Name = "Quantity"; Expression = { [int]$_.Quantity } }, `
  @{Name = "Total"; Expression = { [decimal]$_.Total } }, @{Name = "RunningTotal"; Expression = { [decimal]$_.RunningTotal } }, @{Name = "RunningSold"; Expression = { [decimal]$_.RunningSold } }, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }, `
  @{Name = "TimeOpened"; Expression = { [datetime]$_.TimeOpened } }, @{Name = "TimeClosed"; Expression = { [datetime]$_.TimeClosed } }, @{Name = "TimeUpdated"; Expression = { [datetime]$_.TimeUpdated } }, @{Name = "TimeSpan"; Expression = { [datetime]$_.TimeSpan } }, Status
}
Foreach ($Object in $TemporaryDailyTrades) {
  If (($Totaldailytrades | Where-Object Time -eq $Receivedtime)) {
    Write-Host "This order already exists in memory. Skipping."
    $Object
  }
  Else {
    Write-Host "Adding object from Total Daily Trades Log to memory"
    $TotalDailyTrades += $Object
  }
}

$TotalDailyTrades
$OpenPositions = $TotalDailyTrades | Where-Object Status -eq "Open"

#Clear-host
$limit = (Get-Date).AddMinutes(420)

If ($OpenPositions.count -eq 1) {
  $Message = @"
Good morning,

I'm alive, active and will be posting trades today. 
There is $($OpenPositions.count) open position currently from previous trading sessions.

-Trade Reporter
"@
}
elseif ($OpenPositions.count -gt 1) {
  $Message = @"
Good morning,

I'm alive, active and will be posting trades today. 
There are $($OpenPositions.count) open positions currently from previous trading sessions.

-Trade Reporter
"@
}
else {
  $Message = @"
Good morning,

I'm alive, active and will be posting trades today.

-Trade Reporter
"@
}

If ($DiscordMessage) {
  Live-Message -Message $Message
}
If ($DebugTest) {
  $Trades = Connect-Outlook
  $UnreadCount = ($Trades | Where-Object UnRead -eq $True ).Count
  $Count = 1
  Write-Host "There are $UnreadCount unread emails"
}
do {
  #loop while market is open (bot stops at 4PM)
  $Trades = Connect-Outlook
  If ($Trades) {
    $Trades | Where-Object { $_.UnRead -eq $True } | ForEach-Object {
      #Parse Email Body
      $Body = $_.Body
      [dateTime]$ReceivedTime = $_.ReceivedTime
      $ReceivedTime

      #Save Directories
      $lastFridayOfMonth = last-dayofweek (get-date -f 'yyyy') "Friday" $ReceivedTime.ToString('MM')

      If ($ReceivedTime.ToString('MM-dd-yyyy') -le $lastFridayOfMonth) {
        $Month = $ReceivedTime.ToString('MM') 
        $Year = $ReceivedTime.ToString('yyyy') 
      }
      else {
        $Month = $ReceivedTime
        $Month = $Month.AddMonths(1).ToString('MM')
        If ($Month -eq 01) {
          $Year = $ReceivedTime
          $Year = $Year.AddYears(1).ToString('yyyy')
        }
      }

      $Test = Test-Path -Path "$($MainPath)\Archived"
      If (!$Test) {
        Write-Host "Creating Archived directory"
        New-Item -Path "$($MainPath)\Archived" -ItemType "directory"
      }
      $Test = Test-Path -Path "$($MainPath)\Reports\$($Year)\$($Month)"
      If (!$Test) {
        Write-Host "Creating Month Directory"
        New-Item -Path "$($MainPath)\Reports\$($Year)" -Name "$($Month)" -ItemType "directory"
      }

      $CurrentDayValue = ([datetime]$ReceivedTime).dayofweek.value__
      If ($CurrentDayValue -eq 1) {
        $StoredDayValue = 1
        $FolderNames = (Get-ChildItem "$($MainPath)\Reports\$($Year)\$($Month)").Name
        If ($FolderNames) {
          $Split = $null
          Foreach ($Name in $FolderNames) {
            [array]$Split += $Name.Split("-")[1]
          }
          $WeekNum = ($Split | Measure-Object -Maximum).Maximum
        }
        Else {
          $WeekNum = 1
        }
        $Split = $null
        $Test = Test-Path -Path "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum"
        If ($Test) {
          $ReportNames = (Get-ChildItem "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum").Name
        }
        If ($ReportNames) {
          Foreach ($Name in $ReportNames) {
            $Temp = $Name.Split("-")[-1]
            [array]$Split += $Temp.Split(".")[0]
          }
          If ($Split -gt 1) {
            $WeekNum++
          }
        }
        
        If (!$Test) {
          Write-Host "Creating Week Directory in Month $($Month) folder"
          New-Item -Path "$($MainPath)\Reports\$($Year)\$($Month)" -Name "Week-$($WeekNum)" -ItemType "directory"
        }
      }
      else {
        If ($FolderNames) {
          $Split = $null
          Foreach ($Name in $FolderNames) {
            [array]$Split += $Name.Split("-")[1]
          }
          $WeekNum = ($Split | Measure-Object -Maximum).Maximum
        }
        Else {
          $WeekNum = 1
          $Test = Test-Path -Path "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum"
          If (!$Test) {
            Write-Host "Creating Week directory in Month $($Month) Folder"
            New-Item -Path "$($MainPath)\Reports\$($Year)\$($Month)" -Name "Week-$($WeekNum)" -ItemType "directory"
          }
        }
        if ($CurrentDayValue -gt $StoredDayValue -and $ClosedTrades -and $DebugTest -eq $true) {
          #EOD Report
          $EODObject = EOD-ReportBuild -ClosedTrades $ClosedTrades -TotalDailyTrades $TotalDailyTrades
          If ($DiscordMessage) {
            EOD-Report -TradeObject $EODObject
          }
          If ($EODObject) {
            $EODObject | Export-Csv "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum\EODReport-$StoredDayValue.csv" -NoTypeInformation
          }

          $OpenPositions = $TotalDailyTrades | Where-Object Status -eq "Open" 
          If ($OpenPositions) {
            $TotalDailyTrades | Where-Object Status -eq "Open" | Export-Csv "$MainPath\Logs\OpenPositions.csv" -NoTypeInformation
          }
          #If the EOD is reached then there's no need to keep the memory dump file
          If (Test-Path "$MainPath\Logs\TotalDailyTrades.csv") {
            Remove-Item "$MainPath\Logs\TotalDailyTrades.csv" -Force
          }
          $StoredDayValue = $CurrentDayValue
          $TotalDailyTrades = @()
          $ClosedTrades = @()
          $TemporaryDailyTrades = @()

          if (Test-path "$MainPath\Logs\OpenPositions.csv") {
            $TemporaryDailyTrades += Import-Csv "$MainPath\Logs\OpenPositions.csv" | Select-Object Order, Type, Strike, Expiration, Security, @{Name = "AvgPrice"; Expression = { [decimal]$_.AvgPrice } }, @{Name = "Quantity"; Expression = { [int]$_.Quantity } }, `
            @{Name = "Total"; Expression = { [decimal]$_.Total } }, @{Name = "RunningTotal"; Expression = { [decimal]$_.RunningTotal } }, @{Name = "RunningSold"; Expression = { [decimal]$_.RunningSold } }, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }, `
            @{Name = "TimeOpened"; Expression = { [datetime]$_.TimeOpened } }, @{Name = "TimeClosed"; Expression = { [datetime]$_.TimeClosed } }, @{Name = "TimeUpdated"; Expression = { [datetime]$_.TimeUpdated } }, @{Name = "TimeSpan"; Expression = { [datetime]$_.TimeSpan } }, Status
          }
          Foreach ($Object in $TemporaryDailyTrades) {
            If (($Totaldailytrades | Where-Object Time -eq $Receivedtime)) {
              Write-Host "This order already exists in memory. Skipping."
            }
            Else {
              Write-Host "Adding object from Total Daily Trades Log to memory"
              $TotalDailyTrades += $Object  

            }
          }
        } 
      }

      $TempObject = Body-Parser -Body $Body
      $TradeLog += $TempObject

      # Test for existing orders
      If ($TotalDailyTrades.Security -contains $TempObject.Security) {
        Write-Host "Order with that security has been found"
        $CurrentPosition = $TotalDailyTrades | Where-Object Security -eq $TempObject.Security | Where-Object Type -eq $TempObject.Type | Where-Object Status -eq "Open"
        Write-Host "----------------"
        Write-Host "Current Position"
        $CurrentPosition
        Write-Host "----------------"

        If ($CurrentPosition.Order -ne $TempObject.Order -and $CurrentPosition.Type -eq $TempObject.Type -and $CurrentPosition.Quantity -eq $TempObject.Quantity) {
          # Close Order
          $ReturnObject = Trade-Manager -TradeObject $TempObject -ReceivedTime $ReceivedTime -Action "Close"
          
          $ClosedTrades += $ReturnObject
          If ($DiscordMessage) {
            Close-Message -TradeObject $ReturnObject
          }
          
          Write-Host "----------------"
          Write-Host "Closing Position"
          $ReturnObject
          Write-Host "----------------"

        }
        elseif ($CurrentPosition.Order -eq $TempObject.Order -and $CurrentPosition.Type -eq $TempObject.Type) {
          # Add to Order
          $ReturnObject = Trade-Manager -TradeObject $TempObject -ReceivedTime $ReceivedTime -Action "Add"
          If ($TradeObject.Type -eq 'Short' -and $TradeObject.Order -eq 'Buy') {
            $Color = "16711680"
          }
          else {
            $Color = "4289797"
          }
          If ($DiscordMessage) {
            Position-Update -TradeObject $ReturnObject -Color $Color -Order $($CurrentPosition.Order)
          }
          Write-Host "Position Update Add"
          $ReturnObject
          Write-Host "----------------"

        }
        elseif ($CurrentPosition.Order -ne $TempObject.Order -and $CurrentPosition.Type -eq $TempObject.Type -and $CurrentPosition.Quantity -ne $TempObject.Quantity) {
          # Subtract from Order
          $ReturnObject = Trade-Manager -TradeObject $TempObject -ReceivedTime $ReceivedTime -Action "Subtract"
          If ($TradeObject.Type -eq 'Short' -and $TradeObject.Order -eq 'Sell') {
            $Color = "4289797"
          }
          else {
            $Color = "16711680"
          }
          If ($DiscordMessage) {
            Position-Update -TradeObject $ReturnObject -Color $Color -Order $($CurrentPosition.Order)
          }
          Write-Host "Position Update Subtract"
          $ReturnObject
          Write-Host "----------------"

        }
        elseif ($CurrentPosition.Type -ne $TempObject.Type) {
          # Different Order
          $ReturnObject = Trade-Manager -TradeObject $TempObject -ReceivedTime $ReceivedTime -Action "Different"
          If ($DiscordMessage) {
            Order-Message -TradeObject $ReturnObject
          }
          Write-Host "Different Order"
          $ReturnObject
          Write-Host "----------------"
        }
        else {}
      }
      Else {
        # New Order
        #There are no existing orders
        $TotalDailyTrades += $TempObject
        If ($DiscordMessage) {
          Order-Message -TradeObject $TempObject
        }
        Write-Host "New Order"
        $TempObject
        Write-Host "----------------"
      }
          
      If ($debugTest -eq $False) {
        $_.UnRead = $False
      }

      If ($DiscordMessage) {
        Start-Sleep -Seconds 2
      }

      $Count++
      $Count
    } # End Foreach loop of emails

  }
  Else {
    #Error Message
    $Message = @"
I'm unable to connect to the mailbox and retrieve messages.
"@
    If ($DiscordMessage) {
      Live-Message -Message $Message
    }
  }
  
  # Cleanup Open Trades File
  $OpenPositions = $TotalDailyTrades | Where-Object Status -eq "Open"
  If ($OpenPositions.count -eq 0 -and (test-path "$MainPath\Logs\OpenPositions.csv")) {
    Remove-Item "$MainPath\Logs\OpenPositions.csv" -Force
  }
    
  # Dump memory object to file in case of crash
  If ($TotalDailyTrades) {
    $TotalDailyTrades | Export-Csv "$MainPath\Logs\TotalDailyTrades.csv" -NoTypeInformation
  }
  
  If ($DebugTest -eq $True) {
    #EOD Report
    If ($ClosedTrades) {
      $EODObject = EOD-ReportBuild -ClosedTrades $ClosedTrades -TotalDailyTrades $TotalDailyTrades
      If ($DiscordMessage) {
        EOD-Report -TradeObject $EODObject
      }
      $DayValue = ([datetime]$ClosedTrades[-1].timeclosed).dayofweek.value__
      If ($EODObject) {
        $EODObject | Export-Csv "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum\EODReport-$DayValue.csv" -NoTypeInformation
      }
      
      $OpenPositions = $TotalDailyTrades | Where-Object Status -eq "Open"
      If ($OpenPositions) {
        $TotalDailyTrades | Where-Object Status -eq "Open" | Export-Csv "$MainPath\Logs\OpenPositions.csv" -NoTypeInformation
      }
      #If the EOD is reached then there's no need to keep the memory dump file
      If (Test-Path "$MainPath\Logs\TotalDailyTrades.csv") {
        Remove-Item "$MainPath\Logs\TotalDailyTrades.csv" -Force
      }
      
    }
  }
  Start-Sleep -Seconds 30

} Until ((Get-Date) -ge $limit -or ($debugTest -and $Count -gt $UnreadCount))

#EOD Report
If ($ClosedTrades) {
  $EODObject = EOD-ReportBuild -ClosedTrades $ClosedTrades -TotalDailyTrades $TotalDailyTrades
  If ($DiscordMessage) {
    EOD-Report -TradeObject $EODObject
  }
  $DayValue = ([datetime]$ClosedTrades[-1].timeclosed).dayofweek.value__
  If ($EODObject) {
    $EODObject | Export-Csv "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum\EODReport-$DayValue.csv" -NoTypeInformation
  }

  If ($OpenPositions) {
    $TotalDailyTrades | Where-Object Status -eq "Open" | Export-Csv "$MainPath\Logs\OpenPositions.csv" -NoTypeInformation
  }

  #If the EOD is reached then there's no need to keep the memory dump file
  If (Test-Path "$MainPath\Logs\TotalDailyTrades.csv") {
    Remove-Item "$MainPath\Logs\TotalDailyTrades.csv" -Force
  }
}

If ($DayOfWeek -eq 5 -or $DebugTest -eq $true) {
  $EOWObject = EOW-ReportBuild -EODReportPath "$($MainPath)\Reports\$($Year)\$($Month)\Week-$WeekNum"
  If ($DiscordMessage) {
    EOW-Report -TradeObject $EOWObject
  }
  If ($EOWObject) {
    $EOWObject | Export-Csv "$($MainPath)\Reports\$($Year)\$($Month)\EOWReport-$WeekNum.csv" -NoTypeInformation
  }
}

If ((Get-Date -f 'MM-dd-yyyy') -ge $lastFridayOfMonth) {
  $EOMObject = EOM-ReportBuild -EOWReportPath "$($MainPath)\Reports\$($Year)\$($Month)"
  If ($DiscordMessage) {
    EOM-Report -TradeObject $EOMObject
  }
  If ($EOMObject) {
    Write-Host "Last Friday of the month is: $lastFridayOfMonth"
    $EOMObject | Export-Csv "$($MainPath)\Reports\$($Year)\EOMReport-$Month.csv" -NoTypeInformation
  }
}

$lastFridayOfMonth = last-dayofweek (get-date -f 'yyyy') "Friday" '12'
If ((Get-Date -f 'MM-dd-yyyy') -ge $lastFridayOfMonth) {
  $EOYObject = EOY-ReportBuild -EOYReportPath "$($MainPath)\Reports\$($Year)\$($Month)"
  If ($DiscordMessage) {
    EOM-Report -TradeObject $EOMObject
  }
  If ($EOMObject) {
    Write-Host "Last Friday of the year is: $lastFridayOfMonth"
    $EOMObject | Export-Csv "$($MainPath)\Reports\$($Year)\EOMReport-$Month.csv" -NoTypeInformation
  }
}

#EOD Message
$Message = @"
Welp, that's it for me. Peace out!
"@
If ($DiscordMessage) {
  Live-Message -Message $Message
}
Stop-Transcript
Exit



