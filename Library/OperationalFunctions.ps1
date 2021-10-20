Function Connect-Outlook {
    [CmdletBinding()]
    Param
    (
    )
    BEGIN {
    }
 
    PROCESS {
        $Outlook = New-Object -ComObject Outlook.Application
        $mapi = $Outlook.GetNameSpace("MAPI")
        $Trades = $mapi.folders.item(1).folders.item('Trades').items | Sort-Object receivedtime
    }
    END {
        If ($Trades) {
            Return $Trades
        }
        Else {
            Return Write-Error "Was unable to fetch messages from Outlook folder"
        }
    }
} #END Function

function last-dayofweek {
    param(
     [Int][ValidatePattern("[1-9][0-9][0-9][0-9]")]$year,
     [String][validateset('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')]$dayofweek,
     [int][validateset(1,2,3,4,5,6,7,8,9,10,11,12)]$month
    )
    $date = (Get-Date -Year $year -Month 1 -Day 1)
    while($date.DayOfWeek -ne $dayofweek) {$date = $date.AddDays(1)}
    while($date.year -eq $year) {
        if($date.Month -ne $date.AddDays(7).Month -and $date.Month -eq $month) {$date.ToString("MM-dd-yyyy")}
        $date = $date.AddDays(7)
    }
  }

Function EOD-ReportBuild {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [object]$ClosedTrades,
        [Parameter(Mandatory = $true)]
        [object]$TotalDailyTrades
    )
    BEGIN {
    }
   
    PROCESS {
        $WinCount = 0
        $LossCount = 0
        $TotalCount = $ClosedTrades.Count
        $TotalProfit = $null
        Foreach ($EODProfit in $ClosedTrades.Profit) {
            If ($EODProfit -gt 0) {
                $WinProfit += $EODProfit
                $WinCount++
            }
            else {
                $Losses -= $EODProfit
                $LossCount ++
            }
            $TotalProfit += $EODProfit
        }
          
        $WinRate = ($WinCount / $TotalCount) * 100
        $LossRate = ($LossCount / $TotalCount) * 100
        [string]$WinRatePercent = [math]::Round($WinRate, 2).ToString() + "%"
        [string]$LossRatePercent = [math]::Round($LossRate, 2).ToString() + "%"
        $Securities = $ClosedTrades.Security
        $Profits = $ClosedTrades.Profit
        $Types = $ClosedTrades.Type
        $SecurityString = $null
        $ProfitString = $null
        $TypeString = $null
        $OpenSecuritiesString = $null
        $OpenQuantityString = $null
        $OpenAvgPriceString = $null
      
        Foreach ($Security in $Securities) {
            $SecurityString += $Security.ToString() + "\n"
        }
        Foreach ($Profit in $Profits) {
            If ($Profit -lt 0) {
                $InvertCalc = $Profit * -1
                $TempProfit = '-$' + $InvertCalc.ToString()
            }
            else {
                $TempProfit = '$' + $Profit.ToString()
            }
            $ProfitString += $TempProfit + "\n"
        }
        Foreach ($Type in $Types) {
            $TypeString += $Type.ToString() + "\n"
        }
        $OpenPositions = $TotalDailyTrades | Where-Object Status -eq "Open"
        Foreach ($Position in $OpenPositions) {
            $OpenSecuritiesString += $Position.Security.ToString() + "\n"
            $OpenQuantityString += $Position.Quantity.ToString() + " - " + $Position.Type.ToString() + "\n"
            $OpenAvgPriceString += $Position.AvgPrice.ToString() + "\n"
        }
      
        $EODObject = [PSCustomObject]@{
            TradeCount     = $TotalCount
            WinCount       = $WinCount
            LossCount      = $LossCount
            WinRate        = $WinRatePercent
            LoseRate       = $LossRatePercent
            Profit         = $TotalProfit
            OpenSecurities = $OpenSecuritiesString
            OpenQuantity   = $OpenQuantityString
            OpenAvgPrice   = $OpenAvgPriceString
            Securities     = $SecurityString
            Profits        = $ProfitString
            Types          = $TypeString
        }
        
    }
    END {
        If ($EODObject) {
            Return $EODObject
        }
        Else {
            Return Write-Error "Was unable to build end of day report"
        }
    }
} #END Function

Function EOW-ReportBuild {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$EODReportPath
    )
    BEGIN {
    }
   
    PROCESS {
        $WeeklyReport = @()
        $DailyReports = Get-ChildItem $EODReportPath
        Foreach ($Report in $DailyReports) {
            $WeeklyReport += Import-Csv -Path $Report.FullName | Select-Object @{Name = "TradeCount"; Expression = { [int]$_.TradeCount } }, @{Name = "WinCount"; Expression = { [int]$_.WinCount } }, `
            @{Name = "LossCount"; Expression = { [int]$_.LossCount } }, WinRate, LoseRate, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }
        }

        $TotalTradeCount = $null
        $TotalWinCount = $null
        $TotalLossCount = $null
        $TotalProfit = $null
        Foreach ($Report in $WeeklyReport) {
            $TotalTradeCount += $Report.TradeCount
            $TotalWinCount += $Report.WinCount
            $TotalLossCount += $Report.LossCount
            $TotalProfit += $Report.Profit
        }

        $AvgDailyProfit = $TotalProfit / $DailyReports.Count
    
        $TotalWinRate = ($TotalWinCount / $TotalTradeCount) * 100
        $TotalLossRate = ($TotalLossCount / $TotalTradeCount) * 100
        [string]$TotalWinRatePercent = [math]::Round($TotalWinRate, 2).ToString() + "%"
        [string]$TotalLossRatePercent = [math]::Round($TotalLossRate, 2).ToString() + "%"

        $EOWObject = [PSCustomObject]@{
            TradeCount     = $TotalTradeCount
            WinCount       = $TotalWinCount
            LossCount      = $TotalLossCount
            WinRate        = $TotalWinRatePercent
            LoseRate       = $TotalLossRatePercent
            Profit         = $TotalProfit
            AvgDailyProfit = $AvgDailyProfit
        }
    }
    END {
        If ($EOWObject) {
            Return $EOWObject
        }
        Else {
            Return Write-Error "Was unable to build end of week report"
        }
    }
} #END Function

Function EOM-ReportBuild {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$EOWReportPath
    )
    BEGIN {
    }
   
    PROCESS {
        $MonthlyReport = @()
        $WeeklyReports = Get-ChildItem $EOWReportPath
        Foreach ($Report in $WeeklyReports) {
            $WeeklyReport += Import-Csv -Path $Report.FullName | Select-Object @{Name = "TradeCount"; Expression = { [int]$_.TradeCount } }, @{Name = "WinCount"; Expression = { [int]$_.WinCount } }, `
            @{Name = "LossCount"; Expression = { [int]$_.LossCount } }, WinRate, LoseRate, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }
        }

        $TotalTradeCount = $null
        $TotalWinCount = $null
        $TotalLossCount = $null
        $TotalProfit = $null
        Foreach ($Report in $MonthlyReport) {
            $TotalTradeCount += $Report.TradeCount
            $TotalWinCount += $Report.WinCount
            $TotalLossCount += $Report.LossCount
            $TotalProfit += $Report.Profit
        }

        $AvgWeeklyProfit = $TotalProfit / $WeeklyReports.Count
    
        $TotalWinRate = ($TotalWinCount / $TotalTradeCount) * 100
        $TotalLossRate = ($TotalLossCount / $TotalTradeCount) * 100
        [string]$TotalWinRatePercent = [math]::Round($TotalWinRate, 2).ToString() + "%"
        [string]$TotalLossRatePercent = [math]::Round($TotalLossRate, 2).ToString() + "%"

        $EOMObject = [PSCustomObject]@{
            TradeCount     = $TotalTradeCount
            WinCount       = $TotalWinCount
            LossCount      = $TotalLossCount
            WinRate        = $TotalWinRatePercent
            LoseRate       = $TotalLossRatePercent
            Profit         = $TotalProfit
            AvgWeeklyProfit = $AvgWeeklyProfit
        }
    }
    END {
        If ($EOMObject) {
            Return $EOMObject
        }
        Else {
            Return Write-Error "Was unable to build end of week report"
        }
    }
} #END Function

Function EOY-ReportBuild {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$EOMReportPath
    )
    BEGIN {
    }
   
    PROCESS {
        $YearlyReport = @()
        $MonthlyReports = Get-ChildItem $EOMReportPath
        Foreach ($Report in $MonthlyReports) {
            $YearlyReport += Import-Csv -Path $Report.FullName | Select-Object @{Name = "TradeCount"; Expression = { [int]$_.TradeCount } }, @{Name = "WinCount"; Expression = { [int]$_.WinCount } }, `
            @{Name = "LossCount"; Expression = { [int]$_.LossCount } }, WinRate, LoseRate, @{Name = "Profit"; Expression = { [decimal]$_.Profit } }
        }

        $TotalTradeCount = $null
        $TotalWinCount = $null
        $TotalLossCount = $null
        $TotalProfit = $null
        Foreach ($Report in $MonthlyReport) {
            $TotalTradeCount += $Report.TradeCount
            $TotalWinCount += $Report.WinCount
            $TotalLossCount += $Report.LossCount
            $TotalProfit += $Report.Profit
        }

        $AvgWeeklyProfit = $TotalProfit / $MonthlyReports.Count
    
        $TotalWinRate = ($TotalWinCount / $TotalTradeCount) * 100
        $TotalLossRate = ($TotalLossCount / $TotalTradeCount) * 100
        [string]$TotalWinRatePercent = [math]::Round($TotalWinRate, 2).ToString() + "%"
        [string]$TotalLossRatePercent = [math]::Round($TotalLossRate, 2).ToString() + "%"

        $EOYObject = [PSCustomObject]@{
            TradeCount     = $TotalTradeCount
            WinCount       = $TotalWinCount
            LossCount      = $TotalLossCount
            WinRate        = $TotalWinRatePercent
            LoseRate       = $TotalLossRatePercent
            Profit         = $TotalProfit
            AvgWeeklyProfit = $AvgWeeklyProfit
        }
    }
    END {
        If ($EOYObject) {
            Return $EOYObject
        }
        Else {
            Return Write-Error "Was unable to build end of week report"
        }
    }
} #END Function


Function Body-Parser {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$Body
    )
    BEGIN {
    }
   
    PROCESS {
        $Parse = $Body -split "Duration:"
        $Top = $Parse[0]
        $Parse = $Top -split "Order:"
        $Middle = $Parse[1]
        $FirstLine = ($middle -split '\r?\n')[1 - 1] 
        $Split = $FirstLine.split(' ')
        $Order = $Split[1]
        $Test = $Split[2]
        If ($Test -eq 'Short') {
            [string]$Type = "Short"
            [string]$Security = $Split[4]
        }
        elseif ($Test -eq 'to') {
            [string]$Type = "Short"
            [string]$Security = $Split[5]
        }
        else {
            [string]$Security = $Split[3]
            $OrderName = $Split[4]
            If ($OrderName -eq "@") {
                [string]$Type = "Shares"
            }
            Else {
                [string]$Expiration = $OrderName.Substring(0, 6)
                [string]$OrderType = $OrderName.Substring(6)
                If ($OrderType[0] -eq 'P') {
                    [string]$Type = "Put"
                }
                Else {
                    [string]$Type = "Call"
                }
                [string]$Strike = "$" + $OrderType.Substring(1)
            }
        }
        $QuantitySplit = ($middle -split '\r?\n')[2 - 1]
        [int]$QuantityInt = [int]$QuantitySplit.substring(13)
        $PriceSplit = ($middle -split '\r?\n')[3 - 1]
        [decimal]$PriceInt = [decimal]$PriceSplit.substring(15)
        [decimal]$Total = $PriceInt * $QuantityInt
  
        If ($Type -eq "Put" -or $Type -eq "Call") {
            $PriceInt += $PriceInt * 100
            $Total += $Total * 100
        }
  
        #End Parse and Add Variables to Object
        $TempObject = [PSCustomObject]@{
            Order        = $Order
            Type         = $Type
            Strike       = $Strike
            Expiration   = $Expiration
            Security     = $Security
            AvgPrice     = [math]::Round($PriceInt, 2)
            Quantity     = $QuantityInt
            Total        = [math]::Round($Total, 2)
            RunningTotal = 0
            RunningSold  = 0
            Profit       = ""
            TimeOpened   = $ReceivedTime
            TimeClosed   = ""
            TimeUpdated  = ""
            TimeSpan     = ""
            Status       = "Open"
        }
  
        If ($Order -eq 'Buy') {
            $TempObject.RunningTotal = [math]::Round($Total, 2)
        }
        Else {
            $TempObject.RunningSold = [math]::Round($Total, 2)
        }
  
    }
    END {
        If ($TempObject) {
            Return $TempObject
        }
        Else {
            Return Write-Error "Was unable to create order object"
        }
    }
} #END Function


Function Trade-Manager {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [object]$TradeObject,
        [Parameter(Mandatory = $true)]
        [datetime]$ReceivedTime,
        [Parameter(Mandatory = $true)]
        [string]$Action
    )
    BEGIN {
    }
   
    PROCESS {
        If ($Action -eq "Close") {
            #Close an Open Order
            If ($TradeObject.Type -eq 'Short') {
                [decimal]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").RunningTotal += [math]::Round($TradeObject.Total, 2)
            }
            else {
                [decimal]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").RunningSold += [math]::Round($TradeObject.Total, 2)
            }
            $RunningTotal = ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").RunningTotal
            $RunningSold = ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").RunningSold

            $Profit = [math]::Round($RunningSold, 2) - [math]::Round($RunningTotal, 2)
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").Profit = [math]::Round($Profit, 2)

            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "Open").TimeClosed = $ReceivedTime
            $StartTime = [datetime]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "Open").TimeOpened
            $EndTime = [datetime]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "Open").TimeClosed
            $Span = New-TimeSpan -start $StartTime -end $EndTime
            $TimeSpan = "{0:dd}d:{0:hh}h:{0:mm}m:{0:ss}s" -f $Span
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "Open").TimeSpan = $TimeSpan

            $PositionUpdate = $script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open"
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "Open").Status = "Closed"
            Return $PositionUpdate
        }
        elseif ($Action -eq "Add") {
            #Add to Position - Update existing order and add to Quantity, AvgPrice and Total
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").TimeUpdated = $ReceivedTime
            [int]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").Quantity += $TradeObject.Quantity
            [decimal]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").Total += [math]::Round($TradeObject.Total, 2)
            $AvgPrice = ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").Total / ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").Quantity
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").AvgPrice = [math]::Round($AvgPrice, 2)

            If ($TradeObject.Type -eq 'Short') {
                ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").RunningSold += [math]::Round($TradeObject.Total, 2)
            }
            else {
                ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open").RunningTotal += [math]::Round($TradeObject.Total, 2)
            }
            $PositionUpdate = $script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -eq $TradeObject.Order | Where-Object Status -eq "open"
            Return $PositionUpdate
        }
        elseif ($Action -eq "Subtract") {
            #Sell some of Position - Update existing order and subtract from Quantity, AvgPrice and Total
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").TimeUpdated = $ReceivedTime
            [int]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").Quantity -= $TradeObject.Quantity
            [decimal]($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").Total -= [math]::Round($TradeObject.Total, 2)
            $AvgPrice = ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").Total / ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").Quantity
            ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").AvgPrice = [math]::Round($AvgPrice, 2)
            Write-Host "Selling $($TradeObject.Total)"
            If ($TradeObject.Type -eq 'Short') {
                ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order ne $TradeObject.Order | Where-Object Status -eq "open").RunningTotal += [math]::Round($TradeObject.Total, 2)
            }
            else {
                ($script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open").RunningSold += [math]::Round($TradeObject.Total, 2)
            }
            $PositionUpdate = $script:TotalDailyTrades | Where-Object Security -eq $TradeObject.Security | Where-Object Type -eq $TradeObject.Type | Where-Object Order -ne $TradeObject.Order | Where-Object Status -eq "open"
            Return $PositionUpdate
        }
        elseif ($Action -eq "Different") {
            #Same Security but different order type (ex. Calls existed but then I bought shares of the same security)
            $script:TotalDailyTrades += $TradeObject
            Return $TradeObject  
        }
        else {
        }
    }
    END {
    }
} #END Function

