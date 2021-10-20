Function Order-Message {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [object]$TradeObject
  )
  BEGIN {
    $TempTotal = '$' + $TradeObject.Total.ToString()
    $TempAvgPrice = '$' + $TradeObject.AvgPrice.ToString()
  }
 
  PROCESS {
    $green = "4289797"
    $red = "16711680"
    if ($Order -eq 'Buy') {
      $color = $green
    }
    else {
      $color = $red
    }

    $Content = @"
        {
            "embeds": [
              {
                "title": "$Order Order",
                "description": "$($TradeObject.TimeOpened)",
                "color": "$color",
                "fields": [
                  {
                    "name": "Security",
                    "value": "$($TradeObject.Security)",
                    "inline": true
                  },
                  {
                    "name": "Type",
                    "value": "$($TradeObject.Type)",
                    "inline": true
                  },

"@
    If ($($TradeObject.Type) -eq 'Call' -or $($TradeObject.Type) -eq 'Put') {
      $Content += @"
                  {
                    "name": "Strike",
                    "value": "$($TradeObject.Strike)",
                    "inline": true
                  },
                  {
                    "name": "Expiration",
                    "value": "$($TradeObject.Expiration)",
                    "inline": true
                  },

"@
    }
    $Content += @"
                  {
                    "name": "Quantity",
                    "value": "$($TradeObject.Quantity)",
                    "inline": true
                  },
                  {
                    "name": "Average Price",
                    "value": "$TempAvgPrice",
                    "inline": true
                  },
                  {
                    "name": "Total",
                    "value": "$TempTotal",
                    "inline": true
                  }
                ],
                "footer": {
                  "text": "This is Chris`'s personal trade. This should not be considered a suggestion to buy or sell any security. Trade at your own risk."
                }
              }
            ]
          }
"@

    Invoke-RestMethod -Uri $webHookUrl -Body $content -Method Post -ContentType 'application/json'
  }
  END {

  }
} #END Function

Function Position-Update {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [object]$TradeObject,
    [string]$Color,
    [string]$Order
  )
  BEGIN {
    $TempTotal = '$' + $TradeObject.Total.ToString()
    $TempAvgPrice = '$' + $TradeObject.AvgPrice.ToString()
  }
 
  PROCESS {
    $Content = @"
        {
            "embeds": [
              {
                "title": "Position Update - $Order",
                "description": "$($TradeObject.TimeUpdated)",
                "color": "$color",
                "fields": [
                  {
                    "name": "Security",
                    "value": "$($TradeObject.Security)",
                    "inline": true
                  },
                  {
                    "name": "Type",
                    "value": "$($TradeObject.Type)",
                    "inline": true
                  },

"@
    If ($($TradeObject.Type) -eq 'Call' -or $($TradeObject.Type) -eq 'Put') {
      $Content += @"
                  {
                    "name": "Strike",
                    "value": "$($TradeObject.Strike)",
                    "inline": true
                  },
                  {
                    "name": "Expiration",
                    "value": "$($TradeObject.Expiration)",
                    "inline": true
                  },

"@
    }
    $Content += @"
                  {
                    "name": "Current Position Quantity",
                    "value": "$($TradeObject.Quantity)",
                    "inline": true
                  },
                  {
                    "name": "Current Average Price",
                    "value": "$TempAvgPrice",
                    "inline": true
                  },
                  {
                    "name": "Current Position Total",
                    "value": "$TempTotal",
                    "inline": true
                  }
                ],
                "footer": {
                  "text": "This is Chris`'s personal trade. This should not be considered a suggestion to buy or sell any security. Trade at your own risk."
                }
              }
            ]
          }
"@

    Invoke-RestMethod -Uri $webHookUrl -Body $content -Method Post -ContentType 'application/json'
    
  }
  END {

  }
} #END Function

Function Close-Message {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [object]$TradeObject
  )
  BEGIN {
    $TempTotal = '$' + $TradeObject.Total.ToString()
    $TempRunningTotal = '$' + $TradeObject.RunningTotal.ToString()
    $TempRunningSold = '$' + $TradeObject.RunningSold.ToString()
    If ($TradeObject.Profit -lt 0) {
      $InvertCalc = $TradeObject.Profit * -1
      $TempProfit = '-$' + $InvertCalc.ToString()
    }
    else {
      $TempProfit = '$' + $TradeObject.Profit.ToString()
    }
        
  }
 
  PROCESS {
    $green = "4289797"
    $red = "16711680"
    if ($Profit -ge 0) {
      $color = $green
    }
    else {
      $color = $red
    }

    $Content = @"
        {
            "embeds": [
              {
                "title": "Closed Position",
                "description": "$($TradeObject.TimeClosed)",
                "color": "$color",
                "fields": [
                  {
                    "name": "Security",
                    "value": "$($TradeObject.Security)",
                    "inline": true
                  },
                  {
                    "name": "Type",
                    "value": "$($TradeObject.Type)",
                    "inline": true
                  },

"@
    If ($($TradeObject.Type) -eq 'Call' -or $($TradeObject.Type) -eq 'Put') {
      $Content += @"
                  {
                    "name": "Strike",
                    "value": "$($TradeObject.Strike)",
                    "inline": true
                  },
                  {
                    "name": "Expiration",
                    "value": "$($TradeObject.Expiration)",
                    "inline": true
                  },

"@
    }
    $Content += @"
                  {
                    "name": "Trade Quantity",
                    "value": "$($TradeObject.Quantity)",
                    "inline": true
                  },
                  {
                    "name": "Trade Total",
                    "value": "$TempTotal",
                    "inline": true
                  },
                  {
                    "name": "Total Bought",
                    "value": "$TempRunningTotal",
                    "inline": true
                  },
                  {
                    "name": "Total Sold",
                    "value": "$TempRunningSold",
                    "inline": true
                  },
                  {
                    "name": "Total Profit/Loss",
                    "value": "$TempProfit",
                    "inline": true
                  },
                  {
                    "name": "Approx. Time in Trade",
                    "value": "$($TradeObject.TimeSpan)",
                    "inline": true
                  }
                ],
                "footer": {
                  "text": "This is Chris`'s personal trade. This should not be considered a suggestion to buy or sell any security. Trade at your own risk."
                }
              }
            ]
          }
"@

    Invoke-RestMethod -Uri $webHookUrl -Body $content -Method Post -ContentType 'application/json'
    
  }
  END {

  }
} #END Function

Function EOD-Report {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [object]$TradeObject
  )
  BEGIN {
    $Date = Get-Date -DisplayHint Date
    If ($TradeObject.Profit -lt 0) {
      $InvertCalc = $TradeObject.Profit * -1
      $TempProfit = '-$' + $InvertCalc.ToString()
    }
    else {
      $TempProfit = '$' + $TradeObject.Profit.ToString()
    }
        
  }
 
  PROCESS {
    $Content = @"
        {
            "embeds": [
              {
                "title": "EOD Report",
                "description": "$Date",
                "color": "65535",
                "fields": [
                  {
                    "name": "Trade Count",
                    "value": "$($TradeObject.TradeCount)",
                    "inline": true
                  },
                  {
                    "name": "Win Count",
                    "value": "$($TradeObject.WinCount)",
                    "inline": true
                  },
                  {
                    "name": "Loss Count",
                    "value": "$($TradeObject.LossCount)",
                    "inline": true
                  },
                  {
                    "name": "Win Rate",
                    "value": "$($TradeObject.WinRate)",
                    "inline": true
                  },
                  {
                    "name": "Lose Rate",
                    "value": "$($TradeObject.LoseRate)",
                    "inline": true
                  },
                  {
                    "name": "Total Profit/Loss",
                    "value": "$TempProfit",
                    "inline": true
                  },
                  {
                    "name": "Securities Traded",
                    "value": "$($TradeObject.Securities)",
                    "inline": true
                  },
                  {
                    "name": "Profit/Loss",
                    "value": "$($TradeObject.Profits)",
                    "inline": true
                  },
                  {
                    "name": "Trade Type",
                    "value": "$($TradeObject.Types)",
                    "inline": true
                  }
"@

                  If ($($TradeObject.OpenSecurities)) {
                  $Content += @"
                  ,
                  {
                    "name": "Open Positions",
                    "value": "$($TradeObject.OpenSecurities)",
                    "inline": true
                  },
                  {
                    "name": "Open Quantity",
                    "value": "$($TradeObject.OpenQuantity)",
                    "inline": true
                  },
                  {
                    "name": "Open Avg Price",
                    "value": "$($TradeObject.OpenAvgPrice)",
                    "inline": true
                  }

"@
                }
                $Content += @"
                ],
                "footer": {
                    "text": "This is Chris`'s personal trade. This should not be considered a suggestion to buy or sell any security. Trade at your own risk."
                }
            }
        ]
    }
"@
    Invoke-RestMethod -Uri $webHookUrl -Body $content -Method Post -ContentType 'application/json'
    
  }
  END {

  }
} #END Function

Function EOW-Report {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [object]$TradeObject
  )
  BEGIN {
    $Date = Get-Date -DisplayHint Date
    If ($TradeObject.Profit -lt 0) {
      $InvertCalc = $TradeObject.Profit * -1
      $TempProfit = '-$' + $InvertCalc.ToString()
    }
    else {
      $TempProfit = '$' + $TradeObject.Profit.ToString()
    }
        
  }
 
  PROCESS {
    $Content = @"
        {
            "embeds": [
              {
                "title": "EOW Report",
                "description": "$Date",
                "color": "65535",
                "fields": [
                  {
                    "name": "Trade Count",
                    "value": "$($TradeObject.TradeCount)",
                    "inline": true
                  },
                  {
                    "name": "Win Count",
                    "value": "$($TradeObject.WinCount)",
                    "inline": true
                  },
                  {
                    "name": "Loss Count",
                    "value": "$($TradeObject.LossCount)",
                    "inline": true
                  },
                  {
                    "name": "Win Rate",
                    "value": "$($TradeObject.WinRate)",
                    "inline": true
                  },
                  {
                    "name": "Lose Rate",
                    "value": "$($TradeObject.LoseRate)",
                    "inline": true
                  },
                  {
                    "name": "Total Profit/Loss",
                    "value": "$TempProfit",
                    "inline": true
                  }
                ],
                "footer": {
                    "text": "This is Chris`'s personal trade. This should not be considered a suggestion to buy or sell any security. Trade at your own risk."
                }
            }
        ]
    }
"@
    Invoke-RestMethod -Uri $webHookUrl -Body $content -Method Post -ContentType 'application/json'
    
  }
  END {

  }
} #END Function

Function Live-Message {
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory = $true)]
    [string]$Message
  )
  BEGIN {
  }
 
  PROCESS {
    $payload = [PSCustomObject]@{

      content = $message
        
    }
        
    Invoke-RestMethod -Uri $webHookUrl -Body ($payload | ConvertTo-Json) -Method Post -ContentType 'application/json'
    
  }
  END {

  }
} #END Function
