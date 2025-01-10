[string]$VERSION = "2.1.0"

function ptcg()
{
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        [Parameter(Mandatory=$false)]
        [Alias("n")]
        [string]$Name,
        [Parameter(Mandatory=$false)]
        [Alias("s")]
        [ValidateSet("HIDE", "PLACED", "ALLOCATING", "INVOICING", "PENDING PAYMENT", "PAID", "SHIPPING", "SHIPPED", "COMPLETE")]
        [string]$Status,
        [Parameter(Mandatory=$false)]
        [Alias("p")]
        [string]$Product,
        [Parameter(Mandatory=$false)]
        [Alias("i")]
        [string]$IP,
        [Parameter(Mandatory=$false)]
        [Alias("d")]
        [ValidateSet(1, 2, 3, 4, 5)]
        [UInt16]$Distro=0,
        [Parameter(Mandatory=$false)]
        [Alias("l")]
        [ValidateSet("English")]
        [string]$Lang="English",
        [Parameter(Mandatory=$false)]
        [Alias("rn")]
        [UInt64]$RowNum = 0,
        [Parameter(Mandatory=$false)]
        [Alias("ca")]
        [UInt64]$CaseAmount
    )
    [string]$MASTER_TRACKING_SHEET_URL = "https://docs.google.com/spreadsheets/d/1fWKRk_1i69rFE2ytxEmiAlHqYrPXVmhXSbG3fgGsl_I/export?format=csv"

    # TODO - Add new distro #5 to product command and finish cleaning it up
    [string]$Command = $Command.ToLower()
    Switch($Command) {
        {$_ -in "orders", "o"} {
            [UInt64]$MASTER_TRACKING_SHEET_GID = 1962881622
            [UInt64]$MASTER_TRACKING_SHEET_PREORDER_GID = 232688425
            [UInt64]$MASTER_TRACKING_SHEET_COMPLETE_GID = 2024850616
            [string]$GRID_VIEW_TITLE = "Order Info"

            # Initialize the array of URLs
            $allResponses = @()

            # Conditionally add the CANCELLED tracking sheet URL
            if (($Status -match "HIDE") -or (($null -ne $RowNum) -and ($RowNum -ne 0))) {
                [UInt64]$CANCELLED_TRACKING_SHEET_COMPLETE_GID = 1639309719
                $allResponses += "$($MASTER_TRACKING_SHEET_URL)&gid=$($CANCELLED_TRACKING_SHEET_COMPLETE_GID)"
            }

            # Add other necessary tracking sheet URLs
            $allResponses += @(
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_GID)",
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_PREORDER_GID)",
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_COMPLETE_GID)"
            )

            # Process URLs in parallel with throttle limit
            $Response = $allResponses | ForEach-Object -Parallel {
                Write-Debug "Request: $_"
                try {
                    $response = Invoke-WebRequest -Uri $_ -ErrorAction Stop
                    Write-Debug $response
                    return $response.Content | ConvertFrom-Csv
                } catch {
                    Write-Debug "Error fetching URL: $_, $_"
                    return $null  # Return $null if the request fails
                }
            } -ThrottleLimit 3

            # Filter out null responses if there were any failed requests
            $Response = $Response | Where-Object { $_ -ne $null }

            # Filter contents of the array
            if (($null -ne $RowNum) -and ($RowNum -ne 0)) {
                $Response = $Response | Where-Object { $_."Row Number" -eq $RowNum }
            }
            else {
                if (-not [string]::IsNullOrWhiteSpace($Name)) {
                    $Response = $Response | Where-Object { $_."Name" -match $Name }
                }
                if (-not [string]::IsNullOrWhiteSpace($Product)) {
                    $Response = $Response | Where-Object { $_."Product Requested" -match $Product }
                }
                if ((-not [string]::IsNullOrWhiteSpace($Status)) -and ($Status -notmatch "HIDE")) {
                    $Response = $Response | Where-Object { $_."Status" -match $Status }
                }
                if ($Distro -ne 0) {
                    $Response = $Response | Where-Object { $_."Distro Number" -match $Distro }
                }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Error "No order(s) found"
                return
            }

            # Calculate costs
            if (-not [string]::IsNullOrWhiteSpace($Name) -and ($Status -notmatch "HIDE")) {
                function Get-Amount {
                    param (
                        [string]$status = "",
                        [string]$distroAvailability = "",
                        [string]$columnName = "Total Cost"
                    )
                
                    # Helper function to calculate the rounded sum
                    function Get-Sum($response, $columnName) {
                        if ($response) {
                            return [Math]::Round(
                                ($response | ForEach-Object { 
                                    [decimal]($_.$columnName -replace "[$]", "") 
                                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum), 2
                            )
                        } else {
                            return 0
                        }
                    }
                
                    # Determine the filtered response based on parameters
                    $filteredResponse = if (-not [string]::IsNullOrWhiteSpace($status)) {
                        $Response | Where-Object { $_."Status" -eq $status }
                    } elseif (-not [string]::IsNullOrWhiteSpace($distroAvailability)) {
                        $Response | Where-Object { $_."Distro Availability" -eq $distroAvailability }
                    } else {
                        $Response
                    }
                
                    # Calculate and return the sum
                    return Get-Sum $filteredResponse $columnName
                }                
                
                # Array of statuses to process
                [string[]]$statuses = @("PLACED", "ALLOCATING", "INVOICING", "PENDING PAYMENT", "PAID", "SHIPPING", "SHIP PAY PENDING", "SHIPPED", "COMPLETE")
                
                # Loop through statuses and calculate spend for each
                $statuses | ForEach-Object {
                    try {
                        $spend = Get-Amount -status $_
                        Write-Host "$_ Spend: `$${spend}"
                    } catch {
                        Write-Host "Error processing status '$_': $_" -ForegroundColor Red
                    }
                }                

                $preOrderSpend = Get-Amount -distroAvailability "Pre-Order" -columnName "Total Cost"
                $openOrderSpend = Get-Amount -distroAvailability "Open Order" -columnName "Total Cost"
                $limitOrderSpend = Get-Amount -distroAvailability "Limit Order" -columnName "Total Cost"
                $backOrderSpend = Get-Amount -distroAvailability "Back Order" -columnName "Total Cost"

                Write-Host "`nOpen Order Spend: `$${openOrderSpend}"
                Write-Host "Pre-Order Spend: `$${preOrderSpend}"
                Write-Host "Limited Order Spend: `$${limitOrderSpend}"
                Write-Host "Back Order Spend: `$${backOrderSpend}"
                
                $curShippingCost = Get-Amount -status "SHIPPED" -columnName "Shipping Cost"
                $totalSpend = Get-Amount -columnName "Total Cost"
                $totalShippingCost = Get-Amount -columnName "Shipping Cost"
                $aggregateCost = [Math]::Round($totalSpend + $shippingCost, 2)

                Write-Host "`nCurrent Shipping Cost: `$${curShippingCost}"
                Write-Host "Total Spend: `$${totalSpend}"
                Write-Host "Total Shipping Cost: `$${totalShippingCost}"
                Write-Host "Aggregate Cost: `$${aggregateCost}"
            }

            $Response | Where-Object { $_."Product Requested" -notmatch "Quotes" } | Sort-Object -Property { [int]$_."Row Number" } | Out-GridView -Title $GRID_VIEW_TITLE
        }
        {$_ -in "ranking", "rank", "r"} {
            [UInt64]$RANKING_SHEET_GID = 781716676
            [string]$QUERY = "&gid=$($RANKING_SHEET_GID)"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv

            if (-not [string]::IsNullOrWhiteSpace($CaseAmount)) {
                if ($CaseAmount -gt 0) {
                    $Response = $Response | ForEach-Object {
                        $TotalSpend = ([double]($_."Total Spend" -replace '[$,]', '')) # Strip $ and commas
                    
                        if ($TotalSpend -ge 5000) {
                            [PSCustomObject]@{
                                "Rank" = $_."Rank"
                                "User Name" = $_."User Name"
                                "Estimated Case Count" = (([double]($_."Percent of total spend" -replace '[%,]', '')) / 100) * $CaseAmount
                            }
                        }
                    } | Where-Object { $_ -ne $null }
                }
                else {
                    Write-Error "CaseAmount (ca) must be greater than 0"
                }
            }
            
            if (-not [string]::IsNullOrWhiteSpace($Name)) {
                $Response = $Response | Where-Object { [String]$_."User Name" -match $Name }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Host "No rankings found"
            }
            elseif ([string]::IsNullOrEmpty($Name)) {
                $Response | Out-GridView -Title "Rankings"
            }
            else {
                $Response | Format-Table -AutoSize -Wrap
            }
        }
        {$_ -in "payments", "pay"} {
            $WarningPreference = "SilentlyContinue"
            [UInt64]$PAYMENTS_SHEET_GID = 2061286159
            [string]$QUERY = "&gid=$($PAYMENTS_SHEET_GID)"
            [string]$GRID_VIEW_TITLE = "Payments Info"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv

            if (($null -ne $RowNum) -and ($RowNum -ne 0)) {
                $Response = $Response | Where-Object { $_."Row Number" -eq $RowNum }
                if (-not $Response -or $Response.Count -eq 0) {
                    Write-Host "No payments found for row number $RowNum"
                    return
                }
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Name)) {
                $Response = $Response | Where-Object { $_."Name" -match $Name }
                if (-not $Response -or $Response.Count -eq 0) {
                    Write-Host "No payments found for `"$($Name)`""
                    return
                }

                # Calculate costs
                $totalCost = [Math]::Round($($Response | ForEach-Object {
                    [decimal]($_."Total Cost" -replace "[$]", "")
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum), 2)

                $shippingCost = [Math]::Round($($Response | ForEach-Object {
                    [decimal]($_."Shipping Cost" -replace "[$]", "")
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum), 2)

                $aggregateCost = [Math]::Round($totalCost + $shippingCost, 2)

                Write-Host "Total Cost: `$${totalCost}"
                Write-Host "Shipping Cost: `$${shippingCost}"
                Write-Host "Aggregate Cost: `$${aggregateCost}"
            }

            $Response | Select-Object -Property ( 
                $Response[0].PSObject.Properties.Name | Where-Object { $_ -notmatch "People with payment 1 week overdue|^H\d+$" }
            ) | Out-GridView -Title $GRID_VIEW_TITLE
        }
        {$_ -in "overdue", "due"} {
            $WarningPreference = "SilentlyContinue"
            [UInt64]$PAYMENTS_SHEET_GID = 2061286159
            [string]$QUERY = "&gid=$($PAYMENTS_SHEET_GID)"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv
            $Response | Select-Object -Property "People with payment 1 week overdue" | Where-Object { $_."People with payment 1 week overdue"-match '\S' } | Format-Table -AutoSize -Wrap
        }
        {$_ -in "product", "p"} {
            [UInt64]$SHEET_GID = 0
            [string]$SHEET_URL = "https://docs.google.com/spreadsheets/d/1Qj9aV8ae0MJ7MlBLIqYVh55B8_Ydprryy9zDwNoRJyU/export?format=csv"
            [string]$SHEET_RANGE = ""
            [string]$FULL_IP = ""
            [Boolean]$isPokemon = $false
            [char]$START_COLUMN
            [char]$END_COLUMN
            $HEADERS = @("Product Name", "Price", "Status", "Allocation Due", "Street Date")
            switch($IP) {
                { $_ -match "Pokemon" -or $_ -match "Pokémon" -or $_ -match "Poke" } {
                    $FULL_IP = "Pokémon"
                    $isPokemon = $true
                    $SHEET_URL = "https://docs.google.com/spreadsheets/d/1AnnzLYz1ktCLm0-Mt5o-6p4T8AqE0r2gewv-osqrK0A/export?format=csv"
                    Write-Debug "Getting $FULL_IP Product"
                    if($Lang -eq "English") {
                        $SHEET_GID = 0
                        switch($Distro) {
                            { $_ -eq 1 } {
                                $START_COLUMN = 'F'
                                $END_COLUMN = 'J'
                            }
                            { $_ -eq 2 } {
                                $START_COLUMN = 'L'
                                $END_COLUMN = 'P'
                            }
                            { $_ -eq 3 } {
                                $START_COLUMN = 'R'
                                $END_COLUMN = 'V'
                            }
                            { $_ -eq 4 } {
                                $START_COLUMN = 'W'
                                $END_COLUMN = 'Z'
                            }
                            { $_ -eq 5 } {
                                $START_COLUMN = 'AB'
                                $END_COLUMN = 'AE'
                            }
                            default {
                                Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                                return
                            }
                        }
                        $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                    }
                }
                { $_ -match "Magic The Gathering" -or $_ -match "MTG" -or $_ -match "Magic"  } {
                    $FULL_IP = "Magic The Gathering"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 419743007
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 2 } {
                            $START_COLUMN = 'F'
                            $END_COLUMN = 'J'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'K'
                            $END_COLUMN = 'O'
                        }
                        { $_ -eq 5 } {
                            $START_COLUMN = 'Q'
                            $END_COLUMN = 'U'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Flesh & Blood" -or $_ -match "Flesh And Blood" -or $_ -match "FAB" } {
                    $FULL_IP = "Flesh & Blood"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1539072415
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Grand Archive" -or $_ -match "GA" } {
                    $FULL_IP = "Grand Archive"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1217038948
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Lorcana" -or $_ -match "Lor" } {
                    $FULL_IP = "Lorcana"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 613517122
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)17:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Sorcery" -or $_ -match "Sorc" } {
                    $FULL_IP = "Sorcery"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 80389347
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Star Wars Unlimited" -or $_ -match "SWU" -or $_ -match "Star Wars" } {
                    $FULL_IP = "Star Wars Unlimited"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 879393505
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Weiss Schwarz" -or $_ -match "WS" -or $_ -match "Weiss" } {
                    $FULL_IP = "Weiss Schwarz"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1882644453
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        { $_ -eq 5 } {
                            $START_COLUMN = 'M'
                            $END_COLUMN = 'Q'
                        }
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Yu-Gi-Oh" -or $_ -match "YuGiOh" -or $_ -match "YGO" } {
                    $FULL_IP = "Yu-Gi-Oh"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 103569003
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 2 } {
                            $START_COLUMN = 'H'
                            $END_COLUMN = 'L'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'N'
                            $END_COLUMN = 'R'
                        }
                        { $_ -eq 4 } {
                            $START_COLUMN = 'T'
                            $END_COLUMN = 'X'
                        }
                        { $_ -eq 5 } {
                            $START_COLUMN = 'Z'
                            $END_COLUMN = 'AD'
                        }   
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { @("Bandai", "Dragon Ball Super", "DBS", "Dragon Ball", "Digimon", "Digi", "One Piece", "OP", "Union Arena", "UA", "Gundam", "GCG") -contains $_ }  {
                    $FULL_IP = "Bandai"
                    $HEADERS += "Product Info"
                    Write-Debug "Getting $FULL_IP Product for Distro #$Distro"
                    $SHEET_GID = 24121953
                    switch($Distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }  
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'L'
                        }  
                        { $_ -eq 5 } {
                            $START_COLUMN = 'M'
                            $END_COLUMN = 'Q'
                        }   
                        default {
                            Write-Error "Distro #$Distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -match "Item Request" -or $_ -match "Request" -or $_ -match "IR" } {
                    $SHEET_GID = 1689199249
                }
                { $_ -match "Supplies" -or $_ -match "Supply" } {
                    $SHEET_GID = 1234938269
                }
                default {
                    Write-Error "`"$IP`" is not a valid IP"
                    return
                }
            }

            if (@("Dragon Ball Super", "DBS", "Digimon", "Digi", "One Piece", "OP", "Union Arena", "UA", "Gundam", "GCG") -match $IP) {
                Write-Debug "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)"
                Write-Debug "Filtering Bandai Product"
                switch($IP) {
                    { $_ -match "Dragon Ball Super" -or $_ -match "DBS" -or $_ -match "Dragon Ball"} {
                        $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Dragon Ball Super" }
                    }
                    { $_ -match "Digimon" -or $_ -match "Digi" } {
                        $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Digimon" }
                    }
                    { $_ -match "One Piece" -or $_ -match "OP" } {
                        $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "One Piece" }
                    }
                    { $_ -match "Union Arena" -or $_ -match "UA" } {
                        $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Union Arena" }
                    }
                    { $_ -match "Gundam" -or $_ -match "GCG" } {
                        $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Gundam" }
                    }
                }
            }
            elseif (!$isPokemon) {
                Write-Debug "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)"
                $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS
            }
            else {
                Write-Debug "$($SHEET_URL)&range=$($SHEET_RANGE)"
                $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS
            }

            $Response = $Response | Where-Object { ![string]::IsNullOrWhiteSpace($_."Product Name") }
            if (-not $Response -or $Response.Count -eq 0) {
                Write-Error "No product found for `"$IP`" at Distro #$Distro"
                return
            }
            $Response | Out-GridView -Title "Product"
        }
        {$_ -in "sheets"} {
            Start-Process "https://docs.google.com/spreadsheets/d/1fWKRk_1i69rFE2ytxEmiAlHqYrPXVmhXSbG3fgGsl_I/edit?gid=1962881622#gid=1962881622"
            Start-Process "https://docs.google.com/spreadsheets/d/1AnnzLYz1ktCLm0-Mt5o-6p4T8AqE0r2gewv-osqrK0A/edit?gid=0#gid=0"
            Start-Process "https://docs.google.com/spreadsheets/d/1Qj9aV8ae0MJ7MlBLIqYVh55B8_Ydprryy9zDwNoRJyU/edit?gid=419743007#gid=419743007"
        }
        {$_ -in "faq"} {
            Start-Process "https://docs.google.com/document/d/1K3hmfo1EzLazjQz2-_zFdsqjz-NQnz7POPyAORxO_Wo/edit?tab=t.0"
        }
        {$_ -in "distro"} {
            Start-Process "https://www.southernhobby.com/"
            Start-Process "https://magazine-exchange.com/"
            Start-Process "https://portal.phdgames.com/products?p=preordersdue&page=1&size=20"
            Start-Process "https://madal.com/"
            Start-Process "https://www.gtsdistribution.com/"
        }
        {$_ -in "help", "h"} {
            Write-Host "AUTHOR"
            Write-Host "    Prem (prem8)"
            ""
            Write-Host "VERSION"
            Write-Host "    $VERSION"
            ""
            Write-Host "EXAMPLES"
            Write-Host "    ptcg o -Name `"Eli`" (Prints all orders from the spreadsheet where the name contains `"Eli`")"
            ""
            Write-Host "SYNTAX (Order doesn't mater)"
            Write-Host "    ptcg <Command> [-Parameter(s)] <Value(s)>"
            ""
            Write-Host "COMMANDS"
            Write-Host "    [`"orders`", `"o`"] - Get Master Tracking order info"
            Write-Host "        -Name <String> [Optional] [Alias: n]"
            Write-Host "            Specifies the discord member name to filter orders."
            Write-Host "        -Status <String> [Optional] [Alias: s]"
            Write-Host "            Filters orders based on their status. Valid values are:"
            Write-Host "            PLACED, ALLOCATING, INVOICING, PENDING PAYMENT, PAID, SHIPPING, SHIPPED"
            Write-Host "        -Product <String> [Optional] [Alias: p]"
            Write-Host "            Filters orders based on the product requested."
            Write-Host "        -Distro <UInt16> [Optional] [Alias: d]"
            Write-Host "            Filters the product based on the distributor. Valid values are 1, 2, 3, 4, 5."
            Write-Host "        -RowNum <UInt64> [Optional] [Alias: rn]"
            Write-Host "            Filters orders by the row number."
            Write-Host ""

            Write-Host "    [`"ranking`", `"rank`", `"r`"] - Get rankings"
            Write-Host "        -Name <String> [Optional] [Alias: n]"
            Write-Host "            Specifies the discord member name to filter rankings."
            Write-Host "        -CaseAmount <UInt64> [Optional] [Alias: ca]"
            Write-Host "            The total amount of cases to determine estimated amount for a given user based off spend"
            Write-Host ""

            Write-Host "    [`"payments`", `"pay`"] - Get payments info"
            Write-Host "        -Name <String> [Optional] [Alias: n]"
            Write-Host "            Specifies the discord member name to filter payments."
            Write-Host "        -RowNum <UInt64> [Optional] [Alias: rn]"
            Write-Host "            Filters payments by the row number."
            Write-Host ""

            Write-Host "    [`"overdue`", `"due`"] - Get list of members with payment 1 week overdue"
            Write-Host ""

            Write-Host "    [`"product`", `"p`"] - Get product information for a specific IP and distro"
            Write-Host "        -IP <String> [Optional] [Alias: i]"
            Write-Host "            Filters the product based on its intellectual property (IP)."
            Write-Host "        -Distro <UInt16> [Optional] [Alias: d]"
            Write-Host "            Filters the product based on the distributor. Valid values are 1, 2, 3, 4, 5."
            Write-Host ""

            Write-Host "    [`"sheets`"] - Open the wholesale program google sheets"
            Write-Host ""

            Write-Host "    [`"faq`"] - Open the wholesale program faq google doc"
            Write-Host ""

            Write-Host "    [`"distro`"] - Open all of the distro websites in order"
            Write-Host ""

            Write-Host "    [`"help`", `"h`"] - Displays information about the CLI"
            ""
        }
    }
}