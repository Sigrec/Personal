[string]$VERSION = "3.0.0"

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
        [ValidateSet("HIDE", "PLACED", "ALLOCATING", "INVOICING", "PENDING PAYMENT", "PAID", "SHIPPING", "SHIPPED", "COMPLETE", IgnoreCase=$true)]
        [string]$Status,
        [Parameter(Mandatory=$false)]
        [Alias("p")]
        [string[]]$Product,
        [Parameter(Mandatory=$false)]
        [Alias("i")]
        [ValidateSet("PK", "MTG", "HL", "SCR", "GA", "SWU", "YGO", "LOR", "FAB", "DBS", "DM", "OP", "UA", "GCG", "IR", "Bandai", "Supplies", IgnoreCase=$true)]
        [string]$IP = "PK",
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
        [UInt128]$CaseAmount = 0
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
            $allResponses = @(
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_GID)",
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_PREORDER_GID)",
                "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_COMPLETE_GID)"
            )

            # Step 1: Fetch content from all sheets in parallel
            # Fetch and parse each sheet in parallel
            $parsedSheets = $allResponses | ForEach-Object -Parallel {
                try {
                    $response = Invoke-RestMethod -Uri $_ -ErrorAction Stop
                    $csvRows = $response.Content | ConvertFrom-Csv
                    # Return both the rows and their headers
                    return @{
                        Rows    = $csvRows
                        Headers = if ($csvRows) { $csvRows[0].PSObject.Properties.Name } else { @() }
                    }
                } catch {
                    Write-Debug "Error fetching: $_"
                    return $null
                }
            } -ThrottleLimit 3

            # Initialize header and row collections
            $allRows = [System.Collections.Generic.List[object]]::new()
            $headerSet = [System.Collections.Generic.HashSet[string]]::new()
            $headerTracker = [System.Collections.Generic.List[string]]::new()

            # Flatten results and track headers
            foreach ($sheet in $parsedSheets) {
                if ($null -ne $sheet) {
                    foreach ($header in $sheet.Headers) {
                        if ($headerSet.Add($header)) {
                            $headerTracker.Add($header)
                        }
                    }
                    $allRows.AddRange($sheet.Rows)
                }
            }

            # Normalize rows to ensure consistent headers
            $Response = foreach ($row in $allRows) {
                $newObj = [ordered]@{}
                foreach ($header in $headerTracker) {
                    $newObj[$header] = $row.PSObject.Properties[$header]?.Value
                }
                [PSCustomObject]$newObj
            }

            # Filter out null responses if there were any failed requests
            $Response = $Response | Where-Object { $_ -ne $null }

            # Filter contents of the array
            if (($null -ne $RowNum) -and ($RowNum -ne 0)) { # Get the specific row number
                $Response = $Response | Where-Object { $_."Row Number" -eq $RowNum }
            }
            else {
                if (-not [string]::IsNullOrWhiteSpace($Name)) {
                    $Response = $Response | Where-Object { $_."Name" -ilike "*$Name*" }
                }
                if ($Product.Count -ne 0) {
                    $Response = $Response | Where-Object {
                        $outerProduct = $_."Product Requested"
                        $Product | Where-Object { $outerProduct -ilike "*$_*" } | Measure-Object | Select-Object -ExpandProperty Count | Where-Object { $_ -gt 0 }
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($Status)) {
                    $Response = $Response | Where-Object { $_."Status" -eq $Status }
                }
                if ($Distro -ne 0) {
                    $Response = $Response | Where-Object { $_."Distro Number" -eq $Distro }
                }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Error "No order(s) found"
                return
            }

            # Calculate costs
            if (-not [string]::IsNullOrWhiteSpace($Name) -and ($Status -notmatch "HIDE") -and $Product.Count -eq 0) {
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
            elseif (($Status -notmatch "HIDE") -and ($Product.Count -ne 0) -and [string]::IsNullOrWhiteSpace($Name)) {
                # Initialize hashtable to hold aggregated results
                $aggregatedResultsList = @(
                    @{},
                    @{},
                    @{},
                    @{},
                    @{}
                )
                $totalProductCostList = [decimal[]]::new(5)

                # Iterate through each row in the $Response to manually accumulate the data
                # Write-Host $Response[0].psobject.Properties | ForEach-Object { $_.Name }
                foreach ($row in $Response) {
                    [string]$distro = $row."Distro Number"
                    [string]$product = $row."Product Requested"
                    [uint]$qtyReq = [uint]$row."Qty Req"
                    [decimal]$totalCost = [decimal]($row."Total Cost" -replace '[$,]', '')  # Clean the value

                    # Initialize product in hashtable if not already present
                    if ($distro -ne "Pokewholemart") {
                        $distroIndex = [int]($distro.TrimStart('#')) - 1  # Convert "#1" -> 0, "#2" -> 1, etc.
                    }
                    else {
                        continue
                    }

                    # Update total product cost list
                    $totalProductCostList[$distroIndex] += $totalCost

                    # Initialize product in the hashtable if not already present
                    if (-not $aggregatedResultsList[$distroIndex].ContainsKey($product)) {
                        $aggregatedResultsList[$distroIndex][$product] = @{
                            TotalQtyReq = 0
                            TotalCost   = 0.0
                        }
                    }

                    # Update the quantities and total costs
                    $aggregatedResultsList[$distroIndex][$product].TotalQtyReq += $qtyReq
                    $aggregatedResultsList[$distroIndex][$product].TotalCost += $totalCost
                }

                [decimal]$totalOveralSpend = 0.0
                for ($i = 0; $i -lt $aggregatedResultsList.Count; $i++) {
                    $aggregatedResultsDistro = $aggregatedResultsList[$i]
                    if ($aggregatedResultsDistro.Count -gt 0) {
                        $distro = $i + 1
                        $summary = $aggregatedResultsDistro.GetEnumerator() | ForEach-Object {
                            [PSCustomObject]@{
                                "Distro $distro Product Requested" = $_.Key
                                "Total Qty Req" = [math]::Round($_.Value.TotalQtyReq, 0)
                                "Total Cost" = [math]::Round($_.Value.TotalCost, 2).ToString("#,0.00")
                            }
                        }

                        # Output the summary as a formatted table to the console
                        $summary | Sort-Object "Distro $distro Product Requested" | Format-Table -Property "Distro $distro Product Requested", "Total Qty Req", @{Name="Total Cost";Expression={"$" + $_."Total Cost"}} -AutoSize

                        # Print the total spend for all products combined
                        [string]$totalSpend = "$" + $totalProductCostList[$i].ToString("#,0.00")
                        Write-Host "Distro $distro Total Spend: $totalSpend" -ForegroundColor Yellow
                        $totalOveralSpend += $totalProductCostList[$i]

                    }
                }

                [string]$totalSpend = "$" + $totalOveralSpend.ToString("#,0.00")
                Write-Host "`nTotal Spend: $totalSpend" -ForegroundColor Blue
            }

            $Response | Where-Object { $_."Product Requested" -notmatch "Quotes" } | Sort-Object -Property { [int]$_."Row Number" } | Out-GridView -Title $GRID_VIEW_TITLE
        }
        {$_ -in "ranking", "rank", "r"} {
            [UInt64]$RANKING_SHEET_GID = 781716676
            [string]$QUERY = "&gid=$($RANKING_SHEET_GID)"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-RestMethod -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv

            # Get the header if $IP is selected
            # Mapping abbreviations to full keys
            $ipMap = @{
                "MTG" = "Magic"
                "FAB" = "FaB"
                "GA"  = "Grand Archive"
                "LOR" = "Lorcana"
                "SRC" = "Sorcery"
                "SWU" = "Star Wars Unlimited"
                "YGO" = "Yu-Gi-Oh"
                "OP"  = "Distro 5 Bandai"
                "HL"  = "holoLive"
            }

            # Handle Bandai aliases
            if ($IP -match '^(Bandai|DB|DM|UA)$') {
                $key = "Bandai"
                Write-Debug "Processing Bandai product rankings..."
            }
            elseif ($ipMap.ContainsKey($IP)) {
                $key = $ipMap[$IP]
                Write-Debug "Processing $key product rankings..."
            }
            else {
                $key = "Pokemon"
                Write-Debug "Processing Pokémon product rankings..."
            }

            [decimal]$TotalSpend = 0
            $Response = $Response | ForEach-Object {
                $Spend = [decimal]($_.$key -replace '[$,]', '') # Strip $ and commas
                $TotalSpend += $Spend
                if ($key -eq "PK" -or $Spend -gt 0) {
                    [PSCustomObject]@{
                        "Rank" = 0
                        "User Name" = $_."User Name"
                        "Spend" = $_.$key
                    }
                }
            } | Where-Object { $_ -ne $null }
            
            if (-not [string]::IsNullOrWhiteSpace($Name)) {
                $Response = $Response | Where-Object { [String]$_."User Name" -match $Name }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Host "No rankings found"
            }
            elseif ([string]::IsNullOrWhiteSpace($Name)) {
                $rankCounter = 0
                $Response | 
                    Sort-Object -Property { [decimal]($_."Spend" -replace '[$,]', '') } -Descending | 
                    ForEach-Object {
                        $PercentSpend = [decimal]($_."Spend" -replace '[$,]', '') / $TotalSpend
                        $_ | Add-Member -MemberType NoteProperty -Name "Percent Spend" -Value "$(([math]::Round($PercentSpend * 100, 2)))%"
                        if ($CaseAmount -gt 0) {
                            $_ | Add-Member -MemberType NoteProperty -Name "Case Count" -Value ([math]::Round($PercentSpend * $CaseAmount, 2))
                        }
                        $_.Rank = ++$rankCounter  # Assign the rank value to the existing Rank property
                        $_  # Return the modified object
                    } | 
                    Out-GridView -Title "Rankings"
            }
            else {
                $Response| Format-Table -AutoSize -Wrap
            }
        }
        {$_ -in "payments", "pay"} {
            $WarningPreference = "SilentlyContinue"
            [UInt64]$PAYMENTS_SHEET_GID = 2061286159
            [string]$QUERY = "&gid=$($PAYMENTS_SHEET_GID)"
            [string]$GRID_VIEW_TITLE = "Payments Info"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-RestMethod -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv

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
            $Response = Invoke-RestMethod -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv
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
                "PK" {
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
                "MTG" {
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
                "FAB" {
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
                "GA" {
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
                "LOR" {
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
                "SCR" {
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
                "SWU" {
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
                "WS" {
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
                "YGO" {
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
                { @("Bandai", "DB", "DM", "OP", "UA", "GCG") -contains $_ }  {
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
                "HL" {
                    $FULL_IP = "holoLive"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 941844023
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
                "IR" {
                    $SHEET_GID = 1689199249
                }
                "Supplies" {
                    $SHEET_GID = 1234938269
                }
                default {
                    Write-Error "`"$IP`" is not a valid IP"
                    return
                }
            }

            if (@("DBS", "DM", "OP", "UA", "GCG") -match $IP) {
                Write-Debug "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)"
                Write-Debug "Filtering Bandai Product"
                switch($IP) {
                    "DB" {
                        $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Dragon Ball Super" }
                    }
                    "DM" {
                        $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Digimon" }
                    }
                    "OP" {
                        $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "One Piece" }
                    }
                    "UA" {
                        $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Union Arena" }
                    }
                    "GCG" {
                        $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS | Where-Object { $_."Product Name" -match "Gundam" }
                    }
                }
            }
            elseif (!$isPokemon) {
                Write-Debug "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)"
                $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS
            }
            else {
                Write-Debug "$($SHEET_URL)&range=$($SHEET_RANGE)"
                $Response = Invoke-RestMethod -Uri "$($SHEET_URL)&range=$($SHEET_RANGE)" | ConvertFrom-Csv -Header $HEADERS
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
            Write-Host "        -Status <String> [Optional] [Alias: s] [Case-Insensitive]"
            Write-Host "            Filters orders based on their status. Valid values are:"
            Write-Host "            PLACED, ALLOCATING, INVOICING, PENDING PAYMENT, PAID, SHIPPING, SHIPPED"
            Write-Host "        -Product <String> [Optional] [Alias: p]"
            Write-Host "            Filters orders based on the product requested."
            Write-Host "        -Distro <UInt16> [Optional] [Alias: d] (0)"
            Write-Host "            Filters the product based on the distributor. Valid values are 1, 2, 3, 4, 5."
            Write-Host "        -RowNum <UInt64> [Optional] [Alias: rn] (0)"
            Write-Host "            Filters orders by the row number."
            Write-Host ""

            Write-Host "    [`"ranking`", `"rank`", `"r`"] - Get rankings"
            Write-Host "        -Name <String> [Optional] [Alias: n]"
            Write-Host "            Specifies the discord member name to filter rankings."
            Write-Host "        -IP <String> [Optional] [Alias: i] [Case-Insensitive] (PK)"
            Write-Host "            Filters the product based on its intellectual property (IP)."
            Write-Host "            Valid IPs are  (`"PK`", `"MTG`", `"HL`", `"SCR`", `"GA`", `"SWU`", `"YGO`", `"LOR`", `"FAB`", `"DBS`", `"DM`", `"OP`", `"UA`", `"GCG`", `"IR`", `"Bandai`", `"Supplies`")"
            Write-Host "        -CaseAmount <UInt64> [Optional] [Alias: ca]"
            Write-Host "            The total amount of cases to determine estimated amount for a given user based off spend"
            Write-Host ""

            Write-Host "    [`"payments`", `"pay`"] - Get payments info"
            Write-Host "        -Name <String> [Optional] [Alias: n]"
            Write-Host "            Specifies the discord member name to filter payments."
            Write-Host "        -RowNum <UInt64> [Optional] [Alias: rn] (0)"
            Write-Host "            Filters payments by the row number."
            Write-Host ""

            Write-Host "    [`"overdue`", `"due`"] - Get list of members with payment 1 week overdue"
            Write-Host ""

            Write-Host "    [`"product`", `"p`"] - Get product information for a specific IP and distro"
            Write-Host "        -IP <String> [Optional] [Alias: i] [Case-Insensitive] (PK)"
            Write-Host "            Filters the product based on its intellectual property (IP)."
            Write-Host "            Valid IPs are  (`"PK`", `"MTG`", `"HL`", `"SCR`", `"GA`", `"SWU`", `"YGO`", `"LOR`", `"FAB`", `"DBS`", `"DM`", `"OP`", `"UA`", `"GCG`", `"IR`", `"Bandai`", `"Supplies`")"
            Write-Host "        -Distro <UInt16> [Optional] [Alias: d] (0)"
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