[string]$VERSION = "1.0.0"

# TODO - Add new param to search for a specific order #
function ptcg()
{
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,
        [Parameter(Mandatory=$false)]
        [string]$Name,
        [Parameter(Mandatory=$false)]
        [string]$Status,
        [Parameter(Mandatory=$false)]
        [string]$Product,
        [Parameter(Mandatory=$false)]
        [string]$IP,
        [Parameter(Mandatory=$false)]
        [Alias("d")]
        [ValidateSet(1, 2, 3, 4, 5)]
        [UInt16]$distro,
        [Parameter(Mandatory=$false)]
        [string]$lang="English"
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

            if ($Status -match "HIDE") {
                [UInt64]$CANCELLED_TRACKING_SHEET_COMPLETE_GID = 1639309719
                $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)&gid=$($CANCELLED_TRACKING_SHEET_COMPLETE_GID )" | ConvertFrom-Csv
            }
            else {
                # Fetch URLs in parallel and combine results
                $allResponses = @(
                    "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_GID)",
                    "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_PREORDER_GID)",
                    "$($MASTER_TRACKING_SHEET_URL)&gid=$($MASTER_TRACKING_SHEET_COMPLETE_GID)"
                ) | ForEach-Object -Parallel {
                    Write-Debug "Request: $($_)"
                    $response = Invoke-WebRequest -Uri $_
                    $response.Content | ConvertFrom-Csv
                } -ThrottleLimit 3

                # Combine all data into a single array
                $Response = $allResponses + @()
            }

            # Filter contents of the array
            if (![string]::IsNullOrEmpty($Name)) {
                $Response = $Response | Where-Object { $_."Name" -match $Name }
            }
            if (![string]::IsNullOrEmpty($Product)) {
                $Response = $Response | Where-Object { $_."Product Requested" -match $product }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Error "No order(s) found"
                return
            }

            if (![string]::IsNullOrEmpty($Name) -and ($Status -notmatch "HIDE")) {
                # Calculate costs
                $totalCost = [Math]::Round($($Response | ForEach-Object {
                    [decimal]($_."Total Cost" -replace "[$]", "")
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum), 2)

                $shippingCost = [Math]::Round($($Response | ForEach-Object {
                    [decimal]($_."Shipping Cost" -replace "[$]", "")
                } | Measure-Object -Sum | Select-Object -ExpandProperty Sum), 2)

                $aggregateCost = [Math]::Round($totalCost + $shippingCost, 2)

                Write-Output "Total Cost: `$${totalCost}"
                Write-Output "Shipping Cost: `$${shippingCost}"
                Write-Output "Aggregate Cost: `$${aggregateCost}"
            }

            $Response | Where-Object { $_."Product Requested" -notmatch "Quotes" } | Sort-Object -Property { [int]$_."Row Number" } | Out-GridView -Title $GRID_VIEW_TITLE
        }
        {$_ -in "ranking", "rank", "r"} {
            [UInt64]$RANKING_SHEET_GID = 781716676
            [string]$QUERY = "&gid=$($RANKING_SHEET_GID)"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv
            
            if (![string]::IsNullOrEmpty($Name)) {
              $Response = $Response | Where-Object { $_."User Name" -match $Name }
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Host "No ranking found for `"$($Name)`""
            }
            elseif ([string]::IsNullOrEmpty($Name)) {
                $Response | Out-GridView -Title "Rankings"
            }
            else {
                $Response | Format-Table -AutoSize -Wrap
            }
        }
        {$_ -in "payments", "pay"} {
            if ([string]::IsNullOrEmpty($Name)) {
                Write-Error "'Name' param cannot be empty!"
                break
            } 

            $WarningPreference = "SilentlyContinue"
            [UInt64]$PAYMENTS_SHEET_GID = 2061286159
            [string]$QUERY = "&gid=$($PAYMENTS_SHEET_GID)"
            [string]$GRID_VIEW_TITLE = "Payments Info"

            Write-Debug "Request: $($MASTER_TRACKING_SHEET_URL)$($QUERY)"
            $Response = Invoke-WebRequest -Uri "$($MASTER_TRACKING_SHEET_URL)$($QUERY)" | ConvertFrom-Csv | Where-Object { $_."Name" -match $Name }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Host "No payments found for `"$($Name)`""
            }
            else {
                $Response | Select-Object -Property ( 
                    $Response[0].PSObject.Properties.Name | Where-Object { $_ -notmatch "People with payment 1 week overdue|^H\d+$" }
                ) | Out-GridView -Title $GRID_VIEW_TITLE
            }
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

            switch($IP) {
                { $_ -like "Pokemon" -or $_ -like "Pokémon" -or $_ -like "Poke" } {
                    # TODOD - Finish pokemon
                    $FULL_IP = "Pokémon"
                    $isPokemon = $true
                    $SHEET_URL = "https://docs.google.com/spreadsheets/d/1AnnzLYz1ktCLm0-Mt5o-6p4T8AqE0r2gewv-osqrK0A/export?format=csv"
                    Write-Debug "Getting $FULL_IP Product"
                    if($lang -eq "English") {
                        $SHEET_GID = 0
                        switch($distro) {
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
                            default {
                                Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                                return
                            }
                        }
                        $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                    }
                    else {
                        $SHEET_GID = 2025586520
                    }
                }
                { $_ -like "Magic The Gathering" -or $_ -like "MTG" -or $_ -like "Magic"  } {
                    FULL_IP = "Magic The Gathering"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 419743007
                    switch($distro) {
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
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Flesh & Blood" -or $_ -like "Flesh And Blood" -or $_ -like "FAB"} {
                    $FULL_IP = "Flesh & Blood"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1539072415
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Grand Archive" -or $_ -like "GA"} {
                    $FULL_IP = "Grand Archive"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1217038948
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Lorcana" -or $_ -like "Lor"} {
                    $FULL_IP = "Lorcana"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 613517122
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)16:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Sorcery" -or $_ -like "Sorc"} {
                    $FULL_IP = "Sorcery"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 80389347
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Star Wars Unlimited" -or $_ -like "SWU" -or $_ -like "Star Wars"} {
                    $FULL_IP = "Star Wars Unlimited"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 879393505
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Union Arena" -or $_ -like "UA"} {
                    $FULL_IP = "Union Arena"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 24121953
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Weiss Schwarz" -or $_ -like "WS" -or $_ -like "Weiss"} {
                    $FULL_IP = "Weiss Schwarz"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 1882644453
                    switch($distro) {
                        { $_ -eq 1 } {
                            $START_COLUMN = 'A'
                            $END_COLUMN = 'E'
                        }
                        { $_ -eq 3 } {
                            $START_COLUMN = 'G'
                            $END_COLUMN = 'K'
                        }
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Yu-Gi-Oh" -or $_ -like "YuGiOh" -or $_ -like "YGO"} {
                    $FULL_IP = "Yu-Gi-Oh"
                    Write-Debug "Getting $FULL_IP Product"
                    $SHEET_GID = 103569003
                    switch($distro) {
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
                        default {
                            Write-Error "Distro #$distro does not have `"$FULL_IP`" product"
                            return
                        }
                    }
                    $SHEET_RANGE = "$($START_COLUMN)15:$($END_COLUMN)&tq=SELECT%20*"
                }
                { $_ -like "Item Request" -or $_ -like "Request" -or $_ -like "IR"} {
                    $SHEET_GID = 1689199249
                }
                default {
                    Write-Error "`"$IP`" is not a valid IP"
                    return
                }
            }

            if (!$isPokemon) {
                $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&gid=$($SHEET_GID)&range=$($SHEET_RANGE)" | ConvertFrom-Csv
            }
            else {
                $Response = Invoke-WebRequest -Uri "$($SHEET_URL)&range=$($SHEET_RANGE)" | ConvertFrom-Csv
            }

            if (-not $Response -or $Response.Count -eq 0) {
                Write-Error "No product found for `"$IP`" at Distro #$distro"
                return
            }
            $Response | Where-Object { ![string]::IsNullOrWhiteSpace($_."Name") } | Out-GridView -Title "$IP Product"
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
            Write-Host "    ptcg user -Name `"Eli`" (Prints all orders from the spreadsheet where the name contains `"Eli`")"
            ""
            Write-Host "SYNTAX (Order doesn't mater)"
            Write-Host "    ptcg <Command> [-Parameter(s)] <Value(s)>"
            ""
            Write-Host "COMMANDS"
            Write-Host "    [`"orders`", `"o`"] - Lists the Master Tracking order info for a member"
            Write-Host "    [`"ranking`", `"rank`", `"r`"] - Get current rankings for a specifc member"
            Write-Host "    [`"payments`", `"p`"] - Get payments info for a specific member"
            Write-Host "    [`"help`", `"h`"] - Displays information about the CLI"
            ""
            Write-Host "PARAMETERS"
            Write-Host "    -Name <String> [CaseSensitive]"
            Write-Host "        Specifies the discord member name in the sheet"
            ""
        }
    }
}