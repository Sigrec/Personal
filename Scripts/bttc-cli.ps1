$VERSION = "v1.2.0"

Function bttc()
{
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Command,

        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.List[string]]$Branch,

        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.List[string]]$Artisan,

        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.List[string]]$Archetype,
        
        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.List[string]]$Timezone,

        [Parameter(Mandatory=$false)]
        [System.Collections.Generic.List[string]]$RP
    )

    # Constants
    $URL = "https://docs.google.com/spreadsheets/d/1BPuFezUHKC1mJduXt0SiY3d8f6r98fEVczg9PaCQ14g/gviz/tq?tqx=out:csv&headers=0"
    $START_INDEX = 2
    $END_INDEX = 503
    $GENERIC_SHEET = "Generic Info"
    $ARTISAN_SHEET = "Artisans"
    $ARCHETYPE_SHEET = "Archetypes"

    $Command = $Command.ToLower()
    Switch($Command) {
        {$_ -in "list", "ls", "members"} {
            $ResponseList = [System.Collections.Generic.List[PSCustomObject]]::new()
            $PropertyInputs = [System.Collections.Generic.List[string]]::new()
            $InputList = [System.Collections.Generic.List[string]]::new()

            if (!$Branch -and !$Artisan -and !$Archetype -and !$RP) {
                Write-Debug "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CC%2CF"
                $Response = Invoke-WebRequest -Uri "$($URL)&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CC%2CF" | ConvertFrom-Csv
                if (![string]::IsNullOrEmpty($Timezone)) {
                    $Timezone = Get-Timezone($Timezone)
                    if ($Timezone -eq "Error") { break; }
                    $Response = $Response | Where-Object { $_.Timezone.ToUpper() -eq $Timezone } | Select-Object -Property * -ExcludeProperty Timezone
                }
                $Response | Sort-Object -Property "Discord Name" | Format-Table -AutoSize
                break
            }

            [bool]$HasTimezone = $false
            if (![string]::IsNullOrEmpty($Timezone)) { 
                $HasTimezone = $true
                $Timezone = Get-Timezone($Timezone)
                if ($Timezone -eq "Error") { break; }
                $InputList.Add("($($Timezone -Join ' | '))")
                $TIMEZONE_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CF%20WHERE%20$(Get-Query $Timezone "F")"
            }
            else {
                $TIMEZONE_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CF"
            }
            $PropertyInputs.Add("Timezone")
            Write-Debug "Timezone Query: $TIMEZONE_QUERY"
            $TIMEZONE_RESPONSE = Invoke-WebRequest -Uri "$($URL)$($TIMEZONE_QUERY)" | ConvertFrom-Csv -Header "Discord Name", "Character Name", "Timezone"
            $ResponseList.Add($TIMEZONE_RESPONSE)
            
            if (![string]::IsNullOrEmpty($Branch)) { 
                $Branch = Get-Branch($Branch)
                if ($Branch -eq "Error") { break }

                [string]$BRANCH_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):C$($END_INDEX)&tq=SELECT%20A%2CB%2CC%20WHERE%20$(Get-Query $Branch "C")"
                Write-Debug "Branch Query: $BRANCH_QUERY"

                $Header = @()
                if ($Branch.Count -gt 1) {
                    $PropertyInputs.Add("Branch")
                    $InputList.Add("($($Branch -Join ' | '))")
                    $Header =  "Discord Name", "Character Name", "Branch"
                }
                else {
                    $Header = "Discord Name", "Character Name"
                }
                $Branch_Response = Invoke-WebRequest -Uri "$($URL)$($BRANCH_QUERY)" | ConvertFrom-Csv  -Header $Header
                $ResponseList.Add($Branch_Response)
            }

            if (![string]::IsNullOrEmpty($Artisan)) { 
                for ($i = 0; $i -lt $Artisan.Count; $i++) {
                    $Artisan[$i] = Get-Artisan($Artisan[$i])
                }
                if ($Artisan -eq "Error") { break }

                [string]$ARTISAN_QUERY = "&sheet=$ARTISAN_SHEET&range=A$($START_INDEX):D$($END_INDEX)&tq=SELECT%20A%2CB%2CC%2CD%20WHERE%20"
                foreach ($CurArtisan in $Artisan) {
                    $ARTISAN_QUERY += "(C%20contains%20'$CurArtisan')%20OR%20(D%20contains%20'$CurArtisan')%20OR%20"
                }
                $ARTISAN_QUERY = $ARTISAN_QUERY.TrimEnd("%20OR%20").Trim()
                
                $Header = @()
                if ($Artisan.Count -gt 1) {
                    $PropertyInputs.Add("Grandmaster 1")
                    $PropertyInputs.Add("Grandmaster 2")
                    $InputList.Add("($($Artisan -Join ' | '))")
                    $Header =  "Discord Name", "Character Name", "Grandmaster 1", "Grandmaster 2"
                }
                else {
                    $Header = "Discord Name", "Character Name"
                }

                Write-Debug "Artisan Query: $ARTISAN_QUERY"
                $ArtisanResponse = Invoke-WebRequest -Uri "$($URL)$($ARTISAN_QUERY)" | ConvertFrom-Csv -Header $Header
                $ResponseList.Add($ArtisanResponse)
            }

            if (![string]::IsNullOrEmpty($Archetype)) {
                $Archetype = Get-Archetype($Archetype)
                if ($Archetype -eq "Error") { break }
                
                [string]$ARCHETYPE_QUERY = "&sheet=$ARCHETYPE_SHEET&range=A$($START_INDEX):C$($END_INDEX)&tq=SELECT%20A%2CB%2CC%20WHERE%20$(Get-Query $Archetype "C")"

                Write-Debug "Archetype Query: $ARCHETYPE_QUERY"
                $Header = @()
                if ($Archetype.Count -gt 1) {
                    $PropertyInputs.Add("Primary Archetype")
                    $InputList.Add("($($Archetype -Join ' | '))")
                    $Header =  "Discord Name", "Character Name", "Primary Archetype"
                }
                else {
                    $Header = "Discord Name", "Character Name"
                }
                $ArchetypeResponse = Invoke-WebRequest -Uri "$($URL)$($ARCHETYPE_QUERY)" | ConvertFrom-Csv -Header $Header
                $ResponseList.Add($ArchetypeResponse)
            }

            if (![string]::IsNullOrEmpty($RP)) {
                for ($i = 0; $i -lt $RP.Count; $i++) {
                    $RP[$i] = Get-RPPriority($RP[$i])
                }
                if ($RP -eq "Error") { break }
                
                [string]$RP_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):G$($END_INDEX)&tq=SELECT%20A%2CB%2CG%20WHERE%20$(Get-Query $RP "G")"

                Write-Debug "RP Query: $RP_QUERY"
                $Header = @()
                if ($RP.Count -gt 1) {
                    $PropertyInputs.Add("RP Priority")
                    $InputList.Add("($($RP -Join ' | '))")
                    $Header =  "Discord Name", "Character Name", "RP Priority"
                }
                else {
                    $Header = "Discord Name", "Character Name"
                }
                $RPResponse = Invoke-WebRequest -Uri "$($URL)$($RP_QUERY)" | ConvertFrom-Csv -Header $Header
                $ResponseList.Add($RPResponse)
            }

            while ($ResponseList.Count -gt 1) {
                $Initial_Count = $ResponseList.Count - 1
                for ($i = 0; $i -lt $Initial_Count; $i += 2) {
                    $Response = Compare-Output $PropertyInputs $InputList $ResponseList[$i] $ResponseList[$i + 1] $Artisan
                    $ResponseList.Add($Response)
                }
                if($Initial_Count % 2 -eq 0) {
                    $ResponseList.RemoveRange(0, $Initial_Count)
                } 
                else {
                    $ResponseList.RemoveRange(0, $Initial_Count + 1)
                }
            }
            
            if (($HasTimezone -eq $false) -or ($Timezone.Count -gt 1)) {
                $ResponseList[0] | Sort-Object -Property "Discord Name" -Unique | Format-Table -AutoSize
            }
            else {
                $ResponseList[0] | Select-Object -Property * -ExcludeProperty Timezone | Sort-Object -Property "Discord Name" -Unique | Format-Table -AutoSize
            }
        }
        {$_ -in "help", "h"} { 
            Get-BTTC-Info
        }
        {$_ -in "sheet"} { 
            Start-Process "https://tinyurl.com/bttc-spreadsheet"
        }
        {$_ -in "links"} { 
            Write-Host "LoreForged: https://linktr.ee/loreforged"
            Write-Host "Google Sheet: https://tinyurl.com/bttc-spreadsheet"
        }
        {$_ -in "v", "version" } { 
            Write-Host $VERSION
        }
        {$_ -in "cl", "changelog" } { 
            Get-ChangeLog
        }
        default { 
            Get-BTTC-Info 
        }
    }
}

Function Get-Query(
    [System.Collections.Generic.List[string]]$Response,
    [string]$Column
) {
    [string]$Query = ""
    foreach ($r in $Response) {
        $Query += "($Column%20like%20'$r')%20OR%20"
    }
    return $Query.TrimEnd("%20OR%20").Trim()
}

Function Compare-Output(
    [System.Collections.Generic.List[string]]$PropertyInputs,
    [System.Collections.Generic.List[string]]$InputList, 
    [System.Collections.Generic.List[PSCustomObject]]$Response1, 
    [System.Collections.Generic.List[PSCustomObject]]$Response2,
    [System.Collections.Generic.List[string]]$Artisans
) {
    $Response = [System.Collections.Generic.List[PSCustomObject]]::new()
    [switch]$SkipAddingProperty = $false

    $c1 = $Response1[0].PSobject.Properties.Name.Count
    $c2 = $Response2[1].PSobject.Properties.Name.Count
    if (($c1 -eq 2) -or ($c2 -eq 2)) {
        Write-Debug "Skipping Adding new Property"
        $SkipAddingProperty = $true
    }

    if ($SkipAddingProperty -eq $false) {
        [string]$MissingProperty
        foreach ($prop in $PropertyInputs) {
            if (!($Response1[0].PSobject.Properties.Name -Contains $prop)) {
                $MissingProperty = $prop
                Write-Debug "Missing Property = $MissingProperty"
                break
            }
        }
    }

    foreach ($r1 in $Response1) {
        foreach ($r2 in $Response2) {
            if ($r1."Discord Name" -eq $r2."Discord Name") {
                if ($Artisans.Count -gt 1) {
                    if (![string]::IsNullOrEmpty($r2."Grandmaster 1"))
                    {
                        $Artrisan1 = Get-Artisan($r2."Grandmaster 1")
                        if (($Artrisan1 -in $Artisans) -and !($r1.PSobject.Properties.Name -Contains "Grandmaster 1")) {
                            $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 1" -Value $Artrisan1
                        }
                        else {
                            $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 1" -Value ""
                        }
                    }
                    else {
                        $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 1" -Value ""
                    }
                    if (![string]::IsNullOrEmpty($r2."Grandmaster 2"))
                    {
                        $Artrisan2 = Get-Artisan($r2."Grandmaster 2")
                        if (($Artrisan2 -in $Artisans) -and !($r1.PSobject.Properties.Name -Contains "Grandmaster 2")) {
                            $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 2" -Value $Artrisan2
                        }
                        else {
                            $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 2" -Value ""
                        }
                    }
                    else {
                        $r1 | Add-Member -MemberType NoteProperty -Name "Grandmaster 2" -Value ""
                    }
                    $Response.Add($r1)
                }
                
                if (($SkipAddingProperty -eq $false) -and !($r1.PSobject.Properties.Name -Contains $MissingProperty)) {
                    $r1 | Add-Member -MemberType NoteProperty -Name "$MissingProperty" -Value $($r2."$MissingProperty")
                    $Response.Add($r1)
                }
                elseif (($SkipAddingProperty -eq $true) -and ($c1 -gt $c2)) {
                    $Response.Add($r1)
                }
                elseif (($SkipAddingProperty -eq $true) -and ($c2 -gt $c1)) {
                    $Response.Add($r2)
                }
                break
            }
        }
    }

    if ($Response.Count -eq 0) {
        Write-Host "No Entries Found for `"$($($InputList -Join ' + ').TrimStart(" + "))`""
    }
    return $Response
}

Function Compare-ObjectBool {
    param(
      [Parameter(Mandatory = $true)]
      [PSCustomObject] $firstObject,
  
      [Parameter(Mandatory = $true)]
      [PSCustomObject] $secondObject
    )
    -not (Compare-Object $firstObject.PSObject.Properties $secondObject.PSObject.Properties -Property "Discord Name")
}

Function Get-Archetype(
    [System.Collections.Generic.List[String]]$Archetype
) {
    for ($i = 0; $i -lt $Archetype.Count; $i++) {
        Switch($Archetype[$i]) {
            { $_ -like "Rogue" } {
                $Archetype[$i] = "Rogue"
            }
            { $_ -like "Fighter" } {
                $Archetype[$i] = "Fighter"
            }
            { $_ -like "Cleric" } {
                $Archetype[$i] = "Cleric"
            }
            { $_ -like "Tank" } {
                $Archetype[$i] = "Tank"
            }
            { $_ -like "Summoner" } {
                $Archetype[$i] = "Summoner"
            }
            { $_ -like "Bard" } {
                $Archetype[$i] = "Bard"
            }
            { $_ -like "Ranger" } {
                $Archetype[$i] = "Ranger"
            }
            { $_ -like "Mage" } {
                $Archetype[$i] = "Mage"
            }
            default {
                Write-Error "`"$($Archetype[$i])`" is not a valid Archetype"
                return "Error"
            }
        }
    }
    return $Archetype
}

Function Get-Timezone(
    [System.Collections.Generic.List[String]]$Timezone
) {
    for ($i = 0; $i -lt $Timezone.Count; $i++) {
        Switch($Timezone[$i]) {
            { $_ -like "PST" } {
                $Timezone[$i] = "PST"
            }
            { $_ -like "CST" } {
                $Timezone[$i] = "CST"
            }
            { $_ -like "MST" } {
                $Timezone[$i] = "MST"
            }
            { $_ -like "EST" } {
                $Timezone[$i] = "EST"
            }
            { $_ -like "European" -or $_ -like "Euro" } {
                $Timezone[$i] = "European"
            }
            { $_ -like "New Zealand" -or $_ -like "NZ" } {
                $Timezone[$i] = "New Zealand"
            }
            { $_ -like "Australia" -or $_ -like "AUS" } {
                $Timezone[$i] = "Australia"
            }
            { $_ -like "Something Wild" -or $_ -like "SW" } {
                $Timezone[$i] = "Something Wild"
            }
            default {
                Write-Error "`"$Timezone[$i]`" is not a valid Timezone"
                return "Error"
            }
        }
    }
    return $Timezone
}

Function Get-Branch(
    [System.Collections.Generic.List[String]]$Branch
) {
    for ($i = 0; $i -lt $Branch.Count; $i++) {
        Switch($Branch[$i]) {
            { $_ -like "Warden" } {
                $Branch[$i] = "Warden"
            }
            { $_ -like "Capital" } {
                $Branch[$i] = "Capital"
            }
            { $_ -like "Warborn" } {
                $Branch[$i] = "Warborn"
            }
            default {
                Write-Error "`"$Branch[$i]`" is not a valid Branch"
                return "Error"
            }
        }
    }
    return $Branch
}

Function Get-RPPriority(
    [string]$RP
) {
    switch($RP) {
        { $_ -like "High"  } {
            $RP = "High"
        }
        { $_ -like "Medium"  } {
            $RP = "Medium"
        }
        { $_ -like "Low"} {
            $RP = "Low"
        }
        { $_ -like "None"} {
            $RP = "None"
        }
        default {
            Write-Error "`"$RP`" is not a valid RP Priority"
            return "Error"
        }
    }
    return $RP
}

Function Get-Artisan(
    [string]$Artisan
) {
    Switch($Artisan) {
        { $_ -like "Fishing" -or $_ -like "Fish" -or $_.Contains("Fishing") } {
            $Artisan = "Fishing"
        }
        { $_ -like "Herbalism" -or $_ -like "Herb" -or $_.Contains("Herbalism") } {
            $Artisan = "Herbalism"
        }
        { $_ -like "Hunting" -or $_ -like "Hunt" -or $_.Contains("Hunt") } {
            $Artisan = "Hunting"
        }
        { $_ -like "Lumberjacking" -or $_.Contains("Lumberjacking") } {
            $Artisan = "Lumberjacking"
        }
        { $_ -like "Mining" -or $_.Contains("Mining") } {
            $Artisan = "Mining"
        }
        { $_ -like "Alchemy" -or $_.Contains("Alchemy") } {
            $Artisan = "Alchemy"
        }
        { $_ -like "Animal Husbandry" -or $_ -like "Husbandry" -or $_.Contains("Animal Husbandry") } {
            $Artisan = "Animal Husbandry"
        }
        { $_ -like "Cooking" -or $_ -like "Cook" -or $_.Contains("Cooking") } {
            $Artisan = "Cooking"
        }
        { $_ -like "Farming" -or $_ -like "Farm" -or $_.Contains("Farming") } {
            $Artisan = "Farming"
        }
        { $_ -like "Lumber Milling" -or $_ -like "Milling" -or $_.Contains("Lumber Milling") } {
            $Artisan = "Lumber Milling"
        }
        { $_ -like "Metalworking" -or $_ -like "Metal" -or $_.Contains("Metalworking") } {
            $Artisan = "Metalworking"
        }
        { $_ -like "Stonemasonary" -or $_ -like "Stone" -or $_.Contains("Stonemasonary") } {
            $Artisan = "Stonemasonary"
        }
        { $_ -like "Tanning" -or $_ -like "Tan" -or $_.Contains("Tanning") } {
            $Artisan = "Tanning"
        }
        { $_ -like "Weaving" -or $_ -like "Weave" -or $_.Contains("Weaving") } {
            $Artisan = "Weaving"
        }
        { $_ -like "Arcane Engineering" -or $_ -like "Engineering" -or $_ -like "Arcane" -or $_.Contains("Arcane Engineering") } {
            $Artisan = "Arcane Engineering"
        }
        { $_ -like "Armor Smithing" -or $_ -like "Armor" -or $_.Contains("Armor Smithing") } {
            $Artisan = "Armor Smithing"
        }
        { $_ -like "Carpentry" -or $_.Contains("Carpentry") } {
            $Artisan = "Carpentry"
        }
        { $_ -like "Jewel Cutting" -or $_ -like "Jewel" -or $_.Contains("Jewel Cutting") } {
            $Artisan = "Jewel Cutting"
        }
        { $_ -like "Leatherworking" -or $_ -like "Leather" -or $_.Contains("Leatherworking") } {
            $Artisan = "Leatherworking"
        }
        { $_ -like "Scribing" -or $_ -like "Scribe" -or $_.Contains("Scribing") } {
            $Artisan = "Scribing"
        }
        { $_ -like "Tailoring" -or $_ -like "Tailor" -or $_.Contains("Tailoring") } {
            $Artisan = "Tailoring"
        }
        { $_ -like "Weapon Smithing" -or $_ -like "Weapon" -or $_.Contains("Weapon Smithing") } {
            $Artisan = "Weapon Smithing"
        }
        default {
            Write-Error "`"$Artisan`" is not a valid Artisan"
            return "Error"
        }
    }
    return $Artisan
}

Function Get-ChangeLog() {
    Write-Host @'

Legend
✅ -> Completed new feature/update
🔥 -> Completed bug/hot fix
⌛ -> Completed performance update/fix
✍️ -> Additional Info about a Change
📜 -> Higher level identifier for feature changes
❌ -> In-progress feature will be fixed in later release

v1.2.0 - Sept 14th, 2024
✅ Added new filter "RP" to filter for a BTTC members "RP Priority"
🔥 Fixed "Animal Husbandry" misspelling

v1.1.0 - Sept 8th, 2024
✅ Able to filter for multiple Archetype(s), Branch(es), Artisan(s) and/or Timezone(s) parameter(s) at the same time, this constitues or'ing NOT and'ing
🔥 Fixed help command text
🔥 Fixed issue where "Cook" alias for Artisan would get "Farming" instead of "Cooking"
🔥 Fixed issue where "Farm" alias for Artisan would get "Cooking" instead of "Farming"

v1.0.2 - Sept 7th, 2024
✅ Added new command ["v", "version"] to print the CLI version
✅ Added new command ["sheet"] to open the google sheet in your browser
✅ Added new command [`"cl`", `"changelog`"] to print the changelog for the script to display changes for each version
🔥 Command parameter was labeled "CaseSensitive" when it should be "Not CaseSensitive"
🔥 If only filtering by Timezone it now correctly throws error
📜 Timezone Param Alias Updates
- ✅ "EURO" for "European"
- ✅ "AUS" for "Australia"
- ✅ "NZ" for "New Zealand"
- ✅ "SW" for "Something Wild"
📜 Artisan Param Alias Updates
- ✅ "Metal" for "Metalworking"
- ✅ "Weave" for "Weaving"
- ✅ "Leather" for "Leatherworking"
- ✅ "Scribe" for "Scribing"
- ✅ "Tailor" for "Tailoring"
- ✅ "Hunt" for "Hunting"
- ✅ "Herb" for "Herbalism"
- ✅ "Fish" for "Fishing"
- ✅ "Tan" for "Tanning"
- ✅ "Stone" for "Stonemasonary"
- ✅ "Cook" for "Cooking"
- ✅ "Farm" for "Farming"
'@
}

Function Get-BTTC-Info {
    Write-Host "AUTHOR"
    Write-Host "    Sulix (Prem)"
    ""
    Write-Host "VERSION"
    Write-Host "    $VERSION"
    ""
    Write-Host "EXAMPLES"
    Write-Host "    bttc ls -Artisan Cooking (Prints all members who are planning on GM'ing in `"Cooking`")"
    Write-Host "    bttc members -Branch Warden -Archetype Rogue -Timezone PST (Prints all members who are apart of the `"Warden`" branch located in PST timezone & are planning on a primary archetype of `"Rogue`")"
    ""
    Write-Host "SYNTAX (Order doesn't mater)"
    Write-Host "    bttc <Command> [-Parameter(s)] <Value(s)>"
    ""
    Write-Host "COMMANDS"
    Write-Host "    [`"ls`", `"list`", `"members`"] - Lists the `"Discord Name`", `"Character Name`", `"Timezone`", & `"Branch`" of BTTC members from the sheet, can use `"Branch`", `"Artisan`", `"Timezone`", and/or `"Archetype`" params to filter output"
    Write-Host "    [`"links`"] - Displays various links related to BTTC guild"
    Write-Host "    [`"help`", `"h`"] - Displays information about the CLI"
    Write-Host "    [`"sheet`"] - Opens the BTTC google sheet"
    Write-Host "    [`"v`", `"version`"] - Displays the current CLI version"
    Write-Host "    [`"cl`", `"changelog`"] - Print the changelog for the CLI"
    ""
    Write-Host "PARAMETERS"
    Write-Host "    -Branch <String[]> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the BTTC branch [`"Capital`", `"Warborn`", `"Warden`"]"
    ""
    Write-Host "    -Artisan <String[]> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the AoC artisan, use quotes `"`" if the artisan has a space"
    ""
    Write-Host "    -Archetype <String[]> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the AoC primary archetype"
    ""
    Write-Host "    -Timezone <String[]> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the AoC primary archetype, use quotes `"`" if the artisan has a space"
    ""
    Write-Host "    -RP <String[]> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the BTTC `"RP Priority`""
    ""
}