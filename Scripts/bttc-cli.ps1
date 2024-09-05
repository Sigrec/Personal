Function Get-BTTC-Info {
    Write-Host "AUTHOR"
    Write-Host "    Sulix (Prem)"
    ""
    Write-Host "VERSION"
    Write-Host "    v1.0.0"
    ""
    Write-Host "SYNTAX"
    Write-Host "    bttc [-Command] <string> [-Branch] <string> [-Artisan] <string> [-Archetype] <string> [-Timezone] <string>"
    ""
    Write-Host "COMMANDS"
    Write-Host "    [`"list`", `"members`", `"ls`"] - Lists the `"Discord Name`", `"Character Name`", `"Timezone`", & `"Branch`" of BTTC members from the sheet, can use `"Branch`", `"Artisan`", and/or `"Archetype`" params to filter output"
    Write-Host "    [`"links`"] - Displays various links related to BTTC guild"
    Write-Host "    [`"help`", `"h`"] - Displays information about the binary"
    ""
    Write-Host "PARAMETERS"
    Write-Host "    -Command <String>"
    Write-Host "       Specifies the command being used from [`"list`", `"help`"]"
    ""
    Write-Host "    -Branch <String> (Optional) [Not CaseSensitive]"
    Write-Host "        Specifies the BTTC branch [`"capital`", `"warborn`", `"warden`"], can be combined with `"-Artisan`" or `"-Archetype`" param for additional filtering"
    ""
    Write-Host "    -Artisan <String> (Optional)"
    Write-Host "        Specifies the AoC artisan, use quotes `"`" if the artisan has a space, can be combined with `"-Branch`" param for additional filterin"
    ""
    Write-Host "    -Archetype <String> (Optional)"
    Write-Host "        Specifies the AoC primary archetype, can be combined with `"-Branch`" param for additional filtering"
    ""
    Write-Host "    -Timezone <String> (Optional)"
    Write-Host "        Specifies the AoC primary archetype, use quotes `"`" if the artisan has a space, can be combined with any other param for additional filtering"
    ""
    Write-Host "EXAMPLES"
    Write-Host "    bttc ls -Artisan Cooking (Prints all members who are planning on GM'ing in `"Cooking`")"
    Write-Host "    bttc members -Branch Warden -Archetype Rogue -Timezone PST (Prints all members who are apart of the `"Warden`" branch located in PST & are planning on a primary archetype of `"Rogue`")"
}

Function bttc()
{
    param(
        [Parameter(Mandatory=$true, Position=0)][string]$Command,
        [Parameter(Mandatory=$false)][string]$Branch,
        [Parameter(Mandatory=$false)][string]$Artisan,
        [Parameter(Mandatory=$false)][string]$Archetype,
        [Parameter(Mandatory=$false)][string]$Timezone
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
            if (![string]::IsNullOrEmpty($Branch)) { $Branch = Get-Branch($Branch.ToLower()) }
            if (![string]::IsNullOrEmpty($Artisan)) { $Artisan = Get-Artisan($Artisan.ToLower()) }
            if (![string]::IsNullOrEmpty($Archetype)) { $Archetype = (Get-Culture).TextInfo.ToTitleCase($Archetype.ToLower()) }
            if (![string]::IsNullOrEmpty($Timezone)) { 
                $Timezone = $Timezone.ToUpper() 
                $TIMEZONE_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CF%20WHERE%20upper(F)%20%3D%20'$Timezone'"
            }
            else {
                $TIMEZONE_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CF"
            }
            Write-Debug "Timezone Query = $TIMEZONE_QUERY"

            $CSV_HEADERS = "Discord Name", "Character Name", "Timezone"
            $BRANCH_QUERY = "&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):C$($END_INDEX)&tq=SELECT%20A%2CB%2CC%20WHERE%20C%20%3D%20'$Branch'"
            $ARTISAN_QUERY = "&sheet=$ARTISAN_SHEET&range=A$($START_INDEX):D$($END_INDEX)&tq=SELECT%20A%2CB%2CC%2CD%20WHERE%20C%20contains%20'$Artisan'%20OR%20D%20contains%20'$Artisan'"
            $ARCHETYPE_QUERY = "&sheet=$ARCHETYPE_SHEET&range=A$($START_INDEX):C$($END_INDEX)&tq=SELECT%20A%2CB%2CC%20WHERE%20C%20%3D%20'$Archetype'"
            $TIMEZONE_RESPONSE = Invoke-WebRequest -Uri "$($URL)$($TIMEZONE_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS

            if (!$Branch -and !$Artisan -and !$Archetype) {
                $Response = Invoke-WebRequest -Uri "$($URL)&sheet='$GENERIC_SHEET'&range=A$($START_INDEX):F$($END_INDEX)&tq=SELECT%20A%2CB%2CC%2CF" | ConvertFrom-Csv
                if (![string]::IsNullOrEmpty($Timezone)) {
                    $Response = $Response | Where-Object { $_.Timezone.ToUpper() -eq $Timezone } | Select-Object -Property * -ExcludeProperty Timezone
                }
            }
            elseif ($Branch -and !$Artisan -and !$Archetype) { # Lists all members for a specific branch
                Write-Debug "Branch Query = $BRANCH_QUERY"
                $Branch_Response = Invoke-WebRequest -Uri "$($URL)$($BRANCH_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Response = Compare-Output $Branch $Timezone $Branch_Response $TIMEZONE_RESPONSE
            }
            elseif ($Artisan -and !$Branch -and !$Archetype) { # Lists all members that are planning on GM'ing this artisan
                Write-Debug "Artisan Query = $ARTISAN_QUERY"
                $Artisan_Response = Invoke-WebRequest -Uri "$($URL)$($ARTISAN_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Response = Compare-Output $Artisan $Timezone $Artisan_Response $TIMEZONE_RESPONSE
            }
            elseif ($Branch -and $Artisan -and !$Archetype) { # Lists all members that are planning on GM'ing this artisan for a specific branch
                Write-Debug "Branch Query = $BRANCH_QUERY"
                Write-Debug "Artisan Query = $ARTISAN_QUERY"
                $Branch_Response = Invoke-WebRequest -Uri "$($URL)$($BRANCH_QUERY)"  | ConvertFrom-Csv -Header $CSV_HEADERS
                $Artisan_Response = Invoke-WebRequest -Uri "$($URL)$($ARTISAN_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Response = Compare-Output $Artisan "$Branch + $Timezone" $(Compare-Output $Artisan $Branch $Branch_Response $Artisan_Response) $TIMEZONE_RESPONSE
            }
            elseif (!$Branch -and !$Artisan -and $Archetype) { # Lists all members that are planning a specific primary archetype
                Write-Debug $ARCHETYPE_QUERY
                $Archetype_Response = Invoke-WebRequest -Uri "$($URL)$($ARCHETYPE_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Response = Compare-Output $Archetype $Timezone $Archetype_Response $TIMEZONE_RESPONSE
            }
            elseif ($Branch -and !$Artisan -and $Archetype) { # Lists all members that are planning a specific archetype and are from a specific branch
                Write-Debug "Branch Query = $BRANCH_QUERY"
                Write-Debug $ARCHETYPE_QUERY
                $Branch_Response = Invoke-WebRequest -Uri "$($URL)$($BRANCH_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Archetype_Response = Invoke-WebRequest -Uri "$($URL)$($ARCHETYPE_QUERY)" | ConvertFrom-Csv -Header $CSV_HEADERS
                $Response = Compare-Output $Archetype "$Branch + $Timezone" $(Compare-Output $Archetype $Branch $Branch_Response $Archetype_Response) $TIMEZONE_RESPONSE
            }
            elseif ($Archetype -and $Artisan)
            {
                Write-Error("Can not combine `"Archetype`" & `"Artisan`" Params")
                break
            }
            $Response | Sort-Object -Property "Discord Name" | Format-Table
        }
        {$_ -in "help", "h"} { 
            Get-BTTC-Info
        }
        {$_ -in "links"} { 
            Write-Host "LoreForged: https://linktr.ee/loreforged"
            Write-Host "Google Sheet: https://tinyurl.com/bttc-spreadsheet"
        }
        default { 
            Get-BTTC-Info 
        }
    }
}

Function Compare-Output([string]$Input1, [string]$Input2, [System.Object[]]$Response1, [System.Object[]]$Response2) {
    $Response = @()
    ForEach ($a in $Response1) {
        ForEach ($b in $Response2) {
            if ($a."Discord Name" -eq $b."Discord Name") {
                $Response += $b
                break
            }
        }
    }
    if ($Response.Count -eq 0) {
        if (![string]::IsNullOrEmpty($Input2)) {
            Write-Host "No Entries Found for `"$Input1 + $Input2`""
        }
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

Function Get-Branch($Branch) {
    Switch($Branch) {
        { $_ -like "Warden" } {
            $Branch = "Warden"
        }
        { $_ -like "Capital" } {
            $Branch = "Capital"
        }
        { $_ -like "Warborn" } {
            $Branch = "Warborn"
        }
        default {
            Write-Error "`"$($Branch)`" is not a valid Branch name"
            break
        }
    }
    return $Branch
}

Function Get-Artisan($Artisan) {
    Switch($Artisan) {
        { $_ -like "Fishing" } {
            $Artisan = "Fishing"
        }
        { $_ -like "Herbalism" } {
            $Artisan = "Herbalism"
        }
        { $_ -like "Hunting" } {
            $Artisan = "Hunting"
        }
        { $_ -like "Lumberjacking" } {
            $Artisan = "Lumberjacking"
        }
        { $_ -like "Mining" } {
            $Artisan = "Mining"
        }
        { $_ -like "Alchemy" } {
            $Artisan = "Alchemy"
        }
        { $_ -like "Animal Husbrandry" -or $_ -like "Husbrandry" } {
            $Artisan = "Animal Husbrandry"
        }
        { $_ -like "Cooking" } {
            $Artisan = "Cooking"
        }
        { $_ -like "Farming" } {
            $Artisan = "Farming"
        }
        { $_ -like "Lumber Milling" -or $_ -like "Lumber" -or $_ -like "Milling" } {
            $Artisan = "Lumber Milling"
        }
        { $_ -like "Metalworking" } {
            $Artisan = "Metalworking"
        }
        { $_ -like "Stonemasonary" } {
            $Artisan = "Stonemasonary"
        }
        { $_ -like "Tanning" } {
            $Artisan = "Tanning"
        }
        { $_ -like "Weaving" } {
            $Artisan = "Weaving"
        }
        { $_ -like "Arcane Engineering" -or $_ -like "Engineering" -or $_ -like "Arcane" } {
            $Artisan = "Arcane Engineering"
        }
        { $_ -like "Armor Smithing" -or $_ -like "Armor" } {
            $Artisan = "Armor Smithing"
        }
        { $_ -like "Carpentry" } {
            $Artisan = "Carpentry"
        }
        { $_ -like "Jewel Cutting" -or $_ -like "Jewel" } {
            $Artisan = "Jewel Cutting"
        }
        { $_ -like "Leatherworking" } {
            $Artisan = "Leatherworking"
        }
        { $_ -like "Scribing" } {
            $Artisan = "Scribing"
        }
        { $_ -like "Tailoring" } {
            $Artisan = "Tailoring"
        }
        { $_ -like "Weapon Smithing" -or $_ -like "Weapon" } {
            $Artisan = "Weapon Smithing"
        }
        default {
            Write-Error "`"$($Artisan)`" is not a valid Artisan"
            break
        }
    }
    return $Artisan
}