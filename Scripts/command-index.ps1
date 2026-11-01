function Get-MyCommand {
    <#
    .SYNOPSIS
    Lists every custom command and alias defined in the Scripts folder.

    .DESCRIPTION
    Parses each script in the Scripts folder (without running it) and lists its
    functions with their aliases and synopsis, grouped by file. Because it reads
    the files rather than the session, it also shows commands you have not
    loaded yet.

    .PARAMETER Name
    Wildcard filter on the command name, alias or synopsis, such as git* or *pr*.

    .PARAMETER PassThru
    Return the objects instead of printing a grouped table.

    .EXAMPLE
    cmds

    Shows all custom commands grouped by script.

    .EXAMPLE
    cmds git*

    Shows only the commands whose name, alias or synopsis starts with git.

    .NOTES
    Alias: cmds. For full help on one command run: help <command> -Full
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Position = 0)]
        [string]$Name = '*',

        [switch]$PassThru
    )

    $scriptsDir = $PSScriptRoot
    if (-not $scriptsDir) { $scriptsDir = 'S:\Personal\Scripts' }

    $rows = foreach ($file in Get-ChildItem -Path $scriptsDir -Filter *.ps1) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)

        $aliases = @{}
        foreach ($command in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $false) |
                Where-Object { $_.GetCommandName() -eq 'Set-Alias' }) {
            $elements = $command.CommandElements
            $aliasName = $null
            $target = $null
            for ($i = 1; $i -lt $elements.Count; $i++) {
                $text = $elements[$i].Extent.Text
                if ($text -eq '-Name' -and $i + 1 -lt $elements.Count) { $aliasName = $elements[++$i].Extent.Text }
                elseif ($text -eq '-Value' -and $i + 1 -lt $elements.Count) { $target = $elements[++$i].Extent.Text }
            }
            if ($aliasName -and $target) { $aliases[$target] = $aliasName }
        }

        foreach ($function in $ast.EndBlock.Statements |
                Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] }) {
            [pscustomobject]@{
                Script   = $file.BaseName
                Alias    = $aliases[$function.Name]
                Command  = $function.Name
                Synopsis = ($function.GetHelpContent().Synopsis ?? '').Trim()
            }
        }
    }

    $rows = @($rows | Where-Object {
            $_.Command -like $Name -or $_.Alias -like $Name -or $_.Synopsis -like $Name
        } | Sort-Object Script, Command)

    if ($PassThru) { return $rows }

    if (-not $rows) {
        Write-Host "No custom commands match '$Name'." -ForegroundColor Yellow
        return
    }

    $rows | Format-Table Alias, Command, Synopsis -GroupBy Script -AutoSize | Out-Host
}

Set-Alias -Name cmds -Value Get-MyCommand
