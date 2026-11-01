function Push-WpCode {
    <#
    .SYNOPSIS
    Jumps to the WholesaleProgramPlatform code repo.

    .DESCRIPTION
    Changes to the code repo folder and sets the terminal title to "Code".

    .EXAMPLE
    wpc

    .NOTES
    Alias: wpc
    #>
    [CmdletBinding()]
    param ()

    Set-Location -LiteralPath 'S:\TCG SaaS Platform\WholesaleProgramPlatform'
    $Host.UI.RawUI.WindowTitle = 'Code'
}

function Push-WpWiki {
    <#
    .SYNOPSIS
    Jumps to the WholesaleProgramPlatform wiki repo.

    .DESCRIPTION
    Changes to the wiki repo folder and sets the terminal title to "Wiki".

    .EXAMPLE
    wpw

    .NOTES
    Alias: wpw
    #>
    [CmdletBinding()]
    param ()

    Set-Location -LiteralPath 'S:\TCG SaaS Platform\WholesaleProgramPlatform.wiki'
    $Host.UI.RawUI.WindowTitle = 'Wiki'
}

Set-Alias -Name wpc -Value Push-WpCode
Set-Alias -Name wpw -Value Push-WpWiki
