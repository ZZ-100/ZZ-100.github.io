param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Build', 'Publish')]
    [string]$Action,
    [string]$MessageBase64 = '',
    [string]$ResultPathBase64 = ''
)

$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
. (Join-Path $Root 'blog-lib.ps1')

function ConvertTo-Base64Text([string]$Text) {
    if ($null -eq $Text) { $Text = '' }
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text))
}

try {
    if ($Action -eq 'Build') {
        Invoke-HexoBuild -Root $Root | Out-Null
        Test-PublicLinks -Root $Root | Out-Null
        $result = [pscustomobject]@{
            Success = $true
            Changed = $false
            MessageBase64 = ConvertTo-Base64Text 'Build and link check passed.'
            DetailBase64 = ''
        }
    } else {
        $message = 'Update content'
        if ($MessageBase64) {
            $message = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($MessageBase64))
        }
        $published = [bool](Publish-Git -Message $message -Root $Root)
        $result = [pscustomobject]@{
            Success = $true
            Changed = $published
            MessageBase64 = ConvertTo-Base64Text 'Publish command completed.'
            DetailBase64 = ''
        }
    }
} catch {
    $result = [pscustomobject]@{
        Success = $false
        Changed = $false
        MessageBase64 = ConvertTo-Base64Text $_.Exception.Message
        DetailBase64 = ConvertTo-Base64Text ($_ | Out-String)
    }
}

$json = $result | ConvertTo-Json -Compress
if ($ResultPathBase64) {
    $resultPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($ResultPathBase64))
    [IO.File]::WriteAllText($resultPath, $json, (New-Object Text.UTF8Encoding($false)))
}
$json
if (-not $result.Success) { exit 1 }
