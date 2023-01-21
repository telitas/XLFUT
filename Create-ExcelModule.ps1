if($PSVersionTable.PSVersion.Major -lt 7)
{
    Write-Error -Message 'This script supports PowerShell 7 or later.' -ErrorAction Stop
}
Set-Variable -Name ModulePath -Value (Join-Path -Path $PSScriptRoot -ChildPath 'XLFUT.txt') -Option ReadOnly
Set-Variable -Name PartsDirectory -Value (Join-Path -Path $PSScriptRoot -ChildPath 'src') -Option ReadOnly
Set-Variable -Name PartList -Value @(
    'XLFUT.txt'
    'XLFUT.UTILS.txt'
    'XLFUT.ASSERT.txt'
)
if((git tag --list --contains HEAD) -match '^v(?<version>[0-9]+\.[0-9]+\.[0-9])')
{
    $version = $Matches['version']
}
else
{
    Write-Warning 'The current Commit does not contain versioning tag.'
    $version = '0.0.1'
}
if(Test-Path -Path $ModulePath){
    Write-Warning -Message 'Old Module was removed.'
    Remove-Item -Path $ModulePath -Recurse
}
New-Item -ItemType File -Path $ModulePath > $null

$PartList | ForEach-Object -Process {
    $filePath = Join-Path -Path $PartsDirectory -ChildPath $_
    Get-Content -Raw -Path $filePath | ForEach-Object -Process {
        $content = $_
        switch(Split-Path -Path $filePath -Leaf)
        {
            'XLFUT.txt' {
                $content.Replace('${version}', $version)
            }
            default {
                $content
            }
        }
    } | Out-File -FilePath $ModulePath -Append -Encoding utf8NoBOM
}
