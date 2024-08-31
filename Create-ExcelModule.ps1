if($PSVersionTable.PSVersion.Major -lt 7)
{
    Write-Error -Message 'This script supports PowerShell 7 or later.' -ErrorAction Stop
}
Set-Variable -Name ModulePath -Value (Join-Path -Path $PSScriptRoot -ChildPath 'XLFUT.txt') -Option ReadOnly
Set-Variable -Name PartsDirectory -Value (Join-Path -Path $PSScriptRoot -ChildPath 'src') -Option ReadOnly

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

(
    (
        Get-ChildItem　-Path (Join-Path -Path $PartsDirectory -ChildPath "*.txt") | ForEach-Object -Process {
            $_ | Add-Member -MemberType NoteProperty -Name NameWithoutExtension -Value ([IO.Path]::GetFileNameWithoutExtension($_)) -PassThru
        } | Sort-Object -Property NameWithoutExtension | ForEach-Object -Process {
            $filePath = $_.FullName
            @(
                "/***** $($_.NameWithoutExtension) *****/",
                ""
                (Get-Content -Raw -Path $filePath)
            ) -join "`n"
        }
    ) -join "`n"
).Replace('${version}', $version) | Out-File -FilePath $ModulePath -Encoding utf8NoBOM -NoNewline
