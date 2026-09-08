param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^v\d+\.\d+\.\d+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$instalador = Join-Path $PSScriptRoot "dist\installer\SiPP-Setup-$Version.exe"
$suma = Join-Path $PSScriptRoot 'dist\installer\SHA256SUMS.txt'

if (-not (Test-Path $instalador)) {
    throw "No existe el instalador local: $instalador"
}

$hash = (Get-FileHash $instalador -Algorithm SHA256).Hash.ToLower()
"$hash  SiPP-Setup-$Version.exe" | Set-Content $suma -Encoding ascii

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw 'Instala GitHub CLI y autentica con: gh auth login'
}

gh release view $Version *> $null
if ($LASTEXITCODE -ne 0) {
    gh release create $Version --title "SiPP $Version" --notes "Instalador oficial de SiPP."
}

gh release upload $Version $instalador $suma --clobber
Write-Output "Instalador publicado: $instalador"
Write-Output "Release: https://github.com/dsfenton14-svg/Sipp-Aplicacion/releases/tag/$Version"
