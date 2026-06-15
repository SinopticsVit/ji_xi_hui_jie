<#
.SYNOPSIS
    Install id_hostinger_rikkimortycrypt on a fresh Hostinger VPS and enable ssh pengbo.

.EXAMPLE
    .\install-pengbo-ssh.ps1 -HostName '203.0.113.10' -Password 'one-time-root-password'
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$HostName,

    [Parameter(Mandatory)]
    [string]$Password,

    [int]$Port = 22,

    [string]$User = 'root',

    [string]$Alias = 'pengbo',

    [string]$KeyPath = (Join-Path $env:USERPROFILE '.ssh\id_hostinger_rikkimortycrypt'),

    [switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Write-Step([string]$msg) { Write-Host "==> $msg" -ForegroundColor Cyan }
function Write-Ok  ([string]$msg) { Write-Host "    OK: $msg" -ForegroundColor Green }

$pubPath = "$KeyPath.pub"
$cfgFile = Join-Path $env:USERPROFILE '.ssh\config'
$knownHst = Join-Path $env:USERPROFILE '.ssh\known_hosts'
$toolsDir = Join-Path $PSScriptRoot 'tools'
$plink = Join-Path $toolsDir 'plink.exe'

if (-not (Test-Path -LiteralPath $KeyPath)) { throw "Private key not found: $KeyPath" }
if (-not (Test-Path -LiteralPath $pubPath)) { throw "Public key not found: $pubPath" }

Write-Step 'ACL private key'
& icacls $KeyPath /inheritance:r | Out-Null
& icacls $KeyPath /grant:r "$($env:USERNAME):(R)" | Out-Null
Write-Ok 'ACL applied'

Write-Step "TCP probe $HostName`:$Port"
$tcp = Test-NetConnection -ComputerName $HostName -Port $Port -WarningAction SilentlyContinue
if (-not $tcp.TcpTestSucceeded) { throw "TCP $HostName`:$Port is not reachable" }
Write-Ok 'Port open'

New-Item -ItemType Directory -Force -Path $toolsDir | Out-Null
if (-not (Test-Path -LiteralPath $plink)) {
    Write-Step 'Downloading plink.exe'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri 'https://the.earth.li/~sgtatham/putty/latest/w64/plink.exe' -OutFile $plink -UseBasicParsing
}

Write-Step 'Host key fingerprint'
$scanned = @(& ssh-keyscan -T 10 -t ed25519,rsa -p $Port $HostName 2>$null) | Where-Object { $_ -and $_ -notlike '#*' }
if ($scanned) { Add-Content -LiteralPath $knownHst -Value $scanned }
$kk = & ssh-keygen -l -F $HostName 2>$null
if (-not $kk -or $LASTEXITCODE -ne 0) { throw "Could not read host key for $HostName" }
$fprs = @()
foreach ($line in $kk) {
    if ($line -match 'SHA256:[^\s]+') { $fprs += $matches[0] }
}
$fprs = $fprs | Select-Object -Unique
if ($fprs.Count -eq 0) { throw 'Could not parse host key fingerprints' }
Write-Ok ($fprs -join ', ')

Write-Step 'Install public key on server'
$pubKey = (Get-Content -LiteralPath $pubPath -Raw).Trim()
$probe = & ssh -o BatchMode=yes -o ConnectTimeout=10 -o IdentitiesOnly=yes -o PreferredAuthentications=publickey -i $KeyPath -p $Port "$User@$HostName" 'echo KEY_OK' 2>&1
if ($LASTEXITCODE -eq 0 -and ($probe -join "`n") -match 'KEY_OK') {
    Write-Ok 'Key auth already works'
} else {
    $remote = @"
mkdir -p ~/.ssh && chmod 700 ~/.ssh && \
touch ~/.ssh/authorized_keys && chmod 600 ~/.ssh/authorized_keys && \
grep -qxF '$pubKey' ~/.ssh/authorized_keys || echo '$pubKey' >> ~/.ssh/authorized_keys && \
echo OK_INSTALLED
"@
    $plinkArgs = @('-batch','-ssh','-P', "$Port")
    foreach ($f in $fprs) { $plinkArgs += @('-hostkey', $f) }
    $plinkArgs += @('-pw', $Password, "$User@$HostName", $remote)
    $result = & $plink @plinkArgs 2>&1
    if ($LASTEXITCODE -ne 0 -or ($result -join "`n") -notmatch 'OK_INSTALLED') {
        throw "Failed to install public key via plink"
    }
    Write-Ok 'Public key installed'
}

Write-Step "Update $cfgFile"
$keyForConfig = ($KeyPath -replace '\\', '/')
$block = @"

# Hostinger KVM — Pengbo Aviation VPS (rikkimortycrypt@gmail.com)
Host $Alias
    HostName $HostName
    User $User
    Port $Port
    IdentityFile $keyForConfig
    IdentitiesOnly yes
    PreferredAuthentications publickey
    ServerAliveInterval 60
    ServerAliveCountMax 3
"@

$existingCfg = if (Test-Path -LiteralPath $cfgFile) { Get-Content -LiteralPath $cfgFile -Raw } else { '' }
$pattern = "(?ms)(^|\r?\n)\s*(#[^\r\n]*\r?\n)?Host\s+$([regex]::Escape($Alias))\b.*?(?=(\r?\n\s*Host\s+|\z))"
$hasBlock = [regex]::IsMatch($existingCfg, $pattern)
if ($hasBlock -and -not $Force) {
    $existingCfg = [regex]::Replace($existingCfg, $pattern, '')
}
$newCfg = ($existingCfg.TrimEnd() + "`r`n" + $block.TrimStart() + "`r`n")
[System.IO.File]::WriteAllText($cfgFile, $newCfg, (New-Object System.Text.UTF8Encoding $false))
Write-Ok "Host $Alias configured"

Write-Step "Test ssh $Alias"
$test = & ssh -o BatchMode=yes -o ConnectTimeout=10 $Alias "echo OK_KEY_AUTH; whoami; hostname" 2>&1
Write-Host $test
if ($LASTEXITCODE -ne 0 -or ($test -join "`n") -notmatch 'OK_KEY_AUTH') {
    throw "ssh $Alias failed"
}
Write-Ok "ssh $Alias works"
