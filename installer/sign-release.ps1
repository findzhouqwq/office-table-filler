#Requires -Version 5.1
<#
.SYNOPSIS
    生成带 RSA 签名的发布清单（app.zip + manifest.json + manifest.sig）。

.DESCRIPTION
    安装程序只信任内置公钥签名的清单：
      1. 用 git archive 打包当前提交，保证内容可追溯；
      2. 计算 app.zip 的 SHA256；
      3. 生成 manifest.json；
      4. 用私钥对 manifest.json 的原始字节做 RSA-SHA256(PKCS#1) 签名。

    私钥不随仓库分发。默认从环境变量 TABLEFILLER_SIGNING_KEY 读取 XML 私钥文件路径，
    也可用 -KeyPath 显式指定。

.EXAMPLE
    pwsh -File installer\sign-release.ps1 -Version 1.1.0 -KeyPath D:\keys\tablefiller.xml
#>
param(
    [string]$Version = '1.1.0',
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputDir = (Join-Path (Split-Path -Parent $PSScriptRoot) 'dist'),
    [string]$KeyPath = $env:TABLEFILLER_SIGNING_KEY
)

$ErrorActionPreference = 'Stop'

if (-not $KeyPath -or -not (Test-Path -LiteralPath $KeyPath)) {
    throw '未找到签名私钥。请设置 TABLEFILLER_SIGNING_KEY 或传入 -KeyPath。'
}

$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot '.git'))) {
    throw "不是 git 仓库：$RepoRoot"
}

if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$commit = (& git -C $RepoRoot rev-parse HEAD).Trim()
$dirty = (& git -C $RepoRoot status --porcelain)
if ($dirty) {
    Write-Warning '工作区存在未提交改动；发布包只包含已提交内容（HEAD）。'
}

$zipPath = Join-Path $OutputDir ('app-' + $Version + '.zip')
$manifestPath = Join-Path $OutputDir ('manifest-' + $Version + '.json')
$sigPath = Join-Path $OutputDir ('manifest-' + $Version + '.sig')
Remove-Item -LiteralPath $zipPath, $manifestPath, $sigPath -Force -ErrorAction SilentlyContinue

Write-Host "[1/4] 打包提交 $commit …"
& git -C $RepoRoot archive --format=zip -o $zipPath HEAD
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $zipPath)) {
    throw 'git archive 失败。'
}

$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLower()
$size = (Get-Item -LiteralPath $zipPath).Length

Write-Host "[2/4] 生成清单 …"
$manifest = [ordered]@{
    version = $Version
    commit  = $commit
    created = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    archive = [ordered]@{
        name   = ('app-' + $Version + '.zip')
        sha256 = $hash
        size   = $size
    }
}
$json = $manifest | ConvertTo-Json -Depth 5
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($manifestPath, $json, $utf8)

Write-Host "[3/4] 使用私钥签名清单 …"
$privateXml = [System.IO.File]::ReadAllText($KeyPath)
$rsa = [System.Security.Cryptography.RSA]::Create()
$rsa.FromXmlString($privateXml)
$manifestBytes = [System.IO.File]::ReadAllBytes($manifestPath)
$signature = $rsa.SignData($manifestBytes,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
[System.IO.File]::WriteAllText($sigPath, [Convert]::ToBase64String($signature), $utf8)

Write-Host "[4/4] 用公钥回验签名 …"
$publicRsa = [System.Security.Cryptography.RSA]::Create()
$publicRsa.FromXmlString($rsa.ToXmlString($false))
if (-not $publicRsa.VerifyData($manifestBytes, $signature,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)) {
    throw '签名自检失败。'
}

Write-Host ''
Write-Host '发布文件已生成：'
Write-Host "  $zipPath"
Write-Host "  $manifestPath"
Write-Host "  $sigPath"
Write-Host ''
Write-Host "SHA256: $hash"
Write-Host "大小  : $size 字节"
Write-Host "提交  : $commit"
