<#
.SYNOPSIS
    安装"招采项目审核"技能到当前用户的 Claude Code 全局技能目录。
    安装后可在任何文件夹下使用该技能。

.DESCRIPTION
    将技能文件复制到 ~/.claude/skills/招采项目审核/
    全局技能在所有项目中均可触发，无需每个项目单独配置。

.EXAMPLE
    .\install.ps1
    安装到当前用户

.EXAMPLE
    .\install.ps1 -TargetUser "OtherUser"
    安装到指定用户（需要管理员权限）
#>

param(
    [string]$TargetUser = $env:USERNAME
)

$ErrorActionPreference = "Stop"

# 确定目标路径
if ($TargetUser -eq $env:USERNAME) {
    $destRoot = "$env:USERPROFILE\.claude\skills\招采项目审核"
} else {
    $userProfile = "C:\Users\$TargetUser"
    if (-not (Test-Path $userProfile)) {
        Write-Error "用户 $TargetUser 的目录不存在: $userProfile"
        exit 1
    }
    $destRoot = "$userProfile\.claude\skills\招采项目审核"
}

# 源目录 = 脚本所在目录的父目录
$srcRoot = Split-Path -Parent $PSScriptRoot

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  招采项目审核技能 - 安装程序" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "源目录 : $srcRoot" -ForegroundColor Gray
Write-Host "目标目录: $destRoot" -ForegroundColor Gray
Write-Host "目标用户: $TargetUser" -ForegroundColor Gray
Write-Host ""

# 创建目标目录
New-Item -ItemType Directory -Force -Path $destRoot | Out-Null

# 复制所有文件（排除 install.ps1 自身）
Get-ChildItem -Path $srcRoot -Recurse -File | Where-Object { $_.Name -ne 'install.ps1' } | ForEach-Object {
    $relativePath = $_.FullName.Replace($srcRoot, '').TrimStart('\')
    $destPath = Join-Path $destRoot $relativePath
    $destDir = Split-Path -Parent $destPath
    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Force -Path $destDir | Out-Null
    }
    Copy-Item -Path $_.FullName -Destination $destPath -Force
}

Write-Host ""
Write-Host "✅ 安装完成！" -ForegroundColor Green
Write-Host ""
Write-Host "已安装文件:" -ForegroundColor White
Get-ChildItem $destRoot -Recurse -File | ForEach-Object {
    $relPath = $_.FullName.Replace($destRoot, '').TrimStart('\')
    Write-Host "  $relPath" -ForegroundColor Gray
}

Write-Host ""
Write-Host "使用方法:" -ForegroundColor Yellow
Write-Host "  在任何 Claude Code 会话中输入以下任一关键词即可触发技能:" -ForegroundColor White
Write-Host "    - 审核招采项目" -ForegroundColor Cyan
Write-Host "    - 审核招标文件" -ForegroundColor Cyan
Write-Host "    - 审核投标文件" -ForegroundColor Cyan
Write-Host "    - 审核评标报告" -ForegroundColor Cyan
Write-Host "    - 政府采购审核" -ForegroundColor Cyan
Write-Host "    - 工程招标审核" -ForegroundColor Cyan
Write-Host "    - 串通投标检查" -ForegroundColor Cyan
Write-Host ""
Write-Host "其他用户安装:" -ForegroundColor Yellow
Write-Host "  .\install.ps1 -TargetUser '其他用户名'" -ForegroundColor White
