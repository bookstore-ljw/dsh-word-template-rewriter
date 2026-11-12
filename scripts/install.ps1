# 一键安装（Windows / PowerShell）
#
#   pwsh -ExecutionPolicy Bypass -File .\scripts\install.ps1
#   pwsh -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Editable -WithRender
#
# 做三件事：
#   1. 把本仓库装成 Python 包（提供 dsh-word 命令）；
#   2. 若检测到 DSH 配置目录，把 plugin/skills 下的技能软链/复制进去；
#   3. 跑一遍自检（命令可用、技能可读、单元测试）。
param(
    [switch]$Editable,       # pip install -e .（改代码即时生效）
    [switch]$WithRender,     # 额外安装 pywin32 / pymupdf（Word 版面复核）
    [string]$Python = "python",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
Write-Host "仓库：$root" -ForegroundColor Cyan

# ---- 1. Python 包 -----------------------------------------------------------
Write-Host "`n[1/3] 安装 Python 包" -ForegroundColor Cyan
if ($Editable -and $WithRender) {
    & $Python -m pip install -e ".[render]"
} elseif ($Editable) {
    & $Python -m pip install -e .
} elseif ($WithRender) {
    & $Python -m pip install ".[render]"
} else {
    & $Python -m pip install .
}
if ($LASTEXITCODE -ne 0) { throw "pip install 失败" }

# ---- 2. DSH 技能 ------------------------------------------------------------
Write-Host "`n[2/3] 注册 DSH 技能" -ForegroundColor Cyan
$skillsSrc = Join-Path $root "plugin\skills"
$profiles = Join-Path $env:USERPROFILE ".dsh\profiles"
$registered = 0
if (Test-Path $profiles) {
    foreach ($profile in Get-ChildItem $profiles -Directory) {
        $pluginLink = Join-Path $profile.FullName "node_modules\@local\dsh-word-pipeline"
        if (Test-Path $pluginLink) {
            Write-Host "  已通过插件链接可见：$($profile.Name)"
            $registered++
            continue
        }
        $skillTarget = Join-Path $profile.FullName "skills"
        New-Item -ItemType Directory -Force -Path $skillTarget | Out-Null
        foreach ($skill in Get-ChildItem $skillsSrc -Directory) {
            $dest = Join-Path $skillTarget $skill.Name
            if (-not (Test-Path $dest)) {
                Copy-Item $skill.FullName $dest -Recurse
                Write-Host "  复制技能 $($skill.Name) → $($profile.Name)"
                $registered++
            }
        }
    }
}
if ($registered -eq 0) {
    Write-Host "  未发现 DSH 配置目录，跳过（仅 Python CLI 可用）" -ForegroundColor Yellow
    Write-Host "  技能文件在：$skillsSrc（可手动放到你的 DSH skills 目录）" -ForegroundColor Yellow
}

# ---- 3. 自检 ---------------------------------------------------------------
Write-Host "`n[3/3] 自检" -ForegroundColor Cyan
& $Python -m dsh_word retypeset --help | Select-Object -First 3
if (-not $SkipTests) {
    & $Python -m unittest discover -s tests -p "test_*.py"
}
Write-Host "`n完成。用法示例：" -ForegroundColor Green
Write-Host '  dsh-word retypeset --template 模板.docx --source 原文.docx --out 成稿.docx --profile ieee-to-ttp --verify'
Write-Host '（在仓库目录内也可以用 python -m dsh_word 代替 dsh-word）'
