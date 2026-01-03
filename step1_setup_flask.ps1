# 礼簿管理系统 - Flask版本安装脚本 (PowerShell)

Write-Host "================================" -ForegroundColor Cyan
Write-Host "礼簿管理系统 - Flask版本安装" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# [1/3] 检查Python环境
Write-Host "[1/3] 检查Python环境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host $pythonVersion -ForegroundColor Green
    Write-Host "✓ Python环境正常" -ForegroundColor Green
} catch {
    Write-Host "! Python未安装或未添加到PATH" -ForegroundColor Red
    Write-Host "! 请先安装Python 3.7+" -ForegroundColor Red
    Write-Host ""
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# [2/3] 检查必要的文件夹
Write-Host "[2/3] 检查必要的文件夹..." -ForegroundColor Yellow
$allExist = $true
if (-not (Test-Path "templates")) {
    Write-Host "! templates 文件夹不存在" -ForegroundColor Red
    $allExist = $false
}
if (-not (Test-Path "static")) {
    Write-Host "! static 文件夹不存在" -ForegroundColor Red
    $allExist = $false
}
if (-not (Test-Path "src")) {
    Write-Host "! src 文件夹不存在" -ForegroundColor Red
    $allExist = $false
}
if ($allExist) {
    Write-Host "✓ 所有必要文件夹存在" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "! 缺少必要的文件夹，请检查项目完整性" -ForegroundColor Red
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# [3/3] 安装Python依赖
Write-Host "[3/3] 安装Python依赖..." -ForegroundColor Yellow
try {
    pip install -r config\requirements.txt
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ 依赖安装完成" -ForegroundColor Green
    } else {
        Write-Host "! 依赖安装失败" -ForegroundColor Red
        Write-Host ""
        Write-Host "按任意键退出..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
} catch {
    Write-Host "! 依赖安装失败: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# 完成
Write-Host "================================" -ForegroundColor Cyan
Write-Host "安装完成！" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "使用方法：" -ForegroundColor Yellow
Write-Host "1. 运行 .\step2_start_flask.ps1 启动服务器" -ForegroundColor White
Write-Host "2. 在浏览器中访问 http://127.0.0.1:5000" -ForegroundColor White
Write-Host ""
Write-Host "按任意键退出..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
