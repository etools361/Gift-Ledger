# 数据虚拟化脚本 - 生成示例数据

Write-Host "================================" -ForegroundColor Cyan
Write-Host "礼簿数据虚拟化工具" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "此工具将生成虚拟的礼簿数据，用于：" -ForegroundColor Yellow
Write-Host "  - 项目展示和演示" -ForegroundColor White
Write-Host "  - 功能截图说明" -ForegroundColor White
Write-Host "  - 公开发布和分享" -ForegroundColor White
Write-Host "  - 测试和开发" -ForegroundColor White
Write-Host ""

# 检查Python
Write-Host "检查Python环境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✓ $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "! Python未安装" -ForegroundColor Red
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# 检查openpyxl
Write-Host "检查openpyxl库..." -ForegroundColor Yellow
$hasOpenpyxl = python -c "import openpyxl" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "! openpyxl未安装，正在安装..." -ForegroundColor Yellow
    pip install openpyxl
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ openpyxl安装成功" -ForegroundColor Green
    } else {
        Write-Host "! openpyxl安装失败" -ForegroundColor Red
        Write-Host "按任意键退出..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
} else {
    Write-Host "✓ openpyxl已安装" -ForegroundColor Green
}
Write-Host ""

# 运行虚拟化脚本
Write-Host "================================" -ForegroundColor Cyan
Write-Host "开始生成虚拟数据" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

python tools\virtualize_data.py

Write-Host ""
Write-Host "================================" -ForegroundColor Cyan
Write-Host "按任意键退出..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
