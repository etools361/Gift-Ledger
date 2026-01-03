# 礼簿管理系统 - Flask服务器启动脚本 (PowerShell)

Write-Host "================================" -ForegroundColor Cyan
Write-Host "礼簿管理系统 - Flask服务器" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "正在启动服务器..." -ForegroundColor Yellow
Write-Host "服务器地址: http://127.0.0.1:5000" -ForegroundColor Green
Write-Host ""
Write-Host "按 Ctrl+C 停止服务器" -ForegroundColor Yellow
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

try {
    python src\app_flask.py
} catch {
    Write-Host ""
    Write-Host "! 服务器启动失败" -ForegroundColor Red
    Write-Host "! 请检查：" -ForegroundColor Yellow
    Write-Host "  1. 是否已运行 setup_flask.ps1 安装" -ForegroundColor White
    Write-Host "  2. 是否已安装Python依赖" -ForegroundColor White
    Write-Host "  3. 端口5000是否被占用" -ForegroundColor White
    Write-Host ""
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
