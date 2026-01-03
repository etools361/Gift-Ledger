# 礼簿管理系统 - 打包为EXE

Write-Host "================================" -ForegroundColor Cyan
Write-Host "礼簿管理系统 - 打包为EXE" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# 检查Python环境
Write-Host "[1/6] 检查Python环境..." -ForegroundColor Yellow
try {
    $pythonVersion = python --version 2>&1
    Write-Host $pythonVersion -ForegroundColor Green
} catch {
    Write-Host "! Python未安装" -ForegroundColor Red
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# 安装依赖
Write-Host "[2/6] 安装必要的依赖..." -ForegroundColor Yellow
pip install Flask==3.0.0 openpyxl==3.1.2 pyinstaller Pillow
Write-Host "✓ 依赖安装完成" -ForegroundColor Green
Write-Host ""

# 转换Logo图标
Write-Host "[3/6] 准备应用图标..." -ForegroundColor Yellow
if (Test-Path "config\logo.png") {
    python -c "from PIL import Image; img = Image.open('config/logo.png'); img.save('config/logo.ico', format='ICO', sizes=[(256,256), (128,128), (64,64), (48,48), (32,32), (16,16)])"
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Logo已转换为ICO格式" -ForegroundColor Green
    } else {
        Write-Host "! Logo转换失败，将不使用自定义图标" -ForegroundColor Yellow
    }
} else {
    Write-Host "! 未找到 config\logo.png，将不使用自定义图标" -ForegroundColor Yellow
}
Write-Host ""

# 检查必要的文件夹
Write-Host "[4/6] 检查必要的文件夹..." -ForegroundColor Yellow
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

# 清理旧的打包文件
Write-Host "[5/6] 清理旧的打包文件..." -ForegroundColor Yellow
if (Test-Path "build") {
    Remove-Item -Path "build" -Recurse -Force
}
if (Test-Path "dist") {
    Remove-Item -Path "dist" -Recurse -Force
}
Write-Host "✓ 清理完成" -ForegroundColor Green
Write-Host ""

# 打包
Write-Host "[6/6] 开始打包（这可能需要几分钟）..." -ForegroundColor Yellow
Write-Host "正在打包，请耐心等待..." -ForegroundColor Cyan
pyinstaller --clean config\礼簿.spec

if ($LASTEXITCODE -eq 0) {
    Write-Host "✓ 打包完成！" -ForegroundColor Green
} else {
    Write-Host "! 打包失败" -ForegroundColor Red
    Write-Host "按任意键退出..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
Write-Host ""

# 完成
if (Test-Path "dist\礼簿管理系统.exe") {
    Write-Host "✓ EXE文件已生成" -ForegroundColor Green
    Write-Host ""
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host "打包成功！" -ForegroundColor Green
    Write-Host "================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "生成的文件：" -ForegroundColor Yellow
    Write-Host "  dist\礼簿管理系统.exe" -ForegroundColor White
    Write-Host ""
    Write-Host "使用方法：" -ForegroundColor Yellow
    Write-Host "1. 双击 dist\礼簿管理系统.exe 即可运行" -ForegroundColor White
    Write-Host "2. 浏览器会自动打开" -ForegroundColor White
    Write-Host "3. 数据保存在exe同目录下的 data.json 和 礼簿.xlsx" -ForegroundColor White
    Write-Host ""
    Write-Host "注意事项：" -ForegroundColor Yellow
    Write-Host "- 第一次运行可能会被防火墙拦截，请允许访问" -ForegroundColor White
    Write-Host "- 杀毒软件可能会误报，请添加信任" -ForegroundColor White
    Write-Host "- 可以将整个 dist 文件夹复制到其他电脑使用" -ForegroundColor White
    Write-Host ""
} else {
    Write-Host "! 未找到生成的EXE文件" -ForegroundColor Red
}

Write-Host "按任意键退出..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
