# 礼簿管理系统

一个专为中国传统礼节设计的现代化礼簿管理工具，采用传统竖排格式，喜庆红色主题，支持自动保存、Excel导入导出和打印功能。

## ✨ 主要特性

- 🎨 **传统礼簿格式** - 竖排文字，中国红配色，还原传统纸质礼簿
- 💾 **自动保存** - 每次操作自动保存到 JSON 和 Excel 文件
- 📂 **自动加载** - 启动时自动从 `礼簿.xlsx` 加载已有数据
- 📊 **实时统计** - 自动计算总金额、微信/现金/支付宝分类统计
- 🔤 **中文大写** - 金额自动转换为中文大写（壹贰叁肆...）
- 📱 **双视图模式** - 传统视图（展示）+ 编辑视图（管理）
- 🖨️ **打印导出** - 支持导出打印页和 Excel 文件
- 💰 **支付方式** - 支持微信、支付宝、现金记录，带图标显示
- 🔒 **封面保护** - 可显示封面图片保护隐私信息
- 🎭 **数据虚拟化** - 生成演示数据，保护隐私安全

## 样式
### 封面
![](./src/fig1.png)
### 正文
![](./src/fig2.png)
### 金额统计
![](./src/fig3.png)
### 可打印为pdf
![](./src/fig4.png)

## 🚀 快速开始

### 方式1：Flask服务器（推荐）⭐⭐⭐⭐⭐

适合：日常使用、有Python环境、需要局域网访问、频繁修改代码

```powershell
# 1. 安装（仅首次）
.\step1_setup_flask.ps1

# 2. 启动
.\step2_start_flask.ps1

# 3. 浏览器访问
http://127.0.0.1:5000
```

**优点**：
- ✅ 文件小，不占空间
- ✅ 支持局域网多设备访问
- ✅ 方便调试和修改
- ✅ 自动保存数据到本地文件

### 方式2：双击运行EXE ⭐⭐⭐⭐

适合：分发给他人、不想安装Python、需要像普通软件一样使用

```powershell
# 1. 打包（仅首次需要，需要Python环境）
.\step4_build_exe.ps1

# 2. 运行
双击 dist\礼簿管理系统.exe
```

**优点**：
- ✅ 无需安装Python，双击即用
- ✅ 可以分发给他人使用
- ✅ 像普通软件一样简单
- ✅ 自动保存数据

**注意**：
- 打包后文件约27MB
- 首次运行可能被防火墙拦截，请允许访问
- 杀毒软件可能误报，请添加信任

## 📋 功能清单

### 数据管理
- ✅ **启动自动加载** - 程序启动时自动从 `礼簿.xlsx` 加载数据
- ✅ 添加/编辑/删除记录
- ✅ 姓名重复检测
- ✅ 支付方式分类（微信/支付宝/现金/未记录）
- ✅ 自动保存到本地文件（JSON + Excel双格式）
- ✅ 数据虚拟化工具（生成演示数据）

### 显示功能
- ✅ 传统竖排格式（12条/页）
- ✅ 姓名括号备注显示（如：张三(爸爸)）
- ✅ 金额阿拉伯数字可选显示
- ✅ 支付方式图标显示
- ✅ 封面图片隐藏/显示

### 统计功能
- ✅ 总记录数统计
- ✅ 微信/现金/支付宝分类统计
- ✅ 每页金额汇总
- ✅ 中文大写金额自动转换

### 导入导出
- ✅ Excel 文件导入导出
- ✅ JSON 文件导入导出
- ✅ 打印页面导出（含颜色）
- ✅ 多页汇总打印

## 📊 使用方式对比

| 特性 | Flask版 | EXE版 |
|------|---------|-------|
| 启动方式 | 运行脚本 | 双击exe ✅ |
| 自动保存 | ✅ 完全 | ✅ 完全 |
| 自动加载 | ✅ 完全 | ✅ 完全 |
| 分发性 | ⭐ | ⭐⭐⭐⭐⭐ |
| 安装要求 | Python | 无 ✅ |
| 文件大小 | 小 ✅ | 中等(27MB) |
| 局域网访问 | ✅ 支持 | ⚠️ 需配置 |
| 开发调试 | ✅ 方便 | ❌ 不便 |

## 💻 技术栈

- **前端**：HTML5 + CSS3 + JavaScript (ES6)
- **后端**：Python 3.7+ + Flask 3.0.0
- **数据存储**：JSON + Excel (openpyxl 3.1.2)
- **打包工具**：PyInstaller 6.3.0

## 📁 文件说明

### 核心文件
- `src/app_flask.py` - Flask服务器
- `src/app_exe.py` - EXE版本后端
- `templates/index.html` - HTML模板文件
- `static/app.js` - JavaScript文件
- `static/style.css` - 样式文件

### 工具模块
- `tools/virtualize_data.py` - 数据虚拟化工具

### 数据文件（自动生成）
- `data/data.json` - JSON格式数据
- `data/礼簿.xlsx` - Excel格式数据
- `data/礼簿_示例数据.xlsx` - 虚拟示例数据
- `data/data_示例数据.json` - 虚拟示例数据

### 脚本和配置
- `step1_setup_flask.ps1` - Flask安装脚本
- `step2_start_flask.ps1` - Flask启动脚本
- `step3_generate_demo_data.ps1` - 虚拟数据生成脚本
- `step4_build_exe.ps1` - EXE打包脚本
- `config/礼簿.spec` - PyInstaller配置
- `config/requirements.txt` - Python依赖
- `config/logo.png` - 应用图标源文件
- `config/logo.ico` - EXE图标文件（自动生成）

### 资源文件
- `static/weixin.png` - 微信图标
- `static/favicon.ico` - 支付宝图标
- `static/top.jpg` - 封面图片

## 🎯 快速入门

### 1. 安装和启动

**Flask版：**
```powershell
# 安装
.\step1_setup_flask.ps1

# 启动
.\step2_start_flask.ps1

# 浏览器自动打开 http://127.0.0.1:5000
```

**EXE版：**
```powershell
# 打包（首次）
.\step4_build_exe.ps1

# 运行
双击 dist\礼簿管理系统.exe
```

### 2. 数据加载

**程序启动时会自动加载数据：**
1. 优先从 `data/礼簿.xlsx` 加载（如果存在）
2. 如果Excel不存在，从 `data/data.json` 加载
3. 都不存在则从空白开始

**使用虚拟数据快速体验：**
```powershell
# 1. 生成虚拟数据
.\step3_generate_demo_data.ps1

# 2. 将示例数据重命名为礼簿.xlsx
ren data\礼簿_示例数据.xlsx 礼簿.xlsx

# 3. 启动程序（自动加载虚拟数据）
.\step2_start_flask.ps1
# 或
双击 dist\礼簿管理系统.exe
```

### 3. 添加记录

1. 填写姓名、礼金、支付方式
2. 点击"添加记录"
3. 自动保存，自动跳转到新记录页

### 4. 查看统计

1. 点击"展开汇总信息"查看总统计
2. 点击"展开本页汇总"查看当前页统计
3. 切换到"编辑模式"查看详细列表

### 5. 导出打印

1. 点击"导出打印页"按钮
2. 浏览器自动打开打印预览
3. 可保存为PDF或直接打印

## 🎭 数据虚拟化工具

### 功能说明

生成虚拟的礼簿示例数据，替换真实的姓名和金额信息，用于：

- ✅ **项目展示** - 展示项目功能时保护隐私
- ✅ **功能演示** - 演示软件功能的示例数据
- ✅ **截图说明** - 制作使用教程的截图
- ✅ **公开发布** - GitHub、博客等公开发布
- ✅ **测试开发** - 开发测试时的模拟数据
- ✅ **快速体验** - 生成数据后程序启动时自动加载

### 虚拟数据特点

**姓名虚拟化：**
- 使用常见中文姓氏（张、王、李、刘等30个）
- 80%生成3个字的名字，20%生成2个字的名字
- 50%男性名字，50%女性名字
- 30%概率添加备注，性别智能匹配
- 智能避免唯一性亲戚重复（爸爸、妈妈、大伯等）
- 示例：`张秀英(妈妈)`、`李建国(爸爸)`、`王强(同学)`

**金额虚拟化：**
- 所有金额都是100元的整数倍（百元整）
- 70%使用常见整数金额（100、200、500、1000等）
- 30%使用范围随机金额（按百元递增）
- 金额范围：100-5000元

**支付方式分布：**
- 微信：50% | 支付宝：25% | 现金：20% | 未记录：5%

### 使用方法

**方法1：一键生成（推荐）**
```powershell
.\step3_generate_demo_data.ps1
# 输入要生成的记录数（默认50）
```

**方法2：Python脚本**
```powershell
python tools\virtualize_data.py
```

**方法3：快速体验完整功能**
```powershell
# 一键体验
.\step3_generate_demo_data.ps1    # 生成50条虚拟数据
ren data\礼簿_示例数据.xlsx 礼簿.xlsx
.\step2_start_flask.ps1           # 启动系统，自动加载
```

### 生成的文件

- `data/礼簿_示例数据.xlsx` - Excel格式虚拟数据
- `data/data_示例数据.json` - JSON格式虚拟数据

**注意：**
- 生成的文件名包含"示例数据"，不会覆盖真实数据
- 重命名为 `data/礼簿.xlsx` 后程序启动时会自动加载
- 每次运行生成的数据都不同

## ⚙️ 配置说明

### 修改每页记录数

在网页中点击"参数设置"按钮，修改"每页记录数"（1-50条）

### 修改端口（Flask/EXE版）

编辑 `src/app_flask.py` 或 `src/app_exe.py` 最后一行：
```python
app.run(debug=False, host='127.0.0.1', port=5000)  # 修改端口号
```

### 局域网访问（Flask/EXE版）

```python
app.run(debug=False, host='0.0.0.0', port=5000)  # 改为0.0.0.0
```
其他设备访问：`http://[服务器IP]:5000`

## 📝 使用场景

- 🎊 婚礼现场记录礼金
- 🎂 寿宴、满月酒等喜事
- 🏠 乔迁、开业等庆典
- 💝 其他需要记录人情往来的场合

## ⚠️ 注意事项

### 数据安全
- ✅ **定期备份** `data/data.json` 和 `data/礼簿.xlsx`
- ✅ 重要场合建议提前测试
- ✅ 使用Flask版或EXE版确保数据不丢失

### Flask版注意
- 需要 Python 3.7+ 环境
- 首次需要安装依赖：`pip install -r config\requirements.txt`
- 检查端口5000是否被占用

### EXE版注意
- 打包需要 Python 3.7+ 环境
- 打包时间约3-5分钟
- 生成的exe文件约27MB（包含完整运行环境）
- 首次运行可能被防火墙拦截，请允许访问
- 杀毒软件可能误报，请添加信任
- EXE自带自定义图标（可通过config/logo.png更换）

### 浏览器兼容
- 推荐使用 Chrome、Edge、Firefox 等现代浏览器
- 需要启用JavaScript

## 🔧 故障排除

### 自定义应用图标

如果要更换EXE的图标，只需替换 `config/logo.png` 文件即可：

1. 准备一个PNG格式的图标（建议尺寸：256x256 或更大）
2. 将其命名为 `logo.png` 并放入 `config/` 文件夹
3. 运行 `.\step4_build_exe.ps1` 打包时会自动转换为ICO格式

**注意：**
- 打包脚本会自动安装Pillow库用于图标转换
- 如果图标不显示，运行根目录的清理脚本（见下方）

### Windows图标缓存问题

如果EXE图标在资源管理器中没有更新，这是Windows缓存问题，不是打包问题。

**解决方法：**
1. 注销并重新登录Windows（最简单）
2. 重启电脑
3. 或参考项目中的 `README_图标问题.md` 详细说明

### Flask版启动失败

1. 确认Python版本：`python --version`（需要3.7+）
2. 安装依赖：`pip install -r config\requirements.txt`
3. 检查端口5000是否被占用：`netstat -ano | findstr :5000`
4. 查看控制台错误信息

### EXE版无法启动

1. 检查是否被杀毒软件拦截
2. 尝试以管理员身份运行
3. 查看命令行窗口的错误信息
4. 确保 `templates` 和 `static` 文件夹在exe同目录

### 数据无法保存

1. 检查文件夹权限
2. 确保磁盘空间充足
3. 查看控制台错误信息

### 数据无法加载

1. 检查 `data/礼簿.xlsx` 格式是否正确
2. 确保"汇总"工作表存在
3. 查看控制台加载信息

## 📚 开发说明

### 项目结构

```
礼簿/
├── step1_setup_flask.ps1         # Flask安装脚本
├── step2_start_flask.ps1         # Flask启动脚本
├── step3_generate_demo_data.ps1  # 虚拟数据生成脚本
├── step4_build_exe.ps1           # EXE打包脚本
├── README.md                     # 项目文档
├── src/                          # 源代码文件夹
│   ├── app_flask.py              # Flask服务器
│   └── app_exe.py                # EXE版服务器
├── tools/                        # 工具模块文件夹
│   └── virtualize_data.py        # 数据虚拟化工具
├── config/                       # 配置文件夹
│   ├── 礼簿.spec                  # PyInstaller配置
│   ├── requirements.txt          # Python依赖
│   ├── logo.png                  # 应用图标（源文件）
│   └── logo.ico                  # 应用图标（exe用）
├── data/                         # 数据文件夹（自动生成）
│   ├── data.json                 # JSON格式数据
│   ├── 礼簿.xlsx                  # Excel数据
│   ├── data_示例数据.json         # 虚拟示例数据
│   └── 礼簿_示例数据.xlsx         # 虚拟示例数据
├── templates/                    # Flask模板文件夹
│   └── index.html                # Flask HTML模板
├── static/                       # 静态资源文件夹
│   ├── app.js                    # Flask JS文件
│   ├── style.css                 # 样式文件
│   ├── weixin.png                # 微信图标
│   ├── favicon.ico               # 支付宝图标
│   └── top.jpg                   # 封面图片
└── dist/                         # 打包输出（运行step4后生成）
    └── 礼簿管理系统.exe          # 打包的EXE文件（27MB）
```

### 修改前端

直接编辑以下文件：
- `templates/index.html` - HTML模板
- `static/app.js` - JavaScript代码
- `static/style.css` - 样式

修改后重启Flask服务器即可看到效果。

### 修改后端

1. 编辑 `src/app_flask.py` 或 `src/app_exe.py`
2. 重启服务器或重新打包EXE

### 添加新功能

1. 在 `static/app.js` 中添加前端逻辑
2. 在 `src/app_flask.py` 中添加API端点
3. 在 `templates/index.html` 中添加UI元素
4. 更新文档说明

## 📄 许可证

本项目为个人使用工具，免费开源。

## 🎉 更新日志

### v4.0 (2026-01-02) - 文件结构重组 + 增强功能
- ✅ 新增数据虚拟化工具
- ✅ 智能性别匹配的虚拟名字
- ✅ 唯一性亲戚关系控制
- ✅ 启动时自动从Excel加载数据
- ✅ 项目文件结构重组（src/tools/config/data分文件夹管理）
- ✅ 自定义Logo图标支持（config/logo.png）
- ✅ EXE自动嵌入图标，支持一键更换
- ✅ 优化文档结构和开发流程
- ✅ 删除文件复制步骤，直接使用最终文件
- ✅ 清理重复文件，保持根目录整洁

### v3.0 (2026-01-01) - EXE版本
- ✅ 新增EXE打包支持，双击即用
- ✅ 自动打开浏览器
- ✅ 数据自动保存到文件

### v2.0 (2025-12-29) - Flask版本
- ✅ 使用Flask后端实现真正的文件自动保存
- ✅ 支持局域网多设备访问
- ✅ 保留所有前端功能

### v1.0 (2025-12-29) - 初始版本
- ✅ 实现基础数据录入功能
- ✅ 传统礼簿格式显示
- ✅ Excel导入导出
- ✅ 中文大写转换

---

## 🚀 推荐使用流程

### 首次使用

```powershell
# 1. 安装Flask版（推荐）
.\step1_setup_flask.ps1

# 2. 生成虚拟数据快速体验
.\step3_generate_demo_data.ps1
ren data\礼簿_示例数据.xlsx 礼簿.xlsx

# 3. 启动系统
.\step2_start_flask.ps1

# 4. 体验完整功能！
```

### 日常使用

```powershell
# 直接启动
.\step2_start_flask.ps1
```

### 分发给他人

```powershell
# 1. 打包EXE
.\step4_build_exe.ps1

# 2. 复制dist文件夹给他人
# 3. 对方双击 礼簿管理系统.exe 即可使用
```

---

**享受传统礼簿的数字化管理！** 🎊
