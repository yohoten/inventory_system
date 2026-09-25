# 进销存管理系统

基于Python Tkinter的现代化进销存管理系统，增加了用户权限管理、数据备份、统计分析等高级功能。

## 功能特点

### 🎯 核心功能

- **进货管理**: 记录商品进货信息，自动更新库存
- **出货管理**: 管理商品出库，实时扣减库存
- **库存管理**: 实时查看当前库存状态
- **查询统计**: 按年份查询进货、出货记录和库存情况
- **数据导出**: 支持将查询结果导出为Excel格式
- **打印预览**: 提供数据打印预览功能

## 📁 项目结构

```
Inventory_System/
├── inventory_system.py           # 主程序
├── config.py                     # 系统配置模块
├── utils.py                      # 工具函数模块
├── system_config.json           # 系统配置文件
├── build.py                     # 完整打包脚本
├── simple_build.py              # 简化打包脚本
├── BUILD.bat                    # 批处理打包工具
├── START.bat                    # 启动脚本
└── README/
    ├── README.md                # 主要说明文档
    └── USAGE_GUIDE.md           # 使用指南
```

## 🚀 快速开始

### 开发环境运行
```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python inventory_system.py
```

### 打包为可执行文件
```bash
# 方法1: 图形化打包 (推荐)
双击 BUILD.bat

# 方法2: 完整打包 (功能最全)
python build.py --mode single --windowed --clean

# 方法3: 简化打包 (快速)
python simple_build.py
```
venv\Scripts\activate

### 直接运行
```bash
# Windows用户可以直接双击
START.bat
```

## 📦 打包说明

详细的打包指南请查看 [PACKAGING_README.md](../PACKAGING_README.md)

### 打包输出
- **单文件模式**: `dist/InventorySystem.exe` (推荐)
- **目录模式**: `dist/InventorySystem/` 目录

### 系统要求
- Windows 7/8/10/11
- 无需安装Python环境
- 建议4GB以上内存

## ⚙️ 系统配置

系统默认配置:
- 管理员账号: admin
- 默认密码: 123456
- 公司名称: 进销存管理系统
- 低库存阈值: 10件

配置文件: `system_config.json`

## 🔧 技术栈

- **语言**: Python 3.7+
- **GUI框架**: Tkinter
- **数据库**: SQLite3
- **数据处理**: openpyxl (Excel操作)
- **打包工具**: PyInstaller

## 📖 文档

- [使用指南](USAGE_GUIDE.md) - 详细功能说明
- [打包指南](../PACKAGING_README.md) - 打包部署说明
- [登录美化说明](LOGIN_BEAUTIFICATION.md) - 界面美化特性

## 🤝 贡献

欢迎提交Issue和Pull Request来改进项目。

## 📄 许可证

MIT License
# inventory_system
