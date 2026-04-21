# SearchingPro Mac 安装说明

## 打包完成

您的 SearchingPro 应用程序已成功打包为 Mac 应用程序！

## 文件位置

- **应用程序文件**: `dist/SearchingPro.app`
- **应用程序大小**: 约 76MB
- **支持架构**: ARM64 (Apple Silicon)

## 安装方法

1. **直接使用**:
   - 双击 `dist/SearchingPro.app` 即可运行
   - 或者在终端中使用: `open dist/SearchingPro.app`

2. **安装到应用程序文件夹**:
   ```bash
   cp -R dist/SearchingPro.app /Applications/
   ```

3. **创建桌面快捷方式**:
   ```bash
   ln -s /Applications/SearchingPro.app ~/Desktop/SearchingPro.app
   ```

## 系统要求

- macOS 10.13 或更高版本
- Apple Silicon (M1/M2/M3) 或 Intel 处理器

## 功能特性

- 高级文件搜索功能
- 支持正则表达式搜索
- 文件内容索引和搜索
- 现代化的 PyQt5 用户界面
- 多线程搜索优化

## 注意事项

- 首次运行时，macOS 可能会显示安全警告
- 如果遇到"无法打开，因为它来自身份不明的开发者"，请在系统偏好设置 > 安全性与隐私中允许运行
- 或者在终端中运行: `xattr -cr /Applications/SearchingPro.app`

## 卸载

如需卸载，只需将 SearchingPro.app 移至废纸篓即可。

---

**打包信息**:
- 打包工具: PyInstaller 6.15.0
- Python 版本: 3.11+
- 打包时间: 2025年8月18日 (最新版本，修复了跨平台兼容性问题)
- 更新内容: 修复了Windows特定代码以实现完整的Mac兼容性