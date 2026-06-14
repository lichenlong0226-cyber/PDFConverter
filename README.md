# OfficePDF

> Word / Excel → PDF 批量转换器 · PySide6 GUI

Office 文档转 PDF 的桌面工具，支持批量拖拽、多线程并发、PDF 合并、自动更新。

---

## ✨ 功能

| 功能 | 说明 |
|------|------|
| **批量转换** | 拖拽或添加多个 Word/Excel 文件，一键全部转 PDF |
| **多线程并发** | 最多 3 个任务同时转换，充分利用多核 CPU |
| **高保真输出** | Windows 下调用 Office COM 引擎，排版与原文档一致 |
| **LibreOffice 回退** | 无 Office 时自动使用 LibreOffice 转换 |
| **PDF 合并** | 转换完成后可将所有 PDF 合并为一个文件 |
| **混合支持** | 支持直接拖入已有 PDF，与转换结果一起合并 |
| **拖拽排序** | 文件列表可拖拽调整顺序，合并时保持页面顺序 |
| **SHA256 校验** | 自动更新下载的安装包经过校验，防止篡改 |
| **自动更新** | 内置更新检测，有新版本时一键下载安装 |

### 支持格式

**输入：** .doc .docx .docm .rtf .xls .xlsx .xlsm .xlsb .odt .ods .pdf
**输出：** .pdf

---

## 📦 下载安装

从 [Releases](https://github.com/lichenlong0226-cyber/pdf/releases) 下载最新安装包：

`
PDFConverter-setup-v1.1.6.exe
`

直接双击安装，桌面会生成快捷方式。

> ⚠ 若 Windows Smarrt App Control 拦截，右键安装包 → **属性** → **解除锁定** → 确定

---

## 🚀 使用

1. 启动程序
2. 将 Word/Excel/PDF 拖入文件列表，或点击「添加文件」
3. （可选）选择输出目录，留空则输出到桌面 PDFConverter_output 文件夹
4. （可选）勾选「合并为单个 PDF」
5. 点击「开始转换」
6. 完成后自动打开输出文件夹

---

## 🔧 本地开发

`ash
git clone https://github.com/lichenlong0226-cyber/pdf.git
cd pdf
python -m venv venv
venv\Scripts\activate
pip install PySide6 pypdf requests pywin32
python main.py
`

---

## 🏗 自动构建与发布

每次推送 * 格式的 tag，GitHub Actions 自动完成：

1. PyInstaller 打包为单文件 exe
2. NSIS 生成安装包
3. 计算 SHA256 校验文件
4. 创建 GitHub Release 并上传资产

`ash
# 推送代码
git push origin main

# 打标签触发构建
git tag v1.1.7
git push origin v1.1.7
`

---

## 🧱 技术栈

- Python 3.10 — 核心语言
- PySide6 — 桌面 GUI
- PyInstaller — 打包 exe
- NSIS — 安装包制作
- pypdf — PDF 合并
- pywin32 — Windows Office COM 调用
- GitHub Actions — CI/CD

---

## 📄 开源

MIT License
