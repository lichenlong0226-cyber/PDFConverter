# PDFConverter

Word/Excel -> PDF 转换器（PySide6 GUI）

快速上手
1. 在 main.py 顶部确认 GITHUB_OWNER 与 GITHUB_REPO 已设置为你的仓库（当前已设为 lichenlong0226-cyber/pdf）。
2. 在仓库根放入 app_icon.ico（可选），或删除 workflow 中的 --icon 参数。
3. 本地测试（可选）:
   - python -m venv venv
   - venv\\Scripts\\activate
   - pip install PySide6 pypdf requests pywin32
   - python main.py

自动构建与发布（Actions）
- 将这些文件提交并 push 到仓库，然后创建并 push 一个 tag（例如 v1.0.0）：
  git tag v1.0.0
  git push origin v1.0.0

- GitHub Actions 会在 windows-latest runner 上：
  - 用 PyInstaller 生成 single-file exe
  - 用 NSIS 生成安装器 MyConverter-setup-<version>.exe
  - 计算 SHA256，并把安装器与 .sha256 上传为 Release asset

自动更新（客户端）
- 程序会请求 GitHub Releases 的 latest，如果发现新版本会下载 installer + .sha256 做 SHA256 校验，校验通过则运行安装器。
