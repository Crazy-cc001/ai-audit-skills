# 安装 pdf信息提取 技能所需的Python依赖
$ErrorActionPreference = "Continue"

Write-Host "检查Python环境..."
python --version
if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ 未找到Python。请先安装 Python 3.8+ 并勾选 'Add to PATH'。"
    exit 1
}

Write-Host "安装依赖: pdfplumber PyPDF2 pdf2docx python-docx pandas openpyxl"
python -m pip install -U pdfplumber PyPDF2 pdf2docx python-docx pandas openpyxl

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ 依赖安装完成"
} else {
    Write-Host "❌ 安装失败, 请检查网络或pip配置"
    exit 1
}
