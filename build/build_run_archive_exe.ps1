$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location -LiteralPath $projectRoot

$python = if (Test-Path -LiteralPath ".build-venv\Scripts\python.exe") {
    ".build-venv\Scripts\python.exe"
} elseif (Test-Path -LiteralPath ".venv\Scripts\python.exe") {
    ".venv\Scripts\python.exe"
} else {
    "python"
}

& $python -c "import PyInstaller, PIL, fitz, yaml, docx, openpyxl"
if ($LASTEXITCODE -ne 0) {
    throw "缺少构建依赖，请先运行: $python -m pip install -r requirements-build.txt"
}

$appName = "质保作业申请归档"
$iconPath = Join-Path $projectRoot "assets\docflow_archive.ico"
$distDir = Join-Path $projectRoot "build\delivery"
New-Item -ItemType Directory -Path $distDir -Force | Out-Null
& $python -m PyInstaller `
    --onefile `
    --windowed `
    --clean `
    --noconfirm `
    --name $appName `
    --icon $iconPath `
    --add-data "$iconPath;assets" `
    --paths (Join-Path $projectRoot "src") `
    --distpath $distDir `
    --workpath (Join-Path $projectRoot "build\run_archive") `
    --specpath (Join-Path $projectRoot "build") `
    (Join-Path $projectRoot "run_archive.py")
if ($LASTEXITCODE -ne 0) {
    throw "EXE 构建失败，退出代码: $LASTEXITCODE"
}

$output = Join-Path $distDir "$appName.exe"
if (-not (Test-Path -LiteralPath $output)) {
    throw "构建完成但未找到输出文件: $output"
}
Write-Host "构建完成: $output"
