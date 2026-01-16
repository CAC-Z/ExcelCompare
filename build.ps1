Param(
  [ValidateSet('onedir','onefile')]
  [string]$Mode = 'onedir',
  [string]$Name = 'ExcelCompare',
  [switch]$Clean,
  [switch]$Splash,    # 为 onefile 提供闪屏，缓解等待感
  [switch]$NoUPX      # 关闭 UPX 压缩，提升 onefile 启动速度（体积更大）
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command pyinstaller -ErrorAction SilentlyContinue)) {
  Write-Error "PyInstaller 未安装。请先在虚拟环境内执行: pip install pyinstaller"
}

$commonArgs = @(
  '--noconfirm',
  '--windowed',            # GUI 应用
  '--name', $Name,
  '--icon', 'icons/icon.ico',
  '--add-data', 'icons;icons',  # 打包 icons 目录
  '--add-data', 'icons/icon.png;icons',   # 窗口图标
  '--add-data', 'icons/icon.ico;icons'
  
  # 收集三方库资源，避免运行时缺模块/数据文件
  ,'--collect-all','openpyxl'
  ,'--collect-submodules','pandas'
)

if ($Clean) { $commonArgs += '--clean' }

switch ($Mode) {
  'onefile' { $commonArgs += '--onefile' }
  default   { $commonArgs += '--onedir' }
}

if ($Mode -eq 'onefile') {
  if ($Splash) {
    $splashImg = (Test-Path 'icons/splash.png') ? 'icons/splash.png' : 'icons/icon.png'
    $commonArgs += @('--splash', $splashImg)
  }
  if ($NoUPX) { $commonArgs += '--noupx' }
}

Write-Host "pyinstaller $($commonArgs -join ' ') excel_compare.py" -ForegroundColor Cyan
pyinstaller @commonArgs 'excel_compare.py'

Write-Host "\n完成。产物位于 dist/$Name/ (onedir) 或 dist/$Name.exe (onefile)。" -ForegroundColor Green
