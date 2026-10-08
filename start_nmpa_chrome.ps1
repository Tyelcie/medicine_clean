# 启动供药品网页保存脚本连接的可见 Chrome，验证由用户手动完成。
$ErrorActionPreference = 'Stop'
$chromeCandidates = @(
    'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    'C:\Program Files\Google\Chrome\Application\chrome.exe',
    (Join-Path $env:LOCALAPPDATA 'Google\Chrome\Application\chrome.exe')
)
$chromePath = $chromeCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $chromePath) { throw '未找到 Chrome，请检查安装位置。' }
$profilePath = Join-Path $PSScriptRoot '.nmpa-chrome-profile'
Start-Process -FilePath $chromePath -ArgumentList @(
    '--remote-debugging-address=127.0.0.1',
    '--remote-debugging-port=9222',
    ('--user-data-dir="' + $profilePath + '"'),
    'https://datasearch.nmpa.gov.cn/datasearch/home-index.html#category=yp'
)
Write-Output '已启动可见 Chrome，请手动完成网页验证。保持此窗口打开，再运行 Python 保存脚本。'
