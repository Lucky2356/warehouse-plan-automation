param(
  [Parameter(Mandatory = $true)][string]$Tag
)
# Подписывает выпуск на GitHub: скачивает его .exe, подписывает ключом из хранилища Windows
# этого компьютера и прикладывает к выпуску «‹имя›.exe.sig».
# Запускать после того, как рабочий процесс Release создал выпуск. Без подписи программа
# обновление не ставит.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$work = Join-Path ([IO.Path]::GetTempPath()) ("sign-" + $Tag)

if (Test-Path $work) { Remove-Item -LiteralPath $work -Recurse -Force }
New-Item -ItemType Directory -Path $work | Out-Null

try {
  gh release download $Tag --pattern '*.exe' --dir $work
  $exe = Get-ChildItem -LiteralPath $work -Filter '*.exe' | Select-Object -First 1
  if (-not $exe) { throw "В выпуске $Tag нет .exe" }

  dotnet run --project (Join-Path $root 'tools/ReleaseSigner') -- sign $exe.FullName
  if ($LASTEXITCODE -ne 0) { throw 'Подписать не удалось' }

  dotnet run --no-build --project (Join-Path $root 'tools/ReleaseSigner') -- verify $exe.FullName
  if ($LASTEXITCODE -ne 0) { throw 'Подпись не сходится с ключом, встроенным в программу' }

  gh release upload $Tag ($exe.FullName + '.sig') --clobber
  "Выпуск $Tag подписан: $($exe.Name).sig"
}
finally {
  Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
