$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$installerScript = Join-Path $projectRoot "installer\PnatClicker.iss"

Write-Host "Publishing Pnat Clicker (Release, win-x64)..."
dotnet publish "$projectRoot\ToDoList.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed"
}

$iscc = Get-Command iscc -ErrorAction SilentlyContinue
$isccPath = $null

if ($iscc) {
    $isccPath = $iscc.Source
} else {
    $defaultIsccPaths = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )

    foreach ($path in $defaultIsccPaths) {
        if (Test-Path $path) {
            $isccPath = $path
            break
        }
    }
}

if (-not $isccPath) {
    Write-Host "Inno Setup compiler (iscc) not found."
    Write-Host "Install Inno Setup from: https://jrsoftware.org/isinfo.php"
    Write-Host "After installing, run this script again to generate the installer .exe in .\dist"
    exit 1
}

Write-Host "Building installer..."
& $isccPath $installerScript

if ($LASTEXITCODE -ne 0) {
    throw "Installer build failed"
}

Write-Host "Done. Installer created in .\dist"