# Copia a pasta "Projeto Vendas" para a SUA Area de Trabalho (funciona tambem com OneDrive).
$ErrorActionPreference = 'Stop'
$desktop = [Environment]::GetFolderPath('Desktop')
$origem  = $PSScriptRoot
$destino = Join-Path $desktop 'Projeto Vendas'
Write-Host "Area de Trabalho encontrada: $desktop"
if ($origem -eq $destino) { Write-Host 'A pasta ja esta na Area de Trabalho.'; exit 0 }
New-Item -ItemType Directory -Force -Path $destino | Out-Null
Copy-Item -Path (Join-Path $origem '*') -Destination $destino -Recurse -Force
Write-Host "Pronto! Pasta copiada para: $destino"
Start-Process explorer.exe $destino
