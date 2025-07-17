Write-Output "🔧 Iniciando remoção de Office (crackeado ou não)…"

# Passo 1 – Finaliza processos do Office
$procs = "winword","excel","powerpnt","outlook","onenote","teams","OfficeClickToRun"
foreach($p in $procs){
    Get-Process -Name $p -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# Passo 2 – Tenta remover produtos Click-to-Run via OfficeClickToRun.exe
$ctr = "C:\Program Files\Common Files\Microsoft Shared\Microsoft 365\OfficeClickToRun.exe"
if (!(Test-Path $ctr)) {
    $ctr = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
}

if (Test-Path $ctr) {
    Write-Output "✅ Encontrado ClickToRun em $ctr"
    $languages = @("en-us","es-es","fr-fr")  # ajuste conforme necessário
    foreach ($lang in $languages) {
        $products = "O365HomePremRetail","O365BusinessRetail","O365ProPlusRetail"
        foreach ($product in $products) {
            Write-Output "Tentando remover $product para $lang..."
            Start-Process -FilePath $ctr -ArgumentList "scenario=install","scenariosubtype=ARP","sourcetype=None","productstoremove=${product}.16_$lang_x-none","culture=$lang","DisplayLevel=False" -Wait -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
        }
    }
} else {
    Write-Output "⚠️ ClickToRun.exe não encontrado."
}

# Passo 3 – Remove por UninstallString no registro
$unregs = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* -ErrorAction SilentlyContinue | 
    Where-Object { $_.DisplayName -like "*Microsoft 365*" -or $_.DisplayName -like "*Microsoft Office*" }

foreach ($item in $unregs) {
    $str = $item.UninstallString
    if ($str) {
        Write-Output "⏳ Removendo via UninstallString: $($item.DisplayName)"
        if ($str -match '"([^"]+)"\s*(.*)') {
            $exe = $matches[1]
            $args = $matches[2] + " DisplayLevel=False"
            Start-Process -FilePath $exe -ArgumentList $args -Wait -ErrorAction SilentlyContinue
        } else {
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c",$str,"/qn" -Wait -ErrorAction SilentlyContinue
        }
    }
}

# Passo 4 – Limpeza manual de restos (cracks, pastas, tasks e registro)
$paths = @("C:\Program Files\KMSAuto*","C:\Windows\AutoKMS","C:\ProgramData\KMS*")
foreach ($p in $paths) {
    Try { Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue; Write-Output "🗑️ Limpeza: $p" } Catch {}
}
Get-ScheduledTask | Where-Object {$_.TaskName -like "*KMS*"} | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
foreach ($root in @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Run","HKCU:\Software\Microsoft\Windows\CurrentVersion\Run")) {
    Get-ItemProperty $root | ForEach-Object {
        $_.PSObject.Properties | Where-Object { $_.Value -match "kms" } | ForEach-Object {
            Remove-ItemProperty -Path $root -Name $_.Name -Force -ErrorAction SilentlyContinue
            Write-Output "🧹 Registro removido: $($_.Name)"
        }
    }
}

Write-Output "✅ Script concluído. Reinicie a máquina e verifique em 'Aplicativos e Recursos'."
