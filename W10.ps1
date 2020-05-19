if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

function Banner
{
    Write-Host "====================" -ForegroundColor Green
    Write-Host "|    Creado por    |" -ForegroundColor Green
    Write-Host "|      A. B.       |" -ForegroundColor Green
    Write-Host "|   t.me/audrum    |" -ForegroundColor Green
    Write-Host "====================" -ForegroundColor Green
    Write-Host ""
}

function ActivateWindows
{
    Write-Host "Detectando la versión de Windows..." -ForegroundColor Yellow
    $windows = (Get-WmiObject Win32_OperatingSystem).Caption
    Write-Host "Actualmente tiene" $windows "instalado" -ForegroundColor Green
    Start-Sleep -s 3
    Write-Host "Activando..." -ForegroundColor Yellow
    Start-Sleep -s 2

    switch ($windows)
    {
        'Microsoft Windows 10 Home Single Language'
        {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
            Start-Sleep -s 2
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Start-Sleep -s 2
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Start-Sleep -s 2
            Write-Host "Windows 10 Home Single Language ahora está activado!" -ForegroundColor Green
            Start-Sleep -s 5
        }

        'Microsoft Windows 10 Home'
        {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
            Start-Sleep -s 2
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Start-Sleep -s 2
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Start-Sleep -s 2
            Write-Host "Windows 10 Home ahora está activado!" -ForegroundColor Green
            Start-Sleep -s 5
        }

        'Microsoft Windows 10 Pro'
        {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
            Start-Sleep -s 2
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Start-Sleep -s 2
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Start-Sleep -s 2
            Write-Host "Windows 10 Pro ahora está activado!" -ForegroundColor Green
            Start-Sleep -s 5
        }
    }
}

Banner
ActivateWindows