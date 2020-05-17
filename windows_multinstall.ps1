function Banner
{
    Write-Host "====================" -ForegroundColor Green
    Write-Host "|    Creado por    |" -ForegroundColor Green
    Write-Host "|      A. B.       |" -ForegroundColor Green
    Write-Host "|   t.me/audrum    |" -ForegroundColor Green
    Write-Host "====================" -ForegroundColor Green
    Write-Host ""
    Write-Host "La mayoría de funciones de este script solo funcionan en Windows 10" -ForegroundColor Magenta
}

function Basic
{
    cls
    Write-Host "Verificando entorno..." -ForegroundColor Yellow
            
    if (Test-Path $env:SYSTEMDRIVE\ProgramData\Chocolatey) 
    {
        Write-Host "Entorno listo!" -ForegroundColor Green
        }
            
     else 
     {
        Write-Host "Preparando entorno..." -ForegroundColor Magenta
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        Write-Host "Entorno listo." -ForegroundColor Magenta
        }
            
      Write-Host "Instalando Google Chrome..." -ForegroundColor Yellow
      choco install GoogleChrome -y
      Write-Host "Instalación finalizada" -ForegroundColor Green
      Write-Host ""
      Write-Host "Instalando Firefox..." -ForegroundColor Yellow
      choco install Firefox -y
      Write-Host "Instalación finalizada" -ForegroundColor Green
      Write-Host ""
      Write-Host "Instalando 7-zip..." -ForegroundColor Yellow
      choco install 7zip -y
      Write-Host "Instalación finalizada" -ForegroundColor Green
      Write-Host "Fin de las instalaciones" -ForegroundColor Green
      Start-Sleep -s 5
}

function InstallOffice
{
    cls
    Write-Host "Verificando entorno..." -ForegroundColor Yellow
            
    if (Test-Path $env:SYSTEMDRIVE\ProgramData\Chocolatey) 
    {
        Write-Host "Entorno listo!" -ForegroundColor Green
        }
            
     else 
     {
        Write-Host "Preparando entorno..." -ForegroundColor Magenta
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        Write-Host "Entorno listo." -ForegroundColor Magenta
        }

    Write-Host "1. Descargar e instalar Office 365" -ForegroundColor Yellow
    Write-Host "2. Solo instalar" -ForegroundColor Yellow

    [int]$office = Read-Host "Seleccione una opción"

    switch ($office)
    {
        '1'{
            Write-Host "Iniciando descarga e instalación de Office 365..." -ForegroundColor Yellow
            choco install office365proplus -y
            Write-Host "Instalación  finalizada" -ForegroundColor Green
            Start-Sleep -s 5
            }

         '2'{
            Write-Host "Realice la instalación manualmente y regrese aquí para realizar la activación"
            Write-Host "Regresando al menu en 5 segundos..." -ForegroundColor Yellow
            Start-Sleep -s 5
            cls
            Menu
            }
    }
}

function ActivateWindows
{
    cls
    Write-Host "Detectando la versión de Windows..." -ForegroundColor Yellow
    $windows = (Get-WmiObject Win32_OperatingSystem).Caption
    Write-Host "Actualmente tiene" $windows "instalado"
    Write-Host "Activando..." -ForegroundColor Yellow

    switch ($windows)
    {
        'Microsoft Windows 10 Home Single Language'
        {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Write-Host "Proceso finalizado" -ForegroundColor Green
            Start-Sleep -s 5
            }

         'Microsoft Windows 10 Home'
         {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Write-Host "Proceso finalizado" -ForegroundColor Green
            Start-Sleep -s 5
            }

         'Microsoft Windows 10 Pro'
         {
            Write-Host "Instalando licencia..." -ForegroundColor Yellow
            slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
            Write-Host "Licencia instalada" -ForegroundColor Green
            Write-Host "Conectando con el servidor de activación..." -ForegroundColor Yellow
            slmgr /skms kms8.msguides.com
            Write-Host "Conexion realizada" -ForegroundColor Green 
            Write-Host "Activando" $windows -ForegroundColor Yellow
            slmgr /ato
            Write-Host "Proceso finalizado" -ForegroundColor Green
            Start-Sleep -s 5
            }
    }
}

function ActivateOffice
{
}

function End
{
    Write-Host "1. Regresar al menu"
    Write-Host "2. Salir"
    $pick = Read-Host "Que desea hacer?"

    switch ($pick)
    {
        '1'{ 
                Write-Host "Regresando al menu en 5 segundos..." -ForegroundColor Yellow
                Start-Sleep -s 5
                cls
                Menu
                }

        '2' {
                exit
                }
    }
}

function menu
{
    Banner
    Write-Host ""
    Write-Host "1. Instalar lo básico (Chrome, Firefox, 7-zip)" -ForegroundColor Yellow
    Write-Host "2. Instalar Office" -ForegroundColor Yellow
    Write-Host "3. Activar Windows 10" -ForegroundColor Yellow
    Write-Host "4. Activar Office" -ForegroundColor Yellow
    Write-Host "5. Ejecutar todas las acciones anteriores" -ForegroundColor Yellow
    Write-Host "6. Salir" -ForegroundColor Yellow
    Write-Host ""

    [int]$option = Read-Host "Seleccione una opción"
    
    switch ($option)
    {
        '1' {
                basic
            }

        '2' {
                InstallOffice
            }

        '3' {
                ActivateWindows
            }

        '4' {
                ActivateOffice
            }

        '5' {
                Basic
                InstallOffice
                ActivateWindows
                ActivateOffice
            }

        '6' {
                exit
            }

        }

}

menu  


