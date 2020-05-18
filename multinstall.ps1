function Banner
{
    Write-Host "====================" -ForegroundColor Green
    Write-Host "|    Creado por    |" -ForegroundColor Green
    Write-Host "|      A. B.       |" -ForegroundColor Green
    Write-Host "|   t.me/audrum    |" -ForegroundColor Green
    Write-Host "====================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Probado solo en Windows 10" -ForegroundColor Magenta
}

function ChocoEnvironment
{
    Write-Host "Verificando entorno..." -ForegroundColor Yellow
    
    $testchoco=choco -v
            
    if ($testchoco) 
    {
        Write-Host "Ya tiene instalado chocolatey versión $testchoco" -ForegroundColor Green
        }
            
     else 
     {
        Write-Host "Preparando entorno..." -ForegroundColor Magenta
        Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        refreshenv
        Write-Host "Entorno listo. Se instaló chocolatey versión $testchoco" -ForegroundColor Magenta
        Start-Sleep -s 3
        }
}


function InstallGoogleChrome
{
    if (-not(Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo)
    {
        $title    = 'Instalar Google Chrome'
        $question = 'Google Chrome no está instalado ¿desea instalarlo?'

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) 
        {
            Write-Host "Instalando Google Chrome..." -ForegroundColor Yellow
            cinst GoogleChrome -y
            Write-Host "Google Chrome se ha instalado" -ForegroundColor Green
            Start-Sleep -s 3
            Write-Host ""
        } 
        else 
        {
            Write-Host 'Ha decidido no instalar Google Chrome' -ForegroundColor Yellow
            Write-Host ""
            Start-Sleep -s 3
        }
    }

    else
    {
        Write-Host "Google Chrome ya está instalado, omitiendo instalación" -ForegroundColor Green
        Write-Host ""
        Start-Sleep -s 3
    }
}

function InstallFirefox
{
    if (-not(Test-Path "C:\Program Files\Mozilla Firefox\firefox.exe"))
    {
        $title    = 'Instalar Firefox'
        $question = 'Firefox no está instalado ¿desea instalarlo?'

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) 
        {
            Write-Host "Instalando Firefox..." -ForegroundColor Yellow
            cinst Firefox -y
            Write-Host "Firefox se ha instalado" -ForegroundColor Green
            Start-Sleep -s 3
            Write-Host ""
        } 
        else 
        {
            Write-Host "Ha decidido no instalar Firefox" -ForegroundColor Yellow
            Write-Host ""
            Start-Sleep -s 3
        }
    }

    else
    {
        Write-Host "Firefox ya está instalado, omitiendo instalación" -ForegroundColor Green
        Write-Host ""
        Start-Sleep -s 3
    }

}

function Install7-zip
{
    if (-not(Test-Path "C:\Program Files\7-Zip\7zFM.exe"))
    {
        $title    = "Instalar 7-zip"
        $question = "7-zip no está instalado ¿desea instalarlo?"

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) 
        {
            Write-Host "Instalando 7-zip..." -ForegroundColor Yellow
            cinst 7zip -y
            Write-Host "7-zip se ha instalado" -ForegroundColor Green
            Start-Sleep -s 3
            Write-Host ""
        } 
        else 
        {
            Write-Host "Ha decidido no instalar 7-zip" -ForegroundColor Yellow
            Write-Host ""
            Start-Sleep -s 3
        }
    }

    else
    {
        Write-Host "7-zip ya está instalado, omitiendo instalación" -ForegroundColor Green
        Write-Host ""
        Start-Sleep -s 3
    }
}

function InstallOffice365
{
    $title    = "Instalar Office 365"
    $question = "¿Descargar e instalar Office 365?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) 
    {
        Write-Host "Iniciando descarga e instalación de Office 365..." -ForegroundColor Yellow
        cinst office365proplus -y
        Write-Host "Instalación  finalizada" -ForegroundColor Green
        Start-Sleep -s 3
        Write-Host ""
    } 
        
    else 
    {
        Write-Host "Ha decidido no instalar Office 365" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -s 3
    }
}

function ActivateWindows10
{

    $title    = "Activar Windows 10"
    $question = "¿Desea ejecutar la activación de Windows 10?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) 
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
        
    else 
    {
        Write-Host "Ha decidido no activar Windows 10" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -s 3
    }

    
}

function ActivateOffice365
{
}

function menu
{
    cls
    Banner
    Write-Host ""
    Write-Host "1. Instalar Google Chrome" -ForegroundColor Yellow
    Write-Host "2. Instalar Firefox" -ForegroundColor Yellow
    Write-Host "3. Instalar 7-zip" -ForegroundColor Yellow
    Write-Host "4. Instalar Office" -ForegroundColor Yellow
    Write-Host "5. Activar Windows 10" -ForegroundColor Yellow
    Write-Host "6. Activar Office" -ForegroundColor Yellow
    Write-Host "7. Salir" -ForegroundColor Yellow
    Write-Host ""

    [int]$option = Read-Host "Seleccione la opción que desea ejecutar"
    
    switch ($option)
    {
        '1' {
                InstallGoogleChrome
            }

        '2' {
                InstallFirefox
            }

        '3' {
                Install7-zip
            }

        '4' {
                InstallOffice365
            }

        '5' {
                ActivateWindows10
            }

        '6' {
                ActivateOffice365
            }

        '7' {
                exit
            }

        }

}

function Option
{
    $title    = "Finalizado"
    $question = "Se ha completado el proceso, ¿quiere repetir alguna parte del proceso?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) 
    {
        menu
    } 
        
    else 
    {
        Write-Host "Saliendo..." -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -s 3
        exit
    }
}

Banner
ChocoEnvironment
InstallGoogleChrome
InstallFirefox
Install7-zip
InstallOffice365
ActivateWindows10
Option