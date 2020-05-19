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
    if (-not(Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe'))
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
    if (-not(Test-Path "$env:ProgramFiles\Mozilla Firefox\firefox.exe"))
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
    if (-not(Test-Path "$env:ProgramFiles\7-Zip\7zFM.exe"))
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

function InstallWinRAR
{
    if (-not(Test-Path "$env:ProgramFiles\WinRAR\WinRAR.exe"))
    {
        $title    = "Instalar WinRAR"
        $question = "WinRAR no está instalado ¿desea instalarlo?"

        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        if ($decision -eq 0) 
        {
            Write-Host "Instalando WinRAR..." -ForegroundColor Yellow
            cinst winrar -y
            Write-Host "WinRAR se ha instalado" -ForegroundColor Green
            Start-Sleep -s 3
            Write-Host ""
        } 
        else 
        {
            Write-Host "Ha decidido no instalar WinRAR" -ForegroundColor Yellow
            Write-Host ""
            Start-Sleep -s 3
        }
    }

    else
    {
        Write-Host "WinRAR ya está instalado, omitiendo instalación" -ForegroundColor Green
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

function ActivateOffice
{

    $title    = "Activar Office 365 o 2019"
    $question = "¿Descargar activar Office?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Si'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
    if ($decision -eq 0) 
    {
        if(Test-Path "$env:ProgramFiles\Microsoft Office\Office16")
        {
            Set-Location "$env:ProgramFiles\Microsoft Office\Office16"

            Get-ChildItem "$env:ProgramFiles\Microsoft Office\root\Licenses16\" | Foreach-Object {
            if($_.Name.StartsWith('ProPlus2019VL'))
            {
                cscript ospp.vbs /inslic:"..\root\Licenses16\$_"
            }
            }

            cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
            cscript ospp.vbs /unpkey:BTDRB >nul
            cscript ospp.vbs /unpkey:KHGM9 >nul
            cscript ospp.vbs /unpkey:CPQVG >nul
            cscript ospp.vbs /sethst:kms8.msguides.com
            cscript ospp.vbs /setprt:1688
            cscript ospp.vbs /act

            Write-Host "Office 365 se ha activado satisfactoriamente" -ForegroundColor Green
            Write-Host ""
            Start-Sleep -s 3
        }

        else
        {
            Set-Location "$env:ProgramFiles(x86)\Microsoft Office\Office16"

            Get-ChildItem "$env:ProgramFiles(x86)\Microsoft Office\root\Licenses16" | Foreach-Object {
            if($_.Name.StartsWith('ProPlus2019VL'))
            {
                cscript ospp.vbs /inslic:"..\root\Licenses16\$_"
            }
            }

            cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
            cscript ospp.vbs /unpkey:BTDRB >nul
            cscript ospp.vbs /unpkey:KHGM9 >nul
            cscript ospp.vbs /unpkey:CPQVG >nul
            cscript ospp.vbs /sethst:kms8.msguides.com
            cscript ospp.vbs /setprt:1688
            cscript ospp.vbs /act

            Write-Host "Office se ha activado satisfactoriamente" -ForegroundColor Green
            Write-Host ""
            Start-Sleep -s 3
          }
    } 
        
    else 
    {
        Write-Host "Ha decidido no activar Office" -ForegroundColor Yellow
        Write-Host ""
        Start-Sleep -s 3
    }
    
}

function menu
{
    cls
    Banner
    Write-Host ""
    Write-Host "1. Instalar Google Chrome" -ForegroundColor Yellow
    Write-Host "2. Instalar Firefox" -ForegroundColor Yellow
    Write-Host "3. Instalar 7-zip" -ForegroundColor Yellow
    Write-Host "4. Instalar WinRAR" -ForegroundColor Yellow
    Write-Host "5. Instalar Office" -ForegroundColor Yellow
    Write-Host "6. Activar Windows 10" -ForegroundColor Yellow
    Write-Host "7. Activar Office" -ForegroundColor Yellow
    Write-Host "8. Salir" -ForegroundColor Yellow
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
                InstallWinRAR
            }

        '5' {
                InstallOffice365
            }

        '6' {
                ActivateWindows10
            }

        '7' {
                ActivateOffice365
            }

        '8' {
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
InstallWinRAR
InstallOffice365
ActivateWindows10
ActivateOffice365
Option