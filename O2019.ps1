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

function ActivateOffice
{
    Write-Host "Activando Office..." -ForegroundColor Green
    Start-Sleep -s 3
    if(Test-Path "$env:ProgramFiles\Microsoft Office\Office16")
    {
        Set-Location "$env:ProgramFiles\Microsoft Office\Office16"

        Get-ChildItem "$env:ProgramFiles\Microsoft Office\root\Licenses16\" | Foreach-Object {
        if($_.Name.StartsWith('ProPlus2019VL'))
        {
            cscript ospp.vbs /inslic:"..\root\Licenses16\$_"
        }
        }

        cscript ospp.vbs /setprt:1688
        cscript ospp.vbs /unpkey:6MWKP >nul
        cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
        cscript ospp.vbs /sethst:kms8.msguides.com
        cscript ospp.vbs /act

        Write-Host "Office se ha activado satisfactoriamente" -ForegroundColor Green
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

        cscript ospp.vbs /setprt:1688
        cscript ospp.vbs /unpkey:6MWKP >nul
        cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
        cscript ospp.vbs /sethst:kms8.msguides.com
        cscript ospp.vbs /act

        Write-Host "Office se ha activado satisfactoriamente" -ForegroundColor Green
        Write-Host ""
        Start-Sleep -s 3
    }
}

Banner
ActivateOffice