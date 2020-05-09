@echo off
:menu
cls
echo =============================
echo I                           I
echo I    Creado por A. B.       I
echo I    https://t.me/audrum    I
echo I                           I
echo =============================
echo 1. Activar Windows 10 Professional
echo 2. Activar Windows 10 Home Single Language
echo 3. Activar Windows 10 Home
echo 4. Activar otra version de Windows
echo 5. Mostrar listado de licencias
echo 6. Salir
REM set /p opcion=Digite una opcion:  
choice /c 12345 /n /m "Digite una opcion: "

IF ERRORLEVEL 6 EXIT
IF ERRORLEVEL 5 GOTO LIC
IF ERRORLEVEL 4 GOTO OTHER
IF ERRORLEVEL 3 GOTO W10H
IF ERRORLEVEL 2 GOTO W10HSL
IF ERRORLEVEL 1 GOTO W10P

:W10P
echo Instalando licencia...
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
echo Licencia instalada
echo Conectando con el servidor de activación...
slmgr /skms kms8.msguides.com
echo Coneccion realizada
echo Activando Windows 10 Professional
slmgr /ato
echo Proceso finalizado
pause > nul
cls
goto menu

:W10HSL
echo Instalando licencia...
slmgr /ipk 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
echo Licencia instalada
echo Conectando con el servidor de activación...
slmgr /skms kms8.msguides.com
echo Coneccion realizada
echo Activando Windows 10 Home Single Language
slmgr /ato
echo Proceso finalizado
pause > nul
cls
goto menu

:W10H
echo Instalando licencia...
slmgr /ipk TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
echo Licencia instalada
echo Conectando con el servidor de activación...
slmgr /skms kms8.msguides.com
echo Coneccion realizada
echo Activando Windows 10 Home
slmgr /ato
echo Proceso finalizado
pause > nul
cls
goto menu

:OTHER
set /p licence=Ingrese la licencia: 
echo Instalando licencia...
slmgr /ipk %licence%
echo Licencia instalada
echo Conectando con el servidor de activación...
slmgr /skms kms8.msguides.com
echo Coneccion realizada
echo Activando Windows 10 Home
slmgr /ato
echo Proceso finalizado
pause > nul
cls
goto menu

:LIC
echo =======================================
echo /                                     \
echo /   LICENCIAS DE PRODUCTOS MICROSOFT  \
echo /                                     \  
echo =======================================
echo Windows 10 Home: TX9XD-98N7V-6WMQ6-BX7FG-H8Q99
echo Windows 10 Home N: 3KHY7-WNT83-DGQKR-F7HPR-844BM
echo Windows 10 Home Single Language: 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH
echo Windows 10 Professional: W269N-WFGWX-YVC9B-4J6C9-T83GX
echo Windows 10 Professional N: MH37W-N47XK-V7XM9-C7227-GCQG9
echo Office 2019: NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP
echo Office 2016: XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
echo Windows 8: BN3D2-R7TKB-3YPBD-8DRP2-27GG4
echo Windows 8 Single Language: 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ
echo Windows 8 Professional: NG4HW-VH26C-733KW-K6F98-J8CK4
echo Windows 8 Professional N: XCVCF-2NXM9-723PB-MHCB7-2RYQQ
echo Windows 8 Professional WMC: GNBB8-YVD74-QJHX6-27H4K-8QHDG
echo Windows 8.1: M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK
echo Windows 8.1 N: 7B9N3-D94CG-YTVHR-QBPX3-RJP64
echo Windows 8.1 Single Language: BB6NG-PQ82V-VRDPW-8XVD2-V8P66
echo Windows 8.1 Professional: GCRJD-8NW9H-F2CDX-CCM8D-9D6T9
echo Windows 8.1 Professional N: HMCNV-VVBFX-7HMBH-CTY9B-B4FXY
echo Windows 8.1 Professional WMC: 789NJ-TQK6T-6XTH8-J39CJ-J8D3P
echo Office 2013: YC7DK-G2NP3-2QQC3-J6H88-GVGXT
echo Office 2010: VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB
pause > nul
choice /m "Crear un archivo con estas licencias?"

IF ERRORLEVEL 1 GOTO TXT
IF ERRORLEVEL 2 GOTO menu

:TXT
echo ======================================= > Licencias.txt
echo /                                     \ >> Licencias.txt
echo /   LICENCIAS DE PRODUCTOS MICROSOFT  \ >> Licencias.txt
echo /                                     \ >> Licencias.txt  
echo ======================================= >> Licencias.txt
echo Windows 10 Home: TX9XD-98N7V-6WMQ6-BX7FG-H8Q99 >> Licencias.txt
echo Windows 10 Home N: 3KHY7-WNT83-DGQKR-F7HPR-844BM >> Licencias.txt
echo Windows 10 Home Single Language: 7HNRX-D7KGG-3K4RQ-4WPJ4-YTDFH >> Licencias.txt
echo Windows 10 Professional: W269N-WFGWX-YVC9B-4J6C9-T83GX >> Licencias.txt
echo Windows 10 Professional N: MH37W-N47XK-V7XM9-C7227-GCQG9 >> Licencias.txt
echo Office 2019: NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP >> Licencias.txt
echo Office 2016: XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99 >> Licencias.txt
echo Windows 8: BN3D2-R7TKB-3YPBD-8DRP2-27GG4 >> Licencias.txt
echo Windows 8 Single Language: 2WN2H-YGCQR-KFX6K-CD6TF-84YXQ >> Licencias.txt
echo Windows 8 Professional: NG4HW-VH26C-733KW-K6F98-J8CK4 >> Licencias.txt
echo Windows 8 Professional N: XCVCF-2NXM9-723PB-MHCB7-2RYQQ >> Licencias.txt
echo Windows 8 Professional WMC: GNBB8-YVD74-QJHX6-27H4K-8QHDG >> Licencias.txt
echo Windows 8.1: M9Q9P-WNJJT-6PXPY-DWX8H-6XWKK >> Licencias.txt
echo Windows 8.1 N: 7B9N3-D94CG-YTVHR-QBPX3-RJP64 >> Licencias.txt
echo Windows 8.1 Single Language: BB6NG-PQ82V-VRDPW-8XVD2-V8P66 >> Licencias.txt
echo Windows 8.1 Professional: GCRJD-8NW9H-F2CDX-CCM8D-9D6T9 >> Licencias.txt
echo Windows 8.1 Professional N: HMCNV-VVBFX-7HMBH-CTY9B-B4FXY >> Licencias.txt
echo Windows 8.1 Professional WMC: 789NJ-TQK6T-6XTH8-J39CJ-J8D3P >> Licencias.txt
echo Office 2013: YC7DK-G2NP3-2QQC3-J6H88-GVGXT >> Licencias.txt
echo Office 2010: VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB >> Licencias.txt
echo Archivo Licencias.txt creado
@pause
goto menu