@echo off

IF EXIST C:\Python27\lib\site-packages\ copy /y src\OpenOPC.py C:\Python27\lib\site-packages\

IF NOT "%1"=="" REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_GATE_HOST /d %1 /f
IF NOT "%2"=="" REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_INACTIVE_TIMEOUT /d %2 /f
IF NOT "%3"=="" REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_MAX_CLIENTS /d %3 /f

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_CLASS
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_CLASS /d "Matrikon.OPC.Automation;Graybox.OPC.DAWrapper;HSCOPC.Automation;RSI.OPCAutomation;OPC.Automation"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_CLIENT
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_CLIENT /d "OpenOPC"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_MODE
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_MODE /d "dcom"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_HOST
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_HOST /d "localhost"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_SERVER
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_SERVER /d "Hci.TPNServer;HwHsc.OPCServer;opc.deltav.1;AIM.OPC.1;Yokogawa.ExaopcDAEXQ.1;OSI.DA.1;OPC.PHDServerDA.1;Aspen.Infoplus21_DA.1;National Instruments.OPCLabVIEW;RSLinx OPC Server;KEPware.KEPServerEx.V4;Matrikon.OPC.Simulation;Prosys.OPC.Simulation"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_GATE_HOST
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_GATE_HOST /d "0.0.0.0"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_GATE_PORT
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_GATE_PORT /d "7766"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_INACTIVE_TIMEOUT
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_INACTIVE_TIMEOUT /d "60"

REG QUERY "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_MAX_CLIENTS
if ERRORLEVEL 1 REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v OPC_MAX_CLIENTS /d "25"

regsvr32 /s lib\gbda_aut.dll

net stop zzzOpenOpcService
bin\OpenOPCService remove
bin\OpenOPCService.exe --startup auto install
net start zzzOpenOpcService
