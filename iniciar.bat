@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  py -3 -m venv .venv
  if errorlevel 1 goto erro
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto erro
.venv\Scripts\python.exe -m cadastro_pcm.main
if errorlevel 1 goto erro
exit /b 0
:erro
 echo Falha ao iniciar. Confira a mensagem acima.
 pause
 exit /b 1
