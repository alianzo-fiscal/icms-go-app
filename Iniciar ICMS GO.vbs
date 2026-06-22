Dim oShell, sDir

Set oShell = CreateObject("WScript.Shell")
sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

' Verifica atualizações (silenciosamente)
oShell.Run "cmd /c cd /d """ & sDir & """ && python atualizar.py", 0, True

' Inicia o app Streamlit em segundo plano
oShell.Run "cmd /c cd /d """ & sDir & """ && streamlit run app.py", 0, False

' Aguarda o servidor iniciar e abre o navegador
WScript.Sleep 4000
oShell.Run "http://localhost:8501"

Set oShell = Nothing
