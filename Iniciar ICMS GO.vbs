' Plataforma ICMS/GO — Iniciador silencioso
Dim oShell, sDir

Set oShell = CreateObject("WScript.Shell")

' Pasta onde este .vbs está
sDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

' Inicia Streamlit em segundo plano (janela oculta)
oShell.Run "cmd /c cd /d """ & sDir & """ && streamlit run app.py", 0, False

' Aguarda o servidor iniciar
WScript.Sleep 4000

' Abre o navegador
oShell.Run "http://localhost:8501"

Set oShell = Nothing
