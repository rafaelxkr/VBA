Function logarSAP() As Boolean


    Set objShell = CreateObject("WScript.Shell")
    objShell.Run ("""C:\Program Files (x86)\SAP\FrontEnd\SAPgui\Saplogon.exe""")
    
    'Espera a janela aparecer
    'Do Until objShell.AppActivate("SAP Logon 740")
    '    WScript.Sleep 100
    'Loop
   ' Sleep 15000
'-------------------Aguarda 15 segundos---------------
newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 15
waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime
'---------------------------------------------

    
     
    'Encontrar a opção SAP Lar(SAP LAR R/3 Production)
    objShell.SendKeys "# -E05 - ECC - Produção - SSO", True 'Escolher qual Ambiente de deseja entrar
     
    'Pressiona o botão Logon (Alt + O)
    objShell.SendKeys "%{O}", True
             
    'Aguarda 8 segundos (abertura do R/3)
    'Sleep 8000
'-------------------Aguarda 8 segundos---------------
newHour = Hour(Now())
newMinute = Minute(Now())
newSecond = Second(Now()) + 8
waitTime = TimeSerial(newHour, newMinute, newSecond)
Application.Wait waitTime
'-------------------------------------------------------

    'Inicializa as variáveis de sessão
    If Not IsObject(SapApp) Then
       Set SapGuiAuto = GetObject("SAPGUI")
       Set SapApp = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = SapApp.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject SapApp, "on"
    End If
    
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "login" 'Colocar o login
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Senha" 'Colocar a senha
    session.findById("wnd[0]").sendVKey 0

'' Colocar o código aqui em embaixo

End Function