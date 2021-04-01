Sub SAP_OpenSessionFromLogon()

Dim SapGuiAuto As Object

On Error Resume Next
 Set SapGuiAuto = GetObject("SAPGUI")
 Set App = SapGuiAuto.GetScriptingEngine
 
 
If App.Connections.Count() < 1 Then ' Identifica se aberto ou não

        Dim Applic
        Dim Connection
        Dim session
        Dim WSHShell
        
        Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
        Set WSHShell = CreateObject("WScript.Shell")
        Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
        Loop
        
        Set WSHShell = Nothing
        Set SapGui = GetObject("SAPGUI")
        Set Applic = SapGui.GetScriptingEngine
        Set Connection = Applic.OpenConnection("# -E05 - ECC - Produção - SSO", True)
        Set session = Connection.Children(0)
    Else
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
        session.findById("wnd[0]/tbar[0]/btn[12]").press
        session.findById("wnd[0]/tbar[0]/btn[12]").press
    End If

End Sub
