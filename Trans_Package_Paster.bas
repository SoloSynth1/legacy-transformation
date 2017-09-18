Attribute VB_Name = "Trans_Package_Paster"
Sub Trans_Package_Paster()
    On Error Resume Next
    Dim i, j, k As Integer
    Dim check As String
    
    Set WshShell = CreateObject("WScript.Shell")
    If Not IsObject(SAPApplication) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set SAPApplication = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(SAPConnection) Then
        Set SAPConnection = SAPApplication.Children(0)
    End If
    If Not IsObject(SAPsession) Then
        Set SAPsession = SAPConnection.Children(0)
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject SAPsession, "on"
        WScript.ConnectObject SAPApplication, "on"
    End If

    SAPsession.findById("wnd[0]").resizeWorkingPane 105, 31, False
    SAPsession.findById("wnd[0]/usr/btnTEXT_DRUCKTASTE_WP").press
    SAPsession.findById("wnd[0]/tbar[1]/btn[26]").press
    k = 1

    Do While IsEmpty(Cells(1, k)) = False

        i = 0
        j = 1
        
        SAPsession.findById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While IsEmpty(Cells(j, (k + 1))) = False
            If i = 19 Then
                SAPsession.findById("wnd[0]").sendVKey 0
                SAPsession.findById("wnd[0]/tbar[0]/btn[82]").press
                i = 0
            End If

            SAPsession.findById("wnd[0]/usr/tblSAPLCIDITCTRL_3000/txtRIEWP-KZYK1[0," & CStr(i) & "]").Text = Cells(j, (k + 1)).Value
            i = i + 1
            j = j + 1
        Loop

        SAPsession.findById("wnd[0]").sendVKey 0
        SAPsession.findById("wnd[0]/tbar[1]/btn[19]").press

        k = k + 2

    Loop

End Sub


