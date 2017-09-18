Attribute VB_Name = "Trans_Component_Copier"
Sub Trans_Component_Copier()
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
    SAPsession.findById("wnd[0]/usr/btnTEXT_DRUCKTASTE_MAT").press
    k = 1
    
    Do While 1 = 1
        
        Cells(1, k).Formula = "= """ & SAPsession.findById("wnd[0]/usr/txtPLPOD-VORNR").Text & """"
        
        i = 0
        j = 1
        
        SAPsession.findById("wnd[0]/tbar[0]/btn[80]").press
        
        Do While SAPsession.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i) & "]").Text <> ""
            If i = 20 Then
                check = SAPsession.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i - 1) & "]").Text
                SAPsession.findById("wnd[0]/tbar[0]/btn[82]").press
                If SAPsession.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i - 1) & "]").Text = check Then
                    Exit Do
                End If
                i = 0
            End If
        
            Cells(j, k + 1).Formula = "= """ & SAPsession.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/ctxtRIHSTPX-IDNRK[0," & CStr(i) & "]").Text & """"
            Cells(j, k + 2).Formula = "= """ & SAPsession.findById("wnd[0]/usr/tblSAPLCMDITCTRL_3500/txtRIHSTPX-MENGE[1," & CStr(i) & "]").Text & """"
            i = i + 1
            j = j + 1
        Loop
        
        check = SAPsession.findById("wnd[0]/usr/txtPLPOD-VORNR").Text
        SAPsession.findById("wnd[0]/tbar[1]/btn[19]").press
        
        If SAPsession.findById("wnd[0]/usr/txtPLPOD-VORNR").Text = check Then
            Exit Do
        End If
        k = k + 3
        
    Loop

End Sub

