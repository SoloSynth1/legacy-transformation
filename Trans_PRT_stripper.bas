Attribute VB_Name = "Trans_PRT_stripper"
Sub Trans_PRT_Stripper()
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
    SAPsession.findById("wnd[0]/usr/btnTEXT_DRUCKTASTE_FHM").press

    Do While SAPsession.findById("wnd[0]/usr/txtPLPOD-VORNR").Text <> check

        i = 0
        k = CInt(SAPsession.findById("wnd[0]/usr/txtRC27X-ENTRIES").Text)
        SAPsession.findById("wnd[0]/tbar[1]/btn[33]").press
        Do While i < k
            j = 1
            If SAPsession.findById("wnd[0]/usr/tblSAPLCFDITCTRL_0102/ctxtPLFHD-FHMAR[1," & CStr(i) & "]").Text <> "D" Then
                SAPsession.findById("wnd[0]/usr/tblSAPLCFDITCTRL_0102").getAbsoluteRow(i).Selected = False
            End If
            Do While IsEmpty(Cells(j, 1)) = False
                If CStr(Cells(j, 1).Value) = Left(SAPsession.findById("wnd[0]/usr/tblSAPLCFDITCTRL_0102/txtPLFHD-FHMNR[2," & CStr(i) & "]").Text, Len(CStr(Cells(j, 1).Value))) Then
                    SAPsession.findById("wnd[0]/usr/tblSAPLCFDITCTRL_0102").getAbsoluteRow(i).Selected = False
                    Exit Do
                End If
                j = j + 1
            Loop
            i = i + 1
        Loop
        check = SAPsession.findById("wnd[0]/usr/txtPLPOD-VORNR").Text
        SAPsession.findById("wnd[0]/tbar[1]/btn[14]").press
        SAPsession.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        SAPsession.findById("wnd[0]/tbar[1]/btn[19]").press
    Loop

End Sub


