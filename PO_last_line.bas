Attribute VB_Name = "PO_last_line"
Global SapGuiAuto As Object
Global Connection As Object
Global session As Object

Sub ME22N_Change()

    Set SapGuiAuto = GetObject("SAPGUI")  'Get the SAP GUI Scripting object
    Set SAPApp = SapGuiAuto.GetScriptingEngine 'Get the currently running SAP GUI
    Set SAPCon = SAPApp.Children(0) 'Get the first system that is currently connected
    Set session = SAPCon.Children(0) 'Get the first session (window) on that connection
    
    Dim i, PoLastLine As Long
    Dim ScrollDown As Double: ScrollDown = 1
    Dim Screen_no As String
    
    session.findById("wnd[0]").maximize
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/nme22n"
    session.findById("wnd[0]").sendVKey 0
    
    Screen_no = detect_screen_no(Screen_no, "wnd[0]/usr/subSUB0:SAPLMEGUI:", "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,1]")   'get new table number
    
    For i = 0 To 1000   'loop to see last used line in PO table
        If session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Screen_no & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,1]").Text = "" Then
            PoLastLine = i
            Exit For
        End If
        If i <> 0 Then
            session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Screen_no & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = ScrollDown
            ScrollDown = ScrollDown + 1
        End If
    Next i
    
    session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & Screen_no & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1," & PoLastLine + 1 & "]").Text = "30"
    
End Sub

Function detect_screen_no(Screen_no As String, str1 As String, str2 As String) As String
    On Error Resume Next
        session.findById (str1 & Screen_no & str2)
    If Err.Number = 0 Then
        detect_screen_no = Screen_no
    End If
    For i = 20 To 10 Step -1
        On Error Resume Next
           session.findById (str1 & "00" & CStr(i) & str2)
        If Err.Number = 0 Then
            detect_screen_no = "00" & CStr(i)
            Exit For
        End If
    Next i
    'detect_screen_no = ""
End Function
