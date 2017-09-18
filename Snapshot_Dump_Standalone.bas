Attribute VB_Name = "Snapshot_Dump_Standalone"
'06 Mar 2017: RC1

Sub main()
    Dim path, file As String
    
    path = Snapshot_Dump_Standalone.get_folder()
    xls_path = path & "\xls"
    xlsx_path = path & "\xlsx"
    
    If Len(Dir(xls_path, vbDirectory)) = 0 Then
        MkDir xls_path
    End If
    
    file = Dir(xls_path & "\*.XLS?")

    If file <> "" Then
        MsgBox ".xls or .xlsx files exist at the directory: " & xls_path & "." & vbCrLf & "Aborting.", vbExclamation
        Exit Sub
    End If
    
    Do While file <> ""
        file = Dir()
    Loop
    
    If MsgBox("This action will take extremely long time (~90mins),  are you sure to proceed?", vbOKCancel) = vbOK Then
        Set mysap = getSession()
        dump_plans mysap, Split("H00 H01 H03 H04 H07 H08 HX0"), xls_path
        dump_plans mysap, Split("HI"), xls_path
        xls_to_xlsx path, xls_path, xlsx_path
        
        Kill xls_path & "\*.*"
        RmDir xls_path
        
        Shell "C:\Windows\explorer.exe """ & path & "", vbNormalFocus
    Else
        MsgBox "Action aborted.", vbExclamation
        Exit Sub
    End If
End Sub

Function dump_plans(mysap, ItemList, path)

    On Error Resume Next
    Application.DisplayAlerts = False
    
    For Each item In ItemList
        For i = 0 To 900 Step 100
            j = i + 99
            If item = "HI" Then
                plan1 = "HI*"
                plan2 = ""
                filename = "HI_.XLS"
                i = 900
            Else
                plan1 = item & format(CStr(i), "000")
                plan2 = item & format(CStr(j), "000")
                filename = plan1 & "-" & plan2 & ".XLS"
            End If
            
            With mysap
                .SendCommand ("/nia17")
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-LOW").text = plan1
                .FindById("wnd[0]/usr/ctxtPN_PLNNR-HIGH").text = plan2
                .FindById("wnd[0]/usr/ctxtPN_WERKS-LOW").text = "HK01"
                .FindById("wnd[0]").SendVKey 8
                .FindById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
                If Err.Number = 0 Then
                    .FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                    .FindById("wnd[1]/usr/ctxtDY_PATH").text = path
                    .FindById("wnd[1]/usr/ctxtDY_FILENAME").text = filename
                    .FindById("wnd[1]/tbar[0]/btn[0]").press
                Else
                    Err.Number = 0
                End If
            End With
        Next i
    Next
    
    Application.DisplayAlerts = True
    On Error GoTo 0
    
End Function

Function xls_to_xlsx(ByVal path As String, ByVal xls_path As String, ByVal xlsx_path As String)
    Dim file As String
    Dim a As Object
    
    If Len(Dir(path & "\xlsx", vbDirectory)) = 0 Then
        MkDir path & "\xlsx"
    End If
    
    file = Dir(xls_path & "\*.XLS")
    
    Do While file <> ""
        Workbooks.Open (xls_path & "\" & file)
        Set a = ActiveWorkbook
        NewFile = Left(file, InStr(file, ".")) & "xlsx"
        
        a.SaveAs filename:=(xlsx_path & "\" & NewFile), FileFormat:=xlOpenXMLWorkbook
        a.Close
        file = Dir()
        Set a = Nothing
    Loop
    
End Function

Sub shutdown_after_dump()   'Does what the title say
    main
    Shutdown
End Sub

' getSession() - Get SAPSession object.
' Optional variable "wnd": Selecting which SAPSession (0-5) to grab. Default is the first opened session (0). Will create new ones if the specified session is not opened.
' Optional variable "transaction": Check if the transaction code of the session grabbed matches the variable. This is used in conjuction with "NewSession".
' Optional variable "NewSession": This Boolean determines what will happen when the transcation of the grabbed SAPSession doesn't match the "transaction" variable.
'   When NewSession = False, Error message will pop up and the code would stop itself.
'   When NewSession = True, the code will send a command to change the transaction of the grabbed session to that specified by the "transaction" variable.

Public Function getSession(Optional wnd As Long = -1, Optional transaction As String = "", Optional NewSession As Boolean = False) As Object
    On Error Resume Next
    Do While 1
        Set SapGuiAuto = GetObject("SAPGUI")    'Get Built-in Object "SAPGUI"
        If Err.Number <> 0 Then 'If failed to get SAPGUI object, start up SAPlogon.exe, and get SAPGUI again
            Shell "C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe", vbMinimizedNoFocus
            Err.Number = 0
        Else
            Exit Do
        End If
    Loop
    On Error GoTo 0
    If Not IsObject(SAPApplication) Then
        Set SAPApplication = SapGuiAuto.GetScriptingEngine
        SAPApplication.AllowSystemMessages = False
    End If
    If SAPApplication.Connections.Count = 0 Then
        SAPApplication.OpenConnection ("1.1     PRD     [NEW] International Template")  '10.111.4.53
        NeedLogin = True
    End If
    If Not IsObject(SAPConnection) Then Set SAPConnection = SAPApplication.Children(0)  'Create connection
    If Not IsObject(mysap) Then
        If Not (wnd >= 0 And wnd <= 5) Then
            If Not NeedLogin Then
                If MsgBox("Default Window? [wnd = 0]", vbYesNo) = vbYes Then
                    wnd = 0
                Else
                    Do While Not (wnd >= 0 And wnd <= 5)
                        wnd = InputBox("Input the window you want to control. [0-5]")
                    Loop
                End If
            Else
                wnd = 0
            End If
        End If
        If SAPConnection.Children.Length < (wnd + 1) Then
            Set mysap = SAPConnection.Children(0)
            For i = SAPConnection.Children.Length To wnd
                mysap.createsession
                Do
                    Application.Wait [Now() + "0:00:01"]    'Fucking SAP can't even tell if it is busy, GuiSession.busy only considers in-transaction actions.
                Loop Until SAPConnection.Children.Length = (i + 1)
            Next
        End If
        Set mysap = SAPConnection.Children(CInt(wnd))       'THIS IS THE SHIT
    End If
    Set wnd0 = getWnd(mysap)
    wnd0.resizeWorkingPane 133, 34, False
    wnd0.height = 920
    wnd0.Width = 1000
    
    If NeedLogin Or IsTransaction(mysap, "S000") Then
        logged = AutoLogin(mysap)      'No, not your account
    Else
        logged = True
    End If
    If logged Then
        If NewSession Then
            mysap.SendCommand ("/n" & transaction)
        ElseIf transaction <> "" And UCase(transaction) <> mysap.Info.transaction Then
            MsgBox ("Incorrect SAP Transaction. Current transaction is " & mysap.Info.transaction & ".")
            GoTo Reject
        End If
        Set getSession = mysap      '..-. --- .-.. .-.. --- .-- / - .... . / .-. .- -... -... .. - / .... --- .-.. .
    Else
Reject:
        MsgBox "Failed to initialize. Terminating...", vbCritical, "Error"
    End If
    Set wnd0 = Nothing
End Function

' getXXX() - A cleaner way to pointing objects within SAPSessions. MUST reset the object everytime the UI of SAPSession is updated/changed.

Function getWnd(ByRef mysap As Variant, Optional i As Long = 0) As Object
    Set getWnd = mysap.Children(CInt(i))
End Function

Function AutoLogin(ByRef mysap As Variant, Optional ID As String = "", Optional PW As String = "") As Boolean
    If ID = "" Then InputBox ("Please enter your staff ID.")
    If PW = "" Then InputBox ("Please enter the password.")
    mysap.FindById("wnd[0]/usr/txtRSYST-BNAME").text = ID
    mysap.FindById("wnd[0]/usr/pwdRSYST-BCODE").text = PW
    mysap.FindById("wnd[0]").SendVKey 0
    If mysap.Children.Count > 1 Then    'Terminate previous uncleared login and other pop-ups, if exists
        On Error Resume Next
        mysap.FindById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
        mysap.FindById("wnd[1]/tbar[0]/btn[0]").press
        If Err.Number <> 0 Then
            mysap.FindById("wnd[1]").Close
            Err.Number = 0
        End If
        On Error GoTo 0
    ElseIf IsTransaction(mysap, "S000") Then    'Check if sucessfully logged on
        AutoLogin = False
        Exit Function
    End If
    AutoLogin = True
End Function

Function get_folder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Function
        Else
            get_folder = .SelectedItems(1)
            Exit Function
        End If
    End With
End Function

Function IsTransaction(ByVal mysap As Object, ByVal transaction As String) As Boolean
    If UCase(mysap.Info.transaction) <> UCase(transaction) Then
        IsTransaction = False
    Else
        IsTransaction = True
    End If
End Function
Public Sub Shutdown()
    Shell ("shutdown -s")
End Sub
