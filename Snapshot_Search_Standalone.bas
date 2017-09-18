Attribute VB_Name = "Snapshot_Search_Standalone"
'06 Mar 2017: RC1

'Land of Hope and Glory,
'Mother of the Free;
'How shall we extol thee
'Who are born of thee?
'Wider still and wider
'Shall thy bounds be set;
'God, who made thee mighty,
'Make thee mightier yet!
'God, who made thee mighty,
'Make thee mightier yet.

Sub main()
    'Search strings in the ActiveSheet to all XLSX files in a folder.
    Dim a, result As Object
    Dim master As Workbook
    Dim parent, path, file, strFinal As String
    Dim find_str() As String
    Dim srh_ary() As Variant
    
    ReDim find_str(0)
    
    path = Snapshot_Search_Standalone.get_folder() & "\"
    parent = path & "..\"
    
    file = Dir(path & "*.xlsx")
    
    If file = "" Then
        MsgBox "No xlsx files present. Aborting", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Selection(Yes)/Manual Input(No)?", vbYesNo) = vbYes Then
        For Each num In Selection
            find_str(UBound(find_str)) = num.text
            ReDim Preserve find_str(UBound(find_str) + 1)
        Next
        ReDim Preserve find_str(UBound(find_str) - 1)
    Else
        find_str(0) = InputBox("Please enter the search string you want to find.")
        
addmore:
        If MsgBox("Do you want to add more?", vbYesNo) = vbYes Then
            cin = InputBox("Please enter the search string you want to find.")
            If cin <> "" Then
                ReDim Preserve find_str(UBound(find_str) + 1)
                find_str(UBound(find_str)) = cin
            End If
            GoTo addmore
        End If
    End If
    
    If UBound(find_str) = 0 And find_str(0) = "" Then
        MsgBox "No valid keyword entered.", vbExclamation, "Tosser!"
        Exit Sub
    End If
    
    If UBound(find_str) = 0 Then GoTo skiparray
    
    If MsgBox("Putting search results together?", vbYesNo) = vbYes Then
skiparray:
        ReDim srh_ary(0 To 0)
        srh_ary(0) = find_str
    Else
        ReDim srh_ary(LBound(find_str) To UBound(find_str))
        For i = LBound(find_str) To UBound(find_str)
            Dim sub_ary(0) As String
            sub_ary(0) = find_str(i)
            srh_ary(i) = sub_ary
        Next
    End If
    
    For Each item In srh_ary
        file = Dir(path & "*.xlsx")
        
        Set master = Workbooks.Add
        Set result = master.Worksheets(1)
        result.Name = "Result"
        
        VBATurboMode True
        
        master.Worksheets(3).Delete
        master.Worksheets(2).Delete
        
        Do While file <> ""
            Set a = Workbooks.Open(path & file, , True)
            
            ia17_xlsx_extract_nondatarange master, a, item
    
            a.Close
            file = Dir()
            Set a = Nothing
        Loop
        
        result_generate master, result, item
        
        strFinal = ""
        For Each word In item
            strFinal = strFinal & word & ", "
        Next
        strFinal = Left(strFinal, Len(strFinal) - 2)
        
        If Len(strFinal) >= 50 Then strFinal = Left(strFinal, InStrRev(strFinal, ",")) & "& etc"
        
        master.Sheets("Result").Move Before:=Sheets(1)
        VBATurboMode False
        master.SaveAs (parent & "Search Result of " & filename_normalize(strFinal) & " (" & format(Date, "yyyy-mm-dd") & ").xlsx")
        master.Close
        Set result = Nothing
        Set master = Nothing
    Next
    
    Shell "C:\Windows\explorer.exe """ & parent & "", vbNormalFocus
End Sub

Function finalize(ByVal sheet As Worksheet)
    If IsEmptySheet(sheet) Then
        Application.DisplayAlerts = False
        sheet.Delete
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    With sheet
        .Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range("A1:H1") = Split("Search Hits,Line,Plan,Plan Name,Op,Op Short Text,Workctr.,Package Selected", ",")
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Size = 14
        sheet.Activate
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        ActiveWindow.FreezePanes = True
        
        .Columns.AutoFit
    End With
End Function

Function result_generate(master As Workbook, result As Worksheet, ByRef item As Variant)
    With result
        .Cells(1, 1).Value = "SEARCH RESULT of"
        For Each word In item
            strF = strF & word & vbCrLf
        Next
        .Cells(1, 2).Value = Left(strF, Len(strF) - 1)
        
        .Range("A2:D2") = Split("Plans,Hit Counts,Op. Counts,Plan Counts", ",")
        i = 3
        For Each sheet In master.Sheets
            If sheet.Name <> "Result" Then
                .Cells(i, 1) = sheet.Name
                If IsEmpty(sheet.Cells(1, 2)) Then
                    .Cells(i, 2) = 0
                    .Cells(i, 3) = 0
                    .Cells(i, 4) = 0
                    With .Rows(i).Font
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.499984740745262
                    End With
                Else
                    .Cells(i, 2) = sheet.UsedRange.Rows.Count
                    plan_count = 0
                    op_count = 0
                    count_inv sheet, plan_count, op_count, 0, 1
                    .Cells(i, 3) = op_count
                    .Cells(i, 4) = plan_count
                    .Hyperlinks.Add .Cells(i, 1), "", "'" & sheet.Name & "'!A1", , sheet.Name
                End If
                i = i + 1
                finalize sheet
            End If
        Next
        .Cells(i, 1) = "Total"
        .Cells(i, 2).Formula = "=SUM(B3:B" & i - 1 & ")"
        .Cells(i, 3).Formula = "=SUM(C3:C" & i - 1 & ")"
        .Cells(i, 4).Formula = "=SUM(D3:D" & i - 1 & ")"
        .Rows(2).Font.Bold = True
        .Rows(i).Font.Bold = True
        .Rows(1).Font.Size = 20
        .Activate
        .Range(Cells(1, 3), Cells(1, 4)).Font.ThemeColor = xlThemeColorDark1
        For Each sheet In master.Sheets
            format_lul sheet
            If sheet.Name <> "Result" Then
                SortTable sheet
            End If
        Next
        .Columns.AutoFit
    End With
End Function
Sub SortTable(ByRef sheet As Variant)
    With sheet.ListObjects(1).Sort
        .SortFields.Clear
        .SortFields.Add key:=Range("Table" & Replace(sheet.Name, "-", "_") & "[[#Headers],[Line]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Function ia17_xlsx_extract_nondatarange(master As Workbook, ByVal data As Workbook, find_str As Variant)
    Dim j As Long
    Dim o, n, r As Object
    Set o = data.Sheets(1)
    Set n = master.Sheets.Add
    n.Name = o.Name
    j = 1
    
    Dim textrow() As String
    Dim numrow() As Long
    ReDim textrow(1 To 1)
    ReDim numrow(1 To 1)
    
    For Each item In find_str
        With o.UsedRange
            Set r = .Find(item, LookIn:=xlValues, LookAt:=xlPart, searchorder:=xlByRows, MatchCase:=False)
            If Not r Is Nothing Then
                firstaddress = r.Address
                Do
                    textrow(j) = r.text
                    numrow(j) = r.row
                    Set r = .FindNext(r)
                    j = j + 1
                    ReDim Preserve textrow(1 To j)
                    ReDim Preserve numrow(1 To j)
                Loop While Not r Is Nothing And r.Address <> firstaddress
                n.Cells(1, 1).Resize(j - 1, 1) = Application.Transpose(textrow)
                n.Cells(1, 2).Resize(j - 1, 1) = Application.Transpose(numrow)
            End If
        End With
    Next
    
    Erase textrow, numrow
    
    n.Columns(5).NumberFormat = "@"
    
    MaxR = o.UsedRange.Rows.Count
    MaxL = j - 1
                                            
    w_n = 3
    x_n = 3
    y_v = 1
    y_n = 7
    z_n = 6
    v_v = 1     'v = Workcentre-related
    v_n = 6
    u_n = 5     'u = Maint-related
    u_v = 1
    
    Do While IsEmpty(o.Cells(3, x_n))
        x_n = x_n + 1
    Loop
    Do While Len(o.Cells(3, z_n)) < 3
        z_n = z_n + 1
    Loop
    Do While o.Cells(y_v, 2) <> "Operation"
        y_v = y_v + 1
    Loop
    Do While Len(o.Cells(y_v, y_n)) < 3
        y_n = y_n + 1
    Loop
    Do While IsEmpty(o.Cells(y_v, w_n))
        w_n = w_n + 1
    Loop
    Do While o.Cells(v_v, x_n) <> "Work center"
        v_v = v_v + 1
    Loop
    Do While IsEmpty(o.Cells(v_v, v_n))
        v_n = v_n + 1
    Loop
    
    Do While u_v <= MaxR
        If o.Cells(u_v, 3) = "MntPack." Then
            HavePack = True
            Exit Do
        End If
        u_v = u_v + 1
    Loop
    
    If HavePack Then
        Do While IsEmpty(o.Cells(u_v, u_n))
            u_n = u_n + 1
        Loop
    End If
    
    
    If MaxL <> 0 Then
        Editrange = Range(n.Cells(1, 1), n.Cells(MaxL, 8))
        For x = 1 To MaxL
            
            Z = Editrange(x, 2)
            Do While IsEmpty(o.Cells(Z, 1))
                Z = Z - 1
            Loop
            
            y = Editrange(x, 2)
            Do While o.Cells(y, 2) <> "Operation" And y >= Z
                y = y - 1
            Loop
            
            Editrange(x, 3) = o.Cells(Z, x_n)
            Editrange(x, 4) = o.Cells(Z, z_n)
            Editrange(x, 5) = format(o.Cells(y, w_n), "0000")
            
            If x = 1 Then
                SameOp = False
            Else
                SameOp = (Editrange(x, 3) = Editrange(x - 1, 3)) And (Editrange(x, 4) = Editrange(x - 1, 4)) And (Editrange(x, 5) = Editrange(x - 1, 5))
            End If
            
            If SameOp Then
                Editrange(x, 6) = Editrange(x - 1, 6)
                Editrange(x, 7) = Editrange(x - 1, 7)
                Editrange(x, 8) = maint
            Else
            
                Editrange(x, 6) = o.Cells(y, y_n)
                
                V = y + 1
                Do While IsEmpty(o.Cells(V, v_n))
                    V = V + 1
                Loop
                Editrange(x, 7) = o.Cells(V, v_n)
                
                maint = ""
                u = 5
                Do While o.Cells(V + u, 3) = "MntPack."
                    If maint <> "" Then maint = maint & vbCrLf
                    maint = maint & o.Cells(V + u, u_n) & " " & o.Cells(V + u, v_n)
                    u = u + 1
                Loop
                If maint <> "" Then Editrange(x, 8) = maint
            End If
        Next
        Range(n.Cells(1, 1), n.Cells(MaxL, 8)) = Editrange
    End If
    
    n.UsedRange.Columns.AutoFit
    n.UsedRange.Rows.AutoFit
    
    Set o = Nothing
    Set n = Nothing
End Function

Sub shutdown_after_search()   'Does what the title say
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

Sub VBATurboMode(Enab As Boolean)
    If Enab Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayStatusBar = False
        Application.DisplayAlerts = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayStatusBar = True
        Application.DisplayAlerts = True
    End If
End Sub

Function filename_normalize(ByRef filename As String) As String
    Dim crit() As String
    crit() = Split("\ / * ? "" < > | : [ ]")
    For Each char In crit
        filename = Replace(filename, char, "_")
    Next
    filename_normalize = filename
End Function

Sub count_inv(sheet As Variant, ByRef plan_count As Variant, ByRef op_count As Variant, ByRef hit_count As Variant, Optional StartRow As Long = 2)
    plan_name = ""
    op_name = ""
    With sheet
        Do While Not IsEmpty(.Cells(StartRow, 2))
            hit_count = hit_count + 1
            If plan_name <> .Cells(StartRow, 3) Then
                plan_count = plan_count + 1
                plan_name = .Cells(StartRow, 3)
                op_count = op_count + 1
                op_name = .Cells(StartRow, 5)
            ElseIf op_name <> .Cells(StartRow, 5) Then
                op_count = op_count + 1
                op_name = .Cells(StartRow, 5)
            End If
            StartRow = StartRow + 1
        Loop
    End With
End Sub

Sub format_lul(Optional ByRef shit As Variant = Nothing)
    If shit Is Nothing Then
        Set shit = ActiveSheet
    End If
    
    shit.ListObjects.Add(xlSrcRange, shit.UsedRange, , xlYes).Name = "Table" & shit.Name
    shit.ListObjects("Table" & shit.Name).TableStyle = "TableStyleLight1"
End Sub
Function IsEmptySheet(ByVal shit As Worksheet) As Boolean
    If WorksheetFunction.CountA(shit.Cells) = 0 Then
        IsEmptySheet = True
    Else
        IsEmptySheet = False
    End If
End Function
Public Sub Shutdown()
    Shell ("shutdown -s")
End Sub
