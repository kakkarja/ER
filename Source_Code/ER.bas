Attribute VB_Name = "Module1"
Option Explicit

'''START USERFORM EXPENSE REPORT'''

Sub LK(control As IRibbonControl)
    On Error GoTo Oboy
    ActiveSheet.Activate
    LapKeu.Show
Oboy:
End Sub


'''START USERFORM SET PASSWORD'''

Sub Pswd(control As IRibbonControl)
    On Error GoTo Oboy
    ActiveSheet.Activate
    MsgBox "WARNING!!!" & Chr(10) & _
    "You are about to set Password" _
    & " for your Active Workbook" _
    & ". Please do not forget it." _
    , vbInformation, "Password Setup"
    
    CPass.Show
Oboy:
End Sub


'''USERFORM EXPENSE REPORT'''

Private Sub AddWS_Click()
Dim WB As Workbook, WB1
Dim Ws As Worksheet, WS1 As Worksheet
Dim WSN As Long
Dim C As Long
    Set WB = ThisWorkbook
    Set Ws = WB.Worksheets(2)
    Set WB1 = ActiveWorkbook
        WSN = Worksheets.Count
        CWSN
        On Error Resume Next
        If WSN > 1 Then
            For C = WSN To 1 Step -1
                Set WS1 = WB1.Worksheets(C)
                If Left(WS1.Name, 6) = "ExpRep" Then
                    'Worksheets(C + 1).Activate
                    On Error Resume Next
                    Ws.Copy Worksheets(C + 1)
                    If Err.Number > 0 Then
                        Worksheets.Add , WS1
                        Ws.Copy Worksheets(C + 1)
                        Nama
                        Fn
                        On Error GoTo 0
                        GoTo bye
                    End If
                    Nama
                    Fn
                    GoTo bye
                End If
            Next C
            Ws.Copy WS1
            Nama
            Fn
        Else
            Set WS1 = WB1.Worksheets(WSN)
            If Left(WS1.Name, 6) = "ExpRep" Then
                Worksheets.Add , WS1
                Ws.Copy Worksheets(WSN + 1)
                Nama
                Fn
            Else
                WS1.Activate
                Ws.Copy WS1
                Nama
                Fn
            End If
        End If
        
bye:
Set WB = Nothing
Set Ws = Nothing
Set WS1 = Nothing
WSN = 0
C = 0
End Sub

Private Sub CWSN()
Dim WSNa As Worksheet
    On Error Resume Next
    For Each WSNa In ActiveWorkbook.Worksheets
        If Left(WSNa.Name, 6) = "ExpRep" Then
            If WSNa.Cells(1, 1) = "" Then
                With Application
                    .DisplayAlerts = False
                    WSNa.Delete
                    .DisplayAlerts = True
                End With
            End If
        End If
        If Left(WSNa.Name, 6) = "LapKeu" Then
            With Application
                .DisplayAlerts = False
                WSNa.Delete
                .DisplayAlerts = True
            End With
        End If
    Next WSNa
End Sub

Private Sub Fn()
        ActiveSheet.PageSetup.LeftFooter = "&P/&N"
End Sub

Private Sub Nama()
Dim D As String
    D = Format(Now, "DDMMYYYY HHmmSS")
    With ActiveSheet
        .Name = "ExpRep" & " " & D
        .Tab.ColorIndex = 42
    End With
D = vbNullString
End Sub

Private Sub Ch_Click()
Dim Ch As Chart
Dim DS As String, Ds1 As String
Dim X As Variant
On Error Resume Next
Set Ch = ActiveSheet.Shapes(1).Chart
    If Err.Number = 0 Then
        ActiveSheet.Shapes(1).Delete
        GoTo bye
    End If
    On Error GoTo 0
Set Ch = Nothing
Set Ch = ActiveSheet.Shapes.AddChart2(209, xlColumnClustered, 10).Chart
    DS = ActiveSheet.Name
    X = Cells(Rows.Count, 1).End(xlUp).Row
    With Ch
        .FullSeriesCollection(1).Delete
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(1)
        Ds1 = Replace(Cells(X, 1).End(xlUp).Address, "A", "C")
            .Name = "='" & DS & "'!$C$1"
            .Values = _
            "='" & DS & "'!$C$2:" & Ds1
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(2)
        Ds1 = vbNullString
        Ds1 = Replace(Cells(X, 1).End(xlUp).Address, "A", "D")
            .Name = "='" & DS & "'!$D$1"
            .Values = _
            "='" & DS & "'!$D$2:" & Ds1
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(3)
        Ds1 = vbNullString
        Ds1 = Replace(Cells(X, 1).End(xlUp).Address, "A", "E")
            .Name = "='" & DS & "'!$E$1"
            .Values = _
            "='" & DS & "'!$E$2:" & Ds1
        End With
        .SeriesCollection.NewSeries
        With .FullSeriesCollection(4)
        Ds1 = vbNullString
        Ds1 = Replace(Cells(X, 1).End(xlUp).Address, "A", "G")
            .Name = "='" & DS & "'!$G$1"
            .Values = _
            "='" & DS & "'!$G$2:" & Ds1
        End With
        Ds1 = vbNullString
        Ds1 = Replace(Cells(X, 1).End(xlUp).Address, "A", "B")
        .FullSeriesCollection(4).XValues = _
        "='" & DS & "'!$B$2:" & Ds1
        .ChartTitle.Text = ActiveSheet.Name
        With .ChartArea
            .RoundedCorners = True
            With ActiveWindow
                If .WindowState <> -4137 Then
                    .WindowState = xlMaximized
                End If
            End With
            .Height = ActiveWindow.Height - 350
            .Width = ActiveWindow.Width - 70
        End With
    End With
bye:
Set Ch = Nothing
DS = vbNullString
Ds1 = vbNullString
X = 0
End Sub

Private Sub Del_Click()
Dim Ask As Variant
    With ActiveCell
        If .Value <> "" Or .Offset(, 1) <> "" Then
            Ask = MsgBox _
            ("Delete Record?", vbYesNo, _
            "Laporan Keuangan")
            Select Case Ask
                Case Is = vbYes
                    .Resize(, 4).Clear
            End Select
        End If
    End With
End Sub

Private Sub DelWS_Click()
Dim DWS As String
Dim Ws As Worksheet
Dim WB As Workbook
    Set WB = ActiveWorkbook
    On Error GoTo bye
    DWS = MsgBox("Do you want to delete this worksheet?" _
    , vbYesNo, "Expense Report")
    Select Case DWS
        Case Is = vbYes
            Set Ws = ActiveSheet
            With WB
                .Application.DisplayAlerts = False
                Ws.Delete
                .Application.DisplayAlerts = True
                AcCel
            End With
    End Select
    Set Ws = Nothing
    Set Ws = ActiveSheet
    If Left(Ws.Name, 6) <> "ExpRep" Then
        On Error Resume Next
        Worksheets(Ws.Index - 1).Activate
        If Err.Number > 0 Then
            ActiveWorkbook.Unprotect Environ( _
            "userprofile")
            ActiveSheet.Unprotect Environ( _
            "userprofile")
            End
        End If
    End If
bye:
    CWSN
Set Ws = Nothing
Set WB = Nothing
DWS = vbNullString
End Sub
Private Sub AcCel()
    ActiveCell.Activate
End Sub
Private Sub Kal_Click()
    DCall
End Sub

Private Sub LKCall_Click()
Dim LC As Variant
    LC = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(LC, 1).End(xlUp).Offset(1).Select
    DCall
LC = 0
End Sub

Private Sub LKExp_Change()
    With ActiveCell
        If .Column = 1 Then
            .Offset(, 1).NumberFormat = "@"
            'On Error Resume Next
            If .Offset(, 1).Characters.Count <= 28 Then
                If .Value = "" Then
                    With .Offset(, 1)
                        .Value = LKExp.Text
                        If .Font.Bold = False Then
                            .Font.Bold = True
                        End If
                    End With
                Else
                    With .Offset(, 1)
                        If .Font.Bold = True Then
                            .Font.Bold = False
                        End If
                        .Value = LKExp.Text
                    End With
                End If
            Else
                With .Offset(, 1)
                    .Characters(29).Delete
                    LKExp = .Value
                End With
            End If
        End If
    End With
End Sub

Private Sub LKExp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LKExp.SetFocus
End Sub

Private Sub Pem_Change()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(, 2).Value <= 999999999999# Then
                If .Offset(, 1) <> "" And _
                .Value <> "" Then
                    Cek
                    If IsNumeric(Pem) Or Pem = "" Then
                        With .Offset(, 2)
                            If .Offset(, 1) = "" Then
                                .Value = Pem.Text
                                .NumberFormat = "#,##0"
                            End If
                        End With
                    End If
                End If
            Else
                .Offset(, 2).Value = Left(.Offset(, 2).Value, 12)
                Pem = Left(Pem, 12)
            End If
        End If
    End With
End Sub

Private Sub Cek()
    If ActiveCell.Column = 1 Then
        With ActiveCell.Offset(, 1)
            If .Font.Bold = True Then
                .Font.Bold = False
            End If
        End With
    End If
End Sub

Private Sub Pem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Pem.SetFocus
End Sub

Private Sub Peng_Change()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(, 3).Value <= 999999999999# Then
                If .Offset(, 1) <> "" And _
                .Value <> "" Then
                    Cek
                    If IsNumeric(Peng) Or Peng = "" Then
                        With .Offset(, 3)
                            If .Offset(, -1) = "" Then
                                .Value = Peng.Text
                                .NumberFormat = "#,##0"
                            End If
                        End With
                    End If
                End If
            Else
                .Offset(, 3).Value = Left(.Offset(, 2).Value, 12)
                Peng = Left(Peng, 12)
            End If
        End If
    End With
End Sub

Private Sub Peng_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Peng.SetFocus
End Sub

Private Sub SpinButton1_SpinDown()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(1).Row = "344" Then
                Exit Sub
            Else
                If .Offset(1, 1) = "TOTAL" Then
                    .Offset(2).Select
                    Ap
                Else
                    .Offset(1).Select
                    Ap
                End If
            End If
        End If
    End With
End Sub

Private Sub Ap()
    With ActiveCell
        With .Offset(, 1)
            If .Value <> "" Then
                LKExp = .Value
            Else
                LKExp = ""
            End If
        End With
        With .Offset(, 2)
            If .Value <> "" Then
                Pem = .Value
            Else
                Pem = ""
            End If
        End With
        With .Offset(, 3)
            If .Value <> "" Then
                Peng = .Value
            Else
                Peng = ""
            End If
        End With
    End With
End Sub

Private Sub DCall()
    KalenderK.Show
End Sub

Private Sub SpinButton1_SpinUp()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(-1).Address = "$A$1" Then
                Exit Sub
            Else
                If .Offset(-1, 1) = "TOTAL" Then
                    .Offset(-2).Select
                    Ap
                Else
                    .Offset(-1).Select
                    Ap
                End If
            End If
        End If
    End With
End Sub

Private Sub SpinButton2_SpinDown()
    If ActiveSheet.Index > 1 Then
        Worksheets(ActiveSheet.Index - 1).Activate
    End If
End Sub

Private Sub SpinButton2_SpinUp()
Dim NumWS As Long
    NumWS = Worksheets.Count
    On Error GoTo bye
    If ActiveSheet.Index < NumWS And _
    Left(Worksheets(ActiveSheet.Index + 1).Name, 6) = _
    "ExpRep" Then
        Worksheets(ActiveSheet.Index + 1).Activate
    End If
bye:
NumWS = 0
End Sub

Private Sub UserForm_Click()
    Calc.Show
End Sub

Private Sub UserForm_Initialize()
Dim Ws As Worksheet
Dim WB As Workbook
Dim WSN As Long
Dim C As Long
Fx1
Set WB = ActiveWorkbook
    If WB.ProtectStructure = True Then
        WB.Unprotect Environ("userprofile")
    End If
    WSN = Worksheets.Count
    For C = 1 To WSN
        Set Ws = WB.Worksheets(C)
        If Left(Ws.Name, 6) = "ExpRep" Then
            With Ws
                If .ProtectContents = True Then
                    .Unprotect Environ("userprofile")
                End If
            End With
        End If
    Next C
    Set Ws = Nothing
    Set Ws = ActiveSheet
    With Ws
        If Left(.Name, 6) <> "ExpRep" Then
            AddWS_Click
        End If
    End With
Set Ws = Nothing
Set WB = Nothing
WSN = 0
Fx2
End Sub

Private Sub UserForm_Terminate()
Dim Ws As Worksheet
Dim WSN As Long
Dim C As Long
Dim WB As Workbook
Set WB = ActiveWorkbook
    WSN = Worksheets.Count
        
        For C = 1 To WSN
            Set Ws = WB.Worksheets(C)
            If Left(Ws.Name, 6) = "ExpRep" Then
                With Ws
                    If ActiveCell.Locked = False Then
                        .Cells.Locked = True
                    End If
                    If .ProtectContents = False Then
                        .Protect Environ("userprofile"), _
                        AllowFiltering:=True
                        .EnableSelection = xlNoSelection
                    End If
                End With
            End If
        Next C
        With WB
            .Protect Environ("userprofile")
        End With
Set WB = Nothing
Set Ws = Nothing
WSN = 0
C = 0
End Sub
Private Sub Fast(SU As Boolean, DS As Boolean, C As String, EE As Boolean)
    Application.ScreenUpdating = SU
    Application.DisplayStatusBar = DS
    Application.Calculation = C
    Application.EnableEvents = EE
End Sub
Private Sub Fx1()
Call Fast(False, True, xlCalculationManual, False)
End Sub
Private Sub Fx2()
Call Fast(True, True, xlCalculationAutomatic, True)
End Sub


'''USERFORM SET PASSWORD'''

Private Sub SetP_Click()
    If Pssd = "" Then
        MsgBox "Please submit your password", , _
        "Password Setup"
        Exit Sub
    End If
    With ActiveWorkbook
        If .Path = "" Then
            MsgBox "Please save the workbook first," & _
            " in order to setup a password.", , _
            "Password Setup"
            Unload Me
        Else
            Application.DisplayAlerts = False
            .SaveAs .Path & "\" & .Name, , Pssd
            MsgBox "Please do not forget your password." _
            & " You have just secured your workbook" _
            & " viewing.", vbInformation, _
            "Password Setup"
            Application.DisplayAlerts = True
            Unload Me
        End If
    End With
End Sub


'''USERFORM CALCULATOR'''

Dim i As String
Dim T As String
Dim o As Double

Private Sub AllClear_Click()
    Screen.Text = ""
    i = 0
    T = vbNullString
    o = 0
End Sub

Private Sub Coma_Click()
    If Screen <> "" Then
        i = Screen.Text
        Screen.Text = i & ","
    Else
        Screen.Text = ""
    End If
i = 0
End Sub

Private Sub Divide_Click()
    If Screen <> "" Then
        If o = 0 Then
            o = Screen.Text
        End If
        T = "/"
        Screen.Text = ""
    End If

End Sub

Private Sub Eight_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 8
        Else
            i = Screen.Text
            Screen.Text = i & 8
        End If
    Else
        Screen.Text = 8
    End If

End Sub

Private Sub Equal_Click()
Dim Eq As Double
    On Error Resume Next
    Select Case T
        Case Is = "+"
            Eq = o + CDbl(Screen)
            If Err.Number <> 0 Then
                On Error GoTo 0
                Screen = o
                GoTo bye
            End If
            o = 0
        Case Is = "-"
            If o <> 0 Then
                If o - CLng(Screen) < 0 Then
                    Dim eq2 As LongPtr
                    eq2 = o - CLng(Screen)
                    o = 0
                    Screen.Text = eq2
                    GoTo bye
                Else
                    Eq = o - CDbl(Screen)
                    If Err.Number <> 0 Then
                        On Error GoTo 0
                        Screen = o
                        GoTo bye
                    End If
                    o = 0
                    If Eq < 0 Then
                        Eq = Replace(Eq, "-", "")
                    End If
                End If
            Else
                Dim eq3 As LongPtr
                    eq3 = o - CLng(Screen)
                    o = 0
                    Screen.Text = eq3
                If eq3 < 0 Then
                    eq3 = Replace(eq3, "-", "")
                    Screen.Text = eq3
                End If
                GoTo bye
            End If
        Case Is = "*"
            If o <> 0 Then
                Eq = o * CDbl(Screen)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Screen = o
                    GoTo bye
                End If
                o = 0
            Else
                GoTo bye
            End If
        Case Is = "/"
            If o <> 0 Then
                Eq = o / CDbl(Screen)
                If Err.Number <> 0 Then
                    On Error GoTo 0
                    Screen = o
                    GoTo bye
                End If
                o = 0
            Else
                GoTo bye
            End If
    End Select
    
    Screen.Text = Eq
bye:
eq3 = 0
eq2 = 0
Eq = 0
End Sub

Private Sub ExP_Click()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(, 3).Value <= 999999999999# Then
                If .Offset(, 1) <> "" And _
                .Value <> "" Then
                    Cek
                    With .Offset(, 3)
                        If InStr(Screen, ",") > 0 Then
                            .NumberFormat = "#,#.#0"
                            .Value = Replace _
                            (Replace(Screen.Text, ".", "") _
                            , ",", ".")
                        Else
                            .NumberFormat = "#,##0"
                            .Value = Replace _
                            (Screen.Text, ".", ",")
                        End If
                    End With
                End If
            Else
                .Offset(, 3).Value = Left _
                (.Offset(, 3).Value, 12)
            End If
        End If
        If .Offset(, 2) <> "" Then
            .Offset(, 2) = ""
        End If
    End With
End Sub

Private Sub ExP_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Exp.BackColor <> &HFFFF00 Then
        Exp.BackColor = &HFFFF00
        InC.BackColor = &H8000000F
    End If
End Sub

Private Sub Five_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 5
        Else
            i = Screen.Text
            Screen.Text = i & 5
        End If
    Else
        Screen.Text = 5
    End If

End Sub

Private Sub Four_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 4
        Else
            i = Screen.Text
            Screen.Text = i & 4
        End If
    Else
        Screen.Text = 4
    End If

End Sub

Private Sub InC_Click()
    With ActiveCell
        If .Column = 1 Then
            If .Offset(, 2).Value <= 999999999999# Then
                If .Offset(, 1) <> "" And _
                .Value <> "" Then
                    Cek
                    With .Offset(, 2)
                        If InStr(Screen, ",") > 0 Then
                            .NumberFormat = "#,#.#0"
                            .Value = Replace _
                            (Replace(Screen.Text, ".", "") _
                            , ",", ".")
                        Else
                            .NumberFormat = "#,##0"
                            .Value = Replace _
                            (Screen.Text, ".", ",")
                        End If
                    End With
                End If
            Else
                .Offset(, 2).Value = Left _
                (.Offset(, 2).Value, 12)
            End If
        End If
        If .Offset(, 3) <> "" Then
            .Offset(, 3) = ""
        End If
    End With
End Sub

Private Sub InC_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If InC.BackColor <> &HFFFF00 Then
        InC.BackColor = &HFFFF00
        Exp.BackColor = &H8000000F
    End If
End Sub

Private Sub Minus_Click()
    If Screen <> "" Then
        If o = 0 Then
            o = Screen.Text
        End If
        T = "-"
        Screen.Text = ""
    End If

End Sub

Private Sub Nine_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 9
        Else
            i = Screen.Text
            Screen.Text = i & 9
        End If
    Else
        Screen.Text = 9
    End If

End Sub

Private Sub One_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 1
        Else
            i = Screen.Text
            Screen.Text = i & 1
        End If
    Else
        Screen.Text = 1
    End If
i = 0
End Sub

Private Sub Percentage_Click()
    If Screen > 0 And Screen <> "" Then
        i = Screen.Text * (1 / 100)
        Screen = i
    End If
End Sub

Private Sub Plus_Click()
    If Screen <> "" Then
        If o = 0 Then
            o = Screen.Text
        End If
        T = "+"
        Screen.Text = ""
    End If
End Sub

Private Sub PlusMinus_Click()
    If Screen > 0 And Screen <> "" Then
        i = "-" & Screen.Text
        Screen = i
    Else
        i = Replace(Screen, "-", "")
        Screen = i
    End If
End Sub

Private Sub Screen_Change()
    If Screen <> 0 Then
        If Len(Screen) <= 17 Then
            If InStr(Screen, ",") = 0 Then
                Screen = Format(Screen.Text, "#,##")
            End If
        Else
            Screen = Left(Screen.Text, 17)
        End If
    End If
bye:
i = vbNullString
End Sub

Private Sub Screen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If InC.BackColor = &HFFFF00 Then
        InC.BackColor = &H8000000F
    ElseIf Exp.BackColor = &HFFFF00 Then
        Exp.BackColor = &H8000000F
    End If

End Sub

Private Sub Seven_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 7
        Else
            i = Screen.Text
            Screen.Text = i & 7
        End If
    Else
        Screen.Text = 7
    End If

End Sub

Private Sub Six_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 6
        Else
            i = Screen.Text
            Screen.Text = i & 6
        End If
    Else
        Screen.Text = 6
    End If

End Sub

Private Sub Three_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 3
        Else
            i = Screen.Text
            Screen.Text = i & 3
        End If
    Else
        Screen.Text = 3
    End If

End Sub

Private Sub Times_Click()
    If Screen <> "" Then
        If o = 0 Then
            o = Screen.Text
        End If
        T = "*"
        Screen.Text = ""
    End If

End Sub

Private Sub Two_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 2
        Else
            i = Screen.Text
            Screen.Text = i & 2
        End If
    Else
        Screen.Text = 2
    End If
i = 0
End Sub

Private Sub UserForm_Initialize()
    Screen.Locked = True
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If InC.BackColor = &HFFFF00 Then
        InC.BackColor = &H8000000F
    ElseIf Exp.BackColor = &HFFFF00 Then
        Exp.BackColor = &H8000000F
    End If
End Sub

Private Sub Zero_Click()
    If Screen <> "" Then
        If T <> "" And Screen = "" Then
            Screen.Text = 0
        Else
            i = Screen.Text
            Screen.Text = i & 0
        End If
    Else
        Screen.Text = 0
    End If
End Sub
Private Sub Cek()
    If ActiveCell.Column = 1 Then
        With ActiveCell.Offset(, 1)
            If .Font.Bold = True Then
                .Font.Bold = False
            End If
        End With
    End If
End Sub


'''USERFORM CALENDAR (2017-2019)'''

Dim i As Long
Dim j As Long
Dim k As Long
Dim D As Long
Dim Bul() As Variant
Dim Mon() As Variant
Dim Hit() As Variant
Dim BVal() As Variant
Dim H() As Variant
Dim M As String
Dim Lab As Variant
Dim Rg As Range
Dim Tx As String
Dim CD As Date
Dim CDe As Long
Dim DT As String
Dim Colm As Long
Dim Lr As Variant
Dim Add As Variant

Private Sub Label1_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label1.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label1.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label10_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label10.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label10.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label11_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label11.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label11.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label12_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label12.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label12.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label13_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label13.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label13.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label14_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label14.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label14.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label15_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label15.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label15.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label16_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label16.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label16.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label17_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label17.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label17.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label18_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label18.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label18.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .Select
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label19_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label19.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label19.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label2_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label2.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label2.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label20_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label20.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label20.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label21_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label21.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label21.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label22_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label22.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label22.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label23_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label23.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label23.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label24_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label24.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label24.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label25_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label25.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label25.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label26_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label26.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label26.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label27_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label27.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label27.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label28_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label28.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label28.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label29_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label29.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label29.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label3_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label3.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label3.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label30_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label30.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label30.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label31_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label31.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label31.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label32_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label32.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label32.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label33_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label33.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label33.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label34_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label34.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label34.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label35_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label35.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label35.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label36_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label36.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label36.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label37_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label37.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label37.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label38_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label38.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label38.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label39_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label39.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label39.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label4_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label4.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label4.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label40_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label40.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label40.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label41_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label41.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label41.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label42_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label42.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label42.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label5_Click()
            Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label5.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label5.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label6_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label6.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label6.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label7_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label7.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label7.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label8_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label8.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label8.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub Label9_Click()
        Bul = _
    Array(0, "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12")
    If Not Label9.Caption = "" Then
        CD = DateSerial(CLng(Years.Text), CLng(Bul(SpinButton1.Value)) _
        , CLng(Label9.Caption))
        CDe = CLng(CD)
        With ActiveCell
            .NumberFormat = "@"
            .Value = Format(CDe, "DD/MM/YYYY")
            Columns(.Column).AutoFit
        End With
    End If
    Unload Me
End Sub

Private Sub SpinButton1_Change()
    Select Case Years
        Case Is = "2017"
            Y2017
        Case Is = "2018"
            Y2018
        Case Is = "2019"
            Y2019
        Case Else
            Years = Year(Date)
    End Select
End Sub

Private Sub DD()
    For j = 1 To 42
        Controls("label" & j).Caption = ""
    Next j
    j = 0
End Sub

Private Sub UserForm_Initialize()
    With Years
        .AddItem "2017"
        .AddItem "2018"
        .AddItem "2019"
    End With
    With Controls
        With .Item("Minggu")
            .Caption = Format(Weekday(1, vbSunday), "ddd")
        End With
        With .Item("Senin")
            .Caption = Format(Weekday(2, vbSunday), "ddd")
        End With
        With .Item("Selasa")
            .Caption = Format(Weekday(3, vbSunday), "ddd")
        End With
        With .Item("Rabu")
            .Caption = Format(Weekday(4, vbSunday), "ddd")
        End With
        With .Item("Kamis")
            .Caption = Format(Weekday(5, vbSunday), "ddd")
        End With
        With .Item("Jumat")
            .Caption = Format(Weekday(6, vbSunday), "ddd")
        End With
        With .Item("Sabtu")
            .Caption = Format(Weekday(7, vbSunday), "ddd")
        End With
    End With
    M = Format(Date, "m")
    For i = 1 To 12
        If M = i Then
            SpinButton1.Value = i
            SpinButton1_Change
            Exit Sub
        End If
    Next i
End Sub

Private Sub UserForm_Terminate()
    Cells.Locked = False
    On Error GoTo Good
    Dim X As Range
    For Each X In Cells.SpecialCells(xlCellTypeConstants)
        X.Locked = True
    Next X
    'ActiveSheet.Protect "1Core13"
Good:
'x = 0
End Sub

Private Sub Fast(SU As Boolean, DS As Boolean, C As String, EE As Boolean)
    Application.ScreenUpdating = SU
    Application.DisplayStatusBar = DS
    Application.Calculation = C
    Application.EnableEvents = EE
End Sub
Private Sub Fx1()
Call Fast(False, False, xlCalculationManual, False)
End Sub
Private Sub Fx2()
Call Fast(True, True, xlCalculationAutomatic, True)
End Sub


Private Sub Y2017()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 42736, 42767, 42795, 42826, 42856, 42887, 42917, _
    42948, 42979, 43009, 43040, 43070)
    
    Bul = _
    Array(0, Format(42736, "mmmm"), _
    Format(42767, "mmmm"), Format(42795, "mmmm"), _
    Format(42826, "mmmm"), Format(42856, "mmmm"), _
    Format(42887, "mmmm"), Format(42917, "mmmm"), _
    Format(42948, "mmmm"), Format(42979, "mmmm"), _
    Format(43009, "mmmm"), Format(43040, "mmmm"), _
    Format(43070, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

Private Sub Y2018()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 43101, 43132, 43160, 43191, 43221, 43252, 43282, _
    43313, 43344, 43374, 43405, 43435)
    
    Bul = _
    Array(0, Format(43101, "mmmm"), _
    Format(43132, "mmmm"), Format(43160, "mmmm"), _
    Format(43191, "mmmm"), Format(43221, "mmmm"), _
    Format(43252, "mmmm"), Format(43282, "mmmm"), _
    Format(43313, "mmmm"), Format(43344, "mmmm"), _
    Format(43374, "mmmm"), Format(43405, "mmmm"), _
    Format(43435, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub
Private Sub Y2019()
Dim Minggu As Integer, Senin As Integer, Selasa As Integer _
, Rabu As Integer, Kamis As Integer, Jumat As Integer _
, Sabtu As Integer
    BVal = Array _
    (0, 43466, 43497, 43525, 43556, 43586, 43617, 43647, _
    43678, 43709, 43739, 43770, 43800)
    
    Bul = _
    Array(0, Format(43466, "mmmm"), _
    Format(43497, "mmmm"), Format(43525, "mmmm"), _
    Format(43556, "mmmm"), Format(43586, "mmmm"), _
    Format(43617, "mmmm"), Format(43647, "mmmm"), _
    Format(43678, "mmmm"), Format(43709, "mmmm"), _
    Format(43739, "mmmm"), Format(43770, "mmmm"), _
    Format(43800, "mmmm"))
    
    Hit = Array(0, 31, 28, 31, 30, 31, 30, 31, 31, 30, _
    31, 30, 31)
    
    H = Array(0, Format(Weekday(1, vbSunday), "dddd"), _
    Format(Weekday(2, vbSunday), "dddd"), _
    Format(Weekday(3, vbSunday), "dddd"), _
    Format(Weekday(4, vbSunday), "dddd"), _
    Format(Weekday(5, vbSunday), "dddd"), _
    Format(Weekday(6, vbSunday), "dddd"), _
    Format(Weekday(7, vbSunday), "dddd"))
    
    For i = 1 To UBound(Bul)
        If SpinButton1.Value = i Then
            Bulan.Caption = Bul(i)
                For k = 1 To UBound(H)
                Minggu = 1
                Senin = 2
                Selasa = 3
                Rabu = 4
                Kamis = 5
                Jumat = 6
                Sabtu = 7
                        If Format(BVal(i), "dddd") _
                        = H(k) Then
                            Select Case H(k)
                                Case Format(Weekday(1, vbSunday), "dddd")
                                    DD
                                    For D = Minggu To Hit(i)
                                        With Controls("Label" & D)
                                            .Caption = D
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date)).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(2, vbSunday), "dddd")
                                    DD
                                    For D = Senin To Hit(i) + 1
                                        With Controls("Label" & D)
                                            .Caption = D - 1
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 1).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0

                                    Exit Sub
                                Case Format(Weekday(3, vbSunday), "dddd")
                                    DD
                                    For D = Selasa To Hit(i) + 2
                                        With Controls("Label" & D)
                                            .Caption = D - 2
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 2).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(4, vbSunday), "dddd")
                                    DD
                                    For D = Rabu To Hit(i) + 3
                                        With Controls("Label" & D)
                                            .Caption = D - 3
                                            .Font.Bold = False
                                        End With
                                    Next
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 3).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(5, vbSunday), "dddd")
                                    DD
                                    For D = Kamis To Hit(i) + 4
                                        With Controls("Label" & D)
                                            .Caption = D - 4
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 4).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(6, vbSunday), "dddd")
                                    DD
                                    For D = Jumat To Hit(i) + 5
                                        With Controls("Label" & D)
                                            .Caption = D - 5
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 5).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
                                    Exit Sub
                                Case Format(Weekday(7, vbSunday), "dddd")
                                    DD
                                    For D = Sabtu To Hit(i) + 6
                                        With Controls("Label" & D)
                                            .Caption = D - 6
                                            .Font.Bold = False
                                        End With
                                    Next D
                                    If Bul(i) = Format(Date, "mmmm") And _
                                    Val(Years) = Year(Date) Then
                                        Controls("Label" & Day(Date) + 6).Font.Bold = True
                                    End If
i = 0
k = 0
D = 0
Minggu = 0
Senin = 0
Selasa = 0
Rabu = 0
Kamis = 0
Jumat = 0
Sabtu = 0
 
                                    Exit Sub
                            End Select
                            
                        End If
                Next k
        End If
    Next i
End Sub

Private Sub Years_Change()
    With Years
        If .Locked = True Then
            .Locked = False
            SpinButton1_Change
        Else
            SpinButton1_Change
            .Locked = True
        End If
    End With
End Sub

Private Sub Years_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Years
        If .Locked = True Then
            .Locked = False
        Else
            .Locked = True
        End If
    End With
End Sub
