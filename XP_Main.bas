Attribute VB_Name = "Main"
'// Event routines are in the Workbook object
Option Explicit

Public ctrlXpsearch As CommandBarControl

Private Const TAGXP As String = "XP"         'name of this control Addin
Private Const XPSEARCH As String = "XpSearch"           'name of operational Addin
Public Sub deleteXPcontrols()    '// delete all cutsom controls - 'XP' control
    Dim col As CommandBarControls
    Dim c

    Set col = CommandBars.FindControls(tag:=TAGXP)
    If Not col Is Nothing Then
        For Each c In col
            c.Delete
        Next c
    End If
End Sub

Public Sub xpF3()
    If Not ActiveWindow Is Nothing Then ActiveWindow.ActivateNext
End Sub
Public Sub xpF6()   '// Toggle AutoFilter, TopRow, SplitWindow, Freeze panes
    If Not TypeName(ActiveSheet) = "Worksheet" Then Exit Sub
    With ActiveWindow
        Select Case False
            Case .Split
                .ScrollRow = .ActiveSheet.UsedRange.Cells(1).Row
                .SplitColumn = 0
                .SplitRow = 1
                .FreezePanes = True
            Case .ActiveSheet.AutoFilterMode Or IsEmpty(.ActiveSheet.UsedRange.Cells(1))
                .ActiveSheet.UsedRange.Cells(1).AutoFilter
            Case Else
                .ActiveSheet.AutoFilterMode = False
                .FreezePanes = False
                .Split = False
        End Select
    End With
End Sub

Public Sub xpF5()
    Select Case True
        Case Not ActiveChart Is Nothing     '// on a chart - save to png
            xpChartSavedAsPng
        Case Not ActiveCell Is Nothing      '// in a worksheet cell
            Select Case True
                Case ActiveCell.Hyperlinks.Count = 1
                    xpFollowHyperlink
                Case ActiveCell.ListObject Is Nothing
                Case ActiveCell.ListObject.AutoFilter Is Nothing
                Case Else
                    ActiveCell.ListObject.AutoFilter.ApplyFilter
            End Select
    End Select
End Sub

Public Sub xpF7()
    If TypeName(ActiveSheet) = "Worksheet" Then ActiveSheet.UsedRange.Select
End Sub

Public Sub xpF8()
    If TypeName(ActiveSheet) = "Worksheet" Then ActiveSheet.UsedRange.Select
End Sub

Public Sub xpBuiltInMenusPopup()    'F1 key pulls up All menus :)
    Application.CommandBars("Built-in Menus").ShowPopup
End Sub

Public Sub xpFollowHyperlink()  'F5 follows hyperlink in cell - Same as Regedit 'ForceShellExecute' key
    With ActiveCell
        If .Hyperlinks.Count = 1 Then
            Shell Environ("ProgramW6432") & "\Mozilla Firefox\firefox.exe " & .Hyperlinks(1).Address
            .Font.ThemeColor = xlThemeColorFollowedHyperlink
        ElseIf InStr(.Text, "linkedin.com") > 0 Then
            Shell Environ("ProgramW6432") & "\Mozilla Firefox\firefox.exe " & .Text
            .Font.ThemeColor = xlThemeColorFollowedHyperlink
        End If
    End With
''    ActiveWorkbook.FollowHyperlink "https://www.linkedin.com/in/reidhoffman/"
End Sub

Public Sub xpChartSavedAsPng()
    Dim s As String
    Const EXT As String = "jpg"
''    Const EXT As String = "png" '// 'gif'
''    Const EXT As String = "gif" '// 'gif'
    
''For Each ch In ActiveWorkbook.Charts: ch.Export Filename:=ActiveWorkbook.Path & "\" & ch.Name & ".png": Next
    s = ActiveWorkbook.Path & "\" & ActiveChart.Name & "." & EXT
    ActiveChart.Export Filename:=s, FilterName:=EXT
    MsgBox s
End Sub

''Public Sub XpSearchOff()    '// turn off XpSearch on Excel boot
''    On Error Resume Next
''    If AddIns(XPSEARCH).Installed = True Then AddIns(XPSEARCH).Installed = False
''    If Err.Number <> 0 Then
''        If Err.Number = 9 Then
''            MsgBox ("XpSearch Addin is not installed properly in the directory" & vbNewLine & _
''                "Please check the Application.Addins(n).IsAddIn property is set = True")
''        Else
''            MsgBox ("Uh oh, this error needs debug, dano [Err.Number]=" & Err.Number)
''        End If
''    End If
''    On Error GoTo 0
''''    deleteControls XPSEARCH
''End Sub

''Public Sub XpSearch_OnOff()     'toggle XpSearch addin
''    If Application.EnableEvents = False Then
''        Application.EnableEvents = True: MsgBox "Events are off? - Please Toggle XpSearch on/off"
''    End If
''    If AddIns(XPSEARCH).Installed = True Then   'if On, turn Off
''        On Error Resume Next    'trap error if fails
''        AddIns(XPSEARCH).Installed = False
''        If Err.Number <> 0 Then
''            If Err.Number = 0 Then
''                MsgBox "XpSearch Addin is not installed properly in the directory" & vbNewLine & _
''                    "Please check the Application.IsAddIn property is set = True"
''            Else
''                MsgBox "Uh oh, this error needs debug, dano [Err.Number]=" & Err.Number
''            End If
''        End If
''        On Error GoTo 0
''        With ctrlXpsearch
''            .Caption = XPSEARCH & " On"
''            .TooltipText = "Turn On"
''        End With
''    Else                        'if Off, then Turn On
''        AddIns(XPSEARCH).Installed = True
''        With ctrlXpsearch
''            .Caption = XPSEARCH & " Off"
''            .TooltipText = "Turn Off"
''        End With
''
''    End If
''End Sub
''Public Sub deleteXpControls()
''    deleteControls TAGXP
''End Sub

''Public Sub installXpControl()   'install temporaty XP button on right-click menu & commandbar
''    With CommandBars("Worksheet Menu Bar").Controls.Add(Type:=msoControlButton, Temporary:=True)
''        .Style = msoButtonCaption
''        .Caption = "&" & XPSEARCH & " On"
''        .OnAction = XPSEARCH & "_OnOff"
''        .TooltipText = XPSEARCH & " On"
''        .tag = TAGXP
''''        .Copy CommandBars("cell")     '// loses Temporary property on copy
''    End With
''    Set ctrlXpsearch = CommandBars("Cell").Controls.Add(Type:=msoControlButton, Before:=1, Temporary:=True)
''    With ctrlXpsearch
''        .Style = msoButtonCaption
''        .Caption = "&" & XPSEARCH & " On"
''        .OnAction = XPSEARCH & "_OnOff"
''        .TooltipText = XPSEARCH & " On"
''        .tag = TAGXP
''    End With
''End Sub
''Private Sub deleteControls(str As String)    '// delete all cutsom controls - 'XP' control
''    Dim col As CommandBarControls
''    Dim c
''
''    Set col = CommandBars.FindControls(tag:=str)
''    If Not col Is Nothing Then
''        For Each c In col
''            c.Delete
''        Next c
''    End If
''End Sub
