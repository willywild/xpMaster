Attribute VB_Name = "exportCode"
Option Explicit

'// #INCLUDE: [Microsoft Visual Basic for Applications Extensibility]
'// #INCLUDE: [JS class module]

Private m As thisModule
Private Type thisModule
    s As String
    fldr As String
    gitfldr As String
End Type

Public Sub ExportAllVBAcode()
''    Dim wb As Excel.Workbook
    Dim userAddinsSelected As Boolean
    Dim i As Long
    
    '// Exports all code in open Workbooks and installed Addins
    '// including Worksheet XML and Workbook VBA code
    '// sheet XML files for data to rebuild sheets with formatting and formulas
    
    '// target directory:  %Appdata%\Git\name
    m.gitfldr = Environ("APPDATA") & "\Git\"
    If VBA.Len(VBA.Dir(m.gitfldr, vbDirectory)) = 0 Then VBA.MkDir m.gitfldr
    
    userAddinsSelected = isSelectedAddins '// skip Addins menu at end if no changes
    
    '// 'Trust' VBE object model, then turn off when finished
    If Not isVBEPermissionsOn Then MsgBox "cannot export without VBE permissions, exit", vbInformation: Exit Sub


    '// Note: VBProjects should be accessed using the Excel application and not VBE application
    '//       ... VBE.VBProject cannot refer up to parent workbook or worksheet Excel.Workbook.VBProject can
    
    '   ai always has code -> always export
    '   wb export if HasVBProject is True
    '   - check wb and each ws classes for vb code
    '   - export ws XML if any VBA code
    '   - has Modules or Class Modules or UserForms
    '   json object of properties, etc
    

    Debug.Print "Workbooks:"
    For i = 1 To Excel.Workbooks.Count
        With Workbooks(i)
            Debug.Print , IIf(.HasVBProject, .VBProject.Name, vbTab), IIf(.Saved, vbTab, "not-saved"), .Name
        
            Select Case True
            Case Not .HasVBProject
            Case Not .Saved: MsgBox .Name & " is not saved, skipped"
            Case .VBProject.Protection = vbext_pp_locked: MsgBox .Name & " is protected, skipped"
            Case Else
                exportWorkbook Workbooks(i)
            End Select
            
        End With
    Next i

    Debug.Print "Addins:"
''    Dim ai As Excel.AddIn
''    For Each ai In Excel.AddIns
    For i = 1 To Excel.AddIns.Count
        With AddIns(i)
            Select Case True
            Case Not .Installed: Debug.Print "  -off-", , , .Name
            Case Not Workbooks(.Name).Saved: MsgBox .Name & " is not-saved, skipped"
            Case Else
    ''            Debug.Print , Workbooks(ai.name).VBProject.name, "exported", ai.name
                exportWorkbook Workbooks(.Name)
            End Select
        End With
    Next i
    
    Debug.Print "COM Addins"
''    Dim cm As office.COMAddIn
''    For Each cm In Application.COMAddIns
    For i = 1 To Application.COMAddIns.Count
        With Application.COMAddIns(i)
        Debug.Print .GUID, .progID, .Description
        End With
    Next i
    
    If userAddinsSelected Then Debug.Print isSelectedAddins
    If Not isVBEPermissionsOff Then MsgBox "VBE permissions are on, dangerous", vbCritical
End Sub

Private Sub exportWorkbook(wb As Excel.Workbook)
''    Dim proj As VBIDE.VBProject
    Dim cm As VBIDE.VBComponent
    Dim s As String
    Dim ss As String
    Dim N As Long
    Dim i As Long
    Dim o, oo   '// JScriptTypeInfo
''    Dim sh As Worksheet
    
''    Debug.Print vbLf; "Type", "#lines", "Project_Module.ext"
    
    '[TODO] list wb structure json
    '       workbook, name, VBProject, isAddin, codename, properties, etc..
    '           export workbook component ifis
    '       list sheets, charts, of VBProject type document (not module, class, or form)
    '           export sheetXML ifis
    '           export component ifis
    '       list component ifis not vbp_document (already exported)
    '           export component ifis
    With wb.VBProject
        Set oo = JS.newObject("FileName", wb.Name)
        Debug.Assert JS.addItem(oo, "ProjectName", .Name) > 0 'project name: VBProject or 'JScript'
        Debug.Assert JS.addItem(oo, "isAddin", wb.IsAddin) > 0
        If VBA.Len(.Description) > 0 Then Debug.Assert JS.addItem(oo, "ProjectDescription", .Description) > 0
        If VBA.Len(wb.Author) > 0 Then Debug.Assert JS.addItem(oo, "Author", wb.Author) > 0
        Debug.Assert JS.addItem(oo, "FilePath", wb.Path) > 0
''        Debug.Assert JS.addItem(o, "ThisWorkBook", wb.CodeName) > 0  'ThisWorkbook
        Set o = JS.newObject("meta", oo)
        
        '// Git subfolder name and check it:
        m.s = .Name
        If m.s = "VBAProject" Then                      '// use filename instead of generic VBAProject
            m.s = Replace(.BuildFileName, ".DLL", vbNullString)
            m.s = Split(m.s, "\")(UBound(Split(m.s, "\")))
        End If
        m.fldr = m.gitfldr & m.s & "\"
        If VBA.Len(VBA.Dir(m.fldr, vbDirectory)) = 0 Then VBA.MkDir m.fldr
        
        Dim fso As Object
        Set fso = CreateObject("scripting.filesystemobject")

        Dim arr
        Set oo = JS.newObject("CodeName", wb.CodeName)
        Debug.Assert JS.addItem(oo, "TypeName", VBA.TypeName(wb)) > 0
        Debug.Print JS.arrayPush(arr, oo) > 0

        For i = 1 To wb.Sheets.Count
            With wb.Sheets(i)
                Set oo = JS.newObject("CodeName", .CodeName)
                Debug.Assert JS.addItem(oo, "Name", .Name) > 0
                Debug.Assert JS.addItem(oo, "TypeName", VBA.TypeName(wb.Sheets(i))) > 0
                
                Select Case True
                Case .Type <> xlWorksheet
                Case VBA.IsEmpty(.UsedRange)
                    '// skip blank sheets
                Case Else
                    Debug.Assert JS.addItem(oo, "UsedRange", .UsedRange.AddressLocal) > 0
                    fso.CreateTextFile(m.fldr & m.s & "_" & .CodeName & ".xml").Write .UsedRange.Value(xlRangeValueXMLSpreadsheet)
                End Select
            
                Debug.Assert JS.addItem(arr, i, oo) > 0
            End With
        Next i
        Debug.Assert JS.addItem(o, "ThisWorkbook", arr) > 0
        
        
        Set oo = JS.newObject(0)
        For Each cm In .VBComponents
            With cm
''                Debug.Print cm.Name
''                Debug.Print cm.CodeModule.CountOfLines
                
                Select Case True
                Case .CodeModule.CountOfLines < 3
                    '// skip empty modules which only have 2 lines of 'Option Explicit'
                Case .Type = vbext_ct_StdModule
                    Debug.Print ".", .CodeModule.CountOfLines, s & "_" & .Name & ".bas"
                    .Export m.fldr & m.s & "_" & .Name & ".bas"
                    Debug.Assert JS.addItem(oo, .Name & ".bas", .CodeModule.CountOfLines) > 0
                Case .Type = vbext_ct_Document
                    Debug.Print "wb/ws", .CodeModule.CountOfLines, s & "_" & .Name & ".vb"
                    .Export m.fldr & m.s & "_" & .Name & ".vb"
                    Debug.Assert JS.addItem(oo, .Name & ".vb", .CodeModule.CountOfLines) > 0
                Case .Type = vbext_ct_ClassModule
                    Debug.Print "cls", .CodeModule.CountOfLines, s & "_" & .Name & ".cls"
                    .Export m.fldr & m.s & "_" & .Name & ".cls"
                    Debug.Assert JS.addItem(oo, .Name & ".cls", .CodeModule.CountOfLines) > 0
                Case .Type = vbext_ct_MSForm
                    Debug.Print "frm", .CodeModule.CountOfLines, s & "_" & .Name & ".frm"
                    .Export m.fldr & m.s & "_" & .Name & ".frm"
                    Debug.Assert JS.addItem(oo, .Name & ".frm", .CodeModule.CountOfLines) > 0
                Case Else       '// .Type = vbext_ct_ActiveXDesigner
                    Debug.Assert False
                End Select
                
            End With
        Next cm
        Debug.Assert JS.addItem(o, "CodeModule", oo) > 0
        
        Debug.Assert JS.addItem(o, "References", addReferences(wb.VBProject)) > 0
        fso.CreateTextFile(m.fldr & m.s & ".json").Write JS.stringify(o, "", "  ")
        Debug.Print JS.stringify(o, "", "  ")
        
    End With
End Sub

Public Function addReferences(pj As VBIDE.VBProject) As Variant
    Dim i As Long
    Dim o
    Dim ret
    
    Set ret = JS.newObject(0)
    For i = 1 To pj.References.Count
        With pj.References(i)
            Set o = JS.newObject(0)
            If .IsBroken Then
                MsgBox .Name & " has a broken reference to: " & .Name, vbCritical
                Debug.Assert JS.addItem(o, "isBroken", .IsBroken) > 0
            End If
            Debug.Assert JS.addItem(o, "Description", .Description) > 0
            Debug.Assert JS.addItem(o, "Version", .Major & "." & .Minor) > 0
            Debug.Assert JS.addItem(o, "BuiltIn", .BuiltIn) > 0
            Debug.Assert JS.addItem(o, "GUID", .GUID) > 0
            Debug.Assert JS.addItem(o, "FullPath", .FullPath) > 0
            Debug.Assert .Type = vbext_rk_TypeLib
            Debug.Assert JS.addItem(ret, .Name, o) > 0
        End With
    Next i
    Set addReferences = ret
End Function

Public Function isSelectedAddins() As Boolean
    Dim i As Long
    Dim N As Long
    
    For i = 1 To Excel.AddIns.Count
        If AddIns(i).Installed Then N = N + i
    Next i
    Debug.Print "Select Addins to Export Code"
    Application.Dialogs(xlDialogAddinManager).Show  '// .Dialogs(321).Show
    '// check to see if Addins were selected/deselected
    For i = 1 To Excel.AddIns.Count
        If AddIns(i).Installed Then N = N - i
    Next i
    isSelectedAddins = (N <> 0)
End Function


Public Function isVBEPermissionsOn() As Boolean
        Debug.Print "Enable Trust Access to the VBE Project model"
        On Error Resume Next
            If Not Application.VBE.VBProjects.Count > 0 Then
                Application.CommandBars.ExecuteMso "MacroSecurity"  '// turn off macroSecurity
            '// Application.CommandBars.FindControl(ID:=3627).Execute  '//same thing
            Else
                Debug.Print "... already trusted"
            End If
        isVBEPermissionsOn = IsNumeric(Application.VBE.VBProjects.Count)
End Function

Public Function isVBEPermissionsOff() As Boolean
    Debug.Print "Disable Trust Access to the VBA Project model for safety"
    Application.CommandBars.ExecuteMso "MacroSecurity"
    On Error Resume Next
    Debug.Assert IsNumeric(Application.VBE.VBProjects.Count)
    isVBEPermissionsOff = IIf(Err.Number = 1004, True, False)
End Function
