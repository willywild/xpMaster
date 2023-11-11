Attribute VB_Name = "exportCode"
Option Explicit

Public Sub xpExportAllVBAcode()
'Exports all code in open Workbooks, including Worksheet and Workbook code
'target directory %appdata%\XpSearch\VBexport '20170910
'dano 20200809, 20170910, 20160115
     
'// reference to [Microsoft Visual Basic for Applications Extensibility]

'// check to 'Trust' VBE object model, then turn off
        Debug.Print "Enable Trust Access to the VBA Project model"
        Application.CommandBars.ExecuteMso "MacroSecurity"  '// turn off macroSecurity
'       Application.CommandBars.FindControl(ID:=3627).Execute  '//same thing
        Debug.Print "Select Addins to Export Code"
        Application.Dialogs(321).Show   '// select Addins to export

    Dim proj As VBProject
    Dim comp As VBComponent
    Dim folderName As String
    
''    folderName = Environ("APPDATA") & "\XpSearch\VBexports\"
    folderName = Environ("APPDATA") & "\Git\"
    Debug.Print vbLf; "Type", "#lines", "Project:Module"
    
    For Each proj In Application.VBE.VBProjects
        
        Select Case True
        Case proj.Protection = vbext_pp_locked
            MsgBox proj.Name & " is locked", vbCritical, "VBProject is password protected - skipped"
        Case Not proj.Saved
            MsgBox proj.Name & " is not saved", vbCritical, "Unsaved Project - save and retry"
        Case Else
            For Each comp In proj.VBComponents
                With comp
                    
                Select Case True
                Case .CodeModule.CountOfLines < 4
                    '// skip
                Case .Type = vbext_ct_StdModule
                    Debug.Print ".", .CodeModule.CountOfLines, proj.Name & "_" & .Name & ".bas"
                    .Export folderName & proj.Name & "_" & .Name & ".bas"
                Case .Type = vbext_ct_Document
                    Debug.Print "wb/ws", .CodeModule.CountOfLines, proj.Name & "_" & .Name & ".vb"
                    .Export folderName & proj.Name & "_" & .Name & ".vb"
                Case .Type = vbext_ct_ClassModule
                    Debug.Print "cls", .CodeModule.CountOfLines, proj.Name & "_" & .Name & ".cls"
                    .Export folderName & proj.Name & "_" & .Name & ".cls"
                Case .Type = vbext_ct_MSForm
                    Debug.Print "frm", .CodeModule.CountOfLines, proj.Name & "_" & .Name & ".frm"
                    .Export folderName & proj.Name & "_" & .Name & ".frm"
                Case Else       '// .Type = vbext_ct_ActiveXDesigner
                    Debug.Assert False
                End Select
                    
                End With
            Next comp
        End Select
        
    Next proj
    
    Debug.Print "Disable Trust Access to the VBA Project model for safety"
    Application.CommandBars.ExecuteMso "MacroSecurity"
    Debug.Print "Unselect Unused Addins"
    Application.Dialogs(321).Show   '// select Addins to export
End Sub
