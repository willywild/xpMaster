Attribute VB_Name = "exportCode"
Option Explicit

Public Sub xpExportAllVBAcode()
'Exports all code in open Workbooks, including Worksheet and Workbook code
'target directory %appdata%\XpSearch\VBexport '20170910
'dano 20200809, 20170910, 20160115
     
' reference to "Microsoft Visual Basic for Applications Extensibility"

'// check to 'Trust' VBE object model, then turn off
''    On Error Resume Next
''    If Not Application.VBE.CommandBars.Count > 0 Then
        Debug.Print "Enable Trust Access to the VBA Project model"
        Application.CommandBars.ExecuteMso "MacroSecurity"  '// turn off macroSecurity
'       Application.CommandBars.FindControl(ID:=3627).Execute  '//same thing
        Debug.Print "Select Addins to Export Code"
        Application.Dialogs(321).Show   '// select Addins to export
''    End If
''    On Error GoTo 0

    Dim proj As VBProject
    Dim comp As VBComponent
    Dim folderName As String
    
    folderName = Environ("APPDATA") & "\XpSearch\VBexports\"
    
    
    Debug.Print vbLf; "Type", "#lines", "Project:Module"
    For Each proj In Application.VBE.VBProjects
        Do
            If proj.Protection <> vbext_pp_none Then Exit Do
            For Each comp In proj.VBComponents
                With comp
                    If .CodeModule.CountOfLines > 3 Then 'skips empty "Option Explicit" code modules
                        Debug.Print Switch( _
                            .Type = vbext_ct_StdModule, ".", _
                            .Type = vbext_ct_MSForm, "frm", _
                            .Type = vbext_ct_Document, "cls-wb/ws", _
                            .Type = vbext_ct_ClassModule, "cls", _
                            .Type = vbext_ct_ActiveXDesigner, "ActiveX"), _
                            .CodeModule.CountOfLines, _
                            proj.Name & ":" & .Name
                        If (.Type = vbext_ct_Document) And _
                                (.Name <> "ThisWorkbook") Then
                            .Export folderName & proj.Name & "_" & .Name & "_" & .Properties("Name") & ".vb"
                        Else
                            .Export folderName & proj.Name & "_" & .Name & ".vb"
                        End If
                    End If
                End With
            Next comp
        Loop Until True
    Next proj
    Debug.Print "Disable Trust Access to the VBA Project model for safety"
    Application.CommandBars.ExecuteMso "MacroSecurity"
    Debug.Print "Unselect Unused Addins"
    Application.Dialogs(321).Show   '// select Addins to export
End Sub
