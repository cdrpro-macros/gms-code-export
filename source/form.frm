VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form 
   Caption         =   "CodeExport"
   ClientHeight    =   5520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Microsoft Visual Basic for Applications Extensibility 5.3
Private vb As Object

Private Sub UserForm_Initialize()
    Dim vbp As Object 'VBProject
    Set vb = Application.VBE
    For Each vbp In vb.VBProjects
        list.AddItem vbp.Name
    Next
End Sub

Private Sub exp_Click()
    Dim path$: path = CorelScriptTools.GetFolder
    If Len(path) = 0 Then Exit Sub
    
    Dim vbp As Object 'VBProject
    Set vbp = vb.VBProjects.Item(list.ListIndex + 1)

    Dim vc As Object 'VBComponent
    For Each vc In vbp.VBComponents
        Dim fileName$
        fileName = vc.Name

        Select Case vc.Type
            Case 1 'vbext_ct_StdModule
                fileName = fileName & ".bas"
            Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
                fileName = fileName & ".cls"
            Case 3 'vbext_ct_MSForm
                fileName = fileName & ".frm"
        End Select

        vc.Export path & "\" & fileName
    Next
    
    FileSystem.FileCopy vbp.fileName, path & "\" & vbp.Name & ".gms"
    MsgBox "Done!"
End Sub



