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
' --------- Проверка защиты проекта --------------
    If vbp.Protection Then
        MsgBox "Проект защищен"
        Exit Sub
    End If
'-------------------------------------------------
'------- Для избежания появления неорганизованной кучи 
'--------добавлена возможность создания новой папки для экспортируемого макроса.
'--------Название папки имеет формат ИмяМакроса_Год_Месяц_День_Час_Минута
'--------(так папки сами упорядочиваются в хронологическом порядке)
    If Flag_NewFolder Then ' на форме добавлен чекбокс Flag_NewFolder
        If Len(Month(Now)) = 1 Then
            cMonth = "0" & Month(Now)
        Else
            cMonth = Month(Now)
        End If
        If Len(Day(Now)) = 1 Then
            cDay = "0" & Day(Now)
        Else
            cDay = Day(Now)
        End If
        If Len(Hour(Now)) = 1 Then
            cHour = "0" & Hour(Now)
        Else
            If Len(Hour(Now)) = 0 Then
                cHour = "00" + Hour(Now)
            Else
                cHour = Hour(Now)
            End If
        End If
        If Len(Minute(Now)) = 1 Then
            cMinute = "0" & Minute(Now)
        Else
            If Len(Minute(Now)) = 0 Then
                cMinute = "00" & Minute(Now)
            Else
                cMinute = Minute(Now)
            End If
        End If
        Path = Path & "\" & vbp.Name & "_" & Year(Now) & "_" & cMonth & "_" & cDay & "__" & cHour & "_" & cMinute
        MkDir Path
    End If
'-------------------------------------------------
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



