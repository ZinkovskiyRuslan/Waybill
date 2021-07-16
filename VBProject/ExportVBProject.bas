Attribute VB_Name = "ExportVBProject"
Private Const subPath = "\VBProject\"

Public Sub ExportVBProject()
    Dim objVBComp As Variant
    'If Len(Dir(ActiveWorkbook.path & subPath, vbDirectory)) = 0 Then MkDir ActiveWorkbook.path & subPath
    
    If Len(Dir(ActiveWorkbook.path & subPath, vbDirectory)) <> 0 Then
        For Each varItem In ActiveWorkbook.VBProject.VBComponents
            Set objVBComp = varItem
            Select Case varItem.Type
                Case 1 'vbext_ct_StdModule
                    exportModule objVBComp, ".bas"
                Case 100 'vbext_ct_Document, vbext_ct_ClassModule
                    exportModule objVBComp, ".cls"
                Case vbext_ct_MSForm
                    exportModule objVBComp, ".frm"
                Case Else
                    exportModule objVBComp, ".ign"
            End Select
        Next varItem
    End If
End Sub

Private Sub exportModule(ByVal objVBComp, ext)
    Dim path As String
    Dim ret As Integer
    path = ActiveWorkbook.path & subPath
    objVBComp.Export path & objVBComp.Name & ext
    ret = ChangeFileCharset(path & objVBComp.Name & ext, "utf-8", "Windows-1251")
End Sub

 Function ChangeFileCharset(ByVal filename$, ByVal DestCharset$, _
                           Optional ByVal SourceCharset$) As Boolean
    ' функция перекодировки (смены кодировки) текстового файла
    ' В качестве параметров функция получает путь filename$ к текстовому файлу,
    ' и название кодировки DestCharset$ (в которую будет переведён файл)
    ' Функция возвращает TRUE, если перекодировка прошла успешно
    On Error Resume Next: Err.Clear
    With CreateObject("ADODB.Stream")
        .Type = 2
        If Len(SourceCharset$) Then .Charset = SourceCharset$    ' указываем исходную кодировку
        .Open
        .LoadFromFile filename$    ' загружаем данные из файла
        FileContent$ = .ReadText   ' считываем текст файла в переменную FileContent$
        .Close
        .Charset = DestCharset$    ' назначаем новую кодировку
        .Open
        .WriteText FileContent$
        .SaveToFile filename$, 2   ' сохраняем файл уже в новой кодировке
        .Close
    End With
    ChangeFileCharset = Err = 0
End Function
