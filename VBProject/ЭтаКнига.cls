VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    ExportVBProject.ExportVBProject
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    If Sh.Name = "Диспетчеры" Then
        FillWaybillFormula Sh, "H3"
    End If
    If Sh.Name = "Водители" Then
        FillWaybillFormula Sh, "E3"
    End If
    If Sh.Name = "Мед.Работники" Then
        FillWaybillFormula Sh, "G3"
    End If
    If Sh.Name = "Механики" Then
        FillWaybillFormula Sh, "F3"
    End If
    If Sh.Name = "Организации" Then
        FillWaybillFormula Sh, "C3"
    End If
    If Sh.Name = "Автомобили" Then
        FillWaybillFormula Sh, "D3"
    End If
End Sub

Private Sub FillWaybillFormula(ByVal Sh As Object, RangeList As String)
    Dim Waybill As Object
    Dim i As Integer
    
    i = 1
    Do While Len(Sh.Range("B1").Offset(i, 0).Value) > 0
      Sh.Range("A1").Offset(i, 0).Value = Sh.Range("B1").Offset(i, 0).Value + " " + Sh.Range("C1").Offset(i, 0).Value + " " + Sh.Range("D1").Offset(i, 0).Value
      i = i + 1
    Loop
    ActiveWorkbook.Sheets("Путевые листы").Range(RangeList)(1).Validation.Modify Formula1:="=" + Sh.Name + "!$A$2:$A$" + CStr(i)

End Sub

