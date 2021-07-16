VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Const Waybill = "Путевые листы"
Private Const List = "Списки"
Private Const Controller = "Диспетчеры"
Private Const Driver = "Водители"
Private Const Doctor = "Мед.Работники"
Private Const Mechanic = "Механики"
Private Const Company = "Организации"
Private Const Auto = "Автомобили"
Private Const Pattern4 = "Шаблон грузового авто.стр1"
Private Const Pattern2 = "Шаблон спец. авто.стр1"


Private Sub AddNewWaybill_Click()
    GetSheet(Waybill).Rows("3:3").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow
    GetSheet(Waybill).Range("A3").Value = GetSheet(Waybill).Range("A4").Value
    GetSheet(Waybill).Range("B3").Value = GetSheet(Waybill).Range("B4").Value
    GetSheet(Waybill).Range("I3").Value = "На печать"
    
    Set rangeInterior = GetSheet(Waybill).Range("A3:H3").Interior
    rangeInterior.Pattern = xlNone
    rangeInterior.TintAndShade = 0
    rangeInterior.PatternTintAndShade = 0
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Len(GetSheet(Waybill).Range("B" + CStr(Target.Row)).Value) < 5 Or Target.Column <> 9 Then
        SendKeys ("{Esc}")
        Exit Sub
    End If
    
    SendKeys ("{Esc}")
    Set rangeInterior = GetSheet(Waybill).Range("A" + CStr(Target.Row) + ":H" + CStr(Target.Row)).Interior
    rangeInterior.Pattern = xlSolid
    rangeInterior.PatternColorIndex = xlAutomatic
    rangeInterior.ThemeColor = xlThemeColorDark1
    rangeInterior.TintAndShade = -4.99893185216834E-02
    rangeInterior.PatternTintAndShade = 0
    
    GetSheet(Pattern2).PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
    SendKeys ("{Esc}")
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column = 9 Then
        If GetSheet(Waybill).Range("B3").Value = "Грузовой Автомобиль (0345004)" Then
            FillPattern0345004
        End If
        If GetSheet(Waybill).Range("B3").Value = "Специальный Автомобиль (0345002)" Then
            FillPattern0345002
        End If
    End If
End Sub

Private Sub FillPattern0345004()
    Dim rowNumber As Integer
    Set Pa = GetSheet(Pattern4)
    
    'Date
    GetSheet(Pattern4).Range("BG5").Value = "'" + CStr(DatePart("d", CDate(GetSheet(Waybill).Range("A3"))))
    GetSheet(Pattern4).Range("BP5").Value = MonthName(DatePart("m", CDate(GetSheet(Waybill).Range("A3"))))
    GetSheet(Pattern4).Range("CL5").Value = "'" + CStr(DatePart("yyyy", CDate(GetSheet(Waybill).Range("A3"))))
    
    'Company
    GetSheet(Pattern4).Range("Q6").Value = GetSheet(Waybill).Range("C3").Value
    
    'Auto
    rowNumber = GetRowNumber(Auto, GetSheet(Waybill).Range("D3").Value)
    GetSheet(Pattern4).Range("Q13").Value = GetSheet(Auto).Range("B1").Offset(rowNumber, 0).Value
    GetSheet(Pattern4).Range("AB14").Value = GetSheet(Auto).Range("C1").Offset(rowNumber, 0).Value
    
    'Driver
    rowNumber = GetRowNumber(Driver, GetSheet(Waybill).Range("E3").Value)
    GetSheet(Pattern4).Range("I15").Value = _
        GetSheet(Driver).Range("B1").Offset(rowNumber, 0).Value + " " + _
        GetSheet(Driver).Range("C1").Offset(rowNumber, 0).Value + " " + _
        GetSheet(Driver).Range("D1").Offset(rowNumber, 0).Value
    GetSheet(Pattern4).Range("P17").Value = GetSheet(Driver).Range("E1").Offset(rowNumber, 0).Value
    GetSheet(Pattern4).Range("AS17").Value = GetSheet(Driver).Range("F1").Offset(rowNumber, 0).Value
    
    'Controller
    rowNumber = GetRowNumber(Controller, GetSheet(Waybill).Range("H3").Value)
    GetSheet(Pattern4).Range("V46").Value = _
        GetSheet(Controller).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Controller).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Controller).Range("D1").Offset(rowNumber, 0).Value, 1) + "."
        
    'Doctor
    rowNumber = GetRowNumber(Doctor, GetSheet(Waybill).Range("G3").Value)
    GetSheet(Pattern4).Range("AI50").Value = _
        GetSheet(Doctor).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Doctor).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Doctor).Range("D1").Offset(rowNumber, 0).Value, 1) + "."

    'Mechanic
    rowNumber = GetRowNumber(Mechanic, GetSheet(Waybill).Range("F3").Value)
    GetSheet(Pattern4).Range("CO44").Value = _
        GetSheet(Mechanic).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Mechanic).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Mechanic).Range("D1").Offset(rowNumber, 0).Value, 1) + "."

    'Driver
    rowNumber = GetRowNumber(Driver, GetSheet(Waybill).Range("E3").Value)
    GetSheet(Pattern4).Range("CO46").Value = _
        GetSheet(Driver).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Driver).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Driver).Range("D1").Offset(rowNumber, 0).Value, 1) + "."
End Sub

Private Sub FillPattern0345002()
    Dim rowNumber As Integer
    Set Pattern = GetSheet(Pattern2)
    
    'Fill Date Left
    Pattern.Range("O7").Value = "'" + CStr(DatePart("d", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("X7").Value = MonthName(DatePart("m", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("AQ7").Value = "'" + CStr(DatePart("yyyy", CDate(GetSheet(Waybill).Range("A3"))))
    'Fill Date Right Top
    Pattern.Range("EW6").Value = "'" + CStr(DatePart("d", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("FD6").Value = MonthName(DatePart("m", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("FT6").Value = "'" + CStr(DatePart("yyyy", CDate(GetSheet(Waybill).Range("A3"))))
    'Fill Date Right Down
    Pattern.Range("EX32").Value = "'" + CStr(DatePart("d", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("FD32").Value = MonthName(DatePart("m", CDate(GetSheet(Waybill).Range("A3"))))
    Pattern.Range("FT32").Value = "'" + CStr(DatePart("yyyy", CDate(GetSheet(Waybill).Range("A3"))))
    
    'Fill Company Left
    rowNumber = GetRowNumber(Company, GetSheet(Waybill).Range("C3").Value)
    Pattern.Range("P9").Value = GetSheet(Company).Range("B1").Offset(rowNumber, 0).Value
    Pattern.Range("A12").Value = GetSheet(Company).Range("C1").Offset(rowNumber, 0).Value + _
        ". Тел: " + GetSheet(Company).Range("D1").Offset(rowNumber, 0).Value
    'Fill Company Right Top
    Pattern.Range("FE7").Value = GetSheet(Company).Range("B1").Offset(rowNumber, 0).Value
    'Fill Company Right Down
    Pattern.Range("FD33").Value = GetSheet(Company).Range("B1").Offset(rowNumber, 0).Value
    
    'Auto Left
    rowNumber = GetRowNumber(Auto, GetSheet(Waybill).Range("D3").Value)
    Pattern.Range("W14").Value = GetSheet(Auto).Range("B1").Offset(rowNumber, 0).Value
    Pattern.Range("AK15").Value = GetSheet(Auto).Range("C1").Offset(rowNumber, 0).Value
    'Auto Right Top
    Pattern.Range("FJ9").Value = GetSheet(Auto).Range("B1").Offset(rowNumber, 0).Value
    Pattern.Range("FU10").Value = GetSheet(Auto).Range("C1").Offset(rowNumber, 0).Value
    'Auto Right Down
    Pattern.Range("FI35").Value = GetSheet(Auto).Range("B1").Offset(rowNumber, 0).Value
    Pattern.Range("FW37").Value = GetSheet(Auto).Range("C1").Offset(rowNumber, 0).Value
    
    'Driver
    rowNumber = GetRowNumber(Driver, GetSheet(Waybill).Range("E3").Value)
    Pattern.Range("M16").Value = _
        GetSheet(Driver).Range("B1").Offset(rowNumber, 0).Value + " " + _
        GetSheet(Driver).Range("C1").Offset(rowNumber, 0).Value + " " + _
        GetSheet(Driver).Range("D1").Offset(rowNumber, 0).Value
    Pattern.Range("U18").Value = GetSheet(Driver).Range("E1").Offset(rowNumber, 0).Value
    Pattern.Range("CO18").Value = GetSheet(Driver).Range("F1").Offset(rowNumber, 0).Value

    'Controller
    rowNumber = GetRowNumber(Controller, GetSheet(Waybill).Range("H3").Value)
    Pattern.Range("AD43").Value = _
        GetSheet(Controller).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Controller).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Controller).Range("D1").Offset(rowNumber, 0).Value, 1) + "."

    'Doctor
    rowNumber = GetRowNumber(Doctor, GetSheet(Waybill).Range("G3").Value)
    Pattern.Range("AD47").Value = _
        GetSheet(Doctor).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Doctor).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Doctor).Range("D1").Offset(rowNumber, 0).Value, 1) + "."

    'Mechanic
    rowNumber = GetRowNumber(Mechanic, GetSheet(Waybill).Range("F3").Value)
    Pattern.Range("DM41").Value = _
        GetSheet(Mechanic).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Mechanic).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Mechanic).Range("D1").Offset(rowNumber, 0).Value, 1) + "."
    Pattern.Range("DM48").Value = Pattern.Range("DM41").Value
    
    'Driver
    rowNumber = GetRowNumber(Driver, GetSheet(Waybill).Range("E3").Value)
    Pattern.Range("DM43").Value = _
        GetSheet(Driver).Range("B1").Offset(rowNumber, 0).Value + " " + _
        Left(GetSheet(Driver).Range("C1").Offset(rowNumber, 0).Value, 1) + ". " + _
        Left(GetSheet(Driver).Range("D1").Offset(rowNumber, 0).Value, 1) + "."
    Pattern.Range("DM50").Value = Pattern.Range("DM43").Value
End Sub
Private Function GetSheet(SheetName As String) As Worksheet
    Set GetSheet = ActiveWorkbook.Sheets(SheetName)
End Function

Private Function GetRowNumber(SheetName As String, FindValue) As Integer
    Dim i As Integer
    GetRowNumber = 0
    i = 1
    Set Sh = GetSheet(SheetName)
    Do While Len(Sh.Range("A1").Offset(i, 0).Value) > 0
        If Sh.Range("A1").Offset(i, 0).Value = FindValue Then
            GetRowNumber = i
            Exit Do
        End If
      i = i + 1
    Loop
End Function
