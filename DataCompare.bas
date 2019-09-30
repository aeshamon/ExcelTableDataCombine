Attribute VB_Name = "Module1"
Public CalcState As Long
Public EventState As Boolean
Public PageBreakState As Boolean

Sub OptimizeCode_Begin()

Application.ScreenUpdating = False

EventState = Application.EnableEvents
Application.EnableEvents = False

CalcState = Application.Calculation
Application.Calculation = xlCalculationManual

PageBreakState = ActiveSheet.DisplayPageBreaks
ActiveSheet.DisplayPageBreaks = False

End Sub

Sub OptimizeCode_End()

ActiveSheet.DisplayPageBreaks = PageBreakState
Application.Calculation = CalcState
Application.EnableEvents = EventState
Application.ScreenUpdating = True

End Sub

Sub Test2()
    'Optimize Code
   'Call OptimizeCode_Begin
   
   Dim CEL As Range, RANG As Range
   Dim dict As New Scripting.Dictionary
   With Worksheets("Sheet2")

       ' Build a range (RANG) between cell F2 and the last cell in column F
       Set RANG = Range(.Cells(2, "A"), .Cells(.Rows.Count, "A").End(xlUp))

   End With

   ' For each cell (CEL) in this range (RANG)
   For Each CEL In RANG

       If CEL.Value <> "" Then ' ignore blank cells

           If Not dict.Exists(CEL.Value) Then ' if the value hasn't been seen yet
               dict.Add CEL.Value, CEL ' add the value and first-occurrence-of-value-cell to the dictionary
           Else ' if the value has already been seen
               CEL.Offset(, 30).Value = "Duplicate Found" ' 2nd instance set the value of the cell 1 across to the right of CEL (i.e. column G) as "Duplicate Found"
               dict(CEL.Value).Offset(, 30).Value = "Duplicate Found" '1st instance set the value of the cell 1 across to the right of first-occurrence-of-value-cell (i.e. column G) as "Duplicate Found"
               'New if statemnet to check which of the two cells in the same column has the higher string length and moves it to a new sheet
               Dim i As Integer
               For i = 1 To 29
                   Set FirstInt = dict(CEL.Value).Offset(, i)
                   Set SecondInt = CEL.Offset(, i)
                   If Len(dict(CEL.Value).Offset(, i).Value) >= Len(CEL.Offset(, i)) Then
                       Range(FirstInt.Address).Cut Worksheets("Sheet3").Range(FirstInt.Address)
                   Else
                       Range(SecondInt.Address).Cut Worksheets("Sheet3").Range(FirstInt.Address)
                   End If
                Next i
                i = 0
           
           Set FirstInt = dict(CEL.Value).Offset(, 0)
           Set SecondInt = CEL.Offset(, 0)
           Range(SecondInt.Address).Cut Worksheets("Sheet3").Range(FirstInt.Address)
           dict.Remove (FirstInt)
           FirstInt.EntireRow.ClearContents
           FirstInt.Offset(1, 0).EntireRow.ClearContents
          
        
           

           End If
       End If
   Next CEL
   Set dict = Nothing
   Range("A1").EntireRow.Copy Worksheets("Sheet3").Range("A1")
   Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
   Worksheets("Sheet3").Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
   
   'Optimize Code
  Call OptimizeCode_End

   
End Sub

Sub Combine()
   Dim J As Integer
   Dim s As Worksheet

   On Error Resume Next
   Sheets(1).Select
   Worksheets.Add ' add a sheet in first place
   Sheets(1).Name = "Combined"

   ' copy headings
   Sheets(2).Activate
   Range("A1").EntireRow.Select
   Selection.Copy Destination:=Sheets(1).Range("A1")

   For Each s In ActiveWorkbook.Sheets
       If s.Name <> "Combined" Then
           Application.GoTo Sheets(s.Name).[A1]
           Selection.CurrentRegion.Select
           ' Don't copy the headings
           Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
           Selection.Copy Destination:=Sheets("Combined"). _
             Cells(Rows.Count, 1).End(xlUp)(2)
       End If
   Next
End Sub




