Attribute VB_Name = "Module1"
Sub tickerandtotal()

Dim w As Integer
w = Application.Worksheets.Count

For j = 1 To w

Worksheets(j).Activate

    Dim ticker As String
    Dim yearlychange As Double
    Dim openingprice As Double
    Dim closingprice As Double
    Dim percentchange As Double
    Dim totalstock As Variant
    totalstock = 0
    Dim summarytable As Integer
    summarytable = 2
    Dim firstrow As Boolean
    firstrow = True
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   
    For I = 2 To lastrow
   
            If firstrow = True Then
            openingprice = Cells(I, 3).Value
            firstrow = False
            End If
           
            totalstock = totalstock + Cells(I, 7).Value
       
            'If the next row starts with a different ticker:

            If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
   
            ticker = Cells(I, 1).Value
            closingprice = Cells(I, 6).Value
            firstrow = True
       
            yearlychange = closingprice - openingprice
            percentagechange = (closingprice - openingprice) / openingprice
       
            percentagechange = Format(percentagechange, "0.00%")
            
            'Print into columns
            Range("I" & summarytable).Value = ticker
   
            Range("L" & summarytable).Value = totalstock
       
            Range("J" & summarytable).Value = yearlychange
       
            Range("K" & summarytable).Value = percentagechange
            
            If Cells(summarytable, 10).Value > 0 Then
            Cells(summarytable, 10).Interior.ColorIndex = 4
            ElseIf Cells(summarytable, 10).Value < 0 Then
            Cells(summarytable, 10).Interior.ColorIndex = 3
            End If
            
            If Cells(summarytable, 11).Value > 0 Then
            Cells(summarytable, 11).Interior.ColorIndex = 4
            ElseIf Cells(summarytable, 11).Value < 0 Then
            Cells(summarytable, 11).Interior.ColorIndex = 3
           
            End If
            'add one row to summarytable
            summarytable = summarytable + 1
            
            'reset counter
            totalstock = 0
            End If
           
           'add conditional formatting
           
           
            
           
    Next I
       
     'find the greatest increase/decrease
     
     Range("Q4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
     inc_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
     Range("P4") = Cells(inc_index + 1, 9)
             
     Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
     inc_index = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
     Range("P2") = Cells(inc_index + 1, 9)
     
     Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
     inc_index = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
     Range("P3") = Cells(inc_index + 1, 9)
     
Next j

End Sub


