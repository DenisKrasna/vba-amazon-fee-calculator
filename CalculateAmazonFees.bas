' Module: AmazonFeeCalculator
' Description: Calculates Amazon commission fees for a list of product prices and displays total

Sub CalculateAmazonFees()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws As Worksheet
    Set ws = wb.Worksheets("PriceAmazon") ' Sheet name = "CeneAmazon"
    
    ws.Cells.Clear
    
    Dim header1 As String, header2 As String, totalLabel As String
    header1 = "Product Price"
    header2 = "Commission"
    totalLabel = "Total Commission"
  
    Dim productPrice(1 To 5) As Double
    productPrice(1) = 15.99
    productPrice(2) = 24.5
    productPrice(3) = 39
    productPrice(4) = 12.75
    productPrice(5) = 19.99

    Const commissionRate As Double = 0.15

    ' Write headers
    With ws
        .Range("A1").Value = header1
        .Range("B1").Value = header2
        .Range("A1:B1").Interior.Color = RGB(255, 165, 0)
        .Range("A1:B1").Font.Bold = True
    End With


    ' Write prices and calculate commissions
    Dim i As Integer
    For i = LBound(productPrice) To UBound(productPrice)
        
        ws.Cells(i + 1, 1).Value = Round(productPrice(i), 2)
        ws.Cells(i + 1, 1).NumberFormat = "#,##0.00 €"
        ws.Cells(i + 1, 2).Value = Round(commissionRate * productPrice(i), 2)
        ws.Cells(i + 1, 2).NumberFormat = "#,##0.00 €"

    Next i

    ' Calculate total row position
    Dim lastRow As Long
    lastRow = UBound(productPrice) + 2

    ' Write total
    With ws
        .Range("A" & lastRow).Value = totalLabel
        .Range("A" & lastRow).Interior.Color = RGB(255, 255, 0)
        .Range("B" & lastRow).Interior.Color = RGB(255, 105, 108)
        .Range("A" & lastRow).Font.Bold = True
        .Range("B" & lastRow).Formula = "=SUM(B2:B6)"
        .Range("B" & lastRow).Font.Bold = True
        .Columns("A:B").AutoFit
        .Columns("A:B").HorizontalAlignment = xlCenter
    End With
    
    With ws.Range("A1:B7").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0) 'black border
    End With

    ' Show result
    MsgBox "Total Amazon commission is: " & FormatCurrency(ws.Range("B" & lastRow).Value, 2)

End Sub
