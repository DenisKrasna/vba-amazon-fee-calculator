' Module: AmazonFeeCalculator
' Description: Calculates Amazon commission fees for a list of product prices and displays total

Sub CalculateAmazonFees()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim ws As Worksheet
    Set ws = wb.Worksheets("CeneAmazon") ' Sheet name = "CeneAmazon"

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
    End With

    ' Write prices and calculate commissions
    Dim i As Integer
    For i = LBound(productPrice) To UBound(productPrice)
        ws.Cells(i + 1, 1) = productPrice(i)
        ws.Cells(i + 1, 2) = commissionRate * productPrice(i)
    Next i

    ' Calculate total row position
    Dim lastRow As Long
    lastRow = UBound(productPrice) + 2

    ' Write total
    With ws
        .Range("A" & lastRow).Value = totalLabel
        .Range("B" & lastRow).Formula = "=SUM(B2:B6)"
    End With

    ' Show result
    MsgBox "Total Amazon commission is: " & ws.Range("B" & lastRow).Value

End Sub
