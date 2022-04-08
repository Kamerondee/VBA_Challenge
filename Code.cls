VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockanalysis()
   
    ' Set "ws" As Worksheet
   
    Dim ws As Worksheet
   
    ' Loop through all Worksheets
   
    For Each ws In Worksheets
   
    ' Identify your variables
   
    Dim ticker As String
    Dim lastrow As Long
    Dim sumtbl As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim price_difference As Double
    Dim count As Double
    Dim endrow As Double
   
    ' Bonus variables
   
    Dim greatestvol As Double
    Dim GTV As String
    Dim greatestinc As Double
    Dim GPI As String
    Dim greatestdec As Double
    Dim GPD As String

    ' Set your headers
   
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
   
    ' Set your counters
   
    voltotal = 0
    sumtbl = 2
    open_price = 0
    close_price = 0
    price_difference = 0
   
    ' Declare the last row
   
    lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row

    For i = 2 To lastrow

        ' Start first 'if' statement
       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ticker = ws.Cells(i, 1).Value

            voltotal = voltotal + ws.Cells(i, 7).Value
   
            ws.Range("I" & sumtbl).Value = ticker

            ws.Range("L" & sumtbl).Value = voltotal

            ' Grabs information from next row
           
            sumtbl = sumtbl + 1
           
            ' Reset counter
           
            voltotal = 0

        Else
       
            voltotal = voltotal + ws.Cells(i, 7).Value

        End If

    Next i

        sumtbl = 2

        count = -1

        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                close_price = ws.Cells(i, 6).Value

                price_difference = close_price - open_price

                ws.Range("J" & sumtbl).Value = price_difference

                    If open_price = 0 Then

                        ws.Range("K" & sumtbl).Value = 1

                        ws.Range("K" & sumtbl).NumberFormat = "0.00%"

                        sumtbl = sumtbl + 1

                        count = -1

                    Else

                        ws.Range("K" & sumtbl).Value = (price_difference / open_price)

                        ws.Range("K" & sumtbl).NumberFormat = "0.00%"

                        sumtbl = sumtbl + 1

                        count = -1

                    End If

            Else

                count = count + 1
           
                open_price = ws.Cells(i - count, 3).Value
           
            End If

        Next i
   
            endrow = ws.Cells(Rows.count, 9).End(xlUp).Row

        For i = 2 To endrow

            If ws.Cells(i, 10).Value > 0 Then

                ws.Cells(i, 10).Interior.ColorIndex = 4

            Else

                ws.Cells(i, 10).Interior.ColorIndex = 3

            End If
       
        Next i
       
            greatestvol = ws.Range("L2").Value

                For i = 3 To endrow

                    If ws.Cells(i, 12).Value > greatestvol Then

                        greatestvol = ws.Cells(i, 12).Value
                       
                        GTV = ws.Cells(i, 9).Value

                    End If

                Next i

                    ws.Cells(4, 17).Value = greatestvol
                   
                    ws.Cells(4, 16).Value = GTV
   
                    greatestinc = ws.Cells(2, 11).Value

                        For i = 3 To endrow

                            If (ws.Cells(i, 11).Value > greatestinc) Then

                                greatestinc = ws.Cells(i, 11).Value
                               
                                ws.Range("Q2").NumberFormat = "0.00%"
                               
                                GPI = ws.Cells(i, 9).Value

                            End If

                    Next i

                            ws.Cells(2, 17).Value = greatestinc
                           
                            ws.Cells(2, 16).Value = GPI

                            greatestdec = ws.Cells(2, 11).Value

                        For i = 3 To endrow

                            If (ws.Cells(i, 11).Value < greatestdec) Then

                                greatestdec = ws.Cells(i, 11).Value
                               
                                ws.Range("Q3").NumberFormat = "0.00%"
                               
                                GPD = ws.Cells(i, 9).Value

                            End If

                    Next i

                        ws.Cells(3, 17).Value = greatestdec
                       
                        ws.Cells(3, 16).Value = GPD
       
            Next ws

End Sub
