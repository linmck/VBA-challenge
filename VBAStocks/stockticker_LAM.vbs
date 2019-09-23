Sub Ticker()

Dim ws As Worksheet

For Each ws In Worksheets

  'Declare Ticker
  Dim Ticker As String
  
  'Declare Open & Close Prices
  Dim Close_Price As Double
  Dim Open_Price As Double
  
  'Declare Percent Change
  Dim Prct_Change As Double
  
  'Declare Year Change
  Dim Yr_Change As Double

  'Declare & Initially Define Total Volume As 0
  Dim Tot_Vol As Double
  Tot_Vol = 0
    
  'Find Last Row in Sheet
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  'Label Column Headers
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percent Change"
  Cells(1, 12).Value = "Total Volume"
  
  'Create Summary Table Row Variable to Keep Track of Each Ticker
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  'Loop Through All Tickers
  For i = 2 To LastRow

  'Declare Day Variable
  Dim Day As String
  Day = Right(Cells(i, 2).Value, 4)

    'Find Change in Ticker and Track Closing Values
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      'Define Ticker
      Ticker = Cells(i, 1).Value
      
      'Define Close Price
      Close_Price = Cells(i, 6).Value

      'Add Day's Volume to Total Volume
      Tot_Vol = Tot_Vol + Cells(i, 7).Value

      'Print Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      'Print the Total Volume in the Summary Table
      Range("L" & Summary_Table_Row).Value = Tot_Vol
      
      'Calculate and Print Year Change
      Yr_Change = Close_Price - Open_Price
      Range("J" & Summary_Table_Row).Value = Yr_Change
      If Yr_Change > 0 Then
      Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
      
      Else: Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      End If
      
     'Calculate and Print Percent Change
     Prct_Change = (Close_Price - Open_Price) / Open_Price
     Range("K" & Summary_Table_Row).Value = Prct_Change
     Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

      'Add Another Row to Summary Table
      Summary_Table_Row = Summary_Table_Row + 1
      
      'Set Total Volume Back to 0 to Calculate Next Ticker Correctly
      Tot_Vol = 0

    'Find All Values in Same Ticker...
    Else

      'Add to the Brand Total
      Tot_Vol = Tot_Vol + Cells(i, 7).Value
        
        If Day = "0101" Then
        Open_Price = Cells(i, 3).Value

        End If
          
    End If

  Next i

Dim Increase As Double
Dim Increase_Ticker As String
Dim Decrease As Double
Dim Decrease_Ticker As String
Dim Volume As Double
Dim Volume_Ticker As String

Increase = 0
Decrease = 0
Volume = 0

Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

For j = 2 To LastRow

If Increase < Cells(j + 1, 11).Value Then
Increase = Cells(j + 1, 11)
Increase_Ticker = Cells(j + 1, 1)
End If

If Decrease > Cells(j + 1, 11).Value Then
Decrease = Cells(j + 1, 11)
Decrease_Ticker = Cells(j + 1, 1)
End If

If Volume < Cells(j + 1, 12).Value Then
Volume = Cells(j + 1, 12)
Volume_Ticker = Cells(j + 1, 1)
End If

Next j

Cells(2, 15).Value = Increase_Ticker
Cells(2, 16).Value = Increase
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = Decrease_Ticker
Cells(3, 16).Value = Decrease
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 15).Value = Volume_Ticker
Cells(4, 16).Value = Volume

Next ws

End Sub


