Sub VBAChallenge():

' Define Variables

	Dim Ticker As String
	Dim Volume As Double
	
	Volume = 0

	Dim YearlyChange As Double
	Dim PercentChange As Double

 ' Define table row

	Dim tablerow As Integer


' Set Headers

	Cells(1, 9).Value = "Ticker"
	Cells(1, 10).Value = "Yearly Change"
	Cells(1, 11).Value = "Percent Change"
	Cells(1, 12).Value = "Total Stock Volume"
	

	tablerow = 2

' define last row

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

		
	For i = 2 To Lastrow


		If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

	
		' Find/Define Values


		Ticker = Cells(i, 1).Value

		Volume = Volume + Cells(i, 7).Value

		YearOpen = Cells(i, 3).Value

		YearClose = Cells(i, 6).Value

		YearlyChange = (YearClose - YearOpen)

				If YearOpen <> 0 Then
                    PercentChange = (YearlyChange / YearOpen) * 100

				End If



			If (YearlyChange > 0) Then

				Cells(tablerow, 10).Interior.ColorIndex = 4

			ElseIf (YearlyChange <= 0) Then

				Cells(tablerow, 10).Interior.ColorIndex = 3

			End If
			

		' Insert Values

		Cells(tablerow, 9).Value = Ticker
		Cells(tablerow,10).Value = YearlyChange
		Cells(tablerow,11).Value = PercentChange
		Cells(tablerow,12).Value = Volume

		tablerow = tablerow + 1

		YearlyChange = 0

		YearClose = 0

		YearOpen = Cells(i + 1, 3).Value

		Volume = 0


		Else

			Volume = Volume + Cells(i, 7).Value

		End If

	 Next i


End Sub