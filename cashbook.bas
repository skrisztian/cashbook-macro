REM  *****  BASIC  *****
REM
REM This macro prepares a Cash Book report based on a properly
REM filled and formatted table.
REM
REM Writen by Krisztian Stancz
REM Version: 2017-Mar-14-v1

Type CashBookType
    InvoiceDate as Date
	InvoiceStringLetters as String
    InvoiceStringNumbers as String
    InvoiceStringDigits as Double
    InvoiceNumTo as Double
    InvoiceCount as Long
    LineIncome as Long
    LineExpense as Long
    Comment as String
End Type

Sub Main

	Dim Sheet As Object
	Dim ColumnLastCell, Row, TotalIncome, TotalExpense, TransferRow As Long
	Dim I, J, K, L As Long

	'Get active sheet
	Sheet = ThisComponent.GetCurrentController.ActiveSheet

	'Get the last cell's position in column "A"
	LastRow = GetLastCellOfFirstColumn(Sheet)
	
	'Get the last column
	LastColumn = GetLastColumn(Sheet)

	'Fill in the column after the last one with a header and functions
	'The functions result +1 for income and -1 for expense rows
	LastColumn = LastColumn+1
	Sheet.GetCellByPosition(LastColumn, 0).String = "Tipus"
	For J = 1 To LastRow
		Sheet.GetCellByPosition(LastColumn, J).Formula = "=IF($C" & J+1 &">=$D" & J+1 & ";1;-1)"
	Next J

	'Select range, we skip the headers
	'From Col, From Row To Col, To Row
	Dim SortRange As Object
	SortRange = Sheet.getCellRangeByPosition(0, 1, LastColumn, LastRow)
	
	'Set up the sort parameters
	' Dátum - column 0 - ascending
	' Tipus - lastColumn - descending
	' Számla sorszám - column 1 - ascending
	Dim SortFields(2) As New com.sun.star.util.SortField
	SortFields(0).Field = 0
	SortFields(0).SortAscending = TRUE
	SortFields(1).Field = LastColumn
	SortFields(1).SortAscending = FALSE
	SortFields(2).Field = 1
	SortFields(2).SortAscending = TRUE
	
	'Set up sort descriptor
	Dim SortDesc(0) As New com.sun.star.beans.PropertyValue
	SortDesc(0).Name = "SortFields"
  	SortDesc(0).Value = SortFields()
  	
  	'Sort the range.
 	SortRange.Sort(SortDesc())

	'Delete the "Tipus column"
	Dim ColumnToDelete As Object
	ColumnToDelete = Sheet.getCellRangeByPosition(LastColumn, 0, LastColumn, 0).Columns
	ColumnToDelete.removeByIndex(0, 1)
	LastColumn = LastColumn-1
	
	'Feed the table values into a CashBook data type
	'The initial dimension is likely to be gigger then the number of actual records,
	'but it's OK, as we'll store the last record's position, thus no need to Redim either
	Dim Book(LastRow-1) As New CashBookType

	'For each row in table
	For I = 0 To LastRow-1
		
		Dim InvoiceDate As Date
		Dim InvoiceString, Comment, InvoiceStringLetters, InvoiceStringNumbers As String
		Dim Income, Expense, Counter As Long
		
		InvoiceString = ""
		InvoiceStringNumbers = ""
		InvoiceStringLetters = ""
		InvoiceStringDigits = 0
		Income = 0
		Expense = 0
		Comment = ""
		
		'Read in table data row
		InvoiceDate = Sheet.getCellByPosition(0,I+1).Value
		InvoiceString = Sheet.getCellByPosition(1,I+1).String
		Income = Sheet.getCellByPosition(2,I+1).Value
		Expense = Sheet.getCellByPosition(3,I+1).Value
		Comment = Sheet.getCellByPosition(5,I+1).String

		'Prepare InvoiceString for manipulation
		'by removing trailing, internal and ending spaces
		InvoiceString = Trim(InvoiceString)
		InvoiceString = join(split(InvoiceString, " "), "")

		'Search for "/" in InvoiceString
		'if we do not find it, we split it into letters and numbers
		If InStr(InvoiceString, "/") = 0 Then
		
			StringLength = len(InvoiceString)			
			Counter = 1
	
			Do While IsNumeric(Right(InvoiceString, Counter)) And Counter <= StringLength
				Counter = Counter + 1
			Loop

			If Counter = 1 Then
				'InvoiceString consists only numbers
				InvoiceStringNumbers = InvoiceString
			ElseIf Counter > StringLength Then
				'InvoiceString ends with letters
				InvoiceStringNumbers = InvoiceString
			Else
				'Invoice string starts with letters and ends with numbers
				InvoiceStringLetters = Left(InvoiceString, StringLength-(Counter-1))
				InvoiceStringNumbers = Right(InvoiceString, Counter-1)
				InvoiceStringDigits = CDbl(InvoiceStringNumbers)
			End If
			
   		Else
			InvoiceStringNumbers = InvoiceString
		End If

		'After we split the inoice numbers into letters + numbers parts, we group those invoices together
		'where date and inoice letters part are the same and the numbers are increasing one-by-one. We feed 
		'this into an array.
		
		'We must record the first data set
		If I = 0 Then 
			K = 0
		End If
		
		If Book(K).InvoiceDate = InvoiceDate _
		And Book(K).InvoiceStringLetters = InvoiceStringLetters _
		And Book(K).InvoiceNumTo = InvoiceStringDigits-1 Then

			'Add to the last existing record
    		Book(K).InvoiceNumTo = InvoiceStringNumbers
	    	Book(K).InvoiceCount = Book(K).InvoiceCount+1
	    	Book(K).LineIncome = Book(K).LineIncome + Income
	    	Book(K).LineExpense = Book(K).LineExpense + Expense
		
		Else
			'Create a new record
			If I > 0 Then
				K = K+1
			End If
			Book(K).InvoiceDate = InvoiceDate
			If Len(InvoiceStringLetters) > 0 Then
				Book(K).InvoiceStringLetters = InvoiceStringLetters
			End If
			Book(K).InvoiceStringNumbers = InvoiceStringNumbers
    		Book(K).InvoiceNumTo = InvoiceStringDigits
	    	Book(K).InvoiceCount = 1
	    	Book(K).LineIncome = Income
	    	Book(K).LineExpense = Expense
	    	Book(K).Comment = Comment
		End If
		
	Next I
	
	'Create Cash Book headers
	Sheet.GetCellByPosition(LastColumn+2, 0).String = "Időszaki pénztárjelentés"
	Sheet.GetCellByPosition(LastColumn+2, 0).CharWeight = com.sun.star.awt.FontWeight.BOLD
	Sheet.GetCellByPosition(LastColumn+2, 2).String = "Sorszám"
	Sheet.GetCellByPosition(LastColumn+3, 2).String = "Be/Kif. napja"
	Sheet.GetCellByPosition(LastColumn+4, 2).String = "Bevételi/Kiadási bizonylat sz."
	Sheet.GetCellByPosition(LastColumn+5, 2).String = "Kp. forg. jogc. sz."
	Sheet.GetCellByPosition(LastColumn+6, 2).String = "Szöveg 1"	
	Sheet.GetCellByPosition(LastColumn+7, 2).String = "Szöveg 2"		
	Sheet.GetCellByPosition(LastColumn+8, 2).String = "Bevétel Ft."	
	Sheet.GetCellByPosition(LastColumn+9, 2).String = "Kiadás Ft."	
	
	Sheet.GetCellRangeByPosition(LastColumn+2, 2, LastColumn+9, 2).CharWeight = com.sun.star.awt.FontWeight.BOLD
	Sheet.GetCellRangeByPosition(LastColumn+2, 2, LastColumn+9, 2).IsTextWrapped = TRUE
	
	Dim ColumnNameTitle, ColumnNameIncome, ColumnNameExpense As String 

	ColumnNameTitle = GetColumnName(Sheet.GetCellByPosition(LastColumn+2, 0))
	ColumnNameIncome = GetColumnName(Sheet.GetCellByPosition(LastColumn+8, 0))
	ColumnNameExpense = GetColumnName(Sheet.GetCellByPosition(LastColumn+9, 0))

	'Print the data structure into the spreadsheet	
	'K is the last record's number in the Book array
	'M is the offset of Sorszám to L. Note, it changes by page
	Row = 3
	TotalIncome = 0
	TotalExpense = 0
	SumStart = 4
	TransferRow = 4	'GetByName value

	For L = 0 To K

		Dim Page As Long

		'Sorszám
		Sheet.GetCellByPosition(LastColumn+2, Row).Value = L+1

		'Be/Kif. napja
		Sheet.GetCellByPosition(LastColumn+3, Row).String = Format(Book(L).InvoiceDate, "mm.dd.")
		Sheet.GetCellByPosition(LastColumn+3, Row).HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT

		'B./K. bizonylatsz.
		If Book(L).InvoiceCount > 1 Then
			Sheet.GetCellByPosition(LastColumn+4, Row).String = Book(L).InvoiceStringNumbers _
			& " - " & Book(L).InvoiceNumTo															
		Else
			Sheet.GetCellByPosition(LastColumn+4, Row).String = Book(L).InvoiceStringNumbers
		End If

		'Szöveg 1
		Sheet.GetCellByPosition(LastColumn+6, Row).String = Book(L).InvoiceStringLetters	
		
		'Szöveg 2
		Sheet.GetCellByPosition(LastColumn+7, Row).String = Book(L).Comment

		'Bevétel
		If Book(L).LineIncome > 0 Then
			Sheet.GetCellByPosition(LastColumn+8, Row).Value = Book(L).LineIncome
			TotalIncome = TotalIncome + Book(L).LineIncome
		End If
		Sheet.GetCellByPosition(LastColumn+8, Row).NumberFormat = 3

		'Kiadás
		If Book(L).LineExpense > 0 Then
			Sheet.GetCellByPosition(LastColumn+9, Row).Value = Book(L).LineExpense
			TotalExpense = TotalExpense + Book(L).LineExpense
		End If
		Sheet.GetCellByPosition(LastColumn+9, Row).NumberFormat = 3

		'Sometimes we'll need to insert transfer lines (Átvitel/Áthozat)
		' Do this for every:
		' 26, 66, 106, 146, etc. line
		' 40, 80, 120, 160, etc. line if that is not the last line
		' 0<=line<26, 40<line<66, 80<line<106, etc. lines if that is the last line
		' Lines are counted in the GetCellByPosition value

		Dim Min, Max As Long

		Min = (L\40)*40
		Max = Min+26

		If ((L-26) MOD 40) = 0 _
		Or (L <> K And (L MOD 40) = 0 And L<> 0) _
		Or (L = K And L > Min And L < Max) _
		Or (L = K And L = 0) Then

			'Add extra Átvitel + Áthozat rows
			Sheet.GetCellByPosition(LastColumn+2, Row+1).String = "Átvitel"
			Sheet.GetCellByPosition(LastColumn+2, Row+1).CharWeight = com.sun.star.awt.FontWeight.BOLD
			Sheet.GetCellByPosition(LastColumn+2, Row+3).String = "Áthozat"
			Sheet.GetCellByPosition(LastColumn+2, Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD

			'Bevétel átvitel érték
			Sheet.GetCellByPosition(LastColumn+8, Row+1).Formula = "=SUM(" & ColumnNameIncome & TransferRow & ":" & ColumnNameIncome & Row+1 & ")"
			Sheet.GetCellByPosition(LastColumn+8, Row+1).NumberFormat = 3
			Sheet.GetCellByPosition(LastColumn+8, Row+1).CharWeight = com.sun.star.awt.FontWeight.BOLD
			
			'Bevétel áthozat érték
			Sheet.GetCellByPosition(LastColumn+8, Row+3).Formula = "=" & ColumnNameIncome & Row+2
			Sheet.GetCellByPosition(LastColumn+8, Row+3).NumberFormat = 3
			Sheet.GetCellByPosition(LastColumn+8, Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD
			
			'Kiadás átvitel érték
			ColumnName = GetColumnName(Sheet.GetCellByPosition(LastColumn+9, 0))
			Sheet.GetCellByPosition(LastColumn+9, Row+1).Formula = "=SUM(" & ColumnNameExpense & TransferRow & ":" & ColumnNameExpense & Row+1 & ")"
			Sheet.GetCellByPosition(LastColumn+9, Row+1).NumberFormat = 3
			Sheet.GetCellByPosition(LastColumn+9, Row+1).CharWeight = com.sun.star.awt.FontWeight.BOLD
			
			'Kiadás áthozat érték
			Sheet.GetCellByPosition(LastColumn+9, Row+3).Formula = "=" & ColumnNameExpense & Row+2
			Sheet.GetCellByPosition(LastColumn+9, Row+3).NumberFormat = 3
			Sheet.GetCellByPosition(LastColumn+9, Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD
			
			TransferRow = Row+3+1	'GetCellByName value
			Row = Row + 3			'GetCellByPosition value
			
		End If	
	
	Row = Row + 1
			
	Next L
	
	'Add total row titles
	Sheet.GetCellByPosition(LastColumn+2, Row).String = "Forgalom"
	Sheet.GetCellByPosition(LastColumn+2, Row+1).String = "Kezdő pénzkészlet"
	Sheet.GetCellByPosition(LastColumn+2, Row+2).String = "Záró pénzkészlet"	
	Sheet.GetCellByPosition(LastColumn+2, Row+3).String = "Összesen"	
	Sheet.GetCellRangeByPosition(LastColumn+2, Row, LastColumn+2,Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD

	'Forgalom bevétel
	Sheet.GetCellByPosition(LastColumn+8, Row).Formula = "=SUM(" & ColumnNameIncome & TransferRow & ":" & ColumnNameIncome & Row & ")"
	Sheet.GetCellByPosition(LastColumn+8, Row).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+8, Row).CharWeight = com.sun.star.awt.FontWeight.BOLD

	'Forgalom kiadás
	Sheet.GetCellByPosition(LastColumn+9, Row).Formula = "=SUM(" & ColumnNameExpense & TransferRow & ":" & ColumnNameExpense & Row & ")"
	Sheet.GetCellByPosition(LastColumn+9, Row).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+9, Row).CharWeight = com.sun.star.awt.FontWeight.BOLD
	
	'Kezdő pénz
	Sheet.GetCellByPosition(LastColumn+8, Row+1).Formula = "=E1"
	Sheet.GetCellByPosition(LastColumn+8, Row+1).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+9, Row+1).CellBackColor = RGB(210, 210, 210)
	
	'Záró pénz
	Sheet.GetCellByPosition(LastColumn+9, Row+2).Formula = "=E" & LastRow+1
	Sheet.GetCellByPosition(LastColumn+9, Row+2).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+8, Row+2).CellBackColor = RGB(210, 210, 210)
	
	'Összesen bevétel
	Sheet.GetCellByPosition(LastColumn+8, Row+3).Formula = "=SUM(" & ColumnNameIncome & Row+1 & ":" & ColumnNameIncome & Row+3 & ")"
	Sheet.GetCellByPosition(LastColumn+8, Row+3).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+8, Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD

	'Összesen kiadás
	Sheet.GetCellByPosition(LastColumn+9, Row+3).Formula = "=SUM(" & ColumnNameExpense & Row+1 & ":" & ColumnNameExpense & Row+3 & ")"
	Sheet.GetCellByPosition(LastColumn+9, Row+3).NumberFormat = 3
	Sheet.GetCellByPosition(LastColumn+9, Row+3).CharWeight = com.sun.star.awt.FontWeight.BOLD
	
	'Ellenőrző cella
	Sheet.GetCellByPosition(LastColumn+10, Row+3).Formula = "=IF(" & ColumnNameIncome & Row+4 & "=" & ColumnNameExpense & Row+4 & ";""Rendben"";""Hiba"")"
	
End Sub

Function GetLastCellOfFirstColumn(Sheet As Object) As Long
'Returns the last cell of the first ("A") column of the active sheet
'as a zero based position, i.e GetNameByPosition value

	Dim CellRange As Object
	Dim CellCount As Long

	'Select the first column as a range	
	'CellRange = Sheet.getCellRangeByName("A:A")
	CellRange = Sheet.getCellRangeByPosition(0, 0, 0, 10000)

	'Count the number of cells in the range
	CellCount = CellRange.computeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
	
	'Cell position starts with 0 so we subtract one from the cell count
	GetLastCellOfFirstColumn = CellCount - 1
	
End Function

Function GetLastColumn(Sheet As Object) As Long
'Returns the last used column number as a zero based position, i.e GetNameByPosition value

	Dim Cell, Cursor As Object
	
	Cell = Sheet.GetCellByPosition( 0, 0 )
	Cursor = Sheet.createCursorByRange(Cell)
	Cursor.GotoEndOfUsedArea(FALSE)
	GetLastColumn = Cursor.RangeAddress.EndColumn

End Function

Function GetColumnName(Cell As Object) As String
'Returns a cell's column name's GetCellByName value if given
'GetCellByPosition cell coordinate

	'Absolute cell name e.g.: $Sheet1.$A$1
	aCellName = Split(Cell.AbsoluteName, "$")
	GetColumnName = aCellName(UBound(aCellName)-1)

End Function

