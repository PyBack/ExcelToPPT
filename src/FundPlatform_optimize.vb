Private Sub Optimize_MakeMonthlyRet(ByVal fund_count As Integer)
	
	Dim start_date As Date
	Dim end_date As Date
	Dim now_date As Date
	
	
	Set sht_data = Sheets("Data")
	Set sht_datamonthopt = Sheets("optimize_new")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.Columns.Count
	
	portfolio_column_index = 1 + fund_count
	
	sht_data.Activate
	start_date = sht_data.Range("A2").Value
	end_date = sht_data.Cells(rownum, 1).Value
	Set fund_name_rng = sht_data.Range(Cells(1, 2), Cells(1, 1 + fund_count))
	
	sht_datamonthopt.Activate
	sht_datamonthopt.UsedRange.Delete
	
	sht_datamonthopt.Range("A1").Select
	sht_datamonthopt.Range("A3").Value = "Date"
	
	For i = 1 To fund_count
		fund_name = fund_name_rng.Cells(i).Value
		' fund_weight = weight_rng.Cells(i).Value
		
		sht_datamonthopt.Cells(1, i + 1).Value = 1
		sht_datamonthopt.Cells(3, i + 1).Value = fund_name
	Next i
	
	now_date = dhLastDayInMonth(start_date)
	
	i = 4
	Do While now_date <= end_date
		' Debug.Print now_date
		With sht_datamonthopt.Cells(i, "A")
			.Value = now_date
			.NumberFormat = "yyyy-mm-dd"
		End With
		
		Set vlookup_table_rng = sht_data.Range(sht_data.Cells(2, 1), sht_data.Cells(rownum, portfolio_column_index))
		For j = 1 To fund_count
		
			std_price = Application.WorksheetFunction.VLookup(sht_datamonthopt.Cells(i, "A"), _
															  vlookup_table_rng, _
															  i + j, _
															  False)
			base_price = sht_data.Cells(2, j + 1).Value
			fund_std_price = std_price / base_price
			sht_datamonthopt.Cells(i, 1+ j).Value = fund_std_price
			
		Next j
		
		now_date = DateSerial(Year(now_date) ,Month(now_date) + 2, 0)
		i = i + 1
		
	Loop
	
	
End Sub
Private Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
	'Return the first day in the specified month.
	If dtmDate = 0 Then
		' Did the caller pass in a date? If not, use
		' the current date.
		dtmDate = Date
	End If
	dhFirstDayInMonth = DateSerial(Year(dtmDate), _
								   Month(dtmDate),
								   1)
End Function
Private Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
	' Return the last day in the specified month.
	If dtmdate = 0 Then
		'Did the caller pass in a date? If not, use
		' the current date.
		dtmDate = Date
	End If
	dhLastDayInMonth = DateSerial(Year(dtmDate), _
								  Month(dtmDate) + 1, _
								  0)
End Function
Public Sub Optimize_MakePortfolioRet(ByVal fund_count As Integer)

	Set sht_data = Sheets("Data")
	Set sht_datamonthopt = Sheets("optimize_new")
	
	opt_rownum = sht_datamonthopt.UsedRange.Rows.Count
	opt_data_column_count = sht_datamonthopt.UsedRange.Columns.Count
	
	portfolio_column_index = 1 + fund_count + 1
	
	sht_datamonthopt.Activate
	
	Range("A3").Select
	Range(Selection, Selection.End(xlDown)).Select
	opt_rownum = Selection.Count
	Range("A3").Select
	
	opt_rownum = opt_rownum + 2
	sht_datamonthopt.Cells(3, portfolio_column_index).Value = "포트폴리오 밸류"
	sht_datamonthopt.Cells(3, portfolio_column_index + 1).Value = "포트폴리오 월별수익률"
	
	fund_count_end_col_letter = Col_Letter(1 + fund_count)
	month_ret_col_letter = Col_Letter(portfolio_column_index)
	month_ret_col_letter_addone = Col_Letter(portfolio_column_index + 1)
	
	For i = 1 To fund_count
		ref_col_letter = Col_Letter(i + 1)
		sht_datamonthopt.Cells(2, i + 1).Formula = "=" + ref_col_letter + "$1*5"
	Next i
	
	sht_datamonthopt.Cells(2, portfolio_column_index).Formula = "=SUM(B2:" + fund_count_end_col_letter + "2)"
	sht_datamonthopt.Cells(2, portfolio_column_index + 1).Formula = "=COUNT(B2:" + fund_count_end_col_letter + "2," & _
																	Chr(34) & ">0" & Chr(34) & ")"
																	
	sht_datamonthopt.Cells(4, portfolio_column_index).Formula _
											= "=SUMPRODUCT(B4:" & fund_count_end_col_letter & "4,B$2:" & fund_count_end_col_letter & "$2)"
	sht_datamonthopt.Cells(4, portfolio_column_index).Select
	Selection.AutoFill Destination:=Range(Cells(4, portfolio_column_index), Cells(opt_rownum, portfolio_column_index))
	
	' X5/X4 - 1
	sht_datamonthopt.Cells(5, portfolio_column_index + 1).Formula _
								= "=" + month_ret_col_letter + "5/" _
								      + month_ret_col_letter + "4-1"
	sht_datamonthopt.Cells(5, portfolio_column_index + 1).Select
	Selection.AutoFill Destination:=Range(Cells(5, portfolio_column_index + 1), Cells(opt_rownum, portfolio_column_index + 1))
	
	opt_anul_ret_rownum = opt_rownum + 2
	opt_anul_std_rownum = opt_anul_ret_rownum + 1
	opt_sharpe_rownum = opt_anul_std_rownum + 2
	
	month_ret_addone_start_cell_name = month_ret_col_letter_addone & 5
	month_ret_addone_end_cell_name = month_ret_col_letter_addone & opt_rownum
	
	month_ret_start_cell_name = month_ret_col_letter & 4
	month_ret_end_cell_name = month_ret_col_letter & opt_rownum
	
	
	
	' =(X1311*0.01)^(12/COUNT(A4:A1311))-1
	sht_datamonthopt.Cells(opt_anul_ret_rownum, portfolio_column_index + 1).Formula _
	= "=(" & month_ret_col_letter & opt_rownum & "/" & month_ret_col_letter & "2)^(12/" & Str(opt_rownum -3) & ") -1"
	' =STDEV.S(F3:F36)*SQRT(12)
	sht_datamonthopt.Cells(opt_anul_std_rownum, portfolio_column_index + 1).Formula = "STDEV.S(" & month_ret_addone_start_cell_name & ":" & month_ret_addone_end_cell_name & ")*SQRT(12)"
	sht_datamonthopt.Cells(opt_sharpe_rownum, portfolio_column_index + 1).Formula = "=(" & month_ret_col_letter_addone & opt_anul_ret_rownum & "-0.015)/" & month_ret_col_letter_addone & opt_anul_std_rownum
	
End Sub
Private Function Col_Letter(ByVal col_num As Long) As String
	
	ColLtr = Cells(1, col_num).Address(True, False)
	ColLtr = Replace(ColLtr, "$1", "")
	
	Col_Letter = ColLtr
	
End Function
Private Sub Run_Optimize_MinVariance(ByVal fund_count As Integer)

	Set sht_data = Sheets("Data")
	Set sht_datamonthopt = Sheets("optimize_new")
	
	opt_rownum = sht_datamonthopt.UsedRange.Rows.Count
	opt_data_column_count = sht_datamonthopt.UsedRange.Columns.Count
	
	portfolio_column_index = 1 + fund_count + 1
	
	sht_datamonthopt.Activate
	
	Range("A4").Select
	Range(Selection, Selection.End(xlDown)).Selection
	opt_rownum = Selection.Count
	Range("A4").Select
	
	opt_rownum = opt_rownum + 3
	
	opt_anul_ret_rownum = opt_rownum + 2
	opt_anul_std_rownum = opt_anul_ret_rownum + 1
	opt_sharpe_rownum = opt_anul_std_rownum + 2
	
	sht_datamonthopt.Cells(opt_anul_std_rownum, portfolio_column_index + 1).Select
	
	Application.Run "SolverReset"
	Application.Run "SolverOk", Cells(opt_anul_std_rownum, portfolio_column_index + 1), _
										2, _
										0, _
										Range(Cells(1, 2), Cells(1, 1 + fund_count)), _
										2, _
										"GRG Nonlinear"
										
	For i = 1 To fund_count
		Application.Run "SolverAdd", Cells(1, i + 1), 1, "10"
		Application.Run "SolverAdd", Cells(1, i + 1), 3, "0"
		Application.Run "SolverAdd", Cells(1, i + 1), 4, "정수"
	Next i
	
	Application.Run "SolverAdd", Cells(2, portfolio_column_index), 2, "100"
	'Application.Run "SolverAdd", Cells(2, portfolio_column_index + 1), 2, "3"
	
	Application.Run "SolverSolve", True
	Application.Run "SolverFinishDialog", 1, Array(1)
	
	Set opt_weight = Range(Cells(1, 2), Cells(1, 1 + fund_count))

End Sub
Private Sub Run_Optimize_MaxSharpe(ByVal fund_count As Integer)

	Set sht_data = Sheets("Data")
	Set sht_datamonthopt = Sheets("optimize_new")
	
	opt_rownum = sht_datamonthopt.UsedRange.Rows.Count
	opt_data_column_count = sht_datamonthopt.UsedRange.Columns.Count
	
	portfolio_column_index = 1 + fund_count + 1
	
	sht_datamonthopt.Activate
	
	Range("A4").Select
	Range(Selection, Selection.End(xlDown)).Selection
	opt_rownum = Selection.Count
	Range("A4").Select
	
	opt_rownum = opt_rownum + 3
	
	opt_anul_ret_rownum = opt_rownum + 2
	opt_anul_std_rownum = opt_anul_ret_rownum + 1
	opt_sharpe_rownum = opt_anul_std_rownum + 2
	
	sht_datamonthopt.Cells(opt_sharpe_rownum, portfolio_column_index + 1).Select
	
	Application.Run "SolverReset"
	Application.Run "SolverOk", Cells(opt_sharpe_rownum, portfolio_column_index + 1), _
										1, _
										0, _
										Range(Cells(1, 2), Cells(1, 1 + fund_count)), _
										2, _
										"GRG Nonlinear"
										
	For i = 1 To fund_count
		Application.Run "SolverAdd", Cells(1, i + 1), 1, "10"
		Application.Run "SolverAdd", Cells(1, i + 1), 3, "0"
		Application.Run "SolverAdd", Cells(1, i + 1), 4, "정수"
	Next i
	
	Application.Run "SolverAdd", Cells(2, portfolio_column_index), 2, "100"
	'Application.Run "SolverAdd", Cells(2, portfolio_column_index + 1), 2, "3"
	
	Application.Run "SolverSolve", True
	Application.Run "SolverFinishDialog", 1, Array(1)
	
	Set opt_weight = Range(Cells(1, 2), Cells(1, 1 + fund_count))
	
End Sub
Private Sub Run_Optimize_MaxReturn(ByVal fund_count As Integer)

	Set sht_data = Sheets("Data")
	Set sht_datamonthopt = Sheets("optimize_new")
	
	opt_rownum = sht_datamonthopt.UsedRange.Rows.Count
	opt_data_column_count = sht_datamonthopt.UsedRange.Columns.Count
	
	portfolio_column_index = 1 + fund_count + 1
	
	sht_datamonthopt.Activate
	
	Range("A4").Select
	Range(Selection, Selection.End(xlDown)).Selection
	opt_rownum = Selection.Count
	Range("A4").Select
	
	opt_rownum = opt_rownum + 3
	
	opt_anul_ret_rownum = opt_rownum + 2
	opt_anul_std_rownum = opt_anul_ret_rownum + 1
	opt_sharpe_rownum = opt_anul_std_rownum + 2
	
	sht_datamonthopt.Cells(opt_anul_ret_rownum, portfolio_column_index + 1).Select
	
	Application.Run "SolverReset"
	Application.Run "SolverOk", Cells(opt_anul_ret_rownum, portfolio_column_index + 1), _
										1, _
										0, _
										Range(Cells(1, 2), Cells(1, 1 + fund_count)), _
										2, _
										"GRG Nonlinear"
										
	For i = 1 To fund_count
		Application.Run "SolverAdd", Cells(1, i + 1), 1, "10"
		Application.Run "SolverAdd", Cells(1, i + 1), 3, "0"
		Application.Run "SolverAdd", Cells(1, i + 1), 4, "정수"
	Next i
	
	Application.Run "SolverAdd", Cells(2, portfolio_column_index), 2, "100"
	'Application.Run "SolverAdd", Cells(2, portfolio_column_index + 1), 2, "3"
	
	Application.Run "SolverSolve", True
	Application.Run "SolverFinishDialog", 1, Array(1)
	
	Set opt_weight = Range(Cells(1, 2), Cells(1, 1 + fund_count))
	
Eud Sub
Private Sub optimizer_test()

	' SolverOk (setCell, MaxMinVal, ValueOf, ByChange, Engine, EngineDesc)
	' SetCell -> Set Target Cells
	' MaxMinVal -> MaxMin, and Value options in the solver parameters dialog box.
	'
	' MaxMinVal / Specifies
	' 1			/ Maximize
	' 2			/ Minimize
	' 3			/ Match a specific value
	'
	' ValueOf : if MaxMinVal is 3, you must specify the value to which the target cell is matched.
	' ByChange: Teh cell or range of cells that will be changed so that you will obtain the desired result in the target cell.
	' Engine: The solving method that should be used to solve the problem:
	' 			1 for the Simplex LP method
	'			2 for the GRG Nonlinear method
	'			3 for the Evolutionary method
	'
	' EngineDesc: An alternate way to specify the Solving method that should be used to solve the problem as string
	
	
	Worksheets("Sheets2").Activate
	Application.Run "SolverReset"
	' Application.Run "Solveroptions", precision:=0.001
	Application.Run "SolverOk", Cells(41, 7), _
										1, _
										0, _
										Range(Cells(1, 2), Cells(1, 5)), _
										1, _
										"GRG Nonlinear"
										
										
	' SolverAdd(CellRef, Relation, FormulaText)
	' CellRef: A Reference to a cell or a range of cells that forms the left side of a constraint.
	' Relation: Required Integer. The arithmetic relationship between the left and right sides of the constraint.
	'				If you chooese 4, 5, or 6, CellRef must refer to decision variable cells, and FormulaText should not be specified.
	'
	'
	' Relation	/ Arithmetic relationship
	' 1			/ <=
	' 2			/ =
	' 3			/ >=
	' 4			/ Cells referenced by CellRef must have final values that are integers.
	' 5			/ Cells referenced by CellRef must have final values of either 0 (zero) or 1.
	' 6			/ Cells referenced by CellRef must have final values that are all different and integers.
	'
	' FormulaText: The right side of the constraint.
	
	Application.Run "SolverAdd", Range("B1"), 3, "0"
	Application.Run "SolverAdd", Range("C1"), 3, "0"
	Application.Run "SolverAdd", Range("D1"), 3, "0"
	Application.Run "SolverAdd", Range("E1"), 3, "0"
	Application.Run "SolverAdd", Range("H1"), 2, "1"
	
	Application.Run "SolverSolve", True
	
End Sub
Private Sub Col_Letter_Test()

	MsgBox Col_Letter(23)
	
End Sub
Public Sub optimize_tester()
	
	fund_count = 22
	
	Call Optimize_MakeMonthlyRet(fund_count)
	Call Optimize_MakePortfolioRet(fund_count)
	Call Run_Optimize_MinVariance(fund_count)
	' Call Run_Optimize_MaxSharpe(fund_count)
	' Call Run_Optimize_MaxReturn(fund_count)
	
End Sub