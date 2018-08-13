Public is_msg_box as boolean

Private Type Performance_Report

	cumulative_ret as double
	avg_annual_ret as double
	avg_vol as double
	
	sharpe_ratio as double
	sortino_ratio as double
	
	mdd as double
	
	best_month as double
	worst_month as double
	
	positive_month as double
	negative_month as double
	
end type
Private function isweight100Pct()
	
	Dim fund_count as integer
	dim weight_count as integer
	dim weight_sum as integer
	dim rng as range
	
	set sht_input = sheets("input")
	sht_input.activate
	sht_input.Range("B6").select
	set rng = range(Selection, Selection.end(xldown))
	
	fund_count = rng.count
	
	Set rng = Range(Range("F6"), cells(5+ fund_count, "F"))
	rng.Select
	weight_sum = Application.Sum(Selection) * 10000
	debug.print "fund_count: " & fund_count & " weight_sum: " & weight_sum
	
	' msgbox "fund_count: " & fund_count & " weight_sum: " & weight_sum
	
	if weight_sum = 10000 then
		isweight100Pct = True
	elseif weight_sum <> 1 then
		isweight100Pct = False
	end if
	
end function
Private sub forwardfilldata()

	dim sht as worksheet
	
	set sht_data = sheets("Data")
	sht_data.Activate
	
	rownum = sht_data.usedRange.rows.count
	data_column_count = sht_data.usedRange.columns.count
	
	for i = 3 to rownum
		for j = 2 to data_column_count
			'sht_data.Cells(i, j).select
			if sht_data.Cells(i, j).value = "" then
				sht_data.Cells(i, j).value = sht_data.Cells(i - 1, j).value
			end if
		next j
	next i

end sub
Private sub backfilldata()
	
	dim sht as worksheet
	
	set sht_data = sheets("data")
	sht_data.Activate
	
	rownum = sht_data.usedRange.rows.count
	data_column_count = sht_data.usedRange.columns.count
	
	for i = rownum to 2 step -1
		for j = data_column_count to 2 step -1
			if sht_data.cells(i,j).value = "" then
				sht_data.cells(i, j).select	
				sht_data.Cells(i,j).value = sht_data.Cells(i+1, j).value
			end if
		next j
	next i
	
end sub
Private sub MakestdPrice()


End sub

Private sub MakeStaticPortfolioRet(ByVal fund_count as Integer, ByVal weight_rng as Range)

	Dim sht_input as Worksheet
	Dim sht_data as Worksheet
	Dim sht_port_ret as Worksheet
	
	Set sht_input = sheets("Input")
	Set sht_data = sheets("Data")
	Set sht_port_ret = Sheets("PortfolioReturn")
	
	rownum = sht_data.usedRange.rows.count
	data_column_count = sht_data.usedRange.columns.count
	
	'clear sheet
	sht_port_ret.Activate
	sht_port_ret.Cells.Select
	Selection.ClearContents
	Range("A1").select
	
	' ' copy date
	' sht_data.Activate
	' sht_data.Range(sht_data.Cells(1, "A"), sht_data.Cells(rownum, "A")).select
	' selection.copy
	
	' sht_port_ret.Activate
	' sht_port_ret.Range("A1").select
	' ActiveSheet.Paste
	
	' ' copy fund name
	' sht_data.Activate
	' sht_data.Range(sht_data.Cells(1, "B"), sht_data.Cells(1, 1+ fund_count)).select
	' selection.copy
	
	' sht_port_ret.Activate
	' sht_port_ret.Range("A1").select
	' ActiveSheet.Paste
	
	' copy data
	sht_data.Activate
	sht_data.Cells.select
	selection.copy
	
	sht_port_ret.Activate
	sht_port_ret.Range("A1").select
	ActiveSheet.Paste
	
	sht_port_ret.Range("A1").select
	
	' make cum ret & apply weight_count
	for i = 1 to fund_count
		
		fund_name = sht_data.cells(1, i + 1).value
		base_price = sht_data.cells(2, i + 1).value
		fund_weight = weight_rng.Cells(i).value
		
		debug.print fund_name, base_price, fund_weight
		
		for each item in range(cells(2, i+1), cells(rownum, i+1))
			item.value = item.value / base_price * fund_weight
		next item
		
	next i
	
	' make portfolio return (cumulative)
	sht_port_ret.cells(1, 1 + fund_count + 1).value = "portfolio value"
	sht_port_ret.Cells(1, 1 + fund_count + 2).value = "portfolio cumulative ret"
	
	for i = 2 to rownum
	
		sht_port_ret.Range(cells(i, 2), cells(i, 1 + fund_count)).select
		portfolio_value = Application.sum(selelction)
		portfolio_cum_ret = portfolio_value - 1
		sht_port_ret.cells(i, 1 + fund_count + 1).value = portfolio_value * 100
		sht_port_ret.cells(i, 1 + fund_count + 2).value = portfolio_cum_ret
		
	next i
	
	
End sub

Private Sub MakePerformanceData(ByVal fund_count as Integer, ByVal bm_index_count as Integer)

	dim sht_input as worksheet
	dim sht_data as worksheet
	dim sht_port_ret as worksheet
	dim sht_data_bm as worksheet
	dim sht_performance_data as worksheet
	dim sht_page2 as worksheet
	
	set sht_input = sheets("input")
	set sht_data = sheets("Data")
	Set sht_port_ret = sheets("PortfolioReturn")
	set sht_data_bm = sheets("DataBM")
	set sht_performance_data = sheets("PerformanceData")
	set sht_page2 = sheets("Page2")
	
	rownum = sht_data.usedRange.rows.count
	data_column_count = sht_data.usedRange.columns.count
	
	' Clear sheet
	sht_performance_data.Activate
	sht_performance_data.Cells.Select
	Selection.ClearContents
	Range("A1").Select
	sht_page2.Activate
	
	
	' Copy date
	sht_data.Activate
	sht_data.Range(sht_data.Cells(1, "A"), sht_data.Cells(rownum, "A")).Select
	Selection.Copy
	
	sht_performance_data.Activate
	sht_performance_data.Range("A1").Select
	ActiveSheet.Paste
	
	' Copy portfolio value
	sht_port_ret.Activate
	sht_port_ret.Cells(2 , 1 + fund_count + 1).select
	Set rng = Range(Selection, Selection.End(xldown))
	rng.Select
	Selection.Copy
	
	sht_performance_data.Activate
	sht_performance_data.Range("B1").Value = "Portfolio"
	sht_performance_data.Range("B2").Select
	ActiveSheet.Paste
	
	' Copy BM
	sht_data_bm.Activate
	sht_data_bm.Range("B1").Select
	Range(Selection, Selection.End(xlToRight)).Select
	Range(Selection, Selection.End(xldown)).Select
	Selection.Copy
	
	sht_performance_data.Activate
	sht_performance_data.Range("C1").Select
	ActiveSheet.Paste
	
	' make price
	For i = 1 To bm_index_count
	
		bm_name = sht_performance_data.Cells(1, i + 2).Value
		base_price = sht_performance_data.Cells(2, i + 2).Value
		
		Debug.Print bm_name, base_price
		
		For Each Item In Range(Cells(2, i + 2), Cells(rownum, i + 2))
			Item.Value = Item.Value / base_price * 100
		Next Item
		
	Next i
	
end sub

Private Sub Make_DrawDown(ByVal bm_index_count As Integer)

	Dim sht_input As Worksheet
	Dim sht_data As Worksheet
	Dim sht_performance_data As Worksheet
	
	Set sht_input = Sheets("Input")
	Set sht_data = Sheets("Data")
	Set sht_performance_data = Sheets("PerformanceData")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.columns.Count
	
	sht_performance_data.Activate
	
	' Write Drawdown header
	For i = 1 To bm_index_count + 1
		sht_performance_data.Cells(1, i + 2 + bm_index_count).Value = _
						sht_performance_data.Cells(1, i + 1).Value + " DD"
	Next i
	
	sht_performance_data.Range("A1").Select
	
	For i = 2 To bm_index_count + 2
		For j = 2 To rownum
			max_value = Application.WorksheetFunction.Max( _
														sht_performance_data.Range( _
														sht_performance_data.Cells(2, i), _
														sht_performance_data.Cells(j, i) _
														))
														
			std_value = sht_performance_data.Cells(j, i).Value + 1
			sht_performance_data.Cells(j, bm_index_count + 1 + i).Value = (std_value - (1 + max_value)) / (1 + max_value)
			
		Next j
	Next i
	
End Sub
Private Sub Make_CumulativeRet(ByVal bm_index_count As Integer)

	Dim sht_input As Worksheet
	Dim sht_data As Worksheet
	Dim sht_performance_data As Worksheet
	
	Set sht_input = Sheets("Input")
	Set sht_data = Sheets("Data")
	Set sht_performance_data = Sheets("PerformanceData")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.columns.Count
	
	sht_performance_data.Activate
	
	' Write Cumulative Ret header
	For i = 1 To bm_index_count + 1
		sht_performance_data.Cells(1, i + 2 + bm_index_count).Value = _
						sht_performance_data.Cells(1, i + 1).Value + " Cumulative Ret"
	Next i
	
	sht_performance_data.Range("A1").Select
	
	For i = 1 To bm_index_count + 1
		For j = 2 To rownum
			
			cumulative_ret = sht_performance_data.Cells(j, 1 + i).Value * 0.01 - 1
			sht_performance_data.Cells(j, bm_index_count * 2 + 3 + i).Value = cumulative_ret
			
		Next j
	Next i

End Sub
Private Sub Make_MonthlyReturn(ByVal bm_index_count As Integer)

	Dim start_date As Date
	Dim end_date As Date
	Dim now_date As Date
	
	set sht_input = Sheets("Input")
	Set sht_data = Sheets("Data")
	Set sht_bm_data = Sheets("DataBM")
	Set sht_performance_data = Sheets("PerformanceData")
	Set sht_performance_month_data = Sheets("PerformanceMonthData")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.columns.Count
	
	' Clear sheet
	sht_performance_month_data.Activate
	sht_performance_month_data.Cells.Select
	Selection.ClearContents
	Range("A1").Select
	
	sht_data.Activate
	start_date = sht_data.Range("A2").Value
	end_date = sht_data.Cells(rownum, 1).Value
	
	sht_performance_month_data.Activate
	sht_performance_month_data.Range("A1").Value = "Date"
	
	sht_performance_month_data.Range("B1").Value = "Portfolio Cumulative Ret"
	For i = 1 To bm_index_count
		bm_name = sht_bm_data.Cells(1, 1 + i).Value
		sht_performance_month_data.Cells(1, 2 + i).Value = bm_name & " Cumulative Ret"
	Next i
	
		

End Sub