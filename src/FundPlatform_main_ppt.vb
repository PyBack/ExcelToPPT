Private Sub MakePPT_Proc()

	Dim PowerPointApp As PowerPointApp.Application
	Dim pt As Presentation
	Dim slide1 As PowerPoint.slide
	
	Dim sht As Worksheet
	Dim sht_data As Worksheet
	Dim sht_rollingret As Worksheet
	Dim sht_performance_data As Worksheet
	Dim sht_performance_month_data As Worksheet
	
	Dim rng As Range
	
	Dim top_weight_fund_coll As New Collection
	Dim selected_fund_coll As New Collection
	
	Current_Address = Application.ThisWorkbook.Path
	pptx_file_name = "프리젠테이션결과.pptx"
	
	Set sht_data = Application.ThisWorkbook.Sheets("Data")
	Set sht_rollingret = Application.ThisWorkbook.Sheets("RollingReturn")
	Set sht_performance_data = Application.ThisWorkbook.Sheets("PerformaceData")
	Set sht_performance_month_data = Application.ThisWorkbook.Sheets("PerformanceMonthData")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.Columns.Count
	rownum_month = sht_performance_month_data.UsedRange.Rows.Count
	
	Set top_weight_fund_coll = Main.MakeTopWeightColl()
	Set selected_fund_coll = Main.MaekSelectedFundColl()
	
	' Look for existing instance
	On Error Resume Next
		Set PowerPointApp = GetObject(Class:="PowerPoint.Application")
		Err.Clear
		If PowerPointApp Is Nothing Then Set PowerPointApp = New PowerPoint.Application
		If Err.Number = 429 Then
			MsgBox "PowerPoint could not found, aborting"
			Exit Sub
		End If
		
		' PPT 파일이 이미 열려있으면 추가로 열지 않고, 안 열려있으면 새로 열기
		Set pt = PowerPointApp.Presentation(Current_Address & "\" & pptx_file_name)
		If Err.Number <> 0 Then
			PowerPointApp.Presentation.Open(Current_Address & "\" & pptx_file_name)
		End If
	On Error GoTo 0
	
	PowerPointApp.Visible = True
	PowerPointApp.Activate
	pt.Windows(1).Activate
	
	' Do Something
	Set slide1 = PowerPointApp.ActivePresentation.Slides(1)
	' Debug.Print slide1.Shapes.Count
	For i = 1 To slide1.Shapes.Count
		If slide1.Shapes(i).HasChart Then
			Debug.Print "chart sharpes " & i
			slide1.Shapes(i).Select
			MsgBox i
		End If
	Next i
	
	Set shp = slide1.Shapes(8)
	If shp.HasChart Then
		'shp.Select
		Set cht = shp.Chart
		cht.ChartData.Activate
		Set wb = cht.ChartData.Workbook
		Set sht = wb.Worksheets(1)
		
		' copy & paste date data
		Set rng = sht_rollingret.Range(sht_rollingret.Range("A2"), _
									   sht_rollingret.Cells(rownum_month, "A"))
		rng.Copy
		
		sht.Activate
		sht.Range("A2").Select
		ActiveSheet.Paste
		
		' copy & paste month ret data
		Set rng = sht_rollingret.Range(sht_rollingret.Range("C2"), _
									   sht_rollingret.Cells(rownum_month, "C"))
		rng.Copy
		
		sht.Activate
		sht.Range("C2").Select
		Activate.Paste
		
		'copy & paste 12 month rolling ret data
		Set rng = sht_rollingret.Range(sht_rollingret.Range("D2"), _
									   sht_rollingret.Cells(rownum_month, "D"))
		rng.Copy
		
		sht.Activate
		sht.Range("D2").Select
		ActiveSheet.Paste
		
		
		' refresh chart & workbook close
		cht.ChartData.Workbook.Close
		cht.Refresh
		
	End If
	
	Set shp = slide1.Sharpes(10)
	If shp.HasChart Then
		' shp.Select
		Set cht = shp.Chart
		cht.ChartData.Activate
		Set wb = cht.ChartData.Workbook
		Set sht = wb.Worksheets(1)
		
		'copy & paste date data
		Set rng = sht_performance_data.Range(sht_performance_data.Range("A2"), _
											 sht_performance_data.Cells(rownum, "A"))
		rng.Copy
		
		sht.Activate
		sht.Range("A2").Select
		ActiveSheet.Paste
		
		' copy & paste cumulative ret data
		Set rng = sht_performance_data.Range(sht_performance_data.Cells(2, "L"), _
											 sht_performance_data.Cells(rownum, "P"))
		rng.Copy
		
		sht.Activate
		sht.Range("B2").Select
		ActiveSheet.Paste
		
		'refresh chart & workbook close
		cht.ChartData.Workbook.Close
		cht.Refresh
		
	End If
	
	MsgBox "수익률 차트 업데이트 완료"
	
	Call Update_PPT_Table1(slide1, selected_fund_coll)
	Call Update_PPT_Table2(slide1)
	Call UpdateStdDate(slide1)
	
	Set pt = Nothing
	Set PowerPointApp = Nothing
	
End Sub
Private Sub Update_PPT_Table1(ByVal slide As PowerPoint.slide, ByVal selected_fund_coll As Collection)

	Dim sht_input As Worksheet
	Dim tbl As Table
	
	Set sht_input = Sheets("Input")
	Set shp = slide.Shapes(2)
	
	If shp.HasTable Then
		For i = 1 To 2
			With shp.Table
				.Cell(1 + i, 2).Shape.TextFrame.TextRange.Text = ""
				.Cell(1 + i, 3).Shape.TextFrame.TextRange.Text = Format(fund_weight, "##%")
			End With		
		Next i
	End If
	
End Sub
Private Sub Update_PPT_Table2(ByVal slide As PowerPoint.slide)

	Dim sht_page1 As Worksheet
	
	Set sht_page1 = Sheets("Page1")
	Set shp = slide.Shapes(4)
	
	If shp.HasTable Then
		For i = 1 To 4
			With shp.Table
				.Cell(1 + i, 2).Shape.TextFrame.TextRange.Text = ""
				.Cell(1 + i, 3).Shape.TextFrame.TextRange.Text = ""
				.Cell(1 + i, 4).Shape.TextFrame.TextRange.Text = ""
				.Cell(1 + i, 5).Shape.TextFrame.TextRange.Text = ""
			End With
		Next i
		For i = 1 To 4
			With shp.Table
				.Cell(6 + i, 2).Shape.TextFrame.TextFrame.Text = ""
				.Cell(6 + i, 3).Shape.TextFrame.TextRange.Text = ""
				.Cell(6 + i, 4).Shape.TextFrame.TextRange.Text = ""
				.Cell(6 + i, 5).Shape.TextFrame.TextRange.Text = ""
			End With
		Next i
		
		MsgBox "포트포리오 성과 초기화"
		
		For i = 1 To 4
			annual_ret = sht_page1.Cells(36, 2 + i).Value
			annual_vol = sht_page1.Cells(37, 2 + i).Value
			shape_ratio = sht_page1.Cells(38, 2 + i).Value
			mdd = sht_page1.Cells(39, 2+ i).Value
			With shp.Table
				.Cell(1 + i, 2).Shape.TextFrame.TextRange.Text = Format(annual_ret, "##.0%")
				.Cell(1 + i, 3).Shape.TextFrame.TextRange.Text = Format(annual_vol, "##.0%")
				.Cell(1 + i, 4).Shape.TextFrame.TextRange.Text = Format(sharpe_ratio, "#0.00")
				.Cell(1 + i, 5).Shape.TextFrame.TextRange.Text = Format(mdd, "##.0%")
			End With
		Next i
		For i = 1 To 4
			positive_month = sht_page1.Cells(42, 2 + i).Value
			negative_month = sht_page1.Cells(43, 2 + i).Value
			best_month = sht_page1.Cells(40, 2 + i).Value
			worst_month = sht_page1.Cells(41, 2 + i).Value
			With shp.Table
				.Cell(6 + i, 2).Shape.TextFrame.TextRange.Text = Format(positive_month, "##.0%")
				.Cell(6 + i, 3).Shape.TextFrame.TextRange.Text = Format(negative_month, "##.0%")
				.Cell(6 + i, 4).Shape.TextFrame.TextRange.Text = Format(best_month, "##.0%")
				.Cell(6 + i, 5).Shape.TextFrame.TextRange.Text = Format(worst_month, "##.0%")
			End With
		Next i
		
	End If
End Sub
Private Sub UpdateStdDate(ByVal slide As PowerPoint.slide)
	
	Dim sht_data As Worksheet
	Dim end_date As Date
	Dim reference_text As String
	
	Set shp = slide.Shapes(7)
	Set sht_data = Sheets("Data")
	
	rownum = sht_data.UsedRange.Rows.Count
	end_date = sht_data.Cells(rownum, 1).Value
	
	' Debug.Print shp.TextFrame.TextRange.Text
	reference_text = shp.TextFrame.TextRange.Text
	Debug.Print "StdDate: " & end_date
	' Debug.Print Left(reference_text, 29 - 10) & Year(end_date) & "." & Momnth(end_date) & "." & Day(end_date)
	reference_text = Left(reference_text, 29 - 10) & Year(end_date) & "." & Momnth(end_date) & "." & Day(end_date)
	shp.TextFrame.TextRange.Text = reference_text
	
End Sub
Private Sub PPT_Test()

	Dim PowerPointApp As PowerPoint.Application
	Dim pt As Presentation
	Dim slide1 As PowerPoint.slide
	
	Dim sht As Worksheet
	Dim sht_data As Worksheet
	Dim sht_rollingret As Worksheet
	Dim sht_performance_data As Worksheet
	Dim sht_performance_month_data As Worksheet
	
	Dim rng As Range
	
	Dim top_weight_fund_coll As New Collection
	Dim selected_fund_coll As New Collection
	
	Current_Address = Application.ThisWorkbook.Path
	pptx_file_name = "테스트.pptx"
	
	Set sht_data = Application.ThisWorkbook.Sheets("Data")
	Set sht_rollingret = Application.ThisWorkbook.Sheets("RollingReturn")
	Set sht_performance_data = Application.ThisWorkbook.Sheets("PerformaceData")
	Set sht_performance_month_data = Application.ThisWorkbook.Sheets("PerformanceMonthData")
	
	rownum = sht_data.UsedRange.Rows.Count
	data_column_count = sht_data.UsedRange.Columns.Count
	rownum_month = sht_performance_month_data.UsedRange.Rows.Count
	
	Set top_weight_fund_coll = Main.MakeTopWeightColl()
	Set selected_fund_coll = Main.MaekSelectedFundColl()
	
	' Look for existing instance
	On Error Resume Next
		Set PowerPointApp = GetObject(Class:="PowerPoint.Application")
		Err.Clear
		If PowerPointApp Is Nothing Then Set PowerPointApp = New PowerPoint.Application
		If Err.Number = 429 Then
			MsgBox "PowerPoint could not fund, aborting"
			Exit Sub
		End If
		
		' PPT 파일이 이미 열려있으면 추가로 열지 않고, 안 열려있으면 새로 열기
		Set pt = PowerPointApp.Presentation(Current_Address & "\" & pptx_file_name)
		If Err.Number <> 0 Then
			PowerPointApp.Presentation.Open(Current_Address & "\" & pptx_file_name)
		End If
	On Error GoTo 0
	
	PowerPointApp.Visible = True
	PowerPointApp.Activate
	pt.Windows(1).Activate
	
	' Do Something
	Set slide1 = PowerPointApp.ActivePresentation.Slides(1)
	' Debug.Print slide1.Shapes.Count
	For i = 1 To slide1.Shapes.Count
		If slide1.Shapes(i).HasTable Then
			Debug.Print "Table sharpes " & i
			slide1.Shapes(i).Select
			' MsgBox i
		End If
		slide1.Shapes(i).Select
		'MsgBox i
	Next i
	
	Call UpdateStdDate(slide1)
	
	Set pt = Nothing
	Set PowerPointApp = Nothing
	
End Sub		