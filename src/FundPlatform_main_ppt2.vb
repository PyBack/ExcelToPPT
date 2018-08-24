Global pptx_file_name as string
Global cm_to_px as double

Private Sub MakePPT_Page1(ByVal sld_num as Integer)

	Dim PowerPointApp As PowerPoint.Application
	Dim pt As Presentation
	Dim sld As PowerPoint.slide
	Dim shp As PowerPoint.Shape
	
	Dim sht_page1 As Worksheet
	
	Current_Address = Application.ThisWorkbook.PATH_INFO
	
	' Look for existing instance
	On Error Resume Next
		Set PowerPointApp = GetObject(Class="PowerPoint.Application")
		Err.Clear
		If PowerPoint Is Nothing Then Set PowerPointApp = New PowerPoint.Application
		If Err.Number = 429 Then
			MsgBox "PowerPoint could not found, aborting"
			Exit Sub
		End If
		
		' PPT 파일이 이미 열려있으면 추가로 열지 않고, 안열려있으면 새로 열기
		Set pt = PowerPointApp.Presentations(Current_Address & "\" & pptx_file_name)
		If Err.Number <> 0 Then
			PowerPointApp.Presentation.Open(Current_Address & "\" & pptx_file_name)
		End If
		
		PowerPointApp.Activate
		PowerPoint.Windows(pptx_file_name).Activate
		pt.Windows(1).Activate
		PowerPointApp.Visible = True
		
	On Error GoTo 0
	
	' Do something
	Set sht_page1 = ThisWorkbook.Sheets("Page1")
	Set sld = PowerPointApp.ActivatePresentation.Slides(sld_num)
	
	sld.Select
	
	' PPT 파일에서 기존에 있던 sharpes 삭제
	Debug.Print "ppt_page1 shapes count: " & sld.Shapes.Count
	If sld.Shapes.Count = 9 Then
		For i = 1 To 2
			sld.Shapes(sld.Shapes.Count).Delete
		Next i
	End If
	
	' copy scatter chart from excel page1
	sht_page1.Activate
	ActiveSheet.ChartObjects(1).Activate
	ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
	sld.Shape.Paste
	sld.Shapes(sld.Shapes.Count).LockAspectRatio = msoFalse
	sld.Shapes(sld.Shapes.Count).Height = 5.39 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Width = 24.15 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Left = 1.53 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Top = 5.33 * cm_to_px
	
	' copy table 1 from excel page1
	sht_page1.Range(sht_page1.Range("C28"), sht_page1.Range("AL41")).CopyPicture Appearance:=xlScreen, Format:=xlPicture
	sld.Shapes.Paste
	sld.Shapes(sld.Shapes.Count).LockAspectRatio = msoFalse
	sld.Shapes(sld.Shapes.Count).Height = 5.4 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Width = 24.15 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Left = 1.53 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Top = 12.05 * cm_to_px

End Sub
Private Sub MakePPt_Page2(ByVal sld_num As Integer)

	Dim PowerPointApp As PowerPoint.Application
	Dim pt As Presentation
	Dim sld as PowerPoint.slide
	Dim shp as PowerPoint.Shape
	
	Dim sht_page2 as Worksheet
	
	Current_Address = Application.ThisWorkbook.PATH_INFO
	
	' Look for existing instance
	On Error Resume Next
		Set PowerPointApp = GetObject(Class="PowerPoint.Application")
		Err.Clear
		If PowerPoint Is Nothing Then Set PowerPointApp = New PowerPoint.Application
		If Err.Number = 429 Then
			MsgBox "PowerPoint could not found, aborting"
			Exit Sub
		End If
		
		' PPT 파일이 이미 열려있으면 추가로 열지 않고, 안열려있으면 새로 열기
		Set pt = PowerPointApp.Presentations(Current_Address & "\" & pptx_file_name)
		If Err.Number <> 0 Then
			PowerPointApp.Presentation.Open(Current_Address & "\" & pptx_file_name)
		End If
		
		PowerPointApp.Activate
		PowerPoint.Windows(pptx_file_name).Activate
		pt.Windows(1).Activate
		PowerPointApp.Visible = True
		
	On Error GoTo 0
	
	' Do something
	Set sht_page2 = ThisWorkbook.Sheets("Page2")
	Set sld = PowerPointApp.ActivatePresentation.Slides(sld_num)
	
	sld.Select
	
	' PPT 파일에서 기존에 있던 sharpes 삭제
	Debug.Print "ppt_page2 shapes count: " & sld.Shapes.Count
	If sld.Shapes.Count = 6 Then
		For i = 1 To 2
			sld.Shapes(sld.Shapes.Count).Delete
		Next i
	End If
	
	' copy plot chart from excel page2
	sht_page2.Activate
	ActiveSheet.ChartObjects(1).Activates
	ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
	sld.Shapes.Paste
	sld.Shapes(sld.Shapes.Count).LockAspectRatio = msoFalse
	sld.Shapes(sld.Shapes.Count).Height = 6.85 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Width = 24.4 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Left = 1.76 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Top = 5.27 * cm_to_px
	
	'copy table1 from excel page2
	sht_page2.Range(sht_page2.Range("C28"), sht_page2.Range("AL41")).CopyPicture Appearance:=xlScreen, Format:=xlPicture
	sld.Shapes.Paste
	sld.Shapes(sld.Shapes.Count).LockAspectRatio = msoFalse
	sld.Shapes(sld.Shapes.Count).Height = 4.79 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Width = 24.4 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Left = 1.76 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Top = 12.61 * cm_to_px
	
End Sub
Private Sub MakePPt_Page3(ByVal sld_num As Integer)

	Dim PowerPointApp As PowerPoint.Application
	Dim pt As Presentation
	Dim sld as PowerPoint.slide
	Dim shp as PowerPoint.Shape
	
	Dim sht_page3 as Worksheet
	
	Current_Address = Application.ThisWorkbook.PATH_INFO
	
	' Look for existing instance
	On Error Resume Next
		Set PowerPointApp = GetObject(Class="PowerPoint.Application")
		Err.Clear
		If PowerPoint Is Nothing Then Set PowerPointApp = New PowerPoint.Application
		If Err.Number = 429 Then
			MsgBox "PowerPoint could not found, aborting"
			Exit Sub
		End If
		
		' PPT 파일이 이미 열려있으면 추가로 열지 않고, 안열려있으면 새로 열기
		Set pt = PowerPointApp.Presentations(Current_Address & "\" & pptx_file_name)
		If Err.Number <> 0 Then
			PowerPointApp.Presentation.Open(Current_Address & "\" & pptx_file_name)
		End If
		
		PowerPointApp.Activate
		PowerPoint.Windows(pptx_file_name).Activate
		pt.Windows(1).Activate
		PowerPointApp.Visible = True
		
	On Error GoTo 0
	
	' Do something
	Set sht_page3 = ThisWorkbook.Sheets("Page3")
	Set sld = PowerPointApp.ActivatePresentation.Slides(sld_num)
	
	sld.Select
	
	' PPT 파일에서 기존에 있던 sharpes 삭제
	Debug.Print 'ppt_page3 sharpes count: " & sld.Shapes.Count
	If sld.Shapes.Count = 6 Then
		For i = 1 To 2
			sld.Shapes(sld.Shapes.Count).Delete
		Next i
	End If
	
	'copy plot chart from excel page3
	sht_page3.Activate
	ActiveSheet.ChartObjects(1).Activate
	ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, Format:=xlPicture
	sld.Shapes.Paste
	sld.Shapes(sld.Shapes.Count).LockAspectRatio = msoFalse
	sld.Shapes(sld.Shapes.Count).Height = 4.79 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Width = 24.4 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Left = 24.4 * cm_to_px
	sld.Shapes(sld.Shapes.Count).Top = 12.61 * cm_to_px
	
End Sub
Private Sub MakePPt_Page4(ByVal sld_num As Integer)

	Dim PowerPointApp As PowerPoint.Application
	Dim pt As Presentation
	Dim sld As PowerPoint.slide
	Dim shp As PowerPoint.Shape
	
	Dim sht_page4 As Worksheet
	
	
End Sub