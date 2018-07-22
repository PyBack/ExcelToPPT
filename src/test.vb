Sub PowerPoint_Open1()   ' CreateObject 사용 AP는 Application으로 됨

    Dim Ap As PowerPoint.Application
    Dim Pt As Presentation
     
    Set PPTApp = CreateObject("PowerPoint.Application"): Ap.Visible = True ' 활성화
    Set Pt = PPTApp.Presentations.Open(ThisWorkbook.Path & "\프리젠테이션1.pptx")
    ' Pt.SlideShowSettings.Run    ' 슬라이드쇼 실행
     
    Set Pt = Nothing
    Set Ap = Nothing
    
End Sub
Sub PowerPoint_Open2()    ' New 사용

    Dim Ap As New PowerPoint.Application
    Dim Pt As Presentation
    
    Ap.Visible = True
    Set Pt = Ap.Presentations.Open(ThisWorkbook.Path & "\프리젠테이션1.pptx")
    Pt.SlideShowSettings.Run    ' 슬라이드쇼 실행
     
    Set Pt = Nothing
    Set Ap = Nothing

End Sub
Sub PowerPoint_New1()   ' CreateObject 사용 AP는 Application으로 됨
    Dim Ap As PowerPoint.Application
    Dim Pt As Presentation
    Dim Sd As Slide
     
    Set Ap = CreateObject("PowerPoint.Application"): Ap.Visible = True ' 활성화
    Set Pt = Ap.Presentations.Add
    Set Sd = Pt.Slides.AddSlide(Pt.Slides.Count + 1, Pt.SlideMaster.CustomLayouts(1))
     
    Sd.Shapes.Range(1).TextEffect.Text = "안녕하세요"
    Sd.Shapes.Range(2).TextEffect.Text = "jfree입니다."
     
    Set Sd = Nothing
    Set Pt = Nothing
    Set Ap = Nothing
End Sub

Sub PowerPoint_New2()    ' New 사용
    Dim Ap As New PowerPoint.Application
    Dim Pt As Presentation
    Dim Sd As Slide
     
    Ap.Visible = True ' 활성화
    Set Pt = Ap.Presentations.Add
    Set Sd = Pt.Slides.AddSlide(Pt.Slides.Count + 1, Pt.SlideMaster.CustomLayouts(1))
     
    Sd.Shapes.Range(1).TextEffect.Text = "안녕하세요"
    Sd.Shapes.Range(2).TextEffect.Text = "AidenJFree1004."
     
    Set Sd = Nothing
    Set Pt = Nothing
    Set Ap = Nothing
End Sub

Sub ExcelRangeToPowerPoint()


    Dim PowerPointApp As PowerPoint.Application
    Dim myPresentation As PowerPoint.Presentation
    Dim mySlide As PowerPoint.Slide


    Dim i As Integer
    Dim BoxEntry As PowerPoint.Shape, BoxPronun As PowerPoint.Shape, BoxMean As PowerPoint.Shape, BoxIDX As PowerPoint.Shape
    Dim strEntry As String, strPron As String, strMean As String, strPOS As String, strIDX As String
    Dim r As Range, rng As Range
    
    
    ' Set rng = Sheet1.Range("C2:C33") '리스트 영역
    Set rng = Sheets(1).Range("C2:C4")
    
    
    On Error Resume Next '
        Set PowerPointApp = GetObject(Class:="PowerPoint.Application")
        Err.Clear 'Clear the error between errors

        If PowerPointApp Is Nothing Then Set PowerPointApp = CreateObject(Class:="PowerPoint.Application")
        If Err.Number = 429 Then '
            MsgBox "PowerPoint could not be found, aborting."
            Exit Sub
        End If

    On Error GoTo 0
    
    PowerPointApp.Visible = True
    PowerPointApp.Activate

    Set myPresentation = PowerPointApp.Presentations.Add '새 PPT 문서 생성
    ' Set myLayout = PowerPointApp.PpSlideLayout.PpLayoutBlank
    Set myLayout = myPresentation.Designs(1).SlideMaster.CustomLayouts(7)
    
    i = 0
    
    For Each r In rng

        i = i + 1
        strIDX = Replace(r.Offset(0, -2).Value, "idx", "")
        strEntry = r.Offset(0, 0).Value
        strPron = r.Offset(0, 1).Value
        strPOS = r.Offset(0, 2).Value
        strMean = r.Offset(0, 3).Value
        

        Set mySlide = myPresentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank) '슬라이드1장씩 추가

        With mySlide
            .BackgroundStyle = 1
            .Background.Fill.ForeColor.RGB = RGB(20, 20, 20)
        End With
        
        Set BoxIDX = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=50, Top:=80, Width:=600, Height:=50)
  

        With BoxIDX.TextFrame.TextRange
            .Text = strIDX
            .Font.Bold = True
            .Font.Size = 35
            .Font.Color.RGB = RGB(204, 255, 255)
            .ParagraphFormat.Alignment = ppAlignCenter
          End With

        Set BoxEntry = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=50, Top:=150, Width:=600, Height:=80)

        With BoxEntry.TextFrame.TextRange
            .Text = strEntry
            .Font.Bold = msoCTrue
            .Font.Size = 75
            .Font.Color.RGB = RGB(255, 212, 132)
            .ParagraphFormat.Alignment = ppAlignCenter
        End With

        Set BoxPronun = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=50, Top:=250, Width:=600, Height:=50)

        With BoxPronun.TextFrame.TextRange
            .Text = strPron
            .Font.Size = 40
            .Font.Color.RGB = RGB(204, 255, 204)
            .ParagraphFormat.Alignment = ppAlignCenter
        End With

        Set BoxMean = mySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=50, Top:=330, Width:=600, Height:=50)

        With BoxMean.TextFrame.TextRange
            .Text = "[" & strPOS & "]" & strMean
            .Font.Size = 28
            .Font.Color.RGB = RGB(204, 255, 255)
            .ParagraphFormat.Alignment = ppAlignCenter
        End With
        
    Next r
    
    Set myPresentation = Nothing
    Set PowerPointApp = Nothing
    
    MsgBox i & "장 슬라이드 생성 완료"

End Sub



