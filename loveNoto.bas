Private Function r(i As String, o As String, Optional findBold As Integer)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.Name = i
        '.Font.bold = wdToggle   'mixed
        If (findBold = 1) Then
            .Font.bold = True
        End If
        If (findBold = -1) Then
            .Font.bold = False
        End If
        With .Replacement
            .ClearFormatting
            .Font.Name = o
            If findBold = 1 Then
                .Font.bold = False
            End If
        End With
        .Execute FindText:="", ReplaceWith:="", Format:=True, Replace:=wdReplaceAll
    End With
End Function
Sub loveNoto()
    '
    ' 挚爱思源 宏
    '
    rtn = MsgBox("这是一篇中文文稿吗？按「否」将以西文文稿处理，按「取消」将停止处理。", 3, 文档类型)
    '3 vbYesNoCancel - 显示“是”，“否”和“取消” 按钮。
    If rtn = vbCancel Then
        Exit Sub
    End If
    Call r("Times New Roman", "Adobe Garamond Pro")
    Call r("宋体", "思源宋体 CN Light", -1)
    Call r("宋体", "思源宋体 CN Medium", 1)
    Call r("宋体", "思源宋体 CN")
    Call r("黑体", "Noto Sans CJK SC Regular", -1)
    Call r("黑体", "Noto Sans CJK SC Medium", 1)
    Call r("黑体", "Noto Sans CJK SC Medium")
    Call r("楷体", "方正聚珍新仿简体")
    Call r("楷体_GB2312", "方正聚珍新仿简体")
    Call r("仿宋", "方正清仿宋 简 Bold")
    Call r("仿宋_GB2312", "方正清仿宋 简 Bold")
    Call r("Times New Roman", "Adobe Garamond Pro")
    
    If rtn = vbOK Then
        With ActiveDocument.Content.Find
        .Text = "…"
        With .Replacement
            .Font.Name = "华文中宋"
            .Text = "…"
        End With
        .Execute Replace:=wdReplaceAll
    End With
    End If
End Sub
Sub margin()
'
' 页边距 宏
'
'
    With Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .SectionDirection = wdSectionDirectionLtr
        .LinesPage = 48
        .LayoutMode = wdLayoutModeLineGrid
    End With
End Sub

Sub 排版()
    Call loveNoto
    Call margin
End Sub
