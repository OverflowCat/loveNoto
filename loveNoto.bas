Private Function r(i As String, o As String, Optional findBold As Integer)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Font.Name = i      '查找i体
        .Font.bold = wdToggle   'mixed
        If (findBold = 1) Then
            .Font.bold = True
        End If
        If (findBold = -1) Then
            .Font.bold = False
        End If
        With .Replacement
            .ClearFormatting
            .Font.Name = o    '替换成o体
            If findBold = 1 Then
                .Font.bold = False
            End If
        End With
        .Execute FindText:="", ReplaceWith:="", Format:=True, Replace:=wdReplaceAll
    End With
End Function

Sub 挚爱思源()
    '
    ' 挚爱思源 宏
    '
    rtn = MsgBox("这是一篇中文文稿吗？按「否」将以西文文稿处理，按「取消」将停止处理。", 3, 文档类型)
    '3 vbYesNoCancel - 显示“是”，“否”和“取消” 按钮。
    If rtn = vbCancel Then
        Exit Sub
    End If
    Call r("宋体", "思源宋体 CN Light", -1)
    Call r("宋体", "思源宋体 CN Medium", 1)
    Call r("黑体", "Noto Sans CJK SC Regular", -1)
    Call r("黑体", "Noto Sans CJK SC Medium", 1)
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
