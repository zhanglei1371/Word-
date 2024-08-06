VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WordAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public wd As Word.Application
'word中range、bookmarks、document、filedialog非vb对象，故需修改为object对象，所有对应的常量都应设置为值！
Sub excel写入测试()
Dim ex As Excel.Application
Set ex = GetObject(, "excel.application")
ex.ActiveSheet.Range("D3") = 456789
ex.ActiveSheet.Range("F6").Value = "这是一个测试！！！"
End Sub
Sub Api测试01()
 SleepApi32test
End Sub
Sub Api测试02()
   发送消息测试
End Sub
Sub VB窗体弹出()
Dim frm As New Form1
frm.Show
End Sub
Sub word写入测试()
'On Error Resume Next
Dim wd As Word.Application
Set wd = GetObject(, "word.application")
wd.Selection.TypeText "12345678" & vbNewLine
wd.ActiveDocument.Range.InsertAfter "sdafasdfsadf45856"
MsgBox "sdafsdaf"
End Sub
Sub ppt写入测试()
Dim pt As PowerPoint.Application
Set pt = GetObject(, "powerpoint.application")
pt.ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange = "这是本类模块在ppt中的一个测试！" '写入第一个ppt中第一页的第一个文本框
pt.ActiveWindow.Selection.TextRange.Text = "这个是对ppt中选定文本框写入的一个测试！！！" '写入选定文本框
End Sub
Sub 删除全文空白行()
Dim t#
Set wd = GetObject(, "word.application")
wd.ScreenUpdating = 0
t = Timer
Dim S As Object
Set S = wd.ActiveDocument.Content   '
S.Find.Execute "^13[ 　^t" & ChrW(160) & "^11^13]{1,}", , , 2, , , , , , "^p", 2
Set S = Nothing
wd.ScreenUpdating = 1
End Sub
Sub 计算()
Set wd = GetObject(, "word.application")
a = 4
B = 5
MsgBox a + B
wd.ActiveDocument.Content.InsertAfter a & B
End Sub
Sub 自动编号替换为手动编号()
Dim S As Object
Set wd = GetObject(, "word.application")
If wd.Selection.Type = wdSelectionIP Then wd.Selection.Expand wdParagraph
Set S = wd.Selection.Range
wd.Selection.Range.ListFormat.ConvertNumbersToText
With wd.Selection.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "([0-9]{1,})([.．、^9^32" & ChrW(160) & ChrW(12288) & "]{1,})"
    .Wrap = 0
    .Replacement.Text = "\1. "    '此处可改为顿号或其他
    .MatchWildcards = 1
    .Execute Replace:=wdReplaceAll
End With
End Sub
Sub 每行插入表格n个图()
On Error Resume Next
Set wd = GetObject(, "word.application")
wd.ScreenUpdating = False
Dim D As Object, a, P As InlineShape, t As Word.Table
Dim B$, C$ '这里就体现了不定义变量的缺点：不定义的话生成的图片下面没有名字。
If wd.Selection.Information(wdWithInTable) = True Then MsgBox "请将光标置于表格之外！": Exit Sub
With wd.FileDialog(1) '★★★若是此处写成msoFileDialogFilePicker则不起作用，因为filedialog不是vb的对象，是vba特有，故其所有常量都应写值！！！★★★
    .Title = "请选择..."
    If .Show = -1 Then
    n = InputBox("请输入表格的列数：", "列数", 3)
    m = .SelectedItems.Count
    Debug.Print "共有" & m & "个图片"; m
   h = IIf(m / n = Int(m / n), 2 * m / n, 2 * (Int(m / n) + 1))
   Set t = wd.ActiveDocument.Tables.Add(wd.Selection.Range, h, n)
    t.Borders.Enable = True
t.Borders.OutsideLineStyle = wdLineStyleDouble
        For Each a In .SelectedItems
            B = Split(a, "\")(UBound(Split(a, "\")))
            C = Split(B, ".")(0)
            Set P = wd.Selection.InlineShapes.AddPicture(FileName:=a, SaveWithDocument:=True)
            With P
                w = .Width
                .Width = Int(410 / n)
                .Height = .Width * .Height / w
            End With
            i = i + 1
            wd.Selection.MoveLeft wdCharacter, 1
            wd.Selection.MoveDown wdLine, 1
            wd.Selection.TypeText C
wd.Selection.Cells(1).Select
wd.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '决定了首行居中
wd.Selection.HomeKey
wd.Selection.MoveDown wdLine, -1
            wd.Selection.MoveRight wdCharacter, 2
            Debug.Print i, n
            If i = Val(n) Then
                wd.Selection.MoveRight wdCharacter, 1
                wd.Selection.Cells(1).Select
                wd.Selection.EndKey
                wd.Selection.MoveDown wdLine, 1
                i = 0
            End If
        Next
    End If
End With
wd.ScreenUpdating = True
End Sub
