
Dim Howmuch As Integer
Dim Maxmuch As Integer

Private Sub Command1_Click()
Dim temp As String
Dim name As String
Dim begin As Integer
Dim over As Integer
Dim count As Integer
name = InputBox("请输入座位的前缀")
Do
temp = InputBox("请输入座位的开始序号(请输入整数)")
Loop Until IsNumeric(temp) = True
begin = temp

Do
temp = InputBox("请输入座位的结束序号(请输入整数)")
Loop Until IsNumeric(temp) = True
over = temp


count = ListWhere.ListCount
For I = begin To over
ListWhere.AddItem (name & " " & I)
ListWhere.ItemData(count) = -1
count = count + 1
Next I
Maxmuch = Maxmuch + over - begin + 1
Labelmuch.Caption = Howmuch & " / " & Maxmuch
End Sub

Private Sub Command10_Click()

On Error GoTo h
 Open App.Path & "\output.xls" For Output As #1   '创建/打文件
 Print #1, "姓名" & Chr(9) & "手机号码" & Chr(9) & "座位" & Chr(9) & "签到"
  For I = 0 To ListName.ListCount - 1  '输入文件
  If ListName.ItemData(I) > -1 Then
    Print #1, ListName.List(I) & Chr(9) & ListNum.List(I) & Chr(9) & ListWhere.List(ListName.ItemData(I)) & Chr(9) & ListYesNo.List(I)
Else
 Print #1, ListName.List(I) & Chr(9) & ListNum.List(I) & Chr(9) & "未分配" & Chr(9) & ListYesNo.List(I)
End If
  Next I
  Close #1




  LabelHelp.Caption = "已导出到" & App.Path & "\output.xls"
  Exit Sub
h:
  MsgBox ("请确认文件" & App.Path & "\output.xls 未被占用")
End Sub

Private Sub Command2_Click()


ListSearch.Clear


If Option1.Value = True Then

For I = 0 To ListPY.ListCount - 1
If Left(ListPY.List(I), Len(TextSearch.Text)) = Left(TextSearch.Text, Len(TextSearch.Text)) And Len(TextSearch.Text) <> 0 Then
    ListSearch.AddItem (ListName.List(I))
    ListSearch.ItemData(ListSearch.ListCount - 1) = I
End If
Next I
If ListSearch.ListCount > 0 Then
ListSearch.ListIndex = 0
Else
LabelHelp.Caption = "未搜索到结果"
End If

End If

If Option2.Value = True Then
fuck = 0
For I = 0 To ListWhere.ListCount - 1
If ListWhere.List(I) = TextSearch.Text Then
    If ListWhere.ItemData(I) >= 0 Then
    ListSearch.AddItem (ListName.List(ListWhere.ItemData(I)))
    ListSearch.ItemData(ListSearch.ListCount - 1) = ListWhere.ItemData(I)
    End If
End If
Next I

If ListSearch.ListCount > 0 Then
ListSearch.ListIndex = 0
Else
LabelHelp.Caption = "该座位还没有被分配"
End If

End If

If Option3.Value = True Then

For I = 0 To ListNum.ListCount - 1
If Left(ListNum.List(I), Len(TextSearch.Text)) = Left(TextSearch.Text, Len(TextSearch.Text)) And Len(TextSearch.Text) <> 0 Then
    ListSearch.AddItem (ListName.List(I))
    ListSearch.ItemData(ListSearch.ListCount - 1) = I
End If
Next I
If ListSearch.ListCount > 0 Then
ListSearch.ListIndex = 0
Else
LabelHelp.Caption = "未搜索到结果"
End If

End If

If Option5.Value = True Then

For I = 0 To ListNum.ListCount - 1
If Left(ListName.List(I), Len(TextSearch.Text)) = Left(TextSearch.Text, Len(TextSearch.Text)) And Len(TextSearch.Text) <> 0 Then
    ListSearch.AddItem (ListName.List(I))
    ListSearch.ItemData(ListSearch.ListCount - 1) = I
End If
Next I
If ListSearch.ListCount > 0 Then
ListSearch.ListIndex = 0
Else
LabelHelp.Caption = "未搜索到结果"
End If

End If


End Sub

Private Sub Command3_Click()

If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "Yes" Then


ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No"
Labeltxt.Caption = "未签到"
Labeltxt.ForeColor = &HFF&

ListWhere.ItemData(ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex))) = -1
ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex)) = 0

Dim yes As Integer
Dim no As Integer
yes = 0
no = 0
For I = 0 To ListName.ListCount - 1
If ListYesNo.List(I) = "Yes" Then
yes = yes + 1
Else
no = no + 1
End If
Next I
Label3.Caption = "总人数人数： " & ListName.ListCount & "  已签到人数：  " & yes & "  未签到人数：  " & no & "  签到率： " & Format(CSng(yes / ListName.ListCount), "0.00")
Command3.Enabled = False
Command5.Enabled = True
Howmuch = Howmuch - 1
Labelmuch.Caption = Howmuch & " / " & Maxmuch
Else
LabelHelp.Caption = "还没有签到呢！"
End If
End Sub




Public Function py(mystr As String) As String
    If Asc(mystr) < 0 Then
        If Asc(Left$(mystr, 1)) < Asc("啊") Then
            py = "0"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("啊") And Asc(Left$(mystr, 1)) < Asc("芭") Then
            py = "A"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("芭") And Asc(Left$(mystr, 1)) < Asc("擦") Then
            py = "B"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("擦") And Asc(Left$(mystr, 1)) < Asc("搭") Then
            py = "C"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("搭") And Asc(Left$(mystr, 1)) < Asc("蛾") Then
            py = "D"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("蛾") And Asc(Left$(mystr, 1)) < Asc("发") Then
            py = "E"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("发") And Asc(Left$(mystr, 1)) < Asc("噶") Then
            py = "F"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("噶") And Asc(Left$(mystr, 1)) < Asc("哈") Then
            py = "G"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("哈") And Asc(Left$(mystr, 1)) < Asc("击") Then
            py = "H"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("击") And Asc(Left$(mystr, 1)) < Asc("喀") Then
            py = "J"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("喀") And Asc(Left$(mystr, 1)) < Asc("垃") Then
            py = "K"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("垃") And Asc(Left$(mystr, 1)) < Asc("妈") Then
            py = "L"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("妈") And Asc(Left$(mystr, 1)) < Asc("拿") Then
            py = "M"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("拿") And Asc(Left$(mystr, 1)) < Asc("哦") Then
            py = "N"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("哦") And Asc(Left$(mystr, 1)) < Asc("啪") Then
            py = "O"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("啪") And Asc(Left$(mystr, 1)) < Asc("期") Then
            py = "P"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("期") And Asc(Left$(mystr, 1)) < Asc("然") Then
            py = "Q"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("然") And Asc(Left$(mystr, 1)) < Asc("撒") Then
            py = "R"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("撒") And Asc(Left$(mystr, 1)) < Asc("塌") Then
            py = "S"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("塌") And Asc(Left$(mystr, 1)) < Asc("挖") Then
            py = "T"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("挖") And Asc(Left$(mystr, 1)) < Asc("昔") Then
            py = "W"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("昔") And Asc(Left$(mystr, 1)) < Asc("压") Then
            py = "X"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("压") And Asc(Left$(mystr, 1)) < Asc("匝") Then
            py = "Y"
            Exit Function
        End If
        If Asc(Left$(mystr, 1)) >= Asc("匝") Then
            py = "Z"
            Exit Function
        End If
    Else
        If UCase$(mystr) <= "Z" And UCase$(mystr) >= "A" Then
            py = UCase$(Left$(mystr, 1))
        Else
            py = mystr
        End If
    End If
End Function
Public Function test(str As String) As String
    Dim tmp As String
    For I = 1 To Len(str)
        tmp = tmp & py(Mid$(str, I, 1))
    Next I
    test = tmp
End Function
Private Function Rndz(a As Long, b As Long)
    Randomize
    Rndz = Int((a - b + 1) * Rnd() + b)
End Function



Private Sub Command5_Click()
Dim xss As Integer

If Howmuch < Maxmuch Then

If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No" Then






ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "Yes"
Do
xss = Rndz(0, ListWhere.ListCount + 1)
'MsgBox (xss)
Loop Until ListWhere.ItemData(xss - 1) = -1
ListWhere.ItemData(xss - 1) = ListSearch.ItemData(ListSearch.ListIndex)
ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex)) = xss - 1
Labeltxt.Caption = ListWhere.List(xss - 1)
Labeltxt.ForeColor = &HC000&


Dim yes As Integer
Dim no As Integer
yes = 0
no = 0
For I = 0 To ListName.ListCount - 1
If ListYesNo.List(I) = "Yes" Then
yes = yes + 1
Else
no = no + 1
End If
Next I
Howmuch = yes
Labelmuch.Caption = Howmuch & " / " & Maxmuch


Label3.Caption = "总人数人数： " & ListName.ListCount & "  已签到人数：  " & yes & "  未签到人数：  " & no & "  签到率： " & Format(CSng(yes / ListName.ListCount), "0.00")
Command5.Enabled = False
Command3.Enabled = True
Else
LabelHelp.Caption = "无法再次进行签到！"
End If
Else
LabelHelp.Caption = "座位已满"
End If

End Sub







Private Sub Command9_Click()
If Command9.Caption = "搜索选项" Then
Me.Width = 16005
Command9.Caption = "返回"
Else
Me.Width = 7935
Command9.Caption = "搜索选项"
End If
End Sub

Private Sub Form_Load()
If Dir(App.Path & "\m.txt") = "" Then
MsgBox ("请把数据源存放于" & App.Path & "\m.txt 下")
End
Else
'不存在else'存在end if
    Dim fso  As Object, fs As Object
    Set fso = CreateObject("scripting.Filesystemobject")
    Dim S As String, ss As String
    ss = App.Path & "/m.txt"
    Set fs = fso.opentextfile(ss)  '打文件
    S = fs.readall   '读取所文本
    S = Replace(S, Chr(9), vbCrLf)  '替换文本
    Dim fs2 As Object
    fs.Close   '关闭文本
    Kill ss    '删除文本
    Set fs2 = fso.createtextfile(ss)   '创建文本
    fs2.write S   '写入文本
    fs2.Close    '关闭文本
    ListName.Clear
    ListWhere.Clear
    ListYesNo.Clear
    Dim a() As String
    Dim b() As String
    Dim c() As String
    Dim I As Integer
    I = 1
    Open ss For Input As #1
    On Error GoTo h
    Do While Not EOF(1)
        ReDim Preserve a(I)
        ReDim Preserve b(I)
         Input #1, a(I), b(I)
        I = I + 1
    Loop
h:
    Close #1
    For I = 1 To UBound(a)
       ListName.AddItem (a(I))
       ListNum.AddItem (b(I))
       ListYesNo.AddItem ("No")
       ListPY.AddItem (test(a(I)))
    Next

Dim yes As Integer
Dim no As Integer
yes = 0
no = 0
For I = 0 To ListName.ListCount - 1
If ListYesNo.List(I) = "Yes" Then
yes = yes + 1
Else
no = no + 1
End If
Next I
Label3.Caption = "总人数人数： " & ListName.ListCount & "  已签到人数：  " & yes & "  未签到人数：  " & no & "  签到率： " & Format(CSng(yes / ListName.ListCount), "0.00")
For I = 0 To ListName.ListCount - 1
ListName.ItemData(I) = -1
Next I
End If
End Sub

Private Sub Label2_Click()
Shell "explorer http://www.zhenly.cn"
End Sub

Private Sub ListName_Click()
ListYesNo.ListIndex = ListName.ListIndex
ListNum.ListIndex = ListName.ListIndex
ListPY.ListIndex = ListName.ListIndex
If ListName.ItemData(ListName.ListIndex) >= 0 Then
ListWhere.ListIndex = ListName.ItemData(ListName.ListIndex)
Else
ListWhere.ListIndex = -1
End If
End Sub

Private Sub ListNum_Click()
ListName.ListIndex = ListNum.ListIndex
ListYesNo.ListIndex = ListNum.ListIndex
'ListWhere.ListIndex = ListNum.ListIndex
ListPY.ListIndex = ListNum.ListIndex
End Sub

Private Sub ListPY_Click()
'ListWhere.ListIndex = ListPY.ListIndex
ListName.ListIndex = ListPY.ListIndex
ListNum.ListIndex = ListPY.ListIndex
ListYesNo.ListIndex = ListPY.ListIndex
End Sub

Private Sub ListSearch_Click()
LabelName.Caption = ListName.List(ListSearch.ItemData(ListSearch.ListIndex))
If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No" Then
Labeltxt.Caption = "未签到"
Labeltxt.ForeColor = &HFF&
Command5.Enabled = True
Command3.Enabled = False
Else
Labeltxt.Caption = ListWhere.List(ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex)))
Labeltxt.ForeColor = &HC000&
Command3.Enabled = True
Command5.Enabled = False
End If

LabelNum.Caption = "Phone：" & ListNum.List(ListSearch.ItemData(ListSearch.ListIndex))

End Sub

Private Sub ListYesNo_Click()
'ListWhere.ListIndex = ListYesNo.ListIndex
ListName.ListIndex = ListYesNo.ListIndex
ListNum.ListIndex = ListYesNo.ListIndex
ListPY.ListIndex = ListYesNo.ListIndex
End Sub



Private Sub Option1_Click()
LabelS.Caption = Option1.Caption & ":"
TextSearch.Text = ""
End Sub

Private Sub Option2_Click()

LabelS.Caption = Option2.Caption & ":"
TextSearch.Text = ""
End Sub

Private Sub Option3_Click()

LabelS.Caption = Option3.Caption & ":"
TextSearch.Text = ""
End Sub

Private Sub Option4_Click()
LabelS.Caption = Option4.Caption & ":"
TextSearch.Text = ""
End Sub

Private Sub Option5_Click()
LabelS.Caption = Option5.Caption & ":"
TextSearch.Text = ""
End Sub

Private Sub TextSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command2_Click
    End If
     If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub TextSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Command2_Click
End Sub
