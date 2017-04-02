VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "土木团学学术部专用 签到系统"
   ClientHeight    =   5595
   ClientLeft      =   5925
   ClientTop       =   4845
   ClientWidth     =   15990
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   15990
   Begin VB.ListBox ListGroup 
      Height          =   2895
      Left            =   11280
      TabIndex        =   28
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "增加座位"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   6960
      Width           =   1455
   End
   Begin VB.OptionButton Option5 
      Caption         =   "姓名"
      Height          =   615
      Left            =   13080
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Caption         =   "手机号"
      Height          =   615
      Left            =   11880
      TabIndex        =   19
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "座位号"
      Height          =   615
      Left            =   4320
      TabIndex        =   18
      Top             =   6720
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "姓名首字母"
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   240
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox TextSearch 
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1440
      Width           =   5895
   End
   Begin VB.CommandButton Command10 
      Caption         =   "导出名单"
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "搜索选项"
      Height          =   615
      Left            =   6120
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.ListBox ListNum 
      Height          =   2895
      Left            =   12720
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000FF00&
      Caption         =   "签到"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      MaskColor       =   &H00808080&
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "取消签到"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ListBox ListPY 
      Height          =   2895
      Left            =   7920
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.ListBox ListYesNo 
      Height          =   2895
      Left            =   14520
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ListBox ListWhere 
      Height          =   2895
      Left            =   2040
      TabIndex        =   4
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox ListSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "搜索（Enter）"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin VB.ListBox ListName 
      Height          =   2895
      Left            =   9600
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Labelmuch 
      BackStyle       =   0  'Transparent
      Caption         =   "0 / 0"
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "座位分配情况"
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "姓名首字母       姓名                 分组              手机号码            是否签到"
      Height          =   495
      Left            =   7920
      TabIndex        =   24
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label2 
      Caption         =   "By Zhenly         www.zhenly.cn"
      Height          =   855
      Left            =   13080
      TabIndex        =   22
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label LabelS 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名："
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   1500
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "搜索选项："
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label LabelHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "请先分配座位"
      Height          =   1095
      Left            =   2640
      TabIndex        =   16
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   0
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label LabelName 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   15
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label LabelNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Labeltxt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1215
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "总人数人数：     已签到人数：     未签到人数：    签到率：  "
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   9255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
abc = Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
 Open App.Path & "\output" & abc & ".xls" For Output As #1 '创建/打文件
 Print #1, "姓名" & Chr(9) & "手机号码" & Chr(9) & "组别" & Chr(9) & "签到"
  For I = 0 To ListName.ListCount - 1  '输入文件
    Print #1, ListName.List(I) & Chr(9) & ListNum.List(I) & Chr(9) & ListGroup.List(I) & Chr(9) & ListYesNo.List(I) 'todo
  Next I
  Close #1
  
  
  
  
  LabelHelp.Caption = "已导出到" & App.Path & "\output" & abc & ".xls"
  Exit Sub
h:
  MsgBox ("请确认文件" & App.Path & "\output" & abc & ".xls 未被占用")
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

'If Option2.Value = True Then
'fuck = 0
'For I = 0 To ListWhere.ListCount - 1
'If ListWhere.List(I) = TextSearch.Text Then
'    If ListWhere.ItemData(I) >= 0 Then
'    ListSearch.AddItem (ListName.List(ListWhere.ItemData(I)))
'    ListSearch.ItemData(ListSearch.ListCount - 1) = ListWhere.ItemData(I)
'    End If
'End If
'Next I

'If ListSearch.ListCount > 0 Then
'ListSearch.ListIndex = 0
'Else
'LabelHelp.Caption = "该座位还没有被分配"
'End If

'End If

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

'ListWhere.ItemData(ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex))) = -1
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
TextSearch.SetFocus
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

'If Howmuch < Maxmuch Then

If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No" Then






ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "Yes"
'Do
'xss = Rndz(0, ListWhere.ListCount + 1)
'MsgBox (xss)
'Loop Until ListWhere.ItemData(xss - 1) = -1
'ListWhere.ItemData(xss - 1) = ListSearch.ItemData(ListSearch.ListIndex)
'ListName.ItemData(ListSearch.ItemData(ListSearch.ListIndex)) = xss - 1
Labeltxt.Caption = " " & ListGroup.List(ListSearch.ItemData(ListSearch.ListIndex))
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
'Else
'LabelHelp.Caption = "座位已满"
'End If
TextSearch.SetFocus
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
   ' ListWhere.Clear
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
        ReDim Preserve c(I)
         Input #1, a(I), b(I), c(I)
        I = I + 1
    Loop
h:
    Close #1
    For I = 1 To UBound(a)
       ListName.AddItem (a(I))
       ListNum.AddItem (b(I))
       ListGroup.AddItem (c(I))
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

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo h
abc = Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
 Open App.Path & "\output(auto)" & abc & ".xls" For Output As #1 '创建/打文件
 Print #1, "姓名" & Chr(9) & "手机号码" & Chr(9) & "组别" & Chr(9) & "签到"
  For I = 0 To ListName.ListCount - 1  '输入文件
    Print #1, ListName.List(I) & Chr(9) & ListNum.List(I) & Chr(9) & ListGroup.List(I) & Chr(9) & ListYesNo.List(I) 'todo

  Next I
  Close #1
  
  
  
  
  LabelHelp.Caption = "已导出到" & App.Path & "\output(auto)" & abc & ".xls"
  Exit Sub
h:
  MsgBox ("自动保存失败，请确认文件" & App.Path & "\output(auto)" & abc & ".xls 未被占用")
End Sub

Private Sub Label2_Click()
Shell "explorer http://www.zhenly.cn"
End Sub

Private Sub ListName_Click()
ListYesNo.ListIndex = ListName.ListIndex
ListNum.ListIndex = ListName.ListIndex
ListPY.ListIndex = ListName.ListIndex
'If ListName.ItemData(ListName.ListIndex) >= 0 Then
'ListWhere.ListIndex = ListName.ItemData(ListName.ListIndex)
'Else
'ListWhere.ListIndex = -1
'End If
End Sub

Private Sub ListNum_Click()
ListName.ListIndex = ListNum.ListIndex
ListYesNo.ListIndex = ListNum.ListIndex
'ListWhere.ListIndex = ListNum.ListIndex
ListGroup.ListIndex = ListNum.ListIndex
ListPY.ListIndex = ListNum.ListIndex
End Sub

Private Sub ListPY_Click()
ListGroup.ListIndex = ListPY.ListIndex
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
Labeltxt.Caption = ListGroup.List(ListSearch.ItemData(ListSearch.ListIndex))
Labeltxt.ForeColor = &HC000&
Command3.Enabled = True
Command5.Enabled = False
End If

LabelNum.Caption = "Phone：" & ListNum.List(ListSearch.ItemData(ListSearch.ListIndex))

End Sub

Private Sub ListSearch_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
     If ListSearch.ListCount > 0 Then
     If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No" Then
     Command5_Click
     Else
     Command3_Click
     End If
     End If
End If
End Sub

Private Sub ListYesNo_Click()
ListGroup.ListIndex = ListYesNo.ListIndex
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

Private Sub TextSearch_Change()
Command5.Enabled = False
Command3.Enabled = False
End Sub

Private Sub TextSearch_KeyPress(KeyAscii As Integer)
     If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
        KeyAscii = KeyAscii - 32
    End If
End Sub

Private Sub TextSearch_KeyUp(KeyCode As Integer, Shift As Integer)
Command2_Click
 If KeyCode = 13 Then
     If ListSearch.ListCount > 0 Then
     If ListYesNo.List(ListSearch.ItemData(ListSearch.ListIndex)) = "No" Then
     Command5_Click
     Else
     Command3_Click
     End If
     End If
End If
If KeyCode = 38 Then
    If ListSearch.ListCount > 0 Then
        If ListSearch.ListIndex - 1 >= 0 Then
            ListSearch.ListIndex = ListSearch.ListIndex - 1
            ListSearch.SetFocus
        End If
    End If
    
End If
If KeyCode = 40 Then
    If ListSearch.ListCount > 0 Then
        If ListSearch.ListIndex + 1 <= ListSearch.ListCount Then
            ListSearch.ListIndex = ListSearch.ListIndex + 1
            ListSearch.SetFocus
        End If
    End If
End If
End Sub
