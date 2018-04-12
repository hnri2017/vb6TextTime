VERSION 5.00
Begin VB.UserControl TextTime 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "TextTime.ctx":0000
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "TextTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const tHEIGHT As Long = 300
Private Const INPUTEND As Integer = 2   '每次输入最多只能输入2个数字


Dim intInputCount As Integer    '记录输入次数，2次后归零
Dim strOriginal As String       '更改之前的原始值。分和秒只能在00~59之间，时在00~23之间


Public Property Get Alignment() As Long
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal lNewAt As Long)
'    Text1.Alignment = lNewAt
    Text1.Alignment = IIf(lNewAt = 0 Or lNewAt = 1 Or lNewAt = 2, lNewAt, 0)
End Property


Public Property Get Appearance() As Long
    Appearance = Text1.Appearance
End Property

Public Property Let Appearance(ByVal lNewAp As Long)
    Text1.Appearance = IIf(lNewAp = 0 Or lNewAp = 1, lNewAp, 1)
End Property


Public Property Get BorderStyle() As Long
    BorderStyle = Text1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal lNewBs As Long)
    Text1.BorderStyle = IIf(lNewBs = 0 Or lNewBs = 1, lNewBs, 1)
End Property


Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal bNewEb As Boolean)
    Text1.Enabled = bNewEb
End Property


Public Property Get Hour() As Integer
    If Len(Trim(Text1.Text)) = 8 Then
        Hour = CInt(Mid(Text1.Text, 1, 2))
    Else
        Hour = 0
    End If
End Property


Public Property Get Minute() As Integer
    If Len(Trim(Text1.Text)) = 8 Then
        Minute = CInt(Mid(Text1.Text, 4, 2))
    Else
        Minute = 0
    End If
End Property


Public Property Get Second() As Integer
    If Len(Trim(Text1.Text)) = 8 Then
        Second = CInt(Mid(Text1.Text, 7, 2))
    Else
        Second = 0
    End If
End Property


Public Property Get Time() As Date
    If Len(Trim(Text1.Text)) = 8 Then
        Time = CDate(Text1.Text)
    Else
        Time = Format(Now, "hh:mm:ss")
    End If
End Property

Public Property Let Time(ByVal dNewTime As Date)
    Text1.Text = Format(Date, "hh:mm:ss")
End Property


Private Sub Text1_Click()

    With Text1
        Select Case .SelStart   '=0~8
            Case 0, 1, 2
                .SelStart = 0
                .SelLength = 2  '选中 小时 的两个数字
            Case 3, 4, 5
                .SelStart = 3
                .SelLength = 2  '选中 分钟 的两个数字
            Case 6, 7, 8
                .SelStart = 6
                .SelLength = 2  '选中 秒钟 的两个数字
        End Select
        strOriginal = .SelText  '保存选中时的值，当输入非法值时用此值替代
    End With
    
    intInputCount = 0   '复位输入次数
    
End Sub

Private Sub Text1_GotFocus()
    
    '主要防止非鼠标点击时间文本框时无法定位时、分、秒中的一个
    '代码内容同单击事件内容
    
    With Text1
        Select Case .SelStart   '=0~8
            Case 0, 1, 2
                .SelStart = 0
                .SelLength = 2
            Case 3, 4, 5
                .SelStart = 3
                .SelLength = 2
            Case 6, 7, 8
                .SelStart = 6
                .SelLength = 2
        End Select
        strOriginal = .SelText
    End With
    
    intInputCount = 0
    
End Sub

'Private Sub UserControl_Click()
'    Call Text1_Click
'End Sub
'
'Private Sub UserControl_GotFocus()
'    Call Text1_GotFocus
'End Sub

Private Sub UserControl_Initialize()
    Text1.Text = "00:00:00"     '初始化 设置时间 的文本框值
    Text1.Height = tHEIGHT      '根据需要的高度设置
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '←↑→↓,只处理四个方向键，其它键取消输入
    
    Dim strLastTwo As String    '记录按下键前的 值
    Dim strCurValue As String   '表示按下键改变后的 值
    
    If KeyCode > 36 And KeyCode < 41 Then intInputCount = 0 'Keycode范围限制很重要

    With Text1
        Select Case KeyCode
            Case vbKeyLeft      '←，顺序只能从秒到分，从分到时，到时后再按无效
            
                Select Case .SelStart
                    Case 3, 4, 5    '如果在分上，则向前移至时上。理论上来只3就可以了，4和5防意外，以下同
                        .SelStart = 0
                        .SelLength = 2
                    Case 6, 7, 8    '如果在秒上，则向前移至分上
                        .SelStart = 3
                        .SelLength = 2
                End Select
        
            Case vbKeyRight     '→，顺序只能从时到分，从分到秒，到秒后再按无效
            
                Select Case .SelStart
                    Case 0, 1, 2    '如果在时上，则向后移至分上
                        .SelStart = 3
                        .SelLength = 2
                    Case 3, 4, 5    '如果在分上，则向后移至秒上
                        .SelStart = 6
                        .SelLength = 2
                End Select
            
            Case vbKeyUp, vbKeyDown '↑↓
            
                strLastTwo = .SelText
                Select Case KeyCode
                    Case vbKeyUp    '↑表示按增量1向上加1
                        If Val(strLastTwo) < 9 Then
                            strCurValue = "0" & CStr(Val(strLastTwo) + 1)   '处理01~09
                        Else
                            strCurValue = CStr(Val(strLastTwo) + 1)
                            If .SelStart = 0 Then
                                If Val(strLastTwo) = 23 Then strCurValue = "00" '表示从23 跳至 00
                            ElseIf .SelStart = 3 Or .SelStart = 6 Then
                                If Val(strLastTwo) = 59 Then strCurValue = "00" '表示从59 跳至 00
                            End If
                        End If

                    Case vbKeyDown  '↓表示按增量1向下减1
                        If Val(strLastTwo) > 10 Then
                            strCurValue = CStr(Val(strLastTwo) - 1)
                        Else
                            strCurValue = "0" & CStr(Val(strLastTwo) - 1)   '处理01~09
                            If .SelStart = 0 Then
                                If Val(strLastTwo) = 0 Then strCurValue = "23"  '表示从00 跳至 23
                            ElseIf .SelStart = 3 Or .SelStart = 6 Then
                                If Val(strLastTwo) = 0 Then strCurValue = "59"  '表示从00 跳至 59
                            End If
                        End If
                End Select
                
                .SelText = strCurValue  '将最终有效值替换选中的值
                
                Select Case .SelStart   '再重新定位，以保位置无误
                    Case 0, 1, 2
                        .SelStart = 0
                        .SelLength = 2
                    Case 3, 4, 5
                        .SelStart = 3
                        .SelLength = 2
                    Case 6, 7, 8
                        .SelStart = 6
                        .SelLength = 2
                End Select
                
        End Select
    End With
    
    KeyCode = 0 'with块中有处理过输入值，至此应取消系统输入，否则输入两次值
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    '只能输入0~9这个10个数字
    
    Dim strFull As String   '表示最终输入值
    
    With Text1
        If KeyAscii > 47 And KeyAscii < 59 And intInputCount < INPUTEND And Len(.SelText) = INPUTEND Then

            intInputCount = intInputCount + 1
            
            If intInputCount = 1 Then
                strFull = "0" & Chr(KeyAscii)  '处理01~09
            ElseIf intInputCount = 2 Then
                strFull = Right(.SelText, 1) & Chr(KeyAscii)
                If .SelStart = 0 Then   '当前定位在小时上
                    If Val(strFull) > 23 Then strFull = strOriginal '两次输入组成的数字为非法值时替换为原值
                Else    '当前定位在分钟或秒钟上
                    If Val(strFull) > 59 Then strFull = strOriginal

                End If
                
            End If
            
            .SelText = strFull  '将最终值替换选中值
            
            Select Case .SelStart
                Case 0, 1, 2
                    .SelStart = 0
                    .SelLength = 2
                Case 3, 4, 5
                    .SelStart = 3
                    .SelLength = 2
                Case 6, 7, 8
                    .SelStart = 6
                    .SelLength = 2
            End Select
        End If
        
        If intInputCount = INPUTEND Then intInputCount = 0  '连续输入两次后次数归零
        
    End With
    
    KeyAscii = 0    'with块中有处理过输入值，至此应取消系统输入，否则输入两次值
    
End Sub


Private Sub UserControl_Resize()
    UserControl.Height = tHEIGHT
    Text1.Move 0, 0, UserControl.Width, tHEIGHT
End Sub
