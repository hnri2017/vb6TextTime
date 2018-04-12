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
Private Const INPUTEND As Integer = 2   'ÿ���������ֻ������2������


Dim intInputCount As Integer    '��¼���������2�κ����
Dim strOriginal As String       '����֮ǰ��ԭʼֵ���ֺ���ֻ����00~59֮�䣬ʱ��00~23֮��


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
                .SelLength = 2  'ѡ�� Сʱ ����������
            Case 3, 4, 5
                .SelStart = 3
                .SelLength = 2  'ѡ�� ���� ����������
            Case 6, 7, 8
                .SelStart = 6
                .SelLength = 2  'ѡ�� ���� ����������
        End Select
        strOriginal = .SelText  '����ѡ��ʱ��ֵ��������Ƿ�ֵʱ�ô�ֵ���
    End With
    
    intInputCount = 0   '��λ�������
    
End Sub

Private Sub Text1_GotFocus()
    
    '��Ҫ��ֹ�������ʱ���ı���ʱ�޷���λʱ���֡����е�һ��
    '��������ͬ�����¼�����
    
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
    Text1.Text = "00:00:00"     '��ʼ�� ����ʱ�� ���ı���ֵ
    Text1.Height = tHEIGHT      '������Ҫ�ĸ߶�����
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '��������,ֻ�����ĸ��������������ȡ������
    
    Dim strLastTwo As String    '��¼���¼�ǰ�� ֵ
    Dim strCurValue As String   '��ʾ���¼��ı��� ֵ
    
    If KeyCode > 36 And KeyCode < 41 Then intInputCount = 0 'Keycode��Χ���ƺ���Ҫ

    With Text1
        Select Case KeyCode
            Case vbKeyLeft      '����˳��ֻ�ܴ��뵽�֣��ӷֵ�ʱ����ʱ���ٰ���Ч
            
                Select Case .SelStart
                    Case 3, 4, 5    '����ڷ��ϣ�����ǰ����ʱ�ϡ���������ֻ3�Ϳ����ˣ�4��5�����⣬����ͬ
                        .SelStart = 0
                        .SelLength = 2
                    Case 6, 7, 8    '��������ϣ�����ǰ��������
                        .SelStart = 3
                        .SelLength = 2
                End Select
        
            Case vbKeyRight     '����˳��ֻ�ܴ�ʱ���֣��ӷֵ��룬������ٰ���Ч
            
                Select Case .SelStart
                    Case 0, 1, 2    '�����ʱ�ϣ��������������
                        .SelStart = 3
                        .SelLength = 2
                    Case 3, 4, 5    '����ڷ��ϣ��������������
                        .SelStart = 6
                        .SelLength = 2
                End Select
            
            Case vbKeyUp, vbKeyDown '����
            
                strLastTwo = .SelText
                Select Case KeyCode
                    Case vbKeyUp    '����ʾ������1���ϼ�1
                        If Val(strLastTwo) < 9 Then
                            strCurValue = "0" & CStr(Val(strLastTwo) + 1)   '����01~09
                        Else
                            strCurValue = CStr(Val(strLastTwo) + 1)
                            If .SelStart = 0 Then
                                If Val(strLastTwo) = 23 Then strCurValue = "00" '��ʾ��23 ���� 00
                            ElseIf .SelStart = 3 Or .SelStart = 6 Then
                                If Val(strLastTwo) = 59 Then strCurValue = "00" '��ʾ��59 ���� 00
                            End If
                        End If

                    Case vbKeyDown  '����ʾ������1���¼�1
                        If Val(strLastTwo) > 10 Then
                            strCurValue = CStr(Val(strLastTwo) - 1)
                        Else
                            strCurValue = "0" & CStr(Val(strLastTwo) - 1)   '����01~09
                            If .SelStart = 0 Then
                                If Val(strLastTwo) = 0 Then strCurValue = "23"  '��ʾ��00 ���� 23
                            ElseIf .SelStart = 3 Or .SelStart = 6 Then
                                If Val(strLastTwo) = 0 Then strCurValue = "59"  '��ʾ��00 ���� 59
                            End If
                        End If
                End Select
                
                .SelText = strCurValue  '��������Чֵ�滻ѡ�е�ֵ
                
                Select Case .SelStart   '�����¶�λ���Ա�λ������
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
    
    KeyCode = 0 'with�����д��������ֵ������Ӧȡ��ϵͳ���룬������������ֵ
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    'ֻ������0~9���10������
    
    Dim strFull As String   '��ʾ��������ֵ
    
    With Text1
        If KeyAscii > 47 And KeyAscii < 59 And intInputCount < INPUTEND And Len(.SelText) = INPUTEND Then

            intInputCount = intInputCount + 1
            
            If intInputCount = 1 Then
                strFull = "0" & Chr(KeyAscii)  '����01~09
            ElseIf intInputCount = 2 Then
                strFull = Right(.SelText, 1) & Chr(KeyAscii)
                If .SelStart = 0 Then   '��ǰ��λ��Сʱ��
                    If Val(strFull) > 23 Then strFull = strOriginal '����������ɵ�����Ϊ�Ƿ�ֵʱ�滻Ϊԭֵ
                Else    '��ǰ��λ�ڷ��ӻ�������
                    If Val(strFull) > 59 Then strFull = strOriginal

                End If
                
            End If
            
            .SelText = strFull  '������ֵ�滻ѡ��ֵ
            
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
        
        If intInputCount = INPUTEND Then intInputCount = 0  '�����������κ��������
        
    End With
    
    KeyAscii = 0    'with�����д��������ֵ������Ӧȡ��ϵͳ���룬������������ֵ
    
End Sub


Private Sub UserControl_Resize()
    UserControl.Height = tHEIGHT
    Text1.Move 0, 0, UserControl.Width, tHEIGHT
End Sub
