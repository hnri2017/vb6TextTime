VERSION 5.00
Begin VB.Form frm_��ʱ�ػ� 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʱ�ػ�С����"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "frm_��ʱ�ػ�.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4230
   StartUpPosition =   2  '��Ļ����
   Begin ����1.TextTime Text2 
      Height          =   300
      Left            =   1560
      TabIndex        =   16
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "��ʱ������ػ�"
      Height          =   180
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Left            =   1800
      Top             =   3720
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "��С��������ͼ��"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   270
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   270
      Index           =   1
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1320
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H008080FF&
      Height          =   270
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʱģʽѡ��"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ü�ʱʱ��"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ϵͳ��ǰʱ��"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʱ��ʼʱ��"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʱ����ʱ��"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ʱʣ��ʱ��"
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   9
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "mnuIcon"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "��ʾ����"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "������ʱ��"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "frm_��ʱ�ػ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type NOTIFYICONDATA '����������API����Shell_NotifyIcon�Ĳ���
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Dim tNoti As NOTIFYICONDATA

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, _
    lpData As NOTIFYICONDATA) As Long
Const NIM_ADD = &H0     '���ͼ��
Const NIM_MODIFY = &H1  '�޸�ͼ��
Const NIM_DELETE = &H2  'ɾ��ͼ��

Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203  '��������

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, _
                                              ByVal dwDuration As Long) As Long
Const STRCLOCK As String = "����"
Const STRCOUNT As String = "����ʱ"
Const STRSTART As String = "��ʱ��ʼ"
Const STRSTOP As String = "ֹͣ��ʱ"
Const STROVER As String = "��ʱ������"
Const STRRUN As String = "��ʱ��"
Const INPUTEND As Integer = 2   'ÿ���������ֻ������2������

Dim intInputCount As Integer    '��¼���������2�κ����
Dim strOriginal As String       '����֮ǰ��ԭʼֵ���ֺ���ֻ����00~59֮�䣬ʱ��00~23֮��
Dim strMode As String   '��ʱģʽ
Dim dtStart As Date     '��ʼʱ��
Dim dtEnd As Date       '����ʱ��
Dim blnNotLoad As Boolean   '���Ǵ�����ص�ʱ��
Dim lngBack As Long     'Shell�����Զ��ػ����򷵻ص�ֵ
Dim strWindir As String '����ϵͳ��������Windir��·��


'�������������͡�������API����������ȡ�ļ���ͼ��
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Const MAX_PATH = 1000
Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long          '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80         '  out: type name
End Type

Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_ICON = &H100

Private Type CLSID
    id((123)) As Byte
End Type

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long

Private WithEvents stdPic As PictureBox
Attribute stdPic.VB_VarHelpID = -1

Private Sub Check2_Click()
    If lngBack = 0 Then Exit Sub
    If Check2.Value = 0 Then
        If MsgBox("�Ƿ�ȡ���Զ��ػ���ʾ��", vbQuestion + vbYesNo, "ȡ���ػ���ʾ") = vbYes Then
            If Len(strWindir) > 0 Then
                Shell strWindir & "\System32\shutdown.exe -a", vbNormalFocus
            End If
            lngBack = 0
        End If
    End If
End Sub

Private Sub Combo1_Click()
    strMode = Combo1.Text
End Sub

Private Sub Command1_Click()
'Debug.Print text2.text, Now
    If strMode = STRCLOCK Then
        If Text2.Time < Time Then
            MsgBox "����ʱ�������ڵ�ǰʱ�䣡", vbExclamation, "����ʱ�������쳣"
            Exit Sub
        End If
    ElseIf strMode = STRCOUNT Then
        If Format(Text2.Time, "HH:mm:ss") = "00:00:00" Then
            MsgBox "����ʱʱ���������0�룡", vbExclamation, "��ʱʱ�������쳣"
            Exit Sub
        End If
    End If

    If MsgBox("�Ƿ�������ʱ��", vbQuestion + vbYesNo, "����ѯ��") = vbNo Then Exit Sub
    
    Text1.Item(1).Text = CStr(Time) '��ʾ��ʼʱ��
    dtStart = Time   '��¼��ʼʱ��
    
    If strMode = STRCLOCK Then  'ȷ������ʱ�䡣����ģʽ�µĽ���ʱ���ȡ������һ��
        dtEnd = Text2.Time
    Else
        dtEnd = DateAdd("h", Text2.Hour, dtStart)   '����ʼʱ���Ϸֱ�����趨����ʱ���ʱ���֡���
        dtEnd = DateAdd("n", Text2.Minute, dtEnd)
        dtEnd = DateAdd("s", Text2.Second, dtEnd)
    End If
    
    Timer1.Enabled = True   '���Ѽ�ʱ��ʼ
    
    Text2.Enabled = False
    Command2.ZOrder         '��ʾ������ť
    Label1.Item(6).Caption = STRRUN

End Sub

Private Sub Command2_Click()
    '
    If MsgBox("�Ƿ�ֹͣ��ʱ��", vbQuestion + vbYesNo, "ֹͣѯ��") = vbNo Then Exit Sub
    
    Timer1.Enabled = False
    Text2.Enabled = True
    Label1.Item(6).Caption = STROVER
    Command1.ZOrder
    
    If lngBack > 0 Then '����Ѵ����Զ��ػ�����ȡ���Զ��ػ�
        If Len(strWindir) > 0 Then
            Shell strWindir & "\System32\shutdown.exe -a", vbNormalFocus
        End If
        lngBack = 0
    End If
    
End Sub

Private Sub form_load()
    
    With Combo1
        .AddItem STRCLOCK
        .AddItem STRCOUNT
        .ListIndex = 1
    End With
    
    With Command1
        .Caption = STRSTART
        .ZOrder
        Command2.Caption = STRSTOP
        Command2.Move .Left, .Top, .Width, .Height  '�ÿ�ʼ�������ť�ص�
    End With
    
    Timer1.Interval = 250
    Timer1.Enabled = False
    Timer2.Interval = 250
    Timer2.Enabled = True
    
    Label1.Item(6).Caption = ""
    
    With tNoti  '��ʼ���Զ��������������ʾ����ͼ��
        .cbSize = Len(tNoti)
        .uId = vbNull
        .hWnd = Me.hWnd
        .uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Me.Caption & Chr(0)
    End With
    
    blnNotLoad = True
    
    strWindir = Environ("windir")   '��ȡϵͳ����Windir·��
    
    Dim strIconPath As String
    strIconPath = strWindir & "\system32\taskschd.msc"
    If Len(Dir(strIconPath)) > 0 Then
'        Me.Icon = GetFileIconL(strIconPath)
'
'        Set stdPic = Controls.Add("vb.PictureBox", "stdPic")
'        stdPic.AutoSize = True
'        stdPic.Picture = GetFileIconL(strIconPath)
''''''        SavePicture stdPic.Picture, App.Path & "\Shutdown.ico" 'Picture���Ա��������ͼƬ������̫��
'        SavePicture stdPic.Image, App.Path & "\Shutdown.ico"
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        Call PopupMenu(mnuIcon) '�����˵�
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("�Ƿ��˳���ʱ����", vbQuestion + vbYesNo, "�˳�ѯ��") = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        If blnNotLoad And Check1.Value Then '��С��ʱ���ش��ڲ���ʾ����ͼ��
            Me.Hide
            Shell_NotifyIcon NIM_ADD, tNoti
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, tNoti  'ɾ������ͼ��
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuShow_Click()
    Me.WindowState = vbNormal
    Me.Show
    Call Shell_NotifyIcon(NIM_DELETE, tNoti)
End Sub

Private Sub mnuStart_Click()
    If Label1.Item(6).Caption = STROVER Then
        Command1.Value = True
    End If
End Sub


Private Sub Timer1_Timer()
    '
    Static intCount As Integer
    
    Text1.Item(2).Text = Format(Time - dtStart, "HH:mm:ss")  '��ʾ����ʱ��
    Text1.Item(3).Text = Format(dtEnd - Time, "HH:mm:ss")    '��ʾʣ��ʱ��
    
    intCount = intCount + 1
    Label1.Item(6).Caption = STRRUN & String(intCount, ".")
    If intCount > 5 Then intCount = 0
    
    If Format(Text1.Item(3).Text, "HH:mm:ss") = "00:00:00" Then '��ʱ�����ж�
        Timer1.Enabled = False  'ֹͣ��ʱ
        
        Command1.ZOrder
        Text2.Enabled = True
        Label1.Item(6).Caption = STROVER

        Call mnuShow_Click  '��ʾ����
        Me.SetFocus
        
        If Check2.Value = 1 Then
            If Len(strWindir) > 0 Then
                lngBack = Shell(strWindir & "\System32\shutdown.exe -s -t 60", vbNormalFocus)
            Else
                MsgBox "ϵͳ����·����ȡʧ�ܣ��޷������Զ��ػ�����", vbExclamation, "�ػ�ʧ������"
            End If
        End If
        
    End If
    
End Sub

Private Sub Timer2_Timer()
    Text1.Item(0).Text = CStr(Time)
End Sub


'������������Ϊ��ȡ�ļ�ͼ�ꡣ���� GetFileIconS ��ȡ 16��16 ��ͼ�꣬GetFileIconL ��ȡ 32��32 ��ͼ�ꡣ�������κδ��ڵ��ļ�
'�÷�:Me.Icon = GetFileIconL("C:\WIndows\System32\msvbvm60.dll")
Public Function GetFileIconS(ByVal sFileName As String) As StdPicture
    '��ȡ 16��16 ��ͼ�꣬�������κδ��ڵ��ļ���
    Dim SHinfo As SHFILEINFO
    Dim mTYPEICON As TypeIcon
    Dim mCLSID As CLSID
    Dim lFlag As Long
    
    lFlag = SHGFI_SMALLICON
    If Right(sFileName, 1) <> "\" Then sFileName = sFileName & "\"
    Call SHGetFileInfo(sFileName, 0, SHinfo, Len(SHinfo), SHGFI_ICON + lFlag)
    With mTYPEICON
        .cbSize = Len(mTYPEICON)
        .picType = vbPicTypeIcon
        .hIcon = SHinfo.hIcon
    End With
    With mCLSID
        .id(8) = &HC0
        .id(15) = &H46
    End With
    Call OleCreatePictureIndirect(mTYPEICON, mCLSID, 1, GetFileIconS)
End Function

Public Function GetFileIconL(ByVal sFileName As String) As StdPicture
    '��ȡ 32��32 ��ͼ�ꡣ�������κδ��ڵ��ļ���
    Dim SHinfo As SHFILEINFO
    Dim mTYPEICON As TypeIcon
    Dim mCLSID As CLSID
    Dim lFlag As Long
    
    lFlag = SHGFI_LARGEICON
    If Right(sFileName, 1) <> "\" Then sFileName = sFileName & "\"
    Call SHGetFileInfo(sFileName, 0, SHinfo, Len(SHinfo), SHGFI_ICON + lFlag)
    With mTYPEICON
        .cbSize = Len(mTYPEICON)
        .picType = vbPicTypeIcon
        .hIcon = SHinfo.hIcon
    End With
    With mCLSID
        .id(8) = &HC0
        .id(15) = &H46
    End With
    Call OleCreatePictureIndirect(mTYPEICON, mCLSID, 1, GetFileIconL)
    
End Function

