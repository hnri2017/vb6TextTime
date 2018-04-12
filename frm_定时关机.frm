VERSION 5.00
Begin VB.Form frm_定时关机 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "定时关机小程序"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "frm_定时关机.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4230
   StartUpPosition =   2  '屏幕中心
   Begin 工程1.TextTime Text2 
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
      Caption         =   "计时结束后关机"
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
      Caption         =   "最小化到托盘图标"
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
      Caption         =   "计时模式选择"
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "设置计时时间"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系统当前时间"
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "计时起始时间"
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "计时运行时间"
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "计时剩余时间"
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
         Name            =   "宋体"
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
         Caption         =   "显示窗口"
      End
      Begin VB.Menu mnubar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "重启计时器"
      End
      Begin VB.Menu mnubar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "frm_定时关机"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type NOTIFYICONDATA '下面申明的API函数Shell_NotifyIcon的参数
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
Const NIM_ADD = &H0     '添加图标
Const NIM_MODIFY = &H1  '修改图标
Const NIM_DELETE = &H2  '删除图标

Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDBLCLK = &H203  '按鼠标左键

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, _
                                              ByVal dwDuration As Long) As Long
Const STRCLOCK As String = "定点"
Const STRCOUNT As String = "倒计时"
Const STRSTART As String = "计时开始"
Const STRSTOP As String = "停止计时"
Const STROVER As String = "计时结束！"
Const STRRUN As String = "计时中"
Const INPUTEND As Integer = 2   '每次输入最多只能输入2个数字

Dim intInputCount As Integer    '记录输入次数，2次后归零
Dim strOriginal As String       '更改之前的原始值。分和秒只能在00~59之间，时在00~23之间
Dim strMode As String   '计时模式
Dim dtStart As Date     '起始时间
Dim dtEnd As Date       '结束时间
Dim blnNotLoad As Boolean   '不是窗体加载的时候
Dim lngBack As Long     'Shell调用自动关机程序返回的值
Dim strWindir As String '返回系统环境变量Windir的路径


'以下声明的类型、常量、API函数用于提取文件的图标
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
        If MsgBox("是否取消自动关机提示？", vbQuestion + vbYesNo, "取消关机提示") = vbYes Then
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
            MsgBox "定点时间必须大于当前时间！", vbExclamation, "定点时间设置异常"
            Exit Sub
        End If
    ElseIf strMode = STRCOUNT Then
        If Format(Text2.Time, "HH:mm:ss") = "00:00:00" Then
            MsgBox "倒计时时长必须大于0秒！", vbExclamation, "计时时间设置异常"
            Exit Sub
        End If
    End If

    If MsgBox("是否启动计时？", vbQuestion + vbYesNo, "启动询问") = vbNo Then Exit Sub
    
    Text1.Item(1).Text = CStr(Time) '显示起始时间
    dtStart = Time   '记录起始时间
    
    If strMode = STRCLOCK Then  '确定结束时间。两种模式下的结束时间获取方法不一样
        dtEnd = Text2.Time
    Else
        dtEnd = DateAdd("h", Text2.Hour, dtStart)   '在起始时间上分别加上设定结束时间的时、分、秒
        dtEnd = DateAdd("n", Text2.Minute, dtEnd)
        dtEnd = DateAdd("s", Text2.Second, dtEnd)
    End If
    
    Timer1.Enabled = True   '唤醒计时开始
    
    Text2.Enabled = False
    Command2.ZOrder         '显示结束按钮
    Label1.Item(6).Caption = STRRUN

End Sub

Private Sub Command2_Click()
    '
    If MsgBox("是否停止计时？", vbQuestion + vbYesNo, "停止询问") = vbNo Then Exit Sub
    
    Timer1.Enabled = False
    Text2.Enabled = True
    Label1.Item(6).Caption = STROVER
    Command1.ZOrder
    
    If lngBack > 0 Then '如果已触发自动关机，则取消自动关机
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
        Command2.Move .Left, .Top, .Width, .Height  '让开始与结束按钮重叠
    End With
    
    Timer1.Interval = 250
    Timer1.Enabled = False
    Timer2.Interval = 250
    Timer2.Enabled = True
    
    Label1.Item(6).Caption = ""
    
    With tNoti  '初始化自定义变量，用于显示托盘图标
        .cbSize = Len(tNoti)
        .uId = vbNull
        .hWnd = Me.hWnd
        .uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Me.Caption & Chr(0)
    End With
    
    blnNotLoad = True
    
    strWindir = Environ("windir")   '获取系统变量Windir路径
    
    Dim strIconPath As String
    strIconPath = strWindir & "\system32\taskschd.msc"
    If Len(Dir(strIconPath)) > 0 Then
'        Me.Icon = GetFileIconL(strIconPath)
'
'        Set stdPic = Controls.Add("vb.PictureBox", "stdPic")
'        stdPic.AutoSize = True
'        stdPic.Picture = GetFileIconL(strIconPath)
''''''        SavePicture stdPic.Picture, App.Path & "\Shutdown.ico" 'Picture属性保存出来的图片清晰度太差
'        SavePicture stdPic.Image, App.Path & "\Shutdown.ico"
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then
        Call PopupMenu(mnuIcon) '弹出菜单
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("是否退出计时程序？", vbQuestion + vbYesNo, "退出询问") = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        If blnNotLoad And Check1.Value Then '最小化时隐藏窗口并显示托盘图标
            Me.Hide
            Shell_NotifyIcon NIM_ADD, tNoti
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, tNoti  '删除托盘图标
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
    
    Text1.Item(2).Text = Format(Time - dtStart, "HH:mm:ss")  '显示运行时间
    Text1.Item(3).Text = Format(dtEnd - Time, "HH:mm:ss")    '显示剩余时间
    
    intCount = intCount + 1
    Label1.Item(6).Caption = STRRUN & String(intCount, ".")
    If intCount > 5 Then intCount = 0
    
    If Format(Text1.Item(3).Text, "HH:mm:ss") = "00:00:00" Then '计时结束判断
        Timer1.Enabled = False  '停止计时
        
        Command1.ZOrder
        Text2.Enabled = True
        Label1.Item(6).Caption = STROVER

        Call mnuShow_Click  '显示窗口
        Me.SetFocus
        
        If Check2.Value = 1 Then
            If Len(strWindir) > 0 Then
                lngBack = Shell(strWindir & "\System32\shutdown.exe -s -t 60", vbNormalFocus)
            Else
                MsgBox "系统变量路径获取失败，无法调用自动关机程序！", vbExclamation, "关机失败提醒"
            End If
        End If
        
    End If
    
End Sub

Private Sub Timer2_Timer()
    Text1.Item(0).Text = CStr(Time)
End Sub


'以下两个函数为提取文件图标。其中 GetFileIconS 提取 16×16 的图标，GetFileIconL 提取 32×32 的图标。可以是任何存在的文件
'用法:Me.Icon = GetFileIconL("C:\WIndows\System32\msvbvm60.dll")
Public Function GetFileIconS(ByVal sFileName As String) As StdPicture
    '提取 16×16 的图标，可以是任何存在的文件。
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
    '提取 32×32 的图标。可以是任何存在的文件。
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

