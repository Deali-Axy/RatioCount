VERSION 5.00
Begin VB.Form mainF 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FBFBF7&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   4395
   ClientTop       =   645
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Head 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   15
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   910
      TabIndex        =   0
      Top             =   15
      Width           =   13650
      Begin VB.Label Lbl_Title 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "比例转换计算 By DealiAxy"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3405
      End
   End
   Begin VB.Timer TimMouseM 
      Interval        =   5
      Left            =   13725
      Top             =   990
   End
   Begin VB.Timer TimLoad 
      Interval        =   10
      Left            =   14580
      Top             =   540
   End
   Begin VB.PictureBox Start 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7620
      Left            =   45
      ScaleHeight     =   7620
      ScaleWidth      =   13470
      TabIndex        =   1
      Top             =   600
      Width           =   13470
      Begin VB.TextBox Txt_Res2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Txt_Res1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Txt_Ratio2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   3240
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Txt_Ratio1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "得先输入比例才能进行转换计算！"
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   2400
         Width           =   4275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1320
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "实时转换："
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1320
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入比例："
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   1425
      End
   End
End
Attribute VB_Name = "mainF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FormG As Long, BKpen As Long    '画边框用
Attribute BKpen.VB_VarUserMemId = 1073938432
Dim HeadG As Long, Penx As Long, brushx As Long, brushMin As Long    '画标题栏用
Attribute HeadG.VB_VarUserMemId = 1073938434
Attribute Penx.VB_VarUserMemId = 1073938434
Attribute brushx.VB_VarUserMemId = 1073938434
Attribute brushMin.VB_VarUserMemId = 1073938434
Dim MyCur As POINTAPI    '用于获取鼠标位置判断是否画阴影
Attribute MyCur.VB_VarUserMemId = 1073938438
Dim Color As Integer    '随机一个标题栏颜色
Attribute Color.VB_VarUserMemId = 1073938439


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'用于获取鼠标位置判断是否画阴影
Private Type POINTAPI
    x As Long
    y As Long
End Type
'用于获取鼠标位置判断是否画阴影
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'这个暂时没有用处

Private Function GetRndColor() As Double
    Dim r, g, b As Integer
    r = Rnd * 255
    g = Rnd * 255
    b = Rnd * 255
    GetRndColor = RGB(r, g, b)
End Function

Private Sub Form_Load()
    Randomize
    Color = Rnd * 6    '随机颜色

    Head.Width = Me.Width \ 15 - 2    '调整标题栏

    Start.Left = 2
    Start.Top = Head.Height + 2
    Start.Height = Me.Height \ 15 - Head.Height - 4
    Start.Width = Me.Width \ 15 - 4

    InitGDIPlus
    '创建一大堆要用 不同的画板画笔
    GdipCreateFromHDC Me.hDC, FormG
    GdipCreatePen1 &HFF808080, 1, UnitPixel, BKpen
    GdipCreateFromHDC Head.hDC, HeadG
    'GdipSetSmoothingMode HeadG, SmoothingModeAntiAlias
    GdipCreatePen1 &HFFFFFFFF, 2, UnitPixel, Penx
    GdipCreateSolidFill &HDDDD0000, brushx
    GdipCreateSolidFill &H90808080, brushMin

    '画标题栏
    DrawBk FormG, BKpen, Me
    DrawX 0, HeadG, Penx, brushx, brushMin, Me.Head

End Sub

Private Sub Head_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'操作层

    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    '移动

    If x > Head.Width - 35 And x < Head.Width And y > 0 And y < 35 Then    '关闭
        Form_Unload 0
        End
    ElseIf x > Head.Width - 75 And x < Head.Width - 35 And y > 0 And y < 35 Then    '最小化
        Me.WindowState = vbMinimized
    End If
End Sub

Private Sub DrawX(MouseState As Integer, g As Long, P As Long, RB As Long, BB As Long, Obj As Object)
'绘制层 重画标题栏及按钮
    Select Case Color
    Case Is = 1
        GdipGraphicsClear g, &HFFA8D59D
    Case Is = 2
        GdipGraphicsClear g, &HFFAF88B8
    Case Is = 3
        GdipGraphicsClear g, &HFFF49E9C
    Case Is = 4
        GdipGraphicsClear g, &HFFFACD8A
    Case Is = 5
        GdipGraphicsClear g, &HFF808080
    Case Is = 6
        GdipGraphicsClear g, &HFF8CCCCA
    End Select

    '鼠标停留时的按钮阴影
    Select Case MouseState
    Case Is = 1    '指在关闭键上
        GdipFillRectangleI g, RB, Obj.Width - 33, -1, 35, 32
    Case Is = 2    '指在缩小键上
        GdipFillRectangleI g, BB, Obj.Width - 70, -1, 35, 32
    End Select
    GdipDrawLineI g, P, Obj.Width - 27, 7, Obj.Width - 8, 25
    GdipDrawLineI g, P, Obj.Width - 27, 25, Obj.Width - 8, 7

    GdipDrawLineI g, P, Obj.Width - 65, 22, Obj.Width - 40, 22
    Obj.Refresh
End Sub

Private Sub DrawBk(g As Long, pen As Long, Obj As Object)
'画边框
    GdipDrawRectangleI g, pen, 0, 0, Obj.Width \ 15 - 1, Obj.Height \ 15 - 1
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'删除画板画笔
    GdipDeleteGraphics FormG
    GdipDeletePen BKpen
    GdipDeletePen Penx
    GdipDeleteBrush brushx
    GdipDeleteBrush brushMin
    GdipDeleteGraphics HeadG
    TerminateGDIPlus
End Sub

Private Sub Lbl_Title_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    '移动
End Sub

Private Sub Start_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    '移动
End Sub

Private Sub TimMouseM_Timer()
'判断鼠标位置
    Dim x As Long, y As Long
    GetCursorPos MyCur
    x = MyCur.x - Me.Left \ 15
    y = MyCur.y - Me.Top \ 15
    If x > Head.Width - 35 And x < Head.Width And y > 0 And y < 35 Then
        DrawX 1, HeadG, Penx, brushx, brushMin, Me.Head
    ElseIf x > Head.Width - 75 And x < Head.Width - 35 And y > 0 And y < 35 Then
        DrawX 2, HeadG, Penx, brushx, brushMin, Me.Head
    Else
        DrawX 0, HeadG, Penx, brushx, brushMin, Me.Head
    End If
End Sub

Private Sub Txt_Ratio1_Change()
    If Txt_Ratio1.Text = "0" Then
        MsgBox "The ratio can't be zero!"
        Txt_Ratio1.Text = ""
    End If
    If Len(Txt_Ratio1.Text) > 0 And Len(Txt_Ratio2.Text) > 0 Then
        Txt_Res1.Enabled = True
        Txt_Res2.Enabled = True
        If Val(Txt_Res1.Text) > 0 Then
            Txt_Res2.Text = Val(Txt_Res1.Text) * Val(Txt_Ratio2.Text) / Val(Txt_Ratio1.Text)
        End If
    Else
        Txt_Res1.Enabled = False
        Txt_Res2.Enabled = False
    End If
End Sub

Private Sub Txt_Ratio2_Change()
    If Txt_Ratio2.Text = "0" Then
        MsgBox "The ratio can't be zero!"
        Txt_Ratio2.Text = ""
    End If
    If Len(Txt_Ratio1.Text) > 0 And Len(Txt_Ratio2.Text) > 0 Then
        Txt_Res1.Enabled = True
        Txt_Res2.Enabled = True
        If Val(Txt_Res1.Text) > 0 Then
            Txt_Res2.Text = Val(Txt_Res1.Text) * Val(Txt_Ratio2.Text) / Val(Txt_Ratio1.Text)
        End If
    Else
        Txt_Res1.Enabled = False
        Txt_Res2.Enabled = False
    End If
End Sub

Private Sub Txt_Res1_Change()
    Txt_Res2.Text = Val(Txt_Res1.Text) * Val(Txt_Ratio2.Text) / Val(Txt_Ratio1.Text)
End Sub

Private Sub Txt_Res2_Change()
    Txt_Res1.Text = Val(Txt_Res2.Text) * Val(Txt_Ratio1.Text) / Val(Txt_Ratio2.Text)
End Sub
