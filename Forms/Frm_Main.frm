VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "±ÈÀý¼ÆËã"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Btn_Start 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Æô¶¯"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   80.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Start_Click()
    mainF.Show
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    With Btn_Start
        .Left = (Me.Width - .Width) / 2
        .Top = (Me.Height - .Height) / 2
    End With
    Me.KeyPreview = True
    Me.DrawWidth = 2
    Me.AutoRedraw = True
    Me.Line (0, 0)-(0, Me.Height)
    Me.Line (0, 0)-(Me.Width, 0)
    Me.Line (0, Me.Height)-(Me.Width, Me.Height)
    Me.Line (Me.Width, 0)-(Me.Width, Me.Height)
    
    
End Sub
