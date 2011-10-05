VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Begin VB.Form FrmLoginServer 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Login"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLoginServer.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox TxtServer 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   120
   End
   Begin BasTombol.vbButton vbButton2 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   255
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLoginServer.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "ex: 192.168.1.141"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User Server"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "ex: 192.168.1.141"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Server"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Kosong = (Local) atau IP(192.168.1.41)"
      Top             =   960
      Width           =   975
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   2445
      Left            =   120
      Top             =   600
      Width           =   5295
      _ExtentX        =   9313
      _ExtentY        =   4313
      Image           =   "FrmLoginServer.frx":0038
      Opacity         =   50
      Props           =   5
   End
End
Attribute VB_Name = "FrmLoginServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
TxtServer.Text = ""
TxtUser.Text = ""
TxtPass.Text = ""
End Sub

Private Sub TxtPass_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtServer_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtUser_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton1_Click()
Dim DbName As String
App.Title = "SIP-Toko"
DbName = "SIP-Toko"
SaveSetting App.Title, "Startup", "Server", TxtServer.Text
'if not exists(select * from dbo.sysdatabases where name = 'Testing')create database testing
If Trim(TxtUser.Text) = "" Then
    MsgBox "User Masih Kosong"
    TxtServer.SetFocus
    Exit Sub
ElseIf Trim(TxtPass.Text) = "" Then
    MsgBox "Password Masih Kosong"
    TxtPass.SetFocus
    Exit Sub
End If

If Get_Connection("SIP-Toko", TxtUser, TxtPass) Then
    MsgBox "Database Conneted"
    Connect1 = True
    Unload Me
    CekTabel
    FrmLogin.Show , FrmUtama
End If


Exit Sub
End Sub

Private Sub vbButton1_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub

Private Sub vbButton2_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub
