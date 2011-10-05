VERSION 5.00
Object = "{9CAA1C67-43C4-4FFF-A005-20037C74BF32}#1.0#0"; "AlphaImageControl.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Begin VB.Form FrmLogin 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BasTombol.vbButton CmdLogin 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmLogin.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox TxtPass 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage2 
      Height          =   615
      Left            =   6360
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Image           =   "FrmLogin.frx":001C
      Props           =   5
   End
   Begin VB.Label Label3 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Login Aplikasi"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin AlphaImageControl.aicAlphaImage aicAlphaImage1 
      Height          =   4800
      Left            =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   8467
      Image           =   "FrmLogin.frx":6229
      Props           =   5
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AeroButton1_Click()
Unload Me
End Sub

Private Sub aicAlphaImage1_Click(ByVal Button As Integer)
Unload Me
End Sub

Private Sub aicAlphaImage2_Click(ByVal Button As Integer)
Unload Me
End Sub

Private Sub CmdLogin_Click()
Dim Itung As Integer
If Trim(TxtUser) = "" Then
    MsgBox "User Name Masing Kosong"
    TxtUser.SetFocus
    Exit Sub
ElseIf Trim(TxtPass) = "" Then
    MsgBox "Password Masih Kosong"
    TxtPass.SetFocus
    Exit Sub
End If

    If Itung = 3 Then
        MsgBox "Anda Tidak Berhak Memakai Aplikasi Ini"
        Unload Me
        Unload FrmUtama
    End If
        
    SQL = "select * from UserX where UserID='" & Trim(TxtUser) & "'"
    Set RSFind = DbCon.Execute(SQL)

    If Trim(TxtPass) = Trim(Trans.decryp_pass(25, RSFind!Password)) Then
        MsgBox "Aplikasi Siap Berjalan"
        Connect2 = True
        SaveSetting App.Title, "startup", "login", Trim(TxtUser)
        User.UserId = Trim(TxtUser)
        Unload Me
    Else
        MsgBox "User ID atau Password Kurang Tepat"
        Itung = Itung + 1
        TxtUser.SetFocus
    End If

End Sub

Private Sub CmdLogin_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub Form_Load()
TxtUser = GetSetting(App.Title, "startup", "login")
TxtUser.SelLength = Len(TxtUser)
TxtPass = ""
End Sub

Private Sub TxtPass_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtUser_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub
