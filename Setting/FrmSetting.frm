VERSION 5.00
Begin VB.Form FrmSetting 
   Caption         =   "SQL Server Connector"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   Icon            =   "FrmSetting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2325
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtServerName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
   Begin VB.TextBox TxtLoginname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Tidak Boleh Kosong"
      Top             =   690
      Width           =   2145
   End
   Begin VB.TextBox TxtPassLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1140
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&SET"
      Height          =   465
      Left            =   2985
      TabIndex        =   3
      Top             =   1650
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2100
      Left            =   2205
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label LKoneksi 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   390
      TabIndex        =   7
      Top             =   1770
      Width           =   2610
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   300
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   5
      Top             =   780
      Width           =   1785
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password Ms. SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0079BCFF&
      BackStyle       =   1  'Opaque
      Height          =   2100
      Left            =   120
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileName    As String
Private Const MAX_PATH = 260
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim strBuffer   As String
Dim lngReturn   As Long
Dim x           As New Convert
Dim DbCon       As New ADODB.Connection
Dim strWindowsSystemDirectory As String
Private Sub Command1_Click()

LKoneksi.Caption = ">> Menghubungkan . . ."
LKoneksi.Refresh
Me.MousePointer = vbHourglass
If Trim(TxtLoginname) = "" Then
   MsgBox "User Name Tidak Boleh Kosong"
   TxtLoginname.SetFocus
   Exit Sub
End If
    On Error GoTo Err_Koneksi
With DbCon
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & TxtLoginname & ";password=" & TxtPassLogin & ";Initial Catalog=master;server=" & txtServerName
    .Open
End With
If Len(Dir(FileName)) > 0 Then
   Kill (FileName)
End If
SaveSetting App.Title, "Startup", "server", Trim(txtServerName)

Open FileName For Output As #1
For a = 1 To 2
    Select Case a
    Case 1
       tulis = x.encryp_pass(21, Trim(TxtLoginname))
       
    Case 2
       tulis = x.encryp_pass(21, Trim(TxtPassLogin))
    End Select
    
Print #1, tulis

Next
Close #1
Me.MousePointer = vbNormal
LKoneksi.Caption = ">> SQL Server telah dikoneksikan."
MsgBox "Setting Koneksi berhasil.", vbCritical, "Koneksi SQL Server berhasil"
Unload Me
Exit Sub
Err_Koneksi:
Me.MousePointer = vbNormal
LKoneksi.Caption = ">> Koneksi ke SQL Server Gagal."
If Err.Number = -2147467259 Then
   MsgBox "Nama Server tidak ditemukan." & vbCrLf & "Cek kembali nama Server !!!", vbCritical, "Koneksi Gagal"
   txtServerName.SetFocus
ElseIf Err.Number = -2147217843 Then
   MsgBox "User & Password tidak ditemukan." & vbCrLf & "Cek kembali User & Password !!!", vbCritical, "Koneksi Gagal"
   TxtLoginname.SetFocus
End If
End Sub

Private Sub Command2_Click()
 
End Sub

Private Sub Form_Load()
 Dim x As String
 txtServerName.ToolTipText = "Kosong = (local), " & _
                            "Anda Juga Bisa Memasukkan Nomor IP (ex. 192.168.0.171), " & _
                            "Nama Komputer (Ex. Server)"
 TxtLoginname.ToolTipText = "User Ms. SQL Server, Tidak Boleh Kosong!!!, Not Case Sensitive (ex. UserID=sa) "
 TxtPassLogin.ToolTipText = "Password Login Ke Ms. SQL Server 2000, Case Sensitive (ex. bas <> BAS) "
 App.Title = "SIP-Toko"
 strBuffer = Space$(MAX_PATH)
 lngReturn = GetSystemDirectory(strBuffer, MAX_PATH)
 strWindowsSystemDirectory = Left$(strBuffer, Len(strBuffer) - 1)
 x = Trim(strWindowsSystemDirectory)
 FileName = Left(x, Len(x) - 1) & "\" & "Agus.Said"
 txtServerName = GetSetting(App.Title, "Startup", "Server", "")
 'MsgBox FileName
End Sub

Private Sub TxtLoginname_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtLoginname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxtPassLogin_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub TxtPassLogin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub txtServerName_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub txtServerName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub


