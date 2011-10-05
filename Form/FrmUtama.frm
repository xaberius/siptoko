VERSION 5.00
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "SmartMenuXP.ocx"
Begin VB.Form FrmUtama 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   Caption         =   "Form1 "
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   Icon            =   "FrmUtama.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP1 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   16761024
      BorderStyle     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Data Barang"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "User : None"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   2280
      Left            =   1800
      Picture         =   "FrmUtama.frx":038A
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3840
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Aa As Integer

Private Sub Command1_Click()
Aa = Aa + 1
SmartMenuXP1.BorderStyle = mxpBump
End Sub

Private Sub Form_Activate()
Dim Status1 As String

If Connect1 Then
    Status1 = "Connected"
Else: Status1 = "Disconnect"
End If

Me.BackColor = &HDCB291
Image1.Top = Me.Top + 700
Image1.Left = Me.Left + 100
Image1.Width = Me.Width - 200
Image1.Height = Me.Height - 900
Label3.Caption = "User : " & User.UserId
Label1.Caption = "Database : " & Status1
Label3.Left = Me.Left + Me.Width - 3000
Label1.Left = Me.Left + Me.Width - 3000
End Sub
Private Function getIcon(ByVal iconName As String) As StdPicture
    Set getIcon = LoadPicture(App.Path + "\Icons\" + iconName + ".ico")
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If MsgBox("Akan Keluar Program?", vbYesNo + vbCritical, "Quit") = vbYes Then
        Unload FrmUtama
    End If
End If
End Sub

Private Sub Form_Load()
Connect1 = False
Connect2 = False
Aa = 1
With SmartMenuXP1.MenuItems
        .Add 0, "mnuServer", , "&Login   "
        .Add "mnuServer", "mnuLogin", , "Log&in", getIcon("Login")
        .Add "mnuServer", "mnuLogout", , "&Log&out", getIcon("Exit")
        .Add "mnuServer", , smiSeparator
        .Add "mnuServer", "mnuExit", , "&Exit"

        'TODO : DEFINISI MENU YANG LAIN

        .Add 0, "mnuData", , "&Data   "
        .Add "mnuData", "mnuSupplier", , "&Supplier", getIcon("Supplier")
        .Add "mnuData", "mnuKonsumen", , "&Konsumen", getIcon("Konsumen")
        .Add "mnuData", "mnuSales", , "Sa&les", getIcon("Sales")
        .Add "mnuData", , smiSeparator
        .Add "mnuData", "mnuJenis", , "&Jenis"
        .Add "mnuData", , smiSeparator
        .Add "mnuData", "mnuBarang", , "&Barang", getIcon("Barang")
        .Add "mnuData", , smiSeparator
        .Add "mnuData", "mnuMutasi", , "&Mutasi Barang", getIcon("Mutasi")
        
        .Add 0, "mnuTransaksi", , "&Transaksi   "
        .Add "mnuTransaksi", "mnuGrosir", , "Penjualan &Grosir", getIcon("Grosir")
        .Add "mnuTransaksi", "mnuEceran", , "Penjualan &Eceran", getIcon("Eceran")
        .Add "mnuTransaksi", "mnuReturJual", , "&Retur Penjualan", getIcon("Retur Jual")
        .Add "mnuTransaksi", , smiSeparator
        .Add "mnuTransaksi", "mnuPembelian", , "&Pembelian", getIcon("Pembelian")
        .Add "mnuTransaksi", "mnuReturBeli", , "Retur Pem&belian", getIcon("Retur Beli")
        
        .Add 0, "mnuAbout", , "&About   ", getIcon("Me")
    End With
    
End Sub


Private Sub SmartMenuXP1_Click(ByVal ID As Long)
With SmartMenuXP1.MenuItems
        Select Case .Key(ID)
            Case "mnuLogin":
                If Not Connect1 Then
                    With FrmLoginServer
                        .Show , FrmUtama
                    End With
                Else
                    FrmLogin.Show , FrmUtama
                End If
            Case "mnuLogout": Connect2 = False
            Case "mnuClose": Unload Me
            Case "mnuBarang": CekKonek FrmDataBarang
            Case "mnuAbout": FrmAbout.Show , FrmUtama
            Case "mnuJenis": CekKonek FrmJenisBarang
            Case "mnuSupplier": CekKonek FrmSupplier
            Case "mnuKonsumen": CekKonek FrmKonsumen
            Case "mnuSales": CekKonek FrmSales
            Case "mnuMutasi": CekKonek FrmMutasi
            Case "mnuPembelian": CekKonek FrmBeliBarang
            Case "mnuMicrosoftPowerPoint": 'TODO : something here
            Case "mnuExit": End
        End Select
    End With
End Sub
