VERSION 5.00
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form FrmPerubahanHarga 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   ScaleHeight     =   3930
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtKet 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox TxtHargaLama 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox TxtKodeBarang 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox TxtKwitansi 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "-"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPerubahanHarga.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton2 
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "X"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16744576
      BCOLO           =   16744576
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPerubahanHarga.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton CmdCancel 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmPerubahanHarga.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton CmdSave 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Save"
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
      MICON           =   "FrmPerubahanHarga.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber TxtHarga 
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
      _Version        =   393216
      _ExtentX        =   4895
      _ExtentY        =   556
      Calculator      =   "FrmPerubahanHarga.frx":0070
      Caption         =   "FrmPerubahanHarga.frx":0090
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmPerubahanHarga.frx":00FC
      Keys            =   "FrmPerubahanHarga.frx":011A
      Spin            =   "FrmPerubahanHarga.frx":0164
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Beli Baru"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Beli Lama"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   3255
      Left            =   120
      Top             =   600
      Width           =   5880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Perubahan Harga Beli"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No Kwitansi Pembelian"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "FrmPerubahanHarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
SQL = "Insert into perubahanHarga values('" & Trim(TxtKwitansi) & "','" & Trim(TxtKodeBarang) & _
    "','" & FormatTgl(Date) & "'," & Val(TxtHargaLama) & "," & TxtHarga & ",'" & Trim(TxtKet) & "')"
DbCon.Execute SQL

SQL = "Update barang set hargaBeli=" & TxtHarga & " where kodebarang='" & TxtKodeBarang & "'"
DbCon.Execute SQL

MsgBox "Data Saved"
Unload Me
End Sub

Private Sub Form_Load()
TxtKwitansi = Trim(FrmBeliBarang.TxtKwitansi)
TxtKodeBarang = Trim(FrmBeliBarang.KodesBarang)
End Sub

Private Sub TxtNama_GotFocus()
SQL = "Select namaBarang,HargaBeli from barang where kodeBarang='" & Trim(TxtKodeBarang) & "'"
Set RSFind = DbCon.Execute(SQL)
TxtNama = Trim(RSFind!namaBarang)
TxtHargaLama = Val(RSFind!HargaBeli)
End Sub

Private Sub vbButton1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub
