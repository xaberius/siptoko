VERSION 5.00
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Begin VB.Form FrmGrafik 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   5000
      Left            =   1140
      ScaleHeight     =   5000
      ScaleMode       =   0  'User
      ScaleWidth      =   9945
      TabIndex        =   6
      Top             =   1560
      Width           =   10000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   8880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   8640
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   9120
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   10680
      TabIndex        =   24
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
      MICON           =   "FrmGrafik.frx":0000
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
      Left            =   11160
      TabIndex        =   25
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
      MICON           =   "FrmGrafik.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton4 
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Grosir"
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
      MICON           =   "FrmGrafik.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton5 
      Height          =   375
      Left            =   2880
      TabIndex        =   27
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Eceran"
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
      MICON           =   "FrmGrafik.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton6 
      Height          =   375
      Left            =   10680
      TabIndex        =   28
      ToolTipText     =   "Bulan Depan"
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   ">&>"
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
      MICON           =   "FrmGrafik.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton7 
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      ToolTipText     =   "Bulan Lalu"
      Top             =   840
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&<<"
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
      MICON           =   "FrmGrafik.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton8 
      Height          =   375
      Left            =   9600
      TabIndex        =   30
      ToolTipText     =   "Bulan Lalu"
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Bulan Ini"
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
      MICON           =   "FrmGrafik.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penjualan Grosir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7080
      TabIndex        =   31
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOP 5 Penjualan Barang Bulanan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   360
      TabIndex        =   23
      Top             =   120
      Width           =   5850
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   1140
      TabIndex        =   22
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "450"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   21
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   20
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "350"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   19
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "400"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   18
      Top             =   2400
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "250"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   17
      Top             =   3885
      Width           =   450
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   16
      Top             =   3405
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "150"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   15
      Top             =   4860
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   14
      Top             =   4365
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   720
      TabIndex        =   12
      Top             =   6240
      Width           =   150
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   480
      TabIndex        =   11
      Top             =   5880
      Width           =   300
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Motor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   3135
      TabIndex        =   10
      Top             =   6600
      Width           =   1905
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   5145
      TabIndex        =   9
      Top             =   6600
      Width           =   1830
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Pesawat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   7140
      TabIndex        =   8
      Top             =   6600
      Width           =   1875
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Kapal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1575
      Left            =   9135
      TabIndex        =   7
      Top             =   6600
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   7695
      Left            =   120
      Top             =   720
      Width           =   11400
   End
End
Attribute VB_Name = "FrmGrafik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Angka As Integer
Dim Edit As Boolean
Private Sub Form_Load()
Edit = True
Angka = 0
RefreshData1
Bulanan
End Sub


Private Sub Timer1_Timer()
For I = 1 To 2000
Picture1.Line (0 + I, Text2.Text)-(0 + I, Picture1.Height), vbRed
Picture1.Line (2000 + I, Text3.Text)-(2000 + I, Picture1.Height), vbGreen
Picture1.Line (4000 + I, Text4.Text)-(4000 + I, Picture1.Height), vbBlue
Picture1.Line (6000 + I, Text5.Text)-(6000 + I, Picture1.Height), vbYellow
Picture1.Line (8000 + I, Text6.Text)-(8000 + I, Picture1.Height), vbMagenta
Next I
Timer1.Enabled = False
End Sub

Sub RefreshData2()
Dim Hasil(4) As Integer
Dim Nama(4) As String

SQL = "SELECT     TOP 5 NamaBarang AS [Nama Barang], ISNULL ((SELECT     SUM(b.jumlah) " & _
        "FROM         dtljualgrosir b WHERE     a.kodebarang = b.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) AS [Penjualan Grosir], ISNULL ((SELECT     SUM(c.jumlah) " & _
        "FROM         dtljualecer c WHERE     a.kodebarang = c.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) AS [Penjualan Eceran] " & _
        "FROM         Barang a ORDER BY ISNULL ((SELECT     SUM(b.jumlah) " & _
        "FROM         dtljualgrosir b WHERE     a.kodebarang = b.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) DESC, ISNULL " & _
        "((SELECT     SUM(c.jumlah) FROM         dtljualecer c " & _
        "WHERE     a.kodebarang = c.kodebarang AND month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) DESC"
Set RSFind = DbCon.Execute(SQL)

a = 0
RSFind.MoveFirst
While Not RSFind.EOF
    Hasil(a) = Val(RSFind.Fields(2).Value)
    Nama(a) = Trim(RSFind.Fields(0).Value)
    a = a + 1
RSFind.MoveNext
Wend

Label12.Caption = Nama(0)
Label13.Caption = Nama(1)
Label14.Caption = Nama(2)
Label15.Caption = Nama(3)
Label16.Caption = Nama(4)

Picture1.Cls
Text2.Text = 5000 - Val(Hasil(0)) * 10
Text3.Text = 5000 - Val(Hasil(1)) * 10
Text4.Text = 5000 - Val(Hasil(2)) * 10
Text5.Text = 5000 - Val(Hasil(3)) * 10
Text6.Text = 5000 - Val(Hasil(4)) * 10
Timer1.Enabled = True
End Sub

Sub RefreshData1()
Dim Hasil(4) As Integer
Dim Nama(4) As String

SQL = "SELECT     TOP 5 NamaBarang AS [Nama Barang], ISNULL ((SELECT     SUM(b.jumlah) " & _
        "FROM         dtljualgrosir b WHERE     a.kodebarang = b.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) AS [Penjualan Grosir], ISNULL ((SELECT     SUM(c.jumlah) " & _
        "FROM         dtljualecer c WHERE     a.kodebarang = c.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) AS [Penjualan Eceran] " & _
        "FROM         Barang a ORDER BY ISNULL ((SELECT     SUM(b.jumlah) " & _
        "FROM         dtljualgrosir b WHERE     a.kodebarang = b.kodebarang AND " & _
        "month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) DESC, ISNULL " & _
        "((SELECT     SUM(c.jumlah) FROM         dtljualecer c " & _
        "WHERE     a.kodebarang = c.kodebarang AND month(tgltransaksi) = month(getdate())-'" & Angka & "'), 0) DESC"
Set RSFind = DbCon.Execute(SQL)

a = 0
RSFind.MoveFirst
While Not RSFind.EOF
    Hasil(a) = Val(RSFind.Fields(1).Value)
    Nama(a) = Trim(RSFind.Fields(0).Value)
    a = a + 1
RSFind.MoveNext
Wend

Label12.Caption = Nama(0)
Label13.Caption = Nama(1)
Label14.Caption = Nama(2)
Label15.Caption = Nama(3)
Label16.Caption = Nama(4)

Picture1.Cls
Text2.Text = 5000 - Val(Hasil(0)) * 10
Text3.Text = 5000 - Val(Hasil(1)) * 10
Text4.Text = 5000 - Val(Hasil(2)) * 10
Text5.Text = 5000 - Val(Hasil(3)) * 10
Text6.Text = 5000 - Val(Hasil(4)) * 10
Timer1.Enabled = True
End Sub

Sub Bulanan()
Select Case (Month(Date) - Angka)
    Case 1:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Januari"
    Case 2:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Febuari"
    Case 3:     Label17Caption = "TOP 5 Penjualan Barang Bulan Maret"
    Case 4:     Label17.Caption = "TOP 5 Penjualan Barang Bulan April"
    Case 5:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Mei"
    Case 6:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Juni"
    Case 7:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Juli"
    Case 8:     Label17.Caption = "TOP 5 Penjualan Barang Bulan Agustus"
    Case 9:     Label17.Caption = "TOP 5 Penjualan Barang Bulan September"
    Case 10:    Label17.Caption = "TOP 5 Penjualan Barang Bulan Oktober"
    Case 11:    Label17.Caption = "TOP 5 Penjualan Barang Bulan November"
    Case 12:    Label17.Caption = "TOP 5 Penjualan Barang Bulan Desember"
End Select
End Sub

Private Sub vbButton1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub

Private Sub vbButton4_Click()
Edit = True
Label18.Caption = "Penjualan Grosir"
RefreshData1
End Sub

Private Sub vbButton5_Click()
Edit = False
Label18.Caption = "Penjualan Eceran"
RefreshData2
End Sub
Private Sub vbButton6_Click()
If Angka <= 0 - (12 - Month(Date)) Then
    Exit Sub
Else
    Angka = Angka - 1
    Bulanan
    If Edit Then
        RefreshData1
    Else
        RefreshData2
    End If
End If

End Sub

Private Sub vbButton7_Click()
If Angka >= Month(Date) - 1 Then
    Exit Sub
Else
    Angka = Angka + 1
    Bulanan
    If Edit Then
        RefreshData1
    Else
        RefreshData2
    End If
End If
End Sub

Private Sub vbButton8_Click()
Angka = 0
Bulanan
    If Edit Then
        RefreshData1
    Else
        RefreshData2
    End If
End Sub
