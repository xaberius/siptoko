VERSION 5.00
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmKonsumen 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7305
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKode 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TxtNama 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox TxtAlamat 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox TxtKota 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3240
      TabIndex        =   8
      Top             =   6600
      Width           =   3255
   End
   Begin VB.TextBox TxtTelepon 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   9000
      Left            =   2040
      Top             =   4320
   End
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   7080
      TabIndex        =   10
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
      MICON           =   "FrmKonsumen.frx":0000
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
      Left            =   7560
      TabIndex        =   11
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
      MICON           =   "FrmKonsumen.frx":001C
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
      Left            =   6600
      TabIndex        =   6
      Top             =   1200
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
      MICON           =   "FrmKonsumen.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton CmdCari 
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   6600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "C&ari"
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
      MICON           =   "FrmKonsumen.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   3615
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Untuk Delete Data Tekan Tombol Delete Pada Keyboard"
      Top             =   2640
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   6376
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   16761024
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HFFC0C0&,.bold=0"
      _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1,.appearance=0,.borderColor=&H80000013&"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin BasTombol.vbButton CmdSave 
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   720
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
      MICON           =   "FrmKonsumen.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Data Konsumen"
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
      TabIndex        =   18
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   6615
      Left            =   120
      Top             =   600
      Width           =   7920
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Konsumen"
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
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Konsumen"
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
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat"
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
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Kota"
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
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Konsumen"
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
      TabIndex        =   13
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Telepon"
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
End
Attribute VB_Name = "FrmKonsumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Sub RefreshData()
Grid.DataSource = Nothing
SQL = "Select Kodekonsumen as [Kode Konsumen],namakonsumen as [Nama Konsumen], " & _
    "Alamat,Kota,Telepon from konsumen order by Kodekonsumen"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
End Sub

Private Sub CmdCancel_Click()
Form_Load
End Sub

Function CekIsi() As Boolean
If Trim(TxtNama) = "" Then
    MsgBox "Nama Masih Kosong"
    CekIsi = False
ElseIf Trim(TxtAlamat) = "" Then
    MsgBox "Alamat Masih Kosong"
    CekIsi = False
ElseIf Trim(TxtKota) = "" Then
    MsgBox "Kota Masih Kosong"
    CekIsi = False
ElseIf Trim(TxtTelepon) = "" Then
    MsgBox "Telepon Masih Kosong"
    CekIsi = False
Else
    CekIsi = True
End If
End Function

Private Sub CmdCari_Click()
Dim Kata() As String
Dim Kalimat As String
Dim Query As Variant
If Trim(TxtCari) = "" Then
    MsgBox "Tidak Ada Kriteria"
    TxtCari.SetFocus
    Exit Sub
End If

    Kata() = Split(Trim(TxtCari), " ")

    For Each Query In Kata()
        Kalimat = Kalimat + " or namakonsumen like '%" & Query & "%'"
    Next Query
    Kalimat = Mid(Kalimat, 4, Len(Kalimat) - 3)

SQL = "select Kodekonsumen as [Kode konsumen],NamaKonsumen as [Nama Konsumen],Alamat,Kota,Telepon " & _
    " from konsumen where " & Kalimat & _
    " order by Kodekonsumen"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
TxtCari = ""
Timer1.Enabled = True
End Sub

Private Sub CmdSave_Click()

If Not CekIsi Then Exit Sub
If Edit Then
    SQL = "update konsumen set namakonsumen='" & Trim(TxtNama) & "',alamat='" & _
    Trim(TxtAlamat.Text) & "',kota='" & Trim(TxtKota.Text) & "',telepon='" & _
    Trim(TxtTelepon.Text) & "' where kodekonsumen='" & Trim(TxtKode.Text) & "'"
    DbCon.Execute SQL
    MsgBox "Data Saved"
    Form_Load
ElseIf Not Edit Then
    TxtKode = KodeAuto
    SQL = "Insert Into konsumen values ('" & Trim(TxtKode.Text) & "','" & Trim(TxtNama.Text) & _
    "','" & Trim(TxtAlamat.Text) & "','" & Trim(TxtKota.Text) & "','" & _
    Trim(TxtTelepon.Text) & "')"
    DbCon.Execute SQL
    MsgBox "Data Saved"
    Form_Load
    FrmJualGrosir.CmbKonsumen.Refresh
End If
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub Form_Load()
TxtKode = KodeAuto
TxtKode.Locked = True
Bersih
RefreshData
Edit = False
Timer1.Enabled = False

End Sub

Sub Bersih()
TxtNama = ""
TxtAlamat = ""
TxtKota = ""
TxtTelepon = ""
End Sub


Private Sub Grid_DblClick()

Edit = True
TxtKode = Trim(Grid.Columns(0).Text)
TxtNama = Trim(Grid.Columns(1).Text)
TxtAlamat = Trim(Grid.Columns(2).Text)
TxtKota = Trim(Grid.Columns(3).Text)
TxtTelepon = Trim(Grid.Columns(4).Text)
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("konsumen " & Trim(Grid.Columns(1).Text) & " Yakin Akan Dihapus", vbYesNo + vbCritical) = vbYes Then
        SQL = "delete from konsumen where kodekonsumen='" & Trim(Grid.Columns(0).Text) & "'"
        DbCon.Execute SQL
        MsgBox "Data Deleted"
        Form_Load
    End If
End If
End Sub

Private Sub Timer1_Timer()
RefreshData
End Sub

Private Sub TxtAlamat_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtAlamat_KeyPress(KeyAscii As Integer)
KeyAscii = CekHuruf(KeyAscii)
End Sub

Private Sub TxtKode_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtKota_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtKota_KeyPress(KeyAscii As Integer)
KeyAscii = CekHuruf(KeyAscii)
End Sub

Private Sub TxtNama_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
KeyAscii = CekHuruf(KeyAscii)
End Sub

Private Sub TxtTelepon_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub

Function KodeAuto()
'SQL = "Select No_Urut from ServiceMobil order by No_Urut"
'Set RSFind = DbCon.Execute(SQL)
'If Not RSFind.BOF Then
'   KodeAuto = RSFind!no_urut
'   Exit Function
'End If
SQL = "Select Kodekonsumen from konsumen order by Kodekonsumen Desc"
Set RSFind = DbCon.Execute(SQL)
If RSFind.BOF Then
    KodeAuto = "0001"
Else
    KodeAuto = Format(CInt(Left(RSFind!Kodekonsumen, 4)) + 1, "0000")
End If
End Function


