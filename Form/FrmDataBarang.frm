VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmDataBarang 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   14055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2040
      Top             =   3240
   End
   Begin VB.TextBox TxtCari 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   7680
      TabIndex        =   3
      Top             =   7440
      Width           =   4455
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
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
      AllowRowSizing  =   0   'False
      Appearance      =   0
      BorderStyle     =   1
      ColumnHeaders   =   0   'False
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   ""
      DataField       =   ""
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   16761024
      ValueTranslate  =   0   'False
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
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
      Left            =   360
      TabIndex        =   1
      Top             =   7440
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
      MICON           =   "FrmDataBarang.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   12960
      TabIndex        =   5
      Top             =   240
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
      MICON           =   "FrmDataBarang.frx":001C
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
      Left            =   13440
      TabIndex        =   6
      Top             =   240
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
      MICON           =   "FrmDataBarang.frx":0038
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
      Left            =   1800
      TabIndex        =   2
      Top             =   7440
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
      MICON           =   "FrmDataBarang.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   240
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   11456
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "KodeBarang"
      Columns(0).DataField=   "KodeBarang"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NamaBarang"
      Columns(1).DataField=   "NamaBarang"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Jenis"
      Columns(2).DataField=   "Jenis"
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   33
      Columns(3)._MaxComboItems=   5
      Columns(3).ValueItems(0)._DefaultItem=   0
      Columns(3).ValueItems(0).Value=   "Buah"
      Columns(3).ValueItems(0).Value.vt=   8
      Columns(3).ValueItems(0).DisplayValue=   "Buah"
      Columns(3).ValueItems(0).DisplayValue.vt=   8
      Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(1)._DefaultItem=   0
      Columns(3).ValueItems(1).Value=   "Lusin"
      Columns(3).ValueItems(1).Value.vt=   8
      Columns(3).ValueItems(1).DisplayValue=   "Lusin"
      Columns(3).ValueItems(1).DisplayValue.vt=   8
      Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(2)._DefaultItem=   0
      Columns(3).ValueItems(2).Value=   "Roll"
      Columns(3).ValueItems(2).Value.vt=   8
      Columns(3).ValueItems(2).DisplayValue=   "Roll"
      Columns(3).ValueItems(2).DisplayValue.vt=   8
      Columns(3).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(3)._DefaultItem=   0
      Columns(3).ValueItems(3).Value=   "Keping"
      Columns(3).ValueItems(3).Value.vt=   8
      Columns(3).ValueItems(3).DisplayValue=   "Keping"
      Columns(3).ValueItems(3).DisplayValue.vt=   8
      Columns(3).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(4)._DefaultItem=   0
      Columns(3).ValueItems(4).Value=   "Meter"
      Columns(3).ValueItems(4).Value.vt=   8
      Columns(3).ValueItems(4).DisplayValue=   "Meter"
      Columns(3).ValueItems(4).DisplayValue.vt=   8
      Columns(3).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems(5)._DefaultItem=   0
      Columns(3).ValueItems(5).Value=   "Inch"
      Columns(3).ValueItems(5).Value.vt=   8
      Columns(3).ValueItems(5).DisplayValue=   "Inch"
      Columns(3).ValueItems(5).DisplayValue.vt=   8
      Columns(3).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(3).ValueItems.Count=   6
      Columns(3).Caption=   "Satuan"
      Columns(3).DataField=   "Satuan"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Harga Beli"
      Columns(4).DataField=   "HargaBeli"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Biaya Kirim"
      Columns(5).DataField=   "BiayaKirim"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Harga Pokok"
      Columns(6).DataField=   "HargaPokok"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Harga Grosir"
      Columns(7).DataField=   "HargaGrosir"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Harga Ecer"
      Columns(8).DataField=   "HargaEcer"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Stock Max"
      Columns(9).DataField=   "StockMax"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Stock Min"
      Columns(10).DataField=   "StockMin"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Stock"
      Columns(11).DataField=   "Stock"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(3).Button=1"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(51)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(52)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(10)._EditAlways=0"
      Splits(0)._ColumnProps(56)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(57)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(58)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(60)=   "Column(11)._EditAlways=0"
      Splits(0)._ColumnProps(61)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
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
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HFFC0C0&,.bold=-1"
      _StyleDefs(11)  =   ":id=2,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1,.appearance=0,.borderColor=&H80000013&"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H80000013&,.fgcolor=&H80000012&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin BasTombol.vbButton CmdCari 
      Height          =   375
      Left            =   12360
      TabIndex        =   4
      Top             =   7440
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
      MICON           =   "FrmDataBarang.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton vbButton3 
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   7440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "C&etak Laporan"
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
      MICON           =   "FrmDataBarang.frx":008C
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
      Caption         =   "Form Data Barang"
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
      TabIndex        =   7
      Top             =   240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   7935
      Left            =   120
      Top             =   120
      Width           =   13800
   End
End
Attribute VB_Name = "FrmDataBarang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsTemp1 As New ADODB.Recordset


Private Sub CmdCancel_Click()
Form_Load
End Sub

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
        Kalimat = Kalimat + " or namaBarang like '%" & Query & "%'"
    Next Query
    Kalimat = Mid(Kalimat, 4, Len(Kalimat) - 3)

SQL = "select * from Barang where " & Kalimat & _
    " order by KodeBarang"
Set Grid.DataSource = DbCon.Execute(SQL)
Grid.Refresh
TxtCari = ""
Timer1.Enabled = True
End Sub

Private Sub CmdSave_Click()
Dim Kolom As Integer
Dim Coba As Boolean
   On Error GoTo CmdSave_Click_Error

SQL = "Delete from Barang"
DbCon.Execute SQL

Grid.MoveFirst
While Not Grid.EOF
    For Kolom = 0 To 11
        If Trim(Grid.Columns(Kolom).Text) = "" Then
            Coba = True
        Else: Coba = False
        End If
    Next Kolom
    
    If Coba Then
        MsgBox "Data Tidak Lengkap"
        Grid.Columns(Kolom) = Trim(Grid.Columns(Kolom))
        Exit Sub
    Else
        SQL = "Select * from Barang where KodeBarang='" & Trim(Grid.Columns(0).Text) & "'"
        Set RSFind = DbCon.Execute(SQL)
        If RSFind.BOF Then
        SQL = "Insert into Barang values('" & _
                Trim(Grid.Columns(0).Text) & "','" & Trim(Grid.Columns(1).Text) & _
                "','" & Trim(Grid.Columns(2).Text) & "','" & Trim(Grid.Columns(3).Text) & _
                "'," & Val(Grid.Columns(4).Text) & "," & Val(Grid.Columns(5).Text) & _
                "," & Val(Grid.Columns(6).Text) & "," & Val(Grid.Columns(7).Text) & _
                "," & Val(Grid.Columns(8).Text) & "," & Val(Grid.Columns(9).Text) & _
                "," & Val(Grid.Columns(10).Text) & "," & Val(Grid.Columns(11).Text) & ")"
        DbCon.Execute SQL
        Else
            MsgBox "Kode Barang Sudah Ada. Mohon Diganti"
            Grid.Columns(0) = Trim(Grid.Columns(0).Text)
            Grid.MarqueeStyle = dbgHighlightRow
            Grid.Scroll 1, Grid.Row
            Exit Sub
        End If
    End If
Grid.MoveNext
Wend

MsgBox "Data Saved"
Grid.MarqueeStyle = dbgHighlightCell
Form_Load
FrmBeliBarang.CmbBarang.Refresh
FrmJualGrosir.CmbBarang.Refresh

   On Error GoTo 0
   Exit Sub

CmdSave_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CmdSave_Click of Form FrmDataBarang" & vbCrLf & "Jangan Ada tanda petik atas."
End Sub


Private Sub Form_Activate()
If Connect2 = False Then
    MsgBox "User Tidak Login"
    FrmLogin.Show
End If
End Sub

Private Sub Form_Load()
CreateTbl
Adodc1.ConnectionString = ConDB
Adodc1.RecordSource = "SELECT KodeJenis,NamaJenis FROM Jenis ORDER BY KodeJenis"
Adodc1.Refresh
TDBDropDown1.Refresh
Set TDBDropDown1.DataSource = Adodc1
TDBDropDown1.DataField = "NamaJenis"
TDBDropDown1.Columns(0).Visible = False
Timer1 = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
If RsTemp1.State Then RsTemp1.Close
End Sub


Private Sub Grid_BeforeDelete(Cancel As Integer)
If Grid.EOF Or Grid.BOF Then Cancel = False
End Sub

Private Sub Grid_BeforeUpdate(Cancel As Integer)
RsTemp1.Find "kodeBarang='" & Trim(Grid.Columns(ColIndex).Text) & "'", , adSearchForward, 1
If RsTemp1.BOF Then
Else
    MsgBox "Sudah Ada"
    Grid.Columns(0).Text = ""
    Exit Sub
End If
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Kolom As Integer
Dim Coba As Boolean
Dim Ex As String

If KeyCode = 46 And RsTemp1.RecordCount <> 0 Then
    Grid.Delete
    
ElseIf KeyCode = 40 Then
    For Kolom = 0 To 2
        If Trim(Grid.Columns(Kolom).Text) = "" Then
            Coba = True
        Else: Coba = False
        End If
    Next Kolom
    
    If Coba Then
        MsgBox "Data Tidak Lengkap"
        Grid.AllowAddNew = False
        Exit Sub
    Else
        Grid.AllowAddNew = True
    End If
End If
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
Form_Load
End Sub

Private Sub vbButton1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub



Private Sub CreateTbl(Optional Edit As Boolean = False)

On Error GoTo vb_error

  If RsTemp1.State Then RsTemp1.Close
  With RsTemp1
      .Fields.Append "KodeBarang", adVarChar, 50 '1
      .Fields.Append "NamaBarang", adVarChar, 50 '1
      .Fields.Append "Jenis", adVarChar, 50
      .Fields.Append "Satuan", adVarChar, 50
      .Fields.Append "HargaBeli", adDouble
      .Fields.Append "BiayaKirim", adDouble
      .Fields.Append "HargaPokok", adDouble
      .Fields.Append "HargaGrosir", adDouble
      .Fields.Append "HargaEcer", adDouble
      .Fields.Append "StockMax", adInteger
      .Fields.Append "StockMin", adInteger
      .Fields.Append "Stock", adInteger
      .Open
      If Edit = False Then
          .AddNew
          .Update
      End If
      SQL = "select * from barang order by kodebarang"
      Set RSFind = DbCon.Execute(SQL)
      RSFind.MoveFirst
      While Not RSFind.EOF
          RsTemp1.AddNew
          RsTemp1!KodeBarang = RSFind!KodeBarang
          RsTemp1!namaBarang = RSFind!namaBarang
          RsTemp1!jenis = RSFind!jenis
          RsTemp1!satuan = RSFind!satuan
          RsTemp1!HargaBeli = RSFind!HargaBeli
          RsTemp1!BiayaKirim = RSFind!BiayaKirim
          RsTemp1!HargaPokok = RSFind!HargaPokok
          RsTemp1!HargaGrosir = RSFind!HargaGrosir
          RsTemp1!HargaEcer = RSFind!HargaEcer
          RsTemp1!StockMax = RSFind!StockMax
          RsTemp1!StockMin = RSFind!StockMin
          RsTemp1!Stock = RSFind!Stock
          
          RsTemp1.Update
          RSFind.MoveNext
      Wend
      Set Grid.DataSource = RsTemp1
      Grid.MoveFirst
      Grid.Delete
      Grid.Refresh
  End With        'With rstemp1
        
Exit Sub
vb_error:
MsgBox ErrMessage(Erl, Err.Number, "Procedure : " & Me.Name & ".CreateTbl"), vbExclamation, "Err Number : " & Erl

End Sub

'        untuk membedakan antara baru dan lama
'        SQL = "Select * from Barang where KodeBarang='" & Trim(Grid.Columns(0).Text) & "'"
'        Set RSFind = DbCon.Execute(SQL)
'        If RSFind.BOF Then
'            SQL = "Insert into Barang values('" & _
'                Trim(Grid.Columns(0).Text) & "','" & Trim(Grid.Columns(1).Text) & _
'                "','" & Trim(Grid.Columns(2).Text) & "','" & Trim(Grid.Columns(3).Text) & _
'                "'," & Val(Grid.Columns(4).Text) & "," & Val(Grid.Columns(5).Text) & _
'                "," & Val(Grid.Columns(6).Text) & "," & Val(Grid.Columns(7).Text) & _
'                "," & Val(Grid.Columns(8).Text) & "," & Val(Grid.Columns(9).Text) & _
'                "," & Val(Grid.Columns(10).Text) & "," & Val(Grid.Columns(11).Text) & ")"
'        DbCon.Execute SQL
'        Else
'        End If

Private Sub vbButton3_Click()
 With Grid.PrintInfo

        ' Set the page header

        .PageHeaderFont.Italic = True

        .PageHeader = "Tabel Barang"

        

        ' Column headers will be on every page

        .RepeatColumnHeaders = True

        

        ' Display page numbers (centered)

        .PageFooter = "\tPage: \p"

        ' Invoke Print Preview

        .PrintPreview

    End With


End Sub
