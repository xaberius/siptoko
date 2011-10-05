VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{8B946F6F-F1C6-4F89-A615-115403ACC638}#1.0#0"; "BasTombol.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmMutasi 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtKet 
      Height          =   660
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbBarang 
      Height          =   330
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
      _Version        =   196616
      BackColorOdd    =   16761024
      Columns(0).Width=   3200
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbAsal 
      Height          =   330
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
      DataFieldList   =   "Column 0"
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Row.Count       =   3
      Row(0)          =   "Gudang 1"
      Row(1)          =   "Gudang 2"
      Row(2)          =   "Toko"
      BackColorOdd    =   16761024
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Caption=   "Tempat"
      Columns(0).Name =   "Tempat"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4440
      Top             =   120
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
   Begin TDBNumber6Ctl.TDBNumber TxtJumlah 
      Height          =   330
      Left            =   2040
      TabIndex        =   5
      Top             =   3480
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   582
      Calculator      =   "FrmMutasi.frx":0000
      Caption         =   "FrmMutasi.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMutasi.frx":008C
      Keys            =   "FrmMutasi.frx":00AA
      Spin            =   "FrmMutasi.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99999
      MinValue        =   -99999
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
      MaxValueVT      =   1549533189
      MinValueVT      =   1701707781
   End
   Begin TDBDate6Ctl.TDBDate TxtTgl 
      Height          =   330
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   582
      Calendar        =   "FrmMutasi.frx":011C
      Caption         =   "FrmMutasi.frx":0248
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "FrmMutasi.frx":02B4
      Keys            =   "FrmMutasi.frx":02D2
      Spin            =   "FrmMutasi.frx":0330
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd mmmm yyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm/dd/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "10/03/2011"
      ValidateMode    =   0
      ValueVT         =   1397030919
      Value           =   40819
      CenturyMode     =   0
   End
   Begin BasTombol.vbButton vbButton1 
      Height          =   375
      Left            =   5760
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
      MICON           =   "FrmMutasi.frx":0358
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
      Left            =   6240
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
      MICON           =   "FrmMutasi.frx":0374
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
      Left            =   1560
      TabIndex        =   8
      Top             =   6360
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
      MICON           =   "FrmMutasi.frx":0390
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
      Left            =   240
      TabIndex        =   7
      Top             =   6360
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
      MICON           =   "FrmMutasi.frx":03AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BasTombol.vbButton CmdInput 
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "&Input"
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
      MICON           =   "FrmMutasi.frx":03C8
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
      Height          =   2055
      Left            =   240
      TabIndex        =   9
      Top             =   4080
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Nomor"
      Columns(0).DataField=   "noKet"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Kode Barang"
      Columns(1).DataField=   "KodeBarang"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama Barang"
      Columns(2).DataField=   "NamaBarang"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Jumlah"
      Columns(3).DataField=   "Jumlah"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Named:id=33:Normal"
      _StyleDefs(53)  =   ":id=33,.parent=0"
      _StyleDefs(54)  =   "Named:id=34:Heading"
      _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(56)  =   ":id=34,.wraptext=-1,.appearance=0,.borderColor=&H80000013&"
      _StyleDefs(57)  =   "Named:id=35:Footing"
      _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   "Named:id=36:Selected"
      _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=37:Caption"
      _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(63)  =   "Named:id=38:HighlightRow"
      _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=39:EvenRow"
      _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(67)  =   "Named:id=40:OddRow"
      _StyleDefs(68)  =   ":id=40,.parent=33"
      _StyleDefs(69)  =   "Named:id=41:RecordSelector"
      _StyleDefs(70)  =   ":id=41,.parent=34"
      _StyleDefs(71)  =   "Named:id=42:FilterBar"
      _StyleDefs(72)  =   ":id=42,.parent=33"
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbTujuan 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
      DataFieldList   =   "Column 0"
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Row.Count       =   3
      Row(0)          =   "Gudang 1"
      Row(1)          =   "Gudang 2"
      Row(2)          =   "Toko"
      BackColorOdd    =   16761024
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).Caption=   "Tempat"
      Columns(0).Name =   "Tempat"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      _ExtentX        =   4048
      _ExtentY        =   582
      _StockProps     =   93
      ForeColor       =   0
      BackColor       =   -2147483643
   End
   Begin VB.Label Label7 
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
      TabIndex        =   18
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
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
      Top             =   3480
      Width           =   1575
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
      TabIndex        =   16
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tujuan"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tempat Asal"
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
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Mutasi"
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
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Height          =   6255
      Left            =   120
      Top             =   600
      Width           =   6480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Mutasi Barang"
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
      TabIndex        =   12
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmMutasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Keterangan1 As Integer
Dim RsTemp1 As New ADODB.Recordset

Private Sub CmbAsal_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub CmbBarang_DropDown()
Adodc1.RecordSource = ""
SQL = "Select KodeBarang,NamaBarang,Satuan from Barang order by kodeBarang"
Set RSFind = DbCon.Execute(SQL)
If RSFind.BOF Then Exit Sub
Adodc1.RecordSource = SQL
Adodc1.Refresh
With cmbBarang
    .DataSourceList = Adodc1
    .DataFieldList = "NamaBarang"
    .Columns(0).Visible = False
    .Columns(1).Width = 2000
End With
End Sub


Private Sub CmbBarang_GotFocus()
If IsNull(TxtTgl) Or TxtTgl < Date Then
    MsgBox "Tanggal Tidak Valid"
    TxtTgl.SetFocus
    Exit Sub
ElseIf CmbAsal = CmbTujuan Then
    MsgBox "Tempat Asal Jangan Sama Dengan Tempat Tujuan"
    CmbTujuan = ""
    CmbTujuan.SetFocus
    Exit Sub
ElseIf Trim(CmbAsal) = "" Or Not CmbAsal.IsItemInList Then
    MsgBox "Tempat Asal masih Kosong"
    CmbAsal.SetFocus
    Exit Sub
ElseIf Trim(CmbTujuan) = "" Or Not CmbTujuan.IsItemInList Then
    MsgBox "Tempat Tujuan Masih Kosong"
    CmbTujuan.SetFocus
    Exit Sub
ElseIf Trim(TxtKet) = "" Then
    MsgBox "Keterangan Masih Kosong"
    TxtKet.SetFocus
    Exit Sub
End If
End Sub

Private Sub CmbTujuan_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub CmdCancel_Click()
Form_Load
End Sub

Private Sub CmdInput_Click()
If Trim(cmbBarang) = "" And Not cmbBarang.IsItemInList Then
    MsgBox "Barang Masih Kosong"
    cmbBarang.SetFocus
    Exit Sub
ElseIf TxtJumlah = 0 Then
    MsgBox "Jumlah masih Kosong"
    TxtJumlah.SetFocus
    Exit Sub
End If

RsTmp.Find "namaBarang='" & Trim(cmbBarang) & "'", , adSearchForward, 1
If RsTemp1.EOF Then
    With RsTemp1
        .AddNew
        !noket = RsTemp1.RecordCount
        !namaBarang = Trim(cmbBarang)
        !KodeBarang = Trim(cmbBarang.Columns(0).Text)
        !jumlah = Val(TxtJumlah)
        .Update
    End With
    Grid.DataSource = RsTemp1
    Grid.Refresh
Else
    MsgBox "Barang Sudah Diinputkan."
    Exit Sub
End If
cmbBarang = ""
TxtJumlah = 0
cmbBarang.SetFocus
End Sub


Private Sub CmdSave_Click()
SQL = "insert into Mutasi values('" & FormatTgl(TxtTgl) & "','" & Trim(CmbAsal) & "','" & Trim(CmbTujuan) & _
    "','" & Trim(TxtKet) & "')"
DbCon.Execute SQL
RsTemp1.MoveFirst
While Not RsTemp1.EOF
    With RsTemp1
        SQL = "insert into DtlMutasi values('" & FormatTgl(TxtTgl) & "','" & !noket & "','" & _
            !KodeBarang & "'," & !jumlah & ")"
        DbCon.Execute SQL
        .MoveNext
    End With
Wend
MsgBox "Data Saved"
Bersih
End Sub

Private Sub Form_Load()
Bersih
Adodc1.ConnectionString = ConDB
With RsTemp1
    .Fields.Append "noKet", adInteger, 4
    .Fields.Append "KodeBarang", adVarChar, 50
    .Fields.Append "NamaBarang", adVarChar, 50
    .Fields.Append "Jumlah", adInteger, 4
    .Open
End With

End Sub
Sub Bersih()
TxtTgl = Date
CmbAsal = ""
CmbTujuan = ""
cmbBarang = ""
TxtJumlah = 0
End Sub

Private Sub SSOleDBCombo1_GotFocus()
SendKeys "{F4}"
End Sub

Private Sub Grid_Click()
Keterangan1 = Grid.Columns(0).Text
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If Keterangan1 = 0 Then
    MsgBox "Klik Salah Satu Item Di Tabel"
    Exit Sub
End If
RsTemp1.Find "noket='" & Keterangan1 & "'", , adSearchForward, 1
If Not RsTemp1.EOF Then
    MsgBox RsTemp1!noket & " Dibatalkan"
    RsTemp1.Delete
End If
    Keterangan1 = Keterangan1 + 1
    RsTemp1.Find "noket='" & Keterangan1 & "'", , adSearchForward, 1
    While Not RsTemp1.EOF
        With RsTemp1
            .Clone
            !noket = !noket - 1
            .Update
        End With
        RsTemp1.MoveNext
    Wend
Grid.Refresh
End If
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Keterangan1 = Val(Me.Grid.Columns(0).Text)
End Sub

Private Sub TxtJumlah_GotFocus()
CmbBarang_GotFocus
End Sub

Private Sub TxtJumlah_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub TxtKet_KeyDown(KeyCode As Integer, Shift As Integer)
Enter KeyCode
End Sub

Private Sub vbButton1_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub vbButton2_Click()
Unload Me
End Sub
