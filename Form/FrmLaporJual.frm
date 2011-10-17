VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FrmLaporJual 
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdKeluar 
      BackColor       =   &H0080FFFF&
      Cancel          =   -1  'True
      Caption         =   "&Keluar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1875
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1185
   End
   Begin VB.CommandButton CmdSetting 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1785
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "FrmLaporJual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New Report
Private Sub CmdKeluar_Click()
Unload Me
End Sub

Private Sub CmdSetting_Click()
  Report.PrinterSetup Me.hwnd
  CRViewer1.Refresh
End Sub

'bengkel as c, department as d, dealer as e ,mstpemakai as f
Private Sub Form_Load()
'a.no_urut,a.tgl_urut,a.no_mesin,a.no_rangka,a.no_polisi,a.cabang,a.pemakai,a.department,a.bengkel
If FrmLapJual.TxtLapor = "grosir" Then
    ssql = "SELECT     * FROM         DtlJualGrosir INNER JOIN " & _
        "Barang ON DtlJualGrosir.KodeBarang = Barang.KodeBarang INNER JOIN " & _
        "JualGrosir ON DtlJualGrosir.KodeTransaksi = JualGrosir.KodeTransaksi INNER JOIN " & _
        "Konsumen ON JualGrosir.KodeKonsumen = Konsumen.KodeKonsumen " & _
        "where JualGrosir.kodeTransaksi='" & Trim(FrmLapJual.Grid.Columns(0).Text) & _
        "' order by JualGrosir.kodeTransaksi "
    
    Set Report = New ReportJual
ElseIf FrmLapJual.TxtLapor = "ecer" Then
    ssql = "SELECT     * FROM         DtlJualecer INNER JOIN " & _
        "Barang ON DtlJualecer.KodeBarang = Barang.KodeBarang INNER JOIN " & _
        "Jualecer ON DtlJualecer.KodeTransaksi = Jualecer.KodeTransaksi INNER JOIN " & _
        "Konsumen ON Jualecer.KodeKonsumen = Konsumen.KodeKonsumen " & _
        "where JualEcer.kodeTransaksi='" & Trim(FrmLapJual.Grid.Columns(0).Text) & _
        "' order by JualEcer.kodeTransaksi "
    
    Set Report = New ReportJual2
End If

Set oRS = DbCon.Execute(ssql)
Report.Database.SetDataSource oRS
CRViewer1.ReportSource = Report
Screen.MousePointer = vbHourglass
Me.WindowState = 2


CRViewer1.Zoom 100
CRViewer1.ViewReport
CmdSetting.Left = Screen.Width - (CmdSetting.Width + CmdKeluar.Width)
CmdKeluar.Left = Screen.Width - (CmdKeluar.Width)
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub




