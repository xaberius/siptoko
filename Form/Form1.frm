VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Susu bendera saya satu"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim Coba As Boolean
Dim Kata() As String
Dim Angka As Integer
Dim cooo As Variant
Coba = True

    Kata = Split(Label1.Caption, " ")

    For Each cooo In Kata()
        MsgBox cooo
    Next cooo
    
    
    
'Text1 = InStr(1, Label1.Caption, " ")
'Text1 = Left(Label1.Caption, InStr(Label1.Caption, " "))
'Text2 = Mid(Label1.Caption, InStr(Label1.Caption, " ") + 1)
End Sub

