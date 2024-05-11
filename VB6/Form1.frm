VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucPrinterComboEx Test Form"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Info"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin ucPrinterComboExTest.ucPrinterComboEx ucPrinterComboEx1 
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 
Private Sub Command1_Click()
    Dim sMsg As String
    sMsg = "Selected printer: " & ucPrinterComboEx1.SelectedPrinter & vbCrLf & _
           "Total printers: " & ucPrinterComboEx1.PrinterCount & vbCrLf
    Dim i As Long
    For i = 0 To ucPrinterComboEx1.PrinterCount - 1
        sMsg = sMsg & "Printer " & i & ": " & ucPrinterComboEx1.Printers(i) & vbCrLf
    Next
    MsgBox sMsg, vbOKOnly, App.Title
End Sub

Private Sub ucPrinterComboEx1_PrinterChanged(ByVal sNewPrinterName As String, ByVal sParsingPath As String, ByVal sModelName As String, ByVal sNetworkLocation As String, ByVal sLastStatusMessage As String, ByVal bIsDefaultPrinter As Boolean)
    Debug.Print "Selected printer changed to: " & sNewPrinterName & " (Model: " & sModelName & ", Status=" & sLastStatusMessage & ")"
End Sub
