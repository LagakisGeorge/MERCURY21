VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form helpforma 
   Caption         =   "Form1"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12795
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin AcroPDFLibCtl.AcroPDF AcroPDF1 
      Height          =   9375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   9615
      _cx             =   16960
      _cy             =   16536
   End
   Begin VB.CommandButton Command2 
      Caption         =   "одгциес се WORD"
      Height          =   255
      Left            =   3690
      TabIndex        =   1
      Top             =   45
      Width           =   3105
   End
   Begin VB.CommandButton Command1 
      Caption         =   "жоятыла одгциым ADOBE "
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   45
      Width           =   3105
   End
End
Attribute VB_Name = "helpforma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appWord As Word.Application
Dim WRDdoc As Word.Document

Private Sub Command1_Click()

   AcroPDF1.LoadFile "c:\mercvb\help\help.pdf"
 AcroPDF1.gotoFirstPage
End Sub

Private Sub Command2_Click()
'Start a new document in Word
'    Set mWordA = CreateObject("Word.Application")
'    Set mobjWORD = mWordA.Documents.Open "c:\mercvb\help\help.doc"
  Dim sFileName As String
  
   
sFileName = "c:\mercvb\help\help.doc"
Set appWord = New Word.Application
Set WRDdoc = appWord.Documents.Open(sFileName)
appWord.Visible = True
 
'DOK1.Activate
End Sub

