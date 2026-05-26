VERSION 5.00
Begin VB.Form PAR13 
   Caption         =   "≈Œ¡√Ÿ√«"
   ClientHeight    =   12255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   12255
   ScaleWidth      =   18600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   15255
      Begin VB.CommandButton cmd≈Œœƒœ” 
         Caption         =   "≈Œœƒœ”"
         Height          =   360
         Left            =   12960
         TabIndex        =   33
         Top             =   8640
         Width           =   990
      End
      Begin VB.TextBox RECEPIENT 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   32
         Top             =   8160
         Width           =   8775
      End
      Begin VB.TextBox FINAL_C13_TEXT7 
         Height          =   375
         Left            =   3480
         TabIndex        =   31
         Top             =   7680
         Width           =   8775
      End
      Begin VB.TextBox SHIP_AYTOK 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   7200
         Width           =   8775
      End
      Begin VB.TextBox ETASKG_SKOPOS 
         Height          =   375
         Left            =   3480
         TabIndex        =   29
         Top             =   6600
         Width           =   8775
      End
      Begin VB.TextBox DISCHARGE_PROOR 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   28
         Top             =   6000
         Width           =   8775
      End
      Begin VB.TextBox LOADING_FORTOSH 
         Height          =   375
         Left            =   3480
         TabIndex        =   27
         Top             =   5520
         Width           =   8775
      End
      Begin VB.TextBox FLAG_C12_TEXT6 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   26
         Top             =   5040
         Width           =   8775
      End
      Begin VB.TextBox VESSEL_C3_TEXT3 
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   4680
         Width           =   8775
      End
      Begin VB.TextBox PAYMENT_TRP_ 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   24
         Top             =   3840
         Width           =   8775
      End
      Begin VB.TextBox TOTALAMOUNT 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Top             =   3360
         Width           =   8775
      End
      Begin VB.TextBox TRADETERMS_C2_TEXT2 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   2880
         Width           =   8775
      End
      Begin VB.TextBox GROSS_NUM11_TEXTM5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   2400
         Width           =   8775
      End
      Begin VB.TextBox NET_NUM2_TEXTM4 
         BackColor       =   &H00C0FFC0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   2040
         Width           =   8775
      End
      Begin VB.TextBox QUANTITY_NUM1_TEXTM1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   1560
         Width           =   8775
      End
      Begin VB.TextBox COLI_C1_TEXT0 
         BackColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   960
         Width           =   8775
      End
      Begin VB.TextBox DESTINATION_SKOPOS2_TEXTENTOLEAS 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
      Begin VB.Label lblRECEPIENT 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "RECEPIENT"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   8160
         Width           =   3120
      End
      Begin VB.Label lblFINALDESTINATION 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINAL DESTINATION"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   7800
         Width           =   3120
      End
      Begin VB.Label lblSHIPPINGCOMPANY 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "SHIPPING COMPANY"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   7200
         Width           =   3120
      End
      Begin VB.Label lblETASKG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ETA SKG"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   6600
         Width           =   3120
      End
      Begin VB.Label lblPORTOFDIS 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "PORT OF DISCHARGE"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   6000
         Width           =   3120
      End
      Begin VB.Label lblPORTOF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PORT OF LOADING"
         Height          =   555
         Left            =   360
         TabIndex        =   12
         Top             =   5520
         Width           =   3120
      End
      Begin VB.Label lblFLAG 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "FLAG"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   5040
         Width           =   3120
      End
      Begin VB.Label lblVESSEL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VESSEL"
         Height          =   555
         Left            =   360
         TabIndex        =   10
         Top             =   4680
         Width           =   3120
      End
      Begin VB.Label lblPAYMENTTERMS 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "PAYMENT TERMS:"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   3840
         Width           =   3120
      End
      Begin VB.Label lblTOTALAMOUNT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT:"
         Height          =   555
         Left            =   360
         TabIndex        =   8
         Top             =   3480
         Width           =   3120
      End
      Begin VB.Label lblTRADETERMS 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "TRADE TERMS:"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   2880
         Width           =   3120
      End
      Begin VB.Label lblGROSSWEIGHT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GROSS WEIGHT:"
         Height          =   555
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   3120
      End
      Begin VB.Label lblNETWEIGHT 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "NET WEIGHT:"
         Height          =   315
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label lblQUANTITY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblCOLI 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "COLI"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   960
         Width           =   3120
      End
      Begin VB.Label lblDESTINATION 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINATION"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2940
      End
   End
End
Attribute VB_Name = "PAR13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd≈ÓÔ‰ÔÚ_Click()
PAR1.parat.Text = RECEPIENT.Text
  PAR1.Textm(7).Text = FINAL_C13_TEXT7.Text
  
 PAR1.AYTOKINHTO.Text = SHIP_AYTOK.Text
  'PAR1.SKOPOS.Text = ETASKG_SKOPOS
  
 ' PAR1.PROORISMOS.Text = DISCHARGE_PROOR
 ' PAR1.FORTOSH.Text = LOADING_FORTOSH
  
  
  
  
  PAR1.TC5.Text = DISCHARGE_PROOR
  PAR1.TC6.Text = LOADING_FORTOSH
  PAR1.tc7.Text = ETASKG_SKOPOS
  
  
  PAR1.Textm(6).Text = FLAG_C12_TEXT6
  PAR1.Textm(3).Text = VESSEL_C3_TEXT3
  'PAR1.Combo4.Text = PAYMENT_TRP_.Text
   PAR1.Synolo.Caption = TOTALAMOUNT.Text
  PAR1.Textm(2).Text = TRADETERMS_C2_TEXT2
  PAR1.Textm(5).Text = GROSS_NUM11_TEXTM5
 PAR1.Textm(4).Text = NET_NUM2_TEXTM4
 PAR1.Textm(1).Text = QUANTITY_NUM1_TEXTM1
 PAR1.Textm(0).Text = COLI_C1_TEXT0
  PAR1.TextEntoleas.Text = DESTINATION_SKOPOS2_TEXTENTOLEAS










Unload Me
End Sub

Private Sub Form_Load()
  RECEPIENT.Text = PAR1.parat.Text
  
  FINAL_C13_TEXT7.Text = PAR1.Textm(7).Text
  
  SHIP_AYTOK.Text = PAR1.AYTOKINHTO.Text
  
  ETASKG_SKOPOS = PAR1.tc7.Text
  DISCHARGE_PROOR = PAR1.TC5.Text
  LOADING_FORTOSH = PAR1.TC6.Text
  
  
  
  FLAG_C12_TEXT6 = PAR1.Textm(6).Text
  VESSEL_C3_TEXT3 = PAR1.Textm(3).Text
  PAYMENT_TRP_.Text = PAR1.Combo4.Text
  TOTALAMOUNT.Text = PAR1.Synolo.Caption
  TRADETERMS_C2_TEXT2 = PAR1.Textm(2).Text
  GROSS_NUM11_TEXTM5 = PAR1.Textm(5).Text
  NET_NUM2_TEXTM4 = PAR1.Textm(4).Text
  QUANTITY_NUM1_TEXTM1 = PAR1.Textm(1).Text
  COLI_C1_TEXT0 = PAR1.Textm(0).Text
  DESTINATION_SKOPOS2_TEXTENTOLEAS = PAR1.TextEntoleas.Text
  
  
  
  
  
  
  
End Sub

