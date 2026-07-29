VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form help1 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   13035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14790
   Icon            =   "help.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13035
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   9420
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   8775
   End
   Begin VB.CommandButton cmdVERSIONS 
      Caption         =   "ΕΚΟΔΟΣΕΙΣ"
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   12000
      Width           =   990
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   9855
      Left            =   9000
      TabIndex        =   6
      Top             =   1920
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   17383
      _Version        =   393217
      BackColor       =   16777215
      TextRTF         =   $"help.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   11625
      Begin VB.Image imgLogo 
         Height          =   1185
         Left            =   240
         Picture         =   "help.frx":A517
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company :Lagakis Computers tel.+30 2521022225"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   1
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6000
         TabIndex        =   3
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2400
         TabIndex        =   5
         Top             =   720
         Width           =   2430
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   4320
      End
   End
End
Attribute VB_Name = "help1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdVERSIONS_Click()

100     lblVersion.Caption = MDIForm1.StatusBar1.Panels(3).Text  ' "Version 9.08" ' & App.Major & "." & App.Minor & "." & App.Revision
110     lblProductName.Caption = "MERCURY WIN"  'App.title

        Dim sql As String
116     sql = "    Εκδοση 25.09" + Chr(13)
sql = sql + "Επαναποστολή Mydata από Εισαγωγή Παρ/κου grid" + Chr(13)
sql = sql + "Στο Grid εισαγωγής/Παρουσίασης φαινεται και το Παραστατικό" + Chr(13)



120     sql = sql + "    Εκδοση 10.01" + Chr(13)


130     sql = sql + "Επανεκτύπωση γραμμής επισκευής και ετικετας" + Chr(13)
140     sql = sql + "Ταξινόμηση  οθόνης τελευταίων τιμολογίων " + Chr(13)
150     sql = sql + "Μενου σε οθόνη τελευταίων τιμολογίων " + Chr(13)

160     sql = sql + "    Εκδοση 9.08" + Chr(13)
170     sql = sql + "Επανεκτύπω3ση Απόδειξης Είσπραξης/Πληρωμης με δεξί κλικ στην κίνηση (Καρτέλλα Πελάτη/Προμηθευτή)" + Chr(13)
180     sql = sql + "Το πιστωτικό υπόλοιπο σε προμηθευτές φαίνεται θετικό " + Chr(13)
190     sql = sql + "Kαρτέλλα πελάτη/προμηθευτή σε Excel" + Chr(13)
200     sql = sql + "Aνανέωση οθόνης τελευταίων τιμολογίων σε περίπτωση διαγραφής" + Chr(13)
210     sql = sql + "Κίνηση είδους σε ξένο νόμισμα" + Chr(13)

220     RichTextBox1.Text = sql







End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

        '<EhHeader>
        On Error GoTo Form_KeyPress_Err

        '</EhHeader>
100     If KeyAscii = 27 Then
110         Unload Me
        End If

        '<EhFooter>
        Exit Sub

Form_KeyPress_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.help1.Form_KeyPress " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.help1.Form_KeyPress " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()
List1.AddItem "ΠΡΟΕΙΔΟΠΟΙΕΙ ΓΙΑ ΤΑ ΗΛΕΚΤΡΟΝ.ΔΕΛΤΙΣ ΤΙΜΟΛΟΓΙΟΥ-ΔΑ   28/7/2026"
List1.AddItem "ΑΠΟΘΗΚΕΥΕΙ ΣΤΟ EGGTIM TA CLASSIFICATIONCATEGORY/TYPE   28/7/2026"
List1.AddItem "σδα φορτ=>τιμολογια=>σδα επιστροφης με 9.2-μαρκ σχετ.μαρκ και 10.1 σδα επιστρ 23/7/2026"
List1.AddItem "Στον μετασχηματισμό δίνω κωδ.παρ/κου METASX_DA π.χ. A 23/7/26"
List1.AddItem "ΣΤΑ ΔΑ ΓΙΑ ΛΟΓ/ΣΜΟ ΤΡΙΤΩΝ ΓΥΡΙΖΕΙ ΣΕ ΣΚΟΠΟΣ=2.ΓΙΑ ΛΟΓ.ΤΡΙΤΩΝ-CORR.ENTIT 30/6/2026"
List1.AddItem "ΔΑ ΑΝΑΠΟΔΟ & ΔΕΛΤΙΟ ΠΑΡΑΛΑΒΗΣ ΣΥΣΧ & ΜΗ 16/6/2026"
List1.AddItem "EMAIL OUTLOOK OK. ΨΗΦIAKO MHΔENIZEI TO ΠΑΛIO ΑΦΜ ΝΑ ΜΗΝ ΠΑΕΙ ΣΤΟΝ ΕΠΟΜΕΝΟ 21/5/2026"
List1.AddItem "Προστέθηκε το email στο δεξί κλικ se Eκτύπωση/pdf/email με Outllok 9/5/2026"
List1.AddItem "Nielsen αποστολή στατιστικών 8/5/2026"
List1.AddItem "Προστέθηκε το Δελτίο Ποσοτικής Παραλαβής 10.2-(2.Μόνο διακίνηση)30/4/26"
List1.AddItem "Προστέθηκε το Σκανάρισμα Τιμ.Αγοράς στην εισαγωγή Παραστατικού 30/4/26"
List1.AddItem "ελεγχος ΔΑ με μαρκ GetDeliveryNoteStatus"
List1.AddItem "Μπορώ να βάλω Προεπιλεγμένο κατηγορία ΦΠΑ σε ένα παρ/κό.Π.χ. ΕΕ =>5 "
List1.AddItem "ολοκληρωση διορθωσης δωρεαν υπηρεσιας ηλεκτρονικό πελατολόγιο 24/4/2026"
List1.AddItem "Προεπιλογή Πωλήσεις ή Αγορές στην Εισαγωγή Παραστατικού"
List1.AddItem "Ανάποδα φόρτωση-προορισμός με παράμετρο στην εισαγωγή Παραστατικου"
List1.AddItem "Αποστολή 3.1 τιτλος κτησης με πάροχο ΟΚ και με ψηφιακό τέλος αν χρειάζεται."
List1.AddItem "Αλλος σκοπος=19 φαινεται και στον πάροχο."
List1.AddItem "Ακυρωτικό αυτοματο ακυρωνονταν οταν δεν ειχε internet(καταργ.Transaction ΣΕ ΠΑΡΟΧ0 1-4-2026"
List1.AddItem "ΑΚΥΡΩΤΙΚΟ Δ.Α. ΣΕ ΠΑΡΟΧ0 27-3-2026"
List1.AddItem "ΤΑ ΕΙΔΗ ΧΩΡΙΣ ΜΟΝΑΔΑ ΤΑ ΓΕΜΙΖΩ ΜΕ <ΤΕΜ> ΓΡΑΦΩ ΕΓΩ ΤΙ ΘΕΛΩ  12-3-2026"
List1.AddItem "ΘΕΛΕΙΣ ΤΑ ΕΙΔΗ ΜΕ ΑΓΝ.ΜΟΝΑΔΑ  ΝΑ ΠΡΟΣΤΕΘΟΥΝ ΣΤΗΝ ΛΙΣΤΑ ΤΩΝ ΜΟΝΑΔΩΝ(ΘΑ ΛΕΙΠΟΥΝ ΜΟΝΟ ΤΑ UBL) 12-3-2026"
List1.AddItem "Εργαλείο στα Βοηθητικά 7->Για Μον.Μετρ.UBL 12-3-26 "
List1.AddItem "Βγάζει το σωστό τίτλο παρ/κου οταν στελνω mydata ετεροχρονισμένα 9-3-26"
List1.AddItem "AN ΔEN EXEI TK TO BAZEI AYTOMATA ΑΠΟ ΑΑΔΕ AN EINAI EΛΛHNAΣ ΕΠΑΓΓΕΛΜΑΤΙΑΣ 8/3/26"
List1.AddItem "ΣΤΕΛΝΕΙ ΑΠΟ ΟΘΟΝΗ MYDATA KAI ΣΕ ΠΑΡΟΧΟ ΠΛΗΡΩΣ ΜΕ PDF 8/3/26"
List1.AddItem "ΣΤΑ ΒΕΝΖΙΝΑΔΙΚΑ ΔΕΝ ΒΑΖΕΙ FUELCODE ΣΤΑ ΛΙΠΑΝΤΙΚΑ"

List1.AddItem "ΕΛΕΓΧΕΙ ΑΝ ΥΠΑΡΧΕΙ ΤΟ MERCURYPDF ΟΤΑΝ ΕΧΩ ΠΑΡΟΧΟ ΣΤΙΣ ΠΑΡΑΜΕΤΡΟΥΣ ΠΡΟΓΡΑΜΜΑΤΟΣ"

List1.AddItem "Γέφυρα ΓΛ,param_to_array gParam(500, 2),αριθμηση χρεωπιστωσεων=οκ 22/1/26"
List1.AddItem "Προεπιλεγμένο ΤΑΒ τιμολόγησης :τόπος προορισμού 18/1/26"
List1.AddItem "Εισαγωγή πελατη απο τιμολογηση : αποθηκευει και το ΤΚ 18/1/26"

List1.AddItem "Ταμειακες : 1η στηλη ο α/α γιατι κολλουσε στο backspace 16/1/26"


List1.AddItem "ΕΚΤΥΠΩΣΗ ΚΙΝΗΣΕΩΝ ΜΕ ΦΙΛΤΡΟ ΤΗΝ ΚΙΝΗΣΗ Π.Χ. ΣΕ ΤΡΑΠΕΖΑ ΠΕΙΡΑΙΩΣ 16,1,26"
List1.AddItem "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ ΕΙΣΠΡΑΞΕΙΣ CHECK EPANENHMEROSI=OK EISPR+PARAST 15.1.26"
List1.AddItem "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ ΕΙΣΠΡΑΞΕΙΣ LEFT JOIN XREOPIS ON EIDXPI 15.1.26"
List1.AddItem "ΓΕΝΙΚΗ ΛΟΓΙΣΤΙΚΗ ΕΙΣΠΡΑΞΕΙΣ 15.1.26"
        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

        '<EhFooter>/
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.help1.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.help1.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Frame1_Click()

        '<EhHeader>
        On Error GoTo Frame1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Frame1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.help1.Frame1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        SAVE_ERROR Err.Description & " in ADOMERCNEW.help1.Frame1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

