VERSION 2.00
Begin Form Conv 
   Caption         =   "3GL Convert"
   ClientHeight    =   5640
   ClientLeft      =   1215
   ClientTop       =   4350
   ClientWidth     =   9285
   FontBold        =   -1  'True
   FontItalic      =   0   'False
   FontName        =   "Fixedsys"
   FontSize        =   9
   FontStrikethru  =   0   'False
   FontUnderline   =   0   'False
   Height          =   6045
   Left            =   1155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9285
   Top             =   4005
   Width           =   9405
   Begin ListBox AList 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   1605
      Left            =   1605
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   1365
   End
   Begin CommandButton ClpCopy 
      Caption         =   "Clipboard"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   7875
      TabIndex        =   15
      Top             =   525
      Width           =   1260
   End
   Begin ListBox VList 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   2055
      Left            =   1605
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   1365
   End
   Begin ListBox LList 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   4305
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1275
   End
   Begin ListBox OBJ 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   4305
      Left            =   3105
      TabIndex        =   10
      Top             =   1035
      Visible         =   0   'False
      Width           =   6075
   End
   Begin CommandButton Conv 
      Caption         =   "Convert"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   5505
      TabIndex        =   9
      Top             =   525
      Width           =   1095
   End
   Begin ComboBox Ldest 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   3135
   End
   Begin ComboBox LSrc 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   3135
   End
   Begin CommandButton Clean 
      Caption         =   "Clear"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   6705
      TabIndex        =   4
      Top             =   525
      Width           =   1095
   End
   Begin ListBox Lst 
      FontBold        =   -1  'True
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   4305
      Left            =   3105
      TabIndex        =   3
      Top             =   1035
      Width           =   6075
   End
   Begin CommonDialog CMDialog 
      Filter          =   "*.*|*.*"
      Flags           =   4
      Left            =   5160
      PrinterDefault  =   0   'False
      Top             =   15
   End
   Begin CommandButton GetIn 
      Caption         =   "..."
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   285
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin TextBox FileIn 
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   330
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin CommandButton Load 
      Caption         =   "Load"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   315
      Left            =   4305
      TabIndex        =   0
      Top             =   525
      Width           =   1095
   End
   Begin Label Label5 
      Caption         =   "Arrays"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   1605
      TabIndex        =   17
      Top             =   3480
      Width           =   1155
   End
   Begin Label Label4 
      Caption         =   "Variables"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   180
      Left            =   1620
      TabIndex        =   14
      Top             =   960
      Width           =   1155
   End
   Begin Label Label3 
      Caption         =   "Labels"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   180
      Left            =   210
      TabIndex        =   13
      Top             =   945
      Width           =   765
   End
   Begin Label Label2 
      Caption         =   "To:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin Label Label1 
      Caption         =   "From:"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "Fixedsys"
      FontSize        =   9
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
End
Dim RunFlag As Integer






Sub AList_Click ()
    MsgBox Str(Alist.ItemData(Alist.ListIndex))
End Sub

Sub clean_click ()
  Lst.Clear
  Obj.Clear
  llist.Clear
  Vlist.Clear
  Alist.Clear

  Obj.Visible = False
  Lst.Visible = True

End Sub

Sub ClpCopy_Click ()
Dim x As Integer
Dim a As String

  screen.MousePointer = 11
  a = ""
  For x = 0 To Obj.ListCount
    a = a + Obj.List(x) & Chr(13) & Chr(10)
  Next
  clipboard.SetText a
  screen.MousePointer = 0

End Sub

Sub Conv_Click ()
  
    If RunFlag < 1 Then MsgBox "Nessun programma in memoria": Exit Sub
    
  screen.MousePointer = 11
  
  Obj.Visible = True
  Lst.Visible = False

  Obj.Clear

  Select Case LFrom
  Case 1 'Spectrum
    Spectrum
  Case 6 'Cobol
    Cobol

  Case Else
    MsgBox "Linguaggio non ancora implementato."
  End Select

  screen.MousePointer = 0

End Sub

Sub Form_Load ()
  RunFlag = 0
  LoadLanguages
  For x = 1 To UBound(Lang)
    LSrc.AddItem (Lang(x).name)
    ldest.AddItem (Lang(x).name)
  Next
End Sub

Sub GetIn_Click ()
  CMDialog.Action = 1
  FileIn = CMDialog.Filename
End Sub

Sub Ldest_Click ()
    If Lang(ldest.ListIndex + 1).Structured = False Then MsgBox "Il linguaggio di destinazione scelto non e' strutturato." & Chr(10) & "La conversione puo' essere impossibile.", , "Attenzione !"
End Sub

Sub Load_Click ()

  clean_click

  RunFlag = 1
  screen.MousePointer = 11
  If LSrc.ListIndex = -1 Then MsgBox "Specificare un linguaggio da cui tradurre", , "Errore durante il caricamento": GoTo ConvExit
  If ldest.ListIndex = -1 Then MsgBox "Specificare un linguaggio in cui tradurre", , "Errore durante il caricamento": GoTo ConvExit
  LFrom = LSrc.ListIndex + 1
  LTo = ldest.ListIndex + 1
  On Error GoTo ConvErr

  Dim fnum As Integer
  Dim a As String
  Dim b As String

  fnum = FreeFile
  Open FileIn For Binary As #fnum
  a = " "
  b = ""

  Do While Not EOF(fnum)
    Get #fnum, , a
    b = b + a
    If InStr(b, Lang(LFrom).EOL) > 0 Then
      b = Left(b, Len(b) - Len(Lang(LFrom).EOL))
      Lst.AddItem b
      b = ""
    End If
  Loop
  screen.MousePointer = 0
Exit Sub

ConvErr:
    MsgBox Error, , "Errore durante il caricamento"
ConvExit:
    screen.MousePointer = 0
    RunFlag = 0
    Exit Sub

End Sub

Sub VList_Click ()
  MsgBox Str(Vlist.ItemData(Vlist.ListIndex))
End Sub

