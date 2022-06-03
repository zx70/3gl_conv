Type Language
  name       As String
  FilesName  As String
  Structured As Integer
  EOL        As String
  StrSep     As String
  GT         As String ' >
  GE         As String ' >=
  LT         As String ' <
  LE         As String ' <=
  NE         As String ' <>
  EQ         As String ' =
  AND        As String
  OR         As String
  XOR        As String
  EQV        As String ' = logico (IF a AND B THEN...)
  NOT        As String ' !=

End Type

Global Lang(6) As Language

Global LFrom As Integer
Global LTo As Integer

Sub Array (ByVal VName As String, VType As Integer, ADim() As String)

VName = Trim(VName)

' VType:    0 - Not Defined
'           2 - Array of Integer (Word)
'           4 - Array of Number
'           6 - Array of String

 Dim ArrStr As String

 If pass = 1 And VType <> 0 Then
   ArrStr = VName
   For x = 0 To UBound(ADim)
     ArrStr = ArrStr & " " & ADim(x)
   Next
   conv.AList.AddItem (ArrStr)
   conv.AList.ItemData(conv.AList.NewIndex) = VType
 End If

End Sub

Function CountChars (ByVal St As String, Chars As String) As Integer
Dim pos As Integer
Dim Count As Integer

Count = 0

For pos = 1 To Len(St)
  If Mid(St, pos, 1) = Chars Then Count = Count + 1
Next

CountChars = Count

End Function

Sub DeclVar (ByVal VName As String, VType As Integer)

VName = Trim(VName)

' VType:    0 - Not Defined
'           1 - Integer
'           2 - Number
'           3 - String

 If pass = 1 And VType <> 0 Then
   conv.VList.AddItem (VName)
   conv.VList.ItemData(conv.VList.NewIndex) = VType
 End If


End Sub

'Pos viene impostato a 0 se puntava all'ultima parola
Function GetWord (ByVal Text As String, pos As Integer, ByVal Char As String) As String
  Dim x As Integer
  Dim y As Integer
  Dim z As Integer
  Dim min As Integer
  Dim tst As Integer

    If Len(Text) = 0 Then
      pos = 0
      Exit Function
    End If
    x = pos
    If Mid(Text, x, 1) = " " Then ' Tolgo Spazi
      While Mid(Text, x, 1) = " "
        x = x + 1
        If x = Len(Text) + 1 Then
          pos = 0
          GetWord = ""
          Exit Function
        End If
      Wend
    End If
    If Char = "" Then
      y = InStr(x, Text, " ")
    ElseIf Len(Char) = 1 Then
      y = InStr(x, Text, Char)
    Else
      min = Len(Text) + 1
      For z = 1 To Len(Char)
        tst = InStr(x, Text, Mid(Char, z, 1))
        If tst < min And tst > 0 Then min = tst
      Next
      y = min
    End If
    If y = 0 Then
      pos = 0
      GetWord = Mid(Text, x)
    Else
      GetWord = Mid(Text, x, y - x)
      pos = y
    End If

End Function

Function IsNum (ByVal Num As String) As Integer
  Dim a As Double
  
  On Error GoTo NoNum

  a = Format(Num)
  IsNum = True
  Exit Function

NoNum:
  IsNum = False
  Exit Function

End Function

Function IsStr (St As String) As Integer
  Dim MySt As String
  Dim pos As Integer
  Dim Count As Integer
  Dim Result As Integer

  StSep = Lang(LFrom).StrSep
  
  MySt = Trim(St)
  Result = True

  If Left(MySt, 1) <> StSep Then Result = False
  If Right(MySt, 1) <> StSep Then Result = False
  MySt = Mid(MySt, 2, Len(MySt) - 2)
  
  Count = 0
  For pos = 1 To Len(MySt) - 1
    If Count = 0 Then
      If Mid(MySt, pos, 1) = StSep Then
        Count = 1
        Mid(MySt, pos, 1) = Lang(LTo).StrSep
      End If
    Else
      If Mid(MySt, pos, 1) = StSep Then
        Count = Count + 1
        Mid(MySt, pos, 1) = Lang(LTo).StrSep
      Else
        If Count Mod 2 Then Result = False
        Count = 0
      End If
    End If
  Next

  If Result = True Then
    St = Lang(LTo).StrSep & MySt & Lang(LTo).StrSep
  End If
  
  IsStr = Result

End Function

Sub LoadLanguages ()
  Lang(1).name = "ZX Spectrum Basic"
  Lang(1).FilesName = "SPECTRUM"
  Lang(1).Structured = 0
  Lang(1).EOL = Chr(13) + Chr(10)
  Lang(1).StrSep = Chr(34)
  
  Lang(1).GT = ">"
  Lang(1).GE = ">="
  Lang(1).LT = "<"
  Lang(1).LE = "<="
  Lang(1).NE = "<>"
  Lang(1).EQ = "="
  Lang(1).AND = " AND "
  Lang(1).OR = " OR "
  Lang(1).XOR = " XOR "
  Lang(1).EQV = " AND "
  Lang(1).NOT = " NOT "
  

  Lang(2).name = "Microsoft Visual Basic"
  Lang(2).FilesName = "MSVB"
  Lang(2).Structured = 1
  Lang(2).EOL = Chr(13) + Chr(10)
  Lang(2).StrSep = Chr(34)

  Lang(2).GT = ">"
  Lang(2).GE = ">="
  Lang(2).LT = "<"
  Lang(2).LE = "<="
  Lang(2).NE = "<>"
  Lang(2).EQ = "="
  Lang(2).AND = " AND "
  Lang(2).OR = " OR "
  Lang(2).XOR = " XOR "
  Lang(2).EQV = " AND "
  Lang(2).NOT = " NOT "


  Lang(3).name = "Pascal"
  Lang(3).FilesName = "PASCAL"
  Lang(3).Structured = 1
  Lang(3).EOL = Chr(13) + Chr(10)
  Lang(3).StrSep = "'"

  Lang(3).GT = ">"
  Lang(3).GE = ">="
  Lang(3).LT = "<"
  Lang(3).LE = "<="
  Lang(3).NE = "<>"
  Lang(3).EQ = "="
  Lang(3).AND = " AND "
  Lang(3).OR = " OR "
  Lang(3).XOR = " XOR "
  Lang(3).EQV = " AND "
  Lang(3).NOT = " NOT "
  

  Lang(4).name = "FORTRAN"
  Lang(4).FilesName = "FORTRAN"
  Lang(4).Structured = 0
  Lang(4).EOL = Chr(13) + Chr(10)
  Lang(4).StrSep = Chr(34)

  Lang(4).GT = ".GT."
  Lang(4).GE = ".GE."
  Lang(4).LT = ".LT."
  Lang(4).LE = ".LE."
  Lang(4).NE = ".NE."
  Lang(4).EQ = ".EQ."
  Lang(4).AND = ".AND."
  Lang(4).OR = ".OR."
  Lang(4).XOR = ".NEQV."
  Lang(4).EQV = ".EQV."
  Lang(4).NOT = ".NOT."
  

  Lang(5).name = "C"
  Lang(5).FilesName = "CLANG"
  Lang(5).Structured = 1
  Lang(5).EOL = Chr(13) + Chr(10)
  Lang(5).StrSep = Chr(34)

  Lang(5).GT = ">"
  Lang(5).GE = ">="
  Lang(5).LT = "<"
  Lang(5).LE = "<="
  Lang(5).NE = "!="
  Lang(5).EQ = "=="
  Lang(5).AND = "&"
  Lang(5).OR = "||"
  Lang(5).XOR = " XOR "
  Lang(5).EQV = "&"
  Lang(5).NOT = "!"
  

  Lang(6).name = "Cobol"
  Lang(6).FilesName = "COBOL"
  Lang(6).Structured = 0
  Lang(6).EOL = Chr(13) + Chr(10)
  Lang(6).StrSep = Chr(34)

  Lang(6).GT = ">"
  Lang(6).GE = ">="
  Lang(6).LT = "<"
  Lang(6).LE = "<="
  Lang(6).NE = " NOT EQUAL "
  Lang(6).EQ = "="
  Lang(6).AND = " AND "
  Lang(6).OR = " OR "
  Lang(6).XOR = " XOR "
  Lang(6).EQV = " AND "
  Lang(6).NOT = " NOT "
  


End Sub

Function LogicalExpr (ByVal XPR As String) As String
  Replace XPR, Lang(LFrom).GE, Lang(LTo).GE
  Replace XPR, Lang(LFrom).LE, Lang(LTo).LE
  Replace XPR, Lang(LFrom).NE, Lang(LTo).NE
  Replace XPR, Lang(LFrom).LT, Lang(LTo).LT
  Replace XPR, Lang(LFrom).GT, Lang(LTo).GT
  Replace XPR, Lang(LFrom).EQ, Lang(LTo).EQ
  Replace XPR, Lang(LFrom).AND, Lang(LTo).AND
  Replace XPR, Lang(LFrom).OR, Lang(LTo).OR
  Replace XPR, Lang(LFrom).XOR, Lang(LTo).XOR
  Replace XPR, Lang(LFrom).EQV, Lang(LTo).EQV
  Replace XPR, Lang(LFrom).NOT, Lang(LTo).NOT
  LogicalExpr = XPR
End Function

Sub MathExpr (XPR As String)
  Exit Sub
End Sub

Sub Replace (Text As String, ByVal Find As String, ByVal Repl As String)
  Dim a As Integer
  a = 1

  Do
      a = InStr(a, UCase(Text), UCase(Find))
      If a = 0 Then Exit Sub
      Text = Left(Text, a - 1) & Repl & Mid(Text, a + Len(Find))
      a = a + 1
  Loop

End Sub

Sub StringExpr (XPR As String)
  Exit Sub
End Sub

