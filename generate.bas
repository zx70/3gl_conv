Dim LineBuff As String
Global pass As Integer
Global Expression As String
Dim ForStack() As String
Dim ForLblStack() As String
Dim ZxLineNum As Integer
Dim Indent As Integer
Dim LblCount As Long

Sub Assign (ByVal VName As String, VType As Integer, Vlen As Integer, VVal As String)

VName = Trim(VName)

' VType:    0 - Not Defined
'           1 - Integer
'           2 - Number
'           3 - String

' Vlen      0 - Is an expression (or unknown size)

 If pass = 1 And VType <> 0 Then
   conv.VList.AddItem (VName) & " " & Vlen
   conv.VList.ItemData(conv.VList.NewIndex) = VType
 End If

 Select Case LTo
  Case 1 'Spectrum
    FPut "LET " & VName & " = " & VVal
  Case 2 'VB
    FPut VName & " = " & VVal
    FLine
  Case 3 'Pascal
    FPut VName & " := " & VVal & ";"
    FLine
  Case 4 'FORTRAN
    FPut "      " & VName & " = " & VVal
    FLine
  Case 6 'COBOL
    If Vlen = 0 Then
      FPut "           COMPUTE " & VName & " = " & VVal
    Else
      FPut "           MOVE " & VVal & " TO " & VName
    End If
    FLine
 End Select

End Sub


Sub BasicForEnd ()
Dim VarName As String
Dim LblFor As String

 LblFor = POP(ForStack())
 VarName = POP(ForStack())
 DecIndent
 
 Select Case LTo
  Case 1 'Spectrum
    FPut "NEXT " & VarName & " "
    
  Case 2 'VB
    FPut "Next "
    FLine
  Case 3 'Pascal
    FPut "End;"
    FLine
  
  Case 4 'FORTRAN
    FPut Trim(LblFor) & Space(6 - Len(LblFor)) & "CONTINUE"
    FLine
 
  Case 5 'C
    FPut "}"
    FLine

  Case 6 'Cobol
    FPut "       " & LblFor & "-EX."
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
    FPut "           EXIT."
    FLine
    FPut "       " & LblFor & "-NXT."
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  End Select

End Sub

Sub BasicForIni (ByVal VarName As String, ByVal VFrom As String, ByVal VTo As String, ByVal VStep As String)
 Dim LblFor As String
 Dim LMax As Integer
 Dim LTyper As Integer

 If Len(VFrom) > Len(VTo) Then
   LMax = Len(VFrom)
 Else
   LMax = Len(VTo)
 End If

 'If InStr(VFrom, ".") = 0 And InStr(VTo, ".") = 0 And InStr(VStep, ".") = 0 Then
 '  LType = 1 'Integer
 'Else
 '  LType = 2 'Real
 'End If
 LType = 2 'Real

 If pass = 1 Then
   conv.VList.AddItem (VarName) & " " & LMax
   conv.VList.ItemData(conv.VList.NewIndex) = LType
 End If
 
 Call Push(ForStack(), VarName)
 LblFor = GetSysLbl()
 Call Push(ForStack(), LblFor)

 Select Case LTo
  Case 1 'Spectrum
    FPut "FOR " & VarName & " = " & VFrom & " TO " & VTo & " "
    If trimVStep <> "1" Then
      FPut "STEP " & VStep & " "
    End If
  
  Case 2 'VB
    FPut "For " & VarName & " = " & VFrom & " TO " & VTo & " "
    If VStep <> "1" Then
      FPut "STEP " & VStep & " "
    End If
    FLine
    
  Case 3 'Pascal
    FPut VarName & " := " & VFrom & ";"
    FLine
    FPut "While " & VarName & " <> " & VTo & " DO BEGIN"
    FLine
    If Left(VStep, 1) <> "-" Then
      FPut VarName & " := " & VarName & " + " & VStep & ";"
    Else
      FPut VarName & " := " & VarName & " " & VStep & ";"
    End If
    FLine
  
  Case 4 'FORTRAN
    If VStep <> "1" Then
        FPut "      DO " & LblFor & " " & VarName & " = " & VFrom & "," & VTo & "," & VStep
        FLine
    Else
        FPut "      DO " & LblFor & " " & VarName & " = " & VFrom & "," & VTo
        FLine
    End If
 
  Case 5 'C
    If Left(VStep, 1) <> "-" Then
      FPut "for (" & VarName & " == " & VFrom & " ; " & VarName & "<" & VTo & " ; " & VarName & " == " & VarName & " + " & VStep & ")"
    Else
      FPut "for (" & VarName & " == " & VFrom & " ; " & VarName & ">" & VTo & " ; " & VarName & " == " & VarName & VStep & ")"
    End If
    FLine
    FPut "{"
    FLine

  Case 6 'Cobol
    FPut "           PERFORM " & LblFor & " THRU " & LblFor & "-EX"
    FLine
    FPut "           VARYING " & VarName & " FROM " & VFrom
    If VStep <> "1" Then
    FPut " BY " & VStep
    End If
    FLine
    FPut "           UNTIL " & VarName & " = " & VTo & "."
    FLine
    FPut "           GO TO " & LblFor & "-NXT."
    FLine
    FPut "       " & LblFor & "."
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
 End Select
 
 IncIndent

End Sub

Sub CallProcEnd ()
Select Case LTo
  Case 1 'Spectrum
    FLine
  Case 2 'VB
    FPut " "
    FLine
  Case 3 'Pascal
    FPut ");"
    FLine
  Case 4 'FORTRAN
    FPut ")"
    FLine
  Case 6 'Cobol
    FLine

End Select
End Sub

Sub CallProcIni (ByVal ProcName As String)
Select Case LTo
  Case 1 'Spectrum
    FPut ProcName & " "
  Case 2 'VB
    FPut ProcName & " "
  Case 3 'Pascal
    FPut ProcName & " ("
  Case 4 'FORTRAN
    FPut "      CALL " & ProcName & "("
  Case 6 'Cobol
    FPut "           " & ProcName & " "
End Select
End Sub

Sub CallSub (ByVal Lbl As String)
 If pass = 1 Then
   conv.LList.AddItem (Lbl)
 End If
 Select Case LTo
  Case 1 'Spectrum
    FPut "GO SUB " & Lbl
    FLine
  Case 2 'VB
    FPut "GoSub " & "L" & Lbl
    FLine
  Case 3 'Pascal
    FPut "goto " & "(L" & Lbl & ");"
    FLine
  Case 4 'FORTRAN
    FPut "      CALL " & "L" & Lbl & " ()"
    FLine
  Case 6 'Cobol
    If pass = 2 Then
      FPut "           PERFORM " & "L" & Lbl & " THRU " & "L" & Lbl
    Else
      FPut "           -- CALL: " & "L" & Lbl & " --"
    End If
    FLine
  End Select
End Sub

Sub Comment (ByVal Text As String)
 Select Case LTo
  Case 1 'Spectrum
    FPut "REM " & Text
    FLine
  Case 2 'VB
    FPut "' " & Text
    FLine
  Case 3 'Pascal
    FPut "{ " & Text & " }"
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 4 'FORTRAN
    FPut "C     " & Text
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 5 'C
    FPut "// " & Text
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 6 'Cobol
    FPut "      *" & Text
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
 End Select


End Sub

Sub DecIndent ()
  Indent = Indent - 2
End Sub

Sub ExprText (Text)
  Expression = Expression + Text
End Sub

Sub FLine ()
  If Trim(LineBuff) <> "" Then
     ' Il Fortran e il Sinclair Basic non li indento
     If LTo <> 4 And LTo <> 1 Then
       LineBuff = Space(Indent) & LineBuff
     End If
     
     If LTo = 1 Then 'Spectrum ?
       LineBuff = ZxLineNum & Space(6 - Len(Str(ZxLineNum))) & LineBuff
     End If
     conv.OBJ.AddItem (LineBuff)
     ZxLineNum = ZxLineNum + 5
  End If
  LineBuff = ""
End Sub

Sub FPut (ByVal testo As String)
  LineBuff = LineBuff + testo
End Sub

Function GetSysLbl () As String
 
 Select Case LTo
  Case 1 'Spectrum
    
  Case 2 'VB
    GetSysLbl = "X" & Trim(Str(LblCount))
    LblCount = LblCount + 1
  Case 3 'Pascal
    GetSysLbl = "X" & Trim(Str(LblCount))
    LblCount = LblCount + 1
  Case 4 'FORTRAN
    If LblCount = 0 Then LblCount = 50000
    GetSysLbl = Trim(Str(LblCount))
    LblCount = LblCount + 10
  Case 5 'C
    GetSysLbl = "X" & Trim(Str(LblCount))
    LblCount = LblCount + 1
  Case 6 'Cobol
    GetSysLbl = "X" & Trim(Str(LblCount))
    LblCount = LblCount + 1
 End Select

 
End Function

Sub IfEnd ()

 DecIndent
 
 Select Case LTo
  Case 1 'Spectrum
    
  Case 2 'VB
    FPut "End If"
    FLine
  Case 3 'Pascal
    FPut "End;"
    FLine
  Case 4 'FORTRAN
    FPut "      ENDIF"
    FLine
  Case 5 'C
    FPut "}"
    FLine
  Case 6 'COBOL
    FPut "           END-IF."
    FLine
 End Select

End Sub

Sub IfIni (ByVal Cond As String)
 
 Dim Condition As String
 
 Condition = LogicalExpr(Cond)

 Select Case LTo
  Case 1 'Spectrum
    FPut "IF " & Condition & " THEN "
  Case 2 'VB
    FPut "If " & Condition & " Then "
    FLine
  Case 3 'Pascal
    FPut "If " & Condition & " Then Begin"
    FLine
  Case 4 'FORTRAN
    FPut "      IF (" & Condition & ") THEN "
    FLine
  Case 5 'C
    FPut "if (" & Condition & ")"
    FLine
    FPut "{"
    FLine
  Case 6 'Cobol
    FPut "           IF " & Condition
    FLine
 End Select
 
 IncIndent

End Sub

Sub IncIndent ()
  Indent = Indent + 2
End Sub

Sub Init ()
  ReDim ForStack(1)
  ReDim ForLblStack(1)
  ZxLineNum = 5
  Indent = 2
  LblCount = 0
End Sub

Sub Jump (ByVal Lbl As String)
 If pass = 1 Then
   conv.LList.AddItem (Lbl)
 End If
 Select Case LTo
  Case 1 'Spectrum
    FPut "GO TO " & Lbl
    FLine
  Case 2 'VB
    FPut "GoTo " & "L" & Lbl
    FLine
  Case 3 'Pascal
    FPut "GoTo " & "L" & Lbl & ";"
    FLine
  Case 4 'FORTRAN
    FPut "      GO TO " & Lbl
    FLine
  Case 6 'Cobol
    FPut "           GO TO " & "L" & Lbl
    FLine
 End Select
End Sub

Sub Label (ByVal LName As String)

If pass = 2 And Not LabelUsed(LName) Then
    Exit Sub
End If

Select Case LTo
  Case 1 'ZX
    FPut "-#-" & LName & "-#-"
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 2 'VB
    FPut "L" & LName & ":  "
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 3 'Pascal
    FPut "L" & LName & ":  "
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 4 'FORTRAN
    FPut LName & Space(6 - Len(LName))
    FPut "CnvFOO = 0"
    FLine
  Case 5 'C
    FPut "L" & LName & ":  "
    conv.OBJ.AddItem (LineBuff)
    LineBuff = ""
  Case 6 'Cobol
    FPut "       L" & LName & "."
End Select


End Sub

Sub LabelDeclare ()

Dim x As Integer
Dim Lbl As String

  IncIndent
  For x = 0 To conv.LList.ListCount - 1
    Lbl = Trim(UCase(conv.LList.List(x)))

     Select Case LTo
      
      Case 1 'Spectrum
      
      Case 2 'VB
        
      Case 3 'Pascal
        FPut "L" & Lbl
        If x = conv.LList.ListCount - 1 Then
          FPut ";"
        Else
          FPut ","
        End If
        FLine
      
      Case 4 'FORTRAN
       
      Case 5 'C
    
      Case 6 'Cobol
     
     End Select
  Next
  DecIndent

End Sub

Function LabelUsed (LName) As Integer
  Dim x As Integer
  Dim Result As Integer

  For x = 0 To conv.LList.ListCount
    If UCase(LName) = UCase(conv.LList.List(x)) Then
      LabelUsed = True
      Exit Function
    End If
  Next
  
  LabelUsed = False

End Function

Sub MainEnd ()
 
 
 Select Case LTo
  Case 3 'Pascal
    DecIndent
    FPut "End."
    FLine
  Case 4 'FORTRAN
    DecIndent
    FPut "      END"
    FLine
  Case 5 'C
    DecIndent
    FPut "}"
    FLine
  Case 6 'COBOL
    FPut "           STOP RUN."
    FLine
 End Select
 
End Sub

Sub MainIni ()
 Select Case LTo
  Case 3 'Pascal
    FPut "Begin"
    FLine
    IncIndent
  Case 5 'C
    FPut "Int Main () "
    FLine
    FPut "{"
    FLine
    IncIndent
  Case 6 'Cobol
    FPut "       PROGRAM SECTION."
    FLine
 End Select
 
End Sub

Function POP (Arr() As String) As String
  Dim x As Integer

  x = UBound(Arr)
  POP = Arr(x)
  ReDim Preserve Arr(x - 1)
  
End Function

Sub ProgramEnd ()

If pass = 2 Then
 Select Case LTo
  Case 1 'Spectrum
  FPut "STOP"
  FLine

  Case 2 'VB
  'FPut "End"
  'FLine
    
  Case 3 'Pascal
  FPut "End."
  FLine

  Case 4 'FORTRAN
  FPut "      END"
  FLine

  Case 5 'C

  Case 6 'Cobol
  FPut "           STOP RUN."
  conv.OBJ.AddItem (LineBuff)
 
 End Select

End If

End Sub

Sub ProgramIni ()
Dim PrgName As String
Dim Pos As Integer

PrgName = conv.CMDialog.Filetitle & "."
Pos = InStr(PrgName, ".") - 1
PrgName = Left(PrgName, Pos)

If pass = 2 Then
 Select Case LTo
  
  Case 1 'Spectrum
  
  Case 2 'VB
    
  Case 3 'Pascal
  FPut "Program " & PrgName & ";"
  FLine
  If conv.VList.ListCount > 0 Then
      FPut "Var"
      FLine
      VarDeclare
  End If
  If conv.LList.ListCount > 0 Then
      FPut "Label"
      FLine
      LabelDeclare
  End If

  Case 4 'FORTRAN
  FPut "C     -- Conv support --"
  FLine
  FPut "      INCLUDE      '" & Lang(LFrom).FilesName & ".FI'"
  FLine
  FPut "C     -- ** --"
  FLine
  FPut "      PROGRAM " & PrgName
  FLine
  VarDeclare
  FPut "C     -- Conv support --"
  FLine
  FPut "      INCLUDE      '" & Lang(LFrom).FilesName & ".FD'"
  FLine
  FPut "C     -- ** --"
  FLine

  Case 5 'C

  Case 6 'Cobol
 
 End Select
End If

End Sub

Sub Push (Arr() As String, ByVal Value As String)
  Dim x As Integer

  x = UBound(Arr)
  ReDim Preserve Arr(x + 1)
  Arr(x + 1) = Value
End Sub

Sub Sort ()
 Dim v As String
 Dim x As Integer
 Dim y As Integer

  x = 0
  While x < conv.LList.ListCount
    While UCase(conv.LList.List(x)) = UCase(conv.LList.List(x + 1))
      conv.LList.RemoveItem (x + 1)
    Wend
    x = x + 1
  Wend

  x = 0
  While x < conv.VList.ListCount
    i1 = UCase(conv.VList.List(x))
    i2 = UCase(conv.VList.List(x + 1))
    If Trim(Mid(i1, 1, InStr(i1, " "))) = Trim(Mid(i2, 1, InStr(i2, " "))) Then
      If Val(Mid(i1, InStr(i1, " "))) > Val(Mid(i2, InStr(i2, " "))) Then
        conv.VList.RemoveItem (x + 1)
      Else
        conv.VList.RemoveItem (x)
      End If
    Else
      x = x + 1
    End If
  Wend

  x = 0
  While x < conv.AList.ListCount
    While UCase(conv.AList.List(x)) = UCase(conv.AList.List(x + 1))
      conv.AList.RemoveItem (x + 1)
    Wend
    x = x + 1
  Wend

End Sub

Sub SubReturn (ByVal Lbl As String)
 If pass = 1 Then
   conv.LList.AddItem (Lbl)
 End If
 
 Select Case LTo
  Case 1 'Spectrum
    FPut "RETURN "
    FLine
  Case 2 'VB
    FPut "Return"
    FLine
  Case 3 'Pascal
    FPut "Return;"
    FLine
  Case 4 'FORTRAN
    FPut "      RETURN"
    FLine
  Case 6 'Cobol
    FPut "           EXIT."
    FLine
 End Select

End Sub

Sub VarDeclare ()
Dim x As Integer
Dim Var As String
Dim VName As String
Dim VType As Integer
Dim Vlen As Integer
' VType:    0 - Not Defined
'           1 - Integer
'           2 - Number
'           3 - String

  IncIndent
  For x = 0 To conv.VList.ListCount - 1
    Var = UCase(conv.VList.List(x))
    VName = Trim(Mid(Var, 1, InStr(Var, " ")))
    Vlen = Val(Mid(Var, InStr(Var, " ")))
    VType = conv.VList.ItemData(x)

     Select Case LTo
      
      Case 1 'Spectrum
      
      Case 2 'VB
        
      Case 3 'Pascal
        Select Case VType
        Case 1
          FPut VName & " : Integer;"
          FLine
        Case 2
          FPut VName & " : Real;"
          FLine
        Case 3
          If Vlen = 0 Then
              FPut VName & " : String[255];"
              FLine
          Else
              FPut VName & " : String[" & Vlen & "];"
              FLine
          End If

        End Select
      Case 4 'FORTRAN
        Select Case VType
        Case 1
          FPut "      INTEGER " & VName
          FLine
        Case 2
          FPut "      REAL " & VName
          FLine
        Case 3
          If Vlen = 0 Then
              FPut "      CHARACTER * 255 " & VName
              FLine
          Else
              FPut "      CHARACTER * " & Vlen & " " & VName
              FLine
          End If

        End Select
     
      Case 5 'C
    
      Case 6 'Cobol
     
     End Select
  Next
  DecIndent

End Sub

