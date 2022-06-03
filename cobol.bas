
Sub CloseDivision (Division)
  Select Case Division
    Case "ENVIRONMENT"
      Comment "Ending Environment"
    Case "DECLARATION"
      Comment "Ending Declaration"
    
  End Select
End Sub

Sub Cobol ()
 Dim b As String
 Dim Directive As String
 Dim Section As String
 Dim Division As String
 Dim WSCount As String
 Dim WSName As String
 Dim WSType As String
 Dim WSVal As String
 Dim pos As Integer

 For pass = 1 To 2
   Init
   conv.OBJ.Clear
   If pass = 2 Then Sort: ProgramIni
   MainIni
   For lcount = 0 To conv.Lst.ListCount
     b = conv.Lst.List(lcount)
     
     If Len(b) > 6 Then
        If Mid(b, 7, 1) = "*" Then
          Comment (Mid(b, 7))
          GoTo RigaSuccessiva
        ElseIf Mid(b, 7, 1) = "/" Then
          GoTo RigaSuccessiva
        End If
     End If
     
     Select Case CobolDirective(b, Directive)
       Case 1
         CloseDivision (Division)
         Division = Directive
         GoTo RigaSuccessiva
       Case 2
         Section = Directive
         GoTo RigaSuccessiva
     End Select

     Select Case Division
     Case "IDENTIFICATION"
       Comment (Mid(b, 6))
     Case "ENVIRONMENT"
       Comment (Mid(b, 6))
     Case "DATA"
       Select Case Section
       Case "WORKING-STORAGE"
         pos = 6
         WSCount = GetWord(b, pos, "")
         WSName = GetWord(b, pos, "")
         If pos > 0 Then
           WSType = UCase(GetWord(b, pos, ""))
           If WSType = "PIC" Then
             WSType = UCase(GetWord(b, pos, ""))
             Select Case Left(WSType, 1)
               Case "9"
                 Call DeclVar(WSName, 2)
               Case "X"
                 If pos > 0 Then
                   If UCase(GetWord(b, pos, "")) = "VALUE" Then
                     Call Assign(WSName, 3, 0, GetWord(b, pos, ""))
                   End If
                 Else
                 Call DeclVar(WSName, 3)
                 End If
             End Select
           End If
         End If
         FLine
       End Select
     End Select

RigaSuccessiva:
   Next
 Next
 
 ProgramEnd

End Sub

Function CobolDirective (ByVal Ln As String, Directive As String) As Integer

  Dim p1 As String
  Dim p2 As String
  Dim pos As Integer

  pos = 1
  p1 = GetWord(Ln, pos, "")
  If pos > 0 Then
      p2 = GetWord(Ln, pos, "")
    
      If UCase(p2) = "DIVISION" Or UCase(p2) = "DIVISION." Then
        CobolDirective = 1
        Directive = UCase(p1)
      ElseIf UCase(p2) = "SECTION" Or UCase(p2) = "SECTION." Then
        CobolDirective = 2
        Directive = UCase(p1)
      Else
        CobolDirective = 0
        Directive = ""
      End If
  Else
      CobolDirective = 0
      Directive = ""
  End If

End Function

