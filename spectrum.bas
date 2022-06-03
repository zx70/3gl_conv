Dim ArrayDim() As String

Sub Spectrum ()
 Dim p1 As String
 Dim p2 As String
 Dim p3 As String
 Dim p4 As String
 Dim n1 As Integer
 Dim a, wrd As String
 Dim lcount As Integer
 Dim pos    As Integer
 Dim wpos   As Integer
 Dim IfFlag As Integer
 Dim Lbl As String

 For pass = 1 To 2
 Init
 conv.OBJ.Clear
 If pass = 2 Then Sort: ProgramIni
   MainIni
   For lcount = 0 To conv.Lst.ListCount
     wrd = ""
     b = conv.Lst.List(lcount)
     wpos = 1
     IfFlag = 0
         
     Lbl = GetWord(b, wpos, "")
     If wpos > 0 Then Label (Lbl)
       
     While wpos > 0
       a = GetWord(b, wpos, ":")
       While (CountChars(a, Chr(34)) Mod 2) And wpos > 0
         wpos = wpos + 1
         a = a + ":" + GetWord(b, wpos, ":")
       Wend

       If wpos <> 0 Then Let wpos = wpos + 1
         
       pos = 1

       'While pos > 0
            wrd = GetWord(a, pos, "")
Beginner:
            If wrd = "GO" Then
                wrd = GetWord(a, pos, "")
                p1 = Trim(GetWord(a, pos, ":"))
                If wrd = "TO" Then
                    Jump (p1)
                Else
                    CallSub (p1)
                End If
                wrd = ""
            End If
            
            If wrd = "RETURN" Then
                SubReturn (Lbl)
                wrd = ""
            End If
            
            If wrd = "LET" Then
                wrd = Trim(GetWord(a, pos, "=")) ' Tolgo "="
                p1 = Mid(a, pos + 1)
                Select Case Right(wrd, 1)
                Case "$"
                  If IsStr(p1) Then
                    Call Assign(wrd, 3, Len(p1) - 2, p1)
                  Else
                    Call Assign(wrd, 3, 0, p1)
                  End If
                Case ")"
                  Call Assign(wrd, 0, 0, p1)
                Case Else
                  If IsNum(p1) Then
                    Call Assign(wrd, 2, Len(p1), p1)
                  Else
                    Call Assign(wrd, 2, 0, p1)
                  End If
                End Select
                wrd = ""
            End If

            If wrd = "DIM" Then
                ReDim ArrayDim(1)
                p1 = Trim(GetWord(a, pos, "("))
                pos = pos + 1
                p2 = Trim(GetWord(a, pos, ")"))
                pos = pos + 1
                n1 = 1
                Do
                  p3 = Trim(GetWord(p2, n1, ","))
                  Call Push(ArrayDim(), p3)
                  If n1 = 0 Then Exit Do
                  n1 = n1 + 1
                Loop
                If Right(p1, 1) = "$" Then
                  Call Array(p1, 3, ArrayDim())
                Else
                  Call Array(p1, 2, ArrayDim())
                End If
                wrd = ""
            End If
            
            If wrd = "FOR" Then
                p1 = GetWord(a, pos, "=")
                pos = pos + 1
                p2 = GetWord(a, pos, "")
                wrd = GetWord(a, pos, "")
                p3 = GetWord(a, pos, "")
                If pos <> 0 Then
                    wrd = GetWord(a, pos, "")
                Else
                    wrd = ""
                End If
                p4 = "1"
                If wrd = "STEP" Then
                    p4 = GetWord(a, pos, "")
                End If
                Call BasicForIni(p1, p2, p3, p4)
                wrd = ""
            End If
    
            If wrd = "NEXT" Then
                BasicForEnd
                wrd = ""
            End If
            
            If wrd = "REM" Then
                Comment (Mid(a, pos))
                If wpos <> 0 Then
                  Comment (Mid(b, wpos))
                End If
                wrd = ""
                wpos = 0
            End If

            If wrd = "IF" Then
                wrd = Mid(a, pos + 1, InStr(pos, a, "THEN") - pos - 2)'Prendo COND
                pos = InStr(pos, a, "THEN") + 5
                IfIni (wrd)
                wrd = GetWord(a, pos, "")
                IfFlag = IfFlag + 1
                GoTo Beginner
            End If

            If wrd = "PRINT" Then
                ZXPrint ((Mid(a, pos + 1)))
                wrd = ""
            End If

       If wrd <> "" And pos <> 0 Then
         CallProcIni (wrd)
         FPut (Mid(a, pos + 1))
         CallProcEnd
       End If
     
     Wend
     If IfFlag > 0 Then
       For x = IfFlag To 1 Step -1
         IfEnd
       Next
     End If
   Next
   MainEnd
 Next
 ProgramEnd

End Sub

Sub ZXPrint (ByVal Text As String)
  Dim wrd As String
  Dim pos As Integer

  If Trim(Text) = "" Then
    CallProcIni ("ZXPRINTNL")
    CallProcEnd
    Exit Sub
  End If

  pos = 1
  While pos < Len(Text) And pos <> 0
    wrd = Trim(GetWord(Text, pos, ";'"))
    If pos <> 0 Then pos = pos + 1
    If Left(wrd, 2) = "AT" Then
      CallProcIni ("ZXPRINTAT")
      FPut (Mid(wrd, 4))
      CallProcEnd
    ElseIf Left(wrd, 3) = "TAB" Then
      CallProcIni ("ZXPRINTTAB")
      FPut (Mid(wrd, 5))
      CallProcEnd
    ElseIf Left(wrd, 4) = "OVER" Then
      CallProcIni ("ZXPRINTOVER")
      FPut (Mid(wrd, 6))
      CallProcEnd
    ElseIf Left(wrd, 7) = "INVERSE" Then
      CallProcIni ("ZXPRINTINVRS")
      FPut (Mid(wrd, 9))
      CallProcEnd
    ElseIf Left(wrd, 5) = "FLASH" Then
      CallProcIni ("ZXPRINTFLASH")
      FPut (Mid(wrd, 7))
      CallProcEnd
    ElseIf Left(wrd, 3) = "INK" Then
      CallProcIni ("ZXPRINTINK")
      FPut (Mid(wrd, 5))
      CallProcEnd
    ElseIf Left(wrd, 5) = "PAPER" Then
      CallProcIni ("ZXPRINTPAPER")
      FPut (Mid(wrd, 7))
      CallProcEnd
    Else
      CallProcIni ("ZXPRINT")
      If Left(wrd, 1) = """" Then
        FPut Lang(LTo).StrSep & Mid(wrd, 2, Len(wrd) - 2) & Lang(LTo).StrSep
      Else
        FPut (wrd)
      End If
      CallProcEnd
    End If
    
    If pos <> 0 Then
      If Mid(Text, pos, 1) = "'" Then
          CallProcIni ("ZXPRINTNL")
          CallProcEnd
      End If
    End If
  Wend
End Sub

