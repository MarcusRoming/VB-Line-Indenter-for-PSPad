'*************************************************************************************
' Simple VB-Script Code indenter, from Marcus Roming with Code from gogogadgetscott
' Description   : Simple code beautifier / indenter for Visual Basic Sript
' Version       : 1.00
' Date          : 30.01.15
'*************************************************************************************

Option Explicit
Const module_name   = "VBIndent"
Const module_ver    = "1.00"
Const ConTabLen     = 4


Sub VBIndent
    Dim Columns(100)
    Dim strText, EOL, line, lines, i, intSpace, intCnt, strSpaces, intTh3nPos, strPastThen, intMsgBoxSelection, intErrCnt
    
    intMsgBoxSelection = MsgBox ("Format the code?", vbYesNo+vbQuestion, "Info:")
    intErrCnt = 0
    
    If intMsgBoxSelection = vbNo Then Exit Sub
    '// Get working strText
    strText = handleSelText("")
    
    '// Determine end-of-line
    EOL = ""
    If InStr(strText, Chr(13)) Then
        EOL = EOL & Chr(13)
    End If
    If InStr(strText, Chr(10)) Then
        EOL = EOL & Chr(10)
    End If
    
    '// Get lines
    lines = Split(strText, EOL)
    
    '// Initialize line index
    i = -1
    
    '// Added spacing
    intSpace = 0
    For Each line in lines
        i=i+1
        line = Trim(Replace(line,vbTab,""))                         'Remove spaces and Tabs
        
        ' In the following the elements that are closing a block...
        If UCase(Left(line,7)) = "END SUB" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,12)) = "END FUNCTION" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,12)) = "END PROPERTY" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,9)) = "END CLASS" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,8)) = "END WITH" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,10)) = "END SELECT" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,6)) = "END IF" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,5)) = "NEXT " OR UCase(line) = "NEXT" Then
            intSpace = intSpace - ConTabLen
        End If
        
        
        If UCase(Left(line,5)) = "LOOP " OR UCase(line) = "LOOP" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,5)) = "WEND " OR UCase(line) = "WEND" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,4)) = "ELSE" OR UCase(Left(line,6)) = "ELSEIF" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(line,5)) = "CASE " Then
            intSpace = intSpace - ConTabLen
        End If
        
        'MsgBox CStr(intSpace)
        strSpaces = ""
        If intSpace < 0 Then
            intSpace = 0
            intErrCnt = intErrCnt + 1
            MsgBox "Possible error in line " & CStr(i+1) & " : " & Chr(34) & line & Chr(34) & vbNewLine & "Counterpart of closing declaration not found!", vbExclamation, "Error:"
        End If
        
        For intCnt = 1 To intSpace
            strSpaces = strSpaces & " "    'Create the appropritate number of spaces to be added in front of the line!
        Next
        
        ' In the following the elements that are opening a block...
        If UCase(Left(line,4)) = "SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,6)) = "CLASS " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,13)) = "PUBLIC CLASS " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,14)) = "PRIVATE CLASS " Then
            intSpace = intSpace + ConTabLen
        End If        
        
        If UCase(Left(line,9)) = "FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,17)) = "PRIVATE FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,17)) = "PRIVATE PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,16)) = "PUBLIC FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,16)) = "PUBLIC PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,24)) = "PUBLIC DEFAULT PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,24)) = "PUBLIC DEFAULT FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,19)) = "PUBLIC DEFAULT SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,12)) = "PRIVATE SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,11)) = "PUBLIC SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,3)) = "IF " Then
            intTh3nPos = InStr(UCase(Trim(line))," THEN")
            If len(Trim(line)) < intTh3nPos + 6 Then            'Test if the THEN-Command is a Single line command
                intSpace = intSpace + ConTabLen                 'No, more than one line!
            Else
                'Differentiate bewtween single line command and a following comment!
                strPastThen = Trim(Right(line,len(line)- intTh3nPos -5))
                
                If Left(strPastThen,1) = "'" Then
                    intSpace = intSpace + ConTabLen             'Not a single line command!
                End If
            End If
        End If
        
        If UCase(Left(line,4)) = "FOR " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,3)) = "DO " OR UCase(line) = "DO" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,6)) = "WHILE " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,5)) = "WITH " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,11)) = "SELECT CASE" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,4)) = "ELSE" OR UCase(Left(line,6)) = "ELSEIF" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(line,5)) = "CASE " Then
            intSpace = intSpace + ConTabLen
        End If
        
        lines(i) = strSpaces & line
    Next
    
    '// Replace text
    strText = Join(lines, EOL)
    If intSpace > 0 Then
        MsgBox "Possible error: There may be " & CInt(intSpace\4) & " unclosed blocks!", vbExclamation
        intErrCnt = intErrCnt + 1
    End If
    
    handleSelText strText
    
    If intErrCnt > 0 Then
        MsgBox "Done! Detected " & CStr(intErrCnt) & " Errors", vbExclamation, "Info:"
    Else
        MsgBox "Done! Detected no errors!", vbInformation, "Info:"
    End If
    
End Sub

'// @param string Text to replace selected text
Private Function handleSelText(strText)
    Dim editor
    On Error Resume Next
    Set editor = newEditor()
    editor.assignActiveEditor
    If strText = "" Then
        '// Get selected text
        handleSelText = editor.selText
        If handleSelText = "" Then
            '// No text was select. Get all text and select it.
            handleSelText  = editor.Text
            editor.command "ecSelectAll"
        End If
    Else
        '// Set selected text
        editor.selText strText
    End If
End Function

Sub Init
    addMenuItem "VBIndent", "Format code", "VBIndent"
End Sub