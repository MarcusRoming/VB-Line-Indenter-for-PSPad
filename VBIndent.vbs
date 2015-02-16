'*************************************************************************************
' Simple VB-Script Code indenter, from Marcus Roming with Code from gogogadgetscott
' Description   : Simple code beautifier / indenter for Visual Basic Sript
' Version       : 1.3
' Date          : 04.02.15
'*************************************************************************************

Option Explicit
Const module_name   = "VBIndent"
Const module_ver    = "1.30"
Const ConTabLen     = 4


Sub VBIndent
    Dim Columns(100)
    Dim strText, EOL, line, lines, i, intSpace, intCnt, strSpaces, intThenPos, strPastThen, intMsgBoxSelection, intErrCnt, intSpaceCnt, strTestLine
    
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
        line = Trim(LTrimEx(line))                         'Remove spaces and Tabs
        strTestLine = line
        
        Do                                                          'Remove all unescessary spaces in test string
            strTestLine = Replace(strTestLine,"  "," ")
            intSpaceCnt = InStr(strTestLine,"  ")
        Loop Until intSpaceCnt = 0
        
        ' In the following the elements that are closing a block...
        
        If UCase(Left(strTestLine,4)) = "END " Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(strTestLine,5)) = "NEXT " OR UCase(strTestLine) = "NEXT" Then
            intSpace = intSpace - ConTabLen
        End If
        
        
        If UCase(Left(strTestLine,5)) = "LOOP " OR UCase(strTestLine) = "LOOP" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(strTestLine,5)) = "WEND " OR UCase(strTestLine) = "WEND" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(strTestLine,4)) = "ELSE" OR UCase(Left(strTestLine,6)) = "ELSEIF" Then
            intSpace = intSpace - ConTabLen
        End If
        
        If UCase(Left(strTestLine,5)) = "CASE " Then
            intSpace = intSpace - ConTabLen
        End If
        
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
        If UCase(Left(strTestLine,4)) = "SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,6)) = "CLASS " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,9)) = "PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,9)) = "FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,13)) = "PUBLIC CLASS " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,14)) = "PRIVATE CLASS " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,17)) = "PRIVATE FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,16)) = "PUBLIC FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,17)) = "PRIVATE PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,16)) = "PUBLIC PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,24)) = "PUBLIC DEFAULT PROPERTY " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,24)) = "PUBLIC DEFAULT FUNCTION " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,19)) = "PUBLIC DEFAULT SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,12)) = "PRIVATE SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,11)) = "PUBLIC SUB " Then
            intSpace = intSpace + ConTabLen
        End If
        
        
        If UCase(Left(strTestLine,3)) = "IF " OR UCase(Left(strTestLine,3)) = "IF(" Then
            If Right(strTestLine,1) = "_" Then
                intSpace = intSpace + ConTabLen
            Else
                intThenPos = InStr(UCase(Trim(strTestLine))," THEN")
                If len(Trim(strTestLine)) < intThenPos + 6 Then            'Test if the THEN-Command is a Single line command
                    intSpace = intSpace + ConTabLen                 'No, more than one line!
                Else
                    'Differentiate bewtween single line command and a following comment!
                    strPastThen = Trim(Right(strTestLine,len(strTestLine)- intThenPos -5))
                    
                    If Left(strPastThen,1) = "'" Then
                        intSpace = intSpace + ConTabLen             'Not a single line command!
                    End If
                End If
            End If
        End If
        
        If UCase(Left(strTestLine,4)) = "FOR " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,3)) = "DO " OR UCase(strTestLine) = "DO" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,6)) = "WHILE " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,5)) = "WITH " Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,11)) = "SELECT CASE" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,4)) = "ELSE" OR UCase(Left(strTestLine,6)) = "ELSEIF" Then
            intSpace = intSpace + ConTabLen
        End If
        
        If UCase(Left(strTestLine,5)) = "CASE " Then
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

Function LTrimEx(str)
    Dim re
    Set re = New RegExp
    re.Pattern = "^\s*"
    re.Multiline = False
    LTrimEx = re.Replace(str, "")
End Function

Sub Init
    addMenuItem "VBScriptIndent", "Format code", "VBIndent"
End Sub