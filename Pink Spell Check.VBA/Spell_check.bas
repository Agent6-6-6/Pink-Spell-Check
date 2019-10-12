Attribute VB_Name = "Spell_check"

Private Sub Pink_Spell_Check_onAction(control As IRibbonControl)

'SETUP ERRORHANDLER FOR USER CANCELLING THE SPELLCHECK
    Application.EnableCancelKey = xlErrorHandler

    'INITIALISE VARIABLES
    Dim WS As Worksheet
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    Dim DisplayTime As Integer
    Dim Scope_msg As String

    'INITIALISE NUMBER OF SECONDS TO DISPLAY STATUSBAR MESSAGE
    DisplayTime = 10

    'CHECK FOR PROTECTED SHEETS AND EXIT IF FOUND
    For Each WS In ActiveWorkbook.Worksheets
        If WS.ProtectContents = True Then
            MsgBox "There are protected Sheets in this Workbook. Please unprotect the Sheets to run this macro"
            Exit Sub
        End If
    Next

    'SHOW 'Spelling_form' USERFORM
    Spelling_form.Show
    DoEvents

    'STORE TIME WHEN MACRO STARTS
    StartTime = timer

    If Spelling_form.check_all = True Then
        'CALL SPELLING HIGHLIGHT MACRO WITH 'ActiveSheet.UsedRange' AS SCOPE
        CheckSpelling ActiveSheet.UsedRange
        Scope_msg = "current Sheet"
    Else
        'CALL SPELLING HIGHLIGHT MACRO WITH CURRENT 'Selection' AS SCOPE
        CheckSpelling Selection
        Scope_msg = "current Selection"
    End If

    'DETERMINE HOW MANY SECONDS THE CODE TOOK TO RUN
    SecondsElapsed = Round(timer - StartTime, 2)
    'DISPLAY FINISH MESSAGE IN THE STATUSBAR
    Application.StatusBar = "Highlight of spelling within " & Scope_msg & " completed successfully. Check took " & SecondsElapsed & " seconds"
    DoEvents
    'CLEAR STATUSBAR AFTER 'DisplayTime' SECONDS
    Application.OnTime Now + TimeSerial(0, 0, DisplayTime), "ClearStatusBar_Spelling"

handleCancel:
    If Err = 18 Then
        MsgBox "Spellcheck Cancelled"
    End If


End Sub

Sub CheckSpelling(R As Range)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo handleCancel

    'SETUP ERRORHANDLER FOR USER CANCELLING THE SPELLCHECK ROUTINE
    Application.EnableCancelKey = xlErrorHandler

    'FORMATS MISSPELLED WORDS IN PINK TEXT IN SUPPLIED RANGE 'R'
    '________________________________________________________________________________________________________________
    'Will highlight all incorrect spelling detected by standard spellchecking dialog in excel, its intended that user runs this
    'from time to time as an aid to spotting/correcting spelling mistakes, either correcting them manually or via the built in spell
    'checking dialog as it can be hard to spot the suggested correction required in cells with larger volume of text

    'Specifically checks the following scenarios and applies highlighting:-
    ' - Custom phrases (refer sheet 1 for entry)
    ' - Repeated words
    ' - Misspelt words
    '________________________________________________________________________________________________________________

    'INITIALISE VARIABLES
    Dim cell As Range
    Dim words() As String    'ARRAY FOR SPLIT WORDS FROM EACH CELL
    Dim cell_string As String
    Dim i_words As Long    'SEQUENTIAL NUMBER DESCRIBING THE SEQUENTIAL POSITION FOR EACH WORD WITHIN THE ORIGINAL CELL TEXT
    Dim i_chars As Long    'SEQUENTIAL NUMBER DESCRIBING THE SEQUENTIAL POSITION FOR EACH CHARACTER WITHIN THE ORIGINAL CELL TEXT
    Dim term As Range
    Dim word_total As Long
    Dim current_word As String
    Dim next_word As String
    Dim current_len As Long
    Dim next_len As Long
    Dim combined_len As Long
    Dim correctly_spelled_words As String    'CACHE FOR PREVIOUSLY CORRECTLY SPELLED WORDS
    Dim highlight_colour
    Dim I As Double
    Dim J As Double
    Dim start As Long
    Dim pos As Long
    Dim tempS As String

    'DEFINE DEFAULT HIGHLIGHT COLOUR
    highlight_colour = RGB(255, 0, 255)

    'SETUP OPTIONS FOR CHECKING SPELLING (FROM 'Spelling_form')
    With Application.SpellingOptions
        .IgnoreCaps = Not Spelling_form.Check_all_capitals.Value    'WILL HIGHLIGHT MISSPELLED WORDS ALL IN CAPITALS WHEN SET TO FALSE
        .IgnoreFileNames = Not Spelling_form.Check_filenames.Value    'WILL HIGHLIGHT MISSPELLED FILENAMES WHEN SET TO FALSE
        .IgnoreMixedDigits = Not Spelling_form.Check_mixed_digits.Value    'WILL HIGHLIGHT WORDS WITH MIXED ALPHA AND NUMERIC CHARACTERS (i.e AS/NZS1170 for example)
    End With

    'LOOP THROUGH EACH CELL IN RANGE 'R'
    For Each cell In R
        'IGNORE ANY NON QUALIFYING CELLS CONTAINING FORMULA, NUMERICAL VALUES, OR WHICH ARE EMPTY
        If Not cell.HasFormula And Not IsNumeric(cell.Value) Or IsEmpty(cell.Value) Then
            'GET CONTENTS OF CELL AS A STRING
            cell_string = cell.Value
            'REPLACES 'Alt-Enter' LINE BREAKS WITHIN A CELL STRING
            cell_string = Replace(cell_string, Chr(10), " ")
            'REPLACES NON-BREAKING SPACES WITHIN A CELL STRING
            cell_string = Replace(cell_string, Chr(160), " ")
            'REPLACES DIVIDING FORWARD SLASHES, ALLOWING CHECKING OF 'word1/word2' SCENARIOS
            cell_string = Replace(cell_string, "/", " ")
            'SPLIT CELL STRING IN INDIVIDUAL WORDS SEPARATED BY SPACES
            words = Split(cell_string, " ")
            'INITIALISE CHARACTER COUNT
            i_chars = 1

            'CHECK CUSTOM PHRASES, IF CUSTOM PHRASE IS IN THE CELL STRING THEN FIND EXACT POSITION & LENGTH AND APPLY HIGHLIGHT
            For Each term In ThisWorkbook.Worksheets(1).Range("custom_spell_range")
                If Not term = "" Then
                    start = 1
                    Do
                        pos = InStr(start, cell_string, term, vbBinaryCompare)
                        If pos > 0 Then
                            start = pos + 1
                            tempS = Mid(cell_string, pos, Len(term))
                            If LCase(tempS) = LCase(term) Then
                                cell.Characters(start:=pos, Length:=Len(term)).Font.Color = highlight_colour
                            End If
                        End If
                    Loop While pos > 0
                End If
            Next term

            'INITIALISE TOTAL NUMBER OF WORDS IN A CELL
            word_total = UBound(words)

            'LOOP THROUGH ALL WORDS CHECKING FOR REPEATING WORDS, AND MISSPELLED WORDS
            For i_words = 0 To word_total
                'CHECK IF LAST WORD
                If i_words = word_total Then
                    current_word = words(i_words)
                    current_len = Len(current_word)
                    GoTo skip_check    'AVOIDS REFERENCE GREATER THAN ARRAY SIZE, & SKIPS FURTHER CHECKING FOR NEXT WORD
                Else
                    current_word = words(i_words)
                    next_word = words(i_words + 1)
                    current_len = Len(current_word)
                    next_len = Len(next_word)
                    combined_len = current_len + next_len
                End If
                'CHECKS FOR FOLLOWING WORD BEING IDENTICAL (USING LOWER CASE TO CAPTURE 'The the', ETC)
                If LCase(current_word) = LCase(next_word) Then
                    With cell.Characters(i_chars, 1 + combined_len).Font
                        .Color = highlight_colour
                    End With
                    GoTo skip_check
                End If
                'CHECK FOR NUMBER OF BEGINNING CHARACTERS BEING NON-ALPHANUMERIC IN CURRENT WORD SO THESE CHARACTERS CAN BE IGNORED IN THE HIGHLIGHTING
                J = 1
                Do While J < current_len
                    If IsCharAlphaNumeric(current_word, J, True) = True Then
                        GoTo skip_end_check1
                    Else
                        J = J + 1
                    End If
                Loop

skip_end_check1:
                'CHECK FOR NUMBER OF BEGINNING CHARACTERS BEING NON-ALPHANUMERIC IN CURRENT WORD SO THESE CHARACTERS CAN BE IGNORED IN THE HIGHLIGHTING
                I = 1
                Do While I < next_len
                    If IsCharAlphaNumeric(next_word, I, False) = True Then
                        GoTo skip_end_check2
                    Else
                        I = I + 1
                    End If
                Loop

skip_end_check2:
                'CHECK IF CURRENT WORD AND NEXT WORD ARE THE SAME
                'USED TO PICKUP 'the the.' OR '.the the', ETC
                If LCase(Right(current_word, current_len - J + 1)) = LCase(left(next_word, next_len - I + 1)) Then
                    With cell.Characters(i_chars + J - 1, 1 + 2 * (current_len - J + 1)).Font
                        .Color = highlight_colour
                    End With
                    GoTo skip_check
                End If

skip_check:
                'CHECK IF WORD HAS BEEN PREVIOUSLY CHECKED AND WAS CORRECTLY SPELLED, THEN SKIP RECHECKING SAME WORD AGAIN
                pos = InStr(1, correctly_spelled_words, " " & current_word & " ", vbBinaryCompare)
                If pos > 0 Then
                    'DO NOTHING, THIS WORD HAS BEEN CHECKED PREVIOUSLY
                Else
                    If Not Application.CheckSpelling(Word:=current_word) Then
                        'IF SPELL CHECK IS FALSE THEN WORD IS MISSPELLED
                        'CHECK FOR NUMBER OF END CHARACTERS BEING NON-ALPHANUMERIC IN CURRENT WORD SO THESE CHARACTERS CAN BE IGNORED IN THE HIGHLIGHTING, NOTE BEGINNING POSITION ('J') WAS CHECKED ABOVE AND IS RE-USED
                        I = 1
                        Do While I < current_len
                            If IsCharAlphaNumeric(current_word, I, False) = True Then
                                GoTo skip_end_check3
                            Else
                                I = I + 1
                            End If
                        Loop

skip_end_check3:
                        With cell.Characters(i_chars + J - 1, current_len - J - I + 2).Font
                            .Color = highlight_colour
                        End With
                    Else
                        'STORE WORD IN CACHE IF IT'S SPELLED CORRECTLY, SO IT IS SKIPPED IF IT TURNS UP AGAIN
                        correctly_spelled_words = correctly_spelled_words & current_word & " "
                    End If

                End If
                'SET CHARACTER POSITION TO THE START OF THE NEXT WORD
                i_chars = i_chars + 1 + current_len
            Next i_words
        End If
    Next cell

handleCancel:

    If Err = 18 Then
        MsgBox "Spellcheck Cancelled"
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub ClearStatusBar_Spelling()
'CLEAR MESSAGE FROM STATUS BAR
Application.StatusBar = False
End Sub


Private Function IsCharAlphaNumeric(strValue As String, Position As Double, Left_to_Right As Boolean) As Boolean
    Dim intPos As Integer    'TEMPORARY POSITION WITHIN THE STRING FROM LEFT OR RIGHT END, COUNT STARTS AT 1 THROUGH TO THE LENGTH OF THE STRING

'DETERMINE IF WORKING LEFT TO RIGHT OR RIGHT TO LEFT WITHIN THE STRING
    If Left_to_Right = True Then
        intPos = Position
    Else
        intPos = Len(strValue) + 1 - Position
    End If

    'ISOLATE CHARACTER TO BE CHECKED AT GIVEN POSITION WITHIN THE STRING
    Select Case Asc(Mid(strValue, intPos, 1))
        Case 65 To 90, 97 To 122, 48 To 57
            '65 to 90 = CAPITAL LETTERS, 97 to 122 = LOWERCASE LETTERS, 48 to 57 = NUMBERS, REFER https://www.techonthenet.com/ascii/chart.php
            IsCharAlphaNumeric = True
        Case Else
            IsCharAlphaNumeric = False
    End Select

End Function
