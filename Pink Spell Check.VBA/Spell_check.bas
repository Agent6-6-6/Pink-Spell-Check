Attribute VB_Name = "Spell_check"
Option Explicit

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

    'On Error GoTo handleCancel

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

'    ActiveSheet.Range("Top_Titles").Font.Color = RGB(255, 255, 255)        ' for peer review log as forcing white text to black on black background
'    ActiveSheet.Range("General_specification_comments").Font.Color = RGB(255, 255, 255)
'    ActiveSheet.Range("Design_comments").Font.Color = RGB(255, 255, 255)
'    ActiveSheet.Range("Drawing_comments").Font.Color = RGB(255, 255, 255)
'    ActiveSheet.Range("Modelling_comments").Font.Color = RGB(255, 255, 255)
'_____________________________________________________________________________________
'OLDER CODE OR UNTESTED CODE BELOW THIS POINT

''''https://techniclee.wordpress.com/2010/07/21/isletter-function-for-vba/
'''
'''Function IsLetter(strValue As String) As Boolean
'''    Dim intPos As Integer
'''    For intPos = 1 To Len(strValue)
'''        Select Case Asc(Mid(strValue, intPos, 1))
'''            Case 65 To 90, 97 To 122
'''                IsLetter = True
'''            Case Else
'''                IsLetter = False
'''                Exit For
'''        End Select
'''    Next
'''End Function



'https://stackoverflow.com/questions/41671949/finding-punctuation-within-vba-string-from-the-right-side
Public Sub RunMe()
    Const punc As String = "!""*()-[]{};':@~,./<>?"

Debug.Print InStrRevAny("TE.ST))", punc)
End Sub
'Private Function InStrRevAny(refText As String, chars As String) As Long
'    Dim i As Long, j As Long
'
'    For i = Len(refText) To 1 Step -1
'        For j = 1 To Len(chars)
'            'Debug.Print (Mid(refText, i, 1))
'            If Mid(refText, i, 1) = Mid(chars, j, 1) Then
'                InStrRevAny = i
'                Exit Function
'            End If
'        Next
'    Next
'End Function
Private Function Drop_right_punctuation_end_position(refText As String, chars As String) As Long
    Dim I As Long, J As Long

    For I = 1 To Len(refText) Step 1
        For J = 1 To Len(chars)
            'Debug.Print Len(chars)
            If Mid(refText, I, 1) = Mid(chars, J, 1) Then
                Drop_right_punctuation_end_position = I - 1
                Exit Function
            End If
        Next
    Next
End Function



Private Function InStrRevAny(refText As String, chars As String) As Long
    Dim I As Long, J As Long

    For I = 1 To Len(refText) Step 1
        For J = 1 To Len(chars)
            'Debug.Print (Mid(refText, i, 1))
            If Mid(refText, I, 1) = Mid(chars, J, 1) Then
                InStrRevAny = I
                Exit Function
            End If
        Next
    Next
End Function

Sub HighlightMisspelledCells()
    Dim cl As Range
    For Each cl In ActiveSheet.UsedRange
        If Not Application.CheckSpelling(Word:=cl.Value) Then    '
            cl.Interior.Color = vbRed
        End If
    Next cl

    'only seems to work on cells with low word count
End Sub


Sub HighlightMisspelledCells2()
    Dim cl As Range
    For Each cl In ActiveSheet.UsedRange
        If Not Application.CheckSpelling(Word:=cl.Text) Then    '
            cl.Interior.Color = vbRed
        End If
    Next cl

    'only seems to work on cells with low word count
End Sub


Sub HighlightMisspelledWords()

' Purpose: SpellChecks the entire sheet (or some other specified range) Cell-by-Cell and Word-by-Word,
' highlighting in a color those Words and those Cells with misspelled Words.
' This can run S-L-O-W since it is calling the SpellChecker for each individual Word.

' Optionally can set the Column number that you want the text MISSPELLED inserted into
' so that you can later sort the sheet on that Column to consolidate the problem Rows.
' George Mason


' You can specify a Range by indicating the upper left and the lower right Cell.
' The default Range uses the entire used area of the Sheet.
    Dim oRange As Excel.Range
    'Set oRange = Range("D1:D500")
    Set oRange = ActiveSheet.UsedRange

    Application.ScreenUpdating = True

    ' You can pick which Dictionary Language to spell check against.
    ' There are many, but the following are the most likely.
    Application.SpellingOptions.DictLang = msoLanguageIDEnglishUK    ' value 2057

    ' Other possible spell check options.
    Application.SpellingOptions.SuggestMainOnly = True
    ' The following assumes you want the SpellChecker to ignore Uppercase things, like ACRONYMS.
    Application.SpellingOptions.IgnoreCaps = True
    Application.SpellingOptions.IgnoreMixedDigits = True
    Application.SpellingOptions.IgnoreFileNames = True

    ' For Cell Highlighting, there are 8 named colors you may choose from:
    ' vbBlack, vbWhite, vbRed, vbGreen, vbBlue, vbYellow, vbMagenta, vbCyan.
    Dim lCellHighlightColor As Long
    lCellHighlightColor = vbYellow

    ' For Word Highlighting, there are 8 named colors you may choose from:
    ' vbBlack, vbWhite, vbRed, vbGreen, vbBlue, vbYellow, vbMagenta, vbCyan.
    Dim lWordHighlightColor As Long
    lWordHighlightColor = vbRed

    ' You should set these next 3 items
    ' if you want to have one Column used to mark ANY Cell misspellings for the entire Row.
    Dim bColumnMarker As Boolean
    'bColumnMarker = False
    bColumnMarker = True

    ' Column A = 1, Column B = 2, etc.
    Dim iColumnToMark As Integer
    iColumnToMark = 7

    Dim sMarkerText As String
    sMarkerText = "MISSPELLED"

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' The values for the items above should be modified by the user, as necessary. '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo 0

    Dim oCell As Object
    Dim iLastRowProcessed As Integer
    iLastRowProcessed = 0

    For Each oCell In oRange

        If ((bColumnMarker = True) And _
            (iLastRowProcessed <> oCell.Row)) Then
            ' When beginning to process each new Row, clear out any previous MISSPELLED marker.
            iLastRowProcessed = oCell.Row
            Cells(oCell.Row, iColumnToMark) = ""
        End If
        Rows(oCell.Row).Select

        ' Boolean to track for ANY misspelling in the Cell.
        Dim bResultCell As Boolean
        bResultCell = True

        ' First spell check the entire cell (if less than 256 chars).
        ' This can catch some grammatical errors even if no spelling errors.
        If (Len(oCell.Text) < 256) Then
            bResultCell = Application.CheckSpelling(oCell.Text)
        End If

        Dim iTrackCharPos As Integer
        iTrackCharPos = 1

        ' Split the Text in the Cell into an array of words, using a Space as the delimiter.
        Dim vWords As Variant
        vWords = Split(oCell.Text, Chr(32), -1, vbBinaryCompare)
        Dim I As Integer

        ' Check the spelling of each word in the Cell.
        For I = LBound(vWords) To UBound(vWords)

            Dim iWordLen As Integer
            iWordLen = Len(vWords(I))

            Dim bResultWord As Boolean
            ' Note that a Word longer than 255 characters will generate Error 13.
            ' Any character string without any embedded space is considered a Word.
            bResultWord = Application.CheckSpelling(Word:=vWords(I))

            If (bResultWord = False) Then
                ' Thinks it is misspelled.
                ' Check for trailing punctuation and plural words like ACTION-EVENTs.
                ' The following is crude and should be made more robust when there is time.
                If (iWordLen > 1) Then
                    Dim iWL As Integer
                    For iWL = iWordLen To 1 Step -1
                        If (Not (Mid(vWords(I), iWL, 1) Like "[0-9A-Za-z]")) Then
                            vWords(I) = left(vWords(I), (iWL - 1))
                        Else
                            Exit For
                        End If
                    Next iWL
                    If (Mid(vWords(I), iWL, 1) = "s") Then
                        ' Last letter is lowercase "s".
                        vWords(I) = left(vWords(I), (iWL - 1))
                    End If
                    ' Retest.
                    bResultWord = Application.CheckSpelling(Word:=vWords(I))
                End If
            End If

            If (bResultWord = True) Then
                ' If this is an Uppercased and Hyphenated word, we should split and lowercase then check each portion.
                If ((Len(vWords(I)) > 0) And (vWords(I) = UCase(vWords(I)))) Then
                    ' Word is all Uppercase, check for hyphenation.
                    Dim iHyphenPos As Integer
                    iHyphenPos = InStr(1, vWords(I), "-")
                    If (iHyphenPos > 0) Then
                        ' Word is also hyphenated, split and lowercase then check each portion.
                        Dim vHyphenates As Variant
                        vHyphenates = Split(LCase(vWords(I)), "-", -1, vbBinaryCompare)
                        Dim iH As Integer
                        ' Check the spelling of each newly lowercased portion of the word.
                        For iH = LBound(vHyphenates) To UBound(vHyphenates)
                            bResultWord = Application.CheckSpelling(Word:=vHyphenates(iH))
                            If (bResultWord = False) Then
                                ' As soon as any portion is deemed misspelled, then done.
                                Exit For
                            End If
                        Next iH
                    End If
                End If
            End If

            If (bResultWord = False) Then
                bResultCell = False
                ' Highlight just this misspelled word in the Cell.
                oCell.Characters(iTrackCharPos, iWordLen).Font.Bold = True
                oCell.Characters(iTrackCharPos, iWordLen).Font.Color = lWordHighlightColor
            Else
                ' Clear any previous Highlight on just this word.
                oCell.Characters(iTrackCharPos, iWordLen).Font.Bold = False
                oCell.Characters(iTrackCharPos, iWordLen).Font.Color = vbBlack
            End If

            iTrackCharPos = iTrackCharPos + iWordLen + 1

        Next I

        If (bResultCell = True) Then
            ' The text contents of this Cell are NOT misspelled.
            ' Remove any previous highlighting by setting the Fill Color to the "No Fill" value.
            oCell.Interior.ColorIndex = xlColorIndexNone
        Else
            ' At least some of the text contents of this Cell are misspelled, so highlight the Cell.
            oCell.Interior.Color = lCellHighlightColor
            ' Mark the Row, if requested.
            If (bColumnMarker = True) Then
                Cells(oCell.Row, iColumnToMark) = sMarkerText
            End If
        End If

    Next oCell

End Sub
