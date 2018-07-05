Attribute VB_Name = "Acronyms"
Sub RunAcronymTableMacro()
    
    System.Cursor = wdCursorWait    'Set the cursor to spinning and turn of screen updating while this acronym runs
    Application.ScreenUpdating = False
    
    'Look through all of the tables to find the Acronym or Abbreviations table
    For i = 1 To ActiveDocument.Tables.count
        ActiveDocument.Tables(i).Cell(1, 1).Range.Select
        
        ChkTxt = selection.text
        ChkTxt = Left(ChkTxt, Len(ChkTxt) - 2) 'Remove end of cell markers
                
        If ChkTxt = "Abbreviation" Or ChkTxt = "Abbreviations" Or ChkTxt = "Acronym" Or ChkTxt = "Acronyms" Then ' Check the top left cell of the table to see its name
            Call removeOAndM
            Call acronymBlackMagic  'If it is an acronyms/abbreviations table, start working
            Exit For
        End If
    Next i
    
    Application.ScreenUpdating = True   'Set the cursor back to normal and update the screen
    System.Cursor = wdCursorNormal
End Sub

Private Function acronymBlackMagic()
Attribute acronymBlackMagic.VB_ProcData.VB_Invoke_Func = "Normal.Acronyms.acronymBlackMagic"
    For Each item In getAcronymsNotInTable()    'Gather all of the acronyms that are not in the acronyms table
        Dim acronym As String
        acronym = item
        
        If Not inDudList(acronym) And Not inDictionary(acronym) Then
            With selection
                .Tables(1).Rows.Add
                .Tables(1).Cell(selection.Tables(1).Rows.count, 1).Select
                .InsertAfter (item)
                .Range.HighlightColorIndex = wdYellow
            End With
        End If
    Next item
    
    selection.Tables(1).SortAscending
End Function

Private Function getAcronymsNotInTable() As collection

    Dim acronymsInTable As New collection
    Set acronymsInTable = MarkUnusedAcronyms

    Dim acronymsInDocument As New collection
    Set acronymsInDocument = GetAllAcronymsInDocument
    
    Dim acronymsNotInTable As New collection
        
    For Each item In acronymsInDocument
        Dim acronym As String
        acronym = item

        If Not inCollection(acronymsInTable, acronym) Then
            acronymsNotInTable.Add (acronym)
        End If
    Next item
    
    Set getAcronymsNotInTable = acronymsNotInTable
End Function

Private Function MarkUnusedAcronyms() As collection
    Dim recordedAcronyms As New collection
    Dim NumRows As Integer
    Dim NumCols As Integer
    Dim J As Integer
    Dim K As Integer
    Dim ChkTxt As String

    If Not selection.Information(wdWithInTable) Then
        Exit Function
    End If

    NumRows = selection.Tables(1).Rows.count
    NumCols = selection.Tables(1).Columns.count

    For J = 2 To NumRows
        For K = 1 To NumCols
            selection.Tables(1).Cell(J, K).Range.Select
            
            ChkTxt = selection.text
            ChkTxt = Left(ChkTxt, Len(ChkTxt) - 2) 'Remove end of cell markers
            
            Dim checkCase As Boolean
            checkCase = True
            
            If K = 2 Then checkCase = False
            
            Call highlightIfUnused(ChkTxt, checkCase)
            
            If K = 1 Then recordedAcronyms.Add (ChkTxt)
        Next K
    Next J
    
    Set MarkUnusedAcronyms = recordedAcronyms
End Function

Private Function highlightIfUnused(toFind As String, matchCase As Boolean)

Dim iCount As Integer
Dim strSearch As String

iCount = 0

With ActiveDocument.Content.find
    .text = toFind
    .Format = False
    .Wrap = wdFindStop
    .matchCase = matchCase
    Do While .Execute
        iCount = iCount + 1
    Loop
End With

If iCount = 1 Then
    selection.Range.HighlightColorIndex = wdRed
End If

End Function

Public Function GetAllAcronymsInDocument() As collection
    Dim wd As Range
    Dim coll As New collection
    
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")
    
    For Each wd In ActiveDocument.Words
        Dim thisString As String
        thisString = Trim(wd.text)

        If Len(thisString) > 1 And Len(thisString) < 7 Then
            If isKnownAcronym(thisString) Then
                If Not inCollection(coll, thisString) Then
                    coll.Add (thisString)
                End If
            ElseIf wd.text = UCase(thisString) Then
                If IsAlpha(thisString, excelApp) Then
                    If Not wd.Font.Name = "Courier New" Then
                        If Not inCollection(coll, thisString) Then
                            coll.Add (thisString)
                        End If
                    End If
                End If
            End If
        End If
    Next wd
    
    excelApp.Quit
    
    Set GetAllAcronymsInDocument = coll
End Function

Public Function isKnownAcronym(word As String) As Boolean
    Dim firstLetter As Integer
    firstLetter = Asc(LCase(Left(word, 1)))

    If Len(word) > 0 And firstLetter > 96 And firstLetter < 123 And UCase(word) = word Then
        isKnownAcronym = Contains(readInTextFile(Left(word, 1) + ".csv"), UCase(word))
    End If
End Function

Private Function inDudList(item As String) As Boolean
    
    Dim duds As String
    duds = GetFromWebpage("https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/acronym-duds.txt")
    
    Dim dudList As New collection
    
    For Each dud In Split(duds, vbLf)
        dudList.Add (dud)
    Next dud
    
    inDudList = inCollection(dudList, item)
End Function

Private Function removeOAndM()
    selection.find.ClearFormatting
    selection.find.Replacement.ClearFormatting
    With selection.find
        .text = "O&M"
        .Replacement.text = "IOM"
        .Forward = True
        .Wrap = wdFindContinue
    End With
    selection.find.Execute replace:=wdReplaceAll
End Function


