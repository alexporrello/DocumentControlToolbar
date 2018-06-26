Attribute VB_Name = "Acronyms"
Sub acronymBlackMagic()
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False
    
    For Each item In getAcronymsNotInTable()
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
    
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
End Sub

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
            
            highlightIfUnused (ChkTxt)
            
            If K = 1 Then recordedAcronyms.Add (ChkTxt)
        Next K
    Next J
    
    Set MarkUnusedAcronyms = recordedAcronyms
End Function

Private Function highlightIfUnused(toFind As String)

Dim iCount As Integer
Dim strSearch As String

iCount = 0

With ActiveDocument.Content.find
    .text = toFind
    .Format = False
    .Wrap = wdFindStop
    .MatchCase = True
    Do While .Execute
        iCount = iCount + 1
    Loop
End With

If iCount = 1 Then
    selection.Range.HighlightColorIndex = wdRed
End If

End Function

Private Function GetAllAcronymsInDocument() As collection
    Dim wd As Range
    Dim coll As New collection
    
    Dim excelApp As Object
    Set excelApp = CreateObject("Excel.Application")
    
    For Each wd In ActiveDocument.Words
    
        Dim thisString As String
        thisString = Trim(wd.text)

        If Len(thisString) > 1 And Len(thisString) < 7 Then
            If wd.text = UCase(thisString) Then
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

Private Function IsAlpha(strValue As String, excelApp As Object) As Boolean
    IsAlpha = strValue Like excelApp.WorksheetFunction.Rept("[a-zA-Z]", Len(strValue))
End Function

Private Function inCollection(thisCollection As collection, item As String) As Boolean
    Dim toReturn As Boolean
    toReturn = False
    
    For Each acronym In thisCollection
        If (StrComp(acronym, item, vbTextCompare) = 0) Then
            toReturn = True
            Exit For
        End If
    Next acronym
    
    inCollection = toReturn
    
End Function

Private Function inDudList(item As String) As Boolean
    Dim duds As New collection
        
    duds.Add ("PDF")
    duds.Add ("XX")
    duds.Add ("MM")
    duds.Add ("YY")
    duds.Add ("DD")
    duds.Add ("HH")
    duds.Add ("MM")
    duds.Add ("SS")
    duds.Add ("TBD")
    duds.Add ("JIRA")
    duds.Add ("KAP")
    duds.Add ("CDRL")
    duds.Add ("KTCPRO")
    duds.Add ("SDG")
    duds.Add ("SR")
    duds.Add ("IMG")
    duds.Add ("GB")
    duds.Add ("MB")
    duds.Add ("PNL")
    duds.Add ("PDU")
    duds.Add ("RAM")
    duds.Add ("RESETLOGS")
    duds.Add ("WT")
    
    isDud = inCollection(duds, item)
End Function
Private Function inDictionary(toCheck As String) As Boolean
    inDictionary = Application.checkSpelling(LCase(toCheck))
End Function

Sub TimeApp()
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    'Remember time when macro starts
    StartTime = Timer
    
    Call acronymBlackMagic
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Sub
