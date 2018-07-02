Attribute VB_Name = "Acronyms"
Sub RunAcronymTableMacro()
    Call acronymBlackMagic
End Sub

Private Sub acronymBlackMagic()
Attribute acronymBlackMagic.VB_ProcData.VB_Invoke_Func = "Normal.Acronyms.acronymBlackMagic"
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
                            Debug.Print thisString
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

Private Function inDictionary(toCheck As String) As Boolean
    inDictionary = Application.checkSpelling(LCase(toCheck))
End Function

Private Function inDudList(item As String) As Boolean
    
    Dim duds As String
    duds = GetFromWebpage("https://raw.githubusercontent.com/alexporrello/TWBoilerplateMacros/master/lists/acronym-duds.txt")
    
    Dim dudList As New collection
    
    For Each dud In Split(duds, vbLf)
        dudList.Add (dud)
        Debug.Print dud
    Next dud
    
    inDudList = inCollection(dudList, item)
End Function

Private Function GetFromWebpage(URL As String) As String
On Error GoTo Err_GetFromWebpage

    Dim objWeb As Object
    Dim strXML As String

    ' Instantiate an instance of the web object
    Set objWeb = CreateObject("Microsoft.XMLHTTP")

    ' Pass the URL to the web object, and send the request
    objWeb.Open "GET", URL, False
    objWeb.send

    ' Look at the HTML string returned
    strXML = objWeb.responsetext
        
    GetFromWebpage = strXML
    
End_GetFromWebpage:
    
    ' Clean up after ourselves!
    Set objWeb = Nothing
    Exit Function

Err_GetFromWebpage:
' Just in case there's an error!
MsgBox Err.Description & " (" & Err.Number & ")"
Resume End_GetFromWebpage

End Function
