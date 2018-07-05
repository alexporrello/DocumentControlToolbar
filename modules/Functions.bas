Attribute VB_Name = "Functions"
Public Function startsWith(str As String, prefix As String) As Boolean
    startsWith = Left(str, Len(prefix)) = prefix
End Function

Public Function readInTextFile(csvFileName As String) As String

    Dim fileName As String, textData As String, fileNo As Integer

    fileName = "C:\Users\" + (Environ$("Username")) + "\AppData\Roaming\Microsoft\Templates\" + csvFileName
    fileNo = FreeFile 'Get first free file number
 
    Open fileName For Input As #fileNo
        textData = Input$(LOF(fileNo), fileNo)
    Close #fileNo

    readInTextFile = textData
End Function

Public Function GetFromWebpage(URL As String) As String
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

Public Function inDictionary(toCheck As String) As Boolean
    inDictionary = Application.checkSpelling(LCase(toCheck))
End Function

Public Function IsAlpha(strValue As String, excelApp As Object) As Boolean
    IsAlpha = strValue Like excelApp.WorksheetFunction.Rept("[a-zA-Z]", Len(strValue))
End Function

Public Function Contains(strBaseString As String, strSearchTerm As String) As Boolean
    On Error GoTo ErrorMessage
        Contains = InStr(strBaseString, strSearchTerm)
    Exit Function
ErrorMessage:     MsgBox "The database has generated an error. Please contact the database administrator, quoting the following error message: '" & Err.Description & "'", vbCritical, "Database Error"
End
End Function

Public Function inCollection(thisCollection As collection, item As String) As Boolean
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


Public Function TimeApp()
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    'Remember time when macro starts
    StartTime = Timer
    
    '========================================================

    'TODO Put code here

    '========================================================

    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    Debug.Print "This code ran successfully in " & SecondsElapsed & " seconds"
End Function
