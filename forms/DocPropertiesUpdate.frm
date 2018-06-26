VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocPropertiesUpdate 
   Caption         =   "Document Properties Updater"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "DocPropertiesUpdate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DocPropertiesUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function setProperty(docProperty As String, value As String)
    
    Dim propertyExists As Boolean
    Dim tempObj
    On Error Resume Next
    Set tempObj = ActiveDocument.CustomDocumentProperties.item(docProperty)
        propertyExists = (Err = 0)
    On Error GoTo 0
    
    If propertyExists Then
        Call WriteCustomProp(docProperty, value)
    End If
    
    ActiveDocument.CustomDocumentProperties(docProperty).value = value
End Function
Function WriteCustomProp(sProp As String, sValue As String)
    Dim prop As DocumentProperty
    Dim bExists As Boolean
    
    bExists = False
    
    For Each prop In ActiveDocument.CustomDocumentProperties
        If LCase(prop.Name) = LCase(sProp) Then
            bExists = True
            prop.value = sValue
            Exit For
        End If
    Next
  
    If Not bExists Then
        ActiveDocument.CustomDocumentProperties.Add Name:=sProp, value:=sValue, _
        LinkToContent:=False, Type:=msoPropertyTypeString
    End If
End Function

Private Sub Client_Change()

End Sub

Private Sub ClientAcronym_Change()

End Sub

Private Sub DocAcronym_Change()

End Sub

Private Sub DocReleaseDate_Change()

End Sub

Private Sub DocStatus_Change()

End Sub

Private Sub DocTitle_Change()

End Sub

Private Sub DocVersion_Change()

End Sub

Private Sub Go_Click()
    
System.Cursor = wdCursorWait
Application.ScreenUpdating = False
    
If Not DocPropertiesUpdate.DocTitle = vbNullString Then
    Call setProperty("DocTitle", DocPropertiesUpdate.DocTitle.text)
End If

If Not DocPropertiesUpdate.DocAcronym = vbNullString Then
    Call setProperty("DocAcronym", DocPropertiesUpdate.DocAcronym.text)
End If

If Not DocPropertiesUpdate.DocReleaseDate = vbNullString Then
    Call setProperty("DocReleaseDate", DocPropertiesUpdate.DocReleaseDate.text)
End If

If Not DocPropertiesUpdate.DocVersion = vbNullString Then
    Call setProperty("DocVersion", DocPropertiesUpdate.DocVersion.text)
End If

If Not DocPropertiesUpdate.DocStatus = vbNullString Then
    Call setProperty("DocStatus", DocPropertiesUpdate.DocStatus.text)
End If

'========================================================

If Not DocPropertiesUpdate.DocAuthor = vbNullString Then
    Call setProperty("Author", DocPropertiesUpdate.DocAuthor.text)
End If

If Not DocPropertiesUpdate.PM = vbNullString Then
    Call setProperty("ProjectManager", DocPropertiesUpdate.PM.text)
End If

'========================================================

If Not DocPropertiesUpdate.RoadName = vbNullString Then
    Call setProperty("RoadName", DocPropertiesUpdate.RoadName.text)
End If

If Not DocPropertiesUpdate.SolutionType = vbNullString Then
    Call setProperty("SolutionType", DocPropertiesUpdate.SolutionType.text)
End If

If Not DocPropertiesUpdate.SolutionAcronym = vbNullString Then
    Call setProperty("SolutionAcronym", DocPropertiesUpdate.SolutionAcronym.text)
End If

If Not DocPropertiesUpdate.ClientAcronym = vbNullString Then
    Call setProperty("ClientAcronym", DocPropertiesUpdate.ClientAcronym.text)
End If

If Not DocPropertiesUpdate.Client = vbNullString Then
    Call setProperty("Client", DocPropertiesUpdate.Client.text)
End If

'========================================================

UpdateAllFields

Unload Me

Application.ScreenUpdating = False
System.Cursor = wdCursorNormal

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub SolutionType_Change()

End Sub

Private Sub UserForm_Click()

End Sub
Public Sub UpdateAllFields()
    Dim rngStory As Word.Range
    Dim lngJunk As Long
    Dim oShp As Shape
    
    lngJunk = ActiveDocument.Sections(1).Headers(1).Range.StoryType
    
    For Each rngStory In ActiveDocument.StoryRanges
        Do
        On Error Resume Next
        rngStory.Fields.Update
        Select Case rngStory.StoryType
            Case 6, 7, 8, 9, 10, 11
            If rngStory.ShapeRange.count > 0 Then
                For Each oShp In rngStory.ShapeRange
                    If oShp.TextFrame.HasText Then
                        oShp.TextFrame.TextRange.Fields.Update
                    End If
                Next
            End If
        Case Else
          'Do Nothing
        End Select
        On Error GoTo 0
        'Get next linked story (if any)
        Set rngStory = rngStory.NextStoryRange
        Loop Until rngStory Is Nothing
    Next
End Sub
