VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocControlTools 
   Caption         =   "Document Control Tools"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4065
   OleObjectBlob   =   "DocControlTools.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DocControlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
Private Sub acronymTableUpdaterButton_Click()
    Call RunAcronymTableMacro
End Sub

Private Sub acronymTableUpdaterButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    acronymTableUpdaterButton.BackColor = &H8000000A
End Sub

Private Sub acronymTableUpdaterButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    acronymTableUpdaterButton.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub applyBodyStyleButton_Click()
    Call ApplyBodyText
End Sub

Private Sub applyBodyStyleButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    applyBodyStyleButton.BackColor = &H8000000A
End Sub

Private Sub applyBodyStyleButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    applyBodyStyleButton.BackColor = &H8000000F
End Sub


'==========================================================
Private Sub formatTableButton_Click()
    Call FormatTable
End Sub

Private Sub formatTableButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    formatTableButton.BackColor = &H8000000A
End Sub

Private Sub formatTableButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    formatTableButton.BackColor = &H8000000F
End Sub


'==========================================================
Private Sub keepWithNextButton_Click()
    Call KeepWithNext
End Sub

Private Sub keepWithNextButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    keepWithNextButton.BackColor = &H8000000A
End Sub

Private Sub keepWithNextButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    keepWithNextButton.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub docPropertiesUpdaterButton_Click()
    DocPropertiesUpdate.Show
End Sub

Private Sub docPropertiesUpdaterButton_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    docPropertiesUpdaterButton.BackColor = &H8000000A
End Sub

Private Sub docPropertiesUpdaterButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    docPropertiesUpdaterButton.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub h1_Click()
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("Heading 1,2016_Überschrift 1,Headline 1")
End Sub

Private Sub h1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h1.BackColor = &H8000000A
End Sub

Private Sub h1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h1.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub h2_Click()
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("Heading 2,2016_Überschrift 2,Headline 2")
End Sub

Private Sub h2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h2.BackColor = &H8000000A
End Sub

Private Sub h2_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h2.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub h3_Click()
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("Heading 3,2016_Überschrift 3,Headline 3")
End Sub

Private Sub h3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h3.BackColor = &H8000000A
End Sub

Private Sub h3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h3.BackColor = &H8000000F
End Sub

'==========================================================
Private Sub h4_Click()
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("Heading 4,2016_Überschrift 4,Headline 4")
End Sub

Private Sub h4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h4.BackColor = &H8000000A
End Sub

Private Sub h4_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h4.BackColor = &H8000000F
End Sub
'==========================================================
Private Sub figureCrossRef_Click()
    Call InsertFigureReference
End Sub

Private Sub figureCrossRef_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    figureCrossRef.BackColor = &H8000000A
End Sub

Private Sub figureCrossRef_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    figureCrossRef.BackColor = &H8000000F
End Sub

'========================================================== TODO
Private Sub tableCrossRef_Click()
    Call InsertTableReference
End Sub

Private Sub tableCrossRef_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tableCrossRef.BackColor = &H8000000A
End Sub

Private Sub tableCrossRef_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tableCrossRef.BackColor = &H8000000F
End Sub
'==========================================================
Private Sub h5_Click()
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("Heading 5,2016_Überschrift 5,Headline 5")
End Sub

Private Sub h5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h5.BackColor = &H8000000A
End Sub

Private Sub h5_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    h5.BackColor = &H8000000F
End Sub

Private Sub Run_Click()
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Call FormatTableCaptions
    Call FormatAllTables
    
    Call findReplace("(ClientName)", DocControlTools.ClientNameField.text, DocControlTools.ClientNameField = vbNullString)
    Call findReplace("(ContractName)", DocControlTools.ContractNameField.text, DocControlTools.ContractNameField = vbNullString)
    Call findReplace("(ProjectName)", DocControlTools.ProjectNameField.text, DocControlTools.ProjectNameField = vbNullString)
    Call findReplace("(RoadName)", DocControlTools.RoadNameField.text, DocControlTools.RoadNameField = vbNullString)
    Call findReplace("(Authority)", DocControlTools.AuthorityField.text, DocControlTools.AuthorityField = vbNullString)

    Unload Me
    
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
End Sub

Sub findReplace(toReplace As String, replaceWith As String, isEmpty As Boolean)
    If Not isEmpty Then
        Call Find_and_replace(toReplace, replaceWith)
    End If
End Sub

Private Sub Find_and_replace(find As String, replace As String)
    selection.find.ClearFormatting
    selection.find.Replacement.ClearFormatting
    With selection.find.Replacement.Font
        .Bold = False
        .Italic = False
    End With
    With selection.find
        .text = find
        .Replacement.text = replace
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .matchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    selection.find.Execute replace:=wdReplaceAll
End Sub

Private Sub UserForm_Click()

End Sub
