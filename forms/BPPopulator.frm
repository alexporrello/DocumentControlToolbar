VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BPPopulator 
   Caption         =   "Boilerplate Populator"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   OleObjectBlob   =   "BPPopulator.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "BPPopulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AuthorityField_Change()

End Sub

Private Sub ClientName_Click()

End Sub

Private Sub ClientNameField_Change()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub ProjectNameField_Change()

End Sub

Private Sub Run_Click()
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    Call FormatTableCaptions
    Call FormatAllTables
    
    Call findReplace("(ClientName)", BPPopulator.ClientNameField.text, BPPopulator.ClientNameField = vbNullString)
    Call findReplace("(ContractName)", BPPopulator.ContractNameField.text, BPPopulator.ContractNameField = vbNullString)
    Call findReplace("(ProjectName)", BPPopulator.ProjectNameField.text, BPPopulator.ProjectNameField = vbNullString)
    Call findReplace("(RoadName)", BPPopulator.RoadNameField.text, BPPopulator.RoadNameField = vbNullString)
    Call findReplace("(Authority)", BPPopulator.AuthorityField.text, BPPopulator.AuthorityField = vbNullString)

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
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    selection.find.Execute replace:=wdReplaceAll
End Sub
Private Sub UserForm_Click()

End Sub
