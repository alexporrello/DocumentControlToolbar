Attribute VB_Name = "Boilerplate"
Sub OpenUI()
    BPUserForm.Show
End Sub
Sub InsertTableCrossreference()
    selection.InsertCaption Label:="Table", titleAutoText:="InsertCaption2", _
        Title:="", Position:=wdCaptionPositionAbove, ExcludeLabel:=0
    selection.TypeText text:=vbTab
    selection.Style = ActiveDocument.Styles("2016_Marking")
End Sub
Sub FormatTableCaptions()
    Call InitializeCrossreferenceReplace("#TABLE#")
    
    Dim lastPos As Long
    lastPos = -1
    
    Do While selection.find.Execute = True
        Call InsertCrossreference("#TABLE#", "Table")
    Loop

End Sub
'Kicks off the search for @InsertCrossreferences
Private Sub InitializeCrossreferenceReplace(strTextToFind As String)

    selection.find.ClearFormatting
    selection.find.Replacement.ClearFormatting
    With selection.find
        .text = strTextToFind
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    selection.Style = ActiveDocument.Styles("2016_Marking")

End Sub
' Replaces a tag (such as #TABLE#) with a cross reference.
' @crType can be one of three:
'      1. "Table" for tables
'      2. "Figure" for figures
'      3. "Appendix" for appendices
Private Sub InsertCrossreference(strTextToFind As String, crType As String)
    selection.find.ClearFormatting
    selection.find.Replacement.ClearFormatting
    With selection.find
        .text = strTextToFind
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    With selection
        If .find.Forward = True Then
            .Collapse Direction:=wdCollapseStart
        Else
            .Collapse Direction:=wdCollapseEnd
        End If
            .find.Execute replace:=wdReplaceOne
        
            selection.InsertCaption Label:=crType, titleAutoText:="InsertCaption1", Title:="", Position:=wdCaptionPositionAbove, ExcludeLabel:=0
            selection.TypeText text:=vbTab
            selection.Style = ActiveDocument.Styles("2016_Marking")
    End With
End Sub
Sub FormatAllTables()
    System.Cursor = wdCursorWait
    Application.ScreenUpdating = False

    For i = 1 To ActiveDocument.Tables.count
        FormatTables (i)
    Next i
    
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
End Sub
Sub FormatTables(i As Long)
    ' Apply the MasterTable style.
    
    ActiveDocument.Tables(i).Style = ActiveDocument.Styles("MasterTable")
    ActiveDocument.Tables(i).AutoFitBehavior (wdAutoFitWindow)
    
    ' Sometimes applying the MasterTable style isn't enough.
    ' So, we're going to manually apply the body style to the whole table
    ActiveDocument.Tables(i).Select
    selection.Style = ActiveDocument.Styles("2016_Table | 9pt")
    
    ' Then, we'll select the first row and apply the TableHeader style
    ActiveDocument.Tables(i).Rows(1).Select
    selection.Rows.HeadingFormat = wdToggle
    selection.Style = ActiveDocument.Styles("2016_TableHeader | 10pt bold")
End Sub
