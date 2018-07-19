Attribute VB_Name = "NewMacros"
Public Sub setAllKeyboardShortcuts()
    With Application
        .CustomizationContext = NormalTemplate
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyD), KeyCategory:=wdKeyCategoryCommand, Command:="OpenDocumentControlToolsDialog"
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyA), KeyCategory:=wdKeyCategoryCommand, Command:="AcceptThisChange"
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyB), KeyCategory:=wdKeyCategoryCommand, Command:="ApplyBodyText"
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC), KeyCategory:=wdKeyCategoryCommand, Command:="InsertBlankComment"
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyI), KeyCategory:=wdKeyCategoryCommand, Command:="increaseHeading"
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyT), KeyCategory:=wdKeyCategoryCommand, Command:="FormatTable"
    End With
End Sub
Public Sub KeepWithNext()
    With selection.ParagraphFormat
        .KeepWithNext = True
    End With
End Sub

Public Sub FormatTable()
If selection.Information(wdWithInTable) Then
    Application.ScreenUpdating = False
    
    ' Apply the MasterTable style.
    selection.Tables(1).Style = ActiveDocument.Styles("MasterTable")
    selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    ' Sometimes applying the MasterTable style isn't enough.
    ' So, we're going to manually apply the body style to the whole table
    selection.Tables(1).Select
    selection.Style = ActiveDocument.Styles("2016_Table | 9pt")
    
    ' Then, we'll select the first row and apply the TableHeader style
    selection.Tables(1).Rows(1).Select
    selection.Rows.HeadingFormat = wdToggle
    selection.Style = ActiveDocument.Styles("2016_TableHeader | 10pt bold")
    
    Application.ScreenUpdating = True
End If
End Sub

Private Sub formatBulletedList()
Dim oPara As word.Paragraph

For Each oPara In ActiveDocument.Paragraphs
    If oPara.Range.ListFormat.ListType = WdListType.wdListBullet Then
        Debug.Print oPara.LeftIndent
        If oPara.LeftIndent > 10 And oPara.LeftIndent < 25 Then
            oPara.Style = ActiveDocument.Styles("Body Text enumeration | yellow arrow")
        ElseIf oPara.LeftIndent > 25 And oPara.LeftIndent < 75 Then
            'Opara.Style=
        End If
            'oPara.Style = ActiveDocument.Styles("Body Text enumeration | yellow arrow")
    End If
Next
End Sub

Public Sub ApplyBodyText()
Attribute ApplyBodyText.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.ApplyBodyText"
    selection.Paragraphs(1).Range.Select
    selection.Style = ActiveDocument.Styles("2016_Bodytext | 9pt")
End Sub

Private Sub GenericFindAndReplace()
    Dim toFind As New collection
    Dim toReplace As New collection
    
    toFind.Add ("in house")
    toReplace.Add ("in-house")
    
    toFind.Add ("roll out")
    toReplace.Add ("rollout")
    
    toFind.Add ("roll back")
    toReplace.Add ("rollback")
    
    toFind.Add ("shall")
    toReplace.Add ("will")
    
    toFind.Add ("toll booth")
    toReplace.Add ("tollbooth")
    
    toFind.Add ("toll both")
    toReplace.Add ("tollbooth")
    
    toFind.Add ("in depth")
    toReplace.Add ("in-depth")
    
    toFind.Add ("job site")
    toReplace.Add ("jobsite")

    toFind.Add (".  ")
    toReplace.Add (". ")
    
    Dim arraySize As Integer
    arraySize = toFind.count
    
    For i = 1 To arraySize
        Call FindAndReplaceAll(toFind(i), toReplace(i))
    Next i
    
End Sub
Private Function FindAndReplaceAll(find As String, replace As String)
Attribute FindAndReplaceAll.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.FindAndReplace"
    selection.find.ClearFormatting
    selection.find.Replacement.ClearFormatting
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
End Function
Sub page_break_before()
Attribute page_break_before.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.page_break_before"
    With selection.ParagraphFormat
        .PageBreakBefore = True
    End With
End Sub
Sub InsertBlankComment()
Attribute InsertBlankComment.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.insert_blank_comment"
    selection.Comments.Add Range:=selection.Range
End Sub
Sub size_image()
Attribute size_image.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.size_image"
    With selection.InlineShapes(1)
        .Width = 300
    End With
End Sub
Private Function increaseAllHeadingsByOne()
    System.Cursor = wdCursorWait    'Set the cursor to spinning and turn of screen updating while this acronym runs
    Application.ScreenUpdating = False
    
    Dim oPara As word.Paragraph
    
    For Each oPara In ActiveDocument.Paragraphs
        If oPara.OutlineLevel > 0 And oPara.OutlineLevel < 6 Then
            oPara.Range.Select
            increaseHeading
        End If
    Next oPara
    
    Application.ScreenUpdating = True
    System.Cursor = wdCursorNormal
End Function

Sub increaseHeading()
        If selection.Style = ActiveDocument.Styles("Heading 1,2016_Überschrift 1,Headline 1") Then
            selection.Style = ActiveDocument.Styles("Heading 2,2016_Überschrift 2,Headline 2")
        ElseIf selection.Style = ActiveDocument.Styles("Heading 2,2016_Überschrift 2,Headline 2") Then
            selection.Style = ActiveDocument.Styles("Heading 3,2016_Überschrift 3,Headline 3")
        ElseIf selection.Style = ActiveDocument.Styles("Heading 3,2016_Überschrift 3,Headline 3") Then
            selection.Style = ActiveDocument.Styles("Heading 4,2016_Überschrift 4,Headline 4")
        ElseIf selection.Style = ActiveDocument.Styles("Heading 4,2016_Überschrift 4,Headline 4") Then
            selection.Style = ActiveDocument.Styles("Heading 5,2016_Überschrift 5,Headline 5")
        ElseIf selection.Style = ActiveDocument.Styles("Heading 5,2016_Überschrift 5,Headline 5") Then
            ApplyBodyText
            selection.Font.Bold = True
        End If
End Sub

Sub AcceptThisChange()
Attribute AcceptThisChange.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.AcceptThisChange"
    selection.Range.Revisions.AcceptAll
    selection.NextRevision (True)
End Sub
