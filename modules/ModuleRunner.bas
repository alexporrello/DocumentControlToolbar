Attribute VB_Name = "ModuleRunner"
Sub OpenDocumentControlToolsDialog()
    DocControlTools.Show
End Sub
Sub OpenDocPropertiesUpdater()
    DocPropertiesUpdate.Show
End Sub
Private Sub OpenBoilerplatePopulator()
    BPUserForm.Show
End Sub
Sub SetKeyboardShortcut()
    With Application
        .CustomizationContext = NormalTemplate
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyD), KeyCategory:=wdKeyCategoryCommand, Command:="OpenDocumentControlToolsDialog"
    End With
End Sub

