VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DocControlTools 
   Caption         =   "Document Control Tools"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4440
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
