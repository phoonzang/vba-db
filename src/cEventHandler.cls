VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public uForm As Object

Public EventChange As String
Public EventClick As String
Public EventMouseMove As String
Public EventMouseUp As String

Private mvarControl As MSForms.Control
Private WithEvents mvarTextBox As MSForms.TextBox
Attribute mvarTextBox.VB_VarHelpID = -1
Private WithEvents mvarComboBox As MSForms.ComboBox
Attribute mvarComboBox.VB_VarHelpID = -1
Private WithEvents mvarLabel As MSForms.Label
Attribute mvarLabel.VB_VarHelpID = -1
Private WithEvents mvarOptionButton As MSForms.OptionButton
Attribute mvarOptionButton.VB_VarHelpID = -1
Private WithEvents mvarCheckBox As MSForms.CheckBox
Attribute mvarCheckBox.VB_VarHelpID = -1
Private WithEvents mvarCommandButton As MSForms.CommandButton
Attribute mvarCommandButton.VB_VarHelpID = -1


Public Property Set ctrl(ByRef kontrolka As MSForms.Control)
    Set mvarControl = kontrolka
    Select Case TypeName(kontrolka)
        Case "TextBox"
            Set mvarTextBox = kontrolka
        Case "ComboBox"
            Set mvarComboBox = kontrolka
        Case "OptionButton"
            Set mvarOptionButton = kontrolka
        Case "Label"
            Set mvarLabel = kontrolka
        Case "CheckBox"
            Set mvarCheckBox = kontrolka
        Case "CommandButton"
            Set mvarCommandButton = kontrolka
    End Select
End Property

Private Sub mvarTextBox_Change()
    If EventChange <> Empty Then CallByName uForm, EventChange, VbMethod, mvarTextBox
End Sub

Private Sub mvarCheckBox_Change()
    If EventChange <> Empty Then CallByName uForm, EventChange, VbMethod, mvarCheckBox
End Sub

Private Sub mvarOptionButton_Change()
    If EventChange <> Empty Then CallByName uForm, EventChange, VbMethod, mvarOptionButton
End Sub

Private Sub mvarComboBox_Change()
    If EventChange <> Empty Then CallByName uForm, EventChange, VbMethod, mvarComboBox
End Sub




Private Sub mvarTextBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarTextBox, Array(Button, Shift, x, y)
End Sub

Private Sub mvarCheckBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarCheckBox, Array(Button, Shift, x, y)
End Sub

Private Sub mvarCommandButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarCommandButton, Array(Button, Shift, x, y)
End Sub

Private Sub mvarLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarLabel, Array(Button, Shift, x, y)
End Sub

Private Sub mvarOptionButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarOptionButton, Array(Button, Shift, x, y)
End Sub

Private Sub mvarComboBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseMove <> Empty Then CallByName uForm, EventMouseMove, VbMethod, mvarComboBox, Array(Button, Shift, x, y)
End Sub


Private Sub mvarCheckBox_Click()
    If EventClick <> Empty Then CallByName uForm, EventClick, VbMethod, mvarCheckBox
End Sub

Private Sub mvarCommandButton_Click()
    If EventClick <> Empty Then CallByName uForm, EventClick, VbMethod, mvarCommandButton
End Sub

Private Sub mvarLabel_Click()
    If EventClick <> Empty Then CallByName uForm, EventClick, VbMethod, mvarLabel
End Sub

Private Sub mvarOptionButton_Click()
    If EventClick <> Empty Then CallByName uForm, EventClick, VbMethod, mvarOptionButton
End Sub

Private Sub mvarComboBox_Click()
    If EventClick <> Empty Then CallByName uForm, EventClick, VbMethod, mvarComboBox
End Sub


Private Sub mvarTextBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarTextBox
End Sub

Private Sub mvarCommandButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarCommandButton
End Sub

Private Sub mvarLabel_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarLabel
End Sub

Private Sub mvarOptionButton_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarOptionButton
End Sub

Private Sub mvarCheckBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarCheckBox
End Sub

Private Sub mvarComboBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If EventMouseUp <> Empty Then CallByName uForm, EventMouseUp, VbMethod, mvarComboBox
End Sub
