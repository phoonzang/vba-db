VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event OnAdd(Sender As Object, key As Variant, Item As Variant)
Public Event OnItemLet(Sender As Object, key As Variant, Item As Variant, Control As Object)
Public Event OnItemGet(Sender As Object, key As Variant, Item As Variant)
Public Event OnItemSet(Sender As Object, key As Variant, Item As Variant)

Public Event ControlChanged(Sender As Object, ctrl As Object, ByRef args As Variant)
Public Event ControlMouseMove(Sender As Object, ctrl As Object, ByRef args As Variant)

Public Controls As Collection
Public keyValuePairs As Collection

Private eventHandler As cEventHandler
Attribute eventHandler.VB_VarHelpID = -1
Private cControlsEnabled As Boolean
Public tag As Variant

Public Sub Class_Initialize()
    Set keyValuePairs = New Collection
End Sub

Private Sub Class_Terminate()
    UnBindControls
    Set keyValuePairs = Nothing
    Set Controls = Nothing
End Sub

Public Sub BindControls(conCol As Object, Optional updateControls As Boolean = True)
    Dim ctrl As Object
    
    Set Controls = New Collection
    For Each ctrl In conCol
        If ctrl.tag <> "" Then
            If TypeName(ctrl) = "OptionButton" Then
                If Not Exists(ctrl.GroupName) And Not Exists(ctrl.GroupName & ctrl.tag) Then GoTo CONTINUE
            Else
                If Not Exists(ctrl.tag) Then GoTo CONTINUE
            End If
            If updateControls Then
                Select Case TypeName(ctrl)
                    Case "TextBox"
                        If Not IsNull(Item(ctrl.tag)) And Exists(ctrl.tag) Then
                            ctrl.Text = Item(ctrl.tag)
                        Else
                            ctrl.Text = ""
                        End If
                    Case "ComboBox"
                        If Not IsNull(Item(ctrl.tag)) And Exists(ctrl.tag) Then
                            ctrl.value = Item(ctrl.tag)
                        Else
                            ctrl.value = 0
                        End If
                    Case "Label"
                        If Not IsNull(Item(ctrl.tag)) And Exists(ctrl.tag) Then ctrl.Caption = Item(ctrl.tag)
                    Case "OptionButton"
                        If Not IsNull(Item(ctrl.GroupName)) And Item(ctrl.GroupName) = ctrl.tag Then
                            ctrl.value = True
                        End If
                    Case "Label"
                        If Not IsNull(Item(ctrl.tag)) And Exists(ctrl.tag) Then
                            ctrl.Caption = Item(ctrl.tag)
                        Else
                            ctrl.Caption = ""
                        End If
                End Select
            End If
            Set eventHandler = New cEventHandler
            Set eventHandler.ctrl = ctrl
            Set eventHandler.uForm = Me
            eventHandler.EventChange = "ControlChange"
            eventHandler.EventMouseMove = "ControlMouseMove"
            If TypeName(ctrl) = "OptionButton" Then
                Controls.Add eventHandler, ctrl.GroupName & ctrl.tag
            Else
                Controls.Add eventHandler, ctrl.tag
            End If
        End If
        
        Set eventHandler = Nothing
CONTINUE:
    Next
End Sub

Public Property Let ControlsEnabled(enabled As Boolean)
    cControlsEnabled = enabled
    Dim ctrl As Object
    If Controls Is Nothing Then Exit Property
    For Each ctrl In Controls
        With ctrl
            Select Case TypeName(.ctrl)
                Case "TextBox", "ComboBox"
                    .ctrl.Locked = Not enabled
                    If Not enabled Then
                        .ctrl.BackColor = &H8000000F
                    Else
                        .ctrl.BackColor = &H80000005
                    End If
                Case "OptionButton"
                    .ctrl.enabled = Not enabled
                Case "Label"
                    If Not enabled Then
                        .ctrl.BackColor = &H8000000F
                    Else
                        .ctrl.BackColor = &H80000005
                    End If
            End Select
        End With
    Next
End Property

Public Property Get ControlsEnabled() As Boolean
    ControlsEnabled = cControlsEnabled
End Property

Public Sub ControlMouseMove(ByRef ctrl As Object, Optional ByRef args As Variant)
    If TypeName(ctrl) = "OptionButton" Then
        If Exists(ctrl.GroupName) Or Exists(ctrl.GroupName & ctrl.tag) Then RaiseEvent ControlMouseMove(Me, ctrl, args)
    Else
        If Exists(ctrl.tag) Then RaiseEvent ControlMouseMove(Me, ctrl, args)
    End If
End Sub

Public Sub ControlChange(ByRef ctrl As Object, Optional ByRef args As Variant)
    If Controls Is Nothing Then Exit Sub
    If TypeName(ctrl) = "OptionButton" Then
        If Exists(ctrl.GroupName) Or Exists(ctrl.GroupName & ctrl.tag) Then RaiseEvent ControlChanged(Me, ctrl, args)
    Else
        If Exists(ctrl.tag) Then RaiseEvent ControlChanged(Me, ctrl, args)
    End If
End Sub

Public Sub UnBindControls()
    Dim i As Long
    If Controls Is Nothing Then Exit Sub
    For i = Controls.Count To 1 Step -1
        Controls(i).Clear
        Controls.Remove i
    Next
    Set Controls = Nothing
End Sub

Public Property Get CompareMode() As VbCompareMethod
    CompareMode = vbTextCompare
End Property

Public Property Let Item(key As Variant, Item As Variant)
    Let keyValuePairs.Item(key).value = Item
    If Controls Is Nothing Then
        RaiseEvent OnItemLet(Me, key, Item, Nothing)
    Else
        On Error GoTo ERR
        If inCollection(Controls, key) Then RaiseEvent OnItemLet(Me, key, Item, Controls(key).ctrl)
        Exit Property
ERR:
        If inCollection(Controls, key & Item) Then RaiseEvent OnItemLet(Me, key, Item, Controls(key & Item).ctrl)
        On Error GoTo 0
    End If
End Property

Public Property Set Item(key As Variant, Item As Variant)
    Set keyValuePairs.Item(key).value = Item
    RaiseEvent OnItemSet(Me, key, Item)
End Property

Public Property Get Item(key As Variant) As Variant
    AssignVariable Item, keyValuePairs.Item(CStr(key)).value
    RaiseEvent OnItemGet(Me, key, Item)
End Property

Public Sub Add(key As Variant, Item As Variant, Optional Before As Variant)
    Dim oKVP As cKeyValuePair
    Set oKVP = New cKeyValuePair
    oKVP.key = key
    If IsObject(Item) Then
       Set oKVP.value = Item
    Else
       Let oKVP.value = Item
    End If
    If IsMissing(Before) Or keyValuePairs.Count = 0 Then
        keyValuePairs.Add Item:=oKVP, key:=CStr(key)
    Else
        keyValuePairs.Add Item:=oKVP, key:=CStr(key), Before:=Before
    End If
    RaiseEvent OnAdd(Me, key, Item)
End Sub

Public Property Get Exists(key As Variant) As Boolean
    On Error Resume Next
    Exists = TypeName(keyValuePairs.Item(key)) > ""
End Property

Public Sub Remove(key As Variant)
    keyValuePairs.Remove CStr(key)
End Sub

Public Sub RemoveAll()
    Set keyValuePairs = Nothing
    Set keyValuePairs = New Collection
End Sub

Public Property Get Count() As Long
    Count = keyValuePairs.Count
End Property

Public Property Get Items() As Variant
    Dim vlist As Variant, i As Long
    If Me.Count > 0 Then
        ReDim vlist(0 To Me.Count - 1)
        For i = LBound(vlist) To UBound(vlist)
           AssignVariable vlist(i), keyValuePairs.Item(1 + i).value
        Next i
        Items = vlist
        If TypeOf vlist(0) Is cDictionary Then RaiseEvent OnItemGet(Me, "", vlist(0))
    End If
End Property

Public Property Get Keys() As String()
    Dim vlist() As String, i As Long
    If Me.Count > 0 Then
        ReDim vlist(0 To Me.Count - 1)
        For i = LBound(vlist) To UBound(vlist)
           vlist(i) = keyValuePairs.Item(1 + i).key   '
        Next i
        Keys = vlist
    End If
End Property

Public Property Get KeyValuePair(Index As Long) As Variant
    Set KeyValuePair = keyValuePairs.Item(1 + Index)
End Property

Private Sub AssignVariable(variable As Variant, value As Variant)
    If IsObject(value) Then
       Set variable = value
    Else
       Let variable = value
    End If
End Sub

Private Function inCollection(COL As Object, key As Variant) As Boolean
    Dim obj As Variant
    
    On Error GoTo ERR
    obj = COL(key)
    inCollection = True
    Exit Function
ERR:
    inCollection = False
End Function

Public Sub DebugPrint()
    Dim lItem As Long, lIndex As Long, vItem As Variant, oKVP As cKeyValuePair
    lItem = 0
    For Each oKVP In keyValuePairs
        lItem = lItem + 1
        Debug.Print lItem; oKVP.key; " "; TypeName(oKVP.value);
        If InStr(1, TypeName(oKVP.value), "()") > 0 Then
            vItem = oKVP.value
            Debug.Print "("; CStr(LBound(vItem)); " to "; CStr(UBound(vItem)); ")";
            For lIndex = LBound(vItem) To UBound(vItem)
               Debug.Print " (" & CStr(lIndex) & ")"; TypeName(vItem(lIndex)); "="; vItem(lIndex);
            Next
            Debug.Print
        Else
           Debug.Print "="; oKVP.value
        End If
    Next
End Sub