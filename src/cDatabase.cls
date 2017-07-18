VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rec As New cDictionary
Public dbPath As String
Public dbName As String
Public structure As cDictionary
Public errorDescription As String
Public displayErrors As Boolean
Public recordSet As Object
Public rsArray As Variant
Public connection As Object
Public idField As String
Public autoIdField As Boolean

Private ADOTypes As cDictionary
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LVM_FIRST = &H1000

Public Enum eMode
    admodeRead = 1
    adModeReadWrite = 3
End Enum

Public Enum eConnectionSource
    fromACCDB = 0
    fromMDB = 1
    fromSQLServer = 2
    fromExcel = 3
    fromCSV = 4
End Enum

Public Enum eLockType
    adLockUnspecified = -1      'Unspecified type of lock. Clones inherits lock type from the original Recordset.
    adLockReadOnly = 1          'Read-only records
    adLockPessimistic = 2       'Pessimistic locking, record by record. The provider lock records immediately after editing
    adLockOptimistic = 3        'Optimistic locking, record by record. The provider lock records only when calling update
    adLockBatchOptimistic = 4   'Optimistic batch updates. Required for batch update mode
End Enum

Public Enum eCursorType
    adOpenDynamic = 2       'Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
    adOpenForwardOnly = 0   'Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
    adOpenKeyset = 1        'Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
    adOpenStatic = 3        'Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
    adOpenUnspecified = -1  'Does not specify the type of cursor.
End Enum

Private Sub Class_Initialize()
    Set ADOTypes = New cDictionary
    With ADOTypes
        .Add 2, "Integer"
        .Add 3, "Long"
        .Add 4, "Single"
        .Add 5, "Double"
        .Add 6, "Currency"
        .Add 7, "Date"
        .Add 11, "Boolean"
        .Add 17, "Byte"
        .Add 130, "String"
        .Add 131, "Long"
        .Add 135, "Date"
        .Add 200, "String"
        .Add 201, "String"
        .Add 202, "String"
        .Add 203, "String"
    End With
    autoIdField = True
End Sub

Public Property Get GetADOType(ByRef key As Long) As String
    If ADOTypes.Exists(key) Then GetADOType = ADOTypes.Item(key)
End Property

Public Function GetConnectionString(ByVal connType As eConnectionSource) As String
    Select Case connType
        Case fromACCDB
            GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
        Case fromMDB
            GetConnectionString = "Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
        Case fromSQLServer
            GetConnectionString = "Provider=SQLOLEDB;Data Source=" & dbPath & ";Initial Catalog=" & dbName & ";Integrated Security=SSPI;"
        Case fromExcel
            GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & dbPath
        Case fromCSV
    End Select
End Function

Public Function ReadRecord(ByVal tableName As String, ByVal sField As String, ByVal vValue As Variant) As Boolean
    Dim i As Long
    Dim strSql As String
    Dim ar As Variant
    
    errorDescription = ""
    
    Set rec = New cDictionary
    
    Set structure = New cDictionary
    
    ReadRecord = False
    On Error GoTo ERR
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = admodeRead  '- read, adModeReadWrite=3
       .ConnectionString = GetConnectionString(fromACCDB)  '"Data Source=" & DbPath
       .Open
    End With
    If TypeName(vValue) = "String" Then vValue = "'" & vValue & "'"
    If TypeName(vValue) = "Date" Then vValue = "#" & vValue & "#"
    
    strSql = "SELECT * FROM [" & tableName & "] where [" & sField & "]=" & vValue & ";"
    
    Set recordSet = connection.Execute(strSql)

    If Not recordSet.EOF Then
        For i = 0 To recordSet.Fields.Count - 1
            rec.Add recordSet.Fields(i).Name, recordSet.Fields(i).value
            structure.Add recordSet.Fields(i).Name, recordSet.Fields(i).Type
        Next i
    Else
        ReadRecord = False
        Exit Function
    End If
    
    DeNullStuff

    ReadRecord = True
    Exit Function

ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
    Set recordSet = Nothing
    Set connection = Nothing
End Function

Public Function DeleteRecord(ByVal tableName As String, ByVal sField As String, Optional ByVal value As Variant, Optional connType As eConnectionSource = fromACCDB) As Boolean
    Dim strSql As String
    
    errorDescription = ""
    On Error GoTo ERR
    DeleteRecord = False
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = adModeReadWrite
       .ConnectionString = GetConnectionString(connType)
       .Open
    End With
    
    If TypeName(value) = "String" Then value = "'" & value & "'"
    If TypeName(value) = "Date" Then value = "#" & value & "#"

    strSql = "DELETE * FROM [" & tableName & "] WHERE [" & sField & "]=" & value & ";"
   
    Set recordSet = CreateObject("ADODB.Recordset")
    recordSet.Open strSql, connection, 0, 3, 1

    DeleteRecord = True
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function

Public Function ClearTable(ByVal tableName As String) As Boolean
    Dim strSql As String
    Dim ar As Variant
    
    errorDescription = ""
    ClearTable = False
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = adModeReadWrite
       .ConnectionString = .ConnectionString = GetConnectionString(fromACCDB)
       .Open
    End With
    
    strSql = "DELETE * FROM [" & tableName & "];"
   
    Set recordSet = CreateObject("ADODB.Recordset")
    recordSet.Open strSql, connection, 0, 3, 1
    ClearTable = True
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function

Public Function saveRecord(ByVal tableName As String, Optional ByVal sField As String = "", Optional ByVal vValue As Variant = "", Optional ByVal addNew As Boolean = False, Optional connType As eConnectionSource = fromACCDB, Optional CursorType As eCursorType = adOpenStatic, Optional LockType As eLockType = adLockBatchOptimistic) As Boolean
    Dim strSql As String
    Dim ar As Variant
    Dim tmp As Variant
    Dim i As Long
    Dim key As Variant
    
    errorDescription = ""
    saveRecord = False
    'On Error GoTo ERR
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = adModeReadWrite
       .ConnectionString = GetConnectionString(connType)
       .Open
    End With
    
    If TypeName(vValue) = "String" Then vValue = "'" & vValue & "'"
    If TypeName(vValue) = "Date" Then vValue = "#" & vValue & "#"

    If addNew Then
        strSql = "SELECT * FROM [" & tableName & "]"
    Else
        strSql = "SELECT * FROM [" & tableName & "] where [" & sField & "]=" & vValue & ";"
    End If
    
    Set recordSet = CreateObject("ADODB.Recordset")
    'recordSet.CursorLocation = 3 ' adUseClient=3
    recordSet.Open strSql, connection, CursorType, LockType, 1
    'recordSet.Open strSql, connection, 0, 3, 1
    
    If addNew Then recordSet.addNew
    
    'DeNullStuff
    
    For Each key In rec.Keys
        If inCollection(recordSet.Fields, key) Then
            If Not recordSet.Fields(key).Properties("IsAutoIncrement").value Then
                recordSet.Fields(key).value = rec.Item(key)
            End If
        End If
    Next
    
    recordSet.Update
    
    'If addNew Then rec.Add "ID", recordSet.Fields("ID").value
    
    saveRecord = True
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function

Public Function GetList(ByVal strSql As String) As Variant
    Dim rs As Object
    Dim ar As Variant
    Dim i As Long, j As Long
    
    On Error GoTo ERR
    errorDescription = ""
    Set rec = New cDictionary
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = adModeReadWrite  '- read, adModeReadWrite=3
       .ConnectionString = GetConnectionString(fromACCDB)
       .Open
    End With
       
    Set recordSet = connection.Execute(strSql)
    Set structure = New cDictionary
    
    If Not recordSet.EOF Then
        For i = 0 To recordSet.Fields.Count - 1
            rec.Add i, recordSet.Fields(i).Name
            structure.Add recordSet.Fields(i).Name, recordSet.Fields(i).Type
        Next i
        GetList = recordSet.GetRows
        rsArray = GetList
    End If
 
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function

Public Function GetDictionary(ByVal strSql As String) As Boolean
    RsOpen strSql
    Recordset2Dictionary
    RsClose
    If errorDescription = "" Then
        GetDictionary = True
    Else
        GetDictionary = False
    End If
End Function

Public Function SaveDictionary(ByVal tableName As String, Optional ByVal addNew As Boolean = False) As Boolean
    Dim strSql As String
    Dim ar As Variant
    Dim tmp As Variant
    Dim i As Long
    Dim key As Variant
    Dim key1 As Variant
    errorDescription = ""
    SaveDictionary = False
    On Error GoTo ERR
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = adModeReadWrite
       .ConnectionString = GetConnectionString(fromACCDB)
       .Open
    End With
    
    If addNew Then
        strSql = "SELECT * FROM [" & tableName & "]"
    Else
        For Each key In rec.Keys
            i = i + 1
            ar = ar & rec.Item(key).Item(idField)
            If i <> rec.Count Then ar = ar & ","
        Next
        strSql = "SELECT * FROM [" & tableName & "] WHERE [" & idField & "] IN (" & ar & ");"
    End If
     
    Set recordSet = CreateObject("ADODB.Recordset")
    recordSet.Open strSql, connection, 0, 3, 1
    
    For Each key In rec
        If Not rec.Item(key).Exists(idField) Or addNew Then recordSet.addNew
        For Each key1 In rec.Item(key).Keys
            If inCollection(recordSet.Fields, key) Then
                If Not recordSet.Fields(key1).Properties("IsAutoIncrement").value Then
                    recordSet.Fields(key1).value = rec.Item(key).Item(key1)
                End If
            End If
        Next
    Next

    recordSet.Update

    SaveDictionary = True
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function

Public Sub RsOpen(ByVal sqlString As String, Optional connMode As eMode = admodeRead, Optional connType As eConnectionSource = fromACCDB, Optional CursorType As eCursorType = adOpenStatic, Optional LockType As eLockType = adLockBatchOptimistic)
    Dim i As Long
    
    On Error GoTo ERR
    errorDescription = ""
    
    Set rec = New cDictionary
    Set structure = New cDictionary
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = connMode
       .ConnectionString = GetConnectionString(connType)
       .Open
    End With

    Set recordSet = CreateObject("ADODB.Recordset")
    recordSet.CursorLocation = 3 ' adUseClient=3
    recordSet.Open sqlString, connection, CursorType, LockType, 1
    
    GetIdField
        
    Exit Sub

ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Sub

Public Sub RsClose()
    On Error Resume Next
    recordSet.Close
    On Error Resume Next
    connection.Close
    Set recordSet = Nothing
    Set connection = Nothing
End Sub

Public Sub Recordset2Dictionary()
    Dim tmpDict As cDictionary
    Dim i As Long, j As Long
    
    On Error GoTo ERR
    errorDescription = ""
    
    Set rec = New cDictionary
    idField = ""
    If autoIdField Then GetIdField
    
    Do While Not recordSet.EOF
        Set tmpDict = New cDictionary
        For i = 0 To recordSet.Fields.Count - 1
            tmpDict.Add recordSet.Fields(i).Name, recordSet.Fields(i).value
        Next i
        If idField <> "" Then
            rec.Add tmpDict.Item(idField), tmpDict
        Else
            If Not autoIdField Then
                rec.Add j, tmpDict
            Else
                rec.Add recordSet.Fields(0).value, tmpDict
            End If
        End If
        j = j + 1
        recordSet.MOVENEXT
    Loop
    Exit Sub
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Sub

Public Sub Dictionary2Recordset()
    Dim key1 As Variant
    Dim key2 As Variant
    Dim isInRS As Boolean
    
    'On Error GoTo ERR
    errorDescription = ""
    
    If Not recordSet.EOF Then recordSet.movefirst
    Do While Not recordSet.EOF
        isInRS = False
        If rec.Count > 0 Then
            For Each key1 In rec.Keys
                If rec.Item(key1).Item(idField) = recordSet.Fields(idField).value Then
                    isInRS = True
                    Exit For
                End If
            Next
        End If
        If Not isInRS Then recordSet.Delete
        If Not recordSet.EOF Then recordSet.MOVENEXT
    Loop
    
    GetIdField
    
    If rec.Count > 0 Then
        For Each key1 In rec.Keys
            If IsNull(rec.Item(key1).Item(idField)) Or rec.Item(key1).Item(idField) < 0 Then
                recordSet.addNew
                rec.Item(key1).Item(idField) = recordSet.Fields(idField).value
            Else
                'If Not recordSet.EOF Then
                    recordSet.movefirst
                    recordSet.Find idField & "=" & rec.Item(key1).Item(idField), , , 1
                'End If
            End If
            For Each key2 In rec.Item(key1).Keys
                If Not recordSet.Fields(key2).Properties("IsAutoIncrement").value Then
                    recordSet.Fields(key2).value = rec.Item(key1).Item(key2)
                End If
            Next
        Next
    End If
    Exit Sub
    
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Sub

Public Function BezpiekaTypy(ByRef structure As cDictionary, ByRef ctrlCol As Object) As Boolean
    Dim ctrl As Object
    BezpiekaTypy = True
    For Each ctrl In ctrlCol
        If ctrl.tag <> "" And structure.Exists(ctrl.tag) Then
            If CanConvert(ctrl, cDatabase.GetADOType(structure.Item(ctrl.tag))) Or ctrl = "" Then
                ctrl.BackColor = &H80000005
            Else
                BezpiekaTypy = False
                ctrl.BackColor = &H80FFFF
            End If
        End If
    Next
End Function

Public Function CanConvert(ByRef var As Variant, ByRef varTypeName As String, Optional length As Long = -1) As Boolean
    Dim testVar As Variant
    CanConvert = True
    On Error GoTo ERR
    Select Case varTypeName
        Case "String"
            testVar = CStr(var)
        Case "Long"
            testVar = CLng(var)
        Case "Integer"
            testVar = CInt(var)
        Case "Byte"
            testVar = CByte(var)
        Case "Boolean"
            testVar = CBool(var)
        Case "Single"
            testVar = CSng(var)
        Case "Double"
            testVar = CDbl(var)
        Case "Currency"
            testVar = CCur(var)
        Case "Date"
            testVar = CDate(var)
    End Select
    
    If length > 0 Then
        CanConvert = Len(CStr(var)) = length
    End If
    Exit Function
ERR:
    CanConvert = False
End Function

Public Sub Dictionary2Controls(ByRef rec As cDictionary, ByRef ctrlCol As Object)
    Dim ctrl As Object
    
    For Each ctrl In ctrlCol
        If rec.Exists(ctrl.tag) Or TypeName(ctrl) = "OptionButton" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    If Not IsNull(rec.Item(ctrl.tag)) Then
                        ctrl.Text = rec.Item(ctrl.tag)
                    Else
                        ctrl.Text = ""
                    End If
                Case "ComboBox"
                    'Ctrl.value = 0
                    On Error Resume Next
                    If Not IsNull(rec.Item(ctrl.tag)) And rec.Item(ctrl.tag) <> "" Then
                        ctrl.value = rec.Item(ctrl.tag)
                    Else
                        ctrl.value = 0
                    End If
                    On Error GoTo 0
                Case "Label"
                    If Not IsNull(rec.Item(ctrl.tag)) Then
                        ctrl.Caption = rec.Item(ctrl.tag)
                    Else
                        ctrl.Caption = ""
                    End If
                Case "OptionButton"
                    If ctrl.tag = rec.Item(ctrl.GroupName) Then ctrl.value = True
            End Select
        End If
    Next
    Set ctrl = Nothing
End Sub

Public Function Controls2Dictionary(ByRef rec As cDictionary, ByRef ctrlCol As Object) As Boolean
    Dim ctrl As Object
    Controls2Dictionary = True
    
    For Each ctrl In ctrlCol
        If rec.Exists(ctrl.tag) Or TypeName(ctrl) = "OptionButton" Then
            Select Case TypeName(ctrl)
                Case "TextBox"
                    If ctrl.Text = "" Then
                        rec.Item(ctrl.tag) = Null
                    Else
                        rec.Item(ctrl.tag) = ctrl.Text
                    End If
                Case "ComboBox"
                    If ctrl.value = "" Or ctrl.value = "0" Then
                        rec.Item(ctrl.tag) = Null
                    Else
                        rec.Item(ctrl.tag) = ctrl.value
                    End If
                Case "Label"
                    If ctrl.Caption = "" Then
                        rec.Item(ctrl.tag) = Null
                    Else
                        rec.Item(ctrl.tag) = ctrl.Caption
                    End If
                Case "OptionButton"
                    If ctrl.value Then rec.Item(ctrl.GroupName) = ctrl.tag
            End Select
        End If
    Next
    Set ctrl = Nothing
End Function

Public Sub Dictionary2ComboBox(ByRef ctrl As Control, ByRef dict As cDictionary, fieldName As String)
    Dim key As Variant
    If TypeName(ctrl) = "ComboBox" Then
        ctrl.Clear
        ctrl.AddItem 0
        ctrl.value = 0
        ctrl.List(ctrl.ListCount - 1, 1) = ""
        ctrl.ColumnCount = 2
        ctrl.ColumnWidths = "0;50"
        For Each key In dict.Keys
            If Not dict.Item(key).Exists(fieldName) Then Exit Sub
            If Not IsNull(dict.Item(key).Item(fieldName)) Then
                ctrl.AddItem key
                ctrl.List(ctrl.ListCount - 1, 1) = dict.Item(key).Item(fieldName)
            End If
        Next
        'Ctrl.Text = 0
    End If
End Sub

Private Sub DeNullStuff()
    Dim tmp As Variant
    For Each tmp In rec.Keys
        If IsNull(rec.Item(tmp)) Then rec.Item(tmp) = Empty
    Next
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



Public Sub DisplayOnLV(ByRef ctrl As Object, ByRef source As Variant, Optional autoSize As Boolean = False, Optional ByRef hideCol As Variant, Optional ByRef showCol As Variant)
    Dim x As Long, y As Long, key1 As Variant, key2 As Variant, i As Long, found As Boolean
    
    LockWindowUpdate ctrl.hwnd
    ctrl.Sorted = False
    ctrl.HideSelection = False
    ctrl.ColumnHeaders.Clear
    ctrl.ListItems.Clear
    ctrl.FullRowSelect = True
    ctrl.Gridlines = True
    ctrl.LabelEdit = 1
    ctrl.View = 3

    Select Case TypeName(source)
        Case "cDictionary"
            If source.Count > 0 Then
                For Each key1 In source.Items()(0).Keys
                    ctrl.ColumnHeaders.Add , key1, key1
                    If ctrl.ColumnHeaders.Count > 1 Then
                        Select Case structure.Item(key1)
                            Case 3, 4, 5, 6, 17, 131
                                 ctrl.ColumnHeaders(ctrl.ColumnHeaders.Count).Alignment = lvwColumnRight
                            Case 7, 135
                                 ctrl.ColumnHeaders(ctrl.ColumnHeaders.Count).Alignment = lvwColumnCenter
                            Case Else
                                 ctrl.ColumnHeaders(ctrl.ColumnHeaders.Count).Alignment = lvwColumnLeft
                        End Select
                    End If
                Next
                For Each key1 In source.Keys
                    x = 0
                    For Each key2 In source.Item(key1).Keys
                        If x = 0 Then
                            ctrl.ListItems.Add , "ID" & CStr(key1), key1  'source.Item(key1).Item(key2)
                        Else
                            If Not IsNull(source.Item(key1).Item(key2)) Then
                                If structure.Item(key2) = 5 Then
                                    ctrl.ListItems(ctrl.ListItems.Count).SubItems(ctrl.ColumnHeaders(key2).Index - 1) = FormatNumber(source.Item(key1).Item(key2), 2, , , vbTrue)
                                Else
                                    ctrl.ListItems(ctrl.ListItems.Count).SubItems(ctrl.ColumnHeaders(key2).Index - 1) = source.Item(key1).Item(key2)
                                End If
                            Else
                                ctrl.ListItems(ctrl.ListItems.Count).SubItems(ctrl.ColumnHeaders(key2).Index - 1) = ""
                            End If
                        End If
                        x = x + 1
                    Next
                Next
            End If
            
        Case "Variant()"
            If IsArray(rsArray) Then
                For i = 0 To structure.Count - 1
                    ctrl.ColumnHeaders.Add , , structure.Keys()(i)
                Next
        
                For y = 0 To UBound(rsArray, 2)
                    ctrl.ListItems.Add , , rsArray(0, y)
                    For x = 1 To UBound(rsArray, 1)
                        If Not IsNull(rsArray(x, y)) Then
                            ctrl.ListItems(ctrl.ListItems.Count).SubItems(x) = rsArray(x, y)
                        Else
                            ctrl.ListItems(ctrl.ListItems.Count).SubItems(x) = ""
                        End If
                    Next x
                Next y
            
                ctrl.ListItems.Add , , ctrl.ColumnHeaders(1)
                
                For x = 1 To UBound(rsArray, 1)
                    ctrl.ListItems(ctrl.ListItems.Count).SubItems(x) = ctrl.ColumnHeaders(x + 1)
                Next x
                
            End If
    End Select
    
    If autoSize Then LVAutoSizeColumn ctrl
    
    If Not IsMissing(hideCol) Then
        Select Case True
            Case IsArray(hideCol)
                For i = 0 To UBound(hideCol)
                    ctrl.ColumnHeaders(hideCol(i)).Width = 0
                Next i
            Case Else
                If inCollection(ctrl.ColumnHeaders, hideCol) Then ctrl.ColumnHeaders.Item(hideCol).Width = 0
        End Select
    End If
    
    If Not IsMissing(showCol) Then
        Select Case True
            Case IsArray(showCol)
                For i = 1 To ctrl.ColumnHeaders.Count
                    found = False
                    For x = 0 To UBound(showCol)
                        If showCol(x) = ctrl.ColumnHeaders(i) Or showCol(x) = ctrl.ColumnHeaders(i).Index Then found = True: Exit For
                    Next x
                    If Not found Then ctrl.ColumnHeaders(i).Width = 0
                Next i
            Case Else
                For i = 1 To ctrl.ColumnHeaders.Count
                    If showCol = ctrl.ColumnHeaders(i).Text Or showCol = ctrl.ColumnHeaders(i).Index Then
                    Else
                        ctrl.ColumnHeaders.Item(i).Width = 0
                    End If
                Next i
        End Select
    End If

    If ctrl.ListItems.Count > 0 Then
        ctrl.ListItems.Item(1).Selected = True
        ctrl.ListItems.Item(1).EnsureVisible
        'Ctrl.SetFocus
    End If
    DoEvents
    LockWindowUpdate 0&
End Sub

Public Sub LVAutoSizeColumn(LV As Object, Optional column As Object = Nothing)
    Dim c As Object
    
    If column Is Nothing Then
        For Each c In LV.ColumnHeaders
            SendMessage LV.hwnd, LVM_FIRST + 30, c.Index - 1, -1
        Next
    Else
        SendMessage LV.hwnd, LVM_FIRST + 30, column.Index - 1, -1
    End If
End Sub

Public Function GetStructure(ByRef tableName As String, Optional connType As eConnectionSource = fromACCDB) As cDictionary
    Dim i As Long, ColumnsSchema As Object, key As Variant
    Dim klucze As New cDictionary, wartosci As New cDictionary
    
    
    On Error GoTo ERR
    errorDescription = ""
    
    Set rec = New cDictionary
    Set structure = New cDictionary
    
    Set connection = CreateObject("ADODB.Connection")
    With connection
       .Mode = admodeRead
       .ConnectionString = GetConnectionString(connType)
       .Open
    End With
    
    Set ColumnsSchema = connection.OpenSchema(4, Array(Empty, Empty, tableName))  'adSchemaColumns=4
    Set structure = New cDictionary
    
    Do While Not ColumnsSchema.EOF
        klucze.Add ColumnsSchema.Fields("ORDINAL_POSITION"), ColumnsSchema.Fields("COLUMN_NAME").value
        wartosci.Add ColumnsSchema.Fields("ORDINAL_POSITION"), ColumnsSchema.Fields("DATA_TYPE").value
        ColumnsSchema.MOVENEXT
    Loop
    
    For i = 1 To klucze.Count
        structure.Add klucze.Item(i), wartosci.Item(i)
    Next
    '
    
    
    Set GetStructure = structure
    
    connection.Close
    Set ColumnsSchema = Nothing
    Set klucze = Nothing
    Set wartosci = Nothing
    Exit Function
ERR:
    If displayErrors Then MsgBox ERR.Description
    Debug.Print ERR.Description
    errorDescription = ERR.Description
End Function


Private Sub GetIdField()
    Dim i As Long
    If Not recordSet Is Nothing Then
        
        'If Not recordSet.EOF Then
            idField = ""
            Set structure = New cDictionary
            For i = 0 To recordSet.Fields.Count - 1
                If recordSet.Fields(i).Properties("IsAutoIncrement").value And idField = "" Then
                    idField = recordSet.Fields(i).Name
                End If
                structure.Add recordSet.Fields(i).Name, recordSet.Fields(i).Type
            Next i
        'End If
    End If
End Sub

Public Function DictFromTemplate(ByRef template As cDictionary) As cDictionary
    Dim key As Variant
    Set DictFromTemplate = New cDictionary
    
    For Each key In template.Keys
        DictFromTemplate.Add key, Null
    Next
End Function

Private Sub Class_Terminate()
    RsClose
    Set rec = Nothing
    Set structure = Nothing
    Set ADOTypes = Nothing
End Sub
