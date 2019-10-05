VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Private oConn As New ADODB.Connection
    Private oRs As New ADODB.Recordset
    Private oStrm As New ADODB.Stream
    Private oCmd As New ADODB.Command
    Private sErrMess As String
    Private ncmdTimeOut As Integer
    Private scmdTimeOut As String
    Private sDataserver As String
    Private sDatabasename As String
    Private sText As String
    Private sUsername As String
    Private sUserpassword As String
  


Public Function CloseDatabase() As Boolean
    On Error GoTo CloseDatabaseErr
    
    If oConn.State = adStateOpen Then
        oConn.Close
    End If
    'Set oConn = Nothing
    CloseDatabase = True

Exit Function
CloseDatabaseErr:
    sErrMess = Error
    CloseDatabase = False
End Function
Public Function GetXML(ByVal nToken As Integer, _
        ByVal xmlRoot As String, _
        ByVal sCommand As String) As String
    On Error GoTo xmlerror
    sErrMess = ""

    If Not OpenDatabase() Then
        Exit Function
    End If
    
    ' check for query type
    'If InStr(1, LCase(sCommand), "insert") > 0 Or _
    '        InStr(1, LCase(sCommand), "update") > 0 Or _
    '        InStr(1, LCase(sCommand), "delete") > 0 Then
    '    sErrMess = "That type of query is not allowed."
    '    GetXML = ""
    '    Exit Function
    'End If
    
    ' remove spaces in the root
    xmlRoot = Replace(xmlRoot, " ", "")
    
    Dim sStr As New cStringBuffer
    sStr.Append "<" + xmlRoot + ">"
    With oStrm
        .Type = adTypeText
        .LineSeparator = adCRLF
        .Mode = adModeRead
        .Open
    End With
    On Error GoTo closestreamobject
    
    ncmdTimeOut = 400
    If IsNumeric(scmdTimeOut) Then
        ncmdTimeOut = CInt(scmdTimeOut)
    End If
    With oCmd
        Set .ActiveConnection = oConn
        .Properties("Output Stream").Value = oStrm
        .CommandText = sCommand
        .CommandTimeout = ncmdTimeOut
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
    oConn.Close
    oStrm.Position = 0
 
    Do While Not oStrm.EOS
        sStr.Append oStrm.ReadText(100000)
    Loop
    oStrm.Close
    
    On Error GoTo xmlerror
    sStr.Append "</" + xmlRoot + ">"
    GetXML = sStr.Value
    sStr.Clear
    
    Exit Function
    
closestreamobject:
    sErrMess = Error
    On Error GoTo xmlerror
    If IsObject(oStrm) Then
        If oStrm.State = adStateOpen Then
            oStrm.Close
        End If
    End If
    GetXML = ""
    Exit Function
xmlerror:
    sErrMess = Error
    GetXML = ""
End Function
Private Function OpenDatabase() As Boolean
    On Error GoTo OpenErr
    sErrMess = ""
    If oConn.State >= adStateOpen Then
        OpenDatabase = True
        Exit Function
    End If
    If DataServer = "" Then
        OpenDatabase = False
        sErrMess = "You must specify a data server name."
        Exit Function
    End If
    '        "Persist Security Info=False;"
    If cmdTimeOut = "" Then
      sText = ""
    Else
      sText = ";CommandTimeout=" + cmdTimeOut
    End If
    With oConn
        .ConnectionString = "Provider=SQLOLEDB.1;" + _
            "Integrated Security=SSPI;" + _
            "Initial Catalog=" + DatabaseName + ";" + _
            "Data Source=" + DataServer + sText
        .Open
    End With

    ' set the app role
    On Error GoTo AppContinue:
    'sEPCUserName = ""
    'If NTUserName > " " And LCase(EPCUserName) = DatabaseName + "admin" Then
     '   oConn.Execute ("sp_setapprole '" + NTUserName + "', {Encrypt N'enforce" + NTUserName + "'}, 'odbc'")
    '    nUserAccessLevel = 2
    'End If
    'Set oRS = oConn.Execute("select user 'username'")
   'sEPCUserName = oRS.Fields("username")
AppContinue:
    On Error GoTo OpenErr
    OpenDatabase = True
    
Exit Function
OpenErr:
    OpenDatabase = False
    sErrMess = Error
End Function
Public Function SendXML(ByVal sTableName As String, _
        ByVal nUpdateType As Integer, _
        ByVal sXML As String) As Boolean
    ' Update types:
        ' 1 = update, 2 = insert, 3 = delete
    On Error GoTo xmlerror
    sErrMess = ""
    If Not OpenDatabase() Then
        Exit Function
    End If
    sXML = "up_xmlupdate '" + sTableName + "', " + _
        CStr(nUpdateType) + ", '" + sXML + "'"
    oConn.Execute (sXML)
    SendXML = True
    
    Exit Function
xmlerror:
    sErrMess = Error
    SendXML = False
End Function
Private Function Setup() As Boolean
    sErrMess = ""

    If Not OpenDatabase() Then
        Setup = False
    Else
        Setup = True
    End If
    
End Function
Public Function UserLogin(ByVal sLoginname As String, _
        ByVal sPassword As String) As Boolean
    On Error GoTo LoginErr
    'close the database
    'If Not CloseDatabase() Then
    '    UserLogin = False
    '    Exit Function
    'End If
    sUsername = sLoginname
    sUserpassword = sPassword
    If Not OpenDatabase() Then
        UserLogin = False
        Exit Function
    End If
    UserLogin = True
    
Exit Function
LoginErr:
    sErrMess = Error
    UserLogin = False
End Function
'======= PROPERTIES =================================================
Public Property Get DataServer() As String
    DataServer = sDataserver
End Property
Public Property Let DataServer(ByVal sData As String)
    sDataserver = sData
End Property
Public Property Get DatabaseName() As String
    DatabaseName = sDatabasename
End Property
Public Property Let DatabaseName(ByVal sData As String)
    sDatabasename = sData
End Property
Public Property Get cmdTimeOut() As String
    cmdTimeOut = scmdTimeOut
End Property
Public Property Let cmdTimeOut(ByVal sData As String)
    scmdTimeOut = sData
End Property
Private Property Get UserName() As String
    UserName = sUsername
End Property
Private Property Let UserName(ByVal sData As String)
    sUsername = sData
End Property
Private Property Get UserPassword() As String
    UserPassword = sUserpassword
End Property
Private Property Let UserPassword(ByVal sData As String)
    sUserpassword = sData
End Property
Public Property Get ErrorMessage()
    ErrorMessage = sErrMess
End Property


Private Sub Class_Terminate()
    CloseDatabase
End Sub
