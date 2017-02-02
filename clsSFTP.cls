VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mySession   As Session
Private sServer     As String
Private sUser       As String
Private sPassword   As String
Private sSSHFingerPrint As String
Private bStatus     As Boolean
Private sDetails    As String
Private sSource     As String
Private iErrNumber  As Integer
Private sFile()     As String
Private sDestination As String

 Private Sub Class_Initialize()
    Set mySession = New Session
    bStatus = True
    iErrNumber = 0
 End Sub

Public Sub setServer(ByVal server As String)
    sServer = server
End Sub

Public Sub setUser(ByVal user As String)
    sUser = user
End Sub

Public Sub setPassword(ByVal password As String)
    sPassword = password
End Sub

Public Sub setSSHFingerPrint(ByVal sshFP As String)
    sSSHFingerPrint = sshFP
End Sub

Public Function getStatus() As Boolean
    getStatus = bStatus
End Function

Public Function getDetails() As String
    getDetails = sDetails
End Function

Public Sub setSource(ByVal source As String)
    sSource = source
End Sub

Public Sub setFile(file() As String)
    sFile() = file()
End Sub

Public Sub setDestination(ByVal destination As String)
    sDestination = destination
End Sub

Public Sub validateParameters()
'Dim i As Integer
'
'    For i = LBound(sFile) To UBound(sFile) - 1
'            MsgBox sFile(i)
'    Next
    
    
    If sServer Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado servidor"
        iErrNumber = 1
    ElseIf sUser Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado usuario"
        iErrNumber = 2
    ElseIf sPassword Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado contraseña"
        iErrNumber = 3
    ElseIf sSSHFingerPrint Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado ssh fingerprint"
        iErrNumber = 4
    ElseIf sSource Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado carpeta origen"
        iErrNumber = 5
    ElseIf sDestination Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado carpeta destino"
        iErrNumber = 6
    ElseIf sFile Is Nothing Then
        bStatus = False
        bdetails = "No se ha indicado archivo a transmitir"
        iErrNumber = 7
    End If
    
End Sub


Public Sub Example()
 
    validateParameters
    
    If bStatus = True Then openSession sServer, sUser, sPassword, sSSHFingerPrint
    
    On Error Resume Next
    
    If bStatus = True Then upload "test.fin", "/FTP/", "C:\"
    
    ' Query for errors
    If Err.Number <> 0 Then
        bStatus = False
        sDetails = "Error: " & Err.Description
        iErrNumber = 9
        ' Clear the error
        Err.Clear
    End If
     
    closeSession
    
    ' Restore default error handling
    On Error GoTo 0
    
End Sub


Public Sub openSession(ByVal server As String, ByVal user As String, ByVal password As String, ByVal sshfingerprint As String)
    ' Enable custom error handling
   On Error Resume Next
    
    ' Setup session options
    Dim mySessionOptions As New SessionOptions
    With mySessionOptions
        .Protocol = Protocol_Sftp
        .HostName = server
        .UserName = user
        .password = password
        .SshHostKeyFingerprint = sshfingerprint
    End With
    
    ' Connect
    mySession.Open mySessionOptions
    
    If Err.Number <> 0 Then
        bStatus = False
        sDetails = "Error de conexion. " & Chr(13) & "(" & Err.Description & ")"
        iErrNumber = 8
    End If

End Sub

Public Sub closeSession()

    ' Disconnect, clean up
    mySession.Dispose

End Sub

Public Function upload(ByVal file As String, ByVal pathDest As String, ByVal pathOri As String) As String
    ' Upload files
    Dim myTransferOptions As New TransferOptions
    myTransferOptions.TransferMode = TransferMode_Binary
     
    Dim transferResult As TransferOperationResult
    Set transferResult = mySession.PutFiles(pathOri & file, pathDest & file, False, myTransferOptions)
     
    ' Throw on any error
    transferResult.Check
     
    ' Display results
    Dim transfer As TransferEventArgs
    For Each transfer In transferResult.Transfers
        bStatus = True
        sDetails = "Upload of " & transfer.fileName & " succeeded"
    Next
End Function
