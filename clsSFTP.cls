Option Explicit

'*******************************clsSFTP*********************************
'Description: allows to send files via SFTP using WinSCP .NET wrapper
'Date: 30/01/2017
'Author: FNC
'Version: 0.1b
'Comments: -
'***********************************************************************

Private mySession       As Session
Private sServer         As String
Private sUser           As String
Private sPassword       As String
Private sSSHFingerPrint As String
Private sFile()         As String 'sends all the files in the array
Private sDestination    As String
Private sDetails        As String
Private sSource         As String
Private bStatus         As Boolean
Private iErrNumber      As Integer


'Constructor w/o parameters
 Private Sub Class_Initialize()
    Set mySession = New Session
    bStatus = True
    iErrNumber = 0
 End Sub

'Sets the server address
Public Sub setServer(ByVal server As String)
    sServer = server
End Sub

'Sets the user name
Public Sub setUser(ByVal user As String)
    sUser = user
End Sub

'Sets the password
Public Sub setPassword(ByVal password As String)
    sPassword = password
End Sub

'Sets the SFTP server SSH Finger Print
Public Sub setSSHFingerPrint(ByVal sshFP As String)
    sSSHFingerPrint = sshFP
End Sub

'Sets the source folder of the file
Public Sub setSource(ByVal source As String)
    sSource = source
End Sub

'Sets the file name
Public Sub setFile(file() As String)
    sFile() = file()
End Sub

'Sets the destiny in the SFTP server
Public Sub setDestination(ByVal destination As String)
    sDestination = destination
End Sub

'Gets the status true/false, false means error
Public Function getStatus() As Boolean
    getStatus = bStatus
End Function

'Gets the status detail if anyone
Public Function getDetails() As String
    getDetails = sDetails
End Function

'Gets the error number
Public Function getErrNumber() As Integer
    getErrNumber = iErrNumber
End Function

'Checks if needed parameters are setted
Public Sub validateParameters()
'Dim i As Integer
'
'    For i = LBound(sFile) To UBound(sFile) - 1
'            MsgBox sFile(i)
'    Next
    
    If sServer Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado servidor"
        iErrNumber = 1
    ElseIf sUser Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado usuario"
        iErrNumber = 2
    ElseIf sPassword Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado contraseña"
        iErrNumber = 3
    ElseIf sSSHFingerPrint Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado ssh fingerprint"
        iErrNumber = 4
    ElseIf sSource Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado carpeta origen"
        iErrNumber = 5
    ElseIf sDestination Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado carpeta destino"
        iErrNumber = 6
    ElseIf sFile Is Nothing Then
        bStatus = False
        sDetails = "No se ha indicado archivo a transmitir"
        iErrNumber = 7
    End If
    
End Sub


Public Sub Example()
 
    validateParameters
    
    If bStatus = True Then openSession sServer, sUser, sPassword, sSSHFingerPrint
    
    On Error Resume Next
    
    If bStatus = True Then
        upload "test.fin", "/FTP/", "C:\"
        
    End If
    
    ' Query for errors
    If Err.Number <> 0 Then
        bStatus = False
        sDetails = Err.Description
        iErrNumber = 9
        ' Clear the error
        Err.Clear
    End If
     
    closeSession
    
    ' Restore default error handling
    On Error GoTo 0
    
End Sub

'Opens SFTP session
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
        'sDetails = "Error de conexion. " & Chr(13) & "(" & Err.Description & ")"
        sDetails = Err.Description
        iErrNumber = 8
    End If

End Sub

'Closes SFTP session
Public Sub closeSession()

    ' Disconnect, clean up
    mySession.Dispose

End Sub

'Uploads a file
Private Function upload(ByVal file As String, ByVal pathDest As String, ByVal pathOri As String) As String
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
