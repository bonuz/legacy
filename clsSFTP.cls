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

'*******************************clsSFTP*********************************
'Description: allows to send files via SFTP using WinSCP .NET wrapper
'Date: 30/01/2017
'Author: FNC
'Version: 0.1b
'Comments: -
'***********************************************************************

'Const iUploadType As Integer = 1
'Const iDownloadtype As Integer = 2

Private mySession       As Session
Private sServer         As String
Private sUser           As String
Private sPassword       As String
Private sSSHFingerPrint As String
Private sFile()         As String 'sends all the files in array
Private sDestination    As String
Private sDetails        As String
Private sSource         As String
Private bStatus         As Boolean
Private iTransferType   As Integer
Private iErrNumber      As Integer

'Constructor w/o parameters
Private Sub Class_Initialize()
   Set mySession = New Session
   bStatus = True
   iErrNumber = 0
End Sub

'Sets the server address
Public Sub setTransferType(ByVal type_t As Integer)
    iTransferType = type_t
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

'Checks if needed parameters are setted for upload
Private Sub validateParameters()
   
On Error GoTo errHandler
   
    If sServer = "" Then
        bStatus = False
        sDetails = "No se ha indicado servidor"
        iErrNumber = 1
    ElseIf sUser = "" Then
        bStatus = False
        sDetails = "No se ha indicado usuario"
        iErrNumber = 2
    ElseIf sPassword = "" Then
        bStatus = False
        sDetails = "No se ha indicado contraseña"
        iErrNumber = 3
    ElseIf sSSHFingerPrint = "" Then
        bStatus = False
        sDetails = "No se ha indicado ssh fingerprint"
        iErrNumber = 4
    ElseIf sSource = "" Then
        bStatus = False
        sDetails = "No se ha indicado carpeta origen"
        iErrNumber = 5
    ElseIf sDestination = "" Then
        bStatus = False
        sDetails = "No se ha indicado carpeta destino"
        iErrNumber = 6
    ElseIf sFile(0) = "" Then
        bStatus = False
        sDetails = "No se ha indicado archivo a transmitir"
        iErrNumber = 7
    ElseIf iTransferType = 0 Then
        bStatus = False
        sDetails = "No se ingreso tipo de transferencia"
        iErrNumber = 11
    End If
    
Exit Sub
errHandler:
    bStatus = False
    sDetails = "Error al validar parametros"
    iErrNumber = 10
End Sub


'Opens SFTP session
Private Sub openSession(ByVal server As String, ByVal user As String, ByVal password As String, ByVal sshfingerprint As String)
    ' Enable custom error handling
   On Error Resume Next
    
    ' Setup session options
    Dim mySessionOptions As New SessionOptions
    With mySessionOptions
        .Protocol = Protocol_Sftp
        .HostName = server
        .username = user
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
Private Sub closeSession()

    ' Disconnect, clean up
    mySession.Dispose

End Sub

'Uploads a file
Private Function upload(ByVal file As String, ByVal pathDest As String, ByVal pathOri As String) As String
    On Error Resume Next
    
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

'Download a file
Private Function download(ByVal file As String, ByVal pathDest As String, ByVal pathOri As String) As Boolean
    On Error Resume Next
    
    bStatus = False
    
    ' Download files
    Dim myTransferOptions As New TransferOptions
    myTransferOptions.TransferMode = TransferMode_Binary
 
    Dim transferResult As TransferOperationResult
    Set transferResult = mySession.GetFiles(pathOri + file, pathDest, False, myTransferOptions)
       
    ' Throw on any error
    transferResult.Check
     
    ' Display results
    Dim transfer As TransferEventArgs
    For Each transfer In transferResult.Transfers
        bStatus = True
        sDetails = "Download of " & transfer.fileName & " succeeded"
    Next
    
    download = bStatus
End Function

Public Sub process()
    Dim i As Integer
 
    validateParameters
    
    If bStatus = True Then openSession sServer, sUser, sPassword, sSSHFingerPrint
    
    On Error Resume Next
    
    If bStatus = True Then
        For i = LBound(sFile) To UBound(sFile)
            If iTransferType = 1 Then
                upload sFile(i), sDestination, sSource
            ElseIf iTransferType = 2 Then
                If download(sFile(i), sDestination, sSource) <> True Then
                    Exit For
                End If
            End If
        Next
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
