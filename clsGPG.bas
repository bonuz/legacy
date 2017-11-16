Attribute VB_Name = "GPG"
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private sSHELL      As String

Const sLogRegistrar As String = "\reg.log"
Const sLogFirmar    As String = "\sign.log"
Const sLogListar    As String = "\lst.log"
Const sLogDesCP     As String = "\dcp.log"
Const sErrDesCP     As String = "\dcp.err"
Const sLogDes       As String = "\des.log"
Const sErrDes       As String = "\des.err"
Const okReg         As String = " secret key imported"
Const yaReg         As String = " already in secret keyring"
Const okRegEs       As String = " clave secreta importada"
Const yaRegEs       As String = " ya estaba en el almac‚n secreto"
Const maxI          As Integer = 200

'gpg --import c:\claves\secring.gpg
Public Function registerKey(ByVal sKeyPath As String) As Boolean
    Dim sPathLog As String
    Dim sLineaLog() As String
    Dim i As Integer
    
    i = 0
    
    On Error GoTo errHandler
    
    sPathLog = App.path + sLogRegistrar
    
    'Armo ejecucion del .bat
    sSHELL = App.path + "\registrar.exe """ + sKeyPath + """ """ + sPathLog + """ "
    
    'Ejecuto .bat
    Shell sSHELL, vbHide
    
    'Verifico si el log existe
    Do While Not fileExists(sPathLog)
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            registerKey = False
            'GPG: error al registrar clave privada (timeout fe)
            Logger.WriteToLogFile (ResolveResString(6001))
            Exit Function
        End If
    Loop
    
    'Verifico si el log esta lockeado
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            registerKey = False
            'GPG: error al registrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6002))
            Exit Function
        End If
    Loop
    
    'Obtengo la primer linea del log
    'sLineaLog = Split(LectorINI.leerLinea(sPathLog, 1), ":")
    
    'Verifico el resultado del registro de la clave
'    If sLineaLog(2) = okReg Or sLineaLog(2) = okRegEs Then
'        'ok
'        'GPG: clave registrada correctamente
'        Logger.WriteToLogFile (ResolveResString(6003))
'        registerKey = True
'    ElseIf sLineaLog(2) = yaReg Or sLineaLog(2) = yaRegEs Then
'        'ya registrada
'        'GPG: clave ya registrada, verificar log
'        Logger.WriteToLogFile (ResolveResString(6004) + Logger.guardarLog(sPathLog))
'        registerKey = True
'    Else
'        'error
'        'GPG: error al registrar clave, verificar log
'        Logger.WriteToLogFile (ResolveResString(6005) + Logger.guardarLog(sPathLog))
'        registerKey = False
'    End If
    
    
    'Verifico el resultado del registro de la clave
    If contarEnArchivo(sPathLog, okReg) >= 1 Or contarEnArchivo(sPathLog, okRegEs) >= 1 Then
        'ok
        'GPG: clave registrada correctamente
        Logger.WriteToLogFile (ResolveResString(6003))
        registerKey = True
    ElseIf contarEnArchivo(sPathLog, yaReg) >= 1 Or contarEnArchivo(sPathLog, yaRegEs) >= 1 Then
        'ya registrada
        'GPG: clave ya registrada, verificar log
        Logger.WriteToLogFile (ResolveResString(6004) + Logger.guardarLog(sPathLog))
        registerKey = True
    Else
        'error
        'GPG: error al registrar clave, verificar log
        Logger.WriteToLogFile (ResolveResString(6005) + Logger.guardarLog(sPathLog))
        registerKey = False
    End If
    
    'borrar log
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            registerKey = False
            'GPG: error al registrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6002))
            Exit Function
        End If
    Loop
    
    Kill sPathLog
    
    Exit Function
errHandler:
    registerKey = False
    WriteToLogFile (ResolveResString(1017, "|1", Err.Number, "|2", Err.Description))
End Function

'1: gpg --batch --yes --delete-secret-keys "9E61 795A 8A9E E326 71C3  119D FA7C 6B78 1347 48B3"
'2: gpg --delete-keys DANIOSVALDO
Public Function deregisterKey(ByVal sKeyUser As String) As Boolean

    Dim sPathLog        As String
    Dim sPathErr        As String
    Dim sKeyFingerPrint As String
    Dim sLineaLog()     As String
    Dim bDesCP          As Boolean
    Dim bDes            As Boolean
    Dim i               As Integer
    
    i = 0
    
    On Error GoTo errHandler
    
    sPathLog = App.path + sLogDesCP
    sPathErr = App.path + sErrDesCP
    
    sKeyFingerPrint = GPG.listFP

    'Armo ejecucion del .bat
    sSHELL = App.path + "\desregistrar_cp.bat """ + sKeyFingerPrint + """ """ + sPathLog + """ """ + sPathErr + """"
        
    'Ejecuto .bat
    Shell sSHELL, vbHide
    
    'Verifico si el log existe
    Do While Not fileExists(sPathLog)
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fe)
            Logger.WriteToLogFile (ResolveResString(6006))
            Exit Function
        End If
    Loop
    
    'Verifico si el log esta lockeado
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    'Verifico resultado
    'Siempre devuelve 2 logs, STDOUT y STDERR. Si es correcto ambos estan vacios.
    'Si hay error, STDERR tiene datos.
    If FileLen(sPathErr) = 0 Then
        'ok
        bDesCP = True
    Else
        bDesCP = False
        'GPG: error al desregistrar clave privada, verificar log
        Logger.WriteToLogFile (ResolveResString(6008) + Logger.guardarLog(sPathErr))
        'error
    End If
    
    'borrar log
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    i = 0
    Do While fileLocked(sPathErr) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    Kill sPathLog
    Kill sPathErr
    
    sPathLog = App.path + sLogDes
    sPathErr = App.path + sErrDes
    
    'Armo ejecucion del .bat
    sSHELL = App.path + "\desregistrar.exe " + sKeyUser + " """ + sPathLog + """ """ + sPathErr + """"
    
    'Ejecuto .bat
    Shell sSHELL, vbHide
    
      'Verifico si el log existe
    Do While Not fileExists(sPathLog)
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fe)
            Logger.WriteToLogFile (ResolveResString(6006))
            Exit Function
        End If
    Loop
    
    'Verifico si el log esta lockeado
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    'Verifico resultado
    'Siempre devuelve 2 logs, STDOUT y STDERR. Si es correcto ambos estan vacios.
    'Si hay error, STDERR tiene datos.
    If FileLen(sPathErr) = 0 Then
        'ok
        bDes = True
    Else
        'error
        bDes = False
        'GPG: error al desregistrar clave privada, verificar log
        Logger.WriteToLogFile (ResolveResString(6008) + Logger.guardarLog(sPathErr))
    End If
    
    'borrar log
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    i = 0
    Do While fileLocked(sPathErr) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            deregisterKey = False
            'GPG: error al desregistrar clave privada (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6007))
            Exit Function
        End If
    Loop
    
    Kill sPathLog
    Kill sPathErr
    
    If bDes = True And bDesCP = True Then
        'GPG: clave desregistrada correctamente
        Logger.WriteToLogFile (ResolveResString(6009))
        deregisterKey = True
    Else
        deregisterKey = False
    End If

    Exit Function
errHandler:
    deregisterKey = False
    WriteToLogFile (ResolveResString(1017, "|1", Err.Number, "|2", Err.Description))
End Function

'gpg --sign --batch --yes --passphrase 123456  c:\DigiDoc.log
Public Function signFile(ByVal sFilePath As String, ByVal sPassphrase As String) As String
    Dim sPathLog As String
    Dim i As Integer
    
    i = 0
        
    On Error GoTo errHandler
        
    sPathLog = App.path + sLogFirmar
    
    'armo ejecucion del .bat
    sSHELL = App.path + "\firmar.exe " + sPassphrase + " """ + sFilePath + """ """ + sPathLog + """"
    
    'ejecuto .bat
    Shell sSHELL, vbHide
    
    'Verifico si el log existe
    Do While Not fileExists(sPathLog)
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            signFile = ""
            'GPG: error al firmar archivo (timeout fe)
            Logger.WriteToLogFile (ResolveResString(6010))
            Exit Function
        End If
    Loop
    
    'Verifico si el log esta lockeado
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            signFile = ""
            'GPG: error al firmar archivo (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6011))
            Exit Function
        End If
    Loop
    
    'Log existe y no esta lockeado
    If FileLen(sPathLog) > 0 Then
        'hay error, leer log
        signFile = ""
        'GPG: error al firmar archivo, verificar log
        Logger.WriteToLogFile (ResolveResString(6012) + Logger.guardarLog(sPathLog))
    Else
        'log vacio, todo ok
        'If FSO.fileExists(sFilePath + ".gpg") Then
        If FSO.fileExists(sFilePath + ".sig") Then
            'devuelvo el nombre del archivo
            signFile = sFilePath + ".gpg"
            'GPG: archivo firmado correctamente
            Logger.WriteToLogFile (ResolveResString(6018))
        Else
            'error desconocido
            signFile = ""
            'GPG: error desconocido al firmar archivo, verificar log
            Logger.WriteToLogFile (ResolveResString(6013) + Logger.guardarLog(sPathLog))
        End If
    End If
    
    'borrar log
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            signFile = ""
            'GPG: error al firmar archivo (timeout fl)
            Logger.WriteToLogFile (ResolveResString(6011))
            Exit Function
        End If
    Loop
    
    Kill sPathLog
    Exit Function
errHandler:
    signFile = "2"
    WriteToLogFile (ResolveResString(1017, "|1", Err.Number, "|2", Err.Description))
End Function

'gpg --fingerprint
Public Function listFP() As String
    Dim sPathLog As String
    Dim sLineaLog() As String
    Dim i As Integer
    
    i = 0
    
    On Error GoTo errHandler
    
    sPathLog = App.path + sLogListar
    
    'armo ejecucion del .bat
    sSHELL = App.path + "\listar.exe """ + sPathLog + """"
    
    'ejecuto .bat
    Shell sSHELL, vbHide
    
   'Verifico si el log existe
    Do While Not fileExists(sPathLog)
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            listFP = ""
            'GPG: error al leer fingerprint(timeout fe)
            Logger.WriteToLogFile (ResolveResString(6014))
            Exit Function
        End If
    Loop
    
    'Verifico si el log esta lockeado
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            listFP = ""
            'GPG: error al leer fingerprint(timeout fl)
            Logger.WriteToLogFile (ResolveResString(6015))
            Exit Function
        End If
    Loop
    
    sLineaLog = Split(LectorINI.leerLinea(sPathLog, 4), "=")
    
    If UBound(sLineaLog) >= 0 And LBound(sLineaLog) >= 0 Then
        listFP = Replace(Trim(sLineaLog(1)), " ", "")
    Else
        listFP = ""
        'GPG: no se encontro fingerprint, verifique que la clave se encuente registrada
        Logger.WriteToLogFile (ResolveResString(6016))
    End If
    
    'borrar log
    i = 0
    Do While fileLocked(sPathLog) = True
        If i < maxI Then
            DoEvents
            Sleep (500)
            DoEvents
            i = i + 1
        Else
            listFP = ""
            'GPG: error al leer fingerprint(timeout fl)
            Logger.WriteToLogFile (ResolveResString(6015))
            Exit Function
        End If
    Loop
    
    Kill sPathLog
    
    Exit Function
errHandler:
    listFP = ""
    WriteToLogFile (ResolveResString(1017, "|1", Err.Number, "|2", Err.Description))
End Function

Public Function getKeys(ByVal sServer As String, ByVal sUser As String, ByVal sPwd As String, _
                         ByVal sSSHFingerPrint As String, ByVal sClavePublica As String, _
                         ByVal sClavePrivada As String, ByVal sArchivoDatos As String, _
                         ByVal sPathLocalClaves As String, ByVal sPathClaves As String, _
                         ByVal sOrigen As String) As Boolean
    Dim oSFTP As clsSFTP
    Dim bEstado As Boolean
    Dim sDetalles As String
    Dim sArchivos(2) As String
    Dim sPathLocal As String
    
    On Error GoTo errHandler
    
    Set oSFTP = New clsSFTP
    
    sPathLocal = App.path + sPathLocalClaves
    
    'verifico que la carpeta destino exista
    FSO.VerificarDirectorio (sPathLocal)
    
    sArchivos(0) = sClavePublica
    sArchivos(1) = sClavePrivada
    sArchivos(2) = sArchivoDatos
    
    ' seteo parametros
    oSFTP.setServer sServer
    oSFTP.setUser sUser
    oSFTP.setPassword sPwd
    oSFTP.setSSHFingerPrint sSSHFingerPrint
    
    ' seteo archivos y carpetas origen/destino
    oSFTP.setFile sArchivos
    oSFTP.setDestination App.path + sPathLocalClaves
    oSFTP.setSource sOrigen + sPathClaves
    oSFTP.setTransferType 2
    
    oSFTP.process
    
    bEstado = oSFTP.getStatus
    
    sDetalles = oSFTP.getDetails
    
    'verifico si el estado de la transmision es verdadera y si estan los 3 archivos
    If bEstado And FSO.fileExists(sPathLocal + sClavePublica) _
                And FSO.fileExists(sPathLocal + sClavePrivada) _
                And FSO.fileExists(sPathLocal + sArchivoDatos) Then
        WriteToLogFile ("GPG: claves descargadas correctamente")
        getKeys = True
    Else
        'GPG: error al descargar claves del servidor SFTP
        WriteToLogFile ("SFTP: " + sDetalles)
        getKeys = False
    End If
    
    Set oSFTP = Nothing
    
    Exit Function
errHandler:
    getKeys = False
    WriteToLogFile (ResolveResString(1017, "|1", Err.Number, "|2", Err.Description))
End Function
