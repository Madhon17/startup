VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "RDS Startup"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock ws 
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   11280
      UseMnemonic     =   0   'False
      Width           =   14880
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function addNetworkConnection Lib "Rds Library.dll" (ByVal remoteName As String, ByVal remoteNameLength As Long, ByVal user As String, ByVal userLength As Long, ByVal password As String, ByVal passwordLength As Long) As Long
Private Declare Function cancelNetworkConnection Lib "Rds Library.dll" (ByVal remoteName As String, ByVal remoteNameLength As Long) As Long
Private Declare Function mbedtlsBlowfishDecrypt Lib "Rds Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String, ByVal index0 As Long, ByVal index1 As Long) As Long
Private Declare Function opensslAESDecrypt Lib "Rds Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String, ByVal index0 As Long, ByVal index1 As Long) As Long
Private Declare Function opensslSHA512Digest Lib "Rds Library.dll" (ByVal inputData As String, ByVal inputLength As Long, ByVal outputData As String) As Long
Private Declare Function shellExecuteEx Lib "Rds Library.dll" (ByVal hwnd As Long, ByVal file As String, ByVal fileLength As Long, ByVal wait As Long) As Long

Private wsPasswordBytesTotal As Long
Private wsPasswordDataLength As Long
Private wsPasswordChallenge As Boolean
Private host As String
Private host2 As String
Private copyDirectory As String
Private openFile As String
Private connectTryCount As Long

Private Sub Form_Load()
  
  On Error Resume Next
  
  If App.PrevInstance = True Then
    Exit Sub
  End If
  
  
  host = GetIniSetting("Host", "10.0.0.201")
  host2 = GetIniSetting("Backup", "10.0.0.202")
  
  copyDirectory = GetIniSetting("Directory", "vod\1 Screen")
  openFile = GetIniSetting("Open", "Karaoke.exe")
  
  Sleep GetIniSetting("delayStart", "0")
  
  
  connectTryCount = 1
  
  wsPasswordBytesTotal = 0
  wsPasswordDataLength = 0
  wsPasswordChallenge = False

  ws.RemotePort = 32768
  
  ws.Connect host
  
  lbl.Caption = "Connecting to Outlet Server on " & ws.RemoteHost & ":" & ws.RemotePort
  
  If Err.Number <> 0 Then
    LogError Name, "Form_Load"
  End If
End Sub

Private Function GetIniSetting(Key As String, default As String) As String
  
  On Error Resume Next
  

  Dim strBuffer As String
  Dim lLength As Long
  Dim BufferSize As Long
  
  BufferSize = 2048
  
  strBuffer = Space(BufferSize)

  lLength = GetPrivateProfileString("DataSource", Key, default, strBuffer, BufferSize, App.Path & "\setting-startup.ini")
  
  GetIniSetting = Left(strBuffer, lLength)
End Function

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
  
  On Error Resume Next
  
  Dim dataLength As Long
  Dim Data As String * 32767
  Dim buffer As String * 32767
  Dim challengeCode As String * 127
  Dim byteArray() As Byte
  Dim Key As String
  Dim a As Long
  Dim result As Long
  Dim remoteName As String
  Dim user As String
  Dim sourceDirectory As String
  Dim fso As Scripting.FileSystemObject
  Dim destinationDirectory As String
  Dim executableFile As String

  If wsPasswordBytesTotal = 0 Then
    If wsPasswordDataLength = 0 Then
      ws.GetData wsPasswordDataLength, vbLong, 4
      If bytesTotal > 4 Then
        wsPasswordBytesTotal = bytesTotal - 4
      End If
    Else
      wsPasswordBytesTotal = bytesTotal
    End If
  Else
    wsPasswordBytesTotal = wsPasswordBytesTotal + bytesTotal
  End If
  
  Sleep 2 ' do not remove, strange
  
  If wsPasswordBytesTotal = wsPasswordDataLength Then
    
    Data = Space$(16384)
    
    ws.GetData Data, vbString, wsPasswordDataLength
    
    If wsPasswordChallenge = False Then
      
      lbl.Caption = "Receiving data from " & ws.RemoteHost & " ..."

      buffer = Space$(16384)
      
      dataLength = opensslAESDecrypt(Data, wsPasswordDataLength, buffer, 1073741822, 1073741821)
      If dataLength <> 4 Then
        MsgBox "opensslAESDecrypt error: " & dataLength
        End
      End If
      
      challengeCode = Space$(127)
      
      If opensslSHA512Digest(buffer, dataLength, challengeCode) <> 64 Then
        MsgBox "opensslSHA512Digest error"
        End
      End If
      
      dataLength = 64
      ws.SendData dataLength
      
      ReDim byteArray(63)
      For a = 0 To 63
        byteArray(a) = Asc(Mid(challengeCode, a + 1, 1))
      Next
      ws.SendData byteArray
      
      
      buffer = "getKeyWindows"
      dataLength = 13
      ws.SendData dataLength
      
      ReDim byteArray(dataLength - 1)
      For a = 0 To dataLength - 1
        byteArray(a) = Asc(Mid(buffer, a + 1, 1))
      Next
      ws.SendData byteArray
      
      wsPasswordChallenge = True
    
    Else
      
      buffer = Space$(16384)
  
      result = mbedtlsBlowfishDecrypt(Data, wsPasswordDataLength, buffer, 1073741823, 1073741824)
      If result <> 2048 Then
        MsgBox "mbedtlsBlowfishDecrypt fail: " & result & ", " & Err.Description
        End
      End If
      
      Key = ""
      dataLength = (Asc(Mid(buffer, 1, 1)) * 1) + (Asc(Mid(buffer, 2, 1)) * 256) + (Asc(Mid(buffer, 3, 1)) * 65536) + (Asc(Mid(buffer, 4, 1)) * 16777216)
      For a = 1 To dataLength
        Key = Key & Mid(buffer, a + 4, 1)
      Next
      
      ws.Close
      
      
      remoteName = "\\" & ws.RemoteHost
      user = "Administrator"
      
      lbl.Caption = "Opening file share " & remoteName & " ..."
      DoEvents
      
      cancelNetworkConnection remoteName, Len(remoteName)
  
      result = addNetworkConnection(remoteName, Len(remoteName), user, Len(user), Key, Len(Key))
      If result <> 0 Then
        LogText "addNetworkConnection fail: " & result
      End If
      
      
      lbl.Caption = "Copying file ..."

      Set fso = New Scripting.FileSystemObject
      
      sourceDirectory = remoteName & "\" & copyDirectory
      destinationDirectory = App.Path & "\" & copyDirectory
      
      lbl.Caption = "Copying file from " & sourceDirectory & " to " & destinationDirectory & " ..."
      DoEvents
      
      If fso.FolderExists(sourceDirectory) = False Then
        MsgBox "'" & sourceDirectory & "' not found"
        End
      End If
      
      createDirectory fso, destinationDirectory
      fso.CopyFolder sourceDirectory, destinationDirectory, True
      
      cancelNetworkConnection remoteName, Len(remoteName)
      
      
      executableFile = destinationDirectory & "\" & openFile
      
      lbl.Caption = "Executing program ..."
      DoEvents
      
      shellExecuteEx hwnd, executableFile, Len(executableFile), 0
      
      If Err.Number <> 0 Then
        LogError Name, "ws_DataArrival"
        End
      End If
      
      End
    End If
    
    wsPasswordBytesTotal = 0
    wsPasswordDataLength = 0
  End If
  
  If Err.Number <> 0 Then
    LogError Name, "ws_DataArrival"
  End If
End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  
  On Error Resume Next
  
  If (Number = 10060) Or (Number = 10061) Then
    
    If ws.RemoteHost = host Then
      ws.Close
      ws.Connect host2
    Else
      ws.Close
      ws.Connect host
      connectTryCount = connectTryCount + 1
    End If
    
    lbl.Caption = "Retry " & connectTryCount & " connecting to Outlet Server on " & ws.RemoteHost & ":" & ws.RemotePort
    
  Else
  
    lbl.Caption = "Socket error " & Number & ": " & Description
  
  End If
  
  
  LogError Name, "ws_Error " & Number & ": " & Description
    
  If connectTryCount = 5 Then
    MsgBox "Network error " & Number & ": " & Description, vbCritical Or vbOKOnly
    End
  End If
End Sub

Private Sub createDirectory(fso As Scripting.FileSystemObject, directoryPath As String)
  If fso.FolderExists(directoryPath) = False Then
    createDirectory fso, fso.GetParentFolderName(directoryPath)
    fso.CreateFolder directoryPath
  End If
End Sub


Sub LogText(text As String)

  On Error GoTo hell
  
  Dim fs As New Scripting.FileSystemObject
  Dim drv As Scripting.Drive
  Dim driveSpec As String
  Dim fl As Scripting.TextStream
  Dim fileName As String
  
  fileName = ""
  driveSpec = "D"
  
  If fs.DriveExists(driveSpec) = True Then
    
    Set drv = fs.GetDrive(driveSpec)
    
    If drv.IsReady = True Then
      
      fileName = driveSpec & ":"
    End If
  End If
  
  If fileName = "" Then
    fileName = App.Path
  End If
  
  fileName = fileName & "\log.txt"
  
  Set fl = fs.OpenTextFile(fileName, 8, True)
  fl.WriteLine Format$(Now, "yyyy-MM-dd hh:mm:ss") & " " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & ", " & text
  fl.Close
  Set fl = Nothing
  
  Set fs = Nothing
  
hell:
  
End Sub

Sub LogError(fileName As String, procedureName As String)
  LogText fileName & "." & procedureName & ", Number: " & Err.Number & ", LastDllError: " & Err.LastDllError & ", Description: " & Err.Description
End Sub
