Option Explicit

' Copyright (c) 2016 Masamitsu MURASE
' Released under the MIT license
' https://github.com/masamitsu-murase/external_editor_for_outlook2010/blob/master/README.md

'-----------------------------------------------------------------------
' Path to editor
Private Const EDITOR_PATH As String = "C:\my_program\xyzzy\xyzzy.exe"
' Directory for temporary files
Private Const TEMP_DIR As String = "M:\temp\"
' Remove temporary files
'   For fail safe, keep them now.
Private Const REMOVE_TEMP_FILES As Boolean = False

'-----------------------------------------------------------------------
Private Const POLLING_TIMER_INTERVAL_MS As Long = 500

'-----------------------------------------------------------------------
Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As LongPtr
  hStdInput As LongPtr
  hStdOutput As LongPtr
  hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
  hProcess As LongPtr
  hThread As LongPtr
  dwProcessID As Long
  dwThreadID As Long
End Type

Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
  lpApplicationName As LongPtr, ByVal lpCommandLine As String, ByVal _
  lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, _
  ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
  ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, _
  lpStartupInfo As STARTUPINFO, lpProcessInformation As _
  PROCESS_INFORMATION) As Long

Private Const NORMAL_PRIORITY_CLASS As Long = &H20

Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
  hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long

Private Const WAIT_OBJECT_0 As Long = 0

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
  hObject As LongPtr) As Long

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, _
  ByVal nIDEvent As LongPtr, _
  ByVal uElapse As Long, _
  ByVal lpTimerFunc As LongPtr) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, _
  ByVal nIDEvent As LongPtr) As Long

'-----------------------------------------------------------------------
Private gMailInfo As Object
Private gTimerId As LongPtr

'-----------------------------------------------------------------------
Private Function CommandString(filename As String) As String
  CommandString = EDITOR_PATH & " " & Chr(34) & filename & Chr(34)
End Function

Private Function TempFileName() As String
  Static counter As Long
  Dim filename As String

  counter = counter + 1
  filename = TEMP_DIR & "mail_" & Format(Now(), "yyyyMMdd_HHmmss") & "_" & counter & ".txt"

  TempFileName = filename
End Function

Private Function CurrentMailItem() As Object
  Dim inspector As Object
  Dim Item As Object

  Set inspector = Application.ActiveInspector
  If inspector Is Nothing Then
    Set CurrentMailItem = Nothing
    Exit Function
  End If

  Set Item = inspector.CurrentItem
  If Item Is Nothing Then
    Set CurrentMailItem = Nothing
    Exit Function
  End If

  ' TODO type check

  Set CurrentMailItem = Item
End Function

Private Function SaveMailToTempFile(filename As String, mailItem As Object) As Boolean
  'StreamTypeEnum
  Const adTypeBinary = 1
  Const adTypeText = 2

  'LineSeparatorsEnum
  Const adCR = 13
  Const adCRLF = -1
  Const adLF = 10

  'SaveOptionsEnum
  Const adSaveCreateNotExist = 1
  Const adSaveCreateOverWrite = 2

  Dim outStream As Object
  Set outStream = CreateObject("ADODB.Stream")

  With outStream
    .Type = adTypeText
    .Charset = "UTF-8"
    .LineSeparator = adCRLF
  End With

  Dim body As String
  body = CStr(mailItem.body)

  outStream.Open
  outStream.WriteText body
  outStream.SaveToFile filename, adSaveCreateOverWrite
  outStream.Close

  SaveMailToTempFile = True
End Function

Private Function CreateEditorProcess(filename As String) As LongPtr
  Dim proc As PROCESS_INFORMATION
  Dim start As STARTUPINFO

  start.cb = Len(start)

  Dim ret As Long
  ret = CreateProcessA(0, CommandString(filename), 0, 0, 1, NORMAL_PRIORITY_CLASS, 0, 0, start, proc)

  If ret = 0 Then
    CreateEditorProcess = 0
    Exit Function
  End If

  CreateEditorProcess = proc.hProcess
End Function

Private Sub SaveMailInfo(procHandle As LongPtr, filename As String, mailItem As Object)
  If gMailInfo Is Nothing Then
    Set gMailInfo = CreateObject("Scripting.Dictionary")
  End If

  Dim info As Object
  Set info = CreateObject("Scripting.Dictionary")
  info.Add "procHandle", procHandle
  info.Add "filename", filename
  info.Add "mailItem", mailItem
  gMailInfo.Add procHandle, info
End Sub

Private Sub LoadTempFileToMail(filename As String, mailItem As Object)
  'StreamTypeEnum
  Const adTypeBinary = 1
  Const adTypeText = 2

  'LineSeparatorsEnum
  Const adCR = 13
  Const adCRLF = -1
  Const adLF = 10

  'StreamReadEnum
  Const adReadAll = -1
  Const adReadLine = -2

  ' TODO
  '  Check whether mailItem is active or not.

  Dim outStream As Object
  Set outStream = CreateObject("ADODB.Stream")
  With outStream
    .Type = adTypeText
    .Charset = "UTF-8"
    .LineSeparator = adCRLF
  End With

  outStream.Open
  outStream.LoadFromFile filename
  Dim body As String
  body = CStr(outStream.ReadText(adReadAll))
  outStream.Close

  mailItem.body = body
End Sub

Private Sub RemoveTempFile(filename As String)
  Dim fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(filename) Then
    fso.DeleteFile filename
  End If
End Sub

Private Sub TimerProc(ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wp As LongPtr, ByVal lp As Long)
  If wp <> gTimerId Then
    Exit Sub
  End If

  On Error GoTo TIMER_PROC_ERROR

  If Not gMailInfo Is Nothing Then
    Dim keys
    keys = gMailInfo.keys
    Dim i As Long
    For i = 0 To UBound(keys)
      If WaitForSingleObject(keys(i), 0) = WAIT_OBJECT_0 Then
        CloseHandle keys(i)

        ' Editor is closed!
        Dim mailInfo As Object
        Set mailInfo = gMailInfo.Item(keys(i))
        gMailInfo.Remove keys(i)

        LoadTempFileToMail mailInfo.Item("filename"), mailInfo.Item("mailItem")

        If REMOVE_TEMP_FILES = True Then
          RemoveTempFile mailInfo.Item("filename")
        End If
      End If
    Next

    If gMailInfo.Count = 0 And gTimerId <> 0 Then
      KillTimer 0, gTimerId
      gTimerId = 0
    End If
  End If

  Exit Sub

TIMER_PROC_ERROR:

End Sub

Public Sub OpenInExternalEditor()
  Dim mailItem As Object
  Set mailItem = CurrentMailItem()
  If mailItem Is Nothing Then
    Exit Sub
  End If

  Dim filename As String
  filename = TempFileName()
  If SaveMailToTempFile(filename, mailItem) = False Then
    Exit Sub
  End If

  Dim procHandle As LongPtr
  procHandle = CreateEditorProcess(filename)
  If procHandle = 0 Then
    Exit Sub
  End If

  If gTimerId = 0 Then
    gTimerId = SetTimer(0, 0, POLLING_TIMER_INTERVAL_MS, AddressOf TimerProc)
  End If

  SaveMailInfo procHandle, filename, mailItem
End Sub

Public Sub FinishOpenInExternalEditor()
  If gTimerId <> 0 Then
    KillTimer 0, gTimerId
    gTimerId = 0
  End If

  If Not gMailInfo Is Nothing Then
    Dim keys
    keys = gMailInfo.keys
    Dim i As Long
    For i = 0 To UBound(keys)
      CloseHandle keys(i)
      gMailInfo.Remove keys(i)
    Next
  End If
End Sub
