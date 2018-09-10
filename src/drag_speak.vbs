Option Explicit

If WScript.Arguments(0) = "/ListVoices" Then
  ShowVoices
ElseIf WScript.Arguments(0) = "/PlaySpeakTextFile" Then
  PlaySpeakTextFile WScript.Arguments(1), WScript.Arguments(2)
ElseIf WScript.Arguments(0) = "/SaveSpeakTextFile" Then
  SaveSpeakTextFile WScript.Arguments(1), WScript.Arguments(2), WScript.Arguments(3)
End If




Sub SaveSpeakTextFile(nVoice, sFromFilePath, sToFilePath)
  Dim oFileSystem
  Set oFileSystem = CreateObject("Scripting.FileSystemObject")

  Dim oText
  Set oText = oFileSystem.OpenTextFile(sFromFilePath, 1, False)

  Dim sManuscript
  sManuscript = oText.ReadAll
  oText.Close

  SaveSpeakString nVoice, sManuscript, sToFilePath
End Sub


Sub SaveSpeakString (nVoice, sManuscript, sFilePath)
    Dim oFile
    Set oFile = CreateObject("SAPI.SpFileStream")
    
    Dim oSpVoice
    Set oSpVoice = CreateObject("SAPI.SpVoice")
    Set oSpVoice.Voice = oSpVoice.GetVoices.Item(nVoice)
    oSpVoice.Rate = 1
    
    oFile.Open sFilePath, 3
    Set oSpVoice.AudioOutputStream = oFile
    oSpVoice.Speak sManuscript
    
    oFile.Close    
End Sub


Sub PlaySpeakTextFile(nVoice, sFilePath)
  Dim oFileSystem
  Set oFileSystem = CreateObject("Scripting.FileSystemObject")

  Dim oText
  Set oText = oFileSystem.OpenTextFile(sFilePath, 1, False)

  Dim sManuscript
  sManuscript = oText.ReadAll
  oText.Close

  PlaySpeakString nVoice, sManuscript
End Sub

Sub PlaySpeakString(nVoice, sManuscript)
  Dim oSpVoice
  Set oSpVoice = CreateObject("SAPI.SpVoice")
  Set oSpVoice.Voice = oSpVoice.GetVoices.Item(nVoice)
  oSpVoice.Rate = 1
  oSpVoice.Speak sManuscript
End Sub


Sub ShowVoices
  Dim oSpVoice
  Set oSpVoice = CreateObject("SAPI.SpVoice")

  Dim i, oVoice
  For i = 0 To oSpVoice.GetVoices.Count-1
    Set oVoice = oSpVoice.GetVoices.Item(i)
    WScript.Echo i & ": " & oVoice.GetDescription
  Next
End Sub

