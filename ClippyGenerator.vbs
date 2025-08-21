'===========================
' Clippy Dialogue Generator
'===========================

Option Explicit

Dim objAgent, objClippy
Dim animList, userAnim, userText, confirmRun, preview, exitCmd

' Create MS Agent control
Set objAgent = CreateObject("Agent.Control.2")
objAgent.Connected = True

' Load Clippy (clippit.acs must be in the same folder)
objAgent.Characters.Load "Clippy", "clippit.acs"
Set objClippy = objAgent.Characters("Clippy")

' Predefined list of common animations
animList = "Greet, Wave, RestPose, Idle1, Idle2, GetAttention, Searching, Writing, Congratulate"

' Ask user for an animation
userAnim = InputBox("Enter the animation name for Clippy:" & vbCrLf & animList, "Choose Animation")

' Ask user for optional dialogue (Think bubble only; no TTS)
userText = InputBox("Enter text for Clippy to show (leave blank for none):", "Clippy Dialogue")

' Build confirmation text
preview = "Animation: " & userAnim & vbCrLf
If userText = "" Then
    preview = preview & "Dialogue: (none)"
Else
    preview = preview & "Dialogue: " & userText
End If
preview = preview & vbCrLf & vbCrLf & "Run dialogue now?"

' Confirm before running
confirmRun = MsgBox(preview, vbOKCancel + vbQuestion, "Clippy Dialogue Generator")

If confirmRun = vbOK Then
    ' Show Clippy only when ready to perform
    objClippy.Show

    ' Perform animation if provided
    If userAnim <> "" Then
        On Error Resume Next
        objClippy.Play userAnim
        If Err.Number <> 0 Then
            MsgBox "Invalid animation name. Try one from the list shown.", vbExclamation, "Animation Error"
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' Show optional dialogue (bubble only, no TTS)
    If userText <> "" Then
        objClippy.Think userText
    End If

    ' Loop until user types Delete
    Do
        exitCmd = InputBox("Type DELETE to remove Clippy, or leave blank to keep him around.", "Clippy Control")
        If UCase(exitCmd) = "DELETE" Then
            objClippy.Hide
            Exit Do
        End If
    Loop
End If

Set objClippy = Nothing
Set objAgent = Nothing
