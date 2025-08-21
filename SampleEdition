' ClippyGenerator.vbs
' Ready-to-run sample for Clippy in Windows XP

Option Explicit

Dim clippy
Dim msg, anim

' Create the Clippy object
Set clippy = CreateObject("Agent.Control.2")

' Load Clippy character (make sure CLIPPIT.ACS is in C:\WINDOWS\msagent\chars)
clippy.Characters.Load "CLIPPIT"

' Show Clippy
clippy.Characters("CLIPPIT").Show

' Sample dialogues with animations
' Format: clippy.Characters("CLIPPIT").Play "AnimationName"
' Common animations: Greet, GetAttention, GestureUp, Writing

' Dialogue 1
clippy.Characters("CLIPPIT").Play "Greet"
clippy.Characters("CLIPPIT").Speak "Hello! I'm Clippy, your assistant."

' Dialogue 2
clippy.Characters("CLIPPIT").Play "GetAttention"
clippy.Characters("CLIPPIT").Speak "Want some tips on using Windows XP?"

' Dialogue 3
clippy.Characters("CLIPPIT").Play "GestureUp"
clippy.Characters("CLIPPIT").Speak "Don't forget to save your files often!"

' Dialogue 4
clippy.Characters("CLIPPIT").Play "Writing"
clippy.Characters("CLIPPIT").Speak "Let's have fun with some animations!"

' Optional: keep Clippy on screen until user closes
MsgBox "Clippy is ready! Close this box to hide him.", vbInformation, "Clippy Generator"

' Hide Clippy
clippy.Characters("CLIPPIT").Hide
