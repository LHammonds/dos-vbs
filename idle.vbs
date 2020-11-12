Option Explicit
'#############################################################
'## Name          : idle.vbs
'## Version       : 1.0
'## Date          : 2018-07-11
'## Author        : LHammonds
'## Purpose       : Press the Numlock key to simulate activity to prevent the screensaver from activating.
'## Compatibility : Windows XP/8/10/2000/2003/2008/2012/2016
'## Usage         : cscript.exe idle.vbs
'## Note          : Press CTRL+C to end the script.
'######################## CHANGE LOG #########################
'## DATE       VER WHO WHAT WAS CHANGED
'## ---------- --- --- ---------------------------------------
'## 2018-07-11 1.0 LTH Created program.
'#############################################################

'## Declare variables ##
Dim objResult
Dim objShell
Dim intDelaySeconds
Dim intTicks
Dim intCounter

'## Initialize variables ##
intDelaySeconds = 60
intTicks = intDelaySeconds * 1000
intCounter = 0

Wscript.echo "Pressing NumLock every " & intDelaySeconds & " seconds..."
Wscript.echo "Press CTRL+C to stop the program and infinite loop."
Set objShell = WScript.CreateObject("WScript.Shell")    
'##  Create endless loop ##
Do While True
  intCounter = intCounter + 1
  Wscript.echo " " & intCounter
  '## Simulate pressing the NumLock key twice ##
  objResult = objShell.sendkeys("{NUMLOCK}{NUMLOCK}")
  '## Pause the script for however many seconds are noted in intDelaySeconds ##
  Wscript.Sleep (intTicks)
Loop
