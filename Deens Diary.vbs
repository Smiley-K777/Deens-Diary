'├┘└┐┌┤┬┴┼╕╖╡═║╒╓╔╗╘╛╞╟╢╤╥╦╧╨╩╪╙╜╫╬╠╣╚╝☻
'─ ━ │ ┃ ┄ ┅ ┆ ┇ ╌ ╍ ╎ ╏ ═ ║ ╽ ╿ ▀ ▁ ▂ ▃ ▄ ▅ ▆ ▇ █ ▉ ▊ ▋ ▌ ▍ ▎ ▏ ▐ ░ ▒ ▓ ▔ ▕ ■ ☺ ➆ 〔 〕 ﾌ
',<>!?/\:;()-_+=[]{}*
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Call info
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
FName = inputbox("File Name including extention.","Deens Diary","New File.txt")
if FName = "" then Wscript.Quit
if FName = "New File.txt" then FName = "\New File"
outFile = ScriptDir & FName

Set objFSO = CreateObject("Scripting.FileSystemObject")

If not (objFSO.FileExists(outFile)) Then
	WScript.Echo("File does not exist!")
	Const Hidden = 2
	Set objFile = objFSO.CreateTextFile(outFile, True)
	objFile.Write ".LOG" & vbCrLf
	objFile.Close
	Set mapFile = objFSO.GetFile(outFile)
	mapFile.Attributes = Hidden
	CreateObject("WScript.Shell").Run Chr(34) & ScriptDir & FName & Chr(34)
	WScript.Quit()
End If
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Const ForReading = 1
Const ForWriting = 2
Filename = outFile
Set objFile = objFSO.OpenTextFile(Filename, ForReading)
strText = objFile.ReadAll
objFile.Close

MPass= "5953"
Key=inputbox("Enter Master password for Decryption.","Deens Diary","Lock Me!!!")
if Key="" then
	Wscript.Quit
elseif Key="info" then
	info
elseif Key=MPass then
	UnLock
	CreateObject("WScript.Shell").Run Chr(34) & ScriptDir & FName & Chr(34)
elseif Key="Lock Me!!!" then
	Lock
else
	Busted
end if
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Busted
strNewText = Replace(strText, "", ".LOG")
strNewText = Replace(strNewText,  "Attempted Access was denied due to incorrect [Key] that was provided.","5953")
strNewText = Replace(strNewText,  ".", "~")
strNewText = Replace(strNewText,  "~LOG", ".LOG")
strNewText = Replace(strNewText,  "PM", "▀")
strNewText = Replace(strNewText,  " ", "_")
strNewText = Replace(strNewText,  "'", "`")
strNewText = Replace(strNewText,  "!", "☻")
strNewText = Replace(strNewText,  "?", "☺")
strNewText = Replace(strNewText,  ":","┄")
strNewText = Replace(strNewText,  "/","┅")
strNewText = Replace(strNewText, "A","├獵")
strNewText = Replace(strNewText, "B","┘獵")
strNewText = Replace(strNewText, "C","└獵")
strNewText = Replace(strNewText, "D","┐獵")
strNewText = Replace(strNewText, "E","┌獵")
strNewText = Replace(strNewText, "F","┤獵")
strNewText = Replace(strNewText, "G","┬獵")
strNewText = Replace(strNewText, "H","┴獵")
strNewText = Replace(strNewText, "I","┼獵")
strNewText = Replace(strNewText, "J","╕獵")
strNewText = Replace(strNewText, "K","╖獵")
strNewText = Replace(strNewText, "L","╡獵")
strNewText = Replace(strNewText, "M","╒獵")
strNewText = Replace(strNewText, "N","╓獵")
strNewText = Replace(strNewText, "O","╔獵")
strNewText = Replace(strNewText, "P","╗獵")
strNewText = Replace(strNewText, "Q","╘獵")
strNewText = Replace(strNewText, "R","╛獵")
strNewText = Replace(strNewText, "S","╞獵")
strNewText = Replace(strNewText, "T","╟獵")
strNewText = Replace(strNewText, "U","╢獵")
strNewText = Replace(strNewText, "V","╚獵")
strNewText = Replace(strNewText, "W","╝獵")
strNewText = Replace(strNewText, "X","═獵")
strNewText = Replace(strNewText, "Y","║獵")
strNewText = Replace(strNewText, "Z","╫獵")
strNewText = Replace(strNewText, "a","├")
strNewText = Replace(strNewText, "b","┘")
strNewText = Replace(strNewText, "c","└")
strNewText = Replace(strNewText, "d","┐")
strNewText = Replace(strNewText, "e","┌")
strNewText = Replace(strNewText, "f","┤")
strNewText = Replace(strNewText, "g","┬")
strNewText = Replace(strNewText, "h","┴")
strNewText = Replace(strNewText, "i","┼")
strNewText = Replace(strNewText, "j","╕")
strNewText = Replace(strNewText, "k","╖")
strNewText = Replace(strNewText, "l","╡")
strNewText = Replace(strNewText, "m","╒")
strNewText = Replace(strNewText, "n","╓")
strNewText = Replace(strNewText, "o","╔")
strNewText = Replace(strNewText, "p","╗")
strNewText = Replace(strNewText, "q","╘")
strNewText = Replace(strNewText, "r","╛")
strNewText = Replace(strNewText, "s","╞")
strNewText = Replace(strNewText, "t","╟")
strNewText = Replace(strNewText, "u","╢")
strNewText = Replace(strNewText, "v","╚")
strNewText = Replace(strNewText, "w","╝")
strNewText = Replace(strNewText, "x","═")
strNewText = Replace(strNewText, "y","║")
strNewText = Replace(strNewText, "z","╫")
strNewText = Replace(strNewText, "0","▁")
strNewText = Replace(strNewText, "1","▂")
strNewText = Replace(strNewText, "2"," ▃")
strNewText = Replace(strNewText, "3","▄")
strNewText = Replace(strNewText, "4"," ▅")
strNewText = Replace(strNewText, "5"," ▆")
strNewText = Replace(strNewText, "6","█")
strNewText = Replace(strNewText, "7"," ▉")
strNewText = Replace(strNewText, "8"," ▊")
strNewText = Replace(strNewText, "9"," ▋")
Set objFile = objFSO.OpenTextFile(Filename, ForWriting)
objFile.WriteLine strNewText 
objFile.WriteLine Time & " " & Date
objFile.WriteLine "Attempted Access was denied due to incorrect [Key] that was provided."
objFile.Close
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub Lock
strNewText = Replace(strText, "", ".LOG")
strNewText = Replace(strNewText,  "Attempted Access was denied due to incorrect [Key] that was provided.","5953")
strNewText = Replace(strNewText,  ".", "~")
strNewText = Replace(strNewText,  "~LOG", ".LOG")
strNewText = Replace(strNewText,  "PM", "▀")
strNewText = Replace(strNewText,  " ", "_")
strNewText = Replace(strNewText,  "'", "`")
strNewText = Replace(strNewText,  "!", "☻")
strNewText = Replace(strNewText,  "?", "☺")
strNewText = Replace(strNewText,  ":","┄")
strNewText = Replace(strNewText,  "/","┅")
strNewText = Replace(strNewText, "A","├獵")
strNewText = Replace(strNewText, "B","┘獵")
strNewText = Replace(strNewText, "C","└獵")
strNewText = Replace(strNewText, "D","┐獵")
strNewText = Replace(strNewText, "E","┌獵")
strNewText = Replace(strNewText, "F","┤獵")
strNewText = Replace(strNewText, "G","┬獵")
strNewText = Replace(strNewText, "H","┴獵")
strNewText = Replace(strNewText, "I","┼獵")
strNewText = Replace(strNewText, "J","╕獵")
strNewText = Replace(strNewText, "K","╖獵")
strNewText = Replace(strNewText, "L","╡獵")
strNewText = Replace(strNewText, "M","╒獵")
strNewText = Replace(strNewText, "N","╓獵")
strNewText = Replace(strNewText, "O","╔獵")
strNewText = Replace(strNewText, "P","╗獵")
strNewText = Replace(strNewText, "Q","╘獵")
strNewText = Replace(strNewText, "R","╛獵")
strNewText = Replace(strNewText, "S","╞獵")
strNewText = Replace(strNewText, "T","╟獵")
strNewText = Replace(strNewText, "U","╢獵")
strNewText = Replace(strNewText, "V","╚獵")
strNewText = Replace(strNewText, "W","╝獵")
strNewText = Replace(strNewText, "X","═獵")
strNewText = Replace(strNewText, "Y","║獵")
strNewText = Replace(strNewText, "Z","╫獵")
strNewText = Replace(strNewText, "a","├")
strNewText = Replace(strNewText, "b","┘")
strNewText = Replace(strNewText, "c","└")
strNewText = Replace(strNewText, "d","┐")
strNewText = Replace(strNewText, "e","┌")
strNewText = Replace(strNewText, "f","┤")
strNewText = Replace(strNewText, "g","┬")
strNewText = Replace(strNewText, "h","┴")
strNewText = Replace(strNewText, "i","┼")
strNewText = Replace(strNewText, "j","╕")
strNewText = Replace(strNewText, "k","╖")
strNewText = Replace(strNewText, "l","╡")
strNewText = Replace(strNewText, "m","╒")
strNewText = Replace(strNewText, "n","╓")
strNewText = Replace(strNewText, "o","╔")
strNewText = Replace(strNewText, "p","╗")
strNewText = Replace(strNewText, "q","╘")
strNewText = Replace(strNewText, "r","╛")
strNewText = Replace(strNewText, "s","╞")
strNewText = Replace(strNewText, "t","╟")
strNewText = Replace(strNewText, "u","╢")
strNewText = Replace(strNewText, "v","╚")
strNewText = Replace(strNewText, "w","╝")
strNewText = Replace(strNewText, "x","═")
strNewText = Replace(strNewText, "y","║")
strNewText = Replace(strNewText, "z","╫")
strNewText = Replace(strNewText, "0","▁")
strNewText = Replace(strNewText, "1","▂")
strNewText = Replace(strNewText, "2"," ▃")
strNewText = Replace(strNewText, "3","▄")
strNewText = Replace(strNewText, "4"," ▅")
strNewText = Replace(strNewText, "5"," ▆")
strNewText = Replace(strNewText, "6","█")
strNewText = Replace(strNewText, "7"," ▉")
strNewText = Replace(strNewText, "8"," ▊")
strNewText = Replace(strNewText, "9"," ▋")
Set objFile = objFSO.OpenTextFile(Filename, ForWriting)
objFile.WriteLine strNewText 
objFile.WriteLine Time & " " & Date
objFile.WriteLine "File Encrypted."
objFile.Close
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub UnLock
strNewText = Replace(strText, "", ".LOG")
strNewText = Replace(strNewText,  "~", ".")
strNewText = Replace(strNewText,  "▀", "PM")
strNewText = Replace(strNewText,  "┄",":")
strNewText = Replace(strNewText,  "┅","/")
strNewText = Replace(strNewText,  "_", " ")
strNewText = Replace(strNewText,  "`", "'")
strNewText = Replace(strNewText,  "☻", "!")
strNewText = Replace(strNewText,  "☺", "?")
strNewText = Replace(strNewText, "├獵", "A")
strNewText = Replace(strNewText, "┘獵", "B")
strNewText = Replace(strNewText, "└獵", "C")
strNewText = Replace(strNewText, "┐獵", "D")
strNewText = Replace(strNewText, "┌獵", "E")
strNewText = Replace(strNewText, "┤獵", "F")
strNewText = Replace(strNewText, "┬獵", "G")
strNewText = Replace(strNewText, "┴獵", "H")
strNewText = Replace(strNewText, "┼獵", "I")
strNewText = Replace(strNewText, "╕獵", "J")
strNewText = Replace(strNewText, "╖獵", "K")
strNewText = Replace(strNewText, "╡獵", "L")
strNewText = Replace(strNewText, "╒獵", "M")
strNewText = Replace(strNewText, "╓獵", "N")
strNewText = Replace(strNewText, "╔獵", "O")
strNewText = Replace(strNewText, "╗獵", "P")
strNewText = Replace(strNewText, "╘獵", "Q")
strNewText = Replace(strNewText, "╛獵", "R")
strNewText = Replace(strNewText, "╞獵", "S")
strNewText = Replace(strNewText, "╟獵", "T")
strNewText = Replace(strNewText, "╢獵", "U")
strNewText = Replace(strNewText, "╚獵", "V")
strNewText = Replace(strNewText, "╝獵", "W")
strNewText = Replace(strNewText, "═獵", "X")
strNewText = Replace(strNewText, "║獵", "Y")
strNewText = Replace(strNewText, "╫獵", "Z")
strNewText = Replace(strNewText, "├", "a")
strNewText = Replace(strNewText, "┘", "b")
strNewText = Replace(strNewText, "└", "c")
strNewText = Replace(strNewText, "┐", "d")
strNewText = Replace(strNewText, "┌", "e")
strNewText = Replace(strNewText, "┤", "f")
strNewText = Replace(strNewText, "┬", "g")
strNewText = Replace(strNewText, "┴", "h")
strNewText = Replace(strNewText, "┼", "i")
strNewText = Replace(strNewText, "╕", "j")
strNewText = Replace(strNewText, "╖", "k")
strNewText = Replace(strNewText, "╡", "l")
strNewText = Replace(strNewText, "╒", "m")
strNewText = Replace(strNewText, "╓", "n")
strNewText = Replace(strNewText, "╔", "o")
strNewText = Replace(strNewText, "╗", "p")
strNewText = Replace(strNewText, "╘", "q")
strNewText = Replace(strNewText, "╛", "r")
strNewText = Replace(strNewText, "╞", "s")
strNewText = Replace(strNewText, "╟", "t")
strNewText = Replace(strNewText, "╢", "u")
strNewText = Replace(strNewText, "╚", "v")
strNewText = Replace(strNewText, "╝", "w")
strNewText = Replace(strNewText, "═", "x")
strNewText = Replace(strNewText, "║", "y")
strNewText = Replace(strNewText, "╫", "z")
strNewText = Replace(strNewText, "▁", "0")
strNewText = Replace(strNewText, "▂", "1")
strNewText = Replace(strNewText, " ▃", "2")
strNewText = Replace(strNewText, "▄", "3")
strNewText = Replace(strNewText, " ▅", "4")
strNewText = Replace(strNewText, " ▆", "5")
strNewText = Replace(strNewText, "█", "6")
strNewText = Replace(strNewText, " ▉", "7")
strNewText = Replace(strNewText, " ▊", "8")
strNewText = Replace(strNewText, " ▋", "9")
strNewText = Replace(strNewText,  "5953","Attempted Access was denied due to incorrect [Key] that was provided.")
Set objFile = objFSO.OpenTextFile(Filename, ForWriting)
objFile.WriteLine strNewText 
objFile.WriteLine Time & " " & Date
objFile.WriteLine "File Decrypted."
objFile.Close
Filename =Chr(34) & Filename & Chr(34)
msgbox "{Your file Location}" & VBNewLine & VBNewLine & Filename,,">>> Access Granted."
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
sub info
ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
thisFile = ScriptDir & "\Deens Diary.vbs"
Set objFSO = CreateObject("Scripting.FileSystemObject") 
'Set objFile=objFSO.GetFile("C:\Program Files\Windows Media Player\wmplayer.exe") 
Set objFile=objFSO.GetFile(thisFile)
wscript.echo "Date created: " & objFile.DateCreated            & vbCrlf & _
             "Date last accessed: " & objFile.DateLastAccessed & vbCrlf & _ 
             "Date last modified: " & objFile.DateLastModified & vbCrlf & _ 
             "Drive: " & objFile.Drive                         & vbCrlf & _ 
             "Name: " & objFile.Name                           & vbCrlf & _ 
             "Parent folder: " & objFile.ParentFolder          & vbCrlf & _ 
             "Path: " & objFile.Path                           & vbCrlf & _ 
             "Short name: " & objFile.ShortName                & vbCrlf & _ 
             "Short path: " & objFile.ShortPath                & vbCrlf & _ 
             "Size: " & objFile.Size                           & vbCrlf & _
             "Type: " & objFile.Type
end sub
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
