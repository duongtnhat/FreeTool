#include <File.au3>

Const $MOD = 1234567
Global $hashTable[$MOD]
Global $fileName = "input.txt"

;generateData(1000000, 6000)

Global $readFile = FileOpen(@ScriptDir & "\" & $filename, 0)

Global $hTimer = TimerInit()

Global $aInput
;_FileReadToArray($readFile, $aInput)
$aInput = FileReadToArray($readFile)
FileClose($readFile)
ConsoleWrite("File length: " & UBound($aInput) & " Line" & @CRLF)
For $i = 0 to UBound($aInput) - 1
   $line = StringLower($aInput[$i])
   if not isExist($line) Then
	  addHash($line)
   EndIf
Next
ConsoleWrite("Complete Calculate Time: " & TimerDiff($hTimer) & @CRLF)

$readFile = FileOpen(@ScriptDir & "\" & $filename, 2)
For $data in $hashTable
   if ($data <> "") Then
	  FileWrite($readFile, $data & @CRLF)
   EndIf
Next
FileClose($readFile)

ConsoleWrite("Running Time: " & TimerDiff($hTimer) & @CRLF)

Func getHash($input, $mod)
   Local $aArray = StringToASCIIArray($input)
   Local $res = 0
   For $element in $aArray
	  $res = Mod ($res * 256 + $element, $mod)
   Next
   return $res
EndFunc

Func isExist($input)
   Local $post = getHash($input, $mod)
   while ($hashTable[$post] <> "")
	  if ($hashTable[$post] == $input) Then
		 return True
	  EndIf
	  $post = Mod($post + 1, $MOD)
   WEnd
   Return False
EndFunc

Func addHash($input)
   Local $post = getHash($input, $mod)
   while ($hashTable[$post] <> "")
	  $post = Mod($post + 1, $MOD)
   WEnd
   $hashTable[$post] = $input
EndFunc

Func generateData($number, $duplicate)
   Local $readFile = FileOpen(@ScriptDir & "\" & $filename, 2)
   Local $data[$duplicate]
   Local $lenth = $duplicate
   while($duplicate > 0)
	  $duplicate -= 1
	  $data[$duplicate] = generateString(16)
   WEnd

   While($number > 0)
	  FileWriteLine($readFile, $data[Random(0, $lenth - 1, 1)])
	  $number -= 1
   WEnd

   FileClose($readFile)
EndFunc

Func generateString($length)
   $pwd = ""
   Dim $aSpace[3]
   For $i = 1 To $length
	  $aSpace[0] = Chr(Random(65, 90, 1)) ;A-Z
	  $aSpace[1] = Chr(Random(97, 122, 1)) ;a-z
	  $aSpace[2] = Chr(Random(48, 57, 1)) ;0-9
	  $pwd &= $aSpace[Random(0, 2, 1)]
   Next
   return $pwd
EndFunc


