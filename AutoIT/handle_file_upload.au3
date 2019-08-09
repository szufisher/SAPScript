#include <File.au3>
#include <StringConstants.au3>
Local $sFilePath ="file_in_process.txt"
Local $aRetArray
Local $file
Local $file_ok[1]

While 1
   _FileReadToArray($sFilePath, $aRetArray)
   $file = $aRetArray[1]
   If StringStripWS($file, $STR_STRIPLEADING + $STR_STRIPTRAILING) == "" Then
	  Sleep(25)
   ElseIf $file == "Finished" Then
	  Exit
   Else
	  upload($file)
	  Sleep(2000)
   EndIf
WEnd

Func upload($file)
   $title="Choose File to Upload"
   $hWnd = WinActive($title)
   If $hWnd Then
	  WinActivate($title)
	  Sleep(300)
	  ControlFocus($hWnd,"","Edit1")
	  ControlsetText($hWnd,"","Edit1",$file)
	  Sleep(300)
	  ControlClick($hWnd,"","Button1")
	  $file_ok[0] = $file&"OK"
	  _FileWriteFromArray($sFilePath, $file_ok)
   EndIf
EndFunc
