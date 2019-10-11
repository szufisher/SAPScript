#RequireAdmin	; execute by admin account needed.
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <ListviewConstants.au3>

GUICreate("更新退税发票号", @DesktopWidth-300, 400, 40, 20, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
;GUICtrlCreateObj($oIE, 250, 20, @DesktopWidth-300,350)
GUICtrlCreateLabel("使用须知：", 20, 10, 300, 15)
GUICtrlCreateLabel("1.先登录退税软件，打开出口货物明细申报界面，确保清除了过滤字段。定位到待上传第一条记录且未在修改状态", 20, 25, 900, 25)
GUICtrlCreateLabel("2.将包含序号(需保留前缀0）与发票号两列的上传文件保存为跳格分隔的文本文件并选择该文件，确保上传记录的序号是连续的", 20, 40, 900, 25)
GUICtrlCreateLabel("3.点击upload按钮，程序运行过程中请勿再操作鼠标和键盘", 20, 55, 600, 25)
$Label3 = GUICtrlCreateLabel("Upload File", 20, 85, 60, 25)
$file_input = GUICtrlCreateInput(@WorkingDir & "\refund_invoice.txt", 85,80 ,215 , 25)
$Label4 = GUICtrlCreateLabel("Pos(X/Y)", 20, 115, 60, 25)
$x = GUICtrlCreateInput(210, 85,110 ,40 , 22)
$y = GUICtrlCreateInput(210, 130,110 ,40 , 22)
$Button1 = GUICtrlCreateButton("Select", 300, 80, 45, 25)

global $idButton_upload = GUICtrlCreateButton("Upload", 220, 110, 125, 25)

Global $msg1 = GUICtrlCreateLabel("Processing Status", 20, 140, 150, 25)
Global $msg2 = GUICtrlCreateLabel("OK count", 20, 160, 170, 25)
Global $msg3 = GUICtrlCreateLabel("Failed#", 20, 190, 150, 25)
Global $idListview = GUICtrlCreateListView("#", 20, 205, 150, 160)
GUICtrlSetColor(-1, 0xff0000)

$message = "Select File"
GUISetState(@SW_SHOW) ;Show GUI

While 1
    Local $iMsg = GUIGetMsg()
    Select
	  Case $iMsg = $Button1
			$file = FileOpenDialog($message, @WorkingDir , "Tab Delimited Text(*.txt)", 1)
			GUICtrlSetData($file_input, $file)
      Case $iMsg = $GUI_EVENT_CLOSE
            ExitLoop
			Sleep(1000)
            CheckError("System", @error, @extended)
      Case $iMsg = $idButton_upload
			Upload()
            CheckError("Upload", @error, @extended)
    EndSelect
 WEnd

Func Upload()
   $upload_file = GUICtrlRead($file_input)
   FileOpen($upload_file, 0)

   _GUICtrlListView_DeleteAllItems($idListview)
   $ok_count = 0
   $total_number = _FileCountLines($upload_file)

   For $i=2 to $total_number
	  $data = FileReadLine($upload_file, $i)
	  ;If $data = "" Then ExitLoop
	  $row = StringSplit($data, @TAB)
	  ;MsgBox(0,"info", $row[0] & "i:=" & $i & $data)
	  $msg = "Processing " & String($i-1) & "  Of  " & String($total_number-1)
	  GUICtrlSetData($msg1, $msg)
	  Local $hWnd = WinWait("生产企业出口退税申报系统", "", 10)
	  WinActivate($hWnd)
	  $iMod = Mod($i, 500)
	  ;click next page per 500 records
	  If $i> 2 and $iMod=2 Then ControlClick($hWnd,"","WindowsForms10.Window.8.app.0.33ec00f_r9_ad186")

	  If $i> 2 and $iMod<>2 Then
		 ;ControlFocus($hWnd,"","[CLASS:WindowsForms10.Window.8.app.0.33ec00f_r9_ad1; W:260; H:473]")   ; by instance , x/y/w/h not work
		 MouseClick($MOUSE_CLICK_LEFT, Number(GUICtrlRead($x)), Number(GUICtrlRead($y)))
		 Send("{DOWN}")
	  EndIf

	  ControlClick($hWnd, "", "WindowsForms10.BUTTON.app.0.33ec00f_r9_ad16")  ;click 修改
	  local $ser=ControlGetText($hWnd, "", "WindowsForms10.EDIT.app.0.33ec00f_r9_ad12")  ;check 序号
	  If $row[1] = $ser Then
		 ControlSetText($hWnd, "", "WindowsForms10.EDIT.app.0.33ec00f_r9_ad126",$row[2])
	  Else
		 GUICtrlCreateListViewItem($row[1], $idListview)
	  EndIf
	  ControlClick($hWnd, "", "WindowsForms10.BUTTON.app.0.33ec00f_r9_ad14") ;click 保存
	  Local $hWnd2 = WinWait("系统提示", "", 10)
	  WinActivate($hWnd2)
	  Send("{ENTER}")
	  If $row[2] = ControlGetText($hWnd, "", "WindowsForms10.EDIT.app.0.33ec00f_r9_ad126") Then
		 $ok_count = $ok_count + 1
		 $msg = "Uploaded " & String($ok_count) & " records OK"
		 GUICtrlSetData($msg2, $msg)
	  EndIf
   Next
   ;_FileWriteFromArray($sFilePath, $file_in_process)
   MsgBox(0,"Finished", "Upload finished!")
EndFunc

Func CheckError($sMsg, $iError, $iExtended)
    If $iError Then
        $sMsg = "Error using " & $sMsg & " button (" & $iExtended & ")"
    Else
        $sMsg = ""
    EndIf
    GUICtrlSetData($msg1, $sMsg)
EndFunc   ;==>CheckError

Func ExitScript()
    Exit
EndFunc

GUIDelete()
Exit0008
