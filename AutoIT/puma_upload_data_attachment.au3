#include <GUIConstantsEx.au3>
#include <IE.au3>
#include "WinHttp.au3"
#include <WindowsConstants.au3>
#include <File.au3>
#include <Array.au3>
#include <GuiListView.au3>
#include <ListviewConstants.au3>

If @Compiled Then
   RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", @ScriptName, "REG_DWORD", 11001)
Else
   RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", "AutoIt3.exe", "REG_DWORD", 11001)
   RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION", "autoit3_x64.exe", "REG_DWORD", 11001)
EndIf

Local $oIE = _IECreateEmbedded()
GUICreate("Upload PUMA", @DesktopWidth-50, 550, 40, 20, $WS_OVERLAPPEDWINDOW + $WS_CLIPSIBLINGS + $WS_CLIPCHILDREN)
GUICtrlCreateObj($oIE, 250, 20, @DesktopWidth-300,350)
$Label1 = GUICtrlCreateLabel("Home Page", 20, 20, 60, 25)
$HomePage = GUICtrlCreateInput("https://opw.siemens.com/opw/Default.aspx", 85, 20, 115, 25)
Local $idButton_Go = GUICtrlCreateButton("Go", 200, 20, 45, 25)

$Label3 = GUICtrlCreateLabel("Upload File", 20, 50, 60, 25)
$Puma = GUICtrlCreateInput(@WorkingDir & "\puma.txt", 85,50 ,115 , 25)
$Button1 = GUICtrlCreateButton("Select", 200, 50, 45, 25)

$Label4 = GUICtrlCreateLabel("Steps", 20, 80, 60, 25)
$Step = GUICtrlCreateInput(@WorkingDir & "\step_puma.txt", 85, 80, 115, 25)
$Button2 = GUICtrlCreateButton("Select", 200, 80, 45, 25)

Local $idButton_Start = GUICtrlCreateButton("Upload Actual+Forecast",80, 108, 140, 20)

$Label5 = GUICtrlCreateLabel("Fiscal Year", 20, 135, 60, 25)
$year = GUICtrlCreateInput("", 90, 130, 60, 25)
$idButton_upload_forecast = GUICtrlCreateButton("Upload Forecast", 160, 130, 90, 25)

$listview = GUICtrlCreateListView("Month",20,160,70,260, -1, $LVS_EX_CHECKBOXES )
_GUICtrlListView_SetColumnWidth($listview, 0, $LVSCW_AUTOSIZE_USEHEADER)

$init_months = StringSplit("Oct,Nov,Dec,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep", ",")
dim $months[13]
$months[0]=12
For $i = 1 To UBound($init_months) - 1
   $months[$i]=GUICtrlCreateListViewItem($init_months[$i],$listview)
Next

Global $msg1 = GUICtrlCreateLabel("Processing Status", 100, 160, 150, 25)
Global $msg2 = GUICtrlCreateLabel("OK count", 100, 200, 170, 25)
Global $msg3 = GUICtrlCreateLabel("Failed PUMA#", 100, 240, 150, 25)
Global $idListview = GUICtrlCreateListView("PUMA#", 100, 255, 150, 160)
GUICtrlSetColor(-1, 0xff0000)

$Label6 = GUICtrlCreateLabel("Folder", 20, 430, 60, 25)
$attachment = GUICtrlCreateInput(@WorkingDir & "\upload", 85, 425, 115, 25)
$idButton_attach = GUICtrlCreateButton("Upload", 200, 425, 45, 25)

$message = "Select File"
GUISetState(@SW_SHOW) ;Show GUI
_IENavigate($oIE, "https://opw.siemens.com/opw/Default.aspx")

While 1
    Local $iMsg = GUIGetMsg()
    Select
	  Case $iMsg = $Button1
			$Puma_file = FileOpenDialog($message, @WorkingDir , "Text (*.txt)", 1)
			GUICtrlSetData($Puma, $Puma_file)
	  Case $iMsg = $Button2
			$Step_file = FileOpenDialog($message, @WorkingDir , "Text (*.txt)", 1)
			GUICtrlSetData($Step, $Step_file)
      Case $iMsg = $GUI_EVENT_CLOSE
            ExitLoop
			Sleep(1000)
            CheckError("Home", @error, @extended)
      Case $iMsg = $idButton_Start
			Start(False)
            CheckError("Start", @error, @extended)
	  Case $iMsg = $idButton_upload_forecast
			Start(True)
            CheckError("Start", @error, @extended)
      Case $iMsg = $idButton_attach
			Upload_attachment()
            CheckError("Start", @error, @extended)
	  Case $iMsg = $idButton_Go
			_IENavigate($oIE,GUICtrlRead($HomePage))
            CheckError("Forward", @error, @extended)
    EndSelect
WEnd

Func Upload_attachment()
   _GUICtrlListView_DeleteAllItems($idListview)
   $Folder = GUICtrlRead($attachment)
   $ok_count = 0
   $FileList = _FileListToArray($Folder)
   Local $aRetArray
   Local $file_in_process[1]
   Local $upload_status
   Local $sFilePath ="file_in_process.txt"

   If UBound($FileList) -1>1 Then Run('Handle_File_Upload.exe','',@SW_MAXIMIZE)

   For $i = UBound($FileList) -1 To 1 Step -1
	  $FileName = $FileList[$i]
	  if StringInStr($FileName,".xls")>0 and StringInStr($FileName,"_")>0 Then
		 $puma= StringSplit($FileName,"_")[1]      ;"671954"
	  Else
		 $msg = "File " & $FileName & " ignored"
		 GUICtrlCreateListViewItem($msg, $idListview)
		 ContinueLoop
	  EndIf
	  $file = $Folder & "\" & $FileName
	  ;ConsoleWrite("File:" & $file & " to be processed"  & @CRLF)
	  ProcessOneStep("input","id",	"_adHeader_MainActionSearch_SearchTextBox","setvalue",$puma )
	  ProcessOneStep("button","id","_adHeader_MainActionSearch_ActionSearchButton","click")
	  For $ii = 1 to 500  ;wait till page refreshed
		 $new_action = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionIdValue')
		 If $new_action == $puma Then
			ExitLoop
		 Else
			Sleep(50)
		 EndIf
	  Next
	  $file_in_process[0] = $file
	  _FileWriteFromArray($sFilePath, $file_in_process)
	  $old_overview_status = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionLastupdateValue')
	  ;ProcessOneStep("td","id","ctl00_cphContentHeader_Action_ActionTabPage_AT5","click")
	  ProcessOneStep("td","id","ctl00_cphContentHeader_Action_ActionTabPage_Attachments_AttachmentRibbonMenu_RibbonMenu_DXI0_I","click")
	  ProcessOneStep("input","id","ctl00_cphContentHeader_Action_ActionTabPage_Attachments_FileUploadPopup_UploadFileControl","click")
	  Sleep(2000)
	  $upload_status = "Not OK"
	  For $j =1 to 1000
		 _FileReadToArray($sFilePath, $aRetArray)
		 If $aRetArray[1] == $file&"OK" Then
			$upload_status = "OK"
			ExitLoop
		 EndIf
		 Sleep(50)
	  Next

	  If $upload_status == "OK" Then
		 ProcessOneStep("input","id","ctl00_cphContentHeader_Action_ActionTabPage_Attachments_FileUploadPopup_UploadButton","click")
		 $msg = ""
		 For $k=1 to 150
		   $new_overview_status = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionLastupdateValue')
		   ;ConsoleWrite("old:" & $old_overview_status  & " new :" & $new_overview_status  & @CRLF)
		   If $old_overview_status <> $new_overview_status Then
			  Sleep(100)
			  $ok_count = $ok_count + 1
			  $msg = "Uploaded " & String($ok_count) & " files OK"
			  GUICtrlSetData($msg2, $msg)
			  ExitLoop
		   EndIf
		   Sleep(300)
		 Next
		 If $msg == "" Then GUICtrlCreateListViewItem($puma, $idListview)
	  Else
		 GUICtrlCreateListViewItem($puma, $idListview)
	  EndIf
   Next
   $file_in_process[0] = 'Finished'
   _FileWriteFromArray($sFilePath, $file_in_process)
   MsgBox(0,"Finished", "Upload finished!")
EndFunc


Func Start($upload_forecast)
   $fiscal_year = GUICtrlRead($year)

   $id_forecast_total = 'ctl00_cphContentHeader_Action_ActionTabPage_FactsAndFiguresGrid_cell0_17_RowSummaryLabel'
   $id_fiscal_year ='ctl00_cphContentHeader_Action_ActionTabPage_FiscalYearPicker_FiscalYearDropDownList'
   $id_allocation ='ctl00_cphContentHeader_Action_ActionTabPage_ActionAllocationControl_AllocationGrid_cell0_{index}_tb{year}'
   $sPreviousYear = String(Int($fiscal_year)-1)
   $sPreviousYearIndex = String(Int(StringRight($fiscal_year,2))-11)
   $id_allocation_previous_year = StringReplace($id_allocation, "{index}",$sPreviousYearIndex)
   $id_allocation_previous_year = StringReplace($id_allocation_previous_year, "{year}",$sPreviousYear)

   $sYearIndex = String(Int(StringRight($fiscal_year,2))-10)
   $id_allocation_year = StringReplace($id_allocation, "{index}",$sYearIndex)
   $id_allocation_year = StringReplace($id_allocation_year, "{year}",$fiscal_year)

   If $upload_forecast and $fiscal_year == "" Then
	  MsgBox(0,'Error', 'Please set fiscal year before upload forecast!')
	  Return
   EndIf
   $selected_month = ""
   $iStart_month=0
   $iEnd_month=0
   For $i=1 to $months[0]
	  If BitOR(GUICtrlRead($months[$i], 1), $GUI_CHECKED) = $GUI_CHECKED Then
		 If $iStart_month==0 Then $iStart_month=$i
		 $iEnd_month= $i
		 $aItem = _GUICtrlListView_GetItem($listview, $i-1)
		 $selected_month = $selected_month & $aItem[3]
	  EndIf
   Next
   If $selected_month == "" Then
	  MsgBox(0,'Error', 'Please select at least one month for uploading!')
	  Return
   EndIf
   $file_puma = GUICtrlRead($Puma)
   FileOpen($file_puma, 0)
   $file_step = GUICtrlRead($Step)
   FileOpen($file_step, 0)
   $home_page = GUICtrlRead($HomePage)
   $header_line = FileReadLine($file_puma, 1)
   $header_column = StringSplit($header_line, @TAB)
   $total_number = _FileCountLines($file_puma)
   $ok_count = 0
   For $i = 2 to $total_number
	  $msg = "Processing " & String($i-1) & "  Of  " & String($total_number-1)
	  GUICtrlSetData($msg1, $msg)
	  $line_puma = FileReadLine($file_puma, $i)
	  $columns_puma = StringSplit($line_puma, @TAB)
	  For $j = 2 to _FileCountLines($file_step)
		 $line_step = FileReadLine($file_step, $j)
		 $columns_step = StringSplit($line_step, @TAB)
		 $failed = False
		 $input_value = ''
		 $attr_value = $columns_step[5]
		 If $j == 2 Then
			$input_value = $columns_puma[1]
			If StringLen($input_value) == 0 Then Return  ;puma number empty end of processing
		 ElseIf $j == 5 Then

			If $upload_forecast Then ContinueLoop    ;only upload forecast

			$attr_value = StringReplace($attr_value, "{month}", $header_column[$iStart_month+1])
			$attr_value = StringReplace($attr_value, "{index}", String($iStart_month+4))
			$input_value = $columns_puma[$iStart_month+1]
		 ElseIf $j == 6 Then  ;loop to input selected months data

			If $upload_forecast Then $iStart_month = $iStart_month - 1   ;upload all months forecast

			For $k=$iStart_month+2 to $iEnd_month+1

			   If StringInStr($selected_month, $header_column[$k]) Then
				  $input_value = $columns_puma[$k]
				  $attr_value = StringReplace($columns_step[5], "{month}", $header_column[$k])
				  $attr_value = StringReplace($attr_value, "{index}", String($k+3))	;month index
				  ;ConsoleWrite("month:" & $header_column[$k] & "||" &  @TAB & $input_value & @TAB & $attr_value & @TAB & $columns_puma[1] & @CRLF)
				  If Not ProcessOneStep($columns_step[3],$columns_step[4], $attr_value, $columns_step[6],$input_value) Then
					 ConsoleWrite('Step failed!' & @TAB)
					 $failed = True
					 ExitLoop
				  EndIf
			   EndIf
			Next
		 EndIf

		 If $j <> 6 Then
			If Not ProcessOneStep($columns_step[3],$columns_step[4], $attr_value, $columns_step[6],$input_value) Then
			   If $i == 2 Then
				  MsgBox(0,'Error', 'Please check the template file column header and make sure already login PUMA website before uploading!')
				  Return
			   EndIf
			   ConsoleWrite('Step failed!'& @TAB)
			   $failed = True
			EndIf
     	    ;ConsoleWrite($columns_step[2] & @TAB & $input_value & @TAB & $attr_value & @TAB & $line_puma & @CRLF)
		 EndIf
		 if $failed == True Then
			ExitLoop
		 EndIf
		 If $j == 3 Then ;after click search button wait till the new action ID page refreshed per search and ready for input value
			For $ii = 1 to 500
			   $new_action = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionIdValue')
			   If $new_action == $columns_puma[1] Then
				  ExitLoop
			   Else
				  Sleep(50)
			   EndIf
			Next
			;switch to the fiscal year for upload forecast only
			$old_forecast_total = Get_InnerText_By_ID($id_forecast_total)
			;ConsoleWrite("forecast total old:" & $old_forecast_total & @CRLF)
			Local $oFactsTab = _IEGetObjById($oIE,'ctl00_cphContentHeader_Action_ActionTabPage_AT2')
			_IEAction($oFactsTab, "click")
			Local $oFiscalYear = _IEGetObjById($oIE,$id_fiscal_year)
            _IEFormElementOptionSelect($oFiscalYear, $fiscal_year, 1, "byText")
			;copy previous year allocation to new forecast year
			$allocation_previous_year = _IEGetObjById($oIE, $id_allocation_previous_year)
			$allocation_previous_year = _IEFormElementGetValue($allocation_previous_year)

			$allocation_year = _IEGetObjById($oIE, $id_allocation_year)
			_IEFormElementSetValue($allocation_year, $allocation_previous_year)

			;ConsoleWrite($id_allocation_previous_year &":"  & $id_allocation_year & ":" & $allocation_previous_year & @CRLF)
			For $ii = 1 to 200
			   $new_forecast_total = Get_InnerText_By_ID($id_forecast_total)
			   If $new_forecast_total <> $old_forecast_total Then
				  ExitLoop
			   Else
				  Sleep(50)
			   EndIf
			Next
		 EndIf
	  Next
	  $old_overview_status = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionLastupdateValue')
	  Sleep(1000)
	  $msg = ""
	  For $k=1 to 150
	    $new_overview_status = Get_InnerText_By_ID('cphContentHeader_Action_InformationActionLastupdateValue')
		;ConsoleWrite("old:" & $old_overview_status  & "new :" & $new_overview_status  & @CRLF)
		If $old_overview_status <> $new_overview_status Then
		   Sleep(100)
		   $ok_count = $ok_count + 1
		   $msg = "Changed " & String($ok_count) & " PUMA OK"
		   GUICtrlSetData($msg2, $msg)
		   ExitLoop
		EndIf
		Sleep(300)
	  Next
	  If $msg == "" Then GUICtrlCreateListViewItem($columns_puma[1], $idListview)
   Next
   FileClose($file_puma)
   FileClose($file_step)
EndFunc

Func ProcessOneStep($tag,$attr,$attr_value,$action,$input_value=0)
   For $i= 1 To 200
	  If $attr == 'id' Then
		 $target = _IEGetObjById ($oIE, $attr_value)
	  ElseIf $attr == 'name' Then
		 $target = _IEGetObjByName ($oIE, $attr_value)
	  ElseIf $attr == 'innertext' Then
		 $target = GetObjByInnerText ($oIE, $tag,$attr_value)
	  Else
		 $target = GetObjByAttr($tag,$attr,$attr_value)
	  EndIf

	  If IsObj($target) Then
		 If $action == "setvalue" Then
			_IEFormElementSetValue($target,$input_value)
			_IEAction($target, "focus")
		 ElseIf $action == "click" Then
			_IEAction($target, $action)
		 EndIf
		 Return True
	  EndIf
	  Sleep(50)
   Next
   Return False
EndFunc

Func Get_InnerText_By_ID($ID)
   For $i = 1 To 200
	  $target = _IEGetObjById($oIE,$ID)
	  If IsObj($target) Then
		 Return $target.innertext
	  Else
		 Sleep(50)
	  EndIf
   Next
EndFunc

Func GetObjByInnerText($oIE,$tag, $text)
	$tags = $oIE.document.GetElementsByTagName($tag)
	For $tag in $tags
	  If StringInStr($tag.innertext,$text) Then	return $tag
   Next
   Return 0
EndFunc

Func GetObjByAttr($tag_name, $attr, $attr_value)
	$tags = $oIE.document.GetElementsByTagName($tag_name)
	For $tag in $tags
	  If $attr == 'class' Then
		 If $tag.className == $attr_value Then  return $tag
	  Else
		 If $tag.GetAttribute($attr) == $attr_value Then return $tag
	  EndIf
   Next
   Return 0
EndFunc

Func CheckError($sMsg, $iError, $iExtended)
    If $iError Then
        $sMsg = "Error using " & $sMsg & " button (" & $iExtended & ")"
    Else
        $sMsg = ""
    EndIf
    GUICtrlSetData($msg1, $sMsg)
EndFunc   ;==>CheckError

GUIDelete()
Exit
