#include <SAP.au3>
#include <AutoItConstants.au3>
#include <Constants.au3>
$FileName = @ScriptDir & "\AutoDownloadingReport.xlsm"

If Not FileExists($FileName) Then ; Just a check to be sure..
	MsgBox($MB_SYSTEMMODAL, "Excel Data Test", "Error: Can't find file " & $FileName)
	Exit
EndIf

Local $oExcelDoc = ObjGet($FileName) ; Get an Excel Object from an existing filename

If (Not @error) And IsObj($oExcelDoc) Then ; Check again If everything went well
	; NOTE: $oExcelDoc is a "Workbook Object", Not Excel itself!
   Local $number_rows = $oExcelDoc.ActiveSheet.UsedRange.Rows.Count
   Local $CellRange = "B2:J" & $number_rows
   Local $oDocument = $oExcelDoc.Worksheets(1) ; We use the 'Default' worksheet
   Local $cells = $oDocument.range($CellRange).value
   $oDocument = 0
   $oExcelDoc.saved = 1 ; Prevent questions from excel to save the file
   $oExcelDoc.close ; Get rid of Excel.
   $oExcelDoc = 0

   If IsArray($cells) And UBound($cells, 0) > 0 Then
	   For $x = 0 To UBound($cells, 1) - 1
		  $tcode = $cells[0][$x]  ; [col][row]
		  $ledger = $cells[1][$x]
		  $report = $cells[2][$x]
		  $bu = $cells[3][$x]
		  $cost_object = $cells[4][$x]
		  $fiscal_year = $cells[5][$x]
		  $from_period = $cells[6][$x]
		  $end_period = $cells[7][$x]
		  $folder = $cells[8][$x]
		  $save_file = $folder & $report& "-"& $bu&"-" & $fiscal_year & $end_period
		  If $tcode="" Then
			 ExitLoop
		  EndIf
		  _SAPSessAttach("",$tcode)
		 ;_SAPSessCreate()
		 _SAPObjValueSet("usr/ctxt$HY-BKRS","CN10")
		 _SAPObjValueSet("usr/ctxt$HY-LEDG", $ledger)
		 _SAPObjValueSet("usr/txt$HY-GJAH", $fiscal_year)
		 If $report= "BS" Or $report = "CFS" Then
			_SAPObjValueSet("usr/txt$HY-PERI", $end_period)
		 Else
			_SAPObjValueSet("usr/txt$HY-FR", $from_period)
			_SAPObjValueSet("usr/txt$HY-BPE1",$end_period)
		 EndIf

		 _SAPObjValueSet("usr/ctxt$HY-COBJ",$cost_object)
		 _SAPObjValueSet("usr/txt$HY-PERI", $end_period)
		 _SAPVKeysSend("F8")

		 MouseClick($MOUSE_CLICK_LEFT, 100, 185, 2)
		 Sleep(18000)
		 MouseClick($MOUSE_CLICK_LEFT,294,154)
		 MouseClick($MOUSE_CLICK_LEFT,294,199)
		 Sleep(2000)

		 Local $hSaveAsDlg=WinWaitActive("Save As","",5)
		 If $hSaveAsDlg <>0 then
			ControlSend($hSaveAsDlg,"","[CLASS:Edit; INSTANCE:1]","!n")  ;alt+n place cursor to filename edit box
			ControlSetText($hSaveAsDlg,"","[CLASS:Edit; INSTANCE:1]", $save_file)			
			Sleep(2000)  ; wait to swith the target folder, otherwise excel open error
			ControlClick($hSaveAsDlg,"","&Save")
			Sleep(2000)
		 Else
			MsgBox($MB_SYSTEMMODAL, "Save As Dialog", "Error:Failed activate Save As Popup window" )
		 EndIf
	  Next
   EndIf

Else
	MsgBox($MB_SYSTEMMODAL, "Excel Data Test", "Error: Could Not open " & $FileName & " as an Excel Object.")
EndIf
