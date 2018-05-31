SAP Script tips

1. grid handling, concatenate with variable 
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & screen_no & "/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT12/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/txtMEACCT1000-VPROZ[3," & CStr(row) & "]").Text = percentage '"25"
2. vertical scroll bar
  session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & screen_no & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position = Position + 1
  Position = session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:" & screen_no & "/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").verticalScrollbar.Position
3. how to determine whether need to switch page
	on error resume next
   for grid/table
   PageSize = session.findById("wnd[0]/usr").findByName("SAPLMIGOTV_GOITEM", "GuiTableControl").verticalScrollbar.PageSize
    pageindex = j Mod PageSize
    If pageindex = 1 Then
        If j > 1 Then
               session.findById("wnd[0]/tbar[0]/btn[82]").press   ' click the next page button
              End If
                CurrentRow = 0
            Else
                CurrentRow = CurrentRow + 1
            End If
    for more details, refer to MIGO 261 order _new
4. screen number change due to fold/ unfold status, or screen size, resolution
  Function detect_screen_no(screen_no As String, str1 As String, str2 As String) As String
    On Error Resume Next
        session.findById (str1 & screen_no & str2)
    If Err.Number = 0 Then
        detect_screen_no = screen_no
    End If
    For i = 20 To 10 Step -1
        On Error Resume Next
           session.findById (str1 & "00" & CStr(i) & str2)
        If Err.Number = 0 Then
            detect_screen_no = "00" & CStr(i)
            Exit For
        End If
    Next i
    'detect_screen_no = ""
  End Function
  
  Except findByID, instead findByName(name, controltype) can be used to avoid the dynamic screen number issue, e.g  
  session.findById("wnd[0]/usr").findByName("GODYNPRO-DETAIL_ZEILE", "GuiTextField").Text
  for more details refer to MIGO 261 order New
5. on error resume next to handle exception
        On Error Resume Next
           session.findById (str1 & "00" & CStr(i) & str2)
        If Err.Number = 0 Then
            detect_screen_no = "00" & CStr(i)
            Exit For
        End If
6. status bar message type warning, press ENTER
	MessageType = session.findById("wnd[0]/sbar").MessageType
    If MessageType = "W" Then
        session.findById("wnd[0]").sendVKey 0
    End If
7. datetime , quantity format
8. carriage return embeded in longtext 
   mat_desc = Replace(Replace(mat_desc, Chr(13), ""), Chr(10), "")

9.  considering performance, access the internal object( row, cell) of the item list instead of switch to item details, 
the address/path to the table's active cell: tablecontrol.Rows(rowindex)(columnindex), eg. table_items.Rows(0)(4), means the first row, the 5th column
    assign variable for tablecontrol, row or column will trigger sap auto exit when switch to other transaction or popup window!!!
    
10.  useful video https://www.youtube.com/watch?v=oPPhA14Pm-8 explain SAPGUI script from Excel
