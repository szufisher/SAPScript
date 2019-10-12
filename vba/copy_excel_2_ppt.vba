Sub Daten_ausExcel_holen()
    Dim wb As Workbook, wks As Worksheet
    Dim Folie As Slide, Textfeld As Shape
    Set Folie = ActivePresentation.Slides(1)
    Set Textfeld = Folie.Shapes("dpath")
        
    Set wb = Workbooks.Open(FileName:=Textfeld.TextFrame.TextRange.Text, ReadOnly:=True)
    Set wks = wb.Worksheets("Report")
    Dim FNummer As Integer
    Dim FNummer2 As Integer
    FNummer2 = ActivePresentation.Slides(1).Shapes("Counter").TextFrame.TextRange.Text + 1
    
    For FNummer = 2 To FNummer2
        Set Folie = ActivePresentation.Slides(FNummer)
        Set Textfeld = Folie.Shapes("Title")
        Textfeld.TextFrame.TextRange.Text = "Prio " & wks.Range("AF" & 8 + FNummer).Text & " - " & wks.Range("F" & 8 + FNummer).Text
        Set Textfeld = Folie.Shapes("Subtitle")
        Textfeld.TextFrame.TextRange.Text = "ARE: " & wks.Range("B" & 8 + FNummer).Text & "   " & " Zone: " & wks.Range("E" & 8 + FNummer).Text & "   " & " Owner: " & wks.Range("N" & 8 + FNummer).Text
        
        Set Textfeld = Folie.Shapes("DeficiencyDescription")
        Textfeld.TextFrame.TextRange.Text = wks.Range("AB" & 8 + FNummer).Text
            
        Set Textfeld = Folie.Shapes("PrioText")
        Textfeld.TextFrame.TextRange.Text = wks.Range("AG" & 8 + FNummer).Text
        Set Textfeld = Folie.Shapes("AssessmentType")
        If wks.Range("M" & 8 + FNummer).Text = "DA" Then
            Textfeld.TextFrame.TextRange.Text = "CR " & wks.Range("G" & 8 + FNummer).Text & " (Detailed Assessment)"
        End If
        
        If wks.Range("M" & 8 + FNummer).Text = "DA-ICFR" Then
            Textfeld.TextFrame.TextRange.Text = "CR " & wks.Range("G" & 8 + FNummer).Text & " (Detailed Assessment - ICFR)"
        End If
        
        If wks.Range("M" & 8 + FNummer).Text = "SA" Then
            Textfeld.TextFrame.TextRange.Text = "CR " & wks.Range("G" & 8 + FNummer).Text & " (Self Assessment)"
        End If
        
        If wks.Range("M" & 8 + FNummer).Text = "NSAR" Then
            Textfeld.TextFrame.TextRange.Text = "CR " & wks.Range("G" & 8 + FNummer).Text & " (Non specific assessment required)"
        End If
        
        Set Textfeld = Folie.Shapes("AssessmentText")
        Textfeld.TextFrame.TextRange.Text = wks.Range("H" & 8 + FNummer).Text
        
        Set Textfeld = Folie.Shapes("RemediationText")
        Textfeld.TextFrame.TextRange.Text = wks.Range("BF" & 8 + FNummer).Text
        
        Set Textfeld = Folie.Shapes("DNummer")
        Textfeld.TextFrame.TextRange.Text = wks.Range("AA" & 8 + FNummer).Text
        
        Set Textfeld = Folie.Shapes("DType")
        Textfeld.TextFrame.TextRange.Text = wks.Range("AD" & 8 + FNummer).Text & " - " & wks.Range("BH" & 8 + FNummer).Text
        Set Textfeld = Folie.Shapes("Status")
        Textfeld.TextFrame.TextRange.Text = "Status: " & Format(Date, "dd.mm.yyyy")
    Next FNummer
    
    wb.Close savechanges:=False
End Sub
