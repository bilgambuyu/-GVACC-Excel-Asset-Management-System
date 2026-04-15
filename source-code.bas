Sub Generate_Vendor_Asset_Management_System()
    ' ==================================================================
    ' GVACC - Global Vendor & Asset Command Center
    ' FINAL UNIVERSAL VERSION - All Excel Versions
    ' ==================================================================
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ' ----------------------------------------------------------------
    ' SAFELY DELETE EXISTING SHEETS
    ' ----------------------------------------------------------------
    Dim ws As Worksheet
    Dim sheetCount As Integer
    sheetCount = wb.Sheets.Count
    
    If sheetCount = 1 Then
        wb.Sheets.Add After:=wb.Sheets(wb.Sheets.Count)
    End If
    
    Application.DisplayAlerts = False
    For Each ws In wb.Sheets
        If wb.Sheets.Count > 1 Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    wb.Sheets(1).Name = "Temp_Placeholder"
    
    ' ----------------------------------------------------------------
    ' SHEET 1: Assets_Master
    ' ----------------------------------------------------------------
    Dim wsAssets As Worksheet
    Set wsAssets = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsAssets.Name = "Assets_Master"
    
    wsAssets.Range("A1:K1").Value = Array("Asset_Tag", "Device_Type", "Model", "Serial_Number", _
        "Vendor_Name", "Contract_ID", "Start_Date", "End_Date", "Annual_Cost", "Region", "SLA_Status")
    
    With wsAssets.Range("A1:K1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    With wsAssets.Range("B2:B550").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="Laptop,Server,Network Gear,Mobile,Printer,Desktop"
        .IgnoreBlank = True
    End With
    
    With wsAssets.Range("E2:E550").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="Dell Global,HP Inc,Apple Enterprise,Cisco Systems,Lenovo Pro,Fujitsu,Samsung IT"
        .IgnoreBlank = True
    End With
    
    With wsAssets.Range("J2:J550").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:="AMER,EMEA,APAC,LATAM"
        .IgnoreBlank = True
    End With
    
    Dim i As Long
    Dim vendors As Variant, regions As Variant, devices As Variant, models As Variant
    vendors = Array("Dell Global", "HP Inc", "Apple Enterprise", "Cisco Systems", "Lenovo Pro", "Fujitsu", "Samsung IT")
    regions = Array("AMER", "EMEA", "APAC", "LATAM")
    devices = Array("Laptop", "Server", "Network Gear", "Mobile", "Printer", "Desktop")
    models = Array("Latitude 5430", "EliteBook 840", "MacBook Pro 16", "Catalyst 9300", "ThinkPad X1", "iPhone 14 Pro", "ProLiant DL380")
    
    Randomize
    
    For i = 1 To 520
        Dim vendorIdx As Integer: vendorIdx = Int((UBound(vendors) + 1) * Rnd)
        Dim regionIdx As Integer: regionIdx = Int((UBound(regions) + 1) * Rnd)
        Dim deviceIdx As Integer: deviceIdx = Int((UBound(devices) + 1) * Rnd)
        Dim modelIdx As Integer: modelIdx = Int((UBound(models) + 1) * Rnd)
        
        wsAssets.Cells(i + 1, 1).Value = Left(devices(deviceIdx), 4) & "-" & Left(regions(regionIdx), 3) & "-" & Format(i, "0000")
        wsAssets.Cells(i + 1, 2).Value = devices(deviceIdx)
        wsAssets.Cells(i + 1, 3).Value = models(modelIdx) & " Gen" & Int((5 * Rnd) + 8)
        wsAssets.Cells(i + 1, 4).Value = "SN" & Format(Int((999999 * Rnd) + 100000), "000000")
        wsAssets.Cells(i + 1, 5).Value = vendors(vendorIdx)
        wsAssets.Cells(i + 1, 6).Value = "CT-" & Left(vendors(vendorIdx), 4) & "-" & Year(Date) & "-" & Format(i, "000")
        wsAssets.Cells(i + 1, 7).Value = DateAdd("d", -Int((1095 * Rnd) + 30), Date)
        wsAssets.Cells(i + 1, 8).Value = DateAdd("yyyy", Int((3 * Rnd) + 1), wsAssets.Cells(i + 1, 7).Value)
        
        Dim cost As Double
        Select Case devices(deviceIdx)
            Case "Laptop", "Desktop": cost = 800 + (Rnd * 1200)
            Case "Server": cost = 3500 + (Rnd * 8000)
            Case "Network Gear": cost = 2000 + (Rnd * 5000)
            Case "Mobile": cost = 600 + (Rnd * 600)
            Case "Printer": cost = 400 + (Rnd * 800)
        End Select
        wsAssets.Cells(i + 1, 9).Value = Round(cost, 2)
        wsAssets.Cells(i + 1, 10).Value = regions(regionIdx)
        wsAssets.Cells(i + 1, 11).Formula = "=IF(TODAY()>H" & i + 1 & ",""Expired"",""Active"")"
    Next i
    
    wsAssets.Columns("A:K").AutoFit
    wsAssets.Columns("G:H").NumberFormat = "mm/dd/yyyy"
    wsAssets.Columns("I").NumberFormat = "$#,##0.00"
    
    wsAssets.ListObjects.Add(xlSrcRange, wsAssets.Range("A1:K521"), , xlYes).Name = "Assets_Master"
    
    ' ----------------------------------------------------------------
    ' SHEET 2: Repair_Log
    ' ----------------------------------------------------------------
    Dim wsRepair As Worksheet
    Set wsRepair = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsRepair.Name = "Repair_Log"
    
    wsRepair.Range("A1:J1").Value = Array("Ticket_ID", "Asset_Tag", "Date_Reported", "Date_Resolved", _
        "Vendor", "Region", "SLA_Days_Allowed", "Actual_TAT", "SLA_Breach", "Month")
    
    With wsRepair.Range("A1:J1")
        .Font.Bold = True
        .Interior.Color = RGB(192, 80, 77)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    Dim j As Long, rndAsset As Long
    Dim reportedDate As Date, resolvedDate As Date
    
    For j = 1 To 380
        wsRepair.Cells(j + 1, 1).Value = "TKT-" & Year(Date) & "-" & Format(j, "0000")
        rndAsset = Int((520 * Rnd) + 2)
        wsRepair.Cells(j + 1, 2).Value = wsAssets.Cells(rndAsset, 1).Value
        
        If j < 150 Then
            reportedDate = DateAdd("d", -Int((90 * Rnd) + 60), Date)
        Else
            reportedDate = DateAdd("d", -Int((30 * Rnd) + 5), Date)
        End If
        wsRepair.Cells(j + 1, 3).Value = reportedDate
        
        If j < 150 Then
            resolvedDate = DateAdd("d", Int((4 * Rnd) + 2), reportedDate)
        Else
            resolvedDate = DateAdd("d", Int((2 * Rnd) + 1), reportedDate)
        End If
        wsRepair.Cells(j + 1, 4).Value = resolvedDate
        
        wsRepair.Cells(j + 1, 5).Formula = "=IFERROR(VLOOKUP(B" & j + 1 & ",Assets_Master!A:E,5,FALSE),""Unknown"")"
        wsRepair.Cells(j + 1, 6).Formula = "=IFERROR(VLOOKUP(B" & j + 1 & ",Assets_Master!A:J,10,FALSE),""Unknown"")"
        wsRepair.Cells(j + 1, 7).Value = IIf(InStr(1, wsRepair.Cells(j + 1, 2).Value, "SRV") > 0, 3, 2)
        wsRepair.Cells(j + 1, 8).Formula = "=IF(AND(C" & j + 1 & "<>"""",D" & j + 1 & "<>""""),NETWORKDAYS(C" & j + 1 & ",D" & j + 1 & "),0)"
        wsRepair.Cells(j + 1, 9).Formula = "=IF(H" & j + 1 & ">G" & j + 1 & ",""YES"",""NO"")"
        wsRepair.Cells(j + 1, 10).Formula = "=TEXT(C" & j + 1 & ",""mmm-yyyy"")"
    Next j
    
    wsRepair.ListObjects.Add(xlSrcRange, wsRepair.Range("A1:J381"), , xlYes).Name = "Repair_Log"
    wsRepair.Columns("A:J").AutoFit
    wsRepair.Columns("C:D").NumberFormat = "mm/dd/yyyy"
    
    ' ----------------------------------------------------------------
    ' SHEET 3: Control_Metrics
    ' ----------------------------------------------------------------
    Dim wsControl As Worksheet
    Set wsControl = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsControl.Name = "Control_Metrics"
    
    wsControl.Range("A1:D1").Value = Array("Period", "Total_Tickets", "Late_Tickets", "Late_Percentage")
    With wsControl.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    wsControl.Range("A2").Value = "Historical Baseline"
    wsControl.Range("A3").Value = "Current Period"
    wsControl.Range("A5").Value = "Reduction Achieved:"
    
    wsControl.Range("B2").Formula = "=COUNTIF(Repair_Log[Ticket_ID], ""TKT*"") - COUNTIF(Repair_Log[Ticket_ID], "">TKT-" & Year(Date) & "-0150"")"
    wsControl.Range("C2").Formula = "=COUNTIFS(Repair_Log[SLA_Breach], ""YES"", Repair_Log[Ticket_ID], ""<TKT-" & Year(Date) & "-0150"")"
    wsControl.Range("D2").Formula = "=IF(B2>0, C2/B2, 0)"
    
    wsControl.Range("B3").Formula = "=COUNTIF(Repair_Log[Ticket_ID], "">TKT-" & Year(Date) & "-0150"")"
    wsControl.Range("C3").Formula = "=COUNTIFS(Repair_Log[SLA_Breach], ""YES"", Repair_Log[Ticket_ID], "">TKT-" & Year(Date) & "-0150"")"
    wsControl.Range("D3").Formula = "=IF(B3>0, C3/B3, 0)"
    
    wsControl.Range("B5").Formula = "=IFERROR((D2-D3)/D2, 0)"
    wsControl.Range("B5").NumberFormat = "0.0%"
    
    wsControl.Range("D2:D3").NumberFormat = "0.0%"
    wsControl.Range("B2:C3").NumberFormat = "0"
    wsControl.Columns("A:D").AutoFit
    
    ' ----------------------------------------------------------------
    ' SHEET 4: Pivot_SLA
    ' ----------------------------------------------------------------
    Dim wsPivotSLA As Worksheet
    Set wsPivotSLA = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsPivotSLA.Name = "Pivot_SLA"
    
    Dim pcSLA As PivotCache
    Dim ptSLA As PivotTable
    
    Set pcSLA = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsRepair.ListObjects("Repair_Log").Range)
    Set ptSLA = pcSLA.CreatePivotTable(TableDestination:=wsPivotSLA.Range("A3"), TableName:="SLA_Performance")
    
    With ptSLA
        .PivotFields("Vendor").Orientation = xlRowField
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Ticket_ID").Orientation = xlDataField
        .DataFields(1).Function = xlCount
        .DataFields(1).Name = "Total Repairs"
        
        .PivotFields("SLA_Breach").Orientation = xlDataField
        .DataFields(2).Function = xlCount
        .DataFields(2).Name = "Breaches"
    End With
    
    wsPivotSLA.Range("A1").Value = "Vendor SLA Performance Dashboard"
    wsPivotSLA.Range("A1").Font.Size = 16
    wsPivotSLA.Range("A1").Font.Bold = True
    
    ' ----------------------------------------------------------------
    ' SHEET 5: Pivot_Spend
    ' ----------------------------------------------------------------
    Dim wsPivotSpend As Worksheet
    Set wsPivotSpend = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsPivotSpend.Name = "Pivot_Spend"
    
    Dim pcSpend As PivotCache
    Dim ptSpend As PivotTable
    
    Set pcSpend = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=wsAssets.ListObjects("Assets_Master").Range)
    Set ptSpend = pcSpend.CreatePivotTable(TableDestination:=wsPivotSpend.Range("A3"), TableName:="Regional_Spend")
    
    With ptSpend
        .PivotFields("Region").Orientation = xlRowField
        .PivotFields("Vendor_Name").Orientation = xlRowField
        .PivotFields("Annual_Cost").Orientation = xlDataField
        .DataFields(1).Function = xlSum
        .DataFields(1).NumberFormat = "$#,##0"
        .DataFields(1).Name = "Total Annual Spend"
        
        .PivotFields("Asset_Tag").Orientation = xlDataField
        .DataFields(2).Function = xlCount
        .DataFields(2).Name = "Device Count"
    End With
    
    wsPivotSpend.Range("A1").Value = "Regional Procurement Spend Analysis"
    wsPivotSpend.Range("A1").Font.Size = 16
    wsPivotSpend.Range("A1").Font.Bold = True
    
    ' ----------------------------------------------------------------
    ' SHEET 6: Executive Dashboard
    ' ----------------------------------------------------------------
    Dim wsDash As Worksheet
    Set wsDash = wb.Sheets.Add(Before:=wb.Sheets("Temp_Placeholder"))
    wsDash.Name = "Executive_Dashboard"
    
    ' Title Banner
    With wsDash.Range("A1:H1")
        .Merge
        .Value = "GLOBAL VENDOR & ASSET COMMAND CENTER"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With
    
    ' KPI Cards Section
    wsDash.Range("A3").Value = "KEY PERFORMANCE INDICATORS"
    wsDash.Range("A3:F3").Merge
    wsDash.Range("A3").Font.Bold = True
    wsDash.Range("A3").Font.Size = 12
    wsDash.Range("A3").Interior.Color = RGB(220, 230, 241)
    
    ' KPI 1: Current SLA Breach Rate
    With wsDash.Range("A5:C9")
        .Merge
        .Value = "Current SLA Breach Rate"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(198, 224, 180)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("A7").Value = "=Control_Metrics!D3"
    wsDash.Range("A7").Font.Size = 24
    wsDash.Range("A7").Font.Bold = True
    wsDash.Range("A7").NumberFormat = "0.0%"
    wsDash.Range("A7").HorizontalAlignment = xlCenter
    wsDash.Range("A7").VerticalAlignment = xlCenter
    
    ' KPI 2: Reduction Achieved
    With wsDash.Range("D5:F9")
        .Merge
        .Value = "SLA Improvement"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(155, 194, 230)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("D7").Value = "=Control_Metrics!B5"
    wsDash.Range("D7").Font.Size = 24
    wsDash.Range("D7").Font.Bold = True
    wsDash.Range("D7").NumberFormat = "0.0%"
    wsDash.Range("D7").HorizontalAlignment = xlCenter
    wsDash.Range("D7").VerticalAlignment = xlCenter
    wsDash.Range("D8").Value = "Target: 20%"
    wsDash.Range("D8").Font.Size = 10
    wsDash.Range("D8").Font.Color = RGB(192, 0, 0)
    wsDash.Range("D8").HorizontalAlignment = xlCenter
    
    ' KPI 3: Active Contracts
    With wsDash.Range("G5:I9")
        .Merge
        .Value = "Active Contracts"
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .Interior.Color = RGB(255, 217, 102)
        .Borders.LineStyle = xlContinuous
    End With
    wsDash.Range("G7").Formula = "=COUNTIF(Assets_Master[SLA_Status], ""Active"")"
    wsDash.Range("G7").Font.Size = 24
    wsDash.Range("G7").Font.Bold = True
    wsDash.Range("G7").HorizontalAlignment = xlCenter
    wsDash.Range("G7").VerticalAlignment = xlCenter
    
    ' ----------------------------------------------------------------
    ' CHART 1: SLA Breach Trend (Line Chart)
    ' ----------------------------------------------------------------
    Dim chtSLA As ChartObject
    Set chtSLA = wsDash.ChartObjects.Add(Left:=10, Top:=190, Width:=500, Height:=260)
    With chtSLA.Chart
        .SetSourceData Source:=wsPivotSLA.Range("A4:C10")
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Monthly SLA Breach Rate Trend"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' ----------------------------------------------------------------
    ' CHART 2: Regional Spend (PIE CHART - Safe Version)
    ' ----------------------------------------------------------------
    ' Create summary table on dashboard
    wsDash.Range("M1").Value = "Region"
    wsDash.Range("N1").Value = "Spend"
    
    wsDash.Range("M2").Value = "AMER"
    wsDash.Range("M3").Value = "EMEA"
    wsDash.Range("M4").Value = "APAC"
    wsDash.Range("M5").Value = "LATAM"
    
    wsDash.Range("N2").Formula = "=SUMIF(Assets_Master[Region], M2, Assets_Master[Annual_Cost])"
    wsDash.Range("N3").Formula = "=SUMIF(Assets_Master[Region], M3, Assets_Master[Annual_Cost])"
    wsDash.Range("N4").Formula = "=SUMIF(Assets_Master[Region], M4, Assets_Master[Annual_Cost])"
    wsDash.Range("N5").Formula = "=SUMIF(Assets_Master[Region], M5, Assets_Master[Annual_Cost])"
    
    Dim chtSpend As ChartObject
    Set chtSpend = wsDash.ChartObjects.Add(Left:=520, Top:=190, Width:=480, Height:=260)
    With chtSpend.Chart
        .SetSourceData Source:=wsDash.Range("M1:N5")
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Regional Spend Distribution"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        
        ' Safe way to add data labels - works in all Excel versions
        If Val(Application.Version) >= 12 Then
            .SeriesCollection(1).HasDataLabels = True
            .SeriesCollection(1).DataLabels.ShowPercentage = True
            .SeriesCollection(1).DataLabels.ShowValue = False
        End If
    End With
    
    ' Hide the summary table columns
    wsDash.Columns("M:N").Hidden = True
    
    ' ----------------------------------------------------------------
    ' CHART 3: Vendor Performance (Bar Chart)
    ' ----------------------------------------------------------------
    Dim chtVendor As ChartObject
    Set chtVendor = wsDash.ChartObjects.Add(Left:=10, Top:=470, Width:=500, Height:=220)
    With chtVendor.Chart
        .SetSourceData Source:=wsPivotSLA.Range("A5:B12")
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Vendor Performance (Breach Count)"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
    End With
    
    ' ----------------------------------------------------------------
    ' SLICER FOR INTERACTIVITY
    ' ----------------------------------------------------------------
    On Error Resume Next
    Dim scRegion As SlicerCache
    Set scRegion = wb.SlicerCaches.Add(wsPivotSLA.PivotTables("SLA_Performance"), "Region")
    scRegion.Slicers.Add wsDash, , "Region", "Region Filter", 520, 470, 150, 200
    scRegion.PivotTables.AddPivotTable wsPivotSpend.PivotTables("Regional_Spend")
    On Error GoTo 0
    
    ' ----------------------------------------------------------------
    ' DELETE TEMP PLACEHOLDER
    ' ----------------------------------------------------------------
    Application.DisplayAlerts = False
    wb.Sheets("Temp_Placeholder").Delete
    Application.DisplayAlerts = True
    
    ' ----------------------------------------------------------------
    ' FINAL CLEANUP
    ' ----------------------------------------------------------------
    wsDash.Activate
    wsDash.Range("A1").Select
    
    Application.ScreenUpdating = True
    
    MsgBox "GVACC Excel System Generated Successfully!" & vbCrLf & vbCrLf & _
           "✓ 520 Hardware Assets Created" & vbCrLf & _
           "✓ 380 Repair Tickets with SLA Tracking" & vbCrLf & _
           "✓ 2 Pivot Tables (SLA & Spend)" & vbCrLf & _
           "✓ Executive Dashboard with Charts & Slicer" & vbCrLf & vbCrLf & _
           "Save this file as: GVACC_Executive_Dashboard_v1.xlsm", vbInformation, "Project Complete"
           
End Sub