Imports Microsoft.Office.Interop

Public Class POCMAIN1

    Public Shared Function Get_sql_Integrity_Check()
        Dim sqlIC As String = "Select * from (" & vbCrLf _
            & "Select PO_ORDER_NO, PO_ORDER_LNO, Max (PO_STATUS) PO_STATUS" & vbCrLf _
            & ", Sum (PO_QTY_ORD) PO_QTY_ORD, Sum (PO_QTY_SHP) PO_QTY_SHP" & vbCrLf _
            & ", Sum (PO_QTY_REC) PO_QTY_REC, Sum (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", Sum (PS_QTY_SHP) PS_QTY_SHP, Sum (PS_QTY_REC) PS_QTY_REC, Sum (SHPS) SHPS FROM (" & vbCrLf _
            & "Select PO_ORDER_NO, PO_ORDER_LNO, PO_STATUS, PO_QTY_ORD, PO_QTY_SHP, PO_QTY_REC, PO_QTY_OPN" & vbCrLf _
            & ", 0 PS_QTY_SHP, 0 PS_QTY_REC, 0 SHPS FROM POTORDR2" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO, NULL PO_STATUS" & vbCrLf _
            & ", 0 PO_QTY_ORD, 0 PO_QTY_SHP, 0 PO_QTY_REC, 0 PO_QTY_OPN, " & vbCrLf _
            & " Sum (POTSHIP3.PO_QTY_SHP) PS_QTY_SHP, Sum (POTSHIP3.PO_QTY_REC) PS_QTY_REC, COUNT (*) SHPS" & vbCrLf _
            & " from POTSHIP3 GROUP BY POTSHIP3.PO_ORDER_NO, POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & ") group by PO_ORDER_NO, PO_ORDER_LNO" & vbCrLf _
            & ") where SHPS <> 0" & vbCrLf _
            & " and (PO_QTY_SHP <> PS_QTY_SHP or PO_QTY_REC <> PS_QTY_REC" & vbCrLf _
            & "   or ((NVL(PO_STATUS,'?') = 'O' or NVL(PO_QTY_OPN,0) <> 0) and GREATEST(0,NVL(PO_QTY_ORD,0) - NVL(PO_QTY_SHP,0)) <> NVL(PO_QTY_OPN,0))" & vbCrLf _
            & "   or (NVL(PO_STATUS,'?') = 'O' and PO_QTY_OPN = 0)" & vbCrLf _
            & "   or (NVL(PO_STATUS,'?') = 'C' and PO_QTY_OPN <> 0)" & vbCrLf _
            & ")" & vbCrLf

        Return sqlIC
    End Function

    'Public Shared Sub POSHIPCHK()

    '    Dim m As String = "POs or Shipments Found Out Of Balance" & vbCrLf
    '    m = m & "Please Contact Wayne Immediatly To" & vbCrLf
    '    m = m & "Notify Him About This Message!!" & vbCrLf

    '    ASCMAIN1.sql = "SELECT COUNT(*) FROM (" & vbCrLf _
    '    & " SELECT " & vbCrLf _
    '    & " POTORDR2.PO_ORDER_NO," & vbCrLf _
    '    & " POTORDR2.PO_ORDER_LNO," & vbCrLf _
    '    & " POTSHIP3.PO_SHIPMENT_NO," & vbCrLf _
    '    & " POTSHIP3.PO_SHIPMENT_LNO," & vbCrLf _
    '    & " POTORDR2.STYLE_CODE," & vbCrLf _
    '    & " POTORDR2.COLOR_CODE," & vbCrLf _
    '    & " SUM(POTORDR2.PO_QTY_SHP - POTORDR2.PO_QTY_REC) PO_OPN," & vbCrLf _
    '    & " SUM(POTSHIP3.PO_QTY_SHP - POTSHIP3.PO_QTY_REC) SHP_OPN" & vbCrLf _
    '    & " FROM POTORDR2, POTSHIP3" & vbCrLf _
    '    & " WHERE POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
    '    & " AND PO_DATE_SHIP_BY >= '01-JAN-2001'" & vbCrLf _
    '    & " AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
    '    & " HAVING SUM(POTORDR2.PO_QTY_SHP - POTORDR2.PO_QTY_REC) <> SUM(POTSHIP3.PO_QTY_SHP - POTSHIP3.PO_QTY_REC)" & vbCrLf _
    '    & " GROUP BY" & vbCrLf _
    '    & " POTORDR2.PO_ORDER_NO," & vbCrLf _
    '    & " POTORDR2.PO_ORDER_LNO," & vbCrLf _
    '    & " POTSHIP3.PO_SHIPMENT_NO," & vbCrLf _
    '    & " POTSHIP3.PO_SHIPMENT_LNO," & vbCrLf _
    '    & " POTORDR2.STYLE_CODE," & vbCrLf _
    '    & " POTORDR2.COLOR_CODE," & vbCrLf _
    '    & " POTORDR2.PO_QTY_ORD" & vbCrLf _
    '    & " ORDER BY POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE)"

    '    Dim row As DataRow = ASCDATA1.GetDataRow
    '    If row IsNot Nothing Then
    '        If Val(row.Item(0) & "") > 0 Then
    '            MsgBox(m)
    '        End If
    '    End If
    'End Sub

    Public Shared Sub Check_Status(frm As ASFBASE0)
        ASCMAIN1.sql = "SELECT STYLE_CODE, COLOR_CODE, WHSE_CODE, SUM (ONPO) ONPO, SUM (OPN) OPN FROM (" & vbCrLf _
            & "SELECT STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_ORDER ONPO, 0 OPN FROM ICTSTAT2 WHERE WHSE_QTY_ON_ORDER <> 0 UNION" & vbCrLf _
            & "SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE, 0 ONPO, SUM (POTORDR2.PO_QTY_OPN) OPN" & vbCrLf _
            & "FROM POTORDR2,POTORDR1" & vbCrLf _
            & "WHERE POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
            & "GROUP BY POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
            & ") GROUP BY STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & "HAVING  SUM (ONPO)  <> SUM (OPN)"
        Dim dt As DataTable = ASCDATA1.GetDataTable
        If dt.Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(dt, frm, "Styles with Open PO Out of Balance - please report to ABS")
            End Using
        End If

        ' TO FIX THE ABOVE ISSUE, USE THIS:
        '        "BEGIN DECLARE CURSOR C1 IS" _
        '& "SELECT STYLE_CODE, COLOR_CODE, WHSE_CODE, SUM (ONPO) ONPO, SUM (OPN) OPN FROM (" _
        '& "SELECT STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_ORDER ONPO, 0 OPN FROM ICTSTAT2 WHERE WHSE_QTY_ON_ORDER <> 0 UNION" _
        '& "SELECT POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE, 0 ONPO, SUM (POTORDR2.PO_QTY_OPN) OPN" _
        '& "FROM POTORDR2,POTORDR1" _
        '& "WHERE POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" _
        '& "GROUP BY POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTORDR1.WHSE_CODE" _
        '& ") GROUP BY STYLE_CODE, COLOR_CODE, WHSE_CODE" _
        '& "HAVING  SUM (ONPO)  <> SUM (OPN);" _
        '& "BEGIN FOR R1 IN C1 LOOP" _
        '& "UPDATE ICTSTAT2 SET WHSE_QTY_ON_ORDER = R1.OPN" _
        '& "WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE AND WHSE_CODE = R1.WHSE_CODE;" _
        '& "END LOOP; END; END;"


        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS" & vbCrLf _
            & "SELECT STYLE_CODE, COLOR_CODE, WHSE_CODE, SUM (SHP) SHP, SUM (STA) STA FROM (" & vbCrLf _
            & "SELECT" & vbCrLf _
            & "STYLE_CODE, COLOR_CODE, WHSE_CODE, 0 SHP, WHSE_QTY_TRAN STA" & vbCrLf _
            & "FROM ICTSTAT2 WHERE WHSE_QTY_TRAN <> 0" & vbCrLf _
            & "UNION" & vbCrLf _
            & "SELECT" & vbCrLf _
            & "POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE, SUM (POTSHIP3.PO_QTY_SHP) SHP, 0 STA" & vbCrLf _
            & "FROM POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
            & "WHERE POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & "AND POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & "AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "GROUP BY POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE, POTSHIP1.WHSE_CODE" & vbCrLf _
            & ") GROUP BY STYLE_CODE, COLOR_CODE, WHSE_CODE" & vbCrLf _
            & "HAVING SUM (SHP) <> SUM (STA);" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ICTSTAT2 SET WHSE_QTY_TRAN = R1.SHP " & vbCrLf _
            & "WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE AND WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        Dim dt2 As DataTable = ASCDATA1.GetDataTable
        If dt2.Rows.Count <> 0 Then
            Using F As New ASFMSGBF
                F.Show_grd(dt2, frm, "Styles with Qty Afloat Out of Balance - please report to ABS")
            End Using
        End If
    End Sub

    Public Shared Sub Setup_PO_Change_Details(frmASFBASE0 As ASFBASE0)
        For Each row As DataRow In frmASFBASE0.dst.Tables("POTORDR1").Select("")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Dim PO_PRINTED_IND As String = row.Item("PO_PRINTED_IND") & ""
            Dim PO_HDR_CTR_REV As Int32 = Val(row.Item("PO_HDR_CTR_REV") & "")
            Build_POTORDRZ(frmASFBASE0.dst.Tables("POTORDRZ"), _
                                        PO_ORDER_NO, _
                                        IIf(PO_PRINTED_IND = "1", PO_HDR_CTR_REV - 1, PO_HDR_CTR_REV - 1), _
                                        IIf(PO_PRINTED_IND = "1", PO_HDR_CTR_REV, -1))


        Next
    End Sub

    Public Shared Sub Build_POTORDRZ(tbl As DataTable, PO_ORDER_NO As String, from_Rev As Int32, to_Rev As Int32)
        Dim PO_HDR_CTR_REV As Int32 = from_Rev + 1

        For i As Integer = 0 To 1
            If i = 0 Then
                ASCMAIN1.sql = "Select * from POTORDRZ where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_HDR_CTR_REV = " & CStr(from_Rev)
            Else
                If to_Rev = -1 Then
                    ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                Else
                    ASCMAIN1.sql = "Select * from POTORDRZ where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_HDR_CTR_REV = " & CStr(to_Rev)
                End If
            End If

            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim PO_ORDER_LNO As Int32 = Val(row.Item("PO_ORDER_LNO") & "")
                Dim rowPOTORDRZ As DataRow = tbl.Rows.Find(New Object() {PO_ORDER_NO, PO_HDR_CTR_REV, PO_ORDER_LNO})
                If rowPOTORDRZ Is Nothing Then
                    rowPOTORDRZ = tbl.NewRow
                    rowPOTORDRZ.Item("PO_ORDER_NO") = PO_ORDER_NO
                    rowPOTORDRZ.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV
                    rowPOTORDRZ.Item("PO_ORDER_LNO") = PO_ORDER_LNO
                    tbl.Rows.Add(rowPOTORDRZ)
                End If
                Dim SFX As String = IIf(i = 0, "_PREV", "")
                For Each COLUMN_NAME As String In New String() {"STYLE_CODE", "COLOR_CODE", "PO_QTY_ORD", "PO_COST", "PO_DATE_SHIP_BY", "PO_STATUS", "CARTON_PACK_QTY"}
                    rowPOTORDRZ.Item(COLUMN_NAME & SFX) = row.Item(COLUMN_NAME)
                Next
                If rowPOTORDRZ.Item("STYLE_CODE") & "" = "" Then
                    rowPOTORDRZ.Item("STYLE_CODE") = rowPOTORDRZ.Item("STYLE_CODE_PREV")
                    rowPOTORDRZ.Item("COLOR_CODE") = rowPOTORDRZ.Item("COLOR_CODE_PREV")
                End If
            Next
        Next
    End Sub

    Public Shared Function Generate_Expedite_POs_XLS(frmASFBASE0 As ASFBASE0, EXPEDITE_NO As String, PO_ORDER_NOs As List(Of String)) As String

        Dim pbInt As Integer = 0
        Dim excelFile As String = ""
        Dim FILE_NAME As String = ""

        Dim xlPages As New Dictionary(Of Integer, Integer)

        Dim XLS As Excel.Application = New Excel.Application
        Dim XWB As Excel.Workbook = XLS.Workbooks.Add
        Dim XWS As Excel.Worksheet
        XWS = XWB.Worksheets(3)
        XWS.Delete()
        XWS = XWB.Worksheets(2)
        XWS.Delete()

        Dim rng As Excel.Range

        Dim SHTCOUNT As Integer = 0

        For Each PO_ORDER_NO As String In PO_ORDER_NOs

            Dim rowPOTORDR1 As DataRow = frmASFBASE0.LookUp("POTORDR1", PO_ORDER_NO)
            SHTCOUNT += 1

            'If SHTCOUNT <= 3 Then
            '    XWS = XWB.Sheets(SHTCOUNT)
            'Else
            '    XWS = XWB.Sheets.Add
            'End If

            If SHTCOUNT = 1 Then
                XWS = XWB.Sheets(1)
            Else
                XWS = XWB.Sheets.Add
            End If

            XWS.Name = PO_ORDER_NO

            'insert logo
            Dim LOGO_FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & ".jpg"
            If My.Computer.FileSystem.FileExists(LOGO_FILENAME) Then
                rng = XWS.Range("A" & CStr(1) & ":" & "B" & CStr(6))
                XLS.InsertPictureInRange(LOGO_FILENAME, rng, XWS)
            End If

            rng = XWS.Range("E1:E1")
            rng.FormulaR1C1 = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME")
            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            With rng.Font
                .Name = "Georgia"
                .Size = 18
                .Color = Color.FromArgb(79, 129, 189)
                .Bold = False
            End With

            rng = XWS.Range("E3:E3")
            rng.FormulaR1C1 = "Open PO Report"
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            rng = XWS.Range("E4:E4")
            rng.FormulaR1C1 = "Please Confirm Ship Dates (only) - for All Styles"
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            rng = XWS.Range("E5:E5")
            rng.FormulaR1C1 = "Generated: " & Now.ToShortDateString()
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            rng = XWS.Range("E6:E6")
            rng.FormulaR1C1 = rowPOTORDR1.Item("VEND_NAME")
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            Dim R_Starting As Integer
            Dim COLs As String()

            Dim C As Integer = 0
            Dim R As Integer

            R_Starting = 7
            COLs = {"Rev", "PO No", "Message", "Code", "Cancel"}
            Create_Worksheet_Headers(XWS, R_Starting, COLs)
            R = R_Starting
            R += 1
            XWS.Cells(R, 1).VALUE = (rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
            XWS.Cells(R, 3).VALUE = PO_ORDER_NO
            XWS.Cells(R, 5).VALUE = rowPOTORDR1.Item("PO_MESSAGE")
            XWS.Cells(R, 7).VALUE = rowPOTORDR1.Item("VEND_CODE")
            XWS.Cells(R, 9).VALUE = rowPOTORDR1.Item("PO_DATE_CANCEL") ' Format(rowPOTORDR1.Item("PO_DATE_CANCEL"), "MM/dd/yyyy")

            R_Starting = 9
            COLs = {"Line", "Style", "Description", "Color", "Color Desc", "Qty Ordered", "Qty Open", "Ship Date", "Supplier Comments"} ' , "Price"
            Create_Worksheet_Headers(XWS, R_Starting, COLs)

            'freeze header region
            XWS.Range("A" & CStr(R_Starting + 1), "A" & CStr(R_Starting + 1)).Select()
            XWS.Application.ActiveWindow.FreezePanes = True

            R = R_Starting + 2
            Dim rowsToProcess As Integer = frmASFBASE0.dst.Tables("POTORDR2").Select("").Length + 1
            Dim currentRow As Integer = 1
            Dim currentPage As Integer = 1
            xlPages.Clear()

            For Each row As DataRow In frmASFBASE0.dst.Tables("POTORDR2").Select("")

                pbInt = CInt((currentRow / rowsToProcess) * 100)
                Dim PO_ORDER_LNO As String = Val(row.Item("PO_ORDER_LNO") & "")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE") & ""

                Dim PO_QTY_ORD As String = Val(row.Item("PO_QTY_ORD") & "")
                Dim PO_QTY_OPN As String = Val(row.Item("PO_QTY_OPN") & "")
                Dim PO_DATE_SHIP_BY As Date = row.Item("PO_DATE_SHIP_BY")
                Dim PO_COST As String = Val(row.Item("PO_COST") & "")

                Dim rowICTSTYL1 As DataRow = frmASFBASE0.LookUp("ICTSTYL1", STYLE_CODE)
                Dim STYLE_DESC As String = rowICTSTYL1.Item("STYLE_DESC") & ""

                Dim COLOR_DESC As String = COLOR_CODE
                Dim rowICTCOLR1 As DataRow = frmASFBASE0.LookUp("ICTCOLR1", COLOR_CODE)
                If rowICTCOLR1 IsNot Nothing Then
                    COLOR_DESC = rowICTCOLR1.Item("COLOR_DESC") & ""
                End If

                XWS.Cells(R, 1).VALUE = PO_ORDER_LNO
                XWS.Cells(R, 3).VALUE = STYLE_CODE
                XWS.Cells(R, 5).VALUE = STYLE_DESC
                XWS.Cells(R, 7).VALUE = COLOR_CODE
                XWS.Cells(R, 9).VALUE = COLOR_DESC
                XWS.Cells(R, 11).VALUE = PO_QTY_ORD
                XWS.Cells(R, 13).VALUE = PO_QTY_OPN
                'XWS.Cells(R, 15).VALUE = PO_COST
                XWS.Cells(R, 15).VALUE = PO_DATE_SHIP_BY.ToShortDateString  ' Format(PO_DATE_SHIP_BY, "MM/dd/yyyy")

                'rng = XWS.Range("H" & CStr(R) & ":" & "H" & CStr(R))
                'rng.Font.Bold = True
                'rng.Style = "Currency"

                'For Each cellSet As String In New String() {"D:E", "G:H", "J:K", "M:N"}
                '    Dim xlCells() As String = Split(cellSet, ":")
                '    rng = XWS.Range(xlCells(0) & CStr(R) & ":" & xlCells(1) & CStr(R + 3))
                '    rng.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous, Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin, Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic)
                '    rng = XWS.Range(xlCells(0) & CStr(R) & ":" & xlCells(0) & CStr(R + 3))
                '    With rng.Interior
                '        .Color = Color.FromArgb(242, 242, 242)
                '        .TintAndShade = 0
                '        .PatternTintAndShade = 0
                '    End With
                'Next

                R += 1
                currentRow += 1

                If ((currentPage - 1) * 13) + 12 = currentRow - 1 Then
                    'at current column widths trying to show more than 12 records on the first page
                    'or 13 on subsequent pages scales the record down and creats white space pn the right margin
                    xlPages.Add(currentPage, (R - 1))
                    currentPage += 1
                End If

            Next

            rng = XWS.Range("A:S")
            rng.EntireColumn.AutoFit()

            'rng = XWS.Range(ASCMAIN1.XC(11))
            'rng.EntireColumn.ColumnWidth = 20
            'rng = XWS.Range("J:J")
            'rng.EntireColumn.AutoFit()

            XWS.PageSetup.PrintTitleRows = "$9:$9"

            XWS.PageSetup.FitToPagesWide = 1

            XWS.Application.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlPageBreakPreview

            'replace automatic page breaks with my calculated page breaks
            For i = 1 To XWS.HPageBreaks.Count
                If i > XWS.HPageBreaks.Count Then Exit For
                If xlPages.ContainsKey(i) Then
                    rng = XWS.Range("A" & xlPages(i).ToString)
                    XWS.HPageBreaks(i).Location = rng
                End If
            Next

            'not needed right now
            'rng = XWS.Range("O:O")
            'For i = 1 To XWS.VPageBreaks.Count
            '    XWS.VPageBreaks(i).Location = rng
            'Next

        Next

        XLS.ActiveWindow.View = Microsoft.Office.Interop.Excel.XlWindowView.xlNormalView

        Dim xlsFileName_sfx As String = ""
        Dim xlsFileName As String = ""
        Dim xlsControlNo As String = EXPEDITE_NO

        FILE_NAME = "Open_PO_Report"

        Do
            Try
                xlsFileName = FILE_NAME & "_" & xlsControlNo
                'progressSplash.UpdateProgress("Saving Spreadsheet", xlsFileName & ".xls", "", pbInt)
                excelFile = ASCMAIN1.Folders("Temp") & xlsFileName & ".xls"
                XWB.SaveAs(excelFile)
                ' progressSplash.UpdateProgress("", "", "Done", 100)
                xlsFileName_sfx = ""
            Catch ex As Exception
                xlsFileName_sfx = CStr(Val(xlsFileName_sfx) + 1)
            End Try
        Loop While xlsFileName_sfx <> "" And Val(xlsFileName_sfx) < 10

        XWB.Close()
        XWB = Nothing
        XLS = Nothing
        Return xlsFileName
    End Function

    Public Shared Sub Create_Worksheet_Headers(xlws As Excel.Worksheet, R As Integer, COLs() As String)

        Dim rng As Excel.Range

        Dim C As Integer = 0

        For Each COL As String In COLs
            C += 1
            Dim XX As String = ASCMAIN1.Excel_Cell(0, C)
            rng = xlws.Range(ASCMAIN1.Excel_Cell(0, C))
            Format_Worksheet_Header(COL, xlws, xlws.Range(ASCMAIN1.Excel_Cell(R, C) & ":" & ASCMAIN1.Excel_Cell(R, C)))
            xlws.Cells(R, C).VALUE = COL

            If New String() {"Line", "Qty Ordered", "Qty Open", "Price"}.Contains(COL) Then
                rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                If COL = "Price" Then
                    rng.NumberFormat = "#,##0.00"
                Else
                    rng.NumberFormat = "#,##0"
                End If
            ElseIf New String() {"Ship Date"}.Contains(COL) Then
                rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                rng.NumberFormat = "MM/dd/yyyy"
            End If

            C += 1
            rng = xlws.Range(ASCMAIN1.Excel_Cell(0, C))
            rng.EntireColumn.ColumnWidth = 0.5
        Next
    End Sub

    Public Shared Sub Format_Worksheet_Header(headerText As String, xlws As Excel.Worksheet, headerRange As Excel.Range)
        With headerRange  ' XWS.Range(XC(i - 1, (S - 1) * 3 + 0), XC(i, (S - 1) * 3 + 2))
            .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternLinearGradient
            Dim grd2 As Microsoft.Office.Interop.Excel.LinearGradient
            grd2 = .Interior.Gradient
            Dim cs As Microsoft.Office.Interop.Excel.ColorStop
            cs = grd2.ColorStops.Add(0)
            cs.Color = Color.FromArgb(255, 255, 255)
            cs = grd2.ColorStops.Add(1)
            cs.Color = Color.FromArgb(79, 129, 189)
            cs.TintAndShade = 0
            grd2.Degree = 90
            headerRange.Merge()
            headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter

            With .Font
                .Name = "Calibri"
                .Size = 11
                .Bold = True
            End With
        End With

    End Sub

    Public Shared Sub Dependent_Updates(S As Integer, PO_ORDER_NO As String, Optional cancel_po As Boolean = False)

        ' there is the usual fuzziness around Close PO vs Cancel PO 
        ' since we support Delete (which is like Cancel PO for VANs purposes), we might weant to rename Cancel PO to Close PO

        ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select * from POTORDR1 where PO_ORDER_NO = '" & PO_ORDER_NO & "' for Update;" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Begin" & vbCrLf _
                & "    Declare " & vbCrLf _
                & "     Cursor C2 is Select * from POTORDR2 where PO_ORDER_NO = R1.PO_ORDER_NO for Update;" & vbCrLf _
                & "     QTY Number(8,0);" & vbCrLf _
                & "    Begin" & vbCrLf _
                & "     For R2 in C2 Loop" & vbCrLf _
                & "      QTY := " & CStr(S) & " * NVL(R2.PO_QTY_OPN,0);" & vbCrLf _
                & "      Update ICTSTAT2 Set WHSE_QTY_ON_ORDER = NVL(WHSE_QTY_ON_ORDER,0) + QTY" & vbCrLf _
                & "       where STYLE_CODE = R2.STYLE_CODE and COLOR_CODE = R2.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & "      If SQL%NOTFOUND then" & vbCrLf _
                & "       Insert into ICTSTAT2 (STYLE_CODE, COLOR_CODE, WHSE_CODE, WHSE_QTY_ON_ORDER)" & vbCrLf _
                & "        values (R2.STYLE_CODE, R2.COLOR_CODE, R1.WHSE_CODE, QTY);" & vbCrLf _
                & "      End If;" & vbCrLf _
                & IIf(cancel_po, _
                  "      Update POTORDR2 Set PO_QTY_OPN = 0, PO_STATUS = 'C' where Current of C2;", _
                  "") & vbCrLf _
                & "     End Loop;" & vbCrLf _
                & "    End;" & vbCrLf _
                & "   End;" & vbCrLf _
                & IIf(cancel_po, _
                  "      Update POTORDR1 Set PO_DATE_CANCELLED = TRUNC(SYSDATE), PO_STATUS = 'C' where Current of C1;", _
                  "") & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Shared Function Check_Changed_Fields(frmASFBASE0 As ASFBASE0, rowPOTORDR1 As DataRow) As Boolean

        Dim PO_ORDER_NO As String = rowPOTORDR1.Item("PO_ORDER_NO")

        Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        '  PO_HDR_CTR_REV += 1 ' already done in load record

        ' THIS ROUTINE IS NORMALLY CALLED BY POFORDR1, BUT AT VAN WE ALSO CALL IT FROM POFORDRA

        Dim LAST_DATE As Date = frmASFBASE0.DATETIME_STAMP
        If frmASFBASE0.EntryMode = "N" Then Stop
        Dim REV_LNO As Integer = 0

        Check_Changed_Fields = False

        frmASFBASE0.dst.Tables("POTORDXR").Rows.Clear()

        ASCMAIN1.Progress("Logging Header Changes")

        For i As Integer = 0 To rowPOTORDR1.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = frmASFBASE0.dst.Tables("POTORDR1").Columns(i).ColumnName

            If rowPOTORDR1.Item(COLUMN_NAME) & "" _
            <> rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Check_Changed_Fields = True
                ASCMAIN1.Progress("-", COLUMN_NAME)
                Dim rowPOTORDXR As DataRow = frmASFBASE0.dst.Tables("POTORDXR").NewRow
                With rowPOTORDXR
                    .Item("REV_NO") = PO_HDR_CTR_REV
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_USER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = rowPOTORDR1.Item(COLUMN_NAME)
                    .Item("EMODE") = frmASFBASE0.EntryMode
                End With
                frmASFBASE0.dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Check_Changed_Fields = True
            End If
        Next i

        ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowPOTORDR2_orig As DataRow In dt.Rows
            Dim PO_ORDER_LNO As Int64 = rowPOTORDR2_orig.Item("PO_ORDER_LNO")
            Dim rowPOTORDR2 As DataRow = frmASFBASE0.dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
            If rowPOTORDR2 Is Nothing Then ' Line was Deleted
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    Dim rowPOTORDXR As DataRow = frmASFBASE0.dst.Tables("POTORDXR").NewRow
                    With rowPOTORDXR
                        .Item("REV_NO") = PO_HDR_CTR_REV
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_USER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                        '.Item("NEW_VALUE") = ""
                        .Item("EMODE") = frmASFBASE0.EntryMode
                    End With
                    frmASFBASE0.dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Next

                Check_Changed_Fields = True
            Else
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    If rowPOTORDR2.Item(COLUMN_NAME) & "" <> rowPOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                        ' Value in Column was Changed
                        Dim rowPOTORDXR As DataRow = frmASFBASE0.dst.Tables("POTORDXR").NewRow
                        With rowPOTORDXR
                            .Item("REV_NO") = PO_HDR_CTR_REV
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_USER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                            .Item("NEW_VALUE") = rowPOTORDR2.Item(COLUMN_NAME)
                            .Item("EMODE") = frmASFBASE0.EntryMode
                        End With
                        frmASFBASE0.dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                        Check_Changed_Fields = True
                    End If
                Next
            End If
        Next

        For Each rowPOTORDR2 As DataRow In frmASFBASE0.dst.Tables("POTORDR2").Select("", "", DataViewRowState.Added)
            Dim PO_ORDER_LNO = rowPOTORDR2.Item("PO_ORDER_LNO")
            ' For i As Integer = 0 To dt.Columns.Count - 1
            Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
            Dim rowPOTORDXR As DataRow = frmASFBASE0.dst.Tables("POTORDXR").NewRow
            With rowPOTORDXR
                .Item("REV_NO") = PO_HDR_CTR_REV
                REV_LNO += 1
                .Item("REV_LNO") = REV_LNO
                .Item("PO_ORDER_NO") = PO_ORDER_NO
                .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                .Item("INIT_DATE") = LAST_DATE
                .Item("INIT_USER") = ASCMAIN1.USER_ID
                .Item("COLUMN_NAME") = COLUMN_NAME
                '.Item("OLD_VALUE") = ""
                .Item("NEW_VALUE") = "PO Line Added" ' rowPOTORDR2.Item(COLUMN_NAME)
                .Item("EMODE") = frmASFBASE0.EntryMode
            End With
            frmASFBASE0.dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
            Check_Changed_Fields = True
            'Next
        Next

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Public Shared Function Aged_PO(frmASFBASE0 As ASFBASE0, YP As String)

        Dim DTES(6) As String
        For I As Integer = 0 To 6
            Dim rowGLTPARM2 As DataRow = frmASFBASE0.LookUp("GLTPARM2", ASCMAIN1.Period_Calc(YP, I))
            DTES(I) = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")
        Next

        Dim sqlQTY As String = "THEN PO_QTY_OPN ELSE 0 END"
        Dim sqlAMT As String = "THEN PO_QTY_OPN * PO_COST ELSE 0 END"

        Dim SQL As String = "" _
            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", 0 PO_QTY_XIT" & vbCrLf _
            & ", 0 PO_CST_XIT" & vbCrLf _
            & ", 0 PO_QTY_XIT_CMO" & vbCrLf _
            & ", 0 PO_CST_XIT_CMO" & vbCrLf _
            & ", 0 PO_QTY_XIT_NMO" & vbCrLf _
            & ", 0 PO_CST_XIT_NMO" & vbCrLf _
            & ", 0 PO_QTY_XIT_2NMO" & vbCrLf _
            & ", 0 PO_CST_XIT_2NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (PO_QTY_OPN * PO_COST) PO_CST_OPN" & vbCrLf _
            & ", SUM (CASE WHEN                                     PO_DATE_ETA <= '" & DTES(0) & "' " & sqlQTY & ") PO_QTY_OPN_CMO" & vbCrLf _
            & ", SUM (CASE WHEN                                     PO_DATE_ETA <= '" & DTES(0) & "' " & sqlAMT & ") PO_CST_OPN_CMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(0) & "' and PO_DATE_ETA <= '" & DTES(1) & "' " & sqlQTY & ") PO_QTY_OPN_NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(0) & "' and PO_DATE_ETA <= '" & DTES(1) & "' " & sqlAMT & ") PO_CST_OPN_NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(1) & "' and PO_DATE_ETA <= '" & DTES(2) & "' " & sqlQTY & ") PO_QTY_OPN_2NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(1) & "' and PO_DATE_ETA <= '" & DTES(2) & "' " & sqlAMT & ") PO_CST_OPN_2NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(2) & "' and PO_DATE_ETA <= '" & DTES(3) & "' " & sqlQTY & ") PO_QTY_OPN_3NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(2) & "' and PO_DATE_ETA <= '" & DTES(3) & "' " & sqlAMT & ") PO_CST_OPN_3NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(3) & "' and PO_DATE_ETA <= '" & DTES(4) & "' " & sqlQTY & ") PO_QTY_OPN_4NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(3) & "' and PO_DATE_ETA <= '" & DTES(4) & "' " & sqlAMT & ") PO_CST_OPN_4NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(4) & "' and PO_DATE_ETA <= '" & DTES(5) & "' " & sqlQTY & ") PO_QTY_OPN_5NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(4) & "' and PO_DATE_ETA <= '" & DTES(5) & "' " & sqlAMT & ") PO_CST_OPN_5NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(5) & "'                                      " & sqlQTY & ") PO_QTY_OPN_6NMO" & vbCrLf _
            & ", SUM (CASE WHEN PO_DATE_ETA > '" & DTES(5) & "'                                      " & sqlAMT & ") PO_CST_OPN_6NMO" & vbCrLf _
            & " from POTORDR2" & vbCrLf _
            & " where PO_STATUS = 'O' and PO_QTY_OPN <> 0 " & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
            & ", SUM (POTSHIP3.PO_QTY_SHP) PO_QTY_XIT" & vbCrLf _
            & ", SUM (POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED) PO_CST_XIT" & vbCrLf _
            & ", SUM (CASE WHEN                                     POTSHIP1.PO_SHIP_ETA <= '" & DTES(0) & "' THEN POTSHIP3.PO_QTY_SHP ELSE 0 END) PO_QTY_XIT_CMO" & vbCrLf _
            & ", SUM (CASE WHEN                                     POTSHIP1.PO_SHIP_ETA <= '" & DTES(0) & "' THEN POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED ELSE 0 END) PO_CST_XIT_CMO" & vbCrLf _
            & ", SUM (CASE WHEN POTSHIP1.PO_SHIP_ETA > '" & DTES(0) & "' and POTSHIP1.PO_SHIP_ETA <= '" & DTES(1) & "' THEN POTSHIP3.PO_QTY_SHP ELSE 0 END) PO_QTY_XIT_NMO" & vbCrLf _
            & ", SUM (CASE WHEN POTSHIP1.PO_SHIP_ETA > '" & DTES(0) & "' and POTSHIP1.PO_SHIP_ETA <= '" & DTES(1) & "' THEN POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED ELSE 0 END) PO_CST_XIT_NMO" & vbCrLf _
            & ", SUM (CASE WHEN POTSHIP1.PO_SHIP_ETA > '" & DTES(1) & "'                                      THEN POTSHIP3.PO_QTY_SHP ELSE 0 END) PO_QTY_XIT_2NMO" & vbCrLf _
            & ", SUM (CASE WHEN POTSHIP1.PO_SHIP_ETA > '" & DTES(1) & "'                                      THEN POTSHIP3.PO_QTY_SHP * POTSHIP3.PO_COST_LANDED ELSE 0 END) PO_CST_XIT_2NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN" & vbCrLf _
            & ", 0 PO_CST_OPN" & vbCrLf _
            & ", 0 PO_QTY_OPN_CMO" & vbCrLf _
            & ", 0 PO_CST_OPN_CMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_2NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_2NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_3NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_3NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_4NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_4NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_5NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_5NMO" & vbCrLf _
            & ", 0 PO_QTY_OPN_6NMO" & vbCrLf _
            & ", 0 PO_CST_OPN_6NMO" & vbCrLf _
            & " from POTSHIP1,POTSHIP2,POTSHIP3,POTORDR2" & vbCrLf _
            & " where POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIP_STATUS = 'O'" & vbCrLf _
            & " group by POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf

        SQL = "Select '" & YP & "' OPS_YYYYPP, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", SUM (PO_QTY_XIT) PO_QTY_XIT" & vbCrLf _
            & ", SUM (PO_CST_XIT) PO_CST_XIT" & vbCrLf _
            & ", SUM (PO_QTY_XIT_CMO) PO_QTY_XIT_CMO" & vbCrLf _
            & ", SUM (PO_CST_XIT_CMO) PO_CST_XIT_CMO" & vbCrLf _
            & ", SUM (PO_QTY_XIT_NMO) PO_QTY_XIT_NMO" & vbCrLf _
            & ", SUM (PO_CST_XIT_NMO) PO_CST_XIT_NMO" & vbCrLf _
            & ", SUM (PO_QTY_XIT_2NMO) PO_QTY_XIT_2NMO" & vbCrLf _
            & ", SUM (PO_CST_XIT_2NMO) PO_CST_XIT_2NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & ", SUM (PO_CST_OPN) PO_CST_OPN" & vbCrLf _
            & ", SUM (PO_QTY_OPN_CMO) PO_QTY_OPN_CMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_CMO) PO_CST_OPN_CMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_NMO) PO_QTY_OPN_NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_NMO) PO_CST_OPN_NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_2NMO) PO_QTY_OPN_2NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_2NMO) PO_CST_OPN_2NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_3NMO) PO_QTY_OPN_3NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_3NMO) PO_CST_OPN_3NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_4NMO) PO_QTY_OPN_4NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_4NMO) PO_CST_OPN_4NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_5NMO) PO_QTY_OPN_5NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_5NMO) PO_CST_OPN_5NMO" & vbCrLf _
            & ", SUM (PO_QTY_OPN_6NMO) PO_QTY_OPN_6NMO" & vbCrLf _
            & ", SUM (PO_CST_OPN_6NMO) PO_CST_OPN_6NMO" & vbCrLf _
            & " from (" & SQL & ") group by STYLE_CODE, COLOR_CODE"

        Dim ICTSTKL2 As String = ASCMAIN1.Temp_Table(SQL)

        Return ICTSTKL2
    End Function
End Class