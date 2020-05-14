Imports Microsoft.Office.Interop.Excel
Public Class ICRUPCS1
    Dim cf() As String
    Dim codes() As Object
    Dim cfmax As Integer
    Dim AZ As String
    Dim chkRecap As String
    Dim chkNewPage As String
    Dim chkONHAND As String
    Dim chkSHOWONLYOH As String
    Dim chkSHIPREC As String
    Dim r As Long
    Dim z As String
    Dim zz As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Get Run-Time options

        ASCMAIN1.Progress("Run-Time Options", "")
        AZ = ""




        'excuse me??
        'Call Get_Codes(AZ, cfmax, chkNewPage, cf(), codes(), "")

        ' ---> where are these options coming from??
        'chkONHAND = SRead(opts, "CHKONHAND", 2)
        'chkSHOWONLYOH = SRead(opts, "CHKSHOWONLYOH", 2)
        'chkSHIPREC = SRead(opts, "CHKSHIPREC", 2)

        ' Set up Work File Definition using X's and 0's and 0.01's as required

        ASCMAIN1.Progress("Initialize Work Tables", "")
        sql = "Select 0 RECORD_NO, 0 COUNTER, 0 ON_HAND, STYLE_CODE, COLOR_CODE, " & vbCrLf _
        & Get_GString(1, 30, AZ) & vbCrLf _
         & " from ICTSTYC1 where ROWNUM < 1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ASTSRPT0", 1))

        sql = "Select ICTSTYC4.UPC_CODE, ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC," & vbCrLf _
        & " ICTSTYC4.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTCOLR1.NRF_COLOR_CODE, ICTSTYL1.SUB_UNIT_PACK_QTY," & vbCrLf _
        & " ICTSTYC3.SIZE_CODE, ICTSIZE1.NRF_SIZE_CODE, ICTSTYL1.STYLE_RETAIL" & vbCrLf _
        & " from ICTSTYL1, ICTSTYC4, ICTSTYC3, ICTCOLR1, ICTSIZE1 " & vbCrLf _
        & " where ICTSTYL1.STYLE_CODE = ICTSTYC4.STYLE_CODE" & vbCrLf _
        & " AND ICTSTYL1.STYLE_CODE = ICTSTYC4.STYLE_CODE" & vbCrLf _
        & " AND ICTCOLR1.COLOR_CODE = ICTSTYC4.COLOR_CODE" & vbCrLf _
        & " AND ICTSTYC3.SIZE_CODE = ICTSIZE1.SIZE_CODE" & vbCrLf _
        & " AND ROWNUM < 1"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTSTYL1", 1))


        ' Prepare Work File with Data from Server

        If EMsg <> "" Then
            Exit Sub
        End If

        MyBase.Get_SQL("*")

        ASCMAIN1.Progress("Initialize Server", "")
        sql = " SELECT * FROM (" & vbCrLf _
        & " (SELECT ICTSTYC4.UPC_CODE," & vbCrLf _
        & " ICTSTYC4.STYLE_CODE, ICTSTYL1.STYLE_DESC, " & vbCrLf _
        & " ICTSTYC4.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTCOLR1.NRF_COLOR_CODE" & vbCrLf _
        & ", ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYC3.SIZE_CODE, ICTSIZE1.NRF_SIZE_CODE, " & vbCrLf _
        & " ICTSTYL1.STYLE_RETAIL " & vbCrLf _
        & "  FROM ICTSTYL1, ICTSTYC4, ICTSTYC3, ICTCOLR1, ICTSIZE1" & vbCrLf _
        & "  WHERE " & vbCrLf _
        & "  ICTSTYL1.STYLE_CODE = ICTSTYC4.STYLE_CODE AND" & vbCrLf _
        & "  ICTSTYC4.STYLE_CODE = ICTSTYC3.STYLE_CODE AND" & vbCrLf _
        & "  ICTSTYC4.COLOR_CODE = ICTSTYC3.COLOR_CODE AND" & vbCrLf _
        & "  ICTSTYC4.COLOR_CODE = ICTCOLR1.COLOR_CODE AND" & vbCrLf _
        & "  ICTSTYC3.SIZE_CODE = ICTSIZE1.SIZE_CODE AND" & vbCrLf _
        & "  ICTSTYC4.SIZE_INDEX = ICTSTYC3.SIZE_INDEX" & vbCrLf _
        & " AND ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
        & sql_JOIN & sql_WHERE & vbCrLf _
        & ") " & vbCrLf _
        & "  UNION" & vbCrLf _
        & " (SELECT ICTSTYC2.UPC_CODE, ICTSTYC2.STYLE_CODE," & vbCrLf _
        & " ICTSTYL1.STYLE_DESC, " & vbCrLf _
        & " ICTSTYC2.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTCOLR1.NRF_COLOR_CODE" & vbCrLf _
        & ", ICTSTYL1.SUB_UNIT_PACK_QTY, '90001' SIZE_CODE, '' NRF_SIZE_CODE, " & vbCrLf _
        & " ICTSTYL1.STYLE_RETAIL " & vbCrLf _
        & "  FROM ICTSTYL1, ICTSTYC2, ICTCOLR1" & vbCrLf _
        & "   WHERE " & vbCrLf _
        & "  ICTSTYL1.STYLE_CODE = ICTSTYC2.STYLE_CODE AND" & vbCrLf _
        & "  ICTSTYC2.COLOR_CODE = ICTCOLR1.COLOR_CODE" & vbCrLf _
        & " AND ICTSTYL1.STYLE_STATUS = 'A'" & vbCrLf _
        & sql_JOIN & sql_WHERE & vbCrLf _
        & "))" & vbCrLf _
        & " WHERE UPC_CODE IS NOT NULL" & vbCrLf _
        & " ORDER BY UPC_CODE, STYLE_CODE "

        ASCMAIN1.Progress("Record", "")

        For Each row As DataRow In ASCDATA1.GetDataTable(sql).Select
            r = r + 1
            If r Mod 100 = 0 Then
                z = row.Item(0) & "-" & CStr(r)
                ASCMAIN1.Progress("-", z)
            End If
            WriteRecords(row)
        Next

        ' Prepare Report File, w/Consolidations & Recaps as required

        '----> what do i do with this??
        'Build_Report_File(1, 2, dynASWSRPT0, "ASWSRPT0", cfmax, AZ, "N", "", cf(), "STYLE_CODE,COLOR_CODE")

        '----> and this??
        ASCMAIN1.sql = "Update ASWGROUP set GROUP_CODE = 'STOCK', GROUP_DESC = 'Stock Item'" & vbCrLf _
         & " where GROUP_KEY = 'Customer:' and (GROUP_CODE = '' or GROUP_CODE = 'STOCK')"
        ASCDATA1.ExecuteSQL()

        ' Wrap up

        'Remove all records where there is no On-hand value
        If chkSHOWONLYOH = "1" Then
            For Each row As DataRow In dst.Tables("ASTSRPT1").Select
                If Val(row.Item("SUM_ON_HAND") & "") = 0 Then
                    row.Delete()
                End If
            Next
        End If

        If Absx1.chkFor("CHKEXCEL").Checked Then
            PrepareExcel()
        End If

    End Sub
    Private Sub WriteRecords(ByVal addRow As DataRow)
        Dim rowASTSRPT0 As DataRow = dst.Tables("ASTSRPT0").NewRow
        rowASTSRPT0.Item("RECORD_NO") = r
        rowASTSRPT0.Item("STYLE_CODE") = addRow.Item("STYLE_CODE")
        rowASTSRPT0.Item("COLOR_CODE") = addRow.Item("COLOR_CODE")
        rowASTSRPT0.Item("COUNTER") = 1

        SetKey(addRow)
        dst.Tables("ASTSRPT0").Rows.Add(rowASTSRPT0)

        If Not dst.Tables("").Rows.Contains(addRow.Item("RECORD_NO")) Then
            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").NewRow
            For i As Integer = 0 To rowICTSTYL1.ItemArray.Count - 1
                rowICTSTYL1.Item(i) = addRow.Item(i)
                dst.Tables("ICTSTYL1").Rows.Add(rowICTSTYL1)
            Next
        End If
    End Sub
    Private Sub SetKey(ByVal addRow As DataRow)


        For j As Integer = 1 To cfmax
            zz = addRow.Item(j - 1).Value & ""
            z = Format$(j, "0")
            addRow.Item("G" & z) = cf(2, j) & ":" & zz
        Next j
        Return
    End Sub
    Sub PrepareExcel()
        ASCMAIN1.Progress("Building Excel Files", "")
        Dim objApp As Application
        Dim objBook As Workbook
        Dim objSheet As Worksheet
        'Dim dynwk As Recordset
        Dim txtNote As String = ""
        Dim txtHead As String = ""
        objApp = New Application
        objApp.Visible = False ' SET TO FALSE BEFORE RELEASE
        objApp.UserControl = True
        objBook = objApp.Workbooks.Add
        objApp.DisplayAlerts = False
        Dim i As Integer = 1
        If i > objBook.Worksheets.Count Then
            objBook.Worksheets.Add(after:=objBook.Worksheets(objBook.Worksheets.Count))
        End If
        objSheet = objBook.Worksheets(i)
        Select Case i
            Case 1
                sql = "SELECT ICWSTYL1.* "
                sql = sql & " INTO EXCEL" & Trim(Str(i)) & " From ICWSTYL1 "
                sql = sql & " ORDER BY ICWSTYL1.UPC_CODE, ICWSTYL1.STYLE_CODE"
                objSheet.Range("A1", "A1").Value = "UPC #"
                objSheet.Range("B1", "B1").Value = "STYLE #"
                objSheet.Range("C1", "C1").Value = "DESCRIPTION"
                objSheet.Range("D1", "D1").Value = "COLOR"
                objSheet.Range("F1", "F1").Value = "NRF COLOR"
                objSheet.Range("G1", "G1").Value = "PP"
                objSheet.Range("H1", "H1").Value = "SIZE"
                objSheet.Range("I1", "I1").Value = "NRF SIZE"
                objSheet.Range("J1", "J1").Value = "SUG. RTL"
                objSheet.Range("K1", "K1").Value = "SEL. COD."
                txtNote = ""
                txtHead = "VANDALE INDUSTRIES / RAMPAGE INNERWEAR"
        End Select
        objSheet.Name = "Report" & Trim(Str(i))
        objSheet.Range("A2").CopyFromRecordset(ASCDATA1.GetDataTable(sql))
        objSheet.Range("A2", "K2").Insert(XlDirection.xlToLeft)
        objSheet.Range("A1", "K1").Insert(XlDirection.xlDown)
        With objSheet.Range("A1", "K1")
            .MergeCells = True
            .Value = txtHead
        End With
        With objSheet.Range("A3", "K3")
            .MergeCells = True
            .Value = txtNote
            .AutoFormat(XlRangeAutoFormat.xlRangeAutoFormatSimple)


        End With
        With objSheet.Range("A1", "K3")
            .Font.Bold = True
            .HorizontalAlignment = XlVAlign.xlVAlignCenter
        End With

        With objSheet.Range("D3", "E3")
            .MergeCells = True
        End With

        objSheet.Columns.AutoFit()

        objApp.Visible = True
        objApp.UserControl = True
        objSheet = Nothing
        objBook = Nothing
        objApp = Nothing
    End Sub
    Public Overrides Sub Print_Report()
        SUBT = ""
        CR_params.Add("RECAP", chkRecap)
        CR_params.Add("NEWPAGE", chkNewPage)
        CR_params.Add("RC", aRC)
        CR_params.Add("HG1", "")
        CR_params.Add("HG2", "")
        CR_params.Add("HG3", "")
        CR_params.Add("HG4", "")
        CR_params.Add("HG5", "")
        CR_params.Add("HG6", "")
        CR_params.Add("HG7", "")
        CR_params.Add("LVLS", CStr(cfmax + 1))
        CR_params.Add("ONHAND", chkONHAND)
        CR_params.Add("SHIPFLAG", chkSHIPREC)
        Generate_Report(RPT, , SUBT)
    End Sub
    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub
    'Sub Get_SQL()
    '        Dim jz As String
    '        Dim z As String
    '        Dim y As String
    '        Dim f As String
    '        Dim j As Integer

    '        sqllist = ""
    '        sqllist2 = ""
    '        sqljoin = ""
    '        sqlwhere = ""
    '        sqltables = ""

    '        If Sel(1, 1) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.SALES_DIVISION_CODE " & Sel(2, 1) & " in (" & Sel(1, 1) & ")"
    '        End If
    '        If Sel(1, 2) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.FABRIC_CODE " & Sel(2, 2) & " in (" & Sel(1, 2) & ")"
    '        End If
    '        If Sel(1, 3) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.SEASON_CODE " & Sel(2, 3) & " in (" & Sel(1, 3) & ")"
    '        End If
    '        If Sel(1, 4) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTBODY2.MASTER_BODY_CODE " & Sel(2, 4) & " in (" & Sel(1, 4) & ")"
    '            f = "ICTBODY2"
    '        GoSub Get_SQLx_Join
    '        End If
    '        If Sel(1, 5) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.SUB_BODY_CODE " & Sel(2, 5) & " in (" & Sel(1, 5) & ")"
    '        End If
    '        If Sel(1, 6) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.CUST_CODE " & Sel(2, 6) & " in (" & Sel(1, 6) & ")"
    '        End If
    '        If Sel(1, 7) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.STYLE_CODE " & Sel(2, 7) & " in (" & Sel(1, 7) & ")"
    '        End If
    '        If Sel(1, 8) <> "" Then
    '            sqlwhere = sqlwhere & " and ICTSTYL1.CMT_NO " & Sel(2, 8) & " in (" & Sel(1, 8) & ")"
    '        End If

    '        Exit Sub

    'Get_SQLx_Join:
    '        If InStr(sqltables, "," & f) = 0 Then
    '            sqltables = sqltables & "," & f
    '            Select Case f
    '                Case "ICTBODY2"
    '                    jz = "ICTBODY2.SUB_BODY_CODE (+) = ICTSTYL1.SUB_BODY_CODE"
    '            End Select
    '            sqljoin = sqljoin & "   and " & jz
    '        End If
    '        Return

    'End Sub

End Class