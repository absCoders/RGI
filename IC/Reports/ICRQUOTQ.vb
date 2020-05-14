Imports System.Text

Public Class ICRQUOTQ
    Dim S As New StringBuilder With {.Length = 0}
    Dim dayBreaks As Integer = 120

#Region "Report Standards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        RWU = "N"
        Get_PARM("ICTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()
        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        SUBT = ""
        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        ASCMAIN1.Progress("Now Preparing Dataset")
        Get_SQL("*")
        Dim sql_TABLE_NAMEs_orig As String = sql_TABLE_NAMEs
        Dim sql_JOIN_orig As String = sql_JOIN

        Dim sql_filter2 As String = ""

        '-- Shit you may need here --
        'sql_SELECT_cols, sql_TABLE_NAMEs, sql_WHERE, sql_JOIN, sql_filter, sql_filter2
        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC,")
        S.AppendLine("SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) AS NET_POS,")
        S.AppendLine("SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0)) AS IN_TRANS,")
        S.AppendLine("(SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) - SUM(NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0))) AS NOW_OH,")
        S.AppendLine("0 AS RECEIVED_01,")
        S.AppendLine("0 AS RECEIVED_02,")
        S.AppendLine("0 AS RECEIVED_03,")
        S.AppendLine("0 AS AGED_01,")
        S.AppendLine("0 AS AGED_02,")
        S.AppendLine("0 AS AGED_03")
        S.AppendLine("FROM ICTSTYL1, ICTSTAT2, ICTCOLR1")
        S.AppendLine("WHERE ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE")
        S.AppendLine("AND ICTSTAT2.COLOR_CODE = ICTCOLR1.COLOR_CODE")
        If Absx1.optFor("OPTASN").Value = "S" Then
            S.AppendLine("AND ICTSTYL1.CUST_CODE IS NULL")
        ElseIf Absx1.optFor("OPTASN").Value = "N" Then
            S.AppendLine("AND ICTSTYL1.CUST_CODE IS NOT NULL")
        End If
        S.AppendLine(sql_WHERE)
        S.AppendLine(sql_filter2)
        S.AppendLine("HAVING SUM((NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) - NVL(ICTSTAT2.WHSE_QTY_PICK,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) + NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0))) > 0")
        S.AppendLine("GROUP BY")
        S.AppendLine("ICTSTYL1.STYLE_CODE,")
        S.AppendLine("ICTSTAT2.COLOR_CODE,")
        S.AppendLine("ICTCOLR1.COLOR_DESC")
        ASCMAIN1.sql = S.ToString()
        Create_TDA(dst.Tables.Add, "ICTQUOTQ", "**", 0, False)
        With dst.Tables("ICTQUOTQ").Columns
            .Add("LAST_RCD_DATE")
        End With

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("T2.STYLE_CODE,")
        S.AppendLine("T2.COLOR_CODE,")
        S.AppendLine("C1.COLOR_DESC,")
        S.AppendLine("SUM(CASE")
        S.AppendLine(String.Format("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) <= {0}", dayBreaks))
        S.AppendLine("     THEN T2.TRAN_QTY")
        S.AppendLine("     ELSE 0")
        S.AppendLine("END) AS RECEIVED_01,")
        S.AppendLine("SUM(CASE")
        S.AppendLine(String.Format("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) > {0} AND ROUND((SYSDATE - T1.TRAN_DATE)) <= {1}", dayBreaks, dayBreaks * 2))
        S.AppendLine("     THEN T2.TRAN_QTY")
        S.AppendLine("     ELSE 0")
        S.AppendLine("END) AS RECEIVED_02,")
        S.AppendLine("SUM(CASE")
        S.AppendLine(String.Format("     WHEN ROUND((SYSDATE - T1.TRAN_DATE)) > {0}", dayBreaks * 2))
        S.AppendLine("     THEN T2.TRAN_QTY")
        S.AppendLine("     ELSE 0")
        S.AppendLine("END) AS RECEIVED_03")
        S.AppendLine("FROM ICTTRAN1 T1, ICTTRAN2 T2, ICTCOLR1 C1")
        S.AppendLine("WHERE T1.OPS_YYYYPP = T2.OPS_YYYYPP")
        S.AppendLine("AND T1.TRAN_TYPE = T2.TRAN_TYPE")
        S.AppendLine("AND T1.TRAN_NO = T2.TRAN_NO")
        S.AppendLine("AND T1.TRAN_TYPE = 'R'")
        S.AppendLine("AND T2.COLOR_CODE = C1.COLOR_CODE")
        S.AppendLine("GROUP BY")
        S.AppendLine("T2.STYLE_CODE,")
        S.AppendLine("T2.COLOR_CODE,")
        S.AppendLine("C1.COLOR_DESC")
        ASCMAIN1.sql = S.ToString()
        Create_TDA(dst.Tables.Add, "ICTQUOTD", "**", 0, False)

        Prepare_dst(True, sql_filter)
    End Sub

    Public Overrides Sub Build_Report_File_Pre_Ora2ADO(TT As String)
        MyBase.Build_Report_File_Pre_Ora2ADO(TT)
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()

        'Sticking a stupid row in the table to keep the standards from being an ass.
        Dim newASTSRPT1 As DataRow = dst.Tables("ASTSRPT1").NewRow
        newASTSRPT1.Item("G1") = "XX"
        dst.Tables("ASTSRPT1").Rows.Add(newASTSRPT1)

        For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select("", "STYLE_CODE, COLOR_CODE")
            Dim STYLE_CODE As String = rowICTQUOTQ.Item("STYLE_CODE").ToString()
            Dim COLOR_CODE As String = rowICTQUOTQ.Item("COLOR_CODE").ToString()
            Dim NOW_OH As Int64 = Val(rowICTQUOTQ.Item("NOW_OH").ToString & String.Empty)
            If NOW_OH > 0 Then
                ASCMAIN1.Progress("Now Calculating Aging For Style", String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
                Dim rowFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
                Dim rowICTQUOTD As DataRow = dst.Tables("ICTQUOTD").Select(rowFilter).FirstOrDefault
                Dim RECEIVED_TOTAL As Int64 = 0
                If Not IsNothing(rowICTQUOTD) Then
                    Dim RECEIVED_01 As Int64 = rowICTQUOTD.Item("RECEIVED_01")
                    rowICTQUOTQ.Item("RECEIVED_01") = RECEIVED_01
                    Dim RECEIVED_02 As Int64 = rowICTQUOTD.Item("RECEIVED_02")
                    rowICTQUOTQ.Item("RECEIVED_02") = RECEIVED_02
                    Dim RECEIVED_03 As Int64 = rowICTQUOTD.Item("RECEIVED_03")
                    rowICTQUOTQ.Item("RECEIVED_03") = RECEIVED_03
                    RECEIVED_TOTAL = RECEIVED_01 + RECEIVED_02 + RECEIVED_03
                    If NOW_OH <= RECEIVED_01 Then
                        rowICTQUOTQ.Item("AGED_01") = NOW_OH
                    Else
                        rowICTQUOTQ.Item("AGED_01") = RECEIVED_01
                        NOW_OH = NOW_OH - RECEIVED_01
                        If NOW_OH <= RECEIVED_02 Then
                            rowICTQUOTQ.Item("AGED_02") = NOW_OH
                        Else
                            rowICTQUOTQ.Item("AGED_02") = RECEIVED_02
                            NOW_OH = NOW_OH - RECEIVED_02
                            rowICTQUOTQ.Item("AGED_03") = NOW_OH
                        End If
                    End If
                End If
                If chkGreaterThan.Checked Then
                    If RECEIVED_TOTAL < numGreaterThan.Value Then
                        rowICTQUOTQ.Delete()
                    End If
                End If
            Else
                rowICTQUOTQ.Delete()
            End If
        Next
        dst.Tables("ICTQUOTQ").AcceptChanges()
    End Sub

    Public Overrides Sub Print_Report()

        'CR_params.Add("SUBT", txtDescription.Text & SUBT)
        'Stop 'This is an Excel output report
        'RPT = "ICRQUOTQ"
        'CR_params.Add("SUBT", txtDescription.Text & SUBT)
        'Generate_Report(RPT, , SUBT)
        ASCMAIN1.Progress("Creating Quick Quote Sheet", "")
        Dim XLS_FILENAME1 As String = MakeExcelWorkbook(False)
        Dim XLS_FILENAME2 As String = ""
        If chkBuyerOutput.Checked Then
            ASCMAIN1.Progress("Creating Buyer Quote Sheet", "")
            XLS_FILENAME2 = MakeExcelWorkbook(True)
            Show_Document(XLS_FILENAME1)
            Show_Document(XLS_FILENAME2)
        Else
            Show_Document(XLS_FILENAME1)
        End If
        ASCMAIN1.Progress("", "")
    End Sub

    Private Function MakeExcelWorkbook(ByVal MakeBuyerVersion As Boolean) As String
        Dim XLS_FILENAME As String = ""

        Dim StyleList As List(Of String) = filterDataOnOptions()

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        worksheet.Name = "Style Info"
        Create_Excel_WorkSheet(worksheet, StyleList, MakeBuyerVersion)

        If ASCMAIN1.Folders("Temp").EndsWith("\") Then
            If MakeBuyerVersion Then
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "ICRQUOTQ_2.XLSX"
            Else
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "ICRQUOTQ_1.XLSX"
            End If
        Else
            If MakeBuyerVersion Then
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "\" & "ICRQUOTQ_2.XLSX"
            Else
                XLS_FILENAME = ASCMAIN1.Folders("Temp") & "\" & "ICRQUOTQ_1.XLSX"
            End If
        End If
        Dim success As Boolean = False

        ASCMAIN1.Progress("Now Saving Workbook")

        Do Until success
            Try
                If IO.File.Exists(XLS_FILENAME) Then
                    IO.File.Delete(XLS_FILENAME)
                End If
                workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                success = True
            Catch ex As Exception

            End Try
        Loop
        Return XLS_FILENAME
    End Function

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length > 4 Then
            '    EMsg &= vbCr & "Maximum number of Sort Fields for this report is 4"
            'End If
        End If
    End Sub

    Public Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        If parms.Length > 0 Then
        End If

        EnforceConstraints(False)
        Fill_Records("ICTQUOTQ")
        If chkExcludeWIP.Checked Then
            REMOVE_WIP_STYLES()
        End If
        If chkShowLastRcd.Checked Then
            setLastRcdDate()
        End If

        Fill_Records("ICTQUOTD")
        EnforceConstraints(True)
    End Sub

#End Region

#Region "Form Methods"

#End Region

#Region "Custom Methods"
    Sub Create_Excel_WorkSheet(worksheet As SpreadsheetGear.IWorksheet,
                               ByVal StyleList As List(Of String),
                               ByVal MakeBuyerVersion As Boolean)

        Dim IMAGE_FOLDER As String = Replace(ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR"), "G:", "R:")
        If (ASCMAIN1.Running_in_VS) Then
            If Not IO.Directory.Exists(IMAGE_FOLDER) Then
                Stop 'You Need to Set up Image Folder.
            End If
        End If
        Dim interior As SpreadsheetGear.IInterior
        Dim range As SpreadsheetGear.IRange

        worksheet.Cells("A1:Z1").EntireColumn.Font.Size = 16

        Dim CX As Integer = 0
        Dim RX As Integer = 0

        Dim I As Integer = 0
        I += 4

        Dim CWC() As String = Split("A,B, C,D,E,F,G,H,I,J,K,L, M", ",")
        Dim CWS() As String = Split("1,1,40,6,6,6,6,6,6,6,6,6,20", ",")
        'If optPP.Value & "" = "4/5" Then
        '    CWS(2) = 45
        'End If
        CWS(2) = 45
        For CWCi As Integer = 0 To CWC.Length - 1
            worksheet.Cells(Trim(CWC(CWCi)) & "1").EntireColumn.ColumnWidth = Val(CWS(CWCi))
        Next

        worksheet.Cells(0, 0).EntireColumn.Hidden = True
        worksheet.Cells(0, 1).EntireColumn.Hidden = True

        Dim COL0 As Integer = 6 + 6

        Dim COL As Integer = COL0

        Dim ColVisible(4) As Boolean
        ColVisible(0) = True
        If MakeBuyerVersion Then
            ColVisible(1) = False
            ColVisible(2) = False
            ColVisible(3) = False
        Else
            ColVisible(1) = chkAGED_01.Checked
            ColVisible(2) = chkAGED_02.Checked
            ColVisible(3) = chkAGED_03.Checked
        End If
        ColVisible(4) = False

        For iCol As Integer = 1 To 4
            If ColVisible(iCol) Then
                COL += 1
                With worksheet.Cells(I - 1, COL)
                    .ColumnWidth = 15
                    .EntireColumn.NumberFormat = "#,##0"
                    .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With
            End If
        Next

        COL += 1
        With worksheet.Cells(I - 1, COL)
            .ColumnWidth = 15
            .EntireColumn.NumberFormat = "#,##0"
            .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            '.Value = "All"
            .HorizontalAlignment = SpreadsheetGear.HAlign.Right
        End With

        With worksheet.Cells(I, 0, I, COL)
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With

        Dim I0 As Integer = 0
        Dim IA As Integer = 0
        Dim RT(5) As String
        Dim ROW0 As Integer = I
        Dim style_count As Integer = 0
        Dim pages As Integer = 0

        For Each STYLE_CODE As String In StyleList
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
            ASCMAIN1.Progress("-", STYLE_CODE)

            I += 1
            I0 = I

            COL = COL0

            worksheet.Cells(I, COL - 1).Value = "Color"
            worksheet.Cells(I, COL - 0).Value = "Description"

            For iCol As Integer = 1 To 4
                If ColVisible(iCol) Then
                    COL += 1
                    With worksheet.Cells(I, COL)
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        If iCol = 0 Then
                            .Value = "COL 0"
                        End If

                        If iCol = 1 Then .Value = String.Format("0 - {0}", dayBreaks)
                        If iCol = 2 Then .Value = String.Format("{0} - {1}", dayBreaks + 1, dayBreaks * 2)
                        If iCol = 3 Then .Value = String.Format("Greater {0}", dayBreaks * 2)
                        If iCol = 4 Then .Value = "Never Used"

                    End With
                End If
            Next

            COL += 1
            With worksheet.Cells(I, COL)
                If MakeBuyerVersion Then
                    .Value = "Now"
                Else
                    .Value = "Total"
                End If
                .HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End With

            range = worksheet.Cells(I, COL0 - 1, I, COL)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.Gold

            If chkShowLastRcd.Checked And Not MakeBuyerVersion Then
                COL += 1
                With worksheet.Cells(I, COL)
                    .Value = "Last Recd"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .ColumnWidth = 18
                End With

                range = worksheet.Cells(I, COL0 - 1, I, COL)
                interior = range.Interior
                interior.Color = SpreadsheetGear.Colors.Gold
            End If

            I += 1

            Dim IMAGE_NAME As String = rowICTSTYL1.Item("IMAGE_NAME") & ""

            Dim imageFileStyle As String = IMAGE_FOLDER & "\" & IMAGE_NAME
            If Not IO.File.Exists(imageFileStyle) Then
                IMAGE_NAME = ""
            End If

            Dim ImageRows As Integer = 0
            Dim ImageRowsBig As Integer = 0

            If IMAGE_NAME <> "" _
                AndAlso My.Computer.FileSystem.FileExists(imageFileStyle) Then

                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution / 3
                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution / 3
                Finally
                    imageStyle.Dispose()
                End Try

                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

                Dim col_adj As Decimal = 0
                If heightStyle > widthStyle Then
                    col_adj = 0.3
                Else
                    col_adj = 0.05
                End If

                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0) + col_adj
                Dim topStyle As Double = windowInfoStyle.RowToPoints(I - 1) + 0.1 ' 1.5)

                ImageRows = windowInfoStyle.PointsToRow(heightStyle)
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
            End If

            CX = 1

            With worksheet.Cells(I - 1, 3)
                .Value = "'" & STYLE_CODE
                .Font.Color = SpreadsheetGear.Colors.Purple
                .Font.Size = 24
                .Font.Bold = True
            End With

            CX = 3

            worksheet.Cells(I + 2, CX).Value = "Case Qty"

            range = worksheet.Cells(I + 1, 3, I + 2, 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray

            range = worksheet.Cells(I + 1, 3 + 4, I + 2, 4 + 4)
            interior = range.Interior
            interior.Color = SpreadsheetGear.Colors.LightGray

            CX = 5
            worksheet.Cells(I, CX - 2).Value = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
            worksheet.Cells(I + 2, CX).Value = rowICTSTYL1.Item("CARTON_PACK_QTY")

            Dim SZMAX As Integer = 0
            Dim SZTOT As Integer = 0
            Dim T As String = ""
            Dim CI As Integer = 0
            Dim styleTotal As Int64 = 0
            For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select("STYLE_CODE = '" & STYLE_CODE & "'", "COLOR_CODE")
                CI += 1
                COL = COL0

                If chkShowCost.Checked And Not MakeBuyerVersion Then
                    worksheet.Cells(I + CI - 1, COL - 2).Value = getCostForStyleColor(rowICTQUOTQ.Item("STYLE_CODE") & String.Empty, rowICTQUOTQ.Item("COLOR_CODE") & String.Empty)
                    worksheet.Cells(I + CI - 1, COL - 2).NumberFormat = "###,##0.00"
                End If

                Dim COLOR_CODE As String = rowICTQUOTQ.Item("COLOR_CODE")
                Dim COLOR_DESC As String = rowICTQUOTQ.Item("COLOR_DESC")
                COLOR_DESC = GetAltColorCode(STYLE_CODE, COLOR_CODE, COLOR_DESC)
                worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & COLOR_CODE
                worksheet.Cells(I + CI - 1, COL - 0).Value = COLOR_DESC

                T = ""
                Dim VisCount As Integer = 0
                Dim rowTOTAL As Int64 = 0
                For iCOL As Integer = 1 To 4
                    If MakeBuyerVersion Then
                        If iCOL = 1 And chkAGED_01.Checked Then
                            rowTOTAL += Val(rowICTQUOTQ.Item("AGED_01").ToString & String.Empty)
                        End If
                        If iCOL = 2 And chkAGED_02.Checked Then
                            rowTOTAL += Val(rowICTQUOTQ.Item("AGED_02").ToString & String.Empty)
                        End If
                        If iCOL = 3 And chkAGED_03.Checked Then
                            rowTOTAL += Val(rowICTQUOTQ.Item("AGED_03").ToString & String.Empty)
                        End If
                    End If
                    If ColVisible(iCOL) Then
                        VisCount += 1
                        worksheet.Cells(I + CI - 1, COL + VisCount).Value = Val(rowICTQUOTQ.Item("AGED_0" & iCOL).ToString & String.Empty)
                        T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL + VisCount).Address, "$", "")
                    End If
                Next
                styleTotal += rowTOTAL
                COL += 1
                'This is where you can figure out the row total
                If MakeBuyerVersion Then
                    worksheet.Cells(I + CI - 1, COL + VisCount).Value = rowTOTAL
                Else
                    worksheet.Cells(I + CI - 1, COL + VisCount).Formula = "=" & Mid(T, 2)
                End If

                If chkShowLastRcd.Checked And Not MakeBuyerVersion Then
                    COL += 1
                    With worksheet.Cells(I + CI - 1, COL + VisCount)
                        .Value = rowICTQUOTQ.Item("LAST_RCD_DATE").ToString & String.Empty
                        .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    End With
                End If
            Next

            CI += 1
            COL = COL0

            worksheet.Cells(I - 1, COL - 1, I + CI - 1, COL - 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center

            worksheet.Cells(I + CI - 1, COL - 1).Value = "'" & "***"
            worksheet.Cells(I + CI - 1, COL - 0).Value = "'" & "Total"

            T = ""
            For iCOL As Integer = 1 To 4
                If ColVisible(iCOL) Then
                    COL += 1
                    If CI = 1 Then ' NO COLORS
                        worksheet.Cells(I + CI - 1, COL).Value = 0
                    Else
                        worksheet.Cells(I + CI - 1, COL).Formula = "=sum(" & Replace(worksheet.Cells(I + 1 - 1, COL).Address, "$", "") & ":" & Replace(worksheet.Cells(I + CI - 1 - 1, COL).Address, "$", "") & ")"
                    End If

                    RT(iCOL) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

                    T &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")
                    'COL += 1

                End If
            Next
            COL += 1

            If MakeBuyerVersion Then
                worksheet.Cells(I + CI - 1, COL).Value = styleTotal
            Else
                worksheet.Cells(I + CI - 1, COL).Formula = "=" & Mid(T, 2)
            End If

            RT(5) &= "+" & Replace(worksheet.Cells(I + CI - 1, COL).Address, "$", "")

            worksheet.Cells(I + CI - 1, COL0 - 1, I + CI - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

            With worksheet.Cells(I, COL0 - 1, I + CI - 1, COL)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
            End With

            I += ImageRowsBig

            Dim CJ As Integer = ImageRows ' - 1

            If CJ < 6 Then CJ = 6

            If CI > CJ Then
                I += CI
            Else
                I += CJ
            End If

            style_count += 1

            If (((I - 5) Mod 80) < ((I0 - 5) Mod 80)) Or (style_count >= 5) Or style_count >= 9 Then
                Dim R As SpreadsheetGear.IRange = worksheet.Cells(I0, 0).EntireRow
                worksheet.HPageBreaks.Add(R)
                style_count = 1
                pages += 1
            End If

            With worksheet.Cells(I0, 0, I + 1 - 1, COL)
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            End With
        Next

        I += 2
        COL = COL0

        'worksheet.Cells(I - 1, COL - 1).Value = "'" & "All"
        worksheet.Cells(I - 1, COL - 0).Value = "'" & "Totals"

        Dim GT = ""
        For iCOL As Integer = 1 To 4
            If ColVisible(iCOL) Then
                COL += 1
                worksheet.Cells(I - 1, COL).Formula = "=" & Mid(RT(iCOL), 2)

                GT &= "+" & Replace(worksheet.Cells(I - 1, COL).Address, "$", "")
                COL += 1
            End If
        Next
        COL += 1
        worksheet.Cells(I - 1, COL).Formula = "=" & Mid(GT, 2)


        worksheet.Cells(I - 1, COL0 - 1, I - 1, COL).Interior.Color = SpreadsheetGear.Colors.LightGray

        Dim H0 As Integer = 8 + 6

        worksheet.Cells(0, H0).Value = "Prep"
        worksheet.Cells(1, H0).Value = "By"
        worksheet.Cells(2, H0).Value = "XNo"

        worksheet.Cells(0, H0, 2, H0).Interior.Color = SpreadsheetGear.Colors.LightGray


        worksheet.Cells(0, H0 + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Left
        worksheet.Cells(0, H0 + 1).Value = Now
        worksheet.Cells(0, H0 + 1).NumberFormat = "MM/dd/yy" ' SpreadsheetGear.NumberFormatType.Date

        worksheet.Cells(1, H0 + 1).Value = ASCMAIN1.USER_ID
        worksheet.Cells(2, H0 + 1).Value = "'" & Mid(XNO, 5)

        With worksheet.Cells(0, H0, 2, H0 + 1)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        Dim H1 As Integer = 11

        worksheet.Cells(0, H1).Value = "Note"
        worksheet.Cells(1, H1).Value = "For"

        worksheet.Cells(0, H1, 2, H1).Interior.Color = SpreadsheetGear.Colors.LightGray

        'worksheet.Cells(0, H1 + 1).Value = Format(Absx1.dteFor("QUOTE_DATE").Value, "MM/dd/yyyy")
        worksheet.Cells(0, H1 + 1).NumberFormat = "MM/dd/yy"
        worksheet.Cells(0, H1 + 1).Value = txtNotes.Text
        worksheet.Cells(1, H1 + 1).Value = txtQuoteCUST_CODE.Text
        'worksheet.Cells(1, H1 + 2).Value = Absx1.txtFor("CUST_CODE").Text


        With worksheet.Cells(0, H1, 2, H1 + 2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Font.Color = SpreadsheetGear.Colors.Black
            .Font.Size = 10
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
        End With

        With worksheet.Cells(3, 3)
            .Font.Color = SpreadsheetGear.Colors.Purple
            .Font.Size = 20
            .Font.Bold = True
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            '.Value = Absx1.txtFor("QUOTE_DESC").Text
        End With

        With worksheet.PageSetup
            .TopMargin = 0.25
            .LeftMargin = 0.25
            .RightMargin = 0.25
            .BottomMargin = 0.25
            .FitToPagesWide = 1
            .FitToPagesTall = Nothing
            .PrintTitleRows = "A1:S5"

            .CenterFooter = "&P"
        End With

        'Dim imageFile As String = ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.DBS_COMPANY & "_EMAIL.PNG"

        'worksheet.Cells("C1").EntireColumn.Hidden = True
    End Sub

    Private Function filterDataOnOptions() As List(Of String)
        Dim RetVal As New List(Of String)
        Dim RECEIVED_TOTAL As Int64 = 0
        For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select("", "STYLE_CODE, COLOR_CODE")
            RECEIVED_TOTAL = 0
            'If rowICTQUOTQ.Item("STYLE_CODE").ToString = "502006XIZ" Then Stop
            Dim IncludeColor As Boolean = False
            If chkAGED_01.Checked And Val(rowICTQUOTQ.Item("AGED_01").ToString & String.Empty) > 0 Then
                IncludeColor = True
                RECEIVED_TOTAL = RECEIVED_TOTAL + Val(rowICTQUOTQ.Item("AGED_01").ToString & String.Empty)
            End If
            If chkAGED_02.Checked And Val(rowICTQUOTQ.Item("AGED_02").ToString & String.Empty) > 0 Then
                IncludeColor = True
                RECEIVED_TOTAL = RECEIVED_TOTAL + Val(rowICTQUOTQ.Item("AGED_02").ToString & String.Empty)
            End If
            If chkAGED_03.Checked And Val(rowICTQUOTQ.Item("AGED_03").ToString & String.Empty) > 0 Then
                IncludeColor = True
                RECEIVED_TOTAL = RECEIVED_TOTAL + Val(rowICTQUOTQ.Item("AGED_03").ToString & String.Empty)
            End If
            If chkGreaterThan.Checked Then
                If RECEIVED_TOTAL < numGreaterThan.Value Then
                    IncludeColor = False
                End If
            End If
            If IncludeColor Then
                If Not RetVal.Contains(rowICTQUOTQ.Item("STYLE_CODE").ToString & String.Empty) Then
                    RetVal.Add(rowICTQUOTQ.Item("STYLE_CODE").ToString & String.Empty)
                End If
            Else
                rowICTQUOTQ.Delete()
            End If
        Next
        Return RetVal
    End Function

    Private Function GetAltColorCode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal COLOR_DESC_ORIG As String) As String
        Dim RetVal As String = COLOR_DESC_ORIG
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYL1.Item("SIZE_SCALE") & String.Empty
        Dim MAX_LENGTH As Integer = 60
        Dim I As Integer = InStr(SIZE_SCALE, COLOR_CODE)
        If I <> 0 Then
            Dim S As String = Trim(Mid(SIZE_SCALE, I + 3))
            Dim J As Integer = InStr(Mid(S & "  ", 1, MAX_LENGTH), "  ")
            Dim K As Integer = InStr(Mid(S & vbCrLf, 1, MAX_LENGTH), vbCrLf)
            If J = 0 And K = 0 Then
                J = InStr(Mid(S & " ", 1, MAX_LENGTH), " ")
            End If
            If J = 0 Or J > K Then J = K
            Dim SC As String = ""
            If J <> 0 Then
                SC = Mid(S, 1, J)
                SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                For C As Integer = 1 To SC.Length - 1
                    If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                        Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                    End If
                Next
                If Trim(SC) <> "" Then
                    If SC.Length > 35 Then
                        RetVal = SC.Substring(0, 34)
                    Else
                        RetVal = SC
                    End If

                End If
            End If
        End If
        If RetVal = COLOR_DESC_ORIG Then
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT NVL(STYLE_COLOR_DESC,'') STYLE_COLOR_DESC")
            SQLS.AppendLine("FROM ICTSTYC1")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim COLOR_DESC_MF As String = ASCDATA1.GetDataValue
            If COLOR_DESC_MF.Length > 35 Then
                COLOR_DESC_MF = COLOR_DESC_MF.Substring(0, 35)
            End If
            If COLOR_DESC_MF.Length > 0 Then
                RetVal = COLOR_DESC_MF
            End If
        End If
        Return RetVal
    End Function

    Private Function getCostForStyleColor(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As Double
        Dim Retval As Double = 0
        ASCMAIN1.sql = "Select STYLE_COST from (" & vbCrLf _
                            & "Select STYLE_COST from ICTCOSTA " & vbCrLf _
                            & "where (STYLE_CODE, COLOR_CODE) in (" & vbCrLf _
                            & "Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                            & " from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'" _
                            & " and WHSE_QTY_ON_HAND > 0)" & vbCrLf _
                            & " order by OPS_YYYYPP DESC) where ROWNUM < 2"
        Dim STYLE_COST As Decimal = Val(ASCDATA1.GetDataValue)

        If STYLE_COST = 0 Then
            ASCMAIN1.sql = "Select NVL(PO_COST_LANDED,PO_COST) STYLE_COST" & vbCrLf _
                                & " from (" & vbCrLf _
                                & " Select POTSHIP3.PO_SHIPMENT_NO, POTORDR2.PO_ORDER_NO, " & vbCrLf _
                                & " POTORDR2.PO_COST, POTSHIP3.PO_COST_LANDED, POTSHIP2.PO_DATE_RECEIVED, POTSHIP1.PO_DATE_SHIPPED" & vbCrLf _
                                & " from POTORDR2,POTSHIP3,POTSHIP2,POTSHIP1" & vbCrLf _
                                & " where POTORDR2.STYLE_CODE = '" & STYLE_CODE & "' and POTORDR2.COLOR_CODE = '" & COLOR_CODE & "'" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO" & vbCrLf _
                                & "   and POTSHIP3.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & "   and POTSHIP2.PO_SHIPMENT_LNO (+) = POTSHIP3.PO_SHIPMENT_LNO" & vbCrLf _
                                & "   and POTSHIP1.PO_SHIPMENT_NO (+) = POTSHIP3.PO_SHIPMENT_NO" & vbCrLf _
                                & " order by POTSHIP3.PO_SHIPMENT_NO DESC, POTORDR2.PO_ORDER_NO DESC" & vbCrLf _
                                & ") where ROWNUM <2"
            STYLE_COST = Val(ASCDATA1.GetDataValue)
        End If

        If STYLE_COST <> 0 Then
            Retval = Math.Round(STYLE_COST, 2)
        End If

        Return Retval
    End Function

    Private Sub chkGreaterThan_CheckedChanged(sender As Object, e As EventArgs) Handles chkGreaterThan.CheckedChanged

    End Sub

    Private Sub chkUse180Days_CheckedChanged(sender As Object, e As EventArgs) Handles chkUse180Days.CheckedChanged
        setDayBreaks()
    End Sub

    Private Sub REMOVE_WIP_STYLES()
        For Each rowICTQUOTQ As DataRow In dst.Tables("ICTQUOTQ").Select()
            Dim STYLE_CODE As String = rowICTQUOTQ.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = rowICTQUOTQ.Item("COLOR_CODE") & String.Empty
            Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
            SQLS.AppendLine("SELECT SUM(NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)) AS TOTAL_WIP")
            SQLS.AppendLine("FROM ICTSTAT2")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
            SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = SQLS.ToString()
            Dim TOTAL_WIP As Int64 = Val(ASCDATA1.GetDataValue)
            If TOTAL_WIP > 0 Then
                rowICTQUOTQ.Delete()
            End If
        Next
        dst.Tables("ICTQUOTQ").AcceptChanges()
    End Sub

    Private Sub setDayBreaks()
        If chkUse180Days.Checked Then
            dayBreaks = 180
        Else
            dayBreaks = 120
        End If
        grpAGED_01.Text = String.Format("0 to {0} Days", dayBreaks)
        grpAGED_02.Text = String.Format("{0} to {1} Days", dayBreaks + 1, dayBreaks * 2)
        grpAGED_03.Text = String.Format("Greater that {0} Days", (dayBreaks * 2) + 1)
    End Sub

    Private Sub setLastRcdDate()
        For Each row As DataRow In dst.Tables("ICTQUOTQ").Select("", "STYLE_CODE, COLOR_CODE")
            Dim S As New System.Text.StringBuilder With {.Length = 0}
            Dim STYLE_CODE As String = row.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = row.Item("COLOR_CODE").ToString & String.Empty
            ASCMAIN1.Progress("Calculating Last Rcd Date", String.Format("{0}-{1}", STYLE_CODE, COLOR_CODE))
            S.AppendLine("SELECT NVL(TO_CHAR(MAX(POTSHIP2.PO_DATE_RECEIVED),'MM/DD/YY'),'') PO_DATE_RECEIVED")
            S.AppendLine("FROM POTORDR2, POTSHIP3, POTSHIP2")
            S.AppendLine("WHERE POTSHIP3.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO")
            S.AppendLine("AND POTSHIP3.PO_SHIPMENT_LNO = POTSHIP2.PO_SHIPMENT_LNO")
            S.AppendLine("AND POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO")
            S.AppendLine("AND POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO")
            S.AppendLine(String.Format("AND POTORDR2.STYLE_CODE = '{0}'", STYLE_CODE))
            S.AppendLine(String.Format("AND POTORDR2.COLOR_CODE = '{0}'", COLOR_CODE))
            ASCMAIN1.sql = S.ToString()
            Dim LAST_RCD_DATE As String = ASCDATA1.GetDataValue

            If IsDate(LAST_RCD_DATE) Then
                LAST_RCD_DATE = Format(CDate(LAST_RCD_DATE), "MM/dd/yy")
            Else
                S.Length = 0
                S.AppendLine("SELECT NVL(WHSE_QTY_TRAN,0) AS IN_TRAN")
                S.AppendLine("FROM ICTSTAT2")
                S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                ASCMAIN1.sql = S.ToString()
                Dim IN_TRAN As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                If IN_TRAN > 0 Then
                    LAST_RCD_DATE = "In-Tran"
                Else
                    S.Length = 0
                    S.AppendLine("SELECT NVL(WHSE_QTY_ON_ORDER,0) AS IN_WIP")
                    S.AppendLine("FROM ICTSTAT2")
                    S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    S.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
                    ASCMAIN1.sql = S.ToString()
                    Dim IN_WIP As Int64 = Val(ASCDATA1.GetDataValue & String.Empty)
                    If IN_WIP > 0 Then
                        LAST_RCD_DATE = "In-WIP"
                    Else
                        LAST_RCD_DATE = ""
                    End If
                End If
            End If
            row.Item("LAST_RCD_DATE") = LAST_RCD_DATE
        Next
        ASCMAIN1.Progress("", "")
    End Sub

#End Region

End Class