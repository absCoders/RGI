Public Class WHFIREC1
    Dim WHSE_CODE As String
    Dim LP_CODE As String
    Dim TRANS_SEQ As String
    Dim TRNDTE As String

    Dim SNAPHDR As String
    Dim SNAPDTL As String

    Dim rowICTWHSE1 As DataRow
    Dim rowWHTTPLP1 As DataRow
    Dim ADS_UNPOSTED_ITEM As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            BuildSNAPHDR(True)
            BuildSNAPDTL(True)

            ASCMAIN1.sql = "SELECT * FROM " & SNAPHDR
            Create_TDA(.Tables.Add, "SNAPHDR", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTWHSE1 where LP_CODE is Not Null"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**", 0, False)

            ASCMAIN1.sql = "Select X.*, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & ", ICTSTYL1.SALES_DIVISION_CODE, ICTSTYL1.CUST_CODE, ICTSTYL1.SUB_BODY_CODE, DECODE(ICTSTYL1.CUST_CODE,NULL,'STK','NON') STK_NON, ICTSTYL1.FABRIC_CODE, ICTSTYL1.SEASON_CODE" & vbCrLf _
                & ", ICTCOLR1.COLOR_DESC from (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, SUM(NVL(OUR_QTY,0)) OUR_QTY, SUM(NVL(LP_QTY,0)) LP_QTY from (" & vbCrLf _
                & "(Select STYLE_CODE, COLOR_CODE, 0 OUR_QTY, SUM(LP_QTY) LP_QTY" & vbCrLf _
                & " from (" & vbCrLf _
                & " Select " & vbCrLf _
                & "   DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.STYLE_CODE, WHTSTYLX.STYLE_CODE) STYLE_CODE" & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.COLOR_CODE, WHTSTYLX.COLOR_CODE) COLOR_CODE" & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', TRUNC(100 * WHTPPKM2.PPK_QTY * SNAPDTL.QTY / NVL(WHTPPKM1.PPK_QTY_TOTAL,1))/100, SNAPDTL.QTY) LP_QTY" & vbCrLf _
                & " from " & SNAPDTL & " SNAPDTL,  WHTSTYLX, WHTPPKM2, WHTPPKM1" & vbCrLf _
                & " where WHTSTYLX.LP_CODE   (+) = SNAPDTL.LP_CODE" & vbCrLf _
                & "   and WHTSTYLX.ITEM_CODE (+) = SNAPDTL.ITEM_CODE " & vbCrLf _
                & "   and WHTPPKM2.PPK_CODE  (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
                & "   and WHTPPKM1.PPK_CODE  (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
                & " ) group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ") union (" & vbCrLf _
                & " Select STYLE_CODE, COLOR_CODE" & vbCrLf _
                & " , SUM(WHSE_QTY_ON_HAND) OUR_QTY, 0 LP_QTY" & vbCrLf _
                & " from ICTSTAT2" & vbCrLf _
                & " where WHSE_CODE = :PARM1 and NVL(WHSE_QTY_ON_HAND,0) <> 0" & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE)" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE) X, ICTSTYL1, ICTCOLR1" & vbCrLf _
                & " where X.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & "   and X.COLOR_CODE = ICTCOLR1.COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTIREC1", "**", 0, False, "V", 2)
            With dst.Tables("WHTIREC1")
                .Columns.Add("SHP_QTY", GetType(System.Int64))
                .Columns.Add("BNS_QTY", GetType(System.Int64))
                .Columns.Add("RTN_QTY", GetType(System.Int64))
                .Columns.Add("REC_QTY", GetType(System.Int64))
                .Columns.Add("ADJ_QTY", GetType(System.Int64))
                .Columns.Add("VARIANCE", GetType(System.Int64), "(ISNULL(LP_QTY,0) + ISNULL(SHP_QTY,0) - ISNULL(BNS_QTY,0) - ISNULL(RTN_QTY,0) - ISNULL(REC_QTY,0) - ISNULL(ADJ_QTY,0)) - ISNULL(OUR_QTY,0)")
            End With

            ASCMAIN1.sql = "SELECT " _
                & " SNAPDTL.ITEM_CODE ITEM_CODE," _
                & " SNAPDTL.QTY LP_QTY" _
                & " FROM " & SNAPDTL & " SNAPDTL WHERE" _
                & " SNAPDTL.ITEM_CODE NOT IN (" _
                & " SELECT ITEM_CODE FROM" _
                & " WHTSTYLX" _
                & " WHERE WHTSTYLX.LP_CODE  = :PARM1)"
            Create_TDA(.Tables.Add, "WHTBADI1", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT SNAPDTL.ITEM_CODE ITEM_CODE," _
                & " WHTSTYLX.PPK_CODE, SNAPDTL.QTY LP_QTY" _
                & " FROM " & SNAPDTL & " SNAPDTL,  WHTSTYLX" _
                & " WHERE WHTSTYLX.LP_CODE  (+) = SNAPDTL.LP_CODE" _
                & "  AND WHTSTYLX.ITEM_CODE (+) = SNAPDTL.ITEM_CODE " _
                & "  AND WHTSTYLX.ITEM_TYPE (+) = 'P'" _
                & "  AND WHTSTYLX.PPK_CODE NOT IN (" _
                & " SELECT PPK_CODE FROM WHTPPKM2)"
            Create_TDA(.Tables.Add, "WHTBADP1", "**", 0, False)

            ASCMAIN1.sql = " Select STYLE_CODE, COLOR_CODE, SUM(ADJ_QTY) ADJ_QTY FROM (" & vbCrLf _
                & " SELECT WHTIADJ1.ITEM_CODE " & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.STYLE_CODE, WHTSTYLX.STYLE_CODE) STYLE_CODE" & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.COLOR_CODE, WHTSTYLX.COLOR_CODE) COLOR_CODE" & vbCrLf _
                & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.PPK_QTY * WHTIADJ1.ADJQTY, WHTIADJ1.ADJQTY) ADJ_QTY" & vbCrLf _
                & " FROM  WHTIADJ1, WHTSTYLX, WHTPPKM2" & vbCrLf _
                & " WHERE WHTIADJ1.ABS_STATUS <> 'A' and WHTIADJ1.ABS_STATUS <> 'D'" & vbCrLf _
                & "   AND WHTSTYLX.LP_CODE       = WHTIADJ1.LP_CODE" & vbCrLf _
                & "   AND WHTSTYLX.ITEM_CODE (+) = WHTIADJ1.ITEM_CODE " & vbCrLf _
                & "   AND WHTPPKM2.PPK_CODE  (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
                & "   AND WHTIADJ1.WHSE_CODE     = :PARM1" & vbCrLf _
                & "   AND WHTIADJ1.LP_CODE       = :PARM2" & vbCrLf _
                & " ) GROUP BY STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTIADJ1", "**", 0, False, "VV", 2)
        End With

        Fill_Records("ICTWHSE1")

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSE1")
        grdWHTIREC1.DataSource = DVWs("WHTIREC1")
        grdWHTBADI1.DataSource = dst.Tables("WHTBADI1")
        grdWHTBADP1.DataSource = dst.Tables("WHTBADP1")
        grdSNAPHDR.DataSource = dst.Tables("SNAPHDR")

        Create_Summary(grdWHTIREC1, "STYLE_CODE", "Count")
        Create_Summary(grdWHTIREC1, New String() {"OUR_QTY", "LP_QTY", "SHP_QTY", "BNS_QTY", "RTN_QTY", "REC_QTY", "ADJ_QTY", "VARIANCE"})

        Create_Summary(grdWHTBADI1, "ITEM_CODE", "Count")
        Create_Summary(grdWHTBADI1, "LP_QTY", "Sum")

        Create_Summary(grdWHTBADP1, "ITEM_CODE", "Count")
        Create_Summary(grdWHTBADP1, "LP_QTY", "Sum")

        Create_Summary(grdSNAPHDR, "TRANS_SEQ", "Count")

        Show_Style_Columns(False)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                            EMsg &= vbCr & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " is not set up as a 3PL"
                        Else
                            rowWHTTPLP1 = LookUp("WHTTPLP1", rowICTWHSE1.Item("LP_CODE"))
                            If rowWHTTPLP1 Is Nothing Then
                                EMsg &= vbCrLf & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " Does NOT have a valid value specified for its 3PL"
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    LP_CODE = rowICTWHSE1.Item("LP_CODE")
                    WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
                    BuildSNAPHDR(False)
                    Fill_Records("SNAPHDR")
                    If dst.Tables("SNAPHDR").Rows.Count = 0 Then
                        EMsg &= "No Inventory Snapshots data available"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("WHTSPCK1", grdICTWHSEX.ActiveRow.Cells("WHSE_CODE").Value) Then Exit Sub
                End If

            Case "Update"
                If dst.Tables("SNAPHDR").Select("STATUS = '0'").Length = 0 Then
                    EMsg &= "No Inventory Snapshots data to Reconcile"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Clear_Record()
                Mode_Settings(False)
            Case "Load Phys Counts"
                Create_WHTLOCP1()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load Phys Counts").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load Phys Counts").Visible = IIf(ASCMAIN1.USER_SECURITY_CODEs.Contains("SY"), True, False)
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdICTWHSEX.Visible = Not ScreenMode
        tabREC1.Visible = ScreenMode
        UltraExplorerBar1.Groups("View Options").Visible = ScreenMode
        grpDocInfo.Visible = ScreenMode
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTIREC1", "WHTBADI1", "WHTBADP1", "SNAPHDR", "WHTIADJ1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        LP_CODE = ""
        WHSE_CODE = ""
        TRANS_SEQ = ""
        TRNDTE = ""

        txtLP_CODE.Text = ""
        txtWHSE_CODE.Text = ""
        txtTRANS_SEQ.Text = ""
        txtTRNDTE.Text = ""

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        TRANS_SEQ = dst.Tables("SNAPHDR").Compute("MAX(TRANS_SEQ)", "")
        Dim rowSNAPHDR As DataRow = dst.Tables("SNAPHDR").Select("TRANS_SEQ = '" & TRANS_SEQ & "'")(0)
        TRNDTE = rowSNAPHDR.Item("TRNDTE")

        BuildSNAPDTL(False, TRANS_SEQ, WHSE_CODE, LP_CODE)
        Sort_grdColumns(grdSNAPHDR, "TRNDTE".ToLower)

        Save_Header_Fields(UltraGroupBox1)


        Fill_Records("WHTIREC1", WHSE_CODE)
        Sort_grdColumns(grdWHTIREC1, "STYLE_CODE,COLOR_CODE")
        Fill_Records("WHTBADI1", LP_CODE)
        Fill_Records("WHTBADP1")

        ''ADJUSTMENT TABLES

        ''Dim beginTime As Date = Now
        ''BeginTrans("Retrieving open adjustments from LP ...")
        'TAC.WHCMAIN1.UpdateADSAndImport()
        ''Dim difference As TimeSpan = Now.Subtract(beginTime)
        ''Dim elapsedTime As String = Format(difference.TotalHours, "00") & ":" & Format(difference.TotalMinutes, "00") & ":" & Format(difference.TotalSeconds, "00")
        ''CommitTrans("Import Complete. Elapsed Time = " & elapsedTime)

        'ASCMAIN1.Progress("Now Loading Data ...")

 

        ASCMAIN1.sql = "Select ITEM_CODE, 'X' TRAN_TYPE, 0 TRAN_QTY from WHTSTYLX where ROWNUM < 1"
        ADS_UNPOSTED_ITEM = ASCMAIN1.Temp_Table


        ' Shipped not Billed

        ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1" & vbCrLf _
            & " where SHIP_STATUS = 'P' and SHIP_BOL_NO in " & vbCrLf _
            & " (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL@ADSIIS " & vbCrLf _
            & " where WHSE_CODE = '" & WHSE_CODE & "' and LP_STATUS IN ('V','2'));"

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & ASCMAIN1.sql & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Insert into " & ADS_UNPOSTED_ITEM & " (ITEM_CODE,TRAN_TYPE,TRAN_QTY)" & vbCrLf _
            & "    Select SOTPICK2_3PL.ITEM_CODE, 'S' TRAN_TYPE, SUM (SOTPICK2_3PL.PICK_QTY_CONF) TRAN_QTY" & vbCrLf _
            & "     from ADS.SOTPICK2_3PL@ADSIIS,ADS.SOTPICK1_3PL@ADSIIS" & vbCrLf _
            & "      where SOTPICK2_3PL.PICK_NO = SOTPICK1_3PL.PICK_NO" & vbCrLf _
            & "        and SOTPICK1_3PL.SHIP_BOL_NO = R1.SHIP_BOL_NO group by ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ' Billed not Shipped

        Add_Unposted_ADS("SHP_QTY", "S")

        ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1" & vbCrLf _
            & " where SHIP_STATUS = 'F' and SHIP_BOL_NO in " & vbCrLf _
            & " (Select SHIP_BOL_NO from ADS.SOTSHIP1_3PL@ADSIIS " & vbCrLf _
            & " where WHSE_CODE = '" & WHSE_CODE & "' and LP_STATUS = '1')"
        Dim ADS_BNS As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SUM (SOTPICK2.PICK_QTY_CONF) BNS_QTY" & vbCrLf _
            & "     from SOTPICK2,SOTPICK1,SOTORDR2" & vbCrLf _
            & "      where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "        and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "        and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "        and SOTPICK1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & ADS_BNS & ")" & vbCrLf _
            & " GROUP BY SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        Add_Unposted(ASCDATA1.GetDataTable, "BNS_QTY")


        ' Returns & Receipts

        ASCMAIN1.sql = "   Insert into " & ADS_UNPOSTED_ITEM & " (ITEM_CODE,TRAN_TYPE,TRAN_QTY)" & vbCrLf _
            & "    Select RCPTDTL.ITEM_CODE, DECODE(RCPTHDR.INVTYP,'R','C','R') TRAN_TYPE, SUM (RCPTDTL.RCVQTY) TRAN_QTY" & vbCrLf _
            & "     from ADS.RCPTDTL@ADSIIS,ADS.RCPTHDR@ADSIIS" & vbCrLf _
            & "      where RCPTHDR.TRANS_SEQ = RCPTDTL.TRANS_SEQ" & vbCrLf _
            & "        and RCPTHDR.STATUS in ('0','V')" & vbCrLf _
            & "        and RCPTHDR.LP_CODE = RCPTDTL.LP_CODE " & vbCrLf _
            & "        and RCPTHDR.WHSE_CODE = RCPTDTL.WHSE_CODE " & vbCrLf _
            & "        and RCPTHDR.WHSE_CODE = '" & WHSE_CODE & "'" & vbCrLf _
            & "        group by RCPTDTL.ITEM_CODE, DECODE(RCPTHDR.INVTYP,'R','C','R')"
        ASCDATA1.ExecuteSQL()

        Add_Unposted_ADS("REC_QTY", "R")
        Add_Unposted_ADS("RTN_QTY", "C")


        ' Adjustments

        Fill_Records("WHTIADJ1", New Object() {WHSE_CODE, LP_CODE})
        Add_Unposted(dst.Tables("WHTIADJ1"), "ADJ_QTY")

        'ASCMAIN1.sql = "" _
        '    & "Begin" & vbCrLf _
        '    & " Declare Cursor C1 is " & ASCMAIN1.sql & vbCrLf _
        '    & " Begin" & vbCrLf _
        '    & "  For R1 in C1 Loop" & vbCrLf _
        '    & "   Insert into " & ADS_UNPOSTED_ITEM & " (ITEM_CODE,TRAN_TYPE,TRAN_QTY)" & vbCrLf _
        '    & "    Select INVADJ.ITEM_CODE, 'A' TRAN_TYPE, SUM (INVADJ.ADJQTY) TRAN_QTY" & vbCrLf _
        '    & "     from ADS.INVADJ@ADSIIS" & vbCrLf _
        '    & "      where INVADJ.STATUS in ('0','V')" _
        '    & "        and INVADJ.WHSE_CODE = '" & WHSE_CODE & "'" _
        '    & "        group by ITEM_CODE;" & vbCrLf _
        '    & "  End Loop;" & vbCrLf _
        '    & " End;" & vbCrLf _
        '    & "End;"
        Add_Unposted_ADS("ADJ_QTY", "A")


        If dst.Tables("WHTBADI1").Rows.Count = 0 Then
            grdWHTBADI1.Text = "There are no bad Item Codes"
        End If

        If dst.Tables("WHTBADP1").Rows.Count = 0 Then
            grdWHTBADP1.Text = "There are no bad Pre Pack Numbers"
        End If

        txtLP_CODE.Text = LP_CODE & ""
        txtWHSE_CODE.Text = WHSE_CODE & ""
        txtTRANS_SEQ.Text = TRANS_SEQ & ""
        txtTRNDTE.Text = FormatDateTime(TRNDTE, DateFormat.ShortDate) & ""

        setRowFilter()
        tabREC1Changed()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Create_WHTLOCP1()
        ASCMAIN1.sql = "Truncate Table WHTLOCP1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        'why do we have style codes with null value
        ASCMAIN1.sql = "Insert into WHTLOCP1 (Select '" & WHSE_CODE & "' as WHSE_CODE, 'VIRTL' as LOCATION_CODE, '0000000000' as BAR_CODE, " _
               & " STYLE_CODE, COLOR_CODE, SUM(LP_QTY) LOCATION_QTY" & vbCrLf _
               & " from (" & vbCrLf _
               & " Select " & vbCrLf _
               & "   DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.STYLE_CODE, WHTSTYLX.STYLE_CODE) STYLE_CODE" & vbCrLf _
               & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', WHTPPKM2.COLOR_CODE, WHTSTYLX.COLOR_CODE) COLOR_CODE" & vbCrLf _
               & " , DECODE(WHTSTYLX.ITEM_TYPE, 'P', TRUNC(100 * WHTPPKM2.PPK_QTY * SNAPDTL.QTY / NVL(WHTPPKM1.PPK_QTY_TOTAL,1))/100, SNAPDTL.QTY) LP_QTY" & vbCrLf _
               & " from " & SNAPDTL & " SNAPDTL,  WHTSTYLX, WHTPPKM2, WHTPPKM1" & vbCrLf _
               & " where WHTSTYLX.LP_CODE   (+) = SNAPDTL.LP_CODE" & vbCrLf _
               & "   and WHTSTYLX.ITEM_CODE (+) = SNAPDTL.ITEM_CODE " & vbCrLf _
               & "   and WHTPPKM2.PPK_CODE  (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
               & "   and WHTPPKM1.PPK_CODE  (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
               & " ) Where STYLE_CODE is not Null group by STYLE_CODE, COLOR_CODE)" & vbCrLf
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
    End Sub

    Sub Update_Record()
        Dim TRANS_SEQ As String = String.Empty
        For Each row As DataRow In dst.Tables("SNAPHDR").Select("STATUS = 0")
            TRANS_SEQ = row.Item("TRANS_SEQ")
            ASCDATA1.ExecuteSQL("UPDATE ADS.SNAPHDR@ADSIIS SET STATUS = '1' where WHSE_CODE = '" & WHSE_CODE & "' and STATUS = '0' and TRANS_SEQ = '" & TRANS_SEQ & "'")
        Next
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTIREC1, "SSSBS", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry", "Show Style Master Fields")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show Style Master Fields"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Style_Columns(tlb_sbt.Checked)


        End Select
    End Sub

#End Region

    Sub Show_Style_Columns(show_columns As Boolean)
        For Each COLUMN_NAME As String In New String() {"SALES_DIVISION_CODE", "CUST_CODE", "SUB_BODY_CODE", "STK_NON", "FABRIC_CODE", "SEASON_CODE"}
            grdWHTIREC1.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not show_columns
        Next
    End Sub
#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub
#End Region


    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

    Private Sub setRowFilter()
        Dim rowFilter As String = String.Empty
        If Absx1.chkFor("VARIANCES_ONLY").Checked Then
            rowFilter &= "AND VARIANCE <> 0"
        End If
        If Absx1.chkFor("ADJS_ONLY").Checked Then
            rowFilter &= "AND ADJ_QTY <> 0"
        End If
        DVWs("WHTIREC1").RowFilter = Mid(rowFilter, 5)

    End Sub

    Private Sub AbsCheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles AbsCheckBox1.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        setRowFilter()
    End Sub

    Private Sub AbsCheckBox2_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles AbsCheckBox2.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        setRowFilter()
    End Sub

    Private Sub tabREC1_ActiveTabChanged(sender As Object, e As Infragistics.Win.UltraWinTabControl.ActiveTabChangedEventArgs) Handles tabREC1.ActiveTabChanged
        tabREC1Changed()
    End Sub

    Private Sub tabREC1Changed()
        UltraExplorerBar1.Groups(1).Visible = ScreenMode AndAlso (tabREC1.ActiveTab.Index = 0)
    End Sub

    Private Sub BuildSNAPHDR(initialize As Boolean)

        ASCMAIN1.Progress("Fetching data from LP...")

        Dim ADSTableLocation As String = "ADS.SNAPHDR@ADSIIS"
        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.DBS_SERVER = "" Then
            ADSTableLocation = "SNAPHDR"
        End If

        If initialize Then
            ASCMAIN1.sql = "Select * from " & ADSTableLocation & " where ROWNUM <1"
            SNAPHDR = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SNAPHDR)
            ASCMAIN1.sql = "INSERT INTO " & SNAPHDR & " (" & vbCrLf _
                & " SELECT * FROM  " & ADSTableLocation & vbCrLf _
                & " WHERE STATUS = '0'" & vbCrLf _
                & " AND LP_CODE = '" & LP_CODE & "'" & vbCrLf _
                & " AND WHSE_CODE  = '" & WHSE_CODE & "')"
            ASCDATA1.ExecuteSQL()
        End If
        ASCMAIN1.Progress("")
    End Sub

    Private Sub BuildSNAPDTL(initialize As Boolean, _
                                Optional TRANS_SEQ As String = "", _
                                Optional WHSE_CODE As String = "", _
                                Optional LP_CODE As String = "")

        ASCMAIN1.Progress("Fetching data from LP...")

        Dim ADSTableLocation As String = "ADS.SNAPDTL@ADSIIS"
        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.DBS_SERVER = "" Then
            ADSTableLocation = "SNAPDTL"
        End If

        If initialize Then
            ASCMAIN1.sql = "Select * from " & ADSTableLocation & " where ROWNUM <1"
            SNAPDTL = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SNAPDTL)
            ASCMAIN1.sql = "INSERT INTO " & SNAPDTL & " (" & vbCrLf _
                & " SELECT * FROM  " & ADSTableLocation & vbCrLf _
                & " WHERE TRANS_SEQ = '" & TRANS_SEQ & "'" & vbCrLf _
                & " AND LP_CODE = '" & LP_CODE & "'" & vbCrLf _
                & " AND WHSE_CODE  = '" & WHSE_CODE & "')"
            ASCDATA1.ExecuteSQL()
        End If
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdWHTIREC1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTIREC1.InitializeRow
        If Val(e.Row.Cells("VARIANCE").Value & "") < 0 Then
            e.Row.Cells("VARIANCE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("VARIANCE").Appearance.ForeColor = Drawing.Color.Black
        End If
    End Sub

    Sub Add_Unposted(tbl As DataTable, COLUMN_NAME As String)
        For Each row As DataRow In tbl.Select()
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim keys As Object() = {STYLE_CODE, COLOR_CODE}
            Dim rowWHTIREC1 As DataRow = dst.Tables("WHTIREC1").Rows.Find(keys)
            If rowWHTIREC1 Is Nothing Then
                rowWHTIREC1 = dst.Tables("WHTIREC1").NewRow
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                rowWHTIREC1.Item("STYLE_CODE") = row.Item("STYLE_CODE")
                rowWHTIREC1.Item("COLOR_CODE") = row.Item("COLOR_CODE")
                rowWHTIREC1.Item("STYLE_DESC") = rowICTSTYL1.Item("STYLE_DESC")
                rowWHTIREC1.Item("COLOR_DESC") = rowICTCOLR1.Item("COLOR_DESC")
                dst.Tables("WHTIREC1").Rows.Add(rowWHTIREC1)
            End If
            rowWHTIREC1.Item(COLUMN_NAME) = Val(rowWHTIREC1.Item(COLUMN_NAME) & "") + Val(row.Item(COLUMN_NAME) & "")
        Next
    End Sub

    Sub Add_Unposted_ADS(COLUMN_NAME As String, TRAN_TYPE As String)

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  DECODE(WHTSTYLX.ITEM_TYPE,'P',WHTPPKM2.STYLE_CODE,WHTSTYLX.STYLE_CODE) STYLE_CODE" & vbCrLf _
            & ", DECODE(WHTSTYLX.ITEM_TYPE,'P',WHTPPKM2.COLOR_CODE,WHTSTYLX.COLOR_CODE) COLOR_CODE" & vbCrLf _
            & ", DECODE(WHTSTYLX.ITEM_TYPE,'P',X.TRAN_QTY * WHTPPKM2.PPK_QTY / WHTPPKM1.PPK_QTY_TOTAL,X.TRAN_QTY) " & COLUMN_NAME & vbCrLf _
            & " from " & ADS_UNPOSTED_ITEM & " X,WHTSTYLX,WHTPPKM2,WHTPPKM1" & vbCrLf _
            & " where WHTSTYLX.LP_CODE (+) = '" & LP_CODE & "'" & vbCrLf _
            & "   and WHTSTYLX.ITEM_CODE (+) = X.ITEM_CODE" & vbCrLf _
            & "   and WHTPPKM2.PPK_CODE (+) = WHTSTYLX.PPK_CODE" & vbCrLf _
            & "   and WHTPPKM1.PPK_CODE (+) = WHTPPKM2.PPK_CODE" & vbCrLf _
            & "   and X.TRAN_TYPE = '" & TRAN_TYPE & "'"
        Add_Unposted(ASCDATA1.GetDataTable, COLUMN_NAME)
    End Sub
End Class