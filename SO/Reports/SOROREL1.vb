Public Class SOROREL1

#Region "Declarations"

    Dim PICK_BATCH_NO As String = ""    ' Pick Batch No; this Release will be defined by this control number
    Dim SHIP_BY_DATE As Date    ' Release all orders Scheduled to ship by
    Dim WHSE_CODE As String     ' Ship From Warehouse Code for Released Orders

    Dim SOTORDR0 As String
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String
    Dim SOTRSRV1 As String
    Dim SOTRSRV2 As String

    Dim SOTPICK1 As String
    Dim SOTPICK2 As String
    Dim SOTSHIP1 As String
    Dim SOTCART1 As String
    Dim SOTCART2 As String
    Dim SOTCARM1 As String
    Dim SOTCARM2 As String

    ' WALMART Multi-PO
    Dim sqlSOTSHIP1W As String = ""
    Dim sqlSOTPICK1W As String = ""
    Dim sqlSOTPICK2W As String = ""
    Dim SOTSHIP1W As String = ""
    Dim SOTPICK1W As String = ""
    Dim SOTPICK2W As String = ""

    Dim ARTCUST1 As String
    Dim ICTSTDQ1 As String
    Dim ICTSTDQ2 As String
    Dim ICTSTDQ3 As String

    Dim SOTORDRG_manual As String = ""

    Dim edi850cust As List(Of String)

    Dim PICK_NO_seq As Int64 = 0        ' Temporary Pick Ticket Sequencer
    Dim SHIP_BOL_NO_seq As Int64 = 0    ' Temporary Shipment Sequencer
    Dim CART_NO_seq As Int32 = 0        ' Temporary Carton Sequencer

    Public CUST_CODE_sql As String = ""
    Public SALES_DIVISION_CODE_sql As String = ""
    Public ORDR_GROUP_NO_sql As String = ""
    Public TERM_CODE_sql As String = ""
    Public blnORDR_GROUP_NO_sql_NOT As Boolean = False

    Public RELEASE_SQL As String = ""   ' this is for the Update of SOTORDR1 & SOTORDR2 from Seletions of Customer Groups and Divisions 
    '                                    - used to set all order allocated quantities to 0 if not selected 


    Dim blnALLOCATION_ONLY As Boolean = False
    Dim blnFORCE_PICK As Boolean = False
    Dim blnREL_PAST_CANCEL As Boolean = False
    Dim blnRELEASE_FUT As Boolean = False

    Dim numCANCEL_FUTURE_DAYS As Integer = 0

    Dim selWHSE As String = ""

    Dim blnMANUAL_ONLY As Boolean = False
    Dim manual_release As Boolean = False

    Dim SQL_ins As New Dictionary(Of String, String)
    Dim TABLE_NAMEs As Dictionary(Of String, String) = Nothing

    Dim tblEDT850TM As DataTable
    Private clsTACENCRY As TAC.ASCENCRY

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        clsTACENCRY = New TAC.ASCENCRY
        Dim rowASTPARMP As DataRow = ASCDATA1.GetDataRow("Select * from ASTPARMP WHERE AS_PARM_KEY = 'Z'")
        If rowASTPARMP Is Nothing OrElse Not rowASTPARMP.Table.Columns.Contains("AS_PARM_USE_ENCRYPTION") OrElse rowASTPARMP.Item("AS_PARM_USE_ENCRYPTION") & String.Empty <> "1" Then
            clsTACENCRY.UseEncryption = False
        Else
            clsTACENCRY.UseEncryption = True
        End If

        Set_WHSE()
        Absx1.numFor("FPDCANCEL_FUTURE_DAYS").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_CANCEL_FUTURE_DAYS") & "")
        Absx1.numFor("FPDSHORT_HOR_DAYS").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_SHORT_HOR_DAYS") & "")

        chkRelFutAvail.Checked = (ROWs("SOTPARM1").Item("SO_PARM_RELEASE_FUT_AVAIL") & "" = "1")
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
        Else
            chkRelFutAvail.Visible = False
        End If

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If ASCMAIN1.DBS_SERVER <> ASCMAIN1.DBS_COMPANY Then
                chkRelFutAvail.Visible = True
            End If
        End If

        dteSHIP_DATE.CalendarInfo.MaxSelectedDays = 1
        Dim SO_PARM_RELEASE_DAYS_AHEAD As Int64 = Val(ROWs.Item("SOTPARM1").Item("SO_PARM_RELEASE_DAYS_AHEAD") & "")
        dteSHIP_DATE.CalendarInfo.ActivateDay(Now.Date.AddDays(SO_PARM_RELEASE_DAYS_AHEAD))
        dteSHIP_DATE.CalendarInfo.SelectedDateRanges.Add(dteSHIP_DATE.CalendarInfo.ActiveDay.Date)

        If MENU_ITEM_OBJECT = "SORORELG" Then
            Absx1.chkFor("CHKALLOCATION_ONLY").Checked = True
            Absx1.chkFor("CHKALLOCATION_ONLY").Enabled = False
        End If

        If MENU_ITEM_OBJECT = "" Then
            chkAllocateNoRelease.Checked = True
            chkAllocateNoRelease.Enabled = False
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            chkCheck_OOBAL.Checked = True
        End If

        chkManualOnly.Visible = (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI")

        grpCANCEL_LETTERS.Top = grpREL_DATE.Top
        grpCANCEL_LETTERS.Left = grpREL_DATE.Left
        grpCANCEL_LETTERS.Visible = False

        Set_Date()
        'Setup_Future_Days()

        Dim tblTATSHIPP As DataTable = Nothing
        tblTATSHIPP = ASCDATA1.GetDataTable("SELECT * FROM TATSHIPP WHERE TABLE_NAME = :PARM1 AND KEY_VALUE = :PARM1", "TATSHIPP", "VV", New String() {"SOTPARM1", "Z"})
        grdTATSHIPP.DataSource = tblTATSHIPP


        SOTORDRG_manual = ASCMAIN1.Temp_Table("Select ORDR_GROUP_NO from SOTORDR0 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRG_manual & " Add Primary Key (ORDR_GROUP_NO)")

        Dim tblSOTORDRU As New DataTable
        With tblSOTORDRU
            .Columns.Add("SEL")
            .Columns.Add("USER_ID")
            .Columns.Add("ORDERS", GetType(System.Int32))
        End With

        grdSOTORDRU.DataSource = tblSOTORDRU

        tabOptions.Tabs("Exclusions").Visible = (ASCMAIN1.CLIENT = "VAN")
        tabOptions.Tabs("Multi-PO").Visible = (ASCMAIN1.CLIENT = "VAN")

        If (ASCMAIN1.CLIENT = "VAN") Then
            Create_Allocations_Exceptions()

        End If

        If (ASCMAIN1.CLIENT = "RGI") Then
            tabOptions.Tabs("Exceptions").Visible = True
            Absx1.chkFor("CHKNO_SPECIAL_COMMENTS").Checked = True
            Absx1.chkFor("CHKCOMBINE").Checked = True
            Absx1.chkFor("CHKINTL").Checked = True
            Absx1.numFor("NUMREL_PCT").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_REL_PCT") & String.Empty)
            Absx1.numFor("SO_PARM_REL_CUBE").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_REL_CUBE") & "")
            chkMerge.Checked = True
            chkCustCreditHold.Checked = True

            ASCMAIN1.sql = "Select '0' SEL, ORDR_REL_SHORT_OPER USER_ID, COUNT (*) ORDERS" & vbCrLf _
                & " from SOTORDRG,SOTORDR0" & vbCrLf _
                & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_CNT_OPEN > 0" & vbCrLf _
                & "   and SOTORDRG.ORDR_REL_SHORT = '1'" & vbCrLf _
                & " group by ORDR_REL_SHORT_OPER"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            For Each ROW As DataRow In tbl.Rows
                tblSOTORDRU.Rows.Add(ROW.ItemArray)
            Next
            grdSOTORDRU.Visible = False

        Else
            tabOptions.Tabs("Exceptions").Visible = False
            ' tabOptions.Style = UltraWinTabControl.UltraTabControlStyle.Wizard

            grdSOTORDRU.Visible = False
        End If

        chkECommerce.Visible = (ASCMAIN1.CLIENT = "RGI")

        If ASCMAIN1.CLIENT = "VAN" Then

            ' Build Walmart Multi-PO Selection Table

            ASCMAIN1.sql = "Select X.*, DECODE(EDI_PROMOTION,'POS REPLEN','1','0') SEL from (" & vbCrLf _
                & "Select ORDR_DEPT, ORDR_DATE, EDI_PROMOTION, EDI_CONS_NO, COUNT (*) POS" & vbCrLf _
                & ", MIN (ORDR_SHIP_DATE) ORDR_SHIP_DATE_MIN" & vbCrLf _
                & ", MAX (ORDR_SHIP_DATE) ORDR_SHIP_DATE_MAX" & vbCrLf _
                & ", MIN (ORDR_CANCEL_DATE) ORDR_CANCEL_DATE_MIN" & vbCrLf _
                & ", MAX (ORDR_CANCEL_DATE) ORDR_CANCEL_DATE_MAX" & vbCrLf _
                & " from SOTORDR0,EDT850T1" & vbCrLf _
                & " where CUST_CODE = 'WALMART' AND ORDR_CNT_OPEN <> 0" & vbCrLf _
                & "   and EDI_CONS_NO IS NULL" & vbCrLf _
                & "   and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                & " group by ORDR_DEPT, ORDR_DATE, EDI_CONS_NO, EDI_PROMOTION" & vbCrLf _
                & ") X"

            tblEDT850TM = ASCDATA1.GetDataTable
            tblEDT850TM.Columns("SEL").ReadOnly = False
            grdEDT850TM.DataSource = tblEDT850TM
            For Each gcol As UltraWinGrid.UltraGridColumn In grdEDT850TM.DisplayLayout.Bands(0).Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            With grdEDT850TM.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.True
                .AllowUpdate = DefaultableBoolean.True
            End With

            ' ADD NEW TABLES



        End If


        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {Me.Name, ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If

        btnConsolidateWalmart.Visible = (ASCMAIN1.CLIENT = "VAN")
    End Sub

    Sub Setup_Allocate_No_Release()
        grpReleaseOptions.Visible = Not chkAllocateNoRelease.Checked
        lblAllUnReleasable.Visible = Not chkAllocateNoRelease.Checked
        numHorizonDays.Visible = chkAllocateNoRelease.Checked
        lblHorizonDays.Visible = chkAllocateNoRelease.Checked
        grpREL_DATE.Visible = Not chkAllocateNoRelease.Checked
        chkemailSReps.Visible = chkAllocateNoRelease.Checked And (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA")
        grpCANCEL_LETTERS.Visible = chkAllocateNoRelease.Checked _
            AndAlso (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") _
            AndAlso Absx1.txtFor("WHSE_CODE").Text = "MS"

        If ASCMAIN1.CLIENT = "RGI" Then
            grdSOTORDRU.Visible = Not chkAllocateNoRelease.Checked And chkManualOnly.Checked
            chkECommerce.Visible = Not chkAllocateNoRelease.Checked And Not chkManualOnly.Checked
        End If
    End Sub


    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {"SOROREL1", ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        ORDR_GROUP_NO_sql = SQLA("ORDR_GROUP_NO", , True)
        CUST_CODE_sql = SQLA("CUST_CODE", , True)
        SALES_DIVISION_CODE_sql = SQLA("SALES_DIVISION_CODE", , True)

        If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
            TERM_CODE_sql = SQLA("TERM_CODE", , True)
        End If


        SHIP_BY_DATE = dteSHIP_DATE.CalendarInfo.ActiveDay.Date

        blnALLOCATION_ONLY = Absx1.chkFor("CHKALLOCATION_ONLY").Checked
        blnFORCE_PICK = Absx1.chkFor("CHKFORCE_PICK").Checked
        selWHSE = Absx1.optFor("OPTWHSE").Value

        SQL_ins.Clear()
        SQL_ins.Add("CUST_CODE", SQL_in("SOTORDR1.CUST_CODE"))
        SQL_ins.Add("SALES_DIVISION_CODE", SQL_in("SOTORDR1.SALES_DIVISION_CODE"))
        SQL_ins.Add("ORDR_GROUP_NO", SQL_in("SOTORDR1.ORDR_GROUP_NO"))

        If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
            SQL_ins.Add("TERM_CODE", SQL_in("SOTORDR1.TERM_CODE"))
        End If

        blnORDR_GROUP_NO_sql_NOT = (SQLA("ORDR_GROUP_NO", "EXCLUDE") = "1")
        blnMANUAL_ONLY = chkManualOnly.Checked
        blnREL_PAST_CANCEL = Absx1.chkFor("CHKREL_PAST_CANCEL").Checked
        blnRELEASE_FUT = Absx1.chkFor("CHKRELEASE_FUT").Checked
        numCANCEL_FUTURE_DAYS = Val(Absx1.numFor("FPDCANCEL_FUTURE_DAYS").Value & "")
        TABLE_NAMEs = Nothing

        If grpCANCEL_LETTERS.Visible = False Then
            chkCheck_CANCEL.Checked = False
        End If

        Build_Workfile2()

        If ASCMAIN1.CLIENT = "VAN" Then
            Create_Allocations_Exceptions()
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            ASCDATA1.ExecuteSP("WJZOP", "VV", New String() {Me.Name, ASCMAIN1.USER_ID}, New String() {"FORM_NAME", "USER_ID"})
        End If
    End Sub

    Sub Build_Workfile2()

        ASCMAIN1.Progress("Setting Up Orders (Demand)", "")

        Dim sql_where As String = ""

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" _
        Or ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            sql_where = ""
            If blnALLOCATION_ONLY Or blnFORCE_PICK Or manual_release Then
            Else
                sql_where &= SQL_ins("CUST_CODE")
                sql_where &= SQL_ins("SALES_DIVISION_CODE")
                sql_where &= SQL_ins("ORDR_GROUP_NO")

                If ASCMAIN1.CLIENT = "RGI" Then
                    sql_where &= SQL_ins("TERM_CODE")

                    If blnMANUAL_ONLY Then
                        ASCDATA1.ExecuteSQL("Delete from " & SOTORDRG_manual)
                        Dim USER_IDs As String = ""
                        For Each ROW As DataRow In DirectCast(grdSOTORDRU.DataSource, DataTable).Select("SEL='1'")
                            USER_IDs &= ",'" & ROW.Item("USER_ID") & "'"
                        Next
                        ASCDATA1.ExecuteSQL("Insert into " & SOTORDRG_manual & " Select ORDR_GROUP_NO from SOTORDRG where ORDR_REL_SHORT = '1' and ORDR_REL_SHORT_OPER in (" & Mid(USER_IDs, 2) & ")")

                        sql_where &= " and SOTORDR1.ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDRG_manual & ")"
                    End If
                End If

            End If
        End If

        If ASCMAIN1.CLIENT = "RGI" Then
            If chkECommerce.Checked Then
                ' 03/21/2019 - Permit releasing for all Ecom Customers
                If Absx1.txtFor("ECOM_CODE").Text <> String.Empty Then
                    sql_where &= " and SOTORDR1.ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'"
                Else
                    sql_where &= " and SOTORDR1.ECOM_CODE IS NOT NULL"
                End If
            Else
                sql_where &= " and SOTORDR1.ECOM_CODE IS NULL"
            End If

        End If

        If TABLE_NAMEs Is Nothing Then
            TABLE_NAMEs = TAC.SOCMAIN1.Allocation_Initialization(Me,
                    IIf(selWHSE = "A", "", WHSE_CODE),
                    blnFORCE_PICK,
                    blnALLOCATION_ONLY,
                    True,
                    ORDR_GROUP_NO_sql, SHIP_BY_DATE, sql_where, manual_release)

            edi850cust = TAC.TACMAIN1.Get_EDI_Custs("850")

            SOTPICK1 = Create_Temporary_Table("SOTPICK1", "PICK_NO")
            SOTPICK2 = Create_Temporary_Table("SOTPICK2", "PICK_NO,PICK_LNO")
            SOTSHIP1 = Create_Temporary_Table("SOTSHIP1", "SHIP_BOL_NO")
            SOTCART1 = Create_Temporary_Table("SOTCART1", "CART_NO")
            SOTCART2 = Create_Temporary_Table("SOTCART2", "CART_NO,CART_LNO")

            If ASCMAIN1.CLIENT = "VAN" Then
                ASCDATA1.ExecuteSQL($"Alter Table {SOTCART2} Add CONSOLIDATED VARCHAR2 (1)")
            End If


            If ASCMAIN1.CLIENT = "VAN" Then
                    SOTCARM1 = Create_Temporary_Table("SOTCARM1", "CART_NO")
                    SOTCARM2 = Create_Temporary_Table("SOTCARM2", "CART_NO,CART_LNO")
                    '      ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK2 & " Add ORDR_AMT_PICK NUMBER (13,2)")
                End If

            Else
                For Each TABLE_NAME As String In New String() {SOTPICK1, SOTPICK2, SOTSHIP1, SOTCART1, SOTCART2}
                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME)
            Next
            If ASCMAIN1.CLIENT = "VAN" Then
                For Each TABLE_NAME As String In New String() {SOTCARM1, SOTCARM2}
                    ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME)
                Next
            End If

            For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR0", "ARTCUST1", "ICTSTDQ1", "SOTORDR2", "SOTRSRV1", "SOTRSRV2"}
                ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAMEs(TABLE_NAME))
            Next

            'For Each sql As String In TABLE_NAMEs.Keys
            '    If sql.StartsWith("sql") Then
            '        Dim sqlstmt As String = Replace(TABLE_NAMEs(sql), "'STYLE_CODE'", "'" & STYLE_CODE & "'")
            '        ASCDATA1.ExecuteSQL(sqlstmt)
            '    End If
            'Next

            dst.Tables("SOTSUPP0").Rows.Clear()
                dst.Tables("SOTSUPPI").Rows.Clear()
                dst.Tables("SOTORDR7").Rows.Clear()
                dst.Tables("ICTSTDQ1").Rows.Clear()
                dst.Tables("ICTSTDQ2").Rows.Clear()
            End If

        SOTORDR0 = TABLE_NAMEs("SOTORDR0")
        SOTORDR1 = TABLE_NAMEs("SOTORDR1")
        SOTORDR2 = TABLE_NAMEs("SOTORDR2")
        SOTRSRV1 = TABLE_NAMEs("SOTRSRV1")
        SOTRSRV2 = TABLE_NAMEs("SOTRSRV2")
        ICTSTDQ1 = TABLE_NAMEs("ICTSTDQ1")
        ICTSTDQ2 = TABLE_NAMEs("ICTSTDQ2")
        ARTCUST1 = TABLE_NAMEs("ARTCUST1")


        If ASCMAIN1.CLIENT = "VAN" Then
            ASCMAIN1.sql = "Select SOTPCKP2.* from SOTPCKP2 where PACK_GROUP_STATUS = 'A' and ORDR_GROUP_NO = :PARM1"
            Create_TDA(dst.Tables.Add, "SOTPCKP2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTPCKC4.ORDR_NO, SOTPCKC4.PACK_NO, SOTPCKC4.PACK_CONFIG_NO" & vbCrLf _
                & " from SOTPCKC4 where SOTPCKC4.PACK_NO = :PARM1 "
            Create_TDA(dst.Tables.Add, "SOTPCKC4_ORDR_NO", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPCKC3.*" & vbCrLf _
                & " from SOTPCKC3 where SOTPCKC3.PACK_NO = :PARM1 "
            Create_TDA(dst.Tables.Add, "SOTPCKC3", "**", 0, False, "V", 4)

            ASCMAIN1.sql = "Select WHTPKGM1.*" & vbCrLf _
                & " from WHTPKGM1 where WHTPKGM1.USE_FOR_P2L = '1'"
            Create_TDA(dst.Tables.Add, "WHTPKGM1", "**", 0, False, "", 1)
            Fill_Records("WHTPKGM1")

            ASCMAIN1.sql = "Select ICTBODY2.*" & vbCrLf _
                & " from ICTBODY2"
            Create_TDA(dst.Tables.Add, "ICTBODY2", "**", 0, False, "", 1)
            Fill_Records("ICTBODY2")
        End If

        ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_WEIGHT, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.INNER_PACK_QTY" _
            & " from ICTSTYL1" _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTORDR2 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 1))

        ASCMAIN1.sql = "Select SOTSREP1.* from SOTSREP1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))

        Create_TDA(dst.Tables.Add(), "SOTPICK0", "*")

        Create_Relation("SOTORDR1", "SOTPICK1", "ORDR_NO")

        With dst.Tables("SOTPICK2").Columns
            .Add("STYLE_CODE")
            .Add("COLOR_CODE")
            If ASCMAIN1.CLIENT = "VAN" Then
                .Add("SUB_BODY_CODE", GetType(System.String))
                .Add("STANDARD_CUBE_PER_UNIT", GetType(System.Decimal))
                .Add("CUBE_REQD", GetType(System.Decimal), "PICK_QTY * STANDARD_CUBE_PER_UNIT")
                .Add("CART_NO", GetType(System.String))
            End If
        End With

        With dst.Tables("SOTCART1").Columns
            .Add("PKG_CUBE", GetType(System.Decimal))
            .Add("PKG_CUBE_PACK", GetType(System.Decimal))
        End With
        With dst.Tables("SOTCART2").Columns
            .Add("STYLE_WEIGHT", GetType(System.Decimal))
        End With

        Dim SO_PARM_UPC_VENDOR_ID As String = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID")
        Dim SO_PARM_MAX_CARTON As Integer = Val(ROWs("SOTPARM1").Item("SO_PARM_MAX_CARTON") & "")

        ASCMAIN1.sql = "Select INV_DUE_DATE, INV_BALANCE from ARTOPEN1" _
            & " where CUST_CODE = :PARM1" _
            & " and INV_BALANCE > 0 and INV_TYPE = 'I'" _
            & " and INV_DUE_DATE < :PARM2"
        Create_TDA(dst.Tables.Add, "ARTOPENP", "**", 0, False, "VD", 0)

        Create_TDA(dst.Tables.Add, "SOTCSTP1", "*", 1, False)

        With dst.Tables.Add("SOTSHIP3")
            .Columns.Add("SHIP_BOL_NO")
            .Columns.Add("STYLE_CODE")
            .PrimaryKey = New DataColumn() {.Columns("SHIP_BOL_NO"), .Columns("STYLE_CODE")}
        End With

        With dst.Tables.Add("SOTSHIP2")
            .Columns.Add("SHIP_BOL_NO")
            .Columns.Add("STYLE_CODE")
            .Columns.Add("COLOR_CODE")
            .PrimaryKey = New DataColumn() {.Columns("SHIP_BOL_NO"), .Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
        End With

        ' Main Process

        Dim SOTSUPP1 As String = ""
        Dim SOTDEMD1 As String = ""

        If (ASCMAIN1.DBS_SERVER = "RGIx" Or ASCMAIN1.DBS_COMPANY = "RGIx") And sql_where <> "" Then
            ASCMAIN1.sql = "Select Distinct STYLE_CODE, COLOR_CODE from " & SOTORDR2 & " where ORDR_QTY_OPEN <> 0"

            For Each rowSC As DataRow In ASCDATA1.GetDataTable.Select("", "STYLE_CODE,COLOR_CODE")
                Dim STYLE_CODE As String = rowSC.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSC.Item("COLOR_CODE")

                TAC.SOCMAIN1.Allocation(Me, _
                    blnFORCE_PICK, _
                    blnALLOCATION_ONLY, _
                     IIf(selWHSE = "A", "", WHSE_CODE), _
                     ORDR_GROUP_NO_sql, edi850cust, _
                     SOTSUPP1, SOTDEMD1, TABLE_NAMEs, , , STYLE_CODE, COLOR_CODE, manual_release)
            Next
        Else

            TAC.SOCMAIN1.Allocation(Me, _
                blnFORCE_PICK, _
                blnALLOCATION_ONLY, _
                 IIf(selWHSE = "A", "", WHSE_CODE), _
                 ORDR_GROUP_NO_sql, edi850cust, _
                SOTSUPP1, SOTDEMD1, TABLE_NAMEs, , (ROWs("SOTPARM1").Item("SO_PARM_ALLO_SEQ") & "" = "1"), "", "", manual_release, chkCustCreditHold.Checked)

        End If

        ' Allocation (above) processes all orders within whse selected (if any), and only specific orders if force-pick is going on
        '  If we are doing allocation only, we review allocation results and mark items with inventory shortage codes, and then produce reports
        '  If we are doing Release, we need to pare back the orders to just those in the UI filters by Customer, Division, Order Group, and Date

        If blnALLOCATION_ONLY Then
            Inventory_Shortages()

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Order Data for Reports")

            ASCMAIN1.sql = "Select * from " & ARTCUST1
            Fill_Records("ARTCUST1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & SOTORDR1 & " SOTORDR1"
            Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & SOTORDR2 & " SOTORDR2"
            Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        Else


            If ASCMAIN1.CLIENT = "RGI" Then

                ' delete all orders that are not part of the release scope, 
                ' but are included in SOTORDR1 because they are part of the allocation_only scope

                ASCMAIN1.sql = "Delete from " & SOTORDR2 & " where ORDR_NO in (Select ORDR_NO from " & SOTORDR1 & " where ALLOCATION_ONLY_SCOPE = '1')"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ALLOCATION_ONLY_SCOPE = '1'"
                ASCDATA1.ExecuteSQL()

            End If



            If Order_Release() Then
                Create_PICK_SHIP_CART()
            End If
        End If

        ' Create Header & Detail Records for Order Groups Failing to Release Records for

        Dim sql_filter_when_allocation_only As String = ""
        If blnALLOCATION_ONLY Then
            ' observe filter for specific customers, divisions, or groups for report on allocation only
            sql_filter_when_allocation_only &= SQL_in("SALES_DIVISION_CODE", "SOTORDR1.SALES_DIVISION_CODE")
            sql_filter_when_allocation_only &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
            sql_filter_when_allocation_only &= SQL_in("ORDR_GROUP_NO", "SOTORDR1.ORDR_GROUP_NO")
        End If

        ASCMAIN1.sql = "Select SOTORDRG.*, SOTORDR0.CUST_CODE, SOTORDR0.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.WHSE_CODE" & vbCrLf _
            & ", SOTORDR0.ORDR_CNT_OPEN" & vbCrLf _
            & " from SOTORDRG,SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = SOTORDRG.ORDR_GROUP_NO and ROWNUM < 1"
        Dim SOTORDRG As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRG & " Add Primary Key (ORDR_GROUP_NO)")

        ASCMAIN1.sql = "Select SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.WHSE_CODE" & vbCrLf _
            & ", Min (ORDR_NO) as ORDR_NO_MIN, Max (ORDR_NO) as ORDR_NO_MAX" & vbCrLf _
            & ", Min (ORDR_REL_HOLD_CODES) as ORDR_REL_HOLD_CODES" & vbCrLf _
            & ", Count (*) as ORDR_CNT_OPEN" & vbCrLf _
            & " from " & SOTORDR1 & " SOTORDR1, " & SOTORDR0 & " SOTORDR0" & vbCrLf _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Not Null " & vbCrLf _
            & sql_filter_when_allocation_only _
            & " group by SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.WHSE_CODE"
        ASCMAIN1.sql = "Insert into " & SOTORDRG _
            & " (ORDR_GROUP_NO,CUST_CODE,SALES_DIVISION_CODE,ORDR_DATE,ORDR_CUST_PO" _
            & ",ORDR_SHIP_DATE,ORDR_CANCEL_DATE,WHSE_CODE,ORDR_NO_MIN,ORDR_NO_MAX" _
            & ",ORDR_REL_HOLD_CODES,ORDR_CNT_OPEN)" & vbCrLf _
            & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SOTORDRG & " Set LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & SOTORDRG
        Create_TDA(dst.Tables.Add("SOTORDRG"), SOTORDRG, "**", 0, True, "", 1)
        Fill_Records("SOTORDRG")

        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDRG", 1))

        ' dst.Tables("SOTORDRG").Columns.Add("ORDR_REL_HOLD_CODES")

        '    'Added to include Reservations when allocating only for Gabe by WR - 8/27/05
        '    If blnALLOCATION_ONLY Then
        '        sql = "INSERT INTO SOWORDRG"
        '        & " Select SOWRSRV1.RSRV_NO AS ORDR_GROUP_NO, SOWRSRV1.CUST_CODE,"
        '        & "'" & String$(20, "X") & "'  as ORDR_REL_HOLD_CODES, SOWRSRV1.SALES_DIVISION_CODE, "
        '        & " Min (SOWRSRV1.INIT_DATE) as ORDR_DATE, Min (SOWRSRV1.ORDR_CUST_PO) as ORDR_CUST_PO, "
        '        & " Min (SOWRSRV1.ORDR_SHIP_DATE) as ORDR_SHIP_DATE, Min (SOWRSRV1.ORDR_CANCEL_DATE) as ORDR_CANCEL_DATE, "
        '        & " Min (SOWRSRV1.RSRV_NO) as ORDR_NO_MIN, Max (SOWRSRV1.RSRV_NO) as ORDR_NO_MAX, "
        '        & " Count (*) as ORDR_NOS"
        '        & " from SOWRSRV1"
        '        & " where SOWRSRV1.ORDR_REL_HOLD_CODES is Not Null"
        '        & " group by SOWRSRV1.RSRV_NO, SOWRSRV1.CUST_CODE, '" & String$(20, "X") & "'" & ", SOWRSRV1.SALES_DIVISION_CODE"
        '        AccD.Execute sql
        '    End If
        '    'End of Addition for Gabe.

        'Removed from criteria by WR for Gabe on 4/17/02
        '    If blnALLOCATION_ONLY Then
        '        & "   and SOWORDR1.ORDR_CANCEL_DATE < #" & Format$(Now + Absx1.numFor("FPDSHORT_HOR_DAYS").Value, "MM/DD/YYYY") & "#"
        '    End If

        ASCMAIN1.sql = "Select DISTINCT ORDR_GROUP_NO, ORDR_REL_HOLD_CODES" _
            & " from " & SOTORDR1 _
            & " where ORDR_REL_HOLD_CODES is Not Null "
        'Removed from criteria by WR for Gabe on 4/17/02
        '    If blnALLOCATION_ONLY Then
        '        & "   and SOWORDR1.ORDR_CANCEL_DATE < #" & Format$(Now + Absx1.numFor("FPDSHORT_HOR_DAYS").Value, "MM/DD/YYYY") & "#"
        '    End If
        ASCMAIN1.sql &= " order by ORDR_GROUP_NO"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
            Dim rowSOTORDRG As DataRow = dst.Tables("SOTORDRG").Rows.Find(ORDR_GROUP_NO)
            If rowSOTORDRG IsNot Nothing Then
                Dim ORDR_REL_HOLD_CODES_new = row.Item("ORDR_REL_HOLD_CODES")
                Dim ORDR_REL_HOLD_CODES As String = rowSOTORDRG.Item("ORDR_REL_HOLD_CODES") & ""
                For i As Integer = 1 To Len(ORDR_REL_HOLD_CODES_new)
                    If Not ORDR_REL_HOLD_CODES.Contains(Mid(ORDR_REL_HOLD_CODES_new, i, 1)) Then
                        ORDR_REL_HOLD_CODES &= Mid(ORDR_REL_HOLD_CODES_new, i, 1)
                    End If
                Next i
                If rowSOTORDRG.Item("ORDR_REL_HOLD_CODES") & "" <> ORDR_REL_HOLD_CODES Then
                    rowSOTORDRG.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES
                End If
            End If
        Next

        Update_Record_TDA("SOTORDRG")

        'ASCMAIN1.sql = "Insert into SOTORDRG (ORDR_GROUP_NO, INIT_DATE, INIT_OPER)" & vbCrLf _
        '    & " Select ORDR_GROUP_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "' from " & SOTORDRG & vbCrLf _
        '    & "  where ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDRG & " minus Select ORDR_GROUP_NO from SOTORDRG)"
        'ASCDATA1.ExecuteSQL()

        ' Probably need to open this up to all companies.
        ' THIS ALSO APPLIES TO SOTORDRS CHANGE BELOW
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ASCDATA1.ExecuteSQL("Update SOTORDRG Set ORDR_REL_HOLD_CODES = Null where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")")
        End If

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & SOTORDRG & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTORDRG Set ORDR_REL_HOLD_CODES = R1.ORDR_REL_HOLD_CODES" & vbCrLf _
            & "    , ORDR_NO_MIN = R1.ORDR_NO_MIN, ORDR_NO_MAX = R1.ORDR_NO_MAX" & vbCrLf _
            & "    , LAST_DATE = R1.LAST_DATE, LAST_OPER = R1.LAST_OPER" & vbCrLf _
            & "    where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into SOTORDRG (ORDR_GROUP_NO, ORDR_NO_MIN, ORDR_NO_MAX, ORDR_REL_HOLD_CODES, INIT_DATE, INIT_OPER, LAST_DATE, LAST_OPER)" & vbCrLf _
            & "     values (R1.ORDR_GROUP_NO, R1.ORDR_NO_MIN, R1.ORDR_NO_MAX, R1.ORDR_REL_HOLD_CODES, R1.LAST_DATE, R1.LAST_OPER, R1.LAST_DATE, R1.LAST_OPER);" & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Select" _
                    & " SOTORDR2.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" _
                    & ", Sum (DECODE(SOTORDR1.ORDR_REL_HOLD_CODES,Null,1,0)) NULLS" _
                    & ", Sum (DECODE(SOTORDR1.ORDR_REL_HOLD_CODES,Null,0,1)) NOT_NULLS" _
                    & " from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" _
                    & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
                    & "   and SOTORDR1.ORDR_REL_HOLD_CODEs is Not Null" _
                    & " group by " _
                    & " SOTORDR2.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2 Set WIPIND = 'M'" _
            & " where (ORDR_GROUP_NO, STYLE_CODE, COLOR_CODE) in " _
            & " (Select * from (" & ASCMAIN1.sql & ") where NULLS <> 0 and NOT_NULLS <> 0)"

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE" & vbCrLf _
            & ", SOTORDR2.COLOR_CODE, MIN(SOTORDR2.STYLE_DESC) AS STYLE_DESC, MIN(SOTORDR1.SREP_CODE) AS SREP_CODE" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY) as ORDR_QTY" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_OPEN) as ORDR_QTY_OPEN" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO) as ORDR_QTY_ALLO" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_PICK) as ORDR_QTY_PICK" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_SHIP) as ORDR_QTY_SHIP" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_CANC) as ORDR_QTY_CANC" & vbCrLf _
            & ", Max (SOTORDR2.ORDR_RELEASE_AVAIL) as ORDR_RELEASE_AVAIL" & vbCrLf _
            & ", Max (SOTORDR2.ORDR_RELEASE_SHIP) as ORDR_RELEASE_SHIP" & vbCrLf _
            & ", Max (SOTORDR2.WIP_IND) as WIP_IND" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_CUR) as ORDR_QTY_ALLO_CUR" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_FUT) as ORDR_QTY_ALLO_FUT" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_CXL) as ORDR_QTY_ALLO_CXL" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) as ORDR_AMT" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) as ORDR_AMT_OPEN" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_CUR * SOTORDR2.ORDR_UNIT_PRICE) as ORDR_AMT_ALLO_CUR" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_FUT * SOTORDR2.ORDR_UNIT_PRICE) as ORDR_AMT_ALLO_FUT" & vbCrLf _
            & ", Sum (SOTORDR2.ORDR_QTY_ALLO_CXL * SOTORDR2.ORDR_UNIT_PRICE) as ORDR_AMT_ALLO_CXL" & vbCrLf _
            & ", Max (SOTORDR2.RANGE_STYLE_CODE) as RANGE_STYLE_CODE" & vbCrLf _
            & " from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Not Null " & vbCrLf
        'Removed from criteria by WR for Gabe on 4/17/02
        '    If blnALLOCATION_ONLY Then
        '        & "   and SOWORDR1.ORDR_CANCEL_DATE < #" & Format$(Now + Absx1.numFor("FPDSHORT_HOR_DAYS").Value, "MM/DD/YYYY") & "#"
        '    End If
        ASCMAIN1.sql &= "" _
            & " group by SOTORDR1.ORDR_GROUP_NO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"

        Dim SOTORDRS As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRS & " Add Primary Key (ORDR_GROUP_NO,STYLE_CODE,COLOR_CODE)")

        ASCMAIN1.sql = "UPDATE " & SOTORDRS & " Set WIP_IND = NULL" _
            & " where ORDR_RELEASE_SHIP IS NULL AND ORDR_RELEASE_AVAIL IS NULL AND WIP_IND IS NOT NULL"
        ASCDATA1.ExecuteSQL()

        ' note - at RGI the line below did not clear out all groups when the from was SOTORDRG
        ASCDATA1.ExecuteSQL("Delete from SOTORDRS where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDRS & ")")

        ' note - at NYA the line below did not clear out all groups when the from was SOTORDRS either, so opening it up to all groups in SOTORDR1.  
        ' Probably need to open this up to all companies.
        ' THIS ALSO APPLIES TO SOTORDRG CHANGE ABOVE
        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ASCDATA1.ExecuteSQL("Delete from SOTORDRS where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")")
        End If

        ASCDATA1.ExecuteSQL("Insert into SOTORDRS Select * from " & SOTORDRS)

        ASCMAIN1.sql = "Select * from " & SOTORDRS
        Create_TDA(dst.Tables.Add("SOTORDRS"), SOTORDRS, "**", 0, True, "", 3)
        Fill_Records("SOTORDRS")

        With dst.Tables("SOTORDRS")
            .Columns("WIP_IND").ReadOnly = False
        End With


        '    'Added to include Reservations when allocating only for Gabe by WR - 8/27/05
        '    If blnALLOCATION_ONLY Then
        '        sql = "INSERT INTO SOWORDRS"
        '        & " Select"
        '        & " SOWRSRV1.RSRV_NO AS ORDR_GROUP_NO,"
        '        & " SOWRSRV2.STYLE_CODE ,"
        '        & " SOWRSRV2.COLOR_CODE,"
        '        & " MIN(ICWSTYL1.STYLE_DESC) AS STYLE_DESC,"
        '        & " MIN(SOWRSRV1.SREP_CODE) AS SREP_CODE ,"
        '        & " Sum (SOWRSRV2.RSRV_QTY) as ORDR_QTY ,"
        '        & " Sum (SOWRSRV2.RSRV_QTY_OPEN) as ORDR_QTY_OPEN ,"
        '        & " Sum (SOWRSRV2.RSRV_QTY_ALLO) as ORDR_QTY_ALLO ,"
        '        & " Sum (0) as ORDR_QTY_PICK ,"
        '        & " Sum (0) as ORDR_QTY_SHIP ,"
        '        & " Sum (SOWRSRV2.RSRV_QTY_CANC) as ORDR_QTY_CANC ,"
        '        & " Max (SOWRSRV2.ORDR_RELEASE_AVAIL) as ORDR_RELEASE_AVAIL ,"
        '        & " NULL as ORDR_RELEASE_SHIP ,"
        '        & " NULL as WIP_IND ,"
        '        & " Sum (SOWRSRV2.ORDR_QTY_ALLO_CUR) as ORDR_QTY_ALLO_CUR ,"
        '        & " Sum (SOWRSRV2.ORDR_QTY_ALLO_FUT) as ORDR_QTY_ALLO_FUT ,"
        '        & " Sum (SOWRSRV2.RSRV_QTY_OPEN * SOWRSRV2.ORDR_UNIT_PRICE) as ORDR_AMT_OPEN ,"
        '        & " Max(Null) As RANGE_STYLE_CODE"
        '        & " From SOWRSRV1, SOWRSRV2, ICWSTYL1"
        '        & " Where SOWRSRV1.RSRV_NO = SOWRSRV2.RSRV_NO"
        '        & " AND SOWRSRV2.STYLE_CODE = ICWSTYL1.STYLE_CODE"
        '        & " and SOWRSRV1.ORDR_REL_HOLD_CODES is Not Null"
        '        & " group by SOWRSRV1.RSRV_NO, SOWRSRV2.STYLE_CODE,  SOWRSRV2.COLOR_CODE"
        '        AccD.Execute sql
        '    End If
        '    'End of Addition for Gabe.

        For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select _
                ("ORDR_RELEASE_SHIP IS NULL AND ORDR_RELEASE_AVAIL IS NULL AND WIP_IND IS NOT NULL")
            rowSOTORDRS.Item("WIP_IND") = DBNull.Value
        Next



        If ASCMAIN1.CLIENT = "VAN" Then

            ' Prepare Multi-PO PT Consolidation work tables


            sqlSOTSHIP1W = "Select SOTSHIP1.SHIP_BOL_NO, SOTSHIP1.SHIP_BOL_NO_CONS, EDT850T1.EDI_CONS_NO" & vbCrLf _
                & ", SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT" & vbCrLf _
                & "from SOTORDR0,EDT850T1," & SOTSHIP1 & " SOTSHIP1" & vbCrLf _
                & "where SOTORDR0.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO" & vbCrLf _
                & "  and EDT850T1.EDI_CONS_NO is Not Null" & vbCrLf _
                & "  and SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
                & "  and SOTSHIP1.SHIP_STATUS = 'P'"

            SOTSHIP1W = ASCMAIN1.Temp_Table(sqlSOTSHIP1W.Replace("and SOTSHIP1.SHIP_STATUS = 'P'", _
                                                                 "and SOTSHIP1.SHIP_STATUS = 'P' and ROWNUM <1"))

            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP1W & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Create Index I_" & SOTSHIP1W & "_1 on " & SOTSHIP1W & " (EDI_CONS_NO, CUST_DC_NO)")

            sqlSOTPICK1W = "Select SOTPICK1.PICK_NO, SOTPICK1.PICK_NO_CONS, EDT850T1.EDI_CONS_NO" & vbCrLf _
                & ", SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & "from SOTORDR1,EDT850T1," & SOTPICK1 & " SOTPICK1" & vbCrLf _
                & "where SOTORDR1.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO" & vbCrLf _
                & "  and EDT850T1.EDI_CONS_NO is Not Null" & vbCrLf _
                & "  and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "  and SOTPICK1.PICK_STATUS = 'P'"

            SOTPICK1W = ASCMAIN1.Temp_Table(sqlSOTPICK1W.Replace("and SOTPICK1.PICK_STATUS = 'P'", _
                                                                 "and SOTPICK1.PICK_STATUS = 'P' and ROWNUM <1"))

            ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1W & " Add Primary Key (PICK_NO)")
            ASCDATA1.ExecuteSQL("Create Index I_" & SOTPICK1W & "_1 on " & SOTPICK1W & " (EDI_CONS_NO, CUST_DC_NO, CUST_STORE_NO)")


            sqlSOTPICK2W = "Select SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", EDT850T1.EDI_CONS_NO, SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & "from SOTORDR1,EDT850T1," & SOTPICK1 & " SOTPICK1," & SOTPICK2 & " SOTPICK2,SOTORDR2" & vbCrLf _
                & "where SOTORDR1.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO" & vbCrLf _
                & "  and EDT850T1.EDI_CONS_NO is Not Null" & vbCrLf _
                & "  and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "  and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                & "  and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"

            SOTPICK2W = ASCMAIN1.Temp_Table(sqlSOTPICK2W.Replace("and SOTPICK1.PICK_STATUS = 'P'", _
                                                                 "and SOTPICK1.PICK_STATUS = 'P' and ROWNUM <1"))
            ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK2W & " Add Primary Key (PICK_NO,PICK_LNO)")
        End If


        BeginTrans()

        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" _
        Or ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ' SKIP THIS STEP - IT IS BEING DONE INSIDE SOCMAIN1.ALLOCATION
        Else
            TAC.SOCMAIN1.Update_Status_by_Date(Me, ICTSTDQ1, ICTSTDQ2, ICTSTDQ3, WHSE_CODE, blnALLOCATION_ONLY, SOTORDR2, blnFORCE_PICK, manual_release)
        End If

        If blnALLOCATION_ONLY Then

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is Select * from " & SOTORDR1 & ";" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update SOTORDR1 Set" _
                & "      ORDR_REL_HOLD_CODES = R1.ORDR_REL_HOLD_CODES" _
                & "    where ORDR_NO = R1.ORDR_NO;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" _
                & " Declare Cursor C1 is Select * from " & SOTORDR2 & ";" _
                & " Begin" _
                & "  For R1 in C1 Loop" _
                & "   Update SOTORDR2 Set" _
                & "      ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" _
                & "    , ORDR_RELEASE = R1.ORDR_RELEASE" _
                & "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" _
                & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
                & "  End Loop;" _
                & " End;" _
                & "End;"

            ASCDATA1.ExecuteSQL()

        Else
            Update_Release() ' Update Pick Ticket, Shipment Control & Carton Tables

            ' Special things for Regency Int'l
            If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
                If dst.Tables.Contains("SOTOREMM") Then
                    For Each rowSOTOREMM As DataRow In dst.Tables("SOTOREMM").Select("")
                        rowSOTOREMM.Item("PICK_AMT") = 0
                        rowSOTOREMM.Item("PICK_QTY") = 0

                        Dim ORDR_NO As String = rowSOTOREMM.Item("ORDR_NO") & String.Empty
                        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ORDR_NO = '" & ORDR_NO & "'")
                            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & String.Empty
                            rowSOTOREMM.Item("PICK_AMT") += Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_AMT)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                            rowSOTOREMM.Item("PICK_QTY") += Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                        Next
                    Next
                End If
            End If
        End If

        'If manual_release Then
        '    TAC.SOCMAIN1.Allocation(Me, _
        '        False, _
        '        True, _
        '         IIf(selWHSE = "A", "", WHSE_CODE), _
        '         ORDR_GROUP_NO_sql, edi850cust, _
        '        SOTSUPP1, SOTDEMD1, TABLE_NAMEs, , (ROWs("SOTPARM1").Item("SO_PARM_ALLO_SEQ") & "" = "1"), "", "", False)

        'End If

        CommitTrans()

        If blnALLOCATION_ONLY Then

        Else
        End If

    End Sub

    Function Order_Release() As Boolean

        ASCMAIN1.Progress("Now Releasing Orders for Shipment", "")

        ' Customer Master : Bill-To

        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_PD_GRACE_DAYS, ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CREDIT_HOLD" & vbCrLf _
            & ", ARTCUST1.CUST_CRED_LIMIT_REV, ARTCUST1.CUST_CREDIT_RELEASE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE in (Select Distinct CUST_BILL_TO_CUST from " & ARTCUST1 & ")"
        Dim ARTCUST1_BT As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add Primary Key (CUST_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_PICK NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_OPEN NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_AMT_PICK_NOW NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_BALANCE NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_BALANCE_PD NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_BT & " Add CUST_HOLDS_CREDIT VARCHAR2(20)")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select ARTOPEN1.CUST_CODE" & vbCrLf _
            & ", Sum (ARTOPEN1.INV_BALANCE) CUST_BALANCE" & vbCrLf _
            & ", Sum (CASE WHEN INV_BALANCE > 0 and INV_TYPE = 'I' and ARTOPEN1.INV_DUE_DATE + NVL(ARTCUST1.CUST_PD_GRACE_DAYS,0) +1 < SYSDATE THEN ARTOPEN1.INV_BALANCE ELSE 0 END) CUST_BALANCE_PD" & vbCrLf _
            & "   from ARTOPEN1," & ARTCUST1_BT & " ARTCUST1" & vbCrLf _
            & "   where ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
            & "   group by ARTOPEN1.CUST_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTCUST1_BT & " ARTCUST1 Set" & vbCrLf _
            & "    CUST_BALANCE = R1.CUST_BALANCE" & vbCrLf _
            & "   ,CUST_BALANCE_PD = R1.CUST_BALANCE_PD" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select NVL(ARTCUST1.CUST_BILL_TO_CUST,SOTORDR0.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", Sum (SOTORDR0.ORDR_AMT_OPEN) CUST_AMT_OPEN" & vbCrLf _
            & ", Sum (SOTORDR0.ORDR_AMT_PICK) CUST_AMT_PICK" & vbCrLf _
            & "   from SOTORDR0, ARTCUST1" & vbCrLf _
            & "   where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   group by NVL(ARTCUST1.CUST_BILL_TO_CUST,SOTORDR0.CUST_CODE);" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTCUST1_BT & " ARTCUST1 Set" & vbCrLf _
            & "    CUST_AMT_PICK = R1.CUST_AMT_PICK" & vbCrLf _
            & "   ,CUST_AMT_OPEN = R1.CUST_AMT_OPEN" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'P' where NVL(CUST_BALANCE_PD,0) > 0 and NVL(CUST_CREDIT_RELEASE,'?') NOT IN ('I','N')")
        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'C' where NVL(CUST_CREDIT_HOLD,'0') = '1'")
        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_BT & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'Z' where (NVL(CUST_CREDIT_LIMIT,0) <=0 or CUST_CRED_LIMIT_REV is Null or CUST_CRED_LIMIT_REV < SYSDATE) and NVL(CUST_CREDIT_RELEASE,'?') <> 'N'")


        ' A: Hold Orders where 
        ' Customer is on Credit Hold, 
        ' or is Past Due, 
        ' or has no Credit Limit, 
        ' or Credit Limit has expired, 
        ' or Customer Aging is Past Due beyond Grace Period

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "" _
                & "Select CUST_CODE from " & ARTCUST1 & " where CUST_CREDIT_HOLD = '1'"
        Else
            ASCMAIN1.sql = "" _
                & "Select CUST_CODE from " & ARTCUST1 & " where CUST_CREDIT_HOLD = '1'" & vbCrLf _
                & " union " & vbCrLf _
                & "Select CUST_CODE from " & ARTCUST1_BT & " where CUST_HOLDS_CREDIT is not Null"
        End If

        ASCMAIN1.sql = "Begin Declare Cursor C1 is Select Distinct CUST_CODE from (" & ASCMAIN1.sql & ");" _
            & " Begin For R1 in C1 Loop" _
            & "  Update " & SOTORDR1 & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'A'" _
            & "   where CUST_CODE = R1.CUST_CODE or CUST_BILL_TO_CUST = R1.CUST_CODE or CUST_CREDIT_GROUP_CUST = R1.CUST_CODE;" _
            & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
            If ORDR_GROUP_NO_sql.Length = 0 Then
                ASCMAIN1.sql = "Update " & SOTORDR1 & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'B'" _
                       & " WHERE ORDR_NO IN " _
                       & "(SELECT ORDR_NO" _
                       & " FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & SOTORDR1 & ")" _
                       & " AND CUST_ORDR_CALL_B4_SHIPPING = '1')"
                ASCDATA1.ExecuteSQL()

                If chkECommerce.Checked Then
                    ' DO NOT HOLD ECOM ORDERS IF THEY HAVE INSTRUCTIONS
                Else
                    If Absx1.chkFor("CHKNO_SPECIAL_COMMENTS").Checked Then
                        ASCMAIN1.sql = "Update " & SOTORDR1 & " set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'H'" _
                           & " WHERE ORDR_NO IN " _
                           & "(SELECT ORDR_NO" _
                           & " FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & SOTORDR1 & ")" _
                           & " AND ORDR_SHIP_INSTR IS NOT NULL)"
                        ASCDATA1.ExecuteSQL()
                    End If
                End If


                ' Do Not Release International Shipments (Bill to or Ship To <> US) 3/20/2015
                ' Also Restrict to Only International shipments, No Domestic
                If chkINTL.Checked Then
                    ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_BATCH_NO = Null, ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'T'" _
                        & " WHERE ORDR_GROUP_NO IN " _
                        & "( SELECT SOTORDR1.ORDR_GROUP_NO" _
                        & " FROM " & SOTORDR1 & " SOTORDR1, SOTORDR5" _
                        & " WHERE SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO" _
                        & " AND NVL(UPPER(SOTORDR5.CUST_COUNTRY), 'US') NOT IN ('US', 'USA'))"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                ElseIf chkINTL_ONLY.Checked Then
                    ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_BATCH_NO = Null, ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'U'" _
                        & " WHERE ORDR_GROUP_NO IN " _
                        & "( SELECT SOTORDR1.ORDR_GROUP_NO" _
                        & " FROM " & SOTORDR1 & " SOTORDR1, SOTORDR5" _
                        & " WHERE SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO" _
                        & " AND NVL(UPPER(SOTORDR5.CUST_COUNTRY), 'US') IN ('US', 'USA'))"
                End If
            End If
        End If

        ' Down Below, Hold Orders if Customer is Over Credit Limit, 

        ASCMAIN1.sql = "Select * from " & ARTCUST1_BT
        Create_TDA(dst.Tables.Add("ARTCUST1_BT"), ARTCUST1_BT, "**", 0)
        Fill_Records("ARTCUST1_BT")

        Create_Relation("ARTCUST1_BT", "ARTCUST1", "CUST_CODE", "CUST_BILL_TO_CUST")
        ' NEED TO GET THIS SET UP FOR RGI - ARTCUST1 DOES NOT HAVE THESE CHILD FIELDS RIGHT NOW
        'With dst.Tables("ARTCUST1_BT")
        '    .Columns("CUST_AMT_PICK").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_PICK)"
        '    .Columns("CUST_AMT_OPEN").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_OPEN)"
        '    .Columns("CUST_AMT_PICK_NOW").Expression = "SUM(CHILD(ARTCUST1_BT_ARTCUST1).CUST_AMT_PICK_NOW)"
        'End With

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else
            ASCMAIN1.sql = " Select ORDR_NO" _
                 & " from TATTERM1," & SOTORDR1 & " SOTORDR1, ARTCUST1, TATTERM1 TATTERMC " _
                 & " where SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE" _
                 & " and TATTERM1.TERM_CODE (+) = SOTORDR1.TERM_CODE" _
                 & " and TATTERMC.TERM_CODE (+) = ARTCUST1.TERM_CODE" _
                 & " and (NVL(TATTERM1.TERM_TYPE,'?') = 'R' OR NVL(TATTERMC.TERM_TYPE,'?') = 'R')"

            ' Regency will Queue the Credit Card Auths Up and Allow Customer Servce to process in a batch
            If Not (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
                ASCMAIN1.sql &= " UNION " _
                    & " Select ORDR_NO" & vbCrLf _
                    & " from TATTERM1," & SOTORDR1 & " SOTORDR1" _
                    & " where TATTERM1.TERM_CODE = SOTORDR1.TERM_CODE" & vbCrLf _
                    & "   and TATTERM1.TERM_TYPE = 'D'" & vbCrLf _
                    & "   and (SOTORDR1.CCPA_NO is Null or SOTORDR1.CC_TRANS_ID is Null)"
            End If

            'ASCMAIN1.sql = "Select ORDR_NO" & vbCrLf _
            '    & " from TATTERM1," & SOTORDR1 & " SOTORDR1" _
            '    & " where TATTERM1.TERM_CODE = SOTORDR1.TERM_CODE" & vbCrLf _
            '    & "   and TATTERM1.TERM_TYPE = 'D'" & vbCrLf _
            '    & "   and (SOTORDR1.CCPA_NO is Null or SOTORDR1.CC_TRANS_ID is Null)" _
            '    & " Union " _
            '    & " Select ORDR_NO" _
            '    & " from TATTERM1," & SOTORDR1 & " SOTORDR1, ARTCUST1, TATTERM1 TATTERMC " _
            '    & " where SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE" _
            '    & " and TATTERM1.TERM_CODE (+) = SOTORDR1.TERM_CODE" _
            '    & " and TATTERMC.TERM_CODE (+) = ARTCUST1.TERM_CODE" _
            '    & " and (NVL(TATTERM1.TERM_TYPE,'?') = 'R' OR NVL(TATTERMC.TERM_TYPE,'?') = 'R')"

            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R'" _
                & "   where ORDR_NO in (" & ASCMAIN1.sql & ")"
            ASCDATA1.ExecuteSQL()
        End If

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then

            'ASCMAIN1.sql = "Select ORDR_GROUP_NO FROM SOTORDR0 WHERE SUBSTR(ORDR_GROUP_NO,5,1) = '7'" _
            '      & " AND CUST_CODE IN (SELECT DISTINCT CUST_CODE FROM EDTTRPM1) AND ORDR_CNT_OPEN <> 0"
            'ASCMAIN1.sql = "Update " & SOTORDR1 _
            '      & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || '7'" _
            '      & "   where ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")"
            'ASCDATA1.ExecuteSQL()


            ' Update Credit Authorizations with latest information from EDI Tables

            TAC.SOCMAIN1.Update_Credit_Authorizations()

            ' C: Hold Orders where Credit Authorization is not OK

            ASCMAIN1.sql = "" _
                     & "Select X.*, SOTAUTH1.ORDR_CRED_CLR_BY, SOTAUTH1.ORDR_CRED_CLR_AUTH" & vbCrLf _
                     & ", ARTCUST1.CUST_CREDIT_HOLD, NVL(ARTCUST1.CUST_CREDIT_RELEASE,'M') CUST_CREDIT_RELEASE" & vbCrLf _
                     & ", ARTCUST1.CUST_FACTOR_IND CUST_FACTOR_IND_CUST" & vbCrLf _
                     & " FROM SOTAUTH1,ARTCUST1," & vbCrLf _
                     & "(Select ORDR_GROUP_NO, CUST_BILL_TO_CUST CUST_CODE, MAX (NVL(CUST_FACTOR_IND,'0')) CUST_FACTOR_IND " & vbCrLf _
                     & " from " & SOTORDR1 & " group by ORDR_GROUP_NO, CUST_BILL_TO_CUST) X" & vbCrLf _
                     & " where SOTAUTH1.ORDR_GROUP_NO (+) = X.ORDR_GROUP_NO" & vbCrLf _
                     & "   and ARTCUST1.CUST_CODE = X.CUST_CODE"

            Dim SOTCREDX As String = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCREDX & " Add Primary Key (ORDR_GROUP_NO)")


            ' THERE ARE ORDERS IN SOTORDR1 THAT MIGHT GO BEYOND THE SCOPE OF THE RELEASE (LIKE IF YOU PICKED A SPECIFIC CUSTOMER)
            ' SO TO STAMP ALL OF THE ORDERS WITH NO CREDIT CHECK BEFORE THEY ARE RELEASED IS DANGEROUS - WE MIGHT SWITCH THE CUSTOMER SETTING BEFORE THE DAY THAT THE ORDER RELEASES

            'ASCMAIN1.sql = "Select ORDR_GROUP_NO from " & SOTCREDX & "where CUST_CREDIT_RLEASE = 'N' minus Select ORDR_GROUP_NO from SOTAUTH1"
            'ASCMAIN1.sql = "Insert into SOTAUTH1 (ORDR_GROUP_NO) " & ASCMAIN1.sql
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "" _
            '    & "Begin" & vbCrLf _
            '    & " Declare Cursor C1 is Select * from " & SOTCREDX & " where CUST_CREDIT_RELEASE = 'N';" & vbCrLf _
            '    & " Begin" & vbCrLf _
            '    & "  For R1 in C1 Loop" & vbCrLf _
            '    & "   Update SOTAUTH1 " & vbCrLf _
            '    & "    Set ORDR_CRED_AUTH = 'A'" & vbCrLf _
            '    & "    , ORDR_CRED_CLR_AUTH_TYPE = 'N'" & vbCrLf _
            '    & "    , LAST_OPER = '" & ASCMAIN1.USER_ID & "', LAST_DATE = SYSDATE" & vbCrLf _
            '    & "    , CRED_CODE = Null" & vbCrLf _
            '    & "   where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
            '    & "  End Loop;" & vbCrLf _
            '    & " End;" & vbCrLf _
            '    & "End;"
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select ORDR_GROUP_NO from " & SOTCREDX _
                & " where NVL(ORDR_CRED_CLR_AUTH,'X') <> 'A'" & vbCrLf _
                & "  and (CUST_FACTOR_IND = '1' or CUST_CREDIT_RELEASE = 'M')" & vbCrLf _
                & "  and CUST_CREDIT_RELEASE <> 'N'"
            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'C'" _
                & "   where ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")" & vbCrLf _
                & "     and ORDR_TYPE_CODE <> 'XFR'"
            ASCDATA1.ExecuteSQL()

        End If

        ' P: Hold Orders where Pre-Pack Qty is not evenly divisible by Inner Pack, or Styles are not the same for all PPK Lines

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            'ASCMAIN1.sql = "Select Distinct ORDR_GROUP_NO from " & SOTORDR1 _
            '    & " where CUST_CODE in (Select CUST_CODE from EDTSLSP1 where EDI_SLN_TOT_IND = '1')"

            ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'P'" & vbCrLf _
                & " where ORDR_NO in (" & vbCrLf _
                & "Select Distinct ORDR_NO from (" & vbCrLf _
                    & "Select SOTORDR2.ORDR_NO, SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                    & ", Count (*) PPKS" & vbCrLf _
                    & ", Min (ICTSTYL1.INNER_PACK_QTY) PPK1" & vbCrLf _
                    & ", Max (ICTSTYL1.INNER_PACK_QTY) PPK2" & vbCrLf _
                    & ", Min (SOTORDR2.STYLE_CODE) STYLE1" & vbCrLf _
                    & ", Max (SOTORDR2.STYLE_CODE) STYLE2" & vbCrLf _
                    & ", Min (SOTORDR2.ORDR_UNIT_PRICE) PRICE1" & vbCrLf _
                    & ", Max (SOTORDR2.ORDR_UNIT_PRICE) PRICE2" & vbCrLf _
                    & ", Sum (SOTORDR2.ORDR_QTY_OPEN) QTY" & vbCrLf _
                    & "from SOTORDR1,SOTORDR2,ICTSTYL1" & vbCrLf _
                    & "where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                    & "and SOTORDR1.ORDR_SOURCE = 'E'" & vbCrLf _
                    & "and SOTORDR1.ORDR_GROUP_NO in (" & vbCrLf _
                    & "Select Distinct ORDR_GROUP_NO from SOTORDR1 " & vbCrLf _
                    & "where CUST_CODE in (Select CUST_CODE from EDTSLSP1 where EDI_SLN_TOT_IND = '1')" & vbCrLf _
                    & "and ORDR_STATUS = 'O'" & vbCrLf _
                    & ")" & vbCrLf _
                    & "group by SOTORDR2.ORDR_NO, SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                    & "having Count (*) > 1" & vbCrLf _
                    & ") where STYLE1 <> STYLE2 or  NVL(PRICE1,0) <> NVL(PRICE2,0) or NVL(PPK1,0) <> NVL(PPK2,0) or NVL(PPK1,0) = 0 or Mod(QTY,PPK1) <> 0" & vbCrLf _
                & ")" _
                & " and ORDR_NO <> '0000886345'"
            ASCDATA1.ExecuteSQL()

            ' NOTE HARDCODED LB ORDER ABOVE 0000878114 - THIS IS TO ALLOW ONE TO GET BY AND SEE WHAT IT DOES - SEE LS EMAILS 08/31/15

        End If

        ' S: Hold Orders where Customer is on Sales Hold
        ' S: Hold all Orders for large customers such as SEARS or KMART unless they were explicitly Requested, or specific groups were selected

        If ORDR_GROUP_NO_sql = "" Then
            ASCMAIN1.sql = "Update " & ARTCUST1 _
                & " Set CUST_SALES_HOLD = '1' " _
                & " where CUST_REL_EXPLICITLY = '1'" _
                & "   and CUST_CODE NOT in (" & CUST_CODE_sql & ")"
        End If

        ' O: Hold Orders where Specific Order was indicated to be held

        ASCMAIN1.sql = "Update " & SOTORDR1 _
            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'O'" _
            & " where ORDR_HOLD = '1'"
        ASCDATA1.ExecuteSQL()


        ' Set ORDR_RELEASE flag to Ship Style/Color Short if CUST/STYLE Parameter Table indicates ORDR_RELEASE = '1'

        '    sql = "Update SOWORDR2,SOWORDR1,ICWSTYC1"
        '    & " Set SOWORDR2.ORDR_RELEASE = '1'"
        '    & " where SOWORDR2.ORDR_QTY_OPEN <> SOWORDR2.ORDR_QTY_ALLO"
        '    & "   and SOWORDR2.ORDR_RELEASE is Null"
        '    & "   and SOWORDR2.RANGE_STYLE_CODE is Null"
        '    & "   and SOWORDR1.ORDR_NO = SOWORDR2.ORDR_NO"
        '    & "   and ICWSTYC1.STYLE_CODE = SOWORDR2.STYLE_CODE"
        '    & "   and ICWSTYC1.COLOR_CODE = SOWORDR2.COLOR_CODE"
        '    & "   and ICWSTYC1.ORDR_RELEASE = '1'"

        ' Set ORDR_RELEASE flag to Ship Style/Color Short if CUST/STYLE/COLOR Parameter Table indicates ORDR_RELEASE = '1'

        '    sql = "Update SOWORDR2,SOWORDR1,SOWCSTP1"
        '    & " Set SOWORDR2.ORDR_RELEASE = '1'"
        '    & " where SOWORDR2.ORDR_QTY_OPEN <> SOWORDR2.ORDR_QTY_ALLO"
        '    & "   and SOWORDR2.ORDR_RELEASE is Null"
        '    & "   and SOWORDR2.RANGE_STYLE_CODE is Null"
        '    & "   and SOWORDR1.ORDR_NO = SOWORDR2.ORDR_NO"
        '    & "   and SOWCSTP1.CUST_CODE = SOWORDR2.CUST_CODE"
        '    & "   and SOWCSTP1.STYLE_CODE = SOWORDR2.STYLE_CODE"
        '    & "   and SOWCSTP1.COLOR_CODE = SOWORDR2.COLOR_CODE"
        '    & "   and SOWCSTP1.ORDR_RELEASE = '1'"

        ' I, E, F
        Inventory_Shortages()

        ' N: Hold all orders whose lines total 0 qty allocated

        '& "   and (SOWORDR2.ORDR_RELEASE IS NULL OR (SOWORDR2.ORDR_RELEASE <> 'S' AND SOWORDR2.ORDR_RELEASE <> 'C')) "

        ' If ARTCUST1.CUST_SHIP_COMPLETE_DETAIL AND SOTORDR2.ORDR_QTY <> SOTORDR2.ORDR_QTY_ALLO_CUR then the sales order cannot be released
        ASCMAIN1.sql = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
                & " from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
                & " where NVL(SOTORDR2.ORDR_QTY,0) <> NVL(SOTORDR2.ORDR_QTY_ALLO_CUR,0) AND ORDR_STATUS IN ('O', 'P')"
        ' & "   and SOTORDR2.ORDR_RELEASE is Null"

        ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'M'" & vbCrLf _
            & " where ORDR_NO in (" & ASCMAIN1.sql & ") AND CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE CUST_SHIP_COMPLETE_DETAIL = '1')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO" _
            & " from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Null" _
            & " group by SOTORDR1.ORDR_NO" _
            & " having SUM (SOTORDR2.ORDR_QTY_ALLO_CUR) = 0"
        Dim TBL As DataTable = ASCDATA1.GetDataTable()
        ASCMAIN1.sql = "Update " & SOTORDR1 _
            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'N'" _
            & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ' B: Hold all orders whose lines falls below Minimum Order Amounts. This is the New Feature for Maurice.

            ASCMAIN1.sql = "Select DISTINCT SOTORDR1.ORDR_NO" & vbCrLf _
                & " from " & SOTORDR1 & " SOTORDR1, " & SOTORDR2 & " SOTORDR2, ARTCUST2" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_ADDR_TYPE = SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
                & "   and ARTCUST2.CUST_ADDR_CODE = Decode(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO)" & vbCrLf _
                & "   and (NVL(MIN_ORDR_QTY,0) <> 0" & vbCrLf _
                & "     or NVL(MIN_ORDR_AMT,0) <> 0" & vbCrLf _
                & "     or NVL(MIN_STYLE_QTY,0) <> 0" & vbCrLf _
                & "     or NVL(MIN_STYLE_AMT,0) <> 0)" & vbCrLf _
                & " group by SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.CUST_DC_NO" & vbCrLf _
                & ", ARTCUST2.MIN_ORDR_QTY, ARTCUST2.MIN_ORDR_AMT, ARTCUST2.MIN_STYLE_QTY, ARTCUST2.MIN_STYLE_AMT" & vbCrLf _
                & " having" & vbCrLf _
                & "     Sum (SOTORDR2.ORDR_QTY_OPEN) < ARTCUST2.MIN_ORDR_QTY" & vbCrLf _
                & "      or" & vbCrLf _
                & "     Sum (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) < ARTCUST2.MIN_ORDR_AMT" & vbCrLf _
                & "      or" & vbCrLf _
                & "     Min (SOTORDR2.ORDR_QTY_OPEN) < ARTCUST2.MIN_STYLE_QTY" & vbCrLf _
                & "      or" & vbCrLf _
                & "     Min (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) < ARTCUST2.MIN_STYLE_AMT"

            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'B'" _
                & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
            ASCDATA1.ExecuteSQL()
        End If

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else

            ' Check if Over Credit Limit

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C0 is Select * from " & ARTCUST1_BT & " for Update;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R0 in C0 Loop" & vbCrLf _
                & "   Begin" & vbCrLf _
                & "    Declare" & vbCrLf _
                & "     CUST_AMT_ALLO_NOW NUMBER (13,2);" & vbCrLf _
                & "     Cursor C1 is" & vbCrLf _
                & "      Select SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & "      , SUM (NVL(SOTORDR2.ORDR_QTY_ALLO,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)) ORDR_AMT_ALLO" & vbCrLf _
                & "       from " & SOTORDR1 & " SOTORDR1," & SOTORDR2 & " SOTORDR2" & vbCrLf _
                & "      where SOTORDR1.CUST_BILL_TO_CUST = R0.CUST_CODE" & vbCrLf _
                & "        and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "      group by SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & "      order by SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_GROUP_NO;" & vbCrLf _
                & "    Begin " & vbCrLf _
                & "     CUST_AMT_ALLO_NOW := 0;" & vbCrLf _
                & "     For R1 in C1 Loop" & vbCrLf _
                & "      CUST_AMT_ALLO_NOW := CUST_AMT_ALLO_NOW + NVL(R1.ORDR_AMT_ALLO,0);" & vbCrLf _
                & "      Update " & ARTCUST1_BT & " Set CUST_AMT_PICK_NOW = CUST_AMT_ALLO_NOW where CURRENT of C0;" & vbCrLf _
                & "      If NVL(R0.CUST_CREDIT_LIMIT,0) < NVL(R0.CUST_AMT_PICK,0) + CUST_AMT_ALLO_NOW Then" & vbCrLf _
                & "       Update " & SOTORDR1 & vbCrLf _
                & "        Set ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES || 'L'" & vbCrLf _
                & "        where ORDR_GROUP_NO = R1.ORDR_GROUP_NO;" & vbCrLf _
                & "      End If;" & vbCrLf _
                & "     End Loop;" & vbCrLf _
                & "    End;" & vbCrLf _
                & "   End;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

        End If


        ' C: Integrity Checks

        Dim Records_Updated As Int64 = 0
        Dim ok_to_release As Boolean = True

        ASCMAIN1.sql = "Select ORDR_NO from " & SOTORDR1 & " where ORDR_ADDR_TYPE_ST = 'DC' and CUST_DC_NO is Null"
        ASCMAIN1.sql = "Update " & SOTORDR1 _
            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'C'" _
            & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
        'If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
        '    If Format(Now, "yyyyMMdd") < "20130320" Then
        '        ASCMAIN1.sql &= " and 1<>1"
        '    End If
        'End If
        Records_Updated = ASCDATA1.ExecuteSQL()
        If Records_Updated <> 0 Then ok_to_release = False

        If Records_Updated <> 0 Then
            Dim f As New ASFMSGBF
            Dim tblNoDC As DataTable = ASCDATA1.GetDataTable("Select ORDR_NO, CUST_CODE, CUST_NAME, ORDR_CUST_PO, ORDR_GROUP_NO, CUST_STORE_NO, ORDR_ADDR_TYPE_ST, CUST_DC_NO from " & SOTORDR1 & " where ORDR_ADDR_TYPE_ST = 'DC' and CUST_DC_NO is Null")
            f.Show_grd(tblNoDC, Me, "Orders shipping to DC with No DC")
        End If

        ' X: Hold all orders within a group if both of the following are true:
        '      1) any single order in the group is being held from release
        '      2) the customer is flagged with CUST_SHIP_COMPLETE - DISABLED 2/27/01 BY WJZ - 3 STORES OUT OF 28 NATIO10 RELEASED BECAUSE CUSTOMER WAS NOT MARKED AS SHIP_COMPLETE

        ASCMAIN1.sql = "Select Distinct SOTORDR1.ORDR_GROUP_NO" _
            & " from " & SOTORDR1 & " SOTORDR1," & ARTCUST1 & " ARTCUST1" _
            & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" _
            & " and SOTORDR1.ORDR_REL_HOLD_CODES is Not Null"
        ' ****************************** VAN SPECIFIC SEE #2 ABOVE ****************************
        '& " and ARTCUST1.CUST_SHIP_COMPLETE = '1' "
        ' ****************************** VAN SPECIFIC SEE #2 ABOVE ****************************
        Dim TBL2 As DataTable = ASCDATA1.GetDataTable
        ASCMAIN1.sql = "Update " & SOTORDR1 _
            & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'X'" _
            & " where ORDR_REL_HOLD_CODES is Null and ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")"
        ASCDATA1.ExecuteSQL()

        ' Clear out any holds against any orders which belong to force-picked groups
        ' Identify all Orders which should be part of this Release by setting ORDR_REL_BATCH_NO to XNO



        If blnFORCE_PICK Or manual_release Then

            RELEASE_SQL = " where ORDR_GROUP_NO in (" & ORDR_GROUP_NO_sql & ")"
            ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 " _
                & "Set ORDR_REL_HOLD_CODES = null, ORDR_REL_BATCH_NO = '" & XNO & "'" _
                & RELEASE_SQL

            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_BATCH_NO = '" & XNO & "'" _
                        '& "   and ORDR_REL_HOLD_CODES is Null " ' THIS WAY WE MARK ALL ORDERS THAT WERE ATTEMPTED TO BE RELEASED, WHETHER SUCCESSFUL OR NOT
            RELEASE_SQL = " where ORDR_SHIP_DATE <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'"


            If SALES_DIVISION_CODE_sql <> "" Then
                RELEASE_SQL &= " and SALES_DIVISION_CODE in (" & SALES_DIVISION_CODE_sql & ")"
            End If

            If CUST_CODE_sql <> "" Then
                RELEASE_SQL &= " and CUST_CODE in (" & CUST_CODE_sql & ")"
            End If

            If ORDR_GROUP_NO_sql <> "" Then
                If blnORDR_GROUP_NO_sql_NOT Then
                    RELEASE_SQL &= " and ORDR_GROUP_NO not in (" & ORDR_GROUP_NO_sql & ")"
                Else
                    RELEASE_SQL &= " and ORDR_GROUP_NO in (" & ORDR_GROUP_NO_sql & ")"
                End If
            End If

            If TERM_CODE_sql <> "" Then
                RELEASE_SQL &= " and TERM_CODE in (" & TERM_CODE_sql & ")"
            End If

            If blnMANUAL_ONLY Then

                If ASCMAIN1.CLIENT = "RGI" Then
                    RELEASE_SQL &= " and ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDRG_manual & ")"
                End If

                'Dim sqlU As String = ""
                'For Each row As DataRow In DirectCast(grdSOTORDRU.DataSource, DataTable).Select("SEL='1'")
                '    sqlU &= ",'" & row.Item("USER_ID") & "'"
                'Next
                'If sqlU <> "" Then
                '    sqlU = " and SOTORDRG.ORDR_REL_SHORT_OPER in (" & Mid(sqlU, 2) & ")"
                'End If

                'RELEASE_SQL &= " and ORDR_GROUP_NO in " _
                '    & " (Select SOTORDRG.ORDR_GROUP_NO from SOTORDRG," & SOTORDR1 & " SOTORDR1" _
                '    & "   where SOTORDRG.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO and SOTORDRG.ORDR_REL_SHORT = '1'" & sqlU & ")"
            End If


            ASCMAIN1.sql &= RELEASE_SQL
            ASCDATA1.ExecuteSQL()

            Release_Exceptions()
            ASCMAIN1.Progress(String.Empty, String.Empty)

        End If

        ' 03/21/2019 - Try Credit Card auth here not after release. Mark with Hold Code R if it fails
        If ASCMAIN1.CLIENT = "RGI" Then

            ASCMAIN1.sql = "Select * from " & SOTORDR1 & " where ORDR_REL_BATCH_NO = '" & XNO & "' and ORDR_REL_HOLD_CODES is Null and TERM_CODE IN (SELECT TERM_CODE FROM TATTERM1 WHERE TERM_TYPE = 'D')"
            Dim tblCC As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & SOTORDR2 & " WHERE ORDR_NO IN (" & ASCMAIN1.sql.Replace("*", "ORDR_NO") & ")"
            Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR2")

            For Each rowSOTORDR1_rel As DataRow In tblCC.Select("", "CUST_CODE, ORDR_NO")
                Dim ORDR_NO As String = rowSOTORDR1_rel.Item("ORDR_NO") & String.Empty

                Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_QTY_OPEN <> 0"
                Dim totalSales As Decimal = 0

                For Each rowSOTORDR2_rel As DataRow In tblSOTORDR2.Select(sqlw, "ORDR_LNO")

                    Dim qA As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") & "")
                    Dim qO As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") & "")

                    If chkForcePick.Checked Then
                        qA = qO
                    End If

                    totalSales += (qA * Val(rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE") & String.Empty))
                Next

                If totalSales > 0 Then
                    AuthorizeCreditCards(ORDR_NO, totalSales)
                End If
            Next
        End If

        ASCMAIN1.sql = "Select Distinct SOTORDRG.ORDR_GROUP_NO" _
            & " from " & SOTORDR1 & " SOTORDR1, SOTORDRG" _
            & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" _
            & "   and SOTORDRG.ORDR_REL_SHORT = '1'" _
            & "   and SOTORDR1.ORDR_REL_BATCH_NO = '" & XNO & "'" _
            & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Null"
        ASCMAIN1.sql = "Update SOTORDRG" _
            & " Set ORDR_REL_SHORT = '0', ORDR_REL_BATCH_NO = '" & XNO & "'" _
            & " where ORDR_GROUP_NO in (" & ASCMAIN1.sql & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1 set ORDR_REL_HOLD_CODES = null" _
            & " where ORDR_REL_BATCH_NO is Null"
        ASCDATA1.ExecuteSQL()

        Return ok_to_release

    End Function

    Private Sub Release_Exceptions()

        ' New Order Release Hold Code
        '   T - International Ship To / Bill To

        ' Special things for Regency Int'l
        If Not (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
            Exit Sub
        End If

        ' Here during testing
        'If Not (ASCMAIN1.DBS_COMPANY <> ASCMAIN1.DBS_SERVER) Then
        '    Exit Sub
        'End If

        ' If force pick tickets then no need to do any work.
        If chkForcePick.Checked Then
            Exit Sub
        End If

        If Not chkMerge.Checked Then
            Exit Sub
        End If

        If ORDR_GROUP_NO_sql.Length > 0 Then
            Exit Sub
        End If

        ' if the parameter is null then get out of here. 
        Dim SO_PARM_REL_PCT As Int16 = Absx1.numFor("NUMREL_PCT").Value
        If SO_PARM_REL_PCT <= 0 Then
            Exit Sub
        End If

        ASCMAIN1.Progress("Merging Sales Orders", String.Empty)
        Dim rowARTCUST1 As DataRow = Nothing

        Dim wkGroup As String = String.Empty
        Dim SO_PARM_GROUP_DAYS As Int32 = Val(ROWs("SOTPARM1").Item("SO_PARM_GROUP_DAYS") & String.Empty)

        ' If 0 then default to 4 * 4 * 6 pallet.
        Dim palletCube As Int16 = Val(ROWs("SOTPARM1").Item("SO_PARM_REL_CUBE") & String.Empty)
        If palletCube <= 0 Then
            palletCube = 4 * 4 * 6
        End If

        If Not dst.Tables.Contains("SOTOREMM") Then
            ASCMAIN1.sql = "Select ORDR_NO, WHSE_CODE, ORDR_DATE, CUST_CODE, CUST_NAME, ORDR_CUST_PO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_GROUP_NO,ORDR_REL_HOLD_CODES from SOTORDR1"
            Create_TDA(dst.Tables.Add, "SOTOREMM", ASCMAIN1.sql, 0, False, "", 1)
            With dst.Tables("SOTOREMM")
                .Columns.Add("PICK_AMT", GetType(System.Decimal))
                .Columns.Add("PICK_QTY", GetType(System.Int32))
            End With
        Else
            dst.Tables("SOTOREMM").Rows.Clear()
        End If

        Dim ordrRelHoldCodesToProcess As String = "ORDR_REL_HOLD_CODES IS NULL OR ORDR_REL_HOLD_CODES in ('','I','IE','IEF','IF')"
        Dim oldShipmentShipDate As Date = DateAdd(DateInterval.Day, SO_PARM_GROUP_DAYS * -1, DateTime.Now)

        Dim tblReleaseable As String = String.Empty
        ' Clear the Order Group for those orders that can be grouped
        sql = " SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.ORDR_NO,"
        sql &= " SUM(NVL(SOTORDR2.ORDR_QTY_ALLO_CUR, 0)) ORDR_QTY_ALLO_CUR"
        sql &= " FROM " & SOTORDR1 & " SOTORDR1, " & SOTORDR2 & " SOTORDR2"
        sql &= " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
        sql &= " AND (" & ordrRelHoldCodesToProcess & ")"
        sql &= " AND NVL(ORDR_SHIP_COMPLETE, 0) = '0'"
        sql &= " AND SOTORDR1.ORDR_SHIP_DATE <= '" & SHIP_BY_DATE.ToString("dd-MMM-yyyy") & "'"
        sql &= " GROUP BY SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.ORDR_NO"
        sql &= " HAVING SUM(NVL(SOTORDR2.ORDR_QTY_ALLO_CUR, 0)) > 0 "
        tblReleaseable = ASCMAIN1.Temp_Table(sql)

        ' Combine orders by (Ship Date / Cancle Date) or PO (default) 3/20/2015
        If Absx1.chkFor("CHKCOMBINE").Checked Then

            wkGroup = ASCMAIN1.Temp_Table("SELECT ORDR_NO, ORDR_GROUP_NO FROM " & SOTORDR1)

            ASCMAIN1.sql = "UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = NULL WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & tblReleaseable & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ' Group cleared order groups by  CUST_CODE, CUST_STORE_NO, ORDR_ADDR_TYPE_ST
            ASCMAIN1.sql = "Begin Declare Cursor C1 Is SELECT CUST_CODE, CUST_STORE_NO, ORDR_ADDR_TYPE_ST, ROWNUM SEQ" _
                    & "       FROM (Select CUST_CODE, CUST_STORE_NO, ORDR_ADDR_TYPE_ST, COUNT(*) " _
                    & " FROM " & SOTORDR1 _
                    & " WHERE ORDR_GROUP_NO IS NULL" _
                    & " GROUP BY CUST_CODE, CUST_STORE_NO, ORDR_ADDR_TYPE_ST" _
                    & " HAVING COUNT(*) > 1);" _
                    & " Begin For R1 in C1 Loop " _
                    & "     UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = LPAD(R1.SEQ, 10, '9')" _
                    & "       WHERE CUST_CODE = R1.CUST_CODE" _
                    & "       AND CUST_STORE_NO = R1.CUST_STORE_NO" _
                    & "       AND ORDR_ADDR_TYPE_ST = R1.ORDR_ADDR_TYPE_ST" _
                    & "       AND ORDR_GROUP_NO IS NULL;" _
                    & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ' Number of Shipments for sales order as per customer master (CUST_SHIP_COMPLETE): 0 = Multiple Shipments, 1 = Ship Complete, 2 Max 2 Shipments.
            ' Ship Complete seting on the sales order overrides Ship Complete setting on Customer Master.
            ASCMAIN1.sql = "Begin Declare Cursor C1 Is SELECT * FROM (SELECT " & SOTORDR1 & ".ORDR_NO, NVL(ARTCUST1.CUST_SHIP_COMPLETE, '0') CUST_SHIP_COMPLETE,  " _
                     & " (SELECT COUNT(PICK_NO) FROM SOTPICK1 WHERE PICK_STATUS = 'F' AND ORDR_NO = " & SOTORDR1 & ".ORDR_NO) NUM_SHIP" _
                     & " FROM " & SOTORDR1 & ", ARTCUST1" _
                     & " WHERE " & SOTORDR1 & ".CUST_CODE = ARTCUST1.CUST_CODE) WHERE NVL(NUM_SHIP, 0) > 0 AND CUST_SHIP_COMPLETE = '2';" _
                     & " Begin For R1 in C1 Loop " _
                     & "     UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_NO = R1.ORDR_NO;" _
                     & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            'CUST_ALLOW_BACKORDER
            ASCMAIN1.sql = "Begin Declare Cursor C1 Is SELECT " & SOTORDR1 & ".ORDR_NO" _
                     & " FROM " & SOTORDR1 & ", ARTCUST1" _
                     & " WHERE " & SOTORDR1 & ".CUST_CODE = ARTCUST1.CUST_CODE and NVL(CUST_ALLOW_BACKORDER, '0') = '0';" _
                     & " Begin For R1 in C1 Loop " _
                     & "     UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_NO = R1.ORDR_NO;" _
                     & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_GROUP_NO IS NULL"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_SHIP_COMPLETE = '1'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_HOLD = '1'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_STATUS NOT IN ('O', 'P')"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ' New rule as of 6/19/2015
            ASCMAIN1.sql = "UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_NO NOT IN " _
                & " (SELECT ORDR_NO from " & SOTORDR1 & " WHERE " & ordrRelHoldCodesToProcess & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = " UPDATE " & SOTORDR1 & " SET ORDR_GROUP_NO = ORDR_NO WHERE ORDR_GROUP_NO IS NULL"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        End If

        ' Ranges by %% to see if we can release partial shipments.
        ' Hierarchy.
        ' Customer Settings
        ' State Settings
        ' Sales Order Processing Settings

        ' Load Sales Order Processing Percentage
        ASCMAIN1.sql = "Select Distinct SOTORDR1.CUST_CODE, TATSHIPP.SHIPMENT_AMT, TATSHIPP.SHIPMENT_PERC, UPPER(SOTORDR5.CUST_STATE) CUST_STATE" _
            & " FROM " & SOTORDR1 & " SOTORDR1, SOTORDR5, TATSHIPP" _
            & " WHERE SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO (+)  AND SOTORDR5.CUST_ADDR_TYPE = 'ST'" _
            & " AND TATSHIPP.TABLE_NAME = 'SOTPARM1' AND TATSHIPP.KEY_VALUE = 'Z'"
        Dim TATSHIPP As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ' This allows for missing SOTORDR5 / ST record - In my testing I came across SOTORDR1s without SOTORDR5 ST records.
        ASCMAIN1.sql = "INSERT INTO " & TATSHIPP & " Select Distinct SOTORDR1.CUST_CODE, TATSHIPP.SHIPMENT_AMT, TATSHIPP.SHIPMENT_PERC, NULL CUST_STATE" _
            & " FROM " & SOTORDR1 & " SOTORDR1, TATSHIPP" _
            & " WHERE TATSHIPP.TABLE_NAME = 'SOTPARM1' AND TATSHIPP.KEY_VALUE = 'Z'" _
            & " AND SOTORDR1.CUST_CODE NOT IN (SELECT CUST_CODE FROM " & TATSHIPP & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "INSERT INTO " & TATSHIPP _
            & " SELECT DISTINCT CUST_CODE, 0 SHIPMENT_AMT," & SO_PARM_REL_PCT & " SHIPMENT_PERC, CUST_STATE FROM " & TATSHIPP
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ' Update From State Percentages
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS Select * from TATSTATE where STATE_REL_PCT > 0;" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & "  UPDATE " & TATSHIPP & " SET SHIPMENT_PERC = R1.STATE_REL_PCT WHERE SHIPMENT_AMT = 0 AND CUST_STATE = R1.STATE_CODE; " _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT DISTINCT CUST_CODE, CUST_STATE FROM " & TATSHIPP & " WHERE CUST_STATE IN (Select KEY_VALUE from TATSHIPP WHERE TABLE_NAME = 'TATSTATE');" _
             & " BEGIN FOR R1 IN C1 LOOP " _
             & "    DELETE FROM " & TATSHIPP & " WHERE SHIPMENT_AMT > 0 AND CUST_CODE = R1.CUST_CODE AND CUST_STATE = R1.CUST_STATE;" _
             & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "INSERT INTO " & TATSHIPP _
            & " SELECT DISTINCT CUST_CODE, TATSHIPP.SHIPMENT_AMT, TATSHIPP.SHIPMENT_PERC, CUST_STATE" _
            & " FROM " & TATSHIPP & ", TATSHIPP" _
            & " WHERE TABLE_NAME = 'TATSTATE' AND " & TATSHIPP & ".CUST_STATE = TATSHIPP.KEY_VALUE"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ' Update From Customer Master Settings
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS Select * from ARTCUST1 where CUST_SHIP_PCT_ORDER > 0;" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & "  UPDATE " & TATSHIPP & " SET SHIPMENT_PERC = R1.CUST_SHIP_PCT_ORDER WHERE SHIPMENT_AMT = 0 AND CUST_CODE = R1.CUST_CODE; " _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT DISTINCT CUST_CODE FROM " & TATSHIPP & " WHERE CUST_CODE IN (Select KEY_VALUE from TATSHIPP WHERE TABLE_NAME = 'ARTCUST1');" _
            & " BEGIN FOR R1 IN C1 LOOP " _
            & "    DELETE FROM " & TATSHIPP & " WHERE SHIPMENT_AMT > 0 AND CUST_CODE = R1.CUST_CODE;" _
            & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "INSERT INTO " & TATSHIPP _
            & " SELECT DISTINCT CUST_CODE, TATSHIPP.SHIPMENT_AMT, TATSHIPP.SHIPMENT_PERC, CUST_STATE" _
            & " FROM " & TATSHIPP & ", TATSHIPP" _
            & " WHERE TABLE_NAME = 'ARTCUST1' AND " & TATSHIPP & ".CUST_CODE = TATSHIPP.KEY_VALUE"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim tblSOTORDR1 As DataTable = ASCDATA1.GetDataTable("Select * from " & SOTORDR1)

        ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_REL_BATCH_NO, SOTORDR1.ORDR_REL_HOLD_CODES, ICTSTYL1.CASE_CUBE " _
            & " from " & SOTORDR1 & " SOTORDR1, " & SOTORDR2 & " SOTORDR2, ICTSTYL1" _
            & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & " and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
        Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
        Dim tblTATSHIPP As DataTable = ASCDATA1.GetDataTable("select * from " & TATSHIPP)

        tblSOTORDR2.Columns.Add("ORDER_SALES", GetType(System.Decimal), "ORDR_QTY * ORDR_UNIT_PRICE")
        tblSOTORDR2.Columns.Add("OPEN_SALES", GetType(System.Decimal), "ORDR_QTY_OPEN * ORDR_UNIT_PRICE")
        tblSOTORDR2.Columns.Add("CUR_ALLOC_SALES", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO_CUR, 0) * ORDR_UNIT_PRICE")
        tblSOTORDR2.Columns.Add("FUT_ALLOC_SALES", GetType(System.Decimal), "ISNULL(ORDR_QTY_ALLO_FUT, 0) * ORDR_UNIT_PRICE")
        tblSOTORDR2.Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(ORDR_QTY_ALLO_CUR,0) / ISNULL(CARTON_PACK_QTY,0))")
        tblSOTORDR2.Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")
        tblSOTORDR2.Columns.Add("PALLET_PERC", GetType(System.Decimal), "ISNULL(TOTAL_CUBE,0) / " & palletCube)

        ' Grab Orders With Only Inventory or No Order Hold Codes.
        ' Also the customer needs to allow cback orders; otherwise the items are cancelled.

        ' This is the orders to merge together.
        Dim wkOrderGroupNo As String = ASCMAIN1.Temp_Table("Select Distinct ORDR_GROUP_NO " _
                                                                  & " from " & SOTORDR1 & " SOTORDR1, ARTCUST1" _
                                                                  & " where SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE" _
                                                                  & " and SOTORDR1.ORDR_REL_BATCH_NO Is Not NULL" _
                                                                  & " and (" & ordrRelHoldCodesToProcess & ")" _
                                                                  & " and NVL(ARTCUST1.CUST_ALLOW_BACKORDER, '0') = '1'" _
                                                                  & " AND SOTORDR1.ORDR_SHIP_DATE <= '" & SHIP_BY_DATE.ToString("dd-MMM-yyyy") & "'")

        Dim tblORDR_GROUP_NO As DataTable = ASCDATA1.GetDataTable("SELECT ORDR_GROUP_NO, MIN(ORDR_SHIP_DATE) ORDR_SHIP_DATE" _
                                                                  & " FROM " & SOTORDR1 _
                                                                  & " WHERE ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM " & wkOrderGroupNo & ")" _
                                                                  & " GROUP BY ORDR_GROUP_NO")

        Dim mergedCustomers As New List(Of String)

        For Each rowORDR_GROUP_NO As DataRow In tblORDR_GROUP_NO.Select("", "ORDR_SHIP_DATE DESC, ORDR_GROUP_NO")
            Dim ORDR_GROUP_NO As String = rowORDR_GROUP_NO.Item("ORDR_GROUP_NO")

            Dim openSales As Decimal = 0
            Dim curAlloSales As Decimal = 0
            Dim availPercent As Int16 = 0
            Dim customerRercent As Int16 = 0
            Dim CUST_CODE As String = tblSOTORDR1.Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")(0).Item("CUST_CODE")
            Dim CUST_STORE_NO As String = tblSOTORDR1.Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")(0).Item("CUST_STORE_NO")
            Dim totalPallets As Decimal = 0
            Dim totalCube As Decimal = 0

            If tblSOTORDR1.Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' AND ORDR_REL_BATCH_NO is not NULL").Length = 1 Then
                If tblSOTORDR1.Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' AND ORDR_REL_BATCH_NO is not NULL")(0).Item("ORDR_REL_HOLD_CODES") & String.Empty = String.Empty Then
                    ' No work to be done. Single Order without any Hold Codes
                    Continue For
                End If
            End If

            For Each rowSOTORDR1 As DataRow In tblSOTORDR1.Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' AND ORDR_REL_BATCH_NO is not NULL AND (ORDR_STATUS = 'P' OR ORDR_STATUS = 'O')")
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")

                openSales += Val(tblSOTORDR2.Compute("SUM(OPEN_SALES)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)

                ' Grab only No Hold Codes or Inventory Only Hold Codes
                If (",'','I','IE','IEF','IF',").Contains(",'" & rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & String.Empty & "',") Then
                    curAlloSales += Val(tblSOTORDR2.Compute("SUM(CUR_ALLOC_SALES)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)

                    If chkRelFutAvail.Checked Then
                        curAlloSales += Val(tblSOTORDR2.Compute("SUM(FUT_ALLOC_SALES)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
                    End If

                    totalPallets += Val(tblSOTORDR2.Compute("SUM(PALLET_PERC)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)
                    totalCube += Val(tblSOTORDR2.Compute("SUM(TOTAL_CUBE)", "ORDR_NO = '" & ORDR_NO & "'") & String.Empty)

                End If
            Next

            ' Get the Ship % for this Customer / Shipment
            customerRercent = tblTATSHIPP.Select("CUST_CODE = '" & CUST_CODE & "' AND SHIPMENT_AMT <= " & openSales, "SHIPMENT_AMT DESC")(0).Item("SHIPMENT_PERC")

            availPercent = 0
            If openSales > 0 Then
                availPercent = (curAlloSales / openSales) * 100
            End If

            ' If the Cube Rule is On and we pass the Cube Rule then no further validation - release shipment
            If Absx1.chkFor("CHKNUM_CUBE").Checked _
                    AndAlso totalCube >= Absx1.numFor("SO_PARM_REL_CUBE").Value _
                    AndAlso Absx1.numFor("SO_PARM_REL_CUBE").Value > 0 Then

                ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                            & " Set ORDR_REL_HOLD_CODES = NULL" & vbCrLf _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                            & " AND ORDR_REL_BATCH_NO IS NOT NULL" _
                            & " and (" & ordrRelHoldCodesToProcess & ") " _
                            & " AND ORDR_STATUS IN ('P','O')" _
                            & " and ORDR_NO IN (SELECT ORDR_NO IN " & tblReleaseable & ")"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                If Not mergedCustomers.Contains(CUST_CODE) Then
                    ASCMAIN1.sql = "Select ORDR_NO, WHSE_CODE, ORDR_DATE, CUST_CODE, CUST_NAME, ORDR_CUST_PO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_GROUP_NO,ORDR_REL_HOLD_CODES from " & SOTORDR1 & " WHERE CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "' and ORDR_SHIP_DATE <= '" & SHIP_BY_DATE.ToString("dd-MMM-yyyy") & "'"
                    Fill_Records("SOTOREMM", String.Empty, False, ASCMAIN1.sql)
                End If
                Continue For
            End If

            ' Get a total of only I, IE release hold codes. Since we Clear the I we do not want to release
            ' a subset of orders for a group of orders. For example: IE, IE, I, ID - If we clear the 'I' release code then one order would ship
            ' when in fact it should ship with the other orders
            If (openSales > 0 AndAlso curAlloSales > 0) AndAlso (availPercent >= customerRercent) Then
                ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                    & " Set ORDR_REL_HOLD_CODES = NULL" & vbCrLf _
                    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                    & " AND ORDR_REL_BATCH_NO IS NOT NULL" _
                    & " and (" & ordrRelHoldCodesToProcess & ") " _
                    & " AND ORDR_STATUS IN ('P','O')" _
                    & " and ORDR_NO IN (SELECT ORDR_NO from " & tblReleaseable & ")"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                If Not mergedCustomers.Contains(CUST_CODE) Then
                    ASCMAIN1.sql = "Select ORDR_NO, WHSE_CODE, ORDR_DATE, CUST_CODE, CUST_NAME, ORDR_CUST_PO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_GROUP_NO,ORDR_REL_HOLD_CODES from " & SOTORDR1 & " WHERE CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "' and ORDR_SHIP_DATE <= '" & SHIP_BY_DATE.ToString("dd-MMM-yyyy") & "'"
                    Fill_Records("SOTOREMM", String.Empty, False, ASCMAIN1.sql)
                End If
            Else
                ' What about a group for 4 sales orders, 3 have Hold Codes and one does not. We must prevent the one that does not have hold codes from releasing
                ' but we shouldn't hold back any ecomm orders that are ok for release, note this sub is RGI specific
                If Not chkECommerce.Checked Then
                    ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                            & " Set ORDR_REL_HOLD_CODES = 'M'" & vbCrLf _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_REL_HOLD_CODES is NULL"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If
            End If
        Next

        ' Set the Regency Order Group Number back to the Order Number
        If wkGroup.Length > 0 Then
            ASCDATA1.ExecuteSQL("Update " & SOTORDR1 & " SET ORDR_GROUP_NO = (SELECT ORDR_GROUP_NO FROM " & wkGroup & " WHERE ORDR_NO = " & SOTORDR1 & ".ORDR_NO)")
            For Each rowSOTOREMM As DataRow In dst.Tables("SOTOREMM").Select("")
                rowSOTOREMM.Item("ORDR_GROUP_NO") = rowSOTOREMM.Item("ORDR_NO")
            Next
        End If

        Dim tbl As DataTable = ASCDATA1.SelectDistinct("SOTOREMM", New String() {"CUST_CODE"})
        For Each row As DataRow In tbl.Select("", "CUST_CODE")
            Dim CUST_CODE As String = row.Item("CUST_CODE") & String.Empty

            ASCMAIN1.sql = "ORDR_REL_HOLD_CODES <> '' AND CUST_CODE = '" & CUST_CODE & "'"
            If dst.Tables("SOTOREMM").Select(ASCMAIN1.sql).Length = 0 Then
                For Each rowSOTOREMM As DataRow In dst.Tables("SOTOREMM").Select("CUST_CODE = '" & CUST_CODE & "'")
                    rowSOTOREMM.Delete()
                Next
            End If

            dst.Tables("SOTOREMM").AcceptChanges()
        Next
    End Sub

    Private Sub AuthorizeCreditCards(ByVal ORDR_NO As String, ByVal chargeValue As Decimal)

        If ASCMAIN1.CLIENT <> "RGI" Then
            Exit Sub
        End If

        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.DBS_SERVER Then
            Exit Sub
        End If

        If Not dst.Tables.Contains("ARTCCPA1") Then
            Create_TDA(dst.Tables.Add, "ARTCCPA1", "*")
        Else
            dst.Tables("ARTCCPA1").Rows.Clear()
        End If

        If Not dst.Tables.Contains("SOTORDC1") Then
            Create_TDA(dst.Tables.Add, "SOTORDC1", "*")
            Create_TDA(dst.Tables.Add("SOTORDCX"), "SOTORDC1", "*")
        Else
            dst.Tables("SOTORDC1").Rows.Clear()
            dst.Tables("SOTORDCX").Rows.Clear()
        End If

        Try
            Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = :PARM1", "V", ORDR_NO)
            If rowSOTORDR1 Is Nothing Then
                ASCMAIN1.sql = "Update " & SOTORDR1 _
                        & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R', ORDR_REL_BATCH_NO = NULL where ORDR_NO = '" & ORDR_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                Exit Sub
            End If

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
            Dim CCPA_NO As String = rowSOTORDR1.Item("CCPA_NO") & String.Empty

            ' See if we have an existing auth that will cover these charges
            If CCPA_NO.Length > 0 Then
                ASCMAIN1.sql = " SELECT SUM(CCPA_AMT) CCPA_AMT "
                ASCMAIN1.sql &= " FROM"
                ASCMAIN1.sql &= " ( "
                ASCMAIN1.sql &= " select CCPA_NO, CCPA_AMT from ARTCCPA1 WHERE ORDR_NO = '" & ORDR_NO & "' AND CCPA_STATUS = 'T'"
                ASCMAIN1.sql &= " UNION"
                ASCMAIN1.sql &= " select CCPA_NO, CCPA_AMT * -1 from ARTCCPA1 WHERE ORDR_NO = '" & ORDR_NO & "' AND CCPA_STATUS = 'S'"
                ASCMAIN1.sql &= " )"

                Dim available As Decimal = Val(ASCDATA1.GetDataValue(ASCMAIN1.sql) & String.Empty)

                If available >= chargeValue Then
                    If ASCDATA1.GetDataTable("SELECT * FROM SOTPICK1 WHERE CCPA_NO_AUTH = '" & CCPA_NO & "' and PICK_STATUS = 'P'").Rows.Count = 0 Then
                        Dim rowARTCCPA1_x As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCCPA1 WHERE CCPA_NO = '" & CCPA_NO & "'")
                        If rowARTCCPA1_x IsNot Nothing Then
                            Dim cust_credit_card_exp_date As String = rowARTCCPA1_x.Item("cust_credit_card_exp_date") & String.Empty
                            If cust_credit_card_exp_date.Length = 4 Then
                                cust_credit_card_exp_date = cust_credit_card_exp_date.Substring(2, 2) & cust_credit_card_exp_date.Substring(0, 2)
                                Dim today As String = Now.ToString("yy") & Now.ToString("MM")
                                If Val(cust_credit_card_exp_date) >= Val(today) Then
                                    ASCDATA1.ExecuteSQL("UPDATE " & SOTORDR1 & " SET CCPA_NO = '" & rowARTCCPA1_x.Item("CCPA_NO") & "', CC_TRANS_ID = '" & rowARTCCPA1_x.Item("TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            ASCMAIN1.Progress("Authorize Credit Card for Customer " & rowSOTORDR1.Item("CUST_NAME"), ORDR_NO)

            Dim tblARTCCPA1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ARTCCPA1 WHERE ORDR_NO = :PARM1", "ARTCCPA1", "V", New Object() {ORDR_NO})
            Dim tblARTCUSTC As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ARTCUSTC WHERE CUST_CODE = :PARM1 AND NVL(CUST_CREDIT_CARD_STATUS ,'A') = 'A'", "ARTCCPA1", "V", New Object() {CUST_CODE})

            If clsTACENCRY.UseEncryption Then
                For Each rowARTCCPA1x As DataRow In tblARTCCPA1.Select("")
                    For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                        rowARTCCPA1x.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1x.Item(field & "_E") & String.Empty)
                        rowARTCCPA1x.Item(field & "_E") = DBNull.Value
                    Next
                Next

                For Each rowARTCUSTCx As DataRow In tblARTCUSTC.Select("")
                    For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                        rowARTCUSTCx.Item(field) = clsTACENCRY.DecryptString(rowARTCUSTCx.Item(field & "_E") & String.Empty)
                        rowARTCUSTCx.Item(field & "_E") = DBNull.Value
                    Next
                Next
            End If

            Dim rowARTCCPA1 As DataRow = Nothing
            Dim rowARTCUSTC As DataRow = Nothing

            Dim cardNo As String = String.Empty
            Dim expDate As String = String.Empty
            Dim verCode As String = String.Empty

            ' If show/web $1.00 auth then use that card
            If tblARTCCPA1.Select("CCPA_TYPE = 'A' AND CCPA_AMT = 1").Length > 0 Then
                rowARTCCPA1 = tblARTCCPA1.Select("CCPA_TYPE = 'A' AND CCPA_AMT = 1")(0)
                cardNo = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
                expDate = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                verCode = rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty

                ' evaluate Expiration date
                If expDate.Length = 4 Then
                    Dim expiredate = expDate.Substring(0, 2) & "/28/" & expDate.Substring(2, 2)
                    If IsDate(expiredate) Then
                        expiredate = CDate(expiredate).ToString("yyyyMMdd")
                        If Val(expiredate) < Val(DateTime.Now.ToString("yyyyMMdd")) Then
                            cardNo = String.Empty
                            expDate = String.Empty
                            verCode = String.Empty
                        End If
                    End If
                End If

            End If

            ' If Credit Card on the order then Use that Credit Card.
            If cardNo.Length = 0 AndAlso rowSOTORDR1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                If tblARTCCPA1.Select("CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "' AND RESPONSE_CODE = 'A'").Length > 0 Then
                    rowARTCCPA1 = tblARTCCPA1.Select("CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "' AND RESPONSE_CODE = 'A'")(0)
                    cardNo = rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") & String.Empty
                    expDate = rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                    verCode = rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty
                End If
            End If

            'CUST_CREDIT_CARD_STATUS, CUST_CREDIT_CARD_PREFERRED
            If cardNo.Length = 0 AndAlso tblARTCUSTC.Rows.Count > 0 Then
                If tblARTCUSTC.Select("CUST_CREDIT_CARD_PREFERRED = '1' AND CUST_CREDIT_CARD_STATUS <> 'I'").Length > 0 Then
                    rowARTCUSTC = tblARTCUSTC.Select("CUST_CREDIT_CARD_PREFERRED = '1' AND CUST_CREDIT_CARD_STATUS <> 'I'")(0)
                    cardNo = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO") & String.Empty
                    expDate = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                    verCode = rowARTCUSTC.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty
                    rowARTCCPA1 = tblARTCCPA1.NewRow
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_NO") = rowARTCUSTC.Item("CUST_CREDIT_CARD_NO")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_EXP_DATE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_EXP_DATE")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_VER_CODE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_VER_CODE")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_NAME") = rowARTCUSTC.Item("CUST_CREDIT_CARD_NAME")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_ADDR1") = rowARTCUSTC.Item("CUST_CREDIT_CARD_ADDR1")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_CITY") = rowARTCUSTC.Item("CUST_CREDIT_CARD_CITY")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_STATE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_STATE")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_ZIP_CODE") = rowARTCUSTC.Item("CUST_CREDIT_CARD_ZIP_CODE")
                    rowARTCCPA1.Item("CUST_CREDIT_CARD_COUNTRY") = rowARTCUSTC.Item("CUST_CREDIT_CARD_COUNTRY")
                End If
            End If

            If cardNo.Length = 0 Then
                ASCMAIN1.sql = "Update " & SOTORDR1 _
                    & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R', ORDR_REL_BATCH_NO = NULL where ORDR_NO = '" & ORDR_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                Exit Sub
            End If

            ' Do not permit any additions if the order is locked by someone else
            Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty
            If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , False, , 4) Then
                ASCMAIN1.sql = "Update " & SOTORDR1 _
                    & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R', ORDR_REL_BATCH_NO = NULL where ORDR_NO = '" & ORDR_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                Exit Sub
            End If

            ' Try to get an Authorization
            If Not ProcessCCAuthorization(ORDR_NO, chargeValue, rowARTCCPA1) Then
                ASCMAIN1.sql = "Update " & SOTORDR1 _
                    & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R', ORDR_REL_BATCH_NO = NULL where ORDR_NO = '" & ORDR_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                Exit Sub
            End If

        Catch ex As Exception
            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'R', ORDR_REL_BATCH_NO = NULL" _
                & " where ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Finally
            ASCMAIN1.MultiTask_Release(, , 4)
        End Try
    End Sub

    Private Function ProcessCCAuthorization(ByVal ORDR_NO As String, ByVal shipmentAmount As Decimal, ByVal rowCCAttributes As DataRow) As Boolean

        ProcessCCAuthorization = True

        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM " & SOTORDR1 & " WHERE ORDR_NO = :PARM1", "V", ORDR_NO)
        If rowSOTORDR1 Is Nothing Then
            Return False
        End If

        Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
        Dim FRT_TERMS As String = rowSOTORDR1.Item("FRT_TERMS") & String.Empty
        Dim SHIP_VIA_CODE As String = rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty
        Dim CCPA_NO As String = String.Empty

        Dim freightCost As Decimal = 0
        Dim chargeAgainstAuth As Boolean = False
        Dim rowSOTORDC1 As DataRow = Nothing

        If ORDR_GROUP_NO_sql.Length > 0 Then
            'Return False
        End If

        ' Do Only for RGI !!!!
        If ASCMAIN1.CLIENT <> "RGI" Then
            Return False
        End If

        If ASCMAIN1.Running_in_VS Then
            Stop
        End If

        Dim EMsg As String = String.Empty
        If FRT_TERMS.Length > 0 Then
            If ASCDATA1.GetDataRow("select * from astcode1 where TABLE_NAME = 'SOTORDR1' AND COLUMN_NAME = 'FRT_TERMS' AND T_CODE = '" & FRT_TERMS & "'") Is Nothing Then
                EMsg &= vbCr & "Freight Terms are required to process a credit card."
            End If
        Else
            EMsg &= vbCr & "Freight Terms are required to process a credit card."
        End If

        If SHIP_VIA_CODE.Length > 0 Then
            If ASCDATA1.GetDataRow("SELECT * FROM SOTSVIA1 WHERE SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'") Is Nothing Then
                EMsg &= vbCr & "Ship Via Code is required for credit card processing."
            End If
        Else
            EMsg &= vbCr & "Ship Via Code is required for credit card processing."
        End If

        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("select sotcarr1.carrier_type" _
                                                         & " from sotsvia1, sotcarr1" _
                                                         & " where sotsvia1.carrier_code = sotcarr1.carrier_code" _
                                                         & " and ship_via_code = :PARM1", "V", New Object() {SHIP_VIA_CODE})


        If rowSOTCARR1 Is Nothing Then
            EMsg &= vbCr & "Could not determine carrier for the Ship Via Code."
        End If

        If EMsg.Length > 0 Then
            Return False
        End If

        ' Fedex, UPS and similar pay for freight when freight terms of PPA 
        If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty = "U" OrElse FRT_TERMS.ToUpper = "PPA" Then
            ' New Rule 1/24/2013. 20% or $20 the greater of the two
            freightCost = shipmentAmount * 0.2
            If freightCost < 20 Then
                freightCost = 20
            End If
        End If

        shipmentAmount += freightCost

        If shipmentAmount = 0 Then
            Return True
        End If

        If Not ASCMAIN1.Logical_Lock("ARTCUSTC", CUST_CODE, , False, True, 4) Then
            Return False
        End If

        If Not ASCMAIN1.Logical_Open("ARTCCPA1", "*", , False, True, 4) Then
            Return False
        End If

        Using frmCCProcessor As New TAC.TAFCARDF(Me)
            frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
            frmCCProcessor.CUST_CODE = CUST_CODE
            frmCCProcessor.CCPA_REASON = "O"
            frmCCProcessor.ORDR_NO = ORDR_NO
            frmCCProcessor.TRAN_TYPE = "A"

            With frmCCProcessor.rowARTCCPA1
                .Item("CUST_CODE") = CUST_CODE
                .Item("CCPA_AMT") = shipmentAmount
                .Item("CCPA_NOTE") = "Credit Card Order"

                .Item("CUST_CREDIT_CARD_NO") = rowCCAttributes.Item("CUST_CREDIT_CARD_NO")
                .Item("CUST_CREDIT_CARD_EXP_DATE") = rowCCAttributes.Item("CUST_CREDIT_CARD_EXP_DATE")
                .Item("CUST_CREDIT_CARD_VER_CODE") = rowCCAttributes.Item("CUST_CREDIT_CARD_VER_CODE")
                .Item("CUST_CREDIT_CARD_NAME") = rowCCAttributes.Item("CUST_CREDIT_CARD_NAME")
                .Item("CUST_CREDIT_CARD_ADDR1") = rowCCAttributes.Item("CUST_CREDIT_CARD_ADDR1")
                .Item("CUST_CREDIT_CARD_CITY") = rowCCAttributes.Item("CUST_CREDIT_CARD_CITY")
                .Item("CUST_CREDIT_CARD_STATE") = rowCCAttributes.Item("CUST_CREDIT_CARD_STATE")
                .Item("CUST_CREDIT_CARD_ZIP_CODE") = rowCCAttributes.Item("CUST_CREDIT_CARD_ZIP_CODE")
                .Item("CUST_CREDIT_CARD_COUNTRY") = rowCCAttributes.Item("CUST_CREDIT_CARD_COUNTRY")

            End With

            Try
                frmCCProcessor.CC_Authorize(True)
                Fill_Records("SOTORDC1", String.Empty, True, "Select * from SOTORDC1 where ORDR_NO = '" & ORDR_NO & "'")
                Fill_Records("ARTCCPA1", String.Empty, False, "Select * from ARTCCPA1 where CCPA_NO = '" & frmCCProcessor.CCPA_NO & String.Empty & "'")
                If clsTACENCRY.UseEncryption Then
                    For Each rowARTCCPA1x As DataRow In dst.Tables("ARTCCPA1").Select("")
                        For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_VER_CODE"} ' "CUST_CREDIT_CARD_EXP_DATE",
                            rowARTCCPA1x.Item(field) = clsTACENCRY.DecryptString(rowARTCCPA1x.Item(field & "_E") & String.Empty)
                            rowARTCCPA1x.Item(field & "_E") = DBNull.Value
                        Next
                    Next
                End If

                Dim row As DataRow = dst.Tables("ARTCCPA1").Rows.Find(frmCCProcessor.CCPA_NO & String.Empty)
                If row IsNot Nothing AndAlso (row.Item("CCPA_STATUS") & String.Empty = "T" OrElse row.Item("CCPA_STATUS") & String.Empty = "S") Then
                    rowSOTORDR1.Item("CCPA_NO") = frmCCProcessor.CCPA_NO & String.Empty
                    rowSOTORDR1.Item("CC_TRANS_ID") = row.Item("TRANS_ID")

                    If row.Item("CCPA_STATUS") & String.Empty = "T" Then
                        ASCDATA1.ExecuteSQL("UPDATE SOTORDR1 SET CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "', CC_TRANS_ID = '" & rowSOTORDR1.Item("CC_TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                        ASCDATA1.ExecuteSQL("UPDATE " & SOTORDR1 & " SET CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "', CC_TRANS_ID = '" & rowSOTORDR1.Item("CC_TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                    Else
                        ProcessCCAuthorization = False
                    End If

                    rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                    rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                    rowSOTORDC1.Item("TRANS_TYPE") = "C"
                    rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                    rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                    rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                    rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("BALANCE") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("ACTIVE_IND") = "1"
                    dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)

                ElseIf row IsNot Nothing Then
                    rowSOTORDC1 = dst.Tables("SOTORDC1").NewRow
                    rowSOTORDC1.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDC1.Item("TRANS_NO") = Val(dst.Tables("SOTORDC1").Compute("MAX(TRANS_NO)", "") & String.Empty) + 1
                    rowSOTORDC1.Item("TRANS_TYPE") = "C"
                    rowSOTORDC1.Item("TRANS_DATE") = DateTime.Now
                    rowSOTORDC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDC1.Item("CCPA_NO") = row.Item("CCPA_NO")
                    rowSOTORDC1.Item("CCPA_STATUS") = row.Item("CCPA_STATUS")
                    rowSOTORDC1.Item("AMOUNT") = row.Item("CCPA_AMT")
                    rowSOTORDC1.Item("BALANCE") = 0
                    rowSOTORDC1.Item("ACTIVE_IND") = "0"
                    dst.Tables("SOTORDC1").Rows.Add(rowSOTORDC1)
                    dst.Tables("SOTORDCX").ImportRow(rowSOTORDC1)
                    ProcessCCAuthorization = False
                End If

                Update_Record_TDA("SOTORDC1")

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ProcessCCAuthorization = False
            End Try

        End Using

        ASCMAIN1.MultiTask_Release(, , 4)
        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Function


    Sub Create_PICK_SHIP_CART()

        ' Create Pick Tickets, Shipment BOL, and Cartons
        '   do this for all Sales Orders scheduled to ship on or before SHIP_BY_DATE
        '   also, filter on Division, Customer, and Order Group

        Dim PICK_RELEASED As Date = DATETIME_STAMP
        'PICK_NO_seq = 0

        ASCMAIN1.sql = "Select * from " & ARTCUST1
        Fill_Records("ARTCUST1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR1.*, DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO) SHIP_TO from " & SOTORDR1 & " SOTORDR1"
        Fill_Records("SOTORDR1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from " & SOTORDR2 & " SOTORDR2"
        Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

        Dim tblARTCCPA1 As DataTable = ASCDATA1.GetDataTable("Select ORDR_NO, CCPA_NO" & vbCrLf _
                & " from TATTERM1," & SOTORDR1 & " SOTORDR1" _
                & " where TATTERM1.TERM_CODE = SOTORDR1.TERM_CODE" & vbCrLf _
                & " and TATTERM1.TERM_TYPE = 'D'")

        Dim rowARTCUST1 As DataRow
        Dim CUST_CODE As String = ""
        For Each rowSOTORDR1_rel As DataRow In dst.Tables("SOTORDR1").Select _
                ("ORDR_REL_BATCH_NO is Not Null and ORDR_REL_HOLD_CODES is Null", _
                 "CUST_CODE, ORDR_GROUP_NO, ORDR_ADDR_TYPE_ST, SHIP_TO, ORDR_NO")
            If CUST_CODE <> rowSOTORDR1_rel.Item("CUST_CODE") Then
                CUST_CODE = rowSOTORDR1_rel.Item("CUST_CODE")
                rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            End If

            Dim ORDR_NO As String = rowSOTORDR1_rel.Item("ORDR_NO")
            Dim PICK_SEQ_NO As Integer = Val(rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") & "") + 1
            rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") = PICK_SEQ_NO

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            PICK_NO_seq += 1
            Dim PICK_NO As String = "TEMP" & Format(PICK_NO_seq, "000000")
            With rowSOTPICK1
                .Item("PICK_NO") = PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_PICK_SEQ") = PICK_SEQ_NO
                .Item("PICK_STATUS") = "P"
                .Item("PICK_RELEASED") = PICK_RELEASED
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("SHIP_BOL_NO") = "X"

                ' 04/05/2015 RGI to process CC after Release.
                If (ASCMAIN1.DBS_COMPANY = "RGI" OrElse ASCMAIN1.DBS_SERVER = "RGI") Then
                    If tblARTCCPA1.Select("ORDR_NO = '" & ORDR_NO & "'").Length > 0 Then
                        .Item("CCPA_NO_STATUS") = "1"
                        .Item("CCPA_NO_AUTH") = tblARTCCPA1.Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("CCPA_NO") & String.Empty
                    Else
                        .Item("CCPA_NO_STATUS") = "0"
                    End If
                End If
            End With

            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

            Dim TOTAL_OPEN As Int64 = 0 ' Total Units left OPEN in Order after Release
            Dim TOTAL_PICK As Int64 = 0 ' Total Units in PICK in Order after Release

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "' and ORDR_QTY_OPEN <> 0"
            For Each rowSOTORDR2_rel As DataRow In dst.Tables("SOTORDR2").Select(sqlw, "ORDR_LNO")
                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    .Item("PICK_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")

                    .Item("STYLE_CODE") = rowSOTORDR2_rel.Item("STYLE_CODE")
                    .Item("COLOR_CODE") = rowSOTORDR2_rel.Item("COLOR_CODE")
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE")

                    If ASCMAIN1.CLIENT = "VAN" Then
                        .Item("SUB_BODY_CODE") = rowSOTORDR2_rel.Item("SUB_BODY_CODE")
                        .Item("STANDARD_CUBE_PER_UNIT") = rowSOTORDR2_rel.Item("STANDARD_CUBE_PER_UNIT")
                    End If

                    If ASCMAIN1.Running_in_VS And (.Item("STYLE_CODE") = "MTX59907" Or .Item("STYLE_CODE") = "720R") Then Stop

                    Dim qCANC As Int64 = 0
                    Dim qBACK As Int64 = 0

                    Dim qA As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") & "")
                    Dim qO As Int64 = Val(rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") & "")
                    '  qA = qO ' for testing only

                    If (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") _
                    Or (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") Then
                        If chkForcePick.Checked Then
                            qA = qO
                        End If
                    End If
                    ' If qA <> 0 Then Stop
                    rowSOTORDR2_rel.Item("ORDR_QTY_PICK") = Val(rowSOTORDR2_rel.Item("ORDR_QTY_PICK") & "") + qA
                    .Item("PICK_QTY") = qA
                    TOTAL_PICK = TOTAL_PICK + qA

                    If qO - qA <> 0 Then
                        If rowSOTORDR2_rel.Item("ORDR_BACKORDER") & "" = "Y" Then
                            rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") = qO - qA
                            TOTAL_OPEN = TOTAL_OPEN + qO - qA
                            qBACK = qO - qA
                        Else
                            rowSOTORDR2_rel.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2_rel.Item("ORDR_QTY_CANC") & "") + qO - qA
                            rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") = 0
                            qCANC = qO - qA
                        End If
                    Else
                        rowSOTORDR2_rel.Item("ORDR_QTY_OPEN") = 0
                    End If

                    rowSOTORDR2_rel.Item("ORDR_QTY_ALLO_CUR") = 0

                    .Item("PICK_QTY_CANC_REL") = qCANC
                    .Item("PICK_QTY_BACK_REL") = qBACK

                    'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    '    ' NO CHANGE IN BEHAVIOR
                    '    dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
                    'Else
                    '    If Val(.Item("PICK_QTY") & "") = 0 _
                    '    And Val(.Item("PICK_QTY_CANC_REL") & "") = 0 _
                    '    Then
                    '        ' DO NOT WRITE PICK TICKET DETAIL
                    '    Else
                    '        dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
                    '    End If
                    'End If

                End With
                ' NEXT LINE WRITES ALL PT LINES REGARDLESS IF ANYTHING RELEASED OR GOT CANCELLED
                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
            Next

            If TOTAL_OPEN = 0 Then
                If TOTAL_PICK = 0 Then
                    If Val(rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") & "") <= 1 Then
                        rowSOTORDR1_rel.Item("ORDR_STATUS") = "C"
                    Else
                        rowSOTORDR1_rel.Item("ORDR_STATUS") = "F"
                    End If
                    ' note: the next 2 fields are not going back to oracle; 
                    ' this code was placed here to make cancel in order release equivalent to cancel in order entry; 
                    ' not really a problem since we are not using these 2 fields for anything yet
                    rowSOTORDR1_rel.Item("ORDR_DATE_CLOSED") = DATETIME_STAMP.Date
                    rowSOTORDR1_rel.Item("ORDR_YYYYPP_CLOSED") = ASCMAIN1.CYP
                Else
                    rowSOTORDR1_rel.Item("ORDR_STATUS") = "P"
                End If
            End If
        Next

        Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
        dst.Tables("SOTPICK2").Columns.Add("SHIP_BOL_NO", GetType(System.String), "PARENT(SOTPICK1_SOTPICK2).SHIP_BOL_NO")
        dst.Tables("SOTPICK2").Columns.Add("PICK_AMT", GetType(System.Decimal), "ISNULL(PICK_QTY,0)*ISNULL(PICK_UNIT_PRICE,0)")
        dst.Tables("SOTPICK1").Columns.Add("PICK_AMT", GetType(System.Decimal), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT)")
        dst.Tables("SOTPICK1").Columns.Add("PICK_QTY", GetType(System.Decimal), "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")

        ' Create_Relation("SOTSHIP2", "SOTPICK2", "SHIP_BOL_NO,STYLE_CODE,COLOR_CODE")
        dst.Tables("SOTSHIP2").Columns.Add("PICK_QTY", GetType(System.Int64)) ' , "SUM(CHILD(SOTSHIP2_SOTPICK2).PICK_QTY)")
        Create_Relation("SOTSHIP3", "SOTSHIP2", "SHIP_BOL_NO,STYLE_CODE")
        dst.Tables("SOTSHIP3").Columns.Add("PICK_QTY", GetType(System.Int64), "SUM(CHILD(SOTSHIP3_SOTSHIP2).PICK_QTY)")

        ' when done, regroup pick tickets by group/dc to assign SHIP_BOL_NO

        Dim CART_LNO_seq As Int64 = 0
        SHIP_BOL_NO_seq = 0

        ' TROUBLE PRINTING PICK TICKETS IF WE ALLOW BREAKS ON SHIP VIA NOW
        ' & " SOWORDR1.SHIP_VIA_CODE"
        '            & ", NULL SHIP_VIA_CODE " & vbCrLf _


        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO) SHIP_TO" & vbCrLf _
            & ", MIN (NVL(SOTORDR1.SHIP_VIA_CODE,ARTCUST1.SHIP_VIA_CODE)) SHIP_VIA_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.TERM_CODE) TERM_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.SREP_CODE) SREP_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.FRT_TERMS) FRT_TERMS" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_DEPT) ORDR_DEPT" & vbCrLf _
            & ", MAX (SOTORDR1.CUST_FACTOR_IND) CUST_FACTOR_IND" & vbCrLf _
            & ", MAX (SOTORDR1.EDI_DOC_SEQ_NO) EDI_DOC_SEQ_NO" & vbCrLf _
            & " from " & SOTORDR1 & " SOTORDR1, ARTCUST1" & vbCrLf _
            & " where SOTORDR1.ORDR_REL_BATCH_NO is Not Null" & vbCrLf _
            & "   and SOTORDR1.ORDR_REL_HOLD_CODES is Null" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & " group by SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO)"

        Dim QTY_TO_PACK As Int64         ' Working Variable to Pack PICK_QTY in to Cartons
        Dim PACK_QTY As Int64            ' Qty to Pack into current carton
        Dim MAX_QTY_CTN_cust As Int64    ' Default for the Maximum Unit Count in a Carton, by Customer, defaulting to SO Parameter
        Dim MAX_QTY_CTN_pick As Int64    ' Default for the Maximum Unit Count in a Carton, for the current Pick Ticket, defaulting to MAX_QTY_CTN_cust, but set to lowest MAX_QTY_CTN of all mixable styles on the Pick Ticket with a MAX_QTY_CTN
        Dim ALLOW_SPLIT_TYPE As String = "" 'Used to decide if we can split a styles between cartons.
        'A = Allow Splits.
        'S = Don't Allow Breaks on Styles.
        'C = Don't Allow Breaks on Style Color Combinations.
        Dim LAST_STYLE As String = ""        'Used to Track Last Style for Splitting options.
        Dim LAST_COLOR As String = ""        'Used to Track Last Color for Splitting options.
        Dim NEXT_SPLIT_AMT As Int64      'Total of Next Split Group
        Dim WHSE_CODE As String = ""
        Dim LP_STATUS As String = ""
        Dim ORDR_PICK_TYPE As String = ""
        Dim SHIP_CART_REQD As String = ""

        Dim rowEDTSLSP1 As DataRow = Nothing
        Dim CUST_856_IND As String = ""
        Dim CUST_810_IND As String = ""

        CUST_CODE = ""
        For Each rowSHIPMENT As DataRow In ASCDATA1.GetDataTable.Select("", "WHSE_CODE,CUST_CODE,ORDR_GROUP_NO,ORDR_ADDR_TYPE_ST,SHIP_TO")
            Dim ORDR_ADDR_TYPE_ST As String = rowSHIPMENT.Item("ORDR_ADDR_TYPE_ST")
            Dim ORDR_GROUP_NO As String = rowSHIPMENT.Item("ORDR_GROUP_NO")
            Dim EDI_DOC_SEQ_NO As String = rowSHIPMENT.Item("EDI_DOC_SEQ_NO") & ""

            If ASCMAIN1.CLIENT = "VAN" Then
                Fill_Records("SOTPCKP2", ORDR_GROUP_NO)
                If dst.Tables("SOTPCKP2").Rows.Count <> 0 Then
                    Dim PACK_NO As String = dst.Tables("SOTPCKP2").Compute("MAX(PACK_NO)", "")
                    dst.Tables("SOTPCKC4_ORDR_NO").Rows.Clear()
                    dst.Tables("SOTPCKC3").Rows.Clear()
                    If PACK_NO <> "" Then
                        Fill_Records("SOTPCKC4_ORDR_NO", PACK_NO)
                        Fill_Records("SOTPCKC3", PACK_NO)
                    End If
                End If
            End If

            Dim EDI_DEPT_DESC As String = ""
            Dim EDI_PROMOTION As String = ""
            Dim rowEDT850T1 As DataRow = LookUp("EDT850T1", EDI_DOC_SEQ_NO)
            If rowEDT850T1 IsNot Nothing Then
                EDI_DEPT_DESC = rowEDT850T1.Item("EDI_DEPT_DESC") & ""
                EDI_PROMOTION = rowEDT850T1.Item("EDI_PROMOTION") & ""
            End If

            Dim CUST_FACTOR_TRANS_IND As String = rowSHIPMENT.Item("CUST_FACTOR_IND") & ""

            Dim SHIP_TO As String = rowSHIPMENT.Item("SHIP_TO") & ""
            If SHIP_TO = "" Then SHIP_TO = "MK"
            Dim SHIP_VIA_CODE As String = rowSHIPMENT.Item("SHIP_VIA_CODE") & ""
            Dim TERM_CODE As String = rowSHIPMENT.Item("TERM_CODE") & ""
            Dim FRT_TERMS As String = rowSHIPMENT.Item("FRT_TERMS") & ""
            Dim SREP_CODE As String = rowSHIPMENT.Item("SREP_CODE") & ""
            Dim ORDR_DEPT As String = rowSHIPMENT.Item("ORDR_DEPT") & ""
            If WHSE_CODE <> rowSHIPMENT.Item("WHSE_CODE") & "" Then
                WHSE_CODE = rowSHIPMENT.Item("WHSE_CODE") & ""
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                    LP_STATUS = "0"
                Else
                    LP_STATUS = ""
                End If
            End If

            'ALLOW_SPLIT_TYPE = "A"
            If CUST_CODE <> rowSHIPMENT.Item("CUST_CODE") Then

                CUST_CODE = rowSHIPMENT.Item("CUST_CODE")
                rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

                If rowARTCUST1.Item("CUST_SHIP_BY_CASE") & "" = "1" Then
                    ORDR_PICK_TYPE = "C"
                Else
                    ORDR_PICK_TYPE = "P"
                End If

                SHIP_CART_REQD = rowARTCUST1.Item("CUST_CART_REQD") & ""
                ALLOW_SPLIT_TYPE = "A"

                Fill_Records("SOTCSTP1", CUST_CODE)
                Dim rowSOTCSTP2 As DataRow = LookUp("SOTCSTP2", CUST_CODE)
                If rowSOTCSTP2 IsNot Nothing Then
                    MAX_QTY_CTN_cust = Val(rowSOTCSTP2.Item("MAX_QTY_CTN") & "")
                    If rowSOTCSTP2.Item("NON_SPLIT_STYLE") & "" = "1" Then ALLOW_SPLIT_TYPE = "S"
                    If rowSOTCSTP2.Item("NON_SPLIT_STYLE_COLOR") & "" = "1" Then ALLOW_SPLIT_TYPE = "C"
                Else
                    MAX_QTY_CTN_cust = Val(ROWs("SOTPARM1").Item("SO_PARM_MAX_CARTON") & "")
                End If

                rowEDTSLSP1 = LookUp("EDTSLSP1", CUST_CODE)
                CUST_856_IND = ""
                CUST_810_IND = ""

                If rowEDTSLSP1 IsNot Nothing Then
                    If ASCMAIN1.CLIENT = "NYA" And (WHSE_CODE = "111" Or WHSE_CODE = "114") Then
                    Else
                        If rowEDTSLSP1.Item("EDI_ID_856") & "" <> "" Then CUST_856_IND = "1"
                        If rowEDTSLSP1.Item("EDI_ID_810") & "" <> "" Then CUST_810_IND = "1"
                    End If

                End If

            End If

            Dim SHIP_856_IND As String = ""
            Dim SHIP_810_IND As String = ""

            If EDI_DOC_SEQ_NO <> "" Then
                SHIP_810_IND = CUST_810_IND
                SHIP_856_IND = CUST_856_IND
            End If

            SHIP_BOL_NO_seq += 1
            Dim SHIP_BOL_NO As String = "TEMP" & Format$(SHIP_BOL_NO_seq, "000000")

            Dim sqlw As String = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            If ORDR_ADDR_TYPE_ST = "DC" Then
                sqlw &= " and ORDR_ADDR_TYPE_ST = '" & ORDR_ADDR_TYPE_ST & "'"
                sqlw &= " and CUST_DC_NO = '" & SHIP_TO & "'"
            Else
                sqlw &= " and (ORDR_ADDR_TYPE_ST is Null or ORDR_ADDR_TYPE_ST = '" & ORDR_ADDR_TYPE_ST & "')"
                sqlw &= " and CUST_STORE_NO = '" & SHIP_TO & "'"
            End If

            ' TROUBLE PRINTING PICK TICKETS IF WE ALLOW BREAKS ON SHIP VIA NOW
            'If SHIP_VIA_CODE = "" Then
            '    & "   and SHIP_VIA_CODE is Null"
            'Else
            '    & "   and SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'"
            'End If

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(sqlw)
                Dim rowSOTPICK1 As DataRow = rowSOTORDR1.GetChildRows("SOTORDR1_SOTPICK1")(0)
                rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            Next

            Dim sqlBOL As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

            For Each rowSC As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTPICK2").Select(sqlBOL), _
                 New String() {"STYLE_CODE", "COLOR_CODE"}).Rows

                Dim STYLE_CODE As String = rowSC.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSC.Item("COLOR_CODE")
                Dim sqlSC As String = " and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

                Dim PICK_QTY As Int64 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", sqlBOL & sqlSC) & "")

                Dim rowSOTSHIP3 As DataRow = dst.Tables("SOTSHIP3").Rows.Find(New String() {SHIP_BOL_NO, STYLE_CODE})
                If rowSOTSHIP3 Is Nothing Then
                    rowSOTSHIP3 = dst.Tables("SOTSHIP3").NewRow
                    rowSOTSHIP3.Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    rowSOTSHIP3.Item("STYLE_CODE") = STYLE_CODE
                    dst.Tables("SOTSHIP3").Rows.Add(rowSOTSHIP3)
                End If

                Dim rowSOTSHIP2 As DataRow = dst.Tables("SOTSHIP2").NewRow
                rowSOTSHIP2.Item("SHIP_BOL_NO") = SHIP_BOL_NO
                rowSOTSHIP2.Item("STYLE_CODE") = STYLE_CODE
                rowSOTSHIP2.Item("COLOR_CODE") = COLOR_CODE
                rowSOTSHIP2.Item("PICK_QTY") = PICK_QTY
                dst.Tables("SOTSHIP2").Rows.Add(rowSOTSHIP2)
            Next

            ' DCG SAYS TO CALCULATE THE SHIP_VIA HERE - BUT FROM WHAT, A ROUTING INSTRUCTION IN TEXT?

            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                .Item("SHIP_ADDR_TYPE") = ORDR_ADDR_TYPE_ST
                .Item("SHIP_ADDR_CODE") = SHIP_TO
                .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                .Item("SHIP_STATUS") = "P"
                .Item("TERM_CODE") = TERM_CODE
                .Item("SREP_CODE") = SREP_CODE
                .Item("FRT_TERMS") = FRT_TERMS
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LP_STATUS") = LP_STATUS
                .Item("ORDR_PICK_TYPE") = ORDR_PICK_TYPE
                .Item("SHIP_CART_REQD") = SHIP_CART_REQD
                .Item("ORDR_DEPT") = ORDR_DEPT

                If ASCMAIN1.CLIENT = "VAN" And CUST_CODE = "WALMART" Then
                    .Item("REASON_CODE") = "H002"
                End If

                If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                    ASCMAIN1.sql = "Select ORDR_MESSAGE from SOTORDR1 where ORDR_NO = (Select Min (ORDR_NO) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
                    Dim ORDR_MESSAGE As String = ASCDATA1.GetDataValue
                    .Item("SHIP_NOTES") = ORDR_MESSAGE
                End If

                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                Else
                    ' .Item("ORDR_DEPT") = ORDR_DEPT
                    .Item("CUST_FACTOR_TRANS_IND") = CUST_FACTOR_TRANS_IND
                    .Item("SHIP_856_IND") = SHIP_856_IND
                    .Item("SHIP_810_IND") = SHIP_810_IND
                End If
            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)


            MAX_QTY_CTN_pick = -1
            For Each rowSOTSHIP2 As DataRow In dst.Tables("SOTSHIP2").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                Dim STYLE_CODE As String = rowSOTSHIP2.Item("STYLE_CODE")
                Dim COLOR_CODE As String = rowSOTSHIP2.Item("COLOR_CODE")

                'Perhaps we can add an option for cartonization by range with this, an if and an option.
                'If you do, make sure it is only available for force picks.
                ' STYLE_CODE = rowSOTSHIP2.Item("RANGE_STYLE_CODE")"

                Dim rowSOTCSTP1 As DataRow = dst.Tables("SOTCSTP1").Rows.Find _
                                             (New Object() {CUST_CODE, STYLE_CODE, COLOR_CODE})
                If rowSOTCSTP1 IsNot Nothing Then
                    Dim MAX_QTY_CTN As Int64 = Val(rowSOTCSTP1.Item("MAX_QTY_CTN") & "")    ' Maximum Unit Count in a Carton for a Customer for a Style/Color
                    Dim NON_MIX As String = rowSOTCSTP1.Item("NON_MIX") & ""

                    If MAX_QTY_CTN > 0 And NON_MIX <> "1" Then
                        If (MAX_QTY_CTN_pick = -1 Or MAX_QTY_CTN < MAX_QTY_CTN_pick) Then
                            MAX_QTY_CTN_pick = MAX_QTY_CTN
                        End If
                    End If
                End If
            Next
            If MAX_QTY_CTN_pick = -1 Then
                MAX_QTY_CTN_pick = MAX_QTY_CTN_cust
            End If

            'Override if any combo exceeds max cartons to avoid looping. - disabled by wjz 05/11/13 - don't understand
            'Dim MAX_PICK As Int64 = Val(dst.Tables(IIf(ALLOW_SPLIT_TYPE = "S", "SOTSHIP3", "SOTSHIP2")) _
            '                               .Compute("MAX(PICK_QTY)", "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'") & "")
            'If MAX_QTY_CTN_cust < MAX_PICK Then
            '    ALLOW_SPLIT_TYPE = "A"
            'End If

            Dim CART_NO_ALL_lno As Integer = 0  ' Last Line Number Used for CART_NO_ALL
            Dim CART_PACK_QTY_ALL As Long       ' CART_PACK_QTY of the ALL Carton

            Dim no_mix As Boolean               ' True if RUN_TYPE is not null, implying that mixing is not permitted
            Dim CART_PACK_QTY As Long           ' Running Balance of Qty in Current Carton

            'CART_NO_seq = 0

            Dim vol_ctn_candidate As Boolean = False
            If ASCMAIN1.CLIENT = "VAN" AndAlso
                        ((CUST_CODE = "WALMART" And (EDI_PROMOTION = "POS REPLEN" Or EDI_PROMOTION.StartsWith("POSREPWK"))) Or (CUST_CODE = "KOHLS" And EDI_DEPT_DESC = "PACK BY STORE")) Then
                vol_ctn_candidate = True
            End If

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select _
                ("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                '  If rowSOTPICK1.Item("ORDR_NO") = "0002815634" Then Stop

                Dim CART_NO As String = ""
                Dim CART_NO_ALL As String = ""      ' Carton No to use for ALL Styles except for those with specific Carton Requirements

                'Need to know ORDR_TYPE_CODE for RGI, RGI XFR orders will create multiple cartons
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                Dim ORDR_TYPE_CODE As String = rowSOTORDR1.Item("ORDR_TYPE_CODE") & ""

                Dim rowSOTPCKC4_ORDR_NO As DataRow = Nothing ' Pre-Defined packing by Order
                If ASCMAIN1.CLIENT = "VAN" Then
                    rowSOTPCKC4_ORDR_NO = dst.Tables("SOTPCKC4_ORDR_NO").Rows.Find(ORDR_NO)
                End If

                If rowSOTPCKC4_ORDR_NO IsNot Nothing Then

#Region "If we have special cartonization rules for WALMART pre-packs"

                    Dim PACK_NO As String = rowSOTPCKC4_ORDR_NO.Item("PACK_NO")
                    Dim PACK_CONFIG_NO As String = rowSOTPCKC4_ORDR_NO.Item("PACK_CONFIG_NO")

                    ' rowSOTPICK1.Item("PACK_CONFIG_NO") = PACK_CONFIG_NO

                    Dim PACK_CART_NO As Integer = 0
                    Dim sql_pack As String = "PACK_NO = '" & PACK_NO & "' and PACK_CONFIG_NO = '" & PACK_CONFIG_NO & "'"

                    For Each rowSOTPCKC3 As DataRow In dst.Tables("SOTPCKC3").Select(sql_pack, "PACK_CART_NO,ORDR_LNO")
                        If Val(rowSOTPCKC3.Item("PACK_CART_NO") & "") <> PACK_CART_NO Then
                            PACK_CART_NO = Val(rowSOTPCKC3.Item("PACK_CART_NO") & "")
                            CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                        End If

                        Dim ORDR_LNO As Integer = Val(rowSOTPCKC3.Item("ORDR_LNO") & "")
                        Dim ORDR_QTY_PACK As Integer = Val(rowSOTPCKC3.Item("ORDR_QTY_PACK") & "")
                        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                        Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                        Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)

                        CART_PACK_QTY = CART_PACK_QTY + ORDR_QTY_PACK
                        CART_LNO_seq = CART_LNO_seq + 1

                        Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                        With rowSOTCART2
                            .Item("CART_NO") = CART_NO
                            .Item("CART_LNO") = CART_LNO_seq
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("STYLE_CODE") = STYLE_CODE
                            .Item("COLOR_CODE") = COLOR_CODE
                            .Item("QTY_PACKED") = ORDR_QTY_PACK
                            .Item("QTY_REL") = ORDR_QTY_PACK
                            .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                        End With
                        dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                    Next
#End Region

                Else

                    If vol_ctn_candidate Then
                        TAC.SOCMAIN1.Create_Cartons_For_PICK_NO(Me, PICK_NO, CART_NO_seq, False)
                    Else

#Region "Standard Cartonization"

                        ' max qty per carton
                        ' do not split styles
                        ' do not mix styles

                        LAST_STYLE = ""
                        LAST_COLOR = ""

                        Dim sortby As String = ""
                        If ALLOW_SPLIT_TYPE = "A" Then
                            sortby = "PICK_LNO"
                        Else
                            sortby = "STYLE_CODE,COLOR_CODE"
                        End If


                        Dim last_carton_full As Boolean = False

                        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select _
                            ("PICK_NO = '" & PICK_NO & "'", sortby)
                            'Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                            Dim ORDR_LNO As Int32 = Val(rowSOTPICK2.Item("ORDR_LNO") & "")
                            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE")
                            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE")
                            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                            Dim CARTON_PACK_QTY As Int64 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")

                            Dim MAX_QTY_CTN As Int64 = MAX_QTY_CTN_pick ' Maximum Unit Count in a Carton for a Customer for a Style/Color

                            ' THE FOLLOWING RULE, WHICH IS GOOD FOR ORDERS THAT ARE FULL CASES OF SOLID STYLE/COLORS, 
                            ' PROBABLY WORKS FOR MORE THAN JUST NYA, BUT WON'T WORK FOR VAN WHERE WE DO PICK AND PACK,
                            ' AND NEED CARTONIZATION AFTER RELEASE

                            If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then 'Or (ASCMAIN1.CLIENT = "RGI" And ORDR_TYPE_CODE = "XFR") Then
                                If CARTON_PACK_QTY <> 0 Then
                                    MAX_QTY_CTN = CARTON_PACK_QTY
                                End If
                            End If

                            no_mix = False
                            Dim rowSOTCSTP1 As DataRow = dst.Tables("SOTCSTP1").Rows.Find(New Object() {CUST_CODE, STYLE_CODE, COLOR_CODE})

                            If rowSOTCSTP1 IsNot Nothing Then
                                If rowSOTCSTP1.Item("NON_MIX") & "" = "1" Then
                                    no_mix = True ' CANNOT SUPPORT DIFFERENT MAX_QTY_PER_CTN VALUES ACROSS MULTIPLE STYLES
                                    'MAX_QTY_CTN = Val(tblSOWCSTP1.ITEM("MAX_QTY") & "") - DCG replaced this line with one below 4/26
                                    MAX_QTY_CTN = Val(rowSOTCSTP1.Item("MAX_QTY_CTN") & "")
                                End If
                            End If

                            'If ASCMAIN1.CLIENT = "RGI" And ORDR_TYPE_CODE = "XFR" Then
                            '    no_mix = True
                            'End If

                            QTY_TO_PACK = Val(rowSOTPICK2.Item("PICK_QTY") & "")
                            Dim pack_ctr As Integer = 0

                            'Check here if the next non-split violates the max_cart
                            Dim sqlx As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_NO = '" & PICK_NO & "'"

                            Select Case ALLOW_SPLIT_TYPE
                                Case "S"
                                    If LAST_STYLE <> STYLE_CODE Then
                                        sqlx &= " and STYLE_CODE = '" & STYLE_CODE & "'"
                                        NEXT_SPLIT_AMT = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", sqlx) & "")
                                    End If

                                Case "C"
                                    If LAST_STYLE <> STYLE_CODE Or LAST_COLOR <> COLOR_CODE Then
                                        sqlx &= " and STYLE_CODE = '" & STYLE_CODE & "'"
                                        sqlx &= " and COLOR_CODE = '" & COLOR_CODE & "'"
                                        NEXT_SPLIT_AMT = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", sqlx) & "")
                                    End If

                                Case "A"
                                    NEXT_SPLIT_AMT = 0
                            End Select

                            Do While QTY_TO_PACK <> 0
                                pack_ctr += 1

                                If no_mix Or last_carton_full Then
                                    CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                Else
                                    If CART_NO_ALL <> "" Then
                                        CART_NO = CART_NO_ALL
                                        CART_LNO_seq = CART_NO_ALL_lno
                                        CART_PACK_QTY = CART_PACK_QTY_ALL
                                        If MAX_QTY_CTN <> 0 Then
                                            If CART_PACK_QTY >= MAX_QTY_CTN Then
                                                CART_NO_ALL = ""
                                                CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                            ElseIf pack_ctr = 1 And CART_PACK_QTY > MAX_QTY_CTN / 2 And CART_PACK_QTY = QTY_TO_PACK And 1 = 1 Then ' wait for gerry v
                                                CART_NO_ALL = ""
                                                CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                            Else
                                                ' If last_carton_full Then Stop

                                            End If
                                        End If

                                        'If MAX_QTY_CTN <> 0 And CART_PACK_QTY >= MAX_QTY_CTN Then
                                        '    CART_NO_ALL = ""
                                        '    CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                        'End If
                                    Else
                                        CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                    End If
                                End If

                                If NEXT_SPLIT_AMT <> 0 Then
                                    If (NEXT_SPLIT_AMT + CART_PACK_QTY) > MAX_QTY_CTN And CART_LNO_seq <> 0 Then
                                        ' CART_LNO_seq clause added because an empty carton would result if we just created a new carton a few lines above
                                        'CART_NO = CART_NO_ALL
                                        'CART_LNO_seq = CART_NO_ALL_lno
                                        'CART_PACK_QTY = CART_PACK_QTY_ALL
                                        CART_NO_ALL = ""
                                        CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                    End If
                                End If

                                If MAX_QTY_CTN <> 0 And QTY_TO_PACK > MAX_QTY_CTN - CART_PACK_QTY Then
                                    PACK_QTY = MAX_QTY_CTN - CART_PACK_QTY
                                Else
                                    PACK_QTY = QTY_TO_PACK
                                End If
                                QTY_TO_PACK = QTY_TO_PACK - PACK_QTY
                                CART_PACK_QTY = CART_PACK_QTY + PACK_QTY

                                CART_LNO_seq = CART_LNO_seq + 1
                                Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                                With rowSOTCART2
                                    .Item("CART_NO") = CART_NO
                                    .Item("CART_LNO") = CART_LNO_seq
                                    .Item("ORDR_NO") = ORDR_NO
                                    .Item("ORDR_LNO") = ORDR_LNO
                                    .Item("STYLE_CODE") = STYLE_CODE
                                    .Item("COLOR_CODE") = COLOR_CODE
                                    .Item("QTY_PACKED") = PACK_QTY
                                    .Item("QTY_REL") = PACK_QTY
                                    .Item("STYLE_WEIGHT") = rowICTSTYL1.Item("STYLE_WEIGHT")
                                End With

                                dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)

                                LAST_STYLE = STYLE_CODE
                                LAST_COLOR = COLOR_CODE
                                NEXT_SPLIT_AMT = 0

                                If ASCMAIN1.CLIENT = "NYA" Then ' PROB SHOULD OPEN THIS UP TO ALL
                                    If CART_PACK_QTY = MAX_QTY_CTN And MAX_QTY_CTN <> 0 Then
                                        last_carton_full = True
                                    End If
                                End If

                                If no_mix Then
                                    CART_NO = ""
                                Else
                                    If CART_NO_ALL = "" Then
                                        CART_NO_ALL = CART_NO
                                    End If
                                    CART_NO_ALL_lno = CART_LNO_seq
                                    CART_PACK_QTY_ALL = CART_PACK_QTY
                                End If
                            Loop
                        Next
#End Region

                    End If
                End If
            Next
        Next

        Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
        dst.Tables("SOTCART2").Columns.Add("WGT", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(STYLE_WEIGHT,0)")
        dst.Tables("SOTCART1").Columns.Add("QTY", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        dst.Tables("SOTCART1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).WGT)")
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("QTY")
            rowSOTCART1.Item("CART_TOTAL_UNITS_REL") = rowSOTCART1.Item("QTY")
            rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTCART1.Item("WGT")
        Next

        Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
        dst.Tables("SOTPICK1").Columns.Add("CTNS", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
        dst.Tables("SOTPICK1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_CALC)")
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("CTNS")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("WGT")
        Next


        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_CNT_CARTONS = 0")
            rowSOTPICK1.Item("PICK_STATUS") = "C"
        Next

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_C", GetType(System.Int64), "IIF(PICK_STATUS = 'C',1,0)")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_P", GetType(System.Int64), "IIF(PICK_STATUS = 'P',1,0)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_C", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_C)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_P", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_P)")

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("PICKS_C >0 AND PICKS_P =0")
            rowSOTSHIP1.Item("SHIP_STATUS") = "C"
        Next

    End Sub

    Function New_Carton(PICK_NO As String, ByRef CART_NO_seq As Int32) As String
        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        CART_NO_seq += 1
        Dim CART_NO As String = "TEMP" & Format(CART_NO_seq, "000000")
        rowSOTCART1.Item("CART_NO") = CART_NO
        rowSOTCART1.Item("PICK_NO") = PICK_NO
        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
        Return CART_NO
    End Function

    Function Check_OOB_Styles() As Boolean

        ASCMAIN1.Progress("Checking for out of Balance Styles", "")

        ASCMAIN1.sql = "Select WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", Sum(ORDR_OPEN) AS ORDR_OPEN, Sum(STAT_OPEN) AS STAT_OPEN" & vbCrLf _
            & ", Sum(ORDR_PICK) AS ORDR_PICK, Sum(STAT_PICK) AS STAT_PICK from (" & vbCrLf _
            & "Select WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", Sum(NVL(OPEN_ORDERO,0) + NVL(OPEN_ORDERR,0)) AS ORDR_OPEN" & vbCrLf _
            & ", Sum(STAT_OPEN) AS STAT_OPEN" & vbCrLf _
            & ", 0 AS ORDR_PICK, 0 AS STAT_PICK" & vbCrLf _
            & " from (" & vbCrLf _
            & " Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", Sum(NVL(SOTORDR2.ORDR_QTY_OPEN,0)) AS OPEN_ORDERO, 0 AS OPEN_ORDERR, 0 AS STAT_OPEN" & vbCrLf _
            & " from SOTORDR2,SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO and SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select SOTRSRV1.WHSE_CODE, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
            & ", 0 AS OPEN_ORDERO, Sum(NVL(SOTRSRV2.RSRV_QTY_OPEN,0)) AS OPEN_ORDERR, 0 AS STAT_OPEN" & vbCrLf _
            & " from SOTRSRV2,SOTRSRV1" & vbCrLf _
            & " where SOTRSRV1.RSRV_NO = SOTRSRV2.RSRV_NO" & vbCrLf _
            & " group by SOTRSRV1.WHSE_CODE, SOTRSRV2.STYLE_CODE, SOTRSRV2.COLOR_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ", 0 AS OPEN_ORDERO, 0 AS OPEN_ORDERR, Sum(NVL(WHSE_QTY_OPEN,0)) AS STAT_OPEN" & vbCrLf _
            & " from ICTSTAT2 group by WHSE_CODE, STYLE_CODE, COLOR_CODE)" & vbCrLf _
            & " having Sum(NVL(OPEN_ORDERO, 0) + NVL(OPEN_ORDERR, 0)) - SUM(NVL(STAT_OPEN, 0)) <> 0" & vbCrLf _
            & " group by WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select WHSE_CODE, STYLE_CODE, COLOR_CODE, 0 AS ORDR_OPEN, 0 AS STAT_OPEN" & vbCrLf _
            & ", Sum(NVL(PICK_ORDER,0)) AS ORDR_PICK" & vbCrLf _
            & ", Sum(NVL(STAT_PICK, 0)) As STAT_PICK" & vbCrLf _
            & " from (" & vbCrLf _
            & " Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", Sum(NVL(ORDR_QTY_PICK,0)) AS PICK_ORDER, SUM(0) AS STAT_PICK" & vbCrLf _
            & " from SOTORDR2, SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and (SOTORDR2.ORDR_STATUS = 'O' or SOTORDR2.ORDR_STATUS = 'P')" & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & " Union" & vbCrLf _
            & " Select WHSE_CODE, STYLE_CODE, COLOR_CODE, 0 AS PICK_ORDER, Sum(NVL(WHSE_QTY_PICK,0)) AS STAT_PICK" & vbCrLf _
            & " from ICTSTAT2 group by WHSE_CODE, STYLE_CODE, COLOR_CODE)" & vbCrLf _
            & "Having SUM(NVL(PICK_ORDER, 0)) - SUM(NVL(STAT_PICK, 0)) <> 0" & vbCrLf _
            & " group by WHSE_CODE, STYLE_CODE, COLOR_CODE" & vbCrLf _
            & ") group by WHSE_CODE, STYLE_CODE, COLOR_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTROOB1", 3))

        'Fill_Records("SOTROOB1", "", True, ASCMAIN1.sql)

        'Check table.  If there is anything there then set variable that causes error report to print.
        Return (dst.Tables("SOTROOB1").Rows.Count > 0)
    End Function

    Sub Inventory_Shortages()

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
        Else

            ' Make all E's show up as current allocations
            ' note we are not dealing with F's here - maybe we should

            If blnRELEASE_FUT Then
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is" & vbCrLf _
                    & "  Select * from " & SOTORDR2 & vbCrLf _
                    & "   where ORDR_QTY_ALLO_FUT <> 0 and ORDR_RELEASE_AVAIL is Not Null" & vbCrLf _
                    & "     and ORDR_RELEASE_AVAIL <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
                    & "   For Update;" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update " & SOTORDR2 & vbCrLf _
                    & "    Set ORDR_QTY_ALLO_CUR = NVL(ORDR_QTY_ALLO_CUR,0) + NVL(R1.ORDR_QTY_ALLO_FUT,0)" & vbCrLf _
                    & "     ,  ORDR_QTY_ALLO_FUT = 0" & vbCrLf _
                    & "    where Current of C1;" & vbCrLf _
                    & "  End Loop; " & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                'ASCMAIN1.sql = "Update " & SOTORDR2 & vbCrLf _
                '    & " Set ORDR_QTY_ALLO_CUR = NVL(ORDR_QTY_ALLO_CUR,0) + NVL(ORDR_QTY_ALLO_FUT,0)" & vbCrLf _
                '    & " where SOTORDR2.ORDR_RELEASE_AVAIL <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'"
                'ASCDATA1.ExecuteSQL()
                'ASCMAIN1.sql = "Update " & SOTORDR2 & vbCrLf _
                '    & " Set ORDR_QTY_ALLO_FUT = 0" & vbCrLf _
                '    & " where SOTORDR2.ORDR_RELEASE_AVAIL <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'"
                'ASCDATA1.ExecuteSQL()
            End If
        End If

        ' I: Hold all orders with Inventory Shortage Conditions where ORDR_RELEASE is Null

        If manual_release Then
        Else
            'If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            ASCMAIN1.sql = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
                & " from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
                & " where NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> NVL(SOTORDR2.ORDR_QTY_ALLO_CUR,0)" & vbCrLf _
                & "   and SOTORDR2.ORDR_RELEASE is Null"
            'Else
            '    ASCMAIN1.sql = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
            '        & " from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
            '        & " where NVL(SOTORDR2.ORDR_QTY_OPEN,0) <> NVL(SOTORDR2.ORDR_QTY_ALLO_CUR,0) " _
            '        & IIf(chkRelFutAvail.Checked, "+ CASE WHEN NVL(SOTORDR2.ORDR_RELEASE_AVAIL,SYSDATE) <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "' THEN NVL(SOTORDR2.ORDR_QTY_ALLO_FUT,0) ELSE 0 END", "") & vbCrLf _
            '        & "   and SOTORDR2.ORDR_RELEASE is Null"
            'End If

            ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'I'" & vbCrLf _
                & " where ORDR_NO in (" & ASCMAIN1.sql & ")"
            ASCDATA1.ExecuteSQL()
        End If

        If blnALLOCATION_ONLY Then
            ASCMAIN1.sql = "Select Distinct SOTRSRV2.RSRV_NO" _
                & " from " & SOTRSRV2 & " SOTRSRV2" _
                & " where SOTRSRV2.RSRV_QTY_OPEN <> SOTRSRV2.ORDR_QTY_ALLO_CUR"
            ASCMAIN1.sql = "Update " & SOTRSRV1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'I'" _
                & " where RSRV_NO in (" & ASCMAIN1.sql & ")"
            ASCDATA1.ExecuteSQL()
        End If

        ' E: Hold all orders containing styles whose demand was allocated
        '      by an expected shipment arrival if the calculated in whse date
        '      is later than the SHIP_BY_DATE used to run the release
        ' F: Hold all orders containing styles whose demand was allocated
        '      by an expected shipment arrival if the calculated in whse date
        '      is later than the ORDR_CANCEL_DATE of the Order

        Dim release_E As Boolean = True
        Dim release_F As Boolean = True

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            If Not blnRELEASE_FUT Then
                release_E = False
                release_F = False
            End If
        End If


        If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
            If ASCMAIN1.DBS_SERVER <> ASCMAIN1.DBS_COMPANY Then
                If blnRELEASE_FUT Then
                    release_E = False
                End If
            End If
        End If

        ' These conditions are there to prevent rows from coming up in this result set so that they can be marked for shortage conditions E and F
        ' so, for example, if we are NOT permitting Es to release, then we don't want to allow orders to NOT get marked based on the condition implied in the sql
        Dim SQL_LET_EF_GO As String = ""
        If release_E Then SQL_LET_EF_GO &= " or SOTORDR2.ORDR_RELEASE_AVAIL > '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'"
        If release_F Then SQL_LET_EF_GO &= " or SOTORDR2.ORDR_RELEASE_AVAIL > SOTORDR1.ORDR_CANCEL_DATE"
        If SQL_LET_EF_GO <> "" Then
            SQL_LET_EF_GO = " and (" & Mid(SQL_LET_EF_GO, 5) & ")" & vbCrLf
        End If
        '& "   and (SOTORDR2.ORDR_RELEASE_AVAIL > '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
        '& "     or SOTORDR2.ORDR_RELEASE_AVAIL > SOTORDR1.ORDR_CANCEL_DATE) " & vbCrLf _



        ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", Max (SOTORDR2.ORDR_RELEASE_AVAIL) ORDR_RELEASE_AVAIL" & vbCrLf _
            & " from " & SOTORDR2 & " SOTORDR2, " & SOTORDR1 & " SOTORDR1, " & ARTCUST1 & " ARTCUST1" & vbCrLf _
            & " where SOTORDR2.ORDR_RELEASE_AVAIL is Not Null " & vbCrLf _
            & "   and (SOTORDR2.ORDR_RELEASE is Null or (SOTORDR2.ORDR_RELEASE <> 'C' and SOTORDR2.ORDR_RELEASE <> 'S'))" & vbCrLf _
            & SQL_LET_EF_GO _
            & "    and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "    and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & " group by SOTORDR1.ORDR_NO, SOTORDR1.ORDR_CANCEL_DATE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ORDR_NO As String = row.Item("ORDR_NO")
            Dim ORDR_REL_HOLD_CODE As String = ""
            If row.Item("ORDR_CANCEL_DATE") & "" = "" OrElse Format(row.Item("ORDR_RELEASE_AVAIL"), "yyyyMMdd") _
             > Format(row.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") Then
                ORDR_REL_HOLD_CODE = "F"
            Else
                ORDR_REL_HOLD_CODE = "E"
            End If
            ASCMAIN1.sql = "Update " & SOTORDR1 _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || '" & ORDR_REL_HOLD_CODE & "'" _
                & " where ORDR_NO = '" & ORDR_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next

        If blnALLOCATION_ONLY Then
            ASCMAIN1.sql = "Select SOTRSRV1.RSRV_NO, SOTRSRV1.ORDR_CANCEL_DATE" & vbCrLf _
                & ", Max (SOTRSRV2.ORDR_RELEASE_AVAIL) ORDR_RELEASE_AVAIL" & vbCrLf _
                & " from " & SOTRSRV2 & " SOTRSRV2, " & SOTRSRV1 & " SOTRSRV1, " & ARTCUST1 & " ARTCUST1" & vbCrLf _
                & " where SOTRSRV2.ORDR_RELEASE_AVAIL is Not Null " & vbCrLf _
                & "   and (SOTRSRV2.ORDR_RELEASE_AVAIL > '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "'" & vbCrLf _
                & "     or SOTRSRV2.ORDR_RELEASE_AVAIL > SOTRSRV1.ORDR_CANCEL_DATE) " & vbCrLf _
                & "    and SOTRSRV2.RSRV_NO = SOTRSRV1.RSRV_NO" & vbCrLf _
                & "    and ARTCUST1.CUST_CODE = SOTRSRV1.CUST_CODE" & vbCrLf _
                & " group by SOTRSRV1.RSRV_NO, SOTRSRV1.ORDR_CANCEL_DATE"

            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim RSRV_NO As String = row.Item("RSRV_NO")
                Dim ORDR_REL_HOLD_CODE As String = ""
                If Format(row.Item("ORDR_RELEASE_AVAIL"), "yyyyMMdd") _
                 > Format(row.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") Then
                    ORDR_REL_HOLD_CODE = "F"
                Else
                    ORDR_REL_HOLD_CODE = "E"
                End If
                ASCMAIN1.sql = "Update " & SOTRSRV1 _
                    & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || '" & ORDR_REL_HOLD_CODE & "'" _
                    & " where RSRV_NO = '" & RSRV_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If

        'maybe we should not be overriding code F

        Dim sql As String = "Select Distinct SOTORDRG.ORDR_GROUP_NO" _
            & " from " & SOTORDR1 & " SOTORDR1, SOTORDRG" _
            & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" _
            & "   and SOTORDRG.ORDR_REL_SHORT = '1'" _
            & "   and NVL(SOTORDR1.ORDR_SHIP_COMPLETE,'0') = '0'" _
            & "   and (INSTR(SOTORDR1.ORDR_REL_HOLD_CODES,'I') <> 0 " _
            & "     or INSTR(SOTORDR1.ORDR_REL_HOLD_CODES,'E') <> 0 " _
            & "     or INSTR(SOTORDR1.ORDR_REL_HOLD_CODES,'F') <> 0)"
        ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1" _
            & " Set ORDR_REL_HOLD_CODES = REPLACE(REPLACE(REPLACE(ORDR_REL_HOLD_CODES,'I',''),'E',''),'F','')" _
            & " where ORDR_GROUP_NO in (" & sql & ")"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & SOTORDR1 & " SOTORDR1" _
            & " Set ORDR_REL_HOLD_CODES = NULL" _
            & " where ORDR_GROUP_NO in (" & sql & ")" _
            & "   and ORDR_REL_HOLD_CODES = ''"
        ASCDATA1.ExecuteSQL()
        'For Each rowSOTORDR6 As DataRow In ASCDATA1.GetDataTable.Select("")
        '    Dim ORDR_GROUP_NO As String = rowSOTORDR6.Item("ORDR_GROUP_NO")

        'Next

        If blnALLOCATION_ONLY Then
            'D - Past Cancel orders - Added to report for Gabe by WR on 4/17/02.
            ' Stop ' note that the following lines are not using customer specific cancel grace days

            ASCMAIN1.sql = "Update " & SOTORDR1 & vbCrLf _
                & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'D'" & vbCrLf _
                & " where ORDR_CANCEL_DATE  < '" & Format(DATETIME_STAMP.Date.AddDays(numCANCEL_FUTURE_DAYS), "dd-MMM-yyyy") & "'"
            ASCDATA1.ExecuteSQL()
        Else
            ' D: Hold Orders where Cancel Date is prior to Today + Cancel Grace Days
            If Not blnREL_PAST_CANCEL Then

                Dim ORDR_CANCEL_DATE_cutoff As Date = DATETIME_STAMP.AddDays(numCANCEL_FUTURE_DAYS)
                ASCMAIN1.sql = "Update " & SOTORDR1 _
                    & " Set ORDR_REL_HOLD_CODES = NVL(ORDR_REL_HOLD_CODES,'') || 'D'" _
                    & " where ORDR_CANCEL_DATE < '" & Format(ORDR_CANCEL_DATE_cutoff, "dd-MMM-yyyy") & "'"
                ASCDATA1.ExecuteSQL()

            End If
        End If
    End Sub

    Public Overrides Sub Print_Report()

        If Absx1.chkFor("CHKCHECK_OOBAL").Checked Then
            If Check_OOB_Styles() Then
                RPT = "SORROOB1"
                RPT_TITLE = "Style Out of Balance Report"
                SUBT = "Please Forward A Copy of This Report to ABS"
                'CR_params.Add("OPTDTL", Absx1.optFor("OPTDTL").Value)
                Generate_Report(RPT, RPT_TITLE, SUBT)
            End If
        End If

        ' Get List of Warehouse Codes which had Shipments Released
        Dim WHSE_CODEs As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("SOTSHIP1"), New String() {"WHSE_CODE"}).Select("", "WHSE_CODE")
            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
            WHSE_CODEs.Add(WHSE_CODE)
        Next

        If blnALLOCATION_ONLY Then
            If chkCheck_CANCEL.Checked _
                AndAlso (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") _
                AndAlso Absx1.txtFor("WHSE_CODE").Text = "MS" Then
                Try
                    ASCDATA1.ExecuteSQL("TRUNCATE TABLE SOTORDRC")
                    Dim DTE0 As String = Absx1.dteFor("DTE0").DateTime.ToString("dd-MMM-yyyy")
                    Dim DTE1 As String = Absx1.dteFor("DTE1").DateTime.ToString("dd-MMM-yyyy")

                    ASCDATA1.ExecuteSQL("INSERT INTO SOTORDRC (ORDR_NO, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN, ORDR_QTY_ALLO_CUR, SELECTED, SELECTED_DTL) " & vbCrLf _
                        & " SELECT ORDR_NO, ORDR_LNO, STYLE_CODE, COLOR_CODE, ORDR_QTY_OPEN, ORDR_QTY_ALLO, '1','1'" & vbCrLf _
                        & " from " & SOTORDR2 & " SOTORDR2" & vbCrLf _
                        & " where NVL(SOTORDR2.ORDR_QTY_ALLO,0) < NVL(SOTORDR2.ORDR_QTY_OPEN,0)" & vbCrLf _
                        & " and SOTORDR2.ORDR_NO IN (SELECT ORDR_NO FROM " & SOTORDR1 & " WHERE ORDR_REL_HOLD_CODES LIKE '%I%' AND WHSE_CODE = 'MS'" & vbCrLf _
                        & " and ORDR_DATE BETWEEN '" & DTE0 & "' and '" & DTE1 & "')")
                Catch ex As Exception
                    MessageBox.Show("Error setting up Cancellation Letters Information: " & ex.Message, "Letters", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try
            End If
        Else
            For Each WHSE_CODE As String In WHSE_CODEs
                RPT = "SOROREL1"
                RPT_TITLE = "Released Orders Report"
                If PICK_BATCH_NO <> "" Then
                    SUBT = "Batch " & PICK_BATCH_NO
                Else
                    SUBT = ""
                End If
                If blnFORCE_PICK Then
                    SUBT &= " (Force Picked)"
                End If
                If manual_release Then
                    SUBT &= " (Manual Release)"
                End If
                SUBT &= "  Whse " & WHSE_CODE & "  Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
                Generate_Report(RPT, RPT_TITLE, SUBT, "{SOTORDR1.WHSE_CODE}='" & WHSE_CODE & "'")

                ' Special things for Regency Int'l
                If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
                    If dst.Tables.Contains("SOTOREMM") Then
                        If Not blnFORCE_PICK AndAlso Not manual_release Then
                            RPT = "SORORELM"
                            RPT_TITLE = "Merged Orders Report"
                            Generate_Report(RPT, RPT_TITLE, SUBT, "{SOTOREMM.WHSE_CODE}='" & WHSE_CODE & "'")
                        End If
                    End If
                End If
            Next

            If Not blnFORCE_PICK Then
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    If WHSE_CODEs.Count > 0 Then
                        RPT = "SOROREL8"
                        RPT_TITLE = "Released Minimum Orders Report"
                        SUBT = "Orders Falling Below $25"
                        CR_params.Add("MINVAL", 25)
                        Generate_Report(RPT, RPT_TITLE, SUBT)
                    End If
                End If
            End If

            'If (ASCMAIN1.CLIENT = "RGI") Then
            '    RPT = "SORORECC"
            '    RPT_TITLE = "Released Orders Credit Card Processing Report - Declines"
            '    SUBT = "Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
            '    Generate_Report(RPT, RPT_TITLE, SUBT)
            'End If
        End If


        If (ASCMAIN1.CLIENT = "VAN") Then
            RPT = "SORALEX1"
            RPT_TITLE = "Allocation Exclusions Report"
            SUBT = "Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
            Generate_Report(RPT, RPT_TITLE, SUBT)
        End If


        If blnFORCE_PICK Then
        Else
            'If Not blnALLOCATION_ONLY Then
            ' this report needs to be available in allocation only mode - but the ADO.NET tables are not being updated so that the criteria of the report may be satisfied
            '{SOTORDR2.ORDR_QTY_ALLO}<{SOTORDR2.ORDR_QTY_OPEN} and
            '(Instr({SOTORDR1.ORDR_REL_HOLD_CODES},"I") <> 0 or
            ' Instr({SOTORDR1.ORDR_REL_HOLD_CODES},"E") <> 0 or
            ' Instr({SOTORDR1.ORDR_REL_HOLD_CODES},"F") <> 0)
            RPT = "SOROREL3"
            RPT_TITLE = "Inventory Shortage Report"
            ' THIS SHOULD BE OBSERVING THE OPTION TO SHOW (OR NOT SHOW) ORDERS PAST CANCEL
            If Absx1.chkFor("CHKSHOW_PAST_CANCEL").Checked Then
            End If
            'If blnALLOCATION_ONLY Then
            'SUBT = "Unshippable Orders w/Cancel Date within the next " & CStr(Absx1.numFor("FPDSHORT_HOR_DAYS").Value) & " days"
            'Else
            SUBT = "Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
            'End If

            Generate_Report(RPT, RPT_TITLE, SUBT)
            'End If

            RPT = "SORORELA"
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                RPT = "SOROREL4"
            End If

            If blnALLOCATION_ONLY Then
                If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                    RPT_TITLE = "Orders with Inventory Shortages"
                Else
                    RPT_TITLE = "Un-Releasable Orders Report"
                End If

                SUBT = "Un-Releasable Orders w/Cancel Date within the next " & CStr(Absx1.numFor("FPDSHORT_HOR_DAYS").Value) & " days"
                SUBT = "Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
            Else
                RPT_TITLE = "Orders Not Released Detail Report"
                SUBT = "Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy")
            End If

            CR_params.Add("SORT_SREP", "0")
            CR_params.Add("RELEASE_DATE", Format$(SHIP_BY_DATE, "yyyyMMdd"))
            If blnALLOCATION_ONLY Then
                CR_params.Add("CHKALLOCATION_ONLY", "1")
            Else
                CR_params.Add("CHKALLOCATION_ONLY", "0")
            End If

            Generate_Report(RPT, RPT_TITLE, SUBT)
        End If

        If chkemailSReps.Checked Then
            email_SREP_CODEs()
        End If


        If blnALLOCATION_ONLY Then
        Else
            ' Update_ADS_Tables()
        End If

        Try
            If ASCMAIN1.CLIENT = "NYA" And WHSE_CODEs.Contains("18") Then
                Dim clsASCNOTE1 As New TAC.ASCNOTE1("SOREL18", dst)
                clsASCNOTE1.Note = dst.Tables("SOTSHIP1").Select("WHSE_CODE = '18'").Length & " shipments were released for picking."
                clsASCNOTE1.CreateComponents()
                clsASCNOTE1.EmailDocument()
                clsASCNOTE1 = Nothing
            End If

        Catch ex As Exception
            MessageBox.Show("Error emailing Warehouse 18 there were orders released." & ex.Message, "Email Error", MessageBoxButtons.OK)
        End Try

    End Sub

    Sub email_SREP_CODEs()
        '  Exit Sub
        For Each rowSREP As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), New String("SREP_CODE")).Select("")
            Dim SREP_CODE As String = rowSREP.Item("SREP_CODE")
            Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", SREP_CODE)

            RPT = "SORORELA"
            RPT_TITLE = "Un-Releasable Orders Report"
            CR_params.Add("SORT_SREP", "0")
            CR_params.Add("RELEASE_DATE", Format$(SHIP_BY_DATE, "yyyyMMdd"))
            CR_params.Add("CHKALLOCATION_ONLY", "1")
            Dim FILENAME_body As String = RPT & "_" & SREP_CODE
            Dim REPORT_NO As String = Generate_Report(RPT, RPT_TITLE, SUBT, _
                                                      "{SOTORDRS.SREP_CODE} = '" & SREP_CODE & "'", _
                                                      "PDF", FILENAME_body, False)

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(FILENAME_body & ".pdf", ASCMAIN1.Folders("Temp") & FILENAME_body & ".pdf")

            Dim SUBJECT As String = RPT_TITLE
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
            Dim EMAIL_ADDRESS As String = rowSOTSREP1.Item("SREP_EMAIL") & ""
            EMAIL_ADDRESS = "wjz@absolution.com"
            EMAIL_ADDRESSs.Add(EMAIL_ADDRESS, rowSOTSREP1.Item("SREP_NAME") & "")

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                    SUBJECT, "SORORELA", True, True, SREP_CODE, rowSOTSREP1.Item("SREP_NAME") & "", "Sales Rep")
        Next

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If ASCMAIN1.CLIENT = "VAN" Then
            ' **************** NOT A GOOD IDEA - DO NOT ALLOW THIS
            '    If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and COLUMN_NAME <> 'CUST_CODE' and EXCLUDE = '1'").Length <> 0 Then
            '        EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group & Customer"
            '    End If
            'Else
            If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and EXCLUDE = '1'").Length <> 0 Then
                EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group"
            End If
            'End If

            If ASCMAIN1.CLIENT = "VAN" Then
                Dim row As DataRow
                For Each row In tblASTDSQLA.Select("EXCLUDE='1'")
                    Dim CODE_VALUES As String = row.Item("CODE_VALUES") & ""
                    '  CODE_VALUES = Mid(CODE_VALUES, 2, CODE_VALUES.Length - 2)
                    Dim CODE_VALUE_LIST() As String = CODE_VALUES.Split(",")
                    If CODE_VALUE_LIST.Length > 25 Then
                        EMsg &= vbCr & "You may not use Exclusion for more than 25 items"
                    End If
                Next


                Dim WALMART_customer_selected As Boolean = False

                row = tblASTDSQLA.Rows.Find("CUST_CODE")
                ' remember - you cannot exclude customers, so this is always a list of customers to include
                Dim CUST_CODEs() As String = Split(row.Item("CODE_VALUES") & "", ",")
                If CUST_CODEs.Contains("WALMART") Then
                    WALMART_customer_selected = True
                    If CUST_CODEs.Length > 1 Then
                        EMsg &= vbCr & "You must select WALMART all by itself (no other customers)"
                    End If
                End If

                row = tblASTDSQLA.Rows.Find("ORDR_GROUP_NO")
                Dim OG As String = row.Item("CODE_VALUES") & ""
                Dim sqlOG As String = "'" & Replace(OG, ",", "','") & "" & "'"
                Dim EXCLUDE As String = row.Item("EXCLUDE") & ""

                Dim WALMART_groups_selected As Boolean = False

                If OG <> "" Then
                    'If EXCLUDE = "1" Then
                    Dim ORDR_GROUP_NOs() As String = Split(OG, ",")

                    ASCMAIN1.sql = "Select Distinct CUST_CODE from SOTORDR0 where ORDR_GROUP_NO in (" & sqlOG & ")"
                    Dim OG_CUSTs As New List(Of String)
                    For Each row In ASCDATA1.GetDataTable.Select("")
                        OG_CUSTs.Add(row.Item("CUST_CODE"))
                    Next

                    If OG_CUSTs.Contains("WALMART") Then
                        WALMART_groups_selected = True
                        If OG_CUSTs.Count > 1 Then
                            EMsg &= vbCr & "You must select WALMART orders without mixing other customers"
                        End If
                    End If
                    'Else
                    'End If

                    ''ASCMAIN1.sql = "Select Distinct EDI_CONS_NO from EDT850T1 where ORDR_GROUP_NO in (" & sqlOG & ")"
                    ''Dim EDI_CONS_NOs As New List(Of String)
                    ''For Each row In ASCDATA1.GetDataTable.Select("")
                    ''    EDI_CONS_NOs.Add(row.Item("EDI_CONS_NO"))
                    ''Next

                    ASCMAIN1.sql = "SELECT DISTINCT SOTORDR0.ORDR_GROUP_NO,EDT850T1.EDI_CONS_NO,SOTORDR0.CUST_DC_NO FROM SOTORDR0,EDT850T1" _
                     & " WHERE EDT850T1.EDI_DOC_SEQ_NO(+) = SOTORDR0.EDI_DOC_SEQ_NO" _
                     & " AND (EDT850T1.EDI_CONS_NO,SOTORDR0.CUST_DC_NO) IN (" _
                     & " SELECT DISTINCT EDT850T1.EDI_CONS_NO,SOTORDR0.CUST_DC_NO FROM SOTORDR0,EDT850T1" _
                     & " WHERE EDT850T1.EDI_DOC_SEQ_NO(+) = SOTORDR0.EDI_DOC_SEQ_NO" _
                     & " AND ORDR_GROUP_NO in (" & sqlOG & "))" _
                     & " AND ORDR_GROUP_NO NOT in (" & sqlOG & ")"


                    'Dim EDI_CONS_NOs As New List(Of String)
                    For Each row In ASCDATA1.GetDataTable.Select("")
                        EMsg &= vbCr & $"Missing Order Group No {row.Item("ORDR_GROUP_NO")} For DC {row.Item("CUST_DC_NO")} When Releasing Multi-PO"
                    Next



                End If

                If WALMART_customer_selected And Not chkForcePick.Checked Then
                    EMsg &= vbCr & "You must Force Pick WALMART orders"
                End If

            End If


            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("ORDR_GROUP_NO")

            If Absx1.chkFor("CHKFORCE_PICK").Checked Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    EMsg &= vbCr & "You Must Select Specific Order Groups To Force Pick"
                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Force Picking, you must Select (Not Exclude) Order Groups"
                End If
            End If

            If ASCMAIN1.CLIENT = "RGI" Then
                If Absx1.chkFor("CHKMANUAL_ONLY").Checked Then
                    If DirectCast(grdSOTORDRU.DataSource, DataTable).Select("SEL='1'").Length = 0 Then
                        EMsg &= vbCr & "You Must Select Specific Users when choosing Manually Selected only"
                    End If
                End If
            End If


            If Absx1.chkFor("CHKREL_PAST_CANCEL").Checked Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    If ASCMAIN1.CLIENT = "RGI" And Absx1.chkFor("CHKMANUAL_ONLY").Checked Then
                        If EMsg = "" Then
                            MsgBox("Please be aware that there may be many Orders Past Cancel in your batch of orders", MsgBoxStyle.OkOnly, "Verification")
                        End If
                    Else
                        EMsg &= vbCr & "You Must Select Specific Order Groups to Release Past Cancel"
                    End If

                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Releasing Past Cancel, you must Select (not Exclude) Order Groups"
                End If
            End If

            If optWHSE.Value = "S" Then
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Ship-From Warehouse Specified"
                End If
            End If

            If chkECommerce.Checked Then
                If optWHSE.Value = "A" Then
                    EMsg &= vbCr & "e-Commerce Release must specify a single Warehouse"
                End If

                If Absx1.txtFor("ECOM_CODE").Text = "" Then
                    ' 03/21/2019 - Allow to release for Ecom Sales Orders
                    ' EMsg &= vbCr & "e-Commerce Release must specify a single e-Commerce Partner (non specified)"
                Else
                    If LookUp("ECTECOM1", Absx1.txtFor("ECOM_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "e-Commerce Release must specify a valid e-Commerce Partner"
                    End If
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")

        ' Set_WHSE()
        ' Set_Date()

        'Dim sqlw As String = CStr(parms(0))
        'If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        ' Dim sqlw As String = ""
        Dim ORDR_GROUP_NOs_to_release As New List(Of String)

        If parms.Length > 0 Then
            ORDR_GROUP_NOs_to_release = parms(0)
        End If

        'EnforceConstraints(False)
        'Fill_Records("SOTORDR1")
        'EnforceConstraints(True)

        WHSE_CODE = ""
        ORDR_GROUP_NO_sql = "'" & Join(ORDR_GROUP_NOs_to_release.ToArray, "','") & "'"
        CUST_CODE_sql = ""
        SALES_DIVISION_CODE_sql = ""
        TERM_CODE_sql = ""
        SHIP_BY_DATE = Now.Date.AddYears(1)

        blnALLOCATION_ONLY = False
        blnFORCE_PICK = False
        selWHSE = "A"

        SQL_ins.Clear()
        SQL_ins.Add("CUST_CODE", "")
        SQL_ins.Add("SALES_DIVISION_CODE", "")
        SQL_ins.Add("ORDR_GROUP_NO", "")
        If (ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI") Then
            SQL_ins.Add("TERM_CODE", "")
        End If

        blnORDR_GROUP_NO_sql_NOT = False
        blnMANUAL_ONLY = False
        blnREL_PAST_CANCEL = True
        blnRELEASE_FUT = True
        numCANCEL_FUTURE_DAYS = Val(ROWs("SOTPARM1").Item("SO_PARM_CANCEL_FUTURE_DAYS") & "")
        manual_release = True
        MENU_ITEM_OBJECT = "SOROREL1"
        XNO = ASCMAIN1.Next_Control_No(MENU_ITEM_OBJECT & ".XNO")

        Build_Workfile2()

    End Sub

    Sub Update_Release()

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            PICK_BATCH_NO = ASCMAIN1.Next_Control_No("PICK_BATCH_NO")
        Else
            PICK_BATCH_NO = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")
        End If

        Dim use_BAs As Boolean = True
        '  use_BAs = False

        If SHIP_BOL_NO_seq > 0 Then

            If use_BAs Then

                'Dim RCSOTORDR1_ALL As Int32 = dst.Tables("SOTORDR1").Select("", "").Length
                'Dim RCSOTORDR1_MOD As Int32 = dst.Tables("SOTORDR1").Select("", "", DataViewRowState.ModifiedCurrent).Length

                'Dim RCSOTORDR2_ALL As Int32 = dst.Tables("SOTORDR2").Select("", "").Length
                'Dim RCSOTORDR2_MOD As Int32 = dst.Tables("SOTORDR2").Select("", "", DataViewRowState.ModifiedCurrent).Length

                'Dim RCSOTPICK1_ALL As Int32 = dst.Tables("SOTPICK1").Select("", "").Length
                'Dim RCSOTPICK1_MOD As Int32 = dst.Tables("SOTPICK1").Select("", "", DataViewRowState.ModifiedCurrent).Length
                'Dim RCSOTPICK1_ADD As Int32 = dst.Tables("SOTPICK1").Select("", "", DataViewRowState.Added).Length

                'Dim RCSOTPICK2_ALL As Int32 = dst.Tables("SOTPICK2").Select("", "").Length
                'Dim RCSOTPICK2_MOD As Int32 = dst.Tables("SOTPICK2").Select("", "", DataViewRowState.ModifiedCurrent).Length
                'Dim RCSOTPICK2_ADD As Int32 = dst.Tables("SOTPICK2").Select("", "", DataViewRowState.Added).Length

                'Dim RCSOTCART1_ALL As Int32 = dst.Tables("SOTCART1").Select("", "").Length
                'Dim RCSOTCART1_MOD As Int32 = dst.Tables("SOTCART1").Select("", "", DataViewRowState.ModifiedCurrent).Length
                'Dim RCSOTCART1_ADD As Int32 = dst.Tables("SOTCART1").Select("", "", DataViewRowState.Added).Length

                'Dim RCSOTCART2_ALL As Int32 = dst.Tables("SOTCART2").Select("", "").Length
                'Dim RCSOTCART2_MOD As Int32 = dst.Tables("SOTCART2").Select("", "", DataViewRowState.ModifiedCurrent).Length
                'Dim RCSOTCART2_ADD As Int32 = dst.Tables("SOTCART2").Select("", "", DataViewRowState.Added).Length

                'Dim RCSOTSHIP1_ALL As Int32 = dst.Tables("SOTSHIP1").Select("", "").Length
                'Dim RCSOTSHIP1_MOD As Int32 = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.ModifiedCurrent).Length
                'Dim RCSOTSHIP1_ADD As Int32 = dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.Added).Length

                'Debug.Print("1:" & Now)
                'Dim dt1 As DateTime = Now

                For Each TABLE_NAME As String In New String() {"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2"}
                    Create_BAs(TABLE_NAME, True)
                    Update_BAs(TABLE_NAME, True)
                Next

                'Debug.Print("2:" & Now)
                'Dim dt2 As DateTime = Now
                'Debug.Print(dt2.Subtract(dt1).TotalSeconds)

                ' this next section will delete the rows just written and then load them up the old way
                ' recent test - new way = 1 second for all 5 tables with 40K in SOTPICK2, and took 90 secs for all 5 tables the old way

                'If 1 <> 1 Then
                '    Dim r As Int32 = 0
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTPICK1) : Debug.Print(SOTPICK1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTPICK2) : Debug.Print(SOTPICK2 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTSHIP1) : Debug.Print(SOTSHIP1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTCART1) : Debug.Print(SOTCART1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTCART2) : Debug.Print(SOTCART2 & ":" & r)

                '    Debug.Print("3:" & Now)
                '    Dim dt3 As DateTime = Now
                '    For Each TABLE_NAME As String In New String() {"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2"}
                '        Update_Record_TDA(TABLE_NAME)
                '    Next
                '    Debug.Print("4:" & Now)
                '    Dim dt4 As DateTime = Now
                '    Debug.Print(dt4.Subtract(dt3).TotalSeconds)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTPICK1) : Debug.Print(SOTPICK1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTPICK2) : Debug.Print(SOTPICK2 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTSHIP1) : Debug.Print(SOTSHIP1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTCART1) : Debug.Print(SOTCART1 & ":" & r)
                '    r = ASCDATA1.ExecuteSQL("Delete from " & SOTCART2) : Debug.Print(SOTCART2 & ":" & r)
                '    Stop
                'End If

                ' still doing SOTORDR1/2 the old way for now because i don't have the onions
                ' this should be the next thing we look into for release optimiation for Walmart - because I really think it is ok to use BAs for SOTORDR1/2
                ' we just have to delete all the rows in the tables before inserting them - probably with a truncate before the begin trans
                ' need to look at how many rows are unmodified vs how many are changed.
                For Each TABLE_NAME As String In New String() {"SOTORDR1", "SOTORDR2"}
                    Update_Record_TDA(TABLE_NAME)
                Next

            Else

                For Each TABLE_NAME As String In New String() {"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2", "SOTORDR1", "SOTORDR2"}
                    Update_Record_TDA(TABLE_NAME)
                Next

            End If

            ASCMAIN1.Progress("Updating Pick Ticket && Shipment Tables", "")

            ASCDATA1.ExecuteSQL("Update " & SOTPICK1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")
            ASCDATA1.ExecuteSQL("Update " & SOTSHIP1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")

            For i As Int64 = 1 To SHIP_BOL_NO_seq
                Dim SHIP_BOL_NO As String = ""
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    SHIP_BOL_NO = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
                Else
                    SHIP_BOL_NO = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
                End If

                For Each TABLE_NAME As String In New String() {SOTSHIP1, SOTPICK1}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                        & " where SHIP_BOL_NO = 'TEMP" & Format$(i, "000000") & "'")
                Next
            Next i

            For i As Int64 = 1 To CART_NO_seq
                Dim CART_NO_ctl As String = ""
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    CART_NO_ctl = ASCMAIN1.Next_Control_No("CART_NO")
                Else
                    CART_NO_ctl = ASCMAIN1.Next_Control_No("SOTCART1.CART_NO")
                End If
                Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, CART_NO_ctl, "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                For Each TABLE_NAME As String In New String() {SOTCART1, SOTCART2}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set CART_NO = '" & CART_NO & "'" _
                        & " where CART_NO = 'TEMP" & Format$(i, "000000") & "'")
                Next
            Next i

            For i As Int64 = 1 To PICK_NO_seq
                Dim PICK_NO As String = ""
                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                    PICK_NO = ASCMAIN1.Next_Control_No("PICK_NO")
                Else
                    PICK_NO = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
                End If
                For Each TABLE_NAME As String In New String() {SOTPICK1, SOTPICK2, SOTCART1}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set PICK_NO = '" & PICK_NO & "'" _
                        & " where PICK_NO = 'TEMP" & Format(i, "000000") & "'")
                Next
            Next i


            If ASCMAIN1.CLIENT = "VAN" Then

                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP1W & " " & sqlSOTSHIP1W)
                ASCDATA1.ExecuteSQL("Insert into " & SOTPICK1W & " " & sqlSOTPICK1W)
                ASCDATA1.ExecuteSQL("Insert into " & SOTPICK2W & " " & sqlSOTPICK2W)


                ' use min SHIP_BOL_NO as the SHIP_BOL_NO_CONS for each DC

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is" & vbCrLf _
                    & "Select EDI_CONS_NO, CUST_DC_NO, COUNT (*) PTS" & vbCrLf _
                    & ", MIN (SHIP_BOL_NO) SHIP_BOL_NO_MIN, MAX (SHIP_BOL_NO) SHIP_BOL_NO_MAX" & vbCrLf _
                    & ", MIN (ORDR_DEPT) ORDR_DEPT_MIN, MAX (ORDR_DEPT) ORDR_DEPT_MAX" & vbCrLf _
                    & " from " & SOTSHIP1W & vbCrLf _
                    & " group by EDI_CONS_NO, CUST_DC_NO;" & vbCrLf _
                    & "Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update " & SOTSHIP1W & " Set SHIP_BOL_NO_CONS = R1.SHIP_BOL_NO_MIN" & vbCrLf _
                    & "  where EDI_CONS_NO = R1.EDI_CONS_NO" & vbCrLf _
                    & "    and CUST_DC_NO = R1.CUST_DC_NO;" & vbCrLf _
                    & "End Loop; End; End;"
                ASCDATA1.ExecuteSQL()

                ' use min PICK_NO as the PICK_NO_CONS for each store

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is" & vbCrLf _
                    & "Select EDI_CONS_NO, CUST_DC_NO, CUST_STORE_NO, COUNT (*) PTS" & vbCrLf _
                    & ", MIN (PICK_NO) PICK_NO_MIN, MAX (PICK_NO) PICK_NO_MAX" & vbCrLf _
                    & ", MIN (ORDR_DEPT) ORDR_DEPT_MIN, MAX (ORDR_DEPT) ORDR_DEPT_MAX" & vbCrLf _
                    & " from " & SOTPICK1W & vbCrLf _
                    & " group by EDI_CONS_NO, CUST_DC_NO, CUST_STORE_NO;" & vbCrLf _
                    & "Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update " & SOTPICK1W & " Set PICK_NO_CONS = R1.PICK_NO_MIN" & vbCrLf _
                    & "  where EDI_CONS_NO = R1.EDI_CONS_NO" & vbCrLf _
                    & "    and CUST_DC_NO = R1.CUST_DC_NO" & vbCrLf _
                    & "    and CUST_STORE_NO = R1.CUST_STORE_NO;" & vbCrLf _
                    & "End Loop; End; End;"
                ASCDATA1.ExecuteSQL()


                ' unconsolidate PTs where there are Duplicate Style/Colors

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is" & vbCrLf _
                    & "Select * from (" & vbCrLf _
                    & "Select EDI_CONS_NO, CUST_DC_NO, CUST_STORE_NO,  STYLE_CODE, COLOR_CODE, COUNT (*) DUPS" & vbCrLf _
                    & " from " & SOTPICK2W & vbCrLf _
                    & " group by EDI_CONS_NO, CUST_DC_NO, CUST_STORE_NO,  STYLE_CODE, COLOR_CODE" & vbCrLf _
                    & ") where DUPS > 1;" & vbCrLf _
                    & "Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update " & SOTPICK1W & " Set PICK_NO_CONS = PICK_NO" & vbCrLf _
                    & "  where EDI_CONS_NO = R1.EDI_CONS_NO" & vbCrLf _
                    & "    and CUST_DC_NO = R1.CUST_DC_NO" & vbCrLf _
                    & "    and CUST_STORE_NO = R1.CUST_STORE_NO;" & vbCrLf _
                    & "End Loop; End; End;"
                ASCDATA1.ExecuteSQL()


                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from " & SOTSHIP1W & ";" & vbCrLf _
                    & "Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update " & SOTSHIP1 & " Set SHIP_BOL_NO_CONS = R1.SHIP_BOL_NO_CONS" & vbCrLf _
                    & "  where SHIP_BOL_NO = R1.SHIP_BOL_NO;" & vbCrLf _
                    & "End Loop; End; End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select * from " & SOTPICK1W & ";" & vbCrLf _
                    & "Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update " & SOTPICK1 & " Set PICK_NO_CONS = R1.PICK_NO_CONS" & vbCrLf _
                    & "  where PICK_NO = R1.PICK_NO;" & vbCrLf _
                    & "End Loop; End; End;"
                ASCDATA1.ExecuteSQL()

                ' NEW
                ASCMAIN1.sql = $"SELECT  SOTPICK1.PICK_NO_CONS,  listagg(''''|| pick_no ||'''', ',') within group (order by pick_no) PICK_NOS 
                                FROM {SOTSHIP1} SOTSHIP1, {SOTPICK1} SOTPICK1
                                WHERE SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                AND SOTPICK1.PICK_NO_CONS IS NOT NULL
                                group by SOTPICK1.PICK_NO_CONS"
                For Each rowPICKs As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Rows()
                    Dim PICK_NOs As String = rowPICKs("PICK_NOS")
                    TAC.SOCMAIN1.Create_Cartons_For_PICK_NO_Cons(Me, PICK_NOs, SOTPICK1, SOTPICK2, SOTCART1, SOTCART2, SOTORDR2)
                Next

                For Each TABLE_NAME As String In New String() {"SOTCARM1", "SOTCARM2"}
                    Update_Record_TDA(TABLE_NAME)
                Next
                ASCDATA1.ExecuteSQL("Insert into SOTCARM1 Select * from " & SOTCARM1)
                ASCDATA1.ExecuteSQL("Insert into SOTCARM2 Select * from " & SOTCARM2)


            End If
            If ASCMAIN1.CLIENT = "VAN" Then
                ASCDATA1.ExecuteSQL($"Alter Table {SOTCART2} DROP COLUMN CONSOLIDATED")
            End If

            ASCDATA1.ExecuteSQL("Insert into SOTPICK1 Select * from " & SOTPICK1)
            ASCDATA1.ExecuteSQL("Insert into SOTPICK2 Select * from " & SOTPICK2)
                ASCDATA1.ExecuteSQL("Insert into SOTSHIP1 Select * from " & SOTSHIP1)
                ASCDATA1.ExecuteSQL("Insert into SOTCART1 Select * from " & SOTCART1)
                ASCDATA1.ExecuteSQL("Insert into SOTCART2 Select * from " & SOTCART2)

                Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
                With rowSOTPICK0
                    .Item("PICK_BATCH_NO") = PICK_BATCH_NO
                    .Item("PICK_SHPS") = SHIP_BOL_NO_seq
                    .Item("PICK_CTNS") = CART_NO_seq
                    .Item("PICK_PKTS") = PICK_NO_seq
                    .Item("PICK_BATCH_STATUS") = "O"
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("PICK_SHIP_REL_DATE") = SHIP_BY_DATE
                    If blnFORCE_PICK Then
                        .Item("PICK_FORCED") = "1"
                    End If
                End With
                dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)
                Update_Record_TDA("SOTPICK0")
            End If

            ASCMAIN1.Progress("Updating Order Tables", "")

        'ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2 Set ORDR_STATUS = " _
        '    & " (Select ORDR_STATUS from " & SOTORDR1 & " SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO)"
        ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2 Set ORDR_STATUS = " _
            & " CASE WHEN ORDR_QTY_OPEN > 0 THEN 'O' ELSE CASE WHEN ORDR_QTY_PICK > 0 THEN 'P' ELSE CASE WHEN ORDR_QTY_SHIP > 0 THEN 'F' ELSE 'C' END END END"
        ASCDATA1.ExecuteSQL()



        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select * from " & SOTORDR1 & RELEASE_SQL & ";" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR1 Set" _
            & "      ORDR_PICK_SEQ = R1.ORDR_PICK_SEQ" _
            & "    , ORDR_STATUS = R1.ORDR_STATUS" _
            & "    , ORDR_REL_HOLD_CODES = R1.ORDR_REL_HOLD_CODES" _
            & "    , ORDR_REL_BATCH_NO = R1.ORDR_REL_BATCH_NO" _
            & "    where ORDR_NO = R1.ORDR_NO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select * from " & SOTORDR2 & " WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & SOTORDR1 & RELEASE_SQL & ");" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2 Set" _
            & "      ORDR_STATUS = R1.ORDR_STATUS" _
            & "    , ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" _
            & "    , ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" _
            & "    , ORDR_QTY_PICK = R1.ORDR_QTY_PICK" _
            & "    , ORDR_QTY_CANC = R1.ORDR_QTY_CANC" _
            & "    , ORDR_RELEASE = R1.ORDR_RELEASE" _
            & "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" _
            & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' WHAT ABOUT UPDATING SOTRDRVX LIKE WE UPDATE SOTORDRX?

        ASCMAIN1.Progress("Updating Style Status Tables", "")

        ' note - we are not doing ICTSTAT2 ALLO - and maybe we should
        ' This statement is also in SO.SOFINVHO with minor changes.
        ' If you make a change here, please modify the other code if necessary.

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " & vbCrLf _
            & " Select SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" & vbCrLf _
            & " from SOTORDR2,SOTORDR1," & SOTPICK2 & " SOTPICK2," & SOTPICK1 & " SOTPICK1 " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) + NVL(R1.PICK_QTY,0), " & vbCrLf _
            & "                        WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - NVL(R1.PICK_QTY,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
            & "   where STYLE_CODE = R1.STYLE_CODE" & vbCrLf _
            & "     and COLOR_CODE = R1.COLOR_CODE" & vbCrLf _
            & "     and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Updating Order Group Summary", "")

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is" _
            & "  Select Distinct ORDR_GROUP_NO from SOTORDR1" _
            & "   where ORDR_REL_BATCH_NO = '" & XNO & "';" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "    SOPORDR0_G(R1.ORDR_GROUP_NO);" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from SOTCONF2" _
            & " where ORDR_NO in " _
            & " (Select ORDR_NO from SOTORDR1 " _
            & " where ORDR_REL_BATCH_NO = '" & XNO & "')"
        ASCDATA1.ExecuteSQL()

        ' USED TO UPDATE PICK_BATCH_NO WITH XNO IN VB6
        ASCMAIN1.sql = "Update SOTORDR7 Set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'" _
            & " where ORDR_GROUP_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from SOTORDR1" _
            & "   where ORDR_REL_BATCH_NO = '" & XNO & "'" _
            & "     and ORDR_REL_HOLD_CODES is Null)"
        ASCDATA1.ExecuteSQL()


        If ASCMAIN1.CLIENT = "NYA" Or ASCMAIN1.CLIENT = "RGI" Then
            ASCMAIN1.Progress("Updating 855's", "")
            If ASCMAIN1.CLIENT = "RGI" And WHSE_CODE = "CG" Then
                'RGI - CastleGate no 855's
            Else
                ASCMAIN1.sql = "Select Distinct SOTSHIP1.ORDR_GROUP_NO " _
                    & " from " & SOTSHIP1 & " SOTSHIP1, SOTORDR0 " _
                    & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                    & "   and SOTORDR0.ORDR_SOURCE = 'E'" _
                    & "   and SOTORDR0.CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '855')"
                If ASCMAIN1.CLIENT = "RGI" Then
                    'XFR orders don't get 855 - it may change for different customers
                    ASCMAIN1.sql = ASCMAIN1.sql & " and SOTORDR0.ORDR_TYPE_CODE <> 'XFR'"
                End If
                For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                    Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & String.Empty
                    TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO)
                Next
            End If
        End If

        ' Generate EDI 753s
        If ASCMAIN1.CLIENT = "NYA" Then
            ASCMAIN1.Progress("Generate 753s", "")

            ASCMAIN1.sql = "Select Distinct SOTSHIP1.SHIP_BOL_NO " _
                & " from " & SOTSHIP1 & " SOTSHIP1, SOTORDR0 " _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
                & " and SOTORDR0.CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '753')"
            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                Dim clsEDT75301 As New TAC.EDC753O1(clsASCBASE1)
                clsEDT75301.Generate_753(SHIP_BOL_NO)
            Next
        End If

    End Sub

    Private Sub chkAllocateNoRelease_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAllocateNoRelease.CheckedChanged
        Setup_Allocate_No_Release()
    End Sub

    Private Sub dteSHIP_DATE_AfterPerformAction(sender As Object, e As Infragistics.Win.UltraWinSchedule.AfterMonthViewMultiPerformActionEventArgs) Handles dteSHIP_DATE.AfterPerformAction
        Set_Date()
    End Sub

    Sub Set_Date()
        lblSHIP_DATE.Text = Format(dteSHIP_DATE.CalendarInfo.SelectedDateRanges(0).StartDate, "MM/dd/yy")
        'lblAllUnReleasable.Text = "All Un-Releasable Orders will be Shown (with Ship Date On or Before " & lblSHIP_DATE.Text & ")"
        lblAllUnReleasable.Text = String.Format("All Un-Releasable Orders with a Ship Date of On or Before {0} will be Shown on Exception Reports", lblSHIP_DATE.Text)
    End Sub

    Private Sub dteSHIP_DATE_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dteSHIP_DATE.MouseUp
        Set_Date()
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "ORDR_GROUP_NO"
                sqlw = " SOTORDR0.ORDR_CNT_OPEN > 0"
        End Select
        Return sqlw
    End Function

    Private Sub optWHSE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optWHSE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_WHSE()
    End Sub

    Sub Set_WHSE()
        If optWHSE.Value = "A" Then
            Absx1.txtFor("WHSE_CODE").Text = ""
            Absx1.txtFor("WHSE_CODE").Enabled = False
        Else
            Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
            Absx1.txtFor("WHSE_CODE").Enabled = True
        End If
    End Sub

    Private Sub chkReleasePastCancel_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkReleasePastCancel.CheckedChanged
        ' Setup_Future_Days()
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                grpCANCEL_LETTERS.Visible = chkAllocateNoRelease.Checked _
                    AndAlso (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") _
                    AndAlso Absx1.txtFor("WHSE_CODE").Text = "MS"

        End Select
    End Sub

    'Sub Setup_Future_Days()
    '    numFutureDays.Visible = chkReleasePastCancel.Checked
    '    lblFutureDays.Visible = chkReleasePastCancel.Checked
    'End Sub

    Private Sub chkINTL_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkINTL.CheckedChanged
        If chkINTL.Checked Then
            chkINTL_ONLY.Checked = False
            chkINTL_ONLY.Enabled = False
        Else
            chkINTL_ONLY.Enabled = True
            chkINTL_ONLY.Checked = True
        End If
    End Sub

    Private Sub chkECommerce_CheckedChanged(sender As Object, e As EventArgs) Handles chkECommerce.CheckedChanged
        Absx1.txtFor("ECOM_CODE").Visible = chkECommerce.Checked
    End Sub

    Private Sub chkManualOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkManualOnly.CheckedChanged
        Setup_Allocate_No_Release()
    End Sub

    Sub Create_Allocations_Exceptions()

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1 where CUST_ALLO_EXCL = '1'"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1_X", 1))
        With dst.Tables("ARTCUST1_X")
            .Columns.Add("SEL")
            .Columns("SEL").DefaultValue = "0"
        End With
        grdARTCUST1_X.DataSource = dst.Tables("ARTCUST1_X")

        ASCMAIN1.sql = "Select SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & ", SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTORDR0.ORDR_NO_MIN, SOTORDR0.ORDR_NO_MAX, SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_AMT" & vbCrLf _
            & " from SOTORDR0,SOTORDRG" & vbCrLf _
            & " where SOTORDRG.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTORDR0.ORDR_CNT_OPEN > 0" & vbCrLf _
            & "   and SOTORDRG.ORDR_ALLO_EXCL = '1'"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR0_X", 1))
        With dst.Tables("SOTORDR0_X")
            .Columns.Add("SEL")
            .Columns("SEL").DefaultValue = "0"
        End With
        grdSOTORDR0_X.DataSource = dst.Tables("SOTORDR0_X")
    End Sub

    Private Sub btnConsolidateWalmart_Click(sender As Object, e As EventArgs) Handles btnConsolidateWalmart.Click

        If tblEDT850TM.Rows.Count = 0 Then
            MsgBox("No Walmart Orders appear to be Consolidatable", MsgBoxStyle.OkOnly, "Nothing to Consolidate")
            Exit Sub
        End If

        If tblEDT850TM.Select("SEL = '1'").Length = 0 Then
            MsgBox("No Orders were Selected", MsgBoxStyle.OkOnly, "Nothing to Consolidate")
            Exit Sub
        ElseIf tblEDT850TM.Select("SEL = '1'").Length < 2 Then
            MsgBox("You must select more than 1 order in order to Consolidate", MsgBoxStyle.OkOnly, "Nothing to Consolidate")
            Exit Sub
        End If
         
        Dim r As Integer = 0
        Dim ORDR_DATE As Date
        Dim ORDR_SHIP_DATE_MIN As Date
        Dim ORDR_CANCEL_DATE_MIN As Date
        Dim ORDR_SHIP_DATE_MAX As Date
        Dim ORDR_CANCEL_DATE_MAX As Date
        Dim EDI_PROMOTION As String = ""

        Dim emsg As String = ""

        For Each row As DataRow In tblEDT850TM.Select("SEL = '1'")
            Dim msg As String = ""

            If r = 0 Then
                ORDR_DATE = row.Item("ORDR_DATE")
                ORDR_SHIP_DATE_MIN = row.Item("ORDR_SHIP_DATE_MIN")
                ORDR_CANCEL_DATE_MIN = row.Item("ORDR_CANCEL_DATE_MIN")
                ORDR_SHIP_DATE_MAX = row.Item("ORDR_SHIP_DATE_MAX")
                ORDR_CANCEL_DATE_MAX = row.Item("ORDR_CANCEL_DATE_MAX")
                EDI_PROMOTION = row.Item("EDI_PROMOTION") & ""
            Else
                If Format(ORDR_DATE, "yyyyMMdd") <> Format(row.Item("ORDR_DATE"), "yyyyMMdd") Then
                    msg = "Order Dates do not Match between Dept Orders"
                    If Not emsg.Contains(msg) Then emsg &= vbCr & msg
                End If
                If EDI_PROMOTION <> row.Item("EDI_PROMOTION") & "" Then
                    msg = "Order Type/Events do not Match between Dept Orders"
                    If Not emsg.Contains(msg) Then emsg &= vbCr & msg
                End If
                If Format(ORDR_SHIP_DATE_MIN, "yyyyMMdd") <> Format(row.Item("ORDR_SHIP_DATE_MIN"), "yyyyMMdd") Then
                    msg = "Ship Dates do not Match between Dept Orders"
                    If Not emsg.Contains(msg) Then emsg &= vbCr & msg
                End If
            End If

            If Format(row.Item("ORDR_SHIP_DATE_MIN"), "yyyyMMdd") <> Format(row.Item("ORDR_SHIP_DATE_MAX"), "yyyyMMdd") Then
                msg = "Ship Dates do not Match within a single Dept Order"
                If Not emsg.Contains(msg) Then emsg &= vbCr & msg
            End If
            r += 1
        Next

        If emsg <> "" Then
            If MsgBox(emsg & vbCrLf & vbCrLf & "Click OK to Proceed, or Cancel to avoid PO Consolidation", MsgBoxStyle.OkCancel, "Warning - Order Info does Not Match") = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If
 
        For Each row As DataRow In tblEDT850TM.Select("SEL = '1'")
            ORDR_DATE = row.Item("ORDR_DATE")
            EDI_PROMOTION = row.Item("EDI_PROMOTION") & ""

            ASCMAIN1.sql = "" _
                 & "Select Distinct SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO" & vbCrLf _
                & " from SOTORDR0,EDT850T1,SOTORDR1,SOTPICK1" & vbCrLf _
                & "   where SOTORDR0.CUST_CODE = 'WALMART' and SOTORDR0.ORDR_CNT_OPEN <> 0" & vbCrLf _
                & "     and SOTORDR0.ORDR_DATE = :PARM1 and EDT850T1.EDI_PROMOTION = :PARM2" & vbCrLf _
                & "     and EDT850T1.EDI_CONS_NO IS NULL" & vbCrLf _
                & "     and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                & "     and SOTORDR1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "     and SOTPICK1.PICK_STATUS in ('P','F')"
            Dim ROWS() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "DV", New Object() {ORDR_DATE, EDI_PROMOTION}).Select("")

            If ROWS.Count <> 0 Then
                MsgBox("Some of these Walmart Sales Orders (ex: " & ROWS(0).Item("ORDR_NO") & ") were already Released" _
                       & vbCrLf & " - and are no longer eligible for Consolidation", MsgBoxStyle.OkOnly, "Cancelling Request to Consolidate")
                Exit Sub
            End If
        Next


        BeginTrans()

        Dim EDI_CONS_NO_ASSIGNED As String = ASCMAIN1.Next_Control_No("EDT850T1.EDI_CONS_NO")

        Dim order_count As Integer = 0

        For Each row As DataRow In tblEDT850TM.Select("SEL = '1'")
            ORDR_DATE = row.Item("ORDR_DATE")
            EDI_PROMOTION = row.Item("EDI_PROMOTION") & ""

            ASCMAIN1.sql = "" _
                & "Update EDT850T1 Set EDI_CONS_NO = :PARM1" & vbCrLf _
                & " where EDI_DOC_SEQ_NO in (" & vbCrLf _
                & "  Select Distinct EDT850T1.EDI_DOC_SEQ_NO from SOTORDR0,EDT850T1" & vbCrLf _
                & "   where SOTORDR0.CUST_CODE = 'WALMART' and SOTORDR0.ORDR_CNT_OPEN <> 0" & vbCrLf _
                & "     and SOTORDR0.ORDR_DATE = :PARM2 and EDT850T1.EDI_PROMOTION = :PARM3" & vbCrLf _
                & "     and EDT850T1.EDI_CONS_NO IS NULL" & vbCrLf _
                & "     and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO" & vbCrLf _
                & " )"

            Dim order_count_sel As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDV", New Object() {EDI_CONS_NO_ASSIGNED, ORDR_DATE, EDI_PROMOTION})
            order_count += order_count_sel
        Next

        CommitTrans("Total Number of Orders Consolidated = " & CStr(order_count))

        btnConsolidateWalmart.Enabled = False
        grdEDT850TM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        ' note - to de-consolidate prior to release, all you need to do is to update EDT850T1 and set EDI_CONS_NO = NULL where EDI_CONS_NO = {'00000000099'}



        ' this is what i did to consolidate 2 random orders using PLSQL
        ' it turned out (for these 45 + 45 = 90 POs) that EDI_PROMOTION was POSREPWK43

        ' Begin Declare EDI_CONS_NO_ASSIGNED VARCHAR2(10);
        '                  Begin
        '                   EDI_CONS_NO_ASSIGNED := TAPCTLN1('EDT850T1.EDI_CONS_NO',1);
        '                   UPDATE EDT850T1 SET EDI_CONS_NO = EDI_CONS_NO_ASSIGNED
        '                    where EDI_DOC_SEQ_NO IN (
        '                     Select Distinct EDT850T1.EDI_DOC_SEQ_NO from SOTORDR0,EDT850T1
        '                      where CUST_CODE = 'WALMART' AND ORDR_CNT_OPEN <> 0 AND EDI_PROMOTION = 'POSREPWK43'
        '                        and EDI_CONS_NO IS NULL
        '                        and EDT850T1.EDI_DOC_SEQ_NO = SOTORDR0.EDI_DOC_SEQ_NO
        '                   );
        '                  End;
        '                 End;


    End Sub

    Private Sub grdEDT850TM_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdEDT850TM.InitializeRow
        Dim ORDR_SHIP_DATE_MIN As Date = e.Row.Cells("ORDR_SHIP_DATE_MIN").Value
        Dim ORDR_SHIP_DATE_MAX As Date = e.Row.Cells("ORDR_SHIP_DATE_MAX").Value
        Dim ORDR_CANCEL_DATE_MIN As Date = e.Row.Cells("ORDR_CANCEL_DATE_MIN").Value
        Dim ORDR_CANCEL_DATE_MAX As Date = e.Row.Cells("ORDR_CANCEL_DATE_MAX").Value

        If ORDR_SHIP_DATE_MAX <> ORDR_CANCEL_DATE_MIN Then
            e.Row.Cells("ORDR_SHIP_DATE_MAX").Appearance.ForeColor = Drawing.Color.Red
        End If

        If ORDR_CANCEL_DATE_MAX <> ORDR_CANCEL_DATE_MIN Then
            e.Row.Cells("ORDR_CANCEL_DATE_MAX").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub
End Class