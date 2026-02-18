Public Class SORINVP1

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date

    Dim SQLs As New Dictionary(Of String, String)

    Dim INV_TYPEs As String

    Dim SOTINVP1 As String
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String

    Dim sqlSOTINVP1 As String
    Dim sqlSOTINVH1 As String
    Dim sqlSOTINVH2 As String

    Dim consolidated_invoice As Boolean = False

    Private BATCH_VAN_FOLDER As String = ""

    Dim subUPCSupport As Boolean = (ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI")

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Left = grpDATE_RANGE.Left
        grpPERIOD_RANGE.Top = grpDATE_RANGE.Top

        grpNotPrintedYet.Left = grpDATE_RANGE.Left
        grpNotPrintedYet.Top = grpDATE_RANGE.Top


        If ASCMAIN1.CLIENT = "RGI" Then
            Absx1.chkFor("CHKEXPORT_INFO").Visible = True
        End If

        If ASCMAIN1.CLIENT = "VAN" Then
            grpBATCH_VAN.Visible = True
            chkBATCH_VAN.Checked = False
        Else
            grpBATCH_VAN.Visible = False
            chkBATCH_VAN.Checked = False
        End If
    End Sub


    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices/Credits Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices/Credits Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            If RYP0 = RYP1 Then
                SUBT = "Invoices/Credits Posted in " & RYPLEGEND0
            Else
                SUBT = "Invoices/Credits Posted between " & RYPLEGEND0 & " and " & RYPLEGEND1
            End If
            sqlw = " and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
            RWU = "N"
        End If

        'INV_TYPEs _
        '= IIf(chkTypeS.Checked, ",'I'", "") _
        '& IIf(chkTypeR.Checked, ",'C'", "") 

        If optRANGE.Value = "U" Then
            sqlw &= "   and SOTINVH1.INV_PRINTED is Null" & vbCrLf
            If cmbDIV.Value & "" <> "" Then
                sqlw &= "   and SOTINVH1.SALES_DIVISION_CODE = '" & cmbDIV.Value & "'" & vbCrLf
            End If
            If chkMYINVOICESONLY.Visible And chkMYINVOICESONLY.Checked Then
                sqlw &= "   and SOTINVH1.INIT_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf
            End If
        End If

        sqlw &= SQL_in("SHIP_BOL_NO", "SOTINVH1.SHIP_BOL_NO") & vbCrLf
        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE") & vbCrLf
        sqlw &= SQL_in("SREP_CODE", "SOTINVH1.SREP_CODE") & vbCrLf

        If Absx1.chkFor("CHKCONS_INV").Checked Then
            sqlw &= SQL_in("INV_NO", "SOTINVH1.INV_NO_CONS")
            ' unconsolidated invoices wind up printing * Consolidated Invoice * with these lines enabled - for now, you must uncheck the Consolidated option to print regular invoices.
            'sqlw &= " and (" & Mid(SQL_in("INV_NO", "SOTINVH1.INV_NO"), 5) _
            '    & " or " & Mid(SQL_in("INV_NO", "SOTINVH1.INV_NO_CONS"), 5) & ")" & vbCrLf
        Else
            Dim SQLWX As String = SQL_in("INV_NO", "SOTINVH1.INV_NO")
            If SQLWX <> "" Then sqlw &= " and (SOTINVH1.INV_TYPE = 'I'" & SQLWX & ")" & vbCrLf
        End If

        If Absx1.chkFor("CHKEDI").Checked Then
            sqlw &= " and NVL(SOTORDR1.ORDR_SOURCE,'K') <> 'E'" & vbCrLf
        End If

        'If Absx1.chkFor("CHKBTB").Checked Then
        '    sqlw &= " and SOTINVH1.ORDR_TYPE_CODE = 'BTB'" & vbCrLf
        'Else
        '    sqlw &= " and SOTINVH1.ORDR_TYPE_CODE <> 'BTB'" & vbCrLf
        'End If


        sqlw &= " and SOTINVH1.INV_NO_REV_BY is null" & vbCrLf

        If Absx1.optFor("RANGE").Value = "U" Then
            RWU = "R"
        Else
            RWU = "N"
        End If

        consolidated_invoice = Absx1.chkFor("CHKCONS_INV").Checked

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()
        'Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        'If RPT = "" Then RPT = "SORINVP1"

        Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1"), New String() {"CUST_CODE"})
        If tbl.Rows.Count = 1 Then
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(tbl.Rows(0).Item("CUST_CODE"))
            If rowARTCUST1 IsNot Nothing AndAlso rowARTCUST1.Item("CUST_INV_REPORT") & String.Empty <> String.Empty Then
                RPT = rowARTCUST1.Item("CUST_INV_REPORT")
            End If
        End If

        If chkBATCH_VAN.Checked Then
            If BATCH_VAN_FOLDER.Length > 0 Then
                For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
                    Dim INV_NO As String = rowSOTINVH1.Item("INV_NO").ToString & String.Empty
                    Dim ORDR_CUST_PO As String = rowSOTINVH1.Item("ORDR_CUST_PO").ToString & String.Empty
                    Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE").ToString & String.Empty
                    Dim RPT_NAME As String = String.Format("INV{0}-PO{1}-{2}", INV_NO, ORDR_CUST_PO, CUST_CODE)
                    Dim FILENAME_temp As String = ASCMAIN1.Folders("Temp") & RPT_NAME & ".pdf"
                    Dim INV_SELECT As String = "{SOTINVH1.INV_NO} = '" & INV_NO & "'"
                    Dim RPT_NAME_DEST As String = String.Format("{0}\INV{1}-PO{2}-{3}.pdf", BATCH_VAN_FOLDER, INV_NO, ORDR_CUST_PO, CUST_CODE)
                    CR_params.Add("SUBT", "")
                    CR_params.Add("CONS_INV", IIf(Absx1.chkFor("CHKCONS_INV").Checked, "1", "0"))
                    CR_params.Add("EXPORT_INFO", IIf(Absx1.chkFor("CHKEXPORT_INFO").Checked, "1", "0"))

                    Generate_Report("SORINVP1", "Invoices to be Saved", , INV_SELECT, "PDF", RPT_NAME, False)

                    If IO.File.Exists(RPT_NAME_DEST) Then
                        IO.File.Delete(RPT_NAME_DEST)
                    End If
                    My.Computer.FileSystem.CopyFile(FILENAME_temp, RPT_NAME_DEST, True)

                Next
            Else
                MsgBox("Invalid Folder Selected", vbOKOnly, "Save PDFs")
                Exit Sub
            End If
        Else
            Select Case RPT

                Case "SORINVHZ"
                    CR_params.Add("SUBT", "")
                    'CR_params.Add("CONS_INV", IIf(Absx1.chkFor("CHKCONS_INV").Checked, "1", "0"))
                    'CR_params.Add("EXPORT_INFO", "0")
                    Generate_Report(RPT, , SUBT)
                Case Else
                    CR_params.Add("SUBT", "")
                    CR_params.Add("CONS_INV", IIf(Absx1.chkFor("CHKCONS_INV").Checked, "1", "0"))
                    'CR_params.Add("EXPORT_INFO", "0")
                    CR_params.Add("EXPORT_INFO", IIf(Absx1.chkFor("CHKEXPORT_INFO").Checked, "1", "0"))
                    Generate_Report(RPT, , SUBT)
            End Select
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            Else

            End If
        End If

    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        grpNotPrintedYet.Visible = (optRANGE.Value = "U")
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Enabled = (optRANGE.Value = "P")
        grpNotPrintedYet.Enabled = (optRANGE.Value = "U")

        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

        ' *******************************************************
        ' NOTE
        '   RGIs Credits do not have Order Numbers
        '   therefore the code create a temp order number and SOTORDR5_ST record
        '   so credust show the Ship To Address.
        '   Do Not Update SOTINVH1 from the dataset
        '   edz - 06/04/2015
        ' *******************************************************

        Dim sql As String = "Update SOTINVH1 " _
        & " Set INV_PRINTED = SYSDATE" _
        & " where (INV_TYPE, INV_NO) in (Select INV_TYPE, INV_NO from " & SOTINVP1 & " )"
        ASCDATA1.ExecuteSQL(sql)

        ' WHAT ABOUT DR & CRS
        sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
        & " Select 'SOTORDR1', ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'INVPRT','Invoice Printed', INV_NO" _
        & " from " & SOTINVP1
        ASCDATA1.ExecuteSQL(sql)
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        sqlSOTINVH1 = "Select SOTINVH1.* from SOTINVH1,SOTORDR1 where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO"
        ASCMAIN1.sql = sqlSOTINVH1 & sqlw
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE, INV_NO)")

        sqlSOTINVH2 = "Select SOTINVH2.*,SOTINVH1.PICK_NO,SOTINVH1.ORDR_NO,SOTINVH1.SHIP_BOL_NO,SOTINVH1.SREP_CODE" _
            & " from SOTINVH2, " & SOTINVH1 & " SOTINVH1" _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
        sqlSOTINVH2 = "Select SOTINVH1.*,SOTPICK2.ORDR_LNO from SOTPICK2,(" & sqlSOTINVH2 & ") SOTINVH1" _
            & " where SOTPICK2.PICK_NO (+) = SOTINVH1.PICK_NO and SOTPICK2.PICK_LNO (+) = SOTINVH1.INV_LNO"
        ASCMAIN1.sql = sqlSOTINVH2
        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE, INV_NO, INV_LNO)")

        sqlSOTINVP1 = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_NO, SOTINVH1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.CUST_STORE_NO, SOTINVH1.SHIP_BOL_NO, SOTINVH1.PICK_NO" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1,SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO"
        ASCMAIN1.sql = sqlSOTINVP1 & sqlw
        SOTINVP1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVP1 & " Add Primary Key (INV_TYPE, INV_NO)")

        SQLs.Clear()

        With dst

            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1"
            SQLs.Add("SOTINVH1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0, False, "", 2)
            With .Tables("SOTINVH1").Columns
                .Add("TOTAL_UNITS", GetType(System.Int64))
                .Add("AR_PARM_KEY")
                .Add("BT")
                .Add("ST")
                .Add("CART_TRACKING_NO")
                .Add("TOTAL_CUBE", GetType(System.Decimal))
                .Add("MISC_CHARGES")
                .Add("TARIFFS")
            End With

            ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, MIN (SOTCART1.CART_TRACKING_NO) CART_TRACKING_NO" & vbCrLf _
              & " from " & SOTINVH1 & " SOTINVH1, SOTCART1" & vbCrLf _
              & " where SOTCART1.PICK_NO = SOTINVH1.PICK_NO" & vbCrLf _
              & " group by SOTINVH1.INV_TYPE, SOTINVH1.INV_NO"
            Create_TDA(.Tables.Add, "SOTINVH1_CART_TRACKING_NO", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTINVH2.*, ICTSTYC1.UPC_CODE, '0' TARIFFS" & vbCrLf _
                & " from " & SOTINVH2 & " SOTINVH2, ICTSTYC1" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE (+) = SOTINVH2.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE (+) = SOTINVH2.COLOR_CODE"
            SQLs.Add("SOTINVH2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVH2", "**", 0, False, "", 3)

            If ASCMAIN1.CLIENT = "VAN" Then
                With .Tables("SOTINVH2").Columns
                    .Add("PF_QTY", GetType(System.Int64))
                    .Add("DUTY_HTS_CODE", GetType(System.String))

                End With
                .Tables("SOTINVH2").Columns("LIC_CODE").MaxLength = 20
            End If

            'If ASCMAIN1.CLIENT = "VAN" Then
            With .Tables("SOTINVH1").Columns
                .Add("PF_SHIP_NOTES", GetType(System.String))
                .Add("PF_OVERSEAS_DOMESTIC", GetType(System.String))
                .Add("PF_VIA", GetType(System.String))
                .Add("PO_SHIPMENT_NO", GetType(System.String))
            End With
            'End If

            ' dgj
            Create_Relation("SOTINVH1", "SOTINVH2", "INV_TYPE,INV_NO")
            '.Tables("SOTINVH2").Columns.Add("PICK_NO", GetType(System.String), "PARENT(SOTINVH1_SOTINVH2).PICK_NO")
            '.Tables("SOTINVH2").Columns.Add("ORDR_NO", GetType(System.String), "PARENT(SOTINVH1_SOTINVH2).ORDR_NO")
            '.Tables("SOTINVH2").Columns.Add("ORDR_LNO", GetType(System.Int32))
            .Tables("SOTINVH1").Columns("TOTAL_UNITS").Expression = "SUM(CHILD(SOTINVH1_SOTINVH2).ORDR_QTY_SHIP)"

            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & " from SOTSHIP1" & vbCrLf _
                & " where SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTSHIP1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & " from SOTPICK1" & vbCrLf _
                & " where PICK_NO in (Select Distinct PICK_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTPICK1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "", 1)


            'If ASCMAIN1.CLIENT = "VAN" Then
            With .Tables("SOTPICK1").Columns
                .Add("PF_WEIGHT_UOM", GetType(System.String))
                .Add("PO_SHIPMENT_NO", GetType(System.String))
            End With
            ' End If


            ASCMAIN1.sql = "Select SOTPICK2.*" & vbCrLf _
                & " from SOTPICK2" & vbCrLf _
                & " where PICK_NO in (Select Distinct PICK_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTPICK2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE, SOTPICK1.INV_NO" & vbCrLf _
                & " from SOTPICK1, SOTPICK2, SOTORDR2, ICTSTYL1" & vbCrLf _
                & " where SOTPICK1.PICK_NO = SOTPICK2.PICK_NO AND SOTPICK2.PICK_NO in (Select Distinct PICK_NO from " & SOTINVP1 & ")" & vbCrLf _
                & " and SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO and SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO" & vbCrLf _
                & " and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE"
            SQLs.Add("SOTCUBE1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTCUBE1", "**", 0, False, "", 2)
            .Tables("SOTCUBE1").Columns.Add("TOTAL_CARTONS", GetType(System.Decimal), "IIF(ISNULL(CARTON_PACK_QTY,0)=0,0,ISNULL(PICK_QTY_CONF,0) / ISNULL(CARTON_PACK_QTY,0))")
            .Tables("SOTCUBE1").Columns.Add("TOTAL_CUBE", GetType(System.Decimal), "ISNULL(TOTAL_CARTONS,0) * ISNULL(CASE_CUBE,0)")

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*, '0' INV_UOM" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR9.*" & vbCrLf _
                & " from SOTORDR9" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTORDR9", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR9", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*" & vbCrLf _
                & " from SOTORDR5" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")" & vbCrLf _
                & "   and CUST_ADDR_TYPE = 'BT'"
            SQLs.Add("SOTORDR5_BT", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR5_BT", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*" & vbCrLf _
                & " from SOTORDR5" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")" & vbCrLf _
                & "   and CUST_ADDR_TYPE = 'ST'"
            SQLs.Add("SOTORDR5_ST", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR5_ST", "**", 0, False, "", 2)


            ASCMAIN1.sql = "Select POTORDR1.*" & vbCrLf _
                & " from POTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("POTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & " from POTORDR2" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("POTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1" & vbCrLf _
                & " where STYLE_CODE in (" & vbCrLf _
                & "       (Select Distinct STYLE_CODE from SOTINVH2 where (INV_TYPE,INV_NO) in (Select Distinct INV_TYPE,INV_NO from " & SOTINVP1 & "))" & vbCrLf _
                & " union (Select Distinct STYLE_CODE from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & "))" & vbCrLf _
                & " union (Select Distinct STYLE_CODE_SUB from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & "))" & vbCrLf _
                & " union (Select Distinct STYLE_CODE from SOTINVH2 where (INV_TYPE,INV_NO) in (Select Distinct INV_TYPE,INV_NO from " & SOTINVH1 & "))" & vbCrLf _
                & " union (Select Distinct STYLE_CODE from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVH1 & "))" & vbCrLf _
                & " union (Select Distinct STYLE_CODE_SUB from SOTORDR2 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVH1 & "))" & vbCrLf _
                & ")"
            SQLs.Add("ICTSTYL1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTINVH9.*" _
                & " from SOTINVH9 SOTINVH9, " & SOTINVP1 & " SOTINVP1" _
                & " where SOTINVP1.INV_NO = SOTINVH9.INV_NO"
            SQLs.Add("SOTINVH9", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVH9", "**", 0, False, "", 3)

            ' Non Tariff Misc Charges
            ASCMAIN1.sql = $"Select SOTINVHM.*
                                from SOTINVHM, {SOTINVP1} SOTINVP1
                                where SOTINVHM.INV_NO = SOTINVP1.INV_NO 
                                AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') <> 'T'"
            If ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI" Then
                ASCMAIN1.sql &= " and INV_MISC_CHG <> 0"
            End If
            SQLs.Add("SOTINVHM", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVHM", "**", 0, False, "", 3)

            'Tariff Misc Charges
            ASCMAIN1.sql = $"Select SOTINVHM.*, NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE) COUNTRY_NAME, 
                                ROUND(SOTINVHM.INV_MISC_CHG / SOTINVH2.ORDR_QTY_SHIP, 3) INV_MISC_CHG_UNIT
                                From SOTINVHM, {SOTINVP1} SOTINVP1, TATCNTRY, SOTINVH2
                                Where SOTINVHM.INV_NO = SOTINVP1.INV_NO 
                                AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                                AND SOTINVHM.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE (+)
                                AND SOTINVHM.INV_TYPE = SOTINVH2.INV_TYPE
                                AND SOTINVHM.INV_NO = SOTINVH2.INV_NO
                                AND SOTINVHM.INV_LNO = SOTINVH2.INV_LNO"
            SQLs.Add("SOTINVHT", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVHT", "**", 0, False, "", 4)

            ASCMAIN1.sql = $"SELECT SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.COUNTRY_CODE, SOTINVHM.SURCHARGE_PERC,
                                NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE) COUNTRY_NAME,
                                SUM(SOTINVHM.INV_MISC_CHG) TOTAL_TARIFF
                                FROM SOTINVHM, TATCNTRY, {SOTINVP1} SOTINVP1, SOTINVH2
                                WHERE SOTINVHM.INV_NO = SOTINVP1.INV_NO 
                                AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                                AND SOTINVH2.INV_TYPE = SOTINVHM.INV_TYPE 
                                AND SOTINVH2.INV_NO = SOTINVHM.INV_NO
                                AND SOTINVH2.INV_LNO = SOTINVHM.INV_LNO 
                                AND SOTINVHM.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE (+)
                                GROUP BY SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.COUNTRY_CODE, SOTINVHM.SURCHARGE_PERC,
                                NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE)"
            SQLs.Add("SOTINVHT_SUM", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVHT_SUM", ASCMAIN1.sql, 0, False, "", 3)

            ASCMAIN1.sql = $"Select SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.INV_LNO, SUM(SOTINVHM.INV_MISC_CHG) INV_MISC_CHG_SUM,
                                Round(SUM(SOTINVHM.INV_MISC_CHG / SOTINVH2.ORDR_QTY_SHIP), 4) INV_MISC_CHG_UNIT_SUM
                                From SOTINVHM, {SOTINVP1} SOTINVP1, SOTINVH2
                                Where SOTINVHM.INV_NO = SOTINVP1.INV_NO 
                                AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                                AND SOTINVHM.INV_TYPE = SOTINVH2.INV_TYPE
                                AND SOTINVHM.INV_NO = SOTINVH2.INV_NO
                                AND SOTINVHM.INV_LNO = SOTINVH2.INV_LNO
                                GROUP BY SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.INV_LNO"
            SQLs.Add("SOTINVHT_TS", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVHT_TS", ASCMAIN1.sql, 0, False, "", 3)

            ASCMAIN1.sql = "Select Distinct ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & ", ARTCUST1.CUST_ADDR1, ARTCUST1.CUST_ADDR2, ARTCUST1.CUST_ADDR3" & vbCrLf _
                & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE, ARTCUST1.CUST_DBA_NAME " & vbCrLf _
                & ", ARTCUST1.CUST_XMIT_INV_VIA, ARTCUST1.CUST_INV_COMMENT " & vbCrLf _
                & ", ARTCUST1.CUST_INV_EMAIL, ARTCUST1.CUST_INV_CC" & vbCrLf _
                & ", ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_BILL_SHIP_TO, ARTCUST1.CUST_INCL_INV_SHIP" & vbCrLf _
                & ", ARTCUST1.CUST_VEND_REF, ARTCUST1.CUST_INV_REPORT" & vbCrLf _
                & " from ARTCUST1" & vbCrLf _
                & " where CUST_CODE in" & vbCrLf _
                & " (Select Distinct CUST_CODE from " & SOTINVP1 & " union Select Distinct CUST_BILL_TO_CUST from " & SOTINVP1 & ")"
            SQLs.Add("ARTCUST1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2 where (CUST_CODE,CUST_ADDR_CODE) in " & vbCrLf _
                & " (Select Distinct CUST_CODE,CUST_STORE_NO from " & SOTINVP1 & ")"
            SQLs.Add("ARTCUST2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 3)

            'ASCMAIN1.sql = "Select SOTSREP1.SREP_CODE, SOTSREP1.SREP_NAME" _
            '    & " from SOTSREP1" _
            '    & " where SOTSREP1.SREP_CODE in (Select Distinct SREP_CODE from " & SOTINVP1 & ")"
            'SQLs.Add("SOTSREP1", ASCMAIN1.sql)
            'Create_TDA(.Tables.Add, "SOTSREP1", "**", 0, False, "", 1)

            For Each TABLE_NAME As String In New String() _
            {"TATTERM1", "ICTWHSE1", "SOTSVIA1", "SOTSREP1", "SOTREAS1", "ARTREAS1", "SOTSDIV1", "TATCNTRY", "ICTPORT1"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            With .Tables.Add("SOTINVP0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("AR_PARM_DUNS_NO")
                .Columns.Add("ADDRESS_LINE")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With

            'With .Tables("SOTSDIV1")
            '    .Columns.Add("DIVISION_LOGO", GetType(System.Byte()))
            'End With

            'ASCMAIN1.sql = "Select Distinct SOTSDIV1.DIVISION_CODE, SOTINVP1.*" _
            '& " from SOTINVP1,SOTSDIV1 " _
            '& " where SOTINVP1.CUST_CODE = SOTSDIV1.CUST_CODE"
            'Create_TDA(.Tables.Add, "SOTINVP1_DIV", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select DISTINCT EDT850T1.*" & vbCrLf _
                    & " from EDT850T1" & vbCrLf _
                    & " where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM SOTORDR1 WHERE ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & "))"
            SQLs.Add("EDT850T1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "EDT850T1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select DISTINCT EDT850T2.*" & vbCrLf _
                    & " from EDT850T2" & vbCrLf _
                    & " where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM SOTORDR1 WHERE ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & "))"
            SQLs.Add("EDT850T2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "EDT850T2", "**", 0, False, "", 2)

            If subUPCSupport Then
                ASCMAIN1.sql = $"Select ICTXLSPS.* 
                    from ICTXLSPS, POTORDR2
                    where ICTXLSPS.STYLE_CODE = POTORDR2.STYLE_CODE 
                    AND ICTXLSPS.COLOR_CODE = POTORDR2.COLOR_CODE
                    AND POTORDR2.ORDR_NO in (Select Distinct ORDR_NO from {SOTINVP1})"
                Create_TDA(.Tables.Add, "ICTXLSPS", "**", 0, False, "")
            End If

        End With

        Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
        With ROWs("ARTPARM1")
            rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
            rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
                rowSOTINVP0.Item("REMIT3") = "" _
                    & "  Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
                    & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
            End If
            rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
            If 1 = 1 Then
                rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
            End If
            rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
            rowSOTINVP0.Item("ADDRESS_LINE") = "" _
                & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                  And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                      & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                      & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
        End With
        rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)

        If ASCMAIN1.CLIENT = "NYA" Then
            ' NEED TO RETHINK THIS WHEN WE GET TO NEEDING A 3RD REMIT ADDRESS - RIGHT NOW, BILL SAYS TO DO IT BY CURRENCY
            rowSOTINVP0 = dst.Tables("SOTINVP0").NewRow
            Dim rowARTRMIT1 As DataRow = LookUp("ARTRMIT1", "CAD")
            With rowARTRMIT1
                rowSOTINVP0.Item("AR_PARM_KEY") = "CAD"
                rowSOTINVP0.Item("REMIT0") = .Item("REMIT_NAME") & ""
                rowSOTINVP0.Item("REMIT1") = .Item("REMIT_ADDR1") & ""
                rowSOTINVP0.Item("REMIT2") = .Item("REMIT_CITY") & ", " _
                        & .Item("REMIT_STATE") & " " _
                        & .Item("REMIT_ZIP_CODE") & " " _
                        & .Item("REMIT_COUNTRY")
                If .Item("REMIT_PHONE") & "" <> "" And .Item("REMIT_FAX") & "" <> "" Then
                    rowSOTINVP0.Item("REMIT3") = "" _
                        & "  Tel " & ASCMAIN1.FormatTel(.Item("REMIT_PHONE")) _
                        & ", Fax " & ASCMAIN1.FormatTel(.Item("REMIT_FAX"))
                End If
                rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("REMIT_MESSAGE") & ""
                If 1 = 1 Then
                    '  rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = rowSOTINVP0.Item("REMIT_MESSAGE") '  & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
                End If
                rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("REMIT_DUNS_NO") & ""
                rowSOTINVP0.Item("ADDRESS_LINE") = "" _
                    & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                    & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                    & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                    & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                    & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                      And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                          & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                          & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
            End With
            rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
            dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)



            dst.Tables("SOTINVP0").Columns.Add("AR_PARM_FORM_COUNTRY")
            dst.Tables("SOTINVP0").Columns.Add("AR_PARM_FORM_TAX_ID")
            dst.Tables("SOTINVP0").Columns("AR_PARM_KEY").MaxLength = -1
            dst.Tables("SOTINVH1").Columns("AR_PARM_KEY").MaxLength = -1

            ASCMAIN1.sql = "Select * from SOTCOMP1"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim COMP_CODE As String = row.Item("COMP_CODE")
                Dim rowP As DataRow = dst.Tables("SOTINVP0").NewRow
                rowP.ItemArray = rowSOTINVP0.ItemArray
                rowP.Item("AR_PARM_KEY") = "Z" & COMP_CODE

                With row
                    'rowP.Item("REMIT0") = .Item("COMP_NAME") & ""
                    'rowP.Item("REMIT1") = .Item("COMP_ADDR1") & ""
                    'rowP.Item("REMIT2") = .Item("COMP_CITY") & ", " _
                    '        & .Item("COMP_STATE") & " " _
                    '        & .Item("COMP_ZIP_CODE") & " " _
                    '        & .Item("COMP_COUNTRY")
                    'If .Item("COMP_PHONE") & "" <> "" And .Item("COMP_FAX") & "" <> "" Then
                    '    rowP.Item("REMIT3") = "" _
                    '        & "  Tel " & ASCMAIN1.FormatTel(.Item("COMP_PHONE")) _
                    '        & ", Fax " & ASCMAIN1.FormatTel(.Item("COMP_FAX"))
                    'End If

                    rowP.Item("AR_PARM_DUNS_NO") = "203883207"

                    'rowP.Item("ADDRESS_LINE") = "" _
                    '                    & .Item("COMP_ADDR1") _
                    '                    & ", " & .Item("COMP_CITY") _
                    '                    & ", " & .Item("COMP_STATE") _
                    '                    & " " & .Item("COMP_ZIP_CODE") _
                    '                    & IIf(.Item("COMP_PHONE") & "" <> "" _
                    '                      And .Item("COMP_FAX") & "" <> "", "" _
                    '                          & ", Tel " & ASCMAIN1.FormatTel(.Item("COMP_PHONE") & "") _
                    '                          & ", Fax " & ASCMAIN1.FormatTel(.Item("COMP_FAX") & ""), "")

                End With

                'For Each C As String In New String() {"COMP_NAME", "COMP_ADDR1", "COMP_ADDR2", "COMP_ADDR3", _
                '                                        "COMP_CITY", "COMP_STATE", "COMP_ZIP_CODE", "COMP_COUNTRY", _
                '                                        "COMP_PHONE", "COMP_FAX", "COMP_EMAIL", "COMP_TAX_ID"}
                '    Dim CP As String = Replace(C, "COMP_", "PO_PARM_FORM_")
                '    rowP.Item(CP) = row.Item(C)
                'Next
                rowP.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & "_" & COMP_CODE & ".PNG")
                dst.Tables("SOTINVP0").Rows.Add(rowP)
            Next


        End If

        '  Fill_Records("SOTINVP1_DIV")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        Dim pro_forma As Boolean = False

        If parms.Length > 0 Then
            'Dim INV_NOs As String = parms(0)
            sqlw = parms(0)
            'not nec no more
            'If Not Trim(sqlw).ToUpper.StartsWith("AND") Then
            '    sqlw = " and " & sqlw
            'End If

            Dim INV_TYPE_requested As String = ""
            Dim pfComment As String = ""
            Dim ORDR_QTY_field As String = "DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR2.ORDR_QTY,DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR2.ORDR_QTY_PICK,SOTORDR2.ORDR_QTY_SHIP))"
            Dim sqlw2 As String = ""

            If parms.Length >= 2 Then
                pro_forma = (parms(1) = "1")
            End If
            If parms.Length >= 3 Then
                'If pro_forma And (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") Then
                '    pfComment = parms(2) & ""
                'Else
                INV_TYPE_requested = parms(2) & ""
                'End If
            End If
            If parms.Length >= 4 Then
                If parms(3) & "" <> "" Then
                    ORDR_QTY_field = parms(3) & ""
                    If ORDR_QTY_field = "ORDR_QTY_ALLO_X" Then
                        ORDR_QTY_field = "ORDR_QTY_ALLO"
                        sqlw2 = "   AND NVL(SOTORDR2.ORDR_RELEASE_AVAIL,'01-JAN-1900') = '01-JAN-1900'"
                    End If
                End If
            End If
            If parms.Length >= 5 Then
                consolidated_invoice = (parms(4) = "1")
            End If

            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVP1)
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVH1)
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVH2)

            If pro_forma And INV_TYPE_requested <> "C" Then

                '& ", " & ORDR_QTY_field & " ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_CURR,SOTORDR2.ORDR_NO,SOTORDR2.ORDR_LNO" & vbCrLf _

                If INV_TYPE_requested = "O" OrElse INV_TYPE_requested = "B" Then
                    ASCMAIN1.sql = "Insert into " & SOTINVH2 & vbCrLf _
                            & " (INV_TYPE,INV_NO,INV_LNO,STYLE_CODE,COLOR_CODE,ORDR_UNIT_PRICE,ORDR_QTY_SHIP,ORDR_UNIT_PRICE_CURR,ORDR_NO,ORDR_LNO)" & vbCrLf _
                            & "Select 'P' INV_TYPE, SOTORDR2.ORDR_NO INV_NO, SOTORDR2.ORDR_LNO INV_LNO" & vbCrLf _
                            & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                            & ", " & ORDR_QTY_field & " ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE_CURR,SOTORDR2.ORDR_NO,SOTORDR2.ORDR_LNO" & vbCrLf _
                            & " from SOTORDR2,SOTORDR1,ICTSTYL1" & vbCrLf _
                            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                            & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                            & sqlw _
                            & sqlw2

                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Insert into " & SOTINVH1 & vbCrLf _
                        & "(INV_TYPE,INV_NO,CUST_CODE,CUST_STORE_NO,ORDR_CUST_PO,ORDR_NO,WHSE_CODE," & vbCrLf _
                        & "REASON_CODE,INV_DATE,CUST_BILL_TO_CUST,POST_CODE,SHIP_BOL_NO," & vbCrLf _
                        & "SALES_DIVISION_CODE,TERM_CODE,PICK_NO," & vbCrLf _
                        & "CUST_FACTOR_IND,SREP_CODE,INV_COMMENT," & vbCrLf _
                        & "SREP2_CODE,ORDR_DEPT,CURR_CODE,CURR_EXCH_RATE,ORDR_YYYYPP_UPDATED)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTORDR1.ORDR_NO INV_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                        & ", SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE" & vbCrLf _
                        & ", SOTORDR1.REASON_CODE, SOTORDR1.ORDR_SHIP_DATE INV_DATE" & vbCrLf _
                        & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.POST_CODE, NULL SHIP_BOL_NO" & vbCrLf _
                        & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.TERM_CODE, NULL PICK_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_FACTOR_IND, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                        & ", SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                        & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE, '000000' ORDR_YYYYPP_UPDATED" & vbCrLf _
                        & " from SOTORDR1,ARTCUST1 " & vbCrLf _
                        & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                        & Replace(sqlw, "SOTINVH1.", "SOTORDR1.")

                    ASCMAIN1.sql = "Insert into " & SOTINVH1 & vbCrLf _
                        & "(INV_TYPE,INV_NO,CUST_CODE,CUST_STORE_NO,ORDR_CUST_PO,ORDR_NO,WHSE_CODE," & vbCrLf _
                        & "REASON_CODE,INV_DATE,ORDR_BILL_TO_CUST,POST_CODE,SHIP_BOL_NO," & vbCrLf _
                        & "SALES_DIVISION_CODE,TERM_CODE,PICK_NO," & vbCrLf _
                        & "CUST_FACTOR_IND,SREP_CODE,INV_COMMENT," & vbCrLf _
                        & "SREP2_CODE,ORDR_DEPT,CURR_CODE,CURR_EXCH_RATE)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTORDR1.ORDR_NO INV_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                        & ", SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE" & vbCrLf _
                        & ", SOTORDR1.REASON_CODE, SOTORDR1.ORDR_SHIP_DATE INV_DATE" & vbCrLf _
                        & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.POST_CODE, NULL SHIP_BOL_NO" & vbCrLf _
                        & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.TERM_CODE, NULL PICK_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_FACTOR_IND, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                        & ", SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                        & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE" & vbCrLf _
                        & " from SOTORDR1 " & vbCrLf _
                        & " where SOTORDR1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                        & sqlw

                    ASCDATA1.ExecuteSQL()
                End If

                If INV_TYPE_requested <> "O" OrElse INV_TYPE_requested = "B" Then
                    ASCMAIN1.sql = "Insert into " & SOTINVH2 & vbCrLf _
                                                & " (INV_TYPE,INV_NO,INV_LNO,STYLE_CODE,COLOR_CODE,ORDR_UNIT_PRICE,ORDR_QTY_SHIP,ORDR_UNIT_PRICE_CURR,ORDR_NO,ORDR_LNO)" & vbCrLf _
                                                & "Select 'P' INV_TYPE, SOTPICK2.PICK_NO INV_NO, SOTPICK2.PICK_LNO INV_LNO" & vbCrLf _
                                                & ", SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                                                & ", SOTPICK2.PICK_QTY ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_CURR,SOTORDR2.ORDR_NO,SOTORDR2.ORDR_LNO" & vbCrLf _
                                                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTORDR1,ICTSTYL1" & vbCrLf _
                                                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                                                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                                                & "   and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO " & vbCrLf _
                                                & "   and SOTORDR1.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                                                & "   and ICTSTYL1.STYLE_CODE = SOTORDR2.STYLE_CODE" & vbCrLf _
                                                & Replace(sqlw, "SOTINVH1.", "SOTPICK1.")
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Insert into " & SOTINVH1 & vbCrLf _
                            & "(INV_TYPE,INV_NO,CUST_CODE,CUST_STORE_NO,ORDR_CUST_PO,ORDR_NO,WHSE_CODE," & vbCrLf _
                            & "REASON_CODE,INV_DATE,ORDR_BILL_TO_CUST,POST_CODE,SHIP_BOL_NO," & vbCrLf _
                            & "SALES_DIVISION_CODE,TERM_CODE,PICK_NO," & vbCrLf _
                            & "CUST_FACTOR_IND,SREP_CODE,INV_COMMENT," & vbCrLf _
                            & "SREP2_CODE,ORDR_DEPT,CURR_CODE,CURR_EXCH_RATE)" & vbCrLf _
                            & "Select 'P' INV_TYPE, SOTPICK1.PICK_NO INV_NO" & vbCrLf _
                            & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                            & ", SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE" & vbCrLf _
                            & ", SOTORDR1.REASON_CODE, SOTORDR1.ORDR_SHIP_DATE INV_DATE" & vbCrLf _
                            & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.POST_CODE, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                            & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.TERM_CODE, SOTPICK1.PICK_NO" & vbCrLf _
                            & ", SOTORDR1.CUST_FACTOR_IND, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                            & ", SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                            & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE" & vbCrLf _
                            & " from SOTPICK1,SOTORDR1 " & vbCrLf _
                            & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                            & Replace(sqlw, "SOTINVH1.", "SOTPICK1.")
                    ASCDATA1.ExecuteSQL()
                End If

                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INIT_DATE = SYSDATE, INIT_OPER = '" & ASCMAIN1.USER_ID & "'")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " SOTINVH1 Set INV_SALES = (Select Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) from " & SOTINVH2 & " where INV_TYPE = SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " SOTINVH1 Set INV_SALES_CURR = (Select Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE_CURR) from " & SOTINVH2 & " where INV_TYPE = SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)")

                If ASCMAIN1.CLIENT = "NYA" Then

                    ASCMAIN1.sql = "" _
                        & "BEGIN " & vbCrLf _
                        & " DECLARE CURSOR C1 IS " & vbCrLf _
                        & "  SELECT SOTINVH1.INV_TYPE, SOTINVH1.INV_NO" & vbCrLf _
                        & ", ARTCUST2.STAX_CODE, NVL(ARTSTAX1.STAX_RATE,0) STAX_RATE" & vbCrLf _
                        & ", SOTINVH1.INV_SALES, SOTINVH1.INV_SALES_CURR" & vbCrLf _
                        & "   FROM " & SOTINVH1 & " SOTINVH1,ARTSTAX1,ARTCUST2,SOTORDR5,SOTORDR1" & vbCrLf _
                        & "  WHERE SOTINVH1.CURR_CODE = 'CAD'" & vbCrLf _
                        & "    AND SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
                        & "    AND SOTORDR5.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
                        & "    AND SOTORDR5.CUST_ADDR_TYPE = 'ST'" & vbCrLf _
                        & "    AND ARTCUST2.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                        & "    AND ARTCUST2.CUST_ADDR_TYPE = SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
                        & "    AND ARTCUST2.CUST_ADDR_CODE = SOTORDR5.CUST_ADDR_CODE" & vbCrLf _
                        & "    AND ARTSTAX1.STAX_CODE = ARTCUST2.STAX_CODE;" & vbCrLf _
                        & " BEGIN " & vbCrLf _
                        & "  FOR R1 IN C1 LOOP" & vbCrLf _
                        & "   UPDATE " & SOTINVH1 & " Set" & vbCrLf _
                        & "    STAX_CODE = R1.STAX_CODE," & vbCrLf _
                        & "    STAX_RATE = R1.STAX_RATE," & vbCrLf _
                        & "    INV_STAX = ROUND(NVL(R1.INV_SALES * R1.STAX_RATE,0)/100,2)," & vbCrLf _
                        & "    INV_STAX_CURR = ROUND(NVL(R1.INV_SALES_CURR * R1.STAX_RATE,0)/100,2)," & vbCrLf _
                        & "    GST_TAX = ROUND(NVL(R1.INV_SALES * R1.STAX_RATE,0)/100,2)," & vbCrLf _
                        & "    GST_TAX_CURR = ROUND(NVL(R1.INV_SALES_CURR * R1.STAX_RATE,0)/100,2)" & vbCrLf _
                        & "   WHERE INV_TYPE = R1.INV_TYPE AND INV_NO = R1.INV_NO;" & vbCrLf _
                        & "  END LOOP; " & vbCrLf _
                        & " END; " & vbCrLf _
                        & "END;"

                    ASCDATA1.ExecuteSQL()
                End If

                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " SOTINVH1 Set INV_COGS = (Select Sum (ORDR_QTY_SHIP * ORDR_UNIT_COST) from " & SOTINVH2 & " where INV_TYPE = SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_MISC_CHG = 0")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_FREIGHT = 0")

                If ASCMAIN1.CLIENT = "VAN" Then
                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT = NVL(INV_SALES,0) + NVL(INV_FREIGHT,0) + NVL(GST_TAX,0)")
                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_SALES_CURR = INV_SALES, INV_FREIGHT_CURR = INV_FREIGHT, INV_MISC_CHG_CURR = INV_MISC_CHG, INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT, INV_TOTAL_AMOUNT_CURR = INV_TOTAL_AMOUNT, GST_TAX_CURR = GST_TAX")
                Else
                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT_CURR = NVL(INV_SALES_CURR,0) + NVL(INV_FREIGHT_CURR,0) + NVL(GST_TAX_CURR,0)")
                    'ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT_CURR = NVL(INV_SALES_CURR,0) + NVL(INV_FREIGHT_CURR,0) + NVL(INV_STAX_CURR,0) + NVL(GST_TAX_CURR,0)")
                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT_CURR")

                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT = NVL(INV_SALES,0) + NVL(INV_FREIGHT,0) + NVL(INV_STAX,0)")
                    ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_SALES_CURR = INV_SALES, INV_FREIGHT_CURR = INV_FREIGHT, INV_MISC_CHG_CURR = INV_MISC_CHG, INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT, INV_TOTAL_AMOUNT_CURR = INV_TOTAL_AMOUNT, INV_STAX_CURR = INV_STAX, GST_TAX_CURR = GST_TAX where NVL(CURR_CODE,'USD') = 'USD'")

                    If ASCMAIN1.CLIENT = "RGI" Then
                        ASCDATA1.ExecuteSQL($"Update {SOTINVH1} Set INV_MISC_CHG = (select sum(SOTINVHM.INV_MISC_CHG) from SOTINVHM, {SOTINVH1} SOTINVH1 where SOTINVH1.INV_NO = SOTINVHM.INV_NO and NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T' " & Replace(sqlw, "SOTORDR1.", "SOTINVH1.") & ")")
                        ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT = NVL(INV_SALES,0) + NVL(INV_FREIGHT,0) + NVL(INV_STAX,0) + NVL(INV_MISC_CHG,0)")
                        ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_SALES_CURR = INV_SALES, INV_FREIGHT_CURR = INV_FREIGHT, INV_MISC_CHG_CURR = INV_MISC_CHG, INV_TOTAL_AMT_CURR = INV_TOTAL_AMOUNT, INV_TOTAL_AMOUNT_CURR = INV_TOTAL_AMOUNT, INV_STAX_CURR = INV_STAX, GST_TAX_CURR = GST_TAX where NVL(CURR_CODE,'USD') = 'USD'")
                    End If

                End If
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SOTINVH1 & " " & sqlSOTINVH1 & Replace(sqlw, " ORDR_NO", " SOTORDR1.ORDR_NO"))
                ASCDATA1.ExecuteSQL("Insert into " & SOTINVH2 & " " & sqlSOTINVH2 & sqlw)
            End If

            If pro_forma And pfComment <> "" Then
                Dim commentSql As New Dictionary(Of String, String)
                ASCMAIN1.sql = "Select * from " & SOTINVH1
                For Each rowSOTINVH1_PF As DataRow In ASCDATA1.GetDataTable.Select()
                    Dim invNo As String = rowSOTINVH1_PF.Item("INV_NO") & ""
                    Dim invComment As String = rowSOTINVH1_PF.Item("INV_COMMENT") & ""
                    Dim invComment_PF As String = IIf(invComment <> "", invComment & vbCrLf & pfComment, pfComment)
                    commentSql.Add(invNo, invComment_PF)
                Next
                For Each pfCommentUpdate As KeyValuePair(Of String, String) In commentSql
                    ASCMAIN1.sql = "Update " & SOTINVH1 & " set INV_COMMENT = '" & pfCommentUpdate.Value & "' WHERE " & vbCrLf _
                    & " INV_NO = '" & pfCommentUpdate.Key & "'"
                    ASCDATA1.ExecuteSQL()
                Next
            End If

            ASCDATA1.ExecuteSQL("Insert into " & SOTINVP1 & " " & sqlSOTINVP1 & sqlw) ' Replace(sqlw, "ORDR_NO", "SOTORDR1.ORDR_NO")

        End If

        EnforceConstraints(False)
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")
        Fill_Records("SOTORDR9")
        Fill_Records("SOTORDR5_BT")
        Fill_Records("SOTORDR5_ST")
        Fill_Records("ICTSTYL1")
        Fill_Records("SOTINVH1")
        Fill_Records("SOTINVH1_CART_TRACKING_NO")
        Fill_Records("SOTINVH2")
        Fill_Records("SOTINVH9")
        Fill_Records("SOTINVHM")
        Fill_Records("SOTINVHT")
        Fill_Records("SOTINVHT_SUM")
        Fill_Records("SOTINVHT_TS")

        If pro_forma Then
            sql = $"Select SOTINVHM.*, NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE) COUNTRY_NAME, 
                        ROUND(SOTINVHM.INV_MISC_CHG / SOTINVH2.ORDR_QTY_SHIP, 3) INV_MISC_CHG_UNIT
                        From SOTINVHM, {SOTINVP1} SOTINVP1, TATCNTRY, SOTINVH2, SOTINVH1
                        Where SOTINVH1.ORDR_NO = SOTINVP1.ORDR_NO
                        and SOTINVHM.INV_NO = SOTINVH1.INV_NO 
                        AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                        AND SOTINVHM.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE (+)
                        AND SOTINVHM.INV_TYPE = SOTINVH2.INV_TYPE
                        AND SOTINVHM.INV_NO = SOTINVH2.INV_NO
                        AND SOTINVHM.INV_LNO = SOTINVH2.INV_LNO"
            Fill_Records("SOTINVHT",, True, sql)
            sql = $"SELECT SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.COUNTRY_CODE, SOTINVHM.SURCHARGE_PERC,
                    NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE) COUNTRY_NAME,
                    SUM(SOTINVHM.INV_MISC_CHG) TOTAL_TARIFF
                    FROM SOTINVHM, TATCNTRY, {SOTINVP1} SOTINVP1, SOTINVH2, SOTINVH1
                    Where SOTINVH1.ORDR_NO = SOTINVP1.ORDR_NO
                    and SOTINVHM.INV_NO = SOTINVH1.INV_NO 
                    AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                    AND SOTINVH2.INV_TYPE = SOTINVHM.INV_TYPE 
                    AND SOTINVH2.INV_NO = SOTINVHM.INV_NO
                    AND SOTINVH2.INV_LNO = SOTINVHM.INV_LNO 
                    AND SOTINVHM.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE (+)
                    GROUP BY SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.COUNTRY_CODE, SOTINVHM.SURCHARGE_PERC,
                    NVL(INITCAP(TATCNTRY.COUNTRY_NAME), SOTINVHM.COUNTRY_CODE)"
            Fill_Records("SOTINVHT_SUM",, True, sql)
            sql = $"Select SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.INV_LNO, SUM(SOTINVHM.INV_MISC_CHG) INV_MISC_CHG_SUM,
                    Round(SUM(SOTINVHM.INV_MISC_CHG / SOTINVH2.ORDR_QTY_SHIP), 4) INV_MISC_CHG_UNIT_SUM
                    From SOTINVHM, {SOTINVP1} SOTINVP1, SOTINVH2, SOTINVH1
                    Where SOTINVH1.ORDR_NO = SOTINVP1.ORDR_NO
                    and SOTINVHM.INV_NO = SOTINVH1.INV_NO 
                    AND NVL(SOTINVHM.MISC_CHARGE_TYPE, '?') = 'T'
                    AND SOTINVHM.INV_TYPE = SOTINVH2.INV_TYPE
                    AND SOTINVHM.INV_NO = SOTINVH2.INV_NO
                    AND SOTINVHM.INV_LNO = SOTINVH2.INV_LNO
                    GROUP BY SOTINVHM.INV_TYPE, SOTINVHM.INV_NO, SOTINVHM.INV_LNO"
            Fill_Records("SOTINVHT_TS",, True, sql)
        End If


        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            If pro_forma Then
                INV_NO = ASCDATA1.GetDataValue($"Select INV_NO from SOTINVH1 where ORDR_NO = '{INV_NO}'")
            End If
            If dst.Tables("SOTINVHM").Select($"INV_NO = '{INV_NO}' AND ISNULL(MISC_CHARGE_TYPE, '') = '' ").Length > 0 Then
                rowSOTINVH1.Item("MISC_CHARGES") = "1"
            Else
                rowSOTINVH1.Item("MISC_CHARGES") = "0"
            End If

            If dst.Tables("SOTINVHT_SUM").Select($"INV_NO = '{INV_NO}'").Length > 0 Then
                rowSOTINVH1.Item("TARIFFS") = "1"
            Else
                rowSOTINVH1.Item("TARIFFS") = "0"
            End If
        Next

        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("")
            Dim INV_NO As String = rowSOTINVH2.Item("INV_NO")
            Dim INV_LNO As Int16 = rowSOTINVH2.Item("INV_LNO")
            If pro_forma Then
                INV_NO = ASCDATA1.GetDataValue($"Select INV_NO from SOTINVH1 where ORDR_NO = '{INV_NO}'")
            End If

            If dst.Tables("SOTINVHT").Select($"INV_NO = '{INV_NO}' and INV_LNO = {INV_LNO}").Length > 0 Then
                rowSOTINVH2.Item("TARIFFS") = "1"
            Else
                rowSOTINVH2.Item("TARIFFS") = "0"
            End If
        Next

        ' ALTER TABLE SOTINVHM ADD INV_MISC_CHG_CURR  NUMBER(13,2);
        ' UPDATE SOTINVHM SET INV_MISC_CHG_CURR = INV_MISC_CHG WHERE INV_MISC_CHG_CURR IS NULL;

        For Each rowSOTINVHM As DataRow In dst.Tables("SOTINVHM").Select("")
            Dim INV_MISC_CHG As Decimal = Val(rowSOTINVHM.Item("INV_MISC_CHG") & "")
            Dim INV_MISC_CHG_CURR As Decimal = Val(rowSOTINVHM.Item("INV_MISC_CHG_CURR") & "")
            If INV_MISC_CHG_CURR = 0 And INV_MISC_CHG <> 0 Then
                INV_MISC_CHG_CURR = INV_MISC_CHG ' WORRIED THAT WE HAVE NOT CAUGHT ALL AREAS IN CODE THAT CREATE SOTINVHM RECORDS
                rowSOTINVHM.Item("INV_MISC_CHG_CURR") = INV_MISC_CHG_CURR
            End If
        Next

        Fill_Records("SOTSHIP1")
        Fill_Records("SOTPICK1")
        Fill_Records("SOTPICK2")
        Fill_Records("ARTCUST1")
        Fill_Records("SOTCUBE1")

        Fill_Records("POTORDR1")
        Fill_Records("POTORDR2")

        Fill_Records("EDT850T1")
        Fill_Records("EDT850T2")

        Dim onCounter As Int16 = 1
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            rowSOTINVH1.Item("BT") = "BT"
            rowSOTINVH1.Item("ST") = "ST"

            rowSOTINVH1.Item("AR_PARM_KEY") = "Z"

            If ASCMAIN1.CLIENT = "NYA" Then
                If rowSOTINVH1.Item("CURR_CODE") = "CAD" Then
                    rowSOTINVH1.Item("AR_PARM_KEY") = "CAD"
                End If

                Dim SALES_DIVISION_CODE As String = rowSOTINVH1.Item("SALES_DIVISION_CODE") & ""
                Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(SALES_DIVISION_CODE)
                If rowSOTSDIV1 IsNot Nothing AndAlso rowSOTSDIV1.Item("SEG4_CODE") & "" <> "" Then
                    rowSOTINVH1.Item("AR_PARM_KEY") = "Z" & rowSOTSDIV1.Item("SEG4_CODE")
                End If
            End If

            ' rowSOTINVH1.Item("TOTAL_UNITS") = Val(dst.Tables("SOTINVH2").Compute("SUM(ORDR_QTY_SHIP)", "") & "")
            Dim rowSOTINVH1_CART_TRACKING_NO As DataRow = dst.Tables("SOTINVH1_CART_TRACKING_NO").Rows.Find(New String() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")})
            If rowSOTINVH1_CART_TRACKING_NO IsNot Nothing Then rowSOTINVH1.Item("CART_TRACKING_NO") = rowSOTINVH1_CART_TRACKING_NO.Item("CART_TRACKING_NO")

            rowSOTINVH1.Item("TOTAL_CUBE") = Val(dst.Tables("SOTCUBE1").Compute("SUM(TOTAL_CUBE)", "INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'") & String.Empty)

            Try
                If ASCMAIN1.DBS_SERVER = "RGI" OrElse ASCMAIN1.DBS_COMPANY = "RGI" Then
                    Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                    If ORDR_NO.Length = 0 Then
                        Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE") & String.Empty
                        Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
                        onCounter = onCounter + 1
                        ORDR_NO = "X" & onCounter.ToString.PadLeft(9, "0")

                        Dim rowARTCUST2 As DataRow = dst.Tables("ARTCUST2").Rows.Find(New Object() {CUST_CODE, "MK", CUST_STORE_NO})

                        If rowARTCUST2 Is Nothing Then
                            rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})
                        End If

                        If rowARTCUST2 IsNot Nothing Then
                            rowSOTINVH1.Item("ORDR_NO") = ORDR_NO
                            Dim rowSOTORDR5_ST As DataRow = dst.Tables("SOTORDR5_ST").NewRow
                            rowSOTORDR5_ST.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDR5_ST.Item("CUST_ADDR_TYPE") = rowARTCUST2.Item("CUST_ADDR_TYPE")
                            rowSOTORDR5_ST.Item("CUST_ADDR_CODE") = CUST_STORE_NO
                            rowSOTORDR5_ST.Item("CUST_NAME") = rowARTCUST2.Item("CUST_NAME")
                            rowSOTORDR5_ST.Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_ADDR1")
                            rowSOTORDR5_ST.Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_ADDR2")
                            rowSOTORDR5_ST.Item("CUST_CITY") = rowARTCUST2.Item("CUST_CITY")
                            rowSOTORDR5_ST.Item("CUST_STATE") = rowARTCUST2.Item("CUST_STATE")
                            rowSOTORDR5_ST.Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_ZIP_CODE")
                            rowSOTORDR5_ST.Item("CUST_COUNTRY") = rowARTCUST2.Item("CUST_COUNTRY")
                            rowSOTORDR5_ST.Item("CUST_CONTACT") = rowARTCUST2.Item("CUST_CONTACT")
                            rowSOTORDR5_ST.Item("CUST_PHONE") = rowARTCUST2.Item("CUST_PHONE")
                            rowSOTORDR5_ST.Item("CUST_EXT") = rowARTCUST2.Item("CUST_EXT")
                            rowSOTORDR5_ST.Item("CUST_FAX") = rowARTCUST2.Item("CUST_FAX")
                            rowSOTORDR5_ST.Item("CUST_EMAIL") = rowARTCUST2.Item("CUST_EMAIL")
                            rowSOTORDR5_ST.Item("CUST_ADDR3") = rowARTCUST2.Item("CUST_ADDR3")
                            dst.Tables("SOTORDR5_ST").Rows.Add(rowSOTORDR5_ST)
                        End If
                    End If
                End If
            Catch ex As Exception
                ' nothing - donto want to crash invoice printing with this work around for RGI
            End Try


        Next

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            ' TO SET INV_UOM FOR AHOLD INVOICE WHERE THEY ORDERED IN INNER PACKS
            ' WE NEED TO MAKE THIS OPTION SOMETHING AVAILABLE TO SALES ORDER ENTRY
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {"0000782142", 3})
            If rowSOTORDR2 IsNot Nothing Then
                rowSOTORDR2.Item("INV_UOM") = "I"
            End If
        End If


        'Fill_Records("ARTCUSTZ")
        If subUPCSupport Then
            Fill_Records("ICTXLSPS")
        End If

        EnforceConstraints(True)

        Dim AR_PARM_INVOICE_RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If AR_PARM_INVOICE_RPT <> "" Then RPT = AR_PARM_INVOICE_RPT ' "SORINVP1"

        Prepare_Invoice_Header(sqlw)
    End Sub

    Sub Prepare_Invoice_Header(sqlw As String)

        Check_Invoice_Totals()

        ' Set flag if Invoice has Misc Charges
        'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH5"), "SO_ORDER_NO").Rows
        '    Dim SO_ORDER_NO As String = row.Item("SO_ORDER_NO")
        '    Dim rowSOTINVH1 = dst.Tables("SOTINVH1").Rows.Find(New Object() {SO_ORDER_NO})
        '    rowSOTINVH1.ITEM("MISC") = "1"
        'Next

        'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1"), "SALES_DIVISION_CODE").Rows
        '    Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(row.Item("ORDR_DIV_CODE"))
        '    If rowSOTSDIV1.Item("DIVISION_LOGO_FILENAME") & "" <> "" And rowSOTSDIV1.Item("DIVISION_LOGO").ToString & "" = "" Then
        '        Dim DIVISION_LOGO_FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & rowSOTSDIV1.Item("DIVISION_LOGO_FILENAME")
        '        If My.Computer.FileSystem.FileExists(DIVISION_LOGO_FILENAME) Then
        '            rowSOTSDIV1.Item("DIVISION_LOGO") = ASCMAIN1.GetImageData(DIVISION_LOGO_FILENAME)
        '        End If
        '    End If
        'Next


        ' Load SOTINVH1 - Based on Run-Time Options

        'SOTINVP1 = ASCMAIN1.Temp_Table(sqlw)
        'ASCDATA1.ExecuteSQL("Alter Table " & SOTINVP1 & " Add Primary Key (INV_TYPE, INV_NO)")

        If consolidated_invoice Then Create_Consolidated_Invoice()
    End Sub

    Sub Create_Consolidated_Invoice()

        '  If Absx1.chkFor("CHKCONS_INV").Checked Then
        ASCMAIN1.Progress("Consolidating Invoices")

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("INV_NO_CONS <> ''")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS")
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
            ASCDATA1.DeleteRows(dst.Tables("SOTINVH2"), "INV_NO = '" & INV_NO & "'")
            ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "ORDR_NO = '" & ORDR_NO & "'")
        Next
        ASCDATA1.DeleteRows(dst.Tables("SOTINVH1"), "INV_NO_CONS <> '' and INV_NO <> INV_NO_CONS")

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("INV_NO = INV_NO_CONS")
            Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS")

            Dim PICK_NO As String = rowSOTINVH1.Item("PICK_NO")
            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)

            ASCMAIN1.sql = "Select SUM (PICK_CNT_CARTONS) PICK_CNT_CARTONS, SUM (PICK_TOTAL_WGT) PICK_TOTAL_WGT" & vbCrLf _
                & " from SOTPICK1 where INV_NO IN " _
                & " (Select INV_NO from SOTINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "')"
            Dim row As DataRow = ASCDATA1.GetDataRow

            Dim PICK_CNT_CARTONS As Int64 = Val(row.Item("PICK_CNT_CARTONS") & "")
            Dim PICK_TOTAL_WGT As Int64 = Val(row.Item("PICK_TOTAL_WGT") & "")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = PICK_CNT_CARTONS
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT

            ASCMAIN1.sql = "Select " _
                & "  Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & " from SOTINVH2,SOTINVH1" & vbCrLf _
                & " where SOTINVH1.INV_NO_CONS = '" & INV_NO_CONS & "'" _
                & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
            Dim rowSOTINVH2_Totals As DataRow = ASCDATA1.GetDataRow
            Dim TOTAL_UNITS As Int64 = Val(rowSOTINVH2_Totals.Item("ORDR_QTY_SHIP") & "")

            ASCMAIN1.sql = "Select " _
                & "  Sum (INV_SALES) INV_SALES" & vbCrLf _
                & ", Sum (INV_FREIGHT) INV_FREIGHT" & vbCrLf _
                & ", Sum (INV_MISC_CHG) INV_MISC_CHG" & vbCrLf _
                & ", Sum (INV_TOTAL_AMOUNT) INV_TOTAL_AMOUNT from SOTINVH1" & vbCrLf _
                & " where SOTINVH1.INV_NO_CONS = '" & INV_NO_CONS & "'"
            Dim rowSOTINVH1_CONS As DataRow = ASCDATA1.GetDataRow

            rowSOTINVH1.Item("INV_SALES") = Val(rowSOTINVH1_CONS.Item("INV_SALES") & "")
            rowSOTINVH1.Item("INV_FREIGHT") = Val(rowSOTINVH1_CONS.Item("INV_FREIGHT") & "")
            rowSOTINVH1.Item("INV_MISC_CHG") = Val(rowSOTINVH1_CONS.Item("INV_MISC_CHG") & "")
            rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = Val(rowSOTINVH1_CONS.Item("INV_TOTAL_AMOUNT") & "")

            'rowSOTINVH1.Item("TOTAL_UNITS") = TOTAL_UNITS

            ASCMAIN1.sql = "Select SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.CUST_CODE, SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CUST_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.CUST_SKU" & vbCrLf _
                & ", SUM(NVL(SOTINVH2.ORDR_UNIT_COST,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) AS ORDR_UNIT_COST_X" & vbCrLf _
                & ", SUM(NVL(SOTINVH2.ORDR_UNIT_PRICE,0) * NVL(SOTINVH2.ORDR_QTY_SHIP,0)) AS ORDR_UNIT_PRICE_X" & vbCrLf _
                & ", MAX(NVL(SOTINVH2.ORDR_UNIT_COST,0)) AS ORDR_UNIT_COST_MAX" & vbCrLf _
                & ", MAX(NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) AS ORDR_UNIT_PRICE_MAX" & vbCrLf _
                & ", MAX(NVL(SOTORDR2.CARTON_PACK_QTY,0)) AS CARTON_PACK_QTY" & vbCrLf _
                & ", SUM(SOTINVH2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP From SOTINVH2,SOTORDR2,SOTINVH1,SOTPICK2" & vbCrLf _
                & " where SOTINVH2.INV_TYPE = 'I' and SOTINVH2.INV_NO in" & vbCrLf _
                & " (Select INV_NO from SOTINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "')" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTPICK2.PICK_NO = SOTINVH1.PICK_NO and SOTPICK2.PICK_LNO = SOTINVH2.INV_LNO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & " group by " & vbCrLf _
                & " SOTINVH2.STYLE_CODE, SOTINVH2.COLOR_CODE, SOTINVH2.CUST_CODE, SOTINVH2.ORDR_YYYYPP_UPDATED, SOTINVH2.STYLE_CUST_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.CUST_SKU"

            Dim INV_LNO As Int32 = 0
            For Each rowSOTINVH2_CONS As DataRow In ASCDATA1.GetDataTable.Select _
                    ("", "STYLE_CODE, COLOR_CODE, CUST_CODE, ORDR_YYYYPP_UPDATED, STYLE_CUST_CODE, STYLE_DESC, CUST_SKU")
                INV_LNO += 1
                Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                With rowSOTINVH2
                    .Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE") & ""
                    .Item("INV_NO") = rowSOTINVH1.Item("INV_NO") & ""
                    .Item("INV_LNO") = INV_LNO
                    .Item("STYLE_CODE") = rowSOTINVH2_CONS.Item("STYLE_CODE") & ""
                    .Item("COLOR_CODE") = rowSOTINVH2_CONS.Item("COLOR_CODE") & ""
                    Dim ORDR_QTY_SHP As Int64 = Val(rowSOTINVH2_CONS.Item("ORDR_QTY_SHIP") & "")
                    If ORDR_QTY_SHP = 0 Then
                        .Item("ORDR_UNIT_COST") = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_COST_MAX") & "")
                        .Item("ORDR_UNIT_PRICE") = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_PRICE_MAX") & "")
                    Else
                        .Item("ORDR_UNIT_COST") = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_COST_X") & "") / ORDR_QTY_SHP
                        .Item("ORDR_UNIT_PRICE") = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_PRICE_X") & "") / ORDR_QTY_SHP
                    End If
                    .Item("ORDR_QTY_SHIP") = ORDR_QTY_SHP
                    .Item("CUST_CODE") = rowSOTINVH2_CONS.Item("CUST_CODE") & ""
                    .Item("ORDR_YYYYPP_UPDATED") = rowSOTINVH2_CONS.Item("ORDR_YYYYPP_UPDATED") & ""
                    .Item("STYLE_CUST_CODE") = rowSOTINVH2_CONS.Item("STYLE_CUST_CODE") & ""
                    .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & ""
                    .Item("ORDR_LNO") = INV_LNO
                End With
                dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)

                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                With rowSOTORDR2
                    .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & ""
                    .Item("ORDR_LNO") = INV_LNO
                    .Item("STYLE_CODE") = rowSOTINVH2_CONS.Item("STYLE_CODE") & ""
                    .Item("COLOR_CODE") = rowSOTINVH2_CONS.Item("COLOR_CODE") & ""
                    .Item("STYLE_DESC") = rowSOTINVH2_CONS.Item("STYLE_DESC") & ""
                    .Item("CUST_SKU") = rowSOTINVH2_CONS.Item("CUST_SKU") & ""
                    .Item("CARTON_PACK_QTY") = Val(rowSOTINVH2_CONS.Item("CARTON_PACK_QTY") & "")
                    .Item("ORDR_STATUS") = "F"
                End With
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
            Next

            'ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_DESC, CUST_SKU" & vbCrLf _
            '   & " from SOTORDR2 where ORDR_NO in " & vbCrLf _
            '   & " (Select ORDR_NO from SOTINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "')" & vbCrLf _
            '   & " and ORDR_QTY_SHIP <> 0" & vbCrLf _
            '   & " group by " & vbCrLf _
            '   & " STYLE_CODE, COLOR_CODE, STYLE_DESC, CUST_SKU"

            'Dim ORDR_LNO As Int32 = 0
            'For Each rowSOTORDR2_CONS As DataRow In ASCDATA1.GetDataTable.Select _
            '        ("", "STYLE_CODE, COLOR_CODE")
            '    ORDR_LNO += 1
            '    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
            '    With rowSOTORDR2
            '        .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & ""
            '        .Item("ORDR_LNO") = ORDR_LNO
            '        .Item("STYLE_CODE") = rowSOTORDR2_CONS.Item("STYLE_CODE") & ""
            '        .Item("COLOR_CODE") = rowSOTORDR2_CONS.Item("COLOR_CODE") & ""
            '        .Item("STYLE_DESC") = rowSOTORDR2_CONS.Item("STYLE_DESC") & ""
            '        .Item("CUST_SKU") = rowSOTORDR2_CONS.Item("CUST_SKU") & ""
            '        .Item("ORDR_STATUS") = "F"
            '    End With
            '    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
            'Next
        Next
        '    End If
    End Sub

    Sub Check_Invoice_Totals()

        'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("CURR_CODE <> 'USD'")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_0") = rowSOTINVH1.Item("ORDR_DIV_CODE")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_R") = rowSOTINVH1.Item("ORDR_DIV_CODE")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_R") = "E"
        'Next

        'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select
        '    Dim DIFF As Decimal = 0
        '    DIFF = Val(rowSOTINVH1.Item("ORDR_AMT") & "") _
        '    - (Val(rowSOTINVH1.Item("ORDR_AMT_GROSS") & "") - Val(rowSOTINVH1.Item("ORDR_AMT_ALLOW") & ""))
        '    If System.Math.Abs(System.Math.Round(DIFF, 2)) > 0.01 Then
        '        MsgBox("Order No " & rowSOTINVH1.Item("SO_ORDER_NO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        '    DIFF = Val(rowSOTINVH1.Item("ORDR_AMT") & "") + Val(rowSOTINVH1.Item("ORDR_MISC_CHG") & "") + Val(rowSOTINVH1.Item("ORDR_WD_CHG") & "") <> Val(rowSOTINVH1.Item("ORDR_TOTAL_AMT") & "")
        '    If System.Math.Abs(System.Math.Round(DIFF, 2)) > 0.01 Then
        '        MsgBox("Order No " & rowSOTINVH1.Item("SO_ORDER_NO") & " does not total", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        'Next
        'For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select
        '    If Val(rowSOTINVH2.Item("CASES") & "") <> Val(rowSOTINVH2.Item("QTY_CASES") & "") Then
        '        MsgBox("Order No " & rowSOTINVH2.Item("SO_ORDER_NO") & ", Line " & rowSOTINVH2.Item("SO_ORDER_LNO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        '    If System.Math.Round(Val(rowSOTINVH2.Item("UNITS") & ""), 0) _
        '    <> System.Math.Round(Val(rowSOTINVH2.Item("QTY_UNITS") & ""), 0) Then
        '        MsgBox("Order No " & rowSOTINVH2.Item("SO_ORDER_NO") & ", Line " & rowSOTINVH2.Item("SO_ORDER_LNO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        'Next

    End Sub

    Sub Email_Invoice(ByVal rowSOTINVH1 As DataRow, ByVal rowARTCUST1 As DataRow, ByVal attachment As String)
        Me.Cursor = Cursors.WaitCursor

        Using frmTAFSEND1 As New TAFSEND1(Me)

            With frmTAFSEND1
                .EMAIL_KEY = "INV"
                .SEND_TO = rowARTCUST1.Item("CUST_EMAIL_TO") & ""
                If ASCMAIN1.USER_EMAIL = "" Then
                    .SEND_FROM = "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN")
                Else
                    .SEND_FROM = ASCMAIN1.USER_EMAIL
                End If
                .SEND_FROM_NAME = ASCMAIN1.USER_NAME
                If rowARTCUST1.Item("CUST_EMAIL_CC") & "" <> "" Then
                    .SEND_CC = rowARTCUST1.Item("CUST_EMAIL_CC") & ""
                End If

                Dim customInfo As String = rowARTCUST1.Item("CUSTOM_SUBJECT") & " "
                Dim companyName As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " "
                Dim custPO As String = rowSOTINVH1.Item("ORDR_CUST_PO") & " "
                Dim invNo As String = "INV " & rowSOTINVH1.Item("ORDR_INV_NO") & " "
                Dim invDate As String = rowSOTINVH1.Item("ORDR_INV_DATE") & ""

                Dim subjectLine As String = customInfo & companyName & custPO & invNo & invDate

                .SEND_SUBJECT = subjectLine

                Dim sal As String = ""
                If rowARTCUST1.Item("CUST_SALUTATION") & "" <> "" Then
                    sal = rowARTCUST1.Item("CUST_SALUTATION") & "," & vbCrLf
                Else
                    sal = "To whom it may concern," & vbCrLf
                End If

                Dim body As String = ""
                If "" <> "" Then
                    body = "Please find your invoice attached."
                Else
                    body = rowARTCUST1.Item("CUST_BILLING_NOTE") & "" <> ""
                End If

                .SEND_BODY = sal & body
                .SEND_ATTACHMENT = attachment
                .SEND_METHOD = "E"
                .SEND_ENTITY_CAPTION = "Sold-To"
                .SEND_ENTITY_TABLE = "ARTCUST1"
                .SEND_ENTITY_KEY = rowSOTINVH1.Item("CUST_CODE")
                .SEND_ENTITY_NAME = rowARTCUST1.Item("CUST_NAME") & ""

                .Send_email_automatically(False)

                If .SEND_STATUS <> "S" Then
                    TAC.TACMAIN1.Record_Event("SOTORDR1", rowSOTINVH1.Item("SO_ORDER_NO"), DATETIME_STAMP, ASCMAIN1.USER_ID, "E", "Emailed Invoice to " & .SEND_TO, rowSOTINVH1.Item("CUST_CODE"))
                Else
                    MsgBox("Error Occured: Could Not Send Email for Invoice: " & rowSOTINVH1.Item("ORDR_INV_NO"), MsgBoxStyle.OkOnly, "Error")
                End If
            End With
        End Using
    End Sub

    Private Sub btnBATCH_VAN_Click(sender As Object, e As EventArgs) Handles btnBATCH_VAN.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Bulk Invoice Upload"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("You Will Be Prompted For An Excel File To Use For")
        iMSG.AppendLine("Bulk Upload. There Must Be One Column Tilted")
        iMSG.AppendLine(String.Format("{0}Invoice No{0} On A Sheet Named {0}Invoice No{0}", Chr(34)))
        iMSG.AppendLine("Which Will Be Used To Load The Invoices.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            Dim fileToImport As String = String.Empty
            Dim tableData As New DataTable

            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Open File To Upsert"
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
                openFileDialog1.FilterIndex = 1
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    fileToImport = openFileDialog1.FileName
                End If

                openFileDialog1.Dispose()
            End Using
            If fileToImport.Length = 0 Then
                Exit Sub
            End If
            ASCMAIN1.Progress("Reading File")
            Me.Cursor = Cursors.WaitCursor
            Try
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & "data source=" & fileToImport & ";" & "Extended Properties=Excel 8.0;"
                Using cn As New System.Data.OleDb.OleDbConnection(strConnection)
                    Using cmd As New System.Data.OleDb.OleDbDataAdapter("select * from [Invoice No$]", cn)
                        ' Select the data from Sheet1 of the workbook.
                        cn.Open()
                        cmd.Fill(tableData)
                        cn.Close()
                        cmd.Dispose()
                    End Using
                    cn.Dispose()
                End Using
            Catch ex As Exception
                Try
                    Dim strConnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fileToImport & ";Extended Properties=""Excel 12.0;HDR=YES;IMEX=1"""
                    Using cn As New System.Data.OleDb.OleDbConnection(strConnection)
                        Using cmd As New System.Data.OleDb.OleDbDataAdapter("select * from [Invoice No$]", cn)
                            ' Select the data from Sheet1 of the workbook.
                            cn.Open()
                            cmd.Fill(tableData)
                            cn.Close()
                            cmd.Dispose()
                        End Using
                        cn.Dispose()
                    End Using
                Catch ex2 As Exception
                    MsgBox("Excel Provided Was Not Formatted Correctly", vbOKOnly, "Aborted")
                    Exit Sub
                End Try
            End Try

            Dim ColFound As Boolean = False
            For Each dc As DataColumn In tableData.Columns
                If dc.ColumnName = "Invoice No" Then
                    ColFound = True
                    Exit For
                End If
            Next
            If Not ColFound Then
                MsgBox(String.Format("Could Not Find Column Named {0}Invoice No{0}.", Chr(34)), vbOKOnly, "Aborted")
            Else
                Dim INVOICE_NOS As String = ""
                For Each rowData As DataRow In tableData.Select()
                    If rowData.Item("Invoice No").ToString & String.Empty <> "" Then
                        Dim THIS_INVOICE As String = rowData.Item("Invoice No").ToString & String.Empty
                        If THIS_INVOICE.Length <> 10 Then
                            THIS_INVOICE = THIS_INVOICE.PadLeft(10, "0")
                        End If
                        INVOICE_NOS = INVOICE_NOS & "," & THIS_INVOICE
                    End If
                Next
                If INVOICE_NOS.Length > 0 Then
                    INVOICE_NOS = INVOICE_NOS.Substring(1, INVOICE_NOS.Length - 1)
                    Dim rowASTDSQLA As DataRow = tblASTDSQLA.Select("COLUMN_NAME = 'INV_NO'").FirstOrDefault
                    If Not IsNothing(rowASTDSQLA) Then
                        rowASTDSQLA.Item("CODE_VALUES") = INVOICE_NOS
                    Else
                        MsgBox("Could Not Find Invoice Number In Selector!", vbOKOnly, "Error")
                    End If
                Else
                    MsgBox("Could Not Find Invoice Number To Import!", vbOKOnly, "Error")
                End If
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        Else
            MsgBox("OK.  Let Me Know When You Are Ready.", vbOKOnly, "Aborted")
        End If
    End Sub

    Private Sub chkBATCH_VAN_CheckedChanged(sender As Object, e As EventArgs) Handles chkBATCH_VAN.CheckedChanged
        If chkBATCH_VAN.Checked = True Then
            MsgBox("Please Select A Folder To Store Your PDF's", vbOKOnly, "Save PDFs")
            Dim fbd As New FolderBrowserDialog
            fbd.ShowDialog()
            If fbd.SelectedPath.Length > 0 Then
                BATCH_VAN_FOLDER = fbd.SelectedPath
            Else
                BATCH_VAN_FOLDER = ""
            End If
        End If

    End Sub
End Class