Imports System.Math

Public Class SORPICKE

    Shadows SUBT As String = String.Empty
    Private sqlEDT850T2 As String = String.Empty
    Private Const numDetailItems As Int16 = 11

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "N"
        SUBT = String.Empty

        Dim sqlw As String = String.Empty

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTPICK1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = "SORPICKE"
        SUBT = ""
        Generate_Report(RPT, "Consolidated Packing Slip", SUBT, "", "", "", False)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)

    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        Dim sql As String = String.Empty

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = "ROWNUM < 1"

        With dst

            Create_TDA(dst.Tables.Add, "SOTPICK1", "*")
            Create_TDA(dst.Tables.Add, "SOTPICK2", "*")

            Create_TDA(dst.Tables.Add, "SOTORDR1", "*")
            Create_TDA(dst.Tables.Add, "SOTORDR2", "*")
            .Tables("SOTORDR2").Columns.Add("UPC_CODE", GetType(System.String))
            .Tables("SOTORDR2").Columns.Add("ORDR_UNIT_PRICE_ORIG", GetType(System.Decimal))

            Create_TDA(dst.Tables.Add, "ECTSTYL1", "*")

            Create_TDA(dst.Tables.Add, "WHTMOVE1", "*")
            Create_TDA(dst.Tables.Add, "WHTMOVE2", "*")

            Create_TDA(dst.Tables.Add, "WHTPKGM1", "*")
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "*")

            ASCMAIN1.sql = "Select SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE from SOTORDR2 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "WHTLOCBE", "**", 0, False, "", 2)
            .Tables("WHTLOCBE").Columns.Add("LOCATIONS")

            ASCMAIN1.sql = "SELECT SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTORDR1.ORDR_CUST_PO, SOTPICK2.PICK_QTY" _
                & " FROM SOTPICK1, SOTORDR1, SOTORDR2, SOTPICK2 WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "", 3)

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        ' parms(0) should be Ship Bol Nos
        Dim sqlw As String = String.Empty
        If parms IsNot Nothing AndAlso parms.Length > 0 Then
            sqlw = CStr(parms(0)) & String.Empty
        Else
            sqlw = "'9876543210'"
        End If

        Fill_Records("ECTSTYL1", String.Empty, True, "SELECT * FROM ECTSTYL1")
        Fill_Records("WHTPKGM1", String.Empty, True, "SELECT * FROM WHTPKGM1")

        Dim tempTable As String = ASCMAIN1.Temp_Table("Select PICK_NO, ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO IN (" & sqlw & ")")

        ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FOB" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
                & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE" & vbCrLf _
                & " from SOTPICK1, SOTORDR1, SOTSHIP1 " _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO in (Select PICK_NO FROM " & tempTable & ")"
        Fill_Records("SOTPICK1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTPICK2.*, SOTPICK1.ORDR_NO," & vbCrLf _
                & " SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTORDR2.STYLE_DESC, SOTORDR2.CARTON_PACK_QTY," & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.STYLE_CODE_SUB, ICTCOLR1.COLOR_CODE_LONG, " & vbCrLf _
                & " SOTORDR2.RANGE_STYLE_CODE, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP, ICTSTYL1.CASE_CUBE, SOTPICK1.SHIP_BOL_NO"

        If ASCMAIN1.CLIENT = "RGI" OrElse ASCMAIN1.CLIENT = "NYA" Then
            ASCMAIN1.sql &= ", SOTORDR2.ORDR_PRICE_SOURCE, SOTORDR2.COMM_RATE " & vbCrLf
        End If

        ASCMAIN1.sql &= " from SOTPICK2, SOTPICK1, SOTORDR2, SOTSHIP1, ICTSTYL1, ICTCOLR1" & vbCrLf

        ASCMAIN1.sql &= vbCrLf _
            & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and ICTCOLR1.COLOR_CODE (+) = SOTORDR2.COLOR_CODE" & vbCrLf _
            & "   and SOTPICK2.PICK_NO in (Select PICK_NO FROM " & tempTable & ")"
        Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO in (Select ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_ORIG from SOTORDR2 where ORDR_NO in (Select ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM ICTSTYC1 WHERE (STYLE_CODE, COLOR_CODE) IN (Select STYLE_CODE, COLOR_CODE from SOTORDR2 where ORDR_NO in (Select ORDR_NO FROM " & tempTable & "))"
        Fill_Records("ICTSTYC1", "", True, ASCMAIN1.sql)

        dst.Tables("SOTPICKX").Clear()
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY > 0")

            'SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE, SOTPICK1.PICK_NO
            Dim STYLE_CODE As String = rowSOTPICK2.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = rowSOTPICK2.Item("COLOR_CODE") & String.Empty
            Dim PICK_NO As String = rowSOTPICK2.Item("PICK_NO") & String.Empty

            Dim rowSOTPICKX As DataRow = Nothing
            If dst.Tables("SOTPICKX").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND PICK_NO = '" & PICK_NO & "'").Length > 0 Then
                rowSOTPICKX = dst.Tables("SOTPICKX").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND PICK_NO = '" & PICK_NO & "'")(0)
                rowSOTPICKX.Item("PICK_QTY") = Val(rowSOTPICKX.Item("PICK_QTY") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
            Else
                rowSOTPICKX = dst.Tables("SOTPICKX").NewRow
                rowSOTPICKX.Item("STYLE_CODE") = rowSOTPICK2.Item("STYLE_CODE")
                rowSOTPICKX.Item("COLOR_CODE") = rowSOTPICK2.Item("COLOR_CODE")
                rowSOTPICKX.Item("PICK_NO") = rowSOTPICK2.Item("PICK_NO")
                rowSOTPICKX.Item("ORDR_NO") = rowSOTPICK2.Item("ORDR_NO")

                Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                rowSOTPICKX.Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")

                rowSOTPICKX.Item("PICK_QTY") = rowSOTPICK2.Item("PICK_QTY")
                dst.Tables("SOTPICKX").Rows.Add(rowSOTPICKX)
            End If
        Next

        ' Need to consolidate so it looks like one order
        If dst.Tables("SOTPICK1").Rows.Count > 0 Then
            Dim PICK_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO") & String.Empty
            Dim ORDR_NO As String = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")(0).Item("ORDR_NO") & String.Empty

            For Each row As DataRow In dst.Tables("SOTPICK1").Select("PICK_NO <> '" & PICK_NO & "'")
                row.Delete()
            Next
            dst.Tables("SOTPICK1").AcceptChanges()

            For Each row As DataRow In dst.Tables("SOTORDR1").Select("ORDR_NO <> '" & ORDR_NO & "'")
                row.Delete()
            Next
            dst.Tables("SOTORDR1").AcceptChanges()

            Dim ORDR_LNO As Int16 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & String.Empty) + 1
            PICK_NO = "C" & PICK_NO.Substring(2)
            dst.Tables("SOTPICK1").Rows(0).Item("PICK_NO") = PICK_NO
            dst.Tables("SOTPICK1").Rows(0).Item("INV_NO") = "CONSOLID"

            For Each row As DataRow In dst.Tables("SOTORDR2").Select("", "ORDR_NO, ORDR_LNO")
                Dim PICK_ORDR_NO As String = row.Item("ORDR_NO") & String.Empty
                Dim PICK_ORDR_LNO As Int16 = row.Item("ORDR_LNO") & String.Empty

                row.Item("ORDR_NO") = ORDR_NO
                row.Item("ORDR_LNO") = ORDR_LNO

                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("ORDR_NO = '" & PICK_ORDR_NO & "' AND ORDR_LNO = " & PICK_ORDR_LNO)
                    rowSOTPICK2.Item("PICK_NO") = PICK_NO
                    rowSOTPICK2.Item("PICK_LNO") = Val(dst.Tables("SOTPICK2").Compute("MAX(PICK_LNO)", "PICK_NO = '" & PICK_NO & "'") & String.Empty) + 1

                    rowSOTPICK2.Item("ORDR_NO") = ORDR_NO
                    rowSOTPICK2.Item("ORDR_LNO") = ORDR_LNO
                Next

                ORDR_LNO += 1
            Next

            ' Need to combine the style / color so we have only one entry for each style / color
            Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"STYLE_CODE", "COLOR_CODE"})

            For Each row As DataRow In tbl.Select("")
                Dim STYLE_CODE As String = row.Item("STYLE_CODE") & String.Empty
                Dim COLOR_CODE As String = row.Item("COLOR_CODE") & String.Empty

                Dim totalQty As Int16 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'"))
                If totalQty > 1 Then
                    If dst.Tables("SOTPICK2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'").Length > 1 Then
                        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")
                            rowSOTPICK2.Item("PICK_QTY") = -1
                        Next
                        dst.Tables("SOTPICK2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")(0).Item("PICK_QTY") = totalQty
                    End If
                End If
            Next

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY <= 0")
                rowSOTPICK2.Delete()
            Next

            dst.Tables("SOTPICK2").AcceptChanges()
            dst.Tables("SOTORDR2").AcceptChanges()
        End If

        EnforceConstraints(True)

        BuildWHTLOCBE()

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE") & String.Empty
            Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE") & String.Empty

            Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Rows.Find(New Object() {STYLE_CODE, COLOR_CODE})
            If rowICTSTYC1 IsNot Nothing Then
                rowSOTORDR2.Item("UPC_CODE") = rowICTSTYC1.Item("UPC_CODE")
            End If
        Next

    End Sub

    Private Sub BuildWHTLOCBE()

        Dim WHSE_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("WHSE_CODE") & String.Empty
        If WHSE_CODE.Length = 0 Then
            WHSE_CODE = "MS"
        End If
        dst.Tables("WHTLOCBE").Rows.Clear()
        dst.Tables("WHTMOVE1").Rows.Clear()
        dst.Tables("WHTMOVE2").Rows.Clear()

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim WHSE_LOC_SHIP As String = rowICTWHSE1.Item("WHSE_LOC_SHP") & String.Empty
        If WHSE_LOC_SHIP.Length = 0 Then
            MessageBox.Show("Unable to determine the Warehouse Location Ship.", "Whse Loc Ship", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rowWHTMOVE1 As DataRow = dst.Tables("WHTMOVE1").NewRow
        Dim WHSE_TRAN_NO As String = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

        rowWHTMOVE1.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
        rowWHTMOVE1.Item("WHSE_TRAN_TYPE") = "M"
        rowWHTMOVE1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
        rowWHTMOVE1.Item("WHSE_CODE") = WHSE_CODE
        rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
        rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP
        rowWHTMOVE1.Item("STATUS") = "U"
        dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)

        Dim WHSE_TRAN_LNO As Int16 = 0
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2"), New String() {"STYLE_CODE", "COLOR_CODE"}).Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

            Dim rowWHTMOVE2 As DataRow = dst.Tables("WHTMOVE2").NewRow
            rowWHTMOVE2.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            WHSE_TRAN_LNO += 1
            rowWHTMOVE2.Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
            rowWHTMOVE2.Item("LOCATION_CODE_FROM") = String.Empty
            rowWHTMOVE2.Item("LOCATION_CODE_TO") = WHSE_LOC_SHIP
            rowWHTMOVE2.Item("BAR_CODE") = "0000000000"
            rowWHTMOVE2.Item("WHSE_TRAN_QTY") = 0
            rowWHTMOVE2.Item("STYLE_CODE") = STYLE_CODE
            rowWHTMOVE2.Item("COLOR_CODE") = COLOR_CODE
            rowWHTMOVE2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE2.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTMOVE2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE2.Item("LAST_DATE") = DATETIME_STAMP
            rowWHTMOVE2.Item("STATUS") = "U"
            rowWHTMOVE2.Item("LOAD_NO_FROM") = String.Empty
            rowWHTMOVE2.Item("LOAD_NO_TO") = String.Empty
            rowWHTMOVE2.Item("BAR_CODE_OTHER") = String.Empty
            dst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

            Dim rowWHTLOCBE As DataRow = dst.Tables("WHTLOCBE").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
            If rowWHTLOCBE Is Nothing Then

                Dim LOCATIONS As String = String.Empty

                'ASCMAIN1.sql = "Select * from WHTLOCB1 where WHSE_CODE = '" & WHSE_CODE & "' and STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and LOCATION_QTY > 0"

                ASCMAIN1.sql = " Select WHTLOCB1.*, WHTLOCM1.LOCATION_USE" _
                    & " from WHTLOCB1, WHTLOCM1 " _
                    & " where WHTLOCB1.WHSE_CODE = '" & WHSE_CODE & "'" _
                    & " and WHTLOCB1.STYLE_CODE = '" & STYLE_CODE & "'" _
                    & " and WHTLOCB1.COLOR_CODE = '" & COLOR_CODE & "'" _
                    & " and WHTLOCB1.LOCATION_QTY > 0" _
                    & " AND WHTLOCB1.LOCATION_CODE = WHTLOCM1.LOCATION_CODE" _
                    & " AND WHTLOCB1.WHSE_CODE = WHTLOCM1.WHSE_CODE" _
                    & " AND WHTLOCM1.LOCATION_USE IN ('A', 'E')"

                For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "LOCATION_USE DESC, LOCATION_QTY DESC")
                    Dim LOCATION_CODE As String = row2.Item("LOCATION_CODE")
                    If rowWHTMOVE2.Item("LOCATION_CODE_FROM") & String.Empty = String.Empty Then
                        rowWHTMOVE2.Item("LOCATION_CODE_FROM") = LOCATION_CODE
                    End If
                    LOCATIONS &= "," & LOCATION_CODE
                Next

                rowWHTLOCBE = dst.Tables("WHTLOCBE").NewRow
                rowWHTLOCBE.Item("STYLE_CODE") = STYLE_CODE
                rowWHTLOCBE.Item("COLOR_CODE") = COLOR_CODE
                rowWHTLOCBE.Item("LOCATIONS") = Mid(LOCATIONS, 2)
                dst.Tables("WHTLOCBE").Rows.Add(rowWHTLOCBE)
            End If

            'dst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
        Next
    End Sub

End Class