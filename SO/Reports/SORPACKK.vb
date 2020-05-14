Imports System.Math

Public Class SORPACKK

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

        Check_if_Empty("EDT850T1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = "SORPACKK"
        SUBT = ""
        Generate_Report(RPT, "Kirkland Packing Slip", SUBT, "", "", "", False)
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

            Create_TDA(dst.Tables.Add, "EDT850T1", "*")
            .Tables("EDT850T1").Columns.Add("ORDR_NO", GetType(System.String))

            sql = " SELECT EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ, EDT850T2.EDI_SKU, SOTORDR2.STYLE_CODE REF_NUM, SOTPICK1.PICK_NO, SOTPICK1.LAST_OPER CHECKED_BY,"
            sql &= " NVL(EDI_ITEM_DESC, SOTORDR2.STYLE_DESC) ITEM_DESC, EDT850T2.EDI_TOTAL_QTY, SOTPICK2.PICK_QTY_CONF, SOTORDR2.ORDR_NO, SOTINVH1.INV_DATE"
            sql &= " FROM SOTORDR2, EDT850T2, SOTPICK2, SOTPICK1, SOTINVH1"
            sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO"
            sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ"
            sql &= " AND SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO"
            sql &= " AND SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO"
            sql &= " AND SOTPICK1.PICK_NO = SOTPICK2.PICK_NO"
            sql &= " AND SOTPICK1.INV_NO = SOTINVH1.INV_NO"
            sqlEDT850T2 = sql
            sql &= " AND EDT850T2.EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(dst.Tables.Add, "EDT850T2X", sql, 0, False, "V", 0)

            Create_TDA(dst.Tables.Add("EDT850T5_ST"), "EDT850T5", "*")
            Create_TDA(dst.Tables.Add("EDT850T5_LW"), "EDT850T5", "*")

            dst.Tables("EDT850T5_ST").Columns.Add("FORMATTED_ADDRESS", GetType(System.String))
            dst.Tables("EDT850T5_LW").Columns.Add("FORMATTED_ADDRESS", GetType(System.String))

            ' Create a work table for the form sine there may be multile items
            ' The sections of the report are images of a word document of teh Kirkland's packing slip.

            dst.Tables.Add("PAGE_SECTION")
            With dst.Tables("PAGE_SECTION")
                .Columns.Add("PAGE_NO", GetType(System.Int16))
                .Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))
                .Columns.Add("ORDR_NO", GetType(System.String))
                .Columns.Add("CUST_REF", GetType(System.String))
                .Columns.Add("SHIPPING_DATE", GetType(System.DateTime))
                .Columns.Add("CHECKED_BY", GetType(System.String))
                .Columns.Add("NUM_CARTONS", GetType(System.Int16))
            End With

            dst.Tables.Add("DETAIL_SECTION")
            With dst.Tables("DETAIL_SECTION")
                .Columns.Add("PAGE_NO", GetType(System.Int16))
                .Columns.Add("EDI_DOC_SEQ_NO", GetType(System.String))
                .Columns.Add("ORDR_NO", GetType(System.String))

                For iCtr As Int16 = 1 To numDetailItems
                    .Columns.Add("EDI_DTL_SEQ_" & iCtr, GetType(System.Int16))
                    .Columns.Add("EDI_SKU_" & iCtr, GetType(System.String))
                    .Columns.Add("REF_NUM_" & iCtr, GetType(System.String))
                    .Columns.Add("ITEM_DESC_" & iCtr, GetType(System.String))
                    .Columns.Add("EDI_TOTAL_QTY_" & iCtr, GetType(System.Int16))
                    .Columns.Add("PICK_QTY_CONF_" & iCtr, GetType(System.Int16))
                Next
            End With

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        ' parms(0) should be an Invoice Number
        Dim sqlw As String = String.Empty
        If parms IsNot Nothing AndAlso parms.Length > 0 Then
            sqlw = CStr(parms(0)) & String.Empty
        Else
            sqlw = "'9876543210'"
        End If

        sqlw = sqlw.Replace(Space(1), String.Empty)
        If sqlw.Contains(",") AndAlso Not sqlw.Contains("'") Then
            sqlw = sqlw.Replace(",", "','")
        End If

        If Not sqlw.StartsWith("'") Then
            sqlw = "'" & sqlw
        End If

        If Not sqlw.EndsWith("'") Then
            sqlw &= "'"
        End If

        sql = "Select * from EDT850T1 WHERE EDI_DOC_SEQ_NO in (" & sqlw & ")"
        Fill_Records("EDT850T1", String.Empty, True, sql)

        If dst.Tables("EDT850T1").Rows.Count = 0 Then
            Exit Sub
        End If

        sql = sqlEDT850T2 & " AND EDT850T2.EDI_DOC_SEQ_NO in (" & sqlw & ")"
        Fill_Records("EDT850T2X", String.Empty, True, sql)

        ' As per Maria use BT not LW for the Bill To - 8/9/2018
        sql = "Select * from EDT850T5 WHERE EDI_ADDR_TYPE = 'BT' AND EDI_DOC_SEQ_NO in (" & sqlw & ")"
        Fill_Records("EDT850T5_LW", String.Empty, True, sql)

        sql = "Select * from EDT850T5 WHERE EDI_ADDR_TYPE = 'ST' AND EDI_DOC_SEQ_NO in (" & sqlw & ")"
        Fill_Records("EDT850T5_ST", String.Empty, True, sql)

        For Each tableName As String In New String() {"EDT850T5_LW", "EDT850T5_ST"}
            For Each row As DataRow In dst.Tables(tableName).Select("")
                Dim FORMATTED_ADDRESS As String = String.Empty

                If row.Item("EDI_CUST_NAME_ADR") & String.Empty <> String.Empty Then
                    FORMATTED_ADDRESS &= row.Item("EDI_CUST_NAME_ADR") & String.Empty & Environment.NewLine
                End If

                For Each field As String In New String() {"EDI_ADDRESS1", "EDI_ADDRESS2", "EDI_ADDRESS3"}
                    If row.Item(field) & String.Empty <> String.Empty Then
                        FORMATTED_ADDRESS &= row.Item(field) & String.Empty & Environment.NewLine
                    End If
                Next

                FORMATTED_ADDRESS &= row.Item("EDI_CITY") & ", " & row.Item("EDI_STATE") & "  " & row.Item("EDI_ZIPCODE") & " " & row.Item("EDI_COUNTRY")

                row.Item("FORMATTED_ADDRESS") = FORMATTED_ADDRESS
            Next
        Next

        dst.Tables("PAGE_SECTION").Rows.Clear()
        dst.Tables("DETAIL_SECTION").Rows.Clear()

        Dim PAGE_NO As Int16 = 1
        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim CUST_REF As String = String.Empty
        Dim rowPAGE_SECTION As DataRow = Nothing
        Dim rowDETAIL_SECTION As DataRow = Nothing


        ' REF*OQ*6357038~  Order Number <<< need this on pack slip - suggesting we map this to EDT850T1.EDI_PO_RELEASE_NO
        ' REF*CR*E6357038~ Customer Reference  <<< need this on pack slip - suggesting we map this to EDT850T1.EDI_PROMOTION

        For Each rowEDT850T1 As DataRow In dst.Tables("EDT850T1").Select("", "EDI_DOC_SEQ_NO")
            EDI_DOC_SEQ_NO = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
            ORDR_NO = rowEDT850T1.Item("EDI_PO_RELEASE_NO") & String.Empty
            CUST_REF = rowEDT850T1.Item("EDI_PROMOTION") & String.Empty

            Dim iloop As Int16 = 1
            For Each rowEDT850T2 As DataRow In dst.Tables("EDT850T2X").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_DTL_SEQ")
                If iloop Mod numDetailItems = 1 Then

                    Dim NUM_CARTONS As Int16 = Val(ASCDATA1.GetDataValue("SELECT COUNT(*) FROM SOTCART1 WHERE PICK_NO = '" & dst.Tables("EDT850T2X").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")(0).Item("PICK_NO") & "'") & String.Empty)
                    If NUM_CARTONS = 0 Then
                        NUM_CARTONS = 1
                    End If

                    rowPAGE_SECTION = dst.Tables("PAGE_SECTION").NewRow
                    rowPAGE_SECTION.Item("PAGE_NO") = PAGE_NO
                    rowPAGE_SECTION.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowPAGE_SECTION.Item("ORDR_NO") = ORDR_NO
                    rowPAGE_SECTION.Item("CUST_REF") = CUST_REF
                    rowPAGE_SECTION.Item("SHIPPING_DATE") = rowEDT850T2.Item("INV_DATE") & String.Empty
                    rowPAGE_SECTION.Item("CHECKED_BY") = rowEDT850T2.Item("CHECKED_BY") & String.Empty
                    rowPAGE_SECTION.Item("NUM_CARTONS") = NUM_CARTONS
                    dst.Tables("PAGE_SECTION").Rows.Add(rowPAGE_SECTION)

                    rowDETAIL_SECTION = dst.Tables("DETAIL_SECTION").NewRow
                    rowDETAIL_SECTION.Item("PAGE_NO") = PAGE_NO
                    rowDETAIL_SECTION.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    rowDETAIL_SECTION.Item("ORDR_NO") = ORDR_NO
                    dst.Tables("DETAIL_SECTION").Rows.Add(rowDETAIL_SECTION)

                    PAGE_NO += 1
                    iloop = 1
                End If

                rowDETAIL_SECTION.Item("EDI_DTL_SEQ_" & iloop) = rowEDT850T2.Item("EDI_DTL_SEQ")
                rowDETAIL_SECTION.Item("EDI_SKU_" & iloop) = rowEDT850T2.Item("EDI_SKU")
                rowDETAIL_SECTION.Item("REF_NUM_" & iloop) = rowEDT850T2.Item("REF_NUM")
                rowDETAIL_SECTION.Item("ITEM_DESC_" & iloop) = StrConv(rowEDT850T2.Item("ITEM_DESC") & String.Empty, VbStrConv.ProperCase)
                rowDETAIL_SECTION.Item("EDI_TOTAL_QTY_" & iloop) = rowEDT850T2.Item("EDI_TOTAL_QTY")
                rowDETAIL_SECTION.Item("PICK_QTY_CONF_" & iloop) = rowEDT850T2.Item("PICK_QTY_CONF")
                iloop += 1
            Next
        Next

        EnforceConstraints(True)
    End Sub
End Class