Public Class SORMSCU1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.Get_Dates(ASCMAIN1.CYP).Contains(Now.Date) Then
            ' fill periods with 60 months starting with next month
            Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 1), -60, 0, 0)
        Else
            ' fill periods with 60 months starting with curr month
            Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        End If

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "Y"

        If RWU = "Y" And RYP < ASCMAIN1.CYP Then
            RWU = "N"
        End If

        'Build list of Order Groups for the month for use in FIFO costing function.
        ASCMAIN1.Progress("Building Required Files", "")

        Dim sqlGROUPS As String = "" _
            & "Select DISTINCT SOTORDR1.ORDR_GROUP_NO " & vbCrLf _
            & " FROM SOTINVH1, SOTORDR1" & vbCrLf _
            & " WHERE SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & " AND SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf

        Dim LIVE_DATE As String = "200607"

        ASCMAIN1.sql = "Select DISTINCT STYLE_CODE, COLOR_CODE" & vbCrLf _
         & " FROM SOTORDR2" & vbCrLf _
         & " WHERE ORDR_NO IN " & vbCrLf _
         & "  (" & vbCrLf _
         & "  SELECT ORDR_NO " & vbCrLf _
         & "  FROM SOTORDR1 " & vbCrLf _
         & "  WHERE ORDR_GROUP_NO IN (" & sqlGROUPS & ")" & vbCrLf _
         & "  )" & vbCrLf
        Dim SOTINVHC As String = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "CREATE INDEX I_" & SOTINVHC & "_1 ON " & SOTINVHC & " (STYLE_CODE, COLOR_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = " SELECT I.*" & vbCrLf _
            & " From SOTINVH2 I, " & SOTINVHC & " T" & vbCrLf _
            & " WHERE I.STYLE_CODE = T.STYLE_CODE" & vbCrLf _
            & " AND I.COLOR_CODE = T.COLOR_CODE" & vbCrLf _
            & " AND ORDR_YYYYPP_UPDATED >= '" & LIVE_DATE & "'" & vbCrLf
        Dim SOTINVH2 As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "CREATE INDEX I_" & SOTINVH2 & "_1 ON " & SOTINVH2 & " (INV_TYPE, INV_NO)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "CREATE INDEX I_" & SOTINVH2 & "_2 ON " & SOTINVH2 & " (STYLE_CODE, COLOR_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from SOTINVH1" & vbCrLf _
            & " where INV_NO IN (" & vbCrLf _
            & "    SELECT DISTINCT INV_NO" & vbCrLf _
            & "    from " & SOTINVH2 & ")" & vbCrLf
        Dim SOTINVH1 As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "CREATE INDEX I_" & SOTINVH1 & "_1 ON " & SOTINVH1 & " (INV_TYPE, INV_NO)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Building Masterfile Costs for " & RYP, "")
        Stop ' NEXT LINE WAS REMMED BECAUSE IT WAS DELETE_ME'd
        Dim ICTSCOST As String = "" ' TAC.TACMAIN1.MAKE_COSTMF(Me, RYP, sqlGROUPS, False, True, SOTINVH1, SOTINVH2)

        ''Stop  'This needs to be tested by Wayne.
        'If SOTINVH1 <> "SOTINVH1" And SOTINVH2 <> "SOTINVH2" Then 'JUST MAKE REEEEL SURE!!!
        '    ASCMAIN1.sql = "DROP TABLE " & SOTINVH1 & " PURGE"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "DROP TABLE " & SOTINVH2 & " PURGE"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "DROP TABLE " & SOTINVHC & " PURGE"
        '    ASCDATA1.ExecuteSQL()
        'End If

        ASCMAIN1.Progress("Now Staging Report", "")
        ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" & vbCrLf _
            & ", SOTINVH2.INV_LNO,SOTINVH2.CUST_CODE ,SOTINVH2.STYLE_CODE" & vbCrLf _
            & ", SOTINVH2.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & ", SOTINVH2.ORDR_UNIT_COST, " & ICTSCOST & ".STYLE_COST" & vbCrLf _
            & " from SOTINVH2, ICTSTYL1, " & ICTSCOST & vbCrLf _
            & " where SOTINVH2.STYLE_CODE = " & ICTSCOST & ".STYLE_CODE (+)" & vbCrLf _
            & "   and SOTINVH2.COLOR_CODE = " & ICTSCOST & ".COLOR_CODE (+)" & vbCrLf _
            & "   and SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
            & "   and SOTINVH2.ORDR_UNIT_COST <> " & ICTSCOST & ".STYLE_COST" & vbCrLf _
            & "   and SOTINVH2.ORDR_QTY_SHIP <> 0" & vbCrLf _
            & " group by SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" & vbCrLf _
            & ", SOTINVH2.INV_LNO,SOTINVH2.CUST_CODE ,SOTINVH2.STYLE_CODE" & vbCrLf _
            & ", SOTINVH2.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
            & ", SOTINVH2.ORDR_UNIT_COST, " & ICTSCOST & ".STYLE_COST" & vbCrLf
        Dim SOTINVHX As String = ASCMAIN1.Temp_Table
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTINVHX, "SOTINVHX", 3))

        ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, ORDR_UNIT_COST, STYLE_COST" & vbCrLf _
            & " from " & SOTINVHX & vbCrLf _
            & " group by STYLE_CODE, COLOR_CODE, ORDR_UNIT_COST, STYLE_COST"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTINVHH", 2))

        Dim EXP As String = ""
        With dst.Tables("SOTINVHH").Columns
            EXP = "ISNULL(STYLE_COST,0) - ISNULL(ORDR_UNIT_COST,0)"
            .Add("VARIANCE", GetType(System.Decimal), EXP)
            EXP = "IIF(ISNULL(STYLE_COST,0)=0,100,(ABS(VARIANCE/ISNULL(STYLE_COST,0))*100))"
            .Add("PCT_VAR", GetType(System.Decimal), EXP)
        End With

        ASCMAIN1.Progress("Load Style Data", "")
        ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC from ICTSTYL1" & vbCrLf _
            & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & SOTINVH2 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTYL1", 1))
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = " FOR PERIOD: " & UCase(RYPLegend)
        CR_params.Add("CHKDTL", IIf(Absx1.chkFor("CHKDTL").Checked, "1", "0"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        Dim UPDATE_CNT As Integer = 0

        ASCMAIN1.Progress("Updating Styles", "")

        For Each rowSOTINVHH As DataRow In dst.Tables("SOTINVHH").Select("")
            ASCMAIN1.sql = "Update SOTINVH2" & vbCrLf _
                & " Set ORDR_UNIT_COST = " & rowSOTINVHH.Item("STYLE_COST") & vbCrLf _
                & " WHERE ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
                & " AND STYLE_CODE = '" & rowSOTINVHH.Item("STYLE_CODE") & "'" & vbCrLf _
                & " AND COLOR_CODE = '" & rowSOTINVHH.Item("COLOR_CODE") & "'"
            ASCMAIN1.Progress("Updating Style", rowSOTINVHH.Item("STYLE_CODE"))
            ASCDATA1.ExecuteSQL()
            UPDATE_CNT += 1
        Next

        ASCMAIN1.Progress("Updating Invoices", "")
        For Each rowSOTINVHX As DataRow In dst.Tables("SOTINVHX").Select("")
            ASCMAIN1.sql = "Update SOTINVH1" & vbCrLf _
                & " Set INV_COGS = " & vbCrLf _
                & " (Select SUM(SOTINVH2.ORDR_UNIT_COST * SOTINVH2.ORDR_QTY_SHIP)" & vbCrLf _
                & " from SOTINVH2" & vbCrLf _
                & " where INV_NO = '" & rowSOTINVHX.Item("INV_NO") & "'" & vbCrLf _
                & "   and INV_TYPE = '" & rowSOTINVHX.Item("INV_TYPE") & "')" & vbCrLf _
                & " where INV_NO = '" & rowSOTINVHX.Item("INV_NO") & "'" & vbCrLf _
                & "   and INV_TYPE = '" & rowSOTINVHX.Item("INV_TYPE") & "'"
            ASCDATA1.ExecuteSQL()
        Next

    End Sub
End Class