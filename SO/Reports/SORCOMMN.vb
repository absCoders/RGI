Public Class SORCOMMN

#Region "General Declarations"

    Private xRYP0_legend As String = String.Empty
    Private xRYP0 As String = String.Empty

    Private xRYP1_legend As String = String.Empty
    Private xRYP1 As String = String.Empty

    Private SOTCOMMS As String = String.Empty
    Private querySOTCOMMS As String = String.Empty
    Private alternateSOTCOMMSworktable As String = String.Empty

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        xRYP0_legend = Absx1.cmbFor("RYP0").Value
        xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)

        xRYP1_legend = Absx1.cmbFor("RYP1").Value
        xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)

        If xRYP0 > xRYP1 Then
            MessageBox.Show("Start Period is greater than Ending Period.", "Report", MessageBoxButtons.OK, MessageBoxIcon.Error)
            RWU &= "0"
            xErrMsg = "No Eligible Records"
            Exit Sub
        End If

        SetSubTitle()

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTCOMMS")
    End Sub

    Public Overrides Sub Print_Report()
        'CR_params.Add("SUBT", SUBT)
        'CR_params.Add("SUPPRESS_DETAIL", "0")
        'Generate_Report("SORCOMMN", , SUBT)

        RPT_TITLE = "Commission by Sales Rep / Customer"
        CR_params.Add("SUBT", SUBT)
        CR_params.Add("SUPPRESS_DETAIL", "0")
        Generate_Report("SORCOMMN", RPT_TITLE, SUBT)

        RPT_TITLE = "Commission by Customer / Sales Rep"
        CR_params.Add("SUBT", SUBT)
        CR_params.Add("SUPPRESS_DETAIL", "0")
        Generate_Report("SORCNYAC", RPT_TITLE, SUBT)

        RPT_TITLE = "Commission Summary by Sales Rep"
        CR_params.Add("SUBT", SUBT)
        CR_params.Add("SUPPRESS_DETAIL", "0")
        Generate_Report("SORCNYAT", RPT_TITLE, SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If Absx1.cmbFor("RYP0").Value & "" = "" Then
            EMsg &= vbCr & "You must Specify a Starting Period"
        End If
    End Sub

    Overrides Sub Update_Record()

        ASCDATA1.ExecuteSQL("Insert Into SOTCOMMS Select * from " & SOTCOMMS)

        ' Create Accounts payable records
        For Each SREP_CODE As String In ASCDATA1.SelectDistinct("SOTCOMMS", New String() {"SREP_CODE"}).Rows
            Dim totalCommission As Decimal = Val(dst.Tables("SOTCOMMS").Compute("SUM(SREP_COMM_AMT)", "SREP_CODE = '" & SREP_CODE & "'") & String.Empty)
        Next

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        With dst

            ' SOTCOMMS = work table with the commissions, clone of datatable SOTCOMMS
            SOTCOMMS = TAC.SOCCOMMS.NYAGCommissionsWorktable
            querySOTCOMMS = "Select * from " & TAC.SOCCOMMS.NYAGCommissionsWorktable
            Create_TDA(.Tables.Add, "SOTCOMMS", "*", -1, False)

            'Create_TDA(.Tables.Add, "SOTSREP1", "*")
            Create_TDA(.Tables.Add, "SOTCTYP2", "SELECT 'XX' COMM_CALC_BY, LPAD(35, ' ') COMM_CALC_BY_DESC FROM DUAL WHERE ROWNUM < 1", 0, False)

            Dim field_infos() As Reflection.FieldInfo = GetType(TAC.SOCCOMMS.NYACommissionCalcTypes).GetFields
             For Each field_info As Reflection.FieldInfo In field_infos
                If field_info.IsLiteral Then
                    dst.Tables("SOTCTYP2").Rows.Add(New Object() {CType(field_info.GetValue(Nothing), Integer), field_info.Name})
                 End If
            Next field_info

            Create_TDA(.Tables.Add, "SOTSREP1", "SELECT SREP_CODE, SREP_NAME FROM SOTSREP1")
            Create_TDA(.Tables.Add, "ARTCUST1", "SELECT CUST_CODE, CUST_NAME FROM ARTCUST1")
            Create_TDA(.Tables.Add, "ICTSGRP1", "SELECT STYLE_GROUP_CODE, STYLE_GROUP_DESC FROM ICTSGRP1")
            Create_TDA(.Tables.Add, "SOTSDIV1", "SELECT SALES_DIVISION_CODE, SALES_DIVISION_NAME FROM SOTSDIV1")

            sql = " SELECT SREP_CODE"
            sql &= ", SUM(DECODE(COMM_CALC_BY, 2, SREP_COMM_AMT,0)) DIVISION"
            sql &= ", SUM(DECODE(COMM_CALC_BY, 4, SREP_COMM_AMT,0)) STYLE_GROUP "
            sql &= ", SUM(DECODE(COMM_CALC_BY, 5, SREP_COMM_AMT,0)) CUST_OVERRIDE "
            sql &= ", SUM(DECODE(COMM_CALC_BY, 6, SREP_COMM_AMT,0)) STYLE_GROUP_OVERRIDE"
            sql &= " FROM " & SOTCOMMS
            sql &= " group by SREP_CODE"
            Create_TDA(.Tables.Add, "SOTCOMMT", sql, 0, False, "", 1)

        End With


        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim workTableName As String = String.Empty

        ' Period to load
        If parms.Length > 0 Then
            If parms(0).ToString.Contains("-") Then
                xRYP0 = parms(0).ToString.Split("-")(0).Trim
                xRYP0_legend = ASCDATA1.GetDataValue("SELECT LEGEND FROM GLTPARM2 WHERE OPS_YYYYPP = '" & xRYP0 & "'") & String.Empty

                xRYP1 = parms(1).ToString.Split("-")(0).Trim
                xRYP1_legend = ASCDATA1.GetDataValue("SELECT LEGEND FROM GLTPARM2 WHERE OPS_YYYYPP = '" & xRYP1 & "'") & String.Empty
            Else
                xRYP0 = parms(0)
                xRYP0_legend = ASCDATA1.GetDataValue("SELECT LEGEND FROM GLTPARM2 WHERE OPS_YYYYPP = '" & xRYP0 & "'") & String.Empty
            End If

            SetSubTitle()
        End If

        ' alternate data table to use, most likey a temp oracle table
        alternateSOTCOMMSworktable = String.Empty
        If parms.Length > 1 Then
            If parms(0).ToString.Trim.Length > 0 Then
                alternateSOTCOMMSworktable = parms(1)
            End If
         End If

        EnforceConstraints(False)

        Dim sqlw As String = String.Empty
        If parms.Length = 0 Then
            sqlw &= SQL_in("SALES_DIVISION_CODE")
            sqlw &= SQL_in("STYLE_GROUP_CODE")
            sqlw &= SQL_in("SREP_CODE")
            sqlw &= SQL_in("CUST_CODE")
            sqlw &= SQL_in("INV_NO")
            'sqlw &= SQL_in("COMM_CALC_BY")
        End If

        If sqlw.Length > 0 Then
            RWU = "N"
            sqlw = " WHERE " & sqlw.Substring(4).Trim
        End If

        If alternateSOTCOMMSworktable.Length > 0 Then
            Fill_Records("SOTCOMMS", String.Empty, True, "SELECT * FROM " & alternateSOTCOMMSworktable & sqlw)
            workTableName = alternateSOTCOMMSworktable
        Else
            If xRYP0.Length > 0 AndAlso xRYP0 <> xRYP1 Then
                TAC.SOCCOMMS.GetNYAGCommissions(xRYP0 & "-" & xRYP1)
            Else
                TAC.SOCCOMMS.GetNYAGCommissions(xRYP0)
            End If

            Fill_Records("SOTCOMMS", String.Empty, True, querySOTCOMMS & sqlw)
            workTableName = SOTCOMMS
        End If

        ASCMAIN1.sql = "SELECT SREP_CODE, SREP_NAME FROM SOTSREP1"
        Fill_Records("SOTSREP1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT STYLE_GROUP_CODE, STYLE_GROUP_DESC FROM ICTSGRP1"
        Fill_Records("ICTSGRP1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT SALES_DIVISION_CODE, SALES_DIVISION_NAME FROM SOTSDIV1"
        Fill_Records("SOTSDIV1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT CUST_CODE, CUST_NAME FROM ARTCUST1"
        Fill_Records("ARTCUST1", String.Empty, True, ASCMAIN1.sql)

        Fill_Records("SOTCOMMT")

        ' Show Only summary for sales reps in the main datatable
        For Each rowSOTCOMMT As DataRow In dst.Tables("SOTCOMMT").Select("")
            If dst.Tables("SOTCOMMS").Select("SREP_CODE = '" & rowSOTCOMMT.Item("SREP_CODE") & "'").Length = 0 Then
                rowSOTCOMMT.Delete()
            End If
        Next
        dst.Tables("SOTCOMMT").AcceptChanges()

        EnforceConstraints(True)

    End Sub

    Private Sub SetSubTitle()
        If xRYP0 = xRYP1 Then
            SUBT = "Commissions Posted in " & xRYP0_legend
        Else
            SUBT = "Total Commissions Posted between " & xRYP0_legend & " and " & xRYP1_legend
        End If

        If xRYP0 <> ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) OrElse xRYP0 <> xRYP1 Then
            RWU = "N"
        Else
            If Val(ASCDATA1.GetDataValue("SELECT COUNT(*) FROM SOTCOMMS WHERE OPS_YYYYPP = '" & xRYP0 & "'") & String.Empty) = 0 Then
                SUBT = "Proposed Commissions for " & xRYP0_legend
                RWU = "Y"
            Else
                RWU = "N"
            End If
        End If
    End Sub

End Class