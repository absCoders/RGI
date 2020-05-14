Public Class EDR850I1

#Region "Declarations"

    Dim WHSE_CODE As String ' Ship From Warehouse Code for Released Orders
    Dim SQL_ins As New Dictionary(Of String, String)

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")

        ASCMAIN1.sql = "Select * from EDT850T1 where NVL(EDI_PROCESS_IND,'0') = '0'"
        Dim tblEDT850I As DataTable = ASCDATA1.GetDataTable

        grdEDT850IX.DataSource = tblEDT850I

        Set_WHSE()
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

        SQL_ins.Clear()
        SQL_ins.Add("CUST_CODE", SQL_in("CUST_CODE"))

        ASCMAIN1.Progress("Setting Up Orders (Demand)", "")

        Dim sql_where As String = ""

        'ASCMAIN1.sql = "Select SOTSREP1.* from SOTSREP1"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))

        ' Main Process

        '    BeginTrans()

        '   EDC850I1.Import_orders(Me, ICTSTDQ1, ICTSTDQ2, WHSE_CODE, blnALLOCATION_ONLY, SOTORDR2)
        '   CommitTrans()

    End Sub

    Public Overrides Sub Print_Report()

        RPT = "EDR850I1"
        ' BASED ON SOTORDR0
        RPT_TITLE = "EDI Import"
        SUBT = "Imported Successfully"
        CR_params.Add("EDI_PROCESS_IND", "1")
        Generate_Report(RPT, RPT_TITLE, SUBT)

        RPT = "EDR850I2"
        ' SHOW EXCEPTIONS
        RPT_TITLE = "EDI Import"
        SUBT = "Failed to Imported Successfully"
        CR_params.Add("EDI_PROCESS_IND", "0")
        Generate_Report(RPT, RPT_TITLE, SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If ASCMAIN1.CLIENT = "VAN" Then
            ' **************** NOT A GOOD IDEA - DO NOT ALLOW THIS
            '    If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and COLUMN_NAME <> 'CUST_CODE' and EXCLUDE = '1'").Length <> 0 Then
            '        EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group & Customer"
            '    End If
            'Else

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

        SQL_ins.Clear()
        SQL_ins.Add("CUST_CODE", "")

        '   Build_Workfile2()
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
End Class