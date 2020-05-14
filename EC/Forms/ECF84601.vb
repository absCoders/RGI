Public Class ECF84601
    Private FF As ASFBASE1
    Private STYLE_CODE As String
    Private COLOR_CODE As String
    Private ECOM_NAME As String
    Private ECOM_CODE As String
    Private SQL As New Text.StringBuilder With {.Length = 0}
    Private EDI_PRE As String = ""

    Public Sub New(ByVal frmASFBASE1 As ASFBASE1,
                   ByVal _ECOM_CODE As String,
                   ByVal _STYLE_CODE As String,
                   ByVal _COLOR_CODE As String)
        FF = frmASFBASE1
        STYLE_CODE = _STYLE_CODE
        COLOR_CODE = _COLOR_CODE
        ECOM_CODE = _ECOM_CODE
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim EDI_TP_ID As String = ""
        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            EDI_PRE = "GEN."
        End If
        SQL.Length = 0
        SQL.AppendLine("SELECT C1.ECOM_CODE, C1.ECOM_NAME, C1.EDI_TP_ID")
        SQL.AppendLine("FROM EDTTRPM1 PM, ECTECOM1 C1")
        SQL.AppendLine("WHERE PM.EDI_TP_ID = C1.EDI_TP_ID")
        SQL.AppendLine("AND PM.EDI_DOC_NO = '846'")
        SQL.AppendLine(String.Format("AND C1.ECOM_CODE = '{0}'", ECOM_CODE))
        Dim tblEDTXREF4 As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), "EDTXREF4")
        If tblEDTXREF4.Rows.Count > 0 Then
            EDI_TP_ID = tblEDTXREF4.Rows(0).Item("EDI_TP_ID").ToString & String.Empty
            ECOM_NAME = tblEDTXREF4.Rows(0).Item("ECOM_NAME").ToString & String.Empty
        End If

        SQL.Length = 0
        SQL.AppendLine("SELECT")
        SQL.AppendLine("E2.EDI_OUTBOUND_DOC_NO,")
        SQL.AppendLine("E2.EDI_STYLE,")
        SQL.AppendLine("E2.EDI_COLOR_CODE,")
        SQL.AppendLine("E1.EDI_REPORT_DATE,")
        SQL.AppendLine("E2.EDI_AVAIL_QTY,")
        SQL.AppendLine("DECODE(E2.EDI_MAINT_TYPE_CODE,'001','Active','In-Active') AS EDI_STATUS")
        SQL.AppendLine(String.Format("FROM {0}EDTSYSIH IH, {0}EDT846O1 E1, {0}EDT846O2 E2", EDI_PRE))
        SQL.AppendLine("WHERE IH.COMPANY_CODE = E1.COMPANY_CODE")
        SQL.AppendLine("AND IH.EDI_OUTBOUND_DOC_NO = E1.EDI_OUTBOUND_DOC_NO")
        SQL.AppendLine("AND E1.COMPANY_CODE = E2.COMPANY_CODE")
        SQL.AppendLine("AND E1.EDI_OUTBOUND_DOC_NO = E2.EDI_OUTBOUND_DOC_NO")
        SQL.AppendLine(String.Format("AND E2.EDI_STYLE = '{0}'", STYLE_CODE))
        SQL.AppendLine(String.Format("AND E2.EDI_COLOR_CODE = '{0}'", COLOR_CODE))
        SQL.AppendLine(String.Format("AND IH.EDI_TP_ID = '{0}'", EDI_TP_ID))
        SQL.AppendLine("and E1.EDI_REPORT_DATE > sysdate - 30")
        Dim tblECF84601 As DataTable = ASCDATA1.GetDataTable(SQL.ToString(), "ECF84601")
        grdECF84601.DataSource = tblECF84601
        If tblECF84601.Rows.Count > 0 Then
            grdECF84601.Text = String.Format("846 Transmissions For Partner {0}, Style {1}-{2}", ECOM_NAME, STYLE_CODE, COLOR_CODE)
        Else
            grdECF84601.Text = String.Format("No 846 Transmission Found For Partner {0}, Style {1}-{2}", ECOM_NAME, STYLE_CODE, COLOR_CODE)
        End If

        Sort_grdColumns(grdECF84601, "EDI_REPORT_DATE", False)

        With grdECF84601.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
            For Each grdCol As UltraWinGrid.UltraGridColumn In .Columns
                grdCol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("EDI_REPORT_DATE").Format = "MM/dd/yyyy hh:mm"
            .Columns("EDI_AVAIL_QTY").Format = "###,##0"
        End With

    End Sub

    Private Sub cmdFinished_Click(sender As System.Object, e As System.EventArgs) Handles cmdFinished.Click
        Me.Close()
    End Sub
End Class