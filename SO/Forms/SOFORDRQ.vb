Public Class SOFORDRQ
    Private FF As ASFBASE1
    Private sql As New Text.StringBuilder With {.Length = 0}
    Public CUST_CODE As String = ""
#Region "Form Init"
    Public Sub New(ByVal frmASFBASE1 As ASFBASE1, ByVal in_CUST_CODE As String)
        FF = frmASFBASE1
        CUST_CODE = in_CUST_CODE
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CUST_CODE.Length = 0 Then
            Me.Close()
        End If
        'grdSOTORDRS.DataSource = FF.dst.Tables("SOTORDR1")
    End Sub
#End Region

#Region "Form Controls"
    Private Sub cmdFinished_Click(sender As System.Object, e As System.EventArgs) Handles cmdFinished.Click
        Me.Close()
    End Sub

    Private Sub optQuoteType_ValueChanged(sender As Object, e As EventArgs) Handles optQuoteType.ValueChanged
        If optQuoteType.Value = "M" Then
            setCustPricing()
            pnlMultiPrice.Visible = True
        Else
            pnlMultiPrice.Visible = False
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub setCustPricing()
        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("CUST_CODE,")
        sql.AppendLine("CUST_PRICE_TIER,")
        sql.AppendLine("CUST_DISC_PCT,")
        sql.AppendLine("CUST_PRICE_TIER_PVC,")
        sql.AppendLine("CUST_DISC_PCT_EXTRA")
        sql.AppendLine("FROM ARTCUST1")
        sql.AppendLine("WHERE CUST_CODE = :PARM1")
        Dim tblARTCUST1 As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", CUST_CODE)
        If tblARTCUST1.Rows.Count = 1 Then
            optPRICE_TIER.Value = tblARTCUST1.Rows(0).Item("CUST_PRICE_TIER").ToString & String.Empty
            Select Case optPRICE_TIER.Value
                Case "PC"
                    panEXTRA.Visible = True
                    numDISC_PCT.Visible = False
                    optDISC_PCT_EXTRA.Visible = True
                Case "SP"
                    panEXTRA.Visible = True
                    numDISC_PCT.Visible = True
                    optDISC_PCT_EXTRA.Visible = False
                Case Else
                    panEXTRA.Visible = False
            End Select

            optDISC_PCT_EXTRA.Value = tblARTCUST1.Rows(0).Item("CUST_DISC_PCT_EXTRA").ToString & String.Empty
            numDISC_PCT.Value = tblARTCUST1.Rows(0).Item("CUST_DISC_PCT").ToString & String.Empty
            optPRICE_TIER_PVC.Value = tblARTCUST1.Rows(0).Item("CUST_PRICE_TIER_PVC").ToString & String.Empty
        End If
    End Sub
#End Region
End Class