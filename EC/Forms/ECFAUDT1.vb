
Public Class ECFAUDT1
    Private FF As ASFBASE1
    Private _STYLE_CODE As String = ""
    Private _ECOM_CODE As String = ""

    Public Sub New(ByVal FF As ASFBASE1, ByVal STYLE_CODE As String, ByVal ECOM_CODE As String)
        _STYLE_CODE = STYLE_CODE
        _ECOM_CODE = ECOM_CODE
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        grdASTAUDT1.DataSource = frmASFBASE1.dst.Tables.Item("ASTAUDT1")
        For Each COL As UltraWinGrid.UltraGridColumn In grdASTAUDT1.DisplayLayout.Bands(0).Columns
            COL.Hidden = True
        Next
        With grdASTAUDT1.DisplayLayout.Bands(0).Columns
            .Item("KEY_VALUE").Hidden = False
            .Item("KEY_VALUE2").Hidden = False
            .Item("COLUMN_NAME").Hidden = False
            .Item("USER_ID").Hidden = False
            .Item("INIT_DATE").Hidden = False
            .Item("OLD_VALUE").Hidden = False
            .Item("NEW_VALUE").Hidden = False
            .Item("INIT_DATE").Format = "MM/dd/yy hh:mm:ss"
        End With
        Sort_grdColumns(grdASTAUDT1, "INIT_DATE")
        grdASTAUDT1.DisplayLayout.Bands(0).Columns.Item("INIT_DATE").Format = "MM/dd/yy hh:mm:ss"
        grdASTAUDT1.Text = String.Format("Audit Trail For {0} - {1}", _STYLE_CODE, _ECOM_CODE)
        Dim dvw As DataView = DirectCast(grdASTAUDT1.DataSource, DataTable).DefaultView
        Dim filter As String = String.Format("KEY_VALUE = '{0}' AND KEY_VALUE2 = '{1}'", _STYLE_CODE, _ECOM_CODE)
        dvw.RowFilter = String.Format(filter, "INIT_DATE")
    End Sub

    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        Me.Close()
    End Sub
End Class