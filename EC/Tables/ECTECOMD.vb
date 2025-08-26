Imports Infragistics.Win.UltraWinGrid

Public Class ECTECOMD

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTSVIAW", "*", 1)
            .Tables("SOTSVIAW").Columns.Add("SHIP_VIA_DESC", GetType(System.String), "SHIP_VIA_CODE")

            Create_TDA(.Tables.Add, "ICTSTYCW", "*", 1)
            .Tables("ICTSTYCW").Columns.Add("STYLE_DESC", GetType(System.String))
            .Tables("ICTSTYCW").Columns.Add("COLOR_DESC", GetType(System.String))
            .Tables("ICTSTYCW").Columns.Add("SIZE_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            ASCMAIN1.sql = "SELECT * FROM SOTSVIA1 WHERE CARRIER_CODE IN ('UPS', 'FEDEX') AND CARRIER_PROD_CODE IS NOT NULL AND NVL(SHIP_VIA_STATUS, 'A') = 'A'"
            Fill_Records("SOTSVIA1", String.Empty, True, ASCMAIN1.sql)
        End With

        grdSOTSVIAW.DataSource = dst.Tables("SOTSVIAW")
        ASCMAIN1.Add_Value_List(grdSOTSVIAW, "SHIP_VIA_DESC", "SELECT SHIP_VIA_CODE, SHIP_VIA_DESC FROM SOTSVIA1")
        Create_Summary(grdSOTSVIAW, "SHIP_VIA_CODE", "Count")

        grdICTSTYCW.DataSource = dst.Tables("ICTSTYCW")
        ASCMAIN1.Add_Value_List(grdICTSTYCW, "ECOM_PRODUCT_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "P:Pending"}, 0)
        Create_Summary(grdICTSTYCW, "STYLE_CODE", "Count")
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                grdSOTSVIAW.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdICTSTYCW.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        Update_Record_TDA("SOTSVIAW", "ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'")
        Update_Record_TDA("ICTSTYCW", "ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)

        'ASCMAIN1.sql = $"SELECT SOTSVIAW.*, SOTSVIA1.SHIP_VIA_DESC 
        '                    FROM SOTSVIAW, SOTSVIA1
        '                    WHERE SOTSVIAW.ECOM_CODE = '{Absx1.txtFor("ECOM_CODE").Text}'
        '                    AND SOTSVIAW.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)"
        'Fill_Records("SOTSVIAW", "", True, ASCMAIN1.sql)
        Fill_Records("SOTSVIAW", Absx1.txtFor("ECOM_CODE").Text)

        ASCMAIN1.sql = $"SELECT ICTSTYCW.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYC3.SIZE_CODE
                            FROM ICTSTYCW, ICTSTYL1, ICTCOLR1, ICTSTYC3
                            WHERE ICTSTYCW.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)
                            AND ICTSTYCW.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)
                            AND ICTSTYCW.ECOM_CODE = '{Absx1.txtFor("ECOM_CODE").Text}'
                            AND ICTSTYCW.STYLE_CODE = ICTSTYC3.STYLE_CODE (+)
                            AND ICTSTYCW.COLOR_CODE = ICTSTYC3.COLOR_CODE (+)
                            AND ICTSTYCW.SIZE_INDEX = ICTSTYC3.SIZE_INDEX (+)"
        Fill_Records("ICTSTYCW", "", True, ASCMAIN1.sql)

        EnforceConstraints(True)

        grdSOTSVIAW.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdICTSTYCW.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        Sort_grdColumns(grdSOTSVIAW, "SHIP_VIA_CODE")
        Sort_grdColumns(grdICTSTYCW, "STYLE_CODE,COLOR_CODE,SIZE_INDEX")

        If EntryMode = "New" Then
            optECOM_STATUS.Value = "A"
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTSVIAW").Rows.Clear()
            dst.Tables("ICTSTYCW").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, ByRef Optional sql_where As String = "", ByRef Optional Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "ECOM_CUST_ADDR_CODE"
                sql_where = $"CUST_CODE = '{txtECOM_CUST_CODE.Text}' AND CUST_ADDR_TYPE = 'MK'"

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdICTSTYCW_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTSTYCW.BeforeRowUpdate
        e.Row.Cells("ECOM_CODE").Value = Absx1.txtFor("ECOM_CODE").Text

        Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & String.Empty
        Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & String.Empty
        Dim SIZE_INDEX As String = e.Row.Cells("SIZE_INDEX").Value & String.Empty

        Dim drICTSTYC3 As DataRow = LookUp("ICTSTYC3", {STYLE_CODE, COLOR_CODE, SIZE_INDEX})
        If drICTSTYC3 Is Nothing Then
            MessageBox.Show("Invalid Style, Color, Size combination.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        e.Row.Cells("SIZE_CODE").Value = drICTSTYC3.Item("SIZE_CODE") & String.Empty
    End Sub

    Private Sub grdICTSTYCW_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTSTYCW.ClickCellButton

        Dim sql_where As String = String.Empty

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                sql_where = "STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYC3)"
                grdClickCellButton(grdICTSTYCW, sql_where, False, e.Cell.Column.Key, e.Cell.Column.Key)

            Case "COLOR_CODE"
                sql_where = $"COLOR_CODE IN (SELECT COLOR_CODE FROM ICTSTYC3 WHERE STYLE_CODE = '{e.Cell.Row.Cells("STYLE_CODE").Value & String.Empty}')"
                grdClickCellButton(grdICTSTYCW, sql_where, False, e.Cell.Column.Key, e.Cell.Column.Key)

            Case "SIZE_INDEX"
                Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value & String.Empty
                Dim COLOR_CODE As String = e.Cell.Row.Cells("COLOR_CODE").Value & String.Empty
                ASCMAIN1.sql = $"SELECT SIZE_INDEX, SIZE_CODE FROM ICTSTYC3 WHERE STYLE_CODE = :PARM1 AND COLOR_CODE = :PARM2"

                With ASCMAIN1.CodeSelector
                    .SQL = ASCMAIN1.sql
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "Select Size Index"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    .Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                    .ParamTypes = "VV"
                    .Params = {STYLE_CODE, COLOR_CODE}
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()

                If ASCMAIN1.CodeSelector.Selections = 1 Then
                    Dim SIZE_INDEX As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("SIZE_INDEX") & String.Empty
                    e.Cell.Row.Cells("SIZE_INDEX").Value = SIZE_INDEX
                End If
        End Select
    End Sub

    Private Sub grdICTSTYCW_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdICTSTYCW.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value & String.Empty
                If STYLE_CODE.Length > 0 Then
                    Dim drICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If drICTSTYL1 Is Nothing Then
                        e.Cell.Row.Cells("STYLE_DESC").Value = String.Empty
                    Else
                        e.Cell.Row.Cells("STYLE_DESC").Value = drICTSTYL1.Item("STYLE_DESC") & String.Empty
                    End If
                End If

            Case "COLOR_CODE"
                Dim COLOR_CODE As String = e.Cell.Row.Cells("COLOR_CODE").Value & String.Empty
                If COLOR_CODE.Length > 0 Then
                    Dim drICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                    If drICTCOLR1 Is Nothing Then
                        e.Cell.Row.Cells("COLOR_DESC").Value = String.Empty
                    Else
                        e.Cell.Row.Cells("COLOR_DESC").Value = drICTCOLR1.Item("COLOR_DESC") & String.Empty
                    End If
                End If

            Case "SIZE_INDEX"
                Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value & String.Empty
                Dim COLOR_CODE As String = e.Cell.Row.Cells("COLOR_CODE").Value & String.Empty
                Dim SIZE_INDEX As String = e.Cell.Row.Cells("SIZE_INDEX").Value & String.Empty

                If STYLE_CODE.Length > 0 AndAlso COLOR_CODE.Length > 0 AndAlso SIZE_INDEX.Length > 0 Then
                    Dim drICTSTYC3 As DataRow = LookUp("ICTSTYC3", {STYLE_CODE, COLOR_CODE, SIZE_INDEX})
                    If drICTSTYC3 Is Nothing Then
                        e.Cell.Row.Cells("SIZE_CODE").Value = String.Empty
                    Else
                        e.Cell.Row.Cells("SIZE_CODE").Value = drICTSTYC3.Item("SIZE_CODE") & String.Empty
                    End If
                Else
                    e.Cell.Row.Cells("SIZE_CODE").Value = String.Empty
                End If

        End Select
    End Sub


    Private Sub grdSOTSVIAW_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTSVIAW.BeforeRowUpdate
        e.Row.Cells("ECOM_CODE").Value = Absx1.txtFor("ECOM_CODE").Text
        Dim SHIP_VIA_CODE As String = e.Row.Cells("SHIP_VIA_CODE").Value & String.Empty
        Dim ECOM_SHIP_VIA_CODE As String = e.Row.Cells("ECOM_SHIP_VIA_CODE").Value & String.Empty

        Dim drSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
        If drSOTSVIA1 Is Nothing Then
            MessageBox.Show("Invalid Ship Via Code", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        e.Row.Cells("SHIP_VIA_DESC").Value = drSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
    End Sub

    Private Sub grdSOTSVIAW_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSOTSVIAW.ClickCellButton

        Dim sql_where As String = String.Empty

        Select Case e.Cell.Column.Key
            Case "SHIP_VIA_CODE"
                sql_where = "CARRIER_CODE IN ('UPS', 'FEDEX') AND CARRIER_PROD_CODE IS NOT NULL AND NVL(SHIP_VIA_STATUS, 'A') = 'A'"
                grdClickCellButton(grdSOTSVIAW, sql_where)

        End Select
    End Sub

    Private Sub grdSOTSVIAW_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdSOTSVIAW.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SHIP_VIA_CODE"
                Dim SHIP_VIA_CODE As String = e.Cell.Row.Cells("SHIP_VIA_CODE").Value & String.Empty
                If SHIP_VIA_CODE.Length > 0 Then
                    Dim drSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                    If drSOTSVIA1 Is Nothing Then
                        e.Cell.Row.Cells("SHIP_VIA_DESC").Value = String.Empty
                    Else
                        e.Cell.Row.Cells("SHIP_VIA_DESC").Value = drSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    End If
                End If
        End Select

    End Sub

#End Region

End Class