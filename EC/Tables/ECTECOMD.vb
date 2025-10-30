Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinEditors

Public Class ECTECOMD

    'BEGIN DECLARE CURSOR C1 IS
    'SELECT ICTSTYCW.STYLE_CODE, ICTSTYC3.SIZE_CODE, ICTSTYL1.STYLE_CODE_PLM,
    'SUBSTR(ICTSTYCW.STYLE_CODE, 1, LENGTH(ICTSTYCW.STYLE_CODE) - LENGTH(ICTSTYC3.SIZE_CODE)) AS STYLE_CODE_PLM_NEW
    'FROM ICTSTYCW, ICTSTYL1, ICTCOLR1, ICTSTYC3
    'WHERE ICTSTYCW.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)
    'AND ICTSTYCW.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)
    'AND ICTSTYCW.ECOM_CODE = 'SHOPIFY'
    'AND ICTSTYCW.STYLE_CODE = ICTSTYC3.STYLE_CODE (+)
    'AND ICTSTYCW.COLOR_CODE = ICTSTYC3.COLOR_CODE (+)
    'AND ICTSTYCW.SIZE_INDEX = ICTSTYC3.SIZE_INDEX (+);
    'BEGIN FOR R1 IN C1 LOOP
    '    UPDATE ICTSTYL1 SET STYLE_CODE_PLM = R1.STYLE_CODE_PLM_NEW WHERE STYLE_CODE = R1.STYLE_CODE AND STYLE_CODE_PLM IS NULL;
    'END LOOP; END; END;

    'INSERT INTO ICTSTYCW
    'SELECT 'SHOPIFY' ECOM_CODE, STYLE_CODE, COLOR_CODE, SIZE_INDEX, NULL ECOM_PRODUCT_ID, NULL ECOM_VARIANT_ID, 
    'NULL ECOM_INV_VARIANT_ID, 'A' ECOM_PRODUCT_STATUS, SYSDATE ECOM_PRODUCT_STATUS_DATE, NULL ECOM_PRODUCT_LAST_UPDATED,
    'NULL WEB_DESCRIPTION, NULL BODY_HTML
    'FROM ICTSTYC4 WHERE UPC_CODE IS NOT NULL
    'AND STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYL1 WHERE SALES_DIVISION_CODE = '30')
    'and STYLE_CODE not in (select STYLE_CODE from ICTSTYCW)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTSVIAW", "*", 1)
            .Tables("SOTSVIAW").Columns.Add("SHIP_VIA_DESC", GetType(System.String), "SHIP_VIA_CODE")

            Create_TDA(.Tables.Add, "ICTSTYCW", "*", 1)
            .Tables("ICTSTYCW").Columns.Add("STYLE_CODE_PLM", GetType(System.String))
            .Tables("ICTSTYCW").Columns.Add("STYLE_DESC", GetType(System.String))
            .Tables("ICTSTYCW").Columns.Add("COLOR_DESC", GetType(System.String))
            .Tables("ICTSTYCW").Columns.Add("SIZE_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            ASCMAIN1.sql = "SELECT * FROM SOTSVIA1 WHERE CARRIER_PROD_CODE IS NOT NULL AND NVL(SHIP_VIA_STATUS, 'A') = 'A'"
            Fill_Records("SOTSVIA1", String.Empty, True, ASCMAIN1.sql)

            .Tables.Add("ICTSTYCW_PLM")
            With .Tables("ICTSTYCW_PLM")
                .Columns.Add("ECOM_CODE", GetType(System.String))
                .Columns.Add("STYLE_CODE_PLM", GetType(System.String))
                .Columns.Add("WEB_DESCRIPTION", GetType(System.String))
                .Columns.Add("BODY_HTML", GetType(System.String))
                .Columns.Add("EXISTS", GetType(System.String))

                .PrimaryKey = New DataColumn() { .Columns("ECOM_CODE"), .Columns("STYLE_CODE_PLM")}
            End With

            Create_TDA(.Tables.Add, "WHTPKGMW", "*", 1)
        End With

        grdSOTSVIAW.DataSource = dst.Tables("SOTSVIAW")
        ASCMAIN1.Add_Value_List(grdSOTSVIAW, "SHIP_VIA_DESC", "SELECT SHIP_VIA_CODE, SHIP_VIA_DESC FROM SOTSVIA1")
        Create_Summary(grdSOTSVIAW, "SHIP_VIA_CODE", "Count")

        grdICTSTYCW.DataSource = dst.Tables("ICTSTYCW")
        ASCMAIN1.Add_Value_List(grdICTSTYCW, "ECOM_PRODUCT_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "P:Pending"}, 0)
        Create_Summary(grdICTSTYCW, "STYLE_CODE", "Count")

        grdICTSTYCW_PLM.DataSource = dst.Tables("ICTSTYCW_PLM")
        Create_Summary(grdICTSTYCW_PLM, "STYLE_CODE_PLM", "Count")

        grdWHTPKGMW.DataSource = dst.Tables("WHTPKGMW")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                grdSOTSVIAW.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdICTSTYCW.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdWHTPKGMW.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                Dim lstPLMs As New List(Of String)
                For Each drICTSTYCW As DataRow In dst.Tables("ICTSTYCW").Select
                    If drICTSTYCW.Item("STYLE_CODE_PLM") & String.Empty = String.Empty Then
                        lstPLMs.Add(drICTSTYCW.Item("STYLE_CODE"))
                    End If
                Next

                If lstPLMs.Count > 0 Then
                    Dim zMsg As String = "The following styles do not have a PLM value and will not be sent to the Web."
                    zMsg &= Environment.NewLine & Environment.NewLine
                    zMsg &= String.Join(Environment.NewLine, lstPLMs.ToArray)

                    MessageBox.Show(zMsg, "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""

        For Each dr As DataRow In dst.Tables("ICTSTYCW_PLM").Select("")
            Dim ECOM_CODE As String = dr.Item("ECOM_CODE") & String.Empty
            Dim STYLE_CODE_PLM As String = dr.Item("STYLE_CODE_PLM") & String.Empty

            For Each drICTSTYCW As DataRow In dst.Tables("ICTSTYCW").Select($"ECOM_CODE = '{ECOM_CODE}' AND STYLE_CODE_PLM = '{STYLE_CODE_PLM}'")
                drICTSTYCW.Item("WEB_DESCRIPTION") = dr.Item("WEB_DESCRIPTION")
                drICTSTYCW.Item("BODY_HTML") = dr.Item("BODY_HTML")
            Next
        Next

        Update_Record_TDA("SOTSVIAW", "ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'")
        Update_Record_TDA("ICTSTYCW", "ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'")
        Update_Record_TDA("WHTPKGMW", "ECOM_CODE = '" & Absx1.txtFor("ECOM_CODE").Text & "'")

    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)

        Fill_Records("SOTSVIAW", Absx1.txtFor("ECOM_CODE").Text)
        Fill_Records("WHTPKGMW", Absx1.txtFor("ECOM_CODE").Text)

        Try
            ASCMAIN1.sql = "INSERT INTO ICTSTYCW
                                SELECT 'SHOPIFY' ECOM_CODE, STYLE_CODE, COLOR_CODE, SIZE_INDEX, NULL ECOM_PRODUCT_ID, NULL ECOM_VARIANT_ID, 
                                NULL ECOM_INV_VARIANT_ID, 'A' ECOM_PRODUCT_STATUS, SYSDATE ECOM_PRODUCT_STATUS_DATE, NULL ECOM_PRODUCT_LAST_UPDATED,
                                NULL WEB_DESCRIPTION, NULL BODY_HTML
                                FROM ICTSTYC4 WHERE UPC_CODE IS NOT NULL
                                AND STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYL1 WHERE SALES_DIVISION_CODE = '30')
                                AND STYLE_CODE NOT IN (select STYLE_CODE from ICTSTYCW)"

            Dim numRows As Int32 = ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            If numRows > 0 Then
                'MessageBox.Show($"{numRows} styles were added to the list of styles for the web. You must click Update to save these styles.", "Show Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception

        End Try

        ASCMAIN1.sql = $"SELECT ICTSTYCW.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYC3.SIZE_CODE, ICTSTYL1.STYLE_CODE_PLM
                            FROM ICTSTYCW, ICTSTYL1, ICTCOLR1, ICTSTYC3
                            WHERE ICTSTYCW.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)
                            AND ICTSTYCW.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)
                            AND ICTSTYCW.ECOM_CODE = '{Absx1.txtFor("ECOM_CODE").Text}'
                            AND ICTSTYCW.STYLE_CODE = ICTSTYC3.STYLE_CODE (+)
                            AND ICTSTYCW.COLOR_CODE = ICTSTYC3.COLOR_CODE (+)
                            AND ICTSTYCW.SIZE_INDEX = ICTSTYC3.SIZE_INDEX (+)"
        Fill_Records("ICTSTYCW", "", True, ASCMAIN1.sql)

        Dim tblPLM As DataTable = ASCDATA1.SelectDistinct(dst.Tables("ICTSTYCW"), {"ECOM_CODE", "STYLE_CODE_PLM"})

        dst.Tables("ICTSTYCW_PLM").Rows.Clear()
        For Each drPLM As DataRow In tblPLM.Select()
            Dim ECOM_CODE As String = drPLM.Item("ECOM_CODE") & String.Empty
            Dim STYLE_CODE_PLM As String = drPLM.Item("STYLE_CODE_PLM") & String.Empty

            If dst.Tables("ICTSTYCW").Select($"ECOM_CODE = '{ECOM_CODE}' and STYLE_CODE_PLM = '{STYLE_CODE_PLM}'").Length > 0 Then
                Dim drLookup As DataRow = dst.Tables("ICTSTYCW").Select($"ECOM_CODE = '{ECOM_CODE}' and STYLE_CODE_PLM = '{STYLE_CODE_PLM}'")(0)
                Dim WEB_DESCRIPTION As String = drLookup.Item("WEB_DESCRIPTION") & String.Empty
                Dim BODY_HTML As String = drLookup.Item("BODY_HTML") & String.Empty
                dst.Tables("ICTSTYCW_PLM").Rows.Add({ECOM_CODE, STYLE_CODE_PLM, WEB_DESCRIPTION, BODY_HTML, "1"})
            End If
        Next
        dst.Tables("ICTSTYCW_PLM").AcceptChanges()

        EnforceConstraints(True)

        Sort_grdColumns(grdICTSTYCW_PLM, "STYLE_CODE_PLM")
        Sort_grdColumns(grdSOTSVIAW, "SHIP_VIA_CODE")
        Sort_grdColumns(grdICTSTYCW, "STYLE_CODE,COLOR_CODE,SIZE_INDEX")
        grdSOTSVIAW.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        If EntryMode = "New" Then
            optECOM_STATUS.Value = "A"
        End If
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTSVIAW").Rows.Clear()
            dst.Tables("ICTSTYCW_PLM").Rows.Clear()
            dst.Tables("ICTSTYCW").Rows.Clear()
            dst.Tables("WHTPKGMW").Rows.Clear()
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

            Case "ECOM_DEFAULT_SHIP_VIA "
                sql_where = $"CARRIER_PROD_CODE IS NOT NULL"

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYCW_PLM, "S", "Show Filter")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case = "grdECTECOM1_PARTNER"

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case ""
        End Select

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdPLM_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTSTYCW_PLM.AfterRowActivate
        Try
            WebBrowser1.DocumentText = grdICTSTYCW_PLM.ActiveRow.Cells("BODY_HTML").Value & String.Empty
        Catch ex As Exception
            WebBrowser1.DocumentText = String.Empty
        End Try

        Dim ECOM_CODE As String = grdICTSTYCW_PLM.ActiveRow.Cells("ECOM_CODE").Value & String.Empty
        Dim STYLE_CODE_PLM As String = grdICTSTYCW_PLM.ActiveRow.Cells("STYLE_CODE_PLM").Value & String.Empty

        Dim dvw As DataView = DirectCast(grdICTSTYCW.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"ECOM_CODE = '{ECOM_CODE}' and STYLE_CODE_PLM = '{STYLE_CODE_PLM}'"
        dvw.Sort = "STYLE_CODE,COLOR_CODE,SIZE_INDEX"

    End Sub

    Private Sub grdICTSTYCW_PLM_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICTSTYCW_PLM.AfterRowUpdate


        Dim STYLE_CODE_PLM As String = e.Row.Cells("STYLE_CODE_PLM").Value & String.Empty
        If dst.Tables("ICTSTYCW").Select($"STYLE_CODE_PLM = '{STYLE_CODE_PLM}'").Length = 0 Then

            ASCMAIN1.sql = $"SELECT '{Absx1.txtFor("ECOM_CODE").Text}' ECOM_CODE, ICTSTYL1.STYLE_CODE, ICTSTYC3.COLOR_CODE, ICTSTYC3.SIZE_INDEX,
                            NULL ECOM_PRODUCT_ID, NULL ECOM_VARIANT_ID, NULL ECOM_INV_VARIANT_ID,
                            'A' ECOM_PRODUCT_STATUS, SYSDATE ECOM_PRODUCT_STATUS_DATE, NULL ECOM_PRODUCT_LAST_UPDATED,
                            NULL WEB_DESCRIPTION, NULL BODY_HTML, ICTSTYL1.STYLE_CODE_PLM,
                            ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC, ICTSTYC3.SIZE_CODE
                            FROM ICTSTYL1, ICTSTYC3, ICTCOLR1
                            WHERE ICTSTYL1.STYLE_CODE = ICTSTYC3.STYLE_CODE
                            AND NVL(ICTSTYL1.STYLE_STATUS, 'A') = 'A'
                            AND ICTSTYL1.STYLE_CODE_PLM = '{STYLE_CODE_PLM}'
                            AND ICTSTYC3.COLOR_CODE = ICTCOLR1.COLOR_CODE (+)"

            Fill_Records("ICTSTYCW", String.Empty, False, ASCMAIN1.sql)
        End If

    End Sub

    Private Sub grdICTSTYCW_PLM_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdICTSTYCW_PLM.BeforeRowActivate

        If e.Row.Cells("EXISTS").Value & String.Empty = "1" Then
            grdICTSTYCW_PLM.DisplayLayout.Bands(0).Columns("STYLE_CODE_PLM").CellActivation = Activation.NoEdit
        Else
            grdICTSTYCW_PLM.DisplayLayout.Bands(0).Columns("STYLE_CODE_PLM").CellActivation = Activation.AllowEdit
        End If

    End Sub

    Private Sub grdICTSTYCW_PLM_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTSTYCW_PLM.BeforeRowUpdate

        Try
            Dim STYLE_CODE_PLM As String = (e.Row.Cells("STYLE_CODE_PLM").Value & String.Empty).ToString.Trim
            If STYLE_CODE_PLM.Length = 0 Then
                MessageBox.Show("PLM Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If

            e.Row.Cells("STYLE_CODE_PLM").Value = STYLE_CODE_PLM

            ASCMAIN1.sql = $"SELECT ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC
                            FROM ICTSTYL1, ICTSTYC3
                            WHERE ICTSTYL1.STYLE_CODE = ICTSTYC3.STYLE_CODE
                            AND NVL(ICTSTYL1.STYLE_STATUS, 'A') = 'A'
                            AND ICTSTYL1.STYLE_CODE_PLM = '{STYLE_CODE_PLM}'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {STYLE_CODE_PLM})
            If tbl.Rows.Count = 0 Then
                MessageBox.Show("PLM Code does not contain any Active Styles.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If

            e.Row.Cells("ECOM_CODE").Value = Absx1.txtFor("ECOM_CODE").Text
            If e.Row.Cells("WEB_DESCRIPTION").Value & String.Empty = String.Empty Then
                e.Row.Cells("WEB_DESCRIPTION").Value = StrConv(tbl.Rows(0).Item("STYLE_DESC"), VbStrConv.ProperCase)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End Try
    End Sub

    Private Sub grdICTSTYCW_PLM_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTSTYCW_PLM.ClickCellButton

        Dim sql_where As String = String.Empty

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE_PLM"

                If e.Cell.Row.Cells("EXISTS").Value & String.Empty = "1" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = $"SELECT STYLE_CODE_PLM PLM_CODE, MAX(STYLE_DESC) DESCRIPTION, COUNT(*) NUM_STYLES
                                    FROM ICTSTYL1 
                                    WHERE STYLE_CODE_PLM IS NOT NULL 
                                    AND NVL(STYLE_STATUS, 'A') = 'A'
                                    GROUP BY STYLE_CODE_PLM 
                                    ORDER BY STYLE_CODE_PLM"

                With ASCMAIN1.CodeSelector
                    .SQL = ASCMAIN1.sql
                    .MultipleSelections = False
                    .PreviouslySelectedCodes0 = ""
                    .Caption = "Select Style PLM"
                    .TABLE_NAME = ""
                    .VIEW_NAME = ""
                    .VIEW_DESC = ""
                    .COLUMN_NAME = ""
                    .COLUMN_PREKEYs = New Dictionary(Of String, String)
                    .Custom_sql_where = ""
                    .tblASTVIEW1 = New DataTable
                    .ParamTypes = "VV"
                    .Params = Nothing
                End With
                Dim F As New ASFCODE1
                F.ShowDialog()

                If ASCMAIN1.CodeSelector.Selections = 1 Then
                    Dim STYLE_CODE_PLM As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("PLM_CODE") & String.Empty
                    If dst.Tables("ICTSTYCW_PLM").Select($"STYLE_CODE_PLM = '{STYLE_CODE_PLM}'").Length > 0 Then
                        MessageBox.Show("The Selected PLM already exists.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                    e.Cell.Row.Cells("STYLE_CODE_PLM").Value = STYLE_CODE_PLM
                End If
        End Select
    End Sub


    Private Sub grdICTSTYCW_BeforeRowActivate(sender As Object, e As RowEventArgs) Handles grdICTSTYCW.BeforeRowActivate
        ' Existing rows cannot be modified. They can only be Updated

        'If 1 = 1 Then
        '    Exit Sub
        'End If

        If e.Row.IsAddRow Then
            grdICTSTYCW.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = Activation.AllowEdit
            grdICTSTYCW.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = Activation.AllowEdit
            grdICTSTYCW.DisplayLayout.Bands(0).Columns("SIZE_INDEX").CellActivation = Activation.AllowEdit
        Else
            Dim ECOM_CODE As String = e.Row.Cells("ECOM_CODE").Value & String.Empty
            Dim STYLE_CODE As String = e.Row.Cells("STYLE_CODE").Value & String.Empty
            Dim COLOR_CODE As String = e.Row.Cells("COLOR_CODE").Value & String.Empty
            Dim SIZE_INDEX As String = e.Row.Cells("SIZE_INDEX").Value & String.Empty

            Dim row As DataRow = dst.Tables("ICTSTYCW").Rows.Find({ECOM_CODE, STYLE_CODE, COLOR_CODE, SIZE_INDEX})
            If row Is Nothing OrElse row.RowState = DataRowState.Added Then
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = Activation.AllowEdit
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = Activation.AllowEdit
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("SIZE_INDEX").CellActivation = Activation.AllowEdit
            Else
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = Activation.NoEdit
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("COLOR_CODE").CellActivation = Activation.NoEdit
                grdICTSTYCW.DisplayLayout.Bands(0).Columns("SIZE_INDEX").CellActivation = Activation.NoEdit
            End If
        End If
    End Sub

    Private Sub grdICTSTYCW_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTSTYCW.BeforeRowUpdate

        'If 1 = 1 Then
        '    Exit Sub
        'End If

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

        'If 1 = 1 Then
        '    Exit Sub
        'End If


        Dim sql_where As String = String.Empty

        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"
                sql_where = "STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYC3)"
                grdClickCellButton(grdICTSTYCW, sql_where, False, e.Cell.Column.Key, e.Cell.Column.Key)

            Case "COLOR_CODE"
                sql_where = $"COLOR_CODE IN (SELECT COLOR_CODE FROM ICTSTYC3 WHERE STYLE_CODE = '{e.Cell.Row.Cells("STYLE_CODE").Value & String.Empty}')"
                grdClickCellButton(grdICTSTYCW, sql_where, False, e.Cell.Column.Key, e.Cell.Column.Key)

            Case "SIZE_INDEX"

                If e.Cell.Row.Cells("STYLE_CODE").Activation <> Activation.AllowEdit Then
                    Exit Sub
                End If

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

        'If 1 = 1 Then
        '    Exit Sub
        'End If

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

    End Sub

    Private Sub grdSOTSVIAW_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSOTSVIAW.ClickCellButton

        Dim sql_where As String = String.Empty

        Select Case e.Cell.Column.Key
            Case "SHIP_VIA_CODE"
                sql_where = "CARRIER_PROD_CODE IS NOT NULL AND NVL(SHIP_VIA_STATUS, 'A') = 'A'"
                grdClickCellButton(grdSOTSVIAW, sql_where)
        End Select
    End Sub


    Private Sub grdWHTPKGMW_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdWHTPKGMW.BeforeRowUpdate
        e.Row.Cells("ECOM_CODE").Value = Absx1.txtFor("ECOM_CODE").Text
        Dim PKG_L As Decimal = Val(e.Row.Cells("PKG_L").Value & "")
        Dim PKG_W As Decimal = Val(e.Row.Cells("PKG_W").Value & "")
        Dim PKG_H As Decimal = Val(e.Row.Cells("PKG_H").Value & "")

        ' Sort the values by length, width, height
        If PKG_L <= 0 OrElse PKG_W <= 0 OrElse PKG_H < 0 Then
            MessageBox.Show("All dimensions must be greater than 0", "Update", MessageBoxButtons.OK)
            e.Cancel = True
            Exit Sub
        End If

        Dim dimList As New List(Of Decimal)
        dimList.Add(PKG_L)
        dimList.Add(PKG_W)
        dimList.Add(PKG_H)
        dimList.Sort()
        PKG_L = dimList(2)
        PKG_W = dimList(1)
        PKG_H = dimList(0)

        e.Row.Cells("PKG_L").Value = PKG_L
        e.Row.Cells("PKG_W").Value = PKG_W
        e.Row.Cells("PKG_H").Value = PKG_H
    End Sub

#End Region

End Class