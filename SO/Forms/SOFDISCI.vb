Imports ABSolution
Imports System.Drawing

Public Class SOFDISCI

    Private SOTORDRX As String = String.Empty
    Private ICTSTAT2 As String = String.Empty
    Private SO_PARM_DEF_PICK_WHSE As String = String.Empty
    Private releasedInventoryShortages As Boolean = False
    Private cancelItems As Boolean = False

    Private Const LetterEmailedOrPrinted As String = "2"
    Private Const LettersWithoutDetails As String = "3"
    Private AttachmentFolder As String = String.Empty


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            .Tables("SOTORDR1").Columns.Add("SELECTED", GetType(System.String))
            .Tables("SOTORDR1").Columns.Add("AS_PARM_KEY", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTORDR2", "*", , , , , "ORDR_LINE_CANC")
            .Tables("SOTORDR2").Columns.Add("SELECTED", GetType(System.String))
            .Tables("SOTORDR2").Columns.Add("EXT_PRICE", GetType(System.Decimal), "ORDR_QTY_OPEN * ORDR_UNIT_PRICE")
            .Tables("SOTORDR2").Columns.Add("CANCEL_QTY", GetType(System.Int32), "ISNULL(ORDR_QTY_OPEN, 0) - ISNULL(ORDR_QTY_ALLO, 0)")
            .Tables("SOTORDR2").Columns.Add("EXT_DISC", GetType(System.Decimal), "ISNULL(ORDR_UNIT_PRICE, 0) * ISNULL(CANCEL_QTY, 0)")

            .Relations.Add("SOTORDR1_SOTORDR2", dst.Tables("SOTORDR1").Columns("ORDR_NO"), dst.Tables("SOTORDR2").Columns("ORDR_NO"))

            ' Discontinued Totals
            .Tables("SOTORDR1").Columns.Add("DISC_COUNT", GetType(System.Int32), "COUNT(CHILD.ORDR_LNO)")
            .Tables("SOTORDR1").Columns.Add("DISC_TOTAL", GetType(System.Decimal), "SUM(CHILD.EXT_DISC)")

            .Tables("SOTORDR2").Columns.Add("CUST_CODE", GetType(System.String), "PARENT.CUST_CODE")
            .Tables("SOTORDR2").Columns.Add("CUST_NAME", GetType(System.String), "PARENT.CUST_NAME")

            Create_TDA(.Tables.Add, "SOTORDR5_BT", "SELECT * FROM SOTORDR5", 0, False, "", 2)
            Create_TDA(.Tables.Add, "SOTORDR5_ST", "SELECT * FROM SOTORDR5", 0, False, "", 2)
            Create_TDA(.Tables.Add, "ARTCUST1", "*", 0, False, "", 1)

            Create_TDA(.Tables.Add, "ICTSTYL1", "*")
            .Tables("ICTSTYL1").Columns.Add("STYLE_CLASS_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "ICTSTAT2", "*")
            Create_TDA(.Tables.Add, "SOTORDRC", "*")

            Create_TDA(.Tables.Add, "ASTATTA2", "*")
            Create_Relation("ICTSTYL1", "ICTSTAT2", "STYLE_CODE")

            With .Tables.Add("SOTINVP0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("AR_PARM_DUNS_NO")
                .Columns.Add("ADDRESS_LINE")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With

            Get_PARM("ARTPARM1")
            Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
            With ROWs("ARTPARM1")
                rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
                rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
                rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
                rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                        & .Item("AR_PARM_REMIT_STATE") & " " _
                        & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                        & .Item("AR_PARM_REMIT_COUNTRY")
                If .Item("AR_PARM_REMIT_PHONE") & "" <> "" And .Item("AR_PARM_REMIT_FAX") & "" <> "" Then
                    rowSOTINVP0.Item("REMIT3") = "" _
                        & "  Tel " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_PHONE")) _
                        & ", Fax " & ASCMAIN1.FormatTel(.Item("AR_PARM_REMIT_FAX"))
                End If
                rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
                If 1 = 1 Then
                    rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") & vbCrLf & .Item("AR_PARM_REMIT_MESSAGE_EXPORT")
                End If
                rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
                rowSOTINVP0.Item("ADDRESS_LINE") = "" _
                    & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ADDR1") _
                    & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_CITY") _
                    & ", " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_STATE") _
                    & " " & ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_ZIP_CODE") _
                    & IIf(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "" <> "" _
                      And ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & "" <> "", "" _
                          & ", Tel " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE") & "") _
                          & ", Fax " & ASCMAIN1.FormatTel(ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX") & ""), "")
            End With
            rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
            dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)

            dst.Tables.Add("STATS")
            With dst.Tables("STATS")
                .Columns.Add("SEQ", GetType(System.Int32))
                .Columns.Add("KEY", GetType(System.String))
                .Columns.Add("VALUE", GetType(System.Int32))
            End With

            Create_TDA(.Tables.Add, "SOTSREP1", "*")
            Fill_Records("SOTSREP1", String.Empty, True, "SELECT * FROM SOTSREP1")

        End With

        grdSOTORDRX.DataSource = dst.Tables("SOTORDR1")
        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRX, "DISC_COUNT", "Sum")
        Create_Summary(grdSOTORDRX, "DISC_TOTAL", "Sum")
        Create_Summary(grdSOTORDRX, "SELECTED", "Sum")

        grdSOTORDRX2.DataSource = dst.Tables("SOTORDR2")
        Create_Summary(grdSOTORDRX2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDRX2, "ORDR_QTY_OPEN", "Sum")
        Create_Summary(grdSOTORDRX2, "ORDR_QTY_ALLO", "Sum")
        Create_Summary(grdSOTORDRX2, "CANCEL_QTY", "Sum")
        Create_Summary(grdSOTORDRX2, "EXT_PRICE", "Sum")
        Create_Summary(grdSOTORDRX2, "EXT_DISC", "Sum")

        grdSOTORDRC.DataSource = dst.Tables("SOTORDRC")
        Create_Summary(grdSOTORDRC, "ORDR_NO", "Count")
        Create_Summary(grdSOTORDRC, "EMAILED_TO_CUST", "Sum")
        Create_Summary(grdSOTORDRC, "EMAILED_TO_SREP", "Sum")
        Create_Summary(grdSOTORDRC, "PRINTED", "Sum")

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        Create_Summary(grdICTSTYL1, "STYLE_CODE", "Count")
        For Each field As String In Split("WHSE_QTY_ON_HAND,WHSE_QTY_ON_ORDER,WHSE_QTY_TRAN,WHSE_QTY_OPEN,WHSE_QTY_PICK,WHSE_QTY_ALLO", ",")
            Create_Summary(grdICTSTYL1, field, "Sum", "ICTSTYL1_ICTSTAT2")
        Next
        grdICTSTYL1.DisplayLayout.Bands(1).SummaryFooterCaption = "Totals"

        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = False
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False

        grdStats.DataSource = dst.Tables("STATS")

        AttachmentFolder = ASCMAIN1.Folders("Attach")

        If ASCMAIN1.Running_in_VS Then
            If ASCMAIN1.USER_ID = "edz" Then
                AttachmentFolder = "R:\RGI\Attach\RGI"
            End If
        End If

        If AttachmentFolder.Length > 0 Then
            If Not My.Computer.FileSystem.DirectoryExists(AttachmentFolder) Then
                AttachmentFolder = String.Empty
            Else
                If Not AttachmentFolder.EndsWith("\") Then
                    AttachmentFolder &= "\"
                End If
            End If
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Generate Letters"
                releasedInventoryShortages = False

            Case "Cancel"
                If MessageBox.Show("Do you want to Cancel any changes?",
                                    "Cancel Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                    MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If


            Case "Save Changes"
                If Not releasedInventoryShortages Then

                    If AttachmentFolder.Length = 0 Then
                        EMsg &= vbCr & "AS_PARM_ARCHIVE_FOLDER is not assigned to a valid directory."
                        Exit Select
                    End If

                    If MessageBox.Show("Do you want to Print the Letters for the selected Sales orders?",
                                       "Print Letters", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                       MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub

                    End If

                Else
                    If MessageBox.Show("Do you want to Save your Changes? This menu choice does not cancel items off the Sales Orders.",
                                       "Save Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                       MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub

                    End If
                End If

            Case "View Shortages"
                releasedInventoryShortages = True

            Case "Update Sales Orders"
                If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                If MessageBox.Show("Are you sure you want to cancel the selected items on the selected sales orders?", "Update Sales Orders", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Generate Letters", "View Shortages"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)

            Case "Save Changes", "Update Sales Orders"
                cancelItems = (eItemKey = "Update Sales Orders")
                Update_Record()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View Shortages").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Save Changes").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Update Sales Orders").Settings.Enabled = not_iScreenMode

                .Groups("Screen Control").Items("Generate Letters").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

                If releasedInventoryShortages OrElse not_iScreenMode = 1 Then
                    .Groups("Screen Control").Items("Save Changes").Text = "Save Changes"
                    .Groups("Screen Control").Items("Update Sales Orders").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Save Changes").Text = "Print Letters"
                    .Groups("Screen Control").Items("Update Sales Orders").Settings.Enabled = DefaultableBoolean.False
                End If

                If Not tf Then
                    If Val(ASCDATA1.GetDataValue("select count(*) from sotordrc where EMAILED_TO_CUST = '1' OR PRINTED = '1'") & String.Empty) > 0 Then
                        .Groups("Screen Control").Items("View Shortages").Settings.Enabled = DefaultableBoolean.False
                    End If
                End If

                .Groups("STATS").Visible = tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            Fill_Stats()
        Else
            Clear_Record()
        End If

        splSOTORDRX.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SOTORDR1").Rows.Clear()
        dst.Tables("SOTORDR2").Rows.Clear()
        dst.Tables("ICTSTYL1").Rows.Clear()
        dst.Tables("ICTSTAT2").Rows.Clear()
        dst.Tables("ASTATTA2").Rows.Clear()
        dst.Tables("SOTORDRC").Rows.Clear()

        cancelItems = False
        EnforceConstraints(True)

        releasedInventoryShortages = True

    End Sub

    Private Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        Dim sql As String = String.Empty

        sql = "SELECT DISTINCT SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDRC.SELECTED, SOTORDRC.SELECTED_DTL"
        sql &= " FROM SOTORDR2, SOTORDRC"
        sql &= " Where SOTORDR2.ORDR_NO = SOTORDRC.ORDR_NO  AND SOTORDR2.ORDR_LNO = SOTORDRC.ORDR_LNO"
        sql &= " and SOTORDR2.STYLE_CODE = SOTORDRC.STYLE_CODE AND SOTORDR2.COLOR_CODE = SOTORDRC.COLOR_CODE"

        If Not releasedInventoryShortages Then
            sql &= " AND NVL(SOTORDR2.ORDR_LINE_CANC, '0') <> '1'"
        End If

        If SOTORDRX.Length > 0 Then
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTORDRX)
            ASCDATA1.ExecuteSQL("INSERT INTO " & SOTORDRX & " " & sql)
        Else
            SOTORDRX = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRX & " Add Primary Key(ORDR_NO, ORDR_LNO)")
        End If

        Fill_Records("SOTORDR1", String.Empty, True, "SELECT DISTINCT SOTORDR1.*, SOTORDRX.SELECTED, 'Z' AS_PARM_KEY FROM SOTORDR1, " & SOTORDRX & "  SOTORDRX WHERE SOTORDR1.ORDR_NO = SOTORDRX.ORDR_NO")
        Fill_Records("SOTORDR5_BT", String.Empty, True, "SELECT * FROM SOTORDR5 WHERE CUST_ADDR_TYPE = 'BT' AND ORDR_NO IN (SELECT DISTINCT ORDR_NO FROM " & SOTORDRX & ")")
        Fill_Records("SOTORDR5_ST", String.Empty, True, "SELECT * FROM SOTORDR5 WHERE CUST_ADDR_TYPE = 'ST' AND ORDR_NO IN (SELECT DISTINCT ORDR_NO FROM " & SOTORDRX & ")")

        Dim SOTINVP1 As String = ASCMAIN1.Temp_Table("SELECT DISTINCT CUST_CODE FROM SOTORDR1 WHERE ORDR_NO IN (SELECT DISTINCT ORDR_NO FROM " & SOTORDRX & ")")
        sql = "SELECT * FROM ARTCUST1"
        sql &= " WHERE CUST_CODE IN "
        sql &= " (Select Distinct CUST_CODE from " & SOTINVP1 & " union Select Distinct CUST_BILL_TO_CUST from " & SOTINVP1 & ")"
        Fill_Records("ARTCUST1", String.Empty, True, sql)

        sql = "SELECT DISTINCT SOTORDR2.*, SOTORDRX.SELECTED_DTL SELECTED"
        sql &= " FROM SOTORDR2, " & SOTORDRX & " SOTORDRX"
        sql &= " WHERE  SOTORDR2.ORDR_NO = SOTORDRX.ORDR_NO"
        sql &= " AND SOTORDR2.ORDR_LNO = SOTORDRX.ORDR_LNO"
        Fill_Records("SOTORDR2", String.Empty, True, sql)

        ' Convert Nulls
        For Each row As DataRow In dst.Tables("SOTORDR2").Select("")
            row.Item("ORDR_QTY_ALLO") = Val(row.Item("ORDR_QTY_ALLO") & String.Empty)
        Next

        sql = "SELECT WHSE_CODE, STYLE_CODE, COLOR_CODE FROM ICTSTAT2 WHERE (WHSE_CODE, STYLE_CODE, COLOR_CODE) IN "
        sql &= " (SELECT SOTORDR1.WHSE_CODE, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE"
        sql &= " FROM SOTORDR1, SOTORDR2, " & SOTORDRX & "  SOTORDRX"
        sql &= " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
        sql &= " AND SOTORDR2.ORDR_NO = SOTORDRX.ORDR_NO"
        sql &= " AND SOTORDR2.ORDR_LNO = SOTORDRX.ORDR_LNO)"

        If ICTSTAT2.Length > 0 Then
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ICTSTAT2)
            ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTAT2 & " " & sql)
        Else
            ICTSTAT2 = ASCMAIN1.Temp_Table(sql)
        End If

        Fill_Records("ICTSTYL1", String.Empty, True, "SELECT ICTSTYL1.*, ICTCLAS1.STYLE_CLASS_DESC" _
                     & " FROM ICTSTYL1, ICTCLAS1" _
                     & " WHERE ICTSTYL1.STYLE_CLASS_CODE = ICTCLAS1.STYLE_CLASS_CODE (+) " _
                     & " AND STYLE_CODE IN (SELECT STYLE_CODE FROM " & ICTSTAT2 & ")")

        Fill_Records("ICTSTAT2", String.Empty, True, "SELECT * FROM ICTSTAT2 WHERE (WHSE_CODE, STYLE_CODE, COLOR_CODE) IN (SELECT WHSE_CODE, STYLE_CODE, COLOR_CODE FROM " & ICTSTAT2 & ")")

        For Each field As String In New String() {"EMAILED_TO_CUST", "EMAILED_TO_SREP", "PRINTED"}
            ASCDATA1.ExecuteSQL("UPDATE SOTORDRC SET " & field & " = '0' WHERE " & field & " IS NULL")
        Next
        Fill_Records("SOTORDRC", String.Empty, True, "SELECT * FROM SOTORDRC")

        EnforceConstraints(True)

        Sort_grdColumns(grdSOTORDRX, "CUST_CODE,ORDR_NO")
        Sort_grdColumns(grdSOTORDRX, "ORDR_LNO", False, 1)
        Sort_grdColumns(grdICTSTYL1, "STYLE_CODE")


        If Not releasedInventoryShortages Then
            grdSOTORDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdSOTORDRX2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdSOTORDRX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSOTORDRX2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        If Not releasedInventoryShortages Then
            If Not UpdateCancelOrderData() Then Exit Sub
            If Not CancelItemQtys() Then Exit Sub
        Else
            If Not ClearNoDetailSalesOrders() Then Exit Sub
            Fill_Stats()

            If Not ProcessCustomerEmails() Then Exit Sub

            If Not ProcessPrintedletters() Then Exit Sub

            If Not ProcessSalesRepEmails() Then Exit Sub
            Fill_Stats()

            ' Do this at the end so if there are errors the orders can be reprocessed
            ' The flags on SOTORDRC (EMAILED_TO_CUST, EMAILED_TO_SREP, PRINTED) prevent duplicate processing.
            ' SOTORDR2 is only updated when all the above has been processed successfully.
            Try
                BeginTrans()
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_LINE_CANC = '1'")
                    rowSOTORDR2.Item("ORDR_LINE_CANC") = "0"
                Next
                Update_Record_TDA("SOTORDR2")
                CommitTrans()

            Catch ex As Exception
                Rollback()
                MessageBox.Show($"Error processing Emails Updating SOTORDR2: {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
        End If

        Fill_Stats()
        ASCMAIN1.Progress("", "")
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdICTSTYL1, "SSSPBBBPBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Style Status Inquiry", "Select All Class", "De-Select All Class")

        Load_Popup_Menu(grdSOTORDRX, "SSSPBBBB", "Show Filter", "Show GroupBox", "Show Pins",
                        "Select All", "De-Select All", "Sales Order Inquiry")

        Load_Popup_Menu(grdSOTORDRX2, "SSSPBBBB", "Show Filter", "Show GroupBox", "Show Pins",
                "Select All", "De-Select All", "Style Status Inquiry", "Sales Order Inquiry")

        Load_Popup_Menu(grdSOTORDRC, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name
            'Case "grdSOTORDRX"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Item Status Inquiry"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTORDRX"
                    ' tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = grd.ActiveRow.Band.Index = 0

                    'tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = grd.ActiveRow.Band.Index = 0

                    'tlb_btn = DirectCast(tlb_pop.Tools("Style Status Inquiry"), UltraWinToolbars.ButtonTool)
                    'tlb_btn.SharedProps.Visible = grd.ActiveRow.Band.Index = 1
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")

            Case "Select All Class", "De-Select All Class"

                If Not releasedInventoryShortages Then
                    MessageBox.Show("You are not permitted to perform this action.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim selectedValue As String = "1"
                If e.Tool.Key = "De-Select All Class" Then
                    selectedValue = "0"
                End If

                Dim STYLE_CLASS_CODE As String = String.Empty
                Dim STYLE_CLASS_DESC As String = String.Empty

                Select Case grd.Name
                    Case grdICTSTYL1.Name

                        Select Case grdICTSTYL1.ActiveRow.Band.Key
                            Case grdICTSTYL1.DisplayLayout.Bands(0).Key
                                STYLE_CLASS_CODE = grdICTSTYL1.ActiveRow.Cells("STYLE_CLASS_CODE").Value
                                STYLE_CLASS_DESC = grdICTSTYL1.ActiveRow.Cells("STYLE_CLASS_DESC").Value

                            Case grdICTSTYL1.DisplayLayout.Bands(1).Key
                                STYLE_CLASS_CODE = grdICTSTYL1.ActiveRow.ParentRow.Cells("STYLE_CLASS_CODE").Value
                                STYLE_CLASS_DESC = grdICTSTYL1.ActiveRow.ParentRow.Cells("STYLE_CLASS_DESC").Value
                        End Select

                End Select

                Dim msg As String = "Do you want to '" & e.Tool.Key & "' " & STYLE_CLASS_CODE & "/" & STYLE_CLASS_DESC & "?"
                If MessageBox.Show(msg, e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("STYLE_CLASS_CODE = '" & STYLE_CLASS_CODE & "'")
                    Dim STYLE_CODE As String = rowICTSTYL1.Item("STYLE_CODE") & String.Empty

                    For Each row As DataRow In dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                        row.Item("SELECTED") = selectedValue
                    Next
                Next

                dst.Tables("SOTORDR2").AcceptChanges()

                MessageBox.Show(e.Tool.Key & " Complete", e.Tool.Key, MessageBoxButtons.OK)

            Case "Select All", "De-Select All"

                If Not releasedInventoryShortages Then
                    MessageBox.Show("You are not permitted to perform this action.", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim selectedValue As String = "1"
                If e.Tool.Key = "De-Select All" Then
                    selectedValue = "0"
                End If

                Select Case grd.Name

                    Case grdSOTORDRX.Name
                        For Each row As DataRow In dst.Tables("SOTORDR1").Select("")
                            row.Item("SELECTED") = selectedValue
                        Next
                        dst.Tables("SOTORDR1").AcceptChanges()

                    Case grdSOTORDRX2.Name

                        Dim sql As String = String.Empty

                        If tabGroupBy.SelectedTab.Key = "By Sales Order" Then
                            sql = "ORDR_NO = '" & grdSOTORDRX2.ActiveRow.Cells("ORDR_NO").Value & "'"
                        Else
                            'grdICTSTYL1
                            sql = "STYLE_CODE = '" & grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value & "'"
                        End If

                        For Each row As DataRow In dst.Tables("SOTORDR2").Select(sql)
                            row.Item("SELECTED") = selectedValue
                        Next
                        dst.Tables("SOTORDR2").AcceptChanges()

                    Case grdICTSTYL1.Name
                        Dim STYLE_CODE As String = String.Empty
                        Dim COLOR_CODE As String = String.Empty
                        Select Case grdICTSTYL1.ActiveRow.Band.Key
                            Case grdICTSTYL1.DisplayLayout.Bands(0).Key
                                STYLE_CODE = grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value
                                For Each row As DataRow In dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                                    row.Item("SELECTED") = selectedValue
                                Next

                            Case grdICTSTYL1.DisplayLayout.Bands(1).Key
                                STYLE_CODE = grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value
                                COLOR_CODE = grdICTSTYL1.ActiveRow.Cells("COLOR_CODE").Value

                                For Each row As DataRow In dst.Tables("SOTORDR2").Select("STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "'")
                                    row.Item("SELECTED") = selectedValue
                                Next
                        End Select

                        dst.Tables("SOTORDR2").AcceptChanges()
                End Select

                MessageBox.Show(e.Tool.Key & " Complete", e.Tool.Key, MessageBoxButtons.OK)

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdSOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDRX.AfterRowActivate

        If grdSOTORDRX2.DataSource Is Nothing Then
            Exit Sub
        End If

        If tabGroupBy.SelectedTab.Key <> "By Sales Order" Then
            Exit Sub
        End If

        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = True

        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = False
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False

        Dim view As DataView = DirectCast(grdSOTORDRX2.DataSource, DataTable).DefaultView

        If grdSOTORDRX.ActiveRow IsNot Nothing Then
            view.RowFilter = "ORDR_NO = '" & grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value & "'"
            grdSOTORDRX2.Text = "Details for Sales Order: " & grdSOTORDRX.ActiveRow.Cells("ORDR_NO").Value
        Else
            view.RowFilter = "ORDR_NO = ''"
            grdSOTORDRX2.Text = "No Sales Order Selected"
        End If

        view.Sort = "ORDR_LNO"

        grdSOTORDRX2.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortSingle
    End Sub

    Private Sub grdICTSTYL1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSTYL1.AfterRowActivate

        If grdSOTORDRX2.DataSource Is Nothing Then
            Exit Sub
        End If

        If tabGroupBy.SelectedTab.Key <> "By Style" Then
            Exit Sub
        End If

        If grdICTSTYL1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If Not grdICTSTYL1.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = False
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_CODE").Hidden = False
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("CUST_NAME").Hidden = False

        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("STYLE_CODE").Hidden = True
        grdSOTORDRX2.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = True


        Dim view As DataView = DirectCast(grdSOTORDRX2.DataSource, DataTable).DefaultView

        If grdICTSTYL1.ActiveRow Is Nothing Then
            view.RowFilter = "STYLE_CODE = ''"
            grdSOTORDRX2.Text = "No Style Selected"
            Exit Sub
        End If

        If grdICTSTYL1.ActiveRow.Band.Key = "ICTSTYL1" Then
            view.RowFilter = "STYLE_CODE = '" & grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value & "'"
            grdSOTORDRX2.Text = "Details for Style: " & grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value & " - " & grdICTSTYL1.ActiveRow.Cells("STYLE_DESC").Value
            grdSOTORDRX2.DisplayLayout.Bands(0).Columns("COLOR_CODE").Hidden = False
        Else
            view.RowFilter = "STYLE_CODE = '" & grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value & "' AND COLOR_CODE = '" & grdICTSTYL1.ActiveRow.Cells("COLOR_CODE").Value & "'"
            grdSOTORDRX2.Text = "Details for Style/Color: " & grdICTSTYL1.ActiveRow.Cells("STYLE_CODE").Value & " / " & grdICTSTYL1.ActiveRow.Cells("COLOR_CODE").Value
        End If

        view.Sort = "ORDR_NO, ORDR_LNO"

        grdSOTORDRX2.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortSingle
    End Sub

    Private Sub tabGroupBy_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabGroupBy.SelectedTabChanged

        Select Case tabGroupBy.SelectedTab.Key

            Case "By Sales Order"
                grdSOTORDRX_AfterRowActivate(Nothing, Nothing)
                splSOTORDRX.Panel2Collapsed = False

            Case "By Style"
                grdICTSTYL1_AfterRowActivate(Nothing, Nothing)
                splSOTORDRX.Panel2Collapsed = False

            Case Else
                splSOTORDRX.Panel2Collapsed = True
        End Select
    End Sub

#End Region

#Region "Form Procedures"

    Private Function UpdateCancelOrderData() As Boolean

        Try
            BeginTrans()

            dst.Tables("SOTORDR1").AcceptChanges()
            dst.Tables("SOTORDR2").AcceptChanges()
            ASCMAIN1.Progress("Updating Order", "")
            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "CUST_CODE,ORDR_NO")
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                ASCMAIN1.Progress("-", ORDR_NO)
                For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select("ORDR_NO = '" & ORDR_NO & "'")
                    rowSOTORDRC.Item("SELECTED") = Val(rowSOTORDR1.Item("SELECTED") & String.Empty)
                Next

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                    Dim ORDR_LNO As Integer = rowSOTORDR2.Item("ORDR_LNO") & String.Empty
                    For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)
                        rowSOTORDRC.Item("SELECTED_DTL") = Val(rowSOTORDR2.Item("SELECTED") & String.Empty)
                    Next
                Next
            Next

            dst.Tables("SOTORDRC").AcceptChanges()
            For Each row As DataRow In dst.Tables("SOTORDRC").Select("")
                row.SetAdded()
            Next
            ASCDATA1.ExecuteSQL("DELETE FROM SOTORDRC")
            clsASCBASE1.Create_BAs("SOTORDRC", False)
            clsASCBASE1.Update_BAs("SOTORDRC", False)

            CommitTrans()
            Return True

        Catch ex As Exception
            Rollback()
            MessageBox.Show($"Error Updating SOTORDRC: {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            ASCMAIN1.Progress("", "")
        End Try
    End Function

    Private Function CancelItemQtys() As Boolean

        If Not releasedInventoryShortages Then
            Return True
        End If

        Try
            If cancelItems Then
                CancelSelectedItems()
            End If
            ASCMAIN1.Progress("", "")
            Return True

        Catch ex As Exception
            MessageBox.Show($"Error releasing Inventory Shortages: {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try


    End Function

    Private Function ClearNoDetailSalesOrders() As Boolean

        Try
            ' Get all letters with no details out of the way.
            ASCMAIN1.Progress("Updating User Canceled 'Item Cancellation Letter'", "")
            For Each row As DataRow In dst.Tables("SOTORDR1").Select("SELECTED = '1'", "CUST_CODE,ORDR_NO")
                If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{row.Item("ORDR_NO")}' AND SELECTED = '1'").Length = 0 Then
                    Try
                        BeginTrans()
                        ASCMAIN1.Progress("-", row.Item("ORDR_NO"))
                        TAC.TACMAIN1.Record_Event("SOTORDR1",
                              row.Item("ORDR_NO"),
                              Now,
                              ASCMAIN1.USER_ID,
                              "CANCLC",
                              "User Canceled 'Item Cancellation Letter'")

                        ASCDATA1.ExecuteSQL($"DELETE FROM SOTORDRC WHERE ORDR_NO = :PARM1", "V", New Object() {row.Item("ORDR_NO")})

                        CommitTrans()
                        row.Item("SELECTED") = LettersWithoutDetails

                    Catch ex As Exception
                        Rollback()
                        MessageBox.Show($"Error setting orders with no details (1), Order No {row.Item("ORDR_NO")}: {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End Try
                End If
            Next

            Return True

        Catch ex As Exception
            MessageBox.Show($"Error setting orders with no details (2): {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Private Function ProcessCustomerEmails() As Boolean
        Dim reportFilter As String = String.Empty

        Try
            ' Process Emails
            ' This is done incase there is an error in the processing we do not email or reprint any orders
            ASCMAIN1.Progress("Processing Letters", "")
            Dim ictr As Int16 = 0
            For Each row As DataRow In dst.Tables("SOTORDR1").Select("SELECTED = '1'", "CUST_CODE, ORDR_NO")
                Dim fileAttached As String = String.Empty
                Dim emailedTo As String = String.Empty
                Dim sql As String = String.Empty

                reportFilter = "{SOTORDR1.ORDR_NO}='" & row.Item("ORDR_NO") & "' AND {SOTORDR2.SELECTED}='1'"

                ASCMAIN1.Progress("-", row.Item("ORDR_NO"))

                If EmailCustomer(row.Item("ORDR_NO"), fileAttached, emailedTo, reportFilter, False) Then
                    row.Item("SELECTED") = LetterEmailedOrPrinted

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{row.Item("ORDR_NO")}' AND ORDR_LINE_CANC = '1'")
                        rowSOTORDR2.Item("ORDR_LINE_CANC") = "0"
                    Next

                    'EMAILED_TO_CUST, EMAILED_TO_SREP, PRINTED
                    ASCDATA1.ExecuteSQL($"UPDATE SOTORDRC SET EMAILED_TO_CUST = '1' WHERE ORDR_NO = :PARM1", "V", New Object() {row.Item("ORDR_NO")})
                End If

                ictr += 1
                If ictr Mod 10 = 0 Then
                    Fill_Stats()
                End If

                If fileAttached.Length > 0 Then
                    dst.Tables("ASTATTA2").Rows.Clear()
                    ENTITY.TABLE_NAME = "SOTORDR1"
                    ENTITY.COLUMN_NAME = "ORDR_NO"
                    ENTITY.CODE_VALUE = row.Item("ORDR_NO") & String.Empty

                    fileAttached = "" & fileAttached

                    MyBase.Attach_File(fileAttached, "Printed Cancellation Letter")


                    TAC.TACMAIN1.Record_Event("SOTORDR1",
                          row.Item("ORDR_NO"),
                          Now,
                          ASCMAIN1.USER_ID,
                          "CANCL",
                          "Item Cancellation Letter generated/printed On: " & DateTime.Now)
                End If
            Next

            Return True

        Catch ex As Exception
            MessageBox.Show($"Error processing Emails (2): {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Function

    Private Function ProcessSalesRepEmails() As Boolean

        Dim reportFilter As String = String.Empty


        Try
            ' Email Sales Reps
            Dim SREP_CODE As String = String.Empty
            Dim fileAttached As String = String.Empty
            Dim emailedTo As String = String.Empty
            Dim ictr As Int16 = 0

            For Each row As DataRow In ASCDATA1.SelectDistinct("SOTORDR1", New String() {"SREP_CODE"}).Select("", "SREP_CODE")

                SREP_CODE = row.Item("SREP_CODE") & String.Empty
                SREP_CODE = SREP_CODE.Trim

                If SREP_CODE.Length = 0 Then
                    Continue For
                End If

                ' Only email any that were not previously emailed.
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SREP_CODE = '" & SREP_CODE & "'")
                    Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    If dst.Tables("SOTORDRC").Select($"ORDR_NO = '{rowSOTORDR1.Item("ORDR_NO")}' AND EMAILED_TO_SREP <> '1'").Length > 0 Then
                        rowSOTORDR1.Item("SELECTED") = LetterEmailedOrPrinted
                    End If
                Next

                If dst.Tables("SOTORDR1").Select("SELECTED = '2' AND SREP_CODE = '" & SREP_CODE & "'").Length = 0 Then
                    Continue For
                End If

                reportFilter = "{SOTORDR1.SREP_CODE}='" & SREP_CODE & "' AND {SOTORDR1.SELECTED}='2' AND {SOTORDR2.SELECTED}='1'"

                If EmailCustomer(SREP_CODE, fileAttached, emailedTo, reportFilter, True) Then
                    Try
                        BeginTrans()
                        For Each rowwk As DataRow In dst.Tables("SOTORDR1").Select("SELECTED = '2' AND SREP_CODE = '" & SREP_CODE & "'")
                            TAC.TACMAIN1.Record_Event("SOTORDR1",
                                 rowwk.Item("ORDR_NO"),
                                 Now,
                                 ASCMAIN1.USER_ID,
                                 "CANCL",
                                 "Item Cancellation Letter emailed to Sales Rep on: " & DateTime.Now & " to " & emailedTo)

                            'EMAILED_TO_CUST, EMAILED_TO_SREP, PRINTED
                            ASCDATA1.ExecuteSQL("UPDATE SOTORDRC SET EMAILED_TO_SREP = '1' WHERE ORDR_NO = :PARM1", "V", New Object() {rowwk.Item("ORDR_NO")})
                        Next

                        CommitTrans()

                    Catch ex As Exception
                        Rollback()
                        MessageBox.Show($"Error emailing sales rep ({SREP_CODE}): " & ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End Try
                End If

                ictr += 1
                If ictr Mod 8 = 0 Then
                    Fill_Stats()
                End If
            Next

            Return True

        Catch ex As Exception
            MessageBox.Show($"Error emailing sales reps: {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            ASCMAIN1.Progress("", "")
        End Try


    End Function

    Private Function ProcessPrintedletters()
        ' Process printed letters
        Try
            ' Start with a fresh copy of the data
            Fill_Records("SOTORDRC", String.Empty, True, "SELECT * FROM SOTORDRC")
            For Each row As DataRow In dst.Tables("SOTORDRC").Select("EMAILED_TO_CUST = '1' OR PRINTED = '1'")
                Dim ORDR_NO As String = row.Item("ORDR_NO") & String.Empty
                For Each rowX As DataRow In dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "' AND SELECTED = '1'")
                    rowX.Item("SELECTED") = LetterEmailedOrPrinted
                Next
            Next

            If dst.Tables("SOTORDR1").Select("SELECTED = '1'").Length > 0 Then
                Print_Report_Begin()
                Dim REPORT_NO As String = Generate_Report("SORDISC1", "Modified Sales Order", "", "{SOTORDR1.SELECTED}='1' AND {SOTORDR2.SELECTED}='1'", "RPT", "", False)
                Print_Report_End()
            End If

            For Each row As DataRow In dst.Tables("SOTORDR1").Select("SELECTED = '1'", "CUST_CODE,ORDR_NO")

                row.Item("SELECTED") = LetterEmailedOrPrinted

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{row.Item("ORDR_NO")}' AND ORDR_LINE_CANC = '1'")
                    rowSOTORDR2.Item("ORDR_LINE_CANC") = "0"
                Next

                ASCDATA1.ExecuteSQL("UPDATE SOTORDRC SET PRINTED = '1' WHERE ORDR_NO = :PARM1", "V", New Object() {row.Item("ORDR_NO")})
            Next

            Return True

        Catch ex As Exception
            MessageBox.Show($"Error printing letters (2): {ex.Message}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Function

    Private Function EmailCustomer(ByVal ORDR_NO As String, ByRef fileAttachment As String, ByRef emailedTo As String, ByRef filter As String, ByVal emailToSalesRep As Boolean) As Boolean

        Dim customerEmailFound As Boolean = False

        Try

            If ASCMAIN1.CLIENT <> "RGI" Then
                Return True
            End If

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            If rowSOTORDR1 Is Nothing AndAlso Not emailToSalesRep Then
                Return False
            End If

            ' If we have an attachment record for this sales order then get out.
            Dim rowASTATTA2 As DataRow = Nothing

            If Not emailToSalesRep Then

                ' If there is an entry in ASTATTA2 and the file exisit then get out of here
                ASCMAIN1.sql = "SELECT * FROM ASTATTA2
                                where table_name = 'SOTORDR1'
                                AND COLUMN_NAME = 'ORDR_NO'
                                AND CODE_VALUE = :PARM1
                                AND ATTACHMENT_DESC = 'Printed Cancellation Letter'"
                rowASTATTA2 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ORDR_NO})
                If rowASTATTA2 IsNot Nothing Then
                    Dim ATTACHMENT_NO As String = rowASTATTA2.Item("ATTACHMENT_NO") & String.Empty
                    If ATTACHMENT_NO.Length > 0 Then
                        Dim ATTACHMENT_FILENAME As String = AttachmentFolder & ATTACHMENT_NO
                        If ASCMAIN1.Running_in_VS Then
                            ATTACHMENT_FILENAME = ATTACHMENT_FILENAME.Replace("S:\", "R:\")
                        End If
                        If My.Computer.FileSystem.FileExists(ATTACHMENT_FILENAME) Then
                            fileAttachment = String.Empty
                            Return True
                        End If
                    End If
                End If

                ' This is done incase there is a crash.
                If rowASTATTA2 IsNot Nothing Then
                    If dst.Tables("SOTORDRC").Select("ORDR_NO = '" & ORDR_NO & "' AND PRINTED = '1'").Length > 0 Then
                        fileAttachment = String.Empty
                        Return False
                    End If
                End If
            End If

            Dim CUST_CODE As String = String.Empty
            Dim SREP_CODE As String = String.Empty

            If emailToSalesRep Then
                SREP_CODE = ORDR_NO
            Else
                CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
                SREP_CODE = rowSOTORDR1.Item("SREP_CODE") & String.Empty
            End If

            Dim emailToList As String = String.Empty
            Dim CUST_XMIT_INV_VIA As String = String.Empty

            ' See if the customer receives an acknowledgment
            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

            If rowSOTSREP1 Is Nothing AndAlso rowARTCUST1 Is Nothing Then
                Return False
            End If

            If rowSOTSREP1 IsNot Nothing AndAlso rowSOTSREP1.Item("SREP_EMAIL") & String.Empty <> String.Empty AndAlso emailToSalesRep Then
                emailToList = rowSOTSREP1.Item("SREP_EMAIL") & String.Empty
                customerEmailFound = True
            End If

            If Not emailToSalesRep Then
                If (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim.Length > 0 Then
                    emailToList &= ";" & (rowARTCUST1.Item("CUST_EMAIL") & String.Empty).ToString.Trim
                    customerEmailFound = True
                End If

                Dim tblARTCUSTD As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ARTCUSTD WHERE CONTACT_TYPE = 'B' AND CONTACT_EMAIL IS NOT NULL AND CUST_CODE = '" & CUST_CODE & "'")
                If tblARTCUSTD.Rows.Count > 0 Then
                    If tblARTCUSTD.Select("CONTACT_PRIMARY = '1'").Length > 0 Then
                        customerEmailFound = True
                        For Each row As DataRow In tblARTCUSTD.Select("CONTACT_PRIMARY = '1'")
                            emailToList &= ";" & (row.Item("CONTACT_EMAIL") & String.Empty).ToString.Trim
                        Next
                    Else
                        emailToList &= ";" & (tblARTCUSTD.Rows(0).Item("CONTACT_EMAIL") & String.Empty).ToString.Trim
                        customerEmailFound = True
                    End If
                End If
            End If

            ' remove double semi-colons
            While emailToList.Contains(" ")
                emailToList = emailToList.Replace(" ", "")
            End While
            emailToList = emailToList.Replace(",", ";")

            While emailToList.Contains(";;")
                emailToList = emailToList.Replace(";;", ";")
            End While

            ' should be at least 5 characters
            If emailToList.Replace(";", "").Trim.Length < 5 Then
                ' Return False
            End If

            If emailToSalesRep Then
                fileAttachment = SREP_CODE
            Else
                fileAttachment = rowARTCUST1.Item("CUST_NAME") & " " & ORDR_NO
            End If


            For Each invalidChar As String In New String() {"\", "/", ":", "*", "?", "<", ">", "|", ".", ","}
                fileAttachment = fileAttachment.Replace(invalidChar, "")
            Next
            fileAttachment = fileAttachment.Replace(" ", "_")

            emailedTo = emailToList

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (emailToList).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            ASCMAIN1.Progress("Sending Sales Rep / Customer a copy of the Modified Sales Order", "")

            Print_Report_Begin()

            Dim REPORT_NO As String = Generate_Report("SORDISC1", "Modified Sales Order", "", filter, "PDF", fileAttachment, False)
            Print_Report_End(False, True)

            ' If the file was not there then put it there and get out
            If rowASTATTA2 IsNot Nothing Then
                Dim ATTACHMENT_NO As String = rowASTATTA2.Item("ATTACHMENT_NO") & String.Empty
                My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & fileAttachment & ".pdf", AttachmentFolder & ATTACHMENT_NO)
                fileAttachment = String.Empty
                Return True
            End If

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(fileAttachment & ".pdf", ASCMAIN1.Folders("Temp") & fileAttachment & ".pdf")
            fileAttachment = ASCMAIN1.Folders("Temp") & fileAttachment & ".pdf"

            If EMAIL_ADDRESSs.Count = 0 Then
                Return False
            End If

            ' This is done incase there is a crash.
            If dst.Tables("SOTORDRC").Select("ORDR_NO = '" & ORDR_NO & "' AND (EMAILED_TO_CUST = '1' OR PRINTED = '1')").Length > 0 Then
                Return False
            End If

            Dim SUBJECT As String = String.Empty
            Dim SEND_NO As String = String.Empty

            ' Need to attach the letter to the sales order when we do no thave an email address.
            If emailToList.Replace(";", "").Trim.Length < 5 OrElse EMAIL_ADDRESSs.Count = 0 Then
                Return False
            End If

            If emailToSalesRep Then
                SUBJECT = "Modified Sales Orders for your customers"
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                      (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                       SUBJECT, "ORDRCANC", True, False, SREP_CODE, rowSOTSREP1.Item("SREP_NAME"), "Sales Rep")
            Else
                SUBJECT = "Modified Sales Order (" & ORDR_NO & ") for customer " & rowARTCUST1.Item("CUST_NAME")
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                        SUBJECT, "ORDRCANC", True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")
            End If

            Return (SEND_NO & String.Empty).Length > 0 AndAlso customerEmailFound

        Catch ex As Exception
            Return False
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Function

    Private Sub PrintReport()

    End Sub

    Private Sub CancelSelectedItems()

        Dim ORDR_GROUP_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim ORDR_LNO As String = String.Empty

        Dim tbl As DataTable = Nothing
        Dim detailLine As TAC.SOCORDR1.LineDetail
        Dim salesOrderLineDetails As New List(Of TAC.SOCORDR1.LineDetail)
        Dim clsSOTORDR1 As New TAC.SOCORDR1(Me)

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("SELECTED = '1'")
            ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty
            ORDR_NO = rowSOTORDR1.Item("ORDR_NO") & String.Empty
            ORDR_LNO = String.Empty

            ASCMAIN1.Progress("Processing Order Group: " & ORDR_GROUP_NO, String.Empty)
            salesOrderLineDetails.Clear()

            For Each row As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' and SELECTED = '1'")
                ORDR_LNO &= ", " & row.Item("ORDR_LNO")
            Next

            If ORDR_LNO.Length = 0 Then
                Continue For
            End If

            ORDR_LNO = ORDR_LNO.Substring(1).Trim
            tbl = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR2 WHERE ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO IN (" & ORDR_LNO & ")")

            For Each row As DataRow In tbl.Select("", "ORDR_NO, ORDR_LNO")
                If Val(row.Item("ORDR_QTY_OPEN") & String.Empty) <= 0 Then
                    Continue For
                End If

                detailLine = New TAC.SOCORDR1.LineDetail
                With detailLine
                    .OrderNo = row.Item("ORDR_NO")
                    .OrderLineNo = row.Item("ORDR_LNO")
                    .StyleCode = row.Item("STYLE_CODE")
                    .ColorCode = row.Item("COLOR_CODE")
                    .CancelQuantity = Val(row.Item("ORDR_QTY_OPEN") & String.Empty) - Val(row.Item("ORDR_QTY_ALLO") & String.Empty)
                End With

                If detailLine.CancelQuantity > 0 Then
                    salesOrderLineDetails.Add(detailLine)
                End If
            Next

            If salesOrderLineDetails.Count > 0 Then
                If Not clsSOTORDR1.CancelItemsFormSalesOrder(ORDR_GROUP_NO, salesOrderLineDetails) Then '
                    MessageBox.Show("Error Processing Order " & ORDR_NO & ": " & clsSOTORDR1.LastError, "Update Sales Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    TAC.TACMAIN1.Record_Event("SOTORDR1",
                          ORDR_NO,
                          Now,
                          ASCMAIN1.USER_ID,
                          "MSHORT",
                          "Manufacture Shortages Canceled")
                End If
            End If
        Next
        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Private Sub Fill_Stats()

        ASCMAIN1.sql = "SELECT SUM(NUM_CUST) NUM_CUST, SUM(NUM_ORDERS) NUM_ORDERS, SUM(NUM_SELECTED) NUM_SELECTED, SUM(NUM_EMAILED) NUM_EMAILED, SUM(NUM_PRINTED) NUM_PRINTED, SUM(NUM_EMAILED_SREP) NUM_EMAILED_SREP
                                FROM
                                (
                                Select COUNT(DISTINCT SOTORDR1.CUST_CODE) NUM_CUST, 0 NUM_ORDERS, 0 NUM_SELECTED, 0 NUM_EMAILED, 0 NUM_PRINTED, 0 NUM_EMAILED_SREP
                                From SOTORDRC, SOTORDR1 Where SOTORDRC.ORDR_NO = SOTORDR1.ORDR_NO
                                UNION
                                SELECT 0 NUM_CUST, COUNT(DISTINCT ORDR_NO) NUM_ORDERS, 0 NUM_SELECTED, 0 NUM_EMAILED, 0 NUM_PRINTED, 0 NUM_EMAILED_SREP From SOTORDRC
                                UNION
                                SELECT 0 NUM_CUST, 0 NUM_ORDERS, COUNT(DISTINCT ORDR_NO) NUM_SELECTED, 0 NUM_EMAILED, 0 NUM_PRINTED, 0 NUM_EMAILED_SREP From SOTORDRC WHERE SELECTED_DTL = '1'
                                UNION
                                SELECT 0 NUM_CUST, 0 NUM_ORDERS, 0 NUM_SELECTED, COUNT(DISTINCT ORDR_NO) NUM_EMAILED, 0 NUM_PRINTED, 0 NUM_EMAILED_SREP From SOTORDRC WHERE EMAILED_TO_CUST = '1'
                                UNION
                                SELECT 0 NUM_CUST, 0 NUM_ORDERS, 0 NUM_SELECTED, 0 NUM_EMAILED, COUNT(DISTINCT ORDR_NO) NUM_PRINTED, 0 NUM_EMAILED_SREP From SOTORDRC WHERE PRINTED = '1'
                                UNION
                                SELECT 0 NUM_CUST, 0 NUM_ORDERS, 0 NUM_SELECTED, 0 NUM_EMAILED, 0 NUM_PRINTED, COUNT(DISTINCT ORDR_NO) NUM_EMAILED_SREP From SOTORDRC WHERE EMAILED_TO_SREP = '1'
                                )"
        Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)

        dst.Tables("STATS").Rows.Clear()

        dst.Tables("STATS").Rows.Add(New Object() {1, "# Custs", Val(row.Item("NUM_CUST") & String.Empty)})
        dst.Tables("STATS").Rows.Add(New Object() {2, "# Orders", Val(row.Item("NUM_ORDERS") & String.Empty)})
        dst.Tables("STATS").Rows.Add(New Object() {3, "# Sel Orders", Val(row.Item("NUM_SELECTED") & String.Empty)})
        dst.Tables("STATS").Rows.Add(New Object() {4, "# Emailed", Val(row.Item("NUM_EMAILED") & String.Empty)})
        dst.Tables("STATS").Rows.Add(New Object() {5, "# Printed", Val(row.Item("NUM_PRINTED") & String.Empty)})
        dst.Tables("STATS").Rows.Add(New Object() {6, "# Srep", Val(row.Item("NUM_EMAILED_SREP") & String.Empty)})

        Sort_grdColumns(grdStats, "SEQ", True)
        grdStats.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)


    End Sub

#End Region

End Class