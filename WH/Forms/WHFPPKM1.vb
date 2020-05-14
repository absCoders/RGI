Public Class WHFPPKM1

    Dim PPK_CODE As String

#Region "ABS Standard Routines"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
          
            ASCMAIN1.sql = "Select WHTPPKM1.*" & vbCrLf _
                & ", X.ORDERS" & vbCrLf _
                & ", Y.STYLES, Y.STYLE_CODE" & vbCrLf _
                & ", SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTORDR1,WHTPPKM1" & vbCrLf _
                & ", (Select ITEM_CODE PPK_CODE, Count (*) ORDERS, MAX(ORDR_NO) ORDR_NO from (Select Distinct ITEM_CODE, ORDR_NO from SOTORDR2 where ITEM_CODE is Not Null) group by ITEM_CODE) X" & vbCrLf _
                & ", (Select PPK_CODE, Count (*) STYLES, MIN (STYLE_CODE) STYLE_CODE from WHTPPKM2 group by PPK_CODE) Y" & vbCrLf _
                & " where X.PPK_CODE (+) = WHTPPKM1.PPK_CODE" & vbCrLf _
                & "   and Y.PPK_CODE (+) = WHTPPKM1.PPK_CODE" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO (+) = X.ORDR_NO"
            Create_TDA(.Tables.Add, "WHTPPKMX", "**", 0, False, "", 0)
            .Tables("WHTPPKMX").Columns("ORDERS").DataType = GetType(System.Int32)
            .Tables("WHTPPKMX").Columns("STYLES").DataType = GetType(System.Int32)

            Create_TDA(.Tables.Add, "WHTPPKM1", "*", 1, False)

            ASCMAIN1.sql = "Select WHTPPKM2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from WHTPPKM2,ICTSTYL1,ICTCOLR1" _
                & " where WHTPPKM2.PPK_CODE = :PARM1" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = WHTPPKM2.STYLE_CODE" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = WHTPPKM2.COLOR_CODE" & vbCrLf

            Create_TDA(.Tables.Add, "WHTPPKM2", "**", 0, False, "V", 3)


            ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                & " from SOTORDR1" _
                & " where SOTORDR1.ORDR_NO in (Select Distinct ORDR_NO from SOTORDR2 where SOTORDR2.ITEM_CODE = :PARM1)" & vbCrLf

            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 3)

        End With

        grdWHTPPKMX.DataSource = dst.Tables("WHTPPKMX")
        grdWHTPPKM2.DataSource = dst.Tables("WHTPPKM2")

        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")

        Create_Summary(grdWHTPPKMX, "PPK_CODE", "Count")
        Create_Summary(grdWHTPPKM2, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPPKM2, "PPK_QTY")

        tabWHTPPKMX.Tabs("Purchase Orders").Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                PPK_CODE = Absx1.txtFor("PPK_CODE").Text
                Dim row As DataRow = LookUp("WHTPPKM1", PPK_CODE)
                If row Is Nothing Then
                    EMsg &= vbCr & "Invalid Value Specified for Pre-Pack Code (" & PPK_CODE & ")"
                End If
                If EMsg = "" Then

                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode
        grdWHTPPKMX.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"WHTPPKM1", "WHTPPKM2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If grdWHTPPKMX.Rows.Count = 0 Then
            Load_WHTPPKMX()
        End If
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        EnforceConstraints(False)

        Fill_Records("WHTPPKM1", PPK_CODE)
        Fill_Records("WHTPPKM2", PPK_CODE)
        Sort_grdColumns(grdWHTPPKM2, "STYLE_CODE")
        grdWHTPPKM2.Text = "Pre-Pack Details for " & PPK_CODE

        Fill_Records("SOTORDRX", PPK_CODE)
        Sort_grdColumns(grdSOTORDRX, "ORDR_NO")

        EnforceConstraints(True)
 
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTPPKMX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTORDRX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Sales Order Inquiry")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim filter As String = String.Empty
        Select Case grd.Name

        End Select

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PPK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PPK_CODE"
                Click_Command("View")
        End Select
    End Sub
#End Region


    Sub Load_WHTPPKMX()
        Fill_Records("WHTPPKMX")
        Sort_grdColumns(grdWHTPPKMX, "PPK_CODE".ToLower)
    End Sub

    Private Sub grdWHTPPKMX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTPPKMX.DoubleClickRow
        Absx1.txtFor("PPK_CODE").Text = e.Row.Cells("PPK_CODE").Value & ""
        Click_Command("View")
    End Sub

End Class