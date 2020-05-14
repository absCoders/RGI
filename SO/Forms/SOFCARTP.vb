Imports System.Drawing
Public Class SOFCARTP

#Region "Declarations"
    Dim rowSOTSHIP1 As DataRow
    Dim sqlSOTSHIPX As String
    Dim CUST_CODE As String
    Dim SHIP_BOL_NO As String
    Dim ORDR_GROUP_NO As String
    Dim ORDR_CUST_PO As String
    Dim ORDR_LNOs As New List(Of Int64)

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
 
        Get_PARM("SOTPARM1")

        With dst
            sqlSOTSHIPX = "Select SOTSHIP1.*" & vbCrLf _
              & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" & vbCrLf _
              & " from SOTSHIP1,SOTORDR0" & vbCrLf _
              & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
              & "   and SOTSHIP1.SHIP_STATUS = 'P'" ' and NVL(SOTSHIP1.LP_STATUS,'0') <> '1'
            ASCMAIN1.sql = sqlSOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" & vbCrLf _
            & " from SOTSHIP1,SOTORDR0" & vbCrLf _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
              & " from SOTPICK1,SOTORDR1 " & vbCrLf _
              & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
              & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2, "ITEM_CODE")
 

            ASCMAIN1.sql = "Select SOTPICK1.PICK_NO, SOTORDR1.CUST_STORE_NO, SOTPICK1.ORDR_NO" & vbCrLf _
              & ", SOTPICK1.PICK_STATUS, SOTPICK1.PICK_RELEASED, SOTPICK1.PICK_FREIGHT" & vbCrLf _
              & ", SOTPICK1.PICK_PICKER, SOTPICK1.PICK_NO_REV" & vbCrLf _
              & ", SOTPICK1.PICK_PRINTED, SOTPICK1.PICK_PACKED, SOTPICK1.PICK_SHIPPED" & vbCrLf _
              & ", SOTPICK1.PICK_BATCH_NO, SOTPICK1.SHIP_BOL_NO, SOTPICK1.INV_NO" & vbCrLf _
              & ", SOTPICK1.PICK_CNT_CARTONS, SOTPICK1.PICK_TOTAL_WGT" & vbCrLf _
              & ", SOTPICK1.INIT_OPER, SOTPICK1.LAST_OPER, SOTPICK1.INIT_DATE, SOTPICK1.LAST_DATE" & vbCrLf _
              & ", SOTPICK0.PICK_FORCED" & vbCrLf _
              & " from SOTPICK1,SOTORDR1,SOTPICK0 " & vbCrLf _
              & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
              & "   and SOTPICK0.PICK_BATCH_NO = SOTPICK1.PICK_BATCH_NO" & vbCrLf _
              & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*, SOTORDR2.STYLE_CODE, SOTORDR2.COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & ", SOTPICK1.SHIP_BOL_NO, EDT850T2.EDI_COLOR_CODE" & vbCrLf _
                & ", SOTORDR2.CUST_STYLE_CODE, SOTORDR2.CUST_COLOR_CODE, SOTORDR2.CUST_UPC, SOTORDR2.CUST_SKU" & vbCrLf _
                & " from SOTPICK1,SOTPICK2,SOTORDR2,ICTCOLR1,EDT850T2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and EDT850T2.EDI_DOC_SEQ_NO (+) = SOTORDR2.EDI_DOC_SEQ_NO" & vbCrLf _
                & "   and EDT850T2.EDI_DTL_SEQ (+) = SOTORDR2.EDI_DTL_SEQ" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = SOTORDR2.COLOR_CODE" & vbCrLf _
                & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            Create_Relation("SOTORDR2", "SOTPICK2", "ORDR_NO,ORDR_LNO")
            .Tables("SOTPICK2").Columns.Add("ITEM_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).ITEM_CODE")

            'ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" _
            '    & " where SOTPICK1.SHIP_BOL_NO = :PARM1"
            'Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            'ASCMAIN1.sql = "Select SOTPICK2.* from SOTPICK1,SOTPICK2" _
            '    & " where SOTPICK1.SHIP_BOL_NO = :PARM1" _
            '    & "   and SOTPICK2.CART_NO = SOTPICK1.CART_NO"
            'Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTCART1.* from SOTCART1,SOTPICK1" _
                & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.* from SOTCART1,SOTPICK1,SOTCART2" _
                & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT * FROM WHTPPKM1"
            Create_TDA(.Tables.Add, "WHTPPKM1", "**", 0, False, "", 1)
            .Tables("WHTPPKM1").Columns.Add("PPK_INNERS", GetType(System.Int64))
            .Tables("WHTPPKM1").Columns.Add("MULTIPLE", GetType(System.Int64))

            ASCMAIN1.sql = "SELECT * from WHTPPKM2"
            Create_TDA(.Tables.Add, "WHTPPKM2", "**", 1, False, "", 3)

            Create_Relation("WHTPPKM1", "WHTPPKM2", "PPK_CODE")

            .Tables("WHTPPKM2").Columns.Add("PICK_QTY", GetType(System.Int64))
            .Tables("WHTPPKM2").Columns.Add("PPK_INNERS", GetType(System.Int64), "PARENT(WHTPPKM1_WHTPPKM2).PPK_INNERS")
            .Tables("WHTPPKM2").Columns.Add("MULTIPLE", GetType(System.Int64), "PARENT(WHTPPKM1_WHTPPKM2).MULTIPLE")
            .Tables("WHTPPKM2").Columns.Add("LEFT_QTY", GetType(System.Int64), "ISNULL(PICK_QTY,0) - (ISNULL(PPK_QTY,0) / IIF(ISNULL(PPK_INNERS,0)=0,1,ISNULL(PPK_INNERS,0))) * ISNULL(MULTIPLE,0)")


            ASCMAIN1.sql = "SELECT PPK_CODE FROM WHTPPKM1"
            Create_TDA(.Tables.Add, "WHTPPKMX", "**", 0, False, "", 1)

            With .Tables.Add("WHTPPKMM")
                .Columns.Add("STYLE_CODE")
                .Columns.Add("COLOR_CODE")
                .Columns.Add("PICK_QTY")
                .PrimaryKey = New DataColumn() {.Columns("STYLE_CODE"), .Columns("COLOR_CODE")}
            End With
        End With

        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")

        grdWHTPPKMX.DataSource = dst.Tables("WHTPPKM1")

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")

        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
        Create_Summary(grdSOTPICK2, "PICK_QTY")
        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART2, "QTY_PACKED")
        Create_Summary(grdSOTCART2, "CART_LNO", "Count")
    End Sub


    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        Dim MsgBody As String = ""
        Dim MsgTitle As String = ""

        EMsg = ""

        Select Case eItemKey

            Case "Select"
                If Absx1.txtFor("SHIP_BOL_NO").Text = "" Then
                    EMsg &= "You must First Select a Shipment No"
                Else
                    SHIP_BOL_NO = Absx1.txtFor("SHIP_BOL_NO").Text
                    If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

                    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        EMsg &= String.Format("Invalid Shipment No ({0})", SHIP_BOL_NO)
                    Else
                        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")
                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR0.Item("CUST_CODE")
                        ORDR_CUST_PO = rowSOTORDR0.Item("ORDR_CUST_PO") & ""
                        Dim SHIP_STATUS As String = rowSOTSHIP1.Item("SHIP_STATUS")
                        If SHIP_STATUS <> "P" Then
                            EMsg &= String.Format("{0}Shipment {1} is No Longer in Pick", vbCr, SHIP_BOL_NO)
                        End If

                        If rowSOTSHIP1.Item("LP_STATUS") & "" = "1" Then
                            EMsg &= String.Format("{0}Shipment {1} has already been Transmitted", vbCr, SHIP_BOL_NO)
                        End If

                        If EMsg = "" Then
                            If CUST_CODE = "KMART" Then
                                MsgBody = "You Are Re-Cartonizing KMart Orders." _
                                & vbCrLf & "Please Follow These Cartons Carefully Through The Shipment Process."
                                MsgTitle = "KMart Warning"
                                MsgBox(MsgBody, MsgBoxStyle.Information, MsgTitle)
                            End If

                            If CUST_CODE = "SEARS" Then ' DON'T EVEN WANT TO THINK ABOUT THIS ONE
                                'Only_Weights(True, "Re-Pack not Avail. for SEARS")
                                EMsg = String.Format("{0}{1}Cannot Re-Pack a SEARS Carton", EMsg, vbCr)
                            End If
                        End If


                    
                    End If
                End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Done", "Cancel"
               
            Case "Update"
               
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Select"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                
                Update_Record()
                Mode_Settings(False)

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done", "Cancel"
                Mode_Settings(False)
           
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Select").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Done").Visible = InquiryMode Or (EntryMode = "L" And ScreenMode)
            .Groups("Screen Control").Items("Select").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
            .Groups("Screen Control").Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
            .Groups("Screen Control").Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
            '.Groups("Totals").Visible = ScreenMode
            '.Groups("Splits").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode
        End With

        '  lblStatus.Visible = ScreenMode

        grdSOTSHIPX.Visible = Not tf
        splWHTPPKM1.Visible = tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            'With grdSOTCART1.DisplayLayout.Override
            '    .AllowAddNew = UltraWinGrid.AllowAddNew.No
            '    .AllowUpdate = DefaultableBoolean.False
            '    .AllowDelete = DefaultableBoolean.False
            'End With
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("SHIP_BOL_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        CUST_CODE = ""
        SHIP_BOL_NO = ""
        ORDR_GROUP_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTCART1", "SOTCART2", "SOTPICK1", "SOTPICK2", "WHTPPKM1", "WHTPPKM2", "WHTPPKMM"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_SOTSHIPX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
        Else
            Fill_Records("SOTSHIP1", SHIP_BOL_NO)

            Fill_Records("SOTCART1", SHIP_BOL_NO)
            Fill_Records("SOTCART2", SHIP_BOL_NO)

            Fill_Records("SOTPICK1", SHIP_BOL_NO)
            Fill_Records("SOTPICK2", SHIP_BOL_NO)

            Fill_Records("SOTORDR1", SHIP_BOL_NO)
            Fill_Records("SOTORDR2", SHIP_BOL_NO)

            '  Fill_Records("SOTORDR1", SHIP_BOL_NO)
            'Fill_Records("WHTPPKM1")
            'Fill_Records("WHTPPKM2")
        End If

        grdSOTPICK1.Text = "Pick Tickets for Shipment " & SHIP_BOL_NO

        Display_Totals()
           EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Records()
        For Each TABLE_NAME As String In New String() _
            {"SOTCART2", "SOTCART1"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME _
            & " where CART_NO in (Select CART_NO from SOTCART1 where PICK_NO in " _
            & " (Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'))"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
       
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        BeginTrans()
       
        Update_Record_TDA("SOTORDR2")

        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_BOL_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " AND SOTSHIP1.SHIP_STATUS = 'P'"
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= String.Format(" AND SOTORDR0.CUST_CODE = '{0}'", Absx1.txtFor("CUST_CODE").Text)
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= String.Format(" and SOTORDR0.ORDR_CUST_PO = '{0}'", Absx1.txtFor("ORDR_CUST_PO").Text)
                End If
        End Select
    End Sub
 
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICK2, "B", "Locate Prepack", "Select All", "De-Select All")
        Load_Popup_Menu(grdWHTPPKMX, "B", "Apply Prepack")
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
            Case "grdSOTPICK1"
                If (EntryMode = "E" Or EntryMode = "N") Then
                    e.Cancel = True
                End If

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

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

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICK2.Rows
                    grow.Selected = (e.Tool.Key = "Select All")
                Next

            Case "Apply Prepack"
                Dim PPK_CODE As String = grd.ActiveRow.Cells("PPK_CODE").Value
                If dst.Tables("WHTPPKM2").Select("PPK_CODE = '" & PPK_CODE & "' and ISNULL(PPK_QTY,0) = 0").Length <> 0 Then
                    MsgBox("Style/Colors on Pick Ticket not in Pre-Pack", MsgBoxStyle.OkOnly, "Cannot Apply Pre-Pack Selected")
                    Exit Sub
                End If
                If dst.Tables("WHTPPKM2").Select("PPK_CODE = '" & PPK_CODE & "' and ISNULL(LEFT_QTY,0) <> 0").Length <> 0 Then
                    MsgBox("Pre-Pack is not an exact multiple of Qty's in Style/Colors Selected", MsgBoxStyle.OkOnly, "Cannot Apply Pre-Pack Selected")
                    Exit Sub
                End If

                Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
                For Each ORDR_LNO As Int64 In ORDR_LNOs
                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                    rowSOTORDR2.Item("ITEM_CODE") = PPK_CODE
                Next

                MsgBox("Pre-Pack has been Successfully Applied to the Order Details" & _
                       vbCrLf & " Corresponding to the Pick Ticket Details Selected", MsgBoxStyle.OkOnly, "Verification")

            Case "Locate Prepack"

                If grdSOTPICK2.Selected.Rows.Count = 0 Then
                    MsgBox("No Details Selected", MsgBoxStyle.OkOnly, "Cannot Select Matching Pre-Pack")
                    Exit Sub
                End If

                dst.Tables("WHTPPKMM").Rows.Clear()
                ORDR_LNOs.Clear()

                Dim sql As String = ""
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICK2.Selected.Rows
                    Dim STYLE_CODE As String = grow.Cells("STYLE_CODE").Value
                    Dim COLOR_CODE As String = grow.Cells("COLOR_CODE").Value
                    Dim PICK_QTY As Int64 = Val(grow.Cells("PICK_QTY").Value & "")
                    ORDR_LNOs.Add(Val(grow.Cells("ORDR_LNO").Value & ""))

                    sql &= " INTERSECT "
                    sql &= "Select Distinct PPK_CODE from WHTPPKM2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"

                    With dst.Tables("WHTPPKMM")
                        Dim rowWHTPPKMM As DataRow = .Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                        If rowWHTPPKMM Is Nothing Then
                            .Rows.Add(New Object() {STYLE_CODE, COLOR_CODE, PICK_QTY})
                        Else
                            rowWHTPPKMM.Item("PICK_QTY") = Val(rowWHTPPKMM.Item("PICK_QTY") & "") + PICK_QTY
                        End If
                    End With
                Next
                ASCMAIN1.sql = Mid(sql, 12)

                ASCMAIN1.sql = "Select * from WHTPPKM1 where PPK_CODE in (" & ASCMAIN1.sql & ")"


                EnforceConstraints(False)
                Fill_Records("WHTPPKM1", "", True, ASCMAIN1.sql)

                If dst.Tables("WHTPPKM1").Rows.Count = 0 Then
                    MsgBox("No Pre-Packs Found with all of the Styles Selected", MsgBoxStyle.OkOnly, "Cannot Locate any Matching Pre-Packs")
                    Exit Sub
                End If

                dst.Tables("WHTPPKM2").Rows.Clear()
                Dim COLEXP As String = dst.Tables("WHTPPKM2").Columns("LEFT_QTY").Expression
                dst.Tables("WHTPPKM2").Columns("LEFT_QTY").Expression = ""
                For Each rowWHTPPKM1 As DataRow In dst.Tables("WHTPPKM1").Select("")
                    Dim MULTIPLE As Int64 = 0
                    Dim PPK_CODE As String = rowWHTPPKM1.Item("PPK_CODE")
                    rowWHTPPKM1.Item("PPK_INNERS") = 1
                    Fill_Records("WHTPPKM2", PPK_CODE, False)
                    Dim MIN_QTY As Int64 = Val(dst.Tables("WHTPPKM2").Compute("MIN(PPK_QTY)", "PPK_CODE = '" & PPK_CODE & "'") & "")
                    rowWHTPPKM1.Item("PPK_INNERS") = MIN_QTY

                    For Each rowWHTPPKMM As DataRow In dst.Tables("WHTPPKMM").Select("")
                        Dim STYLE_CODE As String = rowWHTPPKMM.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowWHTPPKMM.Item("COLOR_CODE")
                        Dim rowWHTPPKM2 As DataRow = dst.Tables("WHTPPKM2").Rows.Find(New String() {PPK_CODE, STYLE_CODE, COLOR_CODE})
                        If rowWHTPPKMM Is Nothing Then
                            dst.Tables("WHTPPKM2").Rows.Add(New Object() {PPK_CODE, STYLE_CODE, COLOR_CODE, 0, rowWHTPPKMM.Item("PICK_QTY")})
                        Else

                            rowWHTPPKM2.Item("PICK_QTY") = rowWHTPPKMM.Item("PICK_QTY")
                        End If
                        If MULTIPLE = 0 And Val(rowWHTPPKM2.Item("PPK_QTY") & "") <> 0 Then MULTIPLE = Val(rowWHTPPKMM.Item("PICK_QTY") & "") / (Val(rowWHTPPKM2.Item("PPK_QTY") & "") / Val(rowWHTPPKM2.Item("PPK_INNERS") & ""))
                    Next
                    rowWHTPPKM1.Item("MULTIPLE") = MULTIPLE
                Next

                dst.Tables("WHTPPKM2").Columns("LEFT_QTY").Expression = COLEXP
                EnforceConstraints(True)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Load_SOTSHIPX()
                End If
            Case "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Load_SOTSHIPX()
                End If
            Case "SHIP_BOL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Select")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPX()
                End If
            Case "ORDR_CUST_PO"
                If Not ScreenMode Then
                    Load_SOTSHIPX()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTSHIPX()
            Case "SHIPMENT_NO"
                Click_Command("Select")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
     
#End Region

    Sub Load_SOTSHIPX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim ORDR_CUST_PO As String = Absx1.txtFor("ORDR_CUST_PO").Text
        Dim GridCaption As String = ""
         
        If CUST_CODE = "" Then
            Fill_Records("SOTSHIPX")
            grdSOTSHIPX.Text = "Shipments in Pick"
            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        Else
            If ORDR_CUST_PO = "" Then
                ASCMAIN1.sql = String.Format("{0} and SOTORDR0.CUST_CODE = '{1}'", sqlSOTSHIPX, CUST_CODE)
                GridCaption = String.Format("Shipments in Pick associated with {0}", CUST_CODE)
            Else
                ASCMAIN1.sql = String.Format("{0} AND SOTORDR0.CUST_CODE = '{1}' AND SOTORDR0.ORDR_CUST_PO = '{2}'", sqlSOTSHIPX, CUST_CODE, ORDR_CUST_PO)
                GridCaption = String.Format("Shipments in Pick associated with {0} for PO {1}", CUST_CODE, ORDR_CUST_PO)
            End If
            grdSOTSHIPX.Text = GridCaption
            Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        End If

        grdSOTSHIPX.Visible = True
    End Sub

    Sub Display_Totals()
        
    End Sub

    Private Sub grdSOTSHIPX_DoubleClickRow(sender As System.Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SHIP_BOL_NO").Text = e.Row.Cells("SHIP_BOL_NO").Value & ""
            Click_Command("Select")
        End If
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTPICK2.Visible = False
        Else
            grdSOTPICK2.Visible = True
            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTPICK2.Text = "Style/Color Details for Pick Ticket " & PICK_NO
        End If

        dst.Tables("WHTPPKMM").Rows.Clear()
        ORDR_LNOs.Clear()

    End Sub

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        If grdSOTCART1.ActiveRow Is Nothing OrElse Not grdSOTCART1.ActiveRow.IsDataRow Then
            grdSOTCART2.Visible = False
        Else
            grdSOTCART2.Visible = True
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
            dvw.RowFilter = "CART_NO = '" & CART_NO & "'"
            grdSOTCART2.Text = "Carton Details for Carton " & CART_NO
        End If
    End Sub

    Private Sub grdWHTPPKMX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTPPKMX.InitializeRow
        If e.Row.Band.Key = "WHTPPKM1" Then
            'e.Row.Appearance.ForeColor = Color.Red
        Else
            If Val(e.Row.Cells("LEFT_QTY").Value & "") <> 0 Then
                e.Row.Appearance.ForeColor = Color.Red
            End If
        End If
    End Sub
End Class