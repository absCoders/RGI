Imports System.Drawing
Imports System.Math

Public Class WHFCART2

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.
    Dim Pick_No As String = ""
    Dim Tran_No As String = ""
    Dim Carton_Qty As Int64

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select * from SOTCART1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "", 1)

            ASCMAIN1.sql = " Select C2.CART_NO, C2.CART_LNO, C2.ORDR_NO, C2.ORDR_LNO," _
            & " C2.QTY_PACKED, L1.UPC_CODE, C2.SKU_NO, C2.STYLE_CODE, " _
            & " C2.COLOR_CODE,  C2.SIZE_DESC" _
            & " From SOTCART2 C2, ICVLUPC1 L1 " _
            & " Where Rownum < 1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "", 2)


            ASCMAIN1.sql = " Select C2.CART_NO, C2.CART_LNO, C2.ORDR_NO, C2.ORDR_LNO," _
             & " C2.QTY_PACKED, L1.UPC_CODE, C2.SKU_NO, C2.STYLE_CODE, " _
             & " C2.COLOR_CODE,  C2.SIZE_DESC" _
             & " From SOTCART2 C2, ICVLUPC1 L1 " _
             & " Where Rownum < 1"
            Create_TDA(.Tables.Add, "WHTCARTC", "**", 0, False, "", 2)

            ASCMAIN1.sql = " Select * from WHTAUDT1"
            Create_TDA(.Tables.Add, "WHTAUDT1", "**", 0, True, "", 1)

            ASCMAIN1.sql = " Select * from WHTAUDT1"
            Create_TDA(.Tables.Add, "WHTAUDTL", "**", 0, False, "", 1)

            ASCMAIN1.sql = " Select * from WHTAUDT2"
            Create_TDA(.Tables.Add, "WHTAUDT2", "**", 0, True, "", 2)
            .Tables("WHTAUDT2").Columns.Add("SURPLUS_QTY", GetType(System.Double), "IIF( SCAN_QTY > ORIG_QTY,SCAN_QTY - ORIG_QTY,0)")
            .Tables("WHTAUDT2").Columns.Add("GOOD_QTY", GetType(System.Double), "IIF(SCAN_QTY <= ORIG_QTY,SCAN_QTY,ORIG_QTY)")

            ASCMAIN1.sql = " Select * from WHTAUDT2"
            Create_TDA(.Tables.Add, "WHTAUDTR", "**", 0, False, "", 2)

            With .Tables.Add("SOTPICKS")
                .Columns.Add("PICK_NO", GetType(System.String))
                .Columns.Add("PICK_SCANNED", GetType(System.String))
                .Columns.Add("PICK_SCANNED_BY", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("PICK_NO")}
            End With

            With .Tables.Add("SOTSTATI")
                .Columns.Add("ITEM_CODE", GetType(System.String))
                .Columns.Add("ITEM_DESC", GetType(System.String))
                .Columns.Add("ITEM_QTY", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
            End With

            With .Tables.Add("SOTSTATC")
                .Columns.Add("STATUS_CODE", GetType(System.String))
                .Columns.Add("STATUS_DESC", GetType(System.String))
                .Columns.Add("STATUS_QTY", GetType(System.Int64))
                .PrimaryKey = New DataColumn() {.Columns("STATUS_CODE")}
            End With

        End With
        grdSOTPICKS.DataSource = dst.Tables("SOTPICKS")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdWHTAUDT2.DataSource = dst.Tables("WHTAUDT2")
        grdSOTSTATI.DataSource = dst.Tables("SOTSTATI")
        grdSOTSTATC.DataSource = dst.Tables("SOTSTATC")

        grdSOTSTATI.DisplayLayout.Bands(0).ColHeadersVisible = False
        grdSOTSTATC.DisplayLayout.Bands(0).ColHeadersVisible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Start Scan"
                Dim row As DataRow = clsASCBASE1.LookUp("SOTPICK1", txtPICK_NO.Text)
                If row Is Nothing Then
                    EMsg &= "Invalid Pick No"
                End If
            Case "Scan Complete"

            Case "Cancel"
                'If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                '          "You may have made Changes") = MsgBoxResult.No Then
                '    Exit Sub
                'End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Start Scan"
                EntryMode = "N"
                Pick_No = Absx1.txtFor("PICK_NO").Text
                Load_Record()
                Mode_Settings(True)
            Case "Scan Complete"
                Update_Record("C")
                Mode_Settings(False)
            Case "Cancel"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Start Scan").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Scan Complete").Settings.Enabled = iScreenMode

                End With

                .Groups("Carton Status").Visible = ScreenMode
                .Groups("Scanned Items").Visible = ScreenMode
                .Groups("Picks in DC").Visible = ScreenMode
            End With
        End If
        Set_Read_Only(UltraGroupBox1, ScreenMode)
        tab0.Visible = Not tf
        If ScreenMode Then
        Else
            Clear_Record()
        End If


    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTCART1", "SOTCART2", "SOTSTATI", "SOTSTATC", "WHTCARTC", "WHTAUDT2", "WHTAUDTR", "SOTPICKS"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Pick_No = ""
        Carton_Qty = 0
        Tran_No = ""
        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("CUST_NAME").Text = ""
        Absx1.txtFor("CUST_DC_NO").Text = ""
        Absx1.txtFor("PICK_NO").Text = ""
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        'Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        End If

        Tran_No = ASCMAIN1.Next_Control_No("WHTAUDT1")
        Fill_Temp_Tables()

        Dim Ordr_Group_No As String = ""
        ASCMAIN1.sql = " Select SOTORDR1.* from SOTPICK1, SOTORDR1" _
        & " Where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" _
        & " And SOTPICK1.PICK_NO = '" & Pick_No & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow
        If rowSOTORDR1 IsNot Nothing Then
            Absx1.txtFor("CUST_CODE").Text = rowSOTORDR1.Item("CUST_CODE") & ""
            Absx1.txtFor("CUST_NAME").Text = rowSOTORDR1.Item("CUST_NAME") & ""
            Absx1.txtFor("CUST_DC_NO").Text = rowSOTORDR1.Item("CUST_DC_NO") & ""
            Ordr_Group_No = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
        End If

        Dim rowWHTAUDT1 As DataRow = dst.Tables("WHTAUDT1").NewRow
        With rowWHTAUDT1
            .Item("TRAN_NO") = Tran_No
            .Item("PICK_NO") = Pick_No
            .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
            .Item("CUST_DC_NO") = rowSOTORDR1.Item("CUST_DC_NO") & ""
            .Item("SCAN_DATE") = DATETIME_STAMP
            .Item("SCAN_TIME") = ""
            .Item("SCAN_STATUS") = "C"
        End With
        dst.Tables("WHTAUDT1").Rows.Add(rowWHTAUDT1)

        ASCMAIN1.sql = "Select * from WHTAUDT1 Where ORDR_GROUP_NO = '" & Ordr_Group_No & "'" _
        & " And CUST_DC_NO = '" & Absx1.txtFor("CUST_DC_NO").Text & "'"
        Fill_Records("WHTAUDTL", , , ASCMAIN1.sql)

        ASCMAIN1.sql = "  Select PICK_NO " & vbCrLf _
        & "  from SOTPICK1 Where ORDR_NO in (" & vbCrLf _
        & "  Select ORDR_NO from SOTORDR1 Where (ORDR_GROUP_NO, CUST_DC_NO) in (" & vbCrLf _
        & "  Select ORDR_GROUP_NO, CUST_DC_NO from SOTPICK1, SOTORDR1" & vbCrLf _
        & "  Where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
        & "  And SOTPICK1.PICK_NO = '" & Pick_No & "'))"
        'Fill_Records("SOTPICKS", , , ASCMAIN1.sql)
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSOTPICKS As DataRow = dst.Tables("SOTPICKS").NewRow
            With rowSOTPICKS
                .Item("PICK_NO") = row.Item("PICK_NO")
                .Item("PICK_SCANNED") = IIf(dst.Tables("WHTAUDTL").Select("PICK_NO = '" & row.Item("PICK_NO") & "'").Length <> 0, "1", "0")
                .Item("PICK_SCANNED_BY") = ""
            End With
            dst.Tables("SOTPICKS").Rows.Add(rowSOTPICKS)
        Next
        Sort_grdColumns(grdSOTPICKS, "PICK_NO")
        dst.Tables("SOTPICKS").AcceptChanges()


        ASCMAIN1.sql = " Select C2.CART_NO, C2.CART_LNO, C2.ORDR_NO, C2.ORDR_LNO," & vbCrLf _
        & " C2.QTY_PACKED, L1.UPC_CODE, C2.SKU_NO, C2.STYLE_CODE, " & vbCrLf _
        & " C2.COLOR_CODE,  C2.SIZE_DESC" & vbCrLf _
        & " From SOTCART2 C2, ICVLUPC1 L1 " & vbCrLf _
        & " Where CART_NO in (Select CART_NO from SOTCART1 Where PICK_NO = '" & Pick_No & "')" & vbCrLf _
        & " And C2.STYLE_CODE = L1.STYLE_CODE(+)" & vbCrLf _
        & " And C2.COLOR_CODE = L1.COLOR_CODE(+)" & vbCrLf _
        & " And C2.SIZE_DESC = L1.SIZE_CODE(+)" & vbCrLf _
        & " And QTY_PACKED > 0"


        ASCMAIN1.sql = " Select C2.CART_NO, C2.CART_LNO, C2.ORDR_NO, C2.ORDR_LNO," & vbCrLf _
        & " C2.QTY_PACKED, R2.CUST_UPC as UPC_CODE, R2.CUST_SKU as SKU_NO, C2.STYLE_CODE, " & vbCrLf _
        & " C2.COLOR_CODE,  C2.SIZE_DESC" & vbCrLf _
        & " From SOTCART2 C2, SOTORDR2 R2 " & vbCrLf _
        & " Where CART_NO in (Select CART_NO from SOTCART1 Where PICK_NO = '" & Pick_No & "')" & vbCrLf _
        & " And R2.ORDR_NO(+) = C2.ORDR_NO" & vbCrLf _
        & " And R2.ORDR_LNO(+) = C2.ORDR_LNO" & vbCrLf
        Fill_Records("SOTCART2", , , ASCMAIN1.sql)
        Fill_Records("WHTCARTC", , , ASCMAIN1.sql)

        Carton_Qty = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "") & "")
        Calc_Detail()

        ASCMAIN1.Progress("")
    End Sub

    Sub Fill_Temp_Tables()
        dst.Tables("SOTSTATI").Clear()
        Dim rowSOTSTATI As DataRow = dst.Tables("SOTSTATI").NewRow
        With rowSOTSTATI
            .Item("ITEM_CODE") = "1"
            .Item("ITEM_DESC") = "QTY in Carton"
            .Item("ITEM_QTY") = 0
        End With
        dst.Tables("SOTSTATI").Rows.Add(rowSOTSTATI)

        rowSOTSTATI = dst.Tables("SOTSTATI").NewRow
        With rowSOTSTATI
            .Item("ITEM_CODE") = "2"
            .Item("ITEM_DESC") = "QTY Remaining"
            .Item("ITEM_QTY") = 0
        End With
        dst.Tables("SOTSTATI").Rows.Add(rowSOTSTATI)

        dst.Tables("SOTSTATC").Rows.Clear()
        Dim rowSOTSTATC As DataRow = dst.Tables("SOTSTATC").NewRow
        With rowSOTSTATC
            .Item("STATUS_CODE") = "1"
            .Item("STATUS_DESC") = "Items OK"
            .Item("STATUS_QTY") = 0
        End With
        dst.Tables("SOTSTATC").Rows.Add(rowSOTSTATC)

        rowSOTSTATC = dst.Tables("SOTSTATC").NewRow
        With rowSOTSTATC
            .Item("STATUS_CODE") = "2"
            .Item("STATUS_DESC") = "Surplus Qty"
            .Item("STATUS_QTY") = 0
        End With
        dst.Tables("SOTSTATC").Rows.Add(rowSOTSTATC)

        rowSOTSTATC = dst.Tables("SOTSTATC").NewRow
        With rowSOTSTATC
            .Item("STATUS_CODE") = "3"
            .Item("STATUS_DESC") = "Not in Carton"
            .Item("STATUS_QTY") = 0
        End With
        dst.Tables("SOTSTATC").Rows.Add(rowSOTSTATC)

        rowSOTSTATC = dst.Tables("SOTSTATC").NewRow
        With rowSOTSTATC
            .Item("STATUS_CODE") = "4"
            .Item("STATUS_DESC") = "Not in System"
            .Item("STATUS_QTY") = 0
        End With
        dst.Tables("SOTSTATC").Rows.Add(rowSOTSTATC)


        dst.Tables("SOTSTATI").AcceptChanges()
        dst.Tables("SOTSTATC").AcceptChanges()
        Sort_grdColumns(grdSOTSTATI, "ITEM_CODE", True)
        Sort_grdColumns(grdSOTSTATC, "STATUS_CODE", True)
    End Sub

    Sub Update_Record(Update_Type As String)

        BeginTrans()
        Update_Record_TDA("WHTAUDT1")
        Update_Record_TDA("WHTAUDT2")
        CommitTrans()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where &= " AND WHSE_CTN_CTL = 'C'"
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

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

        Select Case grd.Name
            Case "grdSOTRTRN2"


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case ""

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Location Inquiry"
            '    Dim Style_Code As String = "S:" & grd.ActiveRow.Cells("STYLE_CODE").Text
            '    Context_Launch("Select", Style_Code, e.Tool.Key, "WHFLOCS1", "F", "WHREC")

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PICK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode And Absx1.txtFor("PICK_NO").Text <> "" Then
                        Click_Command("Start Scan", e)
                    End If
                End If
            Case "UPC_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Validate_UPC_Code()
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PICK_NO"
                If Not InquiryMode And Absx1.txtFor("WHSE_CODE").Text <> "" Then
                    ' Click_Command("New")
                End If
        End Select
    End Sub


    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "PICK_NO"
                If txtPICK_NO.TextLength = 10 Then
                    ' Click_Command("Start Scan")
                End If
            Case "UPC_CODE"
                If UCase(txtUPC.Text) = "COMPLETE" Then
                    Click_Command("Scan Complete")
                End If
                If UCase(txtUPC.Text) = "CANCEL" Then
                    Click_Command("Cancel")
                End If
                If txtUPC.TextLength = 12 Then
                    Validate_UPC_Code()
                    txtUPC.Focus()
                End If
        End Select
    End Sub
#End Region

    Sub Validate_UPC_Code()
        Dim UPC_CODE As String = ""
        Dim TRAN_LNO As Integer
        Dim QTY_Scanned As String
        Dim Style As String
        Dim Color As String
        Dim Size As String
        Dim QTY_PACKED As Integer
        Dim ICVLUPC1_Flag As Boolean = True

        QTY_Scanned = "1"
        UPC_CODE = ""
        QTY_PACKED = 0
        Style = ""
        Color = ""
        Size = ""

        If txtUPC.Text <> "" Then
            Dim sql_UPC As String = ""
            UPC_CODE = txtUPC.Text
            TRAN_LNO = Val(dst.Tables("WHTAUDT2").Compute("MAX(TRAN_LNO)", "") & "") + 1
            Check_Duplicate_UPC()


            If ICVLUPC1_Flag = True Then
                ASCMAIN1.sql = "Select UPC_CODE, STYLE_CODE, COLOR_CODE, SIZE_CODE from ICVLUPC1 Where UPC_CODE = '" & UPC_CODE & "'"
            Else
                ASCMAIN1.sql = " Select CUST_UPC UPC_CODE, STYLE_CODE, COLOR_CODE, SIZE_DESC SIZE_CODE from SOTCSTY1" _
                & "  Where CUST_UPC = '" & UPC_CODE & "'"
            End If
            Dim rowICVLUPC1 As DataRow = ASCDATA1.GetDataRow
            If rowICVLUPC1 Is Nothing Then
                Dim iResponse As MsgBoxResult = MsgBox("UPC Code is not in System, do you still wish to add upc to audit", MsgBoxStyle.YesNo, "Proceed")
                If iResponse = MsgBoxResult.Yes Then
                    sql_UPC = "UPC_CODE = '" & UPC_CODE & "'"
                    Add_WHTAUDT2(txtUPC.Text, TRAN_LNO, Style, Color, Size, "U", QTY_Scanned, sql_UPC)
                End If
            Else
                sql_UPC = "UPC_CODE = '" & rowICVLUPC1.Item("UPC_CODE") & "'"
                If dst.Tables("SOTCART2").Select(sql_UPC).Length <> 0 Then
                    QTY_PACKED = Val(dst.Tables("SOTCART2").Select("UPC_CODE = '" & rowICVLUPC1.Item("UPC_CODE") & "'")(0).Item("QTY_PACKED") & "")

                    Style = rowICVLUPC1.Item("STYLE_CODE")
                    Color = rowICVLUPC1.Item("COLOR_CODE")
                    Size = rowICVLUPC1.Item("SIZE_CODE")
                    Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "G", QTY_Scanned, sql_UPC)
                    Update_SOWCART2(rowICVLUPC1.Item("UPC_CODE"), QTY_PACKED, 1, "", "", "", sql_UPC)

                Else
                    sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                    & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                    & " And SIZE_DESC = '" & rowICVLUPC1.Item("SIZE_CODE") & "'" _
                                                    & " And UPC_CODE is null"
                    If dst.Tables("SOTCART2").Select(sql_UPC).Length <> 0 Then

                        QTY_PACKED = Val(dst.Tables("SOTCART2").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                    & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                    & " And SIZE_DESC = '" & rowICVLUPC1.Item("SIZE_CODE") & "'" _
                                                    & " And UPC_CODE is null")(0).Item("QTY_PACKED") & "")
                        Style = rowICVLUPC1.Item("STYLE_CODE")
                        Color = rowICVLUPC1.Item("COLOR_CODE")
                        Size = rowICVLUPC1.Item("SIZE_CODE")

                        Update_SOWCART2(rowICVLUPC1.Item("UPC_CODE"), QTY_PACKED, 1, Style, Color, Size, sql_UPC)
                        Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "G", QTY_Scanned, sql_UPC)

                    Else
                        sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                          & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                          & " And UPC_CODE is null"
                        If dst.Tables("SOTCART2").Select(sql_UPC).Length <> 0 Then

                            QTY_PACKED = Val(dst.Tables("SOTCART2").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                          & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                          & " And UPC_CODE is null")(0).Item("QTY_PACKED") & "")

                            Style = rowICVLUPC1.Item("STYLE_CODE")
                            Color = rowICVLUPC1.Item("COLOR_CODE")
                            Size = ""
                            Update_SOWCART2(rowICVLUPC1.Item("UPC_CODE"), QTY_PACKED, 1, Style, Color, Size, sql_UPC)
                            Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "G", QTY_Scanned, sql_UPC)

                        Else
                            sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                  & " And UPC_CODE is null"
                            If dst.Tables("SOTCART2").Select(sql_UPC).Length <> 0 Then

                                QTY_PACKED = Val(dst.Tables("SOTCART2").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                  & " And UPC_CODE is null")(0).Item("QTY_PACKED") & "")

                                Style = rowICVLUPC1.Item("STYLE_CODE")
                                Color = ""
                                Size = ""
                                Update_SOWCART2(rowICVLUPC1.Item("UPC_CODE"), QTY_PACKED, 1, Style, Color, Size, sql_UPC)
                                Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "G", QTY_Scanned, sql_UPC)
                            Else
                                'Starting at this point i look at copy of SOWCART2 (WHWCARTC) to see if UPC was in table originally and is accounted for
                                ' Stop 'Added new code to check for UPC Code first, run Test
                                If dst.Tables("WHTCARTC").Select("UPC_CODE = '" & rowICVLUPC1.Item("UPC_CODE") & "'").Length <> 0 Then
                                    Dim iResponse As MsgBoxResult = MsgBox("This UPC Code has been accounted for, do you wish to add anyway?", MsgBoxStyle.YesNo, "Proceed")
                                    If iResponse = MsgBoxResult.Yes Then
                                        sql_UPC = "UPC_CODE = '" & rowICVLUPC1.Item("UPC_CODE") & "'"

                                        Style = rowICVLUPC1.Item("STYLE_CODE")
                                        Color = rowICVLUPC1.Item("COLOR_CODE")
                                        Size = rowICVLUPC1.Item("SIZE_CODE")
                                        Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "O", QTY_Scanned, sql_UPC)
                                    End If
                                Else

                                    If dst.Tables("WHTCARTC").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                                    & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                                    & " And SIZE_DESC = '" & rowICVLUPC1.Item("SIZE_CODE") & "'" _
                                                                    & " And UPC_CODE is null").Length <> 0 Then
                                        Dim iResponse As MsgBoxResult = MsgBox("This UPC Code has been accounted for, do you wish to add anyway?", MsgBoxStyle.YesNo, "Proceed")
                                        If iResponse = MsgBoxResult.Yes Then
                                            sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                            & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                            & " And SIZE_DESC = '" & rowICVLUPC1.Item("SIZE_CODE") & "'" _
                                            & " And UPC_CODE is null"

                                            Style = rowICVLUPC1.Item("STYLE_CODE")
                                            Color = rowICVLUPC1.Item("COLOR_CODE")
                                            Size = rowICVLUPC1.Item("SIZE_CODE")
                                            Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "O", QTY_Scanned, sql_UPC)
                                        End If
                                    Else
                                        If dst.Tables("WHTCARTC").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                                        & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                                        & " And UPC_CODE is null").Length <> 0 Then
                                            Dim iResponse As MsgBoxResult = MsgBox("This UPC Code has been accounted for, do you wish to add anyway?", MsgBoxStyle.YesNo, "Proceed")
                                            If iResponse = MsgBoxResult.Yes Then
                                                sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                  & " And COLOR_CODE = '" & rowICVLUPC1.Item("COLOR_CODE") & "'" _
                                                  & " And UPC_CODE is null"

                                                Style = rowICVLUPC1.Item("STYLE_CODE")
                                                Color = rowICVLUPC1.Item("COLOR_CODE")
                                                Size = ""
                                                Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "O", QTY_Scanned, sql_UPC)
                                            End If
                                        Else
                                            If dst.Tables("WHTCARTC").Select("STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "'" _
                                                                             & " And UPC_CODE is null").Length <> 0 Then
                                                Dim iResponse As MsgBoxResult = MsgBox("This UPC Code has been accounted for, do you wish to add anyway?", MsgBoxStyle.YesNo, "Proceed")
                                                If iResponse = MsgBoxResult.Yes Then
                                                    sql_UPC = "STYLE_CODE = '" & rowICVLUPC1.Item("STYLE_CODE") & "' And UPC_CODE is null"

                                                    Style = rowICVLUPC1.Item("STYLE_CODE")
                                                    Color = ""
                                                    Size = ""
                                                    Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "O", QTY_Scanned, sql_UPC)
                                                End If
                                            Else
                                                Dim iResponse As MsgBoxResult = MsgBox("UPC Code is not in Carton, do you still wish to add upc to audit", MsgBoxStyle.YesNo, "Proceed")
                                                If iResponse = MsgBoxResult.Yes Then
                                                    sql_UPC = "UPC_CODE = '" & rowICVLUPC1.Item("UPC_CODE") & "'"

                                                    Style = rowICVLUPC1.Item("STYLE_CODE")
                                                    Color = rowICVLUPC1.Item("COLOR_CODE")
                                                    Size = rowICVLUPC1.Item("SIZE_CODE")
                                                    Add_WHTAUDT2(rowICVLUPC1.Item("UPC_CODE"), TRAN_LNO, Style, Color, Size, "N", QTY_Scanned, sql_UPC)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If


        txtUPC.Text = ""
        txtUPC.Focus()
        Calc_Detail()
    End Sub

    Sub Update_SOWCART2(UPC As String, Qty_Packed As Integer, Qty As Integer, Style As String, Color As String, size As String, sql_UPC As String)
        If Qty_Packed = Qty Then
            ASCDATA1.DeleteRows("SOTCART2", sql_UPC)
        Else
            dst.Tables("SOTCART2").Select(sql_UPC)(0).Item("QTY_PACKED") = Val(dst.Tables("SOTCART2").Select(sql_UPC)(0).Item("QTY_PACKED") & "") - 1
        End If
    End Sub
    Sub Add_WHTAUDT2(UPC_CODE As String, Tran_Lno As Integer, Style As String, color As String, size As String, item_status As String, QTY As Int64, sql_UPC As String)

        If dst.Tables("WHTAUDT2").Select(sql_UPC).Length <> 0 Then
            dst.Tables("WHTAUDT2").Select(sql_UPC)(0).Item("SCAN_QTY") = Val(dst.Tables("WHTAUDT2").Select(sql_UPC)(0).Item("SCAN_QTY") & "") + QTY
        Else
            Dim rowWHTAUDT2 As DataRow = dst.Tables("WHTAUDT2").NewRow
            With rowWHTAUDT2
                .Item("TRAN_NO") = Tran_No
                .Item("TRAN_LNO") = Tran_Lno
                .Item("UPC_CODE") = UPC_CODE
                .Item("STYLE_CODE") = Style
                .Item("COLOR_CODE") = color
                .Item("SIZE_CODE") = size
                .Item("ITEM_STATUS") = item_status
                .Item("SCAN_QTY") = QTY
                If dst.Tables("WHTCARTC").Select(sql_UPC).Length <> 0 Then
                    .Item("ORIG_QTY") = Val(dst.Tables("WHTCARTC").Select(sql_UPC)(0).Item("QTY_PACKED") & "")
                Else
                    .Item("ORIG_QTY") = 0
                End If

            End With
            dst.Tables("WHTAUDT2").Rows.Add(rowWHTAUDT2)
        End If
    End Sub
    Sub Check_Duplicate_UPC()
        'Sql = "Select * From SOTCSTY1 Where CUST_UPC = '" & UPC_CODE & "'"
        'dynSOTCSTY1 = OraD.CreateDynaset(Sql, 8&)

        'If Not dynSOTCSTY1.EOF Then
        '    SOTCSTY1_Flag = True
        'Else
        '    SOTCSTY1_Flag = False
        'End If

        'Sql = "Select * from ICVLUPC1"
        'Sql = Sql & " Where UPC_CODE = '" & UPC_CODE & "'"
        'dynICVLUPC1 = OraD.CreateDynaset(Sql, 8&)
        'If Not dynICVLUPC1.EOF Then
        '    ICVLUPC1_Flag = True
        'Else
        '    ICVLUPC1_Flag = False
        'End If

        'If SOTCSTY1_Flag = True And ICVLUPC1_Flag = True Then
        '    If dynSOTCSTY1.Fields("STYLE_CODE").Value <> dynICVLUPC1.Fields("STYLE_CODE").Value Then
        '        Sql = "Select * from ICTSTYL1 Where STYLE_CODE in "
        '        Sql = Sql & " ('" & dynSOTCSTY1.Fields("STYLE_CODE").Value & "','" & dynICVLUPC1.Fields("STYLE_CODE").Value & "')"
        '        dynICTSTYL1 = OraD.CreateDynaset(Sql, 8&)

        '        Do While Not dynICTSTYL1.EOF
        '            If dynICTSTYL1.Fields("STYLE_CODE").Value = dynSOTCSTY1.Fields("STYLE_CODE").Value Then
        '                If dynICTSTYL1.Fields("STYLE_STATUS").Value <> "A" Then
        '                    SOTCSTY1_Flag = False
        '                End If
        '            End If
        '            If dynICTSTYL1.Fields("STYLE_CODE").Value = dynICVLUPC1.Fields("STYLE_CODE").Value Then
        '                If dynICTSTYL1.Fields("STYLE_STATUS").Value <> "A" Then
        '                    ICVLUPC1_Flag = False
        '                End If
        '            End If
        '            dynICTSTYL1.MoveNext()
        '        Loop
        '        dynICTSTYL1.Close()
        '        If ICVLUPC1_Flag = True And SOTCSTY1_Flag = True Then
        '            MsgBox("UPC Code " & UPC_CODE & " is Active for Style " & dynICVLUPC1.Fields("STYLE_CODE").Value & " and Style " & dynSOTCSTY1.Fields("STYLE_CODE").Value & ". Cannot Proceed", vbOKOnly, "Cannot Proceed")
        '            txtITEMUPC.Text = ""
        '            txtITEMUPC.SetFocus()
        '            Exit Sub
        '        End If
        '        If ICVLUPC1_Flag = False And SOTCSTY1_Flag = False Then
        '            MsgBox("UPC Code " & UPC_CODE & " is In-Active for Style " & dynICVLUPC1.Fields("STYLE_CODE").Value & " and Style " & dynSOTCSTY1.Fields("STYLE_CODE").Value & ". Cannot Proceed", vbOKOnly, "Cannot Proceed")
        '            txtITEMUPC.Text = ""
        '            txtITEMUPC.SetFocus()
        '            Exit Sub
        '        End If
        '    End If
        'End If

    End Sub
    Sub Calc_Detail()
        dst.Tables("SOTSTATI").Select("ITEM_CODE = '1'")(0).Item("ITEM_QTY") = Carton_Qty
        dst.Tables("SOTSTATI").Select("ITEM_CODE = '2'")(0).Item("ITEM_QTY") = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "") & "")


        dst.Tables("SOTSTATC").Select("STATUS_CODE = '1'")(0).Item("STATUS_QTY") = Val(dst.Tables("WHTAUDT2").Compute("SUM(GOOD_QTY)", "ITEM_STATUS = 'G'") & "")
        dst.Tables("SOTSTATC").Select("STATUS_CODE = '2'")(0).Item("STATUS_QTY") = Val(dst.Tables("WHTAUDT2").Compute("SUM(SURPLUS_QTY)", "ITEM_STATUS = 'G'") & "")
        dst.Tables("SOTSTATC").Select("STATUS_CODE = '3'")(0).Item("STATUS_QTY") = Val(dst.Tables("WHTAUDT2").Compute("SUM(SCAN_QTY)", "ITEM_STATUS = 'N'") & "")
        dst.Tables("SOTSTATC").Select("STATUS_CODE = '4'")(0).Item("STATUS_QTY") = Val(dst.Tables("WHTAUDT2").Compute("SUM(SCAN_QTY)", "ITEM_STATUS = 'U'") & "")
    End Sub

End Class