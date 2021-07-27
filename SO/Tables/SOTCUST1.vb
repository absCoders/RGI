Imports Microsoft.Office.Interop
Public Class SOTCUST1
    Dim optCUST_STATUS_Value As String
    Dim OPTCUST_DISC_PCT_EXTRA_Value As String
    Dim sqlARTSREP1 As String
    Dim Remote As New REMOTE(Me)
    Dim IsUSA As Boolean = False
    Dim DiscountsLocked As Boolean = True
    Dim SB As New System.Text.StringBuilder With {.Length = 0}

#Region "ABS Satandards"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            ASCMAIN1.sql = "Select ARTCUST2.* " _
                & " from ARTCUST2 " _
                & " WHERE ARTCUST2.CUST_ADDR_TYPE = 'MK'" _
                & " AND ARTCUST2.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, True, "V", 3)
            .Tables("ARTCUST2").Columns.Add("LAST_VERIFIED", GetType(System.DateTime))

            ASCMAIN1.sql = "Select ARTCUSTD.* " _
                & " from ARTCUSTD " _
                & " where ARTCUSTD.CUST_CODE = :PARM1" _
                & " and NVL(CONTACT_NOTE,'NULL') <> 'DELETED'"
            Create_TDA(.Tables.Add, "ARTCUSTD", "**", 0, True, "V", 2)


            sqlARTSREP1 = "Select ARTSREP1.*, SOTSDIV1.SALES_DIVISION_NAME, SOTSREP1.SREP_NAME" _
            & " from SOTSDIV1,SOTSREP1,ARTSREP1" _
            & " where SOTSDIV1.SALES_DIVISION_CODE = ARTSREP1.SALES_DIVISION_CODE" _
            & "   and SOTSREP1.SREP_CODE = ARTSREP1.SREP_CODE"
            ASCMAIN1.sql = sqlARTSREP1 _
            & "  and ARTSREP1.CUST_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ARTSREP1", "**", 0, True, "V", 3)

            ASCMAIN1.sql = "SELECT SOTORDR1.*,  TO_CHAR(ORDR_DATE, 'YYYY') AS YEAR FROM SOTORDR1 "
            ASCMAIN1.sql += " WHERE SOTORDR1.CUST_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V")
            .Tables("SOTORDRX").Columns.Add("TORDR", GetType(System.String))

            ASCMAIN1.sql = "SELECT * FROM ARTCUSCC WHERE CUST_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ARTCUSCC", "**", 0, False, "V")

            SB.Length = 0
            SB.AppendLine("SELECT * FROM ARTCUSTQ")
            SB.AppendLine(" WHERE CUST_CODE = :PARM1")
            SB.AppendLine(" AND CUST_ADDR_CODE = :PARM2")
            ASCMAIN1.sql = SB.ToString()
            Create_TDA(.Tables.Add, "ARTCUSTQ", "**", 0, True, "VV", 2)

            SB.Length = 0
            SB.AppendLine("SELECT * FROM ARTCUST2")
            SB.AppendLine(" WHERE CUST_CODE = :PARM1")
            SB.AppendLine(" AND CUST_ADDR_CODE = :PARM2")
            SB.AppendLine("  AND CUST_ADDR_TYPE = 'MK'")
            ASCMAIN1.sql = SB.ToString()
            Create_TDA(.Tables.Add, "ARTCUSX2", "**", 0, False, "VV", 3)
            .Tables("ARTCUSX2").Columns.Add("VERIFIED", GetType(System.String))
        End With

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        grdARTCUSTD.DataSource = dst.Tables("ARTCUSTD")
        grdARTSREP1.DataSource = dst.Tables("ARTSREP1")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdARTCUSCC.DataSource = dst.Tables("ARTCUSCC")

        With grdARTCUST2.DisplayLayout.Bands(0)
        End With

        Create_Summary(grdSOTORDRX, "TORDR", "Sum", "", "###,##0.00")
        Sort_grdColumns(grdSOTORDRX, "ORDR_DATE, ORDR_GROUP_NO, ORDR_NO".ToLower(), False)
        grdSOTORDRX.DisplayLayout.Bands(0).Columns("TORDR").Format = "###,##0.00"
        ASCMAIN1.Add_Value_List(grdSOTORDRX, "ORDR_STATUS", , New String() {":", "L:Laptop", "Q:Quote", "C:Cancelled"})

        grdARTCUST2.DisplayLayout.Bands(0).Columns("LAST_VERIFIED").Format = "MM/dd/yy"

        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        'Call InitializeControls(Me)
        ASCMAIN1.Add_Value_List(grdARTCUST2, "CUST_ADDR_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive", "C:Closed"})
        UltraExplorerBar1.Groups("Screen Mode").Items("Audit Trail").Visible = False
        UltraExplorerBar1.Groups("Screen Mode").Items("Multiple Record View").Visible = False
        UltraExplorerBar1.Groups("Begin New Records With").Enabled = False
        UltraExplorerBar1.Groups("Begin New Records With").Expanded = False
        UltraExplorerBar1.Groups("Default Mode").Enabled = False
        UltraExplorerBar1.Groups("Default Mode").Expanded = False

        SplitContainer1.SplitterDistance = SplitContainer1.Parent.Width * 0.25
        SplitContainer3.SplitterDistance = SplitContainer3.Parent.Width * 0.5

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTSREP1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdARTCUST2, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Add Ship-to From Master", "Verify Ship-tos")
        Load_Popup_Menu(grdARTCUSTD, "BB", "Send E-mail", "Make Contact From Main")
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

            Case "grdARTSREP1"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdARTCUSTD"
                tlb_btn = DirectCast(tlb_pop.Tools("Make Contact From Main"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdARTCUST2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Ship-to From Master"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
                tlb_btn = DirectCast(tlb_pop.Tools("Verify Ship-tos"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdARTSREP1" Then
                    'Add_Codes(grdARTSREP1, "SOTSDIV1", "SALES_DIVISION_CODE", "Divisions")
                    Add_Codes(grdARTSREP1, "SOTSREP1", "SREP_CODE", "Sales Reps")
                End If
            Case "Add Ship-to From Master"
                If grd.Name = "grdARTCUST2" Then
                    MakeARTCUST2()
                End If
            Case "Send E-mail"
                If grd.Name = "grdARTCUSTD" Then
                    If grdARTCUSTD.ActiveRow.Cells("CONTACT_EMAIL").Text.Length > 0 Then
                        SendEmail(grdARTCUSTD.ActiveRow.Cells("CONTACT_EMAIL").Text)
                    Else
                        MsgBox("The Selected Row Has No Valid Address", MsgBoxStyle.Critical, "E-Mail")
                    End If
                End If
            Case "Make Contact From Main"
                MakeContactFromMain()
            Case "Verify Ship-tos"
                If grdARTCUST2.ActiveRow.Cells("CUST_ADDR_CODE").Text.Length > 0 Then
                    shipToVerify(True, grdARTCUST2.ActiveRow.Cells("CUST_ADDR_CODE").Text)
                    SetVerified()
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Public Overrides Function Remote_Control( _
        ByVal command As String, _
        Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "Edit"
                Absx1.txtFor("CUST_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                ' EMsg &= vbCr & "Creating & Editing Customers Is Not Supported Until Back Office Is Live."
                'If Absx1.txtFor("CUST_CODE").Text.Length > 0 Then
                '    EMsg &= vbCr & "Customer Code Must Be Left Blank When Creating New Customers"
                '    Absx1.txtFor("CUST_CODE").Text = ""
                'End If
                Generate_Cust_Code()
            Case "Edit"
                ' EMsg &= vbCr & "Creating & Editing Customers Is Not Supported Until Back Office Is Live."
            Case "Update", "Save"
                If optCUST_PRICE_TIER.Value = "SP" Then
                    Dim DISC_PCT_MSG As String = DISC_PCT_CHECK()
                    If DISC_PCT_MSG.Length > 0 Then
                        EMsg &= vbCr & DISC_PCT_MSG
                    End If
                End If
                'Dim CUST_DISC_PCT As Double = Val(Absx1.numFor("CUST_DISC_PCT").Value & "")
                'If CUST_DISC_PCT <> 0 And (CUST_DISC_PCT < 0 Or CUST_DISC_PCT > 60) Then
                '    EMsg &= vbCr & "Customer Disc% Only Be Between 0% and 60%"
                'End If

                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                Dim rowTATSTATE As DataRow = LookUp("TATSTATE", Absx1.txtFor("CUST_STATE").Text)
                If rowTATSTATE Is Nothing Then
                    If IsUSA Then
                        EMsg &= vbCr & "Invalid Value Specified for State"
                    End If
                End If

                ' DO THE FOLLOWING FOR MANDATORY CODES
                Dim NoBlanks As New Dictionary(Of String, String)
                NoBlanks.Add("FRT_TERMS", "FRT Terms")
                NoBlanks.Add("SHIP_VIA_CODE", "Ship Via")
                NoBlanks.Add("CUST_ADDR1", "Address")
                NoBlanks.Add("CUST_CITY", "City")
                NoBlanks.Add("CUST_COUNTRY", "Country")
                If IsUSA Then
                    NoBlanks.Add("CUST_STATE", "State")
                    NoBlanks.Add("CUST_ZIP_CODE", "Zip Code")
                    If Absx1.CtlFor("CUST_ZIP_CODE").Text.Trim().Length > 0 And Absx1.CtlFor("CUST_ZIP_CODE").Text.Trim().Length < 5 Then
                        EMsg &= vbCr & "The Customer Zip Code Must Be At Least 5 Digits"
                    End If
                End If
                NoBlanks.Add("CUST_CONTACT", "Contact Name")
                NoBlanks.Add("CUST_NAME", "Customer Name")
                For Each NoBlanksItem As KeyValuePair(Of String, String) In NoBlanks
                    If Absx1.txtFor(NoBlanksItem.Key).Text.Trim().Length = 0 Then
                        EMsg &= vbCr & String.Format("The {0} Field Can Not Be Left Blank", NoBlanksItem.Value)
                    End If
                Next
                If Absx1.CtlFor("CUST_PHONE").Text.Trim().Length = 0 Then
                    EMsg &= vbCr & "The Contact Phone Number Field Can Not Be Left Blank"
                End If

                Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                If rowSOTSREP1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Value entered for Sales Rep Code"
                End If

                For Each ROW As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTCUSTD"), New String() {"CONTACT_TYPE"}).Rows
                    Dim CONTACT_TYPE As String = ROW.Item("CONTACT_TYPE") & ""
                    Dim sqlw As String = "CONTACT_TYPE = '" & CONTACT_TYPE & "'"
                    Dim c As Integer = Val(dst.Tables("ARTCUSTD").Compute("COUNT(CONTACT_NO)", sqlw & " and CONTACT_PRIMARY = '1'") & "")
                    If c > 1 Then
                        EMsg &= vbCr & "Cannot have > 1 Primary Contact of any Type (see Type " & CONTACT_TYPE & ")"
                    ElseIf c = 0 Then
                        Dim rows() As DataRow = dst.Tables("ARTCUSTD").Select(sqlw)
                        If rows.Length = 1 Then
                            rows(0).Item("CONTACT_PRIMARY") = "1"
                        Else
                            EMsg &= vbCr & "You must select a Primary Contact for each Type of Contact (see Type " & CONTACT_TYPE & ")"
                        End If
                    End If
                Next

                Dim BadStates As Boolean = False
                Dim BadCountries As Boolean = False
                For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select()
                    If Not BadStates Then
                        Dim rowTATSTATE2 As DataRow = LookUp("TATSTATE", rowARTCUST2.Item("CUST_STATE") & "")
                        If rowTATSTATE2 Is Nothing Then
                            If IsUSA Then
                                EMsg &= vbCr & "Invalid State found in Ship-To's"
                                BadStates = True
                            End If
                        End If
                    End If
                    If Not BadCountries Then
                        If Len(rowARTCUST2.Item("CUST_COUNTRY") & "") > 0 Then
                            Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", rowARTCUST2.Item("CUST_COUNTRY") & "")
                            If rowTATCNTRY Is Nothing Then
                                EMsg &= vbCr & "Invalid Country found in Ship-To's"
                                BadCountries = True
                            End If
                        Else
                            EMsg &= vbCr & "Blank Country found in Ship-To's"
                        End If
                    End If
                Next

                If EMsg.Length = 0 Then
                    If dst.Tables.Item("ARTCUST2").Rows.Count = 0 Then
                        Dim Msg As String = ""
                        Msg = "There Are No Ship-To's Defined."
                        Msg = Msg & vbCrLf & "Would You Like Me To Create One From The Bill-To?"
                        Dim iResult As MsgBoxResult = MsgBox(Msg, MsgBoxStyle.YesNo, "Create New Ship-To?")
                        If iResult = MsgBoxResult.Yes Then
                            MakeARTCUST2()
                        Else
                            EMsg &= vbCr & "Please Create At Least One Ship-To For This Customer"
                        End If
                    End If
                End If

                If Absx1.optFor("CUST_XMIT_INV_VIA").Value = "E" Then
                    If Absx1.txtFor("CUST_INV_EMAIL").Text.Trim().Length = 0 Then
                        EMsg &= vbCr & "You Must Provide An Invoice E-mail."
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""

        grdARTCUST2.UpdateData()
        grdARTCUSTD.UpdateData()

        Update_Record_TDA("ARTCUST2")
        Update_Record_TDA("ARTCUSTD")
        Update_Record_TDA("ARTSREP1")

        If EntryMode = "New" Then
            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
            ASCDATA1.ExecuteSQL("Delete from TATCTLN3 where CTL_NO_TYPE = 'ARTCUST1.CUST_CODE' and CTL_NO = '" & CUST_CODE & "'")
        End If

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        Else
            ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        End If

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        For Each TABLE_NAME As String In New String() {"ARTCUST1", "ARTCUST2", "ARTCUSTD"}
            ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & "_L where CUST_CODE = '" & CUST_CODE & "'")
            ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & "_H where CUST_CODE = '" & CUST_CODE & "'")
            ASCDATA1.ExecuteSQL("Insert into " & TABLE_NAME & "_L Select * from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'")
            ASCDATA1.ExecuteSQL("Insert into " & TABLE_NAME & "_H Select * from " & TABLE_NAME & " where CUST_CODE = '" & CUST_CODE & "'")
        Next

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not Remote.IsUserSuper Then
                    sql_where = Remote.SQLWhere & " OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1_L)"
                End If
            Case "SREP_CODE", "SREP2_CODE"
                sql_where = "NVL(SREP_STATUS,'A') = 'A'"
        End Select
    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
            rowASFBASE1.Item("CUST_CREDIT_LIMIT") = Val(ROWs("ARTPARM1").Item("AR_PARM_INITIAL_CR_LIMIT") & "")
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_CRED_LIMIT_EST") = DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CUST_CREDIT_LIMIT_NOTES") = "Initial Credit Limit"
            rowASFBASE1.Item("CUST_STMT_IND") = "M"
            'rowASFBASE1.Item("TERM_CODE") = "CBD"
            rowASFBASE1.Item("TERM_CODE") = "CRED"
            rowASFBASE1.Item("POST_CODE") = ROWs("ARTPARM1").Item("AR_PARM_POST_CODE")
            rowASFBASE1.Item("CUST_STATUS") = "A"
            rowASFBASE1.Item("CUST_CREDIT_HOLD") = "1"

            rowASFBASE1.Item("SREP_CODE") = "HO"
            ASCMAIN1.sql = String.Format("SELECT SREP_CODE FROM TATUSER1 WHERE USER_ID = '{0}'", ASCMAIN1.USER_ID)
            Dim SREP_CODE As String = ASCDATA1.GetDataValue
            If SREP_CODE <> "" Then
                rowASFBASE1.Item("SREP_CODE") = SREP_CODE
            End If

            Absx1.txtFor("SREP_CODE").ReadOnly = False

            rowASFBASE1.Item("WHSE_CODE") = "MS"
            rowASFBASE1.Item("CUST_PRICE_TIER") = "PC"
            rowASFBASE1.Item("CUST_PRICE_TIER_PVC") = "PC"
            If Format(DATETIME_STAMP.Date, "MM/DD/YYYY") <> "01/01/0001" Then
                rowASFBASE1.Item("CUST_STATUS_DATE") = Now.Date ' DATETIME_STAMP.Date
            End If
            rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        Else
            Absx1.txtFor("SREP_CODE").ReadOnly = True
        End If

        EnforceConstraints(False)
        Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text)
        Fill_Records("ARTCUSTD", Absx1.txtFor("CUST_CODE").Text)
        Fill_Records("ARTSREP1", Absx1.txtFor("CUST_CODE").Text)
        Fill_Records("SOTORDRX", Absx1.txtFor("CUST_CODE").Text)
        Fill_Records("ARTCUSCC", Absx1.txtFor("CUST_CODE").Text)
        CalculateOrderTotalX()
        SetVerified()

        EnforceConstraints(True)
        Setup_splUpperRight()
    End Sub

    Private Sub SetVerified()
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select()
            Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text.ToString & String.Empty
            Dim CUST_ADDR_CODE As String = rowARTCUST2.Item("CUST_ADDR_CODE").ToString & String.Empty
            SB.Length = 0
            SB.AppendLine("SELECT LAST_DATE")
            SB.AppendLine("FROM ARTCUSTQ")
            SB.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
            SB.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
            ASCMAIN1.sql = SB.ToString()
            Dim LAST_VERIFIED As String = ASCDATA1.GetDataValue
            If IsDate(LAST_VERIFIED) Then
                rowARTCUST2.Item("LAST_VERIFIED") = CDate(LAST_VERIFIED)
            End If
        Next
        grdARTCUST2.UpdateData()
        grdARTCUST2.Refresh()

    End Sub

    Private Sub CalculateOrderTotalX()
        For Each rowSOTORDRX As DataRow In dst.Tables("SOTORDRX").Select()
            ASCMAIN1.sql = String.Format("select sum(nvl(ordr_unit_price,0) * (nvl(ordr_qty,0) - nvl(ordr_qty_canc,0))) from sotordr2 where ordr_no = '{0}'", rowSOTORDRX.Item("ORDR_NO"))
            rowSOTORDRX.Item("TORDR") = Format(Val(ASCDATA1.GetDataValue), "###,##0.00")
        Next
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ARTCUST2", "ARTCUSTD", "ARTSREP1", "SOTORDRX", "ARTCUSCC"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_NAME"), Not tf)
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCUSTD, grdARTCUST2, grdARTSREP1}
            With grd.DisplayLayout.Override
                If (EntryMode = "New" Or EntryMode = "Edit") Then ' And grd.Name <> "grdARTCUST2" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
        grdARTCUST2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Screen Control").Items("Delete").Visible = False
        'UltraExplorerBar1.Groups("Screen Control").Items("Contacts").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Screen Control").Items("Contacts").Visible = False
        UltraExplorerBar1.Groups("Screen Control").Items("Set Copy-From").Visible = False
        UltraExplorerBar1.Groups("Screen Control").Items("Defaults").Visible = False

        UltraExplorerBar1.Groups("Passwords").Visible = False

        If EntryMode = "Edit" Or EntryMode = "New" Then
            btnUsePassword.Enabled = True
        Else
            btnUsePassword.Enabled = False
        End If
        If ASCMAIN1.USER_SECURITY_CODEs.Contains("X1") Then
            btnShowPassword.Enabled = True
        Else
            btnShowPassword.Enabled = False
        End If

        If (EntryMode = "New" Or EntryMode = "Edit") Then
            'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            lblDiscountLock.Text = "Customer Discounting Is Active"
            optCUST_PRICE_TIER.Enabled = True
            optCUST_PRICE_TIER_PVC.Enabled = True
            grpRGI_Pricing_PVC.Enabled = True
            panRGI_Pricing.Enabled = True
            DiscountsLocked = False
            'Else
            '    lblDiscountLock.Text = "Customer Discounting Is Locked"
            '    optCUST_PRICE_TIER.Enabled = False
            '    optCUST_PRICE_TIER_PVC.Enabled = False
            '    grpRGI_Pricing_PVC.Enabled = False
            '    panRGI_Pricing.Enabled = False
            '    DiscountsLocked = True
            'End If
        Else
            lblDiscountLock.Text = "Customer Discounting Is Locked"
            optCUST_PRICE_TIER.Enabled = False
            optCUST_PRICE_TIER_PVC.Enabled = False
            grpRGI_Pricing_PVC.Enabled = False
            panRGI_Pricing.Enabled = False
            DiscountsLocked = True
        End If
    End Sub

    Public Overrides Sub isDeleteAllowed()
        MyBase.isDeleteAllowed()
        If EMsg = "" Then
            isDeleteAllowed_Check_Aliased_Columns _
            (New String() {"ARTCUST1.CUST_BILL_TO_CUST"})
        End If
    End Sub

    Public Overrides Function Set_Contact_Info() As Boolean
        If ScreenMode Then
            CONTACT_ENTITY_KEY = Absx1.txtFor("CUST_CODE").Text
            CONTACT_ENTITY_NAME = rowASFBASE1.Item("CUST_NAME") & "" ' .txtFor("CUST_NAME").Text
        End If
        Return True
    End Function

#End Region

    Public Sub New()
        InitializeComponent()
    End Sub
#End Region

#Region "grdARTCUST2"

    Private Sub grdARTCUST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.AfterCellUpdate
        'Select Case e.Cell.Column.Key
        '    Case "CUST_CODE"
        '        Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
        '        If row IsNot Nothing Then
        '            grdARTCUST2.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
        '        End If
        'End Select
    End Sub

    Private Sub grdARTCUST2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUST2.AfterRowActivate
        With grdARTCUST2.DisplayLayout.Bands("ARTCUST2")
            If grdARTCUST2.ActiveRow.IsAddRow Then
                .Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_ADDR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdARTCUST2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTCUST2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "CUST_ADDR_CODE"
                'Dim STYLE_CODE As String = Validate_Style(e.NewValue & "")
                'If STYLE_CODE = "" Then
                '    e.Cancel = True
                'End If
        End Select

    End Sub

    Private Sub grdARTCUST2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTCUST2.BeforeExitEditMode
        If grdARTCUST2.ActiveCell IsNot Nothing Then
            With grdARTCUST2.ActiveCell
                Select Case .Column.Key
                    Case "CUST_ADDR_CODE"
                        '.EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        If .EditorResolved.IsValid Then
                            .EditorResolved.Value = .EditorResolved.Value.ToString.PadLeft(6, "0")
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdARTCUST2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUST2.BeforeRowsDeleted

    End Sub

    Private Sub grdARTCUST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        Dim EMsg As String = ""

        If e.Row.Cells("CUST_STATE").Value & "" <> "" Then
            Dim rowTATSTATE As DataRow = LookUp("TATSTATE", e.Row.Cells("CUST_STATE").Value & "")
            If rowTATSTATE Is Nothing Then
                If IsUSA Then
                    EMsg &= vbCr & "Invalid State"
                End If
                'e.Cancel = True
            End If
        End If

        If e.Row.Cells("CUST_COUNTRY").Value & "" <> "" Then
            Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", e.Row.Cells("CUST_COUNTRY").Value & "")
            If rowTATCNTRY Is Nothing Then
                EMsg &= vbCr & "Invalid Country"
                'e.Cancel = True
            End If
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Data Errors")
        End If
        'If row Is Nothing Then
        '    e.Cancel = True
        'End If

        If Not e.Cancel Then
            If e.Row.IsAddRow Then
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
                e.Row.Cells("CUST_ADDR_CODE").Value = GetNextCUST_ADDR_CODE()
            End If
            If e.Row.Cells("CUST_ADDR_CODE").Value = "" Then
                e.Row.Cells("CUST_ADDR_CODE").Value = GetNextCUST_ADDR_CODE()
            End If
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("CUST_ADDR_TYPE").Value = "MK"
        End If
    End Sub

    Private Sub grdARTCUST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        'Dim sql_where As String = ""
        'Call grdClickCellButton(grdARTCUST2, sql_where, True)
    End Sub

    Private Sub grdARTCUST2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUST2.Error
        grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUST2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUST2.InitializeRow
    End Sub

#End Region

#Region "grdARTCUSTD"

    Private Sub grdARTCUSTD_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.AfterCellUpdate
        Select Case e.Cell.Column.Key
            'Case "CUST_CODE"
            '    Dim row As DataRow = LookUp("ARTCUST1", e.Cell.Value)
            '    If row IsNot Nothing Then
            '        grdARTCUSTD.ActiveRow.Cells("CUST_NAME").Value = row.Item("CUST_NAME")
            '    End If
        End Select
    End Sub

    Private Sub grdARTCUSTD_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUSTD.AfterRowActivate
        With grdARTCUSTD.DisplayLayout.Bands("ARTCUSTD")
            'If grdARTCUSTD.ActiveRow.IsAddRow Then
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            'Else
            '    .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            'End If
        End With
    End Sub

    Private Sub grdARTCUSTD_AfterRowInsert(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdARTCUSTD.AfterRowInsert
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ISNULL(CONTACT_NOTE,'NULL') <> 'DELETED'"
    End Sub

    Private Sub grdARTCUSTD_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUSTD.BeforeRowsDeleted
        For Each rowARTCUSTD As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            rowARTCUSTD.Cells.Item("CONTACT_NOTE").Value = "DELETED"
        Next
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ISNULL(CONTACT_NOTE,'NULL') <> 'DELETED'"
        e.Cancel = True
    End Sub

    Private Sub grdARTCUSTD_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTD.BeforeRowUpdate

        'Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        'If row Is Nothing Then
        '    e.Cancel = True
        'End If
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            e.Row.Cells("CONTACT_NO").Value = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        End If
    End Sub

    Private Sub grdARTCUSTD_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTD.ClickCellButton
        'Dim sql_where As String = ""
        'grdClickCellButton(grdARTCUSTD, sql_where, True)
    End Sub

    Private Sub grdARTCUSTD_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTD.Error
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUSTD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTD.InitializeRow
    End Sub

#End Region

#Region "grdARTSREP1"

    Private Sub grdARTSREP1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTSREP1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                grdCodeDesc(grdARTSREP1, "SOTSDIV1", "SALES_DIVISION_CODE", "SALES_DIVISION_NAME")
            Case "SREP_CODE"
                grdCodeDesc(grdARTSREP1, "SOTSREP1", "SREP_CODE", "SREP_NAME")
        End Select
    End Sub

    Private Sub grdARTSREP1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTSREP1.BeforeRowUpdate
        Dim row As DataRow = LookUp("SOTSDIV1", e.Row.Cells("SALES_DIVISION_CODE").Text)
        If row IsNot Nothing Then row = LookUp("SOTSREP1", e.Row.Cells("SREP_CODE").Text)
        If row Is Nothing Then e.Cancel = True

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        End If
    End Sub

    Private Sub grdARTSREP1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTSREP1.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "SALES_DIVISION_CODE"
                Dim sql_where As String = "" ' Get_List_of_Codes("SOTSDIV1.SALES_DIVISION_CODE not in", "ARTSREP1", "SALES_DIVISION_CODE")
                grdClickCellButton(grdARTSREP1, sql_where, True)
            Case "SREP_CODE"
                Dim sql_where As String = "" ' Get_List_of_Codes("SOTSREP1.SREP_CODE not in", "ARTSREP1", "SREP_CODE")
                grdClickCellButton(grdARTSREP1, sql_where, True)
        End Select
    End Sub

    Private Sub grdARTSREP1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTSREP1.InitializeRow
        grd_RowColor(dst.Tables("ARTSREP1"), e.Row)
    End Sub

#End Region

#Region "Form Specific Methods"
#Region "Options"

    Private Sub optCUST_DISC_PCT_EXTRA_Click(sender As Object, e As System.EventArgs) Handles optCUST_DISC_PCT_EXTRA.Click
        OPTCUST_DISC_PCT_EXTRA_Value = optCUST_DISC_PCT_EXTRA.Value
    End Sub

    Private Sub optCUST_DISC_PCT_EXTRA_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optCUST_DISC_PCT_EXTRA.ValueChanged
        Set_Pricing_Visibility()
        'If DiscountsLocked Then
        '    If optCUST_DISC_PCT_EXTRA.Value = "2" Then
        '        optCUST_DISC_PCT_EXTRA.Value = OPTCUST_DISC_PCT_EXTRA_Value
        '        optCUST_DISC_PCT_EXTRA.Update()
        '    End If
        'End If
    End Sub

    Private Sub optCUST_PRICE_TIER_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optCUST_PRICE_TIER.ValueChanged
        Set_Pricing_Visibility()
        'If DiscountsLocked Then
        '    If SELECTION_NO = 0 Then Exit Sub
        '    Set_Pricing_Visibility()
        'Else
        '    If optCUST_PRICE_TIER_PVC.Value <> "PC" Then
        '        optCUST_DISC_PCT_EXTRA.Value = 1
        '    End If
        'End If
    End Sub

    Private Sub optCUST_STATUS_Click(sender As Object, e As System.EventArgs) Handles optCUST_STATUS.Click
        optCUST_STATUS_Value = optCUST_STATUS.Value
    End Sub

    Private Sub optCUST_STATUS_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles optCUST_STATUS.Validating
        If optCUST_STATUS.Value = "I" Then

        End If
    End Sub

    Private Sub optCUST_STATUS_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optCUST_STATUS.ValueChanged
        If optCUST_STATUS.Value = "I" Then
            optCUST_STATUS.Value = optCUST_STATUS_Value
            optCUST_STATUS.Update()
        End If
    End Sub
#End Region

#Region "Buttons"
    Private Sub btnEMail_Click(sender As System.Object, e As System.EventArgs) Handles btnEMail.Click
        If txtCustEmail.Text.Length > 0 Then
            SendEmail(txtCustEmail.Text)
        Else
            MsgBox("No E-Mail Address Specified", MsgBoxStyle.Critical, "E-MAil")
        End If
    End Sub

    Private Sub btnShowPassword_Click(sender As System.Object, e As System.EventArgs) Handles btnShowPassword.Click
        Dim OverRide As String = SOCMAIN2.getTimePassword(10)
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Customer Override"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("The Discount Override Password")
        iMSG.AppendLine("Is : " & OverRide)
        iMSG.AppendLine("It Will Expire at :" & Now.AddMinutes(10).ToShortTimeString)
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
    End Sub

    Private Sub btnUsePassword_Click(sender As System.Object, e As System.EventArgs) Handles btnUsePassword.Click
        Dim OverRide As Boolean = SOCMAIN2.useTimePassword()
        If OverRide Then
            lblDiscountLock.Text = "Customer Discounting Is Active"
            optCUST_PRICE_TIER.Enabled = True
            optCUST_PRICE_TIER_PVC.Enabled = True
            grpRGI_Pricing_PVC.Enabled = True
            panRGI_Pricing.Enabled = True
            DiscountsLocked = False
        Else
            lblDiscountLock.Text = "Customer Discounting Is Locked"
            optCUST_PRICE_TIER.Enabled = False
            optCUST_PRICE_TIER_PVC.Enabled = False
            grpRGI_Pricing_PVC.Enabled = False
            panRGI_Pricing.Enabled = False
            DiscountsLocked = True
        End If
    End Sub
#End Region

#Region "TextBox"
    Private Sub txtCUST_COUNTRY_ValueChanged(sender As System.Object, e As System.EventArgs) Handles txtCUST_COUNTRY.ValueChanged
        Dim PhoneFormat As String = "(###) ###-####"
        'Danny Says Don't Default To USA And Make Country Mandatory.
        'If txtCUST_COUNTRY.Text = "" Or txtCUST_COUNTRY.Text = "USA" Then
        If txtCUST_COUNTRY.Text = "USA" Then
            IsUSA = True
        Else
            IsUSA = False
            PhoneFormat = "##########"
        End If
        txtCUST_PHONE.InputMask = PhoneFormat
        txtCUST_FAX.InputMask = PhoneFormat
    End Sub

    Private Sub txtCUST_STATE_LostFocus(sender As Object, e As System.EventArgs) Handles txtCUST_STATE.LostFocus
        txtCUST_STATE.Text = txtCUST_STATE.Text.ToUpper
    End Sub
#End Region

#Region "Tabs"
    Private Sub tabARTCUST1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabARTCUST1.SelectedTabChanged
        Setup_splUpperRight()
        Me.Name = "ARTCUST1"
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub Generate_Cust_Code()
        Dim EMsg As String = ""

        'ASCMAIN1.sql = "Select Min (CTL_NO) from TATCTLN3 where CTL_NO_TYPE = 'ARTCUST1.CUST_CODE'"
        Dim TATCTLN3 As New TATCTLN3("ARTCUST1.CUST_CODE", Me)
        If Not IsNothing(TATCTLN3.ErrMsg) Then
            MsgBox(TATCTLN3.ErrMsg, MsgBoxStyle.OkOnly, "Problem Getting Next Order Number")
            Exit Sub
        End If
        If TATCTLN3.NumbersRemaining < 5 Then
            Dim msg As String = String.Format("You Only Have {0} Customer Numbers Left", TATCTLN3.NumbersRemaining)
            msg = msg & vbCrLf & "You Should Fetch Some More From The Transfer Screen Soon."
            MsgBox(msg, MsgBoxStyle.Critical, "Running Low On Customer Numbers")
        End If

        Dim CUST_CODE As String = TATCTLN3.Next_ctl_no
        If CUST_CODE = "" Then
            MsgBox("Perhaps you are out of Customer Numbers?" & vbCrLf & vbCrLf & "Get some more in the Data Transfer screen", MsgBoxStyle.OkOnly, "Cannot Generate a new Customer Number")
            Exit Sub
        End If
        If CUST_CODE.Length < 6 Then
            MsgBox("Problem With Customer Code Generated", MsgBoxStyle.OkOnly, "Cannot Generate a new Customer Number")
            Exit Sub
        Else
            CUST_CODE = CUST_CODE.Substring(CUST_CODE.Length - 6, 6)
        End If

        Dim row As DataRow = LookUp("ARTCUST1", CUST_CODE)
        If row IsNot Nothing Then
            EMsg = "Auto-Generated Next Customer Code (" & CUST_CODE & ") already exists"
        Else
            If Not ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE, False, True, True, 2) Then Exit Sub

            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            'Click_Command("New")
        End If

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Auto-Generate Customer")
            ASCMAIN1.MultiTask_Release(, , 2)
        End If

    End Sub

    Private Function GetNextCUST_ADDR_CODE() As String
        Dim RetVal As String = ""
        Dim NextCUST_ADDR_CODE As Integer = 0
        For Each rowARTCUST2 As DataRow In dst.Tables("ARTCUST2").Select("", "CUST_ADDR_CODE")
            If IsNumeric(rowARTCUST2.Item("CUST_ADDR_CODE")) Then
                NextCUST_ADDR_CODE = CInt(rowARTCUST2.Item("CUST_ADDR_CODE"))
            End If
        Next
        NextCUST_ADDR_CODE += 1
        RetVal = Str(NextCUST_ADDR_CODE).Trim().PadLeft(6, "0")
        Return RetVal
    End Function

    Private Sub MakeARTCUST2()
        Dim NextCUST_ADDR_CODE As String = GetNextCUST_ADDR_CODE()

        Dim newARTCUST2 As DataRow = dst.Tables("ARTCUST2").NewRow()
        newARTCUST2.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        newARTCUST2.Item("CUST_ADDR_TYPE") = "MK"
        newARTCUST2.Item("CUST_ADDR_CODE") = NextCUST_ADDR_CODE
        newARTCUST2.Item("CUST_ADDR_STATUS") = "A"
        newARTCUST2.Item("INIT_OPER") = ASCMAIN1.USER_ID
        newARTCUST2.Item("LAST_OPER") = ASCMAIN1.USER_ID
        newARTCUST2.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
        newARTCUST2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
        newARTCUST2.Item("CUST_CONTACT") = Absx1.txtFor("CUST_CONTACT").Text
        newARTCUST2.Item("CUST_PHONE") = txtCUST_PHONE.Text
        newARTCUST2.Item("CUST_EXT") = Absx1.txtFor("CUST_EXT").Text
        newARTCUST2.Item("CUST_FAX") = txtCUST_FAX.Text
        newARTCUST2.Item("CUST_EMAIL") = txtCustEmail.Text
        For Each COLNAME As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY"}
            newARTCUST2.Item(COLNAME) = Absx1.txtFor(COLNAME).Text
        Next
        dst.Tables("ARTCUST2").Rows.Add(newARTCUST2)
        grdARTCUST2.Refresh()
        grdARTCUST2.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        grdARTCUST2.Update()
        If Not IsNothing(grdARTCUST2.ActiveRow) Then
            If Not grdARTCUST2.ActiveRow Is Nothing AndAlso grdARTCUST2.ActiveRow.DataChanged Then
                grdARTCUST2.ActiveRow.CancelUpdate()
            ElseIf Not grdARTCUST2.ActiveRow IsNot Nothing AndAlso grdARTCUST2.ActiveRow.DataChanged Then
                grdARTCUST2.ActiveRow.Update()
            End If
        End If
    End Sub

    Private Sub MakeContactFromMain()
        Dim RowCount = grdARTCUSTD.Rows.Count
        Dim newARTCUSTD As DataRow = dst.Tables.Item("ARTCUSTD").NewRow
        Dim CONTACT_NO As Integer = Val(dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        With newARTCUSTD
            .Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
            .Item("CONTACT_NO") = CONTACT_NO
            .Item("CONTACT_NAME") = Absx1.txtFor("CUST_CONTACT").Text
            .Item("CONTACT_EMAIL") = Absx1.txtFor("CUST_EMAIL").Text
            .Item("CONTACT_PHONE") = txtCUST_PHONE.Text
            .Item("CONTACT_EXT") = Absx1.txtFor("CUST_EXT").Text
            .Item("CONTACT_FAX") = txtCUST_FAX.Text
            If CONTACT_NO = 1 Then
                .Item("CONTACT_TYPE") = "B"
                .Item("CONTACT_PRIMARY") = "1"
            Else
                .Item("CONTACT_TYPE") = "M"
                .Item("CONTACT_PRIMARY") = "0"
            End If
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            '.Item("CONTACT_CELL") = "XXXXXX"
            '.Item("CONTACT_TITLE") = "XXXXXX"
            '.Item("CONTACT_NOTE") = "XXXXXX"
        End With
        dst.Tables.Item("ARTCUSTD").Rows.Add(newARTCUSTD)
        grdARTCUSTD.Refresh()
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.ExitEditMode)
        grdARTCUSTD.Update()
        If Not IsNothing(grdARTCUSTD.ActiveRow) Then
            If Not grdARTCUSTD.ActiveRow Is Nothing AndAlso grdARTCUSTD.ActiveRow.DataChanged Then
                grdARTCUSTD.ActiveRow.CancelUpdate()
            ElseIf Not grdARTCUSTD.ActiveRow IsNot Nothing AndAlso grdARTCUSTD.ActiveRow.DataChanged Then
                grdARTCUSTD.ActiveRow.Update()
            End If
        End If
    End Sub

    Private Sub SendEmail(ByVal EAddress As String)
        Dim OutlookEmail As New Email()
        Dim Subject As String = ""
        OutlookEmail.Message.To = EAddress
        OutlookEmail.Show()
    End Sub

    Sub Set_Pricing_Visibility()
        'Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = Not (optCUST_PRICE_TIER.Value = "FC" Or optCUST_PRICE_TIER.Value = "HC" Or optCUST_PRICE_TIER.Value = "SP")
        'lblCUST_DISC_PCT_EXTRA.Visible = Not (optCUST_PRICE_TIER.Value = "FC" Or optCUST_PRICE_TIER.Value = "HC" Or optCUST_PRICE_TIER.Value = "SP")
        'Absx1.numFor("CUST_DISC_PCT").Visible = (optCUST_PRICE_TIER.Value = "SP")
        If Not IsNothing(optCUST_PRICE_TIER.Value) Then
            Select Case optCUST_PRICE_TIER.Value
                Case "FC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "HC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "PC"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = True
                    lblCUST_DISC_PCT_EXTRA.Visible = True
                    'Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = False
                    Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case "SP"
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Visible = False
                    lblCUST_DISC_PCT_EXTRA.Visible = False
                    Absx1.CtlFor("CUST_DISC_PCT_EXTRA").Text = "0"
                    Absx1.numFor("CUST_DISC_PCT").Visible = True
                    'Absx1.numFor("CUST_DISC_PCT").Value = 0
                Case Else
                    MsgBox("Error In Price Tier Option", vbOKOnly, "Please Let ABS Know")
            End Select
        End If

    End Sub

    Sub Setup_splUpperRight()
        If (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN") _
        Or (ASCMAIN1.DBS_COMPANY = "NYA" Or ASCMAIN1.DBS_SERVER = "NYA") Then
            splUpperRight.Panel1Collapsed = (tabARTCUST1.SelectedTab.Key = "Sales")
            splUpperRight.Panel2Collapsed = Not (tabARTCUST1.SelectedTab.Key = "Sales")
        Else
            splUpperRight.Panel1Collapsed = False
            splUpperRight.Panel2Collapsed = True
        End If

    End Sub

    Private Function shipToVerify(ByVal showForm As Boolean, ByVal CUST_ADDR_CODE As String) As Boolean
        'NOTE: Very Similar Code to This is also in SOFORDR0 with the same name.
        '      If you are making changes here you should consider doing it There
        '      As well.
        Dim RetVal As Boolean = True
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        dst.Tables.Item("ARTCUSTQ").Clear()
        dst.Tables.Item("ARTCUSX2").Clear()
        Dim ORDR_NO As String = "9999999999"
        Fill_Records("ARTCUSTQ", New String() {CUST_CODE, CUST_ADDR_CODE}, False)
        Fill_Records("ARTCUSX2", New String() {CUST_CODE, CUST_ADDR_CODE}, False)

        For Each rowARTCUSX2 As DataRow In dst.Tables("ARTCUSX2").Select()
            Dim Filter As String = String.Format("CUST_CODE = '{0}' AND CUST_ADDR_CODE = '{1}'", CUST_CODE, CUST_ADDR_CODE)
            If dst.Tables("ARTCUSTQ").Select(Filter).Count = 0 Then
                Dim newARTCUSTQ As DataRow = dst.Tables("ARTCUSTQ").NewRow
                newARTCUSTQ.Item("CUST_CODE") = CUST_CODE
                newARTCUSTQ.Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                newARTCUSTQ.Item("LAST_DATE") = CDate(Now().ToShortDateString)
                newARTCUSTQ.Item("LAST_OPER") = ASCMAIN1.USER_ID
                newARTCUSTQ.Item("LAST_ORDR_NO") = Null
                newARTCUSTQ.Item("RESIDENTIAL_ORDR") = "0"
                newARTCUSTQ.Item("INSIDE_REQ") = "0"
                newARTCUSTQ.Item("APPOINTMENT_REQUIRED") = "0"
                newARTCUSTQ.Item("GATE_LIFT_REQ") = "0"
                newARTCUSTQ.Item("LIMITED_ACCESS") = "0"
                newARTCUSTQ.Item("IRREGULAR_HOURS") = "0"
                newARTCUSTQ.Item("APPOINTMENT_REQUIRED") = "0"
                newARTCUSTQ.Item("BROKER") = "0"
                dst.Tables("ARTCUSTQ").Rows.Add(newARTCUSTQ)
            End If
        Next

        If showForm Then
            Dim frmSOFORDRS As New SOFORDRS(Me, Absx1.txtFor("CUST_CODE").Text, ORDR_NO)
            With frmSOFORDRS
                .ShowDialog()
            End With
        End If

        Call Update_Record_TDA("ARTCUSTQ")

        For Each rowARTCUSTQ As DataRow In dst.Tables("ARTCUSTQ").Select()
            Dim LAST_ORDR_NO As String = rowARTCUSTQ.Item("LAST_ORDR_NO").ToString & String.Empty
            Dim NOWDATE As Date = CDate(Now().ToShortDateString)
            Dim LAST_DATE As Date = CDate("01/01/1900")
            If IsDate(rowARTCUSTQ.Item("LAST_DATE").ToString & String.Empty) Then
                LAST_DATE = CDate(CDate((rowARTCUSTQ.Item("LAST_DATE").ToString & String.Empty)).ToShortDateString)
            End If
            If LAST_ORDR_NO = ORDR_NO Or LAST_DATE = NOWDATE Then
                SB.Length = 0
                SB.AppendLine("DELETE FROM ARTCUSTQ_L")
                SB.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                SB.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
                ASCMAIN1.sql = SB.ToString
                ASCDATA1.ExecuteSQL()
                SB.Length = 0
                SB.AppendLine("INSERT INTO ARTCUSTQ_L")
                SB.AppendLine("SELECT * FROM ARTCUSTQ")
                SB.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
                SB.AppendLine(String.Format("AND CUST_ADDR_CODE = '{0}'", CUST_ADDR_CODE))
                ASCMAIN1.sql = SB.ToString
                ASCDATA1.ExecuteSQL()
            Else

                RetVal = False
            End If
        Next

        Return RetVal
    End Function

    Private Function DISC_PCT_CHECK() As String
        'The Only valid Values allowed now are 52, 54, 55, 56, 57 & 59 per Danny - WR 1/23/21
        'This Function Exists in ARTCUST1 and SOTCUST1.  Make Changes To Both Or Suffer The Consequences.
        Dim RETVAL As String = ""
        Dim VALID_PCT As Decimal() = {52, 54, 55, 56, 57, 59}
        Dim VALID_PCT_STR As String = "52, 54, 55, 56, 57 & 59"

        Dim CUST_DISC_PCT As Decimal = Val(Absx1.numFor("CUST_DISC_PCT").Value & "")
        If Not VALID_PCT.Contains(CUST_DISC_PCT) Then
            RETVAL = "Disc % Can Only Be " & VALID_PCT_STR
        End If
        Return RETVAL
    End Function

#End Region
#End Region

End Class