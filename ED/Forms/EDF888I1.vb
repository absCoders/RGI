Imports System.Windows.Forms
Imports ABSolution
Imports Infragistics.Win
Imports System.IO
Imports System.Data

Public Class EDF888I1

#Region "Class Variables"

    Private Const AddCode As String = "003"
    Private Const DeleteCode As String = "002"
    Private Const ChangeCode As String = "001"
    Private wktable As String = String.Empty

    Private Const EDI_PURPOSE_CODE_ADD As String = "02"
    Private Const EDI_PURPOSE_CODE__CHANGE As String = "04"
    Private Const EDI_PURPOSE_CODE_DELETE As String = "03"

    Private Const EDI_APPLICATION_ID As String = "QG"
    Private COMPANY_CODE As String = String.Empty

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    ''' <summary>
    ''' ** Sets up required Data Tables and intializes form controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With MyBase.dst

            If ASCMAIN1.DBS_COMPANY = "VAN" Orelse ASCMAIN1.DBS_SERVER = "VAN" Then
                COMPANY_CODE = "VAN"
            ElseIf ASCMAIN1.DBS_COMPANY = "NYA" Orelse ASCMAIN1.DBS_SERVER = "NYA" Then
                COMPANY_CODE = "NYA"
            ElseIf ASCMAIN1.DBS_COMPANY = "RGI" Orelse ASCMAIN1.DBS_SERVER = "RGI" Then
                COMPANY_CODE = "RGI"
            Else
                COMPANY_CODE = ASCMAIN1.CLIENT
            End If

            Get_PARM("ASTPARM1")

            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT888O1", "*")
            Create_TDA(.Tables.Add, "EDT888O2", "*")
            Create_TDA(.Tables.Add, "EDT888OX", "*")
            .Tables("EDT888OX").Columns.Add("TRANS_TYPE", GetType(System.String))

            ASCMAIN1.sql = "Select '0' SELECTED, ICTWHSE1.*, EDTTRPM1.EDI_OUR_QUAL, EDTTRPM1.EDI_OUR_ID" _
                & " FROM ICTWHSE1, EDTTRPM1" _
                & " WHERE ICTWHSE1.WHSE_EDI_QUAL  = EDTTRPM1.EDI_TP_QUAL" _
                & " AND ICTWHSE1.WHSE_EDI_ID  = EDTTRPM1.EDI_TP_ID" _
                & " AND  EDTTRPM1.EDI_DOC_NO = '888'" _
                & " AND EDTTRPM1.EDI_OUR_QUAL IS NOT NULL" _
                & " AND EDTTRPM1.EDI_OUR_ID IS NOT NULL"
            Create_TDA(.Tables.Add, "ICTWHSE1", "**")
            Fill_Records("ICTWHSE1", String.Empty, True, ASCMAIN1.sql)

            ' NYA uses only one warehouse. that warehouse e=sends teh data to the other warehouses
            If ASCMAIN1.DBS_COMPANY = "NYA" OrElse ASCMAIN1.DBS_SERVER = "NYA" Then
                If dst.Tables("ICTWHSE1").Rows.Count > 1 Then
                    Dim WHSE_CODE As String = dst.Tables("ICTWHSE1").Compute("MIN(WHSE_CODE)", "") & String.Empty
                    For Each rowICTWHSE1 As DataRow In dst.Tables("ICTWHSE1").Select()
                        If rowICTWHSE1.Item("WHSE_CODE") & String.Empty <> WHSE_CODE Then
                            rowICTWHSE1.Delete()
                        End If
                    Next
                    dst.Tables("ICTWHSE1").AcceptChanges()
                End If
            End If

            wktable = ASCMAIN1.Temp_Table("SELECT EDIX_TP_QUAL, EDIX_TP_ID, 'XXX' TRANS_TYPE, EDIX_UPC_CASE_CODE, EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY FROM EDT888OX WHERE ROWNUM < 1")
            Create_TDA(.Tables.Add, "EDT888WK", "SELECT * FROM " & wktable)
            'INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE
            .Tables("EDT888WK").Columns.Add("INIT_OPER", GetType(System.String))
            .Tables("EDT888WK").Columns.Add("INIT_DATE", GetType(System.DateTime))
            .Tables("EDT888WK").Columns.Add("LAST_OPER", GetType(System.String))
            .Tables("EDT888WK").Columns.Add("LAST_DATE", GetType(System.DateTime))

        End With

        grdICTWHSE1.DataSource = dst.Tables("ICTWHSE1")
        Sort_grdColumns(grdICTWHSE1, "WHSE_CODE")

        grdEDT888WK.DataSource = dst.Tables("EDT888WK")

        ASCMAIN1.Add_Value_List(grdEDT888WK, "TRANS_TYPE", Nothing, New String() {":", "001:Changed", "002:Delete", "003:Add"})

        Create_Summary(grdEDT888WK, "EDIX_TP_QUAL", "Count")

     End Sub

    ''' <summary>
    ''' Clear tables and controls based on the current state of the screen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRecord()

        EnforceConstraints(False)

        dst.Tables("EDTSYSIH").Clear()
        dst.Tables("EDT888O1").Clear()
        dst.Tables("EDT888O2").Clear()
        dst.Tables("EDT888OX").Clear()
        dst.Tables("EDT888WK").Clear()

        For Each rowICTWHSE1 As DataRow In dst.Tables("ICTWHSE1").Rows
            rowICTWHSE1.Item("SELECTED") = "1"
        Next

        dst.Tables("ICTWHSE1").AcceptChanges()

        EnforceConstraints(True)
    End Sub

    ''' <summary>
    ''' Load up chnages to go into 832 files
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadRecord()
        Try
            EnforceConstraints(False)
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Create 888 Records", String.Empty)

            Dim sql As String = String.Empty
            Dim processed As List(Of String) = New List(Of String)

            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & wktable)

            For Each rowICTWHSE1 As DataRow In dst.Tables("ICTWHSE1").Select("SELECTED = '1'")
                Dim EDI_OUR_QUAL As String = rowICTWHSE1.Item("EDI_OUR_QUAL") & String.Empty
                Dim EDI_OUR_ID As String = rowICTWHSE1.Item("EDI_OUR_ID") & String.Empty

                If processed.Contains(EDI_OUR_QUAL & "," & EDI_OUR_ID) Then
                    Continue For
                End If

                processed.Add(EDI_OUR_QUAL & "," & EDI_OUR_ID)

                ' Load the new items
                sql = " SELECT '" & EDI_OUR_QUAL & "' EDIX_TP_QUAL, '" & EDI_OUR_ID & "' EDIX_TP_ID, '" & AddCode & "' TRANS_TYPE"
                sql &= ", UPC_CODE EDIX_UPC_CASE_CODE, 'VN' EDIX_ITEM_PROD_QUAL"
                sql &= ", ICTSTYC1.STYLE_CODE || DECODE(NVL(HIDE_COLOR_3PL, '0'), '1', NULL, ICTSTYC1.COLOR_CODE) EDIX_ITEM_PROD_ID"
                sql &= ", ICTSTYL1.STYLE_DESC EDIX_ITEM_DESC, NVL(ICTSTYL1.SUB_UNIT_PACK_QTY, 1) EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM ICTSTYC1, ICTSTYL1"
                sql &= " WHERE ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE"
                sql &= " AND ICTSTYL1.STYLE_STATUS= 'A'"
                sql &= " AND ICTSTYC1.STYLE_COLOR_STATUS = 'A'"
                sql &= " AND NVL(ICTSTYL1.STYLE_HIDE_FROM_3PL, '0') = '0'"
                sql &= " AND ('" & EDI_OUR_QUAL & "', '" & EDI_OUR_ID & "', ICTSTYC1.STYLE_CODE) NOT IN ("
                sql &= " SELECT EDIX_TP_QUAL, EDIX_TP_ID, EDIX_ITEM_PROD_ID FROM EDT888OX)"
                ASCDATA1.ExecuteSQL("INSERT INTO " & wktable & " " & sql)

                '   Changed Items
                sql = " SELECT EDIX_TP_QUAL, EDIX_TP_ID, '" & ChangeCode & "' TRANS_TYPE,"
                sql &= " EDIX_UPC_CASE_CODE, EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, ICTSTYL1.STYLE_DESC EDIX_ITEM_DESC, NVL(ICTSTYL1.SUB_UNIT_PACK_QTY, 1) EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM "
                sql &= " ("
                sql &= " SELECT '" & EDI_OUR_QUAL & "' EDIX_TP_QUAL, '" & EDI_OUR_ID & "' EDIX_TP_ID,"
                sql &= " UPC_CODE EDIX_UPC_CASE_CODE, 'VN' EDIX_ITEM_PROD_QUAL"
                sql &= ", ICTSTYC1.STYLE_CODE || DECODE(NVL(HIDE_COLOR_3PL, '0'), '1', NULL, ICTSTYC1.COLOR_CODE) EDIX_ITEM_PROD_ID"
                sql &= ", ICTSTYL1.STYLE_DESC EDIX_ITEM_DESC, NVL(ICTSTYL1.SUB_UNIT_PACK_QTY, 1) EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM ICTSTYC1, ICTSTYL1"
                sql &= " WHERE ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE"
                sql &= " AND ICTSTYL1.STYLE_STATUS= 'A'"
                sql &= " AND ICTSTYC1.STYLE_COLOR_STATUS = 'A'"
                sql &= " AND NVL(ICTSTYL1.STYLE_HIDE_FROM_3PL, '0') = '0'"
                sql &= " MINUS"
                sql &= "("
                sql &= " SELECT EDIX_TP_QUAL, EDIX_TP_ID, EDIX_UPC_CASE_CODE ,EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM EDT888OX"
                sql &= " WHERE EDIX_TP_QUAL = '" & EDI_OUR_QUAL & "' AND EDIX_TP_ID = '" & EDI_OUR_ID & "'"
                sql &= " Union"
                sql &= " SELECT EDIX_TP_QUAL, EDIX_TP_ID, EDIX_UPC_CASE_CODE ,EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM " & wktable
                sql &= " WHERE EDIX_TP_QUAL = '" & EDI_OUR_QUAL & "' AND EDIX_TP_ID = '" & EDI_OUR_ID & "'"
                sql &= ")"
                sql &= ") X, ICTSTYL1 where ICTSTYL1.STYLE_CODE = X.EDIX_ITEM_PROD_ID"
                ASCDATA1.ExecuteSQL("INSERT INTO " & wktable & " " & sql)

                ' Deleted Items
                sql = " SELECT '" & EDI_OUR_QUAL & "' EDIX_TP_QUAL, '" & EDI_OUR_ID & "' EDIX_TP_ID, '" & DeleteCode & "' TRANS_TYPE,"
                sql &= " UPC_CODE EDIX_UPC_CASE_CODE, 'VN' EDIX_ITEM_PROD_QUAL"
                sql &= ", ICTSTYC1.STYLE_CODE || DECODE(NVL(HIDE_COLOR_3PL, '0'), '1', NULL, ICTSTYC1.COLOR_CODE) EDIX_ITEM_PROD_ID"
                sql &= ", ICTSTYL1.STYLE_DESC EDIX_ITEM_DESC, NVL(ICTSTYL1.SUB_UNIT_PACK_QTY, 1) EDIX_SUB_UNIT_PACK_QTY"
                sql &= " FROM ICTSTYC1, ICTSTYL1"
                sql &= " WHERE ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE"
                sql &= " AND (ICTSTYL1.STYLE_STATUS= 'I'"
                sql &= " OR ICTSTYC1.STYLE_COLOR_STATUS = 'I'"
                sql &= " OR NVL(ICTSTYL1.STYLE_HIDE_FROM_3PL, '0') = '1')"
                sql &= " AND ICTSTYC1.STYLE_CODE IN ("
                sql &= " SELECT EDIX_ITEM_PROD_ID FROM EDT888OX)"
                ASCDATA1.ExecuteSQL("INSERT INTO " & wktable & " " & sql)

            Next

            Fill_Records("EDT888WK", "", True, "SELECT * FROM " & wktable)
            Sort_grdColumns(grdEDT888WK, "EDIX_TP_QUAL, EDIX_TP_ID, TRANS_TYPE, EDIX_ITEM_PROD_ID")

            If dst.Tables("EDT888WK").Rows.Count = 0 Then
                MessageBox.Show("There are no Styles to upload.", "Generate", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ClearRecord()
        End Try

        EnforceConstraints(True)

    End Sub

    ''' <summary>
    ''' Sets up screen based on the form modality, state and type of processing
    ''' </summary>
    ''' <param name="tf"></param>
    ''' <remarks></remarks>
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_Description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Reset Style").Settings.Enabled = not_iScreenMode
            If dst.Tables("EDT888WK").Rows.Count = 0 Then
                .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
            Else
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End If
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
        End With

        If ScreenMode Then
            grdICTWHSE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdEDT888WK.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            grdEDT888WK.DisplayLayout.Bands(0).Columns("TRANS_TYPE").Hidden = False
            grdEDT888WK.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            grdEDT888WK.DisplayLayout.Bands(0).Columns("INIT_OPER").Hidden = True
            grdEDT888WK.DisplayLayout.Bands(0).Columns("LAST_OPER").Hidden = True
            grdEDT888WK.DisplayLayout.Bands(0).Columns("INIT_DATE").Hidden = True
            grdEDT888WK.DisplayLayout.Bands(0).Columns("LAST_DATE").Hidden = True
        Else
            ClearRecord()
            grdICTWHSE1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdEDT888WK.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False
            grdEDT888WK.DisplayLayout.Bands(0).Columns("TRANS_TYPE").Hidden = True
            grdEDT888WK.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            grdEDT888WK.DisplayLayout.Bands(0).Columns("INIT_OPER").Hidden = False
            grdEDT888WK.DisplayLayout.Bands(0).Columns("LAST_OPER").Hidden = False
            grdEDT888WK.DisplayLayout.Bands(0).Columns("INIT_DATE").Hidden = False
            grdEDT888WK.DisplayLayout.Bands(0).Columns("LAST_DATE").Hidden = False

            ASCMAIN1.sql = "SELECT '0' SELECTED, EDIX_TP_QUAL, EDIX_TP_ID, 'XXX' TRANS_TYPE, EDIX_UPC_CASE_CODE, EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE FROM EDT888OX"
            Fill_Records("EDT888WK", String.Empty, True, ASCMAIN1.sql)
        End If

        Clear_All_Filters(grdEDT888WK)
        Show_Filter(grdASFBASEX, False)

    End Sub

    ''' <summary>
    ''' Validates data when a user selects a menu option
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim sql As String = String.Empty

        Select Case eItemKey

            Case "Generate"
                If grdICTWHSE1.Rows.Count = 0 Then
                    EMsg &= vbCr & "There are no warehouses setup to receive an 888"
                End If

                If dst.Tables("ICTWHSE1").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "There are no warehouses Selected to receive an 888"
                End If

            Case "Update"
                If Me.grdICTWHSE1.Rows.Count = 0 Then
                    EMsg = "No records to update."
                End If

            Case "Cancel"

            Case "Reset Style"
                If dst.Tables("EDT888WK").Select("SELECTED = '1'").Length = 0 Then
                    EMsg &= vbCr & "You must select at least one style to reset."
                Else
                    Dim zMsg As String = "Do you want to reset the " & dst.Tables("EDT888WK").Select("SELECTED = '1'").Length & " selected style(s)?"
                    If MessageBox.Show(zMsg, "Reset Styles", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If EMsg <> String.Empty Then
            MessageBox.Show(EMsg, "Cannot Proceed", MessageBoxButtons.OK)
        Else
            Call Proceed(eItemKey)
        End If

    End Sub

    ''' <summary>
    ''' When the user selects a menu option perform the action
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Generate"
                Me.LoadRecord()
                Me.Mode_Settings(True)

            Case "Update"
                Me.UpdateRecord()
                Me.Mode_Settings(False)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Reset Style"
                ResetStyles()

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    ''' <summary>
    ''' Updates data based on the current state of the screen and the type of processing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateRecord()

        Try
            Dim ediOutboundDocNo As String = String.Empty
            Dim EDIX_TP_QUAL As String = String.Empty
            Dim EDI_OUR_ID As String = String.Empty
            Dim EDI_MAINT_TYPE_CODE As String = String.Empty
            Dim WHSE_EDI_ID As String = String.Empty
            Dim EDI_TP_ID As String = String.Empty

            'EDT888O1.COMPANY_CODE = 'VDI'  SHOULD BE 'NYA'
            'EDT888O2.COMPANY_CODE = 'VDI'  SHOULD BE 'NYA'
            'EDT888O1.EDI_WH_ID_CODE = 'VDI' SHOULD BE 'NYA' FROM ICTWHSE1.LP_WHSE_ID
            'EDT888O1.EDI_NAME SHOULD BE ASTPARM1.AS_PARM_INST_NAMe

            For Each rowEDT888WK As DataRow In ASCDATA1.SelectDistinct("EDT888WK", New String() {"EDIX_TP_QUAL", "EDIX_TP_ID", "TRANS_TYPE"}).Rows

                EDIX_TP_QUAL = rowEDT888WK.Item("EDIX_TP_QUAL")
                EDI_OUR_ID = rowEDT888WK.Item("EDIX_TP_ID")
                EDI_MAINT_TYPE_CODE = rowEDT888WK.Item("TRANS_TYPE")
 
                Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Select("EDI_OUR_QUAL = '" & EDIX_TP_QUAL & "' AND EDI_OUR_ID = '" & EDI_OUR_ID & "'")(0)
                EDI_TP_ID = rowICTWHSE1.Item("WHSE_EDI_ID")

                ' Moved from up above
                ediOutboundDocNo = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

                Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
                rowEDTSYSIH.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = ediOutboundDocNo
                rowEDTSYSIH.Item("EDI_APPLICATION_ID") = EDI_APPLICATION_ID
                rowEDTSYSIH.Item("EDI_PROCESS_IND") = "1"
                rowEDTSYSIH.Item("EDI_OUR_ID") = EDI_OUR_ID ' 4012453780
                rowEDTSYSIH.Item("EDI_TP_ID") = EDI_TP_ID 'TAYLORED
                rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
                rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

                Dim rowEDT888O1 As DataRow = dst.Tables("EDT888O1").NewRow
                rowEDT888O1.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT888O1.Item("EDI_OUTBOUND_DOC_NO") = ediOutboundDocNo

                Select Case rowEDT888WK.Item("TRANS_TYPE") & String.Empty
                    Case AddCode
                        rowEDT888O1.Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE_ADD
                    Case DeleteCode
                        rowEDT888O1.Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE_DELETE
                    Case ChangeCode
                        rowEDT888O1.Item("EDI_PURPOSE_CODE") = EDI_PURPOSE_CODE__CHANGE
                End Select

                rowEDT888O1.Item("EDI_REFERENCE_NO") = 1
                rowEDT888O1.Item("EDI_DOC_DATE") = DATETIME_STAMP
                rowEDT888O1.Item("EDI_NAME") = ROWs("ASTPARM1").Item("AS_PARM_INST_NAME") & String.Empty
                rowEDT888O1.Item("EDI_WH_ID_CODE") = rowICTWHSE1.Item("LP_WHSE_ID") & String.Empty
  
                rowEDT888O1.Item("INIT_DATE") = DATETIME_STAMP
                rowEDT888O1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("EDT888O1").Rows.Add(rowEDT888O1)

                Dim EDI_DOC_LNO As Int16 = 0
                For Each rowDETAIL As DataRow In dst.Tables("EDT888WK").Select("EDIX_TP_QUAL = '" & EDIX_TP_QUAL & "'" _
                                                                               & " AND EDIX_TP_ID = '" & EDI_OUR_ID & "'" _
                                                                               & " AND TRANS_TYPE = '" & EDI_MAINT_TYPE_CODE & "'")
                    Dim rowEDT888O2 As DataRow = dst.Tables("EDT888O2").NewRow

                    rowEDT888O2.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT888O2.Item("EDI_OUTBOUND_DOC_NO") = ediOutboundDocNo
                    EDI_DOC_LNO += 1
                    rowEDT888O2.Item("EDI_DOC_LNO") = EDI_DOC_LNO
                    rowEDT888O2.Item("EDI_MAINT_TYPE_CODE") = EDI_MAINT_TYPE_CODE
                    rowEDT888O2.Item("EDI_EFFECTIVE_DATE") = DATETIME_STAMP.ToString("MM/dd/yyyy")
                    rowEDT888O2.Item("EDI_UPC_CASE_CODE") = rowDETAIL.Item("EDIX_UPC_CASE_CODE") & String.Empty
                    rowEDT888O2.Item("EDI_ITEM_PROD_QUAL") = rowDETAIL.Item("EDIX_ITEM_PROD_QUAL") & String.Empty
                    rowEDT888O2.Item("EDI_ITEM_PROD_ID") = rowDETAIL.Item("EDIX_ITEM_PROD_ID") & String.Empty
                    rowEDT888O2.Item("EDI_ITEM_DESC") = rowDETAIL.Item("EDIX_ITEM_DESC") & String.Empty
                    rowEDT888O2.Item("EDI_SUB_UNIT_PACK_QTY") = Val(rowDETAIL.Item("EDIX_SUB_UNIT_PACK_QTY") & String.Empty)
                    dst.Tables("EDT888O2").Rows.Add(rowEDT888O2)
                Next
            Next

            Try
                MyBase.BeginTrans()
                'Update data in EDT888OX using work table

                ASCMAIN1.sql = "DELETE FROM EDT888OX WHERE (EDIX_TP_QUAL, EDIX_TP_ID, EDIX_ITEM_PROD_ID) IN" _
                        & " (SELECT EDIX_TP_QUAL, EDIX_TP_ID, EDIX_ITEM_PROD_ID FROM " & wktable & " WHERE TRANS_TYPE = '" & DeleteCode & "')"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "INSERT INTO EDT888OX " _
                    & " (EDIX_TP_QUAL, EDIX_TP_ID, EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY, " _
                    & " EDIX_UPC_CASE_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE)" _
                    & " SELECT EDIX_TP_QUAL, EDIX_TP_ID, EDIX_ITEM_PROD_QUAL, EDIX_ITEM_PROD_ID, EDIX_ITEM_DESC, EDIX_SUB_UNIT_PACK_QTY, " _
                    & " EDIX_UPC_CASE_CODE, '" & ASCMAIN1.USER_ID & "','" & ASCMAIN1.USER_ID & "', SYSDATE,  SYSDATE" _
                    & " FROM " & wktable _
                    & " WHERE TRANS_TYPE = '" & AddCode & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = " BEGIN DECLARE CURSOR C1 IS SELECT * FROM " & wktable & " WHERE TRANS_TYPE = '" & ChangeCode & "';" _
                    & " BEGIN FOR R1 IN C1 LOOP" _
                    & " 	  UPDATE EDT888OX SET EDIX_ITEM_DESC = R1.EDIX_ITEM_DESC," _
                    & " 	    EDIX_UPC_CASE_CODE = R1.EDIX_UPC_CASE_CODE," _
                    & " 	    LAST_OPER = '" & ASCMAIN1.USER_ID & "'," _
                    & " 	    LAST_DATE = SYSDATE" _
                    & " 	    WHERE EDIX_TP_QUAL = R1.EDIX_TP_QUAL" _
                    & " 	    AND EDIX_TP_ID = R1.EDIX_TP_ID" _
                    & " 	    AND EDIX_ITEM_PROD_ID = R1.EDIX_ITEM_PROD_ID;" _
                    & "     END LOOP;" _
                    & "   END;" _
                    & " END;"
                ASCDATA1.ExecuteSQL()

                MyBase.Update_Record_TDA("EDTSYSIH")
                MyBase.Update_Record_TDA("EDT888O1")
                MyBase.Update_Record_TDA("EDT888O2")
                MyBase.Update_Record_TDA("EDT888OX")
                MyBase.CommitTrans("Update Successful")

            Catch ex As Exception
                MyBase.Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT888WK, "SS", "Show Filter", "Show GroupBox")
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
            Case "grdxxx"

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

        End Select
    End Sub

#End Region

#Region "Private Subs / Functions"

    Private Sub ResetStyles()

        Try
            MyBase.BeginTrans()
            For Each rowEDT888WK As DataRow In dst.Tables("EDT888WK").Select("SELECTED = '1'")
                ASCMAIN1.sql = "DELETE FROM EDT888OX WHERE EDIX_TP_QUAL = :PARM1" & _
                    " AND EDIX_TP_ID = :PARM2" & _
                    " AND EDIX_ITEM_PROD_QUAL = :PARM3 " & _
                    " AND EDIX_ITEM_PROD_ID = :PARM4" & _
                    " AND NVL(EDIX_UPC_CASE_CODE, '*') = :PARM5"

                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVV", New Object() {rowEDT888WK.Item("EDIX_TP_QUAL"), _
                                                                         rowEDT888WK.Item("EDIX_TP_ID"), _
                                                                         rowEDT888WK.Item("EDIX_ITEM_PROD_QUAL"), _
                                                                         rowEDT888WK.Item("EDIX_ITEM_PROD_ID"), _
                                                                         IIf(rowEDT888WK.Item("EDIX_UPC_CASE_CODE") & String.Empty <> String.Empty, rowEDT888WK.Item("EDIX_UPC_CASE_CODE"), "*")})

                rowEDT888WK.Delete()
            Next
            dst.Tables("EDT888WK").AcceptChanges()
            MyBase.CommitTrans("Successful Deletions.")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Form Controls"

#End Region

 End Class