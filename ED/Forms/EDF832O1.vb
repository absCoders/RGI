Imports System.Windows.Forms
Imports ABSolution
Imports Infragistics.Win
Imports System.IO
Imports System.Data

Public Class EDF832O1

#Region "Class Variables"

    Private companyCode As String = ASCMAIN1.DBS_COMPANY
    Private Const ediApplicationId As String = "SC"

    Private catalogueItemSql As String = String.Empty
    Private compareFields As String = String.Empty
    Private sqlEDT832O2 As String = String.Empty
    Private sqlICTITEM1 As String = String.Empty

    Private Const itemAdded As String = "02"
    Private Const itemDeleted As String = "03"
    Private Const itemChanged As String = "04"

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    ''' <summary>
    ''' ** Sets up required Data Tables and intializes form controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Get_PARM("ASTPARM1")
            Create_TDA(.Tables.Add, "EDTSYSIH", "*")
            Create_TDA(.Tables.Add, "EDT832O1", "*", 0, True, String.Empty, 0, String.Empty)
            Create_TDA(.Tables.Add, "EDT832O2", "*")
            Create_TDA(.Tables.Add, "EDT832OX", "*")

            ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDTTRPM1.EDI_DOC_NO = '832' AND EDI_STATUS IN ('P', 'T')"
            Create_TDA(.Tables.Add, "EDTTRPM1", "**")
            Fill_Records("EDTTRPM1", String.Empty, True, ASCMAIN1.sql)

            .Relations.Add("EDT832O1_EDT832O2", _
                           New DataColumn() {dst.Tables("EDT832O1").Columns("COMPANY_CODE"), dst.Tables("EDT832O1").Columns("EDI_OUTBOUND_DOC_NO"), dst.Tables("EDT832O1").Columns("EDI_CATALOG_VERSION")}, _
                           New DataColumn() {dst.Tables("EDT832O2").Columns("COMPANY_CODE"), dst.Tables("EDT832O2").Columns("EDI_OUTBOUND_DOC_NO"), dst.Tables("EDT832O2").Columns("ITEM_SELECTION_CODE")})

        End With


        ASCMAIN1.Add_Value_List(grdEDT832O1, "EDI_PURPOSE_CODE", Nothing, New String() {":", "03:Delete", "02:Add", "04:Change"})

        ASCMAIN1.sql = "  Select 'COMPANY_CODE' COMPANY_CODE, 'EDI_OUR_ID' EDI_OUR_ID, 'EDI_TP_ID' EDI_TP_ID" _
            & " , ICTSTYC1.STYLE_CODE ITEM_CODE, ICTSTYL1.STYLE_DESC ITEM_DESC" _
            & " , NVL(ICTSTYC1.CATALOG_SELECTION_CODE, '001') ITEM_SELECTION_CODE" _
            & " , ICTSTYL1.STYLE_RETAIL ITEM_RETAIL_PRICE" _
            & " , NVL(ICTSTYC1.NRF_COLOR_CODE, '000') ITEM_COLOR_CODE, NVL(ICTCOLRN.NRF_COLOR_DESC, NRF_COLOR_DESC_CUSTOM) ITEM_COLOR_DESC" _
            & " , NVL(ICTSTYC1.NRF_SIZE_CODE, '00000') ITEM_SIZE_CODE, NVL(ICTSIZEN.NRF_SIZE_DESC, NRF_SIZE_DESC_CUSTOM) ITEM_SIZE_DESC" _
            & " , ICTSTYL1.STYLE_SO_QTY_MIN ITEM_SO_QTY_MIN, ICTSTYL1.SUB_UNIT_PACK_QTY ITEM_SO_QTY_MULT" _
            & " , ICTSTYC1.UPC_CODE ITEM_UPC, NULL ITEM_EAN" _
            & " , LPAD(ICTSTYC1.UPC_CODE, 14, '0')  ITEM_GTIN" _
            & " , DECODE(NVL(STYLE_STATUS, 'A'), 'I', NVL(NVL(ICTSTYL1.LAST_DATE, ICTSTYL1.INIT_DATE), SYSDATE), NULL) EDI_DISCONTINUE_DATE" _
            & " , NULL HAZARD_CODE,  ICTSLCT1.CATALOG_SELECTION_DESC" _
            & "  from ICTSTYC1, ICTSTYL1, ICTCOLRN, ICTSIZEN, ICTSLCT1" _
            & "  where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
            & "  AND NVL(ICTSTYC1.NRF_COLOR_CODE, '000') = ICTCOLRN.NRF_COLOR_CODE (+)" _
            & "  and NVL(ICTSTYC1.NRF_SIZE_CODE, '00000') = ICTSIZEN.NRF_SIZE_CODE (+)" _
            & "  and (NVL(ICTSTYC1.HIDE_FROM_CAT, '0') = '0' AND NVL(ICTSTYL1.STYLE_HIDE_FROM_CAT, '0') = '0')" _
            & "  and ICTSTYC1.UPC_CODE IS NOT NULL" _
            & "  and NVL(ICTSTYC1.CATALOG_SELECTION_CODE, '001') = ICTSLCT1.CATALOG_SELECTION_CODE"

        catalogueItemSql = ASCMAIN1.sql

        compareFields = " ITEM_UPC, ITEM_CODE, ITEM_COLOR_CODE, ITEM_SIZE_CODE, ITEM_SELECTION_CODE "

        ASCMAIN1.sql = "Select " & compareFields & vbCr _
           & "  from EDT832OX " & vbCr _
           & "  where COMPANY_CODE = 'COMPANY_CODE'" _
           & "  and EDI_OUR_ID = 'EDI_OUR_ID'" _
           & "  and EDI_TP_ID = 'EDI_TP_ID'"
        sqlEDT832O2 = ASCMAIN1.sql

        sqlICTITEM1 = "SELECT " & compareFields & " FROM (" & catalogueItemSql & ")"

        Create_Summary(grdEDT832O1, "EDI_DOC_LNO", "Count", "EDT832O1_EDT832O2")

        Me.grdEDT832O1.DataSource = dst.Tables("EDT832O1")
        Show_Filter(grdEDT832O1, True)

    End Sub

    ''' <summary>
    ''' Clear tables and controls based on the current state of the screen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRecord()

        EnforceConstraints(False)

        dst.Tables("EDTSYSIH").Clear()
        dst.Tables("EDT832O1").Clear()
        dst.Tables("EDT832O2").Clear()
        dst.Tables("EDT832OX").Clear()

        EnforceConstraints(True)
    End Sub

    ''' <summary>
    ''' Load up changes to go into 832 files
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadRecord()

        ASCMAIN1.Progress("Create 832 Records", String.Empty)

        Dim sql As String = String.Empty
        Dim tblData As DataTable = Nothing

        Try
            EnforceConstraints(False)

            dst.Tables("EDT832OX").Rows.Clear()
            For Each rowEDTTRPM1 As DataRow In dst.Tables("EDTTRPM1").Select()
                Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
                Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

                ASCMAIN1.sql = "Select * from EDT832OX where EDI_OUR_ID = '" & EDI_OUR_ID & "' AND EDI_TP_ID = '" & EDI_TP_ID & "'"
                Fill_Records("EDT832OX", String.Empty, False, ASCMAIN1.sql)

                ' Delete all differences
                ASCMAIN1.Progress("Deletions", "")
                sql = " SELECT X.*, ICTSLCT1.CATALOG_SELECTION_DESC"
                sql &= " FROM"
                sql &= " ("
                sql &= " SELECT  EDT832OX.* "
                sql &= " from EDT832OX "
                sql &= " where (ITEM_UPC, ITEM_CODE, ITEM_COLOR_CODE, ITEM_SIZE_CODE, ITEM_SELECTION_CODE)"
                sql &= " IN"
                sql &= " ("
                sql &= " SELECT ITEM_UPC, ITEM_CODE, ITEM_COLOR_CODE, ITEM_SIZE_CODE, ITEM_SELECTION_CODE from EDT832OX "
                sql &= " MINUS"
                sql &= " SELECT ICTSTYC1.UPC_CODE ITEM_UPC, ICTSTYC1.STYLE_CODE ITEM_CODE, NVL(ICTSTYC1.NRF_COLOR_CODE, '000') ITEM_COLOR_CODE"
                sql &= " , NVL(ICTSTYC1.NRF_SIZE_CODE, '00000') ITEM_SIZE_CODE"
                sql &= " , NVL(ICTSTYC1.CATALOG_SELECTION_CODE, '001') ITEM_SELECTION_CODE"
                sql &= " from ICTSTYC1, ICTCOLRN, ICTSIZEN"
                sql &= "  where NVL(ICTSTYC1.NRF_COLOR_CODE, '000') = ICTCOLRN.NRF_COLOR_CODE (+)"
                sql &= "  and NVL(ICTSTYC1.NRF_SIZE_CODE, '00000') = ICTSIZEN.NRF_SIZE_CODE (+)"
                sql &= " )"
                sql &= " ) X, ICTSLCT1"
                sql &= " WHERE NVL(X.ITEM_SELECTION_CODE, '001') = ICTSLCT1.CATALOG_SELECTION_CODE"
                sql &= " AND X.COMPANY_CODE = '" & companyCode & "'"
                sql &= " AND X.EDI_OUR_ID = '" & EDI_OUR_ID & "'"
                sql &= " AND X.EDI_TP_ID = '" & EDI_TP_ID & "'"


                tblData = ASCDATA1.GetDataTable(sql)
                For Each col As DataColumn In tblData.Columns
                    col.ReadOnly = False
                Next

                ' Set data in deleted rows to data in EDT832OX
                For Each row As DataRow In tblData.Select()
                    ASCMAIN1.sql = "ITEM_CODE = '" & row.Item("ITEM_CODE") & "' and EDI_OUR_ID = '" & row.Item("EDI_OUR_ID") & "' AND EDI_TP_ID = '" & row.Item("EDI_TP_ID") & "'"
                    If dst.Tables("EDT832OX").Select(ASCMAIN1.sql).Length > 0 Then
                        Dim rowEDT832OX As DataRow = dst.Tables("EDT832OX").Select(ASCMAIN1.sql)(0)
                        For Each col As DataColumn In rowEDT832OX.Table.Columns
                            If tblData.Columns.Contains(col.ColumnName) AndAlso col.ColumnName <> "ITEM_CODE" Then
                                row.Item(col.ColumnName) = rowEDT832OX.Item(col.ColumnName)
                            End If
                        Next
                    End If

                Next
                Create832Entries(tblData, itemDeleted, rowEDTTRPM1)

                ' Additions - Item Codes not in the Calalogue
                ASCMAIN1.Progress("Additions", "")
                sql = catalogueItemSql.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " and (NVL(ICTSTYC1.STYLE_COLOR_STATUS, 'A') = 'A' AND  NVL(ICTSTYL1.STYLE_STATUS, 'A') = 'A' ) and ICTSTYC1.STYLE_CODE in (" & vbCr
                sql &= " Select ITEM_CODE from (" & vbCr
                sql &= sqlICTITEM1.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " Minus" & vbCr
                sql &= sqlEDT832O2.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " ))"
                tblData = ASCDATA1.GetDataTable(sql)
                Create832Entries(tblData, itemAdded, rowEDTTRPM1)

                ASCMAIN1.Progress("Changes", "")
                sql = catalogueItemSql.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " and (NVL(ICTSTYC1.STYLE_COLOR_STATUS, 'A') = 'A' AND  NVL(ICTSTYL1.STYLE_STATUS, 'A') = 'A' ) and ICTSTYC1.STYLE_CODE in (" & vbCr
                sql &= " Select ITEM_CODE from (" & vbCr
                sql &= sqlICTITEM1.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " INTERSECT " & vbCr
                sql &= sqlEDT832O2.Replace("'COMPANY_CODE'", "'" & companyCode & "'").Replace("'EDI_OUR_ID' ", "'" & EDI_OUR_ID & "'").Replace("'EDI_TP_ID'", "'" & EDI_TP_ID & "'")
                sql &= " ))"
                tblData = ASCDATA1.GetDataTable(sql)

                Dim nothingChanged As Boolean = True
                Dim Fields As String = "ITEM_DESC,ITEM_RETAIL_PRICE,ITEM_SO_QTY_MIN,ITEM_SO_QTY_MULT,EDI_DISCONTINUE_DATE,ITEM_GTIN"
                For Each rowChanges As DataRow In tblData.Select("")
                    Dim ITEM_CODE As String = rowChanges.Item("ITEM_CODE") & String.Empty
                    For Each rowEDT832OX As DataRow In dst.Tables("EDT832OX").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                        nothingChanged = True
                        For Each field As String In Fields.Split(",")
                            field = field.Trim
                            If field.Length = 0 Then Continue For
                            If rowEDT832OX.Item(field) & String.Empty <> rowChanges.Item(field) & String.Empty Then
                                nothingChanged = False
                                Exit For
                            End If
                        Next
                    Next
                    ' If nothing changed then delete the record
                    If nothingChanged Then rowChanges.Delete()
                Next
                tblData.AcceptChanges()
                Create832Entries(tblData, itemChanged, rowEDTTRPM1)
            Next

            EnforceConstraints(True)
        Catch ex As Exception
            ClearRecord()
            MessageBox.Show(ex.Message)
        Finally

        End Try

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

            If dst.Tables("EDT832O1").Rows.Count = 0 Then
                .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
            Else
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            End If

            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
        End With

        If ScreenMode Then
            grdEDT832O1.Visible = True
        Else
            ClearRecord()
            grdEDT832O1.Visible = False
        End If
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
                If dst.Tables("EDTTRPM1").Rows.Count = 0 Then
                    EMsg &= "There are no Customers setup to receive the 832 Catalog."
                End If

            Case "Update"
                If dst.Tables("EDT832O1").Rows.Count = 0 Then
                    EMsg = "No records to update."
                End If

            Case "Cancel"

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

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    ''' <summary>
    ''' Updates data based on the current state of the screen and the type of processing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateRecord()

        Try
            MyBase.BeginTrans()

            Dim rowEDT832OX As DataRow = Nothing

            For Each rowEDTSYSIH As DataRow In dst.Tables("EDTSYSIH").Select("", "EDI_OUTBOUND_DOC_NO")
                Dim COMPANY_CODE As String = rowEDTSYSIH.Item("COMPANY_CODE")
                Dim EDI_OUTBOUND_DOC_NO As String = rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO")
                Dim EDI_OUR_ID As String = rowEDTSYSIH.Item("EDI_OUR_ID")
                Dim EDI_TP_ID As String = rowEDTSYSIH.Item("EDI_TP_ID")

                For Each rowEDT832O1 As DataRow In dst.Tables("EDT832O1").Select("COMPANY_CODE = '" & COMPANY_CODE & "' AND EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'")

                    If rowEDT832O1.Item("EDI_PURPOSE_CODE") = itemDeleted Then
                        Continue For
                    End If

                    For Each rowEDT832O2 As DataRow In dst.Tables("EDT832O2").Select("COMPANY_CODE = '" & COMPANY_CODE & "' AND EDI_OUTBOUND_DOC_NO = '" & EDI_OUTBOUND_DOC_NO & "'")

                        Dim ITEM_CODE As String = rowEDT832O2.Item("ITEM_CODE")
                        rowEDT832OX = dst.Tables("EDT832OX").Rows.Find(New Object() {COMPANY_CODE, EDI_OUR_ID, EDI_TP_ID, ITEM_CODE})

                        If rowEDT832OX Is Nothing Then
                            rowEDT832OX = dst.Tables("EDT832OX").NewRow
                            rowEDT832OX.Item("COMPANY_CODE") = COMPANY_CODE
                            rowEDT832OX.Item("EDI_OUR_ID") = EDI_OUR_ID
                            rowEDT832OX.Item("EDI_TP_ID") = EDI_TP_ID
                            rowEDT832OX.Item("ITEM_CODE") = ITEM_CODE
                            dst.Tables("EDT832OX").Rows.Add(rowEDT832OX)
                        End If

                        rowEDT832OX.Item("ITEM_DESC") = rowEDT832O2.Item("ITEM_DESC")
                        rowEDT832OX.Item("ITEM_SELECTION_CODE") = rowEDT832O2.Item("ITEM_SELECTION_CODE")
                        rowEDT832OX.Item("ITEM_RETAIL_PRICE") = rowEDT832O2.Item("ITEM_RETAIL_PRICE")
                        rowEDT832OX.Item("ITEM_COLOR_CODE") = rowEDT832O2.Item("ITEM_COLOR_CODE")
                        rowEDT832OX.Item("ITEM_COLOR_DESC") = rowEDT832O2.Item("ITEM_COLOR_DESC")
                        rowEDT832OX.Item("ITEM_SIZE_CODE") = rowEDT832O2.Item("ITEM_SIZE_CODE")
                        rowEDT832OX.Item("ITEM_SIZE_DESC") = rowEDT832O2.Item("ITEM_SIZE_DESC")
                        rowEDT832OX.Item("ITEM_SO_QTY_MIN") = rowEDT832O2.Item("ITEM_SO_QTY_MIN")
                        rowEDT832OX.Item("ITEM_SO_QTY_MULT") = rowEDT832O2.Item("ITEM_SO_QTY_MULT")
                        rowEDT832OX.Item("ITEM_UPC") = rowEDT832O2.Item("ITEM_UPC")
                        rowEDT832OX.Item("ITEM_EAN") = rowEDT832O2.Item("ITEM_EAN")
                        rowEDT832OX.Item("ITEM_GTIN") = rowEDT832O2.Item("ITEM_GTIN")
                        rowEDT832OX.Item("EDI_DISCONTINUE_DATE") = rowEDT832O2.Item("EDI_DISCONTINUE_DATE")
                        rowEDT832OX.Item("HAZARD_CODE") = rowEDT832O2.Item("HAZARD_CODE")
                    Next
                Next
            Next

            MyBase.Update_Record_TDA("EDTSYSIH")
            MyBase.Update_Record_TDA("EDT832O1")
            MyBase.Update_Record_TDA("EDT832O2")
            MyBase.Update_Record_TDA("EDT832OX")

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT STYLE_CODE FROM ICTSTYL1 WHERE NVL(STYLE_STATUS, 'A') = 'I')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT STYLE_CODE FROM ICTSTYC1 WHERE NVL(STYLE_COLOR_STATUS, 'A') = 'I')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT STYLE_CODE FROM ICTSTYL1 WHERE NVL(STYLE_HIDE_FROM_CAT, '0') = '1')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Delete from EDT832OX where ITEM_CODE IN (SELECT STYLE_CODE FROM ICTSTYC1 WHERE NVL(HIDE_FROM_CAT, '0') = '1')"
            ASCDATA1.ExecuteSQL()

            MyBase.CommitTrans("Update Successful")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)

        End Try

    End Sub

#End Region

#Region "Private Subs / Functions"

    ''' <summary>
    ''' Create 832O1 and 832O2 records for the specifiec Purpose Type
    ''' </summary>
    ''' <param name="tbl832Data"></param>
    ''' <param name="EDI_TRANS_PURP_CODE"></param>
    ''' <remarks></remarks>
    Private Sub Create832Entries(ByRef tbl832Data As DataTable, ByVal EDI_TRANS_PURP_CODE As String, ByRef rowEDTTRPM1 As DataRow)

        If tbl832Data Is Nothing OrElse tbl832Data.Rows.Count = 0 Then
            Exit Sub
        End If

        For Each rowSelectionCode As DataRow In ASCDATA1.SelectDistinct(tbl832Data, New String() {"ITEM_SELECTION_CODE"}).Rows
            Dim ITEM_SELECTION_CODE As String = rowSelectionCode.Item("ITEM_SELECTION_CODE") & String.Empty
            Dim rowHeader As DataRow = tbl832Data.Select("ITEM_SELECTION_CODE = '" & ITEM_SELECTION_CODE & "'")(0)

            ' Moved from up above
            Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

            Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
            rowEDTSYSIH.Item("COMPANY_CODE") = companyCode
            rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDTSYSIH.Item("EDI_APPLICATION_ID") = ediApplicationId
            If rowEDTTRPM1.Item("EDI_STATUS") = "T" Then
                rowEDTSYSIH.Item("EDI_PROCESS_IND") = "T"
            Else
                rowEDTSYSIH.Item("EDI_PROCESS_IND") = "1"
            End If
            rowEDTSYSIH.Item("EDI_OUR_ID") = rowHeader.Item("EDI_OUR_ID")
            rowEDTSYSIH.Item("EDI_TP_ID") = rowHeader.Item("EDI_TP_ID")
            rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
            rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

            Dim rowEDT832O1 As DataRow = dst.Tables("EDT832O1").NewRow
            rowEDT832O1.Item("COMPANY_CODE") = companyCode
            rowEDT832O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDT832O1.Item("EDI_CATALOG_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO")
            rowEDT832O1.Item("EDI_CATALOG_VERSION") = ITEM_SELECTION_CODE

            Dim rowICTSLCT1 As DataRow = LookUp("ICTSLCT1", ITEM_SELECTION_CODE)
            If rowICTSLCT1 IsNot Nothing Then
                rowEDT832O1.Item("EDI_CATALOG_DESC") = rowICTSLCT1.Item("CATALOG_SELECTION_DESC")
            Else
                rowEDT832O1.Item("EDI_CATALOG_DESC") = rowHeader.Item("CATALOG_SELECTION_DESC")
            End If

            rowEDT832O1.Item("EDI_PURPOSE_CODE") = EDI_TRANS_PURP_CODE
            rowEDT832O1.Item("EDI_CATALOG_DATE") = DateTime.Now.ToString("MM/dd/yyyy")
            rowEDT832O1.Item("EDI_NAME") = ROWs("ASTPARM1").Item("AS_PARM_INST_NAME")
            rowEDT832O1.Item("INIT_DATE") = DateTime.Now
            rowEDT832O1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowEDT832O1.Item("LAST_DATE") = DateTime.Now
            rowEDT832O1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            dst.Tables("EDT832O1").Rows.Add(rowEDT832O1)

            Dim EDI_DOC_LNO As Int16 = 0
            Dim rowEDT832O2 As DataRow = Nothing
            For Each rowDetail As DataRow In tbl832Data.Select("ITEM_SELECTION_CODE = '" & ITEM_SELECTION_CODE & "'")
                ASCMAIN1.Progress("-", rowDetail.Item("ITEM_CODE") & "")

                rowEDT832O2 = dst.Tables("EDT832O2").NewRow
                rowEDT832O2.Item("COMPANY_CODE") = companyCode
                rowEDT832O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                EDI_DOC_LNO += 1
                rowEDT832O2.Item("EDI_DOC_LNO") = EDI_DOC_LNO
                rowEDT832O2.Item("ITEM_CODE") = rowDetail.Item("ITEM_CODE")
                rowEDT832O2.Item("ITEM_DESC") = rowDetail.Item("ITEM_DESC")
                rowEDT832O2.Item("ITEM_SELECTION_CODE") = rowDetail.Item("ITEM_SELECTION_CODE")
                rowEDT832O2.Item("ITEM_RETAIL_PRICE") = rowDetail.Item("ITEM_RETAIL_PRICE")
                rowEDT832O2.Item("ITEM_COLOR_CODE") = rowDetail.Item("ITEM_COLOR_CODE")
                rowEDT832O2.Item("ITEM_COLOR_DESC") = rowDetail.Item("ITEM_COLOR_DESC")
                rowEDT832O2.Item("ITEM_SIZE_CODE") = rowDetail.Item("ITEM_SIZE_CODE")
                rowEDT832O2.Item("ITEM_SIZE_DESC") = rowDetail.Item("ITEM_SIZE_DESC")
                rowEDT832O2.Item("ITEM_SO_QTY_MIN") = rowDetail.Item("ITEM_SO_QTY_MIN")
                rowEDT832O2.Item("ITEM_SO_QTY_MULT") = rowDetail.Item("ITEM_SO_QTY_MULT")
                rowEDT832O2.Item("ITEM_UPC") = rowDetail.Item("ITEM_UPC")
                rowEDT832O2.Item("ITEM_EAN") = rowDetail.Item("ITEM_EAN")
                rowEDT832O2.Item("ITEM_GTIN") = rowDetail.Item("ITEM_GTIN")
                rowEDT832O2.Item("EDI_DISCONTINUE_DATE") = rowDetail.Item("EDI_DISCONTINUE_DATE")
                rowEDT832O2.Item("HAZARD_CODE") = rowDetail.Item("HAZARD_CODE")
                dst.Tables("EDT832O2").Rows.Add(rowEDT832O2)
            Next
        Next

    End Sub

#End Region

#Region "Form Controls"

#End Region

End Class