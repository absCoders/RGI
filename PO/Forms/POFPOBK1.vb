Public Class POFPOBK1
    Dim POTPOBK1 As String 'TABLE_NAME
    Dim sqlPOTPOBK1 As String

    Dim ETD_to_ETA As Integer


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'just a note so I can redeploy
        Get_PARM("POTPARM1")

        With dst
            ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR2.STYLE_CODE, POTORDR2.PO_ORDER_LNO,  POTORDR2.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & ", POTORDR1.VEND_NAME, POTORDR2.PO_QTY_ORD, POTORDR2.PO_DATE_SHIP_BY,  POTORDR2.PO_DATE_ETA, POTORDR2.PO_BOOK_BY_DATE, 'YES' AS BOOKED, POTORDR2.PO_ON_BOARD_DATE" & vbCrLf _
                & ", 'YES' AS RECEIVED, POTORDR2.LAST_DATE_SHIP_BY, POTORDR2.PO_ON_BOARD_DATE as PO_ETD_DATE, 'YES' AS SHIPPED, 'N' AS UPDATE_ROW, 30 AS ETA_DAYS, POTSHIP1.PO_DATE_SHIPPED as ACT_SHIP_DATE" & vbCr _
                & " from POTORDR1, POTORDR2, ICTSTYL1, POTSHIP3, POTSHIP1" & vbCrLf _
                & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                & " and POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE " & vbCrLf _
                & " and POTORDR2.PO_ORDER_NO = POTSHIP3.PO_ORDER_NO (+) " & vbCrLf _
                & " and POTORDR2.PO_ORDER_LNO = POTSHIP3.PO_ORDER_LNO (+) " & vbCrLf _
                & " and POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+) " & vbCrLf

            sqlPOTPOBK1 = ASCMAIN1.sql

            POTPOBK1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from " & POTPOBK1
            Create_TDA(.Tables.Add("POTPOBK1"), POTPOBK1, "**", 0, True)

        End With


        grdPOTPOBK1.DataSource = dst.Tables("POTPOBK1")

        Create_Summary(grdPOTPOBK1, "PO_ORDER_NO", "Count")

        With grdPOTPOBK1.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_NO").Header.Fixed = True
            .Columns("PO_ORDER_LNO").Header.Fixed = True
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            With .Columns("PO_ORDER_NO")
                .Header.Fixed = True
                .Header.Caption = "PO No"
                .Width = 80
            End With
            With .Columns("STYLE_CODE")
                .Header.Fixed = True
                .Header.Caption = "Style"
                .Width = 140
            End With

            With .Columns("PO_ORDER_LNO")
                .Header.Fixed = True
                .Header.Caption = "PO Lno"
                .Width = 40
            End With

            With .Columns("COLOR_CODE")
                .Header.Fixed = True
                .Header.Caption = "Color"
                .Width = 60
            End With

            With .Columns("STYLE_DESC")
                .Header.Fixed = True
                .Header.Caption = "Desc"
                .Width = 200
            End With

            With .Columns("VEND_NAME")
                .Header.Fixed = True
                .Header.Caption = "Vendor"
                .Width = 160
            End With

            With .Columns("PO_QTY_ORD")
                .Header.Fixed = True
                .Header.Caption = "Qty Ordered"
                .Width = 70
            End With

            With .Columns("PO_DATE_SHIP_BY")
                .Header.Fixed = True
                .Header.Caption = "Ship By"
                .Width = 100
            End With

            With .Columns("PO_DATE_ETA")
                .Header.Fixed = True
                .Header.Caption = "Whse Due"
                .Width = 100
            End With

            With .Columns("PO_ETD_DATE")
                .Header.Fixed = True
                .Header.Caption = "ETD Date"
                .Width = 100
            End With

            With .Columns("PO_BOOK_BY_DATE")
                .Header.Fixed = True
                .Header.Caption = "Book By"
                .Width = 100
            End With

            With .Columns("BOOKED")
                .Header.Fixed = True
                .Header.Caption = "Booked"
                .Width = 60
                .Hidden = True

            End With

            With .Columns("PO_ETD_DATE")
                .Header.Fixed = True
                .Header.Caption = "ETD Date"
                .Width = 100
            End With

            With .Columns("PO_ON_BOARD_DATE")
                .Header.Fixed = True
                .Header.Caption = "On Board"
                .Width = 100
            End With

            With .Columns("RECEIVED")
                .Header.Fixed = True
                .Header.Caption = "Received"
                .Width = 60
                .Hidden = True
            End With

            With .Columns("LAST_DATE_SHIP_BY")
                .Header.Fixed = True
                .Header.Caption = "Last Ship By"
                .Width = 100
            End With

            With .Columns("SHIPPED")
                .Header.Fixed = True
                .Header.Caption = "Shipped"
                .Width = 60
                .Hidden = False
            End With

            With .Columns("UPDATE_ROW")
                .Header.Fixed = True
                .Header.Caption = "Update"
                .Width = 60
                .Hidden = False
            End With

            With .Columns("ETA_DAYS")
                .Header.Fixed = True
                .Header.Caption = "ETA Days"
                .Width = 60
                .Hidden = True
            End With

  


            With .Columns("ACT_SHIP_DATE")
                .Header.Fixed = True
                .Header.Caption = "Act Ship Date"
                .Width = 140
                .Hidden = True
            End With

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                '  If gcol.Key = "NEW_PO_COST" Then
                'gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ' End If
            Next
        End With

        'With grdPOTPOBK1.DisplayLayout.Bands(0)
        '    With .Columns("PO_ORDER_NO")
        '        .Header.Fixed = True
        '        .Header.Caption = "Shipment"
        '        .Width = 100
        '    End With
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        If gcol.Key <> "PO_SHIPMENT_NO" Then
        '            gcol.Width = 90
        '        End If
        '    Next
        'End With



        spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load from Spreadsheet"
                Import_from_Excel()
                If dst.Tables("POTPOBK1").Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Loaded"
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

            Case "Load from Spreadsheet"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
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
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTPOBK1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "E" Then
        Else

            If EntryMode = "H" Then
                ASCMAIN1.sql = "TRUNCATE TABLE " & POTPOBK1
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "INSERT INTO " & POTPOBK1 & " SELECT *  FROM POTPOBK1 "
                ASCDATA1.ExecuteSQL()

                Fill_Records("POTPOBK1")

            Else

                ASCMAIN1.sql = "TRUNCATE TABLE " & POTPOBK1
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "INSERT INTO " & POTPOBK1 & " SELECT X.*,'','','','','' FROM (" & sqlPOTPOBK1 & ") X "
                ASCDATA1.ExecuteSQL()

                'DANAC= INSERT X.*,'','','','','' FROM (   X 


                Fill_Records("POTPOBK1")
            End If
        End If

        Sort_grdColumns(grdPOTPOBK1, "PO_ORDER_NO,PO_ORDER_LNO")

        'For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
        '    rowTATCOLS1.Item("SEL") = "1"
        'Next


        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        ASCMAIN1.sql = "TRUNCATE TABLE " & POTPOBK1
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("POTPOBK1")

        BeginTrans()



        ' NEED TO ADD IN COMMA IF MORE THAN ONE COLUMN
        ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is  Select " & POTPOBK1 & ".*, POTSHIP1.PO_DATE_SHIPPED from  " & POTPOBK1 & ", POTSHIP3, POTSHIP1 where " & POTPOBK1 & ".UPDATE_ROW = 'Y'" _
                & " AND POTSHIP3.PO_ORDER_NO (+) = " & POTPOBK1 & ".PO_ORDER_NO " _
                & " AND POTSHIP3.PO_ORDER_LNO (+) = " & POTPOBK1 & ".PO_ORDER_LNO " _
                & " AND POTSHIP3.PO_SHIPMENT_NO = POTSHIP1.PO_SHIPMENT_NO (+); " _
                & " Begin For R1 in C1 Loop" _
                & " Update POTORDR2 Set  PO_BOOK_BY_DATE = R1.PO_ON_BOARD_DATE, PO_ON_BOARD_DATE = R1.PO_ETD_DATE" _
                & "  where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO  ;" _
                & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()
        'CommitTrans("")

        ' BeginTrans()
        ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is  Select * from  " & POTPOBK1 & " where UPDATE_ROW = 'Y' and SHIPPED = 'NO' and NVL(LAST_DATE_SHIP_BY,'') IS NOT NULL ;" _
                & " Begin For R1 in C1 Loop" _
                & " Update POTORDR2 Set PO_DATE_ETA =  R1.LAST_DATE_SHIP_BY + R1.ETA_DAYS, VEND_CARGO_READY_DATE = R1.PO_DATE_SHIP_BY " _
                & "  where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO  ;" _
                & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ' BeginTrans()
        ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is  Select * from  " & POTPOBK1 & " where NVL(PO_DATE_SHIP_BY,'') IS NOT NULL ;" _
                & " Begin For R1 in C1 Loop" _
                & " Update POTORDR2 Set VEND_CARGO_READY_DATE = R1.PO_DATE_SHIP_BY " _
                & "  where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO  ;" _
                & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()



        'CommitTrans("")

        'ASCMAIN1.sql = "" _
        '& "Begin Declare Cursor C1 is  Select * from  " & POTPOBK1 & " where UPDATE_ROW = 'Y' ;" _
        '& " Begin For R1 in C1 Loop" _
        '& " Update POTORDR2 Set  PO_BOOK_BY_DATE = CASE WHEN R1.RECEIVED = 'YES' THEN R1.PO_BOOK_BY_DATE END," _
        '& "  PO_ON_BOARD_DATE = CASE WHEN R1.SHIPPED = 'NO' THEN R1.PO_ETD_DATE END," _
        '& "  PO_DATE_ETA = CASE WHEN R1.SHIPPED = 'NO' THEN R1.PO_ETD_DATE + R1.ETA_DAYS END " _
        '& "  where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO  ;" _
        '& " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()
        'CommitTrans("")
        'BeginTrans()


        ' 'ASCMAIN1.sql = "" _
        '        & "Begin Declare Cursor C1 is " _
        '        & " Select * from  " & POTPOBK1 & " where SHIPPED = 'YES' ;" & vbCrLf _
        '        & " Begin For R1 in C1 Loop" & vbCrLf _
        '        & " Update POTORDR2 Set " & vbCrLf _
        '        & " PO_ON_BOARD_DATE  = R1.PO_ON_BOARD_DATE " & vbCrLf _
        '        & ", PO_DATE_ETA = R1.PO_ON_BOARD_DATE + 30 " & vbCrLf _
        '        & " where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO  ;" & vbCrLf _
        '        & " End Loop; End; End;"
        'ASCDATA1.ExecuteSQL()

        'EntryMode = ""

        ASCMAIN1.sql = "" _
            & " UPDATE POTPPRM1 SET BOOK_RPT_UPDATE_OPER = '" & ASCMAIN1.USER_ID & "',  BOOK_RPT_UPDATE_DATE = SYSDATE WHERE POTPPRM1_CODE = 'Z' "
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub




#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTPOBK1, "SSBS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Update Column")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name
            Case "grdPOTPOBK1"
                tlb_sbt = DirectCast(tlb_pop.Tools("Update Column"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = False
                If grd.ActiveCell IsNot Nothing Then
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Dim row As DataRow = dst.Tables("TATCOLS1").Rows.Find(COLUMN_NAME)
                    If row IsNot Nothing Then
                        tlb_sbt.SharedProps.Visible = True
                        tlb_sbt.SharedProps.Caption = "Update " & row.Item("COLUMN_CAPTION")
                        tlb_sbt.Tag = ""
                        tlb_sbt.Checked = (row.Item("SEL") = "1")
                        tlb_sbt.Tag = COLUMN_NAME
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdPOTPOBK1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        
        End Select
    End Sub

#End Region

 

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "VEND_CODE"

        End Select

    End Sub


    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

    Sub Import_from_Excel()
        Dim openFileDialog1 As New OpenFileDialog
        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
        openFileDialog1.Filter = "xls files (*.xls)|*.xls"
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            Dim FILENAME As String = openFileDialog1.FileName
            Try
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                "data source=" & FILENAME & ";" & _
                "Extended Properties=Excel 8.0;"
                Dim objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                objConnection.Open()
                Dim dbSchema As DataTable = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                If dbSchema.Rows.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Sub
                End If
                Dim strSQL As String = "SELECT * FROM [" & dbSchema.Rows(0).Item("TABLE_NAME") & "]"
                Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                Dim dt As New DataTable
                objAdapter.Fill(dt)
                objConnection.Close()

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Data from XLS")

                Dim COLs As Int32 = dt.Columns.Count
                Dim PRDmax As Int32 = COLs - 3

                If COLs < 2 Then
                    MsgBox("There appear to be no Records to Import", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else

                End If

                dst.Tables("POTPOBK1").Rows.Clear()

                For Each row As DataRow In dt.Rows
                    Dim PO_ORDER_NO As String = row.Item(0) & ""
                    Dim PO_ORDER_LNO As String = row.Item(2) & ""
                    Dim STYLE_CODE As String = row.Item(1) & ""

                    ETD_to_ETA = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETD_TO_ETA") & "")
                    If PO_ORDER_NO = "155391" Then
                        Dim DANA1 As String = row.Item(0) & ""
                    End If



                    'rowPOTPOBK1.Item("STYLE_CODE") = STYLE_CODE
                    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                    Dim rowPOTORDR2 As DataRow = LookUp("POTORDR2", New String() {PO_ORDER_NO, PO_ORDER_LNO})
 

                    If rowPOTORDR1 Is Nothing Then
                    Else
                        Dim rowPOTPOBK1 As DataRow = dst.Tables("POTPOBK1").NewRow
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        Dim WHSE_CODE As String = rowPOTORDR1.Item("WHSE_CODE") & ""
                        Dim PORT_CODE_ORIG As String = rowPOTORDR1.Item("WHSE_CODE") & ""
                        Dim PO_SHIP_VIA As String = rowPOTORDR1.Item("PO_SHIP_VIA") & ""

                        Dim rowPOTSHIP3 As DataRow = LookUp("POTSHIP3", New String() {PO_ORDER_NO, PO_ORDER_LNO})



                        Dim PO_SHIPMENT_NO As String
                        Dim ACT_SHIP_DATE As String

                        If rowPOTSHIP3 IsNot Nothing Then
                            PO_SHIPMENT_NO = rowPOTSHIP3.Item("PO_SHIPMENT_NO") & ""
                            Dim rowPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                            If rowPOTSHIP1 IsNot Nothing Then
                                ACT_SHIP_DATE = rowPOTSHIP1.Item("PO_DATE_SHIPPED") & ""
                                rowPOTPOBK1.Item("ACT_SHIP_DATE") = rowPOTSHIP1.Item("PO_DATE_SHIPPED") & ""
                            End If
                        End If

                        If PORT_CODE_ORIG & "" & WHSE_CODE & "" <> "" Then
                            Dim rowICTPORT2 As DataRow = LookUp("ICTPORT2", New String() {PORT_CODE_ORIG, WHSE_CODE})
                            If rowICTPORT2 Is Nothing Then
                            Else
                                ETD_to_ETA = Val(rowICTPORT2.Item("ETD_TO_ETA") & "")
                            End If
                        End If

                        If PO_SHIP_VIA & "" <> "" Then
                            Dim rowPOTSVIA1 As DataRow = LookUp("POTSVIA1", PORT_CODE_ORIG)
                            If rowPOTSVIA1 Is Nothing Then
                            Else
                                ETD_to_ETA = Val(rowPOTSVIA1.Item("PO_SHIP_VIA_ETD_TO_ETA") & "")
                            End If
                        End If
                        'Dim rowICTPORT2 As DataRow = LookUp("ICTPORT2", New String() {PORT_CODE_ORIG, WHSE_CODE})

                        ' Dim rowPOTSVIA1 As DataRow = LookUp("POTSVIA1", PORT_CODE_ORIG)
                        If rowICTSTYL1 Is Nothing Then
                            ' LOG ERROR
                        Else
                            Try
                                Dim DANA_COLS As Integer = dst.Tables("POTPOBK1").Columns.Count
                                For I As Integer = 0 To dst.Tables("POTPOBK1").Columns.Count - 5 ' IS THIS CORRECT? DRC/ABS
                                    Dim DANA As String = row.Item(I) & ""
                                    If DANA = "MT18830" Then
                                        DANA = DANA
                                    End If
                                    'Dim II As Integer = 0
                                    'If I > 11 Then
                                    '    II = I + 1
                                    'Else
                                    '    II = I
                                    'End If
                                    If I = 7 Or I = 8 Or I = 9 Or I = 11 Or I = 13 Then
                                        If row.Item(I) & "" = "" Then
                                        Else
                                            If I = 13 Then
                                                If row.Item(I + 1) & "" = "ZZZ" Then
                                                    rowPOTPOBK1.Item(I) = row.Item(I) & ""
                                                Else
                                                    rowPOTPOBK1.Item(I + 1) = row.Item(I) & ""
                                                End If
                                            Else
                                                rowPOTPOBK1.Item(I) = row.Item(I) & ""
                                            End If
                                        End If
                                    Else
                                        If I > 13 Then
                                            rowPOTPOBK1.Item(I + 1) = Trim(row.Item(I) & "")
                                        Else
                                            rowPOTPOBK1.Item(I) = Trim(row.Item(I) & "")
                                        End If

                                    End If

                                    If I = 0 Then
                                        rowPOTPOBK1.Item("ETA_DAYS") = ETD_to_ETA
                                        rowPOTPOBK1.Item("UPDATE_ROW") = "N"
                                    End If

                                    If I = 13 And row.Item(13) & "" = "YES" Then
                                        rowPOTPOBK1.Item("UPDATE_ROW") = "Y"
                                    Else
                                        rowPOTPOBK1.Item("UPDATE_ROW") = "Y"
                                    End If

                                    If I = 14 And row.Item(14) & "" = "NO" Then
                                        rowPOTPOBK1.Item("UPDATE_ROW") = "Y"
                                    Else
                                        rowPOTPOBK1.Item("UPDATE_ROW") = "Y"
                                    End If

                                Next
                                Dim danastyle As String = row.Item(0)
                                Dim dana2 As String = "what!"
                                dst.Tables("POTPOBK1").Rows.Add(rowPOTPOBK1)
                                row.Delete()

                            Catch ex As Exception
                                Stop
                            End Try
                        End If
                    End If
                Next

                If dt.Rows.Count <> 0 Then
                    Dim frmASFMSGBF As New ASFMSGBF

                    frmASFMSGBF.Show_grd(dt, Me, "Records which Failed to Load")

                End If

            Catch ex As Exception

            End Try

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

#End Region
 
  
End Class