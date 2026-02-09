Public Class ARFPYMT1

    Dim rowGLTBANK1 As DataRow
    Dim rowGLTPARM2 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String

    Dim blnFX_Support As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*", 1)

            ASCMAIN1.sql = "SELECT ARTPYMT1.*, X.PYMT_AMT, X.PYMT_AMT_CURR" _
            & " from ARTPYMT1, " _
            & " (Select ARTPYMT2.PYMT_BATCH_NO" _
            & ", Sum (ARTPYMT2.CUST_PYMT_AMT) PYMT_AMT" _
            & ", Sum (ARTPYMT2.CUST_PYMT_AMT_CURR) PYMT_AMT_CURR" _
            & " from ARTPYMT1,ARTPYMT2" _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & "   and ARTPYMT1.STATUS LIKE :PARM1 and ARTPYMT1.INIT_OPER LIKE :PARM2 and OPS_YYYYPP = :PARM3" _
            & " group by ARTPYMT2.PYMT_BATCH_NO) X" _
            & " where ARTPYMT1.PYMT_BATCH_NO = X.PYMT_BATCH_NO"
            Create_TDA(.Tables.Add, "ARTPYMTX", "**", 0, False, "VVV", 1)

        End With

        rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.CYP)

        grdARTPYMTX.DataSource = dst.Tables("ARTPYMTX")
        grdARTPYMT2.DataSource = dst.Tables("ARTPYMT2")

        Create_Summary(grdARTPYMTX, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMTX, "PYMT_AMT")

        Create_Summary(grdARTPYMT2, "PYMT_BATCH_LNO", "Count")
        Create_Summary(grdARTPYMT2, New String() {"CUST_PYMT_AMT", "CUST_PYMT_AMT_CURR"})

        blnFX_Support = ASCDATA1.GetDataTable("*", "TATCURR1").Rows.Count > 1

        If Not blnFX_Support Then
            grpCURR_CODE.Visible = False
            With grdARTPYMTX.DisplayLayout.Bands(0)
                .Columns("CURR_CODE").Hidden = True
                .Columns("CURR_EXCH_RATE").Hidden = True
                .Columns("PYMT_AMT_CURR").Hidden = True
            End With
        Else
            grpCURR_CODE.Visible = True
        End If

        Set_Read_Only(grpCURR_CODE, True)

        With grdARTPYMT2.DisplayLayout.Bands(0)
            .Columns("PYMT_BATCH_LNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_NAME").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("NON_AR").CellAppearance.BackColor = Drawing.Color.Beige

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "CURR_CODE" Or gcol.Key = "CURR_EXCH_RATE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With
        grdARTPYMT2.DisplayLayout.Bands(0).Columns("NON_AR").Tag = "N"
        grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_PYMT_AMT").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & " Amount"

        With grdARTPYMTX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "CURR_CODE" Or gcol.Key = "CURR_EXCH_RATE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        ASCMAIN1.Add_Value_List(grdARTPYMTX, "STATUS", Nothing, New String() {":", "0:Entered", "1:Unapplied", "2:Applied"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                If EMsg = "" Then
                    ' Validate Bank - shouldn't this happen also on Edit?
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If rowGLTBANK1.Item("CURR_CODE") & "" <> Absx1.txtFor("CURR_CODE").Text Then
                        EMsg &= vbCr & "Currency Code defined for Bank (" & rowGLTBANK1.Item("CURR_CODE") & ") does not match Currency for Batch (" & Absx1.txtFor("CURR_CODE").Text & ")"
                    End If
                    If InStr("AR", rowGLTBANK1.Item("BANK_USE") & "") = 0 Then
                        EMsg &= vbCr & "Bank " & Absx1.txtFor("BANK_CODE").Text & " is not defined for use in A/R"
                    End If
                End If

                If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                    If Val(Absx1.numFor("CURR_EXCH_RATE").Value & "") = 0 Then
                        EMsg &= vbCr & "Rate of Exchange required for Foreign Currency (" & Absx1.txtFor("CURR_CODE").Text & ")"
                    End If
                    If EMsg = "" Then
                        If Val(Absx1.numFor("CURR_EXCH_RATE").Value & "") = 1 Then
                            If MsgBox("Rate of Exchange is set to 1.00" & vbCrLf & "OK to continue?", MsgBoxStyle.YesNo, "Rate of Exchange is set to 1.00") = MsgBoxResult.No Then
                                Exit Sub

                            End If
                        End If
                    End If
                End If

                If Absx1.dteFor("PYMT_BATCH_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "No Receipt Date Specified"
                Else
                    Dim PYMT_BATCH_DATE As String = Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd")
                    Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)

                    If PYMT_BATCH_DATE > Format(dates(dates.Length - 1), "yyyyMMdd") Or PYMT_BATCH_DATE < Format(dates(1), "yyyyMMdd") Then
                        EMsg &= vbCr & "Payment Date must be between " & Format(dates(1), "MM/dd/yyyy") & " and " & Format(dates(dates.Length - 1), "MM/dd/yyyy")
                    End If
                End If

            Case "Edit"
                Validate_Code("PYMT_BATCH_NO")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                If grdARTPYMT2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Payments Entered into Batch"
                Else
                    Dim negative_amounts_present As Boolean = False

                    For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Select("", "", DataViewRowState.CurrentRows)
                        If rowARTPYMT2.Item("CUST_PYMT_REF_DATE") & "" <> "" Then
                            If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                             < Format(rowARTPYMT2.Item("CUST_PYMT_REF_DATE"), "yyyyMMdd") Then
                                EMsg &= vbCr & "Customer Payment Date may not be later than Batch Payment Receipt Date (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                            End If
                            If Format(Now.AddYears(-1), "yyyyMMdd") _
                             > Format(rowARTPYMT2.Item("CUST_PYMT_REF_DATE"), "yyyyMMdd") Then
                                EMsg &= vbCr & "Customer Payment Date may not be more than 1 year ago (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                            End If
                        End If

                        Dim CURR_CODE As String = rowARTPYMT2.Item("CURR_CODE") & ""
                        If CURR_CODE = "" Then
                            CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                            rowARTPYMT2.Item("CURR_CODE") = CURR_CODE
                        Else
                            Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", CURR_CODE)
                            If rowTATCURR1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                            Else
                                If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                                    If CURR_CODE <> Absx1.txtFor("CURR_CODE").Text Then
                                        EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                                    End If
                                End If
                                If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                                    Dim CURR_EXCH_RATE As Decimal = Val(rowARTPYMT2.Item("CURR_EXCH_RATE") & "")
                                    If CURR_EXCH_RATE = 1 Or CURR_EXCH_RATE = 0 Then
                                        EMsg &= vbCr & "Invalid Currency Rate of Exchange for Currency " & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                                    End If

                                    'Dim CURR_EXCH_RATE_now As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                                    Dim CURR_EXCH_RATE_now As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), CURR_CODE, Now.Date)
                                    If CURR_EXCH_RATE_now = 0 OrElse CURR_EXCH_RATE = 0 OrElse System.Math.Abs(100 * (CURR_EXCH_RATE - CURR_EXCH_RATE_now) / CURR_EXCH_RATE) > 10 Then
                                        EMsg &= vbCr & "Invalid Currency Rate of Exchange for Currency " & CURR_CODE & " (> 10% variance with current daily rate of " & Format(CURR_EXCH_RATE_now, "#.0000") & ")"
                                    End If
                                End If
                            End If
                        End If

                        If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                            rowARTPYMT2.Item("CURR_EXCH_RATE") = 1
                        End If

                        CUST_CODE = rowARTPYMT2.Item("CUST_CODE") & ""
                        If CUST_CODE <> "" Then
                            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                            Dim CURR_CODE_customer As String = rowARTCUST1.Item("CURR_CODE")
                            If CURR_CODE_customer = "" Then
                                CURR_CODE_customer = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                            End If
                            If CURR_CODE_customer <> CURR_CODE Then
                                EMsg &= vbCr & "Customer " & CUST_CODE & " default Currency Code is " & CURR_CODE_customer
                                'If MsgBox("Customer " & CUST_CODE & " default Currency Code is " & CURR_CODE_customer _
                                '          & vbCrLf & vbCrLf & "Continue Anyway?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.No Then
                                '    Exit Sub
                                'End If
                            End If
                        End If

                        Dim PYMT_AMT = Val(rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") & "")
                        If System.Math.Round(PYMT_AMT, 2) = 0 Then
                            EMsg &= vbCr & "Zero Amount is Not Permitted (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                        End If

                        If System.Math.Round(PYMT_AMT, 2) < 0 Then
                            negative_amounts_present = True
                        End If

                    Next

                    If negative_amounts_present And EMsg = "" Then
                        'Removed by request from Melinda 11/01/2011
                        'If MsgBox("There are Negative Checks in this Batch." & vbCr & "You can see them by sorting the Amount Column." & vbCr & vbCr & "Click YES to Proceed, or NO to return to edit the batch.", MsgBoxStyle.YesNo, "WARNING: Negative Amount Payments") = MsgBoxResult.No Then
                        '    Exit Sub
                        'End If
                    End If
                End If

                If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                 > Format(rowGLTPARM2.Item("PRD_END_DATE"), "yyyyMMdd") Then
                    EMsg &= vbCr & "Payment Receipt Date cannot be later than Current Period End Date (" & Format(rowGLTPARM2.Item("PRD_END_DATE"), "MM/dd/yyyy") & ")"
                Else
                    Dim rowGLTPARM2_LYP = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                    If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                    <= Format(rowGLTPARM2_LYP.Item("PRD_END_DATE"), "yyyyMMdd") Then

                        If MsgBox("Payment Receipt Date is prior to Current Period Start Date (" & Format(CDate(rowGLTPARM2_LYP.Item("PRD_END_DATE")).AddDays(1), "MM/dd/yyyy") & ")" _
                                  & vbCrLf & vbCrLf & "Proceed with this Entry?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                 < Format(CDate(rowGLTPARM2.Item("PRD_END_DATE")).AddYears(-1), "yyyyMMdd") Then
                    EMsg &= vbCr & "Payment Receipt Date cannot be more than 1 year ago"
                End If

            Case "Delete"
                If MsgBox("Are you sure you want to Delete?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If
            Case "Load Excel"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Load Excel"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Ask You For An Excel")
                iMSG.AppendLine("To Auto-Fill The Grid Below.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Are You Ready?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg &= vbCr & "Waiting For You To Be Ready For Auto-Fill."
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Load Excel"
                Dim FILENAME As String = ""
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Auto-Fill With"
                    'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
                    Dim filter As String = "All files (*.*)|*.*"
                    openFileDialog1.Filter = filter
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using
                If FILENAME = "" Then
                    MsgBox("Invalid File", vbCritical, "Did You Pick A File?")
                Else
                    Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
                    Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                    Dim XWS As New Microsoft.Office.Interop.Excel.Worksheet
                    Dim WSFound As Boolean = False
                    XWS = XWB.Worksheets(1)
                    Dim errMsg As New Text.StringBuilder With {.Length = 0}
                    Dim CUR_ROW As Int64 = 1
                    Dim BlankCount As Int64 = 0
                    Dim BAD_CUSTS As New List(Of String)
                    Dim PYMT_BATCH_NO As String = Absx1.txtFor("PYMT_BATCH_NO").Text.ToString & String.Empty
                    Dim PYMT_BATCH_LNO As Int64 = Val(dst.Tables("ARTPYMT2").Compute("MAX(PYMT_BATCH_LNO)", "").ToString & String.Empty)
                    If errMsg.Length = 0 Then
                        For i As Int64 = CUR_ROW To 1000
                            CUR_ROW += 1
                            Dim CUST_CODE As String = getValFromStr(XWS.Cells(CUR_ROW, 1).text.ToString & String.Empty)
                            If CUST_CODE.Length = 0 Or CUST_CODE = "0" Then
                                BlankCount += 1
                                If BlankCount = 10 Then
                                    Exit For
                                End If
                            Else
                                If CUST_CODE.Length <> 6 Then
                                    CUST_CODE = CUST_CODE.PadLeft(6, "0")
                                End If
                                SQLS.Length = 0
                                SQLS.AppendLine($"Select Count(*) AS RECS from ARTCUST1 where CUST_CODE = '{CUST_CODE}'")
                                ASCMAIN1.sql = SQLS.ToString()
                                Dim RECS As Int16 = Val(ASCDATA1.GetDataValue)
                                If RECS <> 1 Then
                                    BAD_CUSTS.Add(CUST_CODE)
                                Else
                                    PYMT_BATCH_LNO += 1
                                    SQLS.Length = 0
                                    SQLS.AppendLine($"Select CUST_NAME from ARTCUST1 where CUST_CODE = '{CUST_CODE}'")
                                    ASCMAIN1.sql = SQLS.ToString()
                                    Dim CUST_NAME As String = ASCDATA1.GetDataValue
                                    Dim CUST_PYMT_REF_NO As String = getValFromStr(XWS.Cells(CUR_ROW, 2).text.ToString & String.Empty)
                                    Dim CUST_PYMT_REF_DATE As Date = Absx1.dteFor("PYMT_BATCH_DATE").DateTime
                                    Dim CUST_PYMT_REF_DATE_S As String = XWS.Cells(CUR_ROW, 3).text.ToString & String.Empty
                                    If IsDate(CUST_PYMT_REF_DATE_S) Then
                                        CUST_PYMT_REF_DATE = CDate(CDate(CUST_PYMT_REF_DATE_S).ToShortDateString)
                                    End If
                                    Dim CUST_PYMT_AMT As Double = 0
                                    Dim CUST_PYMT_AMT_S As String = XWS.Cells(CUR_ROW, 4).text.ToString & String.Empty
                                    CUST_PYMT_AMT_S = CUST_PYMT_AMT_S.Replace("$", "").Replace(",", "")
                                    If IsNumeric(CUST_PYMT_AMT_S) Then
                                        CUST_PYMT_AMT = CUST_PYMT_AMT_S
                                    End If

                                    Dim newARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
                                    newARTPYMT2.Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                                    newARTPYMT2.Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO
                                    newARTPYMT2.Item("CUST_CODE") = CUST_CODE
                                    newARTPYMT2.Item("CUST_NAME") = CUST_NAME
                                    newARTPYMT2.Item("CUST_PYMT_REF_NO") = CUST_PYMT_REF_NO
                                    newARTPYMT2.Item("CUST_PYMT_REF_DATE") = CUST_PYMT_REF_DATE
                                    newARTPYMT2.Item("CUST_PYMT_AMT") = CUST_PYMT_AMT
                                    newARTPYMT2.Item("PYMT_STATUS") = 0
                                    newARTPYMT2.Item("CUST_PYMT_AMT_CURR") = CUST_PYMT_AMT
                                    newARTPYMT2.Item("CURR_CODE") = "USD"
                                    newARTPYMT2.Item("CURR_EXCH_RATE") = 1
                                    dst.Tables("ARTPYMT2").Rows.Add(newARTPYMT2)

                                End If
                            End If
                        Next
                        excel.DisplayAlerts = False
                        XWB.Save()
                        excel.Quit()
                        excel = Nothing
                        If BAD_CUSTS.Count > 0 Then
                            Dim iTitle As String = "Bad Customers"
                            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                            iMSG.AppendLine("The Following Bad Customers Were Skipped:")
                            For Each BAD_CUST As String In BAD_CUSTS
                                iMSG.AppendLine(BAD_CUST)
                            Next
                            MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
                        End If
                        MsgBox("Import Complete", vbOKOnly, "Done")
                        grdARTPYMT2.Update()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = Not This_Record_Inquiry_Only
                    .Items("Done").Visible = This_Record_Inquiry_Only

                    .Items("Update").Visible = ScreenMode And Not This_Record_Inquiry_Only
                    .Items("Delete").Visible = ScreenMode And Not This_Record_Inquiry_Only And (EntryMode = "E")

                    If ASCMAIN1.CLIENT = "RGI" Then
                        .Items("Load Excel").Visible = True
                        .Items("Load Excel").Settings.Enabled = iScreenMode
                    Else
                        .Items("Load Excel").Visible = False
                    End If
                End With

                .Groups("Batch Options").Visible = Not ScreenMode
                .Groups("Find Customer").Visible = ScreenMode And Not This_Record_Inquiry_Only

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpCURR_CODE, True)

        If Not This_Record_Inquiry_Only Then
            Absx1.dteFor("PYMT_BATCH_DATE").ReadOnly = False
        End If

        grdARTPYMT2.Visible = ScreenMode
        grdARTPYMTX.Visible = Not ScreenMode

        If ScreenMode Then
            With grdARTPYMT2.DisplayLayout.Override
                If This_Record_Inquiry_Only Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    .AllowUpdate = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                End If
            End With

            Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
            Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

            Dim COLUMN_NAME As String = "CUST_PYMT_AMT_CURR"
            Dim summary As UltraWinGrid.SummarySettings
            For Each summary In grdARTPYMT2.DisplayLayout.Bands(0).Summaries
                If summary.Key = "CUST_PYMT_AMT_CURR" Then
                    grdARTPYMT2.DisplayLayout.Bands(0).Summaries.Remove(summary)
                End If
            Next

            'If CURR_CODE <> GL_PARM_CURR_CODE Then
            Create_Summary(grdARTPYMT2, New String() {"CUST_PYMT_AMT_CURR"})
            'End If

            With grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_PYMT_AMT_CURR").Header
                If CURR_CODE = GL_PARM_CURR_CODE Then
                    .Caption = "Amount"
                Else
                    .Caption = CURR_CODE & " Amount"
                End If
            End With

            grdARTPYMT2.DisplayLayout.Bands(0).Columns("CURR_CODE").Hidden = (CURR_CODE <> GL_PARM_CURR_CODE) Or Not blnFX_Support
            grdARTPYMT2.DisplayLayout.Bands(0).Columns("CURR_EXCH_RATE").Hidden = (CURR_CODE <> GL_PARM_CURR_CODE) Or Not blnFX_Support
            grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_PYMT_AMT").Hidden = (CURR_CODE <> GL_PARM_CURR_CODE) Or Not blnFX_Support
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ARTPYMT1", "ARTPYMT2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Absx1.txtFor("BANK_CODE").Text = "" Then
            Absx1.txtFor("BANK_CODE").Text = ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE")
            Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            Absx1.numFor("CURR_EXCH_RATE").Value = 1
        End If

        Load_ARTPYMTX()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            HFs("PYMT_BATCH_NO") = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        End If

        Dim rowARTPYMT1 As DataRow = Fill_Record("ARTPYMT1", HFs("PYMT_BATCH_NO"), EntryMode = "N")
        If EntryMode = "N" Then
            rowARTPYMT1.Item("BANK_CODE") = HFs("BANK_CODE")
            rowARTPYMT1.Item("CURR_CODE") = HFs("CURR_CODE")
            rowARTPYMT1.Item("CURR_EXCH_RATE") = HFs("CURR_EXCH_RATE")
            rowARTPYMT1.Item("PYMT_BATCH_DATE") = HFs("PYMT_BATCH_DATE")
            rowARTPYMT1.Item("STATUS") = "0"
            rowARTPYMT1.Item("PYMT_SOURCE") = "MAN"
            rowARTPYMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        Else
            If rowARTPYMT1.Item("STATUS") & "" <> "0" Then
                This_Record_Inquiry_Only = True
            Else
                This_Record_Inquiry_Only = False
            End If
        End If
        Fill_Records("ARTPYMT2", HFs("PYMT_BATCH_NO"))

        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT2.Rows
            If grow.Cells("CUST_CODE").Value & "" = "" Then
                grdARTPYMT2.ActiveRow = grow
                grow.Cells("NON_AR").Value = "1"
                grow.Update()
            End If
        Next
        Sort_grdColumns(grdARTPYMT2, "PYMT_BATCH_LNO")

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        grdARTPYMT2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTPYMT2.DisplayLayout.Bands(0).SortedColumns.Add("PYMT_BATCH_LNO", False)

    End Sub

    Sub Delete_Record()
        BeginTrans()

        Delete_Rows("ARTPYMT1")
        Delete_Rows("ARTPYMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")

        CommitTrans("Records Deleted")
    End Sub

    Sub Update_Record()
        BeginTrans()

        Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
        If CURR_CODE = "" Then
            CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        End If
        Dim CURR_EXCH_RATE As Decimal = Val(Absx1.numFor("CURR_EXCH_RATE").Value & "")
        If CURR_EXCH_RATE = 0 Or CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then CURR_EXCH_RATE = 1

        For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Select("")
            Dim CURR_EXCH_RATE2 As Decimal = Val(rowARTPYMT2.Item("CURR_EXCH_RATE") & "")
            Dim CURR_CODE2 As String = rowARTPYMT2.Item("CURR_CODE")
            If CURR_CODE2 = "" Then
                CURR_CODE2 = CURR_CODE
                rowARTPYMT2.Item("CURR_CODE") = CURR_CODE2
                CURR_EXCH_RATE2 = CURR_EXCH_RATE
            End If
            If CURR_CODE2 = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT2.Item("CURR_CODE") = CURR_CODE2
                CURR_EXCH_RATE2 = 1
            End If
            rowARTPYMT2.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE2
            rowARTPYMT2.Item("CUST_PYMT_AMT") = Val(rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") & "") * CURR_EXCH_RATE2
            rowARTPYMT2.Item("INIT_DATE") = DATETIME_STAMP
            rowARTPYMT2.Item("INIT_OPER") = ASCMAIN1.USER_ID
        Next

        INIT_LAST("ARTPYMT1", , , True)

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")

        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PYMT_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTPYMTX, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdARTPYMT3"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCPA1" Then

            Select Case e.Tool.Key

            End Select
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Click_Command("Edit")

            Case "INV_NUM"
                Dim INV_NUM As String = ASCMAIN1.CodeSelector.SelectedCode
                Dim INV_TYPE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_TYPE")
                Find_Customer_by_Invoice(INV_NUM, INV_TYPE)
        End Select
    End Sub
    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If Absx1.txtFor("BANK_CODE").Text <> "" Then
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text.ToUpper)
                    If rowGLTBANK1 IsNot Nothing AndAlso rowGLTBANK1.Item("CURR_CODE") & "" <> Absx1.txtFor("CURR_CODE").Text Then
                        Absx1.txtFor("CURR_CODE").Text = rowGLTBANK1.Item("CURR_CODE") & ""

                        Get_CURR_EXCH_RATE()
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(sender, UltraWinEditors.UltraDateTimeEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(dtectl)
        Select Case COLUMN_NAME
            Case "PYMT_BATCH_DATE"
                'If Not ScreenMode Then
                Get_CURR_EXCH_RATE()
                'End If
        End Select
    End Sub
#End Region
    Sub Get_CURR_EXCH_RATE()
        If ScreenMode Then
            Exit Sub
        End If

        'Dim CURR_EXCH_RATE As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value)
        Dim CURR_EXCH_RATE As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me.ROWs("GLTPARM1"), Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value)
        Absx1.numFor("CURR_EXCH_RATE").Value = CURR_EXCH_RATE

        'If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
        '    Absx1.numFor("CURR_EXCH_RATE").Value = 1
        'Else
        '    ASCMAIN1.sql = "Select * from TATCURR3" & vbCrLf _
        '      & " where CURR_CODE = :PARM1" & vbCrLf _
        '      & "   and  CURR_DATE = (Select Max(CURR_DATE) from TATCURR3" & vbCrLf _
        '      & " where CURR_CODE = :PARM2 and CURR_DATE <= :PARM3)"
        '    Dim rowTATCURR3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVD", New Object() {Absx1.txtFor("CURR_CODE").Text, Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value})
        '    If rowTATCURR3 IsNot Nothing Then
        '        Dim X As Integer = 30
        '        Dim CURR_DATE As Date = rowTATCURR3.Item("CURR_DATE")
        '        Dim PYMT_BATCH_DATE As Date = Absx1.dteFor("PYMT_BATCH_DATE").DateTime
        '        Dim DAYSOLD As Integer = PYMT_BATCH_DATE.Subtract(CURR_DATE).Days
        '        Dim CURR_EXCH_RATE As Decimal = Val(rowTATCURR3.Item("CURR_EXCH_RATE") & "")
        '        If DAYSOLD < X Then
        '            Absx1.numFor("CURR_EXCH_RATE").Value = CURR_EXCH_RATE
        '        Else
        '            Absx1.numFor("CURR_EXCH_RATE").Value = 0
        '        End If
        '    Else
        '        Absx1.numFor("CURR_EXCH_RATE").Value = 0
        '    End If
        'End If
    End Sub


#Region "grdARTPYMT2"
    Private Sub grdARTPYMT2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTPYMT2.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_PYMT_REF_DATE"
                If "" = "" Then
                    Dim currYear As String = Mid(ASCMAIN1.CYP, 1, 4)
                    'grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
                End If
        End Select
    End Sub
    Private Sub grdARTPYMT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                If e.Cell.Text = "" Then
                    grdARTPYMT2.ActiveRow.Cells("NON_AR").Value = "1"
                    grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = Absx1.txtFor("CURR_CODE").Text
                    grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = Absx1.numFor("CURR_EXCH_RATE").Value
                Else
                    grdARTPYMT2.ActiveRow.Cells("NON_AR").Value = "0"
                    grdCodeDesc(grdARTPYMT2, "ARTCUST1", "CUST_CODE", "CUST_NAME")

                    If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        grdCodeDesc(grdARTPYMT2, "ARTCUST1", "CUST_CODE", "CURR_CODE")
                        If grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                            grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = 1
                        End If
                    Else
                        grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = Absx1.txtFor("CURR_CODE").Text
                        grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = Absx1.numFor("CURR_EXCH_RATE").Value
                    End If

                    If grdARTPYMT2.ActiveRow.Cells("CUST_NAME").Text = "" Then
                        grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
                    Else
                        grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                        'grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                    End If
                End If

            Case "CURR_CODE"
                If grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value & "" = "USD" Then
                    grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = 1
                    grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                End If

            Case "CURR_EXCH_RATE", "CUST_PYMT_AMT_CURR"
                Dim CUST_PYMT_AMT_CURR As Decimal = Val(grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_AMT_CURR").Value & "")
                Dim CURR_EXCH_RATE As Decimal = Val(grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value & "")
                grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_AMT").Value = CUST_PYMT_AMT_CURR * CURR_EXCH_RATE

            Case "NON_AR"
                If e.Cell.Value & "" = "1" Then
                    grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = ""
                    grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
        End Select
    End Sub

    Private Sub grdARTPYMT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT2.AfterRowActivate
        With grdARTPYMT2.DisplayLayout.Bands(0)
            If grdARTPYMT2.ActiveRow.IsAddRow Then
                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
                grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                If grdARTPYMT2.ActiveRow.Cells("NON_AR").Text = "1" Then
                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If
        End With
    End Sub

    Private Sub grdARTPYMT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT2.BeforeRowUpdate
        With grdARTPYMT2
            If e.Row.Cells("NON_AR").Value = "1" Then
                If e.Row.Cells("CUST_NAME").Text = "" Then
                    MsgBox("You Must Enter a Name for Non-AR Payments", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            Else
                If e.Row.Cells("CUST_CODE").Text = "" Then
                    MsgBox("Missing Value for Customer Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
                    If cdr Is Nothing Then
                        MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If
            Dim currCode As String = e.Row.Cells("CURR_CODE").Value & ""
            If currCode = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                If Val(e.Row.Cells("CURR_EXCH_RATE").Value & "") <> 1 Then
                    MsgBox("Invalid Value entered for Currency Exhange Rate (" & e.Row.Cells("CURR_EXCH_RATE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = Absx1.CtlFor("PYMT_BATCH_NO").Text
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = Val(dst.Tables("ARTPYMT2").Compute("Max(PYMT_BATCH_LNO)", "") & "") + 1
                    .ActiveRow.Cells("PYMT_STATUS").Value = "0"
                    If e.Row.Cells("NON_AR").Value = "1" Then
                    Else
                        LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value)
                        e.Row.Cells("CUST_NAME").Value = cdr.Item("CUST_NAME")
                    End If

                End If
            End If

        End With
    End Sub

    Private Sub grdARTPYMT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTPYMT2, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT2.Error
        grdARTPYMT2.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdARTPYMT2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdARTPYMT2.KeyPress
        If grdARTPYMT2.ActiveRow IsNot Nothing Then
            Try
                If grdARTPYMT2.ActiveCell.Column.Key = "CUST_NAME" Then
                    If grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Text <> "" Then
                        e.KeyChar = Chr(0)
                        e.Handled = True
                    End If
                ElseIf grdARTPYMT2.ActiveCell.Column.Key = "CUST_PYMT_REF_DATE" Then

                    If e.KeyChar = "" Then
                        If "" = "" Then
                            Dim year As String = Mid(ASCMAIN1.CYP, 0, 4)
                            grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
                        End If
                    End If

                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub grdARTPYMT2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTPYMT2.BeforeExitEditMode
        With grdARTPYMT2.ActiveCell
            Select Case .Column.Key

                Case "CUST_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If

                Case "CURR_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If

                Case "CURR_EXCH_RATE"
                    If .Row.Cells("CURR_CODE").Value & "" = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        If Val(.Text) <> 1 Then
                            .Value = 1
                        End If
                    End If

                Case "CUST_PYMT_REF_DATE"
                    'Stop
                    Dim DT As String = .EditorResolved.CurrentEditText
                    If grdARTPYMT2.ActiveCell Is Nothing Then
                    Else
                        Try
                            If Len(DT) = 10 And Mid(DT, 7, 4) = "    " And DT <> "  /  /    " Then
                                .Value = Mid(DT, 1, 6) & Now.Year
                            End If
                        Catch ex As Exception

                        End Try
                    End If


            End Select
        End With
    End Sub
#End Region

    Private Sub grdARTPYMTX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTPYMTX.DoubleClickRow
        If grdARTPYMTX.ActiveRow IsNot Nothing AndAlso grdARTPYMTX.ActiveRow.IsDataRow Then
            Absx1.txtFor("PYMT_BATCH_NO").Text = grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_NO").Text
            Click_Command("Edit")
        End If

    End Sub

    Private Sub chkEditableOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEditableOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        cbeYP.Visible = Not chkEditableOnly.Checked
        Load_ARTPYMTX()
    End Sub

    Private Sub chkMyOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMyOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ARTPYMTX()
    End Sub

    Private Sub txtInvoice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.Enter
        txtInvoice.Text = ""
        txtCUST_CODE.Text = ""
    End Sub

    Private Sub txtInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Dim INV_NO As String = ASCMAIN1.Format_Field(txtInvoice.Text, "INV_NO")

            If ASCMAIN1.SOLUTION = "SEA" Then
                txtInvoice.Text = INV_NO
                Dim sql As String = "Select CUST_BILL_TO_CUST from SOTINVH1 where ORDR_INV_NO = '" & INV_NO & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow(sql)
                If row Is Nothing Then
                    txtCUST_CODE.Text = ""
                Else
                    txtCUST_CODE.Text = row.Item("CUST_BILL_TO_CUST")
                    grdARTPYMT2.DisplayLayout.Bands(0).AddNew()
                    grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
                    grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = txtCUST_CODE.Text
                    'grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                End If
            Else
                INV_NO = INV_NO.PadLeft(10, "0")
                txtInvoice.Text = INV_NO
                Find_Customer_by_Invoice(INV_NO, "I")
            End If
        End If
    End Sub

    Sub Find_Customer_by_Invoice(INV_NO As String, INV_TYPE As String)

        Dim row As DataRow = Nothing
        Dim CUST_CODE As String = ""
        Dim sql As String = ""

        sql = "Select NVL(CUST_BILL_TO_CUST,CUST_CODE) CUST_CODE from SOTINVH1 where INV_NO = '" & INV_NO & "'"
        row = ASCDATA1.GetDataRow(sql)
        If row IsNot Nothing Then
            CUST_CODE = row.Item("CUST_CODE") & ""
        End If
        If row Is Nothing Then
            sql = "Select CUST_CODE from ARTOPEN1" _
                & " where INV_TYPE = '" & INV_TYPE & "' and INV_NUM = '" & INV_NO & "'"
            row = ASCDATA1.GetDataRow(sql)
            If row IsNot Nothing Then
                CUST_CODE = row.Item("CUST_CODE")
            End If
        End If
        txtCUST_CODE.Text = CUST_CODE
        If CUST_CODE <> "" Then
            grdARTPYMT2.DisplayLayout.Bands(0).AddNew()
            grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
            grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
            grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
        End If
    End Sub

    Sub Load_ARTPYMTX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Payment Batch Data")

        Dim OPS_YYYYPP As String = ""

        Dim STATUS As String = "0"
        If Not chkEditableOnly.Checked Then
            STATUS = "%"
            OPS_YYYYPP = cbeYP.Value
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", OPS_YYYYPP)
            grdARTPYMTX.Text = "Payments Entered in " & rowGLTPARM2.Item("LEGEND")
        Else
            OPS_YYYYPP = ASCMAIN1.CYP
            grdARTPYMTX.Text = "Payments Entered, awaiting Payment Receipts Journal"
        End If

        Dim USER_ID As String = ASCMAIN1.USER_ID
        If Not chkMyOnly.Checked Then
            USER_ID = "%"
        Else
            grdARTPYMTX.Text &= ", by " & USER_ID
        End If

        Fill_Records("ARTPYMTX", New String() {STATUS, USER_ID, OPS_YYYYPP})
        Sort_grdColumns(grdARTPYMTX, "PYMT_BATCH_NO")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ARTPYMTX()
    End Sub

    Private Function getValFromStr(ByVal inStr As String) As Decimal
        Dim RetVal As Decimal = 0
        Dim isNeg As Boolean = False
        If inStr.Contains("(") And inStr.Contains(")") Then
            isNeg = True
        End If
        Dim REP As String() = {",", "$", "(", ")"}
        For Each R As String In REP
            inStr = inStr.Replace(R, "")
        Next
        If IsNumeric(inStr) Then
            If isNeg And Val(inStr) > 0 Then
                RetVal = Val(inStr) * -1
            Else
                RetVal = Val(inStr)
            End If
        End If
        Return RetVal
    End Function
End Class