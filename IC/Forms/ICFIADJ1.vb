Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid
Imports Infragistics.Win.UltraWinTabControl

Public Class ICFIADJ1
    ' SHOULD PROBABLY ADD A LOCATION_CODE TO ICTIADJ1 AND PROMPT FOR IT IF THE ADJ WHSE IS A LOCATOR - THEN CHG SP TO NOT USE DEFAULT LOC FOR ADJ
    Dim rowICTIADJ1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim tblADJ_REF As String = String.Empty
    Dim CYCLE_STATUS As String = ""
    Dim CYCLE_TYPE As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFIADJI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTIADJ1.*" _
            & " from ICTIADJ1 where ICTIADJ1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTIADJX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            & ", ICTIADJ1.ADJ_DATE, ICTIADJ1.WHSE_CODE, ICTIADJ1.REASON_CODE" _
            & ", ICTIADJ1.ADJ_NOTE, ICTIADJ1.INIT_OPER, ICTIADJ1.INIT_DATE" _
            & ", ICTIADJ1.ADJ_SOURCE, ICTIADJ1.OPS_YYYYPP, ICTIADJ1.RTRN_NO" _
            & " from ICTIADJ1,ICTIADJ3,GLTACCT1 where ICTIADJ1.OPS_YYYYPP = :PARM1" _
            & " and GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE" _
            & " and ICTIADJ3.ADJ_NO = ICTIADJ1.ADJ_NO"
            Create_TDA(.Tables.Add, "ICTIADJG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTIADJ1", "*")

            ASCMAIN1.sql = "Select ICTIADJ2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from ICTIADJ2,ICTSTYL1,ICTCOLR1 where ICTSTYL1.STYLE_CODE = ICTIADJ2.STYLE_CODE" _
            & " and ICTCOLR1.COLOR_CODE = ICTIADJ2.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ2", "**", 1)
            .Tables("ICTIADJ2").Columns.Add("LINE_COSTS", GetType(System.Decimal), "ISNULL(ADJ_QTY,0) * ISNULL(STYLE_COST,0)")

            ASCMAIN1.sql = "Select ICTIADJ3.*, GLTACCT1.ACCT_DESC" _
            & " from ICTIADJ3,GLTACCT1 where GLTACCT1.ACCT_CODE = ICTIADJ3.ACCT_CODE"
            Create_TDA(.Tables.Add, "ICTIADJ3", "**", 1)

            ASCMAIN1.sql = "Select ICTSTAT2.*" _
            & " from ICTSTAT2 where STYLE_CODE = :PARM1 and WHSE_CODE = :PARM2"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "VV")

            .Tables.Add("ICTIADJ0")
            .Tables("ICTIADJ0").Columns.Add("KEY")
            .Tables("ICTIADJ0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTREAS1"
            Create_TDA(.Tables.Add, "ICTREAS1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                'ASCMAIN1.sql = "SELECT DISTINCT WHTCYCL1.* FROM WHTCYCL2,WHTCYCL1" _
                '& " WHERE WHTCYCL1.CYCLE_NO = WHTCYCL2.CYCLE_NO AND NVL(WHTCYCL1.UPDATED_INV_ADJ,0) = 0 AND WHTCYCL1.CYCLE_TYPE = :PARM1" _
                '& " AND WHTCYCL1.CYCLE_RESOLUTION = :PARM2" _
                '& " And WHTCYCL1.INIT_DATE >= :PARM3"
                'Create_TDA(.Tables.Add, "WHTCYCL1", "**", 0, False, "VVD")


                ASCMAIN1.sql = "SELECT DISTINCT WHTCYCL1.* FROM WHTCYCL2,WHTCYCL1" _
                & " WHERE WHTCYCL1.CYCLE_NO = WHTCYCL2.CYCLE_NO AND NVL(WHTCYCL1.UPDATED_INV_ADJ,0) <> 'X'" _
                & " AND WHTCYCL1.CYCLE_TYPE ='V' AND WHTCYCL1.CYCLE_STATUS = 'D'" _
                & " And WHTCYCL1.INIT_DATE >= :PARM1"
                Create_TDA(.Tables.Add, "WHTCYCL1", "**", 0, True, "D")
                .Tables("WHTCYCL1").Columns.Add("SEL")
                .Tables("WHTCYCL1").Columns("SEL").DefaultValue = "0"


                ASCMAIN1.sql = "SELECT * FROM WHTLOCB2 WHERE WHSE_TRAN_TYPE= 'M'" _
                & " AND WHSE_TRAN_NO = :PARM1"
                Create_TDA(.Tables.Add, "WHTLOCB2", "**", 0, False, "V")

            End If




            tblADJ_REF = ASCMAIN1.Temp_Table("SELECT DISTINCT UPPER(ADJ_REF) ADJ_REF FROM ICTIADJ2 WHERE ADJ_REF IS NOT NULL AND LENGTH(ADJ_REF) > 2")
            Create_TDA(.Tables.Add, tblADJ_REF, "*", 0, False)
            Fill_Records(tblADJ_REF)


            If ASCMAIN1.CLIENT = "VAN" And Not InquiryMode Then
                With .Tables.Add("ERROR_TBL")
                    .Columns.Add("SKU", GetType(System.String))
                    .Columns.Add("ERROR_DETAIL", GetType(System.String))
                End With
            End If

        End With

        Set_Read_Only(grpTotals, True)

        Dim CUTOFF As Date = Nothing

        dteDATE_CUTOFF.Value = "06-FEB-2023"
        CUTOFF = dteDATE_CUTOFF.Value
        OptResolution.Value = "U"
        Dim CYCLE_RESOLUTION As String = "U"
        Dim CYCLE_TYPE As String = "V"


        Fill_Records("ICTREAS1")
        Fill_Records("ICTCLAS1")

        If ASCMAIN1.CLIENT = "VAN" Then
            Fill_Records("WHTCYCL1", New Object() {CUTOFF})
        End If


        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIADJ0.DataSource = dst.Tables("ICTIADJ0")
        grdICTIADJ2.DataSource = dst.Tables("ICTIADJ2")
        grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
        grdICTIADJX.DataSource = dst.Tables("ICTIADJX")
        grdICTIADJG.DataSource = dst.Tables("ICTIADJG")
        If ASCMAIN1.CLIENT = "VAN" Then
            grdWHTCYCL1.DataSource = dst.Tables("WHTCYCL1")
            grdWHTLOCB2.DataSource = dst.Tables("WHTLOCB2")
        End If

        Create_Summary(grdICTIADJX, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJX, "TOTAL_COSTS")

        Create_Summary(grdICTIADJG, "ADJ_NO", "Count")
        Create_Summary(grdICTIADJG, "DIST_AMT")

        Create_Summary(grdICTIADJ2, "ADJ_LNO", "Count")
        Create_Summary(grdICTIADJ2, "ADJ_QTY")
        Create_Summary(grdICTIADJ2, "LINE_COSTS")

        Create_Summary(grdICTIADJ3, "ADJ_GNO", "Count")
        Create_Summary(grdICTIADJ3, "DIST_AMT")
        Create_Summary(grdWHTCYCL1, "CYCLE_NO", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_NO", "Count")
        Create_Summary(grdWHTLOCB2, "WHSE_TRAN_QTY")

        With grdICTIADJX.DisplayLayout.Bands("ICTIADJX")
            .Columns("ADJ_NO").Header.Fixed = True
        End With

        With grdICTIADJG.DisplayLayout.Bands("ICTIADJG")
            .Columns("ADJ_NO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdICTIADJX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        ASCMAIN1.Add_Value_List(grdICTIADJG, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTIADJ0.DisplayLayout.Bands(0).ColHeadersVisible = False
        Set_SEGS(grdICTIADJ3, "ICTIADJ3")

        If ASCMAIN1.CLIENT = "VAN" Then
            With grdWHTCYCL1.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.True
            End With
            With grdWHTCYCL1.DisplayLayout.Bands(0)
                .Columns("WHSE_TRAN_NO").CellAppearance.BackColor = Color.Beige
            End With
            With grdWHTLOCB2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.False
            End With

            With grdWHTLOCB2.DisplayLayout.Bands(0)
                .Columns("WHSE_TRAN_NO").CellAppearance.BackColor = Color.Beige
            End With

            With grdWHTCYCL1.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key = "SEL" Then
                        gcol.CellAppearance.BackColor = Drawing.Color.LightGreen
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    End If
                Next
            End With



        End If







        Set_Read_Only(grpTotals, True)
        If InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0 Then
            grpTotals.Visible = False
            With grdICTIADJ2.DisplayLayout.Bands(0)
                .Columns("STYLE_COST").Hidden = True
                .Columns("LINE_COSTS").Hidden = True
                .Columns("STYLE_CLASS_CODE").Hidden = True
                .Columns("SALES_DIVISION_CODE").Hidden = True
            End With
        End If



        grpHeader.Visible = False
        Set_SEGS(grdICTIADJG, "ICTIADJG")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("WHSE_CODE")

                If Absx1.dteFor("ADJ_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Invalid Date Specified for Entry"
                End If

                If Absx1.txtFor("WHSE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        Else
                            If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                                If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                                    If MsgBox("Your entry will cause an out of balance to result with the 3PL." _
                                              & vbCrLf & vbCrLf & "Is this entry authorized by Gabe?",
                                              MsgBoxStyle.YesNo,
                                              "Warning: Warehouse Entered Is A 3PL") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                Else
                                    If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
                                    Else
                                        EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Adjustments Allowed"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            Case "View"
                If Absx1.txtFor("ADJ_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTIADJ1 = LookUp("ICTIADJ1", Absx1.txtFor("ADJ_NO").Text)
                    If rowICTIADJ1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("ADJ_NO").Text & " on File"
                    End If
                End If

            Case "Update"
                If Absx1.txtFor("REASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Reason"
                Else
                    Dim rowICTREAS1 As DataRow = LookUp("ICTREAS1", Absx1.txtFor("REASON_CODE").Text)
                    If rowICTREAS1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Reason"
                    End If
                End If

                If grdICTIADJ2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    For Each rowICTIADJ2 As DataRow In dst.Tables("ICTIADJ2").Select("", "", DataViewRowState.CurrentRows)
                        If rowICTIADJ2.Item("STYLE_CLASS_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Class for " & rowICTIADJ2.Item("STYLE_CODE") & ""
                        End If
                        If rowICTIADJ2.Item("SALES_DIVISION_CODE") & "" = "" Then
                            EMsg &= vbCr & "Unable to determine Division for " & rowICTIADJ2.Item("STYLE_CODE") & ""
                        End If
                    Next
                End If

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTIADJ2", Absx1.txtFor("WHSE_CODE").Text, "ADJ_QTY", 1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK to Continue Anyway?",
                                  MsgBoxStyle.YesNo,
                                  "The following Items do not have Sufficent Qty for this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Reverse"
                If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reverse"
                Set_Up_Reversal()
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") And Not InquiryMode _
                        AndAlso rowICTIADJ1 IsNot Nothing _
                        AndAlso rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") Is DBNull.Value _
                        AndAlso rowICTIADJ1.Item("REVERSES_ADJ_NO") Is DBNull.Value
                End With

                .Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Show if Entered in").Visible = Not ScreenMode ' And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = False ' ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
                .Groups("Damages").Visible = ScreenMode And EntryMode = "N" And (ASCMAIN1.Running_in_VS Or (ASCMAIN1.CLIENT = "RGI" And ASCMAIN1.USER_SECURITY_CODEs.Contains("WS")))
                .Groups("Special Functions").Visible = (ASCMAIN1.CLIENT = "VAN" And (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "wendy")) And Not ScreenMode And Not InquiryMode
                .Groups("Cycle Count Adjustment").Visible = False ' (ASCMAIN1.CLIENT = "VAN" And (ASCMAIN1.USER_ID = "dgj")) And Not ScreenMode And Not InquiryMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode

        tab0.Tabs(2).Visible = ASCMAIN1.CLIENT = "VAN"
        tab0.Tabs(2).Visible = False

        If ScreenMode Then

            With grdICTIADJ2.DisplayLayout.Bands(0)
                If (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") And EntryMode = "N" Then
                    .Columns("WHSE_QTY_ON_HAND").Hidden = False
                Else
                    .Columns("WHSE_QTY_ON_HAND").Hidden = True
                End If
            End With

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            SplitContainer2.Panel2Collapsed = (EntryMode <> "V") Or InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") = 0
            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIADJ2, grdICTIADJ3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
                With grdICTIADJ2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COLOR_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.LightYellow
                End With
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTIADJ2, grdICTIADJ3}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                With grdICTIADJ2.DisplayLayout.Bands(0)
                    .Columns("STYLE_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("COLOR_CODE").CellAppearance.BackColor = Color.Empty
                    .Columns("ADJ_QTY").CellAppearance.BackColor = Color.Empty
                End With
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTIADJ0", "ICTIADJ1", "ICTIADJ2", "ICTIADJ3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If chkGL.Checked Then
            chkGL.Checked = False
        Else
            Refresh_Documents()
        End If
        Setup_tab0_GL()

        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.dteFor("ADJ_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("ADJ_NO").Text = ""

        If ASCMAIN1.CLIENT = "VAN" Then
            CYCLE_TYPE = "V"
            CYCLE_STATUS = "D"
            OptResolution.Value = "A"
            OptResolution.Value = "U"

        End If

        optGL.Tag = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTIADJ1 = dst.Tables("ICTIADJ1").NewRow
            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
                rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("TRAN_NO_A")
            Else
                rowICTIADJ1.Item("ADJ_NO") = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
            End If
            rowICTIADJ1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTIADJ1.Item("ADJ_DATE") = HFs("ADJ_DATE")
            rowICTIADJ1.Item("ADJ_SOURCE") = "E"
            rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("REGISTER_IND") = "0"
            rowICTIADJ1.Item("JOURNAL_IND") = "0"
            dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)
        Else
            Fill_Record("ICTIADJ1", Absx1.txtFor("ADJ_NO").Text)
            dst.AcceptChanges()

            dst.Tables("ICTIADJ0").Rows.Add(New String() {"Entered", Format(rowICTIADJ1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
            dst.Tables("ICTIADJ0").Rows.Add(New String() {"By", rowICTIADJ1.Item("INIT_OPER")})
            dst.Tables("ICTIADJ0").Rows.Add(New String() {"Source", rowICTIADJ1.Item("ADJ_SOURCE")})

            If rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") & "" <> "" Then
                Dim row As DataRow = LookUp("ICTIADJ1", rowICTIADJ1.Item("REVERSED_BY_ADJ_NO"))
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"Reversed", Format(row.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"By", row.Item("INIT_OPER")})
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"using", rowICTIADJ1.Item("REVERSED_BY_ADJ_NO")})
            ElseIf rowICTIADJ1.Item("REVERSES_ADJ_NO") & "" <> "" Then
                dst.Tables("ICTIADJ0").Rows.Add(New String() {"Reverses", rowICTIADJ1.Item("REVERSES_ADJ_NO")})
            End If
        End If


        rowICTWHSE1 = LookUp("ICTWHSE1", rowICTIADJ1.Item("WHSE_CODE"))
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTIADJ2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = True ' Not location_support
            .Columns("LOCATION_CODE").Hidden = Not location_support
        End With

        Fill_Records("ICTIADJ2", Absx1.txtFor("ADJ_NO").Text)
        Fill_Records("ICTIADJ3", Absx1.txtFor("ADJ_NO").Text)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then

            ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                     New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
                     New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})
        End If
        CommitTrans("Update Complete")

    End Sub

    Sub Update_WHTLOCBX()


        Dim rowICTIADJ1 As DataRow = dst.Tables("ICTIADJ1").Rows(0)
        For Each row As DataRow In dst.Tables("ICTIADJ2").Select("")
            Dim TRAN_NO As String = row.Item("ADJ_NO")
            Dim TRAN_LNO As Integer = row.Item("ADJ_LNO")
            Dim WHSE_CODE As String = rowICTIADJ1.Item("WHSE_CODE")
            Dim BAR_CODE As String = "0000000000" ' row.Item("BAR_CODE")
            Dim LOCATION_CODE As String = row.Item("LOCATION_CODE")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim ADJ_QTY As Int64 = Val(row.Item("ADJ_QTY") & "")

            Dim rowWHTLOCB1 As DataRow = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            If rowWHTLOCB1 Is Nothing Then
                Fill_Records("WHTLOCB1", New String() {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE}, False)
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").Rows.Find(New Object() _
                                         {WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE})
            End If

            If rowWHTLOCB1 Is Nothing Then
                rowWHTLOCB1 = dst.Tables("WHTLOCB1").NewRow
                With rowWHTLOCB1
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("LOCATION_CODE") = LOCATION_CODE
                    .Item("BAR_CODE") = BAR_CODE
                    .Item("STYLE_CODE") = STYLE_CODE
                    .Item("COLOR_CODE") = COLOR_CODE
                    .Item("LOCATION_QTY") = ADJ_QTY
                End With
                dst.Tables("WHTLOCB1").Rows.Add(rowWHTLOCB1)
            Else
                rowWHTLOCB1.Item("LOCATION_QTY") = Val(rowWHTLOCB1.Item("LOCATION_QTY") & "") + ADJ_QTY
            End If

            Dim rowWHTLOCB2 As DataRow = dst.Tables("WHTLOCB2").NewRow
            With rowWHTLOCB2
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("LOCATION_CODE") = LOCATION_CODE
                .Item("BAR_CODE") = BAR_CODE
                .Item("STYLE_CODE") = STYLE_CODE
                .Item("COLOR_CODE") = COLOR_CODE
                .Item("WHSE_TRAN_QTY") = ADJ_QTY
                .Item("WHSE_TRAN_TYPE") = "A"
                .Item("WHSE_TRAN_NO") = TRAN_NO
                .Item("WHSE_TRAN_LNO") = TRAN_LNO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LOCATION_CODE_OTHER") = ""
                .Item("SESSION_ID") = ""
            End With
            dst.Tables("WHTLOCB2").Rows.Add(rowWHTLOCB2)
        Next

        Update_Record_TDA("WHTLOCB1")
        Update_Record_TDA("WHTLOCB2")

        dst.Tables("WHTLOCB1").Rows.Clear()
        dst.Tables("WHTLOCB2").Rows.Clear()
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub


    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("ADJ_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTIADJX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTIADJ2, "B", "Style Status Inquiry")
        Load_Popup_Menu(grdICTIADJG, "SS", "Show Filter", "Show GroupBox")
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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "ADJ_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                If Not InquiryMode Then
                    Click_Command("New")
                End If
            Case "ADJ_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTIADJ2"

    Private Sub grdICTIADJ2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIADJ2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

                grdCodeDesc(grdICTIADJ2, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE STYLE_DESC
                If cdr IsNot Nothing Then
                    Dim STYLE_CODE As String = e.Cell.Value
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = cdr.Item("SALES_DIVISION_CODE")
                    e.Cell.Row.Cells("STYLE_DESC").Value = cdr.Item("STYLE_DESC")
                    Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
                    Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
                    Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")


                    e.Cell.Row.Cells("STYLE_CLASS_CODE").Value = STYLE_CLASS_CODE
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = SALES_DIVISION_CODE
                    e.Cell.Row.Cells("STYLE_COST").Value = STYLE_COST
                    If location_support Then
                        e.Cell.Row.Cells("LOCATION_CODE").Value = rowICTWHSE1.Item("WHSE_LOC_ADJ")
                        ' USE ITEM_BIN AS A DEFAULT FOR AHA
                    End If

                    ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
                    Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
                    If rowICTSTYC1s.Length = 1 Then
                        e.Cell.Row.Cells("COLOR_CODE").Value = rowICTSTYC1s(0).Item("COLOR_CODE")
                    End If
                Else
                    grdICTIADJ2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COLOR_CODE"
                grdCodeDesc(grdICTIADJ2, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
                ' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
                If cdr IsNot Nothing Then
                    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")

                    If (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") And EntryMode = "N" Then
                        Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value
                        ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & cdr.Item("COLOR_CODE") & "' and WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
                        Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
                        If rowICTSTAT2 IsNot Nothing Then
                            e.Cell.Row.Cells("WHSE_QTY_ON_HAND").Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
                        End If
                    End If
                End If

            Case "ADJ_QTY"

        End Select
    End Sub

    Private Sub grdICTIADJ2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterExitEditMode
        'Select Case grdICTIADJ2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdICTIADJ2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterRowActivate
        With grdICTIADJ2.DisplayLayout.Bands(0)
            If grdICTIADJ2.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTIADJ2.ActiveCell = grdICTIADJ2.ActiveRow.Cells("STYLE_CODE")
                grdICTIADJ2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("COLOR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If EntryMode = "V" Then
            Show_GL()
        End If
    End Sub

    Private Sub grdICTIADJ2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTIADJ2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTIADJ2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTIADJ2.AfterRowUpdate
        DisplayTotals()
    End Sub


    Private Sub grdICTIADJ2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTIADJ2.BeforeExitEditMode
        If grdICTIADJ2.ActiveCell Is Nothing Then Exit Sub
        With grdICTIADJ2.ActiveCell
            Select Case .Column.Key
                Case "STYLE_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTSTYL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If

                Case "COLOR_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLR1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Color Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                        If Not e.Cancel Then
                            cdr = LookUp("ICTSTYC1", New String() { .Row.Cells("STYLE_CODE").Value, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Color Code (" & .Text & ") not set up with Style (" & .Row.Cells("STYLE_CODE").Value & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If

                Case "LOCATION_CODE"
                    If location_support Then
                        If .Text <> "" Then
                            If .Value IsNot Nothing Then
                                .Value = .Text.ToUpper
                            End If

                        End If
                        If .Text <> "" Then
                            cdr = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, .Text})
                            If cdr Is Nothing Then
                                ASCMAIN1.Progress("Invalid Location Code (" & .Text & ")")
                                If .Value IsNot Nothing Then
                                    .Value = ""
                                End If
                                e.Cancel = True
                            End If
                        End If
                    End If

            End Select
        End With
    End Sub

    Private Sub grdICTIADJ2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTIADJ2.BeforeRowUpdate
        With grdICTIADJ2
            If e.Row.Cells("STYLE_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Row.Cells("COLOR_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTCOLR1", e.Row.Cells("COLOR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Color Code (" & e.Row.Cells("COLOR_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
                If Not e.Cancel Then
                    LookUp("ICTSTYC1", New String() {e.Row.Cells("STYLE_CODE").Text, e.Row.Cells("COLOR_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Color Code (" & e.Row.Cells("COLOR_CODE").Text & ") not set up for Style (" & e.Row.Cells("STYLE_CODE").Text & ")",
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

                If e.Row.Cells("LOCATION_CODE").Text = "" Then
                    e.Cancel = True
                Else
                    LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, e.Row.Cells("LOCATION_CODE").Text})
                    If cdr Is Nothing Then
                        MsgBox("Invalid Value entered for Location Code (" & e.Row.Cells("LOCATION_CODE").Text & ")",
                               MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If

            End If

            If Val(e.Row.Cells("ADJ_QTY").Text) = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("ADJ_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("ADJ_NO").Text = "" Then
                    .ActiveRow.Cells("ADJ_NO").Value = Absx1.CtlFor("ADJ_NO").Text
                    .ActiveRow.Cells("ADJ_LNO").Value = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                End If

                Dim ADJ_REF As String = e.Row.Cells("ADJ_REF").Text
                ADJ_REF = ADJ_REF.Trim
                If ADJ_REF.Length > 0 Then
                    ADJ_REF = ADJ_REF.ToUpper
                    e.Row.Cells("ADJ_REF").Value = ADJ_REF
                    If dst.Tables(tblADJ_REF).Select("ADJ_REF = '" & ADJ_REF & "'").Length = 0 Then
                        ASCDATA1.ExecuteSQL("INSERT INTO " & tblADJ_REF & " VALUES ('" & ADJ_REF & "')")
                        dst.Tables(tblADJ_REF).Rows.Add(New Object() {ADJ_REF})
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub grdICTIADJ2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTIADJ2.ClickCellButton

        If grdICTIADJ2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "STYLE_CODE"

            Case "COLOR_CODE"
                sql_where = "COLOR_CODE in (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE = '" & e.Cell.Row.Cells("STYLE_CODE").Value & "')"

            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"

            Case "ADJ_REF"
                ASCMAIN1.CodeSelector.SQL = "SELECT * FROM " & tblADJ_REF

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.DoNotFilterFirst = True
                    ASCMAIN1.CodeSelector.VIEW_NAME = String.Empty
                    ASCMAIN1.CodeSelector.Custom_sql_where = String.Empty
                    ASCMAIN1.CodeSelector.Custom_sqlkey = String.Empty
                    ASCMAIN1.CodeSelector.ForceFilterFirst = False

                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        Dim ADJ_REF As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item(0)
                        grdICTIADJ2.ActiveRow.Cells("ADJ_REF").Value = ADJ_REF
                    End If
                End If

                Exit Sub
        End Select
        grdClickCellButton(grdICTIADJ2, sql_where, False)

    End Sub

    Private Sub grdICTIADJ2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTIADJ2.Error
        grdICTIADJ2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIADJ2").Compute("SUM(LINE_COSTS)", "") & "")
        Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTIADJX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIADJX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ADJ_NO").Text = e.Row.Cells("ADJ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTIADJG_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTIADJG.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("ADJ_NO").Text = e.Row.Cells("ADJ_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub optGL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optGL.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_GL()
    End Sub

    Sub Show_GL()

        If optGL.Tag <> optGL.Value Or optGL.Value = "L" Then
            optGL.Tag = optGL.Value
            If optGL.Value = "A" Then
                grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
                Dim dvw As DataView = dst.Tables("ICTIADJ3").DefaultView
                dvw.RowFilter = ""
            ElseIf optGL.Value = "L" Then
                grdICTIADJ3.DataSource = dst.Tables("ICTIADJ3")
                Dim dvw As DataView = dst.Tables("ICTIADJ3").DefaultView
                Dim ADJ_LNO As Integer = 0
                If grdICTIADJ2.ActiveRow IsNot Nothing Then
                    ADJ_LNO = Val(grdICTIADJ2.ActiveRow.Cells("ADJ_LNO").Text)
                End If
                dvw.RowFilter = "ADJ_LNO = " & CStr(ADJ_LNO)
            ElseIf optGL.Value = "S" Then
                Dim tbl As DataTable = dst.Tables("ICTIADJ3").Clone
                Dim ADJ_GNO As Integer = 0
                For Each rowA234 As DataRow In ASCDATA1.SelectDistinct _
                ("ICTIADJ3", New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_DESC"}).Rows
                    Dim DIST_AMT As Decimal = dst.Tables("ICTIADJ3").Compute _
                    ("SUM(DIST_AMT)",
                     "ACCT_CODE = '" & rowA234.Item("ACCT_CODE") & "' and SEG2_CODE = '" & rowA234.Item("SEG2_CODE") & "' and SEG3_CODE = '" & rowA234.Item("SEG3_CODE") & "' and SEG4_CODE = '" & rowA234.Item("SEG4_CODE") & "'")
                    Dim row As DataRow = tbl.NewRow
                    row.Item("ADJ_NO") = Absx1.txtFor("ADJ_NO").Text
                    row.Item("ADJ_LNO") = 0
                    ADJ_GNO += 1
                    row.Item("ADJ_GNO") = ADJ_GNO
                    row.Item("ACCT_CODE") = rowA234.Item("ACCT_CODE")
                    row.Item("SEG2_CODE") = rowA234.Item("SEG2_CODE")
                    row.Item("SEG3_CODE") = rowA234.Item("SEG3_CODE")
                    row.Item("SEG4_CODE") = rowA234.Item("SEG4_CODE")
                    row.Item("ACCT_DESC") = rowA234.Item("ACCT_DESC")
                    row.Item("DIST_AMT") = DIST_AMT
                    tbl.Rows.Add(row)
                Next

                grdICTIADJ3.DataSource = tbl
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("ICTIADJX", YP)
        Sort_grdColumns(grdICTIADJX, "ADJ_NO".ToLower)

        grdICTIADJX.Text = "Entered in " & cbeYP.Text
        If chkGL.Checked Then
            Fill_Records("ICTIADJG", YP)
            grdICTIADJG.Text = "Entered in " & cbeYP.Text
        End If
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String,
                       ByVal WHSE_CODE As String,
                       ByVal QTY_FIELD As String,
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            Dim COLOR_CODE As String = row.Item("COLOR_CODE")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            Dim WHSE_QTY_ON_HAND As Integer = 0
            If rowICTSTAT2 IsNot Nothing Then
                WHSE_QTY_ON_HAND = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
            If WHSE_QTY_ON_HAND + S * QTY < 0 Then
                msg &= vbCr & Format("Style/Color " & STYLE_CODE & "/" & COLOR_CODE & " has only " & CStr(WHSE_QTY_ON_HAND) & " On Hand")
            End If
        Next

        Return msg
    End Function

    Private Sub chkGL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGL.CheckedChanged
        Setup_tab0_GL()
    End Sub

    Sub Setup_tab0_GL()
        If Not chkGL.Checked Then
            tab0.Tabs(0).Selected = True
        Else
            Refresh_Documents()
        End If
        tab0.Tabs("GL").Visible = chkGL.Checked

        If chkGL.Checked Then
            tab0.Tabs("GL").Selected = True
        End If
    End Sub

    Sub Set_Up_Reversal()

        Dim REVERSED_BY_ADJ_NO As String = ""
        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            REVERSED_BY_ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            REVERSED_BY_ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        Dim rowICTIADJ1_orig As DataRow = dst.Tables("ICTIADJ1").NewRow
        rowICTIADJ1_orig.ItemArray = rowICTIADJ1.ItemArray

        rowICTIADJ1 = dst.Tables("ICTIADJ1").Rows(0)
        rowICTIADJ1.Item("REVERSED_BY_ADJ_NO") = REVERSED_BY_ADJ_NO
        rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTIADJ1")

        rowICTIADJ1.ItemArray = rowICTIADJ1_orig.ItemArray
        rowICTIADJ1.AcceptChanges()
        rowICTIADJ1.SetAdded()

        With rowICTIADJ1
            .Item("REVERSES_ADJ_NO") = .Item("ADJ_NO")
            .Item("ADJ_NO") = REVERSED_BY_ADJ_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("ADJ_DATE") = DATETIME_STAMP.Date
            .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
        End With

        'Set new RTRN_NO and reverse all quantities for this return.
        For Each row As DataRow In dst.Tables("ICTIADJ2").Rows
            row.Item("ADJ_NO") = REVERSED_BY_ADJ_NO
            If row.Item("ADJ_QTY") IsNot DBNull.Value Then
                row.Item("ADJ_QTY") *= -1
            End If
            row.Item("OPS_YYYYPP") = ASCMAIN1.CYP

            row.AcceptChanges()
            row.SetAdded()
        Next
    End Sub

    Private Sub grdICTIADJ2_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTIADJ2.InitializeLayout

    End Sub

    Private Sub btnDamages_Click(sender As Object, e As EventArgs) Handles btnDamages.Click
        Dim rowICTIADJ2 As DataRow

        'remove rownum with qty > 0
        ASCMAIN1.sql = "select WHTLOCB1.*,ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC from WHTLOCB1,ICTSTYL1,ICTCOLR1 where WHSE_CODE = '" & rowICTWHSE1.Item("WHSE_CODE") & "' and LOCATION_CODE = '" & rowICTWHSE1.Item("WHSE_LOC_DST") & "' and ICTCOLR1.COLOR_CODE = WHTLOCB1.COLOR_CODE and ICTSTYL1.STYLE_CODE = WHTLOCB1.STYLE_CODE  and LOCATION_QTY > 0"


        'ASCMAIN1.sql = "select ICTSTAT2.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTAT2.COLOR_CODE, ICTCOLR1.COLOR_DESC , NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) LOCATION_QTY, '00-ADJ-A' LOCATION_CODE, ICTCOSTA.STYLE_COST from ICTSTAT2,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP = '202101' and ICTCOSTA.STYLE_CODE = ICTSTAT2.STYLE_CODE and ICTCOSTA.COLOR_CODE = ICTSTAT2.COLOR_CODE and WHSE_CODE = 'NJE' and ICTCOLR1.COLOR_CODE = ICTSTAT2.COLOR_CODE and ICTSTYL1.STYLE_CODE = ICTSTAT2.STYLE_CODE  and NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0)<>0"
        'ASCMAIN1.sql = "select DONADJ.STYLE_CODE, ICTSTYL1.STYLE_DESC, DONADJ.COLOR_CODE, ICTCOLR1.COLOR_DESC, -1 * QTY_ADJ LOCATION_QTY, LOCATION_CODE, ICTCOSTA.STYLE_COST, SHEET from DONADJ,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP (+) = '202101' and ICTCOSTA.STYLE_CODE (+) = DONADJ.STYLE_CODE and ICTCOSTA.COLOR_CODE (+) = DONADJ.COLOR_CODE and ICTCOLR1.COLOR_CODE = DONADJ.COLOR_CODE and ICTSTYL1.STYLE_CODE = DONADJ.STYLE_CODE  and NVL(DONADJ.QTY_ADJ,0)<>0"
        'ASCMAIN1.sql = "select DONADJ.STYLE_CODE, ICTSTYL1.STYLE_DESC, DONADJ.COLOR_CODE, ICTCOLR1.COLOR_DESC, -1 * ADJ LOCATION_QTY, '00-ADJ-A' LOCATION_CODE, ICTCOSTA.STYLE_COST, 5 SHEET from DONADJ5 DONADJ,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP (+) = '202101' and ICTCOSTA.STYLE_CODE (+) = DONADJ.STYLE_CODE and ICTCOSTA.COLOR_CODE (+) = DONADJ.COLOR_CODE and ICTCOLR1.COLOR_CODE = DONADJ.COLOR_CODE and ICTSTYL1.STYLE_CODE = DONADJ.STYLE_CODE  and NVL(DONADJ.ADJ,0)<>0"
        'ASCMAIN1.sql = "select DONADJ.STYLE_CODE, ICTSTYL1.STYLE_DESC, DONADJ.COLOR_CODE, ICTCOLR1.COLOR_DESC, -1 * ADJ LOCATION_QTY, '00-ADJ-A' LOCATION_CODE, ICTCOSTA.STYLE_COST, 5 SHEET from DONADJ6 DONADJ,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP (+) = '202101' and ICTCOSTA.STYLE_CODE (+) = DONADJ.STYLE_CODE and ICTCOSTA.COLOR_CODE (+) = DONADJ.COLOR_CODE and ICTCOLR1.COLOR_CODE = DONADJ.COLOR_CODE and ICTSTYL1.STYLE_CODE = DONADJ.STYLE_CODE  and NVL(DONADJ.ADJ,0)<>0"
        'ASCMAIN1.sql = "select D.STYLE_CODE, ICTSTYL1.STYLE_DESC, D.COLOR COLOR_CODE, ICTCOLR1.COLOR_DESC, -1 * ADJ LOCATION_QTY, '00-ADJ-A' LOCATION_CODE, ICTCOSTA.STYLE_COST, 'KOHLS REM' SHEET from DONREM D,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP (+) = '202101' and ICTCOSTA.STYLE_CODE (+) = D.STYLE_CODE and ICTCOSTA.COLOR_CODE (+) = D.COLOR and ICTCOLR1.COLOR_CODE = D.COLOR and ICTSTYL1.STYLE_CODE = D.STYLE_CODE  and NVL(D.ADJ,0)<>0"
        'ASCMAIN1.sql = "select D.STYLE_CODE, ICTSTYL1.STYLE_DESC, D.COLOR COLOR_CODE, ICTCOLR1.COLOR_DESC, -1 * ADJ LOCATION_QTY, '00-ADJ-A' LOCATION_CODE, ICTCOSTA.STYLE_COST, 'KOHLS DISC' SHEET from DONDISC D,ICTSTYL1,ICTCOLR1,ICTCOSTA where ICTCOSTA.OPS_YYYYPP (+) = '202101' and ICTCOSTA.STYLE_CODE (+) = D.STYLE_CODE and ICTCOSTA.COLOR_CODE (+) = D.COLOR and ICTCOLR1.COLOR_CODE = D.COLOR and ICTSTYL1.STYLE_CODE = D.STYLE_CODE  and NVL(D.ADJ,0)<>0"
        'ASCMAIN1.sql = "select whtlocb2.*,ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC from whtlocb2,ICTSTYL1,ICTCOLR1 " & vbCrLf _
        '            & " where whse_tran_type = 'A'  and ICTCOLR1.COLOR_CODE = WHTLOCB2.COLOR_CODE and ICTSTYL1.STYLE_CODE = WHTLOCB2.STYLE_CODE" & vbCrLf _
        '            & " and bar_code in (" & vbCrLf _
        '            & " select distinct bar_code from whtlocb1" & vbCrLf _
        '            & " where bar_code in (" & vbCrLf _
        '            & " select barcode" & vbCrLf _
        '            & " from potlpnl1" & vbCrLf _
        '            & " where po_shipment_no in ('019481','019458','019464','019469')" & vbCrLf _
        '            & " and rec_status =1)" & vbCrLf _
        '            & " group by bar_code" & vbCrLf _
        '            & " having sum(abs(location_qty)) = 0)"
        ' *** Below SQL used for Costco Mergers in to single color
        'ASCMAIN1.sql = "select WHTLOCB1.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC
        '                    From WHTLOCB1, ICTSTYL1, ICTCOLR1
        '                    Where WHTLOCB1.WHSE_CODE = 'NJC'
        '                    And ICTCOLR1.COLOR_CODE = WHTLOCB1.COLOR_CODE
        '                    And ICTSTYL1.STYLE_CODE = WHTLOCB1.STYLE_CODE
        '                    And WHTLOCB1.STYLE_CODE = 'CS36930'
        '                    And WHTLOCB1.COLOR_CODE = '401' 
        '                    And location_qty <> 0"


        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")

            cdr = LookUp("ICTSTYL1", row("STYLE_CODE"))
            Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
            Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
            Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")

            Dim WhseQty = ASCDATA1.GetDataValue("Select WHSE_QTY_ON_HAND from ICTSTAT2 WHERE WHSE_CODE = :PARM1 AND STYLE_CODE = :PARM2 AND COLOR_CODE = :PARM3", "VVV", New Object() {rowICTWHSE1.Item("WHSE_CODE"), row("STYLE_CODE"), row("COLOR_CODE")})
            If WhseQty > 0 Then

                rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
                With rowICTIADJ2
                    .Item("ADJ_NO") = Absx1.CtlFor("ADJ_NO").Text
                    .Item("ADJ_LNO") = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                    .Item("STYLE_CODE") = row("STYLE_CODE")
                    .Item("STYLE_DESC") = row("STYLE_DESC")
                    .Item("COLOR_CODE") = row("COLOR_CODE")
                    .Item("COLOR_DESC") = row("COLOR_DESC")
                    .Item("ADJ_QTY") = Val(row("LOCATION_QTY") & "") * -1  'WHSE_TRAN_QTY
                    .Item("STYLE_COST") = STYLE_COST
                    '.Item("STYLE_COST") = Val(row("STYLE_COST") & "") ' TEMP
                    .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                    .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("LOCATION_CODE") = row("LOCATION_CODE")
                    .Item("BAR_CODE") = row("BAR_CODE")
                    .Item("ADJ_REF") = ""
                    ' .Item("ADJ_REF") = row("SHEET") ' TEMP
                End With
                dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

                'rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
                'With rowICTIADJ2
                '    .Item("ADJ_NO") = Absx1.CtlFor("ADJ_NO").Text
                '    .Item("ADJ_LNO") = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                '    .Item("STYLE_CODE") = row("STYLE_CODE")
                '    .Item("STYLE_DESC") = row("STYLE_DESC")
                '    .Item("COLOR_CODE") = "040"
                '    .Item("COLOR_DESC") = "SILVER"
                '    .Item("ADJ_QTY") = Val(row("LOCATION_QTY") & "") '* -1  'WHSE_TRAN_QTY
                '    .Item("STYLE_COST") = STYLE_COST
                '    '.Item("STYLE_COST") = Val(row("STYLE_COST") & "") ' TEMP
                '    .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                '    .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                '    .Item("LOCATION_CODE") = row("LOCATION_CODE")
                '    .Item("BAR_CODE") = row("BAR_CODE")
                '    .Item("ADJ_REF") = ""
                '    ' .Item("ADJ_REF") = row("SHEET") ' TEMP
                'End With
                'dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
            End If

        Next
        Update_WHSE_QTY_OH()

    End Sub

    Private Sub btnShortage_Click(sender As Object, e As EventArgs) Handles btnShortage.Click
        ' - O status is for shortages in Pick, Negative Adj
        Shortage_N_Found("O")
    End Sub

    Private Sub btnFoundOH_Click(sender As Object, e As EventArgs) Handles btnFoundOH.Click
        ' - A status is for Found Items by Victor, Positive Adj
        Shortage_N_Found("A")
    End Sub
    Private Sub Shortage_N_Found(ByRef STATUS As String)
        Dim rowICTIADJ2 As DataRow
        Dim S As Integer

        If STATUS = "O" Then
            S = -1
        Else
            S = 1
        End If


        BeginTrans()
        ASCMAIN1.sql = $"UPDATE WHTPICKS SET STATUS = 'W', LAST_OPER = '{ASCMAIN1.USER_ID}', LAST_DATE = sysdate WHERE STATUS = '{STATUS}'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = $"SELECT WHTPICKS.PICK_NO ,WHTPICKS.STYLE_CODE ,WHTPICKS.COLOR_CODE ,WHTPICKS.STYLE_DESC ,
                        WHTPICKS.COLOR_DESC ,NVL(WHTPICKS.LOCATION_CODE, WHTLOCM1.LOCATION_CODE)  LOCATION_CODE,
                        WHTPICKS.SHORTAGE ,WHTPICKS.STATUS ,WHTPICKS.INIT_OPER ,WHTPICKS.INIT_DATE ,
                        WHTPICKS.LAST_OPER ,WHTPICKS.LAST_DATE, WHSE_QTY
                        FROM WHTPICKS, (SELECT distinct STYLE_CODE, COLOR_CODE, 
                        SUM(LOCATION_QTY) over(partition by STYLE_CODE, COLOR_CODE) WHSE_QTY,
                        first_value(WHTLOCB1.LOCATION_CODE) over(partition by STYLE_CODE, COLOR_CODE order by WHTLOCB1.LOCATION_QTY desc) LOCATION_CODE
                         FROM WHTLOCB1, WHTLOCM1
                         Where WHTLOCB1.WHSE_CODE =  WHTLOCM1.WHSE_CODE
                         And WHTLOCB1.LOCATION_CODE =  WHTLOCM1.LOCATION_CODE
                         And WHTLOCM1.LOCATION_USE in ('A','E','R')
                         And WHTLOCM1.WHSE_CODE = '{Absx1.txtFor("WHSE_CODE").Text}' ) WHTLOCM1
                         WHERE WHTPICKS.STYLE_CODE = WHTLOCM1.STYLE_CODE
                         And WHTPICKS.COLOR_CODE = WHTLOCM1.COLOR_CODE
                         And WHTPICKS.STATUS = 'W'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")

            cdr = LookUp("ICTSTYL1", row("STYLE_CODE"))
            Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
            Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
            Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")
            Dim ADJ_REF As String = ""

            Dim pickrow As DataRow = LookUp("SOTPICK1", row("PICK_NO") & "")
            If pickrow Is Nothing Then
                ADJ_REF = "Whse Found"
            Else
                ADJ_REF = pickrow("SHIP_BOL_NO") & ""
            End If

            rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
            With rowICTIADJ2
                .Item("ADJ_NO") = Absx1.CtlFor("ADJ_NO").Text
                .Item("ADJ_LNO") = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                .Item("STYLE_CODE") = row("STYLE_CODE")
                .Item("STYLE_DESC") = row("STYLE_DESC")
                .Item("COLOR_CODE") = row("COLOR_CODE")
                .Item("COLOR_DESC") = row("COLOR_DESC")
                .Item("ADJ_QTY") = Val(row("SHORTAGE") & "") * S
                .Item("STYLE_COST") = STYLE_COST
                .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("LOCATION_CODE") = row("LOCATION_CODE")
                .Item("ADJ_REF") = ADJ_REF
            End With
            dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)

        Next

        ASCMAIN1.sql = $"UPDATE WHTPICKS SET STATUS = 'P' WHERE STATUS = 'W'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Update_WHSE_QTY_OH()

        CommitTrans()

    End Sub
    Sub Update_WHSE_QTY_OH()

        For Each grow As UltraWinGrid.UltraGridRow In grdICTIADJ2.Rows

            ASCMAIN1.sql = "Select * from ICTSTAT2 where STYLE_CODE = '" & grow.Cells("STYLE_CODE").Value & "' and COLOR_CODE = '" & grow.Cells("COLOR_CODE").Value & "' and WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
            Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow
            If rowICTSTAT2 IsNot Nothing Then
                grow.Cells("WHSE_QTY_ON_HAND").Value = Val(rowICTSTAT2.Item("WHSE_QTY_ON_HAND") & "")
            End If
        Next

    End Sub
    Sub Import_Amazon_File()

        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            'Dim filter As String = "xlsb files (*.xlsb)|*.xlsx|All files (*.*)|*.*"
            Dim filter As String = "All files (*.*)|*.*"
            openFileDialog1.Filter = filter
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using
        'Try
        If FILENAME <> "" Then
            Try
                ASCMAIN1.Progress("Now Building Adjustment From Excel", "")
                Me.Cursor = Cursors.WaitCursor

                DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
                Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(FILENAME)
                Dim xws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                xws = XWB.Worksheets(1)

                ' Dim xws As Microsoft.Office.Interop.Excel.Worksheet = XWB.Worksheets(1)
                Dim ERROR_CODEs As List(Of String) = New List(Of String)
                Dim ADJ_LNO As Integer = 1
                Dim ADJ_NO As String = ASCMAIN1.Next_Control_No("TRAN_NO_A")
                Dim WHSE_CODE As String = "AMAZ02"
                Dim CUST_STORE_NO As String = "AMAFBA"
                Dim CUST_CODE As String = "AMAZONFBA"
                Dim ADJ_NOTE As String = "AMAZONFBA"
                Dim QTY_COL As Int32 = 0

                If optAuto.Value = "R" Then
                    ADJ_NOTE = "Returned " & XWB.Worksheets(1).NAME & " " & ADJ_NOTE
                    QTY_COL = 10
                ElseIf optAuto.Value = "F" Then
                    ADJ_NOTE = "Found " & XWB.Worksheets(1).NAME & " " & ADJ_NOTE
                    QTY_COL = 11
                ElseIf optAuto.Value = "X" Then
                    ADJ_NOTE = "Removed " & XWB.Worksheets(1).NAME & " " & ADJ_NOTE
                    QTY_COL = 13
                ElseIf optAuto.Value = "L" Then
                    ADJ_NOTE = "Lost " & XWB.Worksheets(1).NAME & " " & ADJ_NOTE
                    QTY_COL = 14
                ElseIf optAuto.Value = "D" Then
                    ADJ_NOTE = "Disposed " & XWB.Worksheets(1).NAME & " " & ADJ_NOTE
                    QTY_COL = 15
                End If

                For i As Integer = 4 To xws.UsedRange.Rows.Count Step +1
                    Dim SKU As String = xws.Cells(i, 1).value.ToString
                    Dim STYLECOLOR As String() = Split(SKU, "-")
                    Dim STYLE_CODE As String = ""
                    Dim STYLE_DESC As String = ""
                    Dim COLOR_CODE As String = ""
                    Dim SALES_DIVISION_CODE As String = ""
                    Dim STYLE_COST As Double = 0
                    Dim STYLE_CLASS_CODE As String = ""
                    Dim ORDR_QTY As Int32 = 0
                    ORDR_QTY = Val(xws.Cells(i, QTY_COL).value.ToString)

                    If ORDR_QTY <> 0 And SKU <> "Grand Total" Then
                        Dim rowSOTCSTY1 As DataRow = LookUp("SOTCSTY1", New String() {CUST_CODE, SKU})
                        If rowSOTCSTY1 Is Nothing Then
                            '                    If Val(xws.Cells(i, 4).value.ToString) <> 0 Then
                            ERROR_CODEs.Add("SKU is missing from Style Cross Reference File " & SKU & " On Line No " & i)

                            Dim rowERROR_TBL As DataRow = Nothing
                            rowERROR_TBL = dst.Tables("ERROR_TBL").NewRow
                            With rowERROR_TBL
                                .Item("SKU") = SKU
                                .Item("ERROR_DETAIL") = "Ln# " & i
                            End With
                            dst.Tables("ERROR_TBL").Rows.Add(rowERROR_TBL)
                            '                  End If

                            STYLE_CODE = ""
                            COLOR_CODE = ""
                            ORDR_QTY = 0
                        Else
                            STYLE_CODE = rowSOTCSTY1.Item("STYLE_CODE")
                            COLOR_CODE = rowSOTCSTY1.Item("COLOR_CODE")
                            ' MANIPULATE 

                        End If


                        If STYLE_CODE <> "" And ORDR_QTY <> 0 Then
                            Dim rowICTSTYL1 As DataRow = clsASCBASE1.LookUp("ICTSTYL1", STYLE_CODE)
                            If rowICTSTYL1 Is Nothing Then
                                ERROR_CODEs.Add("Invalid Style Code " & STYLE_CODE & " for " & SKU & " On Line No " & i)
                                STYLE_CODE = ""
                            Else
                                STYLE_DESC = rowICTSTYL1.Item("STYLE_DESC")
                                SALES_DIVISION_CODE = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                                SALES_DIVISION_CODE = rowICTSTYL1.Item("SALES_DIVISION_CODE")
                                STYLE_COST = rowICTSTYL1.Item("STYLE_COST")
                                STYLE_CLASS_CODE = rowICTSTYL1.Item("STYLE_CLASS_CODE")

                            End If
                        End If

                        'add ICTIADJ2
                        If STYLE_CODE <> "" And ORDR_QTY <> 0 Then
                            Dim rowICTIADJ2 As DataRow = Nothing
                            rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
                            With rowICTIADJ2
                                .Item("ADJ_NO") = ADJ_NO
                                .Item("ADJ_LNO") = ADJ_LNO
                                .Item("STYLE_CODE") = STYLE_CODE
                                .Item("COLOR_CODE") = COLOR_CODE

                                .Item("STYLE_DESC") = STYLE_DESC

                                .Item("STYLE_COST") = STYLE_COST
                                .Item("ADJ_QTY") = ORDR_QTY

                                .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                                .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                                .Item("OPS_YYYYPP") = ""
                                .Item("LOCATION_CODE") = ""

                                .Item("BAR_CODE") = ""
                                .Item("ADJ_REF") = ""

                            End With
                            dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
                            ADJ_LNO = ADJ_LNO + 1

                        End If
                    End If
                Next

                Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTIADJ2").Compute("SUM(LINE_COSTS)", "") & "")


                Dim rowICTIADJ1 As DataRow = Nothing
                rowICTIADJ1 = dst.Tables("ICTIADJ1").NewRow
                With rowICTIADJ1
                    .Item("ADJ_NO") = ADJ_NO
                    .Item("ADJ_DATE") = DATETIME_STAMP.Date
                    .Item("WHSE_CODE") = "AMAZ02"
                    .Item("REASON_CODE") = "STK"
                    .Item("ADJ_NOTE") = ADJ_NOTE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("REGISTER_IND") = ""
                    .Item("REGISTER_XNO") = ""
                    .Item("ADJ_SOURCE") = "A"
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("TOTAL_COSTS") = TOTAL_COSTS
                    .Item("RTRN_NO") = ""
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("REVERSED_BY_ADJ_NO") = ""
                    .Item("REVERSES_ADJ_NO") = ""
                    .Item("ADJ_REF") = ""
                    .Item("JOURNAL_IND") = ""
                    .Item("JOURNAL_XNO") = ""

                End With
                dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)

                If ERROR_CODEs.Count <> 0 Then

                    If dst.Tables("ERROR_TBL").Rows.Count <> 0 Then
                        Using F As New ASFMSGBF
                            F.Show_grd(dst.Tables("ERROR_TBL"), Me, "The following Import Errors have been identified", "DGJ")
                        End Using
                    End If
                    dst.Tables("ICTIADJ1").Rows.Clear()
                    dst.Tables("ICTIADJ2").Rows.Clear()
                    dst.Tables("ERROR_TBL").Rows.Clear()
                Else
                    If dst.Tables("ICTIADJ2").Rows.Count = 0 Then
                        MsgBox("There are no Qtys to Update for this type of import", MsgBoxStyle.OkOnly, "Nothing to  Update from Spreadsheet")

                    Else
                        MsgBox("This Excel File has been successfully Updated to Create an Adjustment",
                          MsgBoxStyle.OkOnly, "Verification")
                        EntryMode = "N"
                        Mode_Settings(True)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message, "Amazon Adj Import, Excel Format Issues", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Import_Amazon_File()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As Object, e As SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        Setup_tab0()
    End Sub
    Sub Setup_tab0()
        If SELECTION_NO = 0 Then Exit Sub

        If tab0.Tabs(2).Selected = True Then
            '.Groups("GL Distribution").Visible = ScreenMode And (EntryMode = "V") And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
            '.Groups("Show if Entered in").Visible = Not ScreenMode ' And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
            '.Groups("Totals").Visible = False ' ScreenMode
            '.Groups("Events").Visible = ScreenMode And (EntryMode <> "N")
            '.Groups("Damages").Visible = ScreenMode And EntryMode = "N" And (ASCMAIN1.Running_in_VS Or (ASCMAIN1.CLIENT = "RGI" And ASCMAIN1.USER_SECURITY_CODEs.Contains("WS")))
            '.Groups("Special Functions").Visible = (ASCMAIN1.CLIENT = "VAN" And (ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "wendy")) And Not ScreenMode And Not InquiryMode
            UltraExplorerBar1.Groups("Cycle Count Adjustment").Visible = True
            UltraExplorerBar1.Groups("Screen Control").Visible = False
            UltraExplorerBar1.Groups("Special Functions").Visible = False
            UltraExplorerBar1.Groups("GL Distribution").Visible = False
            UltraExplorerBar1.Groups("Show if Entered in").Visible = False
            chkGL.Checked = False

            ' UltraExplorerBar1.Groups("Screen Control").Visible = False


            Set_Read_Only(UltraGroupBox1, True)
        Else
            UltraExplorerBar1.Groups("Cycle Count Adjustment").Visible = False
            UltraExplorerBar1.Groups("Screen Control").Visible = True
            UltraExplorerBar1.Groups("Special Functions").Visible = True
            UltraExplorerBar1.Groups("GL Distribution").Visible = True
            UltraExplorerBar1.Groups("Show if Entered in").Visible = True


            Set_Read_Only(UltraGroupBox1, False)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show("Are you sure you want to Refresh based on Date, You will lose any Cycles you have Selected?", "Confirm Refresh",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If

        Dim CUTOFF As Date = Nothing
        CUTOFF = dteDATE_CUTOFF.Value
        Fill_Records("WHTCYCL1", CUTOFF)
        OptResolution.Value = "p"
        OptResolution.Value = "U"
        CYCLE_TYPE = "V"
        CYCLE_STATUS = "D"
        chkUpdated.Checked = False


    End Sub

    Private Sub UltraTextEditor3_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor3.ValueChanged

    End Sub

    Private Sub grdWHTCYCL1_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdWHTCYCL1.ClickCell
        'Dim WHSE_TRAN_NO As String = grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""

        'Fill_Records("WHTLOCB2", WHSE_TRAN_NO)


        'grdWHTLOCB2.Text = "LOCB2 Records for Whse Transaction No" & " " & grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value


        'Dim dvw As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        'dvw.RowFilter = "WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"
    End Sub

    Private Sub cmdUpdateCycles_Click(sender As Object, e As EventArgs) Handles cmdUpdateCycles.Click

        If DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'").Length = 0 Then
            EMsg = "You Must Select a Cycle to Update"
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        If MessageBox.Show("Are you sure you want to Update Cycles Selected?", "Confirm Update",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
            Exit Sub
        End If
        ' update
        Me.Cursor = Cursors.WaitCursor
        Call UPDATE_CYCLE_ADJUSTMENTS()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub optType_ValueChanged(sender As Object, e As EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()
    End Sub
    Sub Load_WHTCYCL1()
        If ASCMAIN1.CLIENT <> "VAN" Then
            Exit Sub
        End If

        Dim SQLW As String = "CYCLE_NO = CYCLE_NO"

        If CYCLE_TYPE = "V" And OptResolution.Value = "U" And CYCLE_STATUS = "D" And chkUpdated.Checked = False Then
            cmdUpdateCycles.Enabled = True
        Else
            cmdUpdateCycles.Enabled = False
        End If

        If OptResolution.Value = "A" Then
        Else
            SQLW = SQLW & " and CYCLE_RESOLUTION = '" & OptResolution.Value & "'"
        End If

        SQLW = SQLW & " and CYCLE_TYPE = '" & CYCLE_TYPE & "'"

        SQLW = SQLW & " and CYCLE_STATUS = '" & CYCLE_STATUS & "'"

        If chkUpdated.Checked = True Then
            SQLW = SQLW & " and UPDATED_INV_ADJ = '1'"
        Else
            SQLW = SQLW & " and UPDATED_INV_ADJ IS NULL"
        End If

        If chkSel.Checked = True Then
            SQLW = SQLW & " and SEL = '1'"
        End If

        Dim dvw As DataView
        dvw = DirectCast(grdWHTCYCL1.DataSource, DataTable).DefaultView
        dvw.RowFilter = SQLW


        If chkUpdated.Checked = True Then
            grdWHTCYCL1.Text = "Cycles That have been Updates"
        Else
            grdWHTCYCL1.Text = "Cycles That have been npot been Updated"
        End If


    End Sub

    Private Sub OptResolution_ValueChanged(sender As Object, e As EventArgs) Handles OptResolution.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()

    End Sub

    Private Sub OptStatus_ValueChanged(sender As Object, e As EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()
    End Sub

    Private Sub chkUpdated_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkUpdated.CheckedValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()
    End Sub
    Sub UPDATE_CYCLE_ADJUSTMENTS()
        Dim WHSE_CODE As String = ""



        dst.Tables("ICTIADJ1").Rows.Clear()
        dst.Tables("ICTIADJ2").Rows.Clear()
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD


        ' PRE UDPATE LOCK CYCLES, CAN'T LOCK, GET OUT
        For Each ROW As DataRow In DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'")
            If WHSE_CODE & "" = "" Then
                WHSE_CODE = ROW.Item("WHSE_CODE") & ""
            End If
            If WHSE_CODE <> ROW.Item("WHSE_CODE") & "" Then
                EMsg &= vbCr & "Multiple Wareshouses Selected, Cannot Proceed"
                Exit Sub
            End If
            Dim CYCLE_NO = ROW.Item("CYCLE_NO") & ""
            If Not ASCMAIN1.Logical_Lock("WHTCYCL1", CYCLE_NO) Then
                ASCMAIN1.MultiTask_Release(, , 1)
                Exit Sub
            End If
            ' CHECK WHTCYCL1 UPDATED FLAG HERE FOR CYCL1
            ASCMAIN1.sql = "Select * from WHTCYCL1 where CYCLE_NO = '" & CYCLE_NO & "' AND UPDATED_INV_ADJ = '1'"
            Dim tblWHTCYCL1 As DataTable = ASCDATA1.GetDataTable()
            If tblWHTCYCL1.Rows.Count > 0 Then
                EMsg &= vbCr & "Cycle " & CYCLE_NO & " Had Already Been Updated, Cannot Proceed"
                Exit Sub
            End If


        Next
        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")



        Dim WHSE_TRAN_NO As String = ""
        Dim LOCATION_CODE As String = ""
        For Each ROW As DataRow In DirectCast(grdWHTCYCL1.DataSource, DataTable).Select("SEL='1'")
            WHSE_TRAN_NO = ROW.Item("WHSE_TRAN_NO") & ""
            LOCATION_CODE = "00-LNF-A"
            Fill_Records("WHTLOCB2", WHSE_TRAN_NO)


            Dim ADJ_NO As String = ASCMAIN1.Next_Control_No("TRAN_NO_A")

            rowICTIADJ1 = dst.Tables("ICTIADJ1").NewRow
            rowICTIADJ1.Item("ADJ_NO") = ADJ_NO
            rowICTIADJ1.Item("WHSE_CODE") = WHSE_CODE
            rowICTIADJ1.Item("ADJ_DATE") = DATETIME_STAMP.Date
            rowICTIADJ1.Item("ADJ_SOURCE") = "E"
            rowICTIADJ1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTIADJ1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTIADJ1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTIADJ1.Item("REGISTER_IND") = "0"
            rowICTIADJ1.Item("JOURNAL_IND") = "0"
            rowICTIADJ1.Item("REASON_CODE") = "WHLOC"

            dst.Tables("ICTIADJ1").Rows.Add(rowICTIADJ1)


            Dim rowICTIADJ2 As DataRow
            For Each rowWHTLOCB2 As DataRow In dst.Tables("WHTLOCB2").Select("WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "' and LOCATION_CODE = '" & LOCATION_CODE & "'")
                cdr = LookUp("ICTSTYL1", rowWHTLOCB2("STYLE_CODE"))
                Dim STYLE_DESC As String = cdr.Item("STYLE_DESC") & ""
                Dim STYLE_CLASS_CODE As String = cdr.Item("STYLE_CLASS_CODE") & ""
                Dim SALES_DIVISION_CODE As String = cdr.Item("SALES_DIVISION_CODE") & ""
                Dim STYLE_COST As Decimal = Val(cdr.Item("STYLE_COST") & "")

                cdr = LookUp("ICTCOLR1", rowWHTLOCB2("COLOR_CODE"))
                Dim COLOR_DESC As String = cdr.Item("COLOR_DESC") & ""


                rowICTIADJ2 = dst.Tables("ICTIADJ2").NewRow
                With rowICTIADJ2
                    .Item("ADJ_NO") = ADJ_NO
                    .Item("ADJ_LNO") = Val(dst.Tables("ICTIADJ2").Compute("Max(ADJ_LNO)", "") & "") + 1
                    .Item("STYLE_CODE") = rowWHTLOCB2("STYLE_CODE")
                    .Item("STYLE_DESC") = STYLE_DESC
                    .Item("COLOR_CODE") = rowWHTLOCB2("COLOR_CODE")
                    .Item("COLOR_DESC") = COLOR_DESC
                    .Item("ADJ_QTY") = Val(rowWHTLOCB2("WHSE_TRAN_QTY") & "") * -1
                    .Item("STYLE_COST") = STYLE_COST
                    ''    '.Item("STYLE_COST") = Val(row("STYLE_COST") & "") ' TEMP
                    .Item("STYLE_CLASS_CODE") = STYLE_CLASS_CODE
                    .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("LOCATION_CODE") = rowWHTLOCB2("LOCATION_CODE")
                    .Item("BAR_CODE") = rowWHTLOCB2("BAR_CODE")
                    .Item("ADJ_REF") = ""
                    .Item("ADJ_REF") = rowWHTLOCB2("WHSE_TRAN_NO")
                End With
                dst.Tables("ICTIADJ2").Rows.Add(rowICTIADJ2)
            Next
            ROW.Item("UPDATED_INV_ADJ") = "1"
            ROW.Item("ADJ_NO") = ADJ_NO
        Next

        ' NEW Update
        BeginTrans()

        Update_Record_TDA("WHTCYCL1")

        ICCMAIN1.Update_Adjustment(Me)

        If location_support Then
            For Each rowICTIADJ1 As DataRow In dst.Tables("ICTIADJ1").Select("")
                ASCDATA1.ExecuteSP("WHPLOCB2", "VVV",
                     New Object() {"A", rowICTIADJ1.Item("ADJ_NO"), ASCMAIN1.SESSION_NO},
                     New String() {"WHSE_TRAN_TYPE_in", "WHSE_TRAN_NO_in", "SESSION_NO_in"})

            Next
        End If
        CommitTrans("Update Complete")

        ASCMAIN1.MultiTask_Release(, , 1)

        Mode_Settings(False)

    End Sub

    Private Sub chkSel_CheckedValueChanged(sender As Object, e As EventArgs) Handles chkSel.CheckedValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_WHTCYCL1()

    End Sub

    Private Sub grdWHTCYCL1_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdWHTCYCL1.InitializeLayout

    End Sub

    Private Sub grdWHTCYCL1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTCYCL1.AfterRowActivate
        Dim WHSE_TRAN_NO As String = grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value & ""

        Fill_Records("WHTLOCB2", WHSE_TRAN_NO)


        grdWHTLOCB2.Text = "LOCB2 Records for Whse Transaction No" & " " & grdWHTCYCL1.ActiveRow.Cells("WHSE_TRAN_NO").Value


        Dim dvw As DataView = DirectCast(grdWHTLOCB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "WHSE_TRAN_NO = '" & WHSE_TRAN_NO & "'"
    End Sub
End Class