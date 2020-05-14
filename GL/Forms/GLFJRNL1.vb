Imports System.Math

Public Class GLFJRNL1

    Dim recurring As Boolean = False
    Dim rowGLTJRNL1 As DataRow
    Dim AUTO_REPORT As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "GLFJRNLI" Then
            InquiryMode = True
        End If

        With dst
            Create_TDA(.Tables.Add, "GLTJRNL1", "*")

            ASCMAIN1.sql = "Select * from GLTDETL1 where JOURNAL_NO = :PARM1"
            Create_TDA(.Tables.Add, "GLTDETL1", "**", 0, True, "V", 3)

            'ASCMAIN1.sql = "Select * from GLTJRNL1 where OPS_YYYYPP = :PARM1"
            'THERE MUST BE SOME REASON WHY WE Did NOT WANT TO SHOW JOURNALS IN THE REVERSING PERIODS
            ASCMAIN1.sql = "Select * from GLTJRNL1 where JOURNAL_NO in (Select Distinct JOURNAL_NO from GLTDETL1 where OPS_YYYYPP = :PARM1)"
            Create_TDA(.Tables.Add, "GLTJRNLP", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO" _
                & ", GLTDETL1.ACCT_CODE, GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" _
                & ", GLTDETL1.DETL_POSTING_AMT" _
                & ", CASE WHEN GLTDETL1.DETL_POSTING_AMT > 0 THEN GLTDETL1.DETL_POSTING_AMT ELSE NULL END AMOUNT_DR" _
                & ", CASE WHEN GLTDETL1.DETL_POSTING_AMT < 0 THEN  -1 * GLTDETL1.DETL_POSTING_AMT ELSE NULL END AMOUNT_CR" _
                & ", GLTDETL1.DETL_DESC, GLTACCT1.ACCT_DESC " _
                & " from GLTDETL1,GLTACCT1,GLTJRNL1 " _
                & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO " _
                & "   and GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE " _
                & "   and GLTDETL1.OPS_YYYYPP = :PARM1"
            '                & "   and GLTJRNL1.OPS_YYYYPP = GLTDETL1.OPS_YYYYPP " _
            ' when we figure out why we should not be showing journals that did not originate in the specified period, we might have to re-enable this
            Create_TDA(.Tables.Add, "GLTJRNLP2", "**", 0, False, "V", 2)


            ASCMAIN1.sql = "Select * from GLTJRNL1 where JOURNAL_TYPE = 'GLRE'"
            Create_TDA(.Tables.Add, "GLTJRNLR", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select GLTJRNL2.*,GLTACCT1.ACCT_DESC" _
                & ", CASE WHEN GLTJRNL2.DETL_POSTING_AMT > 0 THEN GLTJRNL2.DETL_POSTING_AMT ELSE NULL END AMOUNT_DR" _
                & ", CASE WHEN GLTJRNL2.DETL_POSTING_AMT < 0 THEN -1 * GLTJRNL2.DETL_POSTING_AMT ELSE NULL END AMOUNT_CR " _
                & " from GLTJRNL2,GLTJRNL1,GLTACCT1 where GLTACCT1.ACCT_CODE = GLTJRNL2.ACCT_CODE and GLTJRNL1.JOURNAL_NO = GLTJRNL2.JOURNAL_NO and GLTJRNL1.JOURNAL_TYPE = 'GLRE'"
            Create_TDA(.Tables.Add, "GLTJRNLR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select GLTJRNL2.*,GLTACCT1.ACCT_DESC" _
                & ", CASE WHEN GLTJRNL2.DETL_POSTING_AMT > 0 THEN GLTJRNL2.DETL_POSTING_AMT ELSE NULL END AMOUNT_DR" _
                & ", CASE WHEN GLTJRNL2.DETL_POSTING_AMT < 0 THEN -1 * GLTJRNL2.DETL_POSTING_AMT ELSE NULL END AMOUNT_CR " _
                & " from GLTJRNL2,GLTACCT1 where GLTACCT1.ACCT_CODE (+) = GLTJRNL2.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTJRNL2", "**", 1, True)

            .Relations.Add("GLTJRNLP2", _
            New DataColumn() {.Tables("GLTJRNLP").Columns("JOURNAL_NO")}, _
            New DataColumn() {.Tables("GLTJRNLP2").Columns("JOURNAL_NO")})

            .Relations.Add("GLTJRNLR2", _
            New DataColumn() {.Tables("GLTJRNLR").Columns("JOURNAL_NO")}, _
            New DataColumn() {.Tables("GLTJRNLR2").Columns("JOURNAL_NO")})

            ASCMAIN1.sql = "Select Distinct GLTACCT1.* from GLTACCT1,GLTDETL1" _
            & " where GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE and GLTDETL1.JOURNAL_NO = :PARM1"
            Create_TDA(.Tables.Add, "GLTACCT1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select Distinct GLTPARM2.* from GLTPARM2,GLTJRNL1 " _
            & " where (GLTPARM2.OPS_YYYYPP = GLTJRNL1.OPS_YYYYPP " _
            & "     or GLTPARM2.OPS_YYYYPP = NVL(GLTJRNL1.OPS_YYYYPP_REV,'000000')) " _
            & "  and GLTJRNL1.JOURNAL_NO = :PARM1"
            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2," _
            & " (Select Distinct OPS_YYYYPP from GLTDETL1 where JOURNAL_NO = :PARM1) J" _
            & " where J.OPS_YYYYPP = GLTPARM2.OPS_YYYYPP"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)


            .Tables.Add("GLTDETLX")
            With .Tables("GLTDETLX")
                .Columns.Add("DETL_CVX_TYPE", GetType(System.String))
                .Columns.Add("DETL_CVX_NO", GetType(System.String))
                .Columns.Add("DETL_CVX_NAME", GetType(System.String))
                .PrimaryKey = New DataColumn() {.Columns("DETL_CVX_TYPE"), .Columns("DETL_CVX_NO")}
            End With
            'ASCMAIN1.sql = "SELECT GLTDETL1.DETL_CVX_TYPE, GLTDETL1.DETL_CVX_NO " _
            '& ", APTVEND1.VEND_NAME DETL_CVX_NAME" _
            '& " from GLTDETL1,APTVEND1 WHERE ROWNUM < 1"
            '.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLX", 2))

        End With

        Create_Lookup("GLTACCT1")
        Create_Lookup("GLTSEGM1")

        grdGLTJRNL2.DataSource = dst.Tables("GLTJRNL2")
        grdGLTJRNLP.DataSource = dst.Tables("GLTJRNLP")
        grdGLTJRNLR.DataSource = dst.Tables("GLTJRNLR")

        Get_PARM("GLTPARM1")

        Set_SEGS(grdGLTJRNLP, "GLTJRNLP2")
        Set_SEGS(grdGLTJRNLR, "GLTJRNLR2")
        Set_SEGS(grdGLTJRNL2) ' , "GLTJRNL2")

        ' tabGLTJRNL1.Dock = DockStyle.Fill

        Create_Summary(grdGLTJRNL2, "JOURNAL_LNO", "Count")
        Create_Summary(grdGLTJRNL2, "AMOUNT_DR")
        Create_Summary(grdGLTJRNL2, "AMOUNT_CR")
        Create_Summary(grdGLTJRNL2, "DETL_POSTING_AMT")

        Create_Summary(grdGLTJRNLP, "AMOUNT_DR", , "GLTJRNLP2")
        Create_Summary(grdGLTJRNLP, "AMOUNT_CR", , "GLTJRNLP2")
        grdGLTJRNLP.DisplayLayout.Bands("GLTJRNLP2").SummaryFooterCaption = "Journal Totals for [SCROLLTIPFIELD]"
        Create_Summary(grdGLTJRNLR, "AMOUNT_DR", , "GLTJRNLR2")
        Create_Summary(grdGLTJRNLR, "AMOUNT_CR", , "GLTJRNLR2")
        grdGLTJRNLR.DisplayLayout.Bands("GLTJRNLR2").SummaryFooterCaption = "Journal Totals for [SCROLLTIPFIELD]"

        With grdGLTJRNLP.DisplayLayout.Bands(0)
            .Columns("JOURNAL_NO").Header.Fixed = True
            .Columns("JOURNAL_DESC").Header.Fixed = True
        End With

        Dim sql_where As String = ""
        If ASCMAIN1.CLIENT = "RGI" Then
            sql_where = "OPS_YYYYPP >= '" & ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 48) & "'"
        Else
            sql_where = "OPS_YYYYPP >= '" & ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 36) & "'"
        End If
        If ASCMAIN1.Running_in_VS Then
            sql_where = "OPS_YYYYPP >= (Select Min (OPS_YYYYPP) from GLTJRNL1) and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 36) & "'"
        End If
        Load_Drop_Down("OPS_YYYYPP", sql_where)
        Load_Drop_Down("OPS_YYYYPP_REV", sql_where)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"


            Case "New Journal Entry"
                Validate_Code("OPS_YYYYPP")
                grdGLTJRNL2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnBottom
                If Absx1.txtFor("JOURNAL_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Enter a Description"
                End If
                If Absx1.dteFor("JOURNAL_DATE").Text = "" Then
                    EMsg &= vbCr & "You Must Enter a Journal Date"
                End If

            Case "View"
                If Absx1.txtFor("JOURNAL_NO").Text = "" Then
                    EMsg &= vbCr & "No Journal No Specified"
                Else
                    Dim rowGLTJRNL1 As DataRow = LookUp("GLTJRNL1", Absx1.txtFor("JOURNAL_NO").Text)
                    If rowGLTJRNL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Journal No Specified"
                    End If
                End If

            Case "Edit", "Load to Post"

                'Validate_Code("JOURNAL_NO")

                If grdGLTJRNLR.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select a Previously Entered Recurring Journal No"
                Else
                    Absx1.txtFor("JOURNAL_NO").Text = grdGLTJRNLR.Selected.Rows(0).Cells("JOURNAL_NO").Text
                End If

                If Absx1.dteFor("JOURNAL_DATE").Text = "" And eItemKey <> "Edit" Then
                    'Absx1.dteFor("JOURNAL_DATE").Value = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))
                    EMsg &= vbCr & "You Must Enter a Journal Date"
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("JOURNAL_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Post Entry to GL", "Update"

                'grdGLTJRNL2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.TemplateOnBottom

                Save_Header_Fields(UltraGroupBox1)

                Dim T As Decimal = 0
                Dim TT As Decimal = 0
                Dim R As Int64 = 0

                Dim zero_lines As String = ""

                For Each rowGLTJRNL2 As DataRow In dst.Tables("GLTJRNL2").Select("", "", DataViewRowState.CurrentRows)
                    R = R + 1
                    Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", rowGLTJRNL2.Item("ACCT_CODE"))
                    If rowGLTACCT1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Acct (see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                    Else
                        If rowGLTACCT1.Item("ACCT_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Account Status not Active for " & rowGLTJRNL2.Item("ACCT_CODE") & " (see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                        Else
                            If rowGLTACCT1.Item("ACCT_SUB_CTL") & "" = "1" Then
                                EMsg &= vbCr & "Acct Code " & rowGLTJRNL2.Item("ACCT_CODE") & " is a Control Account - no Manual J/E permitted"
                            End If
                        End If
                    End If
                    For i As Integer = 2 To 4
                        Dim COLUMN_NAME As String = "SEG" & CStr(i) & "_CODE"
                        cdr = LookUp("GLTSEGM1", New String() {CStr(i), rowGLTJRNL2.Item(COLUMN_NAME) & ""})
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid " & ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & " (" & rowGLTJRNL2.Item(COLUMN_NAME) & ", Segment " & CStr(i) & " - see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                        Else
                            If cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                                EMsg &= vbCr & "Segment not Permitted for J/E (" & rowGLTJRNL2.Item(COLUMN_NAME) & ", Segment " & CStr(i) & " - see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                            End If
                            If cdr.Item("ACCT_SEG_STATUS") & "" <> "A" Then
                                EMsg &= vbCr & "Segment not Active (" & rowGLTJRNL2.Item(COLUMN_NAME) & ", Segment " & CStr(i) & " - see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                            End If
                            If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "0" Then
                                If rowGLTJRNL2.Item(COLUMN_NAME) <> ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                    EMsg &= vbCr & "Acct " & rowGLTJRNL2.Item("ACCT_CODE") & " requires Default Value (" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & ") for " & ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & " (Segment " & CStr(i) & " - see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                                End If
                            End If
                            If rowGLTACCT1.Item("ACCT_SEG" & CStr(i) & "_MAND") & "" = "1" Then
                                If rowGLTJRNL2.Item(COLUMN_NAME) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) Then
                                    EMsg &= vbCr & "Acct " & rowGLTJRNL2.Item("ACCT_CODE") & " requires non-Default Value (" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & ") for " & ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & " (Segment " & CStr(i) & " - see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                                End If
                            End If
                        End If
                    Next
                    If Round(Val(rowGLTJRNL2.Item("DETL_POSTING_AMT") & ""), 2) = 0 And Not recurring Then
                        zero_lines &= vbCr & "No DR/CR Amount entered (see Line " & rowGLTJRNL2.Item("JOURNAL_LNO") & ")"
                    End If
                    T = T + Round(Val(rowGLTJRNL2.Item("AMOUNT_DR") & "") - Val(rowGLTJRNL2.Item("AMOUNT_CR") & ""), 2)
                    TT = TT + Round(Val(rowGLTJRNL2.Item("DETL_POSTING_AMT") & ""), 2)
                Next
                If Round(T, 2) <> 0 Or Round(TT, 2) <> 0 Then
                    EMsg &= vbCr & "Journal is Out of Balance"
                End If
                If R = 0 Then
                    EMsg &= vbCr & "No Journal Details"
                End If

                If EMsg = "" Then
                    If zero_lines <> "" Then
                        If MsgBox("Do you want these lines removed?", vbYesNo, "There are Lines in this Entry with 0 value") = MsgBoxResult.Yes Then
                            For Each rowGLTJRNL2 As DataRow In _
                            dst.Tables("GLTJRNL2").Select("DETL_POSTING_AMT is Null or (DETL_POSTING_AMT < .005 and DETL_POSTING_AMT > -.005)", "", DataViewRowState.CurrentRows)
                                rowGLTJRNL2.Delete()
                            Next
                        Else
                            EMsg &= vbCr & zero_lines
                        End If
                    End If
                End If

                If Absx1.cmbFor("OPS_YYYYPP").Value & "" = "" Then
                    EMsg &= vbCr & "You must choose a period in which to book the J/E"
                Else
                    If Absx1.cmbFor("OPS_YYYYPP_REV").Value & "" <> "" And (Absx1.cmbFor("OPS_YYYYPP_REV").Value & "" = Absx1.cmbFor("OPS_YYYYPP").Value & "") Then
                        EMsg &= vbCr & "You Cannot Reverse a J/E into the same period that you are Booking it into"
                    End If
                End If

                If Not recurring Then
                    If Absx1.cmbFor("OPS_YYYYPP").Text < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                        EMsg &= vbCr & "Period " & Absx1.txtFor("LEGEND").Text & " is Closed"
                    End If
                    Dim YP_REV As String = Absx1.cmbFor("OPS_YYYYPP_REV").Value & ""
                    If YP_REV <> "" And (YP_REV < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")) Then
                        EMsg &= vbCr & "Period " & Mid(YP_REV, 1, 4) & "-" & Mid(YP_REV, 5, 2) & " (the reversing period) is Closed"
                    End If
                End If

                If Absx1.txtFor("JOURNAL_DESC").Text = "" Then
                    EMsg &= vbCr & "You Must Enter a Description"
                End If
                If Not recurring Then
                    If Absx1.dteFor("JOURNAL_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "You Must Enter a Journal Entry Reference Date"
                    End If
                End If

            Case "Reverse"
                If grdGLTJRNLP.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select a Previously Entered Journal No"
                Else
                    Absx1.txtFor("JOURNAL_NO").Text = grdGLTJRNLP.Selected.Rows(0).Cells("JOURNAL_NO").Text

                    cdr = LookUp("GLTJRNL1", Absx1.txtFor("JOURNAL_NO").Text)
                    If cdr.Item("OPS_YYYYPP") < ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") Then
                        EMsg &= vbCr & "You Cannot Reverse a J/E which was posted into a Closed Period"
                    End If
                    If cdr.Item("OPS_YYYYPP") <> cmbOPS_YYYYPP.Value Then
                        EMsg &= vbCr & "You Cannot Reverse the 'Reversal' of a J/E which was posted into a different Period" _
                        & vbCr & "(J/E " & Absx1.txtFor("JOURNAL_NO").Text & " was originally posted in " & cdr.Item("OPS_YYYYPP") & ")"
                    End If
                    If cdr.Item("JOURNAL_REVERSED_IND") & "" = "1" Then
                        EMsg &= vbCr & "You Cannot Reverse a J/E which was Already Reversed"
                    End If
                    If cdr.Item("JOURNAL_REVERSED_IND") & "" = "2" Then
                        EMsg &= vbCr & "You Cannot Reverse a J/E which was used to Reverse another"
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("JOURNAL_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg = "" Then
                    Dim YPs As String = ""
                    ASCMAIN1.sql = "Select Distinct OPS_YYYYPP from GLTDETL1 where JOURNAL_NO = '" & Absx1.txtFor("JOURNAL_NO").Text & "'"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                        YPs &= "," & row.Item("OPS_YYYYPP")
                    Next
                    If MsgBox("Do you want to Reverse this Entry?" _
                              & vbCrLf & vbCrLf & "Periods impacted: " & Mid(YPs, 2), vbYesNo, "Verification") = MsgBoxResult.Yes Then

                    Else
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                End If

            Case "Create Like"
                If grdGLTJRNLP.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select a Previously Entered Journal No"
                Else
                    Absx1.txtFor("JOURNAL_NO").Text = grdGLTJRNLP.Selected.Rows(0).Cells("JOURNAL_NO").Text
                End If

            Case "Print"
                If grdGLTJRNLP.Selected.Rows.Count <> 1 Then
                    EMsg &= vbCr & "You Must Select a Previously Entered Journal No"
                End If


            Case "Reverse this Entry"

                If EMsg = "" Then
                    Dim YPs As String = ""
                    ASCMAIN1.sql = "Select Distinct OPS_YYYYPP from GLTDETL1 where JOURNAL_NO = '" & Absx1.txtFor("JOURNAL_NO").Text & "'"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                        YPs &= "," & row.Item("OPS_YYYYPP")
                    Next
                    If MsgBox("OK to Reverse this Entry?" _
                              & vbCrLf & vbCrLf & "Periods impacted: " & Mid(YPs, 2), vbYesNo, "Verification") = MsgBoxResult.Yes Then
                    Else
                        EMsg &= vbCr & "Click Cancel to clear the screen"
                    End If
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
                recurring = True
                EntryMode = "N"
                txtModeOfOperation.Text = "You are Entering a New Recurring Journal Entry into the Recurring J/E List."
                Load_Record()
                Mode_Settings(True)

            Case "New Journal Entry"
                recurring = False
                EntryMode = "N"
                txtModeOfOperation.Text = "You are Entering a New Journal Entry into the General Ledger."
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                recurring = True
                EntryMode = "E"
                txtModeOfOperation.Text = "You are Editing a Recurring Journal Entry Definition."
                Load_Record()
                Mode_Settings(True)

            Case "View"
                recurring = False
                EntryMode = "V"
                txtModeOfOperation.Text = "You have Loaded a Journal Entry to View it."
                Load_Record()
                Mode_Settings(True)

            Case "Load to Post"
                recurring = False
                EntryMode = "L"
                txtModeOfOperation.Text = "You have Loaded a Recurring Entry Definition so that you can Post it into the GL after making Edits, as required."
                Load_Record()
                Mode_Settings(True)

            Case "Reverse"
                recurring = False
                EntryMode = "R"
                txtModeOfOperation.Text = "You have Loaded a previously Posted J/E so that you may Reverse it.  If this entry was entered as a B&R Entry, both sides will be reversed."
                Load_Record()
                Mode_Settings(True)

            Case "Create Like"
                recurring = False
                EntryMode = "C"
                txtModeOfOperation.Text = "You have Loaded a previously Posted J/E so that you may enter another one similar to it after making Edits, as required."
                Load_Record()
                Mode_Settings(True)

            Case "Post Entry to GL"
                Update_Record()
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Reverse DR-CR"
                For Each rowGLTJRNL2 As DataRow In dst.Tables("GLTJRNL2").Rows
                    Dim DETL_POSTING_AMT As Decimal = Val(rowGLTJRNL2.Item("DETL_POSTING_AMT") & "")
                    DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
                    rowGLTJRNL2.Item("DETL_POSTING_AMT") = DETL_POSTING_AMT
                    If DETL_POSTING_AMT < 0 Then
                        rowGLTJRNL2.Item("AMOUNT_DR") = Null
                        rowGLTJRNL2.Item("AMOUNT_CR") = -1 * DETL_POSTING_AMT
                    Else
                        rowGLTJRNL2.Item("AMOUNT_DR") = DETL_POSTING_AMT
                        rowGLTJRNL2.Item("AMOUNT_CR") = Null
                    End If
                Next

            Case "Reverse this Entry"
                Reverse_Entry()
                Mode_Settings(False)

            Case "Print View"
                Print_Report(HFs("JOURNAL_NO"))

            Case "Print"
                Dim YP As String = Absx1.cmbFor("OPS_YYYYPP").Value
                EntryMode = "P"
                Dim JOURNAL_NO As String = grdGLTJRNLP.Selected.Rows(0).Cells("JOURNAL_NO").Text
                Fill_Records("GLTJRNL1", JOURNAL_NO)
                Fill_Records("GLTDETL1", JOURNAL_NO)
                Print_Report(JOURNAL_NO)
                dst.Tables("GLTJRNL1").Rows.Clear()
                Absx1.cmbFor("OPS_YYYYPP").Value = YP
                EntryMode = ""
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Recurring Journals")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Load to Post").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    If Not ScreenMode Or EntryMode = "E" Then
                        .Items("Delete").Settings.Enabled = iScreenMode
                    End If

                    .Visible = Not (EntryMode = "R") And Not (EntryMode = "V") And Not InquiryMode
                End With

                With .Groups("Journal Entry")
                    .Items("New Journal Entry").Settings.Enabled = not_iScreenMode
                    .Items("Reverse this Entry").Settings.Enabled = DefaultableBoolean.False

                    If EntryMode <> "R" Then
                        .Items("Post Entry to GL").Settings.Enabled = iScreenMode
                    Else
                        .Items("Reverse this Entry").Settings.Enabled = DefaultableBoolean.True
                    End If
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("View").Visible = Not ScreenMode
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")

                    .Items("Cancel").Visible = Not (EntryMode = "V") And Not InquiryMode
                    .Items("Print View").Visible = (EntryMode = "V")
                    .Items("New Journal Entry").Visible = Not (EntryMode = "V") And Not InquiryMode
                    .Items("Post Entry to GL").Visible = Not (EntryMode = "V") And Not InquiryMode
                    .Items("Reverse this Entry").Visible = Not (EntryMode = "V") And Not InquiryMode

                End With

                If EntryMode <> "R" Then
                    .Groups("Misc Tools").Items("Reverse DR-CR").Settings.Enabled = iScreenMode
                    .Groups("Misc Tools").Visible = Not (EntryMode = "V") And Not InquiryMode
                End If

                With .Groups("Select Journal Entry to")
                    .Items("Reverse").Settings.Enabled = not_iScreenMode
                    .Items("Create Like").Settings.Enabled = not_iScreenMode
                    .Items("Print").Settings.Enabled = not_iScreenMode
                    .Visible = Not tf And Not InquiryMode
                End With

                .Groups("Mode of Operation").Visible = tf And Not InquiryMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If EntryMode <> "R" And EntryMode <> "V" And Not InquiryMode Then
            Absx1.txtFor("JOURNAL_DESC").ReadOnly = False
            Absx1.cmbFor("OPS_YYYYPP").ReadOnly = False
            'Absx1.txtFor("OPS_YYYYPP").ButtonsRight(0).Enabled = True
            Absx1.cmbFor("OPS_YYYYPP_REV").ReadOnly = False
            'Absx1.txtFor("OPS_YYYYPP_REV").ButtonsRight(0).Enabled = True
            Absx1.dteFor("JOURNAL_DATE").ReadOnly = False
        End If

        Set_Read_Only_for_ctl(Absx1.txtFor("JOURNAL_COMMENT"), (EntryMode = "V"))
        Set_Read_Only_for_ctl(Absx1.txtFor("JOURNAL_NO"), ScreenMode)

        If InquiryMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("JOURNAL_DESC"), True)
            Set_Read_Only_for_ctl(Absx1.dteFor("JOURNAL_DATE"), True)
            Set_Read_Only_for_ctl(Absx1.cmbFor("OPS_YYYYPP_REV"), True)
        End If

        tabGLTJRNL1.Visible = Not tf
        grdGLTJRNL2.Visible = True
        SplitContainer1.Visible = tf

        If ScreenMode Then
            If recurring Then
                Absx1.cmbFor("OPS_YYYYPP").Visible = False
                Absx1.dteFor("JOURNAL_DATE").Visible = False
                Absx1.cmbFor("OPS_YYYYPP_REV").Visible = False
                Absx1.txtFor("LEGEND").Visible = False
                lblOPS_YYYYPP.Visible = False
                lblJOURNAL_DATE.Visible = False
                lblOPS_YYYYPP_REV.Visible = False
            End If
        Else
            Absx1.cmbFor("OPS_YYYYPP").Visible = True
            Absx1.dteFor("JOURNAL_DATE").Visible = True
            Absx1.cmbFor("OPS_YYYYPP_REV").Visible = True
            Absx1.txtFor("LEGEND").Visible = True
            lblOPS_YYYYPP.Visible = True
            lblJOURNAL_DATE.Visible = True
            lblOPS_YYYYPP_REV.Visible = True
        End If


        If ScreenMode Then
            If EntryMode = "R" Or EntryMode = "V" Then
                grdGLTJRNL2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdGLTJRNL2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdGLTJRNL2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdGLTJRNL2.DisplayLayout.Bands(0).Columns("DIST_CODE").Hidden = True
            Else
                grdGLTJRNL2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdGLTJRNL2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdGLTJRNL2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdGLTJRNL2.DisplayLayout.Bands(0).Columns("DIST_CODE").Hidden = True ' FOR NOW

                UltraExplorerBar1.Groups("Recurring Journals").Visible = recurring
                UltraExplorerBar1.Groups("Journal Entry").Visible = Not recurring
            End If
        Else
            Clear_Record()
            Setup_Menu()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("GLTJRNL1").Rows.Clear()
        dst.Tables("GLTJRNL2").Rows.Clear()

        Fill_Records("GLTJRNLR")
        Fill_Records("GLTJRNLR2")

        dst.EnforceConstraints = True

        txtModeOfOperation.Text = ""
        Application.DoEvents()
        If HFs.ContainsKey("OPS_YYYYPP") AndAlso HFs("OPS_YYYYPP") <> "" Then
            Absx1.cmbFor("OPS_YYYYPP").Text = HFs("OPS_YYYYPP")
        Else
            Absx1.cmbFor("OPS_YYYYPP").Text = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") ' ASCMAIN1.CYP
        End If
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            HFs("JOURNAL_NO") = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO") ' "GLTJRNL1.JOURNAL_NO")
        End If

        rowGLTJRNL1 = Fill_Record("GLTJRNL1", New String() {HFs("JOURNAL_NO")}, EntryMode = "N")
        Fill_Records("GLTJRNL2", New String() {HFs("JOURNAL_NO")})

        If EntryMode = "V" Then
            ASCMAIN1.sql = "" _
                & "Select GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO" & vbCrLf _
                & ", GLTDETL1.ACCT_CODE, GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" & vbCrLf _
                & ", GLTDETL1.DETL_POSTING_AMT, GLTDETL1.DETL_DESC, NULL DIST_CODE" & vbCrLf _
                & ", GLTACCT1.ACCT_DESC" & vbCrLf _
                & ", CASE WHEN GLTDETL1.DETL_POSTING_AMT > 0 THEN GLTDETL1.DETL_POSTING_AMT ELSE NULL END AMOUNT_DR" & vbCrLf _
                & ", CASE WHEN GLTDETL1.DETL_POSTING_AMT < 0 THEN -1 * GLTDETL1.DETL_POSTING_AMT ELSE NULL END AMOUNT_CR" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1,GLTACCT1 " & vbCrLf _
                & " where GLTACCT1.ACCT_CODE (+) = GLTDETL1.ACCT_CODE" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTJRNL1.OPS_YYYYPP = GLTDETL1.OPS_YYYYPP" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = '" & HFs("JOURNAL_NO") & "'"
            Fill_Records("GLTJRNL2", "", True, ASCMAIN1.sql)

            Fill_Records("GLTDETL1", HFs("JOURNAL_NO"))
        End If

        If EntryMode = "R" Or EntryMode = "C" Then
            For Each row As DataRow In dst.Tables("GLTJRNLP2").Select("JOURNAL_NO = '" & HFs("JOURNAL_NO") & "'")
                Dim rowGLTJRNL2 As DataRow = dst.Tables("GLTJRNL2").NewRow
                For i As Integer = 0 To rowGLTJRNL2.ItemArray.Length - 1
                    Dim COLUMN_NAME As String = dst.Tables("GLTJRNL2").Columns(i).ColumnName
                    If COLUMN_NAME = "DIST_CODE" Then
                    Else
                        rowGLTJRNL2.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    End If
                Next
                dst.Tables("GLTJRNL2").Rows.Add(rowGLTJRNL2)
            Next
        End If

        If EntryMode = "N" Or EntryMode = "L" Then
            With rowGLTJRNL1
                .Item("JOURNAL_NO") = HFs("JOURNAL_NO")
                .Item("JOURNAL_DESC") = HFs("JOURNAL_DESC")
                If recurring Then
                    .Item("JOURNAL_TYPE") = "GLRE"
                Else
                    .Item("JOURNAL_TYPE") = "GLJE"
                    .Item("OPS_YYYYPP") = HFs("OPS_YYYYPP")
                    .Item("OPS_YYYYPP_REV") = HFs("OPS_YYYYPP_REV")
                    If HFs("JOURNAL_DATE") Is Nothing Then
                        .Item("JOURNAL_DATE") = Format(Now, "MM/dd/yyyy")
                    Else
                        .Item("JOURNAL_DATE") = HFs("JOURNAL_DATE")
                    End If
                End If
            End With
        End If

        If EntryMode = "C" Or EntryMode = "L" Then
            HFs("JOURNAL_NO") = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO") '"GLTJRNL1.JOURNAL_NO")
            EnforceConstraints(False)
            With rowGLTJRNL1
                .Item("JOURNAL_NO") = HFs("JOURNAL_NO")
                HFs("JOURNAL_DATE") = Format$(Now, "MM/dd/yyyy")
                .Item("JOURNAL_DATE") = HFs("JOURNAL_DATE")
                .Item("JOURNAL_TYPE") = "GLJE"
                .Item("OPS_YYYYPP") = HFs("OPS_YYYYPP")
                HFs("OPS_YYYYPP_REV") = ""
                .Item("OPS_YYYYPP_REV") = HFs("OPS_YYYYPP_REV")
                .Item("JOURNAL_REVERSED") = Null
                .Item("JOURNAL_REVERSED_IND") = Null
                For Each rowGLTJRNL2 As DataRow In dst.Tables("GLTJRNL2").Rows
                    rowGLTJRNL2.Item("JOURNAL_NO") = HFs("JOURNAL_NO")
                Next
            End With
            EnforceConstraints(True)
        End If

        grdGLTJRNL2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdGLTJRNL2.DisplayLayout.Bands(0).SortedColumns.Add("JOURNAL_LNO", False)

        If EntryMode = "N" Then
            'grdGLTJRNL2.Focus()
            'grdGLTJRNL2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)

            'grdGLTJRNL2.DisplayLayout.Bands(0).AddNew()
            'If grdGLTJRNL2.ActiveRow.IsAddRow Then
            '    grdGLTJRNL2.ActiveCell = grdGLTJRNL2.ActiveRow.Cells("ACCT_CODE")
            '    ' Application.DoEvents()
            '    grdGLTJRNL2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            '    ' grdGLTJRNL2.ActiveCell.IsInEditMode


            'End If
        End If
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            If EntryMode <> "E" Then
                dst.Tables("GLTJRNL1").AcceptChanges()
                rowGLTJRNL1.SetAdded()
                rowGLTJRNL1.Item("INIT_DATE") = DATETIME_STAMP
                rowGLTJRNL1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowGLTJRNL1.Item("LAST_DATE") = Null
                rowGLTJRNL1.Item("LAST_OPER") = Null
            Else
                rowGLTJRNL1.Item("LAST_DATE") = DATETIME_STAMP
                rowGLTJRNL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                ' WHY ARE THE NEXT 2 LINES NEC? ROWSTATE IS UNMODIFIED WHEN CLEARLY THE ROW IS MODIFIED
                dst.Tables("GLTJRNL1").AcceptChanges()
                rowGLTJRNL1.SetModified()
            End If
            Update_Record_TDA("GLTJRNL1")
            If recurring Then
                Update_Record_TDA("GLTJRNL2", "JOURNAL_NO = '" & HFs("JOURNAL_NO") & "'")
            Else
                dst.Tables("GLTDETL1").Rows.Clear()

                For Each rowGLTJRNL2 As DataRow In dst.Tables("GLTJRNL2").Select("", "", DataViewRowState.CurrentRows)
                    Dim rowGLTDETL1 As DataRow = dst.Tables("GLTDETL1").NewRow
                    For i As Integer = 0 To rowGLTJRNL2.ItemArray.Length - 1
                        Dim COLUMN_NAME As String = dst.Tables("GLTJRNL2").Columns(i).ColumnName
                        If COLUMN_NAME = "DIST_CODE" Or COLUMN_NAME = "ACCT_DESC" Or COLUMN_NAME = "AMOUNT_DR" Or COLUMN_NAME = "AMOUNT_CR" Then
                        Else
                            rowGLTDETL1.Item(COLUMN_NAME) = rowGLTJRNL2.Item(COLUMN_NAME)
                        End If
                    Next
                    rowGLTDETL1.Item("OPS_YYYYPP") = rowGLTJRNL1.Item("OPS_YYYYPP")
                    rowGLTDETL1.Item("DETL_CTL_DATE") = rowGLTJRNL1.Item("JOURNAL_DATE")
                    rowGLTDETL1.Item("DETL_EXE_NO") = XNO
                    dst.Tables("GLTDETL1").Rows.Add(rowGLTDETL1)

                    If rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "" <> "" Then
                        Dim rowGLTDETL1_REV As DataRow = dst.Tables("GLTDETL1").NewRow
                        For i As Integer = 0 To rowGLTDETL1.ItemArray.Length - 1
                            rowGLTDETL1_REV.Item(i) = rowGLTDETL1.Item(i)
                        Next
                        rowGLTDETL1_REV.Item("OPS_YYYYPP") = rowGLTJRNL1.Item("OPS_YYYYPP_REV")
                        rowGLTDETL1_REV.Item("DETL_POSTING_AMT") = -1 * Val(rowGLTDETL1_REV.Item("DETL_POSTING_AMT") & "")
                        dst.Tables("GLTDETL1").Rows.Add(rowGLTDETL1_REV)
                    End If
                Next

                InterCompany(HFs("JOURNAL_NO"), HFs("OPS_YYYYPP"))
                If rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "" <> "" Then
                    InterCompany(HFs("JOURNAL_NO"), rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "")
                End If

                Update_Record_TDA("GLTDETL1")

                Update_GLTACCT3(HFs("JOURNAL_NO"), HFs("OPS_YYYYPP"))
                If rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "" <> "" Then
                    Update_GLTACCT3(HFs("JOURNAL_NO"), rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "")
                End If
            End If

            CommitTrans("Update Complete")
            If Not recurring Then
                If AUTO_REPORT Then
                    Print_Report(, True)
                End If
            End If
            '  Print_Report(print_directly:=True)

        Catch ex As Exception
            'Stop
            Rollback("Error Occurred - Please call ABS" & vbCr & vbCr & ex.Message, ex)
        End Try

    End Sub

    Sub Delete_Record()
        BeginTrans()

        Delete_Records("GLTJRNL1")
        Delete_Records("GLTJRNL2")

        CommitTrans("Delete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where JOURNAL_NO = '" & HFs("JOURNAL_NO") & "'")
    End Sub


    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View"
                Absx1.txtFor("JOURNAL_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function


#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTJRNL2, "BBB", "Insert & Replicate", "Load Expensify Report", "Load XLS JE")
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

        Select Case e.SourceControl.Name
            Case "grdGLTJRNL2"
                tlb_btn = DirectCast(tlb_pop.Tools("Insert & Replicate"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" And (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT"))
                tlb_btn = DirectCast(tlb_pop.Tools("Load Expensify Report"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" And (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT"))
                tlb_btn = DirectCast(tlb_pop.Tools("Load XLS JE"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "N" And (ASCMAIN1.CLIENT = "VAN"))
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Select Case e.Tool.Key

            Case "Load XLS JE"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.RestoreDirectory = True

                    '  Excel_Import = -1

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME <> "" Then

                    Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                    Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(1)
                    Dim range As SpreadsheetGear.IRange = Nothing

                    Dim BAD_ACCTS As New List(Of String)

                    '  grdGLTJRNL2.SuspendLayout()
                    grdGLTJRNL2.Visible = False
                    ASCMAIN1.Progress("Now Loading from XLS")

                    Dim r As Integer = 1
                    Do While oSheet.Cells(r, 15).Value & "" <> "" Or oSheet.Cells(r + 1, 15).Value & "" <> ""
                        If oSheet.Cells(r, 15).Value & "" <> "" Then
                            Dim GL_CODE As String = oSheet.Cells(r, 15).Value & ""
                            Dim ACCT_CODE As String = Mid(GL_CODE, 1, 4)
                            Dim rowGLTACCT1 As DataRow = LookUp("GLTACCT1", ACCT_CODE)
                            If rowGLTACCT1 Is Nothing Then
                                If Not BAD_ACCTS.Contains(ACCT_CODE) Then
                                    BAD_ACCTS.Add(ACCT_CODE)
                                End If
                            Else
                                Dim AMT_DR As Decimal = Val(oSheet.Cells(r, 17).Value & "")
                                Dim AMT_cR As Decimal = Val(oSheet.Cells(r, 19).Value & "")

                                With grdGLTJRNL2
                                    If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                                        .ActiveRow = Nothing
                                    End If
                                    .DisplayLayout.Bands(0).AddNew()
                                    With .ActiveRow
                                        .Cells("ACCT_CODE").Value = ACCT_CODE
                                        .Cells("DETL_DESC").Value = oSheet.Cells(r, 13).Value & ""
                                        If AMT_DR > 0 Then
                                            .Cells("AMOUNT_DR").Value = AMT_DR
                                        Else
                                            .Cells("AMOUNT_CR").Value = AMT_cR
                                        End If
                                        .Update()
                                    End With
                                End With
                            End If

                        End If

                        r += 1
                        ASCMAIN1.Progress("-", CStr(r))
                    Loop

                    ASCMAIN1.Progress("")

                    'grdGLTJRNL2.ResumeLayout()
                    grdGLTJRNL2.Visible = True

                    Sort_grdColumns(grdGLTJRNL2, "JOURNAL_LNO")

                    If BAD_ACCTS.Count <> 0 Then
                        MsgBox("The following invalide Accts have been encountered: " & Join(BAD_ACCTS.ToArray, ","), MsgBoxStyle.OkOnly, "Warning")
                    End If
                    MsgBox("XLS JE has been Loaded", MsgBoxStyle.OkOnly, "Success")
                End If

            Case "Load Expensify Report"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsx files (*.xlsx)|*.xlsx"
                    openFileDialog1.RestoreDirectory = True

                    '  Excel_Import = -1

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME <> "" Then

                    Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                    Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
                    Dim range As SpreadsheetGear.IRange = Nothing

                    Dim ACCT_CODEs As New Dictionary(Of String, Decimal)
                    Dim CASHYNs As New Dictionary(Of String, Decimal)
                    Dim r As Integer = 1
                    Do While oSheet.Cells(r, 0).Value & "" <> ""
                        Dim GL_CODE As String = oSheet.Cells(r, 4).Value & ""
                        Dim COLL As String = oSheet.Cells(r, 5).Value & ""
                        If COLL = "" Then COLL = "000"
                        Dim ACCT_CODE As String = GL_CODE & ":" & COLL
                        Dim AMT As Decimal = Val(oSheet.Cells(r, 2).Value & "")
                        Dim CASHYN As String = oSheet.Cells(r, 7).Value & ""

                        If Not ACCT_CODEs.ContainsKey(ACCT_CODE) Then
                            ACCT_CODEs.Add(ACCT_CODE, 0)
                        End If
                        ACCT_CODEs(ACCT_CODE) += AMT

                        If Not CASHYNs.ContainsKey(CASHYN) Then
                            CASHYNs.Add(CASHYN, 0)
                        End If
                        CASHYNs(CASHYN) -= AMT
                        r += 1
                    Loop

                    Dim ACCT_CODE_AMEX As String = ""
                    Dim ACCT_CODE_CASH As String = ""

                    For Each CASHYN As String In CASHYNs.Keys
                        Dim ACCT_CODE As String = ""
                        Dim DETL_DESC As String = ""
                        If CASHYN.ToUpper = "YES" Then ACCT_CODE = "161200" : DETL_DESC = "Employee"
                        If CASHYN.ToUpper = "NO" Then ACCT_CODE = "221400" : DETL_DESC = "Amex"
                        With grdGLTJRNL2
                            If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                                .ActiveRow = Nothing
                            End If
                            .DisplayLayout.Bands(0).AddNew()
                            With .ActiveRow
                                .Cells("ACCT_CODE").Value = ACCT_CODE
                                .Cells("DETL_DESC").Value = DETL_DESC
                                If CASHYNs(CASHYN) >= 0 Then
                                    .Cells("AMOUNT_DR").Value = CASHYNs(CASHYN)
                                Else
                                    .Cells("AMOUNT_CR").Value = -1 * CASHYNs(CASHYN)
                                End If
                                .Update()
                            End With
                        End With
                    Next

                    For Each ACCT_CODE As String In ACCT_CODEs.Keys
                        With grdGLTJRNL2
                            If .ActiveRow IsNot Nothing AndAlso .ActiveRow.IsAddRow Then
                                .ActiveRow = Nothing
                            End If
                            .DisplayLayout.Bands(0).AddNew()
                            With .ActiveRow
                                .Cells("ACCT_CODE").Value = Split(ACCT_CODE, ":")(0)
                                .Cells("SEG4_CODE").Value = Split(ACCT_CODE, ":")(1)
                                If ACCT_CODEs(ACCT_CODE) >= 0 Then
                                    .Cells("AMOUNT_DR").Value = ACCT_CODEs(ACCT_CODE)
                                Else
                                    .Cells("AMOUNT_CR").Value = -1 * ACCT_CODEs(ACCT_CODE)
                                End If
                                .Update()
                            End With
                        End With
                    Next

                    Sort_grdColumns(grdGLTJRNL2, "JOURNAL_LNO")

                    MsgBox("Expensify Report has been Loaded", MsgBoxStyle.OkOnly, "Success")
                End If

            Case "Insert & Replicate"
                If grdGLTJRNL2.ActiveRow Is Nothing Or grdGLTJRNL2.Selected.Rows.Count <> 1 Then
                    MsgBox("Please Select a (single) Row before Inserting", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                Else
                    If Not grdGLTJRNL2.Selected.Rows(0).IsActiveRow Then
                        MsgBox("Please Select a (single) Row before Inserting", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        Exit Sub
                    End If
                End If
                Dim rowGLTJRNL2 As DataRow
                Dim r_start As DataRow = dst.Tables("GLTJRNL2").NewRow
                Dim JOURNAL_LNO_start As Integer = Val(grdGLTJRNL2.ActiveRow.Cells("JOURNAL_LNO").Text)
                For Each rowGLTJRNL2 In dst.Tables("GLTJRNL2").Select("JOURNAL_LNO >= " & CStr(JOURNAL_LNO_start), "JOURNAL_LNO DESC", DataViewRowState.CurrentRows)
                    If Val(rowGLTJRNL2("JOURNAL_LNO")) = JOURNAL_LNO_start Then
                        r_start.ItemArray = rowGLTJRNL2.ItemArray
                    End If
                    rowGLTJRNL2.Item("JOURNAL_LNO") = Val(rowGLTJRNL2.Item("JOURNAL_LNO")) + 1
                Next
                dst.Tables("GLTJRNL2").Rows.Add(r_start)

                grdGLTJRNL2.DisplayLayout.Bands(0).SortedColumns.Clear()
                grdGLTJRNL2.DisplayLayout.Bands(0).SortedColumns.Add("JOURNAL_LNO", False)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "JOURNAL_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "JOURNAL_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If

            Case "JOURNAL_DESC"
                If EntryMode = "" Then
                    If e.KeyCode = Windows.Forms.Keys.Enter Then
                        If tabGLTJRNL1.ActiveTab.Key = "Journals Posted in Period" Then
                            If UltraExplorerBar1.Groups("Journal Entry").Items("New Journal Entry").Settings.Enabled = DefaultableBoolean.True Then
                                Click_Command("New Journal Entry")
                            End If
                        ElseIf tabGLTJRNL1.ActiveTab.Key = "Recurring Journals" Then
                            If UltraExplorerBar1.Groups("Recurring Entry Definitions").Items("New").Settings.Enabled = DefaultableBoolean.True Then
                                Click_Command("New")
                            End If
                        End If
                    End If
                End If

        End Select
    End Sub
#End Region

    Sub Populate_grdGLTJRNLP()
        If Absx1.cmbFor("OPS_YYYYPP").Text = "" Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        Fill_Records("GLTJRNLP", Absx1.cmbFor("OPS_YYYYPP").Text)
        Fill_Records("GLTJRNLP2", Absx1.cmbFor("OPS_YYYYPP").Text)
        EnforceConstraints(True)

        grdGLTJRNLP.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdGLTJRNLP.DisplayLayout.Bands(0).SortedColumns.Add("JOURNAL_NO", True)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Entry()
        BeginTrans()

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO") '"GLTJRNL1.JOURNAL_NO")

        Dim rowGLTJRNL1_REVERSING As DataRow = dst.Tables("GLTJRNL1").NewRow
        With rowGLTJRNL1_REVERSING
            For i As Integer = 0 To rowGLTJRNL1.ItemArray.Length - 1
                .Item(i) = rowGLTJRNL1.Item(i)
            Next
            .Item("JOURNAL_NO") = JOURNAL_NO
            .Item("JOURNAL_DATE") = Format$(Now, "MM/dd/yyyy")
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = Null
            .Item("LAST_DATE") = Null
            .Item("JOURNAL_REVERSED") = HFs("JOURNAL_NO")
            .Item("JOURNAL_REVERSED_IND") = "2"
        End With
        dst.Tables("GLTJRNL1").Rows.Add(rowGLTJRNL1_REVERSING)

        rowGLTJRNL1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowGLTJRNL1.Item("LAST_DATE") = DATETIME_STAMP
        rowGLTJRNL1.Item("JOURNAL_REVERSED") = JOURNAL_NO
        rowGLTJRNL1.Item("JOURNAL_REVERSED_IND") = "1"

        Update_Record_TDA("GLTJRNL1")

        Dim YPs As New List(Of String)
        Fill_Records("GLTDETL1", HFs("JOURNAL_NO"))
        For Each rowGLTDETL1 As DataRow In dst.Tables("GLTDETL1").Rows
            rowGLTDETL1.Item("JOURNAL_NO") = JOURNAL_NO
            rowGLTDETL1.Item("DETL_POSTING_AMT") = -1 * Val(rowGLTDETL1.Item("DETL_POSTING_AMT") & "")
            rowGLTDETL1.AcceptChanges()
            rowGLTDETL1.SetAdded()
            Dim OPS_YYYYPP As String = rowGLTDETL1.Item("OPS_YYYYPP")
            If Not YPs.Contains(OPS_YYYYPP) Then YPs.Add(OPS_YYYYPP)
        Next
        Update_Record_TDA("GLTDETL1")

        For Each YP As String In YPs
            Update_GLTACCT3(JOURNAL_NO, YP)
        Next
        'Update_GLTACCT3(JOURNAL_NO, rowGLTJRNL1.Item("OPS_YYYYPP"))
        'If rowGLTJRNL1.Item("OPS_YYYYPP_REV") & "" <> "" Then
        '    Update_GLTACCT3(JOURNAL_NO, rowGLTJRNL1.Item("OPS_YYYYPP_REV"))
        'End If

        CommitTrans("Reversal Completed using Journal " & JOURNAL_NO)

        If AUTO_REPORT Then
            Print_Report(, True)
        End If

    End Sub

    Private Sub grdGLTJRNL2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTJRNL2.AfterCellUpdate
        If e.Cell.Column.Key = "AMOUNT_DR" And Val(e.Cell.Value & "") <> 0 Then
            grdGLTJRNL2.ActiveRow.Cells("AMOUNT_CR").Value = Null
        ElseIf e.Cell.Column.Key = "AMOUNT_CR" And Val(e.Cell.Value & "") <> 0 Then
            grdGLTJRNL2.ActiveRow.Cells("AMOUNT_DR").Value = Null
        End If
        'If e.Cell.Column.Key = "AMOUNT_DR" Or e.Cell.Column.Key = "AMOUNT_CR" Then
        '    e.Cell.Row.Cells("DETL_POSTING_AMT").Value = Val(e.Cell.Value)
        'End If



        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""
                'e.Cell.Value = ASCMAIN1.Format_Field(ACCT_CODE, e.Cell.Column.Key)

                grdCodeDesc(grdGLTJRNL2, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                If cdr IsNot Nothing Then e.Cell.Row.Cells("ACCT_DESC").Value = cdr.Item("ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"
                'With grdGLTJRNL2.ActiveRow
                '    .Cells("LINE_AMOUNT").Value = Val(.Cells("ORDR_QTY").Value & "") * Val(.Cells("ORDR_UNIT_PRICE").Value & "")
                'End With
            Case "AMOUNT_DR", "AMOUNT_CR"
                With grdGLTJRNL2.ActiveRow
                    .Cells("DETL_POSTING_AMT").Value = Val(.Cells("AMOUNT_DR").Value & "") - Val(.Cells("AMOUNT_CR").Value & "")
                    e.Cell.Row.Cells("DETL_POSTING_AMT").Value = Val(e.Cell.Row.Cells("AMOUNT_DR").Value & "") - Val(e.Cell.Row.Cells("AMOUNT_CR").Value & "")
                End With
        End Select
    End Sub

    Private Sub tabGLTJRNL1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabGLTJRNL1.SelectedTabChanged
        Setup_Menu()
    End Sub

    Sub Setup_Menu()
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Recurring Journals").Visible = (tabGLTJRNL1.ActiveTab.Key = "Recurring Entry Definitions") And Not InquiryMode
                .Groups("Select Journal Entry to").Visible = (tabGLTJRNL1.ActiveTab.Key = "Journals Posted in Period") And Not InquiryMode
                .Groups("Journal Entry").Visible = True
            End With
        End If

    End Sub

    Private Sub cmbOPS_YYYYPP_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOPS_YYYYPP.AfterCloseUp

    End Sub

    Private Sub cmbOPS_YYYYPP_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOPS_YYYYPP.Enter
        ' Stop
    End Sub

    Private Sub cmbOPS_YYYYPP_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles cmbOPS_YYYYPP.InitializeLayout

    End Sub

    Private Sub cmbOPS_YYYYPP_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbOPS_YYYYPP.KeyDown
        ' Stop
    End Sub

    Private Sub cmbOPS_YYYYPP_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOPS_YYYYPP.Leave

    End Sub

    Private Sub cmbOPS_YYYYPP_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbOPS_YYYYPP.ValueChanged
        If EntryMode = "" Then
            Populate_grdGLTJRNLP()
        End If
    End Sub

    Private Sub UltraGroupBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraGroupBox1.Click

    End Sub

    Private Sub grdGLTJRNL2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTJRNL2.AfterExitEditMode
        If grdGLTJRNL2.ActiveCell Is Nothing Then Exit Sub

        Select Case grdGLTJRNL2.ActiveCell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = grdGLTJRNL2.ActiveCell.Text
                If ACCT_CODE <> "" Then
                    grdGLTJRNL2.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTJRNL2.ActiveCell.Column.Key)
                End If
        End Select

    End Sub

    Private Sub grdGLTJRNL2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTJRNL2.AfterRowActivate
        With grdGLTJRNL2
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdGLTJRNL2.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdGLTJRNL2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTJRNL2.BeforeRowUpdate
        'If e.Cell.Column.Key = "AMOUNT_DR" Or e.Cell.Column.Key = "AMOUNT_CR" Then
        '    e.Cell.Row.Cells("DETL_POSTING_AMT").Value = Val(e.Cell.Value)
        'End If
        With grdGLTJRNL2
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    If cdr.Item("ACCT_STATUS") & "" <> "A" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is not Active", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                    If cdr.Item("ACCT_SUB_CTL") & "" = "1" Then
                        MsgBox("Acct Code " & e.Row.Cells("ACCT_CODE").Text & " is a Control Account - no Manual J/E permitted", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        Else
                            If cdr.Item("ACCT_SEG_STATUS") & "" <> "A" Then
                                MsgBox(e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " Code " & e.Row.Cells(COLUMN_NAME).Text & " is not Active", MsgBoxStyle.OkOnly, "Cannot Update Row")
                                e.Cancel = True
                            End If
                            If cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                                MsgBox(e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " Code " & e.Row.Cells(COLUMN_NAME).Text & " is not permitted to be used in a Manual J/E", MsgBoxStyle.OkOnly, "Cannot Update Row")
                                e.Cancel = True
                            End If
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("JOURNAL_NO").Text = "" Then
                    .ActiveRow.Cells("JOURNAL_NO").Value = Absx1.CtlFor("JOURNAL_NO").Text
                    .ActiveRow.Cells("JOURNAL_LNO").Value = Val(dst.Tables("GLTJRNL2").Compute("Max(JOURNAL_LNO)", "") & "") + 1
                End If
            End If
        End With

    End Sub

    Private Sub grdGLTJRNL2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTJRNL2.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdGLTJRNL2.ActiveCell.Column.Key
            Case "ACCT_CODE"
                sql_where = "NVL(ACCT_STATUS,'X') = 'A' and NVL(ACCT_SUB_CTL,'0') <> '1'"

            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"
                sql_where = "NVL(ACCT_SEG_STATUS,'X') = 'A' and NVL(ACCT_SEG_NO_GL,'0') <> '1'"
        End Select

        grdClickCellButton(grdGLTJRNL2, sql_where, False)
    End Sub

    Private Sub grdGLTJRNL2_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdGLTJRNL2.DoubleClickCell
        If e.Cell.Column.Key = "AMOUNT_DR" _
        Or e.Cell.Column.Key = "AMOUNT_CR" _
        And grdGLTJRNL2.ActiveRow.IsAddRow Then
            Dim T As Decimal = Val(dst.Tables("GLTJRNL2").Compute("SUM(DETL_POSTING_AMT)", "") & "")
            If T < 0 Then
                grdGLTJRNL2.ActiveRow.Cells("AMOUNT_DR").Value = -1 * T
            ElseIf T > 0 Then
                grdGLTJRNL2.ActiveRow.Cells("AMOUNT_CR").Value = T
            End If
        End If
    End Sub

    Private Sub grdGLTJRNL2_SummaryValueChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.SummaryValueChangedEventArgs) Handles grdGLTJRNL2.SummaryValueChanged
        'Dim T As Decimal = Val(dst.Tables("GLTJRNL2").Compute("SUM(DETL_POSTING_AMT)", ""))
        Dim tc As Decimal = Val(grdGLTJRNL2.Rows.SummaryValues(grdGLTJRNL2.DisplayLayout.Bands(0).Summaries("DETL_POSTING_AMT")).Value & "")
        If CLng(tc * 100) = 0 Then
            grdGLTJRNL2.DisplayLayout.Bands(0).SummaryFooterCaption = "Journal is in Balance"
        Else
            grdGLTJRNL2.DisplayLayout.Bands(0).SummaryFooterCaption = "Journal is out of Balance.  It needs " & Format(-1 * tc, "##,##0.00DR;##,##0.00CR")
        End If
    End Sub

    Private Sub grdGLTJRNLP_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTJRNLP.DoubleClickRow
        If e.Row.IsDataRow Then
            Dim JOURNAL_NO As String = e.Row.Cells("JOURNAL_NO").Value
            Absx1.txtFor("JOURNAL_NO").Text = JOURNAL_NO
            Click_Command("View")
        End If
    End Sub

    Private Sub grdGLTJRNLP_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTJRNLP.InitializeLayout

    End Sub

    Sub Print_Report( _
    Optional ByVal JOURNAL_NO As String = "", _
    Optional ByVal print_directly As Boolean = False)
        Print_Report_Begin()

        If JOURNAL_NO = "" Then
            JOURNAL_NO = HFs("JOURNAL_NO")
        End If

        Fill_Records("GLTACCT1", JOURNAL_NO)
        Fill_Records("GLTPARM2", JOURNAL_NO)

        If Not dst.Tables.Contains("GLTSEGM1") Then
            dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTSEGM1"))
        End If

        Prepare_GL_Account_Activity_Recaps("GLTDETL1")

        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
        CR_params.Add("SHOW_JRNL_COMMENTS", "1")
        CR_params.Add("SHOW_DETL_DESC", "1")
        CR_params.Add("SHOW_CVX_NAME", "0")
        CR_params.Add("PAGE_BREAK", "1")
        CR_params.Add("ACCT_RECAPS", ROWs("GLTPARM1").Item("GL_PARM_ACCT_RECAPS") & "")

        Generate_Report("GLRJRNL1", "Journal Entry", , , , , False)

        If ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
            '  Print_Report_End(True, , "HP Universal Printing PCL 6", , "192.168.130.45:9100")
            '  Print_Report_End(True, , "HP LaserJet P2055DN", , "")
            'HPLaserJetP2055dn
            '  Print_Report_End(True, , "HP LaserJet P2055DN", , "192.168.135.1:9100")
            Print_Report_End(True, , "HP LaserJet P2055DN", , "192.168.135.1:9100")
            '  Print_Report_End(True,  , "Whse 4250", , "192.168.130.45:9100")
        Else
            Print_Report_End(print_directly)
        End If

    End Sub

    Private Sub grdGLTJRNLR_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTJRNLR.DoubleClickRow
        Click_Command("Load to Post")
    End Sub

    Private Sub cmbOPS_YYYYPP_REV_KeyDown(sender As Object, e As KeyEventArgs) Handles cmbOPS_YYYYPP_REV.KeyDown
        If Not ScreenMode Or EntryMode = "N" Then
            If e.KeyCode = Keys.Delete Then
                cmbOPS_YYYYPP_REV.Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub cmbOPS_YYYYPP_REV_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cmbOPS_YYYYPP_REV.KeyPress

    End Sub
End Class