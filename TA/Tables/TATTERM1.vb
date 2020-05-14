Public Class TATTERM1

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grpTester.Visible = tf

        btnFix.Visible = tf
        If tf And ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            btnFix.Visible = True
        End If
    End Sub

    Private Sub optTERM_DUE_TYPE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTERM_DUE_TYPE.ValueChanged
        Setup_EOM()
    End Sub

    Sub Setup_EOM()
        If optTERM_DUE_TYPE.Value & "" = "E" Then
            lblTERM_DAYS_DUE.Text = "Day of Month Due"
            numTERM_DAYS_DUE.MinValue = 1
            'If numTERM_DAYS_DUE.Value = 0 Then
            '    numTERM_DAYS_DUE.Value = 1
            'End If
        ElseIf optTERM_DUE_TYPE.Value & "" = "D" Then
            lblTERM_DAYS_DUE.Text = "Days til Net Due"
            numTERM_DAYS_DUE.MinValue = 0
        ElseIf optTERM_DUE_TYPE.Value & "" = "S" Then
            numTERM_DAYS_DUE.MinValue = 0
        End If
        lblTERM_DAYS_DUE.Visible = Not (optTERM_DUE_TYPE.Value & "" = "S")
        dteTERM_CUTOFF_DATE.Visible = (optTERM_DUE_TYPE.Value & "" = "S")
        numTERM_DAYS_DUE.Visible = Not (optTERM_DUE_TYPE.Value & "" = "S")
        grpE.Visible = (optTERM_DUE_TYPE.Value & "" = "E")
        Absx1.numFor("TERM_CUTOFF_DAY").Visible = (optTERM_EOM_TYPE.Value = "S")
    End Sub
    Private Sub dteInvoiceDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteInvoiceDate.ValueChanged
        If dteInvoiceDate.Value & "" <> "" Then
            CALC_DUE_DATE()
            ' btnTest.PerformClick()
        End If

    End Sub

    Private Sub optTERM_EOM_TYPE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTERM_EOM_TYPE.ValueChanged
        Absx1.numFor("TERM_CUTOFF_DAY").Visible = (optTERM_EOM_TYPE.Value = "S")
    End Sub

    Sub CALC_DUE_DATE()
        If Absx1.optFor("TERM_DUE_TYPE").Value & "" = "" Then
            MsgBox("Aging Parameters Type Not Defined", MsgBoxStyle.OkOnly, "Cannot Calculate a Due Date")
            Exit Sub
        End If

        If optTERM_DUE_TYPE.Value & "" = "E" And Val(numTERM_DAYS_DUE.Value & "") = 0 Then
            MsgBox("Invalid Day of the Month", MsgBoxStyle.OkOnly, "Cannot Calculate a Due Date")
            Exit Sub
        End If

        Dim INV_DUE_DATE As Object = Nothing
        Dim INV_BASE_DATE As Object = dteInvoiceDate.Value
        dteDiscDueDate.Value = Null
        If INV_BASE_DATE Is Nothing Then
            dteDueDate.Value = Null
            Exit Sub
        End If

        Synch_TABLE_NAME("TATTERM1")
        Dim row As DataRow = dst.Tables("TATTERM1").Rows.Find(Absx1.txtFor("TERM_CODE").Text)
        dteDueDate.Value = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, Absx1.txtFor("TERM_CODE").Text, row, dteInvoiceDate.Value)

        If Val(Absx1.numFor("TERM_DISC_PERC").Value & "") <> 0 Then
            If Absx1.chkFor("TERM_DISC_ELIG_DUE").Checked Then
                dteDiscDueDate.Value = dteDueDate.Value
            Else
                If Val(Absx1.numFor("TERM_DISC_PERC").Value & "") <> 0 Then
                    dteDiscDueDate.Value = DateValue(dteDueDate.Value & "").AddDays(Val(Absx1.numFor("TERM_DAYS_DISC").Value & ""))
                End If
            End If
        End If
    End Sub

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
        CALC_DUE_DATE()
    End Sub
     

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Update"
                If optTERM_DUE_TYPE.CheckedIndex = -1 Then
                    EMsg &= vbCr & "Invalid Terms Type"
                End If

                If Absx1.txtFor("TERM_DESC").Text = "" Then
                    EMsg &= vbCr & "Terms Description Required"
                End If

                If Absx1.optFor("TERM_USE").CheckedIndex = -1 Then
                    EMsg &= vbCr & "Terms Use Required"
                End If

                If Absx1.optFor("TERM_STATUS").CheckedIndex = -1 Then
                    EMsg &= vbCr & "Terms Status Required"
                End If

                If optTERM_DUE_TYPE.Value & "" = "E" Then
                    If Val(numTERM_DAYS_DUE.Value & "") = 0 Then
                        EMsg &= vbCr & "Invalid Day of Month Due (" & numTERM_DAYS_DUE.Value & ")"
                    End If
                    If optTERM_EOM_TYPE.CheckedIndex = -1 Then
                        EMsg &= vbCr & "Invalid EOM-Specific Parameters"
                    End If

                End If
        End Select
    End Sub

    Overrides Sub Show_Record_Special()
        Setup_EOM()
    End Sub

    Private Sub btnFix_Click(sender As Object, e As EventArgs) Handles btnFix.Click

        Synch_TABLE_NAME("TATTERM1")
        Dim rowTATTERM1 As DataRow = dst.Tables("TATTERM1").Rows.Find(Absx1.txtFor("TERM_CODE").Text)
        Dim TERM_CODE As String = Absx1.txtFor("TERM_CODE").Text
        ASCMAIN1.sql = "Select Distinct INV_DATE from ARTOPEN1 where TERM_CODE = '" & TERM_CODE & "'"
        Dim R As Integer = 0
        Dim F As String = ""
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim INV_DATE As Date = row.Item("INV_DATE")
            Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, rowTATTERM1, INV_DATE)
            ASCMAIN1.sql = "Update ARTOPEN1 Set INV_DUE_DATE = :PARM1 where TERM_CODE = :PARM2 and INV_DATE = :PARM3"
            Dim Rs As Integer = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVD", New Object() {INV_DUE_DATE, TERM_CODE, INV_DATE})
            F &= vbCrLf & Format(INV_DATE, "MM/dd/yyyy") & " -> " & Format(INV_DUE_DATE, "MM/dd/yyyy") & ": " & CStr(Rs) & " records"
            R += Rs
        Next
        MsgBox(Mid(F, 3), MsgBoxStyle.OkOnly, "Due Dates corrected for " & CStr(R) & " Records")
    End Sub
End Class