Public Class WHTUSER1
    Dim rowASTPARMP As DataRow
    Dim SECURITY_CODEsW As New List(Of String)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SECURITY_CODEsW.Add("WH")
        SECURITY_CODEsW.Add("WS")

        With dst
            Create_TDA(.Tables.Add, "ASTUSER1", "*")

            ASCMAIN1.sql = "SELECT ASTUSER2.*, ASTSECM1.SECURITY_DESC, '1' SEL " _
            & " FROM ASTUSER2,ASTSECM1 where ASTSECM1.SECURITY_CODE = ASTUSER2.SECURITY_CODE"

            ASCMAIN1.sql = "Select ASTUSER1.USER_ID " _
                & ", DECODE(ASTUSER2.USER_ID,NULL,'0','1') SEL " _
                & ", ASTSECM1.SECURITY_CODE, ASTSECM1.SECURITY_DESC " _
                & " from ASTSECM1, ASTUSER2, ASTUSER1 " _
                & " where ASTUSER2.USER_ID (+) = :PARM1 " _
                & "   and ASTUSER2.SECURITY_CODE (+) = ASTSECM1.SECURITY_CODE " _
                & "   and ASTSECM1.SECURITY_CODE in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "')" _
                & "   and ASTUSER1.USER_ID = :PARM2"
            Create_TDA(.Tables.Add, "ASTUSER2", "**", 0, True, "VV")

        End With

        grdASTUSER2.DataSource = dst.Tables("ASTUSER2")
        grdASTUSER2.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.Default

        rowASTPARMP = ASCDATA1.GetDataRow("Select * from ASTPARMP where AS_PARM_KEY = 'Z'")

    End Sub

    Private Sub Form_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If SELECTION_NO = 0 Then Exit Sub
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            Case "USER_ID"
                sql_where = "USER_ID in (" _
                    & "(Select Distinct USER_ID from ASTUSER2 where SECURITY_CODE in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "'))" _
                    & " minus " _
                    & "(Select Distinct USER_ID from ASTUSER2 where SECURITY_CODE Not in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "'))" _
                    & ")"
        End Select

    End Sub


#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        Dim sql As String = ""

        sql = "Delete from ASTUSER2 where USER_ID = '" & Absx1.txtFor("USER_ID").Text & "'"
        ASCDATA1.ExecuteSQL(sql)
        dst.Tables("ASTUSER2").AcceptChanges()
        For Each row As DataRow In dst.Tables("ASTUSER2").Rows
            If row.Item("SEL") = "1" Then
                row.SetAdded()
            End If
        Next
        Update_Record_TDA("ASTUSER2")

        Dim rowASTUSER1 As DataRow = Fill_Record("ASTUSER1", Absx1.txtFor("USER_ID").Text)
        If rowASTUSER1 Is Nothing Then
            rowASTUSER1 = dst.Tables("ASTUSER1").NewRow
            rowASTUSER1.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
            dst.Tables("ASTUSER1").Rows.Add(rowASTUSER1)
        End If
        rowASTUSER1.Item("USER_NAME") = rowASFBASE1.Item("USER_NAME")
        rowASTUSER1.Item("USER_PASSWORD") = rowASFBASE1.Item("USER_PASSWORD")
        rowASTUSER1.Item("USER_STATUS") = rowASFBASE1.Item("USER_STATUS")
        Write_Audit_Trail(rowASTUSER1)
        Update_Record_TDA("ASTUSER1")
    End Sub

    Overrides Sub Show_Record_Special()

        EnforceConstraints(False)

        Dim rowASTUSER1 As DataRow = Fill_Record("ASTUSER1", Absx1.txtFor("USER_ID").Text)
        If rowASTUSER1 IsNot Nothing Then
            rowASFBASE1.Item("USER_NAME") = rowASTUSER1.Item("USER_NAME")
            rowASFBASE1.Item("USER_PASSWORD") = rowASTUSER1.Item("USER_PASSWORD")
            rowASFBASE1.Item("USER_STATUS") = rowASTUSER1.Item("USER_STATUS")
        End If


        Fill_Records("ASTUSER2", New String() {Absx1.txtFor("USER_ID").Text, Absx1.txtFor("USER_ID").Text})
        Sort_grdColumns(grdASTUSER2, "SECURITY_CODE")
        If EntryMode = "New" And dst.Tables("ASTUSER2").Rows.Count = 0 Then
            ASCMAIN1.sql = "Select * from ASTSECM1 where SECURITY_CODE in ('" & Join(SECURITY_CODEsW.ToArray, "','") & "') order by SECURITY_CODE"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim rowASTUSER2 As DataRow = dst.Tables("ASTUSER2").NewRow
                rowASTUSER2.Item("USER_ID") = Absx1.txtFor("USER_ID").Text
                rowASTUSER2.Item("SEL") = "0"
                rowASTUSER2.Item("SECURITY_CODE") = row.Item("SECURITY_CODE")
                rowASTUSER2.Item("SECURITY_DESC") = row.Item("SECURITY_DESC")
                dst.Tables("ASTUSER2").Rows.Add(rowASTUSER2)
            Next
        End If


        EnforceConstraints(True)
    End Sub

    Sub Load_Report_Form(ByVal FORM_NAME As String)

    End Sub

    Overrides Sub Clear_Record_Special()
        'If ScreenMode Then
        '    Fill_ASTUSERX()
        'End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        grdASTUSER2.Visible = tf

        If tf Then
            If EntryMode = "New" Or EntryMode = "Edit" Then
                grdASTUSER2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Else
                grdASTUSER2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If
        End If

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                Dim USER_ID As String = Absx1.txtFor("USER_ID").Text

                If USER_ID.Length > 0 Then
                    If USER_ID <> USER_ID.ToLower Then
                        EMsg &= vbCr & "User ID should use lowercase letters only"
                    Else
                        For i As Int16 = 1 To USER_ID.Length
                            Dim z As String = USER_ID.Substring(i - 1, 1)
                            If z < "a" Or z > "z" Then
                                If InStr("0123456789.", z) = 0 Then
                                    EMsg &= vbCr & "User ID should use lowercase letters and numbers only"
                                End If
                            End If
                        Next
                    End If
                End If
                ' find all SECURITY_CODEs for an existing user
                Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", USER_ID)
                If rowASTUSER1 IsNot Nothing Then
                    EMsg &= vbCr & "User (" & USER_ID & ") already exists, try edit."
                End If

            Case "Edit"
                Dim USER_ID As String = Absx1.txtFor("USER_ID").Text

                If USER_ID.Length > 0 Then
                    If USER_ID <> USER_ID.ToLower Then
                        EMsg &= vbCr & "User ID should use lowercase letters only"
                    Else
                        For i As Int16 = 1 To USER_ID.Length
                            Dim z As String = USER_ID.Substring(i - 1, 1)
                            If z < "a" Or z > "z" Then
                                If InStr("0123456789.", z) = 0 Then
                                    EMsg &= vbCr & "User ID should use lowercase letters and numbers only"
                                End If
                            End If
                        Next
                    End If
                End If

                ' find all SECURITY_CODEs for an existing user
                Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", USER_ID)
                If rowASTUSER1 IsNot Nothing Then
                    Dim SECURITY_CODEs As New List(Of String)
                    ASCMAIN1.sql = "Select SECURITY_CODE from ASTUSER2 where USER_ID = :PARM1"
                    For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {USER_ID}).Select("")
                        Dim SECURITY_CODE As String = row.Item("SECURITY_CODE")
                        If SECURITY_CODEsW.Contains(SECURITY_CODE) Then
                        Else
                            SECURITY_CODEs.Add(SECURITY_CODE)
                        End If
                    Next

                    If SECURITY_CODEs.Count <> 0 Then
                        EMsg &= vbCr & "Cannot Maintain this User (" & USER_ID & ")"
                    End If
                End If


            Case "Update"

                Dim password_error_checks As String = _
                ASCMAIN1.Validate_User_Password( _
                False, _
                Absx1.txtFor("USER_ID").Text, _
                Absx1.txtFor("USER_PASSWORD").Text, _
                rowASTPARMP)

                If password_error_checks <> "" Then
                    EMsg &= vbCr & "Password Errors:" & vbCr & vbTab & Replace(password_error_checks, vbCr, vbCr & vbTab)
                End If

        End Select

    End Sub
#End Region
End Class