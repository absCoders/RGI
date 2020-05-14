Public Class SOFINVHM

    Public tbl As DataTable
    Public updated As Boolean = False

#Region "Form Events"
    Private Sub TAFLOCM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        'ASCMAIN1.grdInitializeLayout(grdWHTMOVE2)

        With dst
            Create_TDA(.Tables.Add, "SOTINVH1", "*", , , , , "SREP_CODE,SREP2_CODE,TERM_CODE")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*", , , , , "SREP_CODE,SREP2_CODE,TERM_CODE,INV_DUE_DATE")

            ASCMAIN1.sql = "Select * from ARTOPENX where CUST_CODE = :PARM1 and INV_TYPE = :PARM2 and INV_NUM = :PARM3"
            Create_TDA(.Tables.Add, "ARTOPENX", "**", 0, , "VVV", 3, "SREP_CODE,SREP2_CODE,TERM_CODE")

            ' Create_TDA(.Tables.Add, "ASTAUDT1", "*")
            Create_TDA(.Tables.Add, "SOTORDR1", "*", , , , , "SREP_CODE,SREP2_CODE,TERM_CODE")
        End With

        For Each row As DataRow In tbl.Rows
            Dim INV_TYPE As String = row.Item("INV_TYPE")
            Dim INV_NO As String = row.Item("INV_NO")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim rowSOTINVH1 As DataRow = Fill_Record("SOTINVH1", New String() {INV_TYPE, INV_NO}, , False)
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & ""
            Fill_Record("SOTORDR1", New String() {ORDR_NO}, , False)
            Fill_Record("ARTOPEN1", New String() {CUST_CODE, INV_TYPE, INV_NO}, , False)
            Fill_Record("ARTOPENX", New String() {CUST_CODE, INV_TYPE, INV_NO}, , False)
        Next

        grdSOTINVH1.DataSource = dst.Tables("SOTINVH1")

        With grdSOTINVH1.DisplayLayout.Bands(0)
            .Columns("INV_TYPE").Header.Fixed = True
            .Columns("INV_NO").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("INV_DATE").Header.Fixed = True
        End With

        Create_Summary(grdSOTINVH1, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1, "INV_TOTAL_AMOUNT")
    End Sub

    Private Sub TAFLOCM1_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown
        Sort_grdColumns(grdSOTINVH1, "INV_TYPE,INV_NO")

        If grdSOTINVH1.ActiveRow IsNot Nothing Then
            Absx1.txtFor("SREP_CODE").Text = grdSOTINVH1.ActiveRow.Cells("SREP_CODE").Value & ""
            Absx1.txtFor("SREP2_CODE").Text = grdSOTINVH1.ActiveRow.Cells("SREP2_CODE").Value & ""
            Absx1.txtFor("TERM_CODE").Text = grdSOTINVH1.ActiveRow.Cells("TERM_CODE").Value & ""
        End If
    End Sub

#End Region

#Region "Form Prodecures"
     

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As System.Windows.Forms.Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            'Case "LOCATION_CODE"
            '    sql_where = "WHSE_CODE = '" & WHSE_CODE & "'"
        End Select
    End Sub

#End Region

#Region "Buttons"

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click

        Me.Close()
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click

        EMsg = String.Empty
        DATETIME_STAMP = DateTime.Now + ASCMAIN1.NowTSD

        If Absx1.txtFor("SREP_CODE").Text = "" Then
            EMsg &= vbCr & "A Value for Sales Rep is Mandatory"
        Else
            If LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text) Is Nothing Then
                EMsg &= vbCr & "Value Specified for New Sales Rep is Invalid"
            End If
        End If

        If Absx1.txtFor("SREP2_CODE").Text = "" Then
            ' EMsg &= vbCr & "A Value for Sales Rep 2 is Mandatory"
        Else
            If LookUp("SOTSREP1", Absx1.txtFor("SREP2_CODE").Text) Is Nothing Then
                EMsg &= vbCr & "Value Specified for New Sales Rep 2 is Invalid"
            End If
        End If

        If Absx1.txtFor("TERM_CODE").Text = "" Then
            EMsg &= vbCr & "A Value for Terms Code is Mandatory"
        Else
            If LookUp("TATTERM1", Absx1.txtFor("TERM_CODE").Text) Is Nothing Then
                EMsg &= vbCr & "Value Specified for New Terms Code is Invalid"
            End If
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Peform Update")
            Exit Sub
        End If


        For Each row As DataRow In tbl.Select("")
            row.Item("SREP_CODE") = Absx1.txtFor("SREP_CODE").Text
            row.Item("SREP2_CODE") = Absx1.txtFor("SREP2_CODE").Text
            row.Item("TERM_CODE") = Absx1.txtFor("TERM_CODE").Text
        Next

        Dim ORDR_GROUP_NOs As New List(Of String)

        For Each TABLE_NAME As String In New String() {"ARTOPEN1", "ARTOPENX", "SOTINVH1", "SOTORDR1"}
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                Dim KEY_VALUE As String = ""
                If TABLE_NAME = "SOTINVH1" Then
                    KEY_VALUE = row.Item("INV_TYPE") & ":" & row.Item("INV_NO")
                ElseIf TABLE_NAME = "SOTORDR1" Then
                    KEY_VALUE = row.Item("ORDR_NO")
                    Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                    If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                    End If
                Else
                    KEY_VALUE = row.Item("CUST_CODE") & ":" & row.Item("INV_TYPE") & ":" & row.Item("INV_NUM")
                End If

                row.Item("SREP_CODE") = Absx1.txtFor("SREP_CODE").Text
                row.Item("SREP2_CODE") = Absx1.txtFor("SREP2_CODE").Text
                row.Item("TERM_CODE") = Absx1.txtFor("TERM_CODE").Text

                Audit_row(row, TABLE_NAME, "SREP_CODE", KEY_VALUE)
                Audit_row(row, TABLE_NAME, "SREP2_CODE", KEY_VALUE)
                Audit_row(row, TABLE_NAME, "TERM_CODE", KEY_VALUE)

                If TABLE_NAME = "SOTINVH1" Or TABLE_NAME = "SOTORDR1" Then
                Else
                    Dim INV_DATE As Date = row.Item("INV_DATE")
                    Dim TERM_CODE As String = row.Item("TERM_CODE")
                    Dim INV_DUE_DATE As Date = TAC.TACMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, Nothing, INV_DATE)
                    row.Item("INV_DUE_DATE") = INV_DUE_DATE
                    Audit_row(row, TABLE_NAME, "INV_DUE_DATE", KEY_VALUE)
                End If
            Next
        Next

        Try
            BeginTrans()

            Update_Record_TDA("SOTINVH1")
            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("ARTOPEN1")
            Update_Record_TDA("ARTOPENX")
            Update_Record_TDA("ASTAUDT1")

            For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
            Next

            For Each row As DataRow In tbl.Select("")
                ASCMAIN1.sql = "Update SOTINVHS Set SREP_CODE = :PARM1 where INV_TYPE = :PARM2 and INV_NO = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", New Object() {Absx1.txtFor("SREP_CODE").Text, row.Item("INV_TYPE"), row.Item("INV_NO")})
            Next

            CommitTrans("Update Successful")

            updated = True

        Catch ex As Exception
            Rollback(ex.Message)

        End Try

        Me.Close()
    End Sub

    Sub Audit_row(row As DataRow, TABLE_NAME As String, COLUMN_NAME As String, KEY_VALUE As String)
        If row.Item(COLUMN_NAME) & "" <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
            Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
            With rowASTAUDT1
                .Item("TABLE_NAME") = TABLE_NAME
                .Item("KEY_VALUE") = KEY_VALUE
                .Item("COLUMN_NAME") = COLUMN_NAME
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original) & ""
                .Item("NEW_VALUE") = row.Item(COLUMN_NAME) & ""
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("SELECTION_NO") = Me.SELECTION_NO
                .Item("XNO") = Me.XNO
            End With
            dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
        End If
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "LOCATION_CODE"
            '    ValidateSelectedNewLocation()

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "LOCATION_CODE"
            '    ValidateSelectedNewLocation()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "LOCATION_CODE"
            '    ValidateSelectedNewLocation()
        End Select
    End Sub
 
#End Region

End Class