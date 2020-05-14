Imports System.Drawing
Imports System.Text

Public Class WBTSTYLW
    Private WB_PARM_IMAGES_DIR As String
    Private WB_PARM_IMAGES_UPLOADED_DIR As String
    Private STYLE_CODE As String = String.Empty
    Private WB_PARM_MAX_REC_STYLE As Int16 = 0
    Private CurrentImage As String = String.Empty

    Private Sub WBTSTYLW_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SB As New StringBuilder
        With dst
            Get_PARM("WBTPARM1")
            WB_PARM_IMAGES_DIR = (ROWs("WBTPARM1").Item("WB_PARM_IMAGES_DIR") & String.Empty).ToString.Trim
            WB_PARM_IMAGES_UPLOADED_DIR = String.Empty

            If WB_PARM_IMAGES_DIR.Length > 0 Then
                If Not WB_PARM_IMAGES_DIR.EndsWith("\") Then WB_PARM_IMAGES_DIR &= "\"
                WB_PARM_IMAGES_UPLOADED_DIR = WB_PARM_IMAGES_DIR & "Uploaded\"
                If Not My.Computer.FileSystem.DirectoryExists(WB_PARM_IMAGES_DIR) Then
                    WB_PARM_IMAGES_DIR = String.Empty
                End If
            End If

            ASCMAIN1.sql = "Select '0' SELECTED, WBTPAGE1.* FROM WBTPAGE1 WHERE NVL(PAGE_STATUS, 'A') = 'A'"
            Create_TDA(.Tables.Add, "WBTPAGE1", "**", 0, False, "", 2)

            Create_TDA(.Tables.Add, "WBTSTYL3", "*")

        End With

        grdWBTSTYL3.DataSource = dst.Tables("WBTPAGE1")

        Create_Lookup("WBTPAGE1")
        Create_Lookup("WBTSTYL1")
        Create_Lookup("ICTCOLR1")
        Create_Lookup("ICTSIZE1")

        Get_PARM("WBTPARM1")
        If ROWs("WBTPARM1") IsNot Nothing Then
            WB_PARM_MAX_REC_STYLE = Val(ROWs("WBTPARM1").Item("WB_PARM_MAX_REC_STYLE") & String.Empty)
        End If

    End Sub

    Private Sub txt_DisplayImage(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Dim ImageFile As String = txtctl.Text.Trim
        If ImageFile.Length = 0 Then
            Exit Sub
        End If

        CurrentImage = ImageFile
    End Sub

#Region "Overrides"

    Public Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Select Case eItemKey

            Case "Update"

                MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text.Replace(vbCrLf, Space(1)).Trim
                MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text.Replace(Space(2), Space(1)).Trim

                dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_FULL_DESC") = MyBase.Absx1.txtFor("STYLE_FULL_DESC").Text

                Dim wMsg As String = String.Empty

                If wMsg.Length > 0 Then
                    If MessageBox.Show(wMsg & vbCr & vbCr & "Update anyway?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.No Then
                        EMsg = "Update cancelled by user."
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        MyBase.Proceed_PreReq(eItemKey)

        MyBase.Absx1.txtFor("STYLE_CODE").Text = MyBase.Absx1.txtFor("STYLE_CODE").Text.ToUpper.Trim

    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        dst.Tables("WBTSTYL3").Clear()
        dst.Tables("WBTSTYL3").AcceptChanges()

        Dim rowWBTSTYL3 As DataRow = Nothing
        For Each rowWBTPAGE1 As DataRow In dst.Tables("WBTPAGE1").Select("SELECTED = '1'")
            rowWBTSTYL3 = dst.Tables("WBTSTYL3").NewRow
            rowWBTSTYL3.Item("STYLE_CODE") = Absx1.txtFor("STYLE_CODE").Text
            rowWBTSTYL3.Item("PAGE_CODE") = rowWBTPAGE1.Item("PAGE_CODE")
            dst.Tables("WBTSTYL3").Rows.Add(rowWBTSTYL3)
        Next

        Dim seqNo As Integer = 1
        For Each rowICTSTYLR As DataRow In dst.Tables("ICTSTYLR").Select("", "SEQ_NO", DataViewRowState.CurrentRows)
            rowICTSTYLR.Item("SEQ_NO") = seqNo
            seqNo += 1
        Next

        Update_Record_TDA("ICTSTYLR", "DELETE FROM ICTSTYLR WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        Update_Record_TDA("WBTSTYL3", "DELETE FROM WBTSTYL3 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
    End Sub

    Public Overrides Sub Proceed_Update_Special_Post()
        MyBase.Proceed_Update_Special_Post()
        ASCDATA1.ExecuteSQL("UPDATE WBTSTYL1 SET WEB_IND = '1' WHERE STYLE_CODE = :PARM1", "V", STYLE_CODE)
    End Sub

    Public Overrides Sub txt_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs)
        MyBase.txt_EditorButtonClick(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor
        txtctl = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Dim imageFields As String = ",STYLE_IMAGE,STYLE_IMAGE_OTHER1,STYLE_IMAGE_OTHER2, "

        If imageFields.Contains("," & MyBase.Absx1.GetABSColumnName(txtctl) & ",") Then

            Select Case e.Button.Key

                Case "Open"
                    Using fdlg As OpenFileDialog = New OpenFileDialog()
                        fdlg.Title = "Images Open File Dialog"
                        fdlg.InitialDirectory = WB_PARM_IMAGES_DIR
                        fdlg.Filter = "Image Files|*.jpg;*.gif;*.bmp;*.png;*.jpeg|All Files|*.*"
                        fdlg.FilterIndex = 1
                        fdlg.RestoreDirectory = True
                        If fdlg.ShowDialog() = DialogResult.OK Then
                            If My.Computer.FileSystem.FileExists(fdlg.FileName) Then
                                txtctl.Text = My.Computer.FileSystem.GetName(fdlg.FileName)
                            End If
                        End If
                    End Using

                Case "View"
                    txt_DisplayImage(txtctl)

            End Select
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        STYLE_CODE = MyBase.Absx1.txtFor("STYLE_CODE").Text.Trim

        Dim sql As String = String.Empty

        MyBase.EnforceConstraints(False)

        Call Fill_Records("WBTSTYL3", String.Empty, True, "SELECT * FROM WBTSTYL3 WHERE STYLE_CODE = '" & STYLE_CODE & "'")
        Call Fill_Records("ICTSTYLR", STYLE_CODE, )

        Call Fill_Records("WBTPAGE1")

        For Each rowWBTSTYL3 As DataRow In dst.Tables("WBTSTYL3").Rows
            If dst.Tables("WBTPAGE1").Select("PAGE_CODE = '" & rowWBTSTYL3.Item("PAGE_CODE") & "'").Length > 0 Then
                dst.Tables("WBTPAGE1").Select("PAGE_CODE = '" & rowWBTSTYL3.Item("PAGE_CODE") & "'")(0).Item("SELECTED") = 1
            End If
        Next
        dst.Tables("WBTSTYL3").AcceptChanges()

        MyBase.EnforceConstraints(True)

        With grdWBTSTYL3.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("SELECTED", True)

        End With

        If EntryMode = "New" Then
            dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_IMAGE_OTHER1") = String.Empty
            dst.Tables("WBTSTYL1").Rows(0).Item("STYLE_IMAGE_OTHER2") = String.Empty
        End If

    End Sub

    Overrides Sub Clear_Record_Special()

        If ScreenMode Then
            MyBase.EnforceConstraints(False)
            dst.Tables("WBTSTYL3").Rows.Clear()
            dst.Tables("ICTSTYLR").Rows.Clear()
            dst.Tables("WBTPAGE1").Rows.Clear()
            MyBase.EnforceConstraints(True)

            STYLE_CODE = String.Empty
            tabStyle.SelectedTab = tabStyle.Tabs(0)

            CurrentImage = String.Empty
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        tabStyle.Visible = tf

    End Sub

#End Region

#Region "Form Controls"
    Private Sub grdWBTSTYL3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWBTSTYL3.InitializeRow

        Dim pageCode As String = e.Row.Cells("PAGE_CODE").Value

        Try
            e.Row.Cells("PAGE_NAME").Value = MyBase.LookUp("WBTPAGE1", pageCode, True).Item("PAGE_NAME") & String.Empty
        Catch ex As Exception

        End Try

    End Sub

#End Region

    Private Sub UltraTextEditor1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles UltraTextEditor1.KeyUp
        If Not ScreenMode Then
            Dim code As String = Absx1.txtFor("STYLE_CODE").Text.ToUpper
            Absx1.txtFor("STYLE_DESC").Text = Lookup("WBTSTYL1", code, True).Item("STYLE_DESC") & String.Empty
        End If
    End Sub
End Class