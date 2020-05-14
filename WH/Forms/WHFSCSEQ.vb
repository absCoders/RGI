Public Class WHFSCSEQ


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from WHTSCSEQ"
            Create_TDA(.Tables.Add, "WHTSCSEQ", ASCMAIN1.sql, 0, True, 3)
        End With

        '   grdWHTSCSEQ.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdWHTSCSEQ.DataSource = dst.Tables("WHTSCSEQ")
        Create_Summary(grdWHTSCSEQ, "CUST_CODE", "Count")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
            Case "Load from XLS"
                Dim openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True
                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    Dim FILENAME As String = openFileDialog1.FileName
                    ImportSEQ(FILENAME)
                End If

            Case "Update"
            Case "Cancel"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Load from XLS"


            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load from XLS").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

            End With


        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If
    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        dst.Tables("WHTSCSEQ").Rows.Clear()

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)

        Fill_Records("WHTSCSEQ")
        Sort_grdColumns(grdWHTSCSEQ, "CUST_CODE,STYLE_SEQ")

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            ASCMAIN1.sql = "Delete from WHTSCSEQ"
            ASCDATA1.ExecuteSQL()

            Update_Record_TDA("WHTSCSEQ", "1=1")

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTSCSEQ, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

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

        Select Case e.Tool.Key


        End Select
    End Sub

#End Region


    Sub ImportSEQ(fileName As String)

        If fileName <> "" Then
            Dim eMsg As String = ""
            ASCMAIN1.Progress("Now Loading XLS")

            Try
                Dim oWB As SpreadsheetGear.IWorkbook
                oWB = SpreadsheetGear.Factory.GetWorkbook(fileName)
                Dim ws As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)

                dst.Tables("WHTSCSEQ").Rows.Clear()

                For r As Int64 = 1 To ws.UsedRange.RowCount - 1
                    Dim CUST_CODE As String = ws.Cells(r, 0).Text
                    Dim STYLE_CODE As String = ws.Cells(r, 1).Text
                    Dim COLOR_CODE As String = ws.Cells(r, 2).Text
                    Dim STYLE_SEQ As Integer = Val(ws.Cells(r, 3).Text)

                    Dim row As DataRow = dst.Tables("WHTSCSEQ").NewRow
                    row.Item("CUST_CODE") = CUST_CODE
                    row.Item("STYLE_CODE") = STYLE_CODE
                    row.Item("COLOR_CODE") = COLOR_CODE
                    row.Item("STYLE_SEQ") = STYLE_SEQ
                    dst.Tables("WHTSCSEQ").Rows.Add(row)

                Next

            Catch ex As Exception
                MsgBox(ex.Message, vbOKOnly, "Error Occurred")
            End Try

        End If

        ASCMAIN1.Progress("")

    End Sub
End Class