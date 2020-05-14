Imports System.IO
Imports System.Text

Public Class WBFWAYF1
    Dim InquiryOnly As Boolean = False
    Dim TransferFile As String = String.Format("{0}regency.csv", ASCMAIN1.Folders("Temp"))
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Dim SQLs As New System.Text.StringBuilder() With {.Length = 0}
        With dst

            SQLs.Length = 0
            SQLs.AppendLine("SELECT")
            SQLs.AppendLine("'1218' AS Supplier_ID,")
            SQLs.AppendLine("(S1.STYLE_CODE || '-' || S2.COLOR_CODE) AS Item_Number,")
            SQLs.AppendLine("CASE WHEN (NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) < 0 THEN")
            SQLs.AppendLine("  0")
            SQLs.AppendLine("ELSE")
            SQLs.AppendLine("  (NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0))")
            SQLs.AppendLine("END AS On_Hand,")
            SQLs.AppendLine("0 AS Back_Order,")
            SQLs.AppendLine("CASE WHEN (NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) < 0 THEN")
            SQLs.AppendLine("  CASE WHEN ((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) + (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))) < 0 THEN")
            SQLs.AppendLine("    0")
            SQLs.AppendLine("  ELSE")
            SQLs.AppendLine("    ((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) + (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0)))")
            SQLs.AppendLine("  END")
            SQLs.AppendLine("ELSE")
            SQLs.AppendLine("  (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            SQLs.AppendLine("END AS On_Order,")
            SQLs.AppendLine("NULL AS NxtAvailDate,")
            SQLs.AppendLine("DECODE(S1.STYLE_STATUS,'D',1,0) AS STATUS,")
            SQLs.AppendLine("S1.STYLE_DESC AS Description")
            SQLs.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2")
            SQLs.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
            SQLs.AppendLine("AND S2.WHSE_CODE = 'MS'")
            'SQLs.AppendLine("AND S1.STYLE_CODE = 'MT16563'")
            ASCMAIN1.sql = SQLs.ToString()
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 2)
            Fill_Records("ICTSTATX")
        End With

        tab.Visible = False
        'grdSOTORDRX.Parent = tab.Parent

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit"

            Case "Cancel"

            Case "Update"

            Case "Done"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Build File"
                MsgBox("It Is No Longer Required To Run This Program!", MsgBoxStyle.Critical, "No Mass!!")
                Exit Sub
                If My.Computer.FileSystem.FileExists(TransferFile) Then
                    txtProgress.Text = txtProgress.Text & vbCrLf & "Deleting Old File: " & TransferFile
                    My.Computer.FileSystem.DeleteFile(TransferFile)
                End If
                BuildFTPFile()
                txtProgress.Text = txtProgress.Text & vbCrLf & "Transfer File Built: " & TransferFile
                UltraExplorerBar1.Groups("Screen Control").Items("Transfer File").Visible = True
            Case "Transfer File"
                MsgBox("It Is No Longer Required To Run This Program!", MsgBoxStyle.Critical, "No Mass!!")
                Exit Sub
                ftp_File()
                UltraExplorerBar1.Groups("Screen Control").Items("Transfer File").Visible = False
            Case "Cancel"
                txtProgress.Text = vbCrLf & "Process Canceled By User"
                Call Mode_Settings(False)
                UltraExplorerBar1.Groups("Screen Control").Items("Transfer File").Visible = False
            Case "Done"
                txtProgress.Text = vbCrLf & ""
                Call Mode_Settings(False)
                UltraExplorerBar1.Groups("Screen Control").Items("Transfer File").Visible = False
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Edit").Visible = False
                .Groups("Screen Control").Items("Update").Visible = False

                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Cancel").Visible = ScreenMode

                .Groups("Screen Control").Items("Transfer File").Visible = False
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

    End Sub

    Sub Clear_Record()
        dst.Tables("ICTSTYL1").Rows.Clear()
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        'Call Fill_Records("ARTCUST1", Absx1.txtFor("CUST_CODE").Text, True)
        'Call Fill_Records("ARTCUST2", Absx1.txtFor("CUST_CODE").Text, True)

        EnforceConstraints(True)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record(ByVal ORDR_NO As String)
        'Call BeginTrans()
        'For Each TABLE_NAME As String In {"SOTORDR1", "SOTORDR2", "SOTORDR5"}
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        '    ASCMAIN1.sql = "Delete from " & TABLE_NAME & "_L where ORDR_NO = '" & ORDR_NO & "'"
        '    ASCDATA1.ExecuteSQL()
        'Next
        'Call CommitTrans("Order / Quote Deleted")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        'TODO: Remove this stop before going live.
        Stop
        Call CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
            Case "CUST_CODE"
        End Select
    End Sub

    Private Sub Print_Record(ByRef PrintToDefault As Boolean, ByVal ORDR_NO As String, ByVal CUST_CODE As String, Optional ByVal No_Of_Copies As Integer = 1)
        'Print_Report_Begin()
        'Generate_Report("SORORDRO")
        'Print_Report_End()
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSOTORDRX, "SSB", "Show Filter", "Show GroupBox")
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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTORDR1"
            '    If Not InquiryOnly Then
            '        e.Tool.ToolbarsManager.Tools("Edit Ship To").SharedProps.Visible = True
            '    End If
        End Select
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
            'Case "Edit Ship To"
            '    If Not InquiryOnly Then
            '        MsgBox("Edit Ship To Feature Coming Soon", MsgBoxStyle.Exclamation, "Waiting For Feature")
            '    End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "STYLE_CODE"
                'FillStyle()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub BuildFTPFile()
        Dim NewLine As String = ""
        Dim FileForStream As String = "Supplier_ID,Item_Number,On_Hand,Back_Order,On_Order,NxtAvailDate,STATUS,Description" & vbCrLf
        For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select()
            NewLine = ""
            NewLine = String.Format("{0},", rowICTSTATX.Item("Supplier_ID"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("Item_Number"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("On_Hand"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("Back_Order"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("On_Order"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("NxtAvailDate"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("STATUS"))
            NewLine += CleanText(rowICTSTATX.Item("Description").ToString & "") & vbCrLf
            FileForStream += NewLine.ToString
        Next
        Dim fs As FileStream = File.Create(TransferFile)

        Dim info As Byte() = New UTF8Encoding(True).GetBytes(FileForStream)
        fs.Write(info, 0, info.Length)
        fs.Close()
    End Sub

    Public Shared Function CleanText(ByVal TextIn As String) As String
        TextIn = Replace(TextIn, ",", "")
        TextIn = Replace(TextIn, "'", "")
        TextIn = Replace(TextIn, Chr(34).ToString, "")
        Return Trim(TextIn)
    End Function

    Sub ftp_File()

        txtProgress.Text = txtProgress.Text & vbCrLf & "Posting File: " & TransferFile

        Ftp1.User = "EDI_RegencyInternational"
        Ftp1.Password = "Vcbt75EwqaY"
        Ftp1.RemoteHost = "edi.csnstores.com"
        Ftp1.RemotePath = "Inventory"
        Ftp1.Logon()
        Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
        Ftp1.LocalFile = TransferFile
        Ftp1.RemoteFile = "regency.csv"
        Ftp1.Overwrite = True
        Ftp1.Upload()
        Ftp1.Logoff()

        txtProgress.Text = txtProgress.Text & vbCrLf & "File Uploaded to Wayfair"
    End Sub
#End Region
End Class