Imports System.Text

Public Class WBFVANW1
    Dim InquiryOnly As Boolean = False
    Dim SQL As New System.Text.StringBuilder() With {.Length = 0}
    Dim PartialLocations As String = "Z:\Wayne On My Mac\Dropbox\Clients\Vandale\WebPartials\"
    Dim ImageLocations As String = "Z:\Wayne On My Mac\Dropbox\Clients\Vandale\WebPartials\"
#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst
            'SQL.Length = 0
            'SQL.AppendLine("SELECT * FROM TABLE_NAME WHERE COLUMN_NAME = :PARM1")
            'ASCMAIN1.sql = SQL.ToString()
            'Create_TDA(.Tables.Add, "TABLE_NASME", "**", 0, True, "V", 1)
        End With

        Dim CareersLocation As String = PartialLocations & "Careers_partial.cshtml"
        tclCareers.Load(CareersLocation, TXTextControl.StreamType.HTMLFormat)

        Dim News1Location As String = PartialLocations & "news1.cshtml"
        tclNews1.Load(News1Location, TXTextControl.StreamType.HTMLFormat)

        Dim News2Location As String = PartialLocations & "news2.cshtml"
        tclNews2.Load(News2Location, TXTextControl.StreamType.HTMLFormat)

        Dim News3Location As String = PartialLocations & "news3.cshtml"
        tclNews3.Load(News3Location, TXTextControl.StreamType.HTMLFormat)

        pcHero1.ImageLocation = ImageLocations & "hero1.jpg"

        pcHero2.ImageLocation = ImageLocations & "hero2.jpg"

        tab.Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"
            Case "Done"
                Mode_Settings(False)
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)
            Case "Done"
                Call Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        Call Set_ScreenMode_Base(tf)
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Done").Visible = Not ScreenMode
                .Groups("Screen Control").Items("Update").Visible = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If Not tf Then
            Call Clear_Record()
        End If
    End Sub

    Sub Clear_Record()
        'dst.Tables("TABLE_NAME").Rows.Clear()
        'txtPART_HTLM.Text = ""
    End Sub

    Sub Load_Record()
        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

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
        'Call BeginTrans()
        'Update_Record_TDA("TABLE_NAME")
        'Call CommitTrans("Update Complete")
        Dim DTHMS As String = "_" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second

        Dim CareersLocation As String = PartialLocations & "Careers_partial.cshtml"
        Dim CareersCopy As String = PartialLocations & String.Format("Careers_partial{0}.cshtml", DTHMS)
        IO.File.Copy(CareersLocation, CareersCopy)
        tclCareers.Save(CareersLocation, TXTextControl.StreamType.HTMLFormat)

        Dim News1Location As String = PartialLocations & "news1.cshtml"
        Dim News1Copy As String = PartialLocations & String.Format("news1{0}.cshtml", DTHMS)
        IO.File.Copy(News1Location, News1Copy)
        tclNews1.Save(News1Location, TXTextControl.StreamType.HTMLFormat)

        Dim News2Location As String = PartialLocations & "news2.cshtml"
        Dim News2Copy As String = PartialLocations & String.Format("news2{0}.cshtml", DTHMS)
        IO.File.Copy(News2Location, News2Copy)
        tclNews2.Save(News2Location, TXTextControl.StreamType.HTMLFormat)

        Dim News3Location As String = PartialLocations & "news3.cshtml"
        Dim News3Copy As String = PartialLocations & String.Format("news3{0}.cshtml", DTHMS)
        IO.File.Copy(News3Location, News3Copy)
        tclNews3.Save(News3Location, TXTextControl.StreamType.HTMLFormat)


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

    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)

    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        If IsLoading Then
            Exit Sub
        Else
            MyBase.txt_ValueChanged(sender, e)
        End If
    End Sub

    Private Sub btnHero1_Click(sender As Object, e As EventArgs) Handles btnHero1.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        Dim DTHMS As String = "_" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second

        fd.Title = "Select Image To Replace."
        fd.Filter = "All files (*.jpg)|*.jpg"

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            Dim picHero1Copy As String = String.Format(ImageLocations & "hero1{0}.jpg", DTHMS)
            IO.File.Copy(ImageLocations & "hero1.jpg", picHero1Copy)
            IO.File.Delete(ImageLocations & "hero1.jpg")
            IO.File.Copy(fd.FileName, ImageLocations & "hero1.jpg")
            pcHero1.ImageLocation = ImageLocations & "hero1.jpg"
        End If
    End Sub

    Private Sub btnHero2_Click(sender As Object, e As EventArgs) Handles btnHero2.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()

        Dim DTHMS As String = "_" & Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second

        fd.Title = "Select Image To Replace."
        fd.Filter = "All files (*.jpg)|*.jpg"

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            Dim picHero2Copy As String = String.Format(ImageLocations & "hero2{0}.jpg", DTHMS)
            IO.File.Copy(ImageLocations & "hero2.jpg", picHero2Copy)
            IO.File.Delete(ImageLocations & "hero2.jpg")
            IO.File.Copy(fd.FileName, ImageLocations & "hero2.jpg")
            pcHero2.ImageLocation = ImageLocations & "hero2.jpg"
        End If
    End Sub

#End Region

#Region "Custom Methods"

#End Region

End Class