Imports Infragistics.Win.UltraWinGrid

Public Class TAFPMGMM

#Region "General Declarations"
    Public PROGRAM_NO As String = String.Empty
    Public CONV_TOPIC_NO As String = ""
    Public CONV_NO As String = String.Empty
    Public CONV_SUBJECT As String
    Public CONV_NOTES As String
    Public PROGRAM_CATGY_CODE As String
    Public rowPOTPMGM1 As DataRow
    Public rowTATCONV1 As DataRow
    Public rowTATCONV1_PREV As DataRow
    Public frmASFBASE0 As ASFBASE0
    Public use_dst As Boolean = False
    Public CONV_NO_PREV As String = String.Empty
    Public CONV_NO_PREV_NOTES As String = ""
    Public MESSAGE_BY As String = ""

    ' Public EntryMode As String = ""
    Dim reply_color As System.Drawing.Color = Nothing
    Dim reply_in_line As Boolean = False

    Public result As String = ""
#End Region

    Public Sub New(ByVal PROGRAM_NO As String, rowPOTPMGM1 As DataRow, frmASFBASE0 As ASFBASE0, Optional use_dst As Boolean = False)
        Me.PROGRAM_NO = PROGRAM_NO
        Me.frmASFBASE0 = frmASFBASE0
        Me.use_dst = use_dst

        If use_dst Then
            rowPOTPMGM1 = frmASFBASE0.dst.Tables("POTPMGM1").Rows(0)
            Me.PROGRAM_NO = rowPOTPMGM1.Item("PROGRAM_NO")
        End If
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")

        With dst
            Create_TDA(.Tables.Add, "POTPMGM1", "*", , False)
            Create_TDA(.Tables.Add, "POTPMGM2", "*", 1, False)


            ASCMAIN1.sql = "Select * from POTCTOP1" & vbCrLf _
                & " where POTCTOP1.PROGRAM_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTCTOP1", "**", 0, True, "V", 0)

            ASCMAIN1.sql = "Select NVL(NVL(TATCONPU.CONV_NO,:PARM1),'0000000000') CONV_NO" & vbCrLf _
                & ", POTPMGMO.USER_ID, ASTUSER1.USER_EMAIL from TATCONPU, POTPMGMO, ASTUSER1" & vbCrLf _
                & " where TATCONPU.CONV_NO (+) = :PARM1" & vbCrLf _
                & "   and TATCONPU.USER_ID (+) = POTPMGMO.USER_ID" & vbCrLf _
                & "   and ASTUSER1.USER_ID = POTPMGMO.USER_ID"
            Create_TDA(.Tables.Add, "TATCONPU", "**", 0, True, "V", 0)
            .Tables("TATCONPU").Columns.Add("SEL")
            .Tables("TATCONPU").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select NVL(NVL(TATCONPS.CONV_NO,:PARM1),'0000000000') CONV_NO" & vbCrLf _
                & ", POTPMGM2.STYLE_CODE_PLM, POTPMGM2.STYLE_NAME from TATCONPS, POTPMGM2" & vbCrLf _
                & " where TATCONPS.CONV_NO (+) = :PARM1" & vbCrLf _
                & "   and TATCONPS.STYLE_CODE_PLM (+) = POTPMGM2.STYLE_CODE_PLM"
            Create_TDA(.Tables.Add, "TATCONPS", "**", 0, True, "V", 0)
            .Tables("TATCONPS").Columns.Add("SEL")
            .Tables("TATCONPS").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "TATCONV1", "*")

            Create_TDA(.Tables.Add, "POTPMGM9", "*")

        End With

 
        grdTATCONPU.DataSource = dst.Tables("TATCONPU")
        grdTATCONPS.DataSource = dst.Tables("TATCONPS")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdTATCONPU, grdTATCONPS}
            With grd.DisplayLayout.Bands(0).Columns("SEL")
                .Header.CheckBoxVisibility = HeaderCheckBoxVisibility.WhenUsingCheckEditor
                .Header.Caption = ""
            End With

            With grd.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.No
                .AllowDelete = DefaultableBoolean.False
                .AllowUpdate = DefaultableBoolean.True
            End With

            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = Activation.AllowEdit
                Else
                    gcol.CellActivation = Activation.NoEdit
                End If
            Next
        Next

        If use_dst Then
            Dim row1 As DataRow = frmASFBASE0.dst.Tables("POTPMGM1").Rows(0)
            PROGRAM_NO = row1.Item("PROGRAM_NO")
            rowPOTPMGM1 = dst.Tables("POTPMGM1").NewRow
            For Each dcol As DataColumn In dst.Tables("POTPMGM1").Columns
                rowPOTPMGM1.Item(dcol.ColumnName) = row1.Item(dcol.ColumnName)
            Next
            dst.Tables("POTPMGM1").Rows.Add(rowPOTPMGM1)

            Dim row2 As DataRow = frmASFBASE0.dst.Tables("POTCTOP1").Rows.Find(CONV_TOPIC_NO)
            ' CONV_TOPIC_NO = row2.Item("CONV_TOPIC_NO")
            Dim rowPOTCTOP1 As DataRow = dst.Tables("POTCTOP1").NewRow
            For Each dcol As DataColumn In dst.Tables("POTCTOP1").Columns
                rowPOTCTOP1.Item(dcol.ColumnName) = row2.Item(dcol.ColumnName)
            Next
            dst.Tables("POTCTOP1").Rows.Add(rowPOTCTOP1)

            For Each row As DataRow In frmASFBASE0.dst.Tables("POTPMGM2").Select("")
                Dim rowTATCONPS As DataRow = dst.Tables("TATCONPS").NewRow
                rowTATCONPS.Item("CONV_NO") = "0000000000" '  CONV_NO
                rowTATCONPS.Item("STYLE_CODE_PLM") = row.Item("STYLE_CODE_PLM")
                rowTATCONPS.Item("STYLE_NAME") = row.Item("STYLE_NAME")
                dst.Tables("TATCONPS").Rows.Add(rowTATCONPS)
            Next
            ' 
            For Each row As DataRow In frmASFBASE0.dst.Tables("POTPMGMO").Select("")
                Dim rowTATCONPU As DataRow = dst.Tables("TATCONPU").NewRow
                rowTATCONPU.Item("CONV_NO") = "0000000000" '  CONV_NO
                rowTATCONPU.Item("USER_ID") = row.Item("USER_ID")
                Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", row.Item("USER_ID"))
                rowTATCONPU.Item("USER_EMAIL") = rowASTUSER1.Item("USER_EMAIL")
                dst.Tables("TATCONPU").Rows.Add(rowTATCONPU)
            Next

        Else
            rowPOTPMGM1 = Fill_Record("POTPMGM1", PROGRAM_NO)
            Fill_Records("POTPMGM2", PROGRAM_NO)
            Fill_Record("POTCTOP1", CONV_TOPIC_NO)
            Fill_Records("TATCONPU", CONV_NO)
            Sort_grdColumns(grdTATCONPU, "USER_ID")

        End If

        If EntryMode = "N" Then
            Me.CONV_NO = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")

        End If
        '    Me.CONV_SUBJECT = "Kohl's - cost reduction suggestions"

        If CONV_NO_PREV = "" Then

            UltraTabControl1.Tabs("Message").Text = "New Message"
            UltraTabControl1.Tabs("In Reply To").Visible = False
            cmdReplyInLine.Visible = False

        Else
            rowTATCONV1_PREV = LookUp("TATCONV1", CONV_NO_PREV)

            UltraTabControl1.Tabs("Message").Text = "Message Response"
            UltraTabControl1.Tabs("In Reply To").Visible = True
            cmdReplyInLine.Visible = True
            Dim DocumentFileName As String = ASCMAIN1.Folders("Archive") & "CONV_NO\" & CONV_NO_PREV & ".htm"
            If My.Computer.FileSystem.FileExists(DocumentFileName) Then
                'Dim html As String = ""
                'Using sr As New System.IO.StreamReader(DocumentFileName)
                '    html = sr.ReadToEnd
                'End Using
                txtCONV_NOTES_RepliedTo.Load(DocumentFileName, TXTextControl.StreamType.HTMLFormat)

            Else
                txtCONV_NOTES_RepliedTo.Text = Me.CONV_NO_PREV_NOTES
            End If

            Dim reply_colors() As System.Drawing.Color = { _
                System.Drawing.Color.Purple, _
                System.Drawing.Color.Blue, _
                System.Drawing.Color.Green, _
                System.Drawing.Color.Brown, _
                System.Drawing.Color.Aqua, _
                System.Drawing.Color.Fuchsia}
            '  System.Drawing.Color.Red, _

            ASCMAIN1.sql = "Select Count (*) REPLY_COUNT from TATCONV1 where CONV_NO_PREV = '" & CONV_NO_PREV & "'"
            Dim REPLY_COUNT As Integer = ASCDATA1.GetDataValue

            reply_color = reply_colors(REPLY_COUNT)


        End If
 



        If CONV_NO <> "" Then
            If InquiryMode Then
                EntryMode = "I"
                Me.Text &= " - View Existing Plan"
            ElseIf EntryMode = "E" Then
                Me.Text &= " - Respond to an Existing Message"
            ElseIf EntryMode = "N" Then
                Me.Text &= " - New Message"
            End If

            Toggle(True)

            If InquiryMode Then
                If InquiryMode Then cmdCancel.Text = "Done"
                chkRequestReply.Visible = False
                Absx1.txtFor("OUTSRC_SPEC_INSTR").ReadOnly = True
            Else
            End If

        Else
            EntryMode = "N"
            Me.Text &= " - New Message"
            Stop
            Toggle(False)
        End If

        Bind_Controls(grpProgramInfo, "POTPMGM1")
        Bind_Controls(grpProgramInfo, "POTCTOP1")
        Set_Read_Only(grpProgramInfo, True)

        'Dim html As String = ""
        'Using ff As New System.IO.StreamReader("C:\Users\Walter\Desktop\Ashley\text.htm")
        '    html = ff.ReadToEnd
        'End Using

        '  TextControl2.Load("C:\Users\Walter\Desktop\Ashley\text.tx", TXTextControl.StreamType.InternalFormat)
        ' TextControl2.Load("C:\Users\Walter\Desktop\Ashley\text.htm", TXTextControl.StreamType.HTMLFormat)

        If MESSAGE_BY = "" Then
            MESSAGE_BY = ASCMAIN1.USER_ID
        End If

        Me.Text &= "- Message Author: " & MESSAGE_BY

        '  MakeTransparent(chkRequestReply)

        Dim USER_IDs As New List(Of String)
        For Each row As DataRow In dst.Tables("TATCONPU").Select("", "USER_ID")
            USER_IDs.Add(row.Item("USER_ID"))
        Next

        cbeRequestReplyFrom.DataSource = USER_IDs
        If CONV_NO_PREV <> "" Then
            cbeRequestReplyFrom.Value = rowTATCONV1_PREV.Item("INIT_OPER") & ""
        End If
        ' cbeRequestReplyFrom.Value = "aruna"
        dteRequestReplyBy.Value = Now.AddDays(1).Date
        dte.DateTime = Now
        Set_Read_Only(dte, False)
        Console.WriteLine("here")
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case ""
        End Select
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Select Case Absx1.GetABSColumnName(sender)
            Case "X_CODE"
                '   If Not ScreenMode And EntryMode = "N" Then Set_Default_Vendor()

        End Select

    End Sub
    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OUTSRC_DATE_EXPECTED"

        End Select

    End Sub
    Overrides Sub Prepare_for_View_Lookup_Special(
    ByVal ctl As Control,
    ByVal COLUMN_NAME As String,
    Optional ByRef sql_where As String = "",
    Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            'Case "VEND_CODE"
            '    sql_where = "VEND_CODE in (Select VEND_CODE from APTVENDA where VEND_ATTR_CODE = 'OUTSRC')"
        End Select

    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click
        Dim EMsg As String = ""

        'Dim DocumentFileName As String = ""
        'Dim SaveSettings As New TXTextControl.SaveSettings()
        'Dim DocumentStreamType As TXTextControl.StreamType
        'DocumentStreamType = TXTextControl.StreamType.HTMLFormat

        'If DocumentFileName <> "" Then
        '    ' save under same name and type
        '    txtCONV_NOTES.Save(DocumentFileName, DocumentStreamType)
        'Else
        '    ' save as..
        '    txtCONV_NOTES.Save(TXTextControl.StreamType.All, SaveSettings)
        '    DocumentFileName = SaveSettings.SavedFile
        '    DocumentStreamType = SaveSettings.SavedStreamType
        'End If

        ''TextControl1.Save("C:\VS\VDI\ABS\bin\x86\Debug\TEXT.RTF", TXTextControl.StringStreamType.RichTextFormat)
        ''TextControl1.Save("C:\VS\VDI\ABS\bin\x86\Debug\TEXT.HTML", TXTextControl.StringStreamType.HTMLFormat)
        ''TextControl1.Save("C:\VS\VDI\ABS\bin\x86\Debug\TEXT.TXT", TXTextControl.StringStreamType.PlainText)
        'Exit Sub

        Me.Cursor = Cursors.WaitCursor

        Try
            Update_Record()
            result = "U"
        Catch ex As Exception
            MyBase.Rollback(ex.Message)
            result = "X"
        End Try

        Me.Cursor = Cursors.Default
        Me.Close()
    End Sub

    Public Sub Update_Record()
        Try
            MyBase.BeginTrans()

            If use_dst Then
                ' PROBABLY SHOULD JUST SHUFFLE ALL OF THE DATA IN TO THE DATAASET AND NOT TO ORACLE
            End If
 

            DATETIME_STAMP = Now + ASCMAIN1.NowTSD

            DATETIME_STAMP = dte.DateTime ' TEMP


            For Each row As DataRow In dst.Tables("TATCONPS").Select("SEL='1'")
                Dim STYLE_CODE_PLM As String = row.Item("STYLE_CODE_PLM")
                Dim rowPOTPMGM9 As DataRow = dst.Tables("POTPMGM9").NewRow
                With rowPOTPMGM9
                    .Item("PROGRAM_NO") = PROGRAM_NO
                    .Item("STYLE_CODE_PLM") = STYLE_CODE_PLM
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("PROGRAM_COMMENT") = txtCONV_NOTES.Text
                End With
                dst.Tables("POTPMGM9").Rows.Add(rowPOTPMGM9)

                Dim rowTATCONPS As DataRow = dst.Tables("TATCONPS").NewRow
                With rowTATCONPS
                    .Item("CONV_NO") = CONV_NO
                    .Item("STYLE_CODE_PLM") = STYLE_CODE_PLM
                End With
                dst.Tables("TATCONPS").Rows.Add(rowTATCONPS)
            Next

            For Each row As DataRow In dst.Tables("TATCONPU").Select("SEL='1'")
                Dim USER_ID As String = row.Item("USER_ID")
                Dim rowTATCONPU As DataRow = dst.Tables("TATCONPU").NewRow
                With rowTATCONPU
                    .Item("CONV_NO") = CONV_NO
                    .Item("USER_ID") = USER_ID
                End With
                dst.Tables("TATCONPU").Rows.Add(rowTATCONPU)
            Next

            Dim CONV_NOTES As String = Get_Notes()

            Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").NewRow
            With rowTATCONV1
                .Item("CONV_NO") = CONV_NO
                .Item("CONV_DATE") = DATETIME_STAMP
                .Item("CONV_SUBJECT") = CONV_SUBJECT
                .Item("CONV_NOTES") = CONV_NOTES
                .Item("CONV_STATUS") = "A"

                If chkRequestReply.Checked Then
                    .Item("CONV_FOLLOWUP_BY") = "wjz"
                    .Item("CONV_FOLLOWUP_DATE") = dteRequestReplyBy.Value
                End If

                .Item("CONV_NO_PREV") = CONV_NO_PREV
 
                .Item("TABLE_NAME") = "POTCTOP1"
                .Item("TABLE_KEY") = CONV_TOPIC_NO
                .Item("INIT_OPER") = MESSAGE_BY
                .Item("INIT_DATE") = DATETIME_STAMP
            End With
            dst.Tables("TATCONV1").Rows.Add(rowTATCONV1)

            ASCDATA1.DeleteRows(dst.Tables("TATCONPS"), "CONV_NO = '0000000000' or ISNULL(SEL,'0')<>'1'")
            ASCDATA1.DeleteRows(dst.Tables("TATCONPU"), "CONV_NO = '0000000000' or ISNULL(SEL,'0')<>'1'")

            Update_Record_TDA("POTPMGM9")
            Update_Record_TDA("TATCONV1")
            Update_Record_TDA("TATCONPS", "CONV_NO = '" & CONV_NO & "'")
            Update_Record_TDA("TATCONPU", "CONV_NO = '" & CONV_NO & "'")



            Dim CONV_NO_filename As String = ASCMAIN1.Folders("Archive") & "CONV_NO\" & CONV_NO & ".htm"
            Dim SaveSettings As New TXTextControl.SaveSettings()
            Dim DocumentStreamType As TXTextControl.StreamType
            DocumentStreamType = TXTextControl.StreamType.HTMLFormat

            If CONV_NO_filename <> "" Then
                ' save under same name and type
                txtCONV_NOTES.Save(CONV_NO_filename, DocumentStreamType)
            Else
                ' save as..
                txtCONV_NOTES.Save(TXTextControl.StreamType.All, SaveSettings)
                CONV_NO_filename = SaveSettings.SavedFile
                DocumentStreamType = SaveSettings.SavedStreamType
            End If

            Dim CONV_NOTES_html As String = ""
            Dim DocumentStringStreamType As TXTextControl.StringStreamType
            DocumentStringStreamType = TXTextControl.StringStreamType.HTMLFormat
            txtCONV_NOTES.Save(CONV_NOTES_html, DocumentStringStreamType)

            If reply_in_line Then
                Dim CONV_NO_PREV_filename As String = ASCMAIN1.Folders("Archive") & "CONV_NO\" & CONV_NO_PREV & ".htm"
                If My.Computer.FileSystem.FileExists(CONV_NO_PREV_filename) Then
                    My.Computer.FileSystem.RenameFile(CONV_NO_PREV_filename, CONV_NO_PREV & "_" & CONV_NO & ".htm")
                    My.Computer.FileSystem.RenameFile(CONV_NO_filename, CONV_NO_PREV & ".htm")

                    txtCONV_NOTES.Text = ""
                    Add_Text_Field("See Replies in " & reply_color.Name)
                    txtCONV_NOTES.Save(CONV_NO_filename, DocumentStreamType)

                    Dim CONV_NOTES_html_NOW As String = ""
                    txtCONV_NOTES.Save(CONV_NOTES_html_NOW, DocumentStringStreamType)

                    rowTATCONV1.Item("CONV_NOTES") = CONV_NOTES_html_NOW

                    Update_Record_TDA("TATCONV1")

                    Dim rowTATCONV1_PREV As DataRow = Fill_Record("TATCONV1", CONV_NO_PREV, False, False)
                    Dim b1 As Integer = InStr(CONV_NOTES_html, "<body")
                    b1 = b1 + InStr(Mid(CONV_NOTES_html, b1 + 1), ">") + 2

                    Dim b2 As Integer = InStr(CONV_NOTES_html, "</body>")
                    If b1 <> 0 And b2 <> 0 And b2 > b1 Then
                        'CONV_NOTES_html = Mid(CONV_NOTES_html, b1, b2 - b1 + 7)
                        CONV_NOTES_html = Mid(CONV_NOTES_html, b1, b2 - b1)
                    End If

                    Dim r As String = "<p lang='en-US' style='margin-top:0pt;margin-bottom:0pt;'>"
                    r = Replace(r, "'", Chr(34))
                    CONV_NOTES_html = Replace(CONV_NOTES_html, r, "")
                    CONV_NOTES_html = Replace(CONV_NOTES_html, "</p>", "<br />")
                    If CONV_NOTES_html.Length > 980 Then
                        CONV_NOTES_html = Mid(CONV_NOTES_html, 1, 980)
                    End If
                    CONV_NOTES_html = "<body>" & CONV_NOTES_html & "</body>"
                    rowTATCONV1_PREV.Item("CONV_NOTES") = CONV_NOTES_html ' CONV_NOTES
                    Update_Record_TDA("TATCONV1")
                End If
            End If

            result = "U"

            MyBase.CommitTrans("Message " & CONV_NO & " Updated")

        Catch ex As Exception

            MyBase.Rollback(ex.Message)

            result = "C"
            Me.Close()
        End Try
    End Sub

    Function Get_Notes() As String

        Dim CONV_NOTES As String = txtCONV_NOTES.Text
        Dim FLEN As Integer = dst.Tables("TATCONV1").Columns("CONV_NOTES").MaxLength
        If CONV_NOTES.Length > FLEN Then CONV_NOTES = Mid(CONV_NOTES, 1, FLEN)
        Return CONV_NOTES

    End Function

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        result = "C"
        Me.Close()
    End Sub

    Sub Toggle(tf As Boolean)
        ScreenMode = tf
        'cmdUpdate.Visible = tf
    End Sub
     
    Private Sub chkRequestReply_CheckedChanged(sender As Object, e As EventArgs) Handles chkRequestReply.CheckedChanged
        lblRequestReplyBy.Visible = chkRequestReply.Checked
        dteRequestReplyBy.Visible = chkRequestReply.Checked
        lblRequestReplyFrom.Visible = chkRequestReply.Checked
        cbeRequestReplyFrom.Visible = chkRequestReply.Checked
    End Sub

    Private Sub cmdReplyInLine_Click(sender As Object, e As EventArgs) Handles cmdReplyInLine.Click

        Dim DocumentFileName As String = ASCMAIN1.Folders("Archive") & "CONV_NO\" & CONV_NO_PREV & ".htm"
        If My.Computer.FileSystem.FileExists(DocumentFileName) Then
            txtCONV_NOTES.Load(DocumentFileName, TXTextControl.StreamType.HTMLFormat)
        Else
            txtCONV_NOTES.Text = txtCONV_NOTES_RepliedTo.Text
        End If

        For Each textfield As TXTextControl.TextField In txtCONV_NOTES.TextFields
            textfield.Editable = False
        Next

        txtCONV_NOTES.IsSpellCheckingEnabled = True
        'txtCONV_NOTES.SpellChecker = Nothing
        ' txtCONV_NOTES.ForeColor = Color.Red

        cmdReplyInLine.Visible = False
        reply_in_line = True
    End Sub

    Private Sub txtCONV_NOTES_DoubleClick(sender As Object, e As EventArgs) Handles txtCONV_NOTES.DoubleClick
        ''   txtCONV_NOTES.SpellCheckDialog()
        'txtCONV_NOTES.Sections.Add(TXTextControl.SectionBreakKind.BeginAtNewLine)
        'Dim sec As TXTextControl.Section = txtCONV_NOTES.Sections(1)

        'Dim blue As New TXTextControl.ParagraphStyle("reply")
        'blue.ForeColor = Color.Green
        'txtCONV_NOTES.ParagraphStyles.Add(blue)

        '   Add_Text_Field()
 
    End Sub

    Private Sub txtCONV_NOTES_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCONV_NOTES.KeyDown
        If e.KeyCode = Keys.F3 And reply_in_line Then
            Add_Text_Field()
        End If
    End Sub

    Sub Add_Text_Field(Optional msg As String = "")

        Dim canadd As Boolean = False
        canadd = True
        'Try
        '    canadd = txtCONV_NOTES.TextFields.CanAdd
        'Catch ex As Exception

        'End Try

        If canadd Then

            Dim text As String = MESSAGE_BY & " " & Format(DATETIME_STAMP, "MM/dd HH:mm") & "-" & vbLf
            text = MESSAGE_BY & " " & Format(DATETIME_STAMP, "MM/dd") & " - " & msg & vbLf

            Dim NewField As New TXTextControl.TextField(text)
            'NewField.ShowActivated = True
            'NewField.DoubledInputPosition = True
            ' txtCONV_NOTES.InputPosition = new TXTextControl.InputPosition( 1, 3, 0)
            txtCONV_NOTES.TextFields.Add(NewField)
            If NewField.Start = -1 Then Exit Sub

            txtCONV_NOTES.Select(NewField.Start - 1, NewField.Length)
            txtCONV_NOTES.Selection.ForeColor = reply_color
            txtCONV_NOTES.Selection.FontSize = 10 * 20
            txtCONV_NOTES.Select(NewField.Start + NewField.Length - 2, 0)

        End If
    End Sub
    Private Sub txtCONV_NOTES_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtCONV_NOTES.KeyPress

    End Sub

    Private Sub txtCONV_NOTES_SpellCheckText(sender As Object, e As TXTextControl.SpellCheckTextEventArgs) Handles txtCONV_NOTES.SpellCheckText

    End Sub

    Private Sub txtCONV_NOTES_TextChanged(sender As Object, e As EventArgs) Handles txtCONV_NOTES.TextChanged

    End Sub

    Private Sub txtCONV_NOTES_TextFieldChanged(sender As Object, e As TXTextControl.TextFieldEventArgs) Handles txtCONV_NOTES.TextFieldChanged

    End Sub

    Private Sub txtCONV_NOTES_TextFieldClicked(sender As Object, e As TXTextControl.TextFieldEventArgs) Handles txtCONV_NOTES.TextFieldClicked

    End Sub

    Private Sub txtCONV_NOTES_TextFieldEntered(sender As Object, e As TXTextControl.TextFieldEventArgs) Handles txtCONV_NOTES.TextFieldEntered

    End Sub
End Class