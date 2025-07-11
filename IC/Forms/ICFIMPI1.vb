Imports nsoftware.IPWorks

Public Class ICFIMPI1

    Dim WithEvents FTP1 As nsoftware.IPWorks.Ftp
    Dim REMOTEDIRECTORYFILELIST As List(Of String) = New List(Of String)
    Dim displaycontrol As Control = Nothing
    Dim FOLDERNAME As String = ""
    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim MISSING_COLORS_FOUND As String = ""
    Dim MISSING_CUBE_FOUND As String = ""
    Dim rowSOTPARM1 As DataRow
    Dim ICTSTYL1 As String
    Dim ICTSTYV1 As String
    Dim ICTSTYL3 As String




#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
        rowSOTPARM1 = LookUp("SOTPARM1", "Z")
        SO_PARM_UPC_VENDOR_ID = rowSOTPARM1.Item("SO_PARM_UPC_VENDOR_ID") & ""

        With dst

            Create_TDA(.Tables.Add, "ICTSTYL1", "*")

            .Tables("ICTSTYL1").Columns.Add("EXISTING")

            ICTSTYL1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Create_TDA(.Tables.Add, "ICTSTYC1", "*")

            Create_TDA(.Tables.Add, "ICTSTYV1", "*")
            ICTSTYV1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Create_TDA(.Tables.Add, "ICTSTYL3", "*")
            ICTSTYL3 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Create_TDA(.Tables.Add, "ICTSTAT2", "*")

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTCOLR1", 1))

            ASCMAIN1.sql = "Select * from ICTSIZE1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSIZE1", 1))

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTCLAS1", 1))

            ASCMAIN1.sql = "Select * from ICTRSPBX"
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTRSPBX", 1))

        End With

        grdICTSTYL1.DataSource = dst.Tables("ICTSTYL1")
        'grdPOTCENT2.DataSource = dst.Tables("POTCENT2")

        Create_Summary(grdICTSTYL1, "STYLE_CODE", "Count")

        spl.Panel1Collapsed = True

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Import Styles"

                If ASCMAIN1.useUNCPath Then
                    FOLDERNAME = $"{ASCMAIN1.Folders("SharedRoot")}\NEWITEMS\"
                Else
                    FOLDERNAME = "S:\NEWITEMS\"
                End If

                If My.Computer.FileSystem.GetFiles(FOLDERNAME).Count = 0 Then
                    EMsg &= vbCr & "No Files to be processed"
                End If

            Case "Update"

                Dim EXISTING_STYLES_FOUND As Boolean = False


                For Each ROWICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("")
                    Dim STYLE_CODE As String = ROWICTSTYL1.Item("STYLE_CODE")
                    Dim ROW As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If ROW Is Nothing Then
                        ROWICTSTYL1.Item("EXISTING") = "0"
                    Else
                        If Not EXISTING_STYLES_FOUND Then
                            EMsg &= vbCr & "Some Styles to be imported are already in Style Master Table"
                            EXISTING_STYLES_FOUND = True
                        End If
                        ROWICTSTYL1.Item("EXISTING") = "1"
                    End If
                Next
                If MISSING_COLORS_FOUND = "Y" Then
                    EMsg &= vbCr & "Some Colors to be imported are in the Colors Master Table"
                    EXISTING_STYLES_FOUND = True
                End If


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey


            Case "Import Styles"

                PROCESS_FILES()
                Mode_Settings(True)
  

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Import Styles").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ICTSTYL1", "ICTSTYC1", "ICTSTYV1", "ICTSTYL3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        'Load_POTSHIPS()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then

        Else

        End If


        If EntryMode = "N" Then
        Else

        End If

    End Sub
    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()
        Update_Record_TDA("ICTSTYL1")
        Update_Record_TDA("ICTSTYC1") ' unrem 

        Update_Record_TDA("ICTSTYV1")
        Update_Record_TDA("ICTSTYL3")

        Update_Record_TDA("ICTSTAT2") ' un rem 

        CommitTrans("Update Complete")


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "SEASON_CODE"
            '    If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
            '        sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE") & "'"
            '    End If
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()

        ' Load_Popup_Menu(grdEDT855, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Style Status Inquiry")
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

        Select Case e.SourceControl.Name


        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

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

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "LP_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
            '        Load_EDT846T1()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "LP_CODE"
            '    Load_EDT846T1()
        End Select
    End Sub


    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)

        With Absx1.txtFor(COLUMN_NAME)
            Select Case COLUMN_NAME

                'Case "LP_CODE"
                '    Load_EDT846T1()

            End Select

        End With
    End Sub

#End Region

    Sub PROCESS_FILES()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Processing File", "")

        For Each filename As String In My.Computer.FileSystem.GetFiles(FOLDERNAME)
            Process_File(filename)
        Next


        ASCMAIN1.Progress("", "")
        Me.Cursor = Cursors.Default

        MsgBox("All Files Loaded")


 

        'Process_File(FILENAME:=)
    End Sub

    Sub Process_File(FILENAME As String)

        Dim COST_EFF_DATE As String = Format(Absx1.dteFor("COST_DATE").Value, "dd-MMM-yyyy")

        Dim data As String = ""
        Using sr As New System.IO.StreamReader(FILENAME)
            data = sr.ReadToEnd
            Dim datarec() As String = Split(data, vbCrLf)
            Dim DATAROW() As String
            ' Dim offset As Integer = 2
            Dim MPQ As Integer = 0
            Dim BOX As Integer = 0
            Dim LookForNums As String = "0123456789"

            For i As Integer = 2 To UBound(datarec) - 1
                If datarec(i) <> " " Then
                    DATAROW = Split(datarec(i), Chr(9))
                    If Trim(DATAROW(0)) = "" Then

                    Else
                        Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").NewRow
                        Dim DANAC As String = Trim(DATAROW(0))
                        If DANAC = "MT23228" Then
                            Dim RYAN As String = ""
                        End If
                        With rowICTSTYL1

                            Dim STYLE_CLASS As String = Trim(DATAROW(25))
                            Dim CLASS_CODE As String = ""
                            If STYLE_CLASS = "FOLIAGE" Then
                                CLASS_CODE = "FOLLIAGE"
                            ElseIf STYLE_CLASS = "FOL" Then
                                CLASS_CODE = "FOLLIAGE"
                            ElseIf STYLE_CLASS = "GEN" Then
                                CLASS_CODE = "GENERAL"
                            ElseIf STYLE_CLASS = "GARD" Then
                                CLASS_CODE = "GARDEN"
                            ElseIf STYLE_CLASS = "PVC" Then
                                CLASS_CODE = "PVC"
                            ElseIf STYLE_CLASS = "PVC" Then
                                CLASS_CODE = "PVC"
                            ElseIf STYLE_CLASS = "HALL" Then
                                CLASS_CODE = "HALLOWEEN"
                            ElseIf STYLE_CLASS = "HALLOWEEN" Then
                                CLASS_CODE = "HALLOWEEN"
                            ElseIf STYLE_CLASS = "DECO" Then
                                CLASS_CODE = "DECOR"
                            ElseIf STYLE_CLASS = "VALE" Then
                                CLASS_CODE = "VALENTINE"
                            ElseIf STYLE_CLASS = "XMAS" Then
                                CLASS_CODE = "XMAS"
                            ElseIf STYLE_CLASS = "EAST" Then
                                CLASS_CODE = "EASTER"
                            ElseIf STYLE_CLASS = "FLO" Then
                                CLASS_CODE = "FLOWER"
                            ElseIf STYLE_CLASS = "FLOWER" Then
                                CLASS_CODE = "FLOWER"
                            ElseIf STYLE_CLASS = "FALL" Then
                                CLASS_CODE = "FALL"
                            Else
                                CLASS_CODE = "ZZZ"
                            End If
                            Dim CLASS_CODE_L = STYLE_CLASS
                            Dim rowICTCLAS1 As DataRow = LookUp("ICTCLAS1", CLASS_CODE_L)

                            If rowICTCLAS1 Is Nothing Then
                                CLASS_CODE = "ZZZ"
                            Else
                                CLASS_CODE = STYLE_CLASS
                            End If

                            .Item("STYLE_CODE") = Trim(Replace(DATAROW(0), " ", ""))
                            .Item("STYLE_STATUS") = "A"
                            '.Item("STYLE_DESC") = Trim(Replace(DATAROW(1), Chr(34), ""))
                            Dim STYLE_DESC As String = Trim(Mid(DATAROW(1), 1, 48))

                            STYLE_DESC = Replace(Mid(STYLE_DESC, 1, 1), Chr(34), "") & Mid(STYLE_DESC, 2)

                            Dim DANATEST As String = STYLE_DESC
                            STYLE_DESC = Replace(Mid(STYLE_DESC, 1, 1), Chr(34), "") & Mid(STYLE_DESC, 2)
                            Dim DESCLEN As Integer = STYLE_DESC.Length
                            ' If DESCLEN > 0 Then
                            'STYLE_DESC = Mid(STYLE_DESC, 1, DESCLEN - 4) & Replace(Mid(STYLE_DESC, DESCLEN - 3), Chr(34), "")
                            'End If
                            STYLE_DESC = Replace(STYLE_DESC, Chr(34) & Chr(34) & Chr(34), Chr(34))
                            STYLE_DESC = Replace(STYLE_DESC, Chr(34) & Chr(34), Chr(34))

                            Dim dana_desc As String = STYLE_DESC
                            Dim dana_1 As String = Mid(STYLE_DESC, STYLE_DESC.Length, 1)
                            Dim dana_2 As String = Mid(STYLE_DESC, STYLE_DESC.Length - 1, 1)

                            Dim bad_q As Integer = LookForNums.IndexOf(Mid(STYLE_DESC, STYLE_DESC.Length - 1, 1))

                            If Mid(STYLE_DESC, STYLE_DESC.Length, 1) = Chr(34) And bad_q = -1 Then
                                mid(STYLE_DESC, STYLE_DESC.Length, 1) = " "
                            End If
                            .Item("STYLE_DESC") = Mid(STYLE_DESC, 1, 45)
                            Dim drcdesc As String = Mid(STYLE_DESC, 1, 45)


                            .Item("STYLE_ASST_QTY") = Val(Trim(DATAROW(2))) ' added 03/31/2018

                            .Item("STYLE_DESC2") = Trim(Replace(DATAROW(4), Chr(34), ""))

                            .Item("SUB_UNIT_BAG_QTY") = Val(Trim(DATAROW(5)))

                            Dim RESHIPBOX_L = Trim(DATAROW(6))
                            Dim rowICTRSPBX() As DataRow = dst.Tables("ICTRSPBX").Select("RESHIPBOX_DESC = '" & RESHIPBOX_L & "'")

                            Dim RESHIPBOX_CODE As String = ""

                            If RESHIPBOX_L = "" Then
                                RESHIPBOX_CODE = ""
                            Else
                                RESHIPBOX_CODE = rowICTRSPBX(0).Item("RESHIPBOX_CODE")
                            End If

                            .Item("RESHIPBOX_CODE") = RESHIPBOX_CODE
                            '  .Item("CARTON_PACK_QTY") = Val(Trim(DATAROW(7)))

                            .Item("INNER_PACK_QTY") = Val(Trim(DATAROW(8)))
                            .Item("CARTON_PACK_QTY") = Val(Trim(DATAROW(9)))
                            .Item("CASE_CUBE") = Val(Trim(DATAROW(10)))
                            .Item("STYLE_UOM") = Trim(DATAROW(11))
                            .Item("STYLE_PO_QTY_MIN") = Val(Trim(DATAROW(13)))
                            .Item("STYLE_PRICE") = Val(Trim(DATAROW(14)))
                            .Item("LABEL_TYPE_CODE") = Trim(DATAROW(15))
                            .Item("PURCH_NOTES") = Trim(DATAROW(17))
                            .Item("DUTY_RATE_CODE") = Trim(DATAROW(18))
                            ' .Item("STYLE_MATL_DESC") = Trim(DATAROW(0)) ***** NEED TO SUPPORT
                            '.Item("STYLE_SO_QTY_MIN") = Trim(DATAROW(6))

                            MPQ = Val(Trim(DATAROW(7)))
                            BOX = Val(Trim(DATAROW(8)))

                            If MPQ <> 0 Then
                                .Item("STYLE_SO_QTY_MIN") = MPQ
                            Else
                                If BOX <> 0 Then
                                    .Item("STYLE_SO_QTY_MIN") = BOX
                                Else
                                    .Item("STYLE_SO_QTY_MIN") = 1
                                End If
                            End If


                            Dim EXCLUSIVE_STYLE As String = Trim(DATAROW(33))
                            If EXCLUSIVE_STYLE.ToUpper = "YES" Then
                                .Item("EXCLUSIVE_STYLE") = "1"
                            Else
                                .Item("EXCLUSIVE_STYLE") = ""
                            End If

                            Dim STYLE_COUNTRY As String = Trim(DATAROW(19))
                            Dim COUNTRY_CODE As String = ""

                            If STYLE_COUNTRY = "CAMBODIA" Then
                                COUNTRY_CODE = "KHM"
                            ElseIf STYLE_COUNTRY = "CHINA" Then
                                COUNTRY_CODE = "CHN"
                            ElseIf STYLE_COUNTRY = "INDIA" Then
                                COUNTRY_CODE = "IND"
                            ElseIf STYLE_COUNTRY = "PHILIPPINES" Then
                                COUNTRY_CODE = "PHL"
                            ElseIf STYLE_COUNTRY = "POLAND" Then
                                COUNTRY_CODE = "POL"
                            ElseIf STYLE_COUNTRY = "TAIWAN" Then
                                COUNTRY_CODE = "TWN"
                            ElseIf STYLE_COUNTRY = "THAILAND" Then
                                COUNTRY_CODE = "THA"
                            ElseIf STYLE_COUNTRY = "UNITED STATES" Then
                                COUNTRY_CODE = "USA"
                            Else
                                COUNTRY_CODE = "ZZZ"
                            End If
                            .Item("COUNTRY_CODE") = COUNTRY_CODE
                            .Item("VEND_CODE") = Trim(DATAROW(20))

                            .Item("STYLE_CLASS_CODE") = CLASS_CODE

                            Dim style_size As String = Trim(DATAROW(26))

                            Dim rowICTSIZE1 As DataRow = LookUp("ICTSIZE1", style_size)

                            If rowICTSIZE1 Is Nothing Then
                                'COLOR = Mid(COLOR_CODE_LONG, 1, 4)

                            Else
                                style_size = rowICTSIZE1.Item("SIZE_CODE")
                            End If

 
                            .Item("SIZE_CODE") = style_size
                            .Item("SUB_UNIT_PACK_QTY") = 1 ' need to calculate by UOM 

                            .Item("SALES_DIVISION_CODE") = "RIB"
                            .Item("INIT_DATE") = Now
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_DATE") = Now
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        End With

                        If Trim(DATAROW(0)) <> "" Then
                            dst.Tables("ICTSTYL1").Rows.Add(rowICTSTYL1)
                        End If
                        Dim DANA As String = Trim(DATAROW(0))

                        Dim rowICTSTYV1 As DataRow = dst.Tables("ICTSTYV1").NewRow
                        With rowICTSTYV1
                            .Item("STYLE_CODE") = Trim(Replace(DATAROW(0), " ", ""))
                            .Item("VEND_CODE") = Trim(DATAROW(20))
                            .Item("VEND_ITEM_CODE") = Trim(Mid(Replace(DATAROW(3), Chr(34), ""), 1, 25))
                            .Item("PO_COST") = Val(Trim(DATAROW(12)))
                            If Trim(DATAROW(24)) <> "" Then
                                '.Item("PO_COST_DATE") = FormatDATE(Trim(DATAROW(20)))
                                '.Item("PO_COST_DATE") = CDate(Trim(DATAROW(20)))
                                .Item("PO_COST_DATE") = CDate(Trim(COST_EFF_DATE))

                            End If
                            ' CONVERT TO DATE 
                            '.Item("PO_COST_DATE") = "02/17/14"
                        End With
                        If Trim(DATAROW(0)) <> "" Then
                            dst.Tables("ICTSTYV1").Rows.Add(rowICTSTYV1)
                        End If


                        For C As Integer = 0 To 17
                            If Trim(DATAROW(C + 34)) <> "" Then
                                Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").NewRow
                                Dim COLOR_CODE_LONG As String = Trim(DATAROW(C + 34))
                                Dim COLOR_CODE As String = Trim(DATAROW(C + 34))
                                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                                Dim rowICTCOLR1L() As DataRow = dst.Tables("ICTCOLR1").Select("COLOR_CODE_LONG = '" & COLOR_CODE_LONG & "'")
                                'Dim rowICTCOLR1LL As DataRow = dst.Tables("ICTQUOT2").Rows.Find(New String() {QUOTE_NO, STYLE_CODE_PLM})
                                'Dim rowICTCOLR1L() As DataRow = dst.Tables("ICTCOLR1").Select("COLOR_CODE_LONG = '" & COLOR_CODE_LONG & "' AND COLOR_STATUS = 'A'")
                                Dim COLOR As String = ""
                                If rowICTCOLR1 Is Nothing Then
                                    'COLOR = Mid(COLOR_CODE_LONG, 1, 4)
                                    Dim COLOR_CODE_L As String = ""
                                    'If rowICTCOLR1L(0).Item("COLOR_CODE") Is Nothing Then
                                    If rowICTCOLR1L.Length = 0 Then
                                        MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style " & Trim(Replace(DATAROW(0), " ", "") & COLOR_CODE_LONG & " Is Invalid"))
                                        MISSING_COLORS_FOUND = "Y"
                                        ' COLOR = Mid(COLOR_CODE_LONG, 1, 4)
                                    Else
                                        COLOR = rowICTCOLR1L(0).Item("COLOR_CODE")
                                        'MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style " & Trim(Replace(DATAROW(0), " ", "") & COLOR_CODE_LONG & " Is Invalid"))
                                    End If
                                Else
                                    COLOR = rowICTCOLR1.Item("COLOR_CODE")
                                End If


                                With rowICTSTYC1
                                    .Item("STYLE_CODE") = Trim(Replace(DATAROW(0), " ", ""))
                                    .Item("COLOR_CODE") = COLOR
                                    ' REM NEED TO GET UPC CODE
                                    .Item("UPC_CODE") = Get_UPC_Code(rowICTSTYC1.Item("STYLE_CODE"), rowICTSTYC1.Item("COLOR_CODE"))
                                    .Item("STYLE_COLOR_STATUS") = "A"
                                End With

                                If Trim(DATAROW(0)) <> "" Then
                                    dst.Tables("ICTSTYC1").Rows.Add(rowICTSTYC1)
                                End If

                                ' ESTABLISH ICTSTAT2 RECS WITH JUST KEYS - SO ATTRIBUTE LIST WORKS WITH NEW STYLES
                                Dim rowICTSTAT2 As DataRow = dst.Tables("ICTSTAT2").NewRow

                                With rowICTSTAT2
                                    .Item("STYLE_CODE") = Trim(Replace(DATAROW(0), " ", ""))
                                    .Item("COLOR_CODE") = COLOR
                                    .Item("WHSE_CODE") = "MS"
                                End With

                                If Trim(DATAROW(0)) <> "" Then
                                    dst.Tables("ICTSTAT2").Rows.Add(rowICTSTAT2)
                                End If


                            End If
                        Next
                        For A As Integer = 0 To 2
                            If DATAROW(A + 25) <> "" Then
                                Dim rowICTSTYL3 As DataRow = dst.Tables("ICTSTYL3").NewRow
                                With rowICTSTYL3
                                    .Item("STYLE_CODE") = Trim(Replace(DATAROW(0), " ", ""))
                                    .Item("ATTR_CODE") = Trim(Replace(DATAROW(A + 27), ",", ""))
                                End With
                                If Trim(DATAROW(0)) <> "" And Trim(DATAROW(A + 27)) <> "" Then
                                    dst.Tables("ICTSTYL3").Rows.Add(rowICTSTYL3)
                                End If
                            End If
                        Next
                    End If
                End If
            Next


        End Using

        Dim HISTFILE As String = ""
        ' If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_COMPANY = "RGI" Then
        '  HISTFILE = "c:\VS\VDI\NEWITEMSHIST" & Mid(FILENAME, Len(FOLDERNAME))
        ' Else


        If ASCMAIN1.useUNCPath Then
            HISTFILE = $"{ASCMAIN1.Folders("SharedRoot")}\NEWITEMSHIST" & Mid(FILENAME, Len(FOLDERNAME))
        Else
            HISTFILE = "s:\NEWITEMSHIST" & Mid(FILENAME, Len(FOLDERNAME))
        End If
        'End If

        'Dim HISTFILE As String = "s:\NEWITEMSHIST" & Mid(FILENAME, Len(FOLDERNAME))

        My.Computer.FileSystem.MoveFile(FILENAME, HISTFILE, True)


        'Assign_PO_SHIPMENT_NO()

    End Sub

    Function Get_UPC_Code(STYLE_CODE As String, COLOR_CODE As String) As String

        Dim UPC_CODE As String = ""
        Do
            Dim UPC_CODE_CTL_NO As String = ""
            If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("UPC_CODE")
            Else
                UPC_CODE_CTL_NO = ASCMAIN1.Next_Control_No("ICTUPCH1.UPC_CODE")
            End If

            UPC_CODE = TAC.SOCMAIN1.UPC(Me, UPC_CODE_CTL_NO, SO_PARM_UPC_VENDOR_ID, True)
            If LookUp("ICTUPCH1", UPC_CODE) Is Nothing Then Exit Do
        Loop

        ASCMAIN1.sql = "Insert into ICTUPCH1 (UPC_CODE,STYLE_CODE,COLOR_CODE,INIT_DATE,INIT_OPER) " & vbCrLf _
            & " values (:PARM1,:PARM2,:PARM3,SYSDATE,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {UPC_CODE, STYLE_CODE, COLOR_CODE, ASCMAIN1.USER_ID})

        Return UPC_CODE
    End Function



    Public Shared Function FormatDATE(ByVal PDATE As String) As Date
        If Trim(PDATE) <> "" Then
            FormatDATE = CDate((Mid(PDATE, 6, 2) & "/" & Mid(PDATE, 9, 2) & "/" & Mid(PDATE, 1, 4)))
        Else
            FormatDATE = Nothing
        End If


    End Function

End Class