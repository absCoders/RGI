Public Class POFORDRA
    Dim PONO As String = ""
    Dim VAN_REF As String = ""
    Dim VAN_REF_previous As String = ""
    Dim VAN_REF_last_AI As String = ""
    Dim PO_ORDER_NO As String = ""
    Dim rowPOTORDR1 As DataRow
    Dim rowPOTORDRA As DataRow
    Dim tblPOTORDRA As DataTable
    Dim rowpohdr As DataRow
    Dim grdKeys As New Dictionary(Of String, String)
    Dim gcols As New Dictionary(Of String, List(Of String))
    Dim gcols_excluded As New Dictionary(Of String, String())
    Dim gcols_different As New Dictionary(Of String, List(Of String))
    Dim options As ValueList
    Dim grdATs As New Dictionary(Of String, UltraWinGrid.UltraGrid)
    Dim VEND_CODE As String
    Dim poTables() As String = New String() {"pohdr", "pocolor", "pofactory", "posize", "posizedtl", "potrim", "pofabric"}
    Dim PO_PARM_PO_IMG_DIR As String = ""
    Dim PO_ORDER_LNO As Integer = 0
    Dim StyleColors As New Dictionary(Of String, Int32)


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("POTPARM1")
        Get_PARM("SOTPARM1")

        PO_PARM_PO_IMG_DIR = ROWs("POTPARM1").Item("PO_PARM_PO_IMG_DIR") & ""
        If ASCMAIN1.Running_in_VS Then
            PO_PARM_PO_IMG_DIR = "C:\Users\Walter\Desktop\Ashley\PO_Images"
        End If

        With dst
            Create_TDA(.Tables.Add, "POTORDRA", "*")

            'ASCMAIN1.sql = "Select POTORDRA.*" & vbCrLf _
            '& " from POTORDRA where STATUS = :PARM1 or :PARM1 = '*'"

            ASCMAIN1.sql = "SELECT POTORDRA.*" & vbCrLf _
            & ", x.`SentSeq` SENDNO, x.`SendRemarks` REMARKS, x.`StyleNo` STYLE_NO, x.`Style` STYLE_DESC, x.`StyleRef` STYLE_REF, x.`OrderDate` ORDR_DATE, x.`OrderCfmDate` ORDR_CONF_DATE" & vbCrLf _
            & ", x.`VandaleShipDate` SHIP_DATE, x.`TotalQty` TOTAL_QTY, x.`TotalVandaleAmount` TOTAL_AMT" & vbCrLf _
            & ", x.`CreateBy` CREATED_BY, x.`VandaleUser` VANDALE_USER, x.`FollowBy` FOLLOWED_BY, x.`FollowByEmail` FOLLOWED_BY_EMAIL" & vbCrLf _
            & " from POTORDRA, AT.`pohdr` X where X.VAN_REF = POTORDRA.VAN_REF" & vbCrLf _
            & " and (POTORDRA.STATUS = :PARM1 or :PARM1 = '*')"
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", ChrW(34))
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "POTORDR1", "*", 1)
            Create_TDA(.Tables.Add, "POTORDR2", "*", 1)

            ASCMAIN1.sql = "Select * from POTORDXR where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDXR", "**", 0, True, "V")

            For Each TABLE_NAME As String In poTables
                ASCMAIN1.sql = "Select * from AT." & Chr(34) & TABLE_NAME & Chr(34) & " where VAN_REF = :PARM1"
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "V")
                dst.Tables(TABLE_NAME).Columns.Add("ADD")
            Next

            '   Create_Relation("pocolor", "posizedtl", "VAN_REF, POColorKey")
            With .Tables("posizedtl")
                ' .Columns.Add("ColorCode", GetType(System.String), "parent.ColorCode")
                .Columns.Add("ColorCode", GetType(System.String))
            End With

            With .Tables.Add("differences")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("TABLE_KEY", GetType(System.Int64))
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("DIFFERENCE")
                .PrimaryKey = New DataColumn() {.Columns("TABLE_NAME"), .Columns("TABLE_KEY"), .Columns("COLUMN_NAME")}
            End With

            With .Tables.Add("POTORDRB")
                .Columns.Add("VAN_REF")
                .Columns.Add("DATE_RECEIVED", GetType(System.DateTime))
                .Columns.Add("STATUS")
                .Columns.Add("LAST_DATE", GetType(System.DateTime))
                .Columns.Add("LAST_OPER")
                .Columns.Add("MESSAGE")
                .PrimaryKey = New DataColumn() {.Columns("VAN_REF")}
            End With


            If ASCMAIN1.Running_in_VS And False Then

                ASCMAIN1.sql = "Create Table ICTSTYLF (" _
                    & "STYLE_CODE VARCHAR2(12)," _
                    & "BORN DATE," _
                    & "STYLE_DESC VARCHAR2(35)," _
                    & "SIZE_SCALE_ORIG VARCHAR2(255)," _
                    & "SIZE_SCALE VARCHAR2(255)," _
                    & "Primary Key (STYLE_CODE))"
                ' Stop

                ASCMAIN1.sql = "Select STYLE_CODE, TRUNC(INIT_DATE) BORN, STYLE_DESC, SIZE_SCALE SIZE_SCALE_ORIG, SIZE_SCALE" & vbCrLf _
                    & " from ICTSTYL1" & vbCrLf _
                    & " where NVL(INIT_DATE,SYSDATE) > '01-JAN-2000'" & vbCrLf _
                    & "    or STYLE_CODE in (Select STYLE_CODE from ICTSTAT2 where WHSE_QTY_ON_ORDER <> 0 or WHSE_QTY_ON_HAND <> 0)"
                Create_TDA(.Tables.Add, "ICTSTYLF", "**", 0, True, "", 1)
                With .Tables("ICTSTYLF").Columns
                    Dim QTOTAL As String = ""
                    For I As Integer = 1 To 24
                        .Add("S" & CStr(I))
                        .Add("Q" & CStr(I), GetType(System.Int32))
                        QTOTAL &= "+ISNULL(Q" & CStr(I) & ",0)"
                    Next
                    .Add("SQ")
                    .Add("QTOTAL", GetType(System.Int32), Mid(QTOTAL, 2))
                End With

                Create_TDA(.Tables.Add, "ICTSTYLS", "*")

                ASCMAIN1.sql = "Create Table ICTSTYCF (" _
                    & "STYLE_CODE VARCHAR2(12)," _
                    & "COLOR_CODE VARCHAR2(6)," _
                    & "STYLE_COLOR_DESC VARCHAR2(60)," _
                    & "Primary Key (STYLE_CODE, COLOR_CODE))"
                ' Stop

                ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, STYLE_COLOR_DESC" & vbCrLf _
                    & " from ICTSTYC1" & vbCrLf _
                    & " where ICTSTYC1.STYLE_CODE in (" & vbCrLf _
                    & "Select STYLE_CODE from ICTSTYL1" & vbCrLf _
                    & " where NVL(INIT_DATE,SYSDATE) > '01-JAN-2000'" & vbCrLf _
                    & "    or STYLE_CODE in (Select STYLE_CODE from ICTSTAT2 where  WHSE_QTY_ON_ORDER <> 0 or WHSE_QTY_ON_HAND <> 0))"
                Create_TDA(.Tables.Add, "ICTSTYCF", "**", 0, True, "", 2)

                Create_Relation("ICTSTYLF", "ICTSTYCF", "STYLE_CODE")
            End If



            With .Tables.Add("Images")
                .Columns.Add("FILENAME")
                .Columns.Add("IMAGE", GetType(System.Drawing.Bitmap))
                .Columns.Add("IMAGE_TYPE")
                .Columns.Add("IMAGE_DESC")
                .Columns.Add("SOURCE")
                .Columns.Add("POKey", GetType(System.Int32))
                .Columns.Add("POTrimKey", GetType(System.Int32))
            End With



            With .Tables.Add("PDFs")
                .Columns.Add("FILENAME")
                .Columns.Add("FILEDATETIME", GetType(System.DateTime))
            End With


            ASCMAIN1.sql = "Select STYLE_CODE, STYLE_DESC, SIZE_SCALE, STYLE_MATL_DESC, CARTON_PACK_QTY, INNER_PACK_QTY, SUB_UNIT_PACK_QTY, FACTORY_CODE from ICTSTYL1 where STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "Style", "**", 0, False, "V", 0)
            .Tables("Style").Columns.Add("VAN_REF")
            .Tables("Style").Columns.Add("ADD")
            For Each dcol As DataColumn In .Tables("Style").Columns
                dcol.MaxLength = -1
                dcol.AllowDBNull = True
            Next

            ASCMAIN1.sql = "Select ICTSTYC1.STYLE_CODE, ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC, ICTSTYC1.STYLE_COLOR_DESC" & vbCrLf _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1 and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "StyleColor", "**", 0, False, "V", 0)
            .Tables("StyleColor").Columns.Add("VAN_REF")
            .Tables("StyleColor").Columns.Add("ADD")
            For Each dcol As DataColumn In .Tables("StyleColor").Columns
                dcol.MaxLength = -1
                dcol.AllowDBNull = True
            Next

            'With .Tables.Add("StyleSize")
            '    .Columns.Add("STYLE_CODE")
            '    For I As Integer = 1 To 12
            '        Dim SZ As String = "S" & Format(I, "00")
            '        Dim SQ As String = "Q" & Format(I, "00")
            '        .Columns.Add(SZ)
            '        .Columns.Add(SQ, GetType(System.Int32))
            '    Next
            'End With

            ASCMAIN1.sql = "Select ICTSTYLS.* from ICTSTYLS where ICTSTYLS.STYLE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "StyleSize", "**", 0, False, "V", 0)

            .Tables("StyleSize").Columns.Add("VAN_REF")
            .Tables("StyleSize").Columns.Add("ADD")
            For Each dcol As DataColumn In .Tables("StyleSize").Columns
                dcol.MaxLength = -1
                dcol.AllowDBNull = True
            Next

        End With

        grdStyle.DataSource = dst.Tables("Style")
        grdStyleColor.DataSource = dst.Tables("StyleColor")
        grdStyleSize.DataSource = dst.Tables("StyleSize")
        grdStyleSize.DisplayLayout.Bands(0).LevelCount = 2
        grdStyleSize.DisplayLayout.Bands(0).ColHeadersVisible = False



        grdImages.DataSource = dst.Tables("Images")
        grdImages.DisplayLayout.Bands(0).Columns("IMAGE").Hidden = True

        grdPDFs.DataSource = dst.Tables("PDFs")
        'grdImages.DisplayLayout.Bands(0).Columns("IMAGE").Hidden = True

        If ASCMAIN1.Running_in_VS And False Then
            grdICTSTYLF.DataSource = dst.Tables("ICTSTYLF")
            Create_Summary(grdICTSTYLF, "STYLE_CODE", "Count")
        End If

        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdPOTORDRB.DataSource = dst.Tables("POTORDRB")

        grdKeys.Add("pohdr", "POKey")
        grdKeys.Add("pocolor", "POColorKey")
        grdKeys.Add("pofactory", "POFactoryKey")
        grdKeys.Add("posize", "POSizeKey")
        grdKeys.Add("posizedtl", "POSizeKey")
        grdKeys.Add("potrim", "POTrimKey")
        grdKeys.Add("pofabric", "POFabrickey")

        grdATs.Add("pohdr", grdpohdr)
        grdATs.Add("pocolor", grdpocolor)
        grdATs.Add("pofactory", grdpofactory)
        grdATs.Add("posize", grdposize)
        grdATs.Add("posizedtl", grdposizedtl)
        grdATs.Add("potrim", grdpotrim)
        grdATs.Add("pofabric", grdpofabric)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdpohdr, grdpocolor, grdpofactory, grdposize, grdposizedtl, grdpotrim, grdpofabric, grdStyle, grdStyleColor, grdStyleSize}
            Dim t As String = grd.Name.Substring(3)
            grd.DataSource = dst.Tables(t)
            grd.DisplayLayout.Bands(0).Columns("ADD").Hidden = True

            If grd.Name = "grdStyle" Or grd.Name = "grdStyleColor" Or grd.Name = "grdStyleSize" Then
            Else
                grd.DisplayLayout.Bands(0).Columns.Add("XMIT_VERSION")
            End If

            If grd.Name = "grdpohdr" Then
                'If ASCMAIN1.USER_ID = "marilyn" Then
                Dim ci As Integer = -1
                For Each C As String In New String() {"Factory", "VandaleShipDate", "StyleNo", "StyleRef", "Hanger", "Style", "TotalQty", "FactoryCost", "TrimCost", "VandaleCost", "TotalAmount"}
                    ci += 1
                    grd.DisplayLayout.Bands(0).Columns(C).Header.SetVisiblePosition(ci, False)

                Next
                'End If
            End If

            With grd.DisplayLayout.Bands(0).Columns("XMIT_VERSION")
                .Header.VisiblePosition = 0
                .Header.Caption = "Version"
                .Width = 90
                .Header.Fixed = True
                .Header.Appearance.TextHAlign = HAlign.Center
                .CellAppearance.TextHAlign = HAlign.Center
            End With
            grd.DisplayLayout.Bands(0).Columns.Add("XMIT_ACTION")
            With grd.DisplayLayout.Bands(0).Columns("XMIT_ACTION")
                .Header.VisiblePosition = 1
                .Header.Caption = "Action"
                .Width = 80
                .Header.Fixed = True
                .Header.Appearance.TextHAlign = HAlign.Center
                .CellAppearance.TextHAlign = HAlign.Center
            End With

            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            Next

            grd.DisplayLayout.Bands(0).Columns("VAN_REF").Hidden = True
            If grd.Name = "grdStyleSize" Or grd.Name = "grdStyle" Or grd.Name = "grdStyleColor" Then
            Else
                grd.DisplayLayout.Bands(0).Columns(grdKeys(grd.Name.Substring(3))).Hidden = True
            End If
            If grd.Name = "grdpofabric" Then
                grd.DisplayLayout.Bands(0).Columns("Description").CellMultiLine = DefaultableBoolean.True
            End If
        Next

        Dim g As UltraWinGrid.UltraGridGroup
        With grdStyleSize.DisplayLayout.Bands(0)
            g = .Groups.Add("Style_Size")
            g.Header.Caption = ""
            .Columns("XMIT_VERSION").Group = g
            .Columns("XMIT_ACTION").Group = g

            With g.Header.Appearance
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.LightGreen
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With


            For I As Integer = 1 To 12
                Dim SZ As String = "SIZE_" & Format(I, "00")
                Dim SQ As String = "QTY_" & Format(I, "00")
                g = .Groups.Add(SZ)
                g.Header.Caption = "" ' Format(I, "00")
                .Columns(SZ).CellAppearance.TextHAlign = HAlign.Right
                .Columns(SZ).Header.Appearance.TextHAlign = HAlign.Right
                .Columns(SZ).Width = 50
                .Columns(SQ).Width = 50
                g.Width = 50
                .Columns(SZ).Group = g
                .Columns(SQ).Group = g
                .Columns(SQ).Level = 1

                With g.Header.Appearance
                    .TextHAlign = HAlign.Center
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightGreen
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
            Next
        End With



        gcols_excluded.Add("grdpohdr", New String() {"VAN_REF", "POKey", "PONo", "NoofColor", "NoofSize", "NoofCost", _
                           "Picture1", "Picture2", "PictureName1", "PictureName2", "SendStatus", "SendTime", _
                                                    "CreateUser", "CreateTime", "EditUser", "EditTime"})
        gcols_excluded.Add("grdpocolor", New String() {"VAN_REF", "POColorKey", "POKey", "CreateUser", "CreateTime", "EditUser", "EditTime"}) ', ""
        gcols_excluded.Add("grdpofactory", New String() {"VAN_REF", "POFactoryKey", "POKey"}) ', ""
        gcols_excluded.Add("grdposize", New String() {"VAN_REF", "POSizeKey", "POKey"}) ', ""
        gcols_excluded.Add("grdposizedtl", New String() {"VAN_REF", "POSizeKey", "POKey", "POColorKey"}) ', ""
        gcols_excluded.Add("grdpotrim", New String() {"VAN_REF", "POTrimKey", "POKey", "Picture", "PictureName", "TrimPOKey"}) ', ""
        gcols_excluded.Add("grdpofabric", New String() {"VAN_REF", "POFabrickey", "POKey"}) ', ""

        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTORDRX, grdPOTORDR2, grdPOTORDRB}
            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal

                    If grd.Name = "grdPOTORDRX" Then
                        If New String() {"APPROVED_DATE", "APPROVED_BY", "APPROVED_MESSAGE"}.Contains(gcol.Key) Then
                            .BackColor2 = Drawing.Color.LightGreen
                        End If
                        If New String() {"STYLE_NO", "STYLE_DESC", "STYLE_REF", "ORDR_DATE", "ORDR_CONF_DATE", "SHIP_DATE", "TOTAL_QTY", "TOTAL_AMT", "CREATED_BY", "FOLLOWED_BY", "FOLLOWED_BY_EMAIL"}.Contains(gcol.Key) Then
                            .BackColor2 = Drawing.Color.Orange
                        End If
                    End If
                End With
            Next
        Next

        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("VAN_REF").Header.Fixed = True
            .Columns("VANDALE_USER").Header.Fixed = True
            .Columns("DATE_RECEIVED").Header.Fixed = True
            .Columns("SENDNO").Header.Fixed = True
            .Columns("REMARKS").Header.Fixed = True
            .Columns("POKEY").Header.Fixed = True
            .Columns("PONO").Header.Fixed = True
            .Columns("PO_ORDER_NO").Header.Fixed = True
            .Columns("STATUS").Header.Fixed = True
        End With

        With grdPOTORDRB.DisplayLayout.Bands(0).Columns("XMIT_VERSION")
            ' .Width = 90
            .Header.Fixed = True
            .Header.Appearance.TextHAlign = HAlign.Center
            .CellAppearance.TextHAlign = HAlign.Center
        End With

        ASCMAIN1.Add_Value_List(grdPOTORDRX, "STATUS", Nothing, New String() {":", "W:Submitted", "X:Superceded", "A:Approved", "I:Imported", "R:Rejected", "D:Deleted"})

        ASCMAIN1.Add_Value_List(grdPOTORDRB, "STATUS", Nothing, New String() {":", "W:Submitted", "X:Superceded", "A:Approved", "I:Imported", "R:Rejected", "D:Deleted"})

        Create_Summary(grdPOTORDRX, "VAN_REF", "Count")

        Show_Filter(grdPOTORDRX, True)

        Set_Read_Only(grpHeaderData, True)
        grdPOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdPOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdPOTORDR2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

        Bind_Controls(grpImportOptions, "POTORDRA")
        Bind_Controls(grpHeaderData, "POTORDR1")

        options = optImport.ValueList

        cbeWHSE_CODE.DataSource = ASCDATA1.GetDataTable("Select WHSE_CODE,WHSE_DESC from ICTWHSE1 order by WHSE_CODE")


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Select"
                VAN_REF = Absx1.txtFor("VAN_REF").Text

                ASCMAIN1.sql = "Select POTORDRA.* from POTORDRA" & vbCrLf _
                    & " where PONO = (Select PONO from POTORDRA where VAN_REF = :PARM1)" & vbCrLf _
                    & "   and STATUS = '" & optDisplay.Value & "'"
                tblPOTORDRA = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {VAN_REF})
                If tblPOTORDRA.Rows.Count = 0 Then
                    EMsg &= vbCr & "No POs " & optDisplay.Text & " with Transmission No " & VAN_REF
                Else
                    Set_Tentative_Import_Values(tblPOTORDRA)
                    VAN_REF = tblPOTORDRA.Compute("MAX(VAN_REF)", "")
                    rowPOTORDRA = Fill_Record("POTORDRA", VAN_REF)
                    Dim row As DataRow = tblPOTORDRA.Rows.Find(VAN_REF)
                    rowPOTORDRA.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    PONO = rowPOTORDRA.Item("PONO")
                    PO_ORDER_NO = rowPOTORDRA.Item("PO_ORDER_NO") & ""

                    If PO_ORDER_NO = "" Then
                        ASCMAIN1.sql = "Select * from POTORDR1 where VEND_CODE = 'AT' and PO_REFERENCE = :PARM1"
                        Dim rowPONO As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "v", New String() {PONO})
                        If rowPONO IsNot Nothing Then
                            PO_ORDER_NO = rowPONO.Item("PO_ORDER_NO")
                        End If
                    End If

                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("POTORDRA", PONO) Then Exit Sub
                        If PO_ORDER_NO <> "" Then
                            If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                        End If
                    End If
                End If



            Case "View"
                VAN_REF = Absx1.txtFor("VAN_REF").Text

                ASCMAIN1.sql = "Select POTORDRA.* from POTORDRA" & vbCrLf _
                    & " where PONO = (Select PONO from POTORDRA where VAN_REF = :PARM1)"
                tblPOTORDRA = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", New String() {VAN_REF})
                If tblPOTORDRA.Rows.Count = 0 Then
                    EMsg &= vbCr & "No POs " & optDisplay.Text & " with Transmission No " & VAN_REF
                Else
                    Set_Tentative_Import_Values(tblPOTORDRA)
                    VAN_REF = tblPOTORDRA.Compute("MAX(VAN_REF)", "")
                    rowPOTORDRA = Fill_Record("POTORDRA", VAN_REF)
                    Dim row As DataRow = tblPOTORDRA.Rows.Find(VAN_REF)
                    rowPOTORDRA.Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                    PONO = rowPOTORDRA.Item("PONO")
                    PO_ORDER_NO = rowPOTORDRA.Item("PO_ORDER_NO") & ""
                End If

            Case "Update"
                If optImport.Value = "D" Or optImport.Value = "R" Then
                    If Absx1.txtFor("REASON_DESC").Text = "" Then
                        EMsg &= vbCr & "Reason is Required to " & optImport.Text
                    End If
                End If

                If optImport.Value = "R" Then
                    Dim EMAIL_TO As String = Absx1.txtFor("EMAIL_TO").Text.Trim
                    If EMAIL_TO = "" Then
                        EMsg &= vbCr & "email is Required to " & optImport.Text
                    Else
                        If Not ASCMAIN1.ValidateEmail(EMAIL_TO) Then
                            EMsg &= vbCr & "Invalid email address provided"
                        End If
                        If txtEMAIL_CC.Text <> "" Then
                            If Not ASCMAIN1.ValidateEmail(txtEMAIL_CC.Text) Then
                                EMsg &= vbCr & "Invalid cc address provided"
                            End If
                        End If
                    End If
                End If

                If optImport.Value = "A" Then
                    EMsg &= ValidatePO()

                    If EMsg <> "" Then
                        If MsgBox("There are Import Errors:" & vbCrLf & EMsg _
                                  & vbCrLf & vbCrLf & "Do you want to Approve anyway?", _
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            EMsg = ""
                        Else

                            EMsg = vbCr & "Import Errors prevented Approval"
                        End If
                        ' DON WANTS TO SEE THESE MESSAGES, BUT STILL BE ABLE TO APPROVE
                    End If
                End If

                If optImport.Value = "I" Then
                    EMsg &= ValidatePO()

                End If

                If EMsg = "" Then
                    If MsgBox("OK to " & optImport.Text & " this Transmission", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "FOB Cost Sheet"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Refresh"
                Load_POTORDRX()

            Case "Select"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

                ' Me.Close() - if we close, then MC will get bounced out after clicking Done which wastes her time

            Case "Fix Styles"
                Fix_Styles()

            Case "Get PDFs"
                Get_PDFs()

            Case "FOB Cost Sheet"
                Produce_XLS()
        End Select

    End Sub


    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    ' .Items("Select").Settings.Enabled = not_iScreenMode

                    '.Items("Update").Settings.Enabled = iScreenMode
                    '.Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = ScreenMode And (EntryMode <> "V")
                    .Items("Cancel").Visible = ScreenMode And (EntryMode <> "V")
                    .Items("View").Visible = False
                    .Items("Done").Visible = ScreenMode And (EntryMode = "V")

                    .Items("Fix Styles").Visible = ASCMAIN1.Running_in_VS And Not ScreenMode
                    .Items("Get PDFs").Visible = ASCMAIN1.Running_in_VS And Not ScreenMode

                    .Items("FOB Cost Sheet").Visible = ScreenMode

                End With
                .Groups("Import Options").Visible = ScreenMode
                .Groups("Display Options").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl.Panel1Collapsed = Not ScreenMode

        lblStatus.Visible = ScreenMode
        ' grdPOTORDRX.Visible = Not ScreenMode
        splPOTPORDA.Visible = ScreenMode

        grdPOTORDRB.Visible = ScreenMode

        If ScreenMode Then
            grdPOTORDRX.Dock = DockStyle.None
            grdPOTORDRX.Parent = tabPO.Tabs("Transmission History").TabPage
            grdPOTORDRX.Dock = DockStyle.Fill

            Set_Read_Only_for_ctl(chkShowChangesOnly, False)

            '  optImport.ValueList = options

            Dim VLI As ValueListItem = optImport.ValueList.ValueListItems(0)
            If EntryMode = "V" Then
                VLI.DataValue = "N"
                VLI.DisplayText = "Just Viewing"
                optImport.Value = "N"
                VLI = optImport.ValueList.ValueListItems(2)
                optImport.ValueList.ValueListItems.Remove(VLI)
            Else
                If optDisplay.Value = "W" Then
                    VLI.DataValue = "A"
                    VLI.DisplayText = "Approve for Entry"
                    '    optImport.ValueList.ValueListItems.RemoveAt(1)
                    optImport.Value = "A"
                ElseIf optDisplay.Value = "A" Then
                    VLI.DataValue = "I"
                    VLI.DisplayText = "Import"
                    '   optImport.ValueList.ValueListItems.RemoveAt(0)
                    optImport.Value = "I"
                End If
            End If


            Toggle_Message_Block()

        Else
            Clear_Record()
            grdPOTORDRX.Dock = DockStyle.None
            grdPOTORDRX.Parent = spl.Panel2
            grdPOTORDRX.Dock = DockStyle.Fill
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"POTORDRX", "POTORDRA", "POTORDRB", "POTORDR1", "POTORDR2", "POTORDXR", "differences", "Images", "PDFs"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        txtAPPROVED_MESSAGE.Text = ""
        txtREASON_DESC.Text = ""
        Absx1.txtFor("VAN_REF").Text = ""

        Load_POTORDRX()

        Setup_Display()

        chkShowChangesOnly.Checked = False
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Load_POTORDRX(tblPOTORDRA)

        Load_POTORDRB(rowPOTORDRA)

        Dim EMAIL_CC As String = ""
        VEND_CODE = "AT" ' HARD-CODED FOR NOW - NEED A WAY TO DIFFERENTIATE DIFFERENT VENDORS POTORDRA RECORDS IN THE FUTURE
        Dim rowPOTVEND1 As DataRow = LookUp("POTVEND1", VEND_CODE)
        If rowPOTVEND1 IsNot Nothing Then
            EMAIL_CC = rowPOTVEND1.Item("EMAIL_CC")
        End If
        If optDisplay.Value = "A" Then ' Approved, and Awaiting Import
            Dim APPROVED_BY As String = rowPOTORDRA.Item("APPROVED_BY")
            Dim rowASTUSER1_approved_by As DataRow = LookUp("ASTUSER1", APPROVED_BY)
            If rowASTUSER1_approved_by IsNot Nothing Then
                Dim USER_EMAIL_approved_by As String = rowASTUSER1_approved_by.Item("USER_EMAIL") & ""
                If USER_EMAIL_approved_by <> "" Then
                    If USER_EMAIL_approved_by <> EMAIL_CC Then
                        If EMAIL_CC <> "" Then
                            EMAIL_CC &= ";"
                        End If
                        EMAIL_CC &= USER_EMAIL_approved_by
                    End If
                End If
            End If
        End If

        rowPOTORDRA.Item("EMAIL_CC") = EMAIL_CC

        VAN_REF_last_AI = ""

        ASCMAIN1.sql = "Select MAX(VAN_REF) from POTORDRA where PONO = :PARM1 and STATUS in ('A','R','I') and SUPERCEDED_BY is Null and VAN_REF <> '" & VAN_REF & "'"
        If EntryMode = "V" Then
            ASCMAIN1.sql &= " and VAN_REF < '" & VAN_REF & "'"
        End If
        VAN_REF_previous = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PONO})
        If VAN_REF_previous <> "" Then
            If tblPOTORDRA.Rows.Find(VAN_REF_previous) Is Nothing Then


                ASCMAIN1.sql = "Select POTORDRA.* from POTORDRA where VAN_REF = :PARM1"
                Dim row_previous As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {VAN_REF_previous})
                Dim row As DataRow = tblPOTORDRA.Rows.Add(row_previous.ItemArray)
                row.AcceptChanges()
                row.Item("SUPERCEDED_BY") = VAN_REF

                Load_POTORDRB(row_previous)
                chkShowChangesOnly.Visible = True


                If row_previous.Item("STATUS") = "R" Then
                    ASCMAIN1.sql = "Select MAX(VAN_REF) from POTORDRA where PONO = :PARM1 and STATUS in ('A','I') and SUPERCEDED_BY is Null and VAN_REF <> '" & VAN_REF & "' and VAN_REF <> '" & VAN_REF_previous & "'"
                    VAN_REF_last_AI = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {PONO})
                    If VAN_REF_last_AI <> "" Then
                        ASCMAIN1.sql = "Select POTORDRA.* from POTORDRA where VAN_REF = :PARM1"
                        Dim row_last_AI As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {VAN_REF_last_AI})
                        row = tblPOTORDRA.Rows.Add(row_last_AI.ItemArray)
                        row.AcceptChanges()
                        row.Item("SUPERCEDED_BY") = VAN_REF

                        Load_POTORDRB(row_last_AI)
                    End If
                End If
            End If

        Else
            chkShowChangesOnly.Visible = False
        End If

        Sort_grdColumns(grdPOTORDRB, "VAN_REF".ToLower)

        For Each TABLE_NAME As String In poTables
            Fill_Records(TABLE_NAME, VAN_REF)
            
            If VAN_REF_previous <> "" Then
                Fill_Records(TABLE_NAME, VAN_REF_previous, False)

                If VAN_REF_last_AI <> "" Then
                    Fill_Records(TABLE_NAME, VAN_REF_last_AI, False)
                End If

                ' Ensure that we have Curr, Prev and Last records for each Key of each Table
                Dim blnSkipTable As Boolean = False
                If TABLE_NAME = "pocolor" Or TABLE_NAME = "potrim" Or TABLE_NAME = "posize" Then blnSkipTable = True
                If blnSkipTable Then
                Else


                    Dim TKs As New List(Of Int32)
                    Dim TK As String = grdKeys(TABLE_NAME)
                    Dim KEYCOLUMN As String = grdKeys(TABLE_NAME)
                    For Each ROW As DataRow In dst.Tables(TABLE_NAME).Select()
                        Dim TKV As Int32 = Val(ROW.Item(TK))
                        If Not TKs.Contains(TKV) Then
                            TKs.Add(TKV)
                            Dim VR As String = ROW.Item("VAN_REF")

                            For Each VRO As String In New String() {VAN_REF, VAN_REF_previous, VAN_REF_last_AI}
                                If VRO <> "" And VRO <> VR Then ' NEED TO RESOLVE THAT AT CHANGES POCOLORKEY WHEN THEY MAKE A CHANGE TO POCOLOR
                                    Dim SQLW As String = "VAN_REF = '" & VRO & "' and " & TK & " = " & CStr(TKV) & ""
                                    If TABLE_NAME = "pocolor" Then
                                        Dim ColorCode As String = ROW.Item("ColorCode") & ""
                                        Dim c2 As Integer = 2
                                        If ColorCode.StartsWith("# ") Then
                                            c2 = 3
                                        End If
                                        SQLW = "VAN_REF = '" & VRO & "' and (ColorCode = '#  " & Mid(ColorCode, c2) & "' or ColorCode = '# " & Mid(ColorCode, c2) & "' or ColorCode = '" & ColorCode & "')"

                                        Dim Pantone As String = ROW.Item("Pantone") & ""
                                        If Pantone <> "" Then
                                            SQLW &= " and Pantone = '" & ROW.Item("Pantone") & "'"
                                        End If

                                    End If
                                    If TABLE_NAME = "potrim" Then
                                        SQLW = "VAN_REF = '" & VRO & "' and ItemNo = '" & ROW.Item("ItemNo") & "'"
                                    End If
                                    If dst.Tables(TABLE_NAME).Select(SQLW).Length = 0 Then
                                        '  If TABLE_NAME = "posizedtl" Then Stop
                                        Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
                                        row2.Item("VAN_REF") = VRO
                                        row2.Item("POKey") = ROW.Item("POKey")
                                        row2.Item(KEYCOLUMN) = ROW.Item(KEYCOLUMN)
                                        If TABLE_NAME = "posizedtl" Then
                                            row2.Item("POColorKey") = ROW.Item("POColorKey")
                                        End If
                                        row2.Item(TK) = TKV
                                        If VR = VAN_REF Then row2.Item("ADD") = "A"
                                        dst.Tables(TABLE_NAME).Rows.Add(row2)
                                    Else
                                        dst.Tables(TABLE_NAME).Select(SQLW)(0).Item(KEYCOLUMN) = ROW.Item(KEYCOLUMN)
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
            Sort_grdColumns(grdATs(TABLE_NAME), grdKeys(TABLE_NAME) & ",VAN_REF".ToLower)
            grdATs(TABLE_NAME).DisplayLayout.Bands(0).Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        Next

        ' Get Color into size details
        For Each row As DataRow In dst.Tables("posizedtl").Select("")
            Dim POColorKey As Int64 = Val(row.Item("POColorKey") & "")
            Dim rowC() As DataRow = dst.Tables("pocolor").Select("VAN_REF = '" & row.Item("VAN_REF") & "' and POColorKey = " & CStr(POColorKey))
            If rowC.Length = 1 Then
                row.Item("ColorCode") = rowC(0).Item("ColorCode")
            End If
        Next

        Get_List_of_Columns_with_Differences()
        grdpohdr.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdpocolor.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdposize.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdposizedtl.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdpofactory.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdpotrim.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdpofabric.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        rowpohdr = dst.Tables("pohdr").Select("VAN_REF = '" & VAN_REF & "'")(0)

        rowPOTORDRA.Item("EMAIL_TO") = rowpohdr.Item("FollowByEmail")

        If PO_ORDER_NO <> "" Then
            rowPOTORDRA.Item("PO_ORDER_NO") = PO_ORDER_NO
        End If

        lblPOREF.Text = "PO Ref (" & rowPOTORDRA.Item("POKey") & ")"

        gcols.Clear()

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdpohdr, grdpocolor, grdpofactory, grdposize, grdposizedtl, grdpotrim, grdpofabric}
            grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, UltraWinGrid.AutoResizeColumnWidthOptions.All)
            Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)
            Dim cols As New List(Of String)

            For Each row As DataRow In tbl.Select("")
                For Each dcol As DataColumn In tbl.Columns
                    Dim col As String = dcol.ColumnName

                    If Not cols.Contains(col) Then
                        If row.Item(col) & "" <> "" Then
                            cols.Add(col)
                        End If
                    End If
                Next
            Next

            gcols.Add(grd.Name, cols) ' maybe for later when we want to toggle all/meaningful columns

        Next

        Toggle_grd_Columns()



        If VAN_REF_previous = "" Then
            lblStatusXMIT.Text = "Original"
        Else
            lblStatusXMIT.Text = "Revision"
        End If


        If PO_ORDER_NO = "" Then
            rowPOTORDR1 = Nothing
            tabPO.Tabs("VAN PO").Visible = False

            lblStatus.Text = "A New PO will be created"
        Else
            rowPOTORDR1 = Fill_Record("POTORDR1", PO_ORDER_NO)
            Fill_Records("POTORDR2", PO_ORDER_NO)
            tabPO.Tabs("VAN PO").Visible = True
            tabPO.Tabs("VAN PO").Text = "VAN PO " & PO_ORDER_NO

            lblStatus.Text = "PO will be edited"
        End If

        If rowpohdr.Item("POStatus") & "" = "C" Or rowpohdr.Item("SendRemarks") & "" = "Cancel PO" Then
            lblStatus.Text = "AT: Cancel this PO"
            lblStatusXMIT.Text = "Cancel"
        End If

        For Each TABLE_NAME As String In poTables
            Sort_grdColumns(grdATs(TABLE_NAME), grdKeys(TABLE_NAME) & ",VAN_REF".ToLower)
            grdATs(TABLE_NAME).DisplayLayout.Bands(0).Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
        Next

        Dim STYLE_CODE As String = rowpohdr.Item("StyleNo") & ""
        STYLE_CODE = Replace(STYLE_CODE, "-", "")

        Dim rowpofabrics() As DataRow = dst.Tables("pofabric").Select("VAN_REF = '" & VAN_REF & "'")

        dst.Tables("Style").Rows.Clear()
        Dim rowStyle As DataRow = Fill_Record("Style", STYLE_CODE)
        If rowStyle IsNot Nothing Then rowStyle.Item("VAN_REF") = "0000000000"
        rowStyle = dst.Tables("Style").NewRow
        rowStyle.Item("VAN_REF") = rowpohdr.Item("VAN_REF")
        rowStyle.Item("STYLE_DESC") = rowpohdr.Item("Style")
        Dim SIZE_SCALE As String = ""
        rowStyle.Item("SIZE_SCALE") = SIZE_SCALE
        If rowpofabrics.Length = 1 Then
            rowStyle.Item("STYLE_MATL_DESC") = rowpofabrics(0).Item("Description")
        End If
        rowStyle.Item("CARTON_PACK_QTY") = 0 ' NEEDS FIXING
        rowStyle.Item("INNER_PACK_QTY") = Val(rowpohdr.Item("NoofColor") & "") * Val(rowpohdr.Item("NoofSize") & "")
        rowStyle.Item("SUB_UNIT_PACK_QTY") = rowpohdr.Item("NoofPack")
        rowStyle.Item("FACTORY_CODE") = rowpohdr.Item("Factory")
        dst.Tables("Style").Rows.Add(rowStyle)
        Sort_grdColumns(grdStyle, "VAN_REF".ToLower)


        dst.Tables("StyleColor").Rows.Clear()
        Fill_Records("StyleColor", STYLE_CODE)
        For Each row As DataRow In dst.Tables("StyleColor").Select("")
            row.Item("VAN_REF") = "0000000000"
        Next
        For Each row As DataRow In dst.Tables("pocolor").Select("")
            Dim rowStyleColor As DataRow = dst.Tables("StyleColor").NewRow
            rowStyleColor.Item("VAN_REF") = row.Item("VAN_REF") & ""
            Dim ColorCode As String = row.Item("ColorCode") & ""
            rowStyleColor.Item("COLOR_CODE") = ColorCode
            Dim ColorName As String = row.Item("ColorName") & ""
            rowStyleColor.Item("STYLE_COLOR_DESC") = ColorName
            dst.Tables("StyleColor").Rows.Add(rowStyleColor)
        Next
        Sort_grdColumns(grdStyleColor, "COLOR_CODE," & "VAN_REF".ToLower)


        dst.Tables("StyleSize").Rows.Clear()
        Fill_Records("StyleSize", STYLE_CODE)
        For Each row As DataRow In dst.Tables("StyleSize").Select("")
            row.Item("VAN_REF") = "0000000000"
        Next
        Dim rowStyleSize As DataRow = dst.Tables("StyleSize").NewRow
        rowStyleSize.Item("VAN_REF") = "0000000000"
        dst.Tables("StyleSize").Rows.Add(rowStyleSize)


        For Each row As DataRow In dst.Tables("posize").Select("", "VAN_REF,POSizeKey")
            Dim Item As String = row.Item("Item") & ""
            If Item = "" Then
                Dim Color As String = Trim(row.Item("Color") & "")
                If Color = "Size:" Then
                    rowStyleSize = dst.Tables("StyleSize").NewRow
                    rowStyleSize.Item("VAN_REF") = row.Item("VAN_REF")
                    Dim Size As String = Trim(row.Item("Size") & "")
                    If Size.EndsWith("|") Then
                        Size = Size.Substring(0, Size.Length - 1)
                    End If
                    Dim Sizes() As String = Size.Split("|")
                    For i As Integer = 1 To 12
                        Dim SZ As String = "SIZE_" & Format(i, "00")
                        grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Hidden = Not (i <= Sizes.Length)
                        grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Width = 50
                        If (i <= Sizes.Length) Then
                            rowStyleSize.Item(SZ) = Sizes(i - 1)
                        End If
                    Next
                    If ASCMAIN1.Running_in_VS Then
                        dst.Tables("StyleSize").Rows.Add(rowStyleSize)
                    Else
                        Try
                            dst.Tables("StyleSize").Rows.Add(rowStyleSize)
                        Catch ex As Exception

                        End Try
                    End If
                Else
                    Dim Remark As String = Trim(row.Item("Remark") & "")
                    If Remark <> "" Then
                        Dim Size As String = Trim(row.Item("Size") & "")
                        If Size.EndsWith("|") Then
                            Size = Size.Substring(0, Size.Length - 1)
                        End If
                        Dim Sizes() As String = Size.Split("|")

                        For i As Integer = 1 To Sizes.Length
                            Dim QZ As String = "QTY_" & Format(i, "00")
                            If Sizes(i - 1) <> "" Then rowStyleSize.Item(QZ) = Sizes(i - 1)
                        Next
                    End If

                End If
            End If
        Next



        'Dim StyleBySize As Boolean = (Val(rowpohdr.Item("NoofStyle") & "") > 1)
        Dim StyleBySize As Boolean = (rowpohdr.Item("StyleBySize") & "" = "Y")
        btnCreateStyle.Visible = Not StyleBySize

        For Each VRrow As DataRow In ASCDATA1.SelectDistinct("posizedtl", New String() {"VAN_REF"}).Select("", "")

            Dim VR As String = VRrow.Item(0)

            Dim SZPPK As New Dictionary(Of String, Integer)
            Dim SZStyle As New Dictionary(Of String, String)
            For Each row As DataRow In dst.Tables("posizedtl").Select("VAN_REF = '" & VR & "'", "VAN_REF,POSizeKey")
                Dim Size As String = row.Item("Size") & ""
                If Not SZPPK.ContainsKey(Size) Then
                    SZPPK.Add(Size, Val(row.Item("PrePack") & ""))
                    SZStyle.Add(Size, row.Item("Style") & "")
                End If
            Next

            If SZPPK.Count <> 0 Then
                rowStyleSize = dst.Tables("StyleSize").NewRow
                rowStyleSize.Item("VAN_REF") = VR
                rowStyleSize.Item("STYLE_CODE") = STYLE_CODE
                For i As Integer = 1 To 12
                    Dim SZ As String = "SIZE_" & Format(i, "00")
                    If VR = VAN_REF Then

                        If StyleBySize And i <= SZPPK.Count Then
                            Dim Style4Size As String = SZStyle(SZStyle.Keys(i - 1))
                            grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Header.Caption = "EZ"
                            grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Header.ToolTipText = Style4Size
                            grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Header.Tag = SZPPK.Keys(i - 1)
                        Else
                            grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Header.Caption = ""
                            grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Header.ToolTipText = ""
                        End If
                        grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Hidden = Not (i <= SZPPK.Count)
                        grdStyleSize.DisplayLayout.Bands(0).Groups(SZ).Width = 50
                    End If
                    If (i <= SZPPK.Count) Then
                        rowStyleSize.Item(SZ) = SZPPK.Keys(i - 1)
                        Dim QZ As String = "QTY_" & Format(i, "00")
                        rowStyleSize.Item(QZ) = SZPPK(SZPPK.Keys(i - 1))
                    End If
                Next

                If ASCMAIN1.Running_in_VS Then
                    dst.Tables("StyleSize").Rows.Add(rowStyleSize)
                Else
                    Try
                        dst.Tables("StyleSize").Rows.Add(rowStyleSize)
                    Catch ex As Exception

                    End Try
                End If

            End If
        Next

        Sort_grdColumns(grdStyleSize, "VAN_REF".ToLower)



        If StyleBySize Then
            grdStyleSize.DisplayLayout.Bands(0).Groups(0).Header.Caption = "Click Size Header->"
        Else
            grdStyleSize.DisplayLayout.Bands(0).Groups(0).Header.Caption = ""
        End If

        ASCDATA1.DeleteRows(dst.Tables("StyleSize"), "isnull(SIZE_01,'') = ''")
        ' the loop below throws the error I have been looking to demonstrate with van_ref 346
        'For Each row As DataRow In dst.Tables("StyleSize").Select("")
        '    If row.Item("SIZE_01") & "" = "" Then
        '        row.Delete()
        '    End If
        'Next


        Get_Images()

        chkAutoSendEmail.Checked = True

        grpWHSE_CODE.Visible = (optDisplay.Value = "A")
        If optDisplay.Value = "A" Then

            cbeWHSE_CODE.Value = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
        End If

        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Load_POTORDRB(row As DataRow)
        Dim rowPOTORDRB As DataRow = dst.Tables("POTORDRB").NewRow
        With rowPOTORDRB
            .Item("VAN_REF") = row.Item("VAN_REF")
            .Item("DATE_RECEIVED") = row.Item("DATE_RECEIVED")
            .Item("STATUS") = row.Item("STATUS")
            .Item("LAST_DATE") = row.Item("LAST_DATE")
            .Item("LAST_OPER") = row.Item("LAST_OPER")
            If row.Item("STATUS") & "" = "A" Then
                .Item("MESSAGE") = row.Item("APPROVED_MESSAGE")
            ElseIf row.Item("STATUS") & "" = "R" Then
                .Item("MESSAGE") = row.Item("REASON_DESC")
            ElseIf row.Item("STATUS") & "" = "I" Then
                .Item("MESSAGE") = "Imported"
            End If

        End With
        dst.Tables("POTORDRB").Rows.Add(rowPOTORDRB)
    End Sub
    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()

        If rowPOTORDRA.Item("INIT_OPER") & "" = "" Then
            rowPOTORDRA.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTORDRA.Item("INIT_DATE") = DATETIME_STAMP
        End If

        rowPOTORDRA.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowPOTORDRA.Item("LAST_DATE") = DATETIME_STAMP

        Dim commitMsg As String = ""

        Dim sqlw As String = "SUPERCEDED_BY = '" & VAN_REF & "'"
        If optImport.Value = "A" Then
            sqlw &= " and (STATUS = 'X' or STATUS = 'A' or STATUS = 'I')"
        ElseIf optImport.Value = "I" Then
            sqlw &= " and (STATUS = 'X' or STATUS = 'A' or STATUS = 'I')"
        ElseIf optImport.Value = "R" Then
            sqlw &= " and STATUS = 'X'"
        End If

        Dim sqlV As String = "VAN_REF = '" & VAN_REF & "'"

        For Each row As DataRow In dst.Tables("POTORDRX").Select(sqlw) '  & "and " & sqlV)
            Dim STATUS = row.Item("STATUS")
            If STATUS = "X" AndAlso row.Item("INIT_OPER") & "" = "" Then
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("INIT_DATE") = DATETIME_STAMP
            End If
            row.Item("LAST_OPER") = ASCMAIN1.USER_ID
            row.Item("LAST_DATE") = DATETIME_STAMP
            Dim row2 As DataRow = dst.Tables("POTORDRA").NewRow
            For i As Integer = 0 To dst.Tables("POTORDRA").Columns.Count - 1
                Dim dc As DataColumn = dst.Tables("POTORDRA").Columns(i)
                Dim dcname As String = dc.ColumnName
                row2.Item(dcname) = row.Item(dcname)
            Next
            dst.Tables("POTORDRA").Rows.Add(row2)
            row2.AcceptChanges()
            row2.SetModified()
        Next

        rowPOTORDRA.Item("STATUS") = optImport.Value

        If optImport.Value = "A" Then

            commitMsg = " - PO " & PO_ORDER_NO & " has been Approved"

            rowPOTORDRA.Item("APPROVED_DATE") = DATETIME_STAMP
            rowPOTORDRA.Item("APPROVED_BY") = ASCMAIN1.USER_ID
            rowPOTORDRA.Item("APPROVED_MESSAGE") = txtAPPROVED_MESSAGE.Text

        ElseIf optImport.Value = "R" Then

            commitMsg = " - PO " & PO_ORDER_NO & " has been Rejected"

        ElseIf optImport.Value = "I" Then

            dst.Tables("POTORDR1").Rows.Clear()
            dst.Tables("POTORDR2").Rows.Clear()

            If rowpohdr.Item("POStatus") & "" = "C" Or rowpohdr.Item("SendRemarks") & "" = "Cancel PO" Then
                ' totally handled by Dependent Updates
            Else
                '    Dim VEND_CODE As String = "AT"
                Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
                Dim TERM_CODE As String = rowAPTVEND1.Item("TERM_CODE") & ""

                Dim AT_Style_Description As String = rowpohdr.Item("Style") & ""

                '   Dim rowpohdr As DataRow = dst.Tables("pohdr").Rows(0)
                Dim STYLE_CODE As String = rowpohdr.Item("StyleNo") & ""
                STYLE_CODE = STYLE_CODE.Replace("-", "")

                Dim FACTORY_CODE As String = rowpohdr.Item("Factory") & ""
                Dim rowICTFACT1 As DataRow = LookUp("ICTFACT1", FACTORY_CODE)

                Dim PO_CONTACT As String = rowpohdr.Item("FollowBy") & ""

                Dim PO_SHIP_VIA As String = ""
                Dim PORT_CODE_ORIG As String = ""
                Dim WHSE_CODE As String = cbeWHSE_CODE.Value '  "NJE"

                Dim PO_DATE_ORDERED As Date = rowpohdr.Item("OrderDate")

                Dim PO_DATE_SHIP_BY As Date = rowpohdr.Item("VandaleShipDate")

                Dim ETD_to_ETA As Int32 = Calculate_ETD_to_ETA(PORT_CODE_ORIG, WHSE_CODE, PO_SHIP_VIA)
                Dim PO_DATE_ETA As Date = PO_DATE_SHIP_BY.AddDays(ETD_to_ETA)

                Dim COST2 As Decimal = 0
                Dim COST2_ctr As Integer = 0
                For Each rowpofactory As DataRow In dst.Tables("pofactory").Select(sqlV)
                    COST2_ctr += 1
                    If COST2_ctr <> 1 Then ' mc says do not include the 1st line
                        COST2 += Val(rowpofactory.Item("MakerCost") & "")
                    End If
                Next

                PO_ORDER_LNO = 0
                StyleColors.Clear()
                ' Dim NoofStyle As Integer = Val(rowpohdr.Item("NoofStyle") & "")
                Dim StyleBySize As Boolean = (rowpohdr.Item("StyleBySize") & "" = "Y")
                ' WJZ 09/30/21 FOR ME POS, AT IS SENDING STYLEBYSIZE = N - NEED TO TAKE THIS UP WITH EDMUND
                If PONO.StartsWith("ME") Then StyleBySize = False ' True

                Dim PO_COST_VCOST_DZ As Decimal = Val(rowpohdr.Item("FactoryCost") & "") - COST2
                Dim PO_COST_MATLS_DZ As Decimal = 0

                Dim PO_COST_OTHER_DZ As Decimal = Val(rowpohdr.Item("TrimCost") & "") + COST2
                '  Dim PO_COST_COMM_PCT As Decimal = Val(rowpohdr.Item("Commission") & "")
                ' ANNA SAYS TO USE HER PARAMETER AND TO IGNORE THE VALUE IN THE PO FOR COMM%
                ' AT WAS SENDING 3.0 WHEN ANNA HAD 2.5 IN THE PARAMETER TABLE
                Dim PO_COST_COMM_PCT As Decimal = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_COMM") & "")
                'Changed from 2.5 to 2.0 per Anna on 7/7/19 - WR.
                'Changed from 2.0 to 2.5 per Anna on 7/17/19 - WR.
                'PO_COST_COMM_PCT = 2.5 ' PARM FILE IS 0 FOR COMM FOR ALL VENDORS, BUT HARD CODED FOR AT
                'Changed from 2.0 to 0.0 Per Anna 8/16/24 - WR.
                PO_COST_COMM_PCT = 0.0 ' WJZ 10/5/21 - to keep consistent with POFORDR1 for AT

                Dim PO_COST_COMM_PCT_ADD As Decimal = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_BUFFER") & "")
                PO_COST_COMM_PCT_ADD = 1 ' 1 HARD CODED FOR AT

                If PO_ORDER_NO = "" Then ' New PO

                    If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
                        PO_ORDER_NO = ASCMAIN1.Next_Control_No("PO_ORDER_NO")
                    Else
                        PO_ORDER_NO = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO")
                    End If
                    rowPOTORDRA.Item("PO_ORDER_NO") = PO_ORDER_NO

                    commitMsg = " - PO " & PO_ORDER_NO & " has been Created"

                    rowPOTORDR1 = dst.Tables("POTORDR1").NewRow
                    With rowPOTORDR1
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        .Item("VEND_CODE") = VEND_CODE
                        Dim VEND_NAME As String = rowAPTVEND1.Item("VEND_NAME") & ""
                        If VEND_NAME.Length > 35 Then
                            VEND_NAME = Mid(VEND_NAME, 1, 35)
                        End If
                        .Item("VEND_NAME") = VEND_NAME
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("LAST_DATE") = DATETIME_STAMP
                        .Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date ' PO_DATE_ORDERED
                        .Item("PO_REFERENCE") = PONO
                        .Item("WHSE_CODE") = WHSE_CODE
                        .Item("PO_STATUS") = "O"
                        .Item("PO_DATE_SHIP_BY") = PO_DATE_SHIP_BY
                        .Item("PO_DATE_ETA") = PO_DATE_ETA
                        .Item("FOB_CMT") = "F"
                        .Item("FACTORY_CODE") = FACTORY_CODE
                        .Item("PO_CONTACT") = PO_CONTACT
                        ' .Item("PO_NOTES") = AT_Style_Description
                        .Item("TERM_CODE") = TERM_CODE
                        '.Item("PO_MESSAGE") = "?"
                        '  .Item("PO_COMM_PCT") = 0
                        .Item("PO_HAS_PPK") = "0"
                    End With
                    dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

                    For Each rowpocolor As DataRow In dst.Tables("pocolor").Select(sqlV & " and ISNULL(ADD,'?') <> 'A'")
                        Dim COLOR_CODE As String = rowpocolor.Item("ColorCode")
                        If COLOR_CODE.StartsWith("#") Then COLOR_CODE = Trim(COLOR_CODE.Substring(1))

                        Dim POColorKey As Int64 = Val(rowpocolor.Item("POColorKey") & "")
                        Dim STYLE_CODEs As List(Of String) = Get_Style_Codes(STYLE_CODE, POColorKey)

                        For Each STYLE_CODE_to_check As String In STYLE_CODEs
                            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_to_check)
                            Dim SUB_UNIT_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "")
                            Dim StyleColor As String = STYLE_CODE_to_check & ":" & COLOR_CODE


                            Dim PO_QTY_ORD As Int64 = Val(rowpocolor.Item("OrderQty") & "") * 12 / SUB_UNIT_PACK_QTY
                            If StyleBySize Then ' If NoofStyle > 1 Then
                                Dim rowsposizedtl() As DataRow = dst.Tables("posizedtl").Select(sqlV & " and POColorKey = " & CStr(POColorKey) & " and Style = '" & STYLE_CODE_to_check & "'")
                                If rowsposizedtl.Length = 1 Then
                                    PO_QTY_ORD = Val(rowsposizedtl(0).Item("Qty") & "") * 12 / SUB_UNIT_PACK_QTY
                                Else
                                    If ASCMAIN1.Running_in_VS Then Stop
                                    Throw New Exception("Cannot Correlate New Size Detail Record with Style " & STYLE_CODE_to_check & ", POColorKey " & CStr(POColorKey))
                                End If
                            End If

                            Dim OrderUnit As String = rowpocolor.Item("OrderUnit") & ""

                            Dim rowPOTORDR2 As DataRow

                            If StyleColors.ContainsKey(StyleColor) Then
                                rowPOTORDR2 = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, StyleColors(StyleColor)})
                                With rowPOTORDR2
                                    .Item("PO_QTY_ORD") += PO_QTY_ORD
                                    .Item("PO_QTY_OPN") += PO_QTY_ORD
                                End With
                            Else
                                rowPOTORDR2 = Update_Record_POTORDR2_Add(rowICTSTYL1, OrderUnit, STYLE_CODE_to_check, COLOR_CODE, PO_QTY_ORD)
                            End If

                            Update_Record_POTORDR2(rowPOTORDR2, PO_DATE_SHIP_BY, PO_DATE_ETA, _
                               PO_COST_VCOST_DZ, PO_COST_COMM_PCT_ADD, PO_COST_OTHER_DZ, PO_COST_COMM_PCT, PO_COST_MATLS_DZ, _
                               SUB_UNIT_PACK_QTY) ' always do these updates

                        Next
                    Next

                Else ' Editing an existing PO

                    EntryMode = "E" ' SO THAT POCMAIN1.Check_Changed_Fields can have the appropriate value

                    rowPOTORDR1 = Fill_Record("POTORDR1", PO_ORDER_NO)
                    If rowPOTORDR1.Item("PO_STATUS") & "" <> "O" Then
                        Throw New Exception("Cannot change a PO that is not Open")
                    End If

                    Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")

                    If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
                        rowPOTORDR1.Item("PO_PRINTED_IND") = "0"
                        rowPOTORDR1.Item("PO_XMIT_IND") = "0"
                        PO_HDR_CTR_REV += 1
                        rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV
                        rowPOTORDR1.Item("PO_REVISION_NOTE") = ""
                    End If

                    With rowPOTORDR1
                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        .Item("LAST_DATE") = DATETIME_STAMP
                        ' .Item("PO_DATE_ORDERED") = PO_DATE_ORDERED
                        .Item("FACTORY_CODE") = FACTORY_CODE
                        .Item("PO_CONTACT") = PO_CONTACT
                        ' .Item("PO_NOTES") = AT_Style_Description
                        .Item("PO_DATE_CANCEL") = PO_DATE_SHIP_BY.AddDays(60)
                    End With

                    ' MOVING THIS DOWN

                    'If TAC.POCMAIN1.Check_Changed_Fields(Me, rowPOTORDR1) Then
                    '    ' WHAT ABOUT HDR REV CTR - SEE ABOVE
                    '    'Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                    '    'rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV + 1
                    'End If

                    Fill_Records("POTORDR2", PO_ORDER_NO)

                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("", "PO_ORDER_LNO")
                        Dim LNO As Int32 = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
                        Dim STYLE_CODE_po As String = rowPOTORDR2.Item("STYLE_CODE")
                        Dim COLOR_CODE As String = rowPOTORDR2.Item("COLOR_CODE")
                        StyleColors.Add(STYLE_CODE_po & ":" & COLOR_CODE, -1 * LNO)
                        PO_ORDER_LNO = LNO

                        If Format(rowPOTORDR1.Item("PO_DATE_SHIP_BY"), "yyyyMMdd") <> Format(PO_DATE_SHIP_BY, "yyyyMMdd") Then
                            If rowPOTORDR2.Item("PO_ORIG_DATE_SHIP_BY") & "" = "" Then
                                rowPOTORDR2.Item("PO_ORIG_DATE_SHIP_BY") = rowPOTORDR2.Item("PO_DATE_SHIP_BY")
                            End If
                            rowPOTORDR2.Item("PO_DATE_SHIP_BY") = PO_DATE_SHIP_BY
                            rowPOTORDR2.Item("LAST_OPER_SHIP_BY") = ASCMAIN1.USER_ID
                            rowPOTORDR2.Item("LAST_DATE_SHIP_BY") = DATETIME_STAMP
                        End If

                        If Format(rowPOTORDR1.Item("PO_DATE_ETA"), "yyyyMMdd") <> Format(PO_DATE_ETA, "yyyyMMdd") Then
                            If rowPOTORDR2.Item("PO_ORIG_DATE_ETA") & "" = "" Then
                                rowPOTORDR2.Item("PO_ORIG_DATE_ETA") = rowPOTORDR2.Item("PO_DATE_ETA")
                            End If
                            rowPOTORDR2.Item("PO_DATE_ETA") = PO_DATE_ETA
                        End If
                    Next


                    ' how to handle multiple lines, same color, where price is different

                    For Each rowpocolor As DataRow In dst.Tables("pocolor").Select(sqlV)

                        Dim POColorKey As Int64 = Val(rowpocolor.Item("POColorKey") & "")
                        Dim STYLE_CODEs As List(Of String) = Get_Style_Codes(STYLE_CODE, POColorKey)

                        Dim rowPOTORDR2 As DataRow
                        Dim rowadded As Boolean = (rowpocolor.Item("ADD") & "" = "A")
                        If rowadded Then
                            ' change the qty open to 0
                            ' wjz blocking out this sectionn - not sure what to do - leaving the new exception in there
                            'Dim POKey As Int32 = Val(rowpocolor.Item("POKey") & "")
                            'Dim rowLast() As DataRow = dst.Tables("pocolor").Select("VAN_REF = '" & VAN_REF_last_AI & "' and POKey = " & CStr(POKey))

                            'If rowLast.Length = 1 Then
                            '    Dim COLOR_CODE_last As String = rowLast(0).Item("ColorCode")
                            '    If COLOR_CODE_last.StartsWith("#") Then COLOR_CODE_last = trim(COLOR_CODE_last.Substring(1))
                            '    rowPOTORDR2 = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, COLORs(COLOR_CODE_last)})
                            '    With rowPOTORDR2
                            '        .Item("PO_QTY_CXL") = .Item("PO_QTY_OPN")
                            '        .Item("PO_QTY_OPN") = 0
                            '        .Item("PO_STATUS") = "C"
                            '    End With
                            'End If
                            Throw New Exception("Test: Row Added Delete the Row or Change Open Qty to 0")
                        Else

                            Dim COLOR_CODE As String = rowpocolor.Item("ColorCode")
                            If COLOR_CODE.StartsWith("#") Then COLOR_CODE = Trim(COLOR_CODE.Substring(1))

                            For Each STYLE_CODE_to_check As String In STYLE_CODEs
                                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_to_check)
                                Dim SUB_UNIT_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "")

                                Dim StyleColor As String = STYLE_CODE_to_check & ":" & COLOR_CODE

                                Dim PO_QTY_ORD As Int64 = Val(rowpocolor.Item("OrderQty") & "") * 12 / SUB_UNIT_PACK_QTY
                                Dim OrderUnit As String = rowpocolor.Item("OrderUnit") & ""

                                ' the following block, and the code above which handles ME specaially for StyleBySize, added by WJZ 10/5/21 to process ME21255 PO 146810
                                If StyleBySize Then
                                    Dim rowposizedtl() As DataRow = dst.Tables("posizedtl").Select($"Size = '{Split(StyleColor, ":")(0)}' and ColorCode = '{Split(StyleColor, ":")(1)}'")
                                    If rowposizedtl.Length = 1 Then
                                        PO_QTY_ORD = Val(rowposizedtl(0).Item("Qty") & "")
                                    End If
                                End If

                                If StyleColors.ContainsKey(StyleColor) Then
                                    Dim first_time As Boolean = False
                                    If StyleColors(StyleColor) < 0 Then
                                        StyleColors(StyleColor) = -1 * StyleColors(StyleColor)
                                        first_time = True
                                    End If
                                    rowPOTORDR2 = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, StyleColors(StyleColor)})
                                    With rowPOTORDR2
                                        If first_time Then
                                            .Item("PO_QTY_ORD") = 0
                                            .Item("PO_QTY_OPN") = 0
                                        End If

                                        .Item("PO_QTY_ORD") += PO_QTY_ORD
                                        .Item("PO_QTY_OPN") += PO_QTY_ORD
                                    End With
                                Else
                                    rowPOTORDR2 = Update_Record_POTORDR2_Add(rowICTSTYL1, OrderUnit, STYLE_CODE_to_check, COLOR_CODE, PO_QTY_ORD)
                                End If

                                Update_Record_POTORDR2(rowPOTORDR2, PO_DATE_SHIP_BY, PO_DATE_ETA, _
                                   PO_COST_VCOST_DZ, PO_COST_COMM_PCT_ADD, PO_COST_OTHER_DZ, PO_COST_COMM_PCT, PO_COST_MATLS_DZ, _
                                   SUB_UNIT_PACK_QTY) ' always do these updates
                            Next
                        End If
                    Next

                    If TAC.POCMAIN1.Check_Changed_Fields(Me, rowPOTORDR1) Then
                        ' WHAT ABOUT HDR REV CTR - SEE ABOVE
                        'Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                        'rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV + 1
                    End If

                    commitMsg = " - PO " & PO_ORDER_NO & " has been Updated"


                    Dim StyleColorsNotTouched As New List(Of String)
                    For Each StyleColor As String In StyleColors.Keys
                        If StyleColors(StyleColor) < 0 Then
                            StyleColorsNotTouched.Add(StyleColor)
                        End If
                    Next
                    If StyleColorsNotTouched.Count > 0 Then
                        MsgBox("Please Note:" _
                               & vbCrLf & " - the following Style Colors were not edited during this Process" _
                               & vbCrLf & " - Manual Deletion or Editing may be required" _
                               & vbCrLf & Join(StyleColorsNotTouched.ToArray, ","), MsgBoxStyle.OkOnly, "Alert")
                    End If

                End If
            End If
        End If

        BeginTrans()
        '  Stop


        Update_Record_TDA("POTORDRA")

        ' PROBABLY SHOULD JUST email_to_FollowBy() UNLESS DELETING

        If optImport.Value = "R" Then

            ASCMAIN1.sql = "UPDATE POTORDRA SET STATUS = 'X', SUPERCEDED_BY = :PARM1 WHERE PONO = :PARM2 AND STATUS = 'A'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {VAN_REF, PONO})

            ' email_to_FollowBy()

        ElseIf optImport.Value = "A" Then

            ' email_to_FollowBy()

        ElseIf optImport.Value = "I" Then


            If rowpohdr.Item("POStatus") & "" = "C" Or rowpohdr.Item("SendRemarks") & "" = "Cancel PO" Then

                If PO_ORDER_NO <> "" Then
                    TAC.POCMAIN1.Dependent_Updates(-1, PO_ORDER_NO, True)
                    ASCMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, "", Now, ASCMAIN1.USER_ID, "ATPOC", "AT PO Cancelled by AT", VAN_REF)

                    ' Write_Audit_Trail UPDATES ORACLE, WHEREAS WriteAuditTrail UPDATES DST
                    'For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                    '    Write_Audit_Trail(rowPOTORDR2, "E")
                    'Next
                    WriteAuditTrail("POTORDR2")
                End If

            Else

                TAC.POCMAIN1.Dependent_Updates(-1, PO_ORDER_NO)
                If EntryMode = "E" Then
                    ASCMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, "", Now, ASCMAIN1.USER_ID, "ATPOM", "AT PO Revised by AT", VAN_REF)
                Else
                    ASCMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, "", Now, ASCMAIN1.USER_ID, "ATPOE", "AT PO Created by AT", VAN_REF)
                End If

                Update_Record_TDA("POTORDR1")
                Update_Record_TDA("POTORDR2")
                Update_Record_TDA("POTORDXR")

                TAC.POCMAIN1.Dependent_Updates(1, PO_ORDER_NO)

                ' email_to_FollowBy()

            End If

        End If

        CommitTrans("Update Complete" & commitMsg)
    End Sub

    Sub Automatically_Approve()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Automatically Approving")
        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

        dst.Tables("POTORDRA").Rows.Clear()

        Dim SQLW0 As String = "STATUS <> 'X'" ' Superceded
        For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select(SQLW0)
            VAN_REF = rowPOTORDRX.Item("VAN_REF")
            ASCMAIN1.Progress("-", VAN_REF)
            rowPOTORDRA = Fill_Record("POTORDRA", VAN_REF, False, False)

            If rowPOTORDRA.Item("INIT_OPER") & "" = "" Then
                rowPOTORDRA.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowPOTORDRA.Item("INIT_DATE") = DATETIME_STAMP
            End If

            rowPOTORDRA.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTORDRA.Item("LAST_DATE") = DATETIME_STAMP

            Dim sqlw As String = $"SUPERCEDED_BY = '{VAN_REF}' and (STATUS = 'X' or STATUS = 'A' or STATUS = 'I')"

            ' Dim sqlV As String = "VAN_REF = '" & VAN_REF & "'"

            For Each row As DataRow In dst.Tables("POTORDRX").Select(sqlw)
                Dim STATUS = row.Item("STATUS")
                If STATUS = "X" AndAlso row.Item("INIT_OPER") & "" = "" Then
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("INIT_DATE") = DATETIME_STAMP
                End If
                row.Item("LAST_OPER") = ASCMAIN1.USER_ID
                row.Item("LAST_DATE") = DATETIME_STAMP
                Dim row2 As DataRow = dst.Tables("POTORDRA").NewRow
                For i As Integer = 0 To dst.Tables("POTORDRA").Columns.Count - 1
                    Dim dc As DataColumn = dst.Tables("POTORDRA").Columns(i)
                    Dim dcname As String = dc.ColumnName
                    row2.Item(dcname) = row.Item(dcname)
                Next
                dst.Tables("POTORDRA").Rows.Add(row2)
                row2.AcceptChanges()
                row2.SetModified()
            Next

            rowPOTORDRA.Item("STATUS") = "A"

            rowPOTORDRA.Item("APPROVED_DATE") = DATETIME_STAMP
            rowPOTORDRA.Item("APPROVED_BY") = ASCMAIN1.USER_ID
            rowPOTORDRA.Item("APPROVED_MESSAGE") = "Auto-Approved"
        Next

        BeginTrans()
        Update_Record_TDA("POTORDRA")
        CommitTrans()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Update_Record_POTORDR2_Add(rowICTSTYL1 As DataRow, OrderUnit As String, _
                                        STYLE_CODE As String, COLOR_CODE As String, PO_QTY_ORD As Int64) As DataRow

        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").NewRow
        With rowPOTORDR2
            .Item("PO_ORDER_NO") = PO_ORDER_NO
            PO_ORDER_LNO += 1
            StyleColors.Add(STYLE_CODE & ":" & COLOR_CODE, PO_ORDER_LNO)
            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
            .Item("STYLE_CODE") = STYLE_CODE
            .Item("COLOR_CODE") = COLOR_CODE
            .Item("PO_QTY_ORD") = PO_QTY_ORD
            .Item("PO_QTY_OPN") = PO_QTY_ORD

            .Item("PO_STATUS") = "O"

            If OrderUnit = "DOZ" Then
                .Item("PO_QTY_UOM") = 12
            Else
                Throw New Exception("INVALID VALUE FOR OrderUnit")
            End If

            '.Item("STYLE_NOTES") = "?"
            '.Item("PO_COST_QUOTA") = "?"
            '.Item("DFQUOTA") = "?"

            .Item("CARTON_PACK_QTY") = rowICTSTYL1.Item("CARTON_PACK_QTY")
            .Item("PO_CONF_NO") = VAN_REF
            .Item("PO_CONF_DATE") = DATETIME_STAMP.Date
            .Item("INNER_PACK_QTY") = rowICTSTYL1.Item("INNER_PACK_QTY")
            '.Item("PO_LINE_NOTE_INT") = "?"
        End With

        dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)

        Return rowPOTORDR2

    End Function

    Sub Update_Record_POTORDR2(rowPOTORDR2 As DataRow, _
                               PO_DATE_SHIP_BY As Date, _
                               PO_DATE_ETA As Date, _
                               PO_COST_VCOST_DZ As Decimal, _
                               PO_COST_COMM_PCT_ADD As Decimal, _
                               PO_COST_OTHER_DZ As Decimal, _
                               PO_COST_COMM_PCT As Decimal, _
                               PO_COST_MATLS_DZ As Decimal, _
                               SUB_UNIT_PACK_QTY As Integer)
        With rowPOTORDR2
            .Item("PO_COST_VCOST_DZ") = PO_COST_VCOST_DZ '  * SUB_UNIT_PACK_QTY
            .Item("PO_COST_OTHER") = PO_COST_OTHER_DZ
            .Item("PO_COST_COMM") = PO_COST_COMM_PCT

            Dim PO_COST_QUOTA_UN As Decimal = 0 ' UNTIL SOMEONE TELLS ME OTHERWISE
            Dim PO_COST_SUBTOTAL As Decimal = PO_COST_VCOST_DZ + PO_COST_MATLS_DZ + PO_COST_OTHER_DZ
            Dim PO_COST_COMM_UN As Decimal = System.Math.Round(PO_COST_SUBTOTAL * PO_COST_COMM_PCT / 100, 6)
            Dim PO_COST As Decimal = PO_COST_SUBTOTAL + PO_COST_COMM_UN + PO_COST_QUOTA_UN
            .Item("PO_COST") = PO_COST / 12 * SUB_UNIT_PACK_QTY
            .Item("PO_COST_VCOST") = PO_COST_VCOST_DZ / 12 * SUB_UNIT_PACK_QTY
            .Item("PO_COST_BUFFER") = PO_COST_COMM_PCT_ADD
            .Item("SUB_UNIT_PACK_QTY") = SUB_UNIT_PACK_QTY

            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP

            .Item("PO_DATE_SHIP_BY") = PO_DATE_SHIP_BY
            .Item("PO_DATE_ETA") = PO_DATE_ETA
            If .Item("PO_ORIG_DATE_SHIP_BY") & "" = "" Then
                .Item("PO_ORIG_DATE_SHIP_BY") = PO_DATE_SHIP_BY
                .Item("PO_ORIG_DATE_ETA") = PO_DATE_ETA
            End If


            If rowPOTORDR2.RowState = DataRowState.Modified AndAlso .Item("PO_DATE_SHIP_BY", DataRowVersion.Original) & "" <> "" AndAlso Format(.Item("PO_DATE_SHIP_BY", DataRowVersion.Original), "yyyyMMdd") <> Format(PO_DATE_SHIP_BY, "yyyyMMdd") Then
                .Item("LAST_OPER_SHIP_BY") = ASCMAIN1.USER_ID
                .Item("LAST_DATE_SHIP_BY") = Now + ASCMAIN1.NowTSD
            End If




            'SHIP_COST_CHANGE_USER VARCHAR2(20)
            'SHIP_COST_CHANGE_DATE          DATE

        End With
    End Sub
    Function ValidatePO() As String
        Dim EMsg As String = ""
        Dim sqlV As String = "VAN_REF = '" & VAN_REF & "'"

        If PO_ORDER_NO <> "" Then
            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
            If rowPOTORDR1 IsNot Nothing AndAlso rowPOTORDR1.Item("PO_STATUS") & "" <> "O" Then
                '  Throw New Exception("Cannot change a PO that is not Open")
                EMsg &= vbCr & "Cannot change a PO that is not Open: " & PO_ORDER_NO
            End If

            ' CHANGED TO GROUP BY STYLE/COLOR ON 09/30/21 WJZ SO THAT WE CAN IMPORT ME POS
            ASCMAIN1.sql = "Select * from (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, COUNT (*) LINES, SUM (PO_QTY_ORD) ORD, SUM (PO_QTY_SHP) SHP, SUM (PO_QTY_OPN) OPN" & vbCrLf _
                & " from POTORDR2 WHERE PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ") where LINES <> 1 OR NVL(ORD,0) <> NVL(OPN,0) OR NVL(SHP,0) <> 0"
            Dim TBL As DataTable = ASCDATA1.GetDataTable
            If TBL.Rows.Count <> 0 Then
                EMsg &= vbCr & "Cannot change a PO that has been shipped or split: " & PO_ORDER_NO & ", Color: " & TBL.Rows(0).Item("COLOR_CODE")
            End If

        End If
         

        Dim rowpohdr As DataRow = dst.Tables("pohdr").Select(sqlV)(0) '.Rows(0)
        Dim STYLE_CODE As String = rowpohdr.Item("StyleNo") & ""
        STYLE_CODE = STYLE_CODE.Replace("-", "")
        Dim STYLE_CODEs As List(Of String) = Get_Style_Codes(STYLE_CODE, -1)

        Dim SUB_UNIT_PACK_QTY As Decimal = 0
        For Each STYLE_CODE_to_check As String In STYLE_CODEs
            Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE_to_check)

            If rowICTSTYL1 Is Nothing Then
                EMsg &= vbCr & "Invalid Style Code: " & STYLE_CODE_to_check
            Else
                If SUB_UNIT_PACK_QTY = 0 Then
                    SUB_UNIT_PACK_QTY = Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "")
                Else
                    If Val(rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & "") <> SUB_UNIT_PACK_QTY Then
                        EMsg &= vbCr & "Different Sub-Unit Pack Qtys (" & rowICTSTYL1.Item("SUB_UNIT_PACK_QTY") & " vs " & CStr(SUB_UNIT_PACK_QTY) & ") in Styles: " & Join(STYLE_CODEs.ToArray, ".")
                    End If
                End If
            End If
        Next
        

        Dim FACTORY_CODE As String = rowpohdr.Item("Factory") & ""
        Dim rowICTFACT1 As DataRow = LookUp("ICTFACT1", FACTORY_CODE)
        If rowICTFACT1 Is Nothing Then
            EMsg &= vbCr & "Invalid Factory Code: " & FACTORY_CODE
        End If

        If rowpohdr.Item("OrderDate") & "" = "" Then
            EMsg &= vbCr & "Missing Order Date"
        End If

        If rowpohdr.Item("VandaleShipDate") & "" = "" Then
            EMsg &= vbCr & "Missing Vandate Ship Date"
        End If


        Dim AT_COLORS As New Dictionary(Of String, Int32)
        Dim SCs As New List(Of String)

        Dim COST As Decimal = -1

        For Each rowpocolor As DataRow In dst.Tables("pocolor").Select(sqlV)

            Dim OrderQty As Int32 = Val(rowpocolor.Item("OrderQty") & "")
            Dim HangerUnit As Int32 = Val(rowpocolor.Item("HangerUnit") & "")
            Dim QtyDoz As Int32 = Val(rowpocolor.Item("QtyDoz") & "")
            Dim POColorKey As Int32 = Val(rowpocolor.Item("POColorKey") & "")

            If Val(rowpocolor.Item("Cost") & "") <> 0 And COST <> -2 Then
                If COST = -1 Then
                    COST = Val(rowpocolor.Item("Cost") & "")
                Else
                    If COST <> Val(rowpocolor.Item("Cost") & "") Then
                        EMsg &= vbCr & "Multiple Costs in Color grid: " & CStr(COST) & "," & rowpocolor.Item("Cost")
                        COST = -2
                    End If
                End If
            End If


            Dim OrderUnit As String = rowpocolor.Item("OrderUnit") & ""

            If OrderQty <> 0 Then
                Dim COLOR_CODE As String = rowpocolor.Item("ColorCode") & ""

                If Not AT_COLORS.ContainsKey(COLOR_CODE) Then
                    AT_COLORS.Add(COLOR_CODE, HangerUnit)
                End If

                If COLOR_CODE.StartsWith("#") Then COLOR_CODE = Trim(COLOR_CODE.Substring(1))
                Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                If rowICTCOLR1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Color Code: " & COLOR_CODE
                Else
                    If Not SCs.Contains(COLOR_CODE) Then
                        SCs.Add(COLOR_CODE)
                        For Each STYLE_CODE_to_check As String In STYLE_CODEs
                            Dim rowICTSTYC1 As DataRow = LookUp("ICTSTYC1", New String() {STYLE_CODE_to_check, COLOR_CODE})
                            If rowICTSTYC1 Is Nothing Then
                                EMsg &= vbCr & "Missing Style/Color Record for " & STYLE_CODE_to_check & "/" & COLOR_CODE
                            End If
                        Next
                    End If
                End If

                If OrderQty <> QtyDoz Or OrderUnit <> "DOZ" Then
                    EMsg &= vbCr & "Need to investigate the Qtys and UOMs on line " & CStr(POColorKey)
                End If

                If OrderUnit <> "DOZ" Then
                    EMsg &= vbCr & "Need to investigate the UOM (" & OrderUnit & ") on line " & CStr(POColorKey)
                End If
            End If
        Next

        If EMsg = "" Then
            For Each AT_COLOR As String In AT_COLORS.Keys
                Dim HangerUnit As Decimal = Val(AT_COLORS(AT_COLOR))
                Dim OrderQty As Decimal = Val(dst.Tables("pocolor").Compute("Sum(OrderQty)", sqlV & " and ColorCode = '" & AT_COLOR & "'") & "")
                If OrderQty <> 0 Then
                    If HangerUnit = 0 OrElse SUB_UNIT_PACK_QTY <> System.Math.Round(OrderQty * 12 / HangerUnit, 2) Then
                        EMsg &= vbCr & "Need to investigate the HangerUnit (" & CStr(HangerUnit) & ") for Color " & AT_COLOR
                    End If
                End If
            Next
        End If

        ' COST MUST MAKE SENSE


        If rowpohdr.Item("POStatus") & "" = "C" Or rowpohdr.Item("SendRemarks") & "" = "Cancel PO" Then
            EMsg = "" ' no point in bothering if they want to cancel
        End If

        Return EMsg

    End Function
    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Generate_Report("PORWREC2")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PO_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("VAN_REF").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function


#End Region


#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTORDRX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry")
        Load_Popup_Menu(grdPOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "PO Inquiry")
        'Load_Popup_Menu(grdICTSTYLF, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdpohdr, "B", "Style Master File")
        Load_Popup_Menu(grdImages, "B", "Show Image")
        Load_Popup_Menu(grdPDFs, "B", "Show PDF")
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
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdPOTORDRX"
                'tlb_btn = DirectCast(tlb_pop.Tools("Select All VENDOR"), UltraWinToolbars.ButtonTool)
                'If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                '    tlb_btn.SharedProps.Visible = False
                'Else
                '    tlb_btn.SharedProps.Visible = True
                '    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                '    tlb_btn.Tag = VEND_CODE
                '    tlb_btn.SharedProps.Caption = "Select All " & VEND_CODE
                'End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdPOTORDR3"
                '    tlb_sbt = DirectCast(tlb.Tools("Show Cartons"), UltraWinToolbars.StateButtonTool)
                '    e.Tool.SharedProps.Visible = tlb_sbt.Checked

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All"
                For Each rowPOTORDRX As DataRow In dst.Tables("POTORDRX").Select("")
                    rowPOTORDRX.Item("SEL") = "1"
                Next
                'For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                '    If grow.IsDataRow Then
                '        grow.Cells("SEL").Value = "1"
                '        grow.Update()
                '    End If
                'Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text & ""
                If PO_ORDER_NO = "" Then
                    MsgBox("No PO associated with this transmission", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else
                    Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")
                End If

            Case "PO Entry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                If PO_ORDER_NO = "" Then
                    MsgBox("No PO associated with this transmission", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else
                    Context_Launch("Edit", PO_ORDER_NO, e.Tool.Key, "POFORDR1", "F", "POE")
                End If


            Case "Style Master File"

                Dim STYLE_CODE As String = grd.ActiveRow.Cells("StyleNo").Text
                STYLE_CODE = STYLE_CODE.Replace("-", "")
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Dim keys As New Dictionary(Of String, Object)
                    keys.Add("STYLE_CODE", STYLE_CODE)
                    Context_Launch("View", Keys, e.Tool.Key, "ICTSTYL1")
                Else
                    MsgBox("Style does not exist", MsgBoxStyle.OkOnly, "Cannot View Style")
                End If

            Case "Show Image"

                If grd.ActiveRow.IsDataRow Then
                    Dim POKEY As String = rowPOTORDRA.Item("POKEY")
                    Dim FILENAME As String = grd.ActiveRow.Cells("FILENAME").Text
                    Dim file As String = ASCMAIN1.Folders("Temp") & "\" & FILENAME
                    Try
                        If My.Computer.FileSystem.FileExists(file) Then
                            My.Computer.FileSystem.DeleteFile(file)
                        End If
                        My.Computer.FileSystem.CopyFile(PO_PARM_PO_IMG_DIR & "\" & FILENAME, ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                        'Show_Document(PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME)
                        Show_Document(ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Show Image")
                    End Try

                End If

            Case "Show PDF"
                If grd.ActiveRow.IsDataRow Then
                    Dim POKEY As String = rowPOTORDRA.Item("POKEY")
                    Dim FILENAME As String = grd.ActiveRow.Cells("FILENAME").Text
                    Dim file As String = ASCMAIN1.Folders("Temp") & "\" & FILENAME
                    Try
                        If My.Computer.FileSystem.FileExists(file) Then
                            My.Computer.FileSystem.DeleteFile(file)
                        End If
                        My.Computer.FileSystem.CopyFile(PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME, ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                        'Show_Document(PO_PARM_PO_IMG_DIR & "\" & POKEY & "\" & FILENAME)
                        Show_Document(ASCMAIN1.Folders("Temp") & "\" & FILENAME)
                    Catch ex As Exception
                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Show PDF")
                    End Try

                End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub

#End Region

    Private Sub grdPOTORDRX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDRX.AfterRowActivate

    End Sub

    Private Sub grdPOTORDRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDRX.DoubleClickRow
        If ScreenMode Then Exit Sub
        If ScreenMode Or (EntryMode <> "") Then ' Or (optDisplay.Value <> "W" And optDisplay.Value <> "A") Then
            Exit Sub
        End If

        If grdPOTORDRX.ActiveRow IsNot Nothing Then
            Absx1.txtFor("VAN_REF").Text = grdPOTORDRX.ActiveRow.Cells("VAN_REF").Text
            If optDisplay.Value = "*" Then
                Click_Command("View")
            Else
                Click_Command("Select")
            End If

        End If
    End Sub

    Sub Load_POTORDRX(Optional tbl As DataTable = Nothing)
        Dim STATUS As String = optDisplay.Value
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")
        If tbl Is Nothing Then
            Fill_Records("POTORDRX", STATUS)
        Else
            dst.Tables("POTORDRX").Rows.Clear()
            For Each row As DataRow In tbl.Rows
                dst.Tables("POTORDRX").Rows.Add(row.ItemArray)
            Next
        End If

        Set_Tentative_Import_Values()

        grdPOTORDRX.Text = "POs Transmitted by AT - " & optDisplay.Text _
            & IIf(ScreenMode Or (EntryMode <> "") Or (optDisplay.Value <> "W"), "", " - Double-Click Row to Select")
        Sort_grdColumns(grdPOTORDRX, "PONO," & "VAN_REF".ToLower)

        '  grdPOTORDRX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdPOTORDRX.DisplayLayout.Bands(0).ColumnFilters("STATUS").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.DoesNotStartWith, "Superceded")


        UltraExplorerBar1.Groups("Display Options").Visible = True


        If tbl Is Nothing AndAlso optDisplay.Value = "W" Then
            Automatically_Approve()
            optDisplay.Value = "A"
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Set_Tentative_Import_Values(Optional tbl As DataTable = Nothing)
        If tbl Is Nothing Then
            tbl = dst.Tables("POTORDRX")
        End If
        For Each row As DataRow In tbl.Select("", "VAN_REF DESC")
            Dim VAN_REF As String = row.Item("VAN_REF")
            Dim STATUS As String = row.Item("STATUS")
            If STATUS = "W" Then
                Dim PONO As String = row.Item("PONO")

                'ASCMAIN1.sql = "Select * from POTORDR1 where PO_REFERENCE = :PARM1"
                'Dim rowPOTORDR1 = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {PONO})
                'If rowPOTORDR1 IsNot Nothing Then
                '    Dim PO_ORDER_NO As String = rowPOTORDR1.Item("PO_ORDER_NO")
                '    row.Item("PO_ORDER_NO") = PO_ORDER_NO
                'End If

                ASCMAIN1.sql = "Select * from POTORDRA where PONO = :PARM1 and STATUS = 'I'"
                Dim rowPOTORDRA = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New String() {PONO})
                If rowPOTORDRA IsNot Nothing Then
                    Dim PO_ORDER_NO As String = rowPOTORDRA.Item("PO_ORDER_NO")
                    row.Item("PO_ORDER_NO") = PO_ORDER_NO
                End If


                Dim sqlw As String = "PONO = '" & PONO & "' and VAN_REF <> '" & VAN_REF & "' and STATUS = 'W'"
                For Each row2 As DataRow In tbl.Select(sqlw)
                    row2.Item("STATUS") = "X"
                    row2.Item("SUPERCEDED_BY") = VAN_REF
                Next
            End If
        Next
    End Sub

    Private Sub grdPOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRX.InitializeRow
        If e.Row.Cells("STATUS").Value & "" = "W" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Blue
        ElseIf e.Row.Cells("STATUS").Value & "" = "R" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("STATUS").Value & "" = "I" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Purple
        ElseIf e.Row.Cells("STATUS").Value & "" = "A" Then
            e.Row.CellAppearance.ForeColor = Drawing.Color.Green
        Else
            e.Row.CellAppearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub optImport_ValueChanged(sender As Object, e As EventArgs) Handles optImport.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Toggle_Message_Block()
    End Sub

    Sub Toggle_Message_Block()
        lblEMAIL_TO.Visible = (optImport.Value = "R")
        txtEMAIL_TO.Visible = (optImport.Value = "R")
        lblEMAIL_CC.Visible = (optImport.Value = "R")
        txtEMAIL_CC.Visible = (optImport.Value = "R")
        chkAutoSendEmail.Visible = (optImport.Value = "R")
        lblREASON_DESC.Visible = (optImport.Value = "R" Or optImport.Value = "D")
        txtREASON_DESC.Visible = (optImport.Value = "R" Or optImport.Value = "D")
        lblAPPROVED_MESSAGE.Visible = (optImport.Value = "A")
        txtAPPROVED_MESSAGE.Visible = (optImport.Value = "A")
    End Sub

    Private Sub optDisplay_ValueChanged(sender As Object, e As EventArgs) Handles optDisplay.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_Display()
        Load_POTORDRX()
    End Sub

    Sub Setup_Display()
        With grdPOTORDRX.DisplayLayout.Bands(0)
            .Columns("DATE_IMPORTED").Hidden = Not (optDisplay.Value = "*")
            .Columns("INIT_DATE").Hidden = Not (optDisplay.Value = "*")
            .Columns("INIT_OPER").Hidden = Not (optDisplay.Value = "*")
            .Columns("LAST_DATE").Hidden = Not (optDisplay.Value = "*")
            .Columns("LAST_OPER").Hidden = Not (optDisplay.Value = "*")
            .Columns("REASON_DESC").Hidden = Not (optDisplay.Value = "*")
            .Columns("EMAIL_TO").Hidden = Not (optDisplay.Value = "*")

            .Columns("APPROVED_DATE").Hidden = (optDisplay.Value = "W")
            .Columns("APPROVED_BY").Hidden = (optDisplay.Value = "W")
            .Columns("APPROVED_MESSAGE").Hidden = (optDisplay.Value = "W")
        End With
    End Sub

    Private Sub grdpohdr_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdpohdr.InitializeRow
        Color_Version(e, grdpohdr)
    End Sub

    Private Sub grdpocolor_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdpocolor.InitializeRow
        Color_Version(e, grdpocolor)
    End Sub

    Private Sub grdpofactory_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdpofactory.InitializeRow
        Color_Version(e, grdpofactory)
    End Sub

    Private Sub grdposize_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdposize.InitializeRow
        Color_Version(e, grdposize)
    End Sub
    Private Sub grdposizedtl_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdposizedtl.InitializeRow
        Color_Version(e, grdposizedtl)
    End Sub
    Private Sub grdpotrim_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdpotrim.InitializeRow
        Color_Version(e, grdpotrim)
    End Sub

    Private Sub grdpofabric_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdpofabric.InitializeRow
        Color_Version(e, grdpofabric)
    End Sub

    Sub Color_Version(e As UltraWinGrid.InitializeRowEventArgs, grd As UltraWinGrid.UltraGrid)
        If e.Row.Cells("VAN_REF").Value & "" = VAN_REF Then
            e.Row.Cells("XMIT_VERSION").Value = "Curr"
            e.Row.Cells("XMIT_VERSION").Appearance.ForeColor = Drawing.Color.Blue
        ElseIf e.Row.Cells("VAN_REF").Value & "" = VAN_REF_previous Then
            e.Row.Cells("XMIT_VERSION").Value = "Prev"
            e.Row.Cells("XMIT_VERSION").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("VAN_REF").Value & "" = VAN_REF_last_AI Then
            e.Row.Cells("XMIT_VERSION").Value = "Last"
            e.Row.Cells("XMIT_VERSION").Appearance.ForeColor = Drawing.Color.Purple
        ElseIf e.Row.Cells("VAN_REF").Value & "" = "0000000000" Then
            e.Row.Cells("XMIT_VERSION").Value = "Master"
            e.Row.Cells("XMIT_VERSION").Appearance.ForeColor = Drawing.Color.Green
        End If

        If grd.Name = "grdPOTORDRB" Then Exit Sub


        Dim difference_exists As Boolean = False
        Dim TABLE_NAME As String = Mid(grd.Name, 4)

        If TABLE_NAME = "Style" Or TABLE_NAME = "StyleColor" Or TABLE_NAME = "StyleSize" Then
        Else
            If e.Row.Cells("XMIT_VERSION").Value = "Curr" Then
                If TABLE_NAME = "pofactory" Then
                    e.Row.RowSpacingBefore = 2
                End If
            End If
        End If

        If TABLE_NAME = "Style" Or TABLE_NAME = "StyleColor" Or TABLE_NAME = "StyleSize" Then
        Else
            Dim key As Int64 = Val(e.Row.Cells(grdKeys(Mid(grd.Name, 4))).Value & "")
            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                If Not gcol.Hidden Then
                    Dim col As String = gcol.Key
                    Dim rowDifference = dst.Tables("differences").Rows.Find(New Object() {TABLE_NAME, key, col})
                    If rowDifference IsNot Nothing Then
                        e.Row.Cells(col).Appearance.BackColor = Drawing.Color.Yellow
                        difference_exists = True
                    End If
                End If
            Next
        End If

        If e.Row.Cells("VAN_REF").Value & "" = VAN_REF Then ' annotate the Action column for Curr records only
            If VAN_REF_previous = "" Then
                e.Row.Cells("XMIT_ACTION").Value = "Add"
                e.Row.Cells("XMIT_ACTION").Appearance.ForeColor = Drawing.Color.Green
            Else
                If e.Row.Cells("ADD").Value & "" = "A" Then ' Curr Record was added for comparison purposes only
                    e.Row.Cells("XMIT_ACTION").Value = "Delete"
                    e.Row.Cells("XMIT_ACTION").Appearance.ForeColor = Drawing.Color.Red
                Else
                    If difference_exists Then
                        e.Row.Cells("XMIT_ACTION").Value = "Change"
                        e.Row.Cells("XMIT_ACTION").Appearance.ForeColor = Drawing.Color.Blue
                    Else
                        e.Row.Cells("XMIT_ACTION").Value = "-"
                        e.Row.Cells("XMIT_ACTION").Appearance.ForeColor = Drawing.Color.Empty
                    End If
                End If
            End If
        End If


    End Sub

    Sub email_to_FollowBy()
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim EMAIL_TO As String = Absx1.txtFor("EMAIL_TO").Text.Trim

        '   EMAIL_TO = "wjz@absolution.com"

        Dim EMAIL_BODY As String = ""
        Dim REASON_DESC As String = Absx1.txtFor("REASON_DESC").Text

        If optImport.Value = "R" Then
            EMAIL_BODY = "Reason for Rejection: " & vbCrLf & REASON_DESC _
                                        & vbCrLf & vbCrLf _
                                        & vbCrLf & "PO Reference " & PONO _
                                        & vbCrLf & "Vandale Transmission Ctl " & VAN_REF

        ElseIf optImport.Value = "A" Then
            REASON_DESC = ""
            EMAIL_BODY = "PO has been Approved: " & vbCrLf & REASON_DESC _
                            & vbCrLf & vbCrLf _
                            & vbCrLf & "PO Reference " & PONO _
                            & vbCrLf & "Vandale Transmission Ctl " & VAN_REF

        ElseIf optImport.Value = "I" Then
            REASON_DESC = ""
            EMAIL_BODY = "PO has been Imported: " & vbCrLf & REASON_DESC _
                            & vbCrLf & vbCrLf _
                            & vbCrLf & "PO Reference " & PONO _
                            & vbCrLf & "Vandale Transmission Ctl " & VAN_REF
        End If


        Dim VEND_CODE As String = "AT"
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        ' Dim REPORT_NO As String = Print_POs(PO_ORDER_NO, True, PO_ORDER_NO)
        'ATTACHMENTs.Add(PO_ORDER_NO & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO & ".pdf")

        Dim SUBJECT As String = ""

        Dim PFX As String = ""
        If optImport.Value = "R" Then
            PFX = "Rejected "
        ElseIf optImport.Value = "A" Then
            PFX = "Approved "
        ElseIf optImport.Value = "I" Then
            PFX = "Imported "
        End If

        If ASCMAIN1.DBS_COMPANY = "TST" Then
            PFX &= " - TEST - "
        End If

        SUBJECT = PFX & "PO " & PONO

        Dim SEND_CC_to_USER_ID As Boolean = True

        Dim FollowBy As String = rowpohdr.Item("FollowBy") & ""
        Dim FollowByEmail As String = rowpohdr.Item("FollowByEmail") & ""

        Dim EMAIL_NAME As String = FollowBy
        If EMAIL_TO <> FollowByEmail Then
            EMAIL_NAME = "PO Processor"
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(EMAIL_TO, EMAIL_NAME)
        If txtEMAIL_CC.Text <> "" Then
            Dim CCs() As String = txtEMAIL_CC.Text.Split(";")
            For Each CC As String In CCs
                EMAIL_ADDRESSs.Add(CC, "")
            Next
        End If

        Dim SEND_NO As String = ""

        Try
            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                SUBJECT, "ATPOR", chkAutoSendEmail.Checked, SEND_CC_to_USER_ID, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier", EMAIL_BODY)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "error sending email")
        End Try

        If SEND_NO <> "" Then
            rowPOTORDRA.Item("SEND_NO") = SEND_NO
            ASCMAIN1.Record_Event("POTORDRA", VAN_REF, "", Now, ASCMAIN1.USER_ID, "ATPO" & optImport.Value, "AT PO " & PFX, SEND_NO)
        Else
            MsgBox("Please email screenshot to wjz@absolution.com", MsgBoxStyle.OkOnly, "email was not sent")
            ' Throw New Exception("Please email screenshot to wjz@absolution.com")
            '   Stop ' NOW WHAT, BATMAN?
        End If
    End Sub

    Function Calculate_ETD_to_ETA(PORT_CODE_ORIG As String, WHSE_CODE As String, PO_SHIP_VIA As String) As Integer
        'this routine was ripped from POFORDR1 - probably refactor into a TAC class

        Dim ETD_to_ETA As Integer = Val(ROWs("POTPARM1").Item("PO_PARM_DEF_DAYS_ETD_TO_ETA") & "")
        If PORT_CODE_ORIG = "" Or WHSE_CODE = "" Then
        Else
            Dim rowICTPORT2 As DataRow = LookUp("ICTPORT2", New String() {PORT_CODE_ORIG, WHSE_CODE})
            If rowICTPORT2 IsNot Nothing Then
                ETD_to_ETA = Val(rowICTPORT2.Item("ETD_TO_ETA") & "")
            End If
        End If

        If PO_SHIP_VIA <> "" Then
            Dim row As DataRow = LookUp("POTSVIA1", PO_SHIP_VIA)
            If row IsNot Nothing Then
                If row.Item("PO_SHIP_VIA_ETD_TO_ETA") & "" <> "" Then
                    ETD_to_ETA = Val(row.Item("PO_SHIP_VIA_ETD_TO_ETA") & "")
                End If
            End If
        End If

        Return ETD_to_ETA
    End Function

    Private Sub chkShowChangesOnly_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowChangesOnly.CheckedChanged
        Toggle_grd_Columns()
        If chkShowChangesOnly.Checked Then
            tabPO.SelectedTab = tabPO.Tabs("AT Data")
        End If
    End Sub

    Sub Toggle_grd_Columns()

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdpohdr, grdpocolor, grdpofactory, grdposize, grdposizedtl, grdpotrim, grdpofabric}
            grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, UltraWinGrid.AutoResizeColumnWidthOptions.All)
            Dim tbl As DataTable = DirectCast(grd.DataSource, DataTable)
            Dim cols As List(Of String) = gcols(grd.Name)

            For Each dcol As DataColumn In tbl.Columns
                Dim col As String = dcol.ColumnName
                If Not cols.Contains(col) Or gcols_excluded(grd.Name).Contains(col) Then
                    grd.DisplayLayout.Bands(0).Columns(col).Hidden = True
                Else
                    If Not chkShowChangesOnly.Checked Or gcols_different(grd.Name).Contains(col) Then
                        grd.DisplayLayout.Bands(0).Columns(col).Hidden = False
                    Else
                        grd.DisplayLayout.Bands(0).Columns(col).Hidden = True
                    End If
                End If
            Next

            Sort_grdColumns(grd, grdKeys(grd.Name.Substring(3)) & ",VAN_REF".ToLower)
        Next
    End Sub

    Sub Get_List_of_Columns_with_Differences()

        dst.Tables("differences").Rows.Clear()
        gcols_different.Clear()

        For Each TABLE_NAME As String In grdKeys.Keys

            Dim cols As New List(Of String)
            Dim keys As New List(Of Int64)

            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")

                Dim key As Int64 = Val(row.Item(grdKeys(TABLE_NAME)) & "")

                If Not keys.Contains(key) Then

                    keys.Add(key)

                    Dim rowCurr() As DataRow = dst.Tables(TABLE_NAME).Select("VAN_REF = '" & VAN_REF & "' and " & grdKeys(TABLE_NAME) & " = " & CStr(key))
                    Dim rowPrev() As DataRow = dst.Tables(TABLE_NAME).Select("VAN_REF = '" & VAN_REF_previous & "' and " & grdKeys(TABLE_NAME) & " = " & CStr(key))
                    Dim rowLast() As DataRow = dst.Tables(TABLE_NAME).Select("VAN_REF = '" & VAN_REF_last_AI & "' and " & grdKeys(TABLE_NAME) & " = " & CStr(key))

                    Dim DIFFERENCE As String = ""
                    Dim COLUMN_NAME = "*"

                    If rowCurr.Length = 0 Then
                        DIFFERENCE = "Deleted"
                        dst.Tables("differences").Rows.Add(New Object() {TABLE_NAME, key, COLUMN_NAME, DIFFERENCE})

                    ElseIf rowPrev.Length = 0 Or (VAN_REF_last_AI <> "" And rowLast.Length = 0) Then
                        DIFFERENCE = "Added"
                        dst.Tables("differences").Rows.Add(New Object() {TABLE_NAME, key, COLUMN_NAME, DIFFERENCE})

                    Else
                        For Each dc As DataColumn In dst.Tables(TABLE_NAME).Columns
                            COLUMN_NAME = dc.ColumnName
                            If rowPrev(0).Item(dc.ColumnName).Equals(rowCurr(0).Item(dc.ColumnName)) And _
                               (VAN_REF_last_AI = "" OrElse rowLast(0).Item(dc.ColumnName).Equals(rowCurr(0).Item(dc.ColumnName))) Then
                                ' DO NOTHING
                            Else
                                cols.Add(dc.ColumnName)
                                DIFFERENCE = "Value"
                                dst.Tables("differences").Rows.Add(New Object() {TABLE_NAME, key, COLUMN_NAME, DIFFERENCE})
                            End If
                        Next
                    End If
                End If
            Next

            gcols_different.Add("grd" & TABLE_NAME, cols)

        Next
    End Sub

    Private Sub grdPOTORDRB_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdPOTORDRB.InitializeRow
        Color_Version(e, grdPOTORDRB)

        Dim STATUS As String = e.Row.Cells("STATUS").Value & ""
        If STATUS = "A" Then
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Green
        ElseIf STATUS = "R" Then
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf STATUS = "I" Then
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Purple
        Else
            e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub

    Sub Fix_Styles()

        EnforceConstraints(False)
        Fill_Records("ICTSTYLF")
        Fill_Records("ICTSTYCF")

        dst.Tables("ICTSTYLS").Rows.Clear()

        EnforceConstraints(True)

        ASCMAIN1.Progress("Now Fixing Colors and Sizes")
        For Each row As DataRow In dst.Tables("ICTSTYLF").Select("")
            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
            'If STYLE_CODE = "WM173214" Then Stop
            If STYLE_CODE = "EG400132" Then Stop
            ASCMAIN1.Progress("-", STYLE_CODE)
            Do While Fix_Colors(STYLE_CODE)

            Loop
            Fix_Size(STYLE_CODE)
        Next

        grdICTSTYLF.DisplayLayout.Bands(0).Columns("SIZE_SCALE_ORIG").Header.Caption = "Original Colors/Sizes"
        grdICTSTYLF.DisplayLayout.Bands(0).Columns("SIZE_SCALE").Header.Caption = "Unresolved"
        grdICTSTYLF.DisplayLayout.Bands(0).Columns("BORN").Width = 90
        grdICTSTYLF.DisplayLayout.Bands(0).Columns("STYLE_CODE").Width = 90
        grdICTSTYLF.DisplayLayout.Bands(0).Columns("STYLE_DESC").Width = 200

        grdICTSTYLF.DisplayLayout.Bands(0).Columns("SQ").Hidden = True

        grdICTSTYLF.DisplayLayout.Bands(1).Columns("STYLE_CODE").Hidden = True
        grdICTSTYLF.DisplayLayout.Bands(1).Columns("STYLE_COLOR_DESC").ColSpan = 2
        grdICTSTYLF.DisplayLayout.Bands(1).Columns("STYLE_COLOR_DESC").Header.Caption = "Pretty Color Description"

        For SI As Integer = 12 To 1 Step -1
            Dim SX As String = "S" & CStr(SI)
            Dim QX As String = "Q" & CStr(SI)
            grdICTSTYLF.DisplayLayout.Bands(0).Columns(SX).Width = 40
            grdICTSTYLF.DisplayLayout.Bands(0).Columns(QX).Width = 40
            If dst.Tables("ICTSTYLF").Select("isnull(" & SX & ",'') <> ''").Length = 0 Then
                grdICTSTYLF.DisplayLayout.Bands(0).Columns(SX).Hidden = True
                grdICTSTYLF.DisplayLayout.Bands(0).Columns(QX).Hidden = True
            Else
                grdICTSTYLF.DisplayLayout.Bands(0).Columns(SX).Hidden = False
                grdICTSTYLF.DisplayLayout.Bands(0).Columns(QX).Hidden = False
            End If
        Next

        Sort_grdColumns(grdICTSTYLF, "STYLE_CODE")

        UltraExplorerBar1.Groups("Display Options").Visible = False


        grdPOTORDRX.Visible = False
        grdICTSTYLF.Visible = True

        Stop

        ' ICTSTYCF IS A BRAND NEW TABLE I JUST CREATED FOR THIS PURPOSE - AND ITS USE IS ONLY FOR STAGING THESE UPDATES - IT SHOULD BE DROPPED SOMETIME AFTER FIX IS IN -IT IS NOW DROPPED
        ' ICTSTYLF IS A BRAND NEW TABLE I JUST CREATED FOR THIS PURPOSE - AND ITS USE IS ONLY FOR GETTING UPDATES INTO ICTSTYLS - IT SHOULD BE DROPPED SOMETIME AFTER FIX IS IN -IT IS NOW DROPPED
        ' ICTSTYLS IS A BRAND NEW TABLE I CREATED TO HOLD ONTO SIZE SCALES BY STYLE, AND IT IS IMPLEMENTED IN PRODUCTION ICTSTYL1
        ' AND ICTSTYS1 IS A BRAND NEW TABLE I JUST CREATED FOR THIS PURPOSE - BUT APPARENTLY DON'T EVEN USE - IT SHOULD BE DROPPED -IT IS NOW DROPPED

        'ASCMAIN1.sql = "Create Table ICTSTYS1 (" _
        '    & "STYLE_CODE VARCHAR2(16)," _
        '    & "SIZE_INDEX NUMBER(3)," _
        '    & "SIZE_CODE VARCHAR2(6)," _
        '    & "SIZE_QTY NUMBER(3)," _
        '    & "Primary Key (STYLE_CODE, SIZE_INDEX))"


        ASCMAIN1.sql = "Delete from ICTSTYLF"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ICTSTYCF"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from ICTSTYLS"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("ICTSTYLF", "1=1")

        Update_Record_TDA("ICTSTYCF", "1=1")
        Update_Record_TDA("ICTSTYLS", "1=1")

        Stop

        ' THIS WILL REALLY UPDATE THINGS
        ' ON 01/08/19 I VERIFIED THAT THERE WERE 5 RECORDS IN ICTSTYC1 
        ' WITH CONTENT IN STYLE_COLOR_DESC, SO I AM CLAIMING THIS FIELD FOR MY OWN
        'STYLE_CODE	    COLOR_CODE	STYLE_COLOR_DESC
        '447FFFMX	    531	 
        'VC1035116	    270	        270
        '500616HBI	    100	        001
        '501453XIZ38C	270	 
        '500555VBI	    280	        100

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM ICTSTYCF;" & vbCrLf _
            & "BEGIN " & vbCrLf _
            & "UPDATE ICTSTYC1 SET STYLE_COLOR_DESC = NULL;" & vbCrLf _
            & "UPDATE ICTSTYL1 SET SIZE_CODE = NULL;" & vbCrLf _
            & "UPDATE ICTSTYL1 SET SIZE_CODE = SUBSTR((SELECT STYLE_SIZE FROM ICTSTYLS where STYLE_CODE = ICTSTYL1.STYLE_CODE),1,6) where STYLE_CODE IN (SELECT STYLE_CODE FROM ICTSTYLS WHERE STYLE_SIZE IS NOT NULL AND STYLE_SIZE <> 'ASSORTED');" & vbCrLf _
            & "FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ICTSTYC1 SET STYLE_COLOR_DESC = R1.STYLE_COLOR_DESC" & vbCrLf _
            & "WHERE STYLE_CODE = R1.STYLE_CODE AND COLOR_CODE = R1.COLOR_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " & vbCrLf _
            & "select STYLE_CODE, STYLE_SIZE from ictstylS where STYLE_SIZE IS NOT NULL AND STYLE_SIZE LIKE 'SIZE %';" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "UPDATE ICTSTYL1 SET SIZE_CODE = SUBSTR(R1.STYLE_SIZE,6) WHERE STYLE_CODE = R1.STYLE_CODE;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        MsgBox("Fix is In")
        ASCMAIN1.Progress("")
    End Sub


    Sub Update_Fix()

    End Sub

    Sub Get_PDFs()
        Dim FOLDER_FROM As String = "\\192.168.160.100\UploadFolder\"

        For Each row As DataRow In dst.Tables("POTORDRX").Select("")
            Dim POKey As String = row.Item("POKEY")
            Dim PONO As String = row.Item("PONO")

            ASCMAIN1.Progress(PONO)
            Dim FOLDER_TO As String = PO_PARM_PO_IMG_DIR & "\" & POKey & "\"


            For Each filename2copy As String In System.IO.Directory.GetFiles(FOLDER_FROM & POKey & "\")
                If filename2copy.ToUpper.EndsWith(".PDF") Then
                    Dim path2pieces() As String = filename2copy.Split("\")
                    Dim leafname As String = path2pieces(path2pieces.Length - 1)
                    '  Console.WriteLine("Copying " & filename2copy & " to " & FOLDER_TO & FILENAME)
                    Try
                        System.IO.File.Copy(filename2copy, FOLDER_TO & leafname, False)
                    Catch ex As Exception
                        ' Console.WriteLine("Error " & ex.Message & " Copying " & filename2copy & " to " & FOLDER_TO & leafname)
                    End Try
                End If
            Next

        Next
    End Sub

    Sub Fix_Size(STYLE_CODE As String)
        Dim rowICTSTYLF As DataRow = dst.Tables("ICTSTYLF").Rows.Find(STYLE_CODE)

        Dim SIZE_SCALE As String = rowICTSTYLF.Item("SIZE_SCALE") & ""

        Dim rowICTSTYLS As DataRow = dst.Tables("ICTSTYLS").NewRow
        rowICTSTYLS.Item("STYLE_CODE") = STYLE_CODE
        rowICTSTYLS.Item("SIZE_SCALE") = SIZE_SCALE
        dst.Tables("ICTSTYLS").Rows.Add(rowICTSTYLS)

        Do While SIZE_SCALE.EndsWith(vbCrLf)
            SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 2))
        Loop

        If InStr(SIZE_SCALE, vbCrLf) = 0 Then ' THE REMAINDER IS ALL ON 1 LINE

            'If STYLE_CODE = "41001AKMM" Then Stop
            'If STYLE_CODE = "61005EKMXL" Then Stop
            'If STYLE_CODE = "JBS1005AKM" Then Stop
            'If STYLE_CODE = "3146KM38D" Then Stop
            'If STYLE_CODE = "501696XMB" Then Stop

            If STYLE_CODE = "EG400132" Then Stop

            If SIZE_SCALE.EndsWith("PCS)") Then
                If Mid(SIZE_SCALE, Len(SIZE_SCALE) - 5, 1) = "(" Then
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, Len(SIZE_SCALE) - 6))
                ElseIf Mid(SIZE_SCALE, Len(SIZE_SCALE) - 6, 1) = "(" Then
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, Len(SIZE_SCALE) - 7))
                End If

            End If

            SIZE_SCALE = Replace(SIZE_SCALE, vbTab, " ")
            SIZE_SCALE = Replace(SIZE_SCALE, " - ", " ")
            Do While InStr(SIZE_SCALE, "  ") <> 0
                SIZE_SCALE = Replace(SIZE_SCALE, "  ", " ")
            Loop

            SIZE_SCALE = Trim(SIZE_SCALE)
            If SIZE_SCALE = "S M L XL 1-2-2-1" Then
                SIZE_SCALE = "S M L XL = 1-2-2-1"
            End If
            If SIZE_SCALE = "LARGE   2/2/2/2 (S/M/L/XL)" Then
                SIZE_SCALE = "LARGE (S/M/L/XL) 2/2/2/2"
            End If
            If SIZE_SCALE = "MEDIUM (S-M-L-XL-XXL = 1/2/2/1/1)" Then
                SIZE_SCALE = "MEDIUM (S-M-L-XL-XXL) = 1/2/2/1/1"
            End If
            If SIZE_SCALE = "X-LARGE (S/M/L/XL/XXL = 1-2-2-2-1)" Then
                SIZE_SCALE = "X-LARGE (S/M/L/XL/XXL) = 1-2-2-2-1"
            End If
            If SIZE_SCALE = "X-LARGE (S/M/L/XL/XXL = 1-2-2-1-1)" Then
                SIZE_SCALE = "X-LARGE (S/M/L/XL/XXL) = 1-2-2-1-1"
            End If
            If SIZE_SCALE.EndsWith("_") And InStr(SIZE_SCALE, "(") <> 0 And InStr(SIZE_SCALE, ")") = 0 Then
                Mid(SIZE_SCALE, Len(SIZE_SCALE), 1) = ")"
            End If




            Dim STYLE_SIZE As String = ""

            Dim P1 As Integer = InStr(SIZE_SCALE, "(")
            Dim P2 As Integer = InStr(SIZE_SCALE, ")")
            If P1 <> 0 And P2 <> 0 And P2 > P1 And P1 <= 11 Then
                ' If SIZE_SCALE.StartsWith("LARGE ") Then Stop



                Dim SIZE_SCALE_INNER As String = Mid(SIZE_SCALE, P1 + 1, P2 - P1 - 1)
                If System.Text.RegularExpressions.Regex.IsMatch(Replace(Replace(Replace(SIZE_SCALE_INNER, " ", ""), "-", ""), "/", ""), "^[A-Za-z0-9]+$") Then
                    STYLE_SIZE = Mid(SIZE_SCALE, 1, P1 - 1).Trim
                    rowICTSTYLS.Item("STYLE_SIZE") = STYLE_SIZE
                    SIZE_SCALE = Mid(SIZE_SCALE, P1 + 1, P2 - P1 - 1)
                End If
            End If

            If SIZE_SCALE.StartsWith("SMALL S/M/L") Or SIZE_SCALE.StartsWith("MEDIUM S/M/L") Or SIZE_SCALE.StartsWith("LARGE S/M/L") Or SIZE_SCALE.StartsWith("X-LARGE S/M/L") Or SIZE_SCALE.StartsWith("2X-LARGE S/M/L") Then
                P1 = InStr(SIZE_SCALE, " ")
                STYLE_SIZE = Mid(SIZE_SCALE, 1, P1 - 1).Trim
                SIZE_SCALE = Mid(SIZE_SCALE, P1 + 1).Trim
            End If

            rowICTSTYLS.Item("STYLE_SIZE") = STYLE_SIZE

            Dim SIZE_CODEs As New List(Of String)
            Dim SIZE_QTYs As New List(Of Integer)

            If SIZE_SCALE <> "" Then
                If SIZE_SCALE.StartsWith("SIZE ") Then
                    SIZE_SCALE = Mid(SIZE_SCALE, 6)
                ElseIf SIZE_SCALE.StartsWith("SIZE: ") Then
                    SIZE_SCALE = Mid(SIZE_SCALE, 7)
                End If

                If SIZE_SCALE.EndsWith(")") And InStr(SIZE_SCALE, " (") > 0 And InStr(SIZE_SCALE, " (") < 7 Then
                    Dim SPACEI As Integer = InStr(SIZE_SCALE, " (")
                    SIZE_SCALE = Mid(SIZE_SCALE, SPACEI + 2)
                    SIZE_SCALE = Mid(SIZE_SCALE, 1, Len(SIZE_SCALE) - 1)
                End If
            End If


            Try
                If InStr(SIZE_SCALE, "=") <> 0 Then
                    Dim sizes As String = Trim(Split(SIZE_SCALE, "=")(0))
                    Do While InStr(sizes, "//") <> 0
                        sizes = Replace(sizes, "//", "/")
                    Loop
                    If InStr(sizes, "/") = 0 And InStr(sizes, "-") <> 0 Then
                        Do While InStr(sizes, "-") <> 0
                            sizes = Replace(sizes, "-", "/")
                        Loop
                    End If

                    Do While InStr(sizes, vbTab & vbTab) <> 0
                        sizes = Replace(sizes, vbTab & vbTab, vbTab)
                    Loop
                    Do While InStr(sizes, "  ") <> 0
                        sizes = Replace(sizes, "  ", " ")
                    Loop

                    Do While InStr(sizes, "/") <> 0
                        sizes = Replace(sizes, "/", " ")
                    Loop


                    Dim qtys As String = Trim(Split(SIZE_SCALE, "=")(1))
                    Dim dl As String = " "
                    Dim dlQ As String = "/"

                    If InStr(sizes, vbTab) <> 0 Then dl = vbTab
                    If InStr(sizes, "-") <> 0 Then dl = "-"


                    If InStr(qtys, "/") = 0 And InStr(qtys, "-") <> 0 Then qtys = Replace(qtys, "-", "/")

                    Dim SI As Integer = 0
                    For Each S As String In Split(sizes, dl)
                        SI += 1
                        If SI > 6 Then
                            Debug.Print(STYLE_CODE)
                            Exit For
                        End If

                        If InStr(S, " ") <> 0 Then
                            S = Mid(S, 1, InStr(S, " ") - 1)
                        End If
                        rowICTSTYLF.Item("S" & CStr(SI)) = S
                        rowICTSTYLF.Item("Q" & CStr(SI)) = Val(Split(qtys & "//////", "/")(SI - 1) & "")

                        rowICTSTYLS.Item("SIZE_" & Format(SI, "00")) = rowICTSTYLF.Item("S" & CStr(SI))
                        rowICTSTYLS.Item("QTY_" & Format(SI, "00")) = rowICTSTYLF.Item("Q" & CStr(SI))
                    Next
                    rowICTSTYLF.Item("SIZE_SCALE") = ""
                Else
                    Do While InStr(SIZE_SCALE, "  ") <> 0
                        SIZE_SCALE = Replace(SIZE_SCALE, "  ", " ")
                    Loop
                    SIZE_SCALE = Trim(SIZE_SCALE)
                    If Not SIZE_SCALE.Contains(" ") And SIZE_SCALE.Contains("/") Then
                        SIZE_SCALE = Replace(SIZE_SCALE, "/", " ")

                    End If

                    Dim SS() As String = Split(SIZE_SCALE, " ")
                    Dim SI As Integer = 0
                    For Each S As String In SS
                        If S.Length >= 3 AndAlso Mid(S, S.Length - 1, 1) = "/" AndAlso InStr("123456789", Mid(S, S.Length, 1)) <> 0 Then
                            SIZE_CODEs.Add(Mid(S, 1, S.Length - 2))
                            SIZE_QTYs.Add(Val(Mid(S, S.Length, 1)))
                            SI += 1
                            If SI > 6 Then
                                Debug.Print(STYLE_CODE)
                                Exit For
                            End If

                            'If Len(S) = 7 AndAlso (S.EndsWith("DD") And S.Contains("X-")) Then
                            '    S = Mid(S, 1, 2) & Mid(S, 3)
                            'End If
                            'If S = "X-LARGE" Then S = "XL"
                            'If S = "2X-LARGE" Then S = "XXL"
                            'If S = "ASSORTED" Then S = "AST"

                            rowICTSTYLF.Item("S" & CStr(SI)) = Mid(S, 1, S.Length - 2)
                            rowICTSTYLF.Item("Q" & CStr(SI)) = Val(Mid(S, S.Length, 1))

                            rowICTSTYLS.Item("SIZE_" & Format(SI, "00")) = rowICTSTYLF.Item("S" & CStr(SI))
                            rowICTSTYLS.Item("QTY_" & Format(SI, "00")) = rowICTSTYLF.Item("Q" & CStr(SI))

                            Dim SX As Integer = InStr(SIZE_SCALE, S)
                            SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SX - 1) & Mid(SIZE_SCALE, SX + S.Length))

                        End If
                        SIZE_SCALE = Trim(SIZE_SCALE)
                    Next

                    If Trim(Replace(Replace(SIZE_SCALE, " ", ""), "-", "")) = "" Then
                        SIZE_SCALE = ""
                    End If

                    Do While SIZE_SCALE.Contains("--")
                        SIZE_SCALE = Replace(SIZE_SCALE, "--", "-") 'VC1055101
                    Loop
                    If SI = 0 And SIZE_SCALE <> "" Then
                        If SIZE_SCALE = "S-M-L-XL" Or SIZE_SCALE = "S-M-L" _
                            Or SIZE_SCALE = "34B-34C-36B-36C-38B-38C-38D-40B-40C" _
                            Or SIZE_SCALE = "34A-34B-34C-34D-34DD-36B-36C-36D-36DD-38C" _
                            Or SIZE_SCALE = "34A 34B 34C 34D 34DD 36B 36C 36D 36DD 38C" _
                            Or SIZE_SCALE = "34A-34B-36B-36C-38B-38C" _
                            Or SIZE_SCALE = "34B-34C-34D-36B-36C-36D-38C-38D" _
                            Or SIZE_SCALE = "34B-34C-36B-36C-38B-38C-38D-40B-40C" _
                            Or SIZE_SCALE = "34B-34C-36B-36C-38C" _
                            Or SIZE_SCALE = "34B-36B-34C-36C-36D-38C" _
                            Or SIZE_SCALE = "34B-36B-36C-36D-38C-38D" _
                            Or SIZE_SCALE = "34B-36B-36C-38C" _
                            Or SIZE_SCALE = "34B-36B-36C-38C-36D-38D" _
                            Or SIZE_SCALE = "36B-36C-38C-36D" _
                            Or SIZE_SCALE = "32B-32C-32D-32DD-34B-34C-34D-34DD-36B-36C-36D-36DD" _
                            Or SIZE_SCALE = "2X-3X" _
                            Or SIZE_SCALE = "40D-40DD-42C-42D-42DD-44C-44D-44DD" _
                            Or SIZE_SCALE = "40C-40D-42D-38DD-40DD-42DD" _
                            Or SIZE_SCALE = "38DD-40C-40D-40DD-42D-42DD" _
                            Or SIZE_SCALE = "38D-40D-42D-40DD" _
                            Or SIZE_SCALE = "38D-40D-40DD" _
                            Or SIZE_SCALE = "36D-38D-40D-42D-40DD-42DD" _
                            Or SIZE_SCALE = "36B-36B-38C-36D" _
                            Or SIZE_SCALE = "36B-36C-36D-38C" _
                            Or SIZE_SCALE = "34A-34B-34C-36B-36C-38C" _
                            Or SIZE_SCALE = "1X - 2X - 3X" _
                            Or SIZE_SCALE = "1X-2X-3X" _
                            Or SIZE_SCALE = "1X - 2X" _
                            Or SIZE_SCALE = "1X" _
                            Or (Replace(Replace(SIZE_SCALE, " ", ""), "-", "").Split("-").Length <= 12 And System.Text.RegularExpressions.Regex.IsMatch(Replace(Replace(SIZE_SCALE, " ", ""), "-", ""), "^[A-Za-z0-9]+$")) Then

                            SIZE_SCALE = Replace(SIZE_SCALE, " - ", " ")
                            SIZE_SCALE = Replace(SIZE_SCALE, "  ", " ")
                            SIZE_SCALE = Replace(SIZE_SCALE, " ", "-")
                            Do While SIZE_SCALE.Contains("--")
                                SIZE_SCALE = Replace(SIZE_SCALE, "--", "-") 'VC1055101
                            Loop
                            For Each S As String In Split(SIZE_SCALE, "-")
                                SIZE_CODEs.Add(S)
                                SIZE_QTYs.Add(0)
                                SI += 1
                                'If SI > 6 Then
                                '    Debug.Print(STYLE_CODE)
                                '    Exit For
                                'End If

                                If SI <= 12 Then
                                    If Len(S) = 7 AndAlso (S.EndsWith("DD") And S.Contains("X-")) Then
                                        S = Mid(S, 1, 2) & Mid(S, 3)
                                    End If
                                    If S = "X-LARGE" Then S = "XL"
                                    If S = "2X-LARGE" Then S = "XXL"
                                    If S = "ASSORTED" Then S = "AST"
                                    rowICTSTYLS.Item("SIZE_" & Format(SI, "00")) = S
                                    rowICTSTYLS.Item("QTY_" & Format(SI, "00")) = 0
                                End If


                                rowICTSTYLF.Item("S" & CStr(SI)) = S ' Mid(S, 1, S.Length - 2)
                                rowICTSTYLF.Item("Q" & CStr(SI)) = 0 ' Val(Mid(S, S.Length, 1))
                            Next
                            SIZE_SCALE = ""
                        End If
                    End If

                    rowICTSTYLF.Item("SIZE_SCALE") = SIZE_SCALE

                End If
            Catch ex As Exception
                For SI As Integer = 1 To 12

                    rowICTSTYLF.Item("S" & CStr(SI)) = DBNull.Value
                    rowICTSTYLF.Item("Q" & CStr(SI)) = DBNull.Value


                    rowICTSTYLS.Item("SIZE_" & Format(SI, "00")) = rowICTSTYLF.Item("S" & CStr(SI))
                    rowICTSTYLS.Item("QTY_" & Format(SI, "00")) = rowICTSTYLF.Item("Q" & CStr(SI))
                Next
            End Try

        End If
    End Sub

    Function Fix_Colors(STYLE_CODE As String) As Boolean
        Dim fixed As Boolean = False
        Dim rowICTSTYLF As DataRow = dst.Tables("ICTSTYLF").Rows.Find(STYLE_CODE)
        Dim SIZE_SCALE As String = rowICTSTYLF.Item("SIZE_SCALE") & ""
        Dim MAX_LENGTH As Integer = 60
        If SIZE_SCALE <> "" Then
            Dim COLOR_CODEs As New List(Of String)
            For Each row As DataRow In rowICTSTYLF.GetChildRows("ICTSTYLF_ICTSTYCF")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)

                Dim I As Integer = InStr(SIZE_SCALE, COLOR_CODE)
                If I <> 0 Then
                    Dim S As String = Trim(Mid(SIZE_SCALE, I + 3))
                    Dim J As Integer = InStr(Mid(S & "  ", 1, MAX_LENGTH), "  ")
                    Dim K As Integer = InStr(Mid(S & vbCrLf, 1, MAX_LENGTH), vbCrLf)
                    If J = 0 And K = 0 Then
                        J = InStr(Mid(S & " ", 1, MAX_LENGTH), " ")
                    End If
                    If J = 0 Or J > K Then J = K
                    Dim SC As String = ""
                    If J <> 0 Then
                        fixed = True
                        SC = Mid(S, 1, J)
                        SIZE_SCALE = Mid(SIZE_SCALE, 1, I - 1) & Mid(S, J)
                        For C As Integer = 1 To SC.Length - 1
                            If C = 1 Or (C > 1 AndAlso Mid(SC, C + 1, 1) <> " " AndAlso (Mid(SC, C - 1, 1) = " " Or Mid(SC, C - 1, 1) = "/")) Then
                                Mid(SC, C, 1) = Mid(SC, C, 1).ToUpper
                            End If
                        Next
                        If Trim(SC) <> "" Then
                            row.Item("STYLE_COLOR_DESC") = SC
                        End If
                    End If
                End If
            Next
            Dim TF As Boolean = False
            Do
                TF = False
                Do While InStr(SIZE_SCALE, vbCrLf & vbCrLf) <> 0
                    SIZE_SCALE = Replace(SIZE_SCALE, vbCrLf & vbCrLf, vbCrLf)
                    TF = True
                Loop
                Do While SIZE_SCALE.EndsWith(vbCrLf)
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 2))
                    TF = True
                Loop
                Do While SIZE_SCALE.EndsWith("#")
                    SIZE_SCALE = Trim(Mid(SIZE_SCALE, 1, SIZE_SCALE.Length - 1))
                    TF = True
                Loop
            Loop While TF

            rowICTSTYLF.Item("SIZE_SCALE") = Trim(SIZE_SCALE)
        End If
    End Function

    Sub ShowTiles()

        tplStyle.Tiles.Clear()

        For Each row As DataRow In dst.Tables("Images").Select("")
            Dim I As System.Drawing.Bitmap = row.Item("IMAGE")
            Dim IMAGE_DESC As String = row.Item("IMAGE_DESC").ToString
            Dim IMAGE_TYPE As String = row.Item("IMAGE_TYPE").ToString

            If IMAGE_DESC = "" Then IMAGE_DESC = IMAGE_TYPE

            Dim t As New Infragistics.Win.Misc.UltraTile
            Dim P As New UltraWinEditors.UltraPictureBox
            P.Image = I
            t.Control = P
            t.Text = IMAGE_DESC
            tplStyle.Tiles.Add(t)
            P.Visible = True
            t.Visible = True
        Next
 
        tplStyle.Visible = True

    End Sub

    Private Sub grdImages_AfterRowActivate(sender As Object, e As EventArgs) Handles grdImages.AfterRowActivate
        Show_Image()
    End Sub

    Sub Show_Image()
        If grdImages.ActiveRow Is Nothing Then
            imgSTYLE.Image = Nothing
            imgSTYLE.Visible = False
        Else
            Dim FILENAME As String = grdImages.ActiveRow.Cells("FILENAME").Value & ""
            imgSTYLE.Image = ASCMAIN1.Get_Image(PO_PARM_PO_IMG_DIR, FILENAME, True, , , ) ' imgba)
            imgSTYLE.Visible = True
        End If
    End Sub

    Sub Get_Images()

        dst.Tables("Images").Rows.Clear()

        Get_Image(rowpohdr, "PolyBag", "Poly Bag", "PolyBagImg")
        Get_Image(rowpohdr, "ShippingMark", "Shipping Mark", "ShippingMarkImg")
        Get_Image(rowpohdr, "Packing", "Packing", "PackingImg")
        Get_Image(rowpohdr, "Sample", "Sample", "SampleImg")
        Get_Image(rowpohdr, "", "Picture1", "PictureName1")
        Get_Image(rowpohdr, "", "Picture2", "PictureName2")

        For Each rowpotrim As DataRow In dst.Tables("potrim").Select("VAN_REF = '" & VAN_REF & "'")
            Get_Image(rowpotrim, "ItemDesc", "Trim Item", "PictureName")
        Next

        ShowTiles()

        Dim POKey As String = rowpohdr.Item("POKey")
        If System.IO.Directory.Exists(PO_PARM_PO_IMG_DIR & "\" & POKey) Then

            For Each FILENAME As String In System.IO.Directory.GetFiles(PO_PARM_PO_IMG_DIR & "\" & POKey)
                If FILENAME.ToUpper.EndsWith(".PDF") Then
                    Dim path2pieces() As String = FILENAME.Split("\")
                    Dim leafname As String = path2pieces(path2pieces.Length - 1)

                    Dim rowPDFs As DataRow = dst.Tables("PDFs").NewRow
                    rowPDFs.Item("FILENAME") = leafname
                    Dim dt As DateTime = System.IO.Directory.GetCreationTime(FILENAME)
                    rowPDFs.Item("FILEDATETIME") = dt
                    dst.Tables("PDFs").Rows.Add(rowPDFs)
                End If
                '     Dim IMAGE As System.Drawing.Bitmap = ASCMAIN1.Get_Image(PO_PARM_PO_IMG_DIR, FILENAME, True, , , )
            Next
        End If

    End Sub

    Sub Get_Image(row As DataRow, COLUMN_NAME_DESC As String, IMAGE_TYPE As String, COLUMN_NAME_FILENAME As String)

        Dim FILENAME As String = row.Item(COLUMN_NAME_FILENAME) & ""
        If FILENAME = "" Then Return
        'C:\Ashley-System\UploadFolder\190\IZ-BRA(SML) POLYBAG.jpg
        Dim f As Integer = FILENAME.IndexOf("\UploadFolder\")
        If f = 0 Then Return
        FILENAME = FILENAME.Substring(f + "\UploadFolder\".Length)

        
        Dim IMAGE_DESC As String = ""
        If COLUMN_NAME_DESC <> "" Then
            IMAGE_DESC = row.Item(COLUMN_NAME_DESC) & ""
        End If

        'Dim FOLDER_FROM As String = "\\192.168.160.100\UploadFolder\"
        'Dim FOLDER_TO As String = PO_PARM_PO_IMG_DIR & "\"
        'My.Computer.FileSystem.CopyFile(FOLDER_FROM & FILENAME, FOLDER_TO & FILENAME, True)

        Dim IMAGE As System.Drawing.Bitmap = ASCMAIN1.Get_Image(PO_PARM_PO_IMG_DIR, FILENAME, True, , , )

        Dim rowImages As DataRow = dst.Tables("Images").NewRow
        rowImages.Item("FILENAME") = FILENAME
        rowImages.Item("IMAGE_TYPE") = IMAGE_TYPE
        rowImages.Item("IMAGE_DESC") = IMAGE_DESC
        rowImages.Item("SOURCE") = row.Table.TableName
        rowImages.Item("IMAGE") = IMAGE
        If IMAGE_TYPE = "Trim Item" Then
            rowImages.Item("POKey") = row.Item("POKey")
            rowImages.Item("POTrimKey") = row.Item("POTrimKey")
        End If

        dst.Tables("Images").Rows.Add(rowImages)
    End Sub
     
    Private Sub btnCreateStyle_Click(sender As Object, e As EventArgs) Handles btnCreateStyle.Click
        Dim STYLE_CODE As String = rowpohdr.Item("StyleNo")
        STYLE_CODE = STYLE_CODE.Replace("-", "")

        Easy_Button(STYLE_CODE)
         
    End Sub

    Sub Easy_Button(STYLE_CODE As String, Optional StyleBySize As Boolean = False, Optional SIZE_CODE As String = "")

        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)

        Dim fm_mode As String = "Edit"
        If rowICTSTYL1 Is Nothing Then
            fm_mode = "New"
        End If

        If frmASFBASE1s.ContainsKey("ICTSTYL1") AndAlso frmASFBASE1s("ICTSTYL1").ScreenMode Then
            MsgBox("There is a style already being edited by the Easy Button" & vbCrLf & vbCrLf _
                   & "Either Cancel or Update those edits and then retry the Easy Button", _
                   MsgBoxStyle.OkOnly, "Cannot Edit using the Easy Button at this time")
            Exit Sub
        End If

        Dim keys As New Dictionary(Of String, Object)
        keys.Add("STYLE_CODE", STYLE_CODE)

        Dim frmASFBASE1 As ASFBASE1 = Context_Launch(fm_mode, keys, "Style Master File", "ICTSTYL1", , , True)

        If frmASFBASE1s("ICTSTYL1").ScreenMode Then

            Dim keysICTSTYL1 As New Dictionary(Of String, Object)
            Dim sqlv As String = "VAN_REF = '" & VAN_REF & "'"

            ' Material Content

            If dst.Tables("pofabric").Select(sqlv & " and ISNULL(Description,'') <> ''").Length > 0 Then
                'Dim pofabric As DataRow = dst.Tables("pofabric").Select(sqlv & " and ISNULL(Description,'') <> ''")(0)
                'keysICTSTYL1.Add("STYLE_MATL_DESC", pofabric.Item("Description") & "")

                Dim STYLE_MATL_DESC As String = ""
                For Each rowfabric As DataRow In dst.Tables("pofabric").Select(sqlv & " and ISNULL(Description,'') <> ''")

                    STYLE_MATL_DESC &= vbCrLf & rowfabric.Item("Item") & ":" & rowfabric.Item("Description") & ""

                Next
                keysICTSTYL1.Add("STYLE_MATL_DESC", Mid(STYLE_MATL_DESC, 3))

            End If

            ' Factory, Country, Vendor

            keysICTSTYL1.Add("FACTORY_CODE", rowpohdr.Item("Factory") & "")
            keysICTSTYL1.Add("COUNTRY_CODE", "CHI")
            keysICTSTYL1.Add("VEND_CODE", "AT")

            ' Size Scale

            Dim STYLE_COLOR_DESC_maxLength As Integer = frmASFBASE1s("ICTSTYL1").dst.Tables("ICTSTYC1").Columns("STYLE_COLOR_DESC").MaxLength

            Dim row As DataRow
            Dim szs() As String = Nothing
            Dim szq() As String = Nothing

            If StyleBySize Then

                keysICTSTYL1.Add("SIZE_CODE", SIZE_CODE)

                Dim rowsStyleSize() As DataRow = dst.Tables("StyleSize").Select("VAN_REF = '" & VAN_REF & "'")
                If rowsStyleSize.Length = 1 Then
                    Dim szsC As New List(Of String)
                    Dim szqC As New List(Of String)
                    For i As Integer = 1 To 12
                        Dim SIZE As String = rowsStyleSize(0).Item("SIZE_" & Format(i, "00")) & ""
                        If SIZE = "" Then
                            Exit For
                        Else
                            szsC.Add(SIZE)
                            szqC.Add(rowsStyleSize(0).Item("QTY_" & Format(i, "00")) & "")
                        End If
                    Next
                    szs = szsC.ToArray
                    szq = szqC.ToArray
                End If
            Else


                'Dim posize() As DataRow = dst.Tables("posize").Select(sqlv & " and Color='Size:'")
                'Dim posizeqty() As DataRow = dst.Tables("posize").Select(sqlv & " and Color is null")
                'If posize.Length = 1 And posizeqty.Length = 1 Then
                '    Dim sizes As String = posize(0).Item("Size") & ""
                '    szs = Split(sizes, "|")
                '    Dim qtys As String = posizeqty(0).Item("Size") & ""
                '    szq = Split(qtys, "|")
                'End If
                
                Dim ColorKeysSizes As New Dictionary(Of String, List(Of String))
                Dim ColorKeysPrePacks As New Dictionary(Of String, List(Of String))
                For Each rowsizedtl As DataRow In dst.Tables("posizedtl").Select(sqlv & " and IsNull(Size,'') <> ''", "POColorKey, Style, POSizeKey", DataViewRowState.ModifiedCurrent)
                    Dim ColorKey As String = rowsizedtl.Item("POColorKey") & ""
                    Dim Size As String = rowsizedtl.Item("Size") & ""
                    Dim PrePack As Integer = rowsizedtl.Item("PrePack") & ""
                    If Not ColorKeysSizes.ContainsKey(ColorKey) Then
                        ColorKeysSizes.Add(ColorKey, New List(Of String))
                        ColorKeysPrePacks.Add(ColorKey, New List(Of String))
                    End If
                    ColorKeysSizes(ColorKey).Add(Size)
                    ColorKeysPrePacks(ColorKey).Add(PrePack)
                Next
                If ColorKeysSizes.Count > 1 Then
                    szs = ColorKeysSizes(ColorKeysSizes.Keys(0)).ToArray
                    szq = ColorKeysPrePacks(ColorKeysSizes.Keys(0)).ToArray
                End If

            End If

            If szs IsNot Nothing AndAlso szs.Length > 0 Then
                Dim tbl As DataTable = frmASFBASE1s("ICTSTYL1").dst.Tables("ICTSTYLS")
                If tbl.Rows.Count = 0 Then
                    row = tbl.NewRow
                    row.Item("STYLE_CODE") = STYLE_CODE
                    tbl.Rows.Add(row)
                End If

                If tbl.Rows.Count = 1 Then
                    row = tbl.Rows(0)
                    If szs.Length > 0 Then
                        For s As Integer = 1 To szs.Length
                            row.Item("SIZE_" & Format(s, "00")) = szs(s - 1)
                            row.Item("QTY_" & Format(s, "00")) = Val(szq(s - 1))
                        Next
                    End If
                    If szs.Length > 0 And szs.Length < 12 Then
                        For s As Integer = szs.Length + 1 To 12
                            row.Item("SIZE_" & Format(s, "00")) = DBNull.Value
                            row.Item("QTY_" & Format(s, "00")) = DBNull.Value
                        Next
                    End If
                End If
            End If


            ' Extended Color Descriptions

            Dim COLOR_CODEs As New List(Of String)
            Dim bad_COLOR_CODEs As New List(Of String)
            For Each rowpocolor As DataRow In dst.Tables("pocolor").Select(sqlv)
                Dim COLOR_CODE As String = rowpocolor.Item("ColorCode") & ""
                If COLOR_CODE.StartsWith("#") Then
                    COLOR_CODE = Trim(Mid(COLOR_CODE, 2))
                End If

                Dim isColorGood As Boolean = False

                If COLOR_CODE <> "" Then
                    Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                    If rowICTCOLR1 IsNot Nothing Then
                        isColorGood = True
                    Else
                        If Not bad_COLOR_CODEs.Contains(COLOR_CODE) Then
                            bad_COLOR_CODEs.Add(COLOR_CODE)
                            MsgBox("Bad Color Code " & COLOR_CODE)
                        End If
                    End If
                End If

                If isColorGood Then
                    Dim STYLE_COLOR_DESC_po As String = rowpocolor.Item("ColorName") & ""
                    STYLE_COLOR_DESC_po = ASCMAIN1.Make_Caption(STYLE_COLOR_DESC_po)
                    Dim rowICTSTYC1 As DataRow = frmASFBASE1s("ICTSTYL1").dst.Tables("ICTSTYC1").Rows.Find(New String() {STYLE_CODE, COLOR_CODE})
                    If rowICTSTYC1 Is Nothing Then
                        rowICTSTYC1 = frmASFBASE1s("ICTSTYL1").dst.Tables("ICTSTYC1").NewRow
                        rowICTSTYC1.Item("STYLE_CODE") = STYLE_CODE
                        rowICTSTYC1.Item("COLOR_CODE") = COLOR_CODE
                        rowICTSTYC1.Item("STYLE_COLOR_STATUS") = "A"

                        frmASFBASE1s("ICTSTYL1").dst.Tables("ICTSTYC1").Rows.Add(rowICTSTYC1)
                    End If

                    Dim STYLE_COLOR_DESC As String = rowICTSTYC1.Item("STYLE_COLOR_DESC") & ""
                    If Not COLOR_CODEs.Contains(COLOR_CODE) Then
                        STYLE_COLOR_DESC = ""
                        COLOR_CODEs.Add(COLOR_CODE)
                        STYLE_COLOR_DESC = STYLE_COLOR_DESC_po
                    Else
                        STYLE_COLOR_DESC &= ", " & STYLE_COLOR_DESC_po
                    End If



                    If Len(STYLE_COLOR_DESC) > STYLE_COLOR_DESC_maxLength Then
                        MsgBox("Style/Color Description for color " & COLOR_CODE & ":" & vbCrLf & "  " & STYLE_COLOR_DESC _
                               & vbCrLf & " exceeds maximum length of " & CStr(STYLE_COLOR_DESC_maxLength) & " characters." _
                               & vbCrLf & vbCrLf & "Truncating the Description to " _
                               & vbCrLf & Mid(STYLE_COLOR_DESC, 1, STYLE_COLOR_DESC_maxLength))
                        STYLE_COLOR_DESC = Mid(STYLE_COLOR_DESC, 1, STYLE_COLOR_DESC_maxLength)
                    End If

                    rowICTSTYC1.Item("STYLE_COLOR_DESC") = STYLE_COLOR_DESC
                End If
            Next


            Application.DoEvents()

            frmASFBASE1s("ICTSTYL1").RemoteProcedureCall("Fill from AT", keysICTSTYL1)

        End If

        'MsgBox("Style " & STYLE_CODE & " does not exist in the Style Master File (yet)." _
        '       & vbCrLf & vbCrLf & "You must first create Style " & STYLE_CODE & " in the Style Master File." _
        '       & vbCrLf & vbCrLf & "Then, use the 'easy' button to load:" & vbCrLf _
        '       & vbCrLf & " - Material Content" _
        '       & vbCrLf & " - Factory" _
        '       & vbCrLf & " - Country of Origin" _
        '       & vbCrLf & " - Supplier (Vendor Code)" _
        '       & vbCrLf & " - Color Descriptions" _
        '       & "", _
        '       MsgBoxStyle.OkOnly, "Cannot Edit Style")

    End Sub
    Private Sub grdStyle_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdStyle.InitializeRow
        Color_Version(e, grdStyle)
    End Sub

    Private Sub grdStyleColor_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdStyleColor.InitializeRow
        Color_Version(e, grdStyleColor)
    End Sub

    Private Sub grdStyleSize_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdStyleSize.InitializeRow
        Color_Version(e, grdStyleSize)
    End Sub
    Sub Produce_XLS()

        ASCMAIN1.Progress("Now Producing XLS")

        Dim workbook As SpreadsheetGear.IWorkbook = Nothing

        Dim StyleRef As String = rowpohdr.Item("StyleRef") & ""

        Dim clsPOCORDRA As New POCORDRA
        With clsPOCORDRA
            .PO_PARM_PO_IMG_DIR = PO_PARM_PO_IMG_DIR


            If StyleRef = "" Then
                workbook = .Produce_XLS(Me, VAN_REF, workbook)
            Else

                ASCMAIN1.sql = "Select `PONo` PONO, MAX(VAN_REF) VAN_REF from AT.`pohdr` where `StyleRef` = '" & StyleRef & "' group by `PONo`"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PONO")

                    ASCMAIN1.Progress("-", row.Item("PONO"))

                    Dim VAN_REF As String = row.Item("VAN_REF")
                    workbook = .Produce_XLS(Me, VAN_REF, workbook)
                Next

            End If
             

            Dim XLS_FILENAME_base As String = "FOB Cost Sheet " & rowpohdr.Item("PONo") & " Xmit " & VAN_REF
            Dim XLS_FILENAME As String = XLS_FILENAME_base & ".xlsx"
            Dim retryCount As Integer = 0
            Do Until retryCount = -1 Or retryCount > 5
                If retryCount > 0 Then
                    XLS_FILENAME = XLS_FILENAME_base & "_" & CStr(retryCount) & ".xlsx"
                End If
                Try
                    workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
                    workbook.Close()
                    retryCount = -1
                Catch ex As Exception
                    retryCount += 1
                    If retryCount > 5 Then
                        MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Failed to Save Workbook")
                    End If
                End Try
            Loop

            If retryCount = -1 Then
                Show_Document(XLS_FILENAME)
            End If


        End With

        ASCMAIN1.Progress("")
    End Sub


    Function Sizes_and_Colors()

        Dim SCs As String = ""
        Dim SZs As String = ""
        Dim SQs As String = ""
        For I As Integer = 1 To 12
            Dim SIZE_CODE As String = Absx1.txtFor("SIZE_" & Format(I, "00")).Text
            If SIZE_CODE = "" Then
                Exit For
            End If
            SZs &= " " & SIZE_CODE
            Dim QTY As Integer = Val(Absx1.numFor("QTY_" & Format(I, "00")).Value & "")
            SQs &= "/" & CStr(QTY)
        Next

        For Each ROW As DataRow In dst.Tables("ICTSTYC1").Select("", "COLOR_CODE")
            Dim COLOR_CODE = ROW.Item("COLOR_CODE")
            Dim STYLE_COLOR_DESC = ROW.Item("STYLE_COLOR_DESC") & ""
            SCs &= vbCrLf & COLOR_CODE & " " & STYLE_COLOR_DESC
        Next

        If SZs <> "" Then
            SCs = Mid(SZs, 2) & " = " & Mid(SQs, 2) & vbCrLf & SCs
        Else
            SCs = Mid(SCs, 3)
        End If

        Return SCs

    End Function
 
    Private Sub grdStyleSize_MouseDown(sender As Object, e As MouseEventArgs) Handles grdStyleSize.MouseDown
        If btnCreateStyle.Visible Then Exit Sub
        Try
            Dim grid As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
            Dim element As UIElement = grid.DisplayLayout.UIElement.LastElementEntered

            ' See if the element is a HeaderUIElement. This will probably never happen
            ' because the header element is filled with a TexUIElement, but it's best to 
            ' cover all the bases. 
            Dim headerElement As UltraWinGrid.HeaderUIElement = DirectCast(element.Parent, UltraWinGrid.HeaderUIElement)

            If headerElement Is Nothing Then
                ' See if the element has a HeaderUIElement in it's parent chain.
                '    headerElement = element.GetAncestor(typeof( UltraWinGrid.HeaderUIElement)) as UltraWinGrid. HeaderUIElement;
            End If


            ' We failed to find a HeaderUIElement, so we must not be on a header. 
            If headerElement Is Nothing Then
                Return
            Else
                ' A HeaderUIElement could be the element for the grid caption, a group, or a column.
                ' Check if this head has a Group. 
                Dim groupHeader As UltraWinGrid.GroupHeader = DirectCast(headerElement.Header, UltraWinGrid.GroupHeader)
                If groupHeader Is Nothing Then
                    Return
                Else
                    Dim group As UltraWinGrid.UltraGridGroup = groupHeader.Group

                    Dim STYLE_CODE As String = group.Header.ToolTipText
                    Dim SIZE_CODE As String = group.Header.Tag
                    Easy_Button(STYLE_CODE, True, SIZE_CODE)
                    Debug.WriteLine(group.Key + " clicked for " & group.Header.ToolTipText)
                End If
            End If
        Catch ex As Exception

        End Try
      
    End Sub

    Function Get_Style_Codes(STYLE_CODE As String, POColorKey As Int64) As List(Of String)
        Dim sqlV As String = "VAN_REF = '" & VAN_REF & "'"
        If POColorKey <> -1 Then
            sqlV &= " and POColorKey = " & CStr(POColorKey)
        End If

        Dim STYLE_CODEs As New List(Of String)
        ' Dim NoofStyle As Integer = Val(rowpohdr.Item("NoofStyle") & "")
        Dim StyleBySize As Boolean = (rowpohdr.Item("StyleBySize") & "" = "Y")


        ' WJZ 09/30/21 FOR ME POS, AT IS SENDING STYLEBYSIZE = N - NEED TO TAKE THIS UP WITH EDMUND
        ' If PONO.StartsWith("ME") Then StyleBySize = False ' True

        sqlV &= " and Style is not null"
        If StyleBySize Then
            sqlV &= " and Size is not null"
        End If

        If StyleBySize Then ' If NoofStyle > 1 Then
            For Each row As DataRow In dst.Tables("posizedtl").Select(sqlV)
                ' wjz 09/30/2021 -  making the change below because sometimes AT is placing the full style code in the Size column and not the Style column
                Dim col As String = "Style"
                ' If PONO.StartsWith("ME") Then col = "Size"



                If POColorKey = -1 And STYLE_CODEs.Contains(row.Item(col)) Then
                Else
                    STYLE_CODEs.Add(row.Item(col))
                End If
            Next
        Else
            STYLE_CODEs.Add(STYLE_CODE)
        End If
        Return STYLE_CODEs
    End Function

    Private Function STYLE_CODE_to_check() As String
        Throw New NotImplementedException
    End Function

End Class