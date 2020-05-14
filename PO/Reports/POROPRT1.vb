Public Class POROPRT1

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date

    Dim SQLs As New Dictionary(Of String, String)

    Dim POTORDR1 As String
    Dim POTORDR2 As String

    Dim sqlPOTORDR1 As String
    Dim sqlPOTORDR2 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")
        Get_PARM("ICTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 1

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            Absx1.optFor("RANGE").CheckedIndex = 2
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "POs Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "POs Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and POTORDR1.PO_DATE_ORDERED between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "S" Then
            SUBT = "Selected POs"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "U" Then
            SUBT = "All POs not Printed Yet"
            sqlw &= " and POTORDR1.PO_PRINTED_IND = '0'" & vbCrLf

        End If

        sqlw &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
        sqlw &= SQL_in("PO_ORDER_NO", "POTORDR1.PO_ORDER_NO")

        Prepare_dst(True, sqlw)

        Check_if_Empty("POTORDR1")
    End Sub

    Public Overrides Sub Print_Report()
        Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
        If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

        CR_params.Add("SUBT", "")
        CR_params.Add("FORM_TYPE", "P")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            Else

            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")

        If optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()
        Dim sql As String = ""
        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
        Else
            sql = "Update POTORDR1 " _
                & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" _
                & " where (PO_ORDER_NO) in (Select PO_ORDER_NO from " & POTORDR1 & " )"
            ASCDATA1.ExecuteSQL(sql)
        End If

        sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
            & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO_PRT','PO Printed', PO_ORDER_NO" _
            & " from " & POTORDR1
        ASCDATA1.ExecuteSQL(sql)
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("POTPARM1")
        Get_PARM("ICTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        sqlPOTORDR1 = "Select POTORDR1.* from POTORDR1 "
        ASCMAIN1.sql = sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw)
        POTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add Primary Key (PO_ORDER_NO)")

        sqlPOTORDR2 = "Select POTORDR2.* from POTORDR2, " & POTORDR1 _
            & " POTORDR1 where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
        ASCMAIN1.sql = sqlPOTORDR2
        POTORDR2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR2 & " Add Primary Key (PO_ORDER_NO, PO_ORDER_LNO)")

        SQLs.Clear()

        With dst
            ASCMAIN1.sql = "Select POTORDR1.*, 'Z' PO_PARM_KEY" & vbCrLf _
                & " from " & POTORDR1 & " POTORDR1"
            SQLs.Add("POTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & " from " & POTORDR2 & " POTORDR2"
            SQLs.Add("POTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "", 2)
            .Tables("POTORDR2").Columns.Add("DUPLICATE_IMAGE")

            Create_Relation("POTORDR1", "POTORDR2", "PO_ORDER_NO")

            ASCMAIN1.sql = "Select POTORDR6.*" & vbCrLf _
                & " from " & POTORDR1 & " POTORDR1,POTORDR6" _
                & " where POTORDR6.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
            SQLs.Add("POTORDR6", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR6", "**", 0, False, "", 3)
            With .Tables("POTORDR6")
                .Columns.Add("PO_MESSAGE_IMG", GetType(System.Byte()))
            End With

            Create_Relation("POTORDR2", "POTORDR6", "PO_ORDER_NO,PO_ORDER_LNO")
            .Tables("POTORDR2").Columns.Add("MESSAGE_COUNT", GetType(System.Int32), "COUNT(CHILD(POTORDR2_POTORDR6).PO_ORDER_MLNO)")

            ASCMAIN1.sql = "Select POTORDRZ.*" & vbCrLf _
                & " from POTORDRZ" & vbCrLf _
                & " where PO_ORDER_NO in (Select PO_ORDER_NO from " & POTORDR1 & ")"
            SQLs.Add("POTORDRZ", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDRZ", "**", 0, False, "", 3)
            With .Tables("POTORDRZ").Columns
                .Add("STYLE_CODE_PREV")
                .Add("COLOR_CODE_PREV")
                .Add("PO_QTY_ORD_PREV", GetType(System.Int64))
                .Add("PO_COST_PREV", GetType(System.Decimal))
                .Add("PO_DATE_SHIP_BY_PREV", GetType(System.DateTime))
                .Add("PO_STATUS_PREV")
                .Add("CARTON_PACK_QTY_PREV", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.STYLE_COST, ICTSTYL1.STYLE_MATL_DESC, ICTSTYL1.COUNTRY_CODE" & vbCrLf _
                & ", ICTSTYL1.CASE_CUBE, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.LABEL_TYPE_CODE" & vbCrLf _
                & ", ICTSTYL1.DUTY_RATE_CODE, ICTDUTY1.DUTY_HTS_CODE, ICTSTYL1.IMAGE_NAME" & vbCrLf _
                & ", ICTSTYL1.SIZE_SCALE, ICTSTYL1.STYLE_CODE_PLM, ICTPLIN2.DESIGN_STYLE_NO" & vbCrLf _
                & ", ICTSTYL1.PURCH_NOTES, ICTSTYL1.CUST_STYLE_CODE" & vbCrLf _
                & " from ICTSTYL1,ICTDUTY1,ICTPLIN2" & vbCrLf _
                & " where ICTSTYL1.STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & ")" & vbCrLf _
                & "   and ICTDUTY1.DUTY_RATE_CODE (+) = ICTSTYL1.DUTY_RATE_CODE" & vbCrLf _
                & "   and ICTPLIN2.STYLE_CODE_PLM (+) = ICTSTYL1.STYLE_CODE_PLM"
            SQLs.Add("ICTSTYL1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False, "", 1)
            .Tables("ICTSTYL1").Columns.Add("IMAGE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select ICTSTYL4.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" & vbCrLf _
                & " from ICTSTYL4,ICTSTYL1,ICTCOLR1" & vbCrLf _
                & " where (ICTSTYL4.STYLE_CODE,ICTSTYL4.COLOR_CODE) in (Select Distinct STYLE_CODE,COLOR_CODE from " & POTORDR2 & ")" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = ICTSTYL4.STYLE_CODE_COMP" & vbCrLf _
                & "   and ICTCOLR1.COLOR_CODE = ICTSTYL4.COLOR_CODE_COMP"
            SQLs.Add("ICTSTYL4", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYL4", "**", 0, False)

            ASCMAIN1.sql = "Select ICTSTYL5.*, ICTPANTC.PANTONE_DESC, ICTPANTC.RGB" & vbCrLf _
                & " from ICTSTYL5,ICTPANTC" & vbCrLf _
                & " where (ICTSTYL5.STYLE_CODE,ICTSTYL5.COLOR_CODE) in (Select Distinct STYLE_CODE,COLOR_CODE from " & POTORDR2 & ")" & vbCrLf _
                & "   and ICTPANTC.PANTONE_CODE = ICTSTYL5.PANTONE_CODE"
            SQLs.Add("ICTSTYL5", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYL5", "**", 0, False)

            ASCMAIN1.sql = "Select ICTSTYV1.*" & vbCrLf _
                & " from ICTSTYV1" & vbCrLf _
                & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & ")"
            SQLs.Add("ICTSTYV1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYV1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ICTSTYC1.*" & vbCrLf _
                & " from ICTSTYC1" & vbCrLf _
                & " where STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & ")"
            SQLs.Add("ICTSTYC1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ICTCOLR1.*" & vbCrLf _
                & " from ICTCOLR1" & vbCrLf _
                & " where COLOR_CODE in (Select Distinct COLOR_CODE from " & POTORDR2 & ")"
            SQLs.Add("ICTCOLR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select APTVEND1.*" & vbCrLf _
                & " from APTVEND1" & vbCrLf _
                & " where VEND_CODE in (Select Distinct VEND_CODE from " & POTORDR1 & ")"
            SQLs.Add("APTVEND1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTPLIN4.*" & vbCrLf _
                & " from ICTPLIN4" & vbCrLf _
                & " where STYLE_CODE_PLM in (Select Distinct STYLE_CODE_PLM from ICTSTYL1 where STYLE_CODE in (Select Distinct STYLE_CODE from " & POTORDR2 & "))"
            SQLs.Add("ICTPLIN4", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTPLIN4", "**", 0, False, "", 2)

            For Each TABLE_NAME As String In New String() _
            {"TATTERM1", "ICTWHSE1", "ICTPORT1", "TATCNTRY", "ICTCOSTB", "POTMESST"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            ASCMAIN1.sql = "Select POTPARM1.*" & vbCrLf _
                & " from POTPARM1 where PO_PARM_KEY = 'Z'" 
            Create_TDA(.Tables.Add, "POTPARM1", "**", 0, False, "", 1)
            .Tables("POTPARM1").Columns.Add("LOGO", GetType(System.Byte()))


            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & POTORDR1 & ")"
            SQLs.Add("SOTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & POTORDR1 & ")"
            SQLs.Add("SOTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*" & vbCrLf _
                & " from SOTORDR5" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & POTORDR1 & ")" & vbCrLf _
                & "   and CUST_ADDR_TYPE = 'ST'"
            SQLs.Add("SOTORDR5", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR5", "**", 0, False, "", 1) ' PURPOSEFULLY CALLING THIS A 1 PART KEY FOR 1:1 JOIN WITH POTORDR1
        End With

        Try
            Fill_Records("POTPARM1")
            Dim rowPOTPARM1 As DataRow = dst.Tables("POTPARM1").Rows(0)
            rowPOTPARM1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")

            If ASCMAIN1.CLIENT = "NYA" Then
                dst.Tables("POTPARM1").Columns.Add("PO_PARM_FORM_COUNTRY")
                dst.Tables("POTPARM1").Columns.Add("PO_PARM_FORM_TAX_ID")
                dst.Tables("POTPARM1").Columns("PO_PARM_KEY").MaxLength = -1
                dst.Tables("POTORDR1").Columns("PO_PARM_KEY").MaxLength = -1

                ASCMAIN1.sql = "Select * from SOTCOMP1"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim COMP_CODE As String = row.Item("COMP_CODE")
                    Dim rowP As DataRow = dst.Tables("POTPARM1").NewRow
                    rowP.Item("PO_PARM_KEY") = "Z" & COMP_CODE
                    For Each C As String In New String() {"COMP_NAME", "COMP_ADDR1", "COMP_ADDR2", "COMP_ADDR3", _
                                                          "COMP_CITY", "COMP_STATE", "COMP_ZIP_CODE", "COMP_COUNTRY", _
                                                          "COMP_PHONE", "COMP_FAX", "COMP_EMAIL", "COMP_TAX_ID"}
                        Dim CP As String = Replace(C, "COMP_", "PO_PARM_FORM_")
                        rowP.Item(CP) = row.Item(C)
                    Next
                    rowP.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & "_" & COMP_CODE & ".PNG")
                    dst.Tables("POTPARM1").Rows.Add(rowP)
                Next
            End If

        Catch ex As Exception
            Stop
        End Try
 
        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        sqlw = parms(0)

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR1)
        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR2)

        ASCDATA1.ExecuteSQL("Insert into " & POTORDR1 & " " & sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw))
        ASCDATA1.ExecuteSQL("Insert into " & POTORDR2 & " " & sqlPOTORDR2)

        If RWU = "R" Then
            Dim POS_NOT_LOCKABLE As New List(Of String)
            ASCMAIN1.sql = "Select PO_ORDER_NO from " & POTORDR1
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    POS_NOT_LOCKABLE.Add(PO_ORDER_NO)
                End If
            Next
            If POS_NOT_LOCKABLE.Count <> 0 Then
                ASCMAIN1.sql = "Delete from " & POTORDR1 & " where PO_ORDER_NO in (" & Join(POS_NOT_LOCKABLE.ToArray, "','") & ")"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from " & POTORDR2 & " where PO_ORDER_NO in (" & Join(POS_NOT_LOCKABLE.ToArray, "','") & ")"
                ASCDATA1.ExecuteSQL()
            End If
        End If

        EnforceConstraints(False)
        Fill_Records("POTORDR1")

        dst.Tables("POTORDRZ").Rows.Clear()
        TAC.POCMAIN1.Setup_PO_Change_Details(Me)

        Fill_Records("POTORDR2")
        Fill_Records("POTORDR6")
        Fill_Records("ICTSTYL1")

        If ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA" Then
            Fill_Records("ICTPLIN4")

            Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_STYLE_IMG_DIR") & ""
            For Each rowICTSTYL1 As DataRow In dst.Tables("ICTSTYL1").Select("")
                Dim imgba() As Byte = Nothing
                '  TAC.ICCMAIN1.Get_Image(Me, rowICTSTYL1, imgba)
                ASCMAIN1.Get_Image(FOLDER_NAME, rowICTSTYL1.Item("IMAGE_NAME") & "", True, , , imgba) ' imgba)
                rowICTSTYL1.Item("IMAGE") = imgba

                Dim STYLE_CODE_PLM As String = rowICTSTYL1.Item("STYLE_CODE_PLM") & ""
                If dst.Tables("ICTPLIN4").Select("STYLE_CODE_PLM = '" & STYLE_CODE_PLM & "'").Length = 0 Then
                    rowICTSTYL1.Item("STYLE_CODE_PLM") = ""
                End If
            Next
        End If

        Dim PO_ORDER_NO_images As String = ""
        Dim IMAGE_NAME As String = ""
        Dim IMAGE_NAMEs As New List(Of String)
        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("", "PO_ORDER_NO, PO_ORDER_LNO")
            If rowPOTORDR2.Item("PO_ORDER_NO") & "" <> PO_ORDER_NO_images Then
                PO_ORDER_NO_images = rowPOTORDR2.Item("PO_ORDER_NO") & ""
                IMAGE_NAMEs.Clear()
            End If
            Dim rowICTSTYL1 As DataRow = dst.Tables("ICTSTYL1").Rows.Find(rowPOTORDR2.Item("STYLE_CODE"))
            IMAGE_NAME = rowICTSTYL1.Item("IMAGE_NAME") & ""
            If IMAGE_NAMEs.Contains(IMAGE_NAME) Then
                rowPOTORDR2.Item("DUPLICATE_IMAGE") = "1"
            Else
                IMAGE_NAMEs.Add(IMAGE_NAME)
            End If
        Next

        For Each rowPOTORDR6 As DataRow In dst.Tables("POTORDR6").Select("ISNULL(PO_MESSAGE_ATTACHMENT,'') <> ''")
            Dim ATTACHMENT_NO As String = rowPOTORDR6.Item("PO_MESSAGE_ATTACHMENT")
            Dim PO_ORDER_NO As String = rowPOTORDR6.Item("PO_ORDER_NO")
            Dim rowASTATTA2 As DataRow = LookUp("ASTATTA2", New String() {"POTORDR6", "PO_MESSAGE_ATTACHMENT", PO_ORDER_NO, ATTACHMENT_NO})
            If rowASTATTA2 IsNot Nothing Then
                Dim ATTACHMENT_EXT As String = rowASTATTA2.Item("ATTACHMENT_EXT")
                Dim appFileName = ASCMAIN1.Folders("Work") & ATTACHMENT_NO & "." & ATTACHMENT_EXT
                Try
                    My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Attach") & ATTACHMENT_NO, appFileName, True)
                    appFileName = My.Computer.FileSystem.GetFileInfo(appFileName).FullName
                    Dim image_array() As Byte = Nothing
                    Dim x As System.Drawing.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Work"), ATTACHMENT_NO & "." & ATTACHMENT_EXT, True, , , image_array)
                    rowPOTORDR6.Item("PO_MESSAGE_IMG") = image_array
                Catch ex As Exception
                    'MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to Launch Attachment")
                End Try
            End If
        Next


        If ASCMAIN1.CLIENT = "NYA" Then
            For Each row1 As DataRow In dst.Tables("POTORDR1").Select("")
                Dim INIT_DATE As String = Format(row1.Item("INIT_DATE"), "yyyyMMdd")
                'Walter, Below suggestion is OK “All POs entered prior to 05/03 need to show the old logo”. - Pam 07/20

                If INIT_DATE >= "20180503" Then
                    Dim P As String = row1.Item("PO_ORDER_NO")
                    For Each row2 As DataRow In dst.Tables("POTORDR2").Select("PO_ORDER_NO = '" & P & "'")
                        Dim STYLE_CODE As String = row2.Item("STYLE_CODE")
                        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                        Dim SALES_DIVISION_CODE As String = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                        Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE)
                        If rowSOTSDIV1 IsNot Nothing AndAlso rowSOTSDIV1.Item("SEG4_CODE") & "" <> "" Then
                            row1.Item("PO_PARM_KEY") = "Z" & rowSOTSDIV1.Item("SEG4_CODE")
                        End If
                        Exit For
                    Next
                End If
            Next
        End If

        Fill_Records("ICTSTYL4")
        Fill_Records("ICTSTYL5")
        Fill_Records("ICTSTYC1")
        Fill_Records("ICTSTYV1")
        Fill_Records("ICTCOLR1")
        Fill_Records("APTVEND1")
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")
        Fill_Records("SOTORDR5")

        EnforceConstraints(True)
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sqlw = "APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from POTORDR1)"
        End Select
        Return sqlw
    End Function
End Class