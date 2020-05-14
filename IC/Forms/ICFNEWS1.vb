Public Class ICFNEWS1
    Dim ICTSTYLX As String 'TABLE_NAME
    Dim ICTSTYLD As String
    Dim sqlICTSTYLX As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
                & ", STYLE_CLASS_CODE, ICTSTYL1.VEND_CODE, ICTSTYV1.VEND_ITEM_CODE, ICTSTYL1.STYLE_PRICE " & vbCrLf _
                & ", ICTSTYC1.UPC_CODE, ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE, decode (NVL(ICTSTYV1.NEW_PO_COST,0),0,NVL(ICTSTYV1.PO_COST,0),NVL(ICTSTYV1.NEW_PO_COST,0)) PO_COST,0 NEW_PO_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_PO_QTY_MIN, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_MATL_DESC " & vbCrLf _
                & ", T1.LENGTH AS STYLE_LENGTH, T1.WIDTH AS STYLE_WIDTH, T1.HEIGHT AS STYLE_HEIGHT, T1.WEIGHT AS STYLE_WEIGHT" & vbCrLf _
                & ", T2.LENGTH AS INNER_LENGTH, T2.WIDTH AS INNER_WIDTH, T2.HEIGHT AS INNER_HEIGHT, T2.WEIGHT AS INNER_WEIGHT" & vbCrLf _
                & ", T3.LENGTH AS CARTON_LENGTH, T3.WIDTH AS CARTON_WIDTH, T3.HEIGHT AS CARTON_HEIGHT, T2.WEIGHT AS CARTON_WEIGHT" & vbCrLf _
                & ",  SUM( CASE WHEN ICTSTAT2.WHSE_CODE = 'MS' THEN NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) +  NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK ,0) ELSE 0 END) AS FUT_AVAIL_MS " & vbCrLf _
                & ",  SUM( CASE WHEN ICTSTAT2.WHSE_CODE <> 'MS' THEN NVL(ICTSTAT2.WHSE_QTY_ON_HAND,0) + NVL(ICTSTAT2.WHSE_QTY_TRAN,0) +  NVL(ICTSTAT2.WHSE_QTY_ON_ORDER,0) - NVL(ICTSTAT2.WHSE_QTY_OPEN,0) - NVL(ICTSTAT2.WHSE_QTY_PICK ,0) ELSE 0 END) AS FUT_AVAIL_OTHER " & vbCrLf _
                & ", 0 as CART_LEN_SS, 0 as CART_WDTH_SS, 0 as CART_HT_SS,  0 AS CART_WGHT_SS " & vbCrLf _
                & ", 0 as INNER_LEN_SS, 0 as INNER_WDTH_SS, 0 as INNER_HT_SS,  0 AS INNER_WGHT_SS " & vbCrLf _
                & ", 0 as ITM_LEN_SS, 0 as ITM_WDTH_SS, 0 as ITM_HT_SS,  0 AS ITM_WGHT_SS " & vbCrLf _
                & " from ICTSTYL1,ICTSTYC1,ICTSTYV1, ICTSTAT2, ICTSTYLD T1, ICTSTYLD T2, ICTSTYLD T3 " & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE = ICTSTYL1.STYLE_CODE " & vbCrLf _
                & "   and ICTSTYV1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE " & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE " & vbCrLf _
                & "   and ICTSTAT2.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE " & vbCrLf _
                & "   and ICTSTAT2.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE " & vbCrLf _
                & "   and T1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T1.PACK_CODE (+) = 'IT' " & vbCrLf _
                & "   and T2.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T2.PACK_CODE (+) = 'INR' " & vbCrLf _
                & "   and T3.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T3.PACK_CODE (+) = 'INR' " & vbCrLf _
                & " GROUP BY ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS, ICTSTYC1.COLOR_CODE, ICTSTYC1.STYLE_COLOR_STATUS" & vbCrLf _
                & ", STYLE_CLASS_CODE, ICTSTYL1.VEND_CODE, ICTSTYV1.VEND_ITEM_CODE, ICTSTYL1.STYLE_PRICE " & vbCrLf _
                & ", ICTSTYC1.UPC_CODE, ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE,  PO_COST, NEW_PO_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_PO_QTY_MIN, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_MATL_DESC " & vbCrLf _
                & ", T1.LENGTH, T1.WIDTH , T1.HEIGHT , T1.WEIGHT  " & vbCrLf _
                & ", T2.LENGTH  , T2.WIDTH  , T2.HEIGHT , T2.WEIGHT  " & vbCrLf _
                & ", T3.LENGTH  , T3.WIDTH  , T3.HEIGHT , T3.WEIGHT  " & vbCrLf

            sqlICTSTYLX = ASCMAIN1.sql

            ICTSTYLX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add NEW_PRICE VARCHAR2(20)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add FGT  VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add DUTY  VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PFRT VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PCT  VARCHAR2(10)")

            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add CART_LEN_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add CART_WDTH_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add CART_HT_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add CART_WGHT_SS  VARCHAR2(10)")

            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add INNER_LEN_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add INNER_WDTH_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add INNER_HT_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add INNER_WGHT_SS  VARCHAR2(10)")

            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add ITM_LEN_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add ITM_WDTH_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add ITM_HT_SS  VARCHAR2(10)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add ITM_WGHT_SS  VARCHAR2(10)")


            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add NEW_PRICE NUMBER(8,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add FGT NUMBER(8,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add DUTY NUMBER(8,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PFRT NUMBER(8,2)")
            'ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PCT NUMBER(8,2)")

            ASCMAIN1.sql = "Select * from " & ICTSTYLX
            Create_TDA(.Tables.Add("ICTSTYLX"), ICTSTYLX, "**", 0, True)
            '.Tables("ICTSTYLX").Columns.Add("NEW_PRICE")
            '.Tables("ICTSTYLX").Columns.Add("FGT")
            '.Tables("ICTSTYLX").Columns.Add("DUTY")
            '.Tables("ICTSTYLX").Columns.Add("PFRT")
            '.Tables("ICTSTYLX").Columns.Add("PCT")


            With .Tables.Add("TATCOLS1")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_CAPTION")
                .Columns.Add("COLUMN_NAME_ORACLE")
                .Columns.Add("TABLE_NAME")
                .Columns.Add("SEL")
                .Columns.Add("GRID_POS")
                '.Columns.Add("IMAGE", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With

            Create_TDA(.Tables.Add, "ICTSTYLD", "*")


        End With


        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")

        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("COLOR_CODE").Header.Fixed = True
            .Columns("STYLE_STATUS").Header.Fixed = True
            .Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            ' .Columns("COLOR_DESC").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                If gcol.Key = "NEW_PO_COST" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If

                If gcol.Key = "CART_LEN_SS" Or gcol.Key = "CART_WDTH_SS" Or gcol.Key = "CART_HT_SS" Or gcol.Key = "CART_WGHT_SS" Then
                    gcol.Hidden = True
                End If
                If gcol.Key = "INNER_LEN_SS" Or gcol.Key = "INNER_WDTH_SS" Or gcol.Key = "INNER_HT_SS" Or gcol.Key = "INNER_WGHT_SS" Then
                    gcol.Hidden = True
                End If

                If gcol.Key = "ITM_LEN_SS" Or gcol.Key = "ITM_WDTH_SS" Or gcol.Key = "ITM_HT_SS" Or gcol.Key = "ITM_WGHT_SS" Then
                    gcol.Hidden = True
                End If



            Next
        End With

        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_STATUS")
        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_COLOR_STATUS")
        spl.Panel1Collapsed = True

        Dim CC As Integer = dst.Tables("ICTSTYLX").Columns.Count
        For i As Integer = 1 To CC - 1
            Dim DC As DataColumn = dst.Tables("ICTSTYLX").Columns(i)
            '      
            '         

            '      

            '      
            If DC.ColumnName = "STYLE_CODE" Or DC.ColumnName = "STYLE_DESC" Or DC.ColumnName = "UPC_CODE" Or DC.ColumnName = "COLOR_CODE" Or DC.ColumnName = "PCT" Or DC.ColumnName = "PFRT" Or DC.ColumnName = "FGT" Or DC.ColumnName = "STYLE_STATUS" Or DC.ColumnName = "STYLE_COLOR_STATUS" Or DC.ColumnName = "PO_COST" Or DC.ColumnName = "DUTY" Or DC.ColumnName = "VEND_CODE" Or DC.ColumnName = "STYLE_PRICE" Or DC.ColumnName = "STYLE_CLASS_CODE" Or DC.ColumnName = "FUT_AVAIL_MS" Or DC.ColumnName = "FUT_AVAIL_OTHER" Or DC.ColumnName = "CART_LEN_SS" Or DC.ColumnName = "CART_WDTH_SS" Or DC.ColumnName = "CART_HT_SS" Or DC.ColumnName = "CART_WGHT_SS" Or DC.ColumnName = "INNER_LEN_SS" Or DC.ColumnName = "INNER_WDTH_SS" Or DC.ColumnName = "INNER_HT_SS" Or DC.ColumnName = "INNER_WGHT_SS" Or DC.ColumnName = "ITM_LEN_SS" Or DC.ColumnName = "ITM_WDTH_SS" Or DC.ColumnName = "ITM_HT_SS" Or DC.ColumnName = "ITM_WGHT_SS" Then

            Else
                Dim rowTATCOLS1 As DataRow = dst.Tables("TATCOLS1").NewRow
                Dim GC As UltraWinGrid.UltraGridColumn = grdICTSTYLX.DisplayLayout.Bands(0).Columns(DC.ColumnName)
                Dim DANA As String = DC.ColumnName
                If DC.ColumnName = "NEW_PRICE" Then
                    rowTATCOLS1.Item("COLUMN_NAME_ORACLE") = "STYLE_PRICE"
                ElseIf DC.ColumnName = "NEW_PO_COST" Then
                    rowTATCOLS1.Item("COLUMN_NAME_ORACLE") = "NEW_PO_COST"
                ElseIf DC.ColumnName = "NEW_PO_COST_DATE" Then
                    rowTATCOLS1.Item("COLUMN_NAME_ORACLE") = "PO_COST_DATE"
                Else
                    rowTATCOLS1.Item("COLUMN_NAME_ORACLE") = DC
                End If
                rowTATCOLS1.Item("COLUMN_CAPTION") = GC.Header.Caption
                rowTATCOLS1.Item("COLUMN_NAME") = DC
                rowTATCOLS1.Item("SEL") = "1"
                If DC.ColumnName = "NEW_PO_COST" Or DC.ColumnName = "NEW_PO_COST_DATE" Or DC.ColumnName = "VEND_ITEM_CODE" Then
                    rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYV1"
                Else
                    If DC.ColumnName = "STYLE_LENGTH" Or DC.ColumnName = "STYLE_WIDTH" Or DC.ColumnName = "STYLE_HEIGHT" Or DC.ColumnName = "STYLE_WEIGHT" Or DC.ColumnName = "INNER_LENGTH" Or DC.ColumnName = "INNER_WIDTH" Or DC.ColumnName = "INNER_HEIGHT" Or DC.ColumnName = "INNER_WEIGHT" Or DC.ColumnName = "CARTON_LENGTH" Or DC.ColumnName = "CARTON_WIDTH" Or DC.ColumnName = "CARTON_HEIGHT" Or DC.ColumnName = "CARTON_WEIGHT" Then
                        rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYLD"
                    Else
                        rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYL1"
                    End If
                End If


                dst.Tables("TATCOLS1").Rows.Add(rowTATCOLS1)
            End If
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load from Spreadsheet"
                Import_from_Excel()
                CHECK_DIMENSIONS()
                If dst.Tables("ICTSTYLX").Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Loaded"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load from Database"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Load from Spreadsheet"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Load from Hold"
                EntryMode = "H"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                Mode_Settings(True)

            Case "Cancel"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load from Database").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Load from Spreadsheet").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = (ScreenMode And EntryMode = "E")
                    .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTSTYLX", "ICTSTYLD"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        ' Absx1.txtFor("CUST_CODE").Text = ""
        ' Absx1.txtFor("SREP_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "E" Then
        Else

            If EntryMode = "H" Then
                ASCMAIN1.sql = "TRUNCATE TABLE " & ICTSTYLX
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "INSERT INTO " & ICTSTYLX & " SELECT *  FROM ICTSTYLX "
                ASCDATA1.ExecuteSQL()

                Fill_Records("ICTSTYLX")

            Else

                ASCMAIN1.sql = "TRUNCATE TABLE " & ICTSTYLX
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "INSERT INTO " & ICTSTYLX & " SELECT X.*,'','','','','' FROM (" & sqlICTSTYLX & ") X "
                ASCDATA1.ExecuteSQL()

                'DANAC= INSERT X.*,'','','','','' FROM (   X 


                Fill_Records("ICTSTYLX")
            End If
        End If

        Sort_grdColumns(grdICTSTYLX, "STYLE_CODE,COLOR_CODE")

        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            rowTATCOLS1.Item("SEL") = "1"
        Next

        Set_background_colors()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        ASCMAIN1.sql = "TRUNCATE TABLE " & ICTSTYLX
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("ICTSTYLX")

        BeginTrans()

        ' need to add cost date to sql

        Dim COST_EFF_DATE As String = Format(Absx1.dteFor("COST_DATE").Value, "dd-MMM-yyyy")
        'Dim COLS(20) As String
        'COLS(0) = "1"
        'COLS(1) = "1"
        'COLS(2) = "1"
        'COLS(3) = "1"
        'COLS(4) = "1"
        'COLS(5) = "1"
        'COLS(6) = "1"
        'COLS(7) = "1"
        'COLS(8) = "1"
        'COLS(9) = "1"
        'COLS(10) = "1"
        'COLS(11) = "1"
        'COLS(12) = "1"

        Dim SQLV As String = ""
        Dim SQLL As String = ""

        Dim SQL1 As String = ""
        Dim SQLV1 As String = "" ' NEED TO BUILD SUPPORT IF COST UPDATE IS SELECTED

        Dim SQLDC As String = ""
        Dim SQLDI As String = ""
        Dim SQLDS As String = ""


        'If COLS(0) = "1" Then
        '    SQLV = "PO_COST = NVL(NEW_PO_COST,0) "
        'End If

        'If COLS(1) = "1" Then
        '    If SQLV <> "" Then
        '        SQLV = SQLV & ","
        '    End If
        '    SQLV = SQLV & " PO_COST_DATE = NEW_PO_COST_DATE "
        'End If


        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            If rowTATCOLS1.Item("SEL") = "1" Then
                If rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYV1" Then
                    If SQLV = "" Then
                        SQLV = SQLV & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    Else
                        SQLV = SQLV & "," & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    End If
                Else
                    If rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYL1" Then
                        If SQLL = "" Then
                            SQLL = SQLL & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                        Else
                            SQLL = SQLL & "," & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                        End If
                    End If
                End If
            End If
        Next


        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            If rowTATCOLS1.Item("SEL") = "1" Then
                If rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYLD" And Mid(rowTATCOLS1.Item("COLUMN_NAME_ORACLE"), 1, 6) = "CARTON" Then
                    If SQLDC = "" Then
                        SQLDC = SQLDC & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    Else
                        SQLDC = SQLDC & "," & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    End If
                End If
            End If

        Next


        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            If rowTATCOLS1.Item("SEL") = "1" Then
                If rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYLD" And Mid(rowTATCOLS1.Item("COLUMN_NAME_ORACLE"), 1, 5) = "INNER" Then
                    If SQLDI = "" Then
                        SQLDI = SQLDI & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    Else
                        SQLDI = SQLDI & "," & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    End If
                End If
            End If

        Next



        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            If rowTATCOLS1.Item("SEL") = "1" Then
                If rowTATCOLS1.Item("TABLE_NAME") = "ICTSTYLD" And Mid(rowTATCOLS1.Item("COLUMN_NAME_ORACLE"), 1, 5) = "STYLE" Then
                    If SQLDS = "" Then
                        SQLDS = SQLDS & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    Else
                        SQLDS = SQLDS & "," & rowTATCOLS1.Item("COLUMN_NAME_ORACLE") & "= R1." & rowTATCOLS1.Item("COLUMN_NAME")
                    End If
                End If
            End If

        Next



        ' NEED TO ADD IN COMMA IF MORE THAN ONE COLUMN
        If SQLV <> "" And SQLV.Contains("NEW_PO_COST") And chkBYPASS_COSTS.Checked = False Then
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " _
                & " Select STYLE_CODE, VEND_CODE, NEW_PO_COST from  " & ICTSTYLX & " GROUP BY STYLE_CODE, VEND_CODE, NEW_PO_COST;" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & " Update ICTSTYV1 Set " & vbCrLf _
                & " PRV_PO_COST = PO_COST, PRV_PO_COST_DATE = PO_COST_DATE " & vbCrLf _
                & " where STYLE_CODE = R1.STYLE_CODE And VEND_CODE = R1.VEND_CODE And NVL( R1.NEW_PO_COST,0) <> 0 AND NVL(NEW_PO_COST,0) <> 0  ;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is " _
                & " Select STYLE_CODE, VEND_CODE, NEW_PO_COST from  " & ICTSTYLX & " GROUP BY STYLE_CODE, VEND_CODE, NEW_PO_COST;" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & " Update ICTSTYV1 Set " & vbCrLf _
                & " PO_COST = NEW_PO_COST, PO_COST_DATE = NEW_PO_COST_DATE " & vbCrLf _
                & " where STYLE_CODE = R1.STYLE_CODE And VEND_CODE = R1.VEND_CODE And NVL( R1.NEW_PO_COST,0) <> 0 AND NVL(NEW_PO_COST,0) <> 0  ;" & vbCrLf _
                & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If


        SQLV = Replace(SQLV, "R1.NEW_PRICE", "ROUND(R1.NEW_PRICE,1)")


        If SQLV <> "" Then
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 Is " _
                & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
                & " Begin For R1 In C1 Loop" & vbCrLf _
                & " Update ICTSTYV1 Set " & vbCrLf _
                & SQLV & vbCrLf _
                & ", NEW_PO_COST_DATE = '" & COST_EFF_DATE & " '" & vbCrLf _
                & "   where STYLE_CODE = R1.STYLE_CODE and VEND_CODE = R1.VEND_CODE;" & vbCrLf _
                & " End Loop; End; End;"

            ASCDATA1.ExecuteSQL()
        End If



        SQLL = Replace(SQLL, "R1.NEW_PRICE", "ROUND(R1.NEW_PRICE,1)")

        If SQLL <> "" Then
            ASCMAIN1.sql = "" _
                 & "Begin Declare Cursor C1 is " _
                 & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
                 & " Begin For R1 in C1 Loop" & vbCrLf _
                 & " Update ICTSTYL1 Set " & vbCrLf _
                 & SQLL & vbCrLf _
                 & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where STYLE_CODE = R1.STYLE_CODE ;" & vbCrLf _
                 & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If

        'ASCMAIN1.sql = "" _
        '    & "Begin Declare Cursor C1 is " _
        '    & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
        '    & " Begin For R1 in C1 Loop" & vbCrLf _
        '    & " Update ICTSTYL1 Set SUB_UNIT_BAG_QTY = R1.SUB_UNIT_BAG_QTY, INNER_PACK_QTY = R1.INNER_PACK_QTY, CARTON_PACK_QTY = R1.CARTON_PACK_QTY, DUTY_RATE_CODE = R1.DUTY_RATE_CODE " & vbCrLf _
        '    & ", CASE_CUBE = R1.CASE_CUBE, STYLE_PRICE = R1.NEW_PRICE, STYLE_MATL_DESC = r1.STYLE_MATL_DESC, STYLE_PO_QTY_MIN = r1.STYLE_PO_QTY_MIN, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where STYLE_CODE = R1.STYLE_CODE ;" & vbCrLf _
        '    & " Update ICTSTYV1 Set VEND_ITEM_CODE = R1.VEND_ITEM_CODE, NEW_PO_COST = R1.NEW_PO_COST, NEW_PO_COST_DATE = '" & COST_EFF_DATE & "'" & vbCrLf _
        '    & "   where STYLE_CODE = R1.STYLE_CODE and VEND_CODE = R1.VEND_CODE;" & vbCrLf _
        '    & " End Loop; End; End;"


        'If COLS(11) = "1" Then
        '    SQLV1 = SQLV1 & "VEND_ITEM_CODE = R1.VEND_ITEM_CODE,"
        'End If

        'If COLS(12) = "1" Then
        '    SQLV1 = SQLV1 & "NEW_PO_COST = R1.NEW_PO_COST, "
        'End If

        If chkUPDATE_DIMS.Checked Then

            For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("")
                Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE") & ""
                Dim PACK_CODE As String = "CTN"

                ASCMAIN1.sql = "DELETE FROM ICTSTYLD where STYLE_CODE = '" & STYLE_CODE & "'"
                ASCDATA1.ExecuteSQL()

                Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                If rowICTSTYLD Is Nothing Then

                    rowICTSTYLD = dst.Tables("ICTSTYLD").NewRow
                    'If rowICTSTYLD Is Nothing Then
                    rowICTSTYLD.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYLD.Item("PACK_CODE") = PACK_CODE
                    rowICTSTYLD.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("CARTON_LENGTH") & ""), 2)
                    rowICTSTYLD.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WIDTH") & ""), 2)
                    rowICTSTYLD.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("CARTON_HEIGHT") & ""), 2)
                    rowICTSTYLD.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WEIGHT") & ""), 2)
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD)
                Else
                    rowICTSTYLD.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("CARTON_LENGTH") & ""), 2)
                    rowICTSTYLD.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WIDTH") & ""), 2)
                    rowICTSTYLD.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("CARTON_HEIGHT") & ""), 2)
                    rowICTSTYLD.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WEIGHT") & ""), 2)
                    'End If
                End If


                'Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE") & ""
                PACK_CODE = "INR"

                Dim rowICTSTYLD_INNER As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                If rowICTSTYLD_INNER Is Nothing Then

                    rowICTSTYLD_INNER = dst.Tables("ICTSTYLD").NewRow
                    'If rowICTSTYLD Is Nothing Then
                    rowICTSTYLD_INNER.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYLD_INNER.Item("PACK_CODE") = PACK_CODE
                    rowICTSTYLD_INNER.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("INNER_LENGTH") & ""), 2)
                    rowICTSTYLD_INNER.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("INNER_WIDTH") & ""), 2)
                    rowICTSTYLD_INNER.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("INNER_HEIGHT") & ""), 2)
                    rowICTSTYLD_INNER.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("INNER_WEIGHT") & ""), 2)
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD_INNER)
                Else
                    rowICTSTYLD_INNER.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("INNER_LENGTH") & ""), 2)
                    rowICTSTYLD_INNER.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("INNER_WIDTH") & ""), 2)
                    rowICTSTYLD_INNER.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("INNER_HEIGHT") & ""), 2)
                    rowICTSTYLD_INNER.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("INNER_WEIGHT") & ""), 2)
                    'End If

                End If


                PACK_CODE = "IT"

                Dim rowICTSTYLD_ITM As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                If rowICTSTYLD_ITM Is Nothing Then

                    rowICTSTYLD_ITM = dst.Tables("ICTSTYLD").NewRow
                    'If rowICTSTYLD Is Nothing Then
                    rowICTSTYLD_ITM.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYLD_ITM.Item("PACK_CODE") = PACK_CODE
                    rowICTSTYLD_ITM.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("STYLE_LENGTH") & ""), 2)
                    rowICTSTYLD_ITM.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("STYLE_WIDTH") & ""), 2)
                    rowICTSTYLD_ITM.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("STYLE_HEIGHT") & ""), 2)
                    rowICTSTYLD_ITM.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("STYLE_WEIGHT") & ""), 2)
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD_ITM)
                Else
                    rowICTSTYLD_ITM.Item("LENGTH") = Math.Round(Val(rowICTSTYLX.Item("STYLE_LENGTH") & ""), 2)
                    rowICTSTYLD_ITM.Item("WIDTH") = Math.Round(Val(rowICTSTYLX.Item("STYLE_WIDTH") & ""), 2)
                    rowICTSTYLD_ITM.Item("HEIGHT") = Math.Round(Val(rowICTSTYLX.Item("STYLE_HEIGHT") & ""), 2)
                    rowICTSTYLD_ITM.Item("WEIGHT") = Math.Round(Val(rowICTSTYLX.Item("STYLE_WEIGHT") & ""), 2)
                    'End If
                End If
            Next

            Update_Record_TDA("ICTSTYLD")

        End If


        EntryMode = ""


        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub



    Sub SAVE_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Saving ...")

        ASCMAIN1.sql = "TRUNCATE TABLE " & ICTSTYLX
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("ICTSTYLX")

        BeginTrans()


        ' need to add cost date to sql

        Dim COST_EFF_DATE As String = Format(Absx1.dteFor("COST_DATE").Value, "dd-MMM-yyyy")

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " _
            & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & "  Update ICTSTYL1 Set SUB_UNIT_BAG_QTY = R1.SUB_UNIT_BAG_QTY, INNER_PACK_QTY = R1.INNER_PACK_QTY, CARTON_PACK_QTY = R1.CARTON_PACK_QTY,  " & vbCrLf _
            & "  CASE_CUBE = R1.CASE_CUBE, STYLE_PRICE = R1.NEW_PRICE, STYLE_MATL_DESC = r1.STYLE_MATL_DESC, STYLE_PO_QTY_MIN = r1.STYLE_PO_QTY_MIN, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where STYLE_CODE = R1.STYLE_CODE ;" & vbCrLf _
            & " Update ICTSTYV1 Set VEND_ITEM_CODE = R1.VEND_ITEM_CODE, NEW_PO_COST = R1.NEW_PO_COST, NEW_PO_COST_DATE = '" & COST_EFF_DATE & "'" & vbCrLf _
            & "   where STYLE_CODE = R1.STYLE_CODE and VEND_CODE = R1.VEND_CODE;" & vbCrLf _
            & " End Loop; End; End;"

        ASCDATA1.ExecuteSQL()

        EntryMode = ""


        CommitTrans("Save Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYLX, "SSBS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Update Column")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        Select Case e.SourceControl.Name
            Case "grdICTSTYLX"
                tlb_sbt = DirectCast(tlb_pop.Tools("Update Column"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.SharedProps.Visible = False
                If grd.ActiveCell IsNot Nothing Then
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Dim row As DataRow = dst.Tables("TATCOLS1").Rows.Find(COLUMN_NAME)
                    If row IsNot Nothing Then
                        tlb_sbt.SharedProps.Visible = True
                        tlb_sbt.SharedProps.Caption = "Update " & row.Item("COLUMN_CAPTION")
                        tlb_sbt.Tag = ""
                        tlb_sbt.Checked = (row.Item("SEL") = "1")
                        tlb_sbt.Tag = COLUMN_NAME
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdICTSTYLX"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Update Column"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Dim COLUMN_NAME As String = tlb_sbt.Tag & ""
                If COLUMN_NAME <> "" Then
                    Dim row As DataRow = dst.Tables("TATCOLS1").Rows.Find(COLUMN_NAME)
                    If row IsNot Nothing Then
                        If tlb_sbt.Checked Then
                            row.Item("SEL") = "1"
                        Else
                            row.Item("SEL") = "0"
                        End If
                        Set_background_colors()
                    End If
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

    Sub Set_background_colors()
        For Each rowTATCOLS1 As DataRow In dst.Tables("TATCOLS1").Select
            Dim column_name As String = rowTATCOLS1.Item("COLUMN_NAME")
            If rowTATCOLS1.Item("SEL") = "1" Then
                grdICTSTYLX.DisplayLayout.Bands(0).Columns(column_name).CellAppearance.BackColor = Color.LightGreen
            Else
                grdICTSTYLX.DisplayLayout.Bands(0).Columns(column_name).CellAppearance.BackColor = Color.LightGray
            End If

        Next

    End Sub

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"

        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

    Sub CHECK_DIMENSIONS()
        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("")
            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE") & ""

            For Each PACK_CODE As String In New String() {"CTN", "INR", "IT"}


                Dim rowICTSTYLD As DataRow = LookUp("ICTSTYLD", New String() {STYLE_CODE, PACK_CODE})
                'Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                If rowICTSTYLD Is Nothing Then

                Else
                    If PACK_CODE = "CTN" Then

                        rowICTSTYLX.Item("CART_LEN_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_LENGTH")), 2)
                        rowICTSTYLX.Item("CART_WDTH_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WIDTH")), 2)
                        rowICTSTYLX.Item("CART_HT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_HEIGHT")), 2)
                        rowICTSTYLX.Item("CART_WGHT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WEIGHT")), 2)

                        If Val(rowICTSTYLX.Item("CARTON_LENGTH") & "") = 0 Then
                            rowICTSTYLX.Item("CARTON_LENGTH") = Math.Round(Val(rowICTSTYLD.Item("LENGTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("CARTON_WIDTH") & "") = 0 Then
                            rowICTSTYLX.Item("CARTON_WIDTH") = Math.Round(Val(rowICTSTYLD.Item("WIDTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("CARTON_HEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("CARTON_HEIGHT") = Math.Round(Val(rowICTSTYLD.Item("HEIGHT") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("CARTON_WEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("CARTON_WEIGHT") = Math.Round(Val(rowICTSTYLD.Item("WEIGHT") & ""), 2)
                        End If
                    End If
                    If PACK_CODE = "INR" Then
                        rowICTSTYLX.Item("INNER_LEN_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_LENGTH") & ""), 2)
                        rowICTSTYLX.Item("INNER_WDTH_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WIDTH") & ""), 2)
                        rowICTSTYLX.Item("INNER_HT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_HEIGHT") & ""), 2)
                        rowICTSTYLX.Item("INNER_WGHT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WEIGHT") & ""), 2)

                        If Val(rowICTSTYLX.Item("INNER_LENGTH") & "") = 0 Then
                            rowICTSTYLX.Item("INNER_LENGTH") = Math.Round(Val(rowICTSTYLD.Item("LENGTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("INNER_WIDTH") & "") = 0 Then
                            rowICTSTYLX.Item("INNER_WIDTH") = Math.Round(Val(rowICTSTYLD.Item("WIDTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("INNER_HEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("INNER_HEIGHT") = Math.Round(Val(rowICTSTYLD.Item("HEIGHT") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("INNER_WEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("INNER_WEIGHT") = Math.Round(Val(rowICTSTYLD.Item("WEIGHT") & ""), 2)
                        End If

                    End If

                    If PACK_CODE = "IT" Then
                        rowICTSTYLX.Item("ITM_LEN_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_LENGTH")), 2)
                        rowICTSTYLX.Item("ITM_WDTH_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WIDTH")), 2)
                        rowICTSTYLX.Item("ITM_HT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_HEIGHT")), 2)
                        rowICTSTYLX.Item("ITM_WGHT_SS") = Math.Round(Val(rowICTSTYLX.Item("CARTON_WEIGHT")), 2)

                        If Val(rowICTSTYLX.Item("STYLE_LENGTH") & "") = 0 Then
                            rowICTSTYLX.Item("STYLE_LENGTH") = Math.Round(Val(rowICTSTYLD.Item("LENGTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("STYLE_WIDTH") & "") = 0 Then
                            rowICTSTYLX.Item("STYLE_WIDTH") = Math.Round(Val(rowICTSTYLD.Item("WIDTH") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("STYLE_HEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("STYLE_HEIGHT") = Math.Round(Val(rowICTSTYLD.Item("HEIGHT") & ""), 2)
                        End If
                        If Val(rowICTSTYLX.Item("STYLE_WEIGHT") & "") = 0 Then
                            rowICTSTYLX.Item("STYLE_WEIGHT") = Math.Round(Val(rowICTSTYLD.Item("WEIGHT") & ""), 2)
                        End If
                    End If

                End If
            Next

        Next
    End Sub

    Sub Import_from_Excel()
        Dim openFileDialog1 As New OpenFileDialog
        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
        openFileDialog1.Filter = "xls files (*.xls)|*.xls"
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            Dim FILENAME As String = openFileDialog1.FileName
            Try
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" &
                "data source=" & FILENAME & ";" &
                "Extended Properties=Excel 8.0;"

                'Dim strConnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                '"data source=" & FILENAME & ";" &
                '"Extended Properties=Excel 8.0;"

                Dim objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                objConnection.Open()
                Dim dbSchema As DataTable = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                If dbSchema.Rows.Count = 0 Then
                    MsgBox("No Sheets Found")
                    Exit Sub
                End If
                Dim strSQL As String = "SELECT * FROM [" & dbSchema.Rows(0).Item("TABLE_NAME") & "]"
                Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                Dim dt As New DataTable
                objAdapter.Fill(dt)
                objConnection.Close()

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Data from XLS")

                Dim COLs As Int32 = dt.Columns.Count
                Dim PRDmax As Int32 = COLs - 3

                If COLs < 2 Then
                    MsgBox("There appear to be no Records to Import", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Else

                End If

                dst.Tables("ICTSTYLX").Rows.Clear()

                For Each row As DataRow In dt.Rows
                    Dim STYLE_CODE As String = row.Item(0) & ""

                    Dim rowICTSTYLX As DataRow = dst.Tables("ICTSTYLX").NewRow

                    'rowICTSTYLX.Item("STYLE_CODE") = STYLE_CODE
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    If rowICTSTYL1 Is Nothing Then
                        ' LOG ERROR
                    Else
                        Try
                            'For I As Integer = 0 To dst.Tables("ICTSTYLX").Columns.Count - 1 ' IS THIS CORRECT? DRC/ABS
                            For I As Integer = 0 To 38 ' dst.Tables("ICTSTYLX").Columns.Count - 1 ' IS THIS CORRECT? DRC/ABS
                                Dim DANA As String = row.Item(I) & ""
                                'If DANA = "MT16552" Or DANA = "MT16553" Then
                                '    DANA = DANA
                                'End If
                                If I = 1 Then
                                    rowICTSTYLX.Item(I) = Mid(row.Item(I), 1, 45)
                                Else
                                    If I = 2 Or I = 4 Then
                                        rowICTSTYLX.Item(I) = Mid(row.Item(I) & "A", 1, 1)
                                    Else
                                        If I = 99 + 1 Then ' was  If I = 23 + 1 Then
                                            'rowICTSTYLX.Item(I) = "0"
                                        Else
                                            'Dim danatype As String = rowICTSTYLX.
                                            If I = 7 Then
                                                rowICTSTYLX.Item(I) = Mid(row.Item(I), 1, 25)
                                            Else
                                                If I = 8 Or I = 11 Or I = 12 Or I = 13 Or I = 14 Or I = 15 Or I = 16 Or I = 17 Then
                                                    rowICTSTYLX.Item(I) = Val(row.Item(I) & "")
                                                Else

                                                    If I >= 20 And I < 21 Then  ' set new price GOOD
                                                        rowICTSTYLX.Item(I + 26) = Val(row.Item(I) & "")
                                                    Else
                                                        If I >= 21 And I <= 24 Then ' set Carton Dimensions
                                                            rowICTSTYLX.Item(I + 7) = Val(row.Item(I) & "")
                                                        Else
                                                            If I >= 25 And I <= 28 Then ' set innter Dimensions
                                                                rowICTSTYLX.Item(I - 1) = Val(row.Item(I) & "")
                                                            Else
                                                                If I >= 29 And I <= 32 Then ' set Style Dimensions
                                                                    rowICTSTYLX.Item(I - 9) = Val(row.Item(I) & "")
                                                                Else
                                                                    If I >= 33 And I <= 35 Then ' set duty and pfrt AND AR GOOD
                                                                        rowICTSTYLX.Item(I + 2) = Val(row.Item(I) & "")
                                                                    Else
                                                                        If I >= 35 And I < 36 Then ' new price skip
                                                                            'rowICTSTYLX.Item(I + 1) = row.Item(I) & ""
                                                                        Else
                                                                            If I >= 36 And I < 39 Then ' new price skip
                                                                                'rowICTSTYLX.Item(I) = row.Item(I) & ""
                                                                            Else
                                                                                rowICTSTYLX.Item(I) = row.Item(I) & ""
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If




                                    If I = 16 + 1 Or I = 17 + 1 Then
                                        Dim drccell As String = Val(row.Item(I) & "")
                                    End If
                                End If
                            Next
                            Dim danastyle As String = row.Item(0)
                            Dim dana2 As String = "what!"
                            dst.Tables("ICTSTYLX").Rows.Add(rowICTSTYLX)
                            row.Delete()

                        Catch ex As Exception
                            Stop
                        End Try
                    End If
                Next

                If dt.Rows.Count <> 0 Then
                    Dim frmASFMSGBF As New ASFMSGBF

                    frmASFMSGBF.Show_grd(dt, Me, "Records which Failed to Load")

                End If

            Catch ex As Exception

            End Try

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
    End Sub

    Private Sub UltraExplorerBar1_ItemClick(sender As Object, e As UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemClick

    End Sub

#End Region

End Class