Public Class ICFNEWS4
    Dim ICTSTYLX As String 'TABLE_NAME
    Dim sqlICTSTYLX As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, ICTSTYL1.STYLE_STATUS" & vbCrLf _
                & ", STYLE_CLASS_CODE, ICTSTYL1.VEND_CODE, ICTSTYV1.VEND_ITEM_CODE " & vbCrLf _
                & ", ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE" & vbCrLf _
                & ", ICTSTYL1.STYLE_PO_QTY_MIN, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_MATL_DESC " & vbCrLf _
                & ", T3.LENGTH AS CARTON_LENGTH, T3.WIDTH AS CARTON_WIDTH, T3.HEIGHT AS CARTON_HEIGHT, T2.WEIGHT AS CARTON_WEIGHT" & vbCrLf _
                & ", T2.LENGTH AS INNER_LENGTH, T2.WIDTH AS INNER_WIDTH, T2.HEIGHT AS INNER_HEIGHT, T2.WEIGHT AS INNER_WEIGHT" & vbCrLf _
                & ", T1.LENGTH AS STYLE_LENGTH, T1.WIDTH AS STYLE_WIDTH, T1.HEIGHT AS STYLE_HEIGHT, T1.WEIGHT AS STYLE_WEIGHT" & vbCrLf _
                & " from ICTSTYL1,ICTSTYV1, ICTSTYLD T1, ICTSTYLD T2, ICTSTYLD T3" & vbCrLf _
                & " where ICTSTYV1.STYLE_CODE = ICTSTYL1.STYLE_CODE " & vbCrLf _
                & "   and ICTSTYV1.VEND_CODE (+) = ICTSTYL1.VEND_CODE " & vbCrLf _
                & "   and T1.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T1.PACK_CODE (+) = 'IT' " & vbCrLf _
                & "   and T2.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T2.PACK_CODE (+) = 'INR' " & vbCrLf _
                & "   and T3.STYLE_CODE (+) = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   and T3.PACK_CODE (+) = 'CTN' " & vbCrLf


            sqlICTSTYLX = ASCMAIN1.sql

            ICTSTYLX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add NEW_PRICE VARCHAR2(20)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add FGT  VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add DUTY  VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PFRT VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTSTYLX & " Add PCT  VARCHAR2(10)")


            ASCMAIN1.sql = "Select * from " & ICTSTYLX
            Create_TDA(.Tables.Add("ICTSTYLX"), ICTSTYLX, "**", 0, True)

            Create_TDA(.Tables.Add, "ICTSTYLD", "*")

            Create_TDA(.Tables.Add, "ICTSTYL1", "*")

        End With


        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")

        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("STYLE_STATUS").Header.Fixed = True
            '.Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            ' .Columns("COLOR_DESC").Header.Fixed = True

            .Columns("NEW_PRICE").Hidden = True
            .Columns("FGT").Hidden = True
            .Columns("DUTY").Hidden = True
            .Columns("PFRT").Hidden = True
            .Columns("PCT").Hidden = True

            With .Columns("STYLE_CODE")
                .Header.Fixed = True
                .Header.Caption = "Style"
                .Width = 140
                .Header.VisiblePosition = 1
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With
            With .Columns("STYLE_DESC")
                .Header.Fixed = True
                .Header.Caption = "Description"
                .Width = 200
                .Header.VisiblePosition = 2
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_STATUS")
                .Header.Fixed = True
                .Header.Caption = "Status"
                .Header.VisiblePosition = 3
                .Width = 100
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_CLASS_CODE")
                .Header.Fixed = True
                .Header.Caption = "Class"
                .Width = 100
                .Header.VisiblePosition = 4
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("VEND_CODE")
                .Header.Fixed = False
                .Header.Caption = "Supplier"
                .Width = 120
                .Header.VisiblePosition = 5
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("VEND_ITEM_CODE")
                .Header.Fixed = False
                .Header.Caption = "Supplier Style"
                .Width = 160
                .Header.VisiblePosition = 6
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            'With .Columns("STYLE_PRICE")
            '    .Header.Fixed = False
            '    .Header.Caption = "List Price"
            '    .Width = 120
            '    .Header.VisiblePosition = 7
            '    .Header.Appearance.BackColor = Drawing.Color.White
            '    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            'End With


            With .Columns("STYLE_UOM")
                .Header.Fixed = False
                .Header.Caption = "UOM"
                .Width = 80
                .Header.VisiblePosition = 8
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("SUB_UNIT_BAG_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Bag"
                .Width = 80
                .Header.VisiblePosition = 9
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_PACK_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Box"
                .Width = 80
                .Header.VisiblePosition = 10
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_PACK_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Cs"
                .Width = 80
                .Header.VisiblePosition = 11
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CASE_CUBE")
                .Header.Fixed = False
                .Header.Caption = "Cube"
                .Width = 80
                .Header.VisiblePosition = 12
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            'With .Columns("PO_COST")
            '    .Header.Fixed = False
            '    .Header.Caption = "Curr PO"
            '    .Width = 80
            '    .Header.VisiblePosition = 13
            '    .Header.Appearance.BackColor = Drawing.Color.White
            '    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            'End With

            With .Columns("STYLE_PO_QTY_MIN")
                .Header.Fixed = False
                .Header.Caption = "MOQ"
                .Width = 60
                .Header.VisiblePosition = 14
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("DUTY_RATE_CODE")
                .Header.Fixed = False
                .Header.Caption = "HTSUS"
                .Width = 200
                .Header.VisiblePosition = 15
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_MATL_DESC")
                .Header.Fixed = False
                .Header.Caption = "Material"
                .Width = 200
                .Header.VisiblePosition = 16
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            ' dimensions
            With .Columns("CARTON_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Carton Length (In)"
                .Width = 80
                .Header.VisiblePosition = 18
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Cart Width (In)"
                .Width = 100
                .Header.VisiblePosition = 19
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Cart Height (In)"
                .Width = 100
                .Header.VisiblePosition = 20
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Carton Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 21
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With



            With .Columns("INNER_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Inner Length (In)"
                .Width = 100
                .Header.VisiblePosition = 22
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Inner Width (In)"
                .Width = 100
                .Header.VisiblePosition = 23
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Inner Height (In)"
                .Width = 100
                .Header.VisiblePosition = 24
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Inner Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 25
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Style Length (In)"
                .Width = 100
                .Header.VisiblePosition = 26
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Style Width (In)"
                .Width = 100
                .Header.VisiblePosition = 27
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Style Height (In)"
                .Width = 100
                .Header.VisiblePosition = 28
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Style Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 29
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            ' end of dimensions


            With .Columns("FGT")
                .Header.Fixed = False
                .Header.Caption = "FGT"
                .Width = 60
                .Header.VisiblePosition = 30
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With




        End With

        ASCMAIN1.Add_Value_List(grdICTSTYLX, "STYLE_STATUS")
        spl.Panel1Collapsed = True


        Dim CC As Integer = dst.Tables("ICTSTYLX").Columns.Count
        For i As Integer = 1 To CC - 1
            Dim DC As DataColumn = dst.Tables("ICTSTYLX").Columns(i)

            If DC.ColumnName = "STYLE_CODE" Or DC.ColumnName = "STYLE_DESC" Or DC.ColumnName = "UPC_CODE" Or DC.ColumnName = "PCT" Or DC.ColumnName = "PFRT" Or DC.ColumnName = "FGT" Or DC.ColumnName = "STYLE_STATUS" Or DC.ColumnName = "PO_COST" Or DC.ColumnName = "DUTY" Or DC.ColumnName = "VEND_CODE" Or DC.ColumnName = "STYLE_PRICE" Or DC.ColumnName = "STYLE_CLASS_CODE" Then

            Else

            End If
        Next




    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load From Spreadsheet"
                Import_from_Excel()
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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Done"
                Mode_Settings(False)

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
                    ' .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Visible = True
                    .Items("Cancel").Visible = True

                    ' .Items("Update").Visible = (ScreenMode And EntryMode = "E")
                    ' .Items("Cancel").Visible = (ScreenMode And EntryMode = "E")
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
        For Each TABLE_NAME As String In New String() {"ICTSTYLX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "E" Then
        Else


            ASCMAIN1.sql = "TRUNCATE TABLE " & ICTSTYLX
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO " & ICTSTYLX & " SELECT X.*,'','','','','' FROM (" & sqlICTSTYLX & ") X "
            ASCDATA1.ExecuteSQL()

            Fill_Records("ICTSTYLX")
        End If


        Sort_grdColumns(grdICTSTYLX, "STYLE_CODE")

        ASCMAIN1.Progress("")
    End Sub




    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        EntryMode = ""

        BeginTrans()

        For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("")
            Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE") & ""
            Dim PACK_CODE As String = "CTN"
            Dim MOQ As Decimal = Val(rowICTSTYLX.Item("STYLE_PO_QTY_MIN") & "")
            Dim MATL_DESC As String = rowICTSTYLX.Item("STYLE_MATL_DESC") & ""



            If chkUPDATE_DIMS.Checked Then
                ASCMAIN1.sql = "DELETE FROM ICTSTYLD where STYLE_CODE = '" & STYLE_CODE & "'"
                ASCDATA1.ExecuteSQL()

                ' Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                Dim rowICTSTYLD As DataRow = dst.Tables("ICTSTYLD").Rows.Find({STYLE_CODE, PACK_CODE})
                If rowICTSTYLD Is Nothing Then

                    rowICTSTYLD = dst.Tables("ICTSTYLD").NewRow
                    'If rowICTSTYLD Is Nothing Then
                    rowICTSTYLD.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYLD.Item("PACK_CODE") = PACK_CODE
                    rowICTSTYLD.Item("LENGTH") = rowICTSTYLX.Item("CARTON_LENGTH") & ""
                    rowICTSTYLD.Item("WIDTH") = rowICTSTYLX.Item("CARTON_WIDTH") & ""
                    rowICTSTYLD.Item("HEIGHT") = rowICTSTYLX.Item("CARTON_HEIGHT") & ""
                    rowICTSTYLD.Item("WEIGHT") = rowICTSTYLX.Item("CARTON_WEIGHT") & ""
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD)
                Else
                    rowICTSTYLD.Item("LENGTH") = rowICTSTYLX.Item("CARTON_LENGTH") & ""
                    rowICTSTYLD.Item("WIDTH") = rowICTSTYLX.Item("CARTON_WIDTH") & ""
                    rowICTSTYLD.Item("HEIGHT") = rowICTSTYLX.Item("CARTON_HEIGHT") & ""
                    rowICTSTYLD.Item("WEIGHT") = rowICTSTYLX.Item("CARTON_WEIGHT") & ""
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
                    rowICTSTYLD_INNER.Item("LENGTH") = rowICTSTYLX.Item("INNER_LENGTH") & ""
                    rowICTSTYLD_INNER.Item("WIDTH") = rowICTSTYLX.Item("INNER_WIDTH") & ""
                    rowICTSTYLD_INNER.Item("HEIGHT") = rowICTSTYLX.Item("INNER_HEIGHT") & ""
                    rowICTSTYLD_INNER.Item("WEIGHT") = rowICTSTYLX.Item("INNER_WEIGHT") & ""
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD_INNER)
                Else
                    rowICTSTYLD_INNER.Item("LENGTH") = rowICTSTYLX.Item("INNER_LENGTH") & ""
                    rowICTSTYLD_INNER.Item("WIDTH") = rowICTSTYLX.Item("INNER_WIDTH") & ""
                    rowICTSTYLD_INNER.Item("HEIGHT") = rowICTSTYLX.Item("INNER_HEIGHT") & ""
                    rowICTSTYLD_INNER.Item("WEIGHT") = rowICTSTYLX.Item("INNER_WEIGHT") & ""
                    'End If

                End If


                PACK_CODE = "IT"

                Dim rowICTSTYLD_ITM As DataRow = dst.Tables("ICTSTYLD").Rows.Find(New Object() {STYLE_CODE, PACK_CODE})
                If rowICTSTYLD_ITM Is Nothing Then

                    rowICTSTYLD_ITM = dst.Tables("ICTSTYLD").NewRow
                    'If rowICTSTYLD Is Nothing Then
                    rowICTSTYLD_ITM.Item("STYLE_CODE") = STYLE_CODE
                    rowICTSTYLD_ITM.Item("PACK_CODE") = PACK_CODE
                    rowICTSTYLD_ITM.Item("LENGTH") = rowICTSTYLX.Item("STYLE_LENGTH") & ""
                    rowICTSTYLD_ITM.Item("WIDTH") = rowICTSTYLX.Item("STYLE_WIDTH") & ""
                    rowICTSTYLD_ITM.Item("HEIGHT") = rowICTSTYLX.Item("STYLE_HEIGHT") & ""
                    rowICTSTYLD_ITM.Item("WEIGHT") = rowICTSTYLX.Item("STYLE_WEIGHT") & ""
                    dst.Tables("ICTSTYLD").Rows.Add(rowICTSTYLD_ITM)
                Else
                    rowICTSTYLD_ITM.Item("LENGTH") = rowICTSTYLX.Item("STYLE_LENGTH") & ""
                    rowICTSTYLD_ITM.Item("WIDTH") = rowICTSTYLX.Item("STYLE_WIDTH") & ""
                    rowICTSTYLD_ITM.Item("HEIGHT") = rowICTSTYLX.Item("STYLE_HEIGHT") & ""
                    rowICTSTYLD_ITM.Item("WEIGHT") = rowICTSTYLX.Item("STYLE_WEIGHT") & ""
                    'End If
                End If
                Update_Record_TDA("ICTSTYLD")
            End If


            If chkUPDATE_MOQ.Checked Then
                ASCMAIN1.sql = "Update ICTSTYL1 " _
                & " Set STYLE_PO_QTY_MIN = '" & MOQ & "'"
                ASCMAIN1.sql &= " where STYLE_CODE ='" & STYLE_CODE & "'"
                ASCDATA1.ExecuteSQL()
            End If

            If chkUPDATE_MATL.Checked Then
                ASCMAIN1.sql = "Update ICTSTYL1 " _
                & " Set STYLE_MATL_DESC = '" & MATL_DESC & "'"
                ASCMAIN1.sql &= " where STYLE_CODE ='" & STYLE_CODE & "'"
                ASCDATA1.ExecuteSQL()
            End If


        Next





        'If chkUPDATE_MATL.Checked Then

        '    ASCMAIN1.sql = "" _
        '         & "Begin Declare Cursor C1 is " _
        '         & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
        '         & " Begin For R1 in C1 Loop" & vbCrLf _
        '         & " Update ICTSTYL1 Set STYLE_MATL_DESC = r1.STYLE_MATL_DESC" & vbCrLf _
        '         & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where STYLE_CODE = R1.STYLE_CODE ;" & vbCrLf _
        '         & " End Loop; End; End;"
        '    ASCDATA1.ExecuteSQL()

        'End If

        'If chkUPDATE_MOQ.Checked Then
        '    ASCMAIN1.sql = "" _
        '         & "Begin Declare Cursor C1 is " _
        '         & " Select * from  " & ICTSTYLX & ";" & vbCrLf _
        '         & " Begin For R1 in C1 Loop" & vbCrLf _
        '         & " Update ICTSTYL1 Set STYLE_PO_QTY_MIN = r1.STYLE_PO_QTY_MIN" & vbCrLf _
        '         & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where STYLE_CODE = R1.STYLE_CODE ;" & vbCrLf _
        '         & " End Loop; End; End;"
        '    ASCDATA1.ExecuteSQL()
        'End If

        CommitTrans("Update Complete")

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = Lookup("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub

#End Region

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

    Sub Import_from_Excel()
        Dim openFileDialog1 As New OpenFileDialog
        openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
        openFileDialog1.Filter = "xls files (*.xls)|*.xls"
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = DialogResult.OK Then

            Dim FILENAME As String = openFileDialog1.FileName
            Try
                'Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" &
                '"data source=" & FILENAME & ";" &
                '"Extended Properties=Excel 8.0;"

                Dim strConnection As String = "Provider=Microsoft.ACE.OLEDB.12.0;" &
                "data source=" & FILENAME & ";" &
                "Extended Properties=Excel 8.0;"

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
                            For I As Integer = 0 To 25 ' IS THIS CORRECT? DRC/ABS
                                Dim DANA As String = row.Item(I) & ""
                                If DANA = "MT16552" Or DANA = "MT16553" Then
                                    DANA = DANA
                                End If

                                If I = 2 Then
                                    rowICTSTYLX.Item(I) = Mid(row.Item(I) & "A", 1, 1)
                                Else
                                    If I = 99 + 1 Then ' was  If I = 23 + 1 Then
                                        'rowICTSTYLX.Item(I) = "0"
                                    Else
                                        'Dim danatype As String = rowICTSTYLX.
                                        If I < 14 Then
                                            rowICTSTYLX.Item(I) = val(row.Item(I) & "")
                                        Else
                                            rowICTSTYLX.Item(I) = Val(row.Item(I) & "")

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

#End Region

End Class