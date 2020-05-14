Public Class ICFNEWS2
    Dim ICTSTYLX As String 'TABLE_NAME
    Dim sqlICTSTYLX As String


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTSTYL1.STYLE_CODE, ICTSTYL1.STYLE_DESC, REPLACE(REPLACE(ICTSTYL1.PURCH_NOTES,CHR(13),CHR(32)),CHR(10),CHR(32)) AS PURCH_NOTES, ICTSTYL1.STYLE_STATUS" & vbCrLf _
                & ", STYLE_CLASS_CODE, ICTSTYL1.VEND_CODE, ICTSTYV1.VEND_ITEM_CODE, ICTSTYL1.STYLE_PRICE " & vbCrLf _
                & ", ICTSTYL1.STYLE_UOM" & vbCrLf _
                & ", ICTSTYL1.SUB_UNIT_BAG_QTY, ICTSTYL1.INNER_PACK_QTY, ICTSTYL1.CARTON_PACK_QTY, ICTSTYL1.CASE_CUBE, decode (NVL(ICTSTYV1.NEW_PO_COST,0),0,NVL(ICTSTYV1.PO_COST,0),NVL(ICTSTYV1.NEW_PO_COST,0)) PO_COST" & vbCrLf _
                & ", ICTSTYL1.STYLE_PO_QTY_MIN, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_MATL_DESC " & vbCrLf _
                & ", T1.LENGTH AS STYLE_LENGTH, T1.WIDTH AS STYLE_WIDTH, T1.HEIGHT AS STYLE_HEIGHT, T1.WEIGHT AS STYLE_WEIGHT" & vbCrLf _
                & ", T2.LENGTH AS INNER_LENGTH, T2.WIDTH AS INNER_WIDTH, T2.HEIGHT AS INNER_HEIGHT, T2.WEIGHT AS INNER_WEIGHT" & vbCrLf _
                & ", T3.LENGTH AS CARTON_LENGTH, T3.WIDTH AS CARTON_WIDTH, T3.HEIGHT AS CARTON_HEIGHT, T2.WEIGHT AS CARTON_WEIGHT" & vbCrLf _
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

            With .Tables("ICTSTYLX")
                .Columns.Add("ATTR_CODE1")
                .Columns.Add("ATTR_CODE2")
                .Columns.Add("ATTR_CODE3")
                .Columns.Add("ATTR_CODE4")
                .Columns.Add("ATTR_CODE5")
            End With

            Dim s As New System.Text.StringBuilder With {.Length = 0}
            s.AppendLine("SELECT S3.STYLE_CODE,")
            s.AppendLine("S3.ATTR_CODE,")
            s.AppendLine("NVL(A1.ATT_RANK,9) ATT_RANK")
            s.AppendLine("FROM ICTSTYL3 S3, ICTATTR1 A1")
            s.AppendLine("WHERE S3.ATTR_CODE = A1.ATTR_CODE")
            ASCMAIN1.sql = s.ToString
            Create_TDA(.Tables.Add, "ICTSTYL3", "**", 0, False)
        End With

        Fill_Records("ICTSTYL3")

        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")

        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")

        With grdICTSTYLX.DisplayLayout.Bands(0)
            .Columns("STYLE_CODE").Header.Fixed = True
            .Columns("STYLE_DESC").Header.Fixed = True
            .Columns("STYLE_STATUS").Header.Fixed = True
            '.Columns("STYLE_COLOR_STATUS").Header.Fixed = True
            ' .Columns("COLOR_DESC").Header.Fixed = True


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

            With .Columns("PURCH_NOTES")
                .Header.Fixed = True
                .Header.Caption = "Procurement- Purchase Note"
                .Width = 200
                .Header.VisiblePosition = 3
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_STATUS")
                .Header.Fixed = True
                .Header.Caption = "Status"
                .Header.VisiblePosition = 4
                .Width = 100
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_CLASS_CODE")
                .Header.Fixed = True
                .Header.Caption = "Class"
                .Width = 100
                .Header.VisiblePosition = 5
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("VEND_CODE")
                .Header.Fixed = False
                .Header.Caption = "Supplier"
                .Width = 120
                .Header.VisiblePosition = 6
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("VEND_ITEM_CODE")
                .Header.Fixed = False
                .Header.Caption = "Supplier Style"
                .Width = 160
                .Header.VisiblePosition = 7
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_PRICE")
                .Header.Fixed = False
                .Header.Caption = "List Price"
                .Width = 120
                .Header.VisiblePosition = 8
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            With .Columns("STYLE_UOM")
                .Header.Fixed = False
                .Header.Caption = "UOM"
                .Width = 80
                .Header.VisiblePosition = 9
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("SUB_UNIT_BAG_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Bag"
                .Width = 80
                .Header.VisiblePosition = 10
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_PACK_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Box"
                .Width = 80
                .Header.VisiblePosition = 11
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_PACK_QTY")
                .Header.Fixed = False
                .Header.Caption = "#/Cs"
                .Width = 80
                .Header.VisiblePosition = 12
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CASE_CUBE")
                .Header.Fixed = False
                .Header.Caption = "Cube"
                .Width = 80
                .Header.VisiblePosition = 13
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            With .Columns("PO_COST")
                .Header.Fixed = False
                .Header.Caption = "Curr PO"
                .Width = 80
                .Header.VisiblePosition = 14
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_PO_QTY_MIN")
                .Header.Fixed = False
                .Header.Caption = "MOQ"
                .Width = 60
                .Header.VisiblePosition = 15
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("DUTY_RATE_CODE")
                .Header.Fixed = False
                .Header.Caption = "HTSUS"
                .Width = 200
                .Header.VisiblePosition = 16
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_MATL_DESC")
                .Header.Fixed = False
                .Header.Caption = "Material"
                .Width = 200
                .Header.VisiblePosition = 17
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With



            With .Columns("NEW_PRICE")
                .Header.Fixed = False
                .Header.Caption = "New Price"
                .Width = 80
                .Header.VisiblePosition = 18
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            End With

            ' dimensions
            With .Columns("CARTON_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Carton Length (In)"
                .Width = 80
                .Header.VisiblePosition = 19
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Cart Width (In)"
                .Width = 100
                .Header.VisiblePosition = 20
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Cart Height (In)"
                .Width = 100
                .Header.VisiblePosition = 21
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("CARTON_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Carton Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 22
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With



            With .Columns("INNER_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Inner Length (In)"
                .Width = 100
                .Header.VisiblePosition = 23
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Inner Width (In)"
                .Width = 100
                .Header.VisiblePosition = 24
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Inner Height (In)"
                .Width = 100
                .Header.VisiblePosition = 25
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("INNER_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Inner Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 26
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_LENGTH")
                .Header.Fixed = False
                .Header.Caption = "Style Length (In)"
                .Width = 100
                .Header.VisiblePosition = 27
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_WIDTH")
                .Header.Fixed = False
                .Header.Caption = "Style Width (In)"
                .Width = 100
                .Header.VisiblePosition = 28
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_HEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Style Height (In)"
                .Width = 100
                .Header.VisiblePosition = 29
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("STYLE_WEIGHT")
                .Header.Fixed = False
                .Header.Caption = "Style Wght (Lbs)"
                .Width = 100
                .Header.VisiblePosition = 30
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With


            ' end of dimensions


            With .Columns("FGT")
                .Header.Fixed = False
                .Header.Caption = "FGT"
                .Width = 60
                .Header.VisiblePosition = 31
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("DUTY")
                .Header.Fixed = False
                .Header.Caption = "Duty"
                .Width = 60
                .Header.VisiblePosition = 32
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("PFRT")
                .Header.Fixed = False
                .Header.Caption = "PFRT"
                .Width = 60
                .Header.VisiblePosition = 33
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("PCT")
                .Header.Fixed = False
                .Header.Caption = "PCT"
                .Width = 60
                .Header.VisiblePosition = 34
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
            End With

            With .Columns("ATTR_CODE1")
                .Header.Fixed = False
                .Header.Caption = "Attr 1"
                .Width = 60
                .Header.VisiblePosition = 35
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Hidden = True
            End With
            With .Columns("ATTR_CODE2")
                .Header.Fixed = False
                .Header.Caption = "Attr 2"
                .Width = 60
                .Header.VisiblePosition = 36
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Hidden = True
            End With
            With .Columns("ATTR_CODE3")
                .Header.Fixed = False
                .Header.Caption = "Attr 3"
                .Width = 60
                .Header.VisiblePosition = 37
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Hidden = True
            End With
            With .Columns("ATTR_CODE4")
                .Header.Fixed = False
                .Header.Caption = "Attr 4"
                .Width = 60
                .Header.VisiblePosition = 38
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Hidden = True
            End With
            With .Columns("ATTR_CODE5")
                .Header.Fixed = False
                .Header.Caption = "Attr 5"
                .Width = 60
                .Header.VisiblePosition = 39
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = Drawing.Color.LightPink
                .Hidden = True
            End With





            'For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
            '    gcol.Header.Appearance.BackColor = Drawing.Color.White
            '    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
            '    If gcol.Key = "NEW_PO_COST" Then
            '        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            '    End If
            'Next
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

                Fill_Records("ICTSTYLX")
            End If
        End If

        Sort_grdColumns(grdICTSTYLX, "STYLE_CODE")

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        EntryMode = ""

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub



 
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTSTYLX, "SSBSS", "Show Filter", "Show GroupBox", "Style Status Inquiry", "Update Column", "Show All Attributes")
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
            Case "Show All Attributes"
                tlb_sbt = Nothing
                tlb_sbt = DirectCast(tlb.Tools("Show All Attributes"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Me.Cursor = Cursors.WaitCursor
                    For Each rowICTSTYLX As DataRow In dst.Tables("ICTSTYLX").Select("", "STYLE_CODE")
                        Dim STYLE_CODE As String = rowICTSTYLX.Item("STYLE_CODE").ToString.Replace("'", "")
                        Dim rowICTSTYL31 As DataRow = dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK = '1'", STYLE_CODE)).FirstOrDefault
                        If Not IsNothing(rowICTSTYL31) Then
                            rowICTSTYLX.Item("ATTR_CODE1") = rowICTSTYL31.Item("ATTR_CODE")
                        End If
                        Dim nextI As Integer = 2
                        For Each rowICTSTYL3 As DataRow In dst.Tables("ICTSTYL3").Select(String.Format("STYLE_CODE = '{0}' AND ATT_RANK <> '1'", STYLE_CODE), "ATTR_CODE")
                            If nextI > 5 Then Exit For
                            rowICTSTYLX.Item(String.Format("ATTR_CODE{0}", nextI)) = rowICTSTYL3.Item("ATTR_CODE")
                            nextI += 1
                        Next
                    Next
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = False
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = False
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = False
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = False
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE5").Hidden = False
                    grdICTSTYLX.UpdateData()
                Else
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE1").Hidden = True
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE2").Hidden = True
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE3").Hidden = True
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE4").Hidden = True
                    grdICTSTYLX.DisplayLayout.Bands(0).Columns("ATTR_CODE5").Hidden = True
                End If
                Me.Cursor = Cursors.Default
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
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                "data source=" & FILENAME & ";" & _
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
                            For I As Integer = 0 To dst.Tables("ICTSTYLX").Columns.Count - 1 ' IS THIS CORRECT? DRC/ABS
                                Dim DANA As String = row.Item(I) & ""
                                If DANA = "MTX49970L" Then
                                    DANA = DANA
                                End If

                                If I = 2 Or I = 4 Then
                                    rowICTSTYLX.Item(I) = Mid(row.Item(I) & "A", 1, 1)

                                Else
                                    If I = 23 Then
                                        rowICTSTYLX.Item(I) = "0"
                                    Else
                                        'Dim danatype As String = rowICTSTYLX.
                                        rowICTSTYLX.Item(I) = row.Item(I) & ""
                                    End If


                                    If I = 16 Or I = 17 Then
                                        Dim drccell As String = row.Item(I)
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