Imports System.Drawing
Imports Infragistics.Win.Touch
Imports Infragistics.Win.UltraWinGrid

Public Class WHFPHYC1
    Dim rowWHTPHYC1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim WHSE_CODE As String
    Dim TICKET_NO As String
    Dim LOCATION_CODE As String = ""
    Dim BAR_CODE As String = ""

    Dim grdAttention_app As New Infragistics.Win.Appearance
    Dim grdWARNING_app As New Infragistics.Win.Appearance
    Dim SelUser As String

    Dim WHTPHYCZ As String = "" ' ARE WE USING THIS?
    Dim WHTPHYCL As String = ""
    Dim WHTPHYCV As String = ""
    Dim WHTPHYCR As String = ""

    Dim sqlWHTPHYCL As String = ""
    Dim sqlWHTPHYCV As String = ""
    Dim sqlWHTPHYCR As String = ""

    Dim variances_were_rebuilt As Boolean = False

    Dim BAR_CODE_PFX As String = ""

    ' NOTE THAT IF WE DO NOT INITIALIZE COUNTS TABLES AT MONTH END, THAT THIS SCREEN WILL SHOW COUNTS (WHICH IS USEFUL) AFTER THE PI HAS BEEN POSTED
    '  HOWEVER, THE VARIANCE WILL WORK ONLY FOR LOCATABLE WHSES SINCE WE COMPARE TO WHTLOCB0 (SNAPSHOT BY LOCATION, ESTABLISHED AT P/I INIT). 
    '  THE BOOK INVENTORY WILL SHOW WITH BAD DATA FOR NON-LOCATABLE WHSES, SINCE IT IS LOOKING AT ICTSTAT1.
    '  BUT THIS MIGHT BE EASILY FIXED BY USING THE YP OF THE LAST PI UPDATE

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "WHFPHYCI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        Rebuild_Variances(True)

        With dst
            ASCMAIN1.sql = "Select WHTPHYC1.*, WHTPHYC1.STYLE_CODE_MIN STYLE_CODE, WHTPHYC1.SC_MIN SC" & vbCrLf _
                & ", WHTLOCBL.BOOK_CTNS, WHTLOCBL.BOOK_UNITS, WHTLOCBL.BOOK_INVTY_ADJ" & vbCrLf _
                & ", NVL(WHTPHYC1.PHYS_UNITS,0) - (NVL(WHTLOCBL.BOOK_UNITS,0) + NVL(WHTLOCBL.BOOK_INVTY_ADJ,0)) UNIT_VARIANCE" & vbCrLf _
                & ", ABS(NVL(WHTPHYC1.PHYS_UNITS,0) - (NVL(WHTLOCBL.BOOK_UNITS,0) + NVL(WHTLOCBL.BOOK_INVTY_ADJ,0))) ABSOLUTE_VARIANCE" & vbCrLf _
                & ", NVL(A.LAST_ACTIVITY,'01-JAN-1999') LAST_ACTIVITY" & vbCrLf _
                & ", WHTLOCM1.LOCATION_USE, CASE WHEN WHTLOCM1.LOCATION_USE IN('A','L','E') THEN '0' ELSE '1' END VIRTUAL" & vbCrLf _
                & " from WHTLOCM1, WHTPHYC1, WHTLOCBL" & vbCrLf _
                & vbCrLf _
                & ", (Select WHSE_CODE, LOCATION_CODE, MAX(INIT_DATE) LAST_ACTIVITY " & vbCrLf _
                & " from WHTLOCB2 " & vbCrLf _
                & " where WHSE_CODE = :PARM1 " & vbCrLf _
                & " group by WHSE_CODE, LOCATION_CODE) A" & vbCrLf _
                & vbCrLf _
                & " where WHTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and WHTLOCBL.WHSE_CODE (+) = WHTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and WHTLOCBL.LOCATION_CODE (+) = WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & "   and A.WHSE_CODE (+) = WHTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and A.LOCATION_CODE (+) = WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE (+) = WHTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and WHTLOCM1.LOCATION_CODE (+) = WHTPHYC1.LOCATION_CODE"
            Create_TDA(.Tables.Add, "WHTPHYCX", "**", 0, False, "V", 2)
            With .Tables("WHTPHYCX")
                .Columns("UNIT_VARIANCE").DataType = GetType(System.Int64)
                .Columns("ABSOLUTE_VARIANCE").DataType = GetType(System.Int64)
                .Columns.Add("SELECTED")
                .Columns("SELECTED").DefaultValue = "0"
            End With

            ASCMAIN1.sql = $"Select * from {WHTPHYCV}"
            Create_TDA(.Tables.Add, "WHTPHYCV", "**", 0, False, "", 2)
            'Create_TDA(.Tables.Add, "WHTPHYCV", "**", 0, False, "VVVV")

            With .Tables("WHTPHYCV")
                .Columns("BOOK").DataType = GetType(System.Int64)
                .Columns("PHYS").DataType = GetType(System.Int64)
                'If Not .Columns.Contains("VARIANCE") Then
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
                .Columns.Add("VARIANCE_COST", GetType(System.Int64), "ISNULL(STYLE_COST,0) * (ISNULL(PHYS,0) - ISNULL(BOOK,0))")
                .Columns.Add("STYLE_COST_VAR", GetType(System.Int64))
                .Columns.Add("COUNTED_LOCS", GetType(System.String))
                .Columns.Add("BOOKED_LOCS", GetType(System.String))
                'Else
                '    .Columns("VARIANCE").DataType = GetType(System.Int64)
                '    .Columns("VARIANCE_COST").DataType = GetType(System.Int64)
                '    .Columns("STYLE_COST_VAR").DataType = GetType(System.Int64)
                'End If
            End With

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            Create_TDA(.Tables.Add, "WHTPHYC1", "*")

            Create_TDA(.Tables.Add, "WHTPHYC2", "*", 2)

            ASCMAIN1.sql = "Select WHTPHYC3.*, ICTSTYL1.STYLE_DESC" _
               & " from WHTPHYC3,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTPHYC3.STYLE_CODE"
            Create_TDA(.Tables.Add, "WHTPHYC3", "**", 2)
            .Tables("WHTPHYC3").Columns.Add("SC", GetType(System.String), "STYLE_CODE + '-' + COLOR_CODE")

            ASCMAIN1.sql = "Select WHTPHYC3.*" & vbCrLf _
                & ", WHTPHYC1.LOCATION_CODE, WHTPHYC1.INIT_OPER, WHTPHYC1.INIT_DATE" & vbCrLf _
                & " from WHTPHYC3,WHTPHYC1" & vbCrLf _
                & " where WHTPHYC3.WHSE_CODE = :PARM1 and WHTPHYC3.STYLE_CODE = :PARM2 and WHTPHYC3.COLOR_CODE = :PARM3" & vbCrLf _
                & "   and WHTPHYC1.WHSE_CODE = WHTPHYC3.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC1.TICKET_NO = WHTPHYC3.TICKET_NO" & vbCrLf _
                & "   and WHTPHYC1.TICKET_STATUS = 'A'"
            Create_TDA(.Tables.Add, "WHTPHYCI", "**", 0, False, "VVV", 3)

            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE, (WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) LOCATION_QTY ,INIT_DATE ,INIT_OPER ,LAST_DATE ,LAST_OPER ,LOCATION_QTY_WAVE" _
                & " from WHTLOCB0 where WHSE_CODE = :PARM1 and STYLE_CODE = :PARM2 and COLOR_CODE = :PARM3"
            Create_TDA(.Tables.Add, "WHTLOCBA", "**", 0, False, "VVV", 5)

            ASCMAIN1.sql = "Select * from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False)

            ASCMAIN1.sql = $"Select * from {WHTPHYCL}"
            Create_TDA(.Tables.Add, "WHTPHYCL", "**", 0, False, "", 2)
            'Create_TDA(.Tables.Add, "WHTPHYCL", "**", 0, False, "V", 2)
            With .Tables("WHTPHYCL")
                .Columns("TICKETS").DataType = GetType(System.Int64)
                .Columns("PHYS_STYLE_COLORS").DataType = GetType(System.Int64)
                .Columns("BOOK_STYLE_COLORS").DataType = GetType(System.Int64)
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
                .Columns.Add("VARIANCE_COST", GetType(System.Decimal), "ISNULL(PHYS_VALUE,0) - ISNULL(BOOK_VALUE,0)")
                .Columns("PHYS_CTNS").DataType = GetType(System.Int64)
                .Columns("BOOK_CTNS").DataType = GetType(System.Int64)
            End With

            ASCMAIN1.sql = $"Select * from {WHTPHYCR}"
            Create_TDA(.Tables.Add, "WHTPHYCR", "**", 0, False, "", 4)
            ' Create_TDA(.Tables.Add, "WHTPHYCR", "**", 0, False, "V", 4)

            With .Tables("WHTPHYCR")
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("PHYS_VALUE", GetType(System.Decimal), "PHYS_UNITS * STYLE_COST_FIFO")
                .Columns.Add("BOOK_VALUE", GetType(System.Decimal), "BOOK_UNITS * STYLE_COST_FIFO")
                .Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
                .Columns.Add("VARIANCE_COST", GetType(System.Decimal), "ISNULL(PHYS_VALUE,0) - ISNULL(BOOK_VALUE,0)")
            End With

            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE
                , SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS 
                from (
                Select WHTPHYC3.WHSE_CODE, WHTPHYC3.BAR_CODE, WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE
                , WHTPHYC3.PHYS_UNITS, 0 BOOK_UNITS, WHTPHYC1.LOCATION_CODE
                from WHTPHYC3,WHTPHYC1
                where WHTPHYC3.WHSE_CODE = :PARM1
                and WHTPHYC1.LOCATION_CODE = :PARM2
                and WHTPHYC1.TICKET_STATUS = 'A'
                and (NVL(:PARM3,'*') = '*' or (WHTPHYC3.STYLE_CODE = :PARM4 and WHTPHYC3.COLOR_CODE = :PARM5))
                and WHTPHYC1.WHSE_CODE = WHTPHYC3.WHSE_CODE
                and WHTPHYC1.TICKET_NO = WHTPHYC3.TICKET_NO
                UNION
                Select WHTLOCB0.WHSE_CODE, WHTLOCB0.BAR_CODE, WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE
                , 0 PHYS_UNITS, WHTLOCB0.LOCATION_QTY BOOK_UNITS, WHTLOCB0.LOCATION_CODE
                from WHTLOCB0
                where WHTLOCB0.WHSE_CODE = :PARM1 
                and WHTLOCB0.LOCATION_CODE = :PARM2
                and (NVL(:PARM3,'*') = '*' or (WHTLOCB0.STYLE_CODE = :PARM4 and WHTLOCB0.COLOR_CODE = :PARM5))
                ) group by WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTPHYCB", "**", 0, False, "VVVVV", 5)
            With .Tables("WHTPHYCB")
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("VAR_UNITS", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
            End With


            ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE
                , SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS 
                from (
                Select ROWNUM, WHTPHYC3.WHSE_CODE, WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE
                , WHTPHYC3.PHYS_UNITS, 0 BOOK_UNITS, WHTPHYC1.LOCATION_CODE
                from WHTPHYC3,WHTPHYC1
                where WHTPHYC3.WHSE_CODE = :PARM1
                and WHTPHYC1.LOCATION_CODE = :PARM2
                and WHTPHYC1.TICKET_STATUS = 'A'
                and (NVL(:PARM3,'*') = '*' or (WHTPHYC3.STYLE_CODE = :PARM4 and WHTPHYC3.COLOR_CODE = :PARM5))
                and WHTPHYC1.WHSE_CODE = WHTPHYC3.WHSE_CODE
                and WHTPHYC1.TICKET_NO = WHTPHYC3.TICKET_NO
                UNION
                Select ROWNUM, WHTLOCB0.WHSE_CODE, WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE
                , 0 PHYS_UNITS, WHTLOCB0.LOCATION_QTY BOOK_UNITS, WHTLOCB0.LOCATION_CODE
                from WHTLOCB0
                where WHTLOCB0.WHSE_CODE = :PARM1 
                and WHTLOCB0.LOCATION_CODE = :PARM2
                and (NVL(:PARM3,'*') = '*' or (WHTLOCB0.STYLE_CODE = :PARM4 and WHTLOCB0.COLOR_CODE = :PARM5))
                ) group by WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTPHYCS", "**", 0, False, "VVVVV", 5)
            With .Tables("WHTPHYCS")
                .Columns("BOOK_UNITS").DataType = GetType(System.Int64)
                .Columns("PHYS_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("VAR_UNITS", GetType(System.Int64), "ISNULL(PHYS_UNITS,0) - ISNULL(BOOK_UNITS,0)")
            End With

            ASCMAIN1.sql = "Select STYLE_CODE, COLOR_CODE, SUM (LOC_UNITS) LOC_UNITS, SUM (PER_UNITS) PER_UNITS from (" & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, SUM (NVL(LOCATION_QTY,0)) LOC_UNITS, 0 PER_UNITS from WHTLOCB1 where WHSE_CODE = :PARM1 and LOCATION_QTY <> 0 group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, 0 LOC_UNITS, SUM (NVL(WHSE_QTY_ON_HAND,0)) PER_UNITS from ICTSTAT2 where WHSE_CODE = :PARM1 and WHSE_QTY_ON_HAND <> 0 group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ") group by STYLE_CODE, COLOR_CODE"
            Create_TDA(.Tables.Add, "WHTLOCBX", "**", 0, False, "V", 2)
            With .Tables("WHTLOCBX")
                .Columns("LOC_UNITS").DataType = GetType(System.Int64)
                .Columns("PER_UNITS").DataType = GetType(System.Int64)
                .Columns.Add("DIFFERENCE", GetType(System.Int64), "ISNULL(LOC_UNITS,0) - ISNULL(PER_UNITS,0)")
            End With

            ASCMAIN1.sql = "Select * from WHTLOCB0 where WHSE_CODE = :PARM1 and LOCATION_CODE = :PARM2 and LOCATION_QTY <> 0"
            Create_TDA(.Tables.Add, "WHTLOCB0", "**", 0, False, "VV", 5)

            ASCMAIN1.sql = "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'WHTLOCM1' and COLUMN_NAME = 'LOCATION_USE'"
            Create_TDA(.Tables.Add, "WHTLOCM1_LOCATION_USE", "**", 0, False, "", 1)

            With .Tables.Add("WHTPHYC1_TICKET_TYPE")
                .Columns.Add("T_CODE")
                .Columns.Add("T_DESC")
            End With
            With .Tables("WHTPHYC1_TICKET_TYPE")
                .Rows.Add(New String() {"LS", "Load Scan"})
                .Rows.Add(New String() {"TR", "Trailer PC"})
                .Rows.Add(New String() {"RC", "Receiving PC"})
                .Rows.Add(New String() {"SH", "Pick not Ship"})
                .Rows.Add(New String() {"5S", "5-Star"})
                .Rows.Add(New String() {"1S", "1-Stop"})
                .Rows.Add(New String() {"PC", "Pre-Count"})
            End With


        End With

        Fill_Records("WHTLOCM1")

        grdWHTLOCBA.DataSource = dst.Tables("WHTLOCBA")
        grdWHTPHYC2.DataSource = dst.Tables("WHTPHYC2")
        grdWHTPHYC3.DataSource = dst.Tables("WHTPHYC3")
        grdWHTPHYCI.DataSource = dst.Tables("WHTPHYCI")
        grdWHTPHYCX.DataSource = dst.Tables("WHTPHYCX")
        grdWHTPHYCV.DataSource = dst.Tables("WHTPHYCV")
        grdWHTPHYCB.DataSource = dst.Tables("WHTPHYCB")
        grdWHTPHYCS.DataSource = dst.Tables("WHTPHYCS")
        grdWHTPHYCL.DataSource = dst.Tables("WHTPHYCL")
        grdWHTPHYCR.DataSource = dst.Tables("WHTPHYCR")
        grdWHTLOCBX.DataSource = dst.Tables("WHTLOCBX")

        grdWHTLOCM1_LOCATION_USE.DataSource = dst.Tables("WHTLOCM1_LOCATION_USE")
        Fill_Records("WHTLOCM1_LOCATION_USE")
        Sort_grdColumns(grdWHTLOCM1_LOCATION_USE, "T_CODE")

        grdWHTPHYC1_TICKET_TYPE.DataSource = dst.Tables("WHTPHYC1_TICKET_TYPE")


        Create_Summary(grdWHTLOCBX, "STYLE_CODE", "Count")
        Create_Summary(grdWHTLOCBX, New String() {"LOC_UNITS", "PER_UNITS", "DIFFERENCE"})


        Create_Summary(grdWHTPHYCX, "TICKET_NO", "Count")
        Create_Summary(grdWHTPHYCX, New String() {"PHYS_CTNS", "PHYS_UNITS", "UNIT_VARIANCE", "ABSOLUTE_VARIANCE", "BOOK_CTNS", "BOOK_UNITS", "BOOK_INVTY_ADJ", "EMPTY_LOCATION", "VIRTUAL"})

        Create_Summary(grdWHTPHYCL, "LOCATION_CODE", "Count")
        ' Create_Summary(grdWHTPHYCL, "TICKET_NO", "Count")
        Create_Summary(grdWHTPHYCL, New String() {"PHYS_UNITS", "BOOK_UNITS", "PHYS_VALUE", "BOOK_VALUE", "VARIANCE", "VARIANCE_COST", "PHYS_CTNS", "BOOK_CTNS"})

        Create_Summary(grdWHTPHYCR, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTPHYCR, New String() {"PHYS_UNITS", "BOOK_UNITS", "PHYS_VALUE", "BOOK_VALUE", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdWHTPHYCV, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPHYCV, New String() {"BOOK", "PHYS", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdWHTPHYCB, "BAR_CODE", "Count")
        Create_Summary(grdWHTPHYCB, New String() {"BOOK_UNITS", "PHYS_UNITS", "VAR_UNITS"})

        Create_Summary(grdWHTPHYCS, "STYLE_CODE", "Count")
        Create_Summary(grdWHTPHYCS, New String() {"BOOK_UNITS", "PHYS_UNITS", "VAR_UNITS"})

        Create_Summary(grdWHTPHYC2, "BAR_CODE", "Count")
        Create_Summary(grdWHTPHYC3, "PHYS_UNITS")

        Create_Summary(grdWHTPHYCI, "TICKET_NO", "Count")
        Create_Summary(grdWHTPHYCI, "PHYS_UNITS")

        Create_Summary(grdWHTLOCBA, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCBA, "LOCATION_QTY")

        With grdWHTPHYC2.DisplayLayout.Bands("WHTPHYC2")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "BAR_CODE" Or gcol.Key = "COUNTS_BY_UPC" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightGray

            Next
        End With


        With grdWHTPHYC3.DisplayLayout.Bands("WHTPHYC3")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "STYLE_CODE" Or gcol.Key = "PHYS_UNITS" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightGray

            Next
            '.Columns("TICKET_LNO").Header.Fixed = True
            .Columns("STYLE_CODE").Header.Fixed = True
        End With

        With grdWHTPHYCV.DisplayLayout.Bands("WHTPHYCV")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "VARIANCE" Or gcol.Key = "VARIANCE_COST" Or gcol.Key = "STYLE_COST_VAR" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PHYS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "BOOK" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "COLOR_CODE" Or gcol.Key = "STYLE_DESC" Or gcol.Key = "STYLE_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            .Columns("STYLE_CODE").Header.Fixed = True

            .Columns("STYLE_COST_VAR").Hidden = True
            .Columns("COUNTED_LOCS").Hidden = True
            .Columns("BOOKED_LOCS").Hidden = True

        End With



        With grdWHTPHYCB.DisplayLayout.Bands("WHTPHYCB")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "VAR_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PHYS_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "BOOK_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "COLOR_CODE" Or gcol.Key = "STYLE_DESC" Or gcol.Key = "STYLE_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            '.Columns("STYLE_CODE").Header.Fixed = True
        End With

        With grdWHTPHYCS.DisplayLayout.Bands("WHTPHYCS")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "VAR_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PHYS_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "BOOK_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "COLOR_CODE" Or gcol.Key = "STYLE_DESC" Or gcol.Key = "STYLE_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            '.Columns("STYLE_CODE").Header.Fixed = True
        End With

        With grdWHTPHYCX.DisplayLayout.Bands("WHTPHYCX")
            .Columns("TICKET_NO").Header.Fixed = True
            .Columns("TICKET_TYPE").Header.Fixed = True
            .Columns("LOCATION_CODE").Header.Fixed = True
            .Columns("SELECTED").Header.Fixed = True
            '.Columns("SELECTED"). = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "LOCATION_USE" Or gcol.Key = "TICKET_STATUS" Then
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "SELECTED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                If gcol.Key = "SELECTED" Or gcol.Key = "TICKET_NO" Or gcol.Key = "TICKET_TYPE" Or gcol.Key = "LOCATION_CODE" Or gcol.Key = "LOCATION_USE" Or gcol.Key = "VIRTUAL" Or gcol.Key = "VERIFIED_DATE" Or gcol.Key = "VERIFIED_OPER" Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                ElseIf gcol.Key = "BOOK_CTNS" Or gcol.Key = "BOOK_UNITS" Or gcol.Key = "BOOK_INVTY_ADJ" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "SC" Or gcol.Key = "PHYS_SC_COUNT" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "PHYS_CTNS" Or gcol.Key = "PHYS_UNITS" Or gcol.Key = "EMPTY_LOCATION" Or gcol.Key = "TICKET_STATUS" Or gcol.Key = "INVALIDATED_BY" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key = "UNIT_VARIANCE" Or gcol.Key = "ABSOLUTE_VARIANCE" Or gcol.Key = "LAST_ACTIVITY" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            .Columns("SC").Hidden = (ASCMAIN1.CLIENT = "NYA")
            .Columns("STYLE_CODE").Hidden = Not (ASCMAIN1.CLIENT = "NYA")
        End With
        grdAttention_app.BackColor = Drawing.Color.Yellow
        grdWARNING_app.ForeColor = Drawing.Color.Red

        With grdWHTPHYCI.DisplayLayout.Bands("WHTPHYCI")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Turquoise
            Next
        End With


        With grdWHTLOCBA.DisplayLayout.Bands("WHTLOCBA")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Yellow
            Next
        End With

        With grdWHTPHYCL.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                If gcol.Key = "LOCATION_USE" Then
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                End If

                If gcol.Key = "TICKET_NO" Or gcol.Key = "TICKET_TYPE" Or gcol.Key = "LOCATION_CODE" Or gcol.Key = "LOCATION_USE" Or gcol.Key = "VIRTUAL" Or gcol.Key = "TICKETS" Or gcol.Key = "EMPTY" Or gcol.Key = "PHYS_INIT" Or gcol.Key = "PHYS_LAST" Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                ElseIf gcol.Key.StartsWith("PHYS") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key.StartsWith("BOOK") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key.StartsWith("VAR") Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key.StartsWith("STYLE") Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("LOCATION_CODE").Header.Fixed = True
        End With


        With grdWHTPHYCR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "LOCATION_CODE" Or gcol.Key = "TICKET_NO" Or gcol.Key = "INIT-OPER" Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "COLOR_CODE" Or gcol.Key = "STYLE_DESC" Or gcol.Key = "STYLE_COST_FIFO" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                ElseIf gcol.Key.StartsWith("PHYS") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                ElseIf gcol.Key.StartsWith("BOOK") Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key.StartsWith("V") Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            .Columns("LOCATION_CODE").Header.Fixed = True
        End With


        With grdWHTLOCBX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "DIFFERENCE" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PER_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightSalmon
                ElseIf gcol.Key = "LOC_UNITS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf gcol.Key = "STYLE_CODE" Or gcol.Key = "COLOR_CODE" Then
                    gcol.Header.Appearance.BackColor2 = Color.Gold
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
        End With

        'ASCMAIN1.Add_Value_List(grdWHTPHYCX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        'ASCMAIN1.Add_Value_List(grdWHTPHYCV, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grpHeader.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                'Validate_Code("WHSE_CODE")

                'If Absx1.txtFor("WHSE_CODE").Text = "" Then
                '    EMsg &= vbCr & "You must specify a Valid Warehouse"
                'Else
                '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                '    If IsNothing(rowICTWHSE1) Then
                '        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                '    ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                '        EMsg &= vbCr & "Warehouse Entered Is Not Active"
                '        'ElseIf rowICTWHSE1.Item("LP_CODE") & "" <> "" And Not (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
                '        '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Counts Entry Allowed"
                '    ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                '        EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                '    End If
                'End If

                'If Absx1.txtFor("TICKET_NO").Text = "" Then
                '    EMsg &= vbCr & "You must specify a Ticket"
                'Else
                '    Dim rowWHTPHYC1 As DataRow = LookUp("WHTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                '    If rowWHTPHYC1 IsNot Nothing Then
                '        Click_Command("View")
                '        Exit Sub
                '    Else
                '        If Not ASCMAIN1.Logical_Lock("WHTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & Absx1.txtFor("TICKET_NO").Text) Then
                '            Exit Sub
                '        End If
                '    End If
                'End If



            Case "Edit"

                If ASCMAIN1.USER_ID = "rick" Or ASCMAIN1.USER_ID = "dgj" Or ASCMAIN1.USER_ID = "wjz" Then
                Else
                    Exit Sub
                End If

                If grdWHTPHYCX.ActiveRow Is Nothing Then
                    EMsg &= "Select a Ticket from the grid and then click Edit"
                Else


                    TICKET_NO = grdWHTPHYCX.ActiveRow.Cells("TICKET_NO").Value ' Absx1.txtFor("TICKET_NO").Text
                    LOCATION_CODE = grdWHTPHYCX.ActiveRow.Cells("LOCATION_CODE").Value ' Absx1.txtFor("LOCATION_CODE").Text

                    Dim rowWHTPHYC1 As DataRow = LookUp("WHTPHYC1", New String() {WHSE_CODE, TICKET_NO})
                    If rowWHTPHYC1 Is Nothing Then
                        EMsg &= "Ticket is not on File"
                    Else
                        If rowWHTPHYC1.Item("TICKET_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Ticket Is Not Active"
                            'ElseIf rowWHTPHYC1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                            '    EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                        End If

                        'Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                        'If IsNothing(rowICTWHSE1) Then
                        '    EMsg &= vbCr & "Warehouse {} Is Not Valid"
                        '    'ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                        '    '    EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        '    'ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                        '    '    EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                        'End If

                        If Not ASCMAIN1.Logical_Lock("WHTPHYC1", WHSE_CODE & ":" & TICKET_NO) Then
                            Exit Sub
                        End If
                        If Not ASCMAIN1.Logical_Lock("WHTPHYC1", LOCATION_CODE) Then
                            Exit Sub
                        End If
                    End If

                End If

            Case "View"
                If grdWHTPHYCX.ActiveRow Is Nothing Then
                    EMsg &= "Select a Ticket from the grid and then click View"
                Else
                    TICKET_NO = grdWHTPHYCX.ActiveRow.Cells("TICKET_NO").Value ' Absx1.txtFor("TICKET_NO").Text
                    LOCATION_CODE = grdWHTPHYCX.ActiveRow.Cells("LOCATION_CODE").Value ' Absx1.txtFor("LOCATION_CODE").Text

                    Dim rowWHTPHYC1 As DataRow = LookUp("WHTPHYC1", New String() {WHSE_CODE, TICKET_NO})
                    If rowWHTPHYC1 Is Nothing Then
                        EMsg &= "Ticket is not on File"
                    End If
                End If

            Case "Update"
                'Stop
                'Throw New Exception("not implemented")
                'If Not (optMode.Value = "U" Or optMode.Value = "V") Then
                '    If location_support Then
                '        If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                '            EMsg &= vbCr & "You Must Specify a Location"
                '        Else
                '            Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("LOCATION_CODE").Text})
                '            If rowWHTLOCM1 Is Nothing Then
                '                EMsg &= vbCr & "Invalid Value Specified for Location"
                '            End If
                '        End If

                '    End If

                '    If grdWHTPHYC3.Rows.Count = 0 Then
                '        EMsg &= vbCr & "No Details Entered"
                '    Else

                '    End If
                'End If

            Case "Delete"
                If MessageBox.Show("Are you sure you want to Delete this Entry?", "Confirm Deletion",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Update Variances"
                If Not ASCMAIN1.Running_in_VS Then
                    MessageBox.Show("Action Not Allowed", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                    Exit Sub
                End If
                If WHSE_CODE = "" Then Exit Sub


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                ' Stop
                'If optMode.Value = "U" Or optMode.Value = "V" Then
                '    Verify_Counts()
                'Else
                '    Update_Record()
                'End If

                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)


            Case "By Ticket"
                Print_Counts("T")
            Case "By Location"
                Print_Counts("L")
            Case "By Style"
                Print_Counts("S")

            Case "Refresh Tickets"

                Refresh_Tickets(True)

            Case "Rebuild Variances"
                Fill_Records("WHTPHYCX", WHSE_CODE)
                Show_Tickets()
                Sort_grdColumns(grdWHTPHYCX, "TICKET_NO".ToLower)
                variances_were_rebuilt = False
                Rebuild_Variances()

            Case "Update Variances"
                CloseInventory()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If


                    If ScreenMode And (EntryMode <> "N" And EntryMode <> "E") Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Delete").Visible = (ScreenMode And EntryMode = "E")

                    .Items("Refresh Tickets").Visible = Not ScreenMode
                    .Items("Rebuild Variances").Visible = Not ScreenMode
                    .Items("Update Variances").Visible = (WHSE_CODE <> "" And (ASCMAIN1.Running_in_VS))

                End With

                .Groups("Count Reports").Visible = False ' Not ScreenMode

                .Groups("Tickets").Enabled = Not ScreenMode

                For Each tab As UltraWinTabControl.UltraTab In tab0.Tabs
                    tab.Enabled = Not ScreenMode
                    If tab.Key = "Tickets" Then
                        tab.Enabled = True
                    End If
                Next

                Hide_Control_Panel()
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode And (optMode.Value <> "U" And optMode.Value <> "V")
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode
        UltraExplorerBar1.Groups("Tickets").Visible = Not ScreenMode

        If ScreenMode Then

            If EntryMode = "E" Then
                grdWHTPHYC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdWHTPHYC2.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
            Else
                grdWHTPHYC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdWHTPHYC2.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            End If
        Else
            Clear_Record()
            tab0.SelectedTab = tab0.Tabs("Tickets")

            grdWHTPHYCX.Parent = splWHTPHYCX.Panel1
        End If

        With grdWHTPHYCX.DisplayLayout.Bands("WHTPHYCX")
            .Columns("SELECTED").Hidden = Not ScreenMode
        End With

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"WHTPHYC1", "WHTPHYC2", "WHTPHYC3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        Absx1.txtFor("TICKET_NO").Text = ""

        If WHSE_CODE = "" Then
            Absx1.txtFor("WHSE_CODE").Focus()
        Else
            Absx1.txtFor("TICKET_NO").Focus()
        End If

        variances_were_rebuilt = False
        Refresh_Tickets()


        Setup_tab0()
    End Sub

    Sub Load_Record()

        'Stop
        ' Throw New Exception("why are we here")

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        'WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        'TICKET_NO = Absx1.txtFor("TICKET_NO").Text

        'If EntryMode = "N" Then
        '    rowWHTPHYC1 = dst.Tables("WHTPHYC1").NewRow
        '    rowWHTPHYC1.Item("WHSE_CODE") = WHSE_CODE
        '    rowWHTPHYC1.Item("TICKET_NO") = TICKET_NO ' ASCMAIN1.Next_Control_No("WHTPHYC1.TICKET_NO")

        '    rowWHTPHYC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        '    rowWHTPHYC1.Item("INIT_DATE") = DATETIME_STAMP
        '    rowWHTPHYC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        '    rowWHTPHYC1.Item("LAST_DATE") = DATETIME_STAMP
        '    dst.Tables("WHTPHYC1").Rows.Add(rowWHTPHYC1)
        'Else
        '    Fill_Record("WHTPHYC1", New String() {WHSE_CODE, TICKET_NO})
        '    dst.AcceptChanges()
        'End If

        Fill_Record("WHTPHYC1", New String() {WHSE_CODE, TICKET_NO})

        'Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        'location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        'With grdWHTPHYC3.DisplayLayout.Bands(0)
        '    If ASCMAIN1.CLIENT = "RGI" Then
        '        .Columns("BAR_CODE").Hidden = True
        '    Else
        '        .Columns("BAR_CODE").Hidden = Not location_support
        '    End If
        'End With

        Fill_Records("WHTPHYC2", New String() {WHSE_CODE, TICKET_NO})
        Fill_Records("WHTPHYC3", New String() {WHSE_CODE, TICKET_NO})
        Setup_WHTPHYC3("", "")

        Fill_Records("WHTLOCB0", New String() {WHSE_CODE, LOCATION_CODE})

        BAR_CODE_PFX = ""

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        Dim OLD_TICKET_NO As String = TICKET_NO
        ASCMAIN1.sql = "Update WHTPHYC1 set TICKET_STATUS = 'V' Where WHSE_CODE = :PARM1 and TICKET_NO = :PARM2 and TICKET_STATUS in ('A','X')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {WHSE_CODE, OLD_TICKET_NO})

        TICKET_NO = ASCMAIN1.Next_Control_No("WHTPHYC1.TICKET_NO_" & WHSE_CODE)
        Dim rowWHTPHYC1 As DataRow = dst.Tables("WHTPHYC1").Rows.Find(New String() {WHSE_CODE, OLD_TICKET_NO})
        rowWHTPHYC1.Item("TICKET_NO") = TICKET_NO

        If dst.Tables("WHTPHYC2").Rows.Count > 0 Then

            For Each rowWHTPHYC2 As DataRow In dst.Tables("WHTPHYC2").Select("")
                Dim BAR_CODE As String = rowWHTPHYC2.Item("BAR_CODE")
                rowWHTPHYC2.Item("TICKET_NO") = TICKET_NO

                'Check for BAR_CODE in other active tickets
                'and invalidate original ticket

                ASCMAIN1.sql = "Select WHTPHYC1.WHSE_CODE, WHTPHYC1.TICKET_NO from WHTPHYC1, WHTPHYC2" & vbCrLf _
                    & " where WHTPHYC1.WHSE_CODE = WHTPHYC2.WHSE_CODE" & vbCrLf _
                    & "   and WHTPHYC1.TICKET_NO =  WHTPHYC2.TICKET_NO" & vbCrLf _
                    & " and WHTPHYC1.TICKET_STATUS = 'A'" & vbCrLf _
                    & $" and WHTPHYC2.BAR_CODE = '{BAR_CODE}'" & vbCrLf _
                    & $" and WHTPHYC2.TICKET_NO <> '{OLD_TICKET_NO}'" & vbCrLf

                ASCMAIN1.sql = $"Update WHTPHYC1 Set TICKET_STATUS = 'X', INVALIDATED_BY = '{TICKET_NO}'" & vbCrLf &
                     $"   where (WHSE_CODE, TICKET_NO) in ({ASCMAIN1.sql})"

                ASCDATA1.ExecuteSQL()

                ' Get_SC_for_BAR_CODE(BAR_CODE)
            Next

            For Each rowWHTPHYC3 As DataRow In dst.Tables("WHTPHYC3").Select("")
                rowWHTPHYC3.Item("TICKET_NO") = TICKET_NO
            Next
        End If


        Dim PHYS_CTNS As Int16 = dst.Tables("WHTPHYC2").Rows.Count
        Dim PHYS_UNITS As Int16 = Val(dst.Tables("WHTPHYC3").Compute("SUM(PHYS_UNITS)", "") & "")
        Dim PHYS_SC_COUNT As Int16 = ASCDATA1.SelectDistinct(dst.Tables("WHTPHYC3"), New String() {"SC"}).Rows.Count

        Dim STYLE_CODE_MIN As String = dst.Tables("WHTPHYC3").Compute("MIN (STYLE_CODE)", "")
        Dim STYLE_CODE_MAX As String = dst.Tables("WHTPHYC3").Compute("MAX (STYLE_CODE)", "")
        Dim SC_MIN As String = dst.Tables("WHTPHYC3").Compute("MIN (SC)", "")
        Dim SC_MAX As String = dst.Tables("WHTPHYC3").Compute("MAX (SC)", "")

        rowWHTPHYC1.Item("PHYS_CTNS") = PHYS_CTNS
        rowWHTPHYC1.Item("PHYS_UNITS") = PHYS_UNITS
        rowWHTPHYC1.Item("PHYS_SC_COUNT") = PHYS_SC_COUNT

        rowWHTPHYC1.Item("STYLE_CODE_MIN") = STYLE_CODE_MIN
        rowWHTPHYC1.Item("STYLE_CODE_MAX") = STYLE_CODE_MAX
        rowWHTPHYC1.Item("SC_MIN") = SC_MIN
        rowWHTPHYC1.Item("SC_MAX") = SC_MAX

        rowWHTPHYC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTPHYC1.Item("INIT_DATE") = DATETIME_STAMP
        rowWHTPHYC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWHTPHYC1.Item("LAST_DATE") = DATETIME_STAMP

        Update_Record_TDA("WHTPHYC1", $"WHSE_CODE = '{WHSE_CODE}' AND TICKET_NO = '{TICKET_NO}'")
        Update_Record_TDA("WHTPHYC2", $"WHSE_CODE = '{WHSE_CODE}' AND TICKET_NO = '{TICKET_NO}'")
        Update_Record_TDA("WHTPHYC3", $"WHSE_CODE = '{WHSE_CODE}' AND TICKET_NO = '{TICKET_NO}'")
        CommitTrans("Update Complete")


        Dim rowWHTPHYCX As DataRow = dst.Tables("WHTPHYCX").Rows.Find(New String() {WHSE_CODE, OLD_TICKET_NO})
        rowWHTPHYCX.Item("TICKET_STATUS") = "V"
        Dim rowWHTPHYCX_new As DataRow = dst.Tables("WHTPHYCX").NewRow
        rowWHTPHYCX_new.ItemArray = rowWHTPHYCX.ItemArray
        rowWHTPHYCX_new.Item("TICKET_NO") = TICKET_NO
        rowWHTPHYCX_new.Item("TICKET_STATUS") = "A"
        rowWHTPHYCX_new.Item("INIT_OPER") = rowWHTPHYC1.Item("INIT_OPER")
        rowWHTPHYCX_new.Item("LAST_OPER") = rowWHTPHYC1.Item("LAST_OPER")
        rowWHTPHYCX_new.Item("INIT_DATE") = rowWHTPHYC1.Item("INIT_DATE")
        rowWHTPHYCX_new.Item("LAST_DATE") = rowWHTPHYC1.Item("LAST_DATE")
        rowWHTPHYCX_new.Item("EMPTY_LOCATION") = IIf(dst.Tables("WHTPHYC2").Rows.Count = 0, "1", "0")
        rowWHTPHYCX_new.Item("PHYS_CTNS") = rowWHTPHYC1.Item("PHYS_CTNS")
        rowWHTPHYCX_new.Item("PHYS_UNITS") = rowWHTPHYC1.Item("PHYS_UNITS")
        rowWHTPHYCX_new.Item("PHYS_SC_COUNT") = rowWHTPHYC1.Item("PHYS_SC_COUNT")
        rowWHTPHYCX_new.Item("STYLE_CODE_MIN") = rowWHTPHYC1.Item("STYLE_CODE_MIN")
        rowWHTPHYCX_new.Item("STYLE_CODE_MAX") = rowWHTPHYC1.Item("STYLE_CODE_MAX")
        rowWHTPHYCX_new.Item("SC_MIN") = rowWHTPHYC1.Item("SC_MIN")
        rowWHTPHYCX_new.Item("SC_MAX") = rowWHTPHYC1.Item("SC_MAX")
        'rowWHTPHYCX_new.Item("STYLE_CODE") = rowWHTPHYC1.Item("STYLE_CODE")
        rowWHTPHYCX_new.Item("SC") = rowWHTPHYC1.Item("SC_MIN")
        dst.Tables("WHTPHYCX").Rows.Add(rowWHTPHYCX_new)

        Dim rowWHTPHYCL As DataRow = dst.Tables("WHTPHYCL").Select($"LOCATION_CODE = '{LOCATION_CODE}'")(0)


        rowWHTPHYCL.Item("TICKET_NO") = TICKET_NO
        rowWHTPHYCL.Item("TICKETS") = Val(rowWHTPHYCL.Item("TICKETS") & "") + 1

        rowWHTPHYCL.Item("EMPTY") = IIf(dst.Tables("WHTPHYC2").Rows.Count = 0, "1", "0")
        rowWHTPHYCL.Item("PHYS_CTNS") = rowWHTPHYC1.Item("PHYS_CTNS")
        rowWHTPHYCL.Item("PHYS_UNITS") = rowWHTPHYC1.Item("PHYS_UNITS")

        rowWHTPHYCL.Item("PHYS_STYLE_COLORS") = rowWHTPHYC1.Item("PHYS_SC_COUNT")
        rowWHTPHYCL.Item("PHYS_SCMIN") = rowWHTPHYC1.Item("SC_MIN")
        rowWHTPHYCL.Item("PHYS_SCMAX") = rowWHTPHYC1.Item("SC_MAX")

    End Sub

    Sub Verify_Counts()
        BeginTrans()

        ASCMAIN1.sql = "Update WHTPHYC1 " & vbCrLf _
            & " Set WHTPHYC1.VERIFIED_OPER = '" & ASCMAIN1.USER_ID & "', WHTPHYC1.VERIFIED_DATE = SYSDATE " & vbCrLf _
            & " Where WHTPHYC1.WHSE_CODE = '" & WHSE_CODE & "' and WHTPHYC1.TICKET_NO = :PARM1"

        For Each row As DataRow In dst.Tables("WHTPHYCX").Select("SELECTED = '1'")
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", row.Item("TICKET_NO"))
        Next

        CommitTrans("Update Complete, Records Verified")
    End Sub

    Sub Delete_Record()
        Stop
        Throw New Exception("NOT SUPPORTED TO DELETE COUNT")
        BeginTrans()

        'Delete_Records("WHTPHYC1")
        'Delete_Records("WHTPHYC2")
        'Delete_Records("WHTPHYC3")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where WHSE_CODE = '" & WHSE_CODE & "' and TICKET_NO = '" & TICKET_NO & "'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTPHYCX, "SSB", "Show Filter", "Show GroupBox", "Recount Location")
        Load_Popup_Menu(grdWHTPHYC3, "BS", "Style Status Inquiry", "Show Old Counts")
        Load_Popup_Menu(grdWHTPHYCV, "SSB", "Show Filter", "Show GroupBox", "Location Inquiry", "Style Status Inquiry")
        Load_Popup_Menu(grdWHTPHYCL, "SSBB", "Show Filter", "Show GroupBox", "Location Inquiry", "Loc/Style/Color")
        Load_Popup_Menu(grdWHTPHYCR, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdWHTLOCBA, "BB", "Location Inquiry", "Show 0's")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Try

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

            Select Case grd.Name
                Case "grdWHTPHYCX"
                    tlb_btn = DirectCast(tlb_pop.Tools("Recount Location"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E")

            End Select
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
                'e.Cancel = True
            Else
                Select Case e.SourceControl.Name

                    Case "grdWHTLOCBA"
                        '  tlb_sbt = DirectCast(tlb.Tools("Show 0s"), UltraWinToolbars.StateButtonTool)

                End Select

            End If

        Catch ex As Exception
            e.Cancel = True
        End Try

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)
            Case "Show 0's"
                '  tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Setup_WHTLOCBA(True)

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Location Inquiry"
                If grd.Name = "grdWHTPHYCV" Then
                    Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                    Context_Launch("Select", "S:" & STYLE_CODE, e.Tool.Key, "WHFLOCS1")

                ElseIf grd.Name = "grdWHTPHYCL" Then
                    Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Text
                    Context_Launch("Select", "L:" & LOCATION_CODE, e.Tool.Key, "WHFLOCS1")
                End If

            Case "Show Old Counts"
                Set_WHTPHYC3_Filter()

            Case "Loc/Style/Color"
                Dim LOCATION_CODE As String = grd.ActiveRow.Cells("LOCATION_CODE").Text
                If LOCATION_CODE <> "" Then
                    With grdWHTPHYCR.DisplayLayout.Bands(0)
                        .ColumnFilters.ClearAllFilters()
                        .ColumnFilters("LOCATION_CODE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, LOCATION_CODE)
                    End With
                End If
                Show_Filter(grdWHTPHYCR)
                tab0.SelectedTab = tab0.Tabs("Location/Style/Color")

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Refresh_Tickets()
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "TICKET_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        Dim row As DataRow = LookUp("WHTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("WHSE_CODE").Text})
                        If row IsNot Nothing Then
                            Click_Command("View", e)
                        Else
                            Click_Command("New", e)
                        End If
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Refresh_Tickets()
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "TICKET_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                Refresh_Tickets()
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdWHTPHYC3"

    'Private Sub grdWHTPHYC3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPHYC3.AfterCellUpdate
    '    Select Case e.Cell.Column.Key
    '        Case "STYLE_CODE"

    '            grdCodeDesc(grdWHTPHYC3, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
    '            If cdr IsNot Nothing Then
    '                Dim STYLE_CODE As String = e.Cell.Value
    '                'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
    '                Dim COLOR_CODE As String = ""

    '                ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
    '                Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
    '                If rowICTSTYC1s.Length = 1 Then
    '                    COLOR_CODE = rowICTSTYC1s(0).Item("COLOR_CODE")
    '                    e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODE
    '                End If

    '                Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)
    '            Else
    '                grdWHTPHYC3.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
    '            End If

    '        Case "COLOR_CODE"
    '            'grdCodeDesc(grdWHTPHYC3, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
    '            '' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
    '            'If cdr IsNot Nothing Then
    '            '    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
    '            'End If
    '            Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value
    '            Dim COLOR_CODE As String = e.Cell.Value
    '            Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)

    '        Case "COUNT_CTNS"

    '    End Select
    'End Sub

    'Private Sub grdWHTPHYC3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC3.AfterExitEditMode
    '    'Select Case grdWHTPHYC3.ActiveCell.Column.Key

    '    'End Select
    'End Sub

    Private Sub grdWHTPHYC2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC2.AfterRowActivate
        With grdWHTPHYC2.DisplayLayout.Bands(0)
            If grdWHTPHYC2.ActiveRow.IsAddRow Then
                .Columns("BAR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdWHTPHYC2.ActiveCell = grdWHTPHYC2.ActiveRow.Cells("BAR_CODE")
                grdWHTPHYC2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("BAR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        BAR_CODE = grdWHTPHYC2.ActiveRow.Cells("BAR_CODE").Value & ""
        Dim COUNTS_BY_UPC As String = grdWHTPHYC2.ActiveRow.Cells("COUNTS_BY_UPC").Value & ""
        Setup_WHTPHYC2(BAR_CODE, COUNTS_BY_UPC)
    End Sub

    Private Sub grdWHTPHYC2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC2.AfterRowsDeleted
        For Each ROW As DataRow In dst.Tables("WHTPHYC3").Select($"BAR_CODE = '{BAR_CODE}'")
            ROW.Delete()
        Next
    End Sub

    Private Sub grdWHTPHYC2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTPHYC2.AfterRowUpdate
        BAR_CODE = grdWHTPHYC2.ActiveRow.Cells("BAR_CODE").Value & ""
        Dim COUNTS_BY_UPC As String = grdWHTPHYC2.ActiveRow.Cells("COUNTS_BY_UPC").Value & ""
        Setup_WHTPHYC2(BAR_CODE, COUNTS_BY_UPC)

        BAR_CODE_PFX = Mid(BAR_CODE, 1, 1)
    End Sub

    Private Sub grdWHTPHYC2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTPHYC2.BeforeExitEditMode
        If grdWHTPHYC2.ActiveCell Is Nothing Then Exit Sub
        With grdWHTPHYC2.ActiveCell
            Select Case .Column.Key
                Case "BAR_CODE"
                    If location_support Then
                        If .Text <> "" Then

                            If .Value IsNot Nothing Then
                                If Len(.Text) <> 8 Then
                                    .Value = (BAR_CODE_PFX & .Text.PadLeft(7, "0")).ToUpper
                                Else
                                    .Value = .Text.ToUpper
                                End If

                            End If
                        End If
                        BAR_CODE = .Value & ""
                        If BAR_CODE <> "" Then
                            If dst.Tables("WHTPHYC2").Rows.Find(New String() {WHSE_CODE, TICKET_NO, BAR_CODE}) IsNot Nothing Then
                                .Value = ""
                                e.Cancel = True
                            Else
                                cdr = LookUp("WHTBARC1", BAR_CODE)
                                If cdr Is Nothing Then
                                    ASCMAIN1.Progress($"Invalid Bar Code ({BAR_CODE})")
                                    If .Value IsNot Nothing Then
                                        .Value = ""
                                    End If
                                    e.Cancel = True
                                End If
                            End If
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdWHTPHYC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTPHYC2.BeforeRowUpdate
        With grdWHTPHYC2

            BAR_CODE = e.Row.Cells("BAR_CODE").Text

            If BAR_CODE = "" Then
                '                MsgBox("Missing Value for Style Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("WHTBARC1", BAR_CODE)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Bar Code (" & BAR_CODE & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If


            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("TICKET_NO").Text = "" Then
                    .ActiveRow.Cells("WHSE_CODE").Value = WHSE_CODE
                    .ActiveRow.Cells("TICKET_NO").Value = TICKET_NO

                    Add_WHTPHYC3(BAR_CODE)
                End If
            End If
        End With
    End Sub


    Private Sub grdWHTPHYC2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdWHTPHYC2.Error
        grdWHTPHYC3.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdWHTPHYC3"

    'Private Sub grdWHTPHYC3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPHYC3.AfterCellUpdate
    '    Select Case e.Cell.Column.Key
    '        Case "STYLE_CODE"

    '            grdCodeDesc(grdWHTPHYC3, "ICTSTYL1", "STYLE_CODE", "STYLE_DESC")
    '            If cdr IsNot Nothing Then
    '                Dim STYLE_CODE As String = e.Cell.Value
    '                'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
    '                Dim COLOR_CODE As String = ""

    '                ASCMAIN1.sql = "Select * from ICTSTYC1 where STYLE_CODE = '" & STYLE_CODE & "'"
    '                Dim rowICTSTYC1s() As DataRow = ASCDATA1.GetDataTable.Select("")
    '                If rowICTSTYC1s.Length = 1 Then
    '                    COLOR_CODE = rowICTSTYC1s(0).Item("COLOR_CODE")
    '                    e.Cell.Row.Cells("COLOR_CODE").Value = COLOR_CODE
    '                End If

    '                Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)
    '            Else
    '                grdWHTPHYC3.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
    '            End If

    '        Case "COLOR_CODE"
    '            'grdCodeDesc(grdWHTPHYC3, "ICTCOLR1", "COLOR_CODE", "COLOR_DESC")
    '            '' FOR SOME REASON THE ABOVE CALL IS NOT LOADING THE COLOR_DESC
    '            'If cdr IsNot Nothing Then
    '            '    e.Cell.Row.Cells("COLOR_DESC").Value = cdr.Item("COLOR_DESC")
    '            'End If
    '            Dim STYLE_CODE As String = e.Cell.Row.Cells("STYLE_CODE").Value
    '            Dim COLOR_CODE As String = e.Cell.Value
    '            Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)

    '        Case "COUNT_CTNS"

    '    End Select
    'End Sub

    'Private Sub grdWHTPHYC3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC3.AfterExitEditMode
    '    'Select Case grdWHTPHYC3.ActiveCell.Column.Key

    '    'End Select
    'End Sub

    Private Sub grdWHTPHYC3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC3.AfterRowActivate
        With grdWHTPHYC3.DisplayLayout.Bands(0)
            If grdWHTPHYC3.ActiveRow.IsAddRow Then
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdWHTPHYC3.ActiveCell = grdWHTPHYC3.ActiveRow.Cells("STYLE_CODE")
                grdWHTPHYC3.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Dim STYLE_CODE As String = grdWHTPHYC3.ActiveRow.Cells("STYLE_CODE").Value & ""
        Dim COLOR_CODE As String = grdWHTPHYC3.ActiveRow.Cells("COLOR_CODE").Value & ""
        Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)
        'If EntryMode = "V" Then
        '    Show_Variances()
        'End If
    End Sub

    'Private Sub grdWHTPHYC3_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdWHTPHYC3.AfterRowsDeleted
    '    DisplayTotals()
    'End Sub

    'Private Sub grdWHTPHYC3_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTPHYC3.AfterRowUpdate
    '    DisplayTotals()
    'End Sub

    'Private Sub grdWHTPHYC3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTPHYC3.BeforeExitEditMode
    '    If grdWHTPHYC3.ActiveCell Is Nothing Then Exit Sub
    '    With grdWHTPHYC3.ActiveCell
    '        Select Case .Column.Key
    '            Case "STYLE_CODE"
    '                If .Text <> "" Then
    '                    If .Value IsNot Nothing Then
    '                        .Value = .Text.ToUpper
    '                    End If

    '                End If
    '                If .Text <> "" Then
    '                    cdr = LookUp("ICTSTYL1", .Text)
    '                    If cdr Is Nothing Then
    '                        ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
    '                        If .Value IsNot Nothing Then
    '                            .Value = ""
    '                        End If
    '                        e.Cancel = True
    '                    End If
    '                End If
    '                'Case "BAR_CODE"
    '                '    If location_support Then
    '                '        If .Text <> "" Then
    '                '            If .Value IsNot Nothing Then
    '                '                .Value = .Text.ToUpper
    '                '            End If

    '                '        End If
    '                '        If .Text <> "" Then
    '                '            cdr = LookUp("WHTBARC1", .Text)
    '                '            If cdr Is Nothing Then
    '                '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
    '                '                If .Value IsNot Nothing Then
    '                '                    .Value = ""
    '                '                End If
    '                '                e.Cancel = True
    '                '            End If
    '                '        End If
    '                '    End If
    '        End Select
    '    End With
    'End Sub

    'Private Sub grdWHTPHYC3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTPHYC3.BeforeRowUpdate
    '    With grdWHTPHYC3
    '        If e.Row.Cells("STYLE_CODE").Text = "" Then
    '            '                MsgBox("Missing Value for Style Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
    '            e.Cancel = True
    '        Else
    '            LookUp("ICTSTYL1", e.Row.Cells("STYLE_CODE").Text)
    '            If cdr Is Nothing Then
    '                MsgBox("Invalid Value entered for Style Code (" & e.Row.Cells("STYLE_CODE").Text & ")",
    '                       MsgBoxStyle.OkOnly, "Cannot Update Row")
    '                e.Cancel = True
    '            End If
    '        End If

    '        If location_support Then
    '            'If e.Row.Cells("BAR_CODE").Text = "" Then
    '            '    e.Cancel = True
    '            'Else
    '            '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
    '            '    If cdr Is Nothing Then
    '            '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
    '            '               MsgBoxStyle.OkOnly, "Cannot Update Row")
    '            '        e.Cancel = True
    '            '    End If
    '            'End If

    '        End If

    '        If Val(e.Row.Cells("COUNT_CTNS").Value & "") = 0 And Val(e.Row.Cells("PHYS_UNITS").Value & "") = 0 Then
    '            'MsgBox("Invalid Value entered for Count", MsgBoxStyle.OkOnly, "Cannot Update Row")
    '            e.Cancel = True
    '        End If

    '        If e.Cancel Then
    '            e.Row.CancelUpdate()
    '        End If

    '        If Not e.Cancel Then
    '            If e.Row.Cells("TICKET_NO").Text = "" Then
    '                .ActiveRow.Cells("WHSE_CODE").Value = WHSE_CODE
    '                .ActiveRow.Cells("TICKET_NO").Value = Absx1.CtlFor("TICKET_NO").Text
    '                .ActiveRow.Cells("TICKET_LNO").Value = Val(dst.Tables("WHTPHYC3").Compute("Max(TICKET_LNO)", "") & "") + 1
    '            End If
    '        End If
    '    End With
    'End Sub

    'Private Sub grdWHTPHYC3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTPHYC3.ClickCellButton

    '    If grdWHTPHYC3.ActiveRow Is Nothing Then Exit Sub

    '    Dim sql_where As String = ""
    '    Select Case e.Cell.Column.Key
    '        Case "STYLE_CODE"
    '            'Case "LOCATION_CODE"
    '            '    sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
    '    End Select
    '    grdClickCellButton(grdWHTPHYC3, sql_where, False)

    'End Sub

    'Private Sub grdWHTPHYC3_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdWHTPHYC3.Error
    '    grdWHTPHYC3.ActiveRow.CancelUpdate()
    'End Sub

#End Region

    Sub DisplayTotals()
        'Dim TOTAL_COSTS As Decimal = Val(dst.Tables("WHTPHYC3").Compute("SUM(LINE_COSTS)", "") & "")
        'Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

#Region "grdWHTPHYCX"


    Private Sub grdWHTPHYCX_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdWHTPHYCX.BeforeExitEditMode
        If grdWHTPHYCX.ActiveCell Is Nothing Then Exit Sub
        If Not grdWHTPHYCX.ActiveRow.IsDataRow Then Exit Sub
        With grdWHTPHYCX.ActiveCell
            Select Case .Column.Key
                Case "SELECTED"
                    If .Text = "0" And grdWHTPHYCX.ActiveRow.Cells("VERIFIED_OPER").Value & "" <> "" Then
                        e.Cancel = True
                    End If
                Case Else
                    e.Cancel = True
            End Select
        End With
    End Sub


    Private Sub grdWHTPHYCX_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdWHTPHYCX.BeforeRowUpdate
        'Null()
    End Sub

    Private Sub grdWHTPHYCX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTPHYCX.AfterRowActivate
        If Not grdWHTPHYCX.ActiveRow.IsDataRow Then Exit Sub
        Set_WHTPHYCX()

        If tab0.SelectedTab.Key <> "Tickets" Then Exit Sub

        Set_WHTPHYCB()
    End Sub

    Sub Set_WHTPHYCX()
        If EntryMode = "V" Or EntryMode = "" Then Exit Sub
        If grdWHTPHYCX.ActiveRow IsNot Nothing AndAlso grdWHTPHYCX.ActiveRow.IsDataRow Then
            If Not ASCMAIN1.Logical_Lock("WHTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & grdWHTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text) Then
                Exit Sub
            End If
            Dim rowTICKET_NO As String = grdWHTPHYCX.ActiveRow.Cells("TICKET_NO").Text
            Fill_Records("WHTPHYC3", New String() {WHSE_CODE, rowTICKET_NO})
            Set_WHTPHYC3_Filter()
            grdWHTPHYC3.Text = "Details for Ticket " & rowTICKET_NO & ", Location " & grdWHTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text
        End If
    End Sub

    Sub Set_WHTPHYC3_Filter()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Old Counts"), UltraWinToolbars.StateButtonTool)
        Dim dvw As DataView = DirectCast(grdWHTPHYC3.DataSource, DataTable).DefaultView
        If Not tlb_sbt.Checked Then
            dvw.RowFilter = "STATUS IS NULL"
        Else
            dvw.RowFilter = ""
        End If
    End Sub
    Private Sub grdWHTPHYCX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWHTPHYCX.InitializeRow
        If e.Row.IsDataRow Then
            If Not (IsDBNull(e.Row.Cells("LAST_DATE").Value)) AndAlso e.Row.Cells("LAST_ACTIVITY").Value > e.Row.Cells("LAST_DATE").Value Then
                e.Row.Cells("LOCATION_CODE").Appearance = grdWARNING_app
                e.Row.Cells("LOCATION_CODE").ToolTipText = "Location had activity after the count."
                e.Row.Cells("LAST_ACTIVITY").Appearance = grdWARNING_app
                e.Row.Cells("LAST_ACTIVITY").ToolTipText = "Location had activity after the count."
            End If
        End If
    End Sub
    Private Sub grdWHTPHYCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTPHYCX.DoubleClickRow

        If EntryMode = "E" Then Exit Sub

        If e.Row.IsDataRow Then
            Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Text
            'If optMode.Value = "U" Or optMode.Value = "V" Then
            '    Click_Command("Edit")
            'Else
            Click_Command("View")
            'End If
        End If
    End Sub

#End Region
    Private Sub grdWHTPHYCV_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTPHYCV.AfterRowActivate
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        If grdWHTPHYCV.ActiveRow IsNot Nothing AndAlso grdWHTPHYCV.ActiveRow.IsDataRow Then
            STYLE_CODE = grdWHTPHYCV.ActiveRow.Cells("STYLE_CODE").Value
            COLOR_CODE = grdWHTPHYCV.ActiveRow.Cells("COLOR_CODE").Value
        End If
        Setup_WHTPHYC3(STYLE_CODE, COLOR_CODE)
    End Sub

    Private Sub optMode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMode.ValueChanged
        Show_Tickets()
    End Sub

    Private Sub Show_Tickets()
        If SELECTION_NO = 0 Then Exit Sub
        If EntryMode = "E" Then Exit Sub

        Dim dvw As DataView = DirectCast(grdWHTPHYCX.DataSource, DataTable).DefaultView
        If optMode.Value = "A" Then
            dvw.RowFilter = "TICKET_STATUS = 'A'"
            grdWHTPHYCX.Text = "All Active Tickets for Warehouse " & WHSE_CODE
        ElseIf optMode.Value = "*" Then
            dvw.RowFilter = ""
            grdWHTPHYCX.Text = "All Tickets (Active & Voided & Invalidated) for Warehouse " & WHSE_CODE & " (Totals include Voided & Invalidated Tickets)"
            'Else ' view Dirty tickets
            '    dvw.RowFilter = "LAST_ACTIVITY > LAST_DATE"
            '    grdWHTPHYCX.Text = "Active since Physical Counts for Warehouse " & WHSE_CODE
        End If
    End Sub

    Private Sub optVariances_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVariances.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Variances()
    End Sub

    Sub Show_Variances()
        Dim dvw As DataView = DirectCast(grdWHTPHYCV.DataSource, DataTable).DefaultView
        If optVariances.Value = "A" Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "VARIANCE <> 0"
        End If
    End Sub

    Sub Refresh_Tickets(Optional force_refresh As Boolean = False)

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        If WHSE_CODE = Me.WHSE_CODE And Not force_refresh Or WHSE_CODE = "" Then
            Exit Sub
        End If
        Me.WHSE_CODE = WHSE_CODE

        UltraExplorerBar1.Groups("Screen Control").Items("Update Variances").Visible = (WHSE_CODE <> "" And (ASCMAIN1.Running_in_VS))
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor

        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")

        Fill_Records("WHTPHYCX", WHSE_CODE)
        Show_Tickets()
        Sort_grdColumns(grdWHTPHYCX, "TICKET_NO".ToLower)
        tab0.SelectedTab = tab0.Tabs("Tickets")

        variances_were_rebuilt = False
        Rebuild_Variances()
        Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), True)

    End Sub

    Sub Rebuild_Variances(Optional initialize As Boolean = False)


        If initialize Then

            sqlWHTPHYCL = "Select WHTLOCM1.WHSE_CODE, WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & ", WHTLOCM1.LOCATION_USE, CASE WHEN WHTLOCM1.LOCATION_USE IN('A','L','E') THEN '0' ELSE '1' END VIRTUAL" & vbCrLf _
                & ", A.TICKETS, A.PHYS_INIT, A.PHYS_LAST, A.TICKET_NO, A.EMPTY" & vbCrLf _
                & ", WHTPHYC1.TICKET_TYPE, WHTPHYC1.INIT_DATE, WHTPHYC1.INIT_OPER" & vbCrLf _
                & ", B.PHYS_CTNS, B.PHYS_UNITS, B.PHYS_VALUE, B.PHYS_STYLE_COLORS, B.PHYS_SCMIN, B.PHYS_SCMAX" & vbCrLf _
                & ", C.BOOK_CTNS, C.BOOK_UNITS, C.BOOK_VALUE, C.BOOK_STYLE_COLORS, C.BOOK_SCMIN, C.BOOK_SCMAX" & vbCrLf _
                & " from WHTLOCM1,WHTPHYC1" & vbCrLf _
                 & vbCrLf _
                & ", (Select WHTPHYC1.WHSE_CODE, WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", COUNT (*) TICKETS, MIN (INIT_DATE) PHYS_INIT, MAX (INIT_DATE) PHYS_LAST" & vbCrLf _
                & ", MAX (CASE WHEN WHTPHYC1.TICKET_STATUS = 'A' THEN WHTPHYC1.TICKET_NO ELSE NULL END) TICKET_NO" & vbCrLf _
                & ", MAX (CASE WHEN WHTPHYC1.TICKET_STATUS = 'A' AND NVL(WHTPHYC1.EMPTY_LOCATION,'0') = '1' THEN '1' ELSE NULL END) EMPTY" & vbCrLf _
                & " from WHTPHYC1" & vbCrLf _
                & " where WHTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & " group by WHTPHYC1.WHSE_CODE, WHTPHYC1.LOCATION_CODE) A" & vbCrLf _
                & vbCrLf _
                & ", (Select WHTPHYC1.WHSE_CODE, WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", WHTPHYC1.PHYS_CTNS" & vbCrLf _
                & ", SUM (WHTPHYC3.PHYS_UNITS) PHYS_UNITS" & vbCrLf _
                & ", SUM ((WHTPHYC3.PHYS_UNITS) * ICTSTYC1.STYLE_COST_FIFO) PHYS_VALUE" & vbCrLf _
                & ", COUNT (DISTINCT WHTPHYC3.STYLE_CODE || WHTPHYC3.COLOR_CODE) PHYS_STYLE_COLORS" & vbCrLf _
                & ", MIN (WHTPHYC3.STYLE_CODE || '-' || WHTPHYC3.COLOR_CODE) PHYS_SCMIN" & vbCrLf _
                & ", MAX (WHTPHYC3.STYLE_CODE || '-' || WHTPHYC3.COLOR_CODE) PHYS_SCMAX" & vbCrLf _
                & " from WHTPHYC1,WHTPHYC3,ICTSTYC1" & vbCrLf _
                & " where WHTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and WHTPHYC3.WHSE_CODE = WHTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC3.TICKET_NO = WHTPHYC1.TICKET_NO" & vbCrLf _
                & "   and WHTPHYC1.TICKET_STATUS = 'A'" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = WHTPHYC3.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = WHTPHYC3.COLOR_CODE" & vbCrLf _
                & " group by WHTPHYC1.WHSE_CODE, WHTPHYC1.LOCATION_CODE, WHTPHYC1.PHYS_CTNS) B" & vbCrLf _
                & vbCrLf _
                & ", (Select WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", COUNT (DISTINCT WHTLOCB0.BAR_CODE) BOOK_CTNS" & vbCrLf _
                & ", SUM (WHTLOCB0.LOCATION_QTY + WHTLOCB0.BOOK_INVTY_ADJ) BOOK_UNITS" & vbCrLf _
                & ", SUM ((WHTLOCB0.LOCATION_QTY + WHTLOCB0.BOOK_INVTY_ADJ) * ICTSTYC1.STYLE_COST_FIFO) BOOK_VALUE" & vbCrLf _
                & ", COUNT (DISTINCT WHTLOCB0.STYLE_CODE || WHTLOCB0.COLOR_CODE) BOOK_STYLE_COLORS" & vbCrLf _
                & ", MIN (WHTLOCB0.STYLE_CODE || '-' || WHTLOCB0.COLOR_CODE) BOOK_SCMIN" & vbCrLf _
                & ", MAX (WHTLOCB0.STYLE_CODE || '-' || WHTLOCB0.COLOR_CODE) BOOK_SCMAX" & vbCrLf _
                & " from WHTLOCB0,ICTSTYC1" & vbCrLf _
                & " where WHTLOCB0.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and ICTSTYC1.STYLE_CODE = WHTLOCB0.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = WHTLOCB0.COLOR_CODE" & vbCrLf _
                & " group by WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE) C" & vbCrLf _
                & vbCrLf _
                & " where A.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and A.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and WHTPHYC1.WHSE_CODE (+) = A.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC1.TICKET_NO (+) = A.TICKET_NO" & vbCrLf _
                & "   and B.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and B.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and C.WHSE_CODE (+) = WHTLOCM1.WHSE_CODE" & vbCrLf _
                & "   and C.LOCATION_CODE (+) = WHTLOCM1.LOCATION_CODE" & vbCrLf _
                & "   and WHTLOCM1.WHSE_CODE = :PARM1"

            ASCMAIN1.sql = Replace(sqlWHTPHYCL, ":PARM1", "NULL")
            WHTPHYCL = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL($"Alter Table {WHTPHYCL} Add Primary Key (WHSE_CODE, LOCATION_CODE)")


            sqlWHTPHYCV = "Select X.STYLE_CODE, X.COLOR_CODE, ICTSTYL1.STYLE_DESC" & vbCrLf _
                & ", ICTSTYL1.CUST_CODE, CASE WHEN ICTSTYL1.CUST_CODE Is NULL THEN 'STK' ELSE 'NON' END SNS" & vbCrLf _
                & ", ICTSTYC1.STYLE_COST_FIFO STYLE_COST" & vbCrLf _
                & ", X.BOOK, X.PHYS from ICTSTYL1, ICTSTYC1, " & vbCrLf _
                & "(Select STYLE_CODE, COLOR_CODE, Sum (BOOK) BOOK, Sum (PHYS) PHYS from " & vbCrLf _
                & "(" & vbCrLf _
                & "Select WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE, 0 BOOK, Sum (NVL(WHTPHYC3.PHYS_UNITS,0)) PHYS" & vbCrLf _
                & " from WHTPHYC3, WHTPHYC1" & vbCrLf _
                & " where WHTPHYC3.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and WHTPHYC1.WHSE_CODE = WHTPHYC3.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC1.TICKET_NO = WHTPHYC3.TICKET_NO" & vbCrLf _
                & "   and WHTPHYC1.TICKET_STATUS = 'A'" & vbCrLf _
                & " group by WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE" & vbCrLf _
                & vbCrLf _
                & " union " & vbCrLf _
                & vbCrLf _
                & "Select STYLE_CODE, COLOR_CODE, Sum (NVL(LOCATION_QTY,0) + NVL(BOOK_INVTY_ADJ,0)) BOOK, 0 PHYS" & vbCrLf _
                & " from WHTLOCB0 where WHSE_CODE = :PARM1 group by STYLE_CODE, COLOR_CODE" & vbCrLf _
                & ")" & vbCrLf _
                & " group by STYLE_CODE, COLOR_CODE) X" & vbCrLf _
                & vbCrLf _
                & " where X.STYLE_CODE = ICTSTYL1.STYLE_CODE" & vbCrLf _
                & "   And ICTSTYC1.STYLE_CODE (+) = X.STYLE_CODE" & vbCrLf _
                & "   And ICTSTYC1.COLOR_CODE (+) = X.COLOR_CODE"

            '& " union " & vbCrLf _
            '& "Select STYLE_CODE, COLOR_CODE, WHSE_QTY_BEG BOOK, 0 PHYS" & vbCrLf _
            '& " from ICTSTAT1 where WHSE_CODE = :PARM1 and OPS_YYYYPP = :PARM2" & vbCrLf _

            ASCMAIN1.sql = sqlWHTPHYCV
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM1", "NULL")
            'ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM2", "NULL")
            'ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM3", "NULL")
            'ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM4", "NULL")
            WHTPHYCV = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL($"Alter Table {WHTPHYCV} Add Primary Key (STYLE_CODE, COLOR_CODE)")



            sqlWHTPHYCR = "" _
                & "Select X.WHSE_CODE, X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYC1.STYLE_COST_FIFO, ICTSTYL1.STYLE_DESC, MAX(X.TICKET_NO) TICKET_NO, MAX(WHTPHYC1.INIT_OPER) INIT_OPER" & vbCrLf _
                & ", Sum (X.PHYS_UNITS) PHYS_UNITS, Sum (X.BOOK_UNITS) BOOK_UNITS" & vbCrLf _
                & " from ICTSTYC1, ICTSTYL1, WHTPHYC1" & vbCrLf _
                & ", (" & vbCrLf _
                & "Select WHTPHYC3.WHSE_CODE, WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE" & vbCrLf _
                & ", SUM (WHTPHYC3.PHYS_UNITS) PHYS_UNITS, 0 BOOK_UNITS" & vbCrLf _
                & ", MAX (WHTPHYC3.TICKET_NO) TICKET_NO" & vbCrLf _
                & " from WHTPHYC1,WHTPHYC3" & vbCrLf _
                & " where WHTPHYC1.WHSE_CODE = :PARM1" & vbCrLf _
                & "   and WHTPHYC3.WHSE_CODE = WHTPHYC1.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC3.TICKET_NO = WHTPHYC1.TICKET_NO" & vbCrLf _
                & "   and WHTPHYC1.TICKET_STATUS = 'A'" & vbCrLf _
                & " group by WHTPHYC3.WHSE_CODE, WHTPHYC1.LOCATION_CODE" & vbCrLf _
                & ", WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE" & vbCrLf _
                & ", 0 PHYS_UNITS, SUM (WHTLOCB0.LOCATION_QTY - WHTLOCB0.BOOK_INVTY_ADJ) BOOK_UNITS" & vbCrLf _
                & ", NULL TICKET_NO" & vbCrLf _
                & " from WHTLOCB0" & vbCrLf _
                & " where WHTLOCB0.WHSE_CODE = :PARM1" & vbCrLf _
                & " group by WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
                & ", WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE" & vbCrLf _
                & ") X" & vbCrLf _
                & " where ICTSTYC1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and ICTSTYC1.COLOR_CODE = X.COLOR_CODE" & vbCrLf _
                & "   and ICTSTYL1.STYLE_CODE = X.STYLE_CODE" & vbCrLf _
                & "   and WHTPHYC1.WHSE_CODE (+) = X.WHSE_CODE" & vbCrLf _
                & "   and WHTPHYC1.TICKET_NO (+) = X.TICKET_NO" & vbCrLf _
                & " group by X.WHSE_CODE, X.LOCATION_CODE, X.STYLE_CODE, X.COLOR_CODE" & vbCrLf _
                & ", ICTSTYC1.STYLE_COST_FIFO, ICTSTYL1.STYLE_DESC"

            ASCMAIN1.sql = sqlWHTPHYCR
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, ":PARM1", "NULL")
            WHTPHYCR = ASCMAIN1.Temp_Table()
            ASCDATA1.ExecuteSQL($"Alter Table {WHTPHYCR} Add Primary Key (WHSE_CODE, LOCATION_CODE, STYLE_CODE, COLOR_CODE)")

        Else

            If variances_were_rebuilt Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Rebuilding Ticket Data")


            ASCMAIN1.Progress("-", "Locator")
            ASCMAIN1.sql = $"Truncate Table {WHTPHYCL}"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = $"Insert into {WHTPHYCL} " & sqlWHTPHYCL
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {WHSE_CODE})
            Fill_Records("WHTPHYCL")
            Sort_grdColumns(grdWHTPHYCL, "LOCATION_CODE")

            ASCMAIN1.Progress("-", "Variances")
            ASCMAIN1.sql = $"Truncate Table {WHTPHYCV}"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = $"Insert into {WHTPHYCV} " & sqlWHTPHYCV
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {WHSE_CODE})
            'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {WHSE_CODE, ASCMAIN1.CYP})
            'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New String() {WHSE_CODE, "", WHSE_CODE, ASCMAIN1.CYP})
            Fill_Records("WHTPHYCV")
            Sort_grdColumns(grdWHTPHYCV, "STYLE_CODE, COLOR_CODE")

            ASCMAIN1.Progress("-", "Loc/SC")
            ASCMAIN1.sql = $"Truncate Table {WHTPHYCR}"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = $"Insert into {WHTPHYCR} " & sqlWHTPHYCR
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {WHSE_CODE})
            Fill_Records("WHTPHYCR")
            Sort_grdColumns(grdWHTPHYCR, "LOCATION_CODE, STYLE_CODE, COLOR_CODE")
            'Set_WHTPHYCB()

            ASCMAIN1.Progress("-", "Loc v Per")
            Fill_Records("WHTLOCBX", WHSE_CODE)
            Sort_grdColumns(grdWHTLOCBX, "STYLE_CODE, COLOR_CODE")

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")


            variances_were_rebuilt = True

        End If
    End Sub

    Sub Setup_WHTPHYC2(BAR_CODE As String, COUNTS_BY_UPC As String)
        If BAR_CODE = "" Then
            splItemDetails.Visible = False
            grdWHTPHYC3.Visible = False
        Else
            splItemDetails.Visible = True
            grdWHTPHYC3.Visible = True

            Dim dvw As DataView = DirectCast(grdWHTPHYC3.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"BAR_CODE = '{BAR_CODE}'"

            If COUNTS_BY_UPC = "1" And (EntryMode = "E") Then
                grdWHTPHYC3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdWHTPHYC3.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
            Else
                grdWHTPHYC3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdWHTPHYC3.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
            End If
        End If

    End Sub

    Sub Setup_WHTPHYC3(STYLE_CODE As String, COLOR_CODE As String)
        If STYLE_CODE = "" Then
            splItemDetails.Visible = False
        Else
            splItemDetails.Visible = True
            Fill_Records("WHTPHYCI", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
            grdWHTPHYCI.Text = "Tickets with Style " & STYLE_CODE & "-" & COLOR_CODE
            If location_support Then
                Fill_Records("WHTLOCBA", New String() {WHSE_CODE, STYLE_CODE, COLOR_CODE})
                grdWHTLOCBA.Text = "Book Inventory by Location for Style " & STYLE_CODE & "-" & COLOR_CODE
                Setup_WHTLOCBA(False)
            End If
        End If

    End Sub

    Sub Setup_WHTLOCBA(Show_0s As Boolean)
        Dim dvw As DataView = DirectCast(grdWHTLOCBA.DataSource, DataTable).DefaultView
        If Show_0s Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "LOCATION_QTY <> 0"
        End If
    End Sub

    Sub CloseInventory()
        Dim LYP As String = ""
        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

        Stop
        'Set current_period_physical flag to the correct setting
        Dim current_period_physical As Boolean = False
        Dim WHSE_TRAN_NO As String = ""
        Dim ADJ_NO As String = ""
        Dim ADJ_YP As String = ""

        If current_period_physical Then
            ADJ_YP = ASCMAIN1.CYP
        Else
            ADJ_YP = LYP
        End If

        ASCMAIN1.sql = "Select WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE
                        , SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS, sum( nvl(PHYS_UNITS,0) - nvl(BOOK_UNITS,0)) VAR_UNITS
                        from (
                        Select WHTPHYC3.WHSE_CODE, WHTPHYC3.BAR_CODE, WHTPHYC3.STYLE_CODE, WHTPHYC3.COLOR_CODE
                        , WHTPHYC3.PHYS_UNITS, 0 BOOK_UNITS, WHTPHYC1.LOCATION_CODE
                        from WHTPHYC3,WHTPHYC1
                        where WHTPHYC3.WHSE_CODE = '" & WHSE_CODE & "'
                        and WHTPHYC1.TICKET_STATUS = 'A'
                        and WHTPHYC1.WHSE_CODE = WHTPHYC3.WHSE_CODE
                        and WHTPHYC1.TICKET_NO = WHTPHYC3.TICKET_NO
                        UNION
                        Select WHTLOCB0.WHSE_CODE, WHTLOCB0.BAR_CODE, WHTLOCB0.STYLE_CODE, WHTLOCB0.COLOR_CODE
                        , 0 PHYS_UNITS, WHTLOCB0.LOCATION_QTY BOOK_UNITS, WHTLOCB0.LOCATION_CODE
                        from WHTLOCB0
                        where WHTLOCB0.WHSE_CODE = '" & WHSE_CODE & "'
                        ) group by WHSE_CODE, LOCATION_CODE, BAR_CODE, STYLE_CODE, COLOR_CODE"

        Dim WHTPHYV1 As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCMAIN1.sql = $"select  SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS, sum(VAR_UNITS) VAR_UNITS
                        , count(1) RECORDS, sum(case when nvl(PHYS_UNITS,0) - nvl(BOOK_UNITS,0) <> 0 then 1 else 0 end) ADJ_CNT
                        from {WHTPHYV1} WHTPHYV1"
        Dim rowSUMMARY As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)

        If MsgBox("This Action will commit the variances to the current Warehouse Selected for Physical Inventory Processing." & vbCrLf _
                & vbCrLf & vbCrLf & "This includes Updating all data in the Counts Files for this Warehouse," _
                & vbCrLf & " variances will adjust Book  Inventory values by Item/Location," _
                & vbCrLf & vbCrLf & "Booked: " & String.Format("{0:###,###,##0}", rowSUMMARY("BOOK_UNITS")) &
                " Counted: " & String.Format("{0:###,###,##0}", rowSUMMARY("PHYS_UNITS")) &
                " Variance: " & String.Format("{0:###,###,##0}", rowSUMMARY("VAR_UNITS")) & "" _
                & vbCrLf & "A total of " & String.Format("{0:###,###,##0}", rowSUMMARY("ADJ_CNT")) & " adjustment records will be created." _
                & vbCrLf & vbCrLf & "OK to Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        BeginTrans()

        If current_period_physical Then
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select WHSE_CODE, STYLE_CODE, COLOR_CODE " & vbCrLf _
                & " , SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS, sum(VAR_UNITS) VAR_UNITS" & vbCrLf _
                & "  from " & WHTPHYV1 & " Group By WHSE_CODE, STYLE_CODE, COLOR_CODE;" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_PHY = NVL(WHSE_QTY_PHY,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_PHY) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & ASCMAIN1.CYP & "',VAR_UNITS); End If;" & vbCrLf _
                & "   Update ICTSTAT2 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,R1.VAR_UNITS); End If;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "" _
                & "Begin Declare Cursor C1 is Select WHSE_CODE, STYLE_CODE, COLOR_CODE " & vbCrLf _
                & " , SUM (PHYS_UNITS) PHYS_UNITS, SUM (BOOK_UNITS) BOOK_UNITS, sum(VAR_UNITS) VAR_UNITS" & vbCrLf _
                & "  from " & WHTPHYV1 & " Group By WHSE_CODE, STYLE_CODE, COLOR_CODE;" & vbCrLf _
                & " Begin " & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_BEG = nvl(WHSE_QTY_BEG,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_BEG) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & ASCMAIN1.CYP & "',R1.VAR_UNITS); End If;" & vbCrLf _
                & "   Update ICTSTAT1 Set WHSE_QTY_PHY = NVL(WHSE_QTY_PHY,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_PHY) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & LYP & "',R1.VAR_UNITS); End If;" & vbCrLf _
                & "   Update ICTSTAT2 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT2 (STYLE_CODE,COLOR_CODE,WHSE_CODE,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,R1.VAR_UNITS); End If;" & vbCrLf _
                & "   Update ICTSTAT5 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + R1.VAR_UNITS where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
                & "   If SQL%NOTFOUND Then Insert into ICTSTAT5 (STYLE_CODE,COLOR_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_ON_HAND) Values (R1.STYLE_CODE,R1.COLOR_CODE,R1.WHSE_CODE,'" & LYP & "',R1.VAR_UNITS); End If;" & vbCrLf _
                & "  End Loop; " & vbCrLf _
                & " End; " & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")
        If ASCMAIN1.CLIENT = "VAN" Then
            ADJ_NO = ASCMAIN1.Next_Control_No("TRAN_NO_A")
        Else
            ADJ_NO = ASCMAIN1.Next_Control_No("ICTIADJ1.ADJ_NO")
        End If

        ASCMAIN1.sql = $"Begin
                        declare
                        L_DATE DATE := sysdate;
                        Begin
                        -- 01 is Physical Inventory Adj, ADJ_SOURCE - P (Physical)
                        insert into ICTIADJ1 (ADJ_NO, ADJ_DATE, WHSE_CODE, REASON_CODE, ADJ_NOTE, INIT_OPER, INIT_DATE, 
                                             ADJ_SOURCE, OPS_YYYYPP, LAST_OPER, LAST_DATE, ADJ_REF, JOURNAL_IND, JOURNAL_XNO)
                                    values ('{ADJ_NO }', L_DATE, '{WHSE_CODE}', '01','PHYSISCAL INVENTORY ADJ','{ASCMAIN1.USER_ID}',
                                            L_DATE, 'P', '{ADJ_YP}', '{ASCMAIN1.USER_ID}', L_DATE, '{WHSE_TRAN_NO}','1','0000000000');
                        
                        insert into ICTIADJ2 (ADJ_NO, ADJ_LNO, STYLE_CODE, COLOR_CODE, ADJ_QTY, STYLE_COST, STYLE_CLASS_CODE,
                            SALES_DIVISION_CODE, OPS_YYYYPP, LOCATION_CODE, BAR_CODE, ADJ_REF)
                        select '{ADJ_NO}',rownum, STYLE_CODE, COLOR_CODE, VAR_UNITS, STYLE_COST, STYLE_CLASS_CODE, 
                                        SALES_DIVISION_CODE, '{ADJ_YP}', LOCATION_CODE, BAR_CODE, '{WHSE_TRAN_NO}'
                                            from (Select WHTPHYV1.*, ICTSTYL1.STYLE_COST, ICTSTYL1.STYLE_CLASS_CODE, ICTSTYL1.SALES_DIVISION_CODE
                                            from {WHTPHYV1} WHTPHYV1, ICTSTYL1
                                             where NVL(VAR_UNITS,0) <> 0
                                             and WHTPHYV1.STYLE_CODE = ICTSTYL1.STYLE_CODE(+));
                        end;
                        end;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"
        Begin Declare Cursor C1 is Select * from {WHTPHYV1} where VAR_UNITS <> 0;
              Begin 
               For R1 in C1 Loop
                Update WHTLOCB1 Set LOCATION_QTY = NVL(LOCATION_QTY,0) + R1.VAR_UNITS 
                    where STYLE_CODE = R1.STYLE_CODE and COLOR_CODE = R1.COLOR_CODE and WHSE_CODE = R1.WHSE_CODE 
                    and LOCATION_CODE = R1.LOCATION_CODE and BAR_CODE = R1.BAR_CODE;
                If SQL%NOTFOUND Then Insert into WHTLOCB1 (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE,LOCATION_QTY) 
                  Values (R1.WHSE_CODE,R1.LOCATION_CODE,R1.BAR_CODE,R1.STYLE_CODE,R1.COLOR_CODE, R1.VAR_UNITS); End If;
               End Loop; 
               Insert into WHTLOCB2 (WHSE_CODE,LOCATION_CODE,BAR_CODE,STYLE_CODE,COLOR_CODE,WHSE_TRAN_QTY,
                        WHSE_TRAN_TYPE,WHSE_TRAN_NO,WHSE_TRAN_LNO,INIT_DATE,INIT_OPER,LOCATION_CODE_OTHER,SESSION_NO)
                 select V1.WHSE_CODE,V1.LOCATION_CODE,V1.BAR_CODE,V1.STYLE_CODE,V1.COLOR_CODE,V1.VAR_UNITS,'P'
                        ,'{WHSE_TRAN_NO}',rownum,SYSDATE,'{ASCMAIN1.USER_ID}',NULL,'{ASCMAIN1.SESSION_NO}'
                 from {WHTPHYV1} V1 where VAR_UNITS <> 0;        
              End; 
             End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
          & $"Update ICTWHSE1 X Set WHSE_YYYYPP_LAST_PHY = '{LYP}', WHSE_PHYS_STATUS = NULL and WHSE_CODE = '{WHSE_CODE}'"
        ASCDATA1.ExecuteSQL()
        CommitTrans("Inventory CLosed")

    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        If tab0.SelectedTab.Key = "Tickets" Then
            splWHTPHYCB.Parent = splWHTPHYCX.Panel2
            Set_WHTPHYCB()
        ElseIf tab0.SelectedTab.Key = "Variances" Then
            If Load_Variances() Then
                splItemDetails.Parent = splWHTPHYCV.Panel2
                Show_Variances()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        ElseIf tab0.SelectedTab.Key = "Locations" Then
            If Load_Locations() Then
                splWHTPHYCB.Parent = splWHTPHYCL.Panel2
                Set_WHTPHYCB()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        ElseIf tab0.SelectedTab.Key = "Location/Style/Color" Then
            If Load_Location_Style_Colors() Then
                splWHTPHYCB.Parent = splWHTPHYCR.Panel2
                Set_WHTPHYCB()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        Else
            splItemDetails.Parent = splWHTPHYC3.Panel2
        End If

        UltraExplorerBar1.Groups("Variances").Visible = (tab0.SelectedTab.Key = "Variances") And Not ScreenMode
        UltraExplorerBar1.Groups("Tickets").Visible = (tab0.SelectedTab.Key = "Tickets") And Not ScreenMode
        UltraExplorerBar1.Groups("Location Options").Visible = (tab0.SelectedTab.Key = "Locations") And Not ScreenMode

        Hide_Control_Panel()

    End Sub

    Function Load_Variances() As Boolean
        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Variance")
            Return False
        End If

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Compiling Variances")

        If grdWHTPHYCV.ActiveRow Is Nothing Then Setup_WHTPHYC3("", "")

        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")

        Return True

    End Function
    Function Load_Locations() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        ElseIf rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            MsgBox("This option is valid only for Warehouses currently in a Physical Inventory", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        End If

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Compiling Location Control Totals")
        'Fill_Records("WHTPHYCL")
        'Set_Filters_for_WHTPHYCL()
        'Sort_grdColumns(grdWHTPHYCL, "LOCATION_CODE")
        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")

        Set_Filters_for_WHTPHYCL()

        Return True

    End Function

    Sub Set_Filters_for_WHTPHYCL()
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim dvw As DataView = DirectCast(grdWHTPHYCL.DataSource, DataTable).DefaultView
        Dim sql As String = ""

        If optLocZero.Value = "Z" Then
            sql &= " and (ISNULL(PHYS_UNITS,0) = 0 AND ISNULL(BOOK_UNITS,0) = 0)"
        ElseIf optLocZero.Value = "N" Then
            sql &= " and (ISNULL(PHYS_UNITS,0) <> 0 OR ISNULL(BOOK_UNITS,0) <> 0)"
        End If

        If optLocCounted.Value = "U" Then
            sql &= " and TICKET_NO is Null"
        ElseIf optLocZero.Value = "N" Then
            sql &= " and TICKET_NO is Not Null"
        End If

        If chkLocVar.Checked Then
            sql &= " and ISNULL(VARIANCE,0) <> 0"
        End If
        If chkLocVirtual.Checked Then
        Else
            sql &= " and ISNULL(VIRTUAL,'0') = '0'"
        End If

        If sql <> "" Then
            sql = Mid(sql, 5)
        End If

        dvw.RowFilter = sql
    End Sub
    Function Load_Location_Style_Colors() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        ElseIf rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            MsgBox("This option is valid only for Warehouses currently in a Physical Inventory", MsgBoxStyle.OkOnly, "Cannot Show Locations")
            Return False
        End If

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Compiling Location Control Totals")
        'Fill_Records("WHTPHYCR", New String() {WHSE_CODE})
        'Sort_grdColumns(grdWHTPHYCR, "LOCATION_CODE,STYLE_CODE,COLOR_CODE")
        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")

        Return True
    End Function

    'Private Sub grdWHTPHYCI_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdWHTPHYCI.DoubleClickRow
    '    If Not ScreenMode Then
    '        Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Value & ""
    '        Click_Command("View")
    '    End If
    'End Sub

    Sub Print_Counts(BY As String)

        Dim RPT As String = ""
        Dim RPT_TITLE As String = ""
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        If WHSE_CODE = "" Then Exit Sub
        Dim rowICTWHSE1 As DataRow = Fill_Record("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then Exit Sub

        ASCMAIN1.sql = "Select * from WHTPHYC1 where WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("WHTPHYC1", "", True, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select WHTPHYC3.*, ICTSTYL1.STYLE_DESC" _
            & " from WHTPHYC3,ICTSTYL1 where ICTSTYL1.STYLE_CODE = WHTPHYC3.STYLE_CODE" _
            & " and WHTPHYC3.WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("WHTPHYC3", "", True, ASCMAIN1.sql)
        Dim dvwC2 As DataView = DirectCast(grdWHTPHYC3.DataSource, DataTable).DefaultView
        dvwC2.RowFilter = "STATUS IS NULL"

        Select Case BY
            Case "T"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Ticket"
            Case "L"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Location"
            Case "S"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Style"
        End Select

        'Synch_TABLE_NAME("ICTSTYL1")
        Print_Report_Begin()
        CR_params.Add("SORT_BY", BY)
        Generate_Report(RPT, RPT_TITLE, "")
        Print_Report_End()

    End Sub

    Sub Hide_Control_Panel()
        ' MAKING THESE INVISIBLE UNLESS AND UNTIL WE HAVE A REASON WHY WE NEED THESE
        With UltraExplorerBar1
            .Groups("Count Reports").Visible = False
            With .Groups("Screen Control")
                .Items("New").Visible = False
                .Items("View").Visible = (Not ScreenMode) And (tab0.SelectedTab.Key = "Tickets")
                .Items("Edit").Visible = (Not ScreenMode) And (tab0.SelectedTab.Key = "Tickets")
                .Items("Update").Visible = (ScreenMode And EntryMode = "E") And (tab0.SelectedTab.Key = "Tickets")
                .Items("Cancel").Visible = (ScreenMode And EntryMode = "E") And (tab0.SelectedTab.Key = "Tickets")
                .Items("Delete").Visible = (ScreenMode And EntryMode = "E") And (tab0.SelectedTab.Key = "Tickets")
                .Items("Done").Visible = (ScreenMode And EntryMode = "V") And (tab0.SelectedTab.Key = "Tickets")
            End With
        End With
    End Sub

    Private Sub chkLocVirtual_CheckedChanged(sender As Object, e As EventArgs) Handles chkLocVirtual.CheckedChanged
        Set_Filters_for_WHTPHYCL()
    End Sub

    Private Sub chkLocVar_CheckedChanged(sender As Object, e As EventArgs) Handles chkLocVar.CheckedChanged
        Set_Filters_for_WHTPHYCL()
    End Sub

    Private Sub optLocZero_ValueChanged(sender As Object, e As EventArgs) Handles optLocZero.ValueChanged
        Set_Filters_for_WHTPHYCL()
    End Sub

    Private Sub optLocCounted_ValueChanged(sender As Object, e As EventArgs) Handles optLocCounted.ValueChanged
        Set_Filters_for_WHTPHYCL()
    End Sub

    Private Sub grdWHTPHYCR_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTPHYCR.AfterRowActivate
        If grdWHTPHYCR.ActiveRow Is Nothing OrElse Not grdWHTPHYCR.ActiveRow.IsDataRow Then Exit Sub
        If tab0.SelectedTab.Key <> "Location/Style/Color" Then Exit Sub

        Set_WHTPHYCB()
    End Sub

    Sub Set_WHTPHYCB()

        Dim WHSE_CODE As String = ""
        Dim LOCATION_CODE As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""

        If tab0.SelectedTab.Key = "Tickets" Then
            If grdWHTPHYCX.ActiveRow Is Nothing Then
                splWHTPHYCB.Visible = False
                Exit Sub
            Else
                splWHTPHYCB.Visible = True
            End If
            WHSE_CODE = grdWHTPHYCX.ActiveRow.Cells("WHSE_CODE").Text
            LOCATION_CODE = grdWHTPHYCX.ActiveRow.Cells("LOCATION_CODE").Text
            STYLE_CODE = ""
            COLOR_CODE = ""
        ElseIf tab0.SelectedTab.Key = "Locations" Then
            If grdWHTPHYCL.ActiveRow Is Nothing Then
                splWHTPHYCB.Visible = False
                Exit Sub
            Else
                splWHTPHYCB.Visible = True
            End If
            WHSE_CODE = grdWHTPHYCL.ActiveRow.Cells("WHSE_CODE").Text
            LOCATION_CODE = grdWHTPHYCL.ActiveRow.Cells("LOCATION_CODE").Text
            STYLE_CODE = ""
            COLOR_CODE = ""
        ElseIf tab0.SelectedTab.Key = "Location/Style/Color" Then
            If grdWHTPHYCR.ActiveRow Is Nothing Then
                splWHTPHYCB.Visible = False
                Exit Sub
            Else
                splWHTPHYCB.Visible = True
            End If
            WHSE_CODE = grdWHTPHYCR.ActiveRow.Cells("WHSE_CODE").Text
            LOCATION_CODE = grdWHTPHYCR.ActiveRow.Cells("LOCATION_CODE").Text
            STYLE_CODE = grdWHTPHYCR.ActiveRow.Cells("STYLE_CODE").Text
            COLOR_CODE = grdWHTPHYCR.ActiveRow.Cells("COLOR_CODE").Text
        End If

        Dim ALL_STYLES As String = "N"
        If STYLE_CODE = "" Then
            ALL_STYLES = "*"
        End If

        Fill_Records("WHTPHYCB", New String() {WHSE_CODE, LOCATION_CODE, ALL_STYLES, STYLE_CODE, COLOR_CODE})
        Sort_grdColumns(grdWHTPHYCB, "BAR_CODE,STYLE_CODE,COLOR_CODE")
        If ALL_STYLES = "*" Then
            grdWHTPHYCB.Text = $"Physical & Book Units by LPN for Location {LOCATION_CODE}"
        Else
            grdWHTPHYCB.Text = $"Physical & Book Units by LPN for Location {LOCATION_CODE}, Style/Color {STYLE_CODE}/{COLOR_CODE}"
        End If


        Fill_Records("WHTPHYCS", New String() {WHSE_CODE, LOCATION_CODE, ALL_STYLES, STYLE_CODE, COLOR_CODE})
        Sort_grdColumns(grdWHTPHYCS, "STYLE_CODE,COLOR_CODE")
        If ALL_STYLES = "*" Then
            grdWHTPHYCS.Text = $"Physical & Book Units Summary for Location {LOCATION_CODE}"
        Else
            grdWHTPHYCS.Text = $"Physical & Book Units Summary for Location {LOCATION_CODE}, Style/Color {STYLE_CODE}/{COLOR_CODE}"
        End If

    End Sub

    Private Sub grdWHTPHYCL_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTPHYCL.AfterRowActivate
        If Not grdWHTPHYCL.ActiveRow.IsDataRow Then Exit Sub
        If tab0.SelectedTab.Key <> "Locations" Then Exit Sub

        Set_WHTPHYCB()
    End Sub

    Private Sub grdWHTPHYCB_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTPHYCB.InitializeRow
        With e.Row.Cells("VAR_UNITS")
            If Val(.Value & "") < 0 Then
                .Appearance.ForeColor = System.Drawing.Color.Red
            Else
                .Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub

    Private Sub grdWHTPHYCS_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTPHYCS.InitializeRow
        With e.Row.Cells("VAR_UNITS")
            If Val(.Value & "") < 0 Then
                .Appearance.ForeColor = System.Drawing.Color.Red
            Else
                .Appearance.ForeColor = System.Drawing.Color.Empty
            End If
        End With
    End Sub

    Sub Add_WHTPHYC3(BAR_CODE As String)

        Dim SEQ_NO = ASCDATA1.GetDataValue($"SELECT SEQ_NO FROM WHTPHYC4 Where BAR_CODE = '{BAR_CODE}' and STATUS = '0'") & ""
        If SEQ_NO <> "" Then ' we have a closed carton
            ASCMAIN1.sql = $"Select * from WHTPHYC5 WHERE WHSE_CODE = '{WHSE_CODE}' and SEQ_NO = '{SEQ_NO}' and BAR_CODE = '{BAR_CODE}'"
            For Each rowWHTPHYC5 As DataRow In ASCDATA1.GetDataTable().Select("")
                If rowWHTPHYC5("QTY_CNT") > 0 Then
                    Dim rowWHTPHYC3 As DataRow = dst.Tables("WHTPHYC3").NewRow
                    With rowWHTPHYC3
                        .Item("WHSE_CODE") = WHSE_CODE
                        .Item("TICKET_NO") = TICKET_NO
                        .Item("BAR_CODE") = BAR_CODE
                        .Item("STYLE_CODE") = rowWHTPHYC5("STYLE_CODE")
                        .Item("COLOR_CODE") = rowWHTPHYC5("COLOR_CODE")
                        .Item("PHYS_UNITS") = rowWHTPHYC5("QTY_CNT")
                    End With
                    dst.Tables("WHTPHYC3").Rows.Add(rowWHTPHYC3)
                End If
            Next
        Else
            Dim found = Val(dst.Tables("WHTLOCB0").Compute("sum(LOCATION_QTY)", $"BAR_CODE='{BAR_CODE}'") & "")
            Dim negative = Val(dst.Tables("WHTLOCB0").Compute("sum(LOCATION_QTY)", $"BAR_CODE='{BAR_CODE}' and LOCATION_QTY < 0") & "")
            If negative = 0 And found > 0 Then
                For Each rowWHTLOCB1 As DataRow In dst.Tables("WHTLOCB0").Select($"BAR_CODE='{BAR_CODE}'")
                    Dim rowWHTPHYC3 As DataRow = dst.Tables("WHTPHYC3").NewRow
                    With rowWHTPHYC3
                        .Item("WHSE_CODE") = WHSE_CODE
                        .Item("TICKET_NO") = TICKET_NO
                        .Item("BAR_CODE") = BAR_CODE
                        .Item("STYLE_CODE") = rowWHTLOCB1("STYLE_CODE")
                        .Item("COLOR_CODE") = rowWHTLOCB1("COLOR_CODE")
                        .Item("PHYS_UNITS") = rowWHTLOCB1("LOCATION_QTY")
                    End With
                    dst.Tables("WHTPHYC3").Rows.Add(rowWHTPHYC3)
                Next
            Else
                'finally if no open carton counts or clean locb1, goto receiving
                ASCMAIN1.sql = $"Select * from WHTLOCB2 WHERE WHSE_CODE = '{WHSE_CODE}' and BAR_CODE = '{BAR_CODE}' and WHSE_TRAN_TYPE = 'W'"
                For Each rowWHTLOCB2 As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim rowWHTPHYC3 As DataRow = dst.Tables("WHTPHYC3").NewRow
                    With rowWHTPHYC3
                        .Item("WHSE_CODE") = WHSE_CODE
                        .Item("TICKET_NO") = TICKET_NO
                        .Item("BAR_CODE") = BAR_CODE
                        .Item("STYLE_CODE") = rowWHTLOCB2("STYLE_CODE")
                        .Item("COLOR_CODE") = rowWHTLOCB2("COLOR_CODE")
                        .Item("PHYS_UNITS") = rowWHTLOCB2("WHSE_TRAN_QTY")
                    End With
                    dst.Tables("WHTPHYC3").Rows.Add(rowWHTPHYC3)
                Next
            End If
        End If

    End Sub

    Private Sub grdWHTPHYC2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdWHTPHYC2.BeforeRowsDeleted

    End Sub

    Private Sub grdWHTPHYC2_AfterColRegionScroll(sender As Object, e As ColScrollRegionEventArgs) Handles grdWHTPHYC2.AfterColRegionScroll

    End Sub
End Class