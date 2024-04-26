Imports System.Text
Imports System.Math
Imports System.Drawing


Public Class PORPREC1
    Dim ICTPREC1 As String
    Dim RYP_Legend As String = ""
    Dim JOURNAL_LNO As Integer = 0
    Dim NYP As String = ""
    Dim grdASTEXPT2 As New UltraWinGrid.UltraGrid
    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String

    Dim RECEIPT_DATE_FROM As String = ""
    Dim RECEIPT_DATE_TO As String = ""


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("POTPARM1")

        '    Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "N"
        'RYP_Legend = Absx1.cmbFor("RYP0").Value
        'RYP = Mid(RYP_Legend, 1, 4) & Mid(RYP_Legend, 6, 2)

        'NYP = ASCMAIN1.Period_Calc(RYP, 1)

        Dim z As String = ""

        If Absx1.chkFor("CHKPO_DATE_F").Checked Then
                RECEIPT_DATE_FROM = Format(DateSerial(2000, 1, 1), "dd-MMM-yyyy")
            Else
                z = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
                RECEIPT_DATE_FROM  = z
            End If

            If Absx1.chkFor("CHKPO_DATE_L").Checked Then
                RECEIPT_DATE_TO = Format(DateSerial(2100, 1, 1), "dd-MMM-yyyy")
            Else
                z = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
                RECEIPT_DATE_TO  = z
            End If


        'LPeriod = Format(DateSerial(2100, 12, 1), "dd-MMM-yyyy")

  

        ' THIS IS THE SQL TO REPLACE SQL BELOW

        ASCMAIN1.sql = $"
           SELECT ICTIREC2.RECEIPT_NO,ICTIREC2.RECEIPT_LNO,ICTIREC1.RECEIPT_DATE,POTORDR1.VEND_CODE,POTORDR1.VEND_NAME,ICTIREC2.PO_ORDER_NO ,ICTIREC2.PO_ORDER_LNO ,POTORDR1.PO_REFERENCE,POTORDR1.PO_STATUS,
        ICTIREC2.STYLE_CODE,ICTIREC2.COLOR_CODE,ICTIREC2.QTY_REC,ICTIREC2.PO_COST,ICTIREC2.QTY_REC * ICTIREC2.PO_COST PO_COST_EXT,ICTIREC2.QTY_SHP,
        POTSHIP3.PO_COST_LANDED,ICTIREC2.PO_SHIPMENT_NO,ICTIREC2.PO_SHIPMENT_LNO
        FROM ICTIREC1,ICTIREC2,POTORDR1,POTSHIP3,POTORDR2 WHERE
        ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO AND
        POTORDR2.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO AND
        POTORDR2.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO AND
        POTORDR1.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO AND
        POTSHIP3.PO_ORDER_NO = ICTIREC2.PO_ORDER_NO AND
        POTSHIP3.PO_ORDER_LNO = ICTIREC2.PO_ORDER_LNO AND
        ICTIREC1.RECEIPT_DATE >= '" & RECEIPT_DATE_FROM & "'
        AND ICTIREC1.RECEIPT_DATE <= '" & RECEIPT_DATE_TO & "'
        AND ICTIREC1.REVERSED_BY_RECEIPT_NO IS NULL AND ICTIREC1.REVERSES_RECEIPT_NO IS NULL
        ORDER BY ICTIREC2.PO_ORDER_NO,ICTIREC2.PO_ORDER_LNO"
        ICTPREC1 = ASCMAIN1.Temp_Table



        Dim sqlw As String = ""
        sqlw &= SQL_in("VEND_CODE", "ICTPREC1.VEND_CODE")
        If sqlw <> "" Then RWU = "N"

        With dst
            ASCMAIN1.sql = $"Select ICTPREC1.* from {ICTPREC1} ICTPREC1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
            Create_TDA(.Tables.Add, "ICTPREC1", "**", , False)

            ASCMAIN1.sql = $"Select APTVEND1.VEND_CODE, APTVEND1.VEND_NAME from APTVEND1 where APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from {ICTPREC1} ICTPREC1 )"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)
        End With

        Fill_Records("ICTPREC1")
        Fill_Records("APTVEND1")

        '       GL_Interface()

        Check_if_Empty("ICTPREC1")
    End Sub

    Public Overrides Sub Print_Report()

        RPT = "PORPREC1"


        SUBT = "Receipts Posted between " & RECEIPT_DATE_FROM & " and " & RECEIPT_DATE_TO



        ' SUBT = "Report Period "
        ' DO I NEED A REPORT
        Generate_Report(RPT, , SUBT)


        ' Generate_Report(RPT)
        '      Print_GL()
        If ASCMAIN1.CLIENT = "VAN" Then
            Prepare_Data_Extracts()
        End If

    End Sub

    Overrides Sub Update_Record()


    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        '  grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")
        grdASTEXPT1.DataSource = dst.Tables("ICTPREC1")

        grdASTEXPT1.Text = MENU_ITEM_DESC

        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
            grdASTEXPT1.DisplayLayout.Bands(0).Columns("G" & CStr(G)).Header.Fixed = True
        Next
        Set_DX_Column(grdASTEXPT1, "RECEIPT_NO", "Rec No", 80, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_LNO", "Rec Lno", 60, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "RECEIPT_DATE", "Receipt Dt", 90, "MM/dd/yy", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vend Code", 90, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VEND_NAME", "Vend Name", 120, , , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_NO", "PO No", 80, , , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PO_ORDER_LNO", "PO Lno No", 60, , , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PO_REFERENCE", "PO Reference", 100, , , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "PO_STATUS", "PO Stat No", 60)
        Set_DX_Column(grdASTEXPT1, "STYLE_CODE", "Style", 60)
        Set_DX_Column(grdASTEXPT1, "COLOR_CODE", "Color", 60)

        Set_DX_Column(grdASTEXPT1, "QTY_REC", "Qty Rec", 90, "#,###,##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "PO_COST", "PO Cost", 90, "#,##0.0000", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "PO_COST_EXT", "PO Cost Ext", 120, "##,###,##0.00", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "QTY_SHP", "Qty Shp", 90, "#,###,##0", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "PO_COST_LANDED", "Landed Cost", 95, "#,##0.0000", , Color.Pink)
        Set_DX_Column(grdASTEXPT1, "PO_SHIPMENT_NO", "Ship No", 80)
        Set_DX_Column(grdASTEXPT1, "PO_SHIPMENT_LNO", "Ship Lno", 60)




        Create_Summary(grdASTEXPT1, "VEND_CODE", "Count")
        Create_Summary(grdASTEXPT1, New String() {"QTY_REC", "QTY_SHP"})



        '  grdASTEXPT1.DisplayLayout.Bands(0).Columns("STYLE_CODE").Header.Fixed = True

        Sort_grdColumns(grdASTEXPT1, "PO_ORDER_NO,PO_ORDER_LNO")



    End Sub


End Class