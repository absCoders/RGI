Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Net.Mail

Public Class WBFSTYLR
    Private WithEvents Sftp1 As New nsoftware.IPWorks.Ftp
    Private shopSiteFilename As String = String.Empty
    Private itemUploaded As Boolean = False
    Private sqlXmlUpload As String = String.Empty
    Private WB_PARM_SITE_IP As String = String.Empty
    Private WB_PARM_SITE_USER As String = String.Empty
    Private WB_PARM_SITE_PWD As String = String.Empty
    Private WB_PARM_SITE_OUTPUT_DIR As String = String.Empty
    Private WB_PARM_SITE_PRODUCT_POST_URL As String = String.Empty
    Private WB_PARM_SITE_PRODUCT_PUB_URL As String = String.Empty
    Private WB_PARM_PRODUCTS_DIR As String = String.Empty
    Private WB_PARM_INVENTORY_DIR As String = String.Empty
    Private WB_PARM_MASTER_IMAGES As String = String.Empty
    Private WB_PARM_SITE_IMAGES_DIR As String = String.Empty
    Private WB_PARM_IMAGES_DIR As String = String.Empty
    Private WB_PARM_SITE_ORDERS_POST_URL As String = String.Empty
    Private WB_PARM_ORDERS_DIR As String = String.Empty
    Private WB_PARM_RSS_MAX_ENTRIES As Int16 = 0
    Private WB_PARM_RSS_NEW_PAGE As String = String.Empty
    Private WB_PARM_RSS_SALE_PAGE As String = String.Empty
    Private WB_NEW_PROD_MAX As Integer = 0
    Private WB_NEW_PROD_PAGE_CODE As String = String.Empty
    Private docComplete As Boolean = False
    Private styleList As List(Of String) = New List(Of String)
    Private styleListInactive As List(Of String) = New List(Of String)
    Private ICTPEND_TEMP As String = ""

    ' Used for automation 
    Private TASK_NO As String = String.Empty
    Private WBCMAIN1 As New WBCMAIN1
    Dim InAutoMode As Boolean = False
    Dim TickCount As Integer = 0
    Dim FTPImages As Boolean = True
    Dim FTPTables As Boolean = False
    Dim UpLoadType As String = ""
    'Const CandidateFilter As String = "WEB_IND = '1' AND ISNULL(UPLOAD_BATCH,'NULL') = 'NULL'"
    Dim AutoProcessRunning As Boolean = False
    Dim MissingZipFiles As New List(Of String)

    'Wayfair Stuff
    Dim TransferFile As String = String.Format("{0}regency.csv", ASCMAIN1.Folders("Temp"))
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

#Region "ABS Standard Routines"
    Private Sub MoveToParents()
        TabControl1.Parent = UltraTabControl1.Parent
        WebBrowser1.Parent = UltraTabControl1.Parent
        'SplitContainer1.Parent = UltraTabControl1.Parent

        'grdWBTSHOP1.Parent = UltraTabControl1.Parent
        'grdWBTSHOP2.Parent = UltraTabControl1.Parent
        'SplitContainer2.Parent = UltraTabControl1.Parent
    End Sub
    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Dim sql As String = String.Empty

        With dst
            Create_TDA(.Tables.Add, "ICTSTYL3", "*")

            Create_TDA(.Tables.Add, "WBTSTYL1", "*")
            .Tables("WBTSTYL1").Columns.Add("STYLE_CLASS_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "WBTSTYL2", "*")

            Create_TDA(.Tables.Add, "WBTSHOP1", "*")

            Create_TDA(.Tables.Add, "WBTSHOP2", "*")

            Create_TDA(.Tables.Add, "ICTSTAT2", "*")
            .Tables("ICTSTAT2").Columns.Add("ITEM_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "WBTRSSF1", "*")

            Dim sqls As New StringBuilder
            sqls.Length = 0
            sqls.AppendLine("SELECT * FROM")
            sqls.AppendLine("  (")
            sqls.AppendLine("   SELECT C1.STYLE_CODE, C1.COLOR_CODE,")
            sqls.AppendLine("   9999 AS ORDR_QTY,")
            sqls.AppendLine("   C2.COLOR_DESC AS COLOR_CODE_LONG,")
            sqls.AppendLine("   C1.STYLE_COLOR_STATUS,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) < 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE")
            sqls.AppendLine("     WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END)")
            sqls.AppendLine("   END AS MSOH,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'MS'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) <= 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("     CASE WHEN")
            sqls.AppendLine("       SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'MS'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) < 0")
            sqls.AppendLine("     THEN")
            sqls.AppendLine("       0")
            sqls.AppendLine("     ELSE")
            sqls.AppendLine("     SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'MS'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) END")
            sqls.AppendLine("   END AS MSFT,")
            sqls.AppendLine(" CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) < 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE")
            sqls.AppendLine("     WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END)")
            sqls.AppendLine("   END AS SWOH,")
            sqls.AppendLine("   CASE WHEN")
            sqls.AppendLine("   SUM(")
            sqls.AppendLine("     CASE S2.WHSE_CODE WHEN 'SW'")
            sqls.AppendLine("     THEN (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("     ELSE 0")
            sqls.AppendLine("     END) <= 0")
            sqls.AppendLine("   THEN")
            sqls.AppendLine("     0")
            sqls.AppendLine("   ELSE")
            sqls.AppendLine("     CASE WHEN")
            sqls.AppendLine("       SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'SW'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) < 0")
            sqls.AppendLine("     THEN")
            sqls.AppendLine("       0")
            sqls.AppendLine("     ELSE")
            sqls.AppendLine("     SUM(")
            sqls.AppendLine("       CASE S2.WHSE_CODE")
            sqls.AppendLine("       WHEN 'SW'")
            sqls.AppendLine("       THEN (NVL(S2.WHSE_QTY_ON_HAND,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))")
            sqls.AppendLine("       ELSE 0")
            sqls.AppendLine("       END) END")
            sqls.AppendLine("   END AS SWFT")
            sqls.AppendLine("   FROM ICTSTYC1 C1")
            sqls.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            sqls.AppendLine("   ON C1.STYLE_CODE  = S2.STYLE_CODE")
            sqls.AppendLine("   AND C1.COLOR_CODE = S2.COLOR_CODE")
            sqls.AppendLine("   INNER JOIN ICTCOLR1 C2")
            sqls.AppendLine("   ON C1.COLOR_CODE = C2.COLOR_CODE")
            sqls.AppendLine("   GROUP BY C1.STYLE_CODE, C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
            sqls.AppendLine("  )")
            ASCMAIN1.sql = sqls.ToString
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
            Fill_Records("ICTSTYC1")

            Create_TDA(.Tables.Add, "WBTIMGR1", "*")

            ICTPEND_TEMP = MAKE_ICTPEND_TEMP()

            sqls.Length = 0
            sqls.AppendLine("SELECT")
            sqls.AppendLine("'23249' AS Supplier_ID,")
            sqls.AppendLine("(S1.STYLE_CODE || '-' || S2.COLOR_CODE) AS Item_Number,")
            sqls.AppendLine("CASE WHEN SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0) - NVL(P1.ORDR_QTY_PEND,0))) < 4 THEN")
            sqls.AppendLine("  0")
            sqls.AppendLine("ELSE")
            sqls.AppendLine("  SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0)  - NVL(P1.ORDR_QTY_PEND,0)))")
            sqls.AppendLine("END AS On_Hand,")
            sqls.AppendLine("0 AS Back_Order,")
            sqls.AppendLine("0 AS On_Order,")
            sqls.AppendLine("NULL AS NxtAvailDate,")
            sqls.AppendLine("CASE WHEN SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0)  - NVL(P1.ORDR_QTY_PEND,0))) < 4 THEN")
            sqls.AppendLine("  1")
            sqls.AppendLine("ELSE")
            sqls.AppendLine("  0")
            sqls.AppendLine("END AS STATUS,")
            sqls.AppendLine("S1.STYLE_DESC AS DESCRIPTION")
            sqls.AppendLine(String.Format("FROM ICTSTYL1 S1, ICTSTAT2 S2, ECTESTY1 E1, ECTESTY2 E2, {0} P1", ICTPEND_TEMP))
            sqls.AppendLine("WHERE E1.STYLE_CODE = E2.STYLE_CODE")
            sqls.AppendLine("AND NVL(E1.SHIP_DROP,'0') = '1'")
            sqls.AppendLine("AND S1.STYLE_CODE = S2.STYLE_CODE (+)")
            sqls.AppendLine("AND E2.STYLE_CODE = P1.STYLE_CODE (+)")
            sqls.AppendLine("AND E2.COLOR_CODE = P1.COLOR_CODE (+)")
            sqls.AppendLine("AND S2.STYLE_CODE = E2.STYLE_CODE (+)")
            sqls.AppendLine("AND S2.COLOR_CODE = E2.COLOR_CODE (+)")
            sqls.AppendLine("AND E2.ECOM_STYLE_COLOR_STATUS = 'A'")
            sqls.AppendLine("AND S2.WHSE_CODE = 'MS'")
            sqls.AppendLine("AND E1.ECOM_CODE = 'WAYFAIR'") 'This is HC'd until EDI Takes Over.
            sqls.AppendLine("GROUP BY (S1.STYLE_CODE || '-' || S2.COLOR_CODE),")
            sqls.AppendLine("S1.STYLE_DESC")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 2)

            sqls.Length = 0
            sqls.AppendLine("SELECT * FROM WBTSTYL2")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "WBTSTYL3", "**", 0, False, "", 2)

        End With

        grdWBTSTYL1.DataSource = dst.Tables("WBTSTYL1")
        grdWBTSTYL2.DataSource = dst.Tables("WBTSTYL2")
        grdWBTSHOP1.DataSource = dst.Tables("WBTSHOP1")
        grdWBTSHOP2.DataSource = dst.Tables("WBTSHOP2")
        grdMissing.DataSource = dst.Tables("WBTSTYL3")

        MoveToParents()

        UltraTabControl1.Visible = False

        Get_PARM("WBTPARM1")

        WBCMAIN1.DisplayHeaderCheckBox(grdWBTSTYL1, New String() {"WEB_IND"})

        Create_Summary(grdWBTSTYL1, "STYLE_CODE", "Count")
        Create_Summary(grdWBTSTYL1, "WEB_IND", "Sum")

        Create_Summary(grdWBTSHOP1, "STYLE_CODE", "Count")
        Create_Summary(grdWBTSHOP1, "QUANTITY_ON_HAND", "Sum")

        Create_Summary(grdWBTSHOP2, "STYLE_CODE", "Count")
        Create_Summary(grdWBTSHOP2, "QUANTITY_ON_HAND", "Sum")

        WB_PARM_SITE_IP = (ROWs("WBTPARM1").Item("WB_PARM_SITE_IP") & String.Empty).ToString.Trim
        WB_PARM_SITE_USER = (ROWs("WBTPARM1").Item("WB_PARM_SITE_USER") & String.Empty).ToString.Trim
        WB_PARM_SITE_PWD = (ROWs("WBTPARM1").Item("WB_PARM_SITE_PWD") & String.Empty).ToString.Trim

        WB_PARM_SITE_OUTPUT_DIR = (ROWs("WBTPARM1").Item("WB_PARM_SITE_OUTPUT_DIR") & String.Empty).ToString.Trim
        WB_PARM_SITE_PRODUCT_POST_URL = (ROWs("WBTPARM1").Item("WB_PARM_SITE_PRODUCT_POST_URL") & String.Empty).ToString.Trim
        WB_PARM_SITE_ORDERS_POST_URL = (ROWs("WBTPARM1").Item("WB_PARM_SITE_ORDERS_POST_URL") & String.Empty).ToString.Trim
        WB_PARM_SITE_PRODUCT_PUB_URL = (ROWs("WBTPARM1").Item("WB_PARM_SITE_PRODUCT_PUB_URL") & String.Empty).ToString.Trim

        WB_PARM_PRODUCTS_DIR = (ROWs("WBTPARM1").Item("WB_PARM_PRODUCTS_DIR") & String.Empty).ToString.Trim
        WB_PARM_INVENTORY_DIR = (ROWs("WBTPARM1").Item("WB_PARM_INVENTORY_DIR") & String.Empty).ToString.Trim
        WB_PARM_ORDERS_DIR = (ROWs("WBTPARM1").Item("WB_PARM_ORDERS_DIR") & String.Empty).ToString.Trim
        WB_PARM_MASTER_IMAGES = (ROWs("WBTPARM1").Item("WB_PARM_MASTER_IMAGES") & String.Empty).ToString.Trim
        WB_PARM_SITE_IMAGES_DIR = (ROWs("WBTPARM1").Item("WB_PARM_SITE_IMAGES_DIR") & String.Empty).ToString.Trim
        WB_PARM_IMAGES_DIR = (ROWs("WBTPARM1").Item("WB_PARM_IMAGES_DIR") & String.Empty).ToString.Trim

        WB_NEW_PROD_MAX = Val(ROWs("WBTPARM1").Item("WB_NEW_PROD_MAX") & String.Empty)
        WB_NEW_PROD_PAGE_CODE = (ROWs("WBTPARM1").Item("WB_NEW_PROD_PAGE_CODE") & String.Empty).ToString.Trim

        WB_PARM_RSS_MAX_ENTRIES = Convert.ToInt16(ROWs("WBTPARM1").Item("WB_PARM_RSS_MAX_ENTRIES") & String.Empty)
        WB_PARM_RSS_NEW_PAGE = (ROWs("WBTPARM1").Item("WB_PARM_RSS_NEW_PAGE") & String.Empty).ToString.Trim
        WB_PARM_RSS_SALE_PAGE = (ROWs("WBTPARM1").Item("WB_PARM_RSS_SALE_PAGE") & String.Empty).ToString.Trim

        If Not WB_PARM_PRODUCTS_DIR.EndsWith("\") Then WB_PARM_PRODUCTS_DIR &= "\"
        If Not WB_PARM_INVENTORY_DIR.EndsWith("\") Then WB_PARM_INVENTORY_DIR &= "\"
        If Not WB_PARM_ORDERS_DIR.EndsWith("\") Then WB_PARM_ORDERS_DIR &= "\"
        If Not WB_PARM_MASTER_IMAGES.EndsWith("\") Then WB_PARM_MASTER_IMAGES &= "\"
        If Not WB_PARM_SITE_IMAGES_DIR.EndsWith("/") Then WB_PARM_SITE_IMAGES_DIR &= "/"
        If Not WB_PARM_IMAGES_DIR.EndsWith("\") Then WB_PARM_IMAGES_DIR &= "\"

        Dim NextGroup = Convert.ToInt16(ROWs("WBTPARM1").Item("WB_NEXT_GROUP") & String.Empty)
        If NextGroup = 0 Then
            NextGroup = 1
        End If
        txtAutoWait.Text = "30:00"
        txtNextGroup.Text = NextGroup

        Dim SQLU As New System.Text.StringBuilder() With {.Length = 0}
        SQLU.AppendLine("SELECT MAX(STYLE_GROUP) AS RECCNT")
        SQLU.AppendLine("FROM WBTSTYL1")
        SQLU.AppendLine("WHERE STYLE_GROUP < 999")
        ASCMAIN1.sql = SQLU.ToString()
        Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)

        txtMaxGroup.Text = RECCNT
    End Sub

    Private Function MAKE_ICTPEND_TEMP() As String
        Dim RETVAL As String = ""
        Dim s As New StringBuilder With {.Length = 0}
        s.AppendLine("SELECT STYLE_CODE, COLOR_CODE, SUM(ORDR_QTY_PEND) AS ORDR_QTY_PEND")
        s.AppendLine("FROM")
        s.AppendLine("(")
        s.AppendLine("  SELECT L2.STYLE_CODE,")
        s.AppendLine("  L2.COLOR_CODE,")
        s.AppendLine("  SUM(L2.ORDR_QTY_OPEN) ORDR_QTY_PEND")
        s.AppendLine("  FROM SOTORDR1_L L1,")
        s.AppendLine("  SOTORDR2_L L2")
        s.AppendLine("  WHERE L1.ORDR_NO   = L2.ORDR_NO")
        s.AppendLine("  AND L1.ORDR_STATUS = 'O'")
        s.AppendLine("  GROUP BY L2.STYLE_CODE,")
        s.AppendLine("  L2.COLOR_CODE")
        s.AppendLine("  UNION")
        s.AppendLine("  SELECT EDI_STYLE AS STYLE_CODE,")
        s.AppendLine("  EDI_COLOR_CODE AS COLOR_CODE,")
        s.AppendLine("  SUM(EDI_TOTAL_QTY) AS ORDR_QTY_PEND")
        s.AppendLine("  FROM EDTTRPM1 P1,")
        s.AppendLine("  EDT850T1 T1,")
        s.AppendLine("  EDT850T2 T2")
        s.AppendLine("  WHERE T1.EDI_DOC_SEQ_NO = T2.EDI_DOC_SEQ_NO")
        s.AppendLine("  AND P1.EDI_TP_ID = T1.EDI_TP_ID")
        s.AppendLine("  AND P1.EDI_DOC_NO = '850'")
        s.AppendLine("  AND P1.EDI_STATUS = 'P'")
        s.AppendLine("  AND NVL(T1.EDI_PROCESS_IND, '0') = '0'")
        s.AppendLine("  AND NVL(EDI_STYLE,'NULL') <> 'NULL'")
        s.AppendLine("  GROUP BY EDI_STYLE,")
        s.AppendLine("  EDI_COLOR_CODE")
        s.AppendLine(") GROUP BY STYLE_CODE, COLOR_CODE")
        ASCMAIN1.sql = s.ToString
        RETVAL = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RETVAL & " Add Primary Key (STYLE_CODE, COLOR_CODE)")
        Return RETVAL
    End Function

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim zmsg As String = String.Empty

        Select Case eItemKey

            Case "Create Product Exp"
                EMsg = CheckIsAutoMode()
                MyBase.Absx1.txtFor("STYLE_CODE").Text = MyBase.Absx1.txtFor("STYLE_CODE").Text.Trim
                If MyBase.Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
                    Validate_Code("STYLE_CODE")
                Else
                    If dst.Tables("WBTSTYL1").Rows.Count = 0 Then
                        EMsg = "There are no changes to Styles since the last Upload"
                    End If
                End If
                Dim ZeroUploads As Boolean = False
                For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select()
                    If rowWBTSTYL1.Item("WEB_IND") & "" = "1" And rowWBTSTYL1.Item("UPLOAD_BATCH") & "" = "" Then
                        ZeroUploads = True
                        For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select(String.Format("STYLE_CODE = '{0}'", rowWBTSTYL1.Item("STYLE_CODE")))
                            If rowWBTSTYL2.Item("COLOR_STATUS").ToString = "A" Then
                                ZeroUploads = False
                                Exit For
                            Else
                                If Val(rowWBTSTYL2.Item("MSOH").ToString) > 0 Or Val(rowWBTSTYL2.Item("MSFT").ToString) > 0 Then
                                    ZeroUploads = False
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                If ZeroUploads Then
                    EMsg = EMsg & vbCrLf & "Non-Active Items With No Inventory Selected."
                End If
            Case "Upload"
                EMsg = CheckIsAutoMode()
                If EMsg.Length = 0 Then
                    zmsg = "Do you want to Upload the Item Changes?"

                    If EMsg.Length = 0 Then
                        If MessageBox.Show(zmsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                    Dim iResult As MsgBoxResult
                    Dim iTitle As String = "Upload Images"
                    Dim iMSG As New System.Text.StringBuilder

                    If UpLoadType = "P" Then
                        iMSG.AppendLine("Do You Want To Upload Images?")
                        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                        If iResult = MsgBoxResult.Yes Then
                            FTPImages = True
                        Else
                            FTPImages = False
                        End If
                    Else
                        FTPImages = False
                    End If

                    If UpLoadType = "I" Then
                        FTPTables = True
                    Else
                        FTPTables = False
                    End If
                End If

            Case "Create Inventory Exp"

            Case "Done"
                zmsg = "Do you want to Cancel uploading the Item Changes?"

                If zmsg.Length > 0 Then
                    If MessageBox.Show(zmsg, "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If
            Case "Folder Select"

            Case "Compare Shopsite"

            Case "Remove Unwanted"

            Case "Load Records"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Load Records"
                Dim iMSG As New System.Text.StringBuilder
                iMSG.AppendLine("This Will Clear Any Existing Records In Memory")
                iMSG.AppendLine("And Re-Load The Data From The Database.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult <> MsgBoxResult.Yes Then
                    EMsg = EMsg & vbCrLf & "Load Records Cancelled By User."
                End If
            Case "Send Wayfair"
                If chkAutoRefresh.Checked Then
                    EMsg = EMsg & vbCrLf & "You Can Not Send Wayfair While Auto-Refesh Is Running."
                End If
        End Select

        If EMsg <> String.Empty Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Create Product Exp"
                Update_Record(False)
                styleList.Clear()
                styleListInactive.Clear()
                EntryMode = "E"
                Call CreateProductXml()
                Call Mode_Settings(True)
                UpLoadType = "P"
            Case "Create Inventory Exp"
                EntryMode = "I"
                Call Mode_Settings(True)
                UpLoadType = "I"
                Dim FullUpload As Boolean = False
                Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT MAX(STYLE_GROUP) FROM WBTSTYL1 WHERE NVL(STYLE_GROUP,999) <> 999")
                ASCMAIN1.sql = SQLS.ToString()
                Dim STYLE_GROUP As Int16 = Val(ASCDATA1.GetDataValue)

                Dim filter As String = String.Format("STYLE_GROUP = {0}", Val(STYLE_GROUP))
                Dim UpLoadCount As Integer = CreateInventoryXml(1, 100000, filter, False)
                If UpLoadCount > 0 Then
                    SendErrorEMail(String.Format("Regency Update of {0} Items at {1} Began. Including Group {2}", UpLoadCount, Format(Now(), "hh:mm tt"), Val(txtNextGroup.Text)), False)
                    WebBrowser1.Visible = True
                    grdWBTSTYL1.Visible = False
                    grdWBTSTYL2.Visible = False
                    FTPTables = True
                    Call ftpProducts()
                    FTPTables = False
                    SendWayfair()
                    WebBrowser1.Visible = False
                    grdWBTSTYL1.Visible = True
                    grdWBTSTYL2.Visible = True
                    SendErrorEMail(String.Format("Regency Update of {0} Items at {1} Completed.  Including Group {2}", UpLoadCount, Format(Now(), "hh:mm tt"), Val(txtNextGroup.Text)), False)
                Else
                    SendErrorEMail(String.Format("Regency Update Ran at {0} But Had No Inventory To Update.", Format(Now(), "hh:mm tt")), False)
                End If

                Update_Record()
                Call Mode_Settings(False)
            Case "Upload"
                'If EntryMode = "E" Then
                Call ftpProducts()
                'Clear_Record()
                'ElseIf EntryMode = "I" Then
                '    If FTPProducts() Then
                '        Update_Record()
                '    End If
                'End If

                Call Mode_Settings(False)
            Case "Batch 500"
                Dim MaxBatch As Integer = 500
                Dim CurBatch As Integer = 0
                Dim SQLS As New StringBuilder With {.Length = 0}
                SQLS.AppendLine("SELECT NVL(MAX(STYLE_GROUP),0) + 1 FROM WBTSTYL1")
                ASCMAIN1.sql = SQLS.ToString()
                Dim STYLE_GROUP As Int16 = Val(ASCDATA1.GetDataValue)
                For Each grow As UltraWinGrid.UltraGridRow In grdWBTSTYL1.Rows
                    If grow.VisibleIndex <> -1 Then
                        If grow.Cells.Item("WEB_IND").Value <> "1" Then
                            Dim OKToUpload As Boolean = False
                            Dim rowWBTSTYL1 As DataRow = dst.Tables("WBTSTYL1").Select(String.Format("STYLE_CODE = '{0}'", grow.Cells.Item("STYLE_CODE").Text)).FirstOrDefault()
                            If rowWBTSTYL1.Item("STYLE_STATUS").ToString = "A" Then
                                OKToUpload = True
                            Else
                                If Val(rowWBTSTYL1.Item("CURR_ON_HAND").ToString) > 0 Then
                                    OKToUpload = True
                                End If
                            End If
                            If OKToUpload = True Then
                                CurBatch += 1
                                rowWBTSTYL1.Item("WEB_IND") = "1"
                                rowWBTSTYL1.Item("STYLE_GROUP") = STYLE_GROUP
                            End If
                            If CurBatch > MaxBatch Then
                                Exit For
                            End If
                        End If
                    End If
                Next
                grdWBTSTYL1.UpdateData()
                grdWBTSTYL1.Refresh()
                Update_Record()
            Case "Done"
                Update_Record()
                Call Mode_Settings(False)
            Case "Update"
                Update_Record()
                'Clear_Record()
            Case "Folder Select"
                FolderSelect()
            Case "Compare Shopsite"
                CompareShopsite()
            Case "Remove Unwanted"
                RemoveUnwanted()
            Case "Load Records"
                Clear_Record()
                Load_Record()
                Call Mode_Settings(False)
            Case "Nextopia"
                updateNextopia()
                MsgBox("Nextopia Is All Sent", vbOKOnly, "Nextopia Feed")
            Case "Send Wayfair"
                Fill_Records("ICTSTATX")
                SendWayfair()
                MsgBox("Wayfair Feed Is All Sent", vbOKOnly, "Wayfair Feed")
        End Select

    End Sub

    Private Sub updateNextopia()
        'Dim DT As New DataTable
        Dim NForm As New WBCNXTPA(dst.Tables.Item("WBTSTYL1"))
        If File.Exists("\\192.168.110.224\product\other\NEXTOPIA.xml") Then
            File.Delete("\\192.168.110.224\product\other\NEXTOPIA.xml")
        End If
        NForm.XmlDoc.Save("\\192.168.110.224\product\other\NEXTOPIA.xml")

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)


        With UltraExplorerBar1.Groups("Screen Control")
            'Load Records
            .Items("Load Records").Settings.Enabled = not_iScreenMode
            'Create Product Exp
            .Items("Create Product Exp").Settings.Enabled = not_iScreenMode
            'Create Inventory Exp
            .Items("Create Inventory Exp").Settings.Enabled = not_iScreenMode
            'Upload
            .Items("Upload").Settings.Enabled = iScreenMode
            'Remove Unwanted
            .Items("Remove Unwanted").Settings.Enabled = not_iScreenMode
            'Batch 500
            .Items("Batch 500").Settings.Enabled = not_iScreenMode
            'Folder Select
            .Items("Folder Select").Settings.Enabled = not_iScreenMode
            'Update
            .Items("Update").Settings.Enabled = not_iScreenMode
            'Done
            .Items("Done").Settings.Enabled = iScreenMode
            'Send Wayfair
            .Items("Send Wayfair").Settings.Enabled = not_iScreenMode
        End With

        UltraExplorerBar1.Groups("Auto Refresh").Visible = dst.Tables("WBTSTYL1").Rows.Count() > 0

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            WebBrowser1.Visible = True
            grdWBTSTYL1.Visible = False
            grdWBTSTYL2.Visible = False
            grdWBTSHOP1.Visible = False
            grdWBTSHOP2.Visible = False
            grdMissing.Visible = False
        Else
            WebBrowser1.Visible = False
            grdWBTSTYL1.Visible = True
            grdWBTSTYL2.Visible = True
            grdWBTSHOP1.Visible = True
            grdWBTSHOP2.Visible = True
            grdMissing.Visible = True
        End If

        With grdWBTSTYL1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdWBTSTYL1.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTSTYL1.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"WEB_IND", "UPLOAD_BATCH", "STYLE_SORT", "UPLOAD_IMG"}
            grdWBTSTYL1.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In New String() {"UPLOAD_BATCH", "STYLE_SORT"}
            grdWBTSTYL1.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next

        With grdWBTSTYL2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        With grdMissing.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdWBTSTYL2.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTSTYL2.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For i As Integer = 0 To grdMissing.DisplayLayout.Bands(0).Columns.Count - 1
            grdMissing.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i

    End Sub

    Sub Clear_Record()

        dst.Tables("ICTSTYC1").Rows.Clear()
        dst.Tables("WBTSTYL1").Rows.Clear()
        dst.Tables("WBTSTYL2").Rows.Clear()
        dst.Tables("WBTRSSF1").Rows.Clear()
        dst.Tables("ICTSTATX").Rows.Clear()

        Absx1.txtFor("STYLE_CODE").Clear()
        Absx1.txtFor("STYLE_DESC").Clear()

        'UpdateWBTSTYL2()

        shopSiteFilename = String.Empty
        itemUploaded = False

    End Sub

    Sub Load_Record()

        Fill_Records("ICTSTYC1")

        Fill_Records("ICTSTATX")

        Fill_Records("WBTSHOP1")
        Fill_Records("WBTSHOP2")

        Dim SQLW As New StringBuilder
        SQLW.Length = 0
        SQLW.AppendLine("SELECT WBTSTYL1.*, ICTSTYL1.STYLE_CLASS_CODE")
        SQLW.AppendLine("FROM WBTSTYL1, ICTSTYL1")
        SQLW.AppendLine("WHERE WBTSTYL1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        Fill_Records("WBTSTYL1", , , SQLW.ToString)
        Fill_Records("WBTSTYL2", , , "SELECT * FROM WBTSTYL2")
        UpdateWBTSTYLS()
        ShowPrevious(chkShowPrevious.Checked)
        ShowVariance(ckbINVVariance.Checked)
        'MsgBox("Data Loaded.", MsgBoxStyle.Information, "Success")

        Setup_WBTSHOP2()
        Setup_Missing()
    End Sub

    Sub Update_Record(Optional ByVal ShowMsg As Boolean = True)
        Dim MsgShow As String = ""

        If Not InAutoMode Then
            If ShowMsg Then
                MsgShow = "Update Complete"
            End If
        End If


        Try

            MyBase.BeginTrans()

            'If EntryMode = "I" Then
            Update_Record_TDA("WBTSTYL1")
            Update_Record_TDA("WBTSTYL2")

            Dim SQLS As New System.Text.StringBuilder

            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM ICTSTYLW")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO ICTSTYLW")
            SQLS.AppendLine("SELECT DISTINCT ICTSTYL1.STYLE_CODE,")
            SQLS.AppendLine("ICTSTYL1.STYLE_CLASS_CODE")
            SQLS.AppendLine("FROM WBTSTYL1, ICTSTYL1")
            SQLS.AppendLine("WHERE WBTSTYL1.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQLS.AppendLine("AND NVL(ICTSTYL1.STYLE_CLASS_CODE,'NULL') <> 'NULL'")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            'End If

            MyBase.CommitTrans(MsgShow)

        Catch ex As Exception
            If InAutoMode Then
                MyBase.Rollback()
                Exit Sub
            Else
                MyBase.Rollback(ex.Message)
            End If

        End Try
    End Sub

    Public Overrides Sub txt_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_Leave(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                MyBase.Absx1.txtFor("STYLE_CODE").Text = MyBase.Absx1.txtFor("STYLE_CODE").Text.Trim.ToUpper
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub btnResetGroups_Click(sender As System.Object, e As System.EventArgs) Handles btnResetGroups.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Re-Set Groups"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("This Will CLEAR and Re-Set All Items Groups!")
        iMSG.AppendLine("Are You Sure You Know What You Are Doing")
        iMSG.AppendLine("And This Is What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then

            Dim GroupCount As Integer = InputBox("How Many Items Per Group Do You Want?", "Group Qty", 200)
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("SELECT (COUNT(*) / {0}) AS RECCNT", GroupCount))
            SQLS.AppendLine("FROM WBTSTYL1")
            ASCMAIN1.sql = SQLS.ToString()
            Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)
            RECCNT += 1
            txtMaxGroup.Text = RECCNT

            BeginTrans()
            Dim SQLS1 As New System.Text.StringBuilder
            SQLS1.Length = 0
            SQLS1.AppendLine("UPDATE WBTSTYL1")
            SQLS1.AppendLine("SET STYLE_GROUP = 999")
            ASCMAIN1.sql = SQLS1.ToString
            ASCDATA1.ExecuteSQL()
            For i As Integer = 1 To RECCNT
                Dim SQLS2 As New System.Text.StringBuilder
                SQLS2.Length = 0
                SQLS2.AppendLine("UPDATE WBTSTYL1")
                SQLS2.AppendLine("SET STYLE_GROUP = " & i)
                SQLS2.AppendLine("WHERE STYLE_GROUP = 999")
                SQLS2.AppendLine("AND WEB_IND = '1'")
                SQLS2.AppendLine("AND ROWNUM <= " & GroupCount)
                ASCMAIN1.sql = SQLS2.ToString
                ASCDATA1.ExecuteSQL()
            Next
            CommitTrans()
            MsgBox("Finished Re-Set.", MsgBoxStyle.OkOnly, "Re-Set Complete.")
        Else
            MsgBox("Chicken!", MsgBoxStyle.OkOnly, "Re-Set Aborted")
        End If

    End Sub

    Private Sub chkAutoRefresh_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAutoRefresh.CheckedChanged
        'If Not (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        '    MsgBox("Sorry.   Wayne Did Not Enable This Feature Yet", MsgBoxStyle.Critical, "Tell Wayne To Hurry Up!")
        '    Exit Sub
        'End If

        If chkAutoRefresh.Checked Then
            Dim NoNullsFound As Boolean = CheckForNullGroups()
            If NoNullsFound Then
                dtpNextSync.Text = Now()
                TickCount = 0
                AutoSync(True)
            Else
                chkAutoRefresh.Checked = False
            End If
        Else
            AutoSync(False)
            AutoProcessRunning = False
        End If
    End Sub

    Private Sub chkShowPrevious_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowPrevious.CheckedChanged
        ShowPrevious(chkShowPrevious.Checked)
    End Sub

    Private Sub ckbINVVariance_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbINVVariance.CheckedChanged
        ShowVariance(ckbINVVariance.Checked)
    End Sub

    Private Sub grdWBTSTYL1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWBTSTYL1.AfterRowActivate
        Setup_WBTSTYL2()
    End Sub

    Private Sub tmrAutoSync_Tick(sender As System.Object, e As System.EventArgs) Handles tmrAutoSync.Tick
        Dim sleepSec As Integer = (Val(txtAutoWait.Text.Split(":")(1))) + (Val(txtAutoWait.Text.Split(":")(0)) * 60)

        If InAutoMode Then
            If CDate(Now().ToShortTimeString) > CDate(CDate(dtpNextSync.Text).ToShortTimeString) Then
                'Auto Shutdown at 9PM now.  Per Wayne.
                If Now.TimeOfDay.Hours > 21 Then
                    SendErrorEMail("It's After 9PM And The Regency Inventory Process Has Been Working Hard All Day.  Going To Sleep Now.  Don't Forget To Wake Me Back Up In The Morning.", True)
                    'Update_Record()
                    Application.DoEvents()
                    Application.Exit()
                End If

                dtpNextSync.Text = Now().AddSeconds(sleepSec)
                TickCount = 0
                Application.DoEvents()
                Update_Record()
                Application.DoEvents()
                Clear_Record()
                Application.DoEvents()
                Load_Record()
                Application.DoEvents()
                MakeSendZipFile()

                Dim filter As String = String.Format("(LAST_ON_HAND <> CURR_ON_HAND) OR (NVL(STYLE_GROUP,0) = {0}')", Val(txtNextGroup.Text))
                'Dim filter As String = String.Format("(NVL(STYLE_GROUP,0) = {0}')", Val(txtNextGroup.Text))
                Dim UpLoadCount As Integer = CreateInventoryXml(1, 100000, filter, False)
                If UpLoadCount > 0 Then
                    SendErrorEMail(String.Format("Regency Update of {0} Items at {1} Began. Including Group {2}", UpLoadCount, Format(Now(), "hh:mm tt"), Val(txtNextGroup.Text)), False)
                    WebBrowser1.Visible = True
                    grdWBTSTYL1.Visible = False
                    grdWBTSTYL2.Visible = False
                    FTPTables = True
                    Call ftpProducts()
                    FTPTables = False
                    WebBrowser1.Visible = False
                    grdWBTSTYL1.Visible = True
                    grdWBTSTYL2.Visible = True
                    SendErrorEMail(String.Format("Regency Update of {0} Items at {1} Completed.  Including Group {2}", UpLoadCount, Format(Now(), "hh:mm tt"), Val(txtNextGroup.Text)), False)
                Else
                    SendErrorEMail(String.Format("Regency Update Ran at {0} But Had No Inventory To Update.", Format(Now(), "hh:mm tt")), False)
                End If

                If Val(txtNextGroup.Text) = Val(txtMaxGroup.Text) Then
                    txtNextGroup.Text = 1
                Else
                    txtNextGroup.Text = Val(txtNextGroup.Text) + 1
                End If
                Dim SQLS As New System.Text.StringBuilder
                SQLS.Length = 0
                SQLS.AppendLine("UPDATE WBTPARM1 SET WB_NEXT_GROUP = " & Val(txtNextGroup.Text))
                ASCMAIN1.sql = SQLS.ToString
                ASCDATA1.ExecuteSQL()

                If chkWayfair.Checked Then
                    SendWayfair()
                End If

                If chkNextopia.Checked Then
                    updateNextopia()
                End If
            Else
                    TickCount += 1
            End If
        Else
            tmrAutoSync.Stop()
        End If
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        docComplete = True
    End Sub

#End Region

#Region "Custom Methods"

    Private Sub AutoSync(StartSync As Boolean)
        If StartSync Then
            With UltraExplorerBar1.Groups("Screen Control")
                .Items("Load Records").Settings.Enabled = DefaultableBoolean.False
                .Items("Create Product Exp").Settings.Enabled = DefaultableBoolean.False
                .Items("Create Inventory Exp").Settings.Enabled = DefaultableBoolean.False
                .Items("Upload").Settings.Enabled = DefaultableBoolean.False
                .Items("Remove Unwanted").Settings.Enabled = DefaultableBoolean.False
                .Items("Batch 500").Settings.Enabled = DefaultableBoolean.False
                .Items("Folder Select").Settings.Enabled = DefaultableBoolean.False
                .Items("Update").Settings.Enabled = DefaultableBoolean.False
                .Items("Done").Settings.Enabled = DefaultableBoolean.False
                .Items("Send Wayfair").Settings.Enabled = DefaultableBoolean.False
            End With
            InAutoMode = True
            tmrAutoSync.Start()
        Else
            InAutoMode = False
            Mode_Settings(False)
        End If
    End Sub

    Private Sub BuildFTPFile()
        Dim NewLine As String = ""
        Dim FileForStream As String = "Supplier_ID,Item_Number,On_Hand,Back_Order,On_Order,NxtAvailDate,STATUS,Description" & vbCrLf
        For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select()
            'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "whr") Then
            '    If rowICTSTATX.Item("Item_Number").ToString() = "MTX37196" Then Stop
            'End If
            NewLine = ""
            NewLine = String.Format("{0},", rowICTSTATX.Item("Supplier_ID"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("Item_Number"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("On_Hand"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("Back_Order"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("On_Order"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("NxtAvailDate"))
            NewLine += String.Format("{0},", rowICTSTATX.Item("STATUS"))
            NewLine += CleanText(rowICTSTATX.Item("Description").ToString & "") & vbCrLf
            FileForStream += NewLine.ToString
        Next
        Dim fs As FileStream = File.Create(TransferFile)

        Dim info As Byte() = New UTF8Encoding(True).GetBytes(FileForStream)
        fs.Write(info, 0, info.Length)
        fs.Close()
    End Sub

    Private Function CheckIsAutoMode() As String
        Dim RetVal As String = ""
        If InAutoMode Then
            RetVal = "You May Not Run This While Auto-Refresh Is On."
        End If
        Return RetVal
    End Function

    Public Shared Function CleanText(ByVal TextIn As String) As String
        TextIn = Replace(TextIn, ",", "")
        TextIn = Replace(TextIn, "'", "")
        TextIn = Replace(TextIn, Chr(34).ToString, "")
        Return Trim(TextIn)
    End Function

    Private Sub CompareShopsite()
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Select File Downloaded From ShopSite"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("Have You Downloaded The Products And Pages")
        iMSG.AppendLine("Files From ShopSite In XML (.xml) Format")
        iMSG.AppendLine("And Know What Folder They Are In?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Dim fList As New List(Of String)
        Dim fDialog As New OpenFileDialog
        With fDialog
            .CheckFileExists = True
            .CheckPathExists = True
            .DefaultExt = ".csv"
            .Title = "First Select The Pages File Downloaded From ShopSite"
            .Filter = "XML (*.xml) |*.xml"
        End With
        fDialog.ShowDialog()
        Dim PagesFile As String = fDialog.FileName

        With fDialog
            .CheckFileExists = True
            .CheckPathExists = True
            .DefaultExt = ".csv"
            .Title = "Next Select The Products File Downloaded From ShopSite"
            .Filter = "XML (*.xml) |*.xml"
        End With
        fDialog.ShowDialog()
        Dim ProductsFile As String = fDialog.FileName

        If ProductsFile.Length = 0 Or PagesFile.Length = 0 Then
            MsgBox("Invalid Or Missing File Name Provided!", MsgBoxStyle.OkOnly, "Files Names")
            Exit Sub
        Else
            Dim SD As New System.Text.StringBuilder() With {.Length = 0}
            SD.AppendLine("DELETE FROM WBTSHOP1")
            ASCMAIN1.sql = SD.ToString
            ASCDATA1.ExecuteSQL()

            SD.Length = 0
            SD.AppendLine("DELETE FROM WBTSHOP2")
            ASCMAIN1.sql = SD.ToString
            ASCDATA1.ExecuteSQL()

            dst.Tables.Item("WBTSHOP1").Clear()
            dst.Tables.Item("WBTSHOP2").Clear()

            Dim Products_doc As XmlDocument = New XmlDocument
            Dim Product_node_list As XmlNodeList
            Try
                Products_doc.Load(ProductsFile)
            Catch ex As Exception
                MsgBox(ex.Message.ToString, MsgBoxStyle.Critical, ProductsFile)
                Stop
            End Try
            Product_node_list = Products_doc.GetElementsByTagName("Product")
            For Each Product_node As System.Xml.XmlNode In Product_node_list
                'Debug.Print(Product_node.ChildNodes.Count)
                Dim STYLE_CODE As String = ""
                Dim newWBTSHOP1 As DataRow = dst.Tables("WBTSHOP1").NewRow
                newWBTSHOP1.Item("STYLE_STATUS") = "A"
                newWBTSHOP1.Item("WEB_CATEGORY") = ""
                newWBTSHOP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                newWBTSHOP1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                For Each Product_node_child As System.Xml.XmlNode In Product_node

                    Select Case Product_node_child.Name
                        Case Is = "Name"
                            newWBTSHOP1.Item("STYLE_FULL_DESC") = Product_node_child.InnerText
                        Case Is = "SKU"
                            newWBTSHOP1.Item("STYLE_CODE") = Product_node_child.InnerText
                            STYLE_CODE = Product_node_child.InnerText
                        Case Is = "QuantityOnHand"
                            newWBTSHOP1.Item("QUANTITY_ON_HAND") = Val(Product_node_child.InnerText & "")
                        Case Is = "Graphic"
                            newWBTSHOP1.Item("IMAGE") = Product_node_child.InnerText
                        Case Is = "ProductOptions"
                            For Each ProductOptions_node_child As System.Xml.XmlNode In Product_node_child
                                If ProductOptions_node_child.Name = "ProductOption" Then
                                    Dim newWBTSHOP2 As DataRow = dst.Tables("WBTSHOP2").NewRow
                                    newWBTSHOP2.Item("STYLE_CODE") = STYLE_CODE
                                    newWBTSHOP2.Item("COLOR_STATUS") = "A"
                                    newWBTSHOP2.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                    newWBTSHOP2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                                    For Each ProductOption_child As System.Xml.XmlNode In ProductOptions_node_child
                                        Select Case ProductOption_child.Name
                                            Case Is = "name"
                                                Debug.Print(ProductOption_child.Name & " => " & ProductOption_child.InnerText)
                                            Case Is = "AppendText"
                                                newWBTSHOP2.Item("COLOR_CODE") = ProductOption_child.InnerText
                                            Case Is = "QuantityOnHand"
                                                newWBTSHOP2.Item("QUANTITY_ON_HAND") = Val(ProductOption_child.InnerText)
                                            Case Is = "LowStockThreshold"
                                                newWBTSHOP2.Item("LOW_STOCK_THRESHOLD") = Val(ProductOption_child.InnerText)
                                            Case Is = "OutOfStockLimit"
                                                newWBTSHOP2.Item("OUT_OF_STOCK_LIMIT") = Val(ProductOption_child.InnerText)
                                            Case Is = "Image"
                                                newWBTSHOP2.Item("IMAGE") = ProductOption_child.InnerText
                                        End Select
                                    Next
                                    dst.Tables("WBTSHOP2").Rows.Add(newWBTSHOP2)
                                End If

                            Next
                    End Select
                Next
                dst.Tables("WBTSHOP1").Rows.Add(newWBTSHOP1)
            Next

            Dim Pages_doc As XmlDocument = New XmlDocument
            Dim Pages_node_list As XmlNodeList
            Pages_doc.Load(PagesFile)
            Pages_node_list = Pages_doc.GetElementsByTagName("Page")
            For Each Pages_node As System.Xml.XmlNode In Pages_node_list
                Dim PageName As String = ""
                For Each Pages_node_child As System.Xml.XmlNode In Pages_node
                    Select Case Pages_node_child.Name
                        Case Is = "Name"
                            PageName = Pages_node_child.ChildNodes(0).Value
                        Case Is = "ProductLinks"
                            'Stop
                            For Each ProductLinks_node_child As System.Xml.XmlNode In Pages_node_child
                                If ProductLinks_node_child.Name = "Product" Then
                                    For i As Integer = 1 To ProductLinks_node_child.ChildNodes.Count
                                        If ProductLinks_node_child.ChildNodes(i - 1).Name = "SKU" Then
                                            Dim SKU As String = ProductLinks_node_child.ChildNodes(i - 1).InnerText
                                            Dim rowWBTSHOP1 As DataRow = dst.Tables.Item("WBTSHOP1").Select("STYLE_CODE = '" & SKU & "'").FirstOrDefault
                                            ASCMAIN1.Progress("Processing: " & PageName, SKU)
                                            If Not IsNothing(rowWBTSHOP1) Then
                                                rowWBTSHOP1.Item("WEB_CATEGORY") = PageName
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                    End Select
                Next
            Next
            ASCMAIN1.Progress("Posting Data...")

            MyBase.BeginTrans()
            Update_Record_TDA("WBTSHOP1")
            Update_Record_TDA("WBTSHOP2")
            MyBase.CommitTrans("Import Complete")
            ASCMAIN1.Progress("")
        End If

    End Sub

    Private Function CreateInventoryXml(ByVal UploadStart As Integer, ByVal UpLoadFinish As Integer, Optional ByVal TestFilter As String = "", Optional ByVal FullUpload As Boolean = True) As Integer
        Dim RetVal As Integer = 0
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Create XML Document", "")

            Dim inventoryXML As New WBCITEM1
            Dim whseQtyAvail As Int16 = 0

            shopSiteFilename = WB_PARM_INVENTORY_DIR & "INVENTORY_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
            Dim Selector As String = ""

            If InAutoMode Then
                FullUpload = False
                TestFilter = ""
            End If
            Dim RecCnt As Integer = 0
            For Each row As DataRow In dst.Tables("WBTSTYL1").Select(TestFilter)
                Dim ProcessRecord As Boolean = False
                If FullUpload Then
                    ProcessRecord = True
                Else
                    'If (Val(row.Item("LAST_ON_HAND") & "") <> Val(row.Item("CURR_ON_HAND") & "")) Or (row.Item("STYLE_GROUP") & "" = txtNextGroup.Text) Then
                    If TestFilter.Length > 0 Then
                        ProcessRecord = True
                    Else
                        If (row.Item("STYLE_GROUP") & "" = txtNextGroup.Text) Then
                            ProcessRecord = True
                        Else
                            ProcessRecord = False
                        End If
                    End If
                End If
                If ProcessRecord Then
                    If RecCnt >= UploadStart And RetVal <= UpLoadFinish Then
                        row.Item("LAST_ON_HAND") = row.Item("CURR_ON_HAND")
                        RetVal += 1

                        Dim onWeb As Boolean = row.Item("WEB_IND") & "" = "1"
                        Dim hasBatch As Boolean = row.Item("UPLOAD_BATCH") & "" <> ""
                        Dim itemActive As Boolean = row.Item("STYLE_STATUS") & "" = "A"
                        Dim hasInventory As Boolean = Val(row.Item("CURR_ON_HAND")) > 0
                        Dim styleStatus As String = ""
                        'onWeb & hasBatch = U -> Update the item.
                        'onWeb & !hasBatch = S => Add to Web.  This Does Not Get Done Here.  Need to run Inventory load.
                        '!onWeb & hasBatch = X => Remove from Web.
                        '!onWeb & !hasBatch = S => Skip this style.  It's not on the web.
                        If onWeb And hasBatch Then
                            styleStatus = "U"
                        End If
                        If onWeb And Not hasBatch Then
                            styleStatus = "S"
                        End If
                        If Not onWeb And hasBatch Then
                            styleStatus = "X"
                        End If
                        If Not onWeb And Not hasBatch Then
                            styleStatus = "S"
                        End If
                        'If we made it all the way through here and the style has no inventory, remove it from the web.
                        If styleStatus = "U" And Not hasInventory Then
                            styleStatus = "X"
                        End If

                        'If row.Item("STYLE_CODE").ToString = "MTF19619" Then Stop

                        If styleStatus = "U" Then
                            ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
                            styleList.Add(row.Item("STYLE_CODE"))
                            inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
                            MakeInventoryTables(row.Item("STYLE_CODE"))
                        End If

                        If styleStatus = "X" Then
                            ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
                            styleListInactive.Add(row.Item("STYLE_CODE"))
                            inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
                            row.Item("WEB_IND") = "0"
                            row.Item("UPLOAD_BATCH") = Null
                        End If

                        ''Automatically Remove Items From The Web That Are Discontinued and Have No Inventory.
                        'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" And row.Item("STYLE_STATUS") & "" = "D" And Val(row.Item("CURR_ON_HAND")) = 0 Then
                        '    row.Item("WEB_IND") = "0"
                        '    row.Item("UPLOAD_BATCH") = Null
                        'End If

                        ''Automatically Remove Items From The Web That Are Active and Have No Inventory.
                        'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" And row.Item("STYLE_STATUS") & "" = "A" And Val(row.Item("CURR_ON_HAND")) = 0 Then
                        '    row.Item("WEB_IND") = "0"
                        '    row.Item("UPLOAD_BATCH") = Null
                        'End If

                        'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" Then
                        '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
                        '    styleList.Add(row.Item("STYLE_CODE"))
                        '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
                        '    MakeInventoryTable(row.Item("STYLE_CODE"))
                        'End If
                        'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" = "" Then
                        '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
                        '    styleList.Add(row.Item("STYLE_CODE"))
                        '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
                        '    MakeInventoryTable(row.Item("STYLE_CODE"))
                        'End If
                        'If row.Item("WEB_IND") & "" <> "1" And row.Item("UPLOAD_BATCH") & "" <> "" Then
                        '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
                        '    styleListInactive.Add(row.Item("STYLE_CODE"))
                        '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
                        'End If
                    End If
                End If
                RecCnt += 1
            Next

            Dim xmlLabelRequest As XmlDocument = inventoryXML.GetXMLDocument
            ASCMAIN1.Progress("Saving XML Document", "")
            xmlLabelRequest.Save(shopSiteFilename)

            Dim xfileInfo As New FileInfo(shopSiteFilename)
            If xfileInfo.Length <= 500000 Then
                ASCMAIN1.Progress("Loading XML Document in Viewer", "")
                WebBrowser1.Navigate(New Uri(shopSiteFilename))
            Else
                WebBrowser1.Navigate("about:blank")
                Dim HTML As String
                HTML = "<HTML>" & _
                    "<TITLE>XML Inventory Upload</TITLE>" & _
                    "<BODY>" & _
                    "<FONT COLOR = RED>" & _
                    "The XML Document is too " & _
                    "<FONT SIZE = 5>" & _
                    "<B>" & _
                    "Large " & _
                    "</B>" & _
                    "</FONT SIZE>" & _
                    "to display!" & _
                    "</FONT>" & _
                    "</BODY>" & _
                    "</HTML>"

                WebBrowser1.Document.Write(HTML)
            End If
            inventoryXML = Nothing
        Catch ex As Exception
            shopSiteFilename = String.Empty
            If InAutoMode Then
                SendErrorEMail("Error creating XML Document: " & ex.Message)
            Else
                MessageBox.Show("Error creating XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
            End If
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        End Try
        Return RetVal
    End Function

    Private Sub CreateProductXml()

        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Create XML Document", "")

            Dim productXML As New WBCITEM1

            'If MyBase.Absx1.txtFor("STYLE_CODE").Text.Length > 0 Then
            '    ASCMAIN1.Progress("-", MyBase.Absx1.txtFor("STYLE_CODE").Text)
            '    shopSiteFilename = WB_PARM_PRODUCTS_DIR & MyBase.Absx1.txtFor("STYLE_CODE").Text & "_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
            '    productXML.AddStyle(MyBase.Absx1.txtFor("STYLE_CODE").Text)
            '    styleList.Add(MyBase.Absx1.txtFor("STYLE_CODE").Text)
            '    Dim rowWBTSTYL1 As DataRow = ASCDATA1.GetDataRow("Select * From WBTSTYL1 Where Style_Code = :PARM1", "V", MyBase.Absx1.txtFor("STYLE_CODE").Text)
            '    If rowWBTSTYL1 IsNot Nothing AndAlso rowWBTSTYL1.Item("STYLE_STATUS") & String.Empty = "I" Then
            '        styleListInactive.Add(rowWBTSTYL1.Item("STYLE_CODE"))
            '    End If
            'Else
            shopSiteFilename = WB_PARM_PRODUCTS_DIR & "STYLE_CODE_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
            'Do Only One Style
            Dim batchFilter As String = ""
            'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            '    Stop
            '    batchFilter = "UPLOAD_BATCH = '0000008886'"
            'End If
            For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select(batchFilter)
                'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
                '    ASCMAIN1.Progress("-", rowWBTSTYL1.Item("STYLE_CODE"))
                '    styleListInactive.Add(rowWBTSTYL1.Item("STYLE_CODE"))
                '    productXML.AddStyle(rowWBTSTYL1.Item("STYLE_CODE"), styleListInactive)
                'Else
                If rowWBTSTYL1.Item("WEB_IND") & "" = "1" And rowWBTSTYL1.Item("UPLOAD_BATCH") & "" = "" Then
                    ASCMAIN1.Progress("-", rowWBTSTYL1.Item("STYLE_CODE"))
                    styleList.Add(rowWBTSTYL1.Item("STYLE_CODE"))
                    productXML.AddStyle(rowWBTSTYL1.Item("STYLE_CODE"), styleListInactive)
                    MakeInventoryTables(rowWBTSTYL1.Item("STYLE_CODE"))
                End If
                If rowWBTSTYL1.Item("UPLOAD_BATCH") & "" <> "" And rowWBTSTYL1.Item("WEB_IND") & "" <> "1" Then
                    ASCMAIN1.Progress("-", rowWBTSTYL1.Item("STYLE_CODE"))
                    styleListInactive.Add(rowWBTSTYL1.Item("STYLE_CODE"))
                    productXML.AddStyle(rowWBTSTYL1.Item("STYLE_CODE"), styleListInactive)
                End If
                'End If
            Next
            'End If

            Dim xmlLabelRequest As XmlDocument = productXML.GetXMLDocument
            ASCMAIN1.Progress("Saving XML Document", "")
            xmlLabelRequest.Save(shopSiteFilename)

            Dim xfileInfo As New FileInfo(shopSiteFilename)
            If xfileInfo.Length <= 500000 Then
                ASCMAIN1.Progress("Loading XML Document in Viewer", "")
                WebBrowser1.Navigate(New Uri(shopSiteFilename))
            Else
                WebBrowser1.Navigate("about:blank")
                Dim HTML As String
                HTML = "<HTML>" & _
                    "<TITLE>XML Style Upload</TITLE>" & _
                    "<BODY>" & _
                    "<FONT COLOR = RED>" & _
                    "The XML Document is too " & _
                    "<FONT SIZE = 5>" & _
                    "<B>" & _
                    "Large " & _
                    "</B>" & _
                    "</FONT SIZE>" & _
                    "to display!" & _
                    "</FONT>" & _
                    "</BODY>" & _
                    "</HTML>"

                WebBrowser1.Document.Write(HTML)

            End If

        Catch ex As Exception
            shopSiteFilename = String.Empty
            If InAutoMode Then
                SendErrorEMail("Error creating XML Document: " & ex.Message)
            Else
                MessageBox.Show("Error creating XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
            End If
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Sub CreateZipEmail(ByVal ORDR_NO As String, ByVal REQ_EMAIL As String, ByVal Web_Dest_Folder As String)
        Const FROM_ADDRESS As String = "site.admin@regency-rib.com"
        Const FROM_NAME As String = "Regency International"
        Const BCC_ADDRESS As String = "mariog@regency-rib.com"
        Const BCC_NAME As String = "Mario Arenas Jr."
        Const SERVER_IP As String = "192.168.110.221"
        Const SERVER_PORT As Integer = 25
        Const SERVER_ACCOUNT As String = "site.admin@regency-rib.com"
        Const SERVER_PASSWORD As String = "0ff1c3"
        Dim TO_SUBJECT As String = "Requested Images For Order " & ORDR_NO
        Dim HTMLBody As String
        HTMLBody = "Thank you for your request;" & vbCrLf & vbCrLf
        HTMLBody = HTMLBody & String.Format("Below is a link to the images related to order number {0}:", ORDR_NO) & vbCrLf
        HTMLBody = HTMLBody & String.Format("{0}{1}.zip", Web_Dest_Folder, ORDR_NO) & vbCrLf & vbCrLf
        If MissingZipFiles.Count > 0 Then
            HTMLBody = HTMLBody & "The following is a list of images on the order that were not able to be included in the file:" & vbCrLf
            For Each style As String In MissingZipFiles
                HTMLBody = HTMLBody & style & vbCrLf
            Next
            HTMLBody = HTMLBody & vbCrLf & vbCrLf
        End If
        HTMLBody = HTMLBody & "Please note that the file for this link will be active for 10 days from the date of this email.  If you need the files after that date you are more than welcome to request them again." & vbCrLf & vbCrLf
        HTMLBody = HTMLBody & "If you have any additional questions related to this order please contact your sales rep which you can find under the Locate Sales Rep section of our web site." & vbCrLf & vbCrLf
        HTMLBody = HTMLBody & "www.Regency-rib.com"

        Try
            Dim mail As New MailMessage() With {.From = New MailAddress(FROM_ADDRESS, FROM_NAME)}
            mail.To.Add(New MailAddress(REQ_EMAIL, ""))
            mail.Subject = TO_SUBJECT
            mail.IsBodyHtml = False
            mail.Body = HTMLBody
            mail.Bcc.Add(New MailAddress(BCC_ADDRESS, BCC_NAME))

            Dim smtp As New SmtpClient(SERVER_IP, SERVER_PORT)
            If smtp IsNot Nothing Then
                smtp.Credentials = New System.Net.NetworkCredential(SERVER_ACCOUNT, SERVER_PASSWORD)
            Else
                Dim eMsg As String = "SMTP Client could not be created."
                MsgBox(eMsg, MsgBoxStyle.OkOnly, "Error")
            End If

            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                smtp.Send(mail)
            End If
        Catch ex As Exception
            MsgBox("Error Trying to Generate Document" & vbCrLf & ex.Message)
        End Try
    End Sub

    Private Sub CreateZipOrder(ByVal ORDR_NO As String, ByVal Web_Dest_Folder As String)
        Dim pbInt As Integer = 0
        Dim FullDestination As String = String.Format("{0}{1}.zip", Web_Dest_Folder, ORDR_NO)

        Dim SQLP As New System.Text.StringBuilder
        SQLP.Length = 0
        SQLP.AppendLine("select IC_PARM_STYLE_IMG_DIR from ICTPARM1 where IC_PARM_KEY = 'Z'")
        ASCMAIN1.sql = SQLP.ToString()
        Dim IMAGES_FOLDER As String = ASCDATA1.GetDataValue

        If File.Exists(FullDestination) Then
            File.Delete(FullDestination)
        End If

        MissingZipFiles.Clear()

        Try
            Dim Zip1 As New nsoftware.IPWorksZip.Zip
            Zip1.RuntimeLicense = nSoftwareKeys("nSoftwareZipkey")
            Zip1.ArchiveFile = FullDestination

            Dim SQLO As StringBuilder = New StringBuilder() With {.Length = 0}
            SQLO.AppendLine("SELECT STYLE_CODE, COLOR_CODE")
            SQLO.AppendLine("FROM SOTORDR2")
            SQLO.AppendLine(String.Format("WHERE ORDR_NO = '{0}'", ORDR_NO))
            SQLO.AppendLine("GROUP BY STYLE_CODE, COLOR_CODE")
            Dim tbl As DataTable = ASCDATA1.GetDataTable(SQLO.ToString())

            Dim rtp As Integer = tbl.Select().Length
            Dim zipBuffer As Integer = 1
            If rtp > 10 Then
                zipBuffer = CInt(rtp * 0.1)
            End If
            Dim rowsToProcess As Integer = rtp + zipBuffer
            Dim currentRow As Integer = 1

            For Each rowSOTORDR2 As DataRow In tbl.Rows
                pbInt = CInt((currentRow / rowsToProcess) * 100)
                Dim STYLE_CODE As String = rowSOTORDR2.Item("STYLE_CODE") & ""
                Dim COLOR_CODE As String = rowSOTORDR2.Item("COLOR_CODE") & ""

                Dim IMAGE_NAME As String = String.Format("{0}-{1}.jpg", STYLE_CODE, COLOR_CODE)

                If IMAGE_NAME <> "" Then
                    Dim FILENAME As String = IMAGES_FOLDER & IMAGE_NAME
                    If My.Computer.FileSystem.FileExists(FILENAME) Then
                        Zip1.IncludeFiles(FILENAME)
                    Else
                        MissingZipFiles.Add(IMAGE_NAME)
                    End If
                End If
                currentRow += 1
            Next

            Zip1.Compress()
            Zip1.Dispose()
        Catch ex As Exception
            'MsgBox("Error Creating Zip FileG")
        End Try
    End Sub

    Private Sub FolderSelect()

        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Update DataBase From File Folder"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("You Will Be Asked To Select A Folder That Contains JPG Files.")
        iMSG.AppendLine("I Will Update The List To Remove All Batch Numbers From Files")
        iMSG.AppendLine("That Match The Names In That Folder.  You Still Need To Also")
        iMSG.AppendLine("Have Those Images In The Master Folder Before Uploading!!!")
        iMSG.AppendLine("I Will Only Use The Folder You Provide As A List Of Files")
        iMSG.AppendLine("")
        iMSG.AppendLine("Are You Ready?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Dim fList As New List(Of String)
        Dim fDialog As New FolderBrowserDialog
        Dim dInfo As DirectoryInfo
        fDialog.Description = "Please Select The Folder Where The Images Are So I Can Make A List From Them."
        fDialog.ShowDialog()
        dInfo = New DirectoryInfo(fDialog.SelectedPath)
        If fDialog.SelectedPath.Length > 0 Then
            For Each fName As FileInfo In dInfo.GetFiles("*.jpg")
                fList.Add(fName.Name)
            Next fName
            Dim fCount As Integer = 0
            For Each File As String In fList
                Dim STYLE_CODE As String = ""
                Dim COLOR_CODE As String = ""
                Dim PosDash As Integer = File.IndexOf("-")
                Dim PosDot As Integer = File.IndexOf(".")

                If PosDash > 1 And PosDot > 1 Then
                    STYLE_CODE = File.Substring(0, PosDash)
                    COLOR_CODE = File.Substring(PosDash + 1, (PosDot) - (PosDash + 1))
                    For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                        rowWBTSTYL1.Item("UPLOAD_BATCH") = Null
                        fCount += 1
                    Next
                End If
            Next
            If fCount > 0 Then
                MsgBox(String.Format("I Updated {0} Styles", fCount), MsgBoxStyle.Information, "Update Complete")
            Else
                MsgBox("I Didn't Find And Styles To Update", MsgBoxStyle.Information, "Update Complete?")
            End If
        End If
    End Sub

    Sub ftp_File()
        Ftp1.User = "EDI_RegencyInternational"
        Ftp1.Password = "N0wayfa1r!"
        Ftp1.RemoteHost = "edi.wayfair.com"
        Ftp1.RemotePath = "inventory"
        Ftp1.Logon()
        Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
        Ftp1.LocalFile = TransferFile
        Ftp1.RemoteFile = "regency.csv"
        Ftp1.Overwrite = True
        Ftp1.Upload()
        Ftp1.Logoff()
    End Sub

    Private Function ftpProducts() As Boolean
        Try

            'FTPProducts = False

            Me.Cursor = Cursors.WaitCursor

            If shopSiteFilename.Length = 0 Then
                If InAutoMode Then
                    SendErrorEMail("XML document cannot be found")
                Else
                    MessageBox.Show("XML document cannot be found", "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If
            If Not My.Computer.FileSystem.FileExists(shopSiteFilename) Then
                If InAutoMode Then
                    SendErrorEMail("XML document cannot be found")
                Else
                    MessageBox.Show("XML document cannot be found", "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If

            If WB_PARM_SITE_IP.Length = 0 Then
                If InAutoMode Then
                    SendErrorEMail("ftp IP address is missing")
                Else
                    MessageBox.Show("ftp IP address is missing", "ftp Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If

            If WB_PARM_SITE_USER.Length = 0 Then
                If InAutoMode Then
                    SendErrorEMail("ftp User ID is missing")
                Else
                    MessageBox.Show("ftp User ID is missing", "ftp Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If

            If WB_PARM_SITE_PWD.Length = 0 Then
                If InAutoMode Then
                    SendErrorEMail("ftp User Password is missing")
                Else
                    MessageBox.Show("ftp User Password is missing", "ftp Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If

            If WB_PARM_SITE_OUTPUT_DIR.Length > 0 AndAlso Not WB_PARM_SITE_OUTPUT_DIR.EndsWith("/") Then
                WB_PARM_SITE_OUTPUT_DIR &= "/"
            End If


            ASCMAIN1.Progress("Creating FTP Connection to ShopSite", "")

            Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
            ASCMAIN1.Progress("-", "RemoteHost")
            'Sftp1.RemoteHost = WB_PARM_SITE_IP
            Sftp1.RemoteHost = "69.39.227.201"


            ASCMAIN1.Progress("-", "User")
            Sftp1.User = WB_PARM_SITE_USER
            'Ftp1.User = WB_PARM_SITE_USER

            ASCMAIN1.Progress("-", "Password")
            Sftp1.Password = WB_PARM_SITE_PWD
            'Ftp1.Password = WB_PARM_SITE_PWD

            ASCMAIN1.Progress("-", "RemoteFile")
            Sftp1.RemoteFile = String.Empty

            ASCMAIN1.Progress("-", "Timeout")
            Sftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")
            'Sftp1.SSHAuthMode = nsoftware.IPWorks.ftpAuthModes.amPassword

            ' Send the graphics first since they must appear for new products before the Item is sent
            ' ftp 
            If FTPImages Or FTPTables Then
                ASCMAIN1.Progress("Upload Graphics And/Or Tables", "")
                Try

                    If Sftp1.Connected Then
                        Sftp1.Logoff()
                    End If

                    Try
                        Sftp1.Logon()
                    Catch ex As Exception
                        Sftp1.Logoff()
                        Sftp1.Logon()
                    End Try

                    If Sftp1.Connected Then
                        If WB_PARM_MASTER_IMAGES.StartsWith("/") Then WB_PARM_MASTER_IMAGES = WB_PARM_MASTER_IMAGES.Substring(1)

                        ASCMAIN1.Progress("Upload new Images", "")
                        'Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                        ' Check This
                        For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select()
                            If styleList.Contains(rowWBTSTYL1.Item("STYLE_CODE").ToString) Then
                                If rowWBTSTYL1.Item("UPLOAD_IMG").ToString & "" = "1" Then
                                    For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select("STYLE_CODE = '" & rowWBTSTYL1.Item("STYLE_CODE").ToString & "'")
                                        If rowWBTSTYL2.Item("IMG_FOUND").ToString & "" = "1" Then
                                            Sftp1.LocalFile = WB_PARM_MASTER_IMAGES & rowWBTSTYL2.Item("IMG_NAME").ToString '"\\vmware-host\Shared Folders\Documents\RegencyMasterImages\product\"
                                            Dim imageFile As String = My.Computer.FileSystem.GetName(Sftp1.LocalFile)
                                            ASCMAIN1.Progress("-", imageFile)
                                            With imageFile
                                                If .Length > 3 Then
                                                    If .Substring(.Length - 3, 3) = "JPG" Then
                                                        imageFile = .Substring(0, .Length - 3) & "jpg"
                                                    End If
                                                End If
                                            End With
                                            Sftp1.RemoteFile = WB_PARM_SITE_IMAGES_DIR & imageFile '"/www/media/product/"
                                            Sftp1.Overwrite = True
                                            Sftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                                            'Sftp1.DeleteFile(WB_PARM_SITE_IMAGES_DIR & imageFile)
                                            Sftp1.DoEvents()
                                            Sftp1.Upload()
                                            Sftp1.DoEvents()
                                            Do While Not Sftp1.Idle
                                            Loop
                                            rowWBTSTYL1.Item("UPLOAD_IMG") = "0"
                                        End If
                                    Next
                                End If
                                If FTPTables Then
                                    ASCMAIN1.Progress("Upload Tables", rowWBTSTYL1.Item("STYLE_CODE").ToString)
                                    Dim File_Name As String = String.Format("{0}invtbl\{1}.html", WB_PARM_PRODUCTS_DIR, rowWBTSTYL1.Item("STYLE_CODE"))
                                    Sftp1.LocalFile = File_Name
                                    If WB_PARM_SITE_OUTPUT_DIR.StartsWith("/") Then WB_PARM_SITE_OUTPUT_DIR = WB_PARM_SITE_OUTPUT_DIR.Substring(1)
                                    Sftp1.RemoteFile = String.Format("{0}invtbl/{1}.html", WB_PARM_SITE_OUTPUT_DIR, rowWBTSTYL1.Item("STYLE_CODE"))
                                    Sftp1.Overwrite = True
                                    Sftp1.Upload()
                                    Do While Not Sftp1.Idle
                                    Loop
                                End If
                            End If
                        Next
                    Else
                        If InAutoMode Then
                            SendErrorEMail("Could not connect to ShopSite")
                        Else
                            MessageBox.Show("Could not connect to ShopSite", "FTP Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Exit Function
                    End If

                Catch ex As Exception
                    If InAutoMode Then
                        SendErrorEMail("Error ftping Images: " & ex.Message)
                    Else
                        MessageBox.Show("Error ftping Image File: " & ex.Message, "Error", MessageBoxButtons.OK)
                    End If

                End Try
            End If

            ASCMAIN1.Progress("Upload Product File", "")
            Try
                Sftp1.Logoff()
                Sftp1.Logon()
            Catch ex As Exception
                Sftp1.Logoff()
                Sftp1.Logon()
            End Try
            'Ftp1.Logon()

            If Not Sftp1.Connected Then
                If InAutoMode Then
                    SendErrorEMail("Could not connect to ShopSite")
                Else
                    MessageBox.Show("Could not connect to ShopSite", "FTP Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If

            ASCMAIN1.Progress("-", "TransferMode")
            'Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmASCII

            ASCMAIN1.Progress("-", "LocalFile")
            Sftp1.LocalFile = shopSiteFilename

            ASCMAIN1.Progress("-", "RemoteFile")
            If WB_PARM_SITE_OUTPUT_DIR.StartsWith("/") Then WB_PARM_SITE_OUTPUT_DIR = WB_PARM_SITE_OUTPUT_DIR.Substring(1)
            Sftp1.RemoteFile = WB_PARM_SITE_OUTPUT_DIR & My.Computer.FileSystem.GetName(shopSiteFilename)
            Sftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmDefault

            ASCMAIN1.Progress("-", "Upload")
            Sftp1.Upload()

            Dim script As String = WB_PARM_SITE_PRODUCT_POST_URL & My.Computer.FileSystem.GetName(shopSiteFilename)
            docComplete = False

            ASCMAIN1.Progress("-", "Post")
            WebBrowser1.Navigate("")
            WebBrowser1.Navigate(script)

            ' Delay 2 seconds
            System.Threading.Thread.Sleep(2000)
            While Not docComplete
                'Processes all windows messages currently in the message queue
                Application.DoEvents()
            End While
            docComplete = False

            ' Part 3: Automatically click the Login button
            'Dim theWElementCollection As HtmlElementCollection = WebBrowser1.Document.GetElementsByTagName("input")
            For Each curElement As HtmlElement In WebBrowser1.Document.All

                Select Case curElement.GetAttribute("value")
                    Case ""

                    Case "1"

                    Case "2"

                End Select

                'WebBrowser1.Document.GetElementById(???).InvokeMember("Free User")
                If curElement.GetAttribute("value").Equals("Login") Then
                    curElement.InvokeMember("click")
                    Exit For
                End If
            Next '

            ' Delay 5 seconds
            System.Threading.Thread.Sleep(5000)

            ' Auto Publish
            script = WB_PARM_SITE_PRODUCT_PUB_URL
            ASCMAIN1.Progress("-", "Publish")
            WebBrowser1.Navigate("")
            WebBrowser1.Navigate(script)

            ' Delay 2 seconds
            System.Threading.Thread.Sleep(2000)
            While Not docComplete
                'Processes all windows messages currently in the message queue
                Application.DoEvents()
            End While
            docComplete = False

            ' Remove file from shopsite
            ASCMAIN1.Progress("-", "Delete")
            Sftp1.DeleteFile(WB_PARM_SITE_OUTPUT_DIR & My.Computer.FileSystem.GetName(shopSiteFilename))

            shopSiteFilename = String.Empty
            WebBrowser1.Navigate("")

            '' Delete any Inactive styles
            'ASCMAIN1.Progress("Remove Inactive Pages", "")
            'For Each inactiveStyle As String In styleListInactive
            '    Sftp1.RemoteFile = WB_PARM_SITE_OUTPUT_DIR & inactiveStyle.ToLower & ".htm"
            '    ASCMAIN1.Progress("-", inactiveStyle.ToLower & ".htm")
            '    If Sftp1.FileExists Then
            '        Sftp1.DeleteFile(WB_PARM_SITE_OUTPUT_DIR & inactiveStyle.ToLower & ".htm")
            '    End If
            'Next

            Try
                BeginTrans()

                Dim UPLOAD_BATCH As String = ASCMAIN1.Next_Control_No("WBTSTYL1.UPLOAD_BATCH")
                For Each styleCode As String In styleList
                    ASCDATA1.ExecuteSQL(String.Format("Update WBTSTYL1 SET WEB_IND = '1', UPLOAD_BATCH = '{0}' WHERE STYLE_CODE = '{1}'", UPLOAD_BATCH, styleCode))
                    Dim rowWBTSTYL1 As DataRow = dst.Tables.Item("WBTSTYL1").Select("STYLE_CODE = '" & styleCode & "'").FirstOrDefault()
                    If Not IsNothing(rowWBTSTYL1) Then
                        rowWBTSTYL1.Item("UPLOAD_BATCH") = UPLOAD_BATCH
                    End If
                Next
                For Each styleCode As String In styleListInactive
                    ASCDATA1.ExecuteSQL(String.Format("Update WBTSTYL1 SET WEB_IND = '0', UPLOAD_BATCH = NULL WHERE STYLE_CODE = '{0}'", styleCode))
                    Dim rowWBTSTYL1 As DataRow = dst.Tables.Item("WBTSTYL1").Select("STYLE_CODE = '" & styleCode & "'").FirstOrDefault()
                    If Not IsNothing(rowWBTSTYL1) Then
                        rowWBTSTYL1.Item("WEB_IND") = "0"
                        rowWBTSTYL1.Item("UPLOAD_BATCH") = Null
                    End If
                Next
                CommitTrans()
                styleList.Clear()
                ftpProducts = True

            Catch ex As Exception
                If InAutoMode Then
                    Rollback()
                    SendErrorEMail("Error ftping: " & ex.Message)
                    WBCMAIN1.AddTaskDetail(TASK_NO, "Error ftping: " & ex.Message)
                Else
                    Rollback(ex.Message)
                End If
            End Try

        Catch ex As Exception
            If InAutoMode Then
                SendErrorEMail("Error ftping XML Document: " & ex.Message)
            Else
                MessageBox.Show("Error ftping XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
            End If
        Finally


            Me.Cursor = Cursors.Default
            itemUploaded = True
            ASCMAIN1.Progress(String.Empty, String.Empty)

            WebBrowser1.Navigate(String.Empty)
            If Sftp1.Connected Then Sftp1.Logoff()
        End Try
    End Function

    Private Function GetNextDelDate(STYLE_CODE As String, COLOR_CODE As String) As String
        Dim retVal As String = ""
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("SELECT MAX(STATUS_DATE) AS NEXT_DATE")
        SQLS.AppendLine("FROM ICTSTDQ1")
        SQLS.AppendLine("WHERE WHSE_CODE = 'MS'")
        SQLS.AppendLine(String.Format("AND STYLE_CODE = '{0}'", STYLE_CODE))
        SQLS.AppendLine(String.Format("AND COLOR_CODE = '{0}'", COLOR_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim NEXT_DATE As String = ASCDATA1.GetDataValue
        If Not IsNothing(NEXT_DATE) Then
            If IsDate(NEXT_DATE) Then
                retVal = Format(CDate(NEXT_DATE), "MM/dd/yy").ToString
            End If
        End If
        Return retVal
    End Function

    Private Sub MakeInventoryTables(STYLE_CODE As String)
        Dim StyleTable As New System.Text.StringBuilder
        StyleTable.Length = 0
        StyleTable.AppendLine("<table>")
        StyleTable.AppendLine("    <tr>")
        StyleTable.AppendLine("        <td style='border: 1px solid black; background: #E2FFCF; text-align: left'><strong>Color</strong></td>")
        StyleTable.AppendLine("        <td style='border: 1px solid black; background: #E2FFCF; text-align: right'><strong>Qty On Hand</strong></td>")
        StyleTable.AppendLine("        <td style='border: 1px solid black; background: #E2FFCF; text-align: right'><strong>Future Available</strong></td>")
        StyleTable.AppendLine("        <td style='border: 1px solid black; background: #E2FFCF; text-align: right'><strong>Total Available</strong></td>")
        StyleTable.AppendLine("        <td style='border: 1px solid black; background: #E2FFCF; text-align: right'><strong>Future Date</strong></td>")
        StyleTable.AppendLine("    </tr>")
        For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE), "COLOR_CODE_LONG")
            StyleTable.AppendLine("    <tr>")
            StyleTable.AppendLine(String.Format("        <td style='border: 1px solid black; text-align: left'>{0}</td>", rowWBTSTYL2.Item("COLOR_CODE_LONG") & ""))
            StyleTable.AppendLine(String.Format("        <td style='border: 1px solid black; text-align: right'>{0}</td>", Val(rowWBTSTYL2.Item("MSOH") & ""))) 'Qty On Hand
            StyleTable.AppendLine(String.Format("        <td style='border: 1px solid black; text-align: right'>{0}</td>", Val(rowWBTSTYL2.Item("MSFT") & ""))) 'Future Available
            StyleTable.AppendLine(String.Format("        <td style='border: 1px solid black; text-align: right'>{0}</td>", Val(rowWBTSTYL2.Item("MSOH") + Val(rowWBTSTYL2.Item("MSFT")) & ""))) 'Total Available
            'Future Date
            If Val(rowWBTSTYL2.Item("MSFT") & "") > 0 Then
                StyleTable.AppendLine(String.Format("        <td style='border: 1px solid black; text-align: right'>{0}</td>", GetNextDelDate(STYLE_CODE, (rowWBTSTYL2.Item("COLOR_CODE") & "").ToString)))
            Else
                StyleTable.AppendLine("        <td style='border: 1px solid black; text-align: right'>&nbsp;</td>")
            End If
            StyleTable.AppendLine("    </tr>")
        Next
        StyleTable.AppendLine("    <tr>")
        StyleTable.AppendLine("        <td colspan=5 style='border: 1px solid black; text-align: center'>")
        StyleTable.AppendLine("             <small>")
        StyleTable.AppendLine("                 <bold><font color='red'>*</font> Available Quantites May Change And Affect The Amount You Can Purchase <font color='red'>*</font></bold>")
        StyleTable.AppendLine("             </small>")
        StyleTable.AppendLine("        </td>")
        StyleTable.AppendLine("    </tr>")
        StyleTable.AppendLine("</table>")
        Dim File_Name As String = WB_PARM_PRODUCTS_DIR & "invtbl\" & STYLE_CODE & ".html"
        If Not IO.Directory.Exists(WB_PARM_PRODUCTS_DIR & "\invtbl\") Then
            IO.Directory.CreateDirectory(WB_PARM_PRODUCTS_DIR & "\invtbl\")
        End If
        If IO.File.Exists(File_Name) Then
            IO.File.Delete(File_Name)
        End If
        Using outfile As New StreamWriter(File_Name)
            outfile.Write(StyleTable.ToString())
        End Using
    End Sub

    Private Sub MakeSendZipFile()
        Dim Web_Dest_Folder As String = "\\192.168.110.224\product\zip\" 'Parameterize this before live
        Dim Web_Dest_URL As String = "http://api.regency-rib.com:8181/images/product/zip/" 'Parameterize this before live
        Dim SQLW As New StringBuilder
        SQLW.Length = 0
        SQLW.AppendLine("SELECT *")
        SQLW.AppendLine("FROM WBTIMGR1")
        SQLW.AppendLine("WHERE REQ_STATUS = 'W'")
        Fill_Records("WBTIMGR1", , , SQLW.ToString)
        MyBase.BeginTrans()
        For Each rowWBTIMGR1 As DataRow In dst.Tables("WBTIMGR1").Select()
            CreateZipOrder(rowWBTIMGR1.Item("ORDR_NO").ToString, Web_Dest_Folder)
            CreateZipEmail(rowWBTIMGR1.Item("ORDR_NO").ToString, rowWBTIMGR1.Item("REQ_EMAIL").ToString, Web_Dest_URL)
            rowWBTIMGR1.Item("REQ_STATUS") = "A"
        Next

        Update_Record_TDA("WBTIMGR1")
        MyBase.CommitTrans("")
        SQLW.Length = 0
        SQLW.AppendLine("SELECT *")
        SQLW.AppendLine("FROM WBTIMGR1")
        SQLW.AppendLine("WHERE REQ_DATE <= sysdate-10")
        SQLW.AppendLine("AND REQ_STATUS = 'A'")
        Fill_Records("WBTIMGR1", , , SQLW.ToString)
        MyBase.BeginTrans()
        For Each rowWBTIMGR1 As DataRow In dst.Tables("WBTIMGR1").Select()
            System.IO.File.Delete(String.Format("{0}{1}.zip", Web_Dest_Folder, rowWBTIMGR1.Item("ORDR_NO").ToString))
            rowWBTIMGR1.Item("REQ_STATUS") = "X"
        Next
        Update_Record_TDA("WBTIMGR1")
        MyBase.CommitTrans("")
    End Sub

    Private Sub RemoveUnwanted()
        Dim recCount As Integer = 0
        Dim filWHTSTYL1 As String = "WEB_IND = '1' AND ISNULL(UPLOAD_BATCH,'NULL') <> 'NULL'"
        For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select(filWHTSTYL1)
            Dim RemoveStyle As Boolean = False
            If rowWBTSTYL1.Item("STYLE_STATUS") <> "A" Then
                Dim filWHTSTYL2 As String = String.Format("STYLE_CODE = '{0}'", rowWBTSTYL1.Item("STYLE_CODE"))
                RemoveStyle = True
                For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select(filWHTSTYL2)
                    If Val(rowWBTSTYL2.Item("MSOH") & "") > 0 Or Val(rowWBTSTYL2.Item("MSFT") & "") > 0 Then
                        RemoveStyle = False
                        Exit For
                    End If
                Next
            End If
            If RemoveStyle Then
                rowWBTSTYL1.Item("WEB_IND") = Null
                recCount += 1
            End If
        Next
        If recCount > 0 Then
            MsgBox("Removed " & recCount & " Styles", MsgBoxStyle.Information, "Finished")
        Else
            MsgBox("No Styles Found To Remove", MsgBoxStyle.Information, "Finished")
        End If
    End Sub

    Private Sub SendErrorEMail(ByVal MsgBody As String, Optional ByVal StopProcess As Boolean = True)

        Const FROM_ADDRESS As String = "new.accounts@regency-rib.com"
        Const FROM_NAME As String = "Regency Auto-Sync Manager"
        Const SERVER_IP As String = "192.168.110.221"
        Const SERVER_PORT As Integer = 25
        Const SERVER_ACCOUNT As String = "new.accounts@regency-rib.com"
        Const SERVER_PASSWORD As String = "0ff1c3"
        Const EMAIL_ADDRESS As String = "mariog@regency-rib.comt"
        Const EMAIL_NAME As String = "Mario Arenas Jr."
        Dim CC_ADDRESS As String = "whr@waynerichmond.net"
        Dim CC_NAME As String = "Wayne Richmond"


        Try
            Dim HTMLBody As String = "A Message Was Reported From The Regency Auto-Refresh Process As Follows:" & vbCrLf & vbCrLf & MsgBody
            Dim mail As New MailMessage() With {.From = New MailAddress(FROM_ADDRESS, FROM_NAME)}
            mail.To.Add(New MailAddress(EMAIL_ADDRESS, EMAIL_NAME))
            mail.Subject = "Message From Regency Auto-Refresh Process"
            mail.IsBodyHtml = True
            mail.Body = HTMLBody
            'If Not WayneOnly Then
            mail.CC.Add(New MailAddress(CC_ADDRESS, CC_NAME))
            'End If

            Dim smtp As New SmtpClient(SERVER_IP, SERVER_PORT)
            If smtp IsNot Nothing Then
                smtp.Credentials = New System.Net.NetworkCredential(SERVER_ACCOUNT, SERVER_PASSWORD)
            Else
                MsgBox("SMTP Client could not be created.", MsgBoxStyle.OkOnly, "Error")
                InAutoMode = False
            End If

            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                smtp.Send(mail)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Sending Mail")
            InAutoMode = False
        End Try

        If StopProcess Then
            InAutoMode = False
        End If
    End Sub

    Private Sub SendWayfair()
        Try
            ASCMAIN1.Progress("Uploading Inventory", "")
            If My.Computer.FileSystem.FileExists(TransferFile) Then
                My.Computer.FileSystem.DeleteFile(TransferFile)
            End If
            BuildFTPFile()
            ftp_File()
            Dim mBody As String = String.Format("Wayfair Inventory Feed Updated with {0} products at {1}", _
                                                dst.Tables.Item("ICTSTATX").Rows.Count, _
                                                Format(Now(), "hh:mm tt"))
            SendErrorEMail(mBody, False)
        Catch ex As Exception
            SendErrorEMail("Error During Wayfair Inventory Upload: " & ex.Message, False)
        End Try
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub Setup_WBTSTYL2()
        If grdWBTSTYL1.ActiveRow Is Nothing OrElse (Not grdWBTSTYL1.ActiveRow.IsDataRow Or grdWBTSTYL1.ActiveRow.IsAddRow) Then
            grdWBTSTYL2.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdWBTSTYL2.DataSource, DataTable).DefaultView
            Dim STYLE_CODE As String = grdWBTSTYL1.ActiveRow.Cells("STYLE_CODE").Text
            dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "'"
            grdWBTSTYL2.Text = "Color Details for Style " & STYLE_CODE
            grdWBTSTYL2.Visible = True
        End If
    End Sub

    Private Sub ShowPrevious(Show As Boolean)
        Dim dvw As DataView = DirectCast(grdWBTSTYL1.DataSource, DataTable).DefaultView
        If Show Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "ISNULL(UPLOAD_BATCH,'') = ''"
        End If
    End Sub

    Private Sub ShowVariance(Show As Boolean)
        Dim dvw As DataView = DirectCast(grdWBTSTYL1.DataSource, DataTable).DefaultView
        If Show Then
            dvw.RowFilter = "ISNULL(LAST_ON_HAND,0) <> ISNULL(CURR_ON_HAND,0)"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub UpdateWBTSTYLS()
        Dim Numerator As Long = 0
        Dim Denominator As Long = dst.Tables("WBTSTYL2").Select().Count
        Dim Result As Double = 0
        Dim LAST_STYLE_CODE As String = ""
        Dim RUN_AVALABLE As Integer = 0
        For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select("", "STYLE_CODE, COLOR_CODE")
            If rowWBTSTYL2.Item("STYLE_CODE").ToString <> LAST_STYLE_CODE Then
                For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select(String.Format("STYLE_CODE = '{0}'", LAST_STYLE_CODE))
                    rowWBTSTYL1.Item("CURR_ON_HAND") = RUN_AVALABLE
                    RUN_AVALABLE = 0
                Next
            End If
            LAST_STYLE_CODE = rowWBTSTYL2.Item("STYLE_CODE")
            'If rowWBTSTYL2.Item("STYLE_CODE").ToString = "MTX46335" Then
            '    Stop
            'End If
            Numerator += 1
            If Denominator <> 0 Then
                Result = (Numerator / Denominator) * 100
            Else
                Result = 0
            End If
            ASCMAIN1.Progress("Updating Inventory", Format(Result, "##0.0 Pct"))
            Dim MSOH As Integer = 0, MSFT As Integer = 0, SWOH As Integer = 0, SWFT As Integer = 0
            Dim Filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", rowWBTSTYL2.Item("STYLE_CODE"), rowWBTSTYL2.Item("COLOR_CODE"))
            For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select(Filter)
                MSOH = Val(rowICTSTYC1.Item("MSOH") & "")
                MSFT = Val(rowICTSTYC1.Item("MSFT") & "")
                SWOH = Val(rowICTSTYC1.Item("SWOH") & "")
                SWFT = Val(rowICTSTYC1.Item("SWFT") & "")
            Next
            rowWBTSTYL2.Item("MSOH") = Val(MSOH)
            rowWBTSTYL2.Item("MSFT") = Val(MSFT)
            rowWBTSTYL2.Item("SWOH") = Val(SWOH)
            rowWBTSTYL2.Item("SWFT") = Val(SWFT)
            RUN_AVALABLE = RUN_AVALABLE + MSOH + MSFT
            Dim S As New System.Text.StringBuilder
            S.Length = 0
            S.AppendLine("SELECT NVL(STYLE_COLOR_STATUS,'A') AS STYLE_STATUS")
            S.AppendLine("FROM ICTSTYC1")
            S.AppendLine(String.Format("WHERE STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", rowWBTSTYL2.Item("STYLE_CODE"), rowWBTSTYL2.Item("COLOR_CODE")))
            ASCMAIN1.sql = S.ToString()
            rowWBTSTYL2.Item("COLOR_STATUS") = ASCDATA1.GetDataValue
            'If rowWBTSTYL2.Item("IMG_NAME").ToString & "" = "" Then
            Dim imagename As String = rowWBTSTYL2.Item("STYLE_CODE").ToString & "-" & rowWBTSTYL2.Item("COLOR_CODE").ToString & ".JPG"
            Dim imagefile As String = WB_PARM_MASTER_IMAGES & imagename
            If File.Exists(imagefile) Then
                rowWBTSTYL2.Item("IMG_NAME") = imagename
                rowWBTSTYL2.Item("IMG_FOUND") = "1"
                For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select(String.Format("STYLE_CODE = '{0}'", rowWBTSTYL2.Item("STYLE_CODE").ToString))
                    rowWBTSTYL1.Item("UPLOAD_IMG") = "1"
                Next
            Else
                rowWBTSTYL2.Item("IMG_NAME") = Null
                rowWBTSTYL2.Item("IMG_FOUND") = "0"
            End If
            If dst.Tables("WBTSTYL1").Select("STYLE_CODE = '" & rowWBTSTYL2.Item("STYLE_CODE").ToString() & "'").Count() = 1 Then
                If dst.Tables("WBTSTYL1").Select("STYLE_CODE = '" & rowWBTSTYL2.Item("STYLE_CODE").ToString() & "'").FirstOrDefault().Item("DEFAULT_IMAGE").ToString = "" Then
                    dst.Tables("WBTSTYL1").Select("STYLE_CODE = '" & rowWBTSTYL2.Item("STYLE_CODE").ToString() & "'").FirstOrDefault().Item("DEFAULT_IMAGE") = imagename
                End If
            End If
            'End If
        Next

        ASCMAIN1.Progress("Checking All Styles Status", "")
        For Each rowWBTSTYL1 As DataRow In dst.Tables("WBTSTYL1").Select()
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine("SELECT NVL(STYLE_STATUS,'A') AS STYLE_STATUS")
            SQLS.AppendLine("FROM ICTSTYL1")
            SQLS.AppendLine(String.Format("WHERE STYLE_CODE = '{0}'", rowWBTSTYL1.Item("STYLE_CODE")))
            ASCMAIN1.sql = SQLS.ToString()
            rowWBTSTYL1.Item("STYLE_STATUS") = ASCDATA1.GetDataValue
        Next
        ASCMAIN1.Progress("Refreshing Database", "")
        Update_Record_TDA("WBTSTYL1")
        Update_Record_TDA("WBTSTYL2")
        ASCMAIN1.Progress("", "")

    End Sub

    Sub Setup_WBTSHOP2()
        If grdWBTSHOP1.ActiveRow Is Nothing OrElse (Not grdWBTSHOP1.ActiveRow.IsDataRow Or grdWBTSHOP1.ActiveRow.IsAddRow) Then
            'grdWBTSHOP2.Visible = False
        Else
            Dim dvw As DataView = DirectCast(grdWBTSHOP2.DataSource, DataTable).DefaultView
            Dim STYLE_CODE As String = grdWBTSHOP1.ActiveRow.Cells("STYLE_CODE").Value
            dvw.RowFilter = "STYLE_CODE = '" & STYLE_CODE & "'"
            grdWBTSHOP2.Text = "Colors for" & STYLE_CODE
        End If
    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdWBTSTYL1, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdWBTSHOP1, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Select Case e.Tool.Key

            Case "grdWBTSTYL1"
                ' Nothing 
            Case "grdWBTSHOP1"
                ' Nothing 
            Case Else
                e.Cancel = True
                Exit Sub
        End Select

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

    Private Sub grdWBTSHOP1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWBTSHOP1.AfterRowActivate
        Setup_WBTSHOP2()
        Setup_Missing()
    End Sub

    Private Function CheckForNullGroups() As Boolean
        Dim Retval As Boolean = True
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT COUNT(*) AS REC_CNT")
        SQLS.AppendLine("FROM WBTSTYL1")
        SQLS.AppendLine("WHERE NVL(STYLE_GROUP,999) = 999")
        ASCMAIN1.sql = SQLS.ToString()
        Dim REC_CNT As Int16 = Val(ASCDATA1.GetDataValue)

        If REC_CNT > 0 Then
            Dim iTitle As String = "Groups"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("Items Found To Upload With No Groups")
            iMSG.AppendLine("Consider Using the Re-Set Groups feature.")
            iMSG.AppendLine("Proceed Anyway?")
            Dim iResult As MsgBoxResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = vbYes Then
                Retval = True
            Else
                Retval = False
            End If
        End If

        Return Retval
    End Function

    Private Sub Setup_Missing()
        dst.Tables.Item("WBTSTYL3").Clear()
        For Each rowWBTSTYL2 As DataRow In dst.Tables("WBTSTYL2").Select("ISNULL(IMG_NAME,'') = ''")
            Dim rowWBTSTYL3 As DataRow = dst.Tables("WBTSTYL3").NewRow
            For Each COLUMN As DataColumn In dst.Tables("WBTSTYL2").Columns
                Dim colname As String = COLUMN.ColumnName.ToString
                rowWBTSTYL3.Item(colname) = rowWBTSTYL2.Item(colname)
            Next
            dst.Tables.Item("WBTSTYL3").Rows.Add(rowWBTSTYL3)
        Next
        grdMissing.Refresh()
    End Sub


End Class