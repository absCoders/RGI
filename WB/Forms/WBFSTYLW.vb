Imports System.Xml
Imports System.IO
Imports System.Text
Imports System.Net.Mail
Imports Infragistics.Win.UltraWinGrid
'Imports Microsoft.Office.Interop.Excel
'Imports System.Security.AccessControl

Public Class WBFSTYLW
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
    Private ALT_FUT_QTY_LAST As Int64 = -9999
    Private ALT_FUT_DATE_LAST As Date = DateSerial(1900, 1, 1)

    ' Used for automation 
    Private TASK_NO As String = String.Empty
    Private WBCMAIN1 As New WBCMAIN1
    Dim InAutoMode As Boolean = False
    Dim TickCount As Integer = 0
    'Dim FTPImages As Boolean = True
    'Dim FTPTables As Boolean = False
    Dim UpLoadType As String = ""
    'Const CandidateFilter As String = "WEB_IND = '1' AND ISNULL(UPLOAD_BATCH,'NULL') = 'NULL'"
    Dim AutoProcessRunning As Boolean = False
    Dim MissingZipFiles As New List(Of String)

    'Wayfair Stuff
    'Dim TransferFile As String = String.Format("{0}regency.csv", ASCMAIN1.Folders("Temp"))
    'Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

    Dim RecordsLoaded As Boolean = False
    Private LASTMIN As Int64 = 0
    Private FTP_REMOTE_HOST As String = "regency-rib.com"

    Private DISABLED_STYLES As New List(Of String)

#Region "ABS Standard Routines"
    ' These Routines should be found in all Forms which Launch from the Menu.
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Dim sql As String = String.Empty

        'Make Sure There Are No Orphaned Header Records.
        Dim SU As New StringBuilder With {.Length = 0}
        SU.AppendLine("INSERT INTO WBTSTYLH")
        SU.AppendLine("SELECT")
        SU.AppendLine("WD.STYLE_CODE,")
        SU.AppendLine("NULL AS STYLE_DESC_LONG,")
        SU.AppendLine("NULL AS STYLE_DESC_SHORT,")
        SU.AppendLine("NULL AS SHIPPING_DETAILS,")
        SU.AppendLine("NULL AS SEARCH_OVERRIDE,")
        SU.AppendLine("NULL AS SEARCH_KEYWORDS,")
        SU.AppendLine("NULL AS VIDEO_URL,")
        SU.AppendLine("NULL AS MATERIAL_OVERRIDE,")
        SU.AppendLine("NULL AS MATERIALS,")
        SU.AppendLine("NULL AS META_OVERRIDE,")
        SU.AppendLine("NULL AS META_DESC,")
        SU.AppendLine(String.Format("'{0}' AS INIT_OPER,", ASCMAIN1.USER_ID))
        SU.AppendLine(String.Format("'{0}' AS LAST_OPER,", ASCMAIN1.USER_ID))
        SU.AppendLine("SYSDATE AS INIT_DATE,")
        SU.AppendLine("SYSDATE AS LAST_DATE,")
        SU.AppendLine("NULL AS WEB_DESC")
        SU.AppendLine("FROM WBTSTYLD WD")
        SU.AppendLine("WHERE STYLE_CODE NOT IN")
        SU.AppendLine("(")
        SU.AppendLine("  SELECT DISTINCT STYLE_CODE FROM WBTSTYLH")
        SU.AppendLine(")")
        SU.AppendLine("GROUP BY WD.STYLE_CODE")
        ASCMAIN1.sql = SU.ToString
        ASCDATA1.ExecuteSQL()

        With dst
            Dim sqls As New StringBuilder

            Create_TDA(.Tables.Add, "ICTSTYL3", "*")

            sqls.Length = 0
            sqls.AppendLine("SELECT WBTSTYLD.*,")
            sqls.AppendLine("ICTSTYL1.STYLE_DESC,")
            sqls.AppendLine("DECODE(WBTSTYLH.STYLE_DESC_SHORT,'','0','1') AS HAS_DESC_SHORT,")
            sqls.AppendLine("DECODE(WBTSTYLH.STYLE_DESC_LONG,'','0','1') AS HAS_DESC_LONG,")
            sqls.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
            sqls.AppendLine("ICTSTYL1.VEND_CODE,")
            sqls.AppendLine("NVL(PGC.PAGE_CNT,0) PAGE_CNT,")
            sqls.AppendLine("ATR.ATTR_DESC,")
            sqls.AppendLine("STA.WHSE_QTY_ON_ORDER,")
            sqls.AppendLine("STA.WHSE_QTY_TRAN,")
            sqls.AppendLine("STA.WHSE_TOTAL_FUT,")
            sqls.AppendLine("STA.WHSE_QTY_ON_HAND,")
            sqls.AppendLine("STA.WHSE_QTY_PICK,")
            sqls.AppendLine("STA.OPEN_TO_SELL,")
            sqls.AppendLine("STA.WHSE_QTY_OPEN,")
            sqls.AppendLine("STA.FUT_AVAIL,")
            sqls.AppendLine("NVL(ICTSTYL1.STYLE_SO_QTY_MIN,0) STYLE_SO_QTY_MIN")
            sqls.AppendLine("FROM WBTSTYLD, WBTSTYLH, ICTSTYL1, (SELECT STYLE_CODE, COUNT(PAGE_CODE) AS PAGE_CNT FROM WBTPAGED GROUP BY STYLE_CODE) PGC,")
            sqls.AppendLine("(")
            sqls.AppendLine("   SELECT")
            sqls.AppendLine("   A3.STYLE_CODE,")
            sqls.AppendLine("   MIN(A1.ATTR_DESC) AS ATTR_DESC")
            sqls.AppendLine("   FROM ICTATTR1 A1, ICTSTYL3 A3")
            sqls.AppendLine("   WHERE A1.ATTR_CODE = A3.ATTR_CODE")
            sqls.AppendLine("   AND NVL(A1.ATT_RANK,'0') = '1'")
            sqls.AppendLine("   GROUP BY A3.STYLE_CODE")
            sqls.AppendLine(") ATR,")
            sqls.AppendLine("(")
            sqls.AppendLine("    SELECT")
            sqls.AppendLine("    STYLE_CODE,")
            sqls.AppendLine("    COLOR_CODE,")
            sqls.AppendLine("    NVL(WHSE_QTY_ON_ORDER,0) WHSE_QTY_ON_ORDER,")
            sqls.AppendLine("    NVL(WHSE_QTY_TRAN,0) WHSE_QTY_TRAN,")
            sqls.AppendLine("    (NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)) WHSE_TOTAL_FUT,")
            sqls.AppendLine("    NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND,")
            sqls.AppendLine("    NVL(WHSE_QTY_PICK,0) WHSE_QTY_PICK,")
            sqls.AppendLine("    (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0)) OPEN_TO_SELL,")
            sqls.AppendLine("    NVL(WHSE_QTY_OPEN,0) WHSE_QTY_OPEN,")
            sqls.AppendLine("    (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) FUT_AVAIL")
            sqls.AppendLine("    FROM ICTSTAT2")
            sqls.AppendLine("    WHERE WHSE_CODE = 'MS'")
            sqls.AppendLine(") STA")
            sqls.AppendLine("WHERE WBTSTYLD.STYLE_CODE = WBTSTYLH.STYLE_CODE (+)")
            sqls.AppendLine("AND WBTSTYLD.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            sqls.AppendLine("AND WBTSTYLD.STYLE_CODE = PGC.STYLE_CODE (+)")
            sqls.AppendLine("AND WBTSTYLD.STYLE_CODE = ATR.STYLE_CODE (+)")
            sqls.AppendLine("AND WBTSTYLD.STYLE_CODE = STA.STYLE_CODE (+)")
            sqls.AppendLine("AND WBTSTYLD.COLOR_CODE = STA.COLOR_CODE (+)")
            ASCMAIN1.sql = sqls.ToString
            Create_TDA(dst.Tables.Add, "WBTSTYLD", "**", 0, True)
            With .Tables("WBTSTYLD").Columns
                .Add("FILTER_SEL", GetType(System.String))
                .Add("ALL_DISC", GetType(System.String))
            End With
            'Create_TDA(.Tables.Add, "WBTSTYLD", "*")

            Create_TDA(.Tables.Add, "ICTSTAT2", "*")
            .Tables("ICTSTAT2").Columns.Add("ITEM_CODE", GetType(System.String))

            Create_TDA(.Tables.Add, "WBTRSSF1", "*")

            sqls.Length = 0
            sqls.AppendLine("SELECT * FROM")
            sqls.AppendLine("  (")
            sqls.AppendLine("   SELECT UPPER(C1.STYLE_CODE) AS STYLE_CODE, C1.COLOR_CODE,")
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
            sqls.AppendLine("   GROUP BY UPPER(C1.STYLE_CODE), C1.COLOR_CODE, C2.COLOR_DESC, C1.STYLE_COLOR_STATUS")
            sqls.AppendLine("  )")
            ASCMAIN1.sql = sqls.ToString
            Create_TDA(dst.Tables.Add, "ICTSTYC1", "**", 0, False, "", 2)
            Fill_Records("ICTSTYC1")

            Create_TDA(.Tables.Add, "WBTIMGR1", "*")

            sqls.Length = 0
            sqls.AppendLine("SELECT")
            sqls.AppendLine("'23249' AS Supplier_ID,")
            sqls.AppendLine("(S1.STYLE_CODE || '-' || S2.COLOR_CODE) AS Item_Number,")
            sqls.AppendLine("CASE WHEN SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0))) < 0 THEN")
            sqls.AppendLine("  0")
            sqls.AppendLine("ELSE")
            sqls.AppendLine("  SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)- NVL(S2.WHSE_QTY_PICK,0)))")
            sqls.AppendLine("END AS On_Hand,")
            sqls.AppendLine("0 AS Back_Order,")
            sqls.AppendLine("CASE WHEN SUM((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0))) < 0 THEN")
            sqls.AppendLine("  CASE WHEN SUM(((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) + (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0)))) < 0 THEN")
            sqls.AppendLine("    0")
            sqls.AppendLine("  ELSE")
            sqls.AppendLine("    SUM(((NVL(S2.WHSE_QTY_ON_HAND,0)- NVL(S2.WHSE_QTY_OPEN,0)) + (NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0))))")
            sqls.AppendLine("  END")
            sqls.AppendLine("ELSE")
            sqls.AppendLine("  SUM((NVL(S2.WHSE_QTY_ON_ORDER,0) + NVL(S2.WHSE_QTY_TRAN,0)))")
            sqls.AppendLine("END AS On_Order,")
            sqls.AppendLine("NULL AS NxtAvailDate,")
            sqls.AppendLine("DECODE(S1.STYLE_STATUS,'D',1,0) AS STATUS,")
            sqls.AppendLine("S1.STYLE_DESC AS Description")
            sqls.AppendLine("FROM ICTSTYL1 S1, ICTSTAT2 S2")
            sqls.AppendLine("WHERE S1.STYLE_CODE = S2.STYLE_CODE (+)")
            sqls.AppendLine("AND (S2.WHSE_CODE = 'MS' OR S2.WHSE_CODE = 'CG')")
            sqls.AppendLine("GROUP BY (S1.STYLE_CODE || '-' || S2.COLOR_CODE),")
            sqls.AppendLine("DECODE(S1.STYLE_STATUS,'D',1,0),")
            sqls.AppendLine("S1.STYLE_DESC")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 2)

            sqls.Length = 0
            sqls.AppendLine("SELECT WBTSTYLH.*,")
            sqls.AppendLine("ICTSTYL1.STYLE_DESC,")
            sqls.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
            sqls.AppendLine("ICTSTYL1.STYLE_UOM,")
            sqls.AppendLine("ICTSTYL1.CARTON_PACK_QTY,")
            sqls.AppendLine("ICTSTYL1.INNER_PACK_QTY,")
            sqls.AppendLine("ICTSTYL1.SUB_UNIT_PACK_QTY,")
            sqls.AppendLine("ICTSTYL1.SUB_UNIT_BAG_QTY,")
            sqls.AppendLine("ICTSTYL1.DUTY_RATE_CODE,")
            sqls.AppendLine("ICTSTYL1.STYLE_SO_QTY_MIN,")
            sqls.AppendLine("ICTSTYL1.RESHIPBOX_CODE,")
            sqls.AppendLine("ICTSTYL1.EXCLUSIVE_STYLE,")
            sqls.AppendLine("ICTSTYL1.CARTONS_PER_UNIT")
            sqls.AppendLine("FROM WBTSTYLH, ICTSTYL1")
            sqls.AppendLine("WHERE WBTSTYLH.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            sqls.AppendLine("AND WBTSTYLH.STYLE_CODE = :PARM1")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "WBTSTYLH", "**", 0, True, "V", 1)

            sqls.Length = 0
            sqls.AppendLine("SELECT ICTSTYLD.*, ICTSTYLM.PACK_DESC")
            sqls.AppendLine("FROM ICTSTYLD, ICTSTYLM")
            sqls.AppendLine("WHERE ICTSTYLD.PACK_CODE = ICTSTYLM.PACK_CODE")
            sqls.AppendLine("AND ICTSTYLD.STYLE_CODE = :PARM1")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTYLD", "**", 0, True, "V")

            sqls.Length = 0
            sqls.AppendLine("")
            sqls.AppendLine("SELECT")
            sqls.AppendLine("D1.STYLE_CODE,")
            sqls.AppendLine("H1.PAGE_CODE,")
            sqls.AppendLine("H1.PAGE_NAME,")
            sqls.AppendLine("H1.PAGE_STATUS")
            sqls.AppendLine("FROM WBTPAGEH H1, WBTPAGED D1")
            sqls.AppendLine("WHERE H1.PAGE_CODE = D1.PAGE_CODE")
            sqls.AppendLine("AND STYLE_CODE = :PARM1")
            sqls.AppendLine("AND NVL(H1.PAGE_STATUS,'A') = 'A'")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "WBTPAGEX", "**", 0, False, "V")


            sqls.Length = 0
            sqls.AppendLine("SELECT")
            sqls.AppendLine("'0' SEL,")
            sqls.AppendLine("STYLE_CODE,")
            sqls.AppendLine("COLOR_CODE,")
            sqls.AppendLine("STYLE_DESC,")
            sqls.AppendLine("STYLE_STATUS,")
            sqls.AppendLine("STK,")
            sqls.AppendLine("QTY_AVL")
            sqls.AppendLine("FROM")
            sqls.AppendLine("(")
            sqls.AppendLine("  SELECT")
            sqls.AppendLine("  I1.STYLE_CODE,")
            sqls.AppendLine("  S2.COLOR_CODE,")
            sqls.AppendLine("  I1.STYLE_DESC,")
            sqls.AppendLine("  I1.STYLE_STATUS,")
            sqls.AppendLine("  DECODE(NVL(I1.CUST_CODE,'0'),'0','Stock','Non-Stock') STK,")
            sqls.AppendLine("  SUM((NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))) AS QTY_AVL")
            sqls.AppendLine("  FROM ICTSTYL1 I1, ICTSTAT2 S2")
            sqls.AppendLine("  WHERE I1.STYLE_CODE = S2.STYLE_CODE")
            sqls.AppendLine("  AND I1.STYLE_STATUS = 'A'")
            sqls.AppendLine("  GROUP BY")
            sqls.AppendLine("  I1.STYLE_CODE,")
            sqls.AppendLine("  S2.COLOR_CODE,")
            sqls.AppendLine("  I1.STYLE_DESC,")
            sqls.AppendLine("  I1.STYLE_STATUS,")
            sqls.AppendLine("  DECODE(NVL(I1.CUST_CODE,'0'),'0','Stock','Non-Stock')")
            sqls.AppendLine("  UNION")
            sqls.AppendLine("  SELECT")
            sqls.AppendLine("  I1.STYLE_CODE,")
            sqls.AppendLine("  S2.COLOR_CODE,")
            sqls.AppendLine("  I1.STYLE_DESC,")
            sqls.AppendLine("  I1.STYLE_STATUS,")
            sqls.AppendLine("  DECODE(NVL(I1.CUST_CODE,'0'),'0','Stock','Non-Stock') STK,")
            sqls.AppendLine("  SUM((NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))) AS QTY_AVL")
            sqls.AppendLine("  FROM ICTSTYL1 I1, ICTSTAT2 S2")
            sqls.AppendLine("  WHERE I1.STYLE_CODE = S2.STYLE_CODE")
            sqls.AppendLine("  AND I1.STYLE_STATUS <> 'A'")
            sqls.AppendLine("  GROUP BY")
            sqls.AppendLine("  I1.STYLE_CODE,")
            sqls.AppendLine("  S2.COLOR_CODE,")
            sqls.AppendLine("  I1.STYLE_DESC,")
            sqls.AppendLine("  I1.STYLE_STATUS,")
            sqls.AppendLine("  DECODE(NVL(I1.CUST_CODE,'0'),'0','Stock','Non-Stock')")
            sqls.AppendLine("  HAVING SUM((NVL(S2.WHSE_QTY_ON_HAND,0) + NVL(S2.WHSE_QTY_TRAN,0) + NVL(S2.WHSE_QTY_ON_ORDER,0) - NVL(S2.WHSE_QTY_OPEN,0) - NVL(S2.WHSE_QTY_PICK,0))) > 0")
            sqls.AppendLine(")")
            sqls.AppendLine("WHERE (STYLE_CODE, COLOR_CODE) NOT IN (SELECT STYLE_CODE, COLOR_CODE FROM WBTSTYLD)")
            sqls.AppendLine("ORDER BY")
            sqls.AppendLine("STYLE_CODE,")
            sqls.AppendLine("COLOR_CODE")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTYLX", "**", 0, False)
            Fill_Records("ICTSTYLX")

            'sqls.Length = 0
            'sqls.AppendLine("SELECT")
            'sqls.AppendLine("B.BLLT_CODE,")
            'sqls.AppendLine("B.BLLT_DESC, ")
            'sqls.AppendLine("B.BLLT_RANK,")
            'sqls.AppendLine("S.STYLE_CODE")
            'sqls.AppendLine("FROM ICTBLLT1 B, ICTSTYLB S")
            'sqls.AppendLine("WHERE B.BLLT_CODE = S.BLLT_CODE")
            'sqls.AppendLine("AND S.STYLE_CODE = :PARM1")
            'ASCMAIN1.sql = sqls.ToString()
            'Create_TDA(.Tables.Add, "ICTSTYLB", "**", 0, True, "V")

            sqls.Length = 0
            sqls.AppendLine("SELECT *")
            sqls.AppendLine("FROM ICTSTYL1")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTYL1", "**", 0, False)

            sqls.Length = 0
            sqls.AppendLine("")
            sqls.AppendLine("SELECT * FROM")
            sqls.AppendLine("  (")
            sqls.AppendLine("   SELECT")
            sqls.AppendLine("   UPPER(D1.STYLE_CODE) AS STYLE_CODE,")
            sqls.AppendLine("   UPPER(D1.COLOR_CODE) AS COLOR_CODE,")
            sqls.AppendLine("   (UPPER(D1.STYLE_CODE) || '-' || UPPER(D1.COLOR_CODE)) AS SKU,")
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
            sqls.AppendLine("   CASE S2.WHSE_CODE")
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
            sqls.AppendLine("   CASE WHEN")
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
            sqls.AppendLine("   D1.ALT_FUT_QTY,")
            sqls.AppendLine("   D1.ALT_FUT_DATE")
            sqls.AppendLine("   FROM WBTSTYLD D1")
            sqls.AppendLine("   LEFT JOIN ICTSTAT2 S2")
            sqls.AppendLine("   ON D1.STYLE_CODE  = S2.STYLE_CODE")
            sqls.AppendLine("   AND D1.COLOR_CODE = S2.COLOR_CODE")
            sqls.AppendLine("   WHERE D1.WEB_IND = 'W'")
            sqls.AppendLine("   GROUP BY (UPPER(D1.STYLE_CODE) || '-' || UPPER(D1.COLOR_CODE)), UPPER(D1.STYLE_CODE), UPPER(D1.COLOR_CODE), D1.ALT_FUT_QTY, D1.ALT_FUT_DATE")
            sqls.AppendLine("   ORDER BY (UPPER(D1.STYLE_CODE) || '-' || UPPER(D1.COLOR_CODE))")
            sqls.AppendLine("  )")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTINVTR", "**", 0, False)

            sqls.Length = 0
            sqls.AppendLine("SELECT *")
            sqls.AppendLine("FROM ICTSTDQ1")
            sqls.AppendLine("WHERE WHSE_CODE = 'MS'")
            'sqls.AppendLine("AND STYLE_CODE = :PARM1")
            'sqls.AppendLine("AND COLOR_CODE = :PARM2")
            ASCMAIN1.sql = sqls.ToString()
            Create_TDA(.Tables.Add, "ICTSTDQ1", "**", 0, False)
            'Fill_Records("ICTSTDQ1")
        End With

        grdWBTSTYLD.DataSource = dst.Tables("WBTSTYLD")
        grdICTSTYLD.DataSource = dst.Tables("ICTSTYLD")
        grdWBTPAGEX.DataSource = dst.Tables("WBTPAGEX")
        grdICTSTYLX.DataSource = dst.Tables("ICTSTYLX")
        'grdICTSTYLB.DataSource = dst.Tables("ICTSTYLB")

        MoveToParents()

        Get_PARM("WBTPARM1")

        'WBCMAIN1.DisplayHeaderCheckBox(grdWBTSTYLD, New String() {"WEB_IND"})
        With grdWBTSTYLD.DisplayLayout.Bands(0)
            .Columns("PAGE_CNT").Format = "###,##0"
            .Columns("ALT_FUT_QTY").Format = "###,##0"
            '.Columns("QTY_AVL").Format = "###,##0"
            .Columns("STYLE_SO_QTY_MIN").Format = "###,##0"

            .Columns("WHSE_QTY_ON_ORDER").Format = "###,##0"
            .Columns("WHSE_QTY_ON_ORDER").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_QTY_ON_ORDER").Header.Appearance.BackColor2 = Drawing.Color.Yellow

            .Columns("WHSE_QTY_TRAN").Format = "###,##0"
            .Columns("WHSE_QTY_TRAN").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_QTY_TRAN").Header.Appearance.BackColor2 = Drawing.Color.Yellow

            .Columns("WHSE_TOTAL_FUT").Format = "###,##0"
            .Columns("WHSE_TOTAL_FUT").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_TOTAL_FUT").Header.Appearance.BackColor2 = Drawing.Color.Yellow

            .Columns("WHSE_QTY_ON_HAND").Format = "###,##0"
            .Columns("WHSE_QTY_ON_HAND").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_QTY_ON_HAND").Header.Appearance.BackColor2 = Drawing.Color.LightBlue

            .Columns("WHSE_QTY_PICK").Format = "###,##0"
            .Columns("WHSE_QTY_PICK").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_QTY_PICK").Header.Appearance.BackColor2 = Drawing.Color.LightBlue

            .Columns("OPEN_TO_SELL").Format = "###,##0"
            .Columns("OPEN_TO_SELL").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("OPEN_TO_SELL").Header.Appearance.BackColor2 = Drawing.Color.LightGreen

            .Columns("WHSE_QTY_OPEN").Format = "###,##0"
            .Columns("WHSE_QTY_OPEN").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("WHSE_QTY_OPEN").Header.Appearance.BackColor2 = Drawing.Color.Lime

            .Columns("FUT_AVAIL").Format = "###,##0"
            .Columns("FUT_AVAIL").Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            .Columns("FUT_AVAIL").Header.Appearance.BackColor2 = Drawing.Color.Green

            .Columns("DATE_DISABLED").Format = "MM/dd/yy"
        End With

        Create_Summary(grdWBTSTYLD, "STYLE_CODE", "Count")
        Create_Summary(grdWBTPAGEX, "PAGE_CODE", "Count")
        Create_Summary(grdICTSTYLX, "STYLE_CODE", "Count")

        'ASCMAIN1.Add_Value_List(grdWBTSTYLD, "WEB_IND", , New String() {":", "X:Not On Web", "W:On Web", "U:Awaiting Update", "R:Awaiting Removal"})
        ASCMAIN1.Add_Value_List(grdWBTSTYLD, "WEB_IND", , New String() {":", "I:Inactive", "W:On Web", "U:Awaiting Update"})
        ASCMAIN1.Add_Value_List(grdWBTPAGEX, "PAGE_STATUS", , New String() {":", "A:Active"})

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
        SQLU.AppendLine("FROM WBTSTYLD")
        SQLU.AppendLine("WHERE STYLE_GROUP > 0")
        ASCMAIN1.sql = SQLU.ToString()
        Dim RECCNT As Int16 = Val(ASCDATA1.GetDataValue)

        Dim GROUPS As New List(Of Int64)
        For i As Int64 = 1 To RECCNT
            GROUPS.Add(i)
        Next
        GROUPS.Add(99)
        cboGROUPS.DataSource = GROUPS

        txtMaxGroup.Text = RECCNT

        If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "site.admin" Or ASCMAIN1.USER_ID = "mariog") Then
            chkExportTesting.Visible = True
        Else
            chkExportTesting.Visible = False
        End If

        Sort_grdColumns(grdWBTSTYLD, "STYLE_CODE, COLOR_CODE")

        SetTabModes(0)

        Bind_Controls(splWHTSTYLH, "WHTSTYLH")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim zmsg As String = String.Empty

        Select Case eItemKey
            Case "Load Records"

            Case "Update"
            Case "Finish"
            Case "Done"
            Case "Upload" 'Not used anymore but may be revived
                EMsg = CheckIsAutoMode()
                If EMsg.Length = 0 Then
                    zmsg = "Do you want to Upload the Item Changes?"

                    If EMsg.Length = 0 Then
                        If MessageBox.Show(zmsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                    'Dim iResult As MsgBoxResult
                    'Dim iTitle As String = "Upload Images"
                    'Dim iMSG As New System.Text.StringBuilder

                    'If UpLoadType = "P" Then
                    '    iMSG.AppendLine("Do You Want To Upload Images?")
                    '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                    '    If iResult = MsgBoxResult.Yes Then
                    '        FTPImages = True
                    '    Else
                    '        FTPImages = False
                    '    End If
                    'Else
                    '    FTPImages = False
                    'End If

                    'If UpLoadType = "I" Then
                    '    FTPTables = True
                    'Else
                    '    FTPTables = False
                    'End If
                End If
                'Case "Remove Alt Supplier"
                '    If EMsg.Length = 0 Then
                '        Dim iMsg As New StringBuilder With {.Length = 0}
                '        iMsg.AppendLine("This Will Remove All Alternate")
                '        iMsg.AppendLine("Qty and Dates From The Supplier")
                '        iMsg.AppendLine("You Select. Please Make Sure You")
                '        iMsg.AppendLine("Save Any Changes To The Grid Before")
                '        iMsg.AppendLine("You Proceed.")
                '        iMsg.AppendLine("")
                '        iMsg.AppendLine("Are You Ready?")
                '        Dim iResult As MsgBoxResult = MsgBox(iMsg.ToString, MsgBoxStyle.YesNo, "Remove Alt Supplier")
                '        If iResult <> MsgBoxResult.Yes Then
                '            EMsg &= vbCr & "Remove Alt Supplier Cancelled"
                '        End If
                '    End If
        End Select

        If EMsg <> String.Empty Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Load Records"
                Clear_Record()
                Load_Record()
                Call Mode_Settings(False)
                SetTabModes(1)
                RecordsLoaded = True
            Case "Update"
                Update_Record()
            Case "Finish"
                Update_Masterfile()
                SetTabModes(1)
            Case "Done"
                Update_Record()
                Call Mode_Settings(False)
                SetTabModes(0)
                RecordsLoaded = False
            Case "The Rules"
                Dim rules As New StringBuilder With {.Length = 0}
                rules.AppendLine("-- Not On The Web --")
                rules.AppendLine("   * These are items that have been added to the screen but ARE NOT currently on")
                rules.AppendLine("     the website.")
                rules.AppendLine("   * You may remove these items from the screen by right-clicking on them and")
                rules.AppendLine("     selecting 'Remove From Screen'")
                rules.AppendLine("")
                rules.AppendLine("-- Awaiting Update --")
                rules.AppendLine("   * These are items that ARE NOT currently on the web but are waiting to be")
                rules.AppendLine("      transferred to the web when their batch is updated.  At that time, if")
                rules.AppendLine("      there is ample inventorythey will be sent to the web.")
                rules.AppendLine("   * You may only set their status to 'Not on Web' or right-click to")
                rules.AppendLine("     'Add To Web Immediately'.")
                rules.AppendLine("")
                rules.AppendLine("-- Awaiting Removal --")
                rules.AppendLine("   * These are items that ARE currently on the web awaiting removal.")
                rules.AppendLine("     They will be removed from the web the next time their batch is updated.")
                rules.AppendLine("   * You may only right-click and select 'Remove From Web Immediately'.")
                rules.AppendLine("")
                rules.AppendLine("-- On The Web --")
                rules.AppendLine("   * These are items that ARE currently on the web.")
                rules.AppendLine("   * They will be updated with current infomation including inventory the next")
                rules.AppendLine("     time their batch is updated.  If their future availablility falls below")
                rules.AppendLine("     zero they will be removed from the web.")
                rules.AppendLine("   * You may only set their status to 'waiting Removal' or right-click and")
                rules.AppendLine("     select 'Remove From Web Immediately'.")
                MsgBox(rules.ToString, vbOKOnly, "Web Items Status Rules")
                'Case "Remove Alt Supplier"
                '    Dim S As New Text.StringBuilder With {.Length = 0}
                '    S.AppendLine("SELECT VEND_CODE, VEND_NAME")
                '    S.AppendLine("FROM APTVEND1")
                '    S.AppendLine("WHERE VEND_TYPE = 'S'")
                '    With ASCMAIN1.CodeSelector
                '        .SQL = S.ToString
                '        .MultipleSelections = False
                '        .PreviouslySelectedCodes0 = ""
                '        .Caption = "Suppliers"
                '        .TABLE_NAME = ""
                '        .VIEW_NAME = ""
                '        .VIEW_DESC = ""
                '        .COLUMN_NAME = ""
                '        .COLUMN_PREKEYs = New Dictionary(Of String, String)
                '        .Custom_sql_where = ""
                '        .tblASTVIEW1 = New DataTable
                '    End With
                '    Dim F As New ASFCODE1
                '    F.ShowDialog()
                '    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                '        Dim VEND_CODE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("VEND_CODE") & ""
                '        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
                '        SQLS.AppendLine("")
                '        SQLS.AppendLine("UPDATE WBTSTYLD")
                '        SQLS.AppendLine("SET ALT_FUT_QTY = NULL,")
                '        SQLS.AppendLine("ALT_FUT_DATE = NULL")
                '        SQLS.AppendLine("WHERE (STYLE_CODE, COLOR_CODE) IN")
                '        SQLS.AppendLine("(")
                '        SQLS.AppendLine("  SELECT")
                '        SQLS.AppendLine("  WD.STYLE_CODE, WD.COLOR_CODE")
                '        SQLS.AppendLine("  FROM ICTSTYL1 S1, WBTSTYLD WD")
                '        SQLS.AppendLine("  WHERE S1.STYLE_CODE = WD.STYLE_CODE")
                '        SQLS.AppendLine(String.Format("  AND S1.VEND_CODE = '{0}'", VEND_CODE))
                '        SQLS.AppendLine("  AND NVL(ALT_FUT_QTY,0) > 0")
                '        SQLS.AppendLine(")")
                '        ASCMAIN1.sql = SQLS.ToString
                '        ASCDATA1.ExecuteSQL()

                '        Application.DoEvents()
                '        MsgBox("Alternates Updated.  Please Wait While Data Is Refreshed", vbOKOnly, "Updated")
                '        Clear_Record()
                '        Application.DoEvents()
                '        Load_Record()
                '        Application.DoEvents()
                '    End If
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)


        'With UltraExplorerBar1.Groups("Screen Control")
        '    .Items("Load Records").Settings.Enabled = not_iScreenMode
        '    .Items("Update").Settings.Enabled = not_iScreenMode
        '    .Items("Finish").Settings.Enabled = iScreenMode
        '    .Items("Done").Settings.Enabled = iScreenMode
        'End With

        'UltraExplorerBar1.Groups("Auto Refresh").Visible = dst.Tables("WBTSTYLD").Rows.Count() > 0

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            WebBrowser1.Visible = True
            grdWBTSTYLD.Visible = False
            splWHTSTYLH.Visible = False
        Else
            WebBrowser1.Visible = False
            grdWBTSTYLD.Visible = True
            splWHTSTYLH.Visible = False
        End If

        With grdWBTSTYLD.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.True
        End With
        For i As Integer = 0 To grdWBTSTYLD.DisplayLayout.Bands(0).Columns.Count - 1
            grdWBTSTYLD.DisplayLayout.Bands(0).Columns(i).CellActivation = UltraWinGrid.Activation.NoEdit
        Next i
        For Each COLNAME As String In New String() {"UPLOAD_BATCH", "STYLE_SORT", "UPLOAD_IMG", "FULL_UPLOAD", "ALT_FUT_QTY", "ALT_FUT_DATE", "FLAG_NEW"}
            grdWBTSTYLD.DisplayLayout.Bands(0).Columns(COLNAME).CellActivation = UltraWinGrid.Activation.AllowEdit
        Next
        For Each COLNAME As String In New String() {"UPLOAD_BATCH", "STYLE_SORT", "ALT_FUT_QTY"}
            grdWBTSTYLD.DisplayLayout.Bands(0).Columns(COLNAME).CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Next
    End Sub

    Sub Clear_Record()

        dst.Tables("ICTSTYC1").Rows.Clear()
        dst.Tables("WBTSTYLH").Rows.Clear()
        dst.Tables("WBTSTYLD").Rows.Clear()
        dst.Tables("WBTRSSF1").Rows.Clear()
        dst.Tables("ICTSTATX").Rows.Clear()
        dst.Tables("WBTPAGEX").Rows.Clear()

        Absx1.txtFor("STYLE_CODE").Clear()
        Absx1.txtFor("STYLE_DESC").Clear()

        shopSiteFilename = String.Empty
        itemUploaded = False

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor

        Fill_Records("ICTSTYL1")

        Fill_Records("ICTSTYC1")

        Fill_Records("ICTSTATX")

        Fill_Records("ICTSTDQ1")

        Dim SQLW As New StringBuilder
        SQLW.AppendLine("SELECT WBTSTYLD.*,")
        SQLW.AppendLine("ICTSTYL1.STYLE_DESC,")
        SQLW.AppendLine("DECODE(WBTSTYLH.STYLE_DESC_SHORT,'','0','1') AS HAS_DESC_SHORT,")
        SQLW.AppendLine("DECODE(WBTSTYLH.STYLE_DESC_LONG,'','0','1') AS HAS_DESC_LONG,")
        SQLW.AppendLine("ICTSTYL1.STYLE_CLASS_CODE,")
        SQLW.AppendLine("ICTSTYL1.VEND_CODE,")
        SQLW.AppendLine("NVL(PGC.PAGE_CNT,0) PAGE_CNT,")
        SQLW.AppendLine("ATR.ATTR_DESC,")
        SQLW.AppendLine("STA.WHSE_QTY_ON_ORDER,")
        SQLW.AppendLine("STA.WHSE_QTY_TRAN,")
        SQLW.AppendLine("STA.WHSE_TOTAL_FUT,")
        SQLW.AppendLine("STA.WHSE_QTY_ON_HAND,")
        SQLW.AppendLine("STA.WHSE_QTY_PICK,")
        SQLW.AppendLine("STA.OPEN_TO_SELL,")
        SQLW.AppendLine("STA.WHSE_QTY_OPEN,")
        SQLW.AppendLine("STA.FUT_AVAIL,")
        SQLW.AppendLine("NVL(ICTSTYL1.STYLE_SO_QTY_MIN,0) STYLE_SO_QTY_MIN")
        SQLW.AppendLine("FROM WBTSTYLD, WBTSTYLH, ICTSTYL1, (SELECT STYLE_CODE, COUNT(PAGE_CODE) AS PAGE_CNT FROM WBTPAGED GROUP BY STYLE_CODE) PGC,")
        SQLW.AppendLine("(")
        SQLW.AppendLine("   SELECT")
        SQLW.AppendLine("   A3.STYLE_CODE,")
        SQLW.AppendLine("   MIN(A1.ATTR_DESC) AS ATTR_DESC")
        SQLW.AppendLine("   FROM ICTATTR1 A1, ICTSTYL3 A3")
        SQLW.AppendLine("   WHERE A1.ATTR_CODE = A3.ATTR_CODE")
        SQLW.AppendLine("   AND NVL(A1.ATT_RANK,'0') = '1'")
        SQLW.AppendLine("   GROUP BY A3.STYLE_CODE")
        SQLW.AppendLine(") ATR,")
        SQLW.AppendLine("(")
        SQLW.AppendLine("    SELECT")
        SQLW.AppendLine("    STYLE_CODE,")
        SQLW.AppendLine("    COLOR_CODE,")
        SQLW.AppendLine("    NVL(WHSE_QTY_ON_ORDER,0) WHSE_QTY_ON_ORDER,")
        SQLW.AppendLine("    NVL(WHSE_QTY_TRAN,0) WHSE_QTY_TRAN,")
        SQLW.AppendLine("    (NVL(WHSE_QTY_ON_ORDER,0) + NVL(WHSE_QTY_TRAN,0)) WHSE_TOTAL_FUT,")
        SQLW.AppendLine("    NVL(WHSE_QTY_ON_HAND,0) WHSE_QTY_ON_HAND,")
        SQLW.AppendLine("    NVL(WHSE_QTY_PICK,0) WHSE_QTY_PICK,")
        SQLW.AppendLine("    (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0)) OPEN_TO_SELL,")
        SQLW.AppendLine("    NVL(WHSE_QTY_OPEN,0) WHSE_QTY_OPEN,")
        SQLW.AppendLine("    (NVL(WHSE_QTY_ON_HAND,0) - NVL(WHSE_QTY_PICK,0) + NVL(WHSE_QTY_TRAN,0) + NVL(WHSE_QTY_ON_ORDER,0) - NVL(WHSE_QTY_OPEN,0)) FUT_AVAIL")
        SQLW.AppendLine("    FROM ICTSTAT2")
        SQLW.AppendLine("    WHERE WHSE_CODE = 'MS'")
        SQLW.AppendLine(") STA")
        SQLW.AppendLine("WHERE WBTSTYLD.STYLE_CODE = WBTSTYLH.STYLE_CODE (+)")
        SQLW.AppendLine("AND WBTSTYLD.STYLE_CODE = ICTSTYL1.STYLE_CODE")
        SQLW.AppendLine("AND WBTSTYLD.STYLE_CODE = PGC.STYLE_CODE (+)")
        SQLW.AppendLine("AND WBTSTYLD.STYLE_CODE = ATR.STYLE_CODE (+)")
        SQLW.AppendLine("AND WBTSTYLD.STYLE_CODE = STA.STYLE_CODE (+)")
        SQLW.AppendLine("AND WBTSTYLD.COLOR_CODE = STA.COLOR_CODE (+)")
        Fill_Records("WBTSTYLD", , , SQLW.ToString)


        UpdateInventory()

        Dim updatesFound As Boolean = UpdateStatus()
        If updatesFound Then
            BeginTrans()
            Update_Record_TDA("WBTSTYLD")
            CommitTrans()
        End If
        markAllDisc()

        Me.Cursor = Cursors.Default

        'MsgBox("Data Loaded.", MsgBoxStyle.Information, "Success")
    End Sub

    Private Sub markAllDisc()
        Dim STYLE_DISC As New Dictionary(Of String, String)
        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("", "STYLE_CODE")
            Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
            If Not STYLE_DISC.Keys.Contains(STYLE_CODE) Then
                STYLE_DISC.Add(STYLE_CODE, "0")
            End If
        Next
        For Each SK As KeyValuePair(Of String, String) In STYLE_DISC
            Dim FLT As String = $"STYLE_CODE = '{SK.Key}'"
            Dim ALL_DISC As String = "1"
            For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(FLT, "STYLE_CODE")
                Dim STYLE_STATUS As String = rowWBTSTYLD.Item("STYLE_STATUS").ToString & String.Empty
                Dim FUT_AVAIL As Int64 = Val(rowWBTSTYLD.Item("FUT_AVAIL").ToString & String.Empty)
                If STYLE_STATUS <> "D" Or FUT_AVAIL > 0 Then
                    ALL_DISC = "0"
                    Exit For
                End If
            Next
            Dim FLT2 As String = $"STYLE_CODE = '{SK.Key}'"
            For Each rowWBTSTYLD2 As DataRow In dst.Tables("WBTSTYLD").Select(FLT2, "STYLE_CODE")
                rowWBTSTYLD2.Item("ALL_DISC") = ALL_DISC
            Next
        Next
        For Each grow As UltraWinGrid.UltraGridRow In grdWBTSTYLD.Rows
            If grow.Cells.Item("ALL_DISC").Value = "1" And grow.Cells.Item("WEB_IND").Value <> "I" Then
                grow.Cells.Item("ALL_DISC").Appearance.BackColor = Drawing.Color.Salmon
            Else
                grow.Cells.Item("ALL_DISC").Appearance.BackColor = Drawing.Color.Empty
            End If
            If IsNumeric(grow.Cells.Item("FUT_AVAIL").Value) And IsNumeric(grow.Cells.Item("STYLE_SO_QTY_MIN").Value) Then
                If Val(grow.Cells.Item("FUT_AVAIL").Value) > 0 Then
                    If Val(grow.Cells.Item("FUT_AVAIL").Value) < Val(grow.Cells.Item("STYLE_SO_QTY_MIN").Value) Then
                        grow.Cells.Item("STYLE_SO_QTY_MIN").Appearance.BackColor = Drawing.Color.MediumVioletRed
                    Else
                        grow.Cells.Item("STYLE_SO_QTY_MIN").Appearance.BackColor = Drawing.Color.Empty
                    End If
                Else
                    grow.Cells.Item("STYLE_SO_QTY_MIN").Appearance.BackColor = Drawing.Color.Empty
                End If
            End If
        Next
    End Sub

    Private Function UpdateStatus() As Boolean
        Dim RetVal As Boolean = False
        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
            Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowWBTSTYLD.Item("COLOR_CODE").ToString & String.Empty
            'If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            '    If STYLE_CODE = "MT25031" Then Stop
            'End If
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(filter).FirstOrDefault
            If Not IsNothing(rowICTSTYC1) Then
                If rowWBTSTYLD.Item("STYLE_STATUS") <> rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & String.Empty Then
                    RetVal = True
                    rowWBTSTYLD.Item("STYLE_STATUS") = rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & String.Empty
                    If rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString & String.Empty <> "A" Then
                        rowWBTSTYLD.Item("ALT_FUT_QTY") = Null
                        rowWBTSTYLD.Item("ALT_FUT_DATE") = Null
                    End If
                End If
            End If
        Next
        grdWBTSTYLD.UpdateData()
        grdWBTSTYLD.Refresh()
        Return RetVal
    End Function
    Private Sub UpdateInventory()
        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
            Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowWBTSTYLD.Item("COLOR_CODE").ToString & String.Empty
            If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                'If STYLE_CODE = "MTX53294" And COLOR_CODE = "APGR" Then Stop
            End If
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            Dim rowICTSTYC1 As DataRow = dst.Tables("ICTSTYC1").Select(filter).FirstOrDefault
            If Not IsNothing(rowICTSTYC1) Then
                rowWBTSTYLD.Item("CURR_ON_HAND") = Val(rowICTSTYC1.Item("MSOH").ToString & String.Empty) 'Val(rowICTSTYC1.Item("MSFT").ToString & String.Empty)
            End If
        Next
        grdWBTSTYLD.UpdateData()
        grdWBTSTYLD.Refresh()
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
            Update_Record_TDA("WBTSTYLH")
            Update_Record_TDA("WBTSTYLD")

            Dim SQLS As New System.Text.StringBuilder

            SQLS.Length = 0
            SQLS.AppendLine("DELETE FROM ICTSTYLW")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            SQLS.Length = 0
            SQLS.AppendLine("INSERT INTO ICTSTYLW")
            SQLS.AppendLine("SELECT DISTINCT ICTSTYL1.STYLE_CODE,")
            SQLS.AppendLine("ICTSTYL1.STYLE_CLASS_CODE")
            SQLS.AppendLine("FROM WBTSTYLD, ICTSTYL1")
            SQLS.AppendLine("WHERE WBTSTYLD.STYLE_CODE = ICTSTYL1.STYLE_CODE")
            SQLS.AppendLine("AND NVL(ICTSTYL1.STYLE_CLASS_CODE,'NULL') <> 'NULL'")
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()

            UpdateWBTPARM1()

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

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Call Load_Popup_Menu(grdWBTSTYLD, "SSBBBB", "Show Filter", "Show GroupBox", "Add To Web Immediately", "Remove From Web Immediately", "Style Status Inquiry", "Remove Style")
        Call Load_Popup_Menu(grdWBTSTYLD, "SSBBBBBBBB", "Show Filter", "Show GroupBox", "Create XML For This Style", "Select All For Full Upload", "Select None For Full Upload", "Select All As New", "Clear All New", "Move Selected To Inactive Group")
        Call Load_Popup_Menu(grdICTSTYLX, "SSB", "Show Filter", "Show GroupBox", "Add Selected Styles")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdWBTSTYLD"
                tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                tlb_btn = DirectCast(tlb_pop.Tools("Move Selected To Inactive Group"), UltraWinToolbars.ButtonTool)
                Dim WEB_IND As New List(Of String)
                Dim ALL_DISC As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdWBTSTYLD.Selected.Rows
                    If Not IsNothing(grow.Cells) Then
                        If Not WEB_IND.Contains(grow.Cells.Item("WEB_IND").Value) Then
                            WEB_IND.Add(grow.Cells.Item("WEB_IND").Value)
                        End If
                        If Not ALL_DISC.Contains(grow.Cells.Item("ALL_DISC").Value) Then
                            ALL_DISC.Add(grow.Cells.Item("ALL_DISC").Value)
                        End If
                    End If
                Next
                If WEB_IND.Count = 1 And ALL_DISC.Count = 1 Then
                    If WEB_IND(0) = "W" And ALL_DISC(0) = "1" Then
                        If (ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog") Then
                            tlb_btn.SharedProps.Visible = True
                        Else
                            tlb_btn.SharedProps.Visible = False
                        End If
                    Else
                        tlb_btn.SharedProps.Visible = False
                    End If
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdICTSTYLX"
            Case Else

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

            'Case "Add To Web Immediately"
            '    StyleImmediate("A")
            'Case "Remove From Web Immediately"
            '    StyleImmediate("R")
            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Move Selected To Inactive Group"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Inactive Group"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Move All Selected Styles")
                iMSG.AppendLine("And Their Related Colors To The")
                iMSG.AppendLine("Inactive Group.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is This What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    Dim STYLE_LIST As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdWBTSTYLD.Selected.Rows
                        If Not STYLE_LIST.Contains(grow.Cells.Item("STYLE_CODE").Text) Then
                            STYLE_LIST.Add(grow.Cells.Item("STYLE_CODE").Text)
                        End If
                    Next
                    If STYLE_LIST.Count > 0 Then
                        For Each STYLE_CODE As String In STYLE_LIST
                            For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select($"STYLE_CODE = '{STYLE_CODE}'")
                                rowWBTSTYLD.Item("WEB_IND") = "I"
                                rowWBTSTYLD.Item("STYLE_GROUP") = "999"
                            Next
                        Next
                    End If
                    Update_Record(False)
                    Application.DoEvents()
                    MsgBox("Style(s) Updated.  Please Wait While Data Is Refreshed", vbOKOnly, "Added")
                    Clear_Record()
                    Application.DoEvents()
                    Load_Record()
                    Application.DoEvents()
                End If
            Case "Remove Style From Web"
                'This was removed from menu, so you should never get here.
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                'Dim AllX As Boolean = True
                'For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                '    If rowWBTSTYLD.Item("WEB_IND").ToString & "" <> "X" Then
                '        AllX = False
                '    End If
                'Next
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Remove Style From Web"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine(String.Format("This Will Remove {0} And All It's Accociated", STYLE_CODE))
                iMSG.AppendLine("Colors From The Web Upload Process.  You May Add It Later")
                iMSG.AppendLine("If You Wish To Include It Again.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    BeginTrans()
                    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE))
                        rowWBTSTYLD.Delete()
                    Next
                    Dim SQLS As New StringBuilder With {.Length = 0}
                    SQLS.AppendLine(String.Format("DELETE FROM WBTSTYLH WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                    Update_Record_TDA("WBTSTYLD")
                    CommitTrans()
                End If
            Case "Create XML For This Style"
                If Not ASCMAIN1.Running_in_VS Then
                    MsgBox("Only Wayne Runs This Feature For Now", vbOKOnly, "Nope")
                Else
                    Stop
                    Dim STYLE_UPDATE As String = grd.ActiveRow.Cells("STYLE_CODE").Value
                    'If CreateProductXml(STYLE_UPDATE, "A", False, True, 99) Then
                    '    For Each STYLE_CODE As String In styleList
                    '        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("STYLE_CODE = '" & STYLE_CODE & "'")
                    '            rowWBTSTYLD.Item("LAST_ON_HAND") = rowWBTSTYLD.Item("CURR_ON_HAND")
                    '            rowWBTSTYLD.Item("LAST_UPDATE") = Now()
                    '            rowWBTSTYLD.Item("LAST_UPDATE_REMARKS") = "Right Click From Grid"
                    '        Next
                    '    Next
                    'End If

                    'MsgBox("Shopsite Created.  Upload File Below To Site", vbOKOnly, "Complete")
                End If
            Case "Add Selected Styles"
                MsgBox("Waiting On Wayne", vbOKOnly, "New Feature")
            Case "Select All For Full Upload"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Select All For Full Upload"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Mark All Style/Colors")
                iMSG.AppendLine("To Be Uploaded On The Next")
                iMSG.AppendLine("Full Upload Push.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    'For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
                    '    rowWBTSTYLD.Item("FULL_UPLOAD") = "1"
                    'Next
                    For Each grow As UltraGridRow In grdWBTSTYLD.Rows
                        If Not grow.IsFilteredOut Then
                            grow.Cells.Item("FULL_UPLOAD").Value = "1"
                        End If
                    Next
                    grdWBTSTYLD.UpdateData()
                    grdWBTSTYLD.Refresh()
                    Me.Cursor = Cursors.Default
                End If
            Case "Select None For Full Upload"
                Dim iResult As MsgBoxResult
                Dim iTitle As String = "Select All For Full Upload"
                Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
                iMSG.AppendLine("This Will Mark All Style/Colors")
                iMSG.AppendLine("To Be Removed From The Next")
                iMSG.AppendLine("Full Upload Push.")
                iMSG.AppendLine("")
                iMSG.AppendLine("Is That What You Want?")
                iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
                If iResult = MsgBoxResult.Yes Then
                    Me.Cursor = Cursors.WaitCursor
                    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
                        rowWBTSTYLD.Item("FULL_UPLOAD") = "0"
                    Next
                    Me.Cursor = Cursors.Default
                End If
            'Case "Use Last Alt Vals"
            '    If (ALT_FUT_DATE_LAST = DateSerial(1900, 1, 1) Or ALT_FUT_QTY_LAST = -9999) Then
            '        Dim iTitle As String = "Use Last Alt Vals"
            '        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            '        iMSG.AppendLine("You Must Set Alt Values Once")
            '        iMSG.AppendLine("Before Using this Feature.")
            '        MsgBox(iMSG.ToString(), MsgBoxStyle.OkOnly, iTitle)
            '    Else
            '        If grd.Selected.Rows.Count = 0 Then
            '            grd.ActiveRow.Cells("ALT_FUT_QTY").Value = ALT_FUT_QTY_LAST
            '            grd.ActiveRow.Cells("ALT_FUT_DATE").Value = ALT_FUT_DATE_LAST
            '        Else
            '            For Each thisRow As UltraGridRow In grd.Selected.Rows
            '                thisRow.Cells("ALT_FUT_QTY").Value = ALT_FUT_QTY_LAST
            '                thisRow.Cells("ALT_FUT_DATE").Value = ALT_FUT_DATE_LAST
            '            Next
            '        End If
            '    End If
            'Case "Clear All Alt Vals"
            '    Dim iResult As MsgBoxResult
            '    Dim iTitle As String = "Clear All Alt Vals"
            '    Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            '    iMSG.AppendLine("This Will Clear All Alternate")
            '    iMSG.AppendLine("Dates And Qty From The System")
            '    iMSG.AppendLine("And CAN NOT Be Undone.")
            '    iMSG.AppendLine("")
            '    iMSG.AppendLine("Is That What You Want?")
            '    iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            '    If iResult = MsgBoxResult.Yes Then
            '        Me.Cursor = Cursors.WaitCursor
            '        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
            '            rowWBTSTYLD.Item("ALT_FUT_QTY") = Null
            '            rowWBTSTYLD.Item("ALT_FUT_DATE") = Null
            '        Next
            '        Me.Cursor = Cursors.Default
            '    End If
            Case "Select All As New"
                For Each thisRow As UltraGridRow In grd.Selected.Rows
                    thisRow.Cells("FLAG_NEW").Value = "1"
                Next
                grdWBTSTYLD.UpdateData()
                grdWBTSTYLD.Refresh()
            Case "Clear All New"
                For Each thisRow As UltraGridRow In grd.Selected.Rows
                    thisRow.Cells("FLAG_NEW").Value = "0"
                Next
                grdWBTSTYLD.UpdateData()
                grdWBTSTYLD.Refresh()
                Me.Cursor = Cursors.Default
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub btnAddStyles_Click(sender As Object, e As EventArgs) Handles btnAddStyles.Click
        If Not RecordsLoaded Then
            MsgBox("You 1st Have To Load Records Before Adding New Ones", vbOKOnly, "Adding Styles")
        Else
            AddStyleColors()
        End If
    End Sub

    'Private Sub btnUpdateShopsite_Click(sender As Object, e As EventArgs) Handles btnUpdateShopsite.Click
    '    Me.Cursor = Cursors.WaitCursor
    '    Dim DoneMsg As String = "Shopsite Created.  Upload File Below To Site."
    '    If chkInventoryFeed.Checked Then
    '        DoneMsg = "Inventory File Loaded To Shopsite."
    '        uploadShopsiteInventory()
    '    Else
    '        Dim LAST_UPDATE As DateTime = Now()
    '        Val(cboGROUPS.Text)
    '        Dim UploadInventoryOnly As Boolean = True
    '        If chkFullUpload.Checked Then
    '            UploadInventoryOnly = False
    '        End If

    '        Dim UpdatePricing As Boolean = False
    '        If chkUpdatePricing.Checked Then
    '            UpdatePricing = True
    '        End If

    '        If chkArchive.Checked Then
    '            Dim di As New DirectoryInfo(WB_PARM_PRODUCTS_DIR)
    '            Dim fiArr As FileInfo() = di.GetFiles()
    '            Dim fri As FileInfo
    '            For Each fri In fiArr
    '                Console.WriteLine(fri.Name)
    '                System.IO.File.Move(WB_PARM_PRODUCTS_DIR & fri.Name, WB_PARM_PRODUCTS_DIR & "Archives\" & fri.Name)
    '            Next
    '        End If

    '        If CreateProductXml("", "A", False, UploadInventoryOnly, Val(cboGROUPS.Text), UpdatePricing) Then
    '            If Not (chkFullUpload.Checked = True Or chkUseFilter.Checked = True) Then
    '                WebBrowser1.Visible = True
    '                grdWBTSTYLD.Visible = False
    '                Call FTPProducts()
    '                WebBrowser1.Visible = False
    '                grdWBTSTYLD.Visible = True

    '                'Put this back if you ever want to get upload images put on your laptop.
    '                'For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
    '                '    Dim IMAGE_NAME As String = rowWBTSTYLD.Item("DEFAULT_IMAGE").ToString & String.Empty
    '                '    Dim LOCALPATH As String = "Z:\Wayne On My Mac\Dropbox\Regency\Shopsite\WebImages\"
    '                '    Dim IMAGEPATH As String = "\\192.168.110.221\Shared\Images\"
    '                '    If IO.File.Exists(LOCALPATH & IMAGE_NAME) Then
    '                '        IO.File.Delete(LOCALPATH & IMAGE_NAME)
    '                '    End If
    '                '    If Not IO.File.Exists(IMAGEPATH & IMAGE_NAME) Then
    '                '        Stop
    '                '    End If
    '                '    IO.File.Copy(IMAGEPATH & IMAGE_NAME, LOCALPATH & IMAGE_NAME)
    '                'Next

    '                'End If

    '                'For Each STYLE_CODE As String In styleListInactive
    '                '    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("STYLE_CODE = '" & STYLE_CODE & "'")
    '                '        rowWBTSTYLD.Item("LAST_ON_HAND") = rowWBTSTYLD.Item("CURR_ON_HAND")
    '                '        rowWBTSTYLD.Item("LAST_UPDATE") = LAST_UPDATE
    '                '        rowWBTSTYLD.Item("WEB_IND") = "X"
    '                '        rowWBTSTYLD.Item("LAST_UPDATE_REMARKS") = "Removed From Site - No Inv"
    '                '    Next
    '                'Next

    '                For Each STYLE_CODE As String In styleList
    '                    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("STYLE_CODE = '" & STYLE_CODE & "'")
    '                        rowWBTSTYLD.Item("LAST_ON_HAND") = rowWBTSTYLD.Item("CURR_ON_HAND")
    '                        rowWBTSTYLD.Item("LAST_UPDATE") = LAST_UPDATE
    '                        rowWBTSTYLD.Item("LAST_UPDATE_REMARKS") = "Update Shopsite Button"
    '                    Next
    '                Next
    '            End If
    '        End If
    '    End If
    '    Me.Cursor = Cursors.Default
    '    MsgBox(DoneMsg, vbOKOnly, "Complete")
    'End Sub

    Private Sub uploadShopsiteInventory()
        ASCMAIN1.Progress("Uploading Invetory", Now.ToShortTimeString)
        Dim RemotePath As String = "www/inventory"
        Dim FileName As String = "inventory.csv"
        Dim str As New StringBuilder
        Dim sql As New StringBuilder With {.Length = 0}
        Fill_Records("ICTINVTR")
        Fill_Records("ICTSTDQ1")

        str.Append(Chr(34) & "SKU" & Chr(34) & ",")
        str.Append(Chr(34) & "INVENTORY" & Chr(34) & ",")
        str.Append(Chr(34) & "FUTURE" & Chr(34) & ",")
        str.Replace(",", vbNewLine, str.Length - 1, 1)
        For Each rowICTINVTR As DataRow In dst.Tables("ICTINVTR").Select()
            'Fill_Records("ICTSTDQ1", New String() {rowICTINVTR.Item("STYLE_CODE").ToString & String.Empty, rowICTINVTR.Item("COLOR_CODE").ToString & String.Empty})
            Dim STYLE_CODE As String = rowICTINVTR.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowICTINVTR.Item("COLOR_CODE").ToString & String.Empty
            Dim SKU As String = rowICTINVTR.Item("SKU").ToString & String.Empty
            str.Append(Chr(34) & SKU & Chr(34) & ",")
            Dim CURR_QTY_AVAIL As Int64 = 0
            Dim FUT_QTY_AVAIL As Int64 = 0
            Dim FUT_DATE As String = ""
            Dim MSOH As Int64 = 0
            If IsNumeric(rowICTINVTR.Item("MSOH").ToString & String.Empty) Then
                MSOH = Val(rowICTINVTR.Item("MSOH").ToString & String.Empty)
            End If
            Dim SFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            '    If STYLE_CODE = "MTX66844" Then Stop
            'End If

            For Each rowICTSTDQ1 As DataRow In dst.Tables.Item("ICTSTDQ1").Select(SFilter, "STATUS_DATE")
                If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) Then
                    If CDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) <= Now().AddDays(1) Then
                        CURR_QTY_AVAIL = CURR_QTY_AVAIL + Val(rowICTSTDQ1.Item("QTY_ATS").ToString & String.Empty)
                    Else
                        If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) Then
                            FUT_DATE = CDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty).ToShortDateString
                            FUT_QTY_AVAIL = FUT_QTY_AVAIL + Val(rowICTSTDQ1.Item("QTY_ATS").ToString & String.Empty)
                        End If
                    End If
                End If
            Next
            If CURR_QTY_AVAIL = 0 And MSOH > 0 Then
                'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
                CURR_QTY_AVAIL = MSOH
            End If

            'If STYLE_CODE = "MTF24040" Then
            '    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            'End If


            'Lower Inventory For Items Not Divisable by MOQ.
            Dim MOQ As Int64 = Val(dst.Tables.Item("ICTSTYL1").Select($"STYLE_CODE = '{STYLE_CODE}'").FirstOrDefault.Item("STYLE_SO_QTY_MIN").ToString & String.Empty)
            If MOQ > 0 And CURR_QTY_AVAIL > 0 Then
                Dim DIV_QTY As Double = CURR_QTY_AVAIL / MOQ
                Dim DIV_QTY_INT As Int64 = Math.Floor(DIV_QTY)
                If DIV_QTY <> DIV_QTY_INT Then
                    CURR_QTY_AVAIL = MOQ * DIV_QTY_INT
                End If
            End If

            str.Append(CURR_QTY_AVAIL & ",")

            'This used to use real future if exists and if not use Alternate future.
            'This was changed on 2/19 to use alternate if it exists and if not use
            'Any real inventory Per Danny and Rich. - WR 
            'If IsDate(FUT_DATE) And Val(FUT_QTY_AVAIL) > 0 Then
            '    FAVL = FUT_DATE & "|" & FUT_QTY_AVAIL
            'Else
            '    If IsDate(rowICTINVTR.Item("ALT_FUT_DATE").ToString & String.Empty) And Val(rowICTINVTR.Item("ALT_FUT_QTY").ToString & String.Empty) > 0 Then
            '        FAVL = CDate(rowICTINVTR.Item("ALT_FUT_DATE").ToString & String.Empty).ToShortDateString & "|" & Val(rowICTINVTR.Item("ALT_FUT_QTY").ToString & String.Empty)
            '    End If
            'End If
            Dim FAVL As String = ""
            If IsDate(rowICTINVTR.Item("ALT_FUT_DATE").ToString & String.Empty) And Val(rowICTINVTR.Item("ALT_FUT_QTY").ToString & String.Empty) > 0 Then
                FAVL = CDate(rowICTINVTR.Item("ALT_FUT_DATE").ToString & String.Empty).ToShortDateString & "|" & Val(rowICTINVTR.Item("ALT_FUT_QTY").ToString & String.Empty)
            Else
                If IsDate(FUT_DATE) And Val(FUT_QTY_AVAIL) > 0 Then
                    FAVL = FUT_DATE & "|" & FUT_QTY_AVAIL
                End If
            End If
            str.Append(Chr(34) & FAVL & Chr(34) & ",")
            str.Replace(",", vbNewLine, str.Length - 1, 1)
        Next
        Dim localFile As String = ASCMAIN1.Folders("Temp")
        If Not localFile.EndsWith("\") Then
            localFile = localFile & "\"
        End If
        localFile = localFile & FileName
        If IO.File.Exists(localFile) Then
            IO.File.Delete(localFile)
        End If
        My.Computer.FileSystem.WriteAllText(localFile, str.ToString, False)

        If (ASCMAIN1.Running_in_VS) Then
            Stop
        End If

        If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop

        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                If .Connected = True Then
                    .Logoff()
                End If
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = WB_PARM_SITE_USER
                .Password = WB_PARM_SITE_PWD
                .RemoteHost = FTP_REMOTE_HOST
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = FileName
                '.Overwrite = False
                .Overwrite = True
                'If Not .FileExists() Then
                .Upload()
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
                'End If
            Catch ex As Exception
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
            End Try
        End With

        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub chkAutoRefresh_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAutoRefresh.CheckedChanged
        If chkAutoRefresh.Checked Then
            'If Not (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
            '    MsgBox("Sorry.   Wayne Did Not Enable This Feature Yet", MsgBoxStyle.Critical, "Tell Wayne To Hurry Up!")
            '    chkAutoRefresh.Checked = False
            '    Exit Sub
            'Else
            '    Stop
            'End If
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

    Private Sub chkMATERIAL_OVERRIDE_CheckedChanged(sender As Object, e As EventArgs)
        UpdateOverrides()
    End Sub

    Private Sub chkMETA_OVERRIDE_CheckedChanged(sender As Object, e As EventArgs)
        UpdateOverrides()
    End Sub

    Private Sub chkSEARCH_KEYWORDS_CheckedChanged(sender As Object, e As EventArgs)
        UpdateOverrides()
    End Sub

    'Private Sub StyleImmediate(ByVal AR As String)
    '    MsgBox("You Must Select A Row To Remove", vbOKOnly, "Removal")
    '    'A = Add Immediately
    '    'R = Remove Immediately
    '    If AR <> "A" And AR <> "R" Then
    '        MsgBox("Error in StyleImmediate", vbOKOnly, "Error")
    '        Exit Sub
    '    Else
    '        If grdWBTSTYLD.Selected.Rows.Count = 1 Then
    '            Dim STYLE_CODE As String = grdWBTSTYLD.Selected.Rows(0).Cells.Item("STYLE_CODE").Text
    '            If CreateProductXml(STYLE_CODE, AR,, False) Then
    '                'FTPTables = True
    '                Call FTPProducts()
    '                'FTPTables = False
    '            End If
    '        Else
    '            MsgBox("You Must Select A Row To Remove", vbOKOnly, "Removal")
    '        End If
    '    End If
    'End Sub

    Private Sub tmrAutoSync_Tick(sender As System.Object, e As System.EventArgs) Handles tmrAutoSync.Tick
        Dim sleepSec As Integer = (Val(txtAutoWait.Text.Split(":")(1))) + (Val(txtAutoWait.Text.Split(":")(0)) * 60)
        Stop 'This should No Longer Be Used
        Exit Sub

        'If InAutoMode Then
        '    If CDate(Now().ToShortTimeString) > CDate(CDate(dtpNextSync.Text).ToShortTimeString) Then
        '        Me.Cursor = Cursors.WaitCursor
        '        Dim LAST_UPDATE As DateTime = Now()
        '        'Auto Shutdown at 9PM now.  Per Wayne.
        '        If LAST_UPDATE.TimeOfDay.Hours > 21 Then
        '            SendErrorEMail("It's After 9PM And The Regency Inventory Process Has Been Working Hard All Day.  Going To Sleep Now.  Don't Forget To Wake Me Back Up In The Morning.", True)
        '            'Update_Record()
        '            Application.DoEvents()
        '            Application.Exit()
        '        End If

        '        dtpNextSync.Text = LAST_UPDATE.AddSeconds(sleepSec)
        '        TickCount = 0

        '        Application.DoEvents()
        '        Update_Record()
        '        Application.DoEvents()
        '        Clear_Record()
        '        Application.DoEvents()
        '        Load_Record()
        '        Application.DoEvents()

        '        Dim UpdatePricing As Boolean = False
        '        If chkUpdatePricing.Checked Then
        '            UpdatePricing = True
        '        End If

        '        If CreateProductXml("", "A", False, True, Val(txtNextGroup.Text), UpdatePricing) Then
        '            WebBrowser1.Visible = True
        '            grdWBTSTYLD.Visible = False
        '            'FTPTables = True
        '            Call FTPProducts()
        '            'FTPTables = False
        '            WebBrowser1.Visible = False
        '            grdWBTSTYLD.Visible = True

        '            For Each STYLE_CODE As String In styleList
        '                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("STYLE_CODE = '" & STYLE_CODE & "'")
        '                    rowWBTSTYLD.Item("LAST_ON_HAND") = rowWBTSTYLD.Item("CURR_ON_HAND")
        '                    rowWBTSTYLD.Item("LAST_UPDATE") = LAST_UPDATE
        '                    rowWBTSTYLD.Item("LAST_UPDATE_REMARKS") = "Auto-Sync"
        '                    'rowWBTSTYLD.Item("WEB_IND") = "W" 'USED WHEN ADDING NEW STYLES
        '                Next
        '            Next
        '        End If

        '        If Val(txtNextGroup.Text) = Val(txtMaxGroup.Text) Then
        '            txtNextGroup.Text = 1
        '        Else
        '            txtNextGroup.Text = Val(txtNextGroup.Text) + 1
        '        End If
        '        UpdateWBTPARM1()
        '        Application.DoEvents()
        '        Update_Record()
        '        Application.DoEvents()
        '        Clear_Record()
        '        Application.DoEvents()
        '        Load_Record()
        '        Application.DoEvents()
        '        Me.Cursor = Cursors.Default
        '    Else
        '        TickCount += 1
        '    End If
        'Else
        '    tmrAutoSync.Stop()
        'End If
    End Sub
    Private Sub UpdateWBTPARM1()
        Dim SQLS As New System.Text.StringBuilder
        SQLS.Length = 0
        SQLS.AppendLine("UPDATE WBTPARM1 SET WB_NEXT_GROUP = " & Val(txtNextGroup.Text))
        ASCMAIN1.sql = SQLS.ToString
        ASCDATA1.ExecuteSQL()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        docComplete = True
    End Sub

#End Region

#Region "grdWBTSTYLD"
    Private Sub grdWBTSTYLD_DoubleClick(sender As Object, e As EventArgs) Handles grdWBTSTYLD.DoubleClick
        Dim STYLE_CODE As String = grdWBTSTYLD.ActiveRow.Cells.Item("STYLE_CODE").Text
        Dim STYLE_DESC As String = grdWBTSTYLD.ActiveRow.Cells.Item("STYLE_DESC").Text
        txtSTYLE_CODE.Text = STYLE_CODE
        txtSTYLE_DESC.Text = STYLE_DESC
        Fill_Records("WBTSTYLH", STYLE_CODE)
        Fill_Records("ICTSTYLD", STYLE_CODE)
        Fill_Records("WBTPAGEX", STYLE_CODE)
        'Fill_Records("ICTSTYLB", STYLE_CODE)
        'txtSTYLE_CODE.Text = dst.Tables.Item("WBTSTYLH").Rows(0).Item("STYLE_CODE").ToString & ""
        'txtSTYLE_DESC.Text = dst.Tables.Item("WBTSTYLH").Rows(0).Item("STYLE_DESC").ToString & ""
        Dim RecCntH As Int16 = dst.Tables.Item("WBTSTYLH").Rows.Count
        If RecCntH = 0 Then
            Dim newWBTSTYLH As DataRow = dst.Tables.Item("WBTSTYLH").NewRow
            newWBTSTYLH.Item("STYLE_CODE") = STYLE_CODE
            newWBTSTYLH.Item("STYLE_DESC_LONG") = Null
            newWBTSTYLH.Item("SHIPPING_DETAILS") = Null
            newWBTSTYLH.Item("SEARCH_OVERRIDE") = "0"
            newWBTSTYLH.Item("SEARCH_KEYWORDS") = STYLE_CODE
            newWBTSTYLH.Item("VIDEO_URL") = Null
            newWBTSTYLH.Item("MATERIAL_OVERRIDE") = "0"
            newWBTSTYLH.Item("MATERIALS") = GetMaterials(STYLE_CODE)
            newWBTSTYLH.Item("META_OVERRIDE") = "0"
            newWBTSTYLH.Item("META_DESC") = STYLE_DESC
            newWBTSTYLH.Item("INIT_OPER") = ASCMAIN1.USER_ID
            newWBTSTYLH.Item("LAST_OPER") = ASCMAIN1.USER_ID
            newWBTSTYLH.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
            newWBTSTYLH.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
            dst.Tables.Item("WBTSTYLH").Rows.Add(newWBTSTYLH)
        End If
        Bind_Controls(splWHTSTYLH, "WBTSTYLH")
        SetOverrides()
        SetTabModes(2)
        Dim SQLP As New System.Text.StringBuilder
        SQLP.Length = 0
        SQLP.AppendLine("select IC_PARM_STYLE_IMG_DIR from ICTPARM1 where IC_PARM_KEY = 'Z'")
        ASCMAIN1.sql = SQLP.ToString()
        Dim IMAGES_FOLDER As String = ASCDATA1.GetDataValue
        Dim IMAGE_NAME As String = grdWBTSTYLD.ActiveRow.Cells.Item("STYLE_CODE").Text & "-" & grdWBTSTYLD.ActiveRow.Cells.Item("COLOR_CODE").Text & ".JPG"
        imgSTYLE.Image = ASCMAIN1.Get_Image(IMAGES_FOLDER, IMAGE_NAME, True, , , ) ' imgba)
    End Sub

#End Region

#Region "Custom Methods - Subs"
    Private Sub AddStyleColors()
        Dim S As New Text.StringBuilder With {.Length = 0}
        S.AppendLine("SELECT")
        S.AppendLine("SL.STYLE_CODE,")
        S.AppendLine("CL.COLOR_CODE,")
        S.AppendLine("SL.STYLE_DESC,")
        S.AppendLine("CL.STYLE_COLOR_STATUS,")
        S.AppendLine("DECODE(NVL(SL.CUST_CODE,'X'),'X','Stock','Non-Stock') AS TYPE,")
        S.AppendLine("((NVL(ST.WHSE_QTY_ON_HAND,0) - NVL(ST.WHSE_QTY_PICK,0)) + NVL(ST.WHSE_QTY_TRAN,0) + NVL(ST.WHSE_QTY_ON_ORDER,0) - NVL(ST.WHSE_QTY_OPEN,0)) AS FTR_AVAIL")
        S.AppendLine("FROM ICTSTYL1 SL, ICTSTYC1 CL, ICTSTAT2 ST")
        S.AppendLine("WHERE SL.STYLE_CODE = CL.STYLE_CODE")
        S.AppendLine("AND CL.STYLE_CODE = ST.STYLE_CODE (+)")
        S.AppendLine("AND CL.COLOR_CODE = ST.COLOR_CODE(+)")
        S.AppendLine("AND ST.WHSE_CODE (+) = 'MS'")
        If chkSTOCKONLY.Checked = True Then
            S.AppendLine("AND NVL(SL.CUST_CODE,'X') = 'X'")
        End If
        If chkNoDNR.Checked = True Then
            S.AppendLine("AND NOT ")
            S.AppendLine(" (")
            S.AppendLine("  SL.STYLE_STATUS = 'D'")
            S.AppendLine("  AND")
            S.AppendLine("  (")
            S.AppendLine("  ((NVL(ST.WHSE_QTY_ON_HAND,0) - NVL(ST.WHSE_QTY_PICK,0)) + NVL(ST.WHSE_QTY_TRAN,0) + NVL(ST.WHSE_QTY_ON_ORDER,0) - NVL(ST.WHSE_QTY_OPEN,0)) = 0")
            S.AppendLine("  )")
            S.AppendLine(" )")
        End If
        'Put these back if you want to go back to limiting Active with future. 
        'S.AppendLine("AND SL.STYLE_STATUS = 'A'")
        'S.AppendLine("AND CL.STYLE_COLOR_STATUS = 'A'")
        'S.AppendLine("AND ((NVL(ST.WHSE_QTY_ON_HAND,0) - NVL(ST.WHSE_QTY_PICK,0)) + NVL(ST.WHSE_QTY_TRAN,0) + NVL(ST.WHSE_QTY_ON_ORDER,0) - NVL(ST.WHSE_QTY_OPEN,0)) > 0")
        If chkUSEUPLOADS.Checked Then
            S.AppendLine("AND (SL.STYLE_CODE,CL.COLOR_CODE)")
            S.AppendLine("IN")
            S.AppendLine("(")
            S.AppendLine("  SELECT STYLE_CODE, COLOR_CODE FROM WBTUPLD1 WHERE (STYLE_CODE, COLOR_CODE) NOT IN (SELECT STYLE_CODE, COLOR_CODE FROM WBTSTYLD)")
            S.AppendLine(")")
        End If
        'S.AppendLine("")
        'S.AppendLine("AND SL.STYLE_CODE")
        'S.AppendLine("NOT IN")
        'S.AppendLine("(")
        'S.AppendLine("  SELECT STYLE_CODE FROM WBTSTYLH GROUP BY  STYLE_CODE")
        'S.AppendLine(")")

        With ASCMAIN1.CodeSelector
            .SQL = S.ToString
            .MultipleSelections = True
            .PreviouslySelectedCodes0 = ""
            .Caption = "Web Style Selection"
            .TABLE_NAME = ""
            .VIEW_NAME = ""
            .VIEW_DESC = ""
            .COLUMN_NAME = ""
            .COLUMN_PREKEYs = New Dictionary(Of String, String)
            .Custom_sql_where = ""
            .tblASTVIEW1 = New DataTable
        End With
        Dim F As New ASFCODE1
        F.ShowDialog()
        If ASCMAIN1.CodeSelector.Selections <> 0 Then
            For Each rowSEL As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", rowSEL.Item("STYLE_CODE"), rowSEL.Item("COLOR_CODE"))
                If dst.Tables.Item("WBTSTYLD").Select(filter).Count = 0 Then
                    Dim newWBTSTYLD As DataRow = dst.Tables.Item("WBTSTYLD").NewRow
                    newWBTSTYLD.Item("STYLE_CODE") = rowSEL.Item("STYLE_CODE") & ""
                    newWBTSTYLD.Item("COLOR_CODE") = rowSEL.Item("COLOR_CODE") & ""
                    newWBTSTYLD.Item("STYLE_STATUS") = rowSEL.Item("STYLE_COLOR_STATUS") & ""
                    'newWBTSTYLD.Item("STYLE_FULL_DESC") = rowSEL.Item("STYLE_DESC") & ""
                    newWBTSTYLD.Item("WEB_IND") = "U"
                    newWBTSTYLD.Item("UPLOAD_BATCH") = 99
                    newWBTSTYLD.Item("DEFAULT_IMAGE") = rowSEL.Item("STYLE_CODE") & "-" & rowSEL.Item("COLOR_CODE") & ".jpg"
                    newWBTSTYLD.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    newWBTSTYLD.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    newWBTSTYLD.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    newWBTSTYLD.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                    newWBTSTYLD.Item("LAST_ON_HAND") = Val(rowSEL.Item("FTR_AVAIL") & "")
                    newWBTSTYLD.Item("CURR_ON_HAND") = Val(rowSEL.Item("FTR_AVAIL") & "")
                    newWBTSTYLD.Item("STYLE_SORT") = rowSEL.Item("STYLE_CODE") & "-" & rowSEL.Item("COLOR_CODE")
                    newWBTSTYLD.Item("STYLE_GROUP") = 0
                    newWBTSTYLD.Item("UPLOAD_IMG") = "1"
                    newWBTSTYLD.Item("LAST_UPDATE") = Now()
                    newWBTSTYLD.Item("LAST_UPDATE_REMARKS") = "Initial Addition of Item"
                    dst.Tables.Item("WBTSTYLD").Rows.Add(newWBTSTYLD)
                End If
                Dim SC As New System.Text.StringBuilder With {.Length = 0}
                SC.AppendLine(String.Format("SELECT COUNT(*) FROM WBTSTYLH WHERE STYLE_CODE = '{0}'", rowSEL.Item("STYLE_CODE")))
                ASCMAIN1.sql = SC.ToString()
                Dim RecCnt As Int16 = Val(ASCDATA1.GetDataValue)
                Dim filterH As String = String.Format("STYLE_CODE = '{0}'", rowSEL.Item("STYLE_CODE"))
                Dim RecCntH As Int16 = dst.Tables.Item("WBTSTYLH").Rows.Count
                If RecCnt = 0 And RecCntH = 0 Then
                    Dim newWBTSTYLH As DataRow = dst.Tables.Item("WBTSTYLH").NewRow
                    newWBTSTYLH.Item("STYLE_CODE") = rowSEL.Item("STYLE_CODE") & ""
                    newWBTSTYLH.Item("STYLE_DESC_LONG") = Null
                    newWBTSTYLH.Item("SHIPPING_DETAILS") = Null
                    newWBTSTYLH.Item("SEARCH_OVERRIDE") = "0"
                    newWBTSTYLH.Item("SEARCH_KEYWORDS") = GetKeyWords(rowSEL.Item("STYLE_CODE") & "")
                    newWBTSTYLH.Item("VIDEO_URL") = Null
                    newWBTSTYLH.Item("MATERIAL_OVERRIDE") = "0"
                    newWBTSTYLH.Item("MATERIALS") = GetMaterials(rowSEL.Item("STYLE_CODE") & "")
                    newWBTSTYLH.Item("META_OVERRIDE") = "0"
                    newWBTSTYLH.Item("META_DESC") = rowSEL.Item("STYLE_DESC") & ""
                    newWBTSTYLH.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    newWBTSTYLH.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    newWBTSTYLH.Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                    newWBTSTYLH.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                    dst.Tables.Item("WBTSTYLH").Rows.Add(newWBTSTYLH)
                End If
            Next
            Update_Record(False)
            Application.DoEvents()
            MsgBox("Style(s) Added.  Please Wait While Data Is Refreshed", vbOKOnly, "Added")
            Clear_Record()
            Application.DoEvents()
            Load_Record()
            Application.DoEvents()
        End If
    End Sub

    Private Sub AutoSync(StartSync As Boolean)
        If StartSync Then
            With UltraExplorerBar1.Groups("Screen Control")
                .Items("Load Records").Settings.Enabled = DefaultableBoolean.False
                .Items("Update").Settings.Enabled = DefaultableBoolean.False
                .Items("Finish").Visible = False
                .Items("Done").Settings.Enabled = DefaultableBoolean.False
                '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.False
            End With
            chkShowOnlyDiff.Checked = False
            grpUploads.Enabled = False
            btnAddStyles.Enabled = False
            InAutoMode = True
            tmrAutoSync.Start()
        Else
            InAutoMode = False
            grpUploads.Enabled = True
            Mode_Settings(False)
            SetTabModes(0)
            btnAddStyles.Enabled = True
        End If
    End Sub

    'Private Sub BuildFTPFile()
    '    Dim NewLine As String = ""
    '    Dim FileForStream As String = "Supplier_ID,Item_Number,On_Hand,Back_Order,On_Order,NxtAvailDate,STATUS,Description" & vbCrLf
    '    For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select()
    '        'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "whr") Then
    '        '    If rowICTSTATX.Item("Item_Number").ToString() = "MTX37196" Then Stop
    '        'End If
    '        NewLine = ""
    '        NewLine = String.Format("{0},", rowICTSTATX.Item("Supplier_ID"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("Item_Number"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("On_Hand"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("Back_Order"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("On_Order"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("NxtAvailDate"))
    '        NewLine += String.Format("{0},", rowICTSTATX.Item("STATUS"))
    '        NewLine += CleanText(rowICTSTATX.Item("Description").ToString & "") & vbCrLf
    '        FileForStream += NewLine.ToString
    '    Next
    '    Dim fs As FileStream = File.Create(TransferFile)

    '    Dim info As Byte() = New UTF8Encoding(True).GetBytes(FileForStream)
    '    fs.Write(info, 0, info.Length)
    '    fs.Close()
    'End Sub

    Private Function CreateProductXml(Optional ByVal STYLE_COLOR As String = "",
                                      Optional ByVal AR As String = "",
                                      Optional ByVal TestingOnly As Boolean = False,
                                      Optional UploadInventoryOnly As Boolean = True,
                                      Optional STYLE_GROUP As Int64 = 99,
                                      Optional UpdatePricing As Boolean = False) As Boolean
        Stop 'THIS SHOULD NO LONGER BE USED
        'Dim Retval As Boolean = True
        'Dim batchFilter As String = ""
        'Dim singleStyle As Boolean = False
        'styleList.Clear()
        'styleListInactive.Clear()

        'If STYLE_COLOR.Length > 0 Then
        '    If AR.Length > 0 Then
        '        batchFilter = String.Format("STYLE_CODE = '{0}'", STYLE_COLOR)
        '        singleStyle = True
        '    Else
        '        MsgBox("Add / Remove Not Specified", vbOKOnly, "Error")
        '        Retval = False
        '        Return Retval
        '        Exit Function
        '    End If
        'Else
        '    If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        '        Stop
        '        'batchFilter = String.Format("WEB_IND = '{0}'", "U")
        '        'batchFilter = String.Format("WEB_IND = '{0}'", "W")
        '        'batchFilter = String.Format("STYLE_CODE = '{0}'", "MTX47222")
        '        'batchFilter = String.Format("WEB_IND = '{0}' AND CURR_ON_HAND = 0", "W")
        '        'batchFilter = "WEB_IND = 'R' or (WEB_IND = 'W' AND CURR_ON_HAND <> LAST_ON_HAND)"
        '        'batchFilter = "WEB_IND = 'U'"
        '        'batchFilter = String.Format("(LAST_ON_HAND <> CURR_ON_HAND) OR (STYLE_GROUP = {0})", Val(txtNextGroup.Text))
        '    End If
        '    If chkUseFilter.Checked Then
        '        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select()
        '            rowWBTSTYLD.Item("FILTER_SEL") = "0"
        '        Next
        '        For Each grow As UltraWinGrid.UltraGridRow In grdWBTSTYLD.Rows
        '            If Not grow.IsFilteredOut Then
        '                Dim STYLE_CODE As String = grow.Cells.Item("STYLE_CODE").Text & String.Empty
        '                Dim SFilter As String = String.Format("STYLE_CODE = '{0}'", STYLE_CODE)
        '                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(SFilter)
        '                    rowWBTSTYLD.Item("FILTER_SEL") = "1"
        '                Next
        '            End If
        '        Next
        '        batchFilter = "FILTER_SEL = '1'"
        '    Else
        '        If STYLE_GROUP = 99 Then
        '            'batchFilter = String.Format("WEB_IND = '{0}' AND (LAST_ON_HAND <> CURR_ON_HAND)", "W")
        '            batchFilter = String.Format("WEB_IND = '{0}'", "W")
        '        Else
        '            'batchFilter = String.Format("WEB_IND = '{0}' AND STYLE_GROUP = {1} AND (LAST_ON_HAND <> CURR_ON_HAND)", "W", STYLE_GROUP)
        '            If chkFullUpload.Checked Then
        '                UploadInventoryOnly = False
        '                batchFilter = "FULL_UPLOAD = '1'"
        '            Else
        '                If UploadInventoryOnly Then
        '                    'batchFilter = String.Format("WEB_IND = '{0}' AND STYLE_GROUP = {1} AND (LAST_ON_HAND <> CURR_ON_HAND)", "W", STYLE_GROUP)
        '                    batchFilter = String.Format("WEB_IND = '{0}' AND STYLE_GROUP = {1}", "W", STYLE_GROUP)
        '                Else
        '                    'batchFilter = String.Format("WEB_IND = '{0}' AND STYLE_GROUP = {1}", "W", STYLE_GROUP)
        '                    batchFilter = String.Format("STYLE_GROUP = {0}", STYLE_GROUP)
        '                End If
        '            End If

        '        End If
        '    End If
        'End If

        'Try
        '    Me.Cursor = Cursors.WaitCursor

        '    ASCMAIN1.Progress("Create XML Document", "")

        '    'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        '    '    'Stop
        '    '    'batchFilter = String.Format("STYLE_CODE = '{0}'", "MTX62444")
        '    '    ''batchFilter = "WEB_IND = 'W' AND CURR_ON_HAND > 0"
        '    '    'batchFilter = "WEB_IND = 'W'"
        '    'End If
        '    Dim productXML As New WBCITEM2(dst.Tables("WBTSTYLD"))
        '    'Dim productXML As New WBCITEMW

        '    shopSiteFilename = WB_PARM_PRODUCTS_DIR & "STYLE_CODE_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
        '    Dim StyleCount As Int64 = 0
        '    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(batchFilter)

        '        Dim ProcessStyleColor As Boolean = False
        '        Select Case rowWBTSTYLD.Item("WEB_IND")
        '            Case "X" 'Not On Web
        '                If singleStyle Then
        '                    MsgBox("Style Selected Not On Web", vbOKOnly, "Error")
        '                    ProcessStyleColor = False
        '                Else
        '                    ProcessStyleColor = False
        '                End If
        '            Case "W" 'On Web
        '                If singleStyle Then
        '                    If AR = "A" Then
        '                        If Not styleList.Contains(rowWBTSTYLD.Item("STYLE_CODE")) Then
        '                            styleList.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                        End If
        '                        ProcessStyleColor = True
        '                    Else
        '                        styleListInactive.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                        ProcessStyleColor = True
        '                    End If
        '                Else
        '                    If Not styleList.Contains(rowWBTSTYLD.Item("STYLE_CODE")) Then
        '                        styleList.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                    End If
        '                    ProcessStyleColor = True
        '                End If
        '            Case "U" 'Waiting For Upload
        '                If singleStyle Then
        '                    If AR = "A" Then
        '                        styleList.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                        ProcessStyleColor = True
        '                    Else
        '                        MsgBox("Style Selected Not On Web", vbOKOnly, "Error")
        '                        ProcessStyleColor = False
        '                    End If
        '                Else
        '                    styleList.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                    ProcessStyleColor = True
        '                End If
        '            Case "R" 'Waiting For Removal
        '                If singleStyle Then
        '                    If AR = "A" Then
        '                        MsgBox("Style Selected Already on Web Waiting For Removal", vbOKOnly, "Error")
        '                        ProcessStyleColor = False
        '                    Else
        '                        styleListInactive.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                        ProcessStyleColor = True
        '                    End If
        '                Else
        '                    styleListInactive.Add(rowWBTSTYLD.Item("STYLE_CODE"))
        '                    ProcessStyleColor = True
        '                End If
        '        End Select
        '        If ProcessStyleColor Then
        '            ASCMAIN1.Progress("-", rowWBTSTYLD.Item("STYLE_CODE"))
        '            Application.DoEvents()
        '            StyleCount += 1
        '            Dim rowCnt As Int64 = productXML.AddStyle(rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE"), styleListInactive, UploadInventoryOnly, , UpdatePricing)
        '            'MsgBox("STYLES: " & rowCnt, vbOKOnly, "Added")
        '            'productXML.AddStyle(rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE"), styleListInactive, UploadInventoryOnly)
        '            'productXML.AddStyle(rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE"), styleListInactive, False)
        '        End If
        '    Next
        '    'MsgBox("Style Count: " & styleList.Count)
        '    For Each style As String In styleList
        '        ASCMAIN1.Progress("Parents", style)
        '        'Dim DefaultColorFilter As String = String.Format("STYLE_CODE = '{0}' AND WEB_IND = 'W' AND CURR_ON_HAND <> 0", style)
        '        'GO_LIVE_CHANGES
        '        Dim DefaultColorFilter As String = String.Format("STYLE_CODE = '{0}' AND CURR_ON_HAND > 0", style)
        '        'Dim DefaultColor As String = dst.Tables("WBTSTYLD").Select(DefaultColorFilter).FirstOrDefault.Item("COLOR_CODE").ToString & String.Empty
        '        Dim DefaultColor As String = ""
        '        Dim rowColor As DataRow = dst.Tables("WBTSTYLD").Select(DefaultColorFilter).FirstOrDefault
        '        If IsNothing(rowColor) Then
        '            'Stop
        '            Dim DefaultColorFilter2 As String = String.Format("STYLE_CODE = '{0}'", style)
        '            Dim rowColor2 As DataRow = dst.Tables("WBTSTYLD").Select(DefaultColorFilter2).FirstOrDefault
        '            If Not IsNothing(rowColor2) Then
        '                DefaultColor = rowColor2.Item("COLOR_CODE").ToString & String.Empty
        '                If DefaultColor = "" Then
        '                    Stop
        '                End If
        '            End If
        '        Else
        '            DefaultColor = rowColor.Item("COLOR_CODE").ToString & String.Empty
        '            If DefaultColor = "" Then
        '                Stop
        '            End If
        '        End If
        '        StyleCount += 1
        '        productXML.AddStyle(style, DefaultColor, styleListInactive, UploadInventoryOnly, True, UpdatePricing)
        '        'productXML.AddStyle(style, DefaultColor, styleListInactive, False, True)
        '    Next

        '    If styleListInactive.Count + styleList.Count = 0 Then
        '        Retval = False
        '    Else
        '        Retval = True
        '    End If

        '    If Retval = True Then
        '        Dim xmlLabelRequest As XmlDocument = productXML.GetXMLDocument
        '        ASCMAIN1.Progress("Saving XML Document", "")
        '        xmlLabelRequest.Save(shopSiteFilename)
        '    End If

        '    'txtOutputFile.Text = shopSiteFilename
        '    txtOutputFile.Text = WB_PARM_PRODUCTS_DIR
        '    Dim xfileInfo As New FileInfo(shopSiteFilename)
        '    If xfileInfo.Length <= 500000 Then
        '        ASCMAIN1.Progress("Loading XML Document in Viewer", "")
        '        WebBrowser1.Navigate(New Uri(shopSiteFilename))
        '    Else
        '        WebBrowser1.Navigate("about:blank")
        '        Dim HTML As String
        '        HTML = "<HTML>" &
        '                "<TITLE>XML Style Upload</TITLE>" &
        '                "<BODY>" &
        '                "<FONT COLOR = RED>" &
        '                "The XML Document is too " &
        '                "<FONT SIZE = 5>" &
        '                "<B>" &
        '                "Large " &
        '                "</B>" &
        '                "</FONT SIZE>" &
        '                "to display!" &
        '                "</FONT>" &
        '                "</BODY>" &
        '                "</HTML>"

        '        WebBrowser1.Document.Write(HTML)
        '    End If

        'Catch ex As Exception
        '    shopSiteFilename = String.Empty
        '    If InAutoMode Then
        '        SendErrorEMail("Error creating XML Document: " & ex.Message)
        '    Else
        '        MessageBox.Show("Error creating XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
        '    End If
        'Finally
        '    Me.Cursor = Cursors.Default
        '    ASCMAIN1.Progress(String.Empty, String.Empty)
        'End Try

        'Return Retval
    End Function

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
                'Stop
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

    'Private Sub ftp_File()
    '    Ftp1.User = "EDI_RegencyInternational"
    '    Ftp1.Password = "N0wayfa1r!"
    '    Ftp1.RemoteHost = "edi.wayfair.com"
    '    Ftp1.RemotePath = "inventory"
    '    Ftp1.Logon()
    '    Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
    '    Ftp1.LocalFile = TransferFile
    '    Ftp1.RemoteFile = "regency.csv"
    '    Ftp1.Overwrite = True
    '    Ftp1.Upload()
    '    Ftp1.Logoff()
    'End Sub

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

    Private Sub MoveToParents()
        WebBrowser1.Parent = grdWBTSTYLD.Parent
        splWHTSTYLH.Parent = grdWBTSTYLD.Parent
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
                'Stop
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

    'Private Sub SendWayfair()
    '    Try
    '        ASCMAIN1.Progress("Uploading Inventory", "")
    '        If My.Computer.FileSystem.FileExists(TransferFile) Then
    '            My.Computer.FileSystem.DeleteFile(TransferFile)
    '        End If
    '        BuildFTPFile()
    '        ftp_File()
    '        Dim mBody As String = String.Format("Wayfair Inventory Feed Updated with {0} products at {1}",
    '                                            dst.Tables.Item("ICTSTATX").Rows.Count,
    '                                            Format(Now(), "hh:mm tt"))
    '        SendErrorEMail(mBody, False)
    '    Catch ex As Exception
    '        SendErrorEMail("Error During Wayfair Inventory Upload: " & ex.Message, False)
    '    End Try
    '    ASCMAIN1.Progress("", "")
    'End Sub

    Private Sub SetTabModes(ByVal TabMode As Integer)
        Select Case TabMode
            Case 0 'Data Not Loaded
                txtSTYLE_CODE.Visible = False
                txtSTYLE_DESC.Visible = False
                grdWBTSTYLD.Visible = False
                splWHTSTYLH.Visible = False
                WebBrowser1.Visible = False
                UltraTabControl1.Visible = False
                UltraTabControl1.Tabs(0).Text = ""
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Load Records").Settings.Enabled = DefaultableBoolean.True
                    .Items("Update").Settings.Enabled = DefaultableBoolean.False
                    .Items("Finish").Visible = False
                    .Items("Done").Settings.Enabled = DefaultableBoolean.False
                    '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.False
                End With
                UltraExplorerBar1.Groups("Inventory").Visible = False
                UltraExplorerBar1.Groups("Alt Inventory").Visible = False
                'UltraExplorerBar1.Groups("Auto Refresh").Visible = False
                UltraExplorerBar1.Groups("Shopsite Upload").Visible = False
            Case 1 'Styles To Upload
                txtSTYLE_CODE.Visible = False
                txtSTYLE_DESC.Visible = False
                UltraTabControl1.Tabs(0).Text = "Styles To Upload"
                grdWBTSTYLD.Visible = True
                splWHTSTYLH.Visible = False
                WebBrowser1.Visible = False
                UltraTabControl1.Visible = True
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Load Records").Settings.Enabled = DefaultableBoolean.False
                    .Items("Update").Settings.Enabled = DefaultableBoolean.True
                    .Items("Finish").Visible = False
                    .Items("Done").Settings.Enabled = DefaultableBoolean.True
                    '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.True
                End With
                UltraExplorerBar1.Groups("Inventory").Visible = True
                UltraExplorerBar1.Groups("Alt Inventory").Visible = True
                UltraExplorerBar1.Groups("Shopsite Upload").Visible = True
                'UltraExplorerBar1.Groups("Auto Refresh").Visible = True
            Case 2 'Style Details
                txtSTYLE_CODE.Visible = True
                txtSTYLE_DESC.Visible = True
                UltraTabControl1.Tabs(0).Text = "Style Details"
                grdWBTSTYLD.Visible = False
                splWHTSTYLH.Visible = True
                WebBrowser1.Visible = False
                UltraTabControl1.Visible = True
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Load Records").Settings.Enabled = DefaultableBoolean.False
                    .Items("Update").Visible = False
                    .Items("Finish").Settings.Enabled = DefaultableBoolean.True
                    .Items("Finish").Visible = True
                    .Items("Done").Settings.Enabled = DefaultableBoolean.False
                    '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.False
                End With
                UltraExplorerBar1.Groups("Inventory").Visible = False
                UltraExplorerBar1.Groups("Alt Inventory").Visible = False
                'UltraExplorerBar1.Groups("Auto Refresh").Visible = False
                UltraExplorerBar1.Groups("Shopsite Upload").Visible = False
            Case 3 'Browser
                txtSTYLE_CODE.Visible = False
                txtSTYLE_DESC.Visible = False
                UltraTabControl1.Tabs(0).Text = "Shopsite Browser"
                grdWBTSTYLD.Visible = False
                splWHTSTYLH.Visible = False
                WebBrowser1.Visible = True
                UltraTabControl1.Visible = True
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Load Records").Settings.Enabled = DefaultableBoolean.False
                    .Items("Update").Visible = False
                    .Items("Finish").Settings.Enabled = DefaultableBoolean.True
                    .Items("Finish").Visible = False
                    .Items("Done").Settings.Enabled = DefaultableBoolean.False
                    '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.False
                End With
                UltraExplorerBar1.Groups("Inventory").Visible = False
                UltraExplorerBar1.Groups("Alt Inventory").Visible = False
                'UltraExplorerBar1.Groups("Auto Refresh").Visible = False
                UltraExplorerBar1.Groups("Shopsite Upload").Visible = False
            Case Else
                txtSTYLE_CODE.Visible = False
                txtSTYLE_DESC.Visible = False
                UltraTabControl1.Tabs(0).Text = ""
                UltraTabControl1.Visible = False
                grdWBTSTYLD.Visible = True
                splWHTSTYLH.Visible = False
                WebBrowser1.Visible = False
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("Load Records").Settings.Enabled = DefaultableBoolean.True
                    .Items("Update").Settings.Enabled = DefaultableBoolean.False
                    .Items("Finish").Visible = False
                    .Items("Done").Settings.Enabled = DefaultableBoolean.False
                    '.Items("Remove Alt Supplier").Settings.Enabled = DefaultableBoolean.False
                End With
                UltraExplorerBar1.Groups("Inventory").Visible = False
                UltraExplorerBar1.Groups("Alt Inventory").Visible = False
                'UltraExplorerBar1.Groups("Auto Refresh").Visible = False
                UltraExplorerBar1.Groups("Shopsite Upload").Visible = False
        End Select
    End Sub

    Private Sub SetOverrides()
        If dst.Tables.Item("WBTSTYLH").Rows.Count = 1 Then
            Dim rowWBTSTYLH As DataRow = dst.Tables.Item("WBTSTYLH").Rows(0)
            If rowWBTSTYLH.Item("MATERIAL_OVERRIDE") & "" = "1" Then
                chkMATERIAL_OVERRIDE.Checked = True
                txtMATERIALS.Enabled = True
            Else
                chkMATERIAL_OVERRIDE.Checked = False
                txtMATERIALS.Enabled = False
            End If
            If rowWBTSTYLH.Item("SEARCH_OVERRIDE") & "" = "1" Then
                chkSEARCH_KEYWORDS.Checked = True
                txtSEARCH_KEYWORDS.Enabled = True
            Else
                chkSEARCH_KEYWORDS.Checked = False
                txtSEARCH_KEYWORDS.Enabled = False
            End If
            If rowWBTSTYLH.Item("META_OVERRIDE") & "" = "1" Then
                chkMETA_OVERRIDE.Checked = True
                txtMETA_DESC.Enabled = True
            Else
                chkMETA_OVERRIDE.Checked = False
                txtMETA_DESC.Enabled = False
            End If
        End If

    End Sub

    Private Sub Update_Masterfile()
        UpdateOverrides()
        'Update_Record_TDA("ICTSTYLB")
        Update_Record_TDA("WBTSTYLH")
        Dim filter As String = String.Format("STYLE_CODE = '{0}'", txtSTYLE_CODE.Text)

        Dim STYLE_DESC_LONG As String = "0"
        If Absx1.txtFor("STYLE_DESC_LONG").Text <> "" Then
            STYLE_DESC_LONG = "1"
        End If

        Dim STYLE_DESC_SHORT As String = "0"
        If Absx1.txtFor("STYLE_DESC_SHORT").Text <> "" Then
            STYLE_DESC_SHORT = "1"
        End If

        For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(filter)
            rowWBTSTYLD.Item("HAS_DESC_LONG") = STYLE_DESC_LONG
            rowWBTSTYLD.Item("HAS_DESC_SHORT") = STYLE_DESC_SHORT
        Next

        dst.Tables("WBTSTYLH").Rows.Clear()
        dst.Tables("ICTSTYLD").Rows.Clear()
        'dst.Tables("ICTSTYLB").Rows.Clear()
    End Sub

    Private Sub UpdateOverrides()
        If dst.Tables.Item("WBTSTYLH").Rows.Count = 1 Then
            Dim rowWBTSTYLH As DataRow = dst.Tables.Item("WBTSTYLH").Rows(0)
            If chkMATERIAL_OVERRIDE.Checked = True Then
                txtMATERIALS.Enabled = True
                rowWBTSTYLH.Item("MATERIAL_OVERRIDE") = "1"
            Else
                txtMATERIALS.Enabled = False
                rowWBTSTYLH.Item("MATERIAL_OVERRIDE") = "0"
            End If

            If chkSEARCH_KEYWORDS.Checked = True Then
                txtSEARCH_KEYWORDS.Enabled = True
                rowWBTSTYLH.Item("SEARCH_OVERRIDE") = "1"
            Else
                txtSEARCH_KEYWORDS.Enabled = False
                rowWBTSTYLH.Item("SEARCH_OVERRIDE") = "0"
            End If

            If chkMETA_OVERRIDE.Checked = True Then
                txtMETA_DESC.Enabled = True
                rowWBTSTYLH.Item("META_OVERRIDE") = "1"
            Else
                txtMETA_DESC.Enabled = False
                rowWBTSTYLH.Item("META_OVERRIDE") = "0"
            End If
        End If

    End Sub

#End Region

#Region "Custom Methods - Functions"
    Private Function CheckForNullGroups() As Boolean
        Dim Retval As Boolean = True
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("Select COUNT(*) As REC_CNT")
        SQLS.AppendLine("FROM WBTSTYLD")
        SQLS.AppendLine("WHERE NVL(STYLE_GROUP,0) = 0")
        SQLS.AppendLine("AND WEB_IND <> 'X'")
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

    Private Function CheckIsAutoMode() As String
        Dim RetVal As String = ""
        If InAutoMode Then
            RetVal = "You May Not Run This While Auto-Refresh Is On."
        End If
        Return RetVal
    End Function

    Private Function CleanText(ByVal TextIn As String) As String
        TextIn = Replace(TextIn, ",", "")
        TextIn = Replace(TextIn, "'", "")
        TextIn = Replace(TextIn, Chr(34).ToString, "")
        Return Trim(TextIn)
    End Function

    'Private Function CreateInventoryXml(ByVal UploadStart As Integer, ByVal UpLoadFinish As Integer, Optional ByVal TestFilter As String = "", Optional ByVal FullUpload As Boolean = True) As Integer
    '    Dim RetVal As Integer = 0
    '    Try
    '        Me.Cursor = Cursors.WaitCursor

    '        ASCMAIN1.Progress("Create XML Document", "")

    '        Stop 'This should not be done.
    '        Dim inventoryXML As New WBCITEM1
    '        Dim whseQtyAvail As Int16 = 0

    '        shopSiteFilename = WB_PARM_INVENTORY_DIR & "INVENTORY_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".xml"
    '        Dim Selector As String = ""

    '        If InAutoMode Then
    '            FullUpload = False
    '            TestFilter = ""
    '        End If
    '        Dim RecCnt As Integer = 0
    '        For Each row As DataRow In dst.Tables("WBTSTYLD").Select(TestFilter)
    '            Dim ProcessRecord As Boolean = False
    '            If FullUpload Then
    '                ProcessRecord = True
    '            Else
    '                'If (Val(row.Item("LAST_ON_HAND") & "") <> Val(row.Item("CURR_ON_HAND") & "")) Or (row.Item("STYLE_GROUP") & "" = txtNextGroup.Text) Then
    '                If TestFilter.Length > 0 Then
    '                    ProcessRecord = True
    '                Else
    '                    If (row.Item("STYLE_GROUP") & "" = txtNextGroup.Text) Then
    '                        ProcessRecord = True
    '                    Else
    '                        ProcessRecord = False
    '                    End If
    '                End If
    '            End If
    '            If ProcessRecord Then
    '                If RecCnt >= UploadStart And RetVal <= UpLoadFinish Then
    '                    row.Item("LAST_ON_HAND") = row.Item("CURR_ON_HAND")
    '                    RetVal += 1

    '                    Dim onWeb As Boolean = row.Item("WEB_IND") & "" = "W"
    '                    Dim hasBatch As Boolean = row.Item("UPLOAD_BATCH") & "" <> ""
    '                    Dim itemActive As Boolean = row.Item("STYLE_STATUS") & "" = "A"
    '                    Dim hasInventory As Boolean = Val(row.Item("CURR_ON_HAND")) > 0
    '                    Dim styleStatus As String = ""
    '                    Stop
    '                    'WEB_IND has been revided with new meaning as of 4/10/17
    '                    'This has to all be re-thought.

    '                    'onWeb & hasBatch = U -> Update the item.
    '                    'onWeb & !hasBatch = S => Add to Web.  This Does Not Get Done Here.  Need to run Inventory load.
    '                    '!onWeb & hasBatch = X => Remove from Web.
    '                    '!onWeb & !hasBatch = S => Skip this style.  It's not on the web.
    '                    If onWeb And hasBatch Then
    '                        styleStatus = "U"
    '                    End If
    '                    If onWeb And Not hasBatch Then
    '                        styleStatus = "S"
    '                    End If
    '                    If Not onWeb And hasBatch Then
    '                        styleStatus = "X"
    '                    End If
    '                    If Not onWeb And Not hasBatch Then
    '                        styleStatus = "S"
    '                    End If
    '                    'If we made it all the way through here and the style has no inventory, remove it from the web.
    '                    If styleStatus = "U" And Not hasInventory Then
    '                        styleStatus = "X"
    '                    End If

    '                    'If row.Item("STYLE_CODE").ToString = "MTF19619" Then Stop

    '                    If styleStatus = "U" Then
    '                        ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
    '                        styleList.Add(row.Item("STYLE_CODE"))
    '                        inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
    '                    End If

    '                    If styleStatus = "X" Then
    '                        ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
    '                        styleListInactive.Add(row.Item("STYLE_CODE"))
    '                        inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
    '                        row.Item("WEB_IND") = "X"
    '                        row.Item("UPLOAD_BATCH") = Null
    '                    End If

    '                    ''Automatically Remove Items From The Web That Are Discontinued and Have No Inventory.
    '                    'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" And row.Item("STYLE_STATUS") & "" = "D" And Val(row.Item("CURR_ON_HAND")) = 0 Then
    '                    '    row.Item("WEB_IND") = "0"
    '                    '    row.Item("UPLOAD_BATCH") = Null
    '                    'End If

    '                    ''Automatically Remove Items From The Web That Are Active and Have No Inventory.
    '                    'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" And row.Item("STYLE_STATUS") & "" = "A" And Val(row.Item("CURR_ON_HAND")) = 0 Then
    '                    '    row.Item("WEB_IND") = "0"
    '                    '    row.Item("UPLOAD_BATCH") = Null
    '                    'End If

    '                    'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" <> "" Then
    '                    '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
    '                    '    styleList.Add(row.Item("STYLE_CODE"))
    '                    '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
    '                    '    MakeInventoryTable(row.Item("STYLE_CODE"))
    '                    'End If
    '                    'If row.Item("WEB_IND") & "" = "1" And row.Item("UPLOAD_BATCH") & "" = "" Then
    '                    '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
    '                    '    styleList.Add(row.Item("STYLE_CODE"))
    '                    '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
    '                    '    MakeInventoryTable(row.Item("STYLE_CODE"))
    '                    'End If
    '                    'If row.Item("WEB_IND") & "" <> "1" And row.Item("UPLOAD_BATCH") & "" <> "" Then
    '                    '    ASCMAIN1.Progress("-", row.Item("STYLE_CODE"))
    '                    '    styleListInactive.Add(row.Item("STYLE_CODE"))
    '                    '    inventoryXML.AddInventory(row.Item("STYLE_CODE").ToString, styleListInactive)
    '                    'End If
    '                End If
    '            End If
    '            RecCnt += 1
    '        Next

    '        Dim xmlLabelRequest As XmlDocument = inventoryXML.GetXMLDocument
    '        ASCMAIN1.Progress("Saving XML Document", "")
    '        xmlLabelRequest.Save(shopSiteFilename)

    '        Dim xfileInfo As New FileInfo(shopSiteFilename)
    '        If xfileInfo.Length <= 500000 Then
    '            ASCMAIN1.Progress("Loading XML Document in Viewer", "")
    '            WebBrowser1.Navigate(New Uri(shopSiteFilename))
    '        Else
    '            WebBrowser1.Navigate("about:blank")
    '            Dim HTML As String
    '            HTML = "<HTML>" &
    '                "<TITLE>XML Inventory Upload</TITLE>" &
    '                "<BODY>" &
    '                "<FONT COLOR = RED>" &
    '                "The XML Document is too " &
    '                "<FONT SIZE = 5>" &
    '                "<B>" &
    '                "Large " &
    '                "</B>" &
    '                "</FONT SIZE>" &
    '                "to display!" &
    '                "</FONT>" &
    '                "</BODY>" &
    '                "</HTML>"

    '            WebBrowser1.Document.Write(HTML)
    '        End If
    '        inventoryXML = Nothing
    '    Catch ex As Exception
    '        shopSiteFilename = String.Empty
    '        If InAutoMode Then
    '            SendErrorEMail("Error creating XML Document: " & ex.Message)
    '        Else
    '            MessageBox.Show("Error creating XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
    '        End If
    '    Finally
    '        Me.Cursor = Cursors.Default
    '        ASCMAIN1.Progress("", "")
    '    End Try
    '    Return RetVal
    'End Function

    Private Function FTPProducts() As Boolean

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
        'Stop
        Dim IsLive As Boolean = True
        If IsLive Then
            'LIVE SETTINGS
            ASCMAIN1.Progress("-", "FTP To LIVE Site")
            Sftp1.RemoteHost = "69.39.227.201"
            Sftp1.User = WB_PARM_SITE_USER
            Sftp1.Password = WB_PARM_SITE_PWD
        Else
            'DEMO SITE SETTINGS
            ASCMAIN1.Progress("-", "FTP To Demo Site")
            Sftp1.RemoteHost = "216.38.11.230"
            Sftp1.User = "regdemo"
            Sftp1.Password = "re67hg34"
        End If

        Sftp1.RemoteFile = String.Empty
        Sftp1.Timeout = 300

        ASCMAIN1.Progress("-", "Logon")
        'Sftp1.SSHAuthMode = nsoftware.IPWorks.ftpAuthModes.amPassword

        If IsLive Then
            ASCMAIN1.Progress("Uploading Product File", "")
            Try
                Sftp1.Logoff()
                Sftp1.Logon()
            Catch ex As Exception
                Sftp1.Logoff()
                Sftp1.Logon()
            End Try
            If IsNumeric(txtFTPTIME.Text) Then
                Sftp1.Timeout = (Val(txtFTPTIME.Text) * 60)
            Else
                Sftp1.Timeout = (20 * 60)
            End If

            If Not Sftp1.Connected Then
                If InAutoMode Then
                    SendErrorEMail("Could not connect to ShopSite")
                Else
                    MessageBox.Show("Could not connect to ShopSite", "FTP Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Exit Function
            End If
            Sftp1.LocalFile = shopSiteFilename
            If WB_PARM_SITE_OUTPUT_DIR.StartsWith("/") Then
                WB_PARM_SITE_OUTPUT_DIR = WB_PARM_SITE_OUTPUT_DIR.Substring(1)
            End If
            Sftp1.RemoteFile = WB_PARM_SITE_OUTPUT_DIR & My.Computer.FileSystem.GetName(shopSiteFilename)
            Sftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmDefault
            Sftp1.Passive = True
            Sftp1.Upload()
            Dim script As String = WB_PARM_SITE_PRODUCT_POST_URL & My.Computer.FileSystem.GetName(shopSiteFilename)
            docComplete = False

            If Sftp1.Connected Then Sftp1.Logoff()

            ASCMAIN1.Progress("-", "Post")
            WebBrowser1.Navigate("")
            WebBrowser1.Navigate(script)

            System.Threading.Thread.Sleep(2000)
            While Not docComplete
                Application.DoEvents()
            End While
            docComplete = False

            For Each curElement As HtmlElement In WebBrowser1.Document.All

                Select Case curElement.GetAttribute("value")
                    Case ""

                    Case "1"

                    Case "2"

                End Select

                If curElement.GetAttribute("value").Equals("Login") Then
                    curElement.InvokeMember("click")
                    Exit For
                End If
            Next

            System.Threading.Thread.Sleep(5000)

            script = WB_PARM_SITE_PRODUCT_PUB_URL
            ASCMAIN1.Progress("-", "Publish")
            WebBrowser1.Navigate("")
            WebBrowser1.ScriptErrorsSuppressed = True
            WebBrowser1.Navigate(script)

            System.Threading.Thread.Sleep(2000)
            While Not docComplete
                Application.DoEvents()
            End While
            docComplete = False

            ASCMAIN1.Progress("-", "Delete")
            Sftp1.Logon()
            Sftp1.DeleteFile(WB_PARM_SITE_OUTPUT_DIR & My.Computer.FileSystem.GetName(shopSiteFilename))
            If Sftp1.Connected Then Sftp1.Logoff()

            shopSiteFilename = String.Empty
            WebBrowser1.Navigate("")
        Else
            Debug.Print(shopSiteFilename)
            MsgBox("Manually Upload File: " & shopSiteFilename, vbOKOnly, "You Are In Test")
            'Stop 'This needs to be done manually using the shopSiteFilename
        End If

        Me.Cursor = Cursors.Default
        itemUploaded = True
        ASCMAIN1.Progress(String.Empty, String.Empty)

        WebBrowser1.Navigate(String.Empty)
        If Sftp1.Connected Then Sftp1.Logoff()
    End Function

    Private Function FTPInventoryTable(ByVal STYLE_CODE As String) As Boolean
        Try
            Me.Cursor = Cursors.WaitCursor

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
            'LIVE SETTINGS
            ASCMAIN1.Progress("-", "FTP To LIVE Site")
            Sftp1.RemoteHost = "69.39.227.201"
            Sftp1.User = WB_PARM_SITE_USER
            Sftp1.Password = WB_PARM_SITE_PWD
            'DEMO SITE SETTINGS
            'ASCMAIN1.Progress("-", "FTP To Demo Site")
            'Sftp1.RemoteHost = "216.38.11.230"
            'Sftp1.User = "regdemo"
            'Sftp1.Password = "re67hg34"

            Sftp1.RemoteFile = String.Empty
            Sftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")


            ASCMAIN1.Progress("Upload Tables", STYLE_CODE)
            Dim File_Name As String = String.Format("{0}invtbl\{1}.html", WB_PARM_PRODUCTS_DIR, STYLE_CODE)
            Sftp1.LocalFile = File_Name
            If WB_PARM_SITE_OUTPUT_DIR.StartsWith("/") Then WB_PARM_SITE_OUTPUT_DIR = WB_PARM_SITE_OUTPUT_DIR.Substring(1)
            Sftp1.RemoteFile = String.Format("{0}shop/invtbl/{1}.html", WB_PARM_SITE_OUTPUT_DIR, STYLE_CODE)
            Sftp1.Overwrite = True
            Sftp1.Upload()
            Do While Not Sftp1.Idle
            Loop

            Me.Cursor = Cursors.Default

            ASCMAIN1.Progress(String.Empty, String.Empty)

            If Sftp1.Connected Then Sftp1.Logoff()

        Catch ex As Exception
            If InAutoMode Then
                SendErrorEMail("Error ftping table: " & STYLE_CODE & " - " & ex.Message)
            Else
                MessageBox.Show("Error ftping table: " & STYLE_CODE & " - " & ex.Message, "Error", MessageBoxButtons.OK)
            End If
        Finally

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)

            If Sftp1.Connected Then Sftp1.Logoff()
        End Try
    End Function

    Private Function GetMaterials(ByVal STYLE_CODE As String) As String
        Dim Retval As String = ""
        Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        If Not IsNothing(rowICTSTYL1) Then
            Retval = rowICTSTYL1.Item("STYLE_MATL_DESC") & ""
        End If
        If Retval.Length > 100 Then
            Retval = Retval.Substring(0, 99)
        End If
        Return Retval
    End Function

    Private Function GetKeyWords(ByVal STYLE_CODE As String) As String
        Dim Retval As String = ""
        Dim sql As New StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("L3.STYLE_CODE,")
        sql.AppendLine("LOWER(L3.ATTR_CODE) AS ATTR_CODE,")
        sql.AppendLine("LOWER(A1.ATTR_DESC) AS ATTR_DESC")
        sql.AppendLine("FROM ICTSTYL3 L3, ICTATTR1 A1")
        sql.AppendLine("WHERE L3.ATTR_CODE = A1.ATTR_CODE")
        sql.AppendLine(String.Format("AND L3.STYLE_CODE = '{0}'", STYLE_CODE))
        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        For Each rowICTSTYL1 As DataRow In tbl.Rows
            Retval = Retval + ", " + rowICTSTYL1.Item("ATTR_DESC") & ""
        Next
        If Retval.Length >= 2 Then
            Retval = Retval.Substring(2, Retval.Length - 2)
        End If
        Return Retval
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

    Private Sub grdWBTSTYLD_BeforeCellUpdate(sender As Object, e As BeforeCellUpdateEventArgs) Handles grdWBTSTYLD.BeforeCellUpdate
        If Not e.Cell.IsFilterRowCell Then
            If e.Cell.Column.Key = "WEB_IND" Then
                MsgBox("No Changes Allowed To This Column.", vbOKOnly, "Change Cancelled.")
                e.Cancel = True
                'Dim valFrom As String = e.Cell.Value
                'Dim valTo As String = e.Cell.Text
                'Dim cancelUpdate As Boolean = False
                'Dim msg As New StringBuilder With {.Length = 0}
                'Select Case valFrom
                '    Case "W"
                '        If valTo <> "Awaiting Removal" Then
                '            msg.AppendLine("You Can Only Set Items On Web To Awaiting Removal")
                '            cancelUpdate = True
                '        End If
                '    Case "X"
                '        If valTo <> "Awaiting Update" Then
                '            msg.AppendLine("You Can Only Set Items Not On Web To Awaiting Update")
                '            cancelUpdate = True
                '        End If
                '    Case "U"
                '        If valTo <> "Not On Web" Then
                '            msg.AppendLine("You Can Only Set Items Awaiting Update To Not On Web")
                '            cancelUpdate = True
                '        End If
                '    Case "R"
                '        msg.AppendLine("You Can Not Modify Items Awaiting Removal")
                '        cancelUpdate = True
                'End Select
                'If cancelUpdate Then
                '    MsgBox(msg.ToString, vbOKOnly, "Change Cancelled.  See The Rules.")
                '    e.Cancel = cancelUpdate
                'End If

            End If
        End If

    End Sub

    Private Sub grdWBTSTYLD_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdWBTSTYLD.AfterCellUpdate
        If Not e.Cell.IsFilterRowCell Then
            Select Case e.Cell.Column.Key
                Case "WEB_IND"
                    Dim thisGroup As Integer = 0
                    If e.Cell.Value = "X" Then
                        e.Cell.Row.Cells.Item("STYLE_GROUP").Value = thisGroup
                    Else
                        If e.Cell.Row.Cells.Item("STYLE_GROUP").Value = 0 Then
                            thisGroup = dst.Tables.Item("WBTSTYLD").Compute("MAX(STYLE_GROUP)", "")
                            If thisGroup = 0 Then
                                thisGroup = 1
                            End If
                            e.Cell.Row.Cells.Item("STYLE_GROUP").Value = thisGroup
                        End If
                    End If
                    SetRelatedStyles(e.Cell.Row.Cells.Item("STYLE_CODE").Value, e.Cell.Row.Cells.Item("COLOR_CODE").Value, e.Cell.Value, thisGroup)
                Case "ALT_FUT_QTY"
                    If IsNumeric(e.Cell.Row.Cells.Item("ALT_FUT_QTY").Value) Then
                        ALT_FUT_QTY_LAST = e.Cell.Row.Cells.Item("ALT_FUT_QTY").Value
                    End If
                Case "ALT_FUT_DATE"
                    If IsDate(e.Cell.Row.Cells.Item("ALT_FUT_DATE").Value) Then
                        ALT_FUT_DATE_LAST = e.Cell.Row.Cells.Item("ALT_FUT_DATE").Value
                    End If

            End Select
        End If
    End Sub

    Private Sub SetRelatedStyles(ByVal STYLE_CODE As String,
                                 ByVal COLOR_CODE As String,
                                 ByVal NewVal As String,
                                 NewGroup As Integer)
        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE <> '{1}'", STYLE_CODE, COLOR_CODE)
        For Each rowWHTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(filter)
            rowWHTSTYLD.Item("WEB_IND") = NewVal
            If rowWHTSTYLD.Item("STYLE_GROUP") <> NewGroup Then
                rowWHTSTYLD.Item("STYLE_GROUP") = NewGroup
            End If
        Next
    End Sub

    Private Sub chkShowOnlyDiff_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowOnlyDiff.CheckedChanged
        Dim rowFilter As String = ""
        If chkShowOnlyDiff.Checked Then
            rowFilter = String.Format("WEB_IND = '{0}' AND (LAST_ON_HAND <> CURR_ON_HAND)", "W")
        End If
        Dim dvw As DataView = DirectCast(grdWBTSTYLD.DataSource, DataTable).DefaultView
        dvw.RowFilter = rowFilter
    End Sub

    Private Sub imgSTYLE_Click(sender As Object, e As EventArgs) Handles imgSTYLE.Click
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
        If STYLE_CODE.Length > 0 Then
            Dim frmIMAGE As New TAC.TAFIMGV1(Me, STYLE_CODE, "", "M")
            With frmIMAGE
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub chkUseFilter_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseFilter.CheckedChanged
        lblGroup.Visible = (chkUseFilter.Checked = False)
        cboGROUPS.Visible = (chkUseFilter.Checked = False)
    End Sub

    Private Sub btnRunAllGroups_Click(sender As Object, e As EventArgs) Handles btnRunAllGroups.Click
        Me.Cursor = Cursors.WaitCursor
        Dim DoneMsg As String = "Shopsite Created.  Upload File Below To Site."
        Dim LAST_UPDATE As DateTime = Now()
        Dim UploadInventoryOnly As Boolean = True
        Dim UpdatePricing As Boolean = True

        Dim di As New DirectoryInfo(WB_PARM_PRODUCTS_DIR)
        Dim fiArr As FileInfo() = di.GetFiles()
        Dim fri As FileInfo
        For Each fri In fiArr
            Console.WriteLine(fri.Name)
            System.IO.File.Move(WB_PARM_PRODUCTS_DIR & fri.Name, WB_PARM_PRODUCTS_DIR & "Archives\" & fri.Name)
        Next

        Dim DS As DataSet = MakeClassData()
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine("SELECT MIN(STYLE_GROUP) FROM WBTSTYLD")
        ASCMAIN1.sql = SQLS.ToString()
        Dim MIN_GP As Int64 = Val(ASCDATA1.GetDataValue)

        SQLS.Length = 0
        SQLS.AppendLine("SELECT MAX(STYLE_GROUP) FROM WBTSTYLD WHERE STYLE_GROUP < 900")
        ASCMAIN1.sql = SQLS.ToString()
        Dim MAX_GP As Int64 = Val(ASCDATA1.GetDataValue)

        DISABLED_STYLES.Clear()
        For grp As Int64 = MIN_GP To MAX_GP
            CreateProductXmlALL(grp, DS)
        Next
        If DISABLED_STYLES.Count > 0 Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Disabled Items"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine($"There Were {DISABLED_STYLES.Count} Styles Sent As Disabled.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Do You Want To Mark Them As Inactive In Absoution?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                For Each DISABLED_STYLE As String In DISABLED_STYLES
                    Dim fltr As String = $"STYLE_CODE = '{DISABLED_STYLE}'"
                    For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(fltr)
                        rowWBTSTYLD.Item("DATE_DISABLED") = Format(Now(), "MM/dd/yyyy")
                        rowWBTSTYLD.Item("WEB_IND") = "I"
                        rowWBTSTYLD.Item("STYLE_GROUP") = "999"
                    Next
                    'Also remove Style From Web Maint Screen.  9/5/24. W.R.
                    Dim S As New System.Text.StringBuilder With {.Length = 0}
                    S.AppendLine($"DELETE FROM WBTPAGED WHERE STYLE_CODE = '{DISABLED_STYLE}'")
                    ASCMAIN1.sql = SQLS.ToString
                    ASCDATA1.ExecuteSQL()
                Next
            End If
        End If
        Me.Cursor = Cursors.Default
        MsgBox("GetType Your File" & vbCrLf & WB_PARM_PRODUCTS_DIR, vbOKOnly, "Complete")
    End Sub

    Private Sub btnRun999Group_Click(sender As Object, e As EventArgs) Handles btnRun999Group.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Run 999 Group"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("This will Archive All Pending Shopsite")
        iMSG.AppendLine("Upload Files And Generate A New One")
        iMSG.AppendLine("For Only Group 999.")
        iMSG.AppendLine("")
        iMSG.AppendLine("They Should All Be Discontinued.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Is That What You Want?")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult <> MsgBoxResult.Yes Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        Dim DoneMsg As String = "Shopsite Created.  Upload File Below To Site."
        Dim LAST_UPDATE As DateTime = Now()
        Dim UploadInventoryOnly As Boolean = True
        Dim UpdatePricing As Boolean = True

        Dim di As New DirectoryInfo(WB_PARM_PRODUCTS_DIR)
        Dim fiArr As FileInfo() = di.GetFiles()
        Dim fri As FileInfo
        For Each fri In fiArr
            Console.WriteLine(fri.Name)
            System.IO.File.Move(WB_PARM_PRODUCTS_DIR & fri.Name, WB_PARM_PRODUCTS_DIR & "Archives\" & fri.Name)
        Next

        Dim DS As DataSet = MakeClassData()
        CreateProductXmlALL(999, DS)
        Me.Cursor = Cursors.Default
        MsgBox("Get Your File" & vbCrLf & WB_PARM_PRODUCTS_DIR, vbOKOnly, "Complete")
    End Sub

    Private Function MakeClassData() As DataSet
        Dim DS As New DataSet
        Dim sql As New Text.StringBuilder With {.Length = 0}

        sql.Length = 0
        sql.AppendLine("SELECT * FROM ARTCUST1")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ARTCUST1"))

        sql.Length = 0
        sql.AppendLine("SELECT * FROM ICTSTYL1")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTSTYL1"))

        sql.Length = 0
        sql.AppendLine("SELECT * FROM ICTCLAS1")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTCLAS1"))

        sql.Length = 0
        sql.AppendLine("SELECT S1.STYLE_CODE, C1.COLOR_GROUP_CODE")
        sql.AppendLine("FROM ICTSTYC1 S1, ICTCOLR1 C1")
        sql.AppendLine("WHERE S1.COLOR_CODE = C1.COLOR_CODE")
        sql.AppendLine("AND NVL(C1.COLOR_GROUP_CODE,'NULL') <> 'NULL'")
        sql.AppendLine("GROUP BY S1.STYLE_CODE, C1.COLOR_GROUP_CODE")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "COLOR_GROUP_CODE"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("WD.STYLE_CODE,")
        sql.AppendLine("WD.COLOR_CODE,")
        sql.AppendLine("WH.STYLE_DESC_SHORT,")
        sql.AppendLine("S1.STYLE_DESC")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1")
        sql.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        sql.AppendLine("AND WH.STYLE_CODE = S1.STYLE_CODE")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "COLORS"))

        sql.Length = 0
        sql.AppendLine("Select")
        sql.AppendLine("WD.STYLE_CODE, ")
        sql.AppendLine("WD.COLOR_CODE, ")
        sql.AppendLine("WH.STYLE_DESC_SHORT, ")
        sql.AppendLine("S1.STYLE_DESC")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1")
        sql.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        sql.AppendLine("And WH.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("And CURR_ON_HAND > 0")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "CROSS"))

        sql.Length = 0
        sql.AppendLine("SELECT * FROM WBTSTYLD")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "WBTSTYLD"))

        sql.Length = 0
        sql.AppendLine("Select")
        sql.AppendLine("P2.STYLE_CODE,")
        sql.AppendLine("P2.PROMO_UNIT_PRICE,")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("And P1.PROMO_END_DATE >= SYSDATE")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "PROMOS"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("S1.STYLE_CODE, A1.ATTR_DESC")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTSTYL3 S3, ICTATTR1 A1")
        sql.AppendLine("WHERE WH.STYLE_CODE (+) = WD.STYLE_CODE")
        sql.AppendLine("AND WD.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND S1.STYLE_CODE = S3.STYLE_CODE")
        sql.AppendLine("AND S3.ATTR_CODE = A1.ATTR_CODE")
        sql.AppendLine("AND NVL(A1.ATT_RANK,0) = 1")
        sql.AppendLine("GROUP BY S1.STYLE_CODE, A1.ATTR_DESC")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ATTR_DESC"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("S1.STYLE_CODE, Z1.SIZE_CODE")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTSIZE1 Z1")
        sql.AppendLine("WHERE WH.STYLE_CODE = WD.STYLE_CODE")
        sql.AppendLine("AND WH.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND S1.SIZE_CODE = Z1.SIZE_CODE")
        sql.AppendLine("AND S1.STYLE_CLASS_CODE = 'PVC'")
        sql.AppendLine("GROUP BY S1.STYLE_CODE, Z1.SIZE_CODE")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "SIZE_CODE"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("S1.STYLE_CODE, C1.COLOR_DESC")
        sql.AppendLine("FROM WBTSTYLH WH, WBTSTYLD WD, ICTSTYL1 S1, ICTCOLR1 C1")
        sql.AppendLine("WHERE WH.STYLE_CODE (+) = WD.STYLE_CODE")
        sql.AppendLine("AND WD.STYLE_CODE = S1.STYLE_CODE")
        sql.AppendLine("AND WD.COLOR_CODE = C1.COLOR_CODE")
        sql.AppendLine("GROUP BY S1.STYLE_CODE, C1.COLOR_DESC")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "COLOR_DESC"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("WBTPAGED.STYLE_CODE, WBTPAGEH.PAGE_CODE, WBTPAGEH.PAGE_NAME")
        sql.AppendLine("FROM WBTPAGEH, WBTPAGED")
        sql.AppendLine("WHERE WBTPAGEH.PAGE_CODE = WBTPAGED.PAGE_CODE")
        sql.AppendLine("AND NVL(WBTPAGEH.PAGE_STATUS,'A') = 'A'")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "WBTPAGEX"))

        sql.Length = 0
        sql.AppendLine("SELECT")
        sql.AppendLine("SL.STYLE_CODE,")
        sql.AppendLine("SC.COLOR_CODE,")
        sql.AppendLine("SL.STYLE_DESC,")
        sql.AppendLine("C1.COLOR_DESC,")
        sql.AppendLine("C1.COLOR_CODE_LONG,")
        sql.AppendLine("SL.STYLE_STATUS,")
        sql.AppendLine("WS.WEB_IND,")
        sql.AppendLine("SC.STYLE_COLOR_STATUS,")
        sql.AppendLine("SL.INNER_PACK_QTY,")
        sql.AppendLine("SL.CARTON_PACK_QTY,")
        sql.AppendLine("SL.STYLE_UOM,")
        sql.AppendLine("SL.SUB_UNIT_PACK_QTY,")
        sql.AppendLine("SL.STYLE_CLASS_CODE,")
        sql.AppendLine("CL.STYLE_CLASS_DESC,")
        sql.AppendLine("SL.STYLE_SO_QTY_MIN,")
        sql.AppendLine("SL.STYLE_MATL_DESC,")
        sql.AppendLine("WH.STYLE_DESC_SHORT,")
        sql.AppendLine("NVL(SL.SIZE_CODE,'') AS SIZE_CODE,")
        sql.AppendLine("SC.UPC_CODE,")
        sql.AppendLine("WH.SEARCH_KEYWORDS,")
        sql.AppendLine("WH.MATERIALS,")
        sql.AppendLine("WH.META_DESC,")
        sql.AppendLine("NVL(SC.THEME_CODE,'') AS THEME_CODE,")
        sql.AppendLine("0 as CURR_QTY_AVAIL,")
        sql.AppendLine("0 as FUT_QTY_AVAIL,")
        sql.AppendLine("'          ' AS FUT_DATE,")
        sql.AppendLine("NVL(WS.ALT_FUT_QTY,0) AS ALT_FUT_QTY,")
        sql.AppendLine("WS.ALT_FUT_DATE,")
        sql.AppendLine("WS.FLAG_NEW")
        sql.AppendLine("FROM WBTSTYLD WS, WBTSTYLH WH, ICTSTYL1 SL, ICTSTYC1 SC, ICTSTAT2 ST, ICTCLAS1 CL, ICTCOLR1 C1")
        sql.AppendLine("WHERE WS.STYLE_CODE = SL.STYLE_CODE")
        sql.AppendLine("AND WS.STYLE_CODE = WH.STYLE_CODE (+)")
        sql.AppendLine("AND WS.COLOR_CODE = SC.COLOR_CODE")
        sql.AppendLine("AND SL.STYLE_CODE = SC.STYLE_CODE")
        sql.AppendLine("AND SC.STYLE_CODE = ST.STYLE_CODE (+)")
        sql.AppendLine("AND SC.COLOR_CODE = ST.COLOR_CODE (+)")
        sql.AppendLine("AND SL.STYLE_CLASS_CODE = CL.STYLE_CLASS_CODE")
        sql.AppendLine("AND SC.COLOR_CODE = C1.COLOR_CODE")
        sql.AppendLine("AND ST.WHSE_CODE (+) = 'MS'")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "STATUS"))
        DS.Tables("STATUS").Columns("CURR_QTY_AVAIL").ReadOnly = False
        DS.Tables("STATUS").Columns("FUT_QTY_AVAIL").ReadOnly = False
        DS.Tables("STATUS").Columns("FUT_DATE").ReadOnly = False

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTSTDQ1")
        sql.AppendLine("WHERE WHSE_CODE = 'MS'")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTSTDQ1"))

        For Each rowSTAT As DataRow In DS.Tables("STATUS").Select()
            Dim SC As String = rowSTAT.Item("STYLE_CODE").ToString & String.Empty
            Dim CC As String = rowSTAT.Item("COLOR_CODE").ToString & String.Empty
            Dim CURR_QTY_AVAIL As Int64 = 0
            Dim FUT_QTY_AVAIL As Int64 = 0
            Dim FUT_DATE As String = ""

            Dim CFilter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", SC, CC)
            For Each rowICTSTDQ1 As DataRow In DS.Tables("ICTSTDQ1").Select(CFilter, "STATUS_DATE")
                If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) Then
                    If CDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) <= Now().AddDays(1) Then
                        CURR_QTY_AVAIL = CURR_QTY_AVAIL + Val(rowICTSTDQ1.Item("QTY_ATS").ToString & String.Empty)
                    Else
                        If IsDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty) Then
                            FUT_DATE = CDate(rowICTSTDQ1.Item("STATUS_DATE").ToString & String.Empty).ToShortDateString
                            FUT_QTY_AVAIL = FUT_QTY_AVAIL + Val(rowICTSTDQ1.Item("QTY_ATS").ToString & String.Empty)
                        End If
                    End If
                End If
            Next
            rowSTAT.Item("CURR_QTY_AVAIL") = CURR_QTY_AVAIL
            If IsDate(rowSTAT.Item("ALT_FUT_DATE").ToString & String.Empty) And Val(rowSTAT.Item("ALT_FUT_QTY").ToString & String.Empty) > 0 Then
                rowSTAT.Item("FUT_QTY_AVAIL") = Val(Val(rowSTAT.Item("ALT_FUT_QTY").ToString & String.Empty))
                rowSTAT.Item("FUT_DATE") = CDate(rowSTAT.Item("ALT_FUT_DATE").ToString & String.Empty).ToShortDateString
            Else
                rowSTAT.Item("FUT_QTY_AVAIL") = FUT_QTY_AVAIL
                rowSTAT.Item("FUT_DATE") = FUT_DATE
            End If
        Next

        sql.Length = 0
        sql.AppendLine("SELECT * FROM (")
        sql.AppendLine("SELECT I3.STYLE_CODE, A1.ATTR_DESC")
        sql.AppendLine("FROM ICTSTYL3 I3, ICTATTR1 A1")
        sql.AppendLine("WHERE I3.ATTR_CODE = A1.ATTR_CODE")
        sql.AppendLine("AND NVL(A1.ATTR_DESC,'NULL') <> 'NULL'")
        sql.AppendLine("GROUP BY I3.STYLE_CODE, A1.ATTR_DESC")
        sql.AppendLine("UNION")
        sql.AppendLine("SELECT S1.STYLE_CODE, C1.STYLE_CLASS_DESC AS ATTR_DESC")
        sql.AppendLine("FROM ICTSTYL1 S1, ICTCLAS1 C1")
        sql.AppendLine("WHERE S1.STYLE_CLASS_CODE = C1.STYLE_CLASS_CODE")
        sql.AppendLine("GROUP BY S1.STYLE_CODE, C1.STYLE_CLASS_DESC")
        sql.AppendLine("UNION")
        sql.AppendLine("SELECT S1.SIZE_CODE, Z1.SIZE_DESC AS ATTR_DESC")
        sql.AppendLine("FROM ICTSTYL1 S1, ICTSIZE1 Z1")
        sql.AppendLine("WHERE S1.SIZE_CODE = Z1.SIZE_CODE")
        sql.AppendLine("GROUP BY S1.SIZE_CODE, Z1.SIZE_DESC")
        sql.AppendLine("UNION")
        sql.AppendLine("SELECT I3.STYLE_CODE, W1.ATTR_DESC")
        sql.AppendLine("FROM ICTSTYL3 I3, ICTATTR1 A1, WBTATTR1 W1")
        sql.AppendLine("WHERE I3.ATTR_CODE = A1.ATTR_CODE")
        sql.AppendLine("AND A1.ATTR_CODE = W1.ATTR_CODE")
        sql.AppendLine("AND NVL(A1.ATTR_DESC,'NULL') <> 'NULL'")
        sql.AppendLine("GROUP BY I3.STYLE_CODE, W1.ATTR_DESC")
        sql.AppendLine(")")
        DS.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ATTR_CODE"))

        Return DS
    End Function

    Private Function CreateProductXmlALL(ByVal GroupNo As Int64, ByRef DS As DataSet) As Boolean
        Dim Retval As Boolean = True
        Dim batchFilter As String = ""
        Dim styleListAll As List(Of String) = New List(Of String)
        Dim styleListInactiveAll As List(Of String) = New List(Of String)

        'If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then
        '    Stop
        '    batchFilter = String.Format("STYLE_CODE = '{0}'", "MTF24040")
        'Else
        If GroupNo = 999 Then
            batchFilter = String.Format("STYLE_GROUP = 999")
        Else
            batchFilter = String.Format("WEB_IND = '{0}' AND STYLE_GROUP = {1}", "W", GroupNo)
        End If
        'End If

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Create Group " & GroupNo, "")
            'Dim productXML As New WBCITEM2(dst.Tables("WBTSTYLD"))
            shopSiteFilename = WB_PARM_PRODUCTS_DIR & "SHOP_" & DateTime.Now.ToString("yyyyMMddhhmm") & "_" & GroupNo & ".xml"
            With New WBCITEMA(DS, DISABLED_STYLES)
                Dim StyleCount As Int64 = 0
                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select(batchFilter)

                    If Not styleListAll.Contains(rowWBTSTYLD.Item("STYLE_CODE")) Then
                        styleListAll.Add(rowWBTSTYLD.Item("STYLE_CODE"))
                    End If
                    StyleCount += 1
                    ASCMAIN1.Progress("-", "Color: " & StyleCount)
                    'ASCMAIN1.Progress("-", rowWBTSTYLD.Item("STYLE_CODE"))
                    Application.DoEvents()

                    Dim rowCnt As Int64 = .AddStyle(rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE"), styleListInactiveAll, False, , True)
                Next

                For Each style As String In styleListAll
                    'ASCMAIN1.Progress("Parents", style)
                    ASCMAIN1.Progress("-", "Parent: " & StyleCount)
                    Dim DefaultColorFilter As String = String.Format("STYLE_CODE = '{0}' AND CURR_ON_HAND > 0", style)

                    Dim DefaultColor As String = ""
                    Dim rowColor As DataRow = dst.Tables("WBTSTYLD").Select(DefaultColorFilter).FirstOrDefault
                    If IsNothing(rowColor) Then
                        Dim DefaultColorFilter2 As String = String.Format("STYLE_CODE = '{0}'", style)
                        Dim rowColor2 As DataRow = dst.Tables("WBTSTYLD").Select(DefaultColorFilter2).FirstOrDefault
                        If Not IsNothing(rowColor2) Then
                            DefaultColor = rowColor2.Item("COLOR_CODE").ToString & String.Empty
                            If DefaultColor = "" Then
                                MsgBox("Error With Default Color", vbOKOnly, "Error")
                                Stop
                            End If
                        End If
                    Else
                        DefaultColor = rowColor.Item("COLOR_CODE").ToString & String.Empty
                        If DefaultColor = "" Then
                            MsgBox("Error With Default Color", vbOKOnly, "Error")
                            Stop
                        End If
                    End If
                    StyleCount += 1

                    .AddStyle(style, DefaultColor, styleListInactiveAll, False, True, True)

                Next

                If styleListInactiveAll.Count + styleListAll.Count = 0 Then
                    Retval = False
                Else
                    Retval = True
                End If

                If Retval = True Then
                    Dim xmlLabelRequest As XmlDocument = .GetXMLDocument
                    ASCMAIN1.Progress("Saving XML Document", "")
                    With xmlLabelRequest
                        .Save(shopSiteFilename)
                    End With
                End If

            End With

        Catch ex As Exception
            shopSiteFilename = String.Empty
            MessageBox.Show("Error creating XML Document: " & ex.Message, "Error", MessageBoxButtons.OK)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

        Return Retval
    End Function

    Private Sub chkInventoryRunning_CheckedChanged(sender As Object, e As EventArgs) Handles chkInventoryRunning.CheckedChanged
        If chkInventoryRunning.Checked Then
            tmrInventory.Start()
        Else
            tmrInventory.Stop()
        End If

    End Sub

    Private Sub tmrInventory_Tick(sender As Object, e As EventArgs) Handles tmrInventory.Tick


        Dim HR As Int64 = Now().Hour
        Dim MN As Int64 = Now().Minute
        Dim INT As New List(Of Int64)
        INT.Clear()
        INT.Add(0)
        For i As Int64 = 10 To 50 Step 10
            INT.Add(i)
        Next

        If ASCMAIN1.Running_in_VS Then Stop
        If INT.Contains(MN) And MN <> LASTMIN Then
            If ASCMAIN1.Running_in_VS Then Stop
            LASTMIN = MN
            txtInventoryLast.Text = String.Format("Last: {0}", Now().ToShortTimeString)
            uploadShopsiteInventory()
            uploadShipTos()
            sendCustomerPricing()
            sendImagesEmail()
        Else
            If txtInventoryLast.Text = "" Then
                txtInventoryLast.Text = "Waiting...."
            End If
        End If
    End Sub

    Private Sub sendImagesEmail()
        Dim ZIP_FOLDER As String = "\\192.168.110.233\c$\ShopsiteService\Data\OrderImageZips\"
        Dim TXT_FOLDER As String = "\\192.168.110.233\c$\ShopsiteService\Data\Files\"
        Dim BASE_URL As String = "https://www.regency-rib.com/images/"
        'Dim ORDR_NO As String = "0000771989"
        'Dim GIVENNAME As String = "Mario (Big Pappa)"
        Dim ATTACHMENTs As New Dictionary(Of String, String)
        'ATTACHMENTs.Add($"{ORDR_NO}.zip", $"\\192.168.110.233\c$\ShopsiteService\Data\OrderImageZips\{ORDR_NO}.zip")

        Dim SQLW As New StringBuilder
        SQLW.Length = 0
        SQLW.AppendLine("SELECT *")
        SQLW.AppendLine("FROM WBTIMGR1")
        Fill_Records("WBTIMGR1", , , SQLW.ToString)

        'Dim di As DirectoryInfo = System.IO.DirectoryInfo(ZIP_FOLDER)
        Dim files As String() = Directory.GetFiles(ZIP_FOLDER)
        For Each file As String In files
            If file.ToUpper.EndsWith(".ZIP") Then
                Dim ORDR_NO As String = file.Replace(".ZIP", "").Replace(".zip", "").Replace(ZIP_FOLDER, "")
                Dim FLTR As String = $"ORDR_NO = '{ORDR_NO}'"
                If dst.Tables.Item("WBTIMGR1").Select(FLTR).Count = 0 Then
                    Dim TXT_FILE As String = $"{TXT_FOLDER}{ORDR_NO}.txt"
                    If System.IO.File.Exists(TXT_FILE) Then
                        Dim fileContent As String = System.IO.File.ReadAllText(TXT_FILE)
                        Dim fileData As String() = fileContent.Split("|")
                        If fileData.Length >= 1 Then
                            Dim EMAIL As String = fileData(0).ToUpper
                            SQLW.Length = 0
                            SQLW.AppendLine("SELECT GIVENNAME")
                            SQLW.AppendLine("FROM WBTCUST1")
                            SQLW.AppendLine($"WHERE UPPER(EMAIL) = UPPER('{EMAIL}')")
                            ASCMAIN1.sql = SQLW.ToString()
                            Dim GIVENNAME As String = ASCDATA1.GetDataValue

                            SQLW.Length = 0
                            SQLW.AppendLine("SELECT FAMILYNAME")
                            SQLW.AppendLine("FROM WBTCUST1")
                            SQLW.AppendLine($"WHERE UPPER(EMAIL) = UPPER('{EMAIL}')")
                            ASCMAIN1.sql = SQLW.ToString()
                            Dim FAMILYNAME As String = ASCDATA1.GetDataValue

                            If GIVENNAME.Length > 0 Then
                                Dim SUBJECT As String = $"Requested Images For Order {ORDR_NO}"
                                Dim BDY As New StringBuilder With {.Length = 0}
                                If GIVENNAME.Length > 0 Then
                                    BDY.AppendLine($"Hi {GIVENNAME};")
                                Else
                                    BDY.AppendLine("Hi;")
                                End If
                                BDY.AppendLine("")
                                BDY.AppendLine($"Please follow the link below to view to download images for order #{ORDR_NO}:")
                                BDY.AppendLine($"<a href={BASE_URL}{ORDR_NO}.zip >Click Here To Download Your File</a>")
                                BDY.AppendLine("")
                                BDY.AppendLine("Thank You,")
                                BDY.AppendLine("")
                                BDY.AppendLine("Regency International")
                                BDY.AppendLine("800.782.7810")
                                BDY.AppendLine("")
                                BDY.AppendLine("For questions, please contact <a href='mailto:hq@regency-rib.com'>Customer Service</a> or your <a href='https://www.regency-rib.com/locate-sales-representative.html/'>Sales Representative</a> directly.")

                                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                                'EMAIL_ADDRESSs.Add("whr@waynerichmond.net", "Wayne Richmond")
                                EMAIL_ADDRESSs.Add("mariog@regency-rib.com", "Mario Arenas")
                                EMAIL_ADDRESSs.Add(EMAIL, $"{GIVENNAME} {FAMILYNAME}")

                                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                                        SUBJECT, "SHOPSI", True, False, ORDR_NO, "Order Images", "Order Images", BDY.ToString)
                                If SEND_NO.Length > 0 Then
                                    Dim newWBTIMGR1 As DataRow = dst.Tables("WBTIMGR1").NewRow
                                    newWBTIMGR1.Item("ORDR_NO") = ORDR_NO
                                    newWBTIMGR1.Item("REQ_DATE") = Now()
                                    newWBTIMGR1.Item("REQ_EMAIL") = EMAIL.ToUpper
                                    newWBTIMGR1.Item("REQ_STATUS") = "X"
                                    dst.Tables("WBTIMGR1").Rows.Add(newWBTIMGR1)
                                    Update_Record_TDA("WBTIMGR1")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next


    End Sub

    Private Sub uploadShipTos()
        Dim UserName As String = "regency-rib"
        Dim Password As String = "joydHUJ3"
        Dim RemoteHost As String = "regency-rib.com" '69.39.227.201
        Dim RemotePath As String = "www/customers/shipAddresses"
        'Dim ServerFilePath As String = "S:\RGI\Archive\Shopsite\"

        Dim _WBCSHIPT As New WBCSHIPT()
        Dim FileName As String = _WBCSHIPT.MakeFile(ASCMAIN1.Folders("Temp").ToString)
        If _WBCSHIPT.ErrMsg.Length = 0 Then
            Dim FtpShopSite As New nsoftware.IPWorks.Ftp
            With FtpShopSite
                Try
                    If .Connected = True Then
                        .Logoff()
                    End If
                    .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                    .User = UserName
                    .Password = Password
                    .RemoteHost = RemoteHost
                    .RemotePath = RemotePath
                    .Logon()
                    .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                    .LocalFile = FileName
                    .RemoteFile = _WBCSHIPT.FileNameCSV
                    .Overwrite = True
                    .Upload()
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                Catch ex As Exception
                    MsgBox(ex.Message.ToString, vbExclamation, "Error Creating Ship To File")
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                End Try
            End With
        Else
            MsgBox(_WBCSHIPT.ErrMsg, vbExclamation, "Error Creating Ship To File")
        End If
    End Sub

    Private Sub btnCheckInventory_Click(sender As Object, e As EventArgs) Handles btnCheckInventory.Click
        ASCMAIN1.Progress("Checking Inventory", Now.ToShortTimeString)
        Dim RemotePath As String = "www/inventory"
        Dim LocalFileName As String = "inventory_tmp.csv"
        Dim RemoteFileName As String = "inventory.csv"
        Dim str As New StringBuilder
        Dim sql As New StringBuilder With {.Length = 0}
        Dim LastFileTime As String = ""

        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                If .Connected = True Then
                    .Logoff()
                End If
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = WB_PARM_SITE_USER
                .Password = WB_PARM_SITE_PWD
                .RemoteHost = FTP_REMOTE_HOST
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                '.LocalFile = localFolder
                .RemoteFile = RemoteFileName
                '.Overwrite = False
                .Overwrite = False
                'If Not .FileExists() Then
                '.Download()
                LastFileTime = .FileTime
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
                If LastFileTime.Length = 0 Then
                    MsgBox("Error Getting FTP File", vbExclamation, "Check Inventory File")
                Else
                    MsgBox("Last Modified: " & LastFileTime, vbInformation, "Check Inventory File")
                End If
            Catch ex As Exception
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
                MsgBox("Error Getting FTP File", vbExclamation, "Check Inventory File")
            End Try
        End With

        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub btnRemoveInventory_Click(sender As Object, e As EventArgs) Handles btnRemoveInventory.Click
        Dim iResult As MsgBoxResult
        Dim iTitle As String = "Remove Inventory File?"
        Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
        iMSG.AppendLine("This Will Remove The Inventory File From")
        iMSG.AppendLine("The FTP Site.  This Will Cause Shopsite")
        iMSG.AppendLine("To Deduct Waiting Orders From Inventory")
        iMSG.AppendLine("On The Site Until A New Inventory File")
        iMSG.AppendLine("Is Uploaded.")
        iMSG.AppendLine("")
        iMSG.AppendLine("Are You Ready??")
        iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
        If iResult = MsgBoxResult.Yes Then
            ASCMAIN1.Progress("Uploading Invetory", Now.ToShortTimeString)
            Dim RemotePath As String = "www/inventory"
            'Dim LocalFileName As String = "inventory_tmp.csv"
            Dim RemoteFileName As String = "inventory.csv"
            Dim str As New StringBuilder
            Dim sql As New StringBuilder With {.Length = 0}
            'Dim LastFileTime As String = ""

            Dim FtpShopSite As New nsoftware.IPWorks.Ftp
            With FtpShopSite
                Try
                    If .Connected = True Then
                        .Logoff()
                    End If
                    .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                    .User = WB_PARM_SITE_USER
                    .Password = WB_PARM_SITE_PWD
                    .RemoteHost = FTP_REMOTE_HOST
                    .RemotePath = RemotePath
                    .Logon()
                    .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                    '.LocalFile = localFolder
                    '.RemoteFile = RemoteFileName
                    '.Overwrite = False
                    '.Overwrite = True
                    'If Not .FileExists() Then
                    '.Download()
                    'LastFileTime = .FileTime
                    .DeleteFile(RemoteFileName)
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                Catch ex As Exception
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                    MsgBox("Error Removing FTP File", vbExclamation, "Check Inventory File")
                End Try
            End With
        Else
            MsgBox("Removal Aborted", vbExclamation, "Chicken!")
        End If

        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub rdoALTMODE_ADD_CheckedChanged(sender As Object, e As EventArgs) Handles rdoALTMODE_ADD.CheckedChanged
        If rdoALTMODE_ADD.Checked Then
            btnAltInventory.Text = "Add Alternates"
            lblQTY.Visible = True
            numQTY.Visible = True
            numQTY.Text = "0"
            lblALTDATE.Visible = True
            dteALTDATE.Visible = True
            dteALTDATE.Value = Now().AddMonths(3)
        End If
    End Sub

    Private Sub rdoALTMODE_DEL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoALTMODE_DEL.CheckedChanged
        If rdoALTMODE_DEL.Checked Then
            btnAltInventory.Text = "Remove Alternates"
            lblQTY.Visible = False
            numQTY.Visible = False
            numQTY.Text = "0"
            lblALTDATE.Visible = False
            dteALTDATE.Visible = False
            dteALTDATE.Value = Now().AddMonths(-3)
        End If
    End Sub

    Private Sub btnAltInventory_Click(sender As Object, e As EventArgs) Handles btnAltInventory.Click
        Dim eMsg As New StringBuilder With {.Length = 0}
        Dim SelRows As Int64 = grdWBTSTYLD.Selected.Rows.Count
        If SelRows < 1 Then
            eMsg.AppendLine("You Must Select At Least One Row.")
        End If
        If rdoALTMODE_ADD.Checked Then
            If Not IsNumeric(numQTY.Text) Then
                eMsg.AppendLine("Qty Must Be Numeric.")
            Else
                If rdoALTMODE_ADD.Checked Then
                    If Val(numQTY.Text) < 1 Or Val(numQTY.Text) > 20000 Then
                        eMsg.AppendLine("Qty Must Be Between 1 - 20000.")
                    End If
                End If
            End If
            If Not IsDate(dteALTDATE.Value) Then
                eMsg.AppendLine("Invalid Date.")
            Else
                If CDate(dteALTDATE.Value) < Now() Or CDate(dteALTDATE.Value) > Now().AddYears(1) Then
                    eMsg.AppendLine("Date Can Only Be Now to 1 year")
                End If
            End If
        End If
        If eMsg.Length > 0 Then
            MsgBox(eMsg.ToString, vbExclamation, "Can Not Update")
        Else
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Update Alternates"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("You Are About To Update The")
            iMSG.AppendLine(String.Format("Alternate Values On {0} Style/Colors.", SelRows))
            iMSG.AppendLine("")
            iMSG.AppendLine("Are You Sure?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult = MsgBoxResult.Yes Then
                For Each thisRow As UltraGridRow In grdWBTSTYLD.Selected.Rows
                    If rdoALTMODE_ADD.Checked Then
                        thisRow.Cells("ALT_FUT_QTY").Value = Val(numQTY.Text)
                        thisRow.Cells("ALT_FUT_DATE").Value = CDate(dteALTDATE.Value)
                    Else
                        thisRow.Cells("ALT_FUT_QTY").Value = Null
                        thisRow.Cells("ALT_FUT_DATE").Value = Null
                    End If
                Next
                grdWBTSTYLD.UpdateData()
                grdWBTSTYLD.Refresh()
            Else
                MsgBox("Aborted", vbOKOnly, "Chicken")
            End If
        End If
    End Sub

#End Region

#Region "Upload Web Customer Terms"
    Private Sub sendCustomerPricing()
        Dim OutBoundFile As String = "accounts.csv"
        Dim ErrMsg As New StringBuilder With {.Length = 0}
        Dim TempFolder As String = ASCMAIN1.Folders("Temp").ToString
        If Not TempFolder.EndsWith("\") Then
            TempFolder = TempFolder & "\"
        End If
        Dim LocalFile As String = String.Format("{0}{1}", TempFolder, OutBoundFile)
        If File.Exists(LocalFile) Then
            File.Delete(LocalFile)
        End If

        If ErrMsg.Length = 0 Then
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            WebCustOutboundCheck(ErrMsg, LocalFile, OutBoundFile)
        End If

        If ErrMsg.Length = 0 Then
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            WebCustOutboundCreate(ErrMsg, LocalFile)
        End If

        If ErrMsg.Length = 0 Then
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            WebCustOutboundSend(ErrMsg, LocalFile, OutBoundFile)
        End If

        If ErrMsg.Length = 0 Then
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            WebCustTMPDelete(ErrMsg, LocalFile)
        End If

        If ErrMsg.Length = 0 Then
            If (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wayne") Then Stop
            WebBrowser1.Visible = False
            WebBrowser1.Navigate(New Uri("https://www.regency-rib.com/customers/import.php"))
            WebBrowser1.Navigate("about:blank")
        End If
    End Sub
    Private Sub WebCustOutboundCheck(errMsg As StringBuilder, ByVal localFile As String, ByVal OutBoundFile As String)
        Dim UserName As String = "regency-rib"
        Dim Password As String = "joydHUJ3"
        Dim RemoteHost As String = "regency-rib.com" '69.39.227.201
        Dim RemotePath As String = "www/customers"
        Dim ServerFilePath As String = "S:\RGI\Archive\Shopsite\"

        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = OutBoundFile
                .Overwrite = False
                If .FileExists() Then
                    errMsg.AppendLine("New Customer File Still Waiting On ShopSite.")
                    .DoEvents()
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                End If
            Catch ex As Exception
                errMsg.AppendLine(ex.Message.ToString)
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
            End Try
        End With

    End Sub

    Private Sub WebCustOutboundCreate(errMsg As StringBuilder, localFile As String)
        Try
            Dim Retval As Boolean = False
            Dim str As New StringBuilder
            Dim sql As New StringBuilder With {.Length = 0}
            sql.AppendLine("SELECT")
            sql.AppendLine(String.Format("A1.CUST_CODE {0}Regency Account #{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_NAME {0}Business Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.GIVENNAME {0}First Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.FAMILYNAME {0}Last Name{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.FULLNAME {0}Contact Name{0},", Chr(34)))
            sql.AppendLine(String.Format("'' {0}Contact Number{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.EMAIL {0}Email Address{0},", Chr(34)))
            sql.AppendLine(String.Format("W1.PASSWORD {0}Password{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ADDR1 {0}Business Address Line 1{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ADDR2 {0}Business Address Line 2{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_CITY {0}City{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_STATE {0}State{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_ZIP_CODE {0}Zip Code{0},", Chr(34)))
            sql.AppendLine(String.Format("A1.CUST_COUNTRY {0}Country{0},", Chr(34)))
            sql.AppendLine(String.Format("'00' {0}Price Group{0},", Chr(34)))
            sql.AppendLine(String.Format("'1' {0}Welcome Type{0},", Chr(34)))
            sql.AppendLine(String.Format("'1' {0}Terms{0}", Chr(34)))
            sql.AppendLine("FROM WBTCUST1 W1, ARTCUST1 A1")
            sql.AppendLine("WHERE W1.CUST_CODE_ACTUAL = A1.CUST_CODE")
            sql.AppendLine("AND (STATUS = 'A' OR STATUS = 'U')")
            sql.AppendLine("ORDER BY W1.FULLNAME")
            'str.Append(Chr(34))
            Dim tblAccounts As DataTable = ASCDATA1.GetDataTable(sql.ToString())
            For Each dc As DataColumn In tblAccounts.Columns
                str.Append(Chr(34) & dc.ColumnName.ToString & Chr(34) & ",")
            Next
            str.Replace(",", vbNewLine, str.Length - 1, 1)
            tblAccounts.Columns.Item("Price Group").ReadOnly = False
            tblAccounts.Columns.Item("Terms").ReadOnly = False
            For Each rowACCOUNTS As DataRow In tblAccounts.Rows
                Dim EMAIL As String = rowACCOUNTS.Item("Email Address").ToString.ToUpper & String.Empty
                Dim CUST_CODE As String = rowACCOUNTS.Item("Regency Account #").ToString.ToUpper & String.Empty
                rowACCOUNTS.Item("PASSWORD") = ""
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If Not IsNothing(rowARTCUST1) Then
                    rowACCOUNTS.Item("Price Group") = CalculatepriceGroup(rowARTCUST1)
                    rowACCOUNTS.Item("Terms") = CalculateTerms(rowARTCUST1)
                End If
                Dim colIndex As Integer = 0
                For Each field As Object In rowACCOUNTS.ItemArray
                    If colIndex = 6 Or colIndex = 14 Or colIndex = 16 Then
                        str.Append(Chr(34) & field.ToString & Chr(34) & ",")
                    Else
                        str.Append(Chr(34) & "" & Chr(34) & ",")
                    End If
                    colIndex += 1
                Next
                str.Replace(",", vbNewLine, str.Length - 1, 1)
            Next
            Try
                My.Computer.FileSystem.WriteAllText(localFile, str.ToString, False)
                Retval = True
            Catch ex As Exception
                MsgBox("Error Creating Web Customer Output File", vbExclamation, "Error")
                Retval = False
            End Try
        Catch ex As Exception
            errMsg.AppendLine(ex.Message.ToString)
        End Try
    End Sub

    Private Sub WebCustOutboundSend(errMsg As StringBuilder, localFile As String, ByVal OutBoundFile As String)
        Dim UserName As String = "regency-rib"
        Dim Password As String = "joydHUJ3"
        Dim RemoteHost As String = "regency-rib.com" '69.39.227.201
        Dim RemotePath As String = "www/customers"
        Dim ServerFilePath As String = "S:\RGI\Archive\Shopsite\"
        Dim FtpShopSite As New nsoftware.IPWorks.Ftp
        With FtpShopSite
            Try
                If .Connected = True Then
                    .Logoff()
                End If
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
                .User = UserName
                .Password = Password
                .RemoteHost = RemoteHost
                .RemotePath = RemotePath
                .Logon()
                .TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
                .LocalFile = localFile
                .RemoteFile = OutBoundFile
                .Overwrite = False
                If Not .FileExists() Then
                    .Upload()
                    .Logoff()
                    Do While .Connected
                        .DoEvents()
                    Loop
                End If
            Catch ex As Exception
                errMsg.AppendLine(ex.Message.ToString)
                .Logoff()
                Do While .Connected
                    .DoEvents()
                Loop
            End Try
        End With
    End Sub

    Private Sub WebCustTMPDelete(ByRef errMsg As StringBuilder, ByVal localFile As String)
        'Delete tmp file.
        Try
            If File.Exists(localFile) Then
                File.Delete(localFile)
            Else
                errMsg.AppendLine("No Local File Found To Delete.")
            End If
        Catch ex As Exception
            errMsg.AppendLine(ex.Message.ToString)
        End Try
    End Sub

    Private Function CalculatepriceGroup(ByRef rowARTCUST1 As DataRow) As String
        Dim RetVal As String = "0"
        Dim CUST_PRICE_TIER As String = rowARTCUST1.Item("CUST_PRICE_TIER").ToString & String.Empty
        Dim CUST_DISC_PCT_EXTRA As String = rowARTCUST1.Item("CUST_DISC_PCT_EXTRA").ToString & String.Empty
        Dim CUST_DISC_PCT As Int64 = Val(rowARTCUST1.Item("CUST_DISC_PCT").ToString & String.Empty)
        If CUST_DISC_PCT_EXTRA = "" Then
            CUST_DISC_PCT_EXTRA = "0"
        End If
        Select Case CUST_PRICE_TIER
            Case "PC"
                Select Case CUST_DISC_PCT_EXTRA
                    Case "1"
                        RetVal = "1"
                    Case "2"
                        RetVal = "2"
                End Select
            Case "HC"
                RetVal = "3"
            Case "FC"
                RetVal = "4"
            Case "SP"
                Select Case CUST_DISC_PCT
                    Case 52
                        RetVal = "5"
                    Case 54
                        RetVal = "6"
                    Case 55
                        RetVal = "7"
                    Case 56
                        RetVal = "8"
                    Case 57
                        RetVal = "9"
                    Case 59
                        RetVal = "10"
                End Select
        End Select
        Return RetVal
    End Function

    Private Function CalculateTerms(ByRef rowARTCUST1 As DataRow) As String
        Dim RetVal As String = "1" 'Credit Card Only.
        Dim TERM_CODE As String = rowARTCUST1.Item("TERM_CODE").ToString & String.Empty
        If TERM_CODE = "N30" Then
            RetVal = "2" 'Net 30 or Credit Card.
        End If
        Return RetVal
    End Function

    Private Sub btnReGroup_Click(sender As Object, e As EventArgs) Handles btnReGroup.Click
        If (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "site.admin" Or ASCMAIN1.USER_ID = "mariog") Then
            Dim iResult As MsgBoxResult
            Dim iTitle As String = "Make New Groups"
            Dim iMSG As New System.Text.StringBuilder With {.Length = 0}
            iMSG.AppendLine("This Will Move All Syles Waiting For")
            iMSG.AppendLine("Upload To On The Web And Create New")
            iMSG.AppendLine("Groups For All Of Them Except 999.")
            iMSG.AppendLine("")
            iMSG.AppendLine("You Will Be Asked For How Many Styles")
            iMSG.AppendLine("You Want in The Group.  The Default")
            iMSG.AppendLine("Is 500 With A Range Between 200-1000.")
            iMSG.AppendLine("")
            iMSG.AppendLine("Are You Ready?")
            iResult = MsgBox(iMSG.ToString(), MsgBoxStyle.YesNo, iTitle)
            If iResult <> MsgBoxResult.Yes Then
                Exit Sub
            End If
            Dim frmASFMSGBF As New ASFMSGBF
            Dim STYLES_GROUP As Int64 = 0
            STYLES_GROUP = frmASFMSGBF.Get_numint_from_User("How Many Style / Group", "Groups", 1000, 200, 500)
            If STYLES_GROUP < 200 Or STYLES_GROUP > 1000 Then
                Exit Sub
            End If

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Creating Groups", "")
            Application.DoEvents()

            Dim STYLE_LIST As New List(Of String)
            For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select("STYLE_GROUP < 999", "STYLE_CODE, COLOR_CODE")
                Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
                If Not STYLE_LIST.Contains(STYLE_CODE) Then
                    STYLE_LIST.Add(STYLE_CODE)
                End If
            Next

            Dim THIS_GRP As Int64 = 1
            Dim GRP_CNT As Int64 = 0
            For Each STYLE As String In STYLE_LIST
                GRP_CNT += 1
                If GRP_CNT >= STYLES_GROUP Then
                    GRP_CNT = 1
                    THIS_GRP += 1
                End If
                ASCMAIN1.Progress("Now Creating Groups", THIS_GRP)
                For Each rowWBTSTYLD As DataRow In dst.Tables("WBTSTYLD").Select($"STYLE_CODE = '{STYLE}'", "COLOR_CODE")
                    If rowWBTSTYLD.Item("WEB_IND").ToString & String.Empty = "U" Then
                        rowWBTSTYLD.Item("WEB_IND") = "W"
                    End If
                    rowWBTSTYLD.Item("STYLE_GROUP") = THIS_GRP
                Next
            Next

            Update_Record(False)
            Application.DoEvents()
            MsgBox("Style(s) Updated.  Please Wait While Data Is Refreshed", vbOKOnly, "Complete")
            Clear_Record()
            Application.DoEvents()
            Load_Record()
            Application.DoEvents()
            Me.Cursor = Cursors.Default
        End If
    End Sub
#End Region

End Class