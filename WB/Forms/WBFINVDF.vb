Imports Infragistics.Win

Public Class WBFINVDF

    ' Used for automation 
    Private automatedExport As Boolean = False
    Private TASK_NO As String = String.Empty
    Private clsSOCMAIN1 As New TAC.SOCMAIN1

    Private PARTNER_CODE As String = String.Empty
    Private PARTNER_STATUS As String = String.Empty
    Private PARTNER_NAME As String = String.Empty
    Private PARTNER_ORDR_SOURCE_CODE As String = String.Empty
    Private PARTNER_SITE_IP As String = String.Empty
    Private PARTNER_SITE_USER As String = String.Empty
    Private PARTNER_SITE_PWD As String = String.Empty
    Private PARTNER_SITE_OUTPUT_DIR As String = String.Empty
    Private PARTNER_ORDERS_DIR As String = String.Empty
    Private PARTNER_LAST_SALES_ORDER As String = String.Empty
    Private PARTNER_SITE_ORDERS_POST_URL As String = String.Empty
    Private PARTNER_OUR_ID As String = String.Empty
    Private PARTNER_OUR_SITE_NAME As String = String.Empty
    Private PARTNER_SHIP_CONF_FILENAME As String = String.Empty
    Private PARTNER_SHIP_CONF_LOCAL_DIR As String = String.Empty
    Private PARTNER_SHIP_CONF_IP As String = String.Empty
    Private PARNTER_SHIP_CONF_USER As String = String.Empty
    Private PARTNER_SHIP_CONF_PASS As String = String.Empty
    Private PARTNER_SHIP_CONF_REMOTE_DIR As String = String.Empty
    Private PARTNER_PRODUCT_FILENAME As String = String.Empty
    Private PARTNER_PRODUCT_LOCAL_DIR As String = String.Empty
    Private PARTNER_PRODUCT_LAST_EXTRACT As String = String.Empty
    Private PARTNER_PRODUCT_IP As String = String.Empty
    Private PARTNER_PRODUCT_USER As String = String.Empty
    Private PARTNER_PRODUCT_PASS As String = String.Empty
    Private PARTNER_PRODUCT_REMOTE_DIR As String = String.Empty
    Private PARTNER_PRODUCT_OUR_ID As String = String.Empty
    Private PARTNER_PRODUCT_OUR_SUBID As String = String.Empty
    Private PARTNER_PRODUCT_INV_PCT As Double = 0
    Private PARTNER_PRODUCT_AID As String = String.Empty
    Private PARTNER_PRODUCT_SHIP_RATE As Double = 0
    Private PARTNER_PRODUCT_PROMO_TEXT As String = String.Empty
    Private PARTNER_PRODUCT_SHIP_RATE_GR As Double = 0
    Private PARTNER_PRODUCT_SHIP_RATE_2D As Double = 0
    Private PARTNER_PRODUCT_SHIP_RATE_ND As Double = 0
    Private PARTNER_PRODUCT_INV_MIN As Int16 = 0
    Private PARTNER_PRODUCT_ARCHIVE_DAYS As Int16 = 0
    Private SEND_ONCE_DAY As Boolean = False

    Private WHSE_ZIP_CODE As String = String.Empty

    Private WithEvents Ftp1 As New nsoftware.IPWorks.Ftp
    Private WithEvents Sftp1 As New nsoftware.IPWorksSSH.Sftp
    Private ftpFileList As List(Of String)
    Private itemsDatatable As DataTable

    Private Const imageDir As String = "http://www.webundies.com/images/"
    Private Const imageMediaDir As String = "http://www.webundies.com/media/products/"
    Private siteURL As String = String.Empty

    Private STYLE_CODE As String = String.Empty
    Private COLOR_CODE As String = String.Empty
    Private SIZE_CODE As String = String.Empty
    Private DEPT_CODE As String = String.Empty
    Private ITEM_TYPE_CODE As String = String.Empty
    Private PREVIOUS_STYLE_CODE As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        Get_PARM("WBTPARM1")

        With dst
            Create_TDA(.Tables.Add, "SOTPART1", "SELECT DECODE(PARTNER_STATUS, 'A', '1','0') SEL, '0' PROCESSED, SOTPART1.* FROM SOTPART1 WHERE PARTNER_PRODUCT_FILENAME IS NOT NULL", 0, False, "", 0)
            Create_TDA(.Tables.Add, "SOTPART2", "*", 1, False)
            Create_TDA(.Tables.Add, "SOTPART3", "*", 1, False)

            sql = " SELECT ICTSTYL2.STYLE_CODE, ICTSTYL2.FEATURE_CODE, ICTFEAT1.FEATURE_DESC "
            sql &= " FROM ICTSTYL2, ICTFEAT1"
            sql &= " WHERE ICTSTYL2.FEATURE_CODE = ICTFEAT1.FEATURE_CODE"
            Create_TDA(.Tables.Add, "ICTFEATX", sql, 0, False)

            Create_TDA(.Tables.Add, "ICTSTAT2", "*")

            Create_Lookup("ICTBRAN1")

        End With

        WHSE_ZIP_CODE = ASCDATA1.GetDataValue("SELECT WHSE_ZIP_CODE FROM ICTWHSE1 WHERE WHSE_CODE = '001'") & String.Empty
        grdSOTPART1.DataSource = dst.Tables("SOTPART1")
        siteURL = ROWs("WBTPARM1").Item("WB_PARM_SITE_NAME")

        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")
        Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")

    End Sub

    Private Sub SOFPBIMP_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Environment.GetCommandLineArgs.Count >= 5 AndAlso MENU_ITEM_OBJECT = "WBFINVDF" Then
            If Environment.GetCommandLineArgs.ElementAt(4) = MENU_ITEM_OBJECT Then
                Dim company As String = Environment.GetCommandLineArgs.ElementAt(1)

                If ASCMAIN1.DBS_COMPANY = company.ToUpper Then
                    Try
                        automatedExport = True
                        TASK_NO = clsSOCMAIN1.UpdateTask(String.Empty, MENU_ITEM_OBJECT)
                        MyBase.Click_Command("Load")
                        MyBase.Click_Command("Done")
                        TASK_NO = clsSOCMAIN1.UpdateTask(TASK_NO, MENU_ITEM_OBJECT)
                    Catch ex As Exception
                    Finally
                        Me.BeginInvoke(New MethodInvoker(AddressOf Me.Close))
                    End Try
                End If
            End If
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If dst.Tables("SOTPART1").Select("ISNULL(SEL, '0') = '1' AND ISNULL(PROCESSED, '0') = '0'").Length = 0 Then
                    EMsg &= vbCr & "No partners selected."
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)


        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        If tf Then
            grdSOTPART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdSOTPART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False

        dst.Tables("SOTPART2").Rows.Clear()
        dst.Tables("ICTFEATX").Rows.Clear()
        Fill_Records("SOTPART1")
        With grdSOTPART1.DisplayLayout.Bands(0).SortedColumns
            .Clear()
            .Add("PARTNER_NAME", False)
        End With

        dst.EnforceConstraints = True

        If dst.Tables.Contains("ASTSQLX1") Then
            dst.Tables("ASTSQLX1").Clear()
        End If

    End Sub

    Sub Load_Record()

        Try
            Me.Cursor = Cursors.WaitCursor
            Call ASCMAIN1.Progress("Now Loading Partner Data")

            grdSOTPART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

            MyBase.EnforceConstraints(False)
            Fill_Records("SOTPART2")
            MyBase.EnforceConstraints(True)

            Dim PARTNER_CODE As String = String.Empty
            Dim iRecCount As Int16 = 0
            Dim dataLine As String = String.Empty
            Dim setActiveMode As Boolean = False

            For Each rowSOTPART1 As DataRow In dst.Tables("SOTPART1").Select("ISNULL(SEL, '0') = '1' AND ISNULL(PROCESSED, '0') = '0'", "PARTNER_CODE")

                PARTNER_CODE = rowSOTPART1.Item("PARTNER_CODE")
                Call ASCMAIN1.Progress("Now Loading " & PARTNER_CODE & " Data", "0")
                iRecCount = 0
                dataLine = String.Empty

                STYLE_CODE = String.Empty
                COLOR_CODE = String.Empty
                SIZE_CODE = String.Empty
                DEPT_CODE = String.Empty
                ITEM_TYPE_CODE = String.Empty
                PREVIOUS_STYLE_CODE = String.Empty
                setActiveMode = False

                LoadPartnerParameters(PARTNER_CODE)
                Fill_Records("SOTPART3", PARTNER_CODE)

                If PARTNER_PRODUCT_FILENAME.Length = 0 Then
                    If automatedExport Then
                        clsSOCMAIN1.AddTaskDetail(TASK_NO, PARTNER_CODE & " DataFeed: No PARTNER_PRODUCT_FILENAME")
                    Else
                        MessageBox.Show(PARTNER_CODE & " DataFeed: No PARTNER_PRODUCT_FILENAME", "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Continue For
                End If

                If PARTNER_PRODUCT_LOCAL_DIR.Length = 0 Then
                    If automatedExport Then
                        clsSOCMAIN1.AddTaskDetail(TASK_NO, PARTNER_CODE & " DataFeed: No PARTNER_PRODUCT_LOCAL_DIR")
                    Else
                        MessageBox.Show(PARTNER_CODE & " DataFeed: No PARTNER_PRODUCT_LOCAL_DIR", "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Continue For
                ElseIf Not My.Computer.FileSystem.DirectoryExists(PARTNER_PRODUCT_LOCAL_DIR) Then
                    If automatedExport Then
                        clsSOCMAIN1.AddTaskDetail(TASK_NO, PARTNER_CODE & " DataFeed: Cannot locate directory, " & PARTNER_PRODUCT_LOCAL_DIR)
                    Else
                        MessageBox.Show(PARTNER_CODE & " DataFeed: Cannot locate directory, " & PARTNER_PRODUCT_LOCAL_DIR, "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Continue For
                ElseIf SEND_ONCE_DAY = True AndAlso IsDate(PARTNER_PRODUCT_LAST_EXTRACT) _
                    AndAlso DateTime.Now.ToString("yyyyMMdd") = CDate(PARTNER_PRODUCT_LAST_EXTRACT).ToString("yyyyMMdd") Then
                    If automatedExport Then
                        clsSOCMAIN1.AddTaskDetail(TASK_NO, PARTNER_CODE & " DataFeed: Can be sent once a day.")
                    Else
                        MessageBox.Show(PARTNER_CODE & " DataFeed: Can be sent once a day.", "Export Data", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                    Continue For
                End If

                If Not PARTNER_PRODUCT_LOCAL_DIR.EndsWith("\") Then
                    PARTNER_PRODUCT_LOCAL_DIR &= "\"
                End If

                ' Get Up to date Inventory and Features
                Fill_Records("ICTSTAT2", String.Empty, True, "SELECT * FROM ICTSTAT2")
                MyBase.Fill_Records("ICTFEATX")

                Select Case PARTNER_CODE

                    Case "COMMJUNC"
                        GetItemsToExport(False, PARTNER_CODE)
                        setActiveMode = False

                    Case "AMAZON"
                        If Not PARTNER_PRODUCT_REMOTE_DIR.EndsWith("\") Then PARTNER_PRODUCT_REMOTE_DIR &= "\"
                        If My.Computer.FileSystem.FileExists(PARTNER_PRODUCT_REMOTE_DIR & PARTNER_PRODUCT_FILENAME) Then
                            Continue For
                        End If
                        GetItemsToExport(False, PARTNER_CODE)
                        setActiveMode = False

                    Case "BUY"
                        GetItemsToExport(True, String.Empty)
                        setActiveMode = True

                    Case "BUYI"
                        GetItemsToExport(False, PARTNER_CODE)
                        setActiveMode = True

                    Case Else
                        GetItemsToExport(True, String.Empty)
                        setActiveMode = False
                End Select

                Using sw As New System.IO.StreamWriter(PARTNER_PRODUCT_LOCAL_DIR & PARTNER_PRODUCT_FILENAME, False)

                    For Each rowICTSTYL1 As DataRow In itemsDatatable.Select("", "STYLE_CODE, SIZE_SEQ_NO, COLOR_CODE")

                        iRecCount += 1
                        If iRecCount Mod 100 = 0 Then
                            ASCMAIN1.Progress("-", iRecCount)
                            Application.DoEvents()
                        End If

                        STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty
                        COLOR_CODE = rowICTSTYL1.Item("COLOR_CODE") & String.Empty
                        SIZE_CODE = rowICTSTYL1.Item("SIZE_CODE") & String.Empty
                        DEPT_CODE = rowICTSTYL1.Item("DEPT_CODE") & String.Empty
                        ITEM_TYPE_CODE = rowICTSTYL1.Item("ITEM_TYPE_CODE") & String.Empty

                        Select Case rowSOTPART1.Item("PARTNER_CODE")

                            Case "AMAZON"
                                dataLine = AmazonDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "BING"
                                dataLine = BingDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "BIZRATE"
                                dataLine = BizrateDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "BUY"
                                dataLine = BuyDotComFeed(rowICTSTYL1, iRecCount = 1)

                            Case "BUYI"
                                dataLine = BuyDotComInventoryFeed(rowICTSTYL1, iRecCount = 1)

                            Case "CATALOGS"
                                dataLine = CatalogsDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "COMMJUNC"
                                dataLine = CommJunctionDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "CPC"
                                dataLine = CPCDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "GOOGLE"
                                dataLine = GoogleDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "MSHOPPER"
                                dataLine = MShopperDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "NEXTAG"
                                dataLine = NextTagDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "PRICEGRAB"
                                dataLine = PriceGrabberDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "PRONTO"
                                dataLine = ProntoDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "SEARCHSPR"
                                dataLine = SearchSpringDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "SAS"
                                dataLine = ShareASaleFeed(rowICTSTYL1, iRecCount = 1)

                            Case "SHOP"
                                dataLine = ShopDotComDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "SHOPPING"
                                dataLine = ShoppingDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "SLI"
                                dataLine = SLIDataFeed(rowICTSTYL1, iRecCount = 1)

                            Case "THEFIND"
                                dataLine = TheFindDataFeed(rowICTSTYL1, iRecCount = 1)

                        End Select

                        PREVIOUS_STYLE_CODE = STYLE_CODE

                        If dataLine.Length > 0 Then
                            sw.WriteLine(dataLine)
                        End If
                    Next

                    sw.Close()

                End Using

                Dim ftpSuccessful As Boolean = True

                If Not (PARTNER_PRODUCT_USER.Length = 0 OrElse PARTNER_PRODUCT_PASS.Length = 0 OrElse PARTNER_PRODUCT_IP.Length = 0) AndAlso chkTestMode.Checked = False Then
                    ftpSuccessful = ftpData(PARTNER_PRODUCT_LOCAL_DIR & PARTNER_PRODUCT_FILENAME, setActiveMode)
                ElseIf PARTNER_PRODUCT_REMOTE_DIR.Length > 0 AndAlso My.Computer.FileSystem.DirectoryExists(PARTNER_PRODUCT_REMOTE_DIR) AndAlso chkTestMode.Checked = False Then
                    If Not PARTNER_PRODUCT_REMOTE_DIR.EndsWith("\") Then PARTNER_PRODUCT_REMOTE_DIR &= "\"
                    My.Computer.FileSystem.CopyFile(PARTNER_PRODUCT_LOCAL_DIR & PARTNER_PRODUCT_FILENAME, PARTNER_PRODUCT_REMOTE_DIR & PARTNER_PRODUCT_FILENAME)
                End If

                If Not chkTestMode.Checked AndAlso ftpSuccessful Then
                    My.Computer.FileSystem.MoveFile(PARTNER_PRODUCT_LOCAL_DIR & PARTNER_PRODUCT_FILENAME, PARTNER_PRODUCT_LOCAL_DIR & "PROCESSED\" & DateTime.Now.ToString("yyyyMMdd_hhmmss") & "_" & PARTNER_PRODUCT_FILENAME)
                    UpdatePartner(PARTNER_CODE)
                End If

                ' Clean Out Archive Log
                Try
                    If PARTNER_PRODUCT_ARCHIVE_DAYS > 0 Then
                        For Each file As String In My.Computer.FileSystem.GetFiles(PARTNER_PRODUCT_LOCAL_DIR & "PROCESSED\")
                            Dim fFile As New System.IO.FileInfo(file)
                            If DateDiff(DateInterval.Day, fFile.CreationTime, DateTime.Now) >= PARTNER_PRODUCT_ARCHIVE_DAYS Then
                                fFile.Delete()
                            End If
                        Next
                    End If
                Catch ex As Exception
                    ' Nothing
                End Try
            Next

        Catch ex As Exception

            If automatedExport Then
                clsSOCMAIN1.AddTaskDetail(TASK_NO, "Error creating " & PARTNER_CODE & " DataFeed: " & ex.Message)
            Else
                MessageBox.Show("Error creating " & PARTNER_CODE & " DataFeed: " & ex.Message, "Error", MessageBoxButtons.OK)
            End If

        Finally

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress(String.Empty, String.Empty)
            Me.Cursor = Cursors.Default
            EmailErrors()
        End Try
    End Sub

#End Region

#Region "Form Procedures"

    Private Function AmazonDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim parentRecord As Int16 = 1
        Dim childRecord As Int16 = 2
        Dim rowSOTPART3 As DataRow = Nothing
        Dim rowICTBRAN1 As DataRow = Nothing

        If includeHeader Then

            Dim firstLine As String = "TemplateType=ClothingAccessories,Version=1.4,This row for Amazon.com use only.  Do not modify or delete.,,,,,,,,,Offer Information - These attributes are required to make your item buyable for customers on the site.,,,,,,,,,,,Sales Price information,,,Item discovery information - Affects how customers can find your product on the site.,,,,,,Image Information - see Image Info tab for details.,,,,,,,,,,FBA - make use of these columns if you are participating in the Fulfillment by Amazon program,,,,,,,,Variation information,,,,ClothingAccories Product Information - these attributes are specific to certain product types.  Please use associated Valid Values for more detail.,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,Product Dimensions,,,,,,Common infrequently used attributes ,,,,,,,,,"
            Dim secondLine As String = "sku,product-name,product-id,product-id-type,brand,product-description,bullet-point1,bullet-point2,bullet-point3,bullet-point4,bullet-point5,item-price,msrp,currency,product-tax-code,shipping-weight,shipping-weight-unit-measure,leadtime-to-ship,launch-date,release-date,restock-date,quantity,sale-price,sale-from-date,sale-through-date,search-terms1,search-terms2,search-terms3,search-terms4,search-terms5,item-type,main-image-url,other-image-url1,other-image-url2,other-image-url3,other-image-url4,other-image-url5,other-image-url6,other-image-url7,other-image-url8,swatch-image-url,fulfillment-center-id,package-height,package-width,package-length,package-length-unit-of-measure,package-weight,package-weight-unit-of-measure,max-aggregate-ship-quantity,parent-child,parent-sku,relationship-type,variation-theme,apparel-closure-type,belt-style,bottom-style,button-quantity,character,chest-size,chest-size-unit-of-measure,collar-type,color,color-map,control-type,cpsia-warning1,cpsia-warning2,cpsia-warning3,cpsia-warning4,cpsia-warning-description,cuff-type,cup-size,department,fabric-wash,fit-type,front-pleat-type,inseam-length,inseam-length-unit-of-measure,is-stain-resistant,item-package-quantity,item-rise,item-rise-unit-of-measure,laptop-capacity,leg-diameter,leg-diameter-unit-of-measure,leg-style,material-fabric1,material-fabric2,material-fabric3,material-opacity,neck-size,neck-size-unit-of-measure,neck-style,number-of-items,number-of-pieces,occasion-lifestyle,pattern-style,pocket-description,rise-style,size,size-map,size-modifier,sleeve-length,sleeve-length-unit-of-measure,sleeve-type,special-feature1,strap-type,style-name,theme,toe-style,top-style,underwire-type,waist-size,waist-size-unit-of-measure,water-resistance-level,wheel-type,item-weight-unit-of-measure,item-weight,item-length-unit-of-measure,item-length,item-width,item-height,is-gift-message-available,is-gift-wrap-available,is-discontinued-by-manufacturer,registered-parameter,platinum-keywords1,platinum-keywords2,platinum-keywords3,platinum-keywords4,platinum-keywords5,update-delete"
            For Each dataString As String In New String() {firstLine, secondLine}
                For Each colHeader As String In dataString.Split(",")
                    dataLine &= colHeader & vbTab
                Next
                dataLine = dataLine.Substring(0, dataLine.Length - 1)
                dataLine &= Environment.NewLine
            Next

            dataLine = dataLine.Replace("Fulfillment by Amazon", Chr(34) & "Fulfillment by Amazon" & Chr(34))
        End If

        ' Uses A Style Header Record for each Style
        Dim numLoops As Int16 = 2
        If PREVIOUS_STYLE_CODE <> rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            numLoops = 1
        End If

        For iloop As Int16 = numLoops To 2

            If iloop = parentRecord Then
                dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab ' sku 
                dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' product-name 
            Else
                dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' sku 
                dataLine &= rowICTSTYL1.Item("STYLE_DESC_XML") & vbTab ' product-name 
            End If

            If iloop = parentRecord Then
                dataLine &= vbTab ' product-id
                dataLine &= vbTab ' product-id-type
            Else
                Select Case (rowICTSTYL1.Item("ITEM_GTIN") & String.Empty).ToString.Length
                    Case 12
                        dataLine &= rowICTSTYL1.Item("ITEM_GTIN") & String.Empty & vbTab ' product-id
                        dataLine &= "UPC" & vbTab ' product-id-type
                    Case 13
                        dataLine &= rowICTSTYL1.Item("ITEM_GTIN") & String.Empty & vbTab ' product-id
                        dataLine &= "EAN" & vbTab ' product-id-type
                    Case Else
                        dataLine &= vbTab ' product-id
                        dataLine &= vbTab ' product-id-type
                End Select
            End If

            rowICTBRAN1 = LookUp("ICTBRAN1", rowICTSTYL1.Item("BRAND_CODE") & String.Empty)

            If rowICTBRAN1 IsNot Nothing AndAlso (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim.Length > 0 Then
                dataLine &= (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim & vbTab ' brand
            Else
                dataLine &= "WebUndies" & vbTab ' brand
            End If

            If iloop = parentRecord Then
                dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' product-description 
            Else
                dataLine &= vbTab
            End If

            dataLine &= GetCareInstructions(rowICTSTYL1) & vbTab ' bullet-point1
            dataLine &= rowICTSTYL1.Item("SOURCE_DESC") & vbTab ' bullet-point2
            dataLine &= GetFeatures(STYLE_CODE) & vbTab ' bullet-point3
            dataLine &= vbTab ' bullet-point4
            dataLine &= vbTab ' bullet-point5

            If iloop = childRecord Then
                dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab ' item-price
                dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab ' msrp
            Else
                dataLine &= vbTab ' item-price
                dataLine &= vbTab ' msrp
            End If

            dataLine &= "USD" & vbTab ' currency

            If iloop = childRecord Then
                dataLine &= "G_GEN_NOTAX" & vbTab ' product-tax-code
            Else
                dataLine &= vbTab ' product-tax-code
            End If

            dataLine &= vbTab ' shipping-weight
            dataLine &= vbTab ' shipping-weight-unit-measure

            ' leadtime-to-ship
            If iloop = childRecord Then
                dataLine &= "1" & vbTab
            Else
                dataLine &= vbTab
            End If

            dataLine &= vbTab ' launch-date
            dataLine &= vbTab ' release-date
            dataLine &= vbTab ' restock-date

            If iloop = childRecord Then
                If PARTNER_PRODUCT_INV_PCT <= 0 OrElse PARTNER_PRODUCT_INV_PCT > 100 Then PARTNER_PRODUCT_INV_PCT = 100
                dataLine &= Math.Floor((GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) * (PARTNER_PRODUCT_INV_PCT / 100))) & vbTab  ' quantity
            Else
                dataLine &= vbTab ' quantity  
            End If

            '            When sending Sale Price, we need to send start/end date

            'When sending Sale-price
            'Set Sale-From-Date to Current date
            'Set Sale-Through-Date to Current date + 30

            'Dates should be formatted as : 2010-07-10


            ' sale-price	
            If "SC".Contains(rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty) AndAlso iloop = childRecord Then
                dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab
                dataLine &= DateTime.Now.ToString("yyyy-MM-dd") & vbTab ' sale-from-date
                dataLine &= DateAdd(DateInterval.Day, 30, DateTime.Now).ToString("yyyy-MM-dd") & vbTab ' sale-through-date
            Else
                dataLine &= vbTab
                dataLine &= vbTab ' sale-from-date
                dataLine &= vbTab ' sale-through-date
            End If

            ' search-terms1, search-terms2, search-terms3, search-terms4, search-terms5
            Dim numSearch As Int16 = 0
            For Each searchTerm As String In (rowICTSTYL1("STYLE_ADDL_KEYWORDS") & String.Empty).ToString.Trim.Split(",")
                If searchTerm.Trim.Length > 0 Then
                    dataLine &= searchTerm.Trim & vbTab
                    numSearch += 1
                End If
                If numSearch >= 5 Then Exit For
            Next
            While numSearch < 5
                dataLine &= vbTab
                numSearch += 1
            End While

            dataLine &= GetPartnerCategory("AMAZON", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' item-type
            dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' main-image-url

            ' other-image-url1, other-image-url2
            For Each altImage As String In New String() {"STYLE_IMAGE_OTHER1", "STYLE_IMAGE_OTHER2"}
                If rowICTSTYL1.Item(altImage) & String.Empty <> String.Empty Then
                    dataLine &= imageMediaDir & rowICTSTYL1.Item(altImage).ToString.ToLower & vbTab ' URL
                Else
                    dataLine &= vbTab
                End If
            Next

            dataLine &= vbTab ' other-image-url3
            dataLine &= vbTab ' other-image-url4
            dataLine &= vbTab ' other-image-url5
            dataLine &= vbTab ' other-image-url6
            dataLine &= vbTab ' other-image-url7
            dataLine &= vbTab ' other-image-url8
            dataLine &= vbTab ' swatch-image-url
            dataLine &= vbTab ' fulfillment-center-id	
            dataLine &= vbTab ' package-height
            dataLine &= vbTab ' package-width
            dataLine &= vbTab ' package-length
            dataLine &= vbTab ' package-length-unit-of-measure	
            dataLine &= vbTab ' package-weight
            dataLine &= vbTab ' package-weight-unit-of-measure
            dataLine &= vbTab ' max-aggregate-ship-quantity

            If iloop = parentRecord Then
                dataLine &= "Parent" & vbTab ' parent-child	 
                dataLine &= vbTab ' parent-sku	
                dataLine &= vbTab ' relationship-type
            Else
                dataLine &= "Child" & vbTab ' parent-child	 
                dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab ' parent-sku	
                dataLine &= "Variation" & vbTab ' relationship-type	 
            End If

            dataLine &= "Size" & vbTab ' variation-theme
            dataLine &= vbTab ' apparel-closure-type
            dataLine &= vbTab ' belt-style
            dataLine &= vbTab ' bottom-style
            dataLine &= vbTab ' button-quantity
            dataLine &= vbTab ' character
            dataLine &= vbTab ' chest-size
            dataLine &= vbTab ' chest-size-unit-of-measure
            dataLine &= vbTab ' collar-type
            'dataLine &= "Multi" & vbTab ' color
            dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' color
            'dataLine &= "Multicoloured" & vbTab ' color-map
            dataLine &= rowICTSTYL1.Item("STYLE_COLOR_STANDARD") & String.Empty & vbTab ' color-map
            dataLine &= vbTab ' control-type
            dataLine &= vbTab ' cpsia-warning1
            dataLine &= vbTab ' cpsia-warning2
            dataLine &= vbTab ' cpsia-warning3
            dataLine &= vbTab ' cpsia-warning4
            dataLine &= vbTab ' cpsia-warning-description
            dataLine &= vbTab ' cuff-type
            dataLine &= vbTab ' cup-size
            dataLine &= rowICTSTYL1.Item("DEPT_DESC") & vbTab ' department
            dataLine &= vbTab ' fabric-wash
            dataLine &= vbTab ' fit-type
            dataLine &= vbTab ' front-pleat-type
            dataLine &= vbTab ' inseam-length
            dataLine &= vbTab ' inseam-length-unit-of-measure
            dataLine &= vbTab ' is-stain-resistant
            dataLine &= vbTab ' item-package-quantity
            dataLine &= vbTab ' item-rise
            dataLine &= vbTab ' item-rise-unit-of-measure
            dataLine &= vbTab ' laptop-capacity
            dataLine &= vbTab ' leg-diameter
            dataLine &= vbTab ' leg-diameter-unit-of-measure
            dataLine &= vbTab ' leg-style

            dataLine &= rowICTSTYL1.Item("MATL_DESC") & vbTab  ' material-fabric1
            dataLine &= vbTab ' material-fabric2
            dataLine &= vbTab ' material-fabric3
            dataLine &= vbTab ' material-opacity

            dataLine &= vbTab ' neck-size
            dataLine &= vbTab ' neck-size-unit-of-measure
            dataLine &= vbTab ' neck-style
            dataLine &= vbTab ' number-of-items
            dataLine &= vbTab ' number-of-pieces
            dataLine &= vbTab ' occasion-lifestyle
            dataLine &= vbTab ' pattern-style
            dataLine &= vbTab ' pocket-description
            dataLine &= vbTab ' rise-style

            If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
                rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
            End If

            If iloop = parentRecord Then
                dataLine &= vbTab ' size
                dataLine &= vbTab ' size-map
                dataLine &= vbTab ' size-modifier
            ElseIf rowSOTPART3 IsNot Nothing Then
                dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab ' size
                dataLine &= vbTab ' size-map
                dataLine &= rowSOTPART3.Item("PARTNER_SIZE_MODIFIER") & vbTab ' size-modifier
            Else
                dataLine &= rowICTSTYL1.Item("SIZE_CODE") & String.Empty & vbTab ' size
                dataLine &= vbTab ' size-map
                dataLine &= vbTab ' size-modifier
            End If

            dataLine &= vbTab ' sleeve-length
            dataLine &= vbTab ' sleeve-length-unit-of-measure
            dataLine &= vbTab ' sleeve-type
            dataLine &= vbTab ' special-feature1
            dataLine &= vbTab ' strap-type
            dataLine &= vbTab ' style-name
            dataLine &= vbTab ' theme
            dataLine &= vbTab ' toe-style
            dataLine &= vbTab ' top-style
            dataLine &= vbTab ' underwire-type
            dataLine &= vbTab ' waist-size
            dataLine &= vbTab ' waist-size-unit-of-measure
            dataLine &= vbTab ' water-resistance-level
            dataLine &= vbTab ' wheel-type
            dataLine &= vbTab ' item-weight-unit-of-measure
            dataLine &= vbTab ' item-weight
            dataLine &= vbTab ' item-length-unit-of-measure
            dataLine &= vbTab ' item-length
            dataLine &= vbTab ' item-width
            dataLine &= vbTab ' item-height
            dataLine &= vbTab ' is-gift-message-available
            dataLine &= vbTab ' is-gift-wrap-available
            dataLine &= vbTab ' is-discontinued-by-manufacturer
            dataLine &= "PrivateLabel" & vbTab ' registered-parameter

            If rowICTSTYL1.Item("STYLE_KEYWORD_1") & String.Empty <> String.Empty Then
                dataLine &= rowICTSTYL1.Item("STYLE_KEYWORD_1") & vbTab ' platinum-keywords1
            Else
                dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' platinum-keywords1
            End If
            dataLine &= rowICTSTYL1.Item("STYLE_KEYWORD_2") & vbTab ' platinum-keywords2
            dataLine &= rowICTSTYL1.Item("STYLE_KEYWORD_3") & vbTab ' platinum-keywords3	
            dataLine &= vbTab ' platinum-keywords4	
            dataLine &= vbTab ' platinum-keywords5	

            ' update-delete
            If rowICTSTYL1.Item("STYLE_STATUS") <> "A" Then
                dataLine &= "Delete"
            ElseIf DateTime.Compare(rowICTSTYL1.Item("INIT_DATE"), CDate(PARTNER_PRODUCT_LAST_EXTRACT)) > 0 OrElse iloop = parentRecord Then
                dataLine &= "Update"
            ElseIf rowICTSTYL1.Item("ITEM_STATUS") <> "A" Then
                dataLine &= "Delete"
            Else
                dataLine &= "Update"
            End If

            If iloop = parentRecord Then
                dataLine &= Environment.NewLine
            End If
        Next

        Return dataLine
    End Function

    Private Function BingDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine = "MerchantProductID" & vbTab
            dataLine &= "Title" & vbTab
            dataLine &= "ImageURL" & vbTab
            dataLine &= "ProductURL" & vbTab
            dataLine &= "Price" & vbTab
            dataLine &= "Description" & vbTab
            dataLine &= "Brand" & vbTab
            dataLine &= "Availability" & vbTab
            dataLine &= "Shipping" & vbTab
            dataLine &= "Condition" & vbTab
            dataLine &= "MerchantCategory" & vbTab
            dataLine &= "B_Category" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        ' MerchantProductID = STYLE_CODE
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab
        'Title = ICTSTYL1.STYLE_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab
        'ImageURL(IMAGE_URL)
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        'ProductURL(PRODUCT_URL)
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        'Price = STYLE_PRICE_TYPE = "C" or "S" use STYLE_PRICE_SALE else use STYLE_PRICE_OUR
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        'Description = ICTSTYL1.STYLE_FULL_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' Description
        'Brand = "WebUndies"
        dataLine &= "WebUndies" & vbTab
        'Availability = "In Stock"
        dataLine &= "In Stock" & vbTab
        'Shipping = SOTPART1.PARTNER_PRODUCT_SHIP_RATE
        dataLine &= formatNumber(Val(GetPartnerAttribute("BING", "PARTNER_PRODUCT_SHIP_RATE") & String.Empty)) & vbTab
        'Condition = "New"
        dataLine &= "New" & vbTab
        'MerchantCategory = SOTPART2.PARTNER_CATEGORY
        dataLine &= GetPartnerCategory("BING", DEPT_CODE, ITEM_TYPE_CODE) & vbTab
        'B_Category = "Clothing & Shoes"
        dataLine &= "Clothing & Shoes"

        Return dataLine

    End Function

    Private Function BizrateDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty

        If includeHeader Then
            dataLine = "Category" & vbTab
            dataLine &= "Manufacturer" & vbTab
            dataLine &= "Title" & vbTab
            dataLine &= "Description" & vbTab
            dataLine &= "Link" & vbTab
            dataLine &= "Image" & vbTab
            dataLine &= "SKU" & vbTab
            dataLine &= "Quantity on Hand" & vbTab
            dataLine &= "Condition" & vbTab
            dataLine &= "Shipping Weight (In Pounds, Not Required)" & vbTab
            dataLine &= "Shipping Cost (Not Required)" & vbTab
            dataLine &= "Bid (Not Required)" & vbTab
            dataLine &= "Promo Text (Not Required)" & vbTab
            dataLine &= "Other (Not Required)" & vbTab
            dataLine &= "Price" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If


        dataLine &= GetPartnerCategory("BIZRATE", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' Category
        dataLine &= "WebUndies" & vbTab ' Manufacturer
        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_DESC_EXT") & Chr(34) & vbTab ' Title
        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_FULL_DESC") & Chr(34) & vbTab ' Description
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' Image
        dataLine &= STYLE_CODE & vbTab ' SKU
        dataLine &= "In Stock" & vbTab ' Quantity on Hand
        dataLine &= "New" & vbTab ' Condition
        dataLine &= vbTab ' Shipping Weight
        dataLine &= vbTab ' Shipping Cost
        dataLine &= vbTab ' Bid
        dataLine &= vbTab ' Promo Text
        dataLine &= vbTab ' Other
        dataLine &= GetItemPrice(rowICTSTYL1, "C") ' Price

        Return (dataLine)

    End Function

    Private Function BuyDotComFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine = "seller-id" & vbTab
            dataLine &= "gtin" & vbTab
            dataLine &= "isbn" & vbTab
            dataLine &= "mfg-name" & vbTab
            dataLine &= "mfg-part-number" & vbTab
            dataLine &= "asin" & vbTab
            dataLine &= "seller-sku" & vbTab
            dataLine &= "title" & vbTab
            dataLine &= "description" & vbTab
            dataLine &= "main-image" & vbTab
            dataLine &= "additional-images" & vbTab
            dataLine &= "weight" & vbTab
            dataLine &= "features" & vbTab
            dataLine &= "listing-price" & vbTab
            dataLine &= "msrp" & vbTab
            dataLine &= "category-id" & vbTab
            dataLine &= "keywords" & vbTab
            dataLine &= "product-set-id" & vbTab
            'dataLine &= "Age Range" & vbTab
            dataLine &= "Age Segment" & vbTab
            'dataLine &= "Apparel Material" & vbTab
            'dataLine &= "Apparel Occasion" & vbTab
            'dataLine &= "Assembled Dimension Height (in)" & vbTab
            'dataLine &= "Assembled Dimension Length (in)" & vbTab
            'dataLine &= "Assembled Dimension Width (in)" & vbTab
            'dataLine &= "Assembled Weight (lbs)" & vbTab
            'dataLine &= "Belt Type" & vbTab
            'dataLine &= "Bra Type" & vbTab
            'dataLine &= "Brassiere Cup Size (US)" & vbTab
            'dataLine &= "Brassiere Size (US)" & vbTab
            'dataLine &= "Casual" & vbTab
            'dataLine &= "Character" & vbTab
            'dataLine &= "Character / Series Type" & vbTab
            'dataLine &= "Closure Type" & vbTab
            dataLine &= "Color" & vbTab
            'dataLine &= "Color Class" & vbTab
            'dataLine &= "Cover Up" & vbTab
            'dataLine &= "Dance Type" & vbTab
            'dataLine &= "Dress Type" & vbTab
            'dataLine &= "Dress/Skirt Length" & vbTab
            'dataLine &= "Eyewear Type" & vbTab
            'dataLine &= "Footwear Size (Child US)" & vbTab
            'dataLine &= "Footwear Size (Female US)" & vbTab
            'dataLine &= "Footwear Size (Male US)" & vbTab
            'dataLine &= "Footwear Size Code" & vbTab
            'dataLine &= "Footwear Type" & vbTab
            'dataLine &= "Footwear Width" & vbTab
            'dataLine &= "Footwear Width Class" & vbTab
            'dataLine &= "Footwear Width Code" & vbTab
            'dataLine &= "Formal" & vbTab
            dataLine &= "Gender" & vbTab
            'dataLine &= "Gloves Type" & vbTab
            'dataLine &= "Hat & Headwear Type" & vbTab
            'dataLine &= "Heel Height" & vbTab
            'dataLine &= "Inseam Length" & vbTab
            'dataLine &= "Jacket Type" & vbTab
            'dataLine &= "Lace" & vbTab
            'dataLine &= "Legal Notice" & vbTab
            'dataLine &= "Lingerie" & vbTab
            'dataLine &= "Manufacturer Suggested Age" & vbTab
            'dataLine &= "Manufacturer Suggested Age Max" & vbTab
            'dataLine &= "Metallic" & vbTab
            'dataLine &= "MLB Team" & vbTab
            'dataLine &= "MLS Team" & vbTab
            'dataLine &= "NASCAR Driver" & vbTab
            'dataLine &= "NBA Team" & vbTab
            'dataLine &= "NCAA Team" & vbTab
            'dataLine &= "Neck Size" & vbTab
            'dataLine &= "Neck Type" & vbTab
            'dataLine &= "NFL Team" & vbTab
            'dataLine &= "NHL Team" & vbTab
            'dataLine &= "Pants Style" & vbTab
            'dataLine &= "Pants Type" & vbTab
            'dataLine &= "Pantyhose Type" & vbTab
            'dataLine &= "Pattern" & vbTab
            'dataLine &= "Raingear" & vbTab
            'dataLine &= "Reading Glasses Power" & vbTab
            'dataLine &= "Season" & vbTab
            'dataLine &= "Shapewear" & vbTab
            'dataLine &= "Shirt Type" & vbTab
            'dataLine &= "Shoe Lace Length (inches)" & vbTab
            'dataLine &= "Shorts Type" & vbTab
            dataLine &= "Size" & vbTab
            'dataLine &= "Size (Boys US)" & vbTab
            'dataLine &= "Size (Children US)" & vbTab
            'dataLine &= "Size (Girls US)" & vbTab
            'dataLine &= "Size (Infant US)" & vbTab
            'dataLine &= "Size (Junior US)" & vbTab
            'dataLine &= "Size (Men US)" & vbTab
            'dataLine &= "Size (Misses US)" & vbTab
            'dataLine &= "Size (Toddler US)" & vbTab
            'dataLine &= "Size (Women US)" & vbTab
            'dataLine &= "Size Code (Child US)" & vbTab
            'dataLine &= "Size Code (Men US)" & vbTab
            'dataLine &= "Size Code (Women US)" & vbTab
            'dataLine &= "Size Modifier" & vbTab
            'dataLine &= "Sleeve Length" & vbTab
            'dataLine &= "Sleeve Type" & vbTab
            'dataLine &= "Slip Type" & vbTab
            'dataLine &= "Sock Length" & vbTab
            'dataLine &= "Sports" & vbTab
            'dataLine &= "Sports League" & vbTab
            'dataLine &= "Stain Resistant" & vbTab
            'dataLine &= "Strap Type" & vbTab
            'dataLine &= "Sweathshirt Type" & vbTab
            'dataLine &= "Swimwear Type" & vbTab
            'dataLine &= "Therapeutic Type" & vbTab
            'dataLine &= "Tights Type" & vbTab
            'dataLine &= "Toe Type" & vbTab
            'dataLine &= "Top Length" & vbTab
            'dataLine &= "Trend/Style" & vbTab
            dataLine &= "Underwear Type" & Environment.NewLine
            'dataLine &= "Waist Size" & vbTab
            'dataLine &= "Wallet Type" & vbTab
            'dataLine &= "Waterproof" & vbTab
            'dataLine &= "Wet Suit Type" & vbTab
            'dataLine &= "WNBA Team" & vbTab
            'dataLine &= "Wrinkle Resistant" 
        End If

        dataLine &= PARTNER_PRODUCT_OUR_ID & vbTab ' seller-id
        dataLine &= rowICTSTYL1.Item("ITEM_GTIN") & vbTab ' gtin
        dataLine &= "" & vbTab ' isbn

        'dataLine &= "" & vbTab ' mfg-name
        Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", rowICTSTYL1.Item("BRAND_CODE") & String.Empty)
        If rowICTBRAN1 IsNot Nothing AndAlso (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim.Length > 0 Then
            dataLine &= (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim & vbTab
        Else
            dataLine &= "WebUndies" & vbTab
        End If

        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' mfg-part-number
        dataLine &= "" & vbTab ' asin
        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' seller-sku
        dataLine &= rowICTSTYL1.Item("STYLE_DESC") & vbTab ' title
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' description
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' main-image
        If (rowICTSTYL1.Item("STYLE_IMAGE_OTHER1") & String.Empty).ToString.Trim.Length > 0 Then
            dataLine &= imageMediaDir & rowICTSTYL1.Item("STYLE_IMAGE_OTHER1") & vbTab ' additional-images
        Else
            dataLine &= "" & vbTab ' additional-images
        End If
        dataLine &= IIf(Val(rowICTSTYL1.Item("ITEM_WEIGHT") & String.Empty) > 0, Val(rowICTSTYL1.Item("ITEM_WEIGHT") & String.Empty), 0.5) & vbTab ' weight
        dataLine &= GetFeatures(STYLE_CODE, "|") & vbTab ' features
        dataLine &= GetItemPrice(rowICTSTYL1, "") & vbTab ' listing-price
        dataLine &= "" & vbTab ' msrp
        dataLine &= GetPartnerCategoryId("BUY", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' category-id

        ' Keyword are comma delimeted
        Dim keywords As String = (rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") & String.Empty).ToString
        Dim keywordsArray() As String = keywords.Split(",")
        keywords = String.Empty
        For Each keywordExpr As String In keywordsArray
            keywordExpr = keywordExpr.Trim
            If keywordExpr.Length = 0 Then Continue For
            ' convert pipes to a space and convert double space to single space
            keywordExpr = keywordExpr.Replace("|", " ").Replace("  ", " ")
            If keywordExpr.Length > 40 Then
                keywordExpr = keywordExpr.Substring(0, 40).Trim
            End If
            ' Buy dot com delimits keywords with a pipe
            If keywords.Length + keywordExpr.Length <= 250 Then
                keywords &= keywordExpr & "|"
            End If
        Next
        If keywords.EndsWith("|") Then
            keywords = keywords.Substring(0, keywords.Length - 1)
        End If

        dataLine &= keywords & vbTab ' keywords
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab ' product-set-id
        'dataLine &= "" & vbTab ' Age Range
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper ' Age Segment
            Case "MENS", "WOMENS", "UNISEX"
                dataLine &= "Adult" & vbTab
            Case "BOYS", "GIRLS", "CHILDREN"
                Select Case rowICTSTYL1.Item("SIZE_CODE") & String.Empty
                    Case "1218M", "12M", "18M", "24M"
                        dataLine &= "Infant" & vbTab
                    Case "2T/3T", "2T", "3T", "4T", "5T"
                        dataLine &= "Toddler" & vbTab
                    Case Else
                        dataLine &= "Child" & vbTab
                End Select

            Case Else
                dataLine &= "" & vbTab
        End Select

        'dataLine &= "" & vbTab ' Apparel Material
        'dataLine &= "" & vbTab ' Apparel Occasion
        'dataLine &= "" & vbTab ' Assembled Dimension Height (in)
        'dataLine &= "" & vbTab ' Assembled Dimension Length (in)
        'dataLine &= "" & vbTab ' Assembled Dimension Width (in)
        'dataLine &= "" & vbTab ' Assembled Weight (lbs)
        'dataLine &= "" & vbTab ' Belt Type
        'dataLine &= "" & vbTab ' Bra Type
        'dataLine &= "" & vbTab ' Brassiere Cup Size (US)
        'dataLine &= "" & vbTab ' Brassiere Size (US)
        'dataLine &= "" & vbTab ' Casual
        'dataLine &= "" & vbTab ' Character
        'dataLine &= "" & vbTab ' Character / Series Type
        'dataLine &= "" & vbTab ' Closure Type
        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' Color
        'dataLine &= "" & vbTab ' Color Class
        'dataLine &= "" & vbTab ' Cover Up
        'dataLine &= "" & vbTab ' Dance Type
        'dataLine &= "" & vbTab ' Dress Type
        'dataLine &= "" & vbTab ' Dress/Skirt Length
        'dataLine &= "" & vbTab ' Eyewear Type
        'dataLine &= "" & vbTab ' Footwear Size (Child US)
        'dataLine &= "" & vbTab ' Footwear Size (Female US)
        'dataLine &= "" & vbTab ' Footwear Size (Male US)
        'dataLine &= "" & vbTab ' Footwear Size Code
        'dataLine &= "" & vbTab ' Footwear Type
        'dataLine &= "" & vbTab ' Footwear Width
        'dataLine &= "" & vbTab ' Footwear Width Class
        'dataLine &= "" & vbTab ' Footwear Width Code
        'dataLine &= "" & vbTab ' Formal
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper ' Gender
            Case "MENS", "BOYS"
                dataLine &= "Male" & vbTab
            Case "WOMENS", "GIRLS"
                dataLine &= "Female" & vbTab
            Case Else ' Currently all others go as male
                dataLine &= "Male" & vbTab
        End Select

        'dataLine &= "" & vbTab ' Gloves Type
        'dataLine &= "" & vbTab ' Hat & Headwear Type
        'dataLine &= "" & vbTab ' Heel Height
        'dataLine &= "" & vbTab ' Inseam Length
        'dataLine &= "" & vbTab ' Jacket Type
        'dataLine &= "" & vbTab ' Lace
        'dataLine &= "" & vbTab ' Legal Notice
        'dataLine &= "" & vbTab ' Lingerie
        'dataLine &= "" & vbTab ' Manufacturer Suggested Age
        'dataLine &= "" & vbTab ' Manufacturer Suggested Age Max
        'dataLine &= "" & vbTab ' Metallic
        'dataLine &= "" & vbTab ' MLB Team
        'dataLine &= "" & vbTab ' MLS Team
        'dataLine &= "" & vbTab ' NASCAR Driver
        'dataLine &= "" & vbTab ' NBA Team
        'dataLine &= "" & vbTab ' NCAA Team
        'dataLine &= "" & vbTab ' Neck Size
        'dataLine &= "" & vbTab ' Neck Type
        'dataLine &= "" & vbTab ' NFL Team
        'dataLine &= "" & vbTab ' NHL Team
        'dataLine &= "" & vbTab ' Pants Style
        'dataLine &= "" & vbTab ' Pants Type
        'dataLine &= "" & vbTab ' Pantyhose Type
        'dataLine &= "" & vbTab ' Pattern
        'dataLine &= "" & vbTab ' Raingear
        'dataLine &= "" & vbTab ' Reading Glasses Power
        'dataLine &= "" & vbTab ' Season
        'dataLine &= "" & vbTab ' Shapewear
        'dataLine &= "" & vbTab ' Shirt Type
        'dataLine &= "" & vbTab ' Shoe Lace Length (inches)
        'dataLine &= "" & vbTab ' Shorts Type
        dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab ' Size
        'dataLine &= "" & vbTab ' Size (Boys US)
        'dataLine &= "" & vbTab ' Size (Children US)
        'dataLine &= "" & vbTab ' Size (Girls US)
        'dataLine &= "" & vbTab ' Size (Infant US)
        'dataLine &= "" & vbTab ' Size (Junior US)
        'dataLine &= "" & vbTab ' Size (Men US)
        'dataLine &= "" & vbTab ' Size (Misses US)
        'dataLine &= "" & vbTab ' Size (Toddler US) 
        'dataLine &= "" & vbTab ' Size (Women US)
        'dataLine &= "" & vbTab ' Size Code (Child US)
        'dataLine &= "" & vbTab ' Size Code (Men US)
        'dataLine &= "" & vbTab ' Size Code (Women US)
        'dataLine &= "" & vbTab ' Size Modifier
        'dataLine &= "" & vbTab ' Sleeve Length
        'dataLine &= "" & vbTab ' Sleeve Type
        'dataLine &= "" & vbTab ' Slip Type
        'dataLine &= "" & vbTab ' Sock Length
        'dataLine &= "" & vbTab ' Sports
        'dataLine &= "" & vbTab ' Sports League
        'dataLine &= "" & vbTab ' Stain Resistant
        'dataLine &= "" & vbTab ' Strap Type
        'dataLine &= "" & vbTab ' Sweathshirt Type
        'dataLine &= "" & vbTab ' Swimwear Type
        'dataLine &= "" & vbTab ' Therapeutic Type
        'dataLine &= "" & vbTab ' Tights Type
        'dataLine &= "" & vbTab ' Toe Type
        'dataLine &= "" & vbTab ' Top Length
        'dataLine &= "" & vbTab ' Trend/Style
        dataLine &= GetPartnerCategory(PARTNER_CODE, DEPT_CODE, ITEM_TYPE_CODE) ' Underwear Type
        'dataLine &= "" & vbTab ' Waist Size
        'dataLine &= "" & vbTab ' Wallet Type
        'dataLine &= "" & vbTab ' Waterproof
        'dataLine &= "" & vbTab ' Wet Suit Type
        'dataLine &= "" & vbTab ' WNBA Team
        'dataLine &= "" ' Wrinkle Resistant

        Return dataLine
    End Function

    Private Function BuyDotComInventoryFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine = "##Type=Inventory;Version=5.0" & Environment.NewLine
            dataLine &= "ListingId" & vbTab
            dataLine &= "ProductId" & vbTab
            dataLine &= "ProductIdType" & vbTab
            dataLine &= "ItemCondition" & vbTab
            dataLine &= "Price" & vbTab
            dataLine &= "MAP" & vbTab
            dataLine &= "MAPType" & vbTab
            dataLine &= "Quantity" & vbTab
            dataLine &= "OfferExpeditedShipping" & vbTab
            dataLine &= "Description" & vbTab
            dataLine &= "ShippingRateStandard" & vbTab
            dataLine &= "ShippingRateExpedited" & vbTab
            dataLine &= "ShippingLeadTime" & vbTab
            dataLine &= "OfferTwoDayShipping" & vbTab
            dataLine &= "ShippingRateTwoDay" & vbTab
            dataLine &= "OfferOneDayShipping" & vbTab
            dataLine &= "ShippingRateOneDay" & vbTab
            dataLine &= "OfferSameDayShipping" & vbTab
            dataLine &= "ShippingRateSameDay" & vbTab
            dataLine &= "OfferLocalDeliveryShippingRates" & vbTab
            dataLine &= "ReferenceId" & Environment.NewLine
        End If

        dataLine &= PARTNER_PRODUCT_OUR_ID & vbTab ' ListingId
        dataLine &= rowICTSTYL1("ITEM_CODE") & vbTab ' ProductId
        dataLine &= "3" & vbTab ' ProductIdType
        dataLine &= "1" & vbTab ' ItemCondition
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' MAP
        dataLine &= "0" & vbTab ' MAPType

        If rowICTSTYL1.Item("ITEM_STATUS") <> "A" Then
            dataLine &= "0" & vbTab
        Else
            If PARTNER_PRODUCT_INV_PCT <= 0 OrElse PARTNER_PRODUCT_INV_PCT > 100 Then PARTNER_PRODUCT_INV_PCT = 100
            dataLine &= Math.Floor((GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) * (PARTNER_PRODUCT_INV_PCT / 100))) & vbTab  ' Quantity
        End If

        dataLine &= "1" & vbTab ' OfferExpeditedShipping
        dataLine &= "New" & vbTab ' Description
        dataLine &= GetPartnerAttribute("BUYI", "PARTNER_PRODUCT_SHIP_RATE") & vbTab ' ShippingRateStandard
        dataLine &= GetPartnerAttribute("BUYI", "PARTNER_PRODUCT_SHIP_RATE_ND") & vbTab ' ShippingRateExpedited
        dataLine &= "" & vbTab ' ShippingLeadTime
        dataLine &= "0" & vbTab ' OfferTwoDayShipping
        dataLine &= "" & vbTab ' ShippingRateTwoDay
        dataLine &= "0" & vbTab ' OfferOneDayShipping
        dataLine &= "" & vbTab ' ShippingRateOneDay
        dataLine &= "0" & vbTab ' OfferSameDayShipping
        dataLine &= "" & vbTab ' ShippingRateSameDay
        dataLine &= "0" & vbTab ' OfferLocalDeliveryShippingRates
        dataLine &= rowICTSTYL1("ITEM_CODE") & "" 'ReferenceId

        Return dataLine
    End Function

    Private Function CommJunctionDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine = "&CID=" & GetPartnerAttribute("COMMJUNC", "PARTNER_PRODUCT_OUR_ID") & Environment.NewLine
            dataLine &= "&SUBID=" & GetPartnerAttribute("COMMJUNC", "PARTNER_PRODUCT_OUR_SUBID") & Environment.NewLine
            'dataLine &= "&DATEFMT=" & DateTime.Now.ToString("MM/dd/yyyy hh:mm tt") & Environment.NewLine
            dataLine &= "&PROCESSTYPE=OVERWRITE" & Environment.NewLine
            dataLine &= "&AID=" & GetPartnerAttribute("COMMJUNC", "PARTNER_PRODUCT_AID") & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_DESC_EXT") & Chr(34) & vbTab

        temp = rowICTSTYL1.Item("STYLE_KEYWORD_1") & "," & rowICTSTYL1.Item("STYLE_KEYWORD_2") & "," & rowICTSTYL1.Item("STYLE_KEYWORD_3") & "," & rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS")
        While temp.Contains(",,")
            temp = temp.Replace(",,", ",")
        End While

        If temp.StartsWith(",") AndAlso temp.Length = 1 Then
            temp = String.Empty
        ElseIf temp.StartsWith(",") Then
            temp = temp.Substring(1).Trim & String.Empty
        End If
        dataLine &= Chr(34) & temp & Chr(34) & vbTab

        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_FULL_DESC") & Chr(34) & vbTab
        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_CODE") & Chr(34) & vbTab
        dataLine &= Chr(34) & siteURL & STYLE_CODE.ToLower & ".htm" & Chr(34) & vbTab

        If rowICTSTYL1.Item("STYLE_STATUS") & String.Empty = "A" Then
            dataLine &= Chr(34) & "Yes" & Chr(34) & vbTab
        Else
            dataLine &= Chr(34) & "No" & Chr(34) & vbTab
        End If

        dataLine &= Chr(34) & imageMediaDir & STYLE_CODE.ToLower & ".jpg" & Chr(34) & vbTab
        dataLine &= Chr(34) & GetItemPrice(rowICTSTYL1, "O") & Chr(34) & vbTab
        dataLine &= vbTab

        If "SC".Contains(rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty) Then
            dataLine &= Chr(34) & GetItemPrice(rowICTSTYL1, "C") & Chr(34) & vbTab
        Else
            dataLine &= vbTab
        End If

        dataLine &= Chr(34) & "USD" & Chr(34) & vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= ""

        dataLine = dataLine.Replace(vbTab, ",")

        Return (dataLine)

    End Function

    Private Function formatNumber(ByVal number As Double, Optional ByVal numDecimals As Int16 = 2) As String

        If numDecimals <= 0 Then
            numDecimals = 0
        Else
            If numDecimals > 10 Then numDecimals = 10
        End If

        Return String.Format("{0:n" & numDecimals.ToString.Trim & "}", number)

    End Function

    Private Sub EmailErrors()
        Try
            If Not automatedExport Then Exit Sub
            If TASK_NO.Length = 0 Then Exit Sub

            Dim EMAIL_BODY As String = String.Empty
            For Each row As DataRow In ASCDATA1.GetDataTable("SELECT * FROM ASTTASK2 WHERE TASK_NO = '" & TASK_NO & "'").Select("", "TASK_LNO")
                EMAIL_BODY &= row.Item("TASK_DETAIL") & Environment.NewLine
            Next

            If EMAIL_BODY.Length = 0 Then Exit Sub
            Dim EMAIL_FROM As String = "service@webundies.com"
            Dim EMAIL_SUBJECT As String = "Inventory Feed Errors"

            For Each E_MAIL As String In ROWs("WBTPARM1").Item("WB_PARM_ERRORS_EMAIL").ToString.Split(";")
                E_MAIL = E_MAIL.Trim
                If E_MAIL.Length = 0 Then Continue For
                clsSOCMAIN1.SendEmailMessage(EMAIL_FROM, E_MAIL, EMAIL_SUBJECT, EMAIL_BODY, String.Empty, String.Empty)
            Next

        Catch ex As Exception

        End Try
    End Sub

    Private Function ftpData(ByVal fileToUpload As String, ByVal setActiveMode As Boolean) As Boolean

        If chkTestMode.Checked Then Exit Function

        Try
            If PARTNER_PRODUCT_USER.Length = 0 OrElse PARTNER_PRODUCT_PASS.Length = 0 OrElse PARTNER_PRODUCT_IP.Length = 0 Then
                Return False
            End If

            If PARTNER_PRODUCT_IP.ToUpper = "www.webundies.com".ToUpper Then
                Return SecureftpData(fileToUpload)
            End If

            ' Need this since the FTP control keeps the previous settings.
            Ftp1 = New nsoftware.IPWorks.Ftp
            Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

            Ftp1.User = PARTNER_PRODUCT_USER
            Ftp1.Password = PARTNER_PRODUCT_PASS
            Ftp1.RemoteHost = PARTNER_PRODUCT_IP
            Ftp1.Logon()

            If PARTNER_PRODUCT_REMOTE_DIR.Length > 0 Then
                Ftp1.RemotePath = PARTNER_PRODUCT_REMOTE_DIR
            End If

            ASCMAIN1.Progress("Uploading: " & fileToUpload, String.Empty)
            Ftp1.LocalFile = fileToUpload
            Ftp1.RemoteFile = My.Computer.FileSystem.GetName(fileToUpload)
            Ftp1.Overwrite = True
            If setActiveMode Then
                Ftp1.Passive = False
            End If
            Ftp1.Upload()

            ftpData = True

        Catch ex As Exception
            If automatedExport Then
                clsSOCMAIN1.AddTaskDetail(TASK_NO, "Error uploading order files (" & PARTNER_NAME & "): " & ex.Message)
            Else
                MessageBox.Show("Error uploading order files (" & PARTNER_NAME & "): " & ex.Message, "Error", MessageBoxButtons.OK)
            End If

            ftpData = False

        Finally
            Ftp1.Logoff()
            Ftp1.Dispose()
        End Try

    End Function

    Private Function SecureftpData(ByVal fileToUpload As String) As Boolean

        If chkTestMode.Checked Then Exit Function

        Try
            If PARTNER_PRODUCT_USER.Length = 0 OrElse PARTNER_PRODUCT_PASS.Length = 0 OrElse PARTNER_PRODUCT_IP.Length = 0 Then
                Return False
            End If

            ' Need this since the FTP control keeps the previous settings.
            Sftp1 = New nsoftware.IPWorksSSH.Sftp

            Sftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
            ASCMAIN1.Progress("-", "RemoteHost")

            ASCMAIN1.Progress("-", "User")
            Sftp1.SSHUser = PARTNER_PRODUCT_USER

            ASCMAIN1.Progress("-", "Password")
            Sftp1.SSHPassword = PARTNER_PRODUCT_PASS

            ASCMAIN1.Progress("-", "RemoteFile")
            Sftp1.RemoteFile = String.Empty

            ASCMAIN1.Progress("-", "Timeout")
            Sftp1.Timeout = 300

            ASCMAIN1.Progress("-", "Logon")
            Sftp1.SSHAuthMode = nsoftware.IPWorksSSH.SftpSSHAuthModes.amPassword
            Try
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_PRODUCT_IP, 22)
            Catch ex As Exception
                Sftp1.SSHLogoff()
                Sftp1.SSHLogon(PARTNER_PRODUCT_IP, 22)
            End Try

            If PARTNER_PRODUCT_REMOTE_DIR.Length > 0 Then
                Sftp1.RemotePath = PARTNER_PRODUCT_REMOTE_DIR
            End If

            ASCMAIN1.Progress("Uploading: " & fileToUpload, String.Empty)
            Sftp1.LocalFile = fileToUpload
            Sftp1.RemoteFile = My.Computer.FileSystem.GetName(fileToUpload)
            Sftp1.Overwrite = True
            Sftp1.Upload()

            SecureftpData = True

        Catch ex As Exception
            If automatedExport Then
                clsSOCMAIN1.AddTaskDetail(TASK_NO, "Error uploading order files (" & PARTNER_CODE & "): " & ex.Message)
            Else
                MessageBox.Show("Error uploading order files (" & PARTNER_CODE & "): " & ex.Message, "Error", MessageBoxButtons.OK)
            End If

            SecureftpData = False

        Finally
            Sftp1.SSHLogoff()
            Sftp1.Dispose()
        End Try

    End Function

    Private Function GetCareInstructions(ByRef rowICTSTYL1 As DataRow) As String

        Dim careInstr As String = String.Empty

        If rowICTSTYL1.Item("STYLE_CARE_MW") & String.Empty = "1" Then
            careInstr &= "; Machine Washable"
        End If

        If rowICTSTYL1.Item("STYLE_CARE_EC") & String.Empty = "1" Then
            careInstr &= "; Easy Care"
        End If

        If rowICTSTYL1.Item("STYLE_CARE_CW") & String.Empty = "1" Then
            careInstr &= "; Cold Water"
        End If

        If rowICTSTYL1.Item("STYLE_CARE_DC") & String.Empty = "1" Then
            careInstr &= "; Delicate Cycle"
        End If

        If careInstr.Length > 0 Then
            careInstr = careInstr.Substring(1).Trim
        End If

        Return careInstr

    End Function

    Private Function GetFeatures(ByVal STYLE_CODE As String, Optional ByVal separator As String = "") As String
        Dim features As String = String.Empty

        If separator = "" Then separator = ";"

        For Each rowICTFEATX As DataRow In dst.Tables("ICTFEATX").Select("STYLE_CODE = '" & STYLE_CODE & "'", "FEATURE_DESC")
            features &= separator & rowICTFEATX.Item("FEATURE_DESC") & String.Empty
        Next
        If features.Length > 0 Then
            features = features.Substring(1).Trim
        End If

        Return features

    End Function

    Private Function GetPartnerCategory(ByVal PARTNER_CODE As String, ByVal DEPT_CODE As String, ByVal ITEM_TYPE_CODE As String) As String
        Dim sql As String = String.Empty

        sql = "PARTNER_CODE = '" & PARTNER_CODE & "' AND DEPT_CODE = '" & DEPT_CODE & "' AND ITEM_TYPE_CODE = '" & ITEM_TYPE_CODE & "'"
        If dst.Tables("SOTPART2").Select(sql).Length > 0 Then
            Return dst.Tables("SOTPART2").Select(sql)(0).Item("PARTNER_CATEGORY") & String.Empty
        Else
            sql = "PARTNER_CODE = '" & PARTNER_CODE & "' AND DEPT_CODE = '" & DEPT_CODE & "' AND ITEM_TYPE_CODE = '*'"
            If dst.Tables("SOTPART2").Select(sql).Length > 0 Then
                Return dst.Tables("SOTPART2").Select(sql)(0).Item("PARTNER_CATEGORY") & String.Empty
            Else
                Return String.Empty
            End If
        End If

    End Function

    Private Function GetPartnerCategoryExtensions(ByVal STYLE_CODE As String, ByVal FieldName As String, ByVal Separator As String, Optional ByVal ConversionChar As String = "") As String

        Dim sql As String = String.Empty
        Dim result As String = String.Empty

        sql = " Select Distinct WBTPAGE1." & FieldName
        sql &= " From ICTSTYL1,ICTSTYL3, WBTPAGE1"
        sql &= " Where ICTSTYL1.STYLE_CODE = ICTSTYL3.STYLE_CODE"
        sql &= " and STYLE_STATUS = 'A'"
        sql &= " and ICTSTYL3.PAGE_CODE = WBTPAGE1.PAGE_CODE "
        sql &= " and  ICTSTYL1.STYLE_CODE = :PARM1 "
        sql &= " and WBTPAGE1." & FieldName & " IS NOT NULL"
        sql &= " order by WBTPAGE1." & FieldName

        For Each row As DataRow In ASCDATA1.GetDataTable(sql, "", "V", New Object() {STYLE_CODE}).Rows
            result &= Separator & row.Item(FieldName) & String.Empty
        Next

        ' the character '-' is a char separating data. Some partners want a different code.
        If ConversionChar.Length > 0 Then
            result = result.Replace("-", ConversionChar)
        End If

        If result.Length > 0 AndAlso result.StartsWith(Separator) Then
            result = result.Substring(1)
        End If

        Return result
    End Function

    Private Function GetPartnerCategoryId(ByVal PARTNER_CODE As String, ByVal DEPT_CODE As String, ByVal ITEM_TYPE_CODE As String) As String
        Dim sql As String = String.Empty

        sql = "PARTNER_CODE = '" & PARTNER_CODE & "' AND DEPT_CODE = '" & DEPT_CODE & "' AND ITEM_TYPE_CODE = '" & ITEM_TYPE_CODE & "'"
        If dst.Tables("SOTPART2").Select(sql).Length > 0 Then
            Return dst.Tables("SOTPART2").Select(sql)(0).Item("PARTNER_CATEGORY_ID") & String.Empty
        Else
            sql = "PARTNER_CODE = '" & PARTNER_CODE & "' AND DEPT_CODE = '" & DEPT_CODE & "' AND ITEM_TYPE_CODE = '*'"
            If dst.Tables("SOTPART2").Select(sql).Length > 0 Then
                Return dst.Tables("SOTPART2").Select(sql)(0).Item("PARTNER_CATEGORY_ID") & String.Empty
            Else
                Return String.Empty
            End If
        End If

    End Function

    Private Function GetInventoryLevel(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal SIZE_CODE As String) As Int16

        Dim whseQtyAvail As Int16 = 0
        Dim row As DataRow = Nothing
        Dim sql As String = "STYLE_CODE = '" & STYLE_CODE & "' AND COLOR_CODE = '" & COLOR_CODE & "' AND SIZE_CODE = '" & SIZE_CODE & "' AND WHSE_CODE = '001'"

        If dst.Tables("ICTSTAT2").Select(sql).Length > 0 Then
            row = dst.Tables("ICTSTAT2").Select(sql)(0)
            whseQtyAvail = Val(row.Item("WHSE_QTY_ON_HAND") & String.Empty) - (Val(row.Item("WHSE_QTY_PICK") & String.Empty) + Val(row.Item("WHSE_QTY_OPEN") & String.Empty))
        End If

        If whseQtyAvail < 0 Then whseQtyAvail = 0

        Return whseQtyAvail

    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="ActiveOnly">Only select items where the Style and Item are active</param>
    ''' <param name="lastDateforPartner">Partner Code to get the last datetime for Inactive Styles</param>
    ''' <remarks></remarks>
    Private Sub GetItemsToExport(ByVal ActiveOnly As Boolean, ByVal lastDateforPartner As String)

        Dim sql As String
        Static ICTSTYL1 As String = String.Empty

        sql = "  SELECT ICTSTYL1.*, ICTITEM1.ITEM_CODE, ICTITEM1.COLOR_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.ITEM_GTIN"
        sql &= "  , NVL(ICTITEM1.INIT_DATE, SYSDATE) I_INIT_DATE, NVL(ICTITEM1.LAST_DATE, SYSDATE) I_LAST_DATE"
        sql &= "  , ICTITEM1.ITEM_STATUS, ICTSIZE1.SIZE_DESC, NVL(ICTSIZE1.SIZE_SEQ_NO, 0) SIZE_SEQ_NO"
        sql &= "  , TRIM(ICTSTYL1.STYLE_DESC)"
        sql &= "  || DECODE(ICTDEPT1.DEPT_CODE, 'MENS', ' for men','WOMENS', ' for women', 'GIRLS' , ' for girls', 'BOYS', ' for boys', '')"
        sql &= "  || DECODE(ICTSIZE1.SIZE_DESC, NULL, '', ' (' || ICTSIZE1.SIZE_DESC || ')') STYLE_DESC_XML"
        sql &= "  , ICTLICE1.LICENSE_DESC, ICTTYPE1.ITEM_TYPE_DESC, ICTTYPE1.ITEM_WEIGHT, ICTDEPT1.DEPT_DESC, ICTMATL1.MATL_DESC, ICTSRCE1.SOURCE_DESC"
        sql &= "  , ICTSCOL1.STYLE_COLOR_DESC, ICTSCOL1.STYLE_COLOR_STANDARD"
        sql &= "  , TRIM(ICTSTYL1.STYLE_DESC) || DECODE(ICTDEPT1.DEPT_CODE, 'MENS', ' for men','WOMENS', ' for women', 'GIRLS' , ' for girls', 'BOYS', ' for boys', '') STYLE_DESC_EXT"
        sql &= "  FROM ICTSTYL1, ICTITEM1, ICTSIZE1, ICTLICE1, ICTTYPE1, ICTDEPT1, ICTMATL1, ICTSRCE1, ICTSCOL1"
        sql &= "  WHERE ICTSTYL1.STYLE_CODE = ICTITEM1.STYLE_CODE"
        sql &= "  AND ICTITEM1.SIZE_CODE = ICTSIZE1.SIZE_CODE (+)"
        sql &= "  AND NVL(ICTSTYL1.STYLE_EXCL_FROM_DATAFEED, '0') = '0'"
        sql &= "  AND ICTSTYL1.LICENSE_CODE = ICTLICE1.LICENSE_CODE (+)"
        sql &= "  AND ICTSTYL1.ITEM_TYPE_CODE = ICTTYPE1.ITEM_TYPE_CODE (+)"
        sql &= "  AND ICTSTYL1.DEPT_CODE = ICTDEPT1.DEPT_CODE (+)"
        sql &= "  AND ICTSTYL1.MATL_CODE = ICTMATL1.MATL_CODE (+)"
        sql &= "  AND ICTSTYL1.SOURCE_CODE = ICTSRCE1.SOURCE_CODE (+)"
        sql &= "  AND ICTSTYL1.STYLE_COLOR_CODE = ICTSCOL1.STYLE_COLOR_CODE (+)"

        If ActiveOnly Then
            sql &= " AND ICTITEM1.ITEM_STATUS = 'A' AND ICTSTYL1.STYLE_STATUS = 'A'"
        End If

        ' Use temp table for Status that went inactive since the last Update
        If lastDateforPartner.Length > 0 AndAlso Not ActiveOnly Then
            If ICTSTYL1.Length = 0 Then
                ICTSTYL1 = ASCMAIN1.Temp_Table(sql)
            Else
                ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ICTSTYL1)
                ASCDATA1.ExecuteSQL("INSERT INTO " & ICTSTYL1 & " " & sql)
            End If

            sql = "DELETE FROM " & ICTSTYL1
            sql &= "  WHERE NVL(STYLE_STATUS, 'A') <> 'A'"
            sql &= " AND LAST_DATE < (SELECT PARTNER_PRODUCT_LAST_EXTRACT FROM SOTPART1 WHERE PARTNER_CODE = '" & lastDateforPartner & "')"
            ASCDATA1.ExecuteSQL(sql)

            sql = "DELETE FROM " & ICTSTYL1
            sql &= "  WHERE NVL(ITEM_STATUS, 'A') <> 'A'"
            sql &= " AND I_LAST_DATE < (SELECT PARTNER_PRODUCT_LAST_EXTRACT FROM SOTPART1 WHERE PARTNER_CODE = '" & lastDateforPartner & "')"
            ASCDATA1.ExecuteSQL(sql)

            sql = "SELECT * FROM " & ICTSTYL1
        End If

        itemsDatatable = ASCDATA1.GetDataTable(sql)

    End Sub

    Private Function GetItemPrice(ByRef rowICTSTYL1 As DataRow, ByVal priceType As String) As String

        Dim price As Double = 0

        If rowICTSTYL1 Is Nothing Then
            Return price
        End If

        Select Case priceType
            Case "M"
                price = Val(rowICTSTYL1.Item("STYLE_PRICE_MFR") & String.Empty)

            Case "O"
                price = Val(rowICTSTYL1.Item("STYLE_PRICE_OUR") & String.Empty)

            Case Else
                ' Current Price
                If "SC".Contains(rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty) Then
                    price = Val(rowICTSTYL1.Item("STYLE_PRICE_SALE") & String.Empty)
                Else
                    price = Val(rowICTSTYL1.Item("STYLE_PRICE_OUR") & String.Empty)
                End If
        End Select

        Return formatNumber(price)

    End Function

    Private Function GetPartnerAttribute(ByVal PARTNER_CODE As String, ByVal fieldName As String) As Object

        If dst.Tables("SOTPART1").Select("PARTNER_CODE = '" & PARTNER_CODE & "'").Length = 0 Then
            Return String.Empty
        End If

        If Not dst.Tables("SOTPART1").Columns.Contains(fieldName) Then
            Return String.Empty
        End If

        Return dst.Tables("SOTPART1").Select("PARTNER_CODE = '" & PARTNER_CODE & "'")(0).Item(fieldName) & String.Empty
    End Function

    Private Function GoogleDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "link" & vbTab
            dataLine &= "title" & vbTab
            dataLine &= "description" & vbTab
            dataLine &= "price" & vbTab
            dataLine &= "image_link" & vbTab
            dataLine &= "google_product_category" & vbTab 'required 9/2011 must be valid in Google Taxonomy
            dataLine &= "product_type" & vbTab 'WUN version - category with ITEM_TYPE_DESC
            dataLine &= "c:department" & vbTab
            dataLine &= "c:style" & vbTab
            dataLine &= "id" & vbTab 'item
            dataLine &= "item group id" & vbTab 'parent style new 9/2011 to provide by item
            'dataLine &= "quantity" & vbTab 'removed by Google 9/2011
            dataLine &= "size" & vbTab ' new 9/2011
            dataLine &= "age group" & vbTab ' new 9/2011
            dataLine &= "gender" & vbTab ' new 9/2011
            dataLine &= "brand" & vbTab ' new 9/2011
            dataLine &= "color" & vbTab ' new 9/2011
            dataLine &= "material" & vbTab ' new 9/2011
            dataLine &= "availability" & vbTab ' new 9/2011
            dataLine &= "condition" & Environment.NewLine
        End If

        '' Only uses Styles; therefore, only send style once
        'If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
        '    Return String.Empty
        'End If

        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' Title
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' Description
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' Image
        dataLine &= GetPartnerCategory("GOOGLE", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' Category
        dataLine &= GetPartnerCategory("GOOGLE", DEPT_CODE, ITEM_TYPE_CODE) & " > " & rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Type
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & vbTab ' Department
        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Style
        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' ID
        dataLine &= STYLE_CODE & vbTab ' Item Group ID
        'dataLine &= "50" & vbTab ' Quantity
        ' size
        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If

        ' age group  
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "BOYS", "GIRLS"
                dataLine &= "kids" & vbTab
            Case "MENS", "WOMENS"
                dataLine &= "adult" & vbTab
            Case Else
                dataLine &= "adult" & vbTab
        End Select

        ' gender 
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "MENS", "BOYS"
                dataLine &= "male" & vbTab
            Case "WOMENS", "GIRLS"
                dataLine &= "female" & vbTab
            Case Else
                dataLine &= "unisex" & vbTab
        End Select

        ' brand
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' color
        dataLine &= rowICTSTYL1.Item("MATL_DESC") & String.Empty & vbTab 'material

        ' Availability
        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) > 0 Then
            dataLine &= "in stock" & vbTab
        Else
            dataLine &= "out of stock" & vbTab
        End If

        dataLine &= "New"  ' Condition

        Return dataLine

    End Function

    Private Function CPCDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String
        'from Google with two added fields
        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "link" & vbTab
            dataLine &= "title" & vbTab
            dataLine &= "description" & vbTab
            dataLine &= "price" & vbTab
            dataLine &= "image_link" & vbTab
            dataLine &= "product_category" & vbTab
            dataLine &= "product_type" & vbTab
            dataLine &= "c:department" & vbTab
            dataLine &= "c:style" & vbTab
            dataLine &= "id" & vbTab 'item
            dataLine &= "item group id" & vbTab
            dataLine &= "size" & vbTab
            dataLine &= "age group" & vbTab
            dataLine &= "gender" & vbTab
            dataLine &= "brand" & vbTab
            dataLine &= "color" & vbTab
            dataLine &= "material" & vbTab
            dataLine &= "availability" & vbTab
            dataLine &= "condition" & vbTab
            dataLine &= "original price" & vbTab
            dataLine &= "keywords" & Environment.NewLine

        End If

        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' Title
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' Description
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' Image
        dataLine &= GetPartnerCategory("CPC", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' Category
        dataLine &= GetPartnerCategory("CPC", DEPT_CODE, ITEM_TYPE_CODE) & " > " & rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Type
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & vbTab ' Department
        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Style
        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' ID
        dataLine &= STYLE_CODE & vbTab ' Item Group ID

        ' size
        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If

        ' age group  
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "BOYS", "GIRLS"
                dataLine &= "kids" & vbTab
            Case "MENS", "WOMENS"
                dataLine &= "adult" & vbTab
            Case Else
                dataLine &= "adult" & vbTab
        End Select

        ' gender 
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "MENS", "BOYS"
                dataLine &= "male" & vbTab
            Case "WOMENS", "GIRLS"
                dataLine &= "female" & vbTab
            Case Else
                dataLine &= "unisex" & vbTab
        End Select

        ' brand
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' color
        dataLine &= rowICTSTYL1.Item("MATL_DESC") & String.Empty & vbTab 'material

        ' Availability
        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) > 0 Then
            dataLine &= "in stock" & vbTab
        Else
            dataLine &= "out of stock" & vbTab
        End If

        dataLine &= "New" & vbTab ' Condition
        dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") ' Keywords

        Return dataLine

    End Function

    Private Function SearchSpringDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String
        'from Google with two added fields
        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "link" & vbTab
            dataLine &= "title" & vbTab
            dataLine &= "description" & vbTab
            dataLine &= "price" & vbTab
            dataLine &= "image_link" & vbTab
            dataLine &= "product_category" & vbTab
            dataLine &= "product_type" & vbTab
            dataLine &= "c:department" & vbTab
            dataLine &= "c:style" & vbTab
            dataLine &= "id" & vbTab 'item
            dataLine &= "item group id" & vbTab
            dataLine &= "size" & vbTab
            dataLine &= "age group" & vbTab
            dataLine &= "gender" & vbTab
            dataLine &= "brand" & vbTab
            dataLine &= "color" & vbTab
            dataLine &= "material" & vbTab
            dataLine &= "availability" & vbTab
            dataLine &= "condition" & vbTab
            dataLine &= "original price" & vbTab
            dataLine &= "keywords" & Environment.NewLine

        End If

        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' Title
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' Description
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' Image
        dataLine &= GetPartnerCategory("SEARCHSPR", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' Category
        dataLine &= GetPartnerCategory("SEARCHSPR", DEPT_CODE, ITEM_TYPE_CODE) & " > " & rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Type
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & vbTab ' Department
        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' Style
        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' ID
        dataLine &= STYLE_CODE & vbTab ' Item Group ID

        ' size
        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If

        ' age group  
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "BOYS", "GIRLS"
                dataLine &= "kids" & vbTab
            Case "MENS", "WOMENS"
                dataLine &= "adult" & vbTab
            Case Else
                dataLine &= "adult" & vbTab
        End Select

        ' gender 
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "MENS", "BOYS"
                dataLine &= "male" & vbTab
            Case "WOMENS", "GIRLS"
                dataLine &= "female" & vbTab
            Case Else
                dataLine &= "unisex" & vbTab
        End Select

        ' brand
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' color
        dataLine &= rowICTSTYL1.Item("MATL_DESC") & String.Empty & vbTab 'material

        ' Availability
        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) > 0 Then
            dataLine &= "in stock" & vbTab
        Else
            dataLine &= "out of stock" & vbTab
        End If

        dataLine &= "New" & vbTab ' Condition
        dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") ' Keywords

        Return dataLine

    End Function

    Private Function MShopperDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "BRANDNAME" & vbTab
            dataLine &= "PRODUCTNAME" & vbTab
            dataLine &= "SHORTDESCRIPTION" & vbTab
            dataLine &= "LONGDESCRIPTION" & vbTab
            dataLine &= "MODELNUMBER" & vbTab
            dataLine &= "UPC" & vbTab
            dataLine &= "SKU" & vbTab
            dataLine &= "SIZE" & vbTab
            dataLine &= "CATEGORY" & vbTab
            dataLine &= "SUBCATEGORY" & vbTab
            dataLine &= "IMAGEURL" & vbTab
            dataLine &= "IMAGEURL2" & vbTab
            dataLine &= "IMAGEURL3" & vbTab
            dataLine &= "BUYURL" & vbTab
            dataLine &= "PRICE" & vbTab
            dataLine &= "SALEPRICE" & vbTab
            dataLine &= "SHIPCOST" & vbTab
            dataLine &= "SHIPZIP" & vbTab
            dataLine &= "KEYWORDS" & vbTab
            dataLine &= "QUANTITY" & vbTab
            dataLine &= "CONDITION" & Environment.NewLine
        End If

        ' Do not send over 0 quantities
        If PARTNER_PRODUCT_INV_PCT <= 0 OrElse PARTNER_PRODUCT_INV_PCT > 100 Then PARTNER_PRODUCT_INV_PCT = 100
        Dim qty As Int16 = Math.Floor((GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) * (PARTNER_PRODUCT_INV_PCT / 100)))
        If qty <= 0 Then
            Return dataLine
        End If

        dataLine &= "WebUndies" & vbTab 'BRANDNAME
        dataLine &= rowICTSTYL1.Item("STYLE_DESC") & vbTab ' PRODUCTNAME
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' SHORTDESCRIPTION
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' LONGDESCRIPTION
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab ' MODELNUMBER
        dataLine &= "" & vbTab ' UPC is blank
        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' SKU

        'SIZE
        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If

        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & vbTab ' CATEGORY
        dataLine &= GetPartnerCategory("MSHOPPER", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' SUBCATEGORY

        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' IMAGEURL

        ' IMAGEURL 1
        If rowICTSTYL1.Item("STYLE_IMAGE_OTHER1") & String.Empty <> String.Empty Then
            dataLine &= imageMediaDir & rowICTSTYL1.Item("STYLE_IMAGE_OTHER1") & vbTab
        Else
            dataLine &= vbTab
        End If

        ' IMAGEURL 2
        If rowICTSTYL1.Item("STYLE_IMAGE_OTHER2") & String.Empty <> String.Empty Then
            dataLine &= imageMediaDir & rowICTSTYL1.Item("STYLE_IMAGE_OTHER2") & vbTab
        Else
            dataLine &= vbTab
        End If

        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' BUYURL
        dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab ' PRICE

        ' SALEPRICE
        If rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty = "S" Then
            dataLine &= GetItemPrice(rowICTSTYL1, "S") & vbTab
        Else
            dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab
        End If

        ' SHIPCOST
        dataLine &= formatNumber(Val(GetPartnerAttribute("MSHOPPER", "PARTNER_PRODUCT_SHIP_RATE") & String.Empty)) & vbTab
        dataLine &= WHSE_ZIP_CODE & vbTab ' SHIPZIP
        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") & vbTab ' KEYWORDS

        ' QUANTITY
        If PARTNER_PRODUCT_INV_PCT <= 0 OrElse PARTNER_PRODUCT_INV_PCT > 100 Then PARTNER_PRODUCT_INV_PCT = 100
        dataLine &= qty.ToString.Trim & vbTab

        dataLine &= "New" ' CONDITION

        Return dataLine

    End Function

    Private Function CatalogsDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String
        'Catalogs.com - copy of Google to start

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "id" & vbTab 'item 
            dataLine &= "title" & vbTab
            dataLine &= "link" & vbTab
            dataLine &= "price" & vbTab
            dataLine &= "description" & vbTab
            dataLine &= "brand" & vbTab
            dataLine &= "image_link" & vbTab
            dataLine &= "product_category" & vbTab 'use same as Google Taxonomy
            dataLine &= "inventory" & vbTab
            dataLine &= "color" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        dataLine &= STYLE_CODE & vbTab ' ID
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' Title
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' Link
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab ' Price
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' Description

        ' brand
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' Image
        'dataLine &= GetPartnerCategory("CATALOGS", DEPT_CODE, ITEM_TYPE_CODE) & GetParnterCategoryExtensions(STYLE_CODE) & vbTab ' Category
        dataLine &= GetPartnerCategoryExtensions(STYLE_CODE, "PAGE_GROUP", "^") & vbTab ' Category
        dataLine &= "50" & vbTab ' Inventory
        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty ' color
        Return dataLine

    End Function

    Private Sub LoadPartnerParameters(ByVal partnerCode As String)

        Dim rowSOTPART1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPART1 WHERE PARTNER_CODE = :PARM1", True, "V", New String() {partnerCode})

        PARTNER_CODE = rowSOTPART1.Item("PARTNER_CODE") & String.Empty
        PARTNER_STATUS = rowSOTPART1.Item("PARTNER_STATUS") & String.Empty
        PARTNER_NAME = rowSOTPART1.Item("PARTNER_NAME") & String.Empty
        PARTNER_ORDR_SOURCE_CODE = rowSOTPART1.Item("PARTNER_ORDR_SOURCE_CODE") & String.Empty
        PARTNER_SITE_IP = rowSOTPART1.Item("PARTNER_SITE_IP") & String.Empty
        PARTNER_SITE_USER = rowSOTPART1.Item("PARTNER_SITE_USER") & String.Empty
        PARTNER_SITE_PWD = rowSOTPART1.Item("PARTNER_SITE_PWD") & String.Empty
        PARTNER_SITE_OUTPUT_DIR = rowSOTPART1.Item("PARTNER_SITE_OUTPUT_DIR") & String.Empty
        PARTNER_ORDERS_DIR = rowSOTPART1.Item("PARTNER_ORDERS_DIR") & String.Empty
        PARTNER_LAST_SALES_ORDER = rowSOTPART1.Item("PARTNER_LAST_SALES_ORDER") & String.Empty
        PARTNER_SITE_ORDERS_POST_URL = rowSOTPART1.Item("PARTNER_SITE_ORDERS_POST_URL") & String.Empty
        PARTNER_OUR_ID = rowSOTPART1.Item("PARTNER_OUR_ID") & String.Empty
        PARTNER_OUR_SITE_NAME = rowSOTPART1.Item("PARTNER_OUR_SITE_NAME") & String.Empty
        PARTNER_SHIP_CONF_FILENAME = rowSOTPART1.Item("PARTNER_SHIP_CONF_FILENAME") & String.Empty
        PARTNER_SHIP_CONF_LOCAL_DIR = rowSOTPART1.Item("PARTNER_SHIP_CONF_LOCAL_DIR") & String.Empty
        PARTNER_SHIP_CONF_IP = rowSOTPART1.Item("PARTNER_SHIP_CONF_IP") & String.Empty
        PARNTER_SHIP_CONF_USER = rowSOTPART1.Item("PARNTER_SHIP_CONF_USER") & String.Empty
        PARTNER_SHIP_CONF_PASS = rowSOTPART1.Item("PARTNER_SHIP_CONF_PASS") & String.Empty
        PARTNER_SHIP_CONF_REMOTE_DIR = rowSOTPART1.Item("PARTNER_SHIP_CONF_REMOTE_DIR") & String.Empty
        PARTNER_PRODUCT_FILENAME = rowSOTPART1.Item("PARTNER_PRODUCT_FILENAME") & String.Empty
        PARTNER_PRODUCT_LOCAL_DIR = rowSOTPART1.Item("PARTNER_PRODUCT_LOCAL_DIR") & String.Empty
        PARTNER_PRODUCT_LAST_EXTRACT = rowSOTPART1.Item("PARTNER_PRODUCT_LAST_EXTRACT") & String.Empty
        PARTNER_PRODUCT_IP = rowSOTPART1.Item("PARTNER_PRODUCT_IP") & String.Empty
        PARTNER_PRODUCT_USER = rowSOTPART1.Item("PARTNER_PRODUCT_USER") & String.Empty
        PARTNER_PRODUCT_PASS = rowSOTPART1.Item("PARTNER_PRODUCT_PASS") & String.Empty
        PARTNER_PRODUCT_REMOTE_DIR = rowSOTPART1.Item("PARTNER_PRODUCT_REMOTE_DIR") & String.Empty
        PARTNER_PRODUCT_OUR_ID = rowSOTPART1.Item("PARTNER_PRODUCT_OUR_ID") & String.Empty
        PARTNER_PRODUCT_OUR_SUBID = rowSOTPART1.Item("PARTNER_PRODUCT_OUR_SUBID") & String.Empty
        PARTNER_PRODUCT_INV_PCT = Val(rowSOTPART1.Item("PARTNER_PRODUCT_INV_PCT") & String.Empty)
        PARTNER_PRODUCT_AID = rowSOTPART1.Item("PARTNER_PRODUCT_AID") & String.Empty
        PARTNER_PRODUCT_SHIP_RATE = Val(rowSOTPART1.Item("PARTNER_PRODUCT_SHIP_RATE") & String.Empty)
        PARTNER_PRODUCT_PROMO_TEXT = rowSOTPART1.Item("PARTNER_PRODUCT_PROMO_TEXT") & String.Empty
        PARTNER_PRODUCT_SHIP_RATE_GR = Val(rowSOTPART1.Item("PARTNER_PRODUCT_SHIP_RATE_GR") & String.Empty)
        PARTNER_PRODUCT_SHIP_RATE_2D = Val(rowSOTPART1.Item("PARTNER_PRODUCT_SHIP_RATE_2D") & String.Empty)
        PARTNER_PRODUCT_SHIP_RATE_ND = Val(rowSOTPART1.Item("PARTNER_PRODUCT_SHIP_RATE_ND") & String.Empty)
        PARTNER_PRODUCT_INV_MIN = Val(rowSOTPART1.Item("PARTNER_PRODUCT_INV_MIN") & String.Empty)
        PARTNER_PRODUCT_ARCHIVE_DAYS = Val(rowSOTPART1.Item("PARTNER_PRODUCT_ARCHIVE_DAYS") & String.Empty)
        SEND_ONCE_DAY = Val(rowSOTPART1.Item("SEND_ONCE_DAY") & String.Empty) = 1

        If PARTNER_PRODUCT_FILENAME.Contains("{datetime}") Then
            PARTNER_PRODUCT_FILENAME = PARTNER_PRODUCT_FILENAME.Replace("{datetime}", DateTime.Now.ToString("yyyyMMdd_hhmmss"))
        End If

        If PARTNER_PRODUCT_LAST_EXTRACT.Length = 0 Then
            rowSOTPART1.Item("PARTNER_PRODUCT_LAST_EXTRACT") = DateTime.Now.AddDays(-30)
            PARTNER_PRODUCT_LAST_EXTRACT = rowSOTPART1.Item("PARTNER_PRODUCT_LAST_EXTRACT")
            ASCDATA1.ExecuteSQL("UPDATE SOTPART1 SET PARTNER_PRODUCT_LAST_EXTRACT = SYSDATE - 30  WHERE PARTNER_CODE = :PARM1", "V", New String() {partnerCode})
        End If

        Fill_Records("SOTPART2", partnerCode)

    End Sub

    Private Function NextTagDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty

        If includeHeader Then
            dataLine = "Product Title" & vbTab
            dataLine &= "Manufacturer Name" & vbTab
            dataLine &= "Price" & vbTab
            dataLine &= "URL of product page" & vbTab
            dataLine &= "Manufacturer SKU" & vbTab
            dataLine &= "Product Category" & vbTab
            dataLine &= "Image URL of product" & vbTab
            dataLine &= "Product Description" & vbTab
            dataLine &= "Distributor ID" & vbTab
            dataLine &= "MSRP Price" & vbTab
            dataLine &= "Stock Status" & vbTab
            dataLine &= "Ground shipping" & vbTab
            dataLine &= "2nd day shipping" & vbTab
            dataLine &= "Overnight" & vbTab
            dataLine &= "Product Weight" & vbTab
            dataLine &= "UPC" & vbTab
            dataLine &= "Marketing Message" & vbTab
            dataLine &= "Warranty" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        'Product Title                      ICTSTYL1.STYLE_NAME
        'Manufacturer Name      WebUndies
        'Price                                      STYLE_PRICE_TYPE = "C" or "S" use STYLE_PRICE_SALE else use STYLE_PRICE_OUR
        'URL of product page       http://www.webundies.com/17fp005.htm
        'Manufacturer SKU          ICTSTYL1.STYLE_CODE
        'Product Category             SOTPART2.PARTNER_CATEGORY - DATA IS LOADED UP
        'Image URL of product    http://www.webundies.com/media/products/17fp005.jpg
        'Product Description        ICTSTYL1.STYLE_FULL_DESC
        'Distributor ID                     null
        'MSRP Price                         ICTSTYL1.STYLE_PRICE_MFR
        'Stock Status                       Yes
        'Ground shipping              PARTNER_PRODUCT_SHIP_RATE             <<< use SHIP_RATE not SHIP_RATE_GR for now 
        '2nd day shipping              PARTNER_PRODUCT_SHIP_RATE_2D
        'Overnight                            PARTNER_PRODUCT_SHIP_RATE_ND
        'Product Weight                null
        'UPC                                       null
        'Marketing Message        SOTPART1.PARTNER_PRODUCT_PROMO_TEXT
        'Warranty                             null


        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab
        dataLine &= "WebUndies" & vbTab
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab
        dataLine &= GetPartnerCategory("NEXTAG", DEPT_CODE, ITEM_TYPE_CODE) & vbTab
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab
        dataLine &= vbTab
        dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab
        dataLine &= "Yes" & vbTab
        dataLine &= formatNumber(Val(GetPartnerAttribute("NEXTAG", "PARTNER_PRODUCT_SHIP_RATE") & String.Empty)) & vbTab
        dataLine &= formatNumber(Val(GetPartnerAttribute("NEXTAG", "PARTNER_PRODUCT_SHIP_RATE_2D") & String.Empty)) & vbTab
        dataLine &= formatNumber(Val(GetPartnerAttribute("NEXTAG", "PARTNER_PRODUCT_SHIP_RATE_ND") & String.Empty)) & vbTab
        dataLine &= vbTab
        dataLine &= vbTab
        dataLine &= GetPartnerAttribute("NEXTAG", "PARTNER_PRODUCT_PROMO_TEXT") & String.Empty & vbTab
        dataLine &= ""

        Return (dataLine)

    End Function

    Private Function PriceGrabberDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty

        'PRODUCTNAME	PRICE	AVAIL	URL	SKU	IMAGE URL	CATEGORY	MARKETING TEXT

        'Disney's Phineas & Ferb Secret Agent Perry Lounge Pants for men	
        '17.99
        'Yes	
        'http://www.webundies.com/cgi-bin/redirect?goto=/17fp005.htm&code=pg12025	
        '17FP005	
        'http://www.webundies.com/images/17fp005.jpg	
        'Clothing > Men's Lounge & Underwear > Pajamas & Robes>pajama bottoms	
        'These lounge pants for men feature Secret Agent Perry, Phineas and Ferb's platypus pet, 
        'from Disney Channel's award winning show in an all-over print on a blue background. 
        'Machine washable with button fly and covered elastic waistband with drawstring tie. 	

        If includeHeader Then
            dataLine = "PRODUCTNAME" & vbTab
            dataLine &= "PRICE" & vbTab
            dataLine &= "AVAIL" & vbTab
            dataLine &= "URL" & vbTab
            dataLine &= "SKU" & vbTab
            dataLine &= "IMAGE URL" & vbTab
            dataLine &= "CATEGORY" & vbTab
            dataLine &= "MARKETING TEXT" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab ' PRODUCTNAME
        dataLine &= GetItemPrice(rowICTSTYL1, "C") ' PRICE

        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) > 0 Then
            dataLine &= "Yes" & vbTab  ' AVAIL
        Else
            dataLine &= "No" & vbTab  ' AVAIL
        End If

        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' URL
        dataLine &= STYLE_CODE & vbTab ' SKU
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' IMAGE URL

        dataLine &= GetPartnerCategory("PRICEGRAB", DEPT_CODE, ITEM_TYPE_CODE) & vbTab ' Category

        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab ' MARKETING TEXT

        Return (dataLine)

    End Function

    Private Function ProntoDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine &= "Title" & vbTab
            dataLine &= "SalePrice" & vbTab
            dataLine &= "URL" & vbTab
            dataLine &= "Description" & vbTab
            dataLine &= "Category" & vbTab
            dataLine &= "ImageURL" & vbTab
            dataLine &= "Condition" & vbTab
            dataLine &= "Brand" & vbTab
            dataLine &= "Keywords" & vbTab
            dataLine &= "ISBN" & vbTab
            dataLine &= "ArtistAuthor" & vbTab
            dataLine &= "ProductSKU" & vbTab
            dataLine &= "Outlet" & vbTab
            dataLine &= "InStock" & vbTab
            dataLine &= "ShippingCost" & vbTab
            dataLine &= "ShippingWeight" & vbTab
            dataLine &= "ZipCode" & vbTab
            dataLine &= "ProntoCategoryID" & vbTab
            dataLine &= "Other" & vbTab
            dataLine &= "ProductBid" & vbTab
            dataLine &= "RetailPrice" & vbTab
            dataLine &= "SpecialOffer" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        'Title = ICTSTYL1.STYLE_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab
        ' Sale_Price
        If rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty = "S" Then
            dataLine &= GetItemPrice(rowICTSTYL1, "S") & vbTab
        Else
            dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab
        End If
        '(PRODUCT_URL)
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab
        'Description = ICTSTYL1.STYLE_FULL_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab
        'Category = SOTPART2.PARTNER_CATEGORY
        dataLine &= GetPartnerCategory("PRONTO", DEPT_CODE, ITEM_TYPE_CODE) & vbTab
        'Image_Link(IMAGE_URL)
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        ' Condition
        dataLine &= "new" & vbTab
        'Brand = "WebUndies"
        dataLine &= "WebUndies" & vbTab
        ' Keywords
        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") & vbTab
        ' ISBN
        dataLine &= "" & vbTab
        ' ArtistAuthor
        dataLine &= "" & vbTab
        'SKU = STYLE_CODE
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab
        ' Outlet
        dataLine &= "" & vbTab
        ' In Stock
        dataLine &= "yes" & vbTab
        'Shipping_Cost = SOTPART1.PARTNER_PRODUCT_SHIP_RATE
        dataLine &= formatNumber(Val(GetPartnerAttribute("PRONTO", "PARTNER_PRODUCT_SHIP_RATE") & String.Empty)) & vbTab
        'ShippingWeight
        dataLine &= "" & vbTab
        ' Zip Code
        dataLine &= WHSE_ZIP_CODE & vbTab
        'Category ID= SOTPART2.PARTNER_CATEGORY_ID
        dataLine &= GetPartnerCategoryId("PRONTO", DEPT_CODE, ITEM_TYPE_CODE) & vbTab
        ' Other
        dataLine &= "" & vbTab
        ' ProductBid
        dataLine &= "" & vbTab
        ' Retail Price
        dataLine &= "" & vbTab
        ' Special Offer
        dataLine &= GetPartnerAttribute("PRONTO", "PARTNER_PRODUCT_PROMO_TEXT") & String.Empty & vbTab

        Return dataLine

    End Function

    Private Function ShopDotComDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine = "GROUP_NAME" & vbTab
            dataLine &= "GROUP_DESCRIPTION" & vbTab
            dataLine &= "LINE_ITEM_CODE" & vbTab
            dataLine &= "LINE_ITEM_NAME" & vbTab
            dataLine &= "LINE_ITEM_PRICE" & vbTab
            dataLine &= "IMAGE_URL" & vbTab
            dataLine &= "FIRST_LEVEL_DEPARTMENT" & vbTab
            dataLine &= "SECOND_LEVEL_DEPARTMENT" & vbTab
            dataLine &= "THIRD_LEVEL_DEPARTMENT" & vbTab
            dataLine &= "KEYWORDS" & vbTab
            dataLine &= "LINE_ITEM_SALE_PRICE" & vbTab
            dataLine &= "OPTION_TYPE" & vbTab
            dataLine &= "OPTION_VALUE1" & vbTab
            dataLine &= "PERMUTATION" & vbTab
            dataLine &= "PERMUTATION_ITEM_CODE" & vbTab
            dataLine &= "PERMUTATION_INVENTORY_STATUS" & vbTab
            dataLine &= "ALTERNATE_IMAGE_PROMPT" & vbTab
            dataLine &= "ALTERNATE_IMAGE_REFERENCE" & Environment.NewLine
        End If

        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) <= PARTNER_PRODUCT_INV_MIN Then
            Return String.Empty
        End If

        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & String.Empty & vbTab
        dataLine &= STYLE_CODE & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & String.Empty & vbTab
        dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & vbTab

        dataLine &= GetPartnerCategory("SHOP", DEPT_CODE, ITEM_TYPE_CODE) & vbTab

        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") & vbTab
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab
        dataLine &= "Size" & vbTab

        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If

        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab
        dataLine &= "0" & vbTab

        If rowICTSTYL1.Item("STYLE_IMAGE_OTHER1") & String.Empty <> String.Empty Then
            dataLine &= "Additional images" & vbTab
            dataLine &= imageMediaDir & rowICTSTYL1.Item("STYLE_IMAGE_OTHER1")
        Else
            dataLine &= vbTab
        End If

        Return dataLine

    End Function

    Private Function ShoppingDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine = "Unique Merchant SKU" & vbTab
            dataLine &= "MPN/ISBN" & vbTab
            dataLine &= "UPC" & vbTab
            dataLine &= "Manufacturer" & vbTab
            dataLine &= "Product Name" & vbTab
            dataLine &= "Product URL" & vbTab
            dataLine &= "Mobile URL" & vbTab
            dataLine &= "Current Price" & vbTab
            dataLine &= "Original Price" & vbTab
            dataLine &= "Category ID" & vbTab
            dataLine &= "Category" & vbTab
            'dataLine &= "Sub-Category" & vbTab
            dataLine &= "Parent SKU" & vbTab
            dataLine &= "Parent Name" & vbTab
            dataLine &= "Product Description" & vbTab
            dataLine &= "Stock Description" & vbTab
            dataLine &= "Product Bullet Point 1" & vbTab
            dataLine &= "Product Bullet Point 2" & vbTab
            dataLine &= "Product Bullet Point 3" & vbTab
            dataLine &= "Product Bullet Point 4" & vbTab
            dataLine &= "Product Bullet Point 5" & vbTab
            dataLine &= "Image URL" & vbTab
            dataLine &= "Alternative Image URL 1" & vbTab
            dataLine &= "Alternative Image URL 2" & vbTab
            dataLine &= "Alternative Image URL 3" & vbTab
            dataLine &= "Alternative Image URL 4" & vbTab
            dataLine &= "Alternative Image URL 5" & vbTab
            dataLine &= "Product Type" & vbTab
            'dataLine &= "Style" & vbTab
            dataLine &= "Condition" & vbTab
            dataLine &= "Gender" & vbTab
            'dataLine &= "Department" & vbTab
            dataLine &= "Age Range" & vbTab
            dataLine &= "Color" & vbTab
            dataLine &= "Material" & vbTab
            'dataLine &= "Format" & vbTab
            dataLine &= "Team" & vbTab
            dataLine &= "League" & vbTab
            dataLine &= "Fan Gear Type" & vbTab
            'dataLine &= "Platform" & vbTab
            'dataLine &= "Software Type" & vbTab
            'dataLine &= "Watch Display Type" & vbTab
            'dataLine &= "Phone Type" & vbTab
            'dataLine &= "Cell Phone Service Provider" & vbTab
            'dataLine &= "Cell Phone Plan Type" & vbTab
            'dataLine &= "Cell Phone Usage Profile" & vbTab
            dataLine &= "Size" & vbTab
            dataLine &= "Size Unit of Measure" & vbTab
            dataLine &= "Product Length" & vbTab
            dataLine &= "Length Unit of Measure" & vbTab
            dataLine &= "Product Width" & vbTab
            dataLine &= "Width Unit of Measure" & vbTab
            dataLine &= "Product Height" & vbTab
            dataLine &= "Height Unit of Measure" & vbTab
            dataLine &= "Product Weight" & vbTab
            dataLine &= "Weight Unit of Measure" & vbTab
            dataLine &= "Unit Price" & vbTab
            dataLine &= "Top Seller Rank" & vbTab
            dataLine &= "Product Launch Date" & vbTab
            dataLine &= "Stock Availability" & vbTab
            dataLine &= "Shipping Rate" & vbTab
            dataLine &= "Shipping Weight" & vbTab
            dataLine &= "Zip Code" & vbTab
            dataLine &= "Estimated Ship Date" & vbTab
            dataLine &= "Coupon Code" & vbTab
            dataLine &= "Coupon Code Description" & vbTab
            dataLine &= "Merchandising Type" & vbTab
            dataLine &= "Bundle" & vbTab
            dataLine &= "Related Products" & Environment.NewLine
        End If

        dataLine &= rowICTSTYL1.Item("ITEM_CODE") & vbTab ' SKU
        dataLine &= vbTab ' MPN or ISBN
        dataLine &= vbTab ' UPC / EAN
        dataLine &= "WebUndies" & vbTab ' Brand
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_XML") & vbTab ' Product Name
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab ' URL
        dataLine &= vbTab ' Mobile URL

        ' Current Price
        dataLine &= GetItemPrice(rowICTSTYL1, "C") & vbTab

        ' Original Price
        dataLine &= GetItemPrice(rowICTSTYL1, "M") & vbTab
        dataLine &= "31515" & vbTab ' Category ID
        dataLine &= "Clothing and Accessories > Clothing" & vbTab ' Category 
        'dataLine &= "Clothing" & vbTab ' Sub Category Name - not valid after 9/1/11
        dataLine &= STYLE_CODE & vbTab ' Parent SKU

        dataLine &= rowICTSTYL1.Item("STYLE_DESC" & String.Empty) & vbTab ' Parent Name
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC" & String.Empty) & vbTab ' Product Desc

        dataLine &= PARTNER_PRODUCT_PROMO_TEXT & vbTab ' Stock Desc
        'as per Shopping.com STOCK DESCRIPTION 	This field should be used for promtional text like free shipping on orders over $100 etc...

        ' Prod Bullet Point 1
        dataLine &= GetCareInstructions(rowICTSTYL1) & vbTab
        'Product Bullet Point 2
        dataLine &= GetFeatures(STYLE_CODE) & vbTab

        ' Product Bullet Point 3
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Trim
        If temp.Length > 0 Then
            temp = "Licensed by: " & temp
            temp = temp.Replace("<sup>TM</sup>", " TM ")
            temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
            temp = temp.Trim
            dataLine &= temp
        End If
        dataLine &= vbTab

        dataLine &= vbTab ' Product Bullet Point 4
        dataLine &= vbTab ' Product Bullet Point 5

        ' Alt Image 1, 2
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab ' URL

        For Each altImage As String In New String() {"STYLE_IMAGE_OTHER1", "STYLE_IMAGE_OTHER2"}
            If rowICTSTYL1.Item(altImage) & String.Empty <> String.Empty Then
                dataLine &= imageMediaDir & rowICTSTYL1.Item(altImage).ToString.ToLower & vbTab ' URL
            Else
                dataLine &= vbTab
            End If
        Next
        dataLine &= vbTab ' Alt Image 3
        dataLine &= vbTab ' Alt Image 4
        dataLine &= vbTab ' Alt Image 5

        ' Product Type
        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & String.Empty
        dataLine &= vbTab

        ' dataLine &= vbTab ' Style - not valid after 9/1/11
        dataLine &= "New" & vbTab ' Condition

        ' Gender
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "MENS"
                dataLine &= "Men"
            Case "WOMENS"
                dataLine &= "Women"
            Case Else
                dataLine &= rowICTSTYL1.Item("DEPT_DESC") & String.Empty
        End Select
        dataLine &= vbTab

        'dataLine &= vbTab ' Department - not valid as of 9/1/11

        ' Age Range
        Select Case (rowICTSTYL1.Item("DEPT_CODE") & String.Empty).ToString.Trim.ToUpper
            Case "BOYS", "GIRLS"
                dataLine &= "Youth" & vbTab
            Case "MENS", "WOMENS"
                dataLine &= "Adult" & vbTab
            Case Else
                dataLine &= vbTab
        End Select

        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab ' color

        ' Material
        dataLine &= rowICTSTYL1.Item("MATL_DESC") & String.Empty
        dataLine &= vbTab

        'dataLine &= vbTab ' Format
        dataLine &= vbTab ' Team
        dataLine &= vbTab ' League
        dataLine &= vbTab ' Fan Gear Type
        'dataLine &= vbTab ' Platform
        'dataLine &= vbTab ' Software Type
        'dataLine &= vbTab ' Watch Display True
        'dataLine &= vbTab ' Phone Type
        'dataLine &= vbTab ' Cell Phone Service provider
        'dataLine &= vbTab ' Cell Phone Plan Type
        'dataLine &= vbTab ' Cell Phone usage Profile

        dataLine &= rowICTSTYL1.Item("SIZE_CODE") & String.Empty & vbTab

        dataLine &= vbTab ' Size Unit of Measure
        dataLine &= vbTab ' Product Length
        dataLine &= vbTab ' Length Unit of Measure
        dataLine &= vbTab ' Product Width
        dataLine &= vbTab ' Width UOM
        dataLine &= vbTab ' Prod height
        dataLine &= vbTab ' Height UOM   
        dataLine &= vbTab ' Prod weight
        dataLine &= vbTab ' Weight UOM   
        dataLine &= vbTab ' Unit Price
        dataLine &= vbTab ' Top Seller Rank
        dataLine &= vbTab ' Prod Launch date

        ' Stock Availability
        If GetInventoryLevel(STYLE_CODE, COLOR_CODE, SIZE_CODE) > 0 Then
            dataLine &= "Yes" & vbTab
        Else
            dataLine &= "No" & vbTab
        End If

        ' Shipping Rate
        dataLine &= formatNumber(PARTNER_PRODUCT_SHIP_RATE) & vbTab
        dataLine &= vbTab ' Shipping weight

        ' Zip Code
        dataLine &= WHSE_ZIP_CODE & vbTab

        dataLine &= vbTab ' Est Ship Date  
        dataLine &= vbTab ' Coupon Code
        dataLine &= vbTab ' Coupon Code Desc  
        dataLine &= vbTab ' Merchandising Type
        dataLine &= vbTab ' Bundle
        dataLine &= String.Empty  ' Related Products

        Return dataLine
    End Function

    Private Function SLIDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim rowSOTPART3 As DataRow = Nothing

        If includeHeader Then
            dataLine &= "Style" & vbTab ' STYLE_CODE
            dataLine &= "Title" & vbTab ' STYLE_DESC
            dataLine &= "Description" & vbTab ' STYLE_FULL_DESC
            dataLine &= "KeyWords" & vbTab ' STYLE_ADDL_KEYWORDS
            dataLine &= "Department" & vbTab ' DEPT_DESC
            dataLine &= "ItemType" & vbTab ' ITEM_TYPE_DESC
            dataLine &= "Material" & vbTab ' MATL_DESC
            dataLine &= "License" & vbTab ' LICENSE_DESC
            dataLine &= "Price" & vbTab ' STYLE_PRICE
            dataLine &= "SalePrice" & vbTab ' SALE_PRICE
            dataLine &= "Color" & vbTab ' STYLE_COLOR_DESC
            dataLine &= "Brand" & vbTab  ' BRAND_DESC
            dataLine &= "ImageUrl" & vbTab
            dataLine &= "ProductUrl" & vbTab
            dataLine &= "ProductCategory" & vbTab
            dataLine &= "Size" & vbTab ' SIZE
            dataLine &= "Sku" & Environment.NewLine ' ITEM_CODE
        End If

        ' Only uses Styles; therefore, only send style once
        'If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
        '    Return String.Empty
        'End If

        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_DESC") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("STYLE_ADDL_KEYWORDS") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("ITEM_TYPE_DESC") & String.Empty & vbTab
        dataLine &= rowICTSTYL1.Item("MATL_DESC") & String.Empty & vbTab

        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= temp & vbTab

        dataLine &= GetItemPrice(rowICTSTYL1, "O") & String.Empty & vbTab
        ' SALEPRICE
        If rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty = "S" Then
            dataLine &= GetItemPrice(rowICTSTYL1, "S") & vbTab
        Else
            dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab
        End If
        dataLine &= rowICTSTYL1.Item("STYLE_COLOR_DESC") & String.Empty & vbTab

        Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", rowICTSTYL1.Item("BRAND_CODE") & String.Empty)

        If rowICTBRAN1 IsNot Nothing AndAlso (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim.Length > 0 Then
            dataLine &= (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim & vbTab
        Else
            dataLine &= "WebUndies" & vbTab
        End If

        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab

        dataLine &= GetPartnerCategoryExtensions(STYLE_CODE, "PAGE_NAME", "|", " >") & vbTab ' Category

        ' size
        If dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'").Length > 0 Then
            rowSOTPART3 = dst.Tables("SOTPART3").Select("PARTNER_CODE = '" & PARTNER_CODE & "' AND SIZE_CODE = '" & rowICTSTYL1.Item("SIZE_CODE") & "'")(0)
        End If
        If rowSOTPART3 IsNot Nothing Then
            dataLine &= rowSOTPART3.Item("PARTNER_SIZE_DESC") & vbTab
        Else
            dataLine &= rowICTSTYL1.Item("SIZE_DESC") & vbTab
        End If
        dataLine &= rowICTSTYL1.Item("ITEM_CODE")  ' Sku

        Return dataLine

    End Function

    Private Function TheFindDataFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty

        If includeHeader Then
            dataLine &= "Title" & vbTab
            dataLine &= "Description" & vbTab
            dataLine &= "Image_Link" & vbTab
            dataLine &= "Page_URL" & vbTab
            dataLine &= "Price" & vbTab
            dataLine &= "SKU" & vbTab
            dataLine &= "Sale" & vbTab
            dataLine &= "Sale_Price" & vbTab
            dataLine &= "Shipping_Cost" & vbTab
            dataLine &= "Online_Only" & vbTab
            dataLine &= "Brand" & vbTab
            dataLine &= "Condition" & vbTab
            dataLine &= "Categories" & vbTab
            dataLine &= "Department" & Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        'Title = ICTSTYL1.STYLE_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_DESC_EXT") & vbTab
        'Description = ICTSTYL1.STYLE_FULL_DESC
        dataLine &= rowICTSTYL1.Item("STYLE_FULL_DESC") & vbTab
        'Image_Link(IMAGE_URL)
        dataLine &= imageMediaDir & STYLE_CODE.ToLower & ".jpg" & vbTab
        'Page_URL(PRODUCT_URL)
        dataLine &= siteURL & STYLE_CODE.ToLower & ".htm" & vbTab
        'Price = STYLE_PRICE_TYPE = "C" or "S" use STYLE_PRICE_SALE else use STYLE_PRICE_OUR "O"
        dataLine &= GetItemPrice(rowICTSTYL1, "O") & vbTab
        'SKU = STYLE_CODE
        dataLine &= rowICTSTYL1.Item("STYLE_CODE") & vbTab

        ' Sale / Sale_Price
        If rowICTSTYL1.Item("STYLE_PRICE_TYPE") & String.Empty = "S" Then
            dataLine &= "Yes" & vbTab
            dataLine &= GetItemPrice(rowICTSTYL1, "S") & vbTab
        Else
            dataLine &= "" & vbTab
            dataLine &= "" & vbTab
        End If

        'Shipping_Cost = SOTPART1.PARTNER_PRODUCT_SHIP_RATE
        dataLine &= formatNumber(Val(GetPartnerAttribute("THEFIND", "PARTNER_PRODUCT_SHIP_RATE") & String.Empty)) & vbTab
        ' Online_Only
        dataLine &= "Yes" & vbTab
        'Brand = "WebUndies"
        dataLine &= "WebUndies" & vbTab
        'Condition = "New"
        dataLine &= "New" & vbTab
        'Category = SOTPART2.PARTNER_CATEGORY
        dataLine &= GetPartnerCategory("THEFIND", DEPT_CODE, ITEM_TYPE_CODE) & vbTab
        'Department = ICTDEPT1.DEPT_DESC
        dataLine &= rowICTSTYL1.Item("DEPT_DESC") & String.Empty

        Return dataLine

    End Function

    Private Function UpdatePartner(ByVal PARTNER_CODE As String)
        Try

            If dst.Tables("SOTPART1").Select("ISNULL(SEL, '0') = '1' AND PARTNER_CODE = '" & PARTNER_CODE & "'").Length > 0 Then
                dst.Tables("SOTPART1").Select("ISNULL(SEL, '0') = '1' AND PARTNER_CODE = '" & PARTNER_CODE & "'")(0).Item("PROCESSED") = "1"
                dst.Tables("SOTPART1").Select("ISNULL(SEL, '0') = '1' AND PARTNER_CODE = '" & PARTNER_CODE & "'")(0).Item("PARTNER_PRODUCT_LAST_EXTRACT") = DateTime.Now
            End If
            ASCDATA1.ExecuteSQL("UPDATE SOTPART1 SET PARTNER_PRODUCT_LAST_EXTRACT = SYSDATE WHERE PARTNER_CODE = :PARM1", "V", New String() {PARTNER_CODE})
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function ShareASaleFeed(ByVal rowICTSTYL1 As DataRow, ByVal includeHeader As Boolean) As String

        Dim dataLine As String = String.Empty
        Dim temp As String = String.Empty
        Dim delimeter As String = Chr(34) & ","

        ' No header record. legt code incase it changes
        If includeHeader AndAlso 1 = 2 Then
            dataLine = Chr(34) & "SKU" & delimeter
            dataLine &= Chr(34) & "Name" & delimeter
            dataLine &= Chr(34) & "URL to product" & delimeter
            dataLine &= Chr(34) & "Price" & delimeter
            dataLine &= Chr(34) & "Retail Price" & delimeter
            dataLine &= Chr(34) & "URL to image" & delimeter
            dataLine &= Chr(34) & "URL to thumbnail image" & delimeter
            dataLine &= Chr(34) & "Commission" & delimeter
            dataLine &= Chr(34) & "Category" & delimeter
            dataLine &= Chr(34) & "SubCategory" & delimeter
            dataLine &= Chr(34) & "Description" & delimeter
            dataLine &= Chr(34) & "SearchTerms" & delimeter
            dataLine &= Chr(34) & "Status" & delimeter
            dataLine &= Chr(34) & "Your MerchantID" & delimeter
            dataLine &= Chr(34) & "Custom 1" & delimeter
            dataLine &= Chr(34) & "Custom 2" & delimeter
            dataLine &= Chr(34) & "Custom 3" & delimeter
            dataLine &= Chr(34) & "Custom 4" & delimeter
            dataLine &= Chr(34) & "Custom 5" & delimeter
            dataLine &= Chr(34) & "Manufacturer" & delimeter
            dataLine &= Chr(34) & "PartNumber" & delimeter
            dataLine &= Chr(34) & "MerchantCategory" & delimeter
            dataLine &= Chr(34) & "MerchantSubcategory" & delimeter
            dataLine &= Chr(34) & "ShortDescription" & delimeter
            dataLine &= Chr(34) & "ISBN" & delimeter
            dataLine &= Chr(34) & "UPC" & delimeter
            dataLine &= Chr(34) & "CrossSell" & delimeter
            dataLine &= Chr(34) & "MerchantGroup" & delimeter
            dataLine &= Chr(34) & "MerchantSubgroup" & delimeter
            dataLine &= Chr(34) & "CompatibleWith" & delimeter
            dataLine &= Chr(34) & "CompareTo" & delimeter
            dataLine &= Chr(34) & "QuantityDiscount" & delimeter
            dataLine &= Chr(34) & "Bestseller" & delimeter
            dataLine &= Chr(34) & "AddToCartURL" & delimeter
            dataLine &= Chr(34) & "ReviewsRSSURL" & delimeter
            dataLine &= Chr(34) & "Option1" & delimeter
            dataLine &= Chr(34) & "Option2" & delimeter
            dataLine &= Chr(34) & "Option3" & delimeter
            dataLine &= Chr(34) & "Option4" & delimeter
            dataLine &= Chr(34) & "Option5" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & delimeter
            dataLine &= Chr(34) & "ReservedForFutureUse" & Chr(34)
            dataLine &= Environment.NewLine
        End If

        ' Only uses Styles; therefore, only send style once
        If PREVIOUS_STYLE_CODE = rowICTSTYL1.Item("STYLE_CODE") & String.Empty Then
            Return String.Empty
        End If

        'SKU = STYLE_CODE
        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_CODE") & delimeter
        'Name = ICTSTYL1.STYLE_DESC
        dataLine &= Chr(34) & (rowICTSTYL1.Item("STYLE_DESC") & String.Empty).ToString.Replace(Chr(34), "'") & delimeter
        'URL to product (PRODUCT_URL)
        dataLine &= Chr(34) & siteURL & STYLE_CODE.ToLower & ".htm" & delimeter
        'Price = STYLE_PRICE_TYPE = "C" or "S" use STYLE_PRICE_SALE else use STYLE_PRICE_OUR
        dataLine &= Chr(34) & GetItemPrice(rowICTSTYL1, "C") & delimeter
        'Retail Price = STYLE_PRICE_TYPE = "C" or "S" use STYLE_PRICE_SALE else use STYLE_PRICE_OUR
        dataLine &= Chr(34) & "" & delimeter
        'URL to image ImageURL(IMAGE_URL)
        dataLine &= Chr(34) & imageMediaDir & STYLE_CODE.ToLower & ".jpg" & delimeter
        'URL to thumbnail image
        dataLine &= Chr(34) & "" & delimeter
        'Commission
        dataLine &= Chr(34) & "" & delimeter
        'Category
        dataLine &= Chr(34) & "8" & delimeter
        'SUBCategory = SOTPART2.PARTNER_CATEGORY
        dataLine &= Chr(34) & GetPartnerCategoryId("SAS", DEPT_CODE, ITEM_TYPE_CODE) & delimeter
        'Description = ICTSTYL1.STYLE_FULL_DESC
        dataLine &= Chr(34) & (rowICTSTYL1.Item("STYLE_FULL_DESC") & String.Empty).ToString.Replace(Chr(34), "'") & delimeter

        'Search Terms
        Dim searchData As String = String.Empty
        For Each searchTerm As String In (rowICTSTYL1("STYLE_ADDL_KEYWORDS") & String.Empty).ToString.Trim.Split(",")
            searchTerm = searchTerm.Trim.Replace(Chr(34), "")
            If searchTerm.Length > 0 AndAlso searchData.Length + 1 + searchTerm.Length < 255 Then
                searchData &= "," & searchTerm
            End If
        Next

        If searchData.Length > 0 Then
            searchData = searchData.Substring(1).Trim
        End If
        dataLine &= Chr(34) & searchData & delimeter

        ' Status
        dataLine &= Chr(34) & "instock" & delimeter
        'Merchant ID
        dataLine &= Chr(34) & PARTNER_PRODUCT_OUR_ID & delimeter

        ' Custom 1 - License
        temp = rowICTSTYL1.Item("LICENSE_DESC") & String.Empty
        temp = temp.Replace("<sup>TM</sup>", " TM ")
        temp = temp.Replace("<sup>", String.Empty).Replace("</sup>", String.Empty)
        temp = temp.Trim
        dataLine &= Chr(34) & temp.Replace(Chr(34), "") & delimeter

        ' Custom 2 - material
        dataLine &= Chr(34) & rowICTSTYL1.Item("MATL_DESC") & delimeter

        ' Custom 3 - Color
        dataLine &= Chr(34) & rowICTSTYL1.Item("STYLE_COLOR_DESC") & delimeter

        ' Custom 4 - Care
        temp = String.Empty
        For Each field As String In New String() {"STYLE_CARE_MW", "STYLE_CARE_EC", "STYLE_CARE_CW", "STYLE_CARE_DC"}
            If rowICTSTYL1.Item(field) & String.Empty = "1" Then
                Select Case field
                    Case "STYLE_CARE_MW" : temp &= ", " & "Machine Washable"
                    Case "STYLE_CARE_EC" : temp &= ", " & "Easy Care"
                    Case "STYLE_CARE_CW" : temp &= ", " & "Cold Water"
                    Case "STYLE_CARE_DC" : temp &= ", " & "Delicate Cycle"
                End Select
            End If
        Next
        If temp.Length > 0 Then
            temp = temp.Substring(1).Trim
        End If
        dataLine &= Chr(34) & temp & delimeter

        ' Custom  5
        dataLine &= Chr(34) & "" & delimeter

        'Manufacturer = "WebUndies"
        Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", rowICTSTYL1.Item("BRAND_CODE") & String.Empty)
        If rowICTBRAN1 IsNot Nothing AndAlso (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim.Length > 0 Then
            dataLine &= Chr(34) & (rowICTBRAN1.Item("BRAND_DESC") & String.Empty).ToString.Trim & delimeter ' brand
        Else
            dataLine &= Chr(34) & "WebUndies" & delimeter ' brand
        End If

        'PartNumber
        dataLine &= Chr(34) & "" & delimeter
        'MerchantCategory
        dataLine &= Chr(34) & rowICTSTYL1.Item("DEPT_DESC") & delimeter
        'MerchantSubcategory
        dataLine &= Chr(34) & rowICTSTYL1.Item("ITEM_TYPE_DESC") & delimeter
        'ShortDescription
        dataLine &= Chr(34) & "" & delimeter
        'ISBN
        dataLine &= Chr(34) & "" & delimeter
        'UPC
        dataLine &= Chr(34) & "" & delimeter
        'CrossSell
        dataLine &= Chr(34) & "" & delimeter
        'MerchantGroup
        dataLine &= Chr(34) & "" & delimeter
        'MerchantSubgroup
        dataLine &= Chr(34) & "" & delimeter
        'CompatibleWith
        dataLine &= Chr(34) & "" & delimeter
        'CompareTo
        dataLine &= Chr(34) & "" & delimeter
        'QuantityDiscount
        dataLine &= Chr(34) & "" & delimeter
        'Bestseller
        dataLine &= Chr(34) & "" & delimeter
        'AddToCartURL
        dataLine &= Chr(34) & "" & delimeter
        'ReviewsRSSURL
        dataLine &= Chr(34) & "" & delimeter
        'Option1
        dataLine &= Chr(34) & "" & delimeter
        'Option2
        dataLine &= Chr(34) & "" & delimeter
        'Option3
        dataLine &= Chr(34) & "" & delimeter
        'Option4
        dataLine &= Chr(34) & "" & delimeter
        'Option5
        dataLine &= Chr(34) & "" & delimeter

        'ReservedForFutureUse 1 - 9
        For res As Integer = 1 To 9
            dataLine &= Chr(34) & "" & delimeter
        Next
        ' ReservedForFutureUse 10
        dataLine &= Chr(34) & "" & Chr(34)

        Return dataLine

    End Function


#End Region

End Class
