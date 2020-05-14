
Imports System.Xml
Imports System.IO

Public Class WBCNXTPN
    Public Property Feedready As Boolean
    Public Property XmlDoc As New XmlDocument
    Private _tblWBTSTYLD As DataTable
    Private _Classes As New DataTable
    Private _Atributes As New DataTable
    Private _Colors As New DataTable
    Private _Styles As New DataTable
    Private _Discounts As DataTable
    Private _useDemo As Boolean = False
    Private _HTTPAddress_demo As String = "http://216.38.11.230"
    Private _HTTPAddress As String = "https://www.regency-rib.com/shop"

#Region "Contructor"
    Public Sub New(ByVal tblWBTSTYLD As DataTable, Optional ByVal useDemo As Boolean = False)
        _tblWBTSTYLD = tblWBTSTYLD
        _useDemo = useDemo
        'If _useDemo Then
        '    _HTTPAddress = _HTTPAddress_demo
        'End If
        Feedready = False
        createClasses()
        createAttributes()
        createColors()
        createStyles()
        createDiscounts()
        CreateFeedBody()
        Feedready = True
    End Sub

    Private Sub createAttributes()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT L3.STYLE_CODE, A1.ATTR_DESC, NVL(A1.ATT_RANK,999) AS ATT_RANK")
        sql.AppendLine("FROM ICTSTYL3 L3, ICTATTR1 A1")
        sql.AppendLine("WHERE L3.ATTR_CODE = A1.ATTR_CODE")
        _Atributes = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
    End Sub

    Private Sub createColors()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        'sql.AppendLine("SELECT")
        'sql.AppendLine("ICTSTYC1.STYLE_CODE,")
        'sql.AppendLine("ICTSTYC1.COLOR_CODE,")
        'sql.AppendLine("ICTCOLR1.COLOR_DESC")
        'sql.AppendLine("FROM ICTSTYC1, ICTCOLR1, ICTSTAT2")
        'sql.AppendLine("WHERE ICTSTYC1.COLOR_CODE = ICTCOLR1.COLOR_CODE")
        'sql.AppendLine("AND ICTSTYC1.STYLE_CODE = ICTSTAT2.STYLE_CODE (+)")
        'sql.AppendLine("AND ICTSTYC1.COLOR_CODE = ICTSTAT2.COLOR_CODE (+)")
        'sql.AppendLine("AND (ICTSTYC1.STYLE_COLOR_STATUS = 'A' OR ICTSTAT2.WHSE_QTY_ON_HAND > 0)")
        'sql.AppendLine("AND ICTSTAT2.WHSE_CODE = 'MS'")
        'sql.AppendLine("GROUP BY")
        'sql.AppendLine("ICTSTYC1.STYLE_CODE,")
        'sql.AppendLine("ICTSTYC1.COLOR_CODE,")
        'sql.AppendLine("ICTCOLR1.COLOR_DESC")
        sql.AppendLine("SELECT")
        sql.AppendLine("WD.STYLE_CODE,")
        sql.AppendLine("WD.COLOR_CODE,")
        sql.AppendLine("C1.COLOR_DESC")
        sql.AppendLine("FROM WBTSTYLD WD, ICTCOLR1 C1")
        sql.AppendLine("WHERE WD.COLOR_CODE = C1.COLOR_CODE")
        _Colors = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
        For Each rowColors As DataRow In _Colors.Rows
            Dim STYLE_CODE As String = rowColors.Item("STYLE_CODE").ToString & String.Empty
            Dim COLOR_CODE As String = rowColors.Item("COLOR_CODE").ToString & String.Empty
            Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
            Dim rowWBTSTYLD As DataRow = _tblWBTSTYLD.Select(filter).FirstOrDefault
            If Not IsNothing(rowWBTSTYLD) Then
                If Val(rowWBTSTYLD.Item("CURR_ON_HAND").ToString & String.Empty) <= 0 Then
                    rowColors.Delete()
                End If
            End If
        Next
    End Sub

    Private Sub createClasses()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT * FROM ICTCLAS1")
        _Classes = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
    End Sub

    Private Sub createDiscounts()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT * FROM ICTDISC1")
        _Discounts = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
    End Sub

    Private Sub createStyles()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT * FROM ICTSTYL1")
        _Styles = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
    End Sub
#End Region

#Region "Custom Methods"
    Private Sub CreateFeedBody()
        ' to create the XmlDocument... '
        'XmlDoc.AppendChild(XmlDoc.CreateXmlDeclaration("1.0", "UTF-8", String.Empty))
        'XmlDoc.PreserveWhitespace = True

        Dim rssElement As System.Xml.XmlElement = XmlDoc.CreateElement("rss")
        XmlDoc.AppendChild(rssElement)

        Dim channelElement As System.Xml.XmlElement = XmlDoc.CreateElement("channel")
        channelElement.AppendChild(makeChannelNode("title", "Regency International"))
        channelElement.AppendChild(makeChannelNode("link", _HTTPAddress))
        channelElement.AppendChild(makeChannelNode("description", "Product feed for Nextopia"))
        Dim LiveMode As String = "W"
        Dim filter As String = String.Format("WEB_IND = '{0}'", LiveMode)
        For Each dr As DataRow In _tblWBTSTYLD.Select(filter)
            ASCMAIN1.Progress("Creating Nextopia Feed", dr.Item("STYLE_CODE").ToString)
            Dim itemElement As XmlNode = makeItemNode(dr)
            channelElement.AppendChild(itemElement)
        Next
        ASCMAIN1.Progress("", "")

        rssElement.AppendChild(channelElement)
        'rssElement.AppendChild(n1Element.CloneNode(deep:=True))

        ' to update the XmlDocument (simple example)... '
        'Dim s1Element As Xml.XmlElement = xmlDoc.SelectSingleNode("foo/n1/s1")
        'If Not s1Element Is Nothing Then s1Element.InnerText = "some value"

    End Sub

    Private Function GetWEB_DESC(ByRef rowWBTSTYLD As DataRow) As String
        Dim RetVal As String = ""
        Dim STYLE_CODE As String = rowWBTSTYLD.Item("STYLE_CODE").ToString & String.Empty
        Dim SQLS As New System.Text.StringBuilder With {.Length = 0}
        SQLS.AppendLine(String.Format("SELECT NVL(WEB_DESC,'') AS WEB_DESC FROM WBTSTYLH WHERE STYLE_CODE = '{0}'", STYLE_CODE))
        ASCMAIN1.sql = SQLS.ToString()
        Dim WEB_DESC As String = ASCDATA1.GetDataValue
        If WEB_DESC.Length > 0 Then
            RetVal = WEB_DESC
        Else
            RetVal = rowWBTSTYLD.Item("STYLE_DESC") & String.Empty
        End If
        Return RetVal
    End Function

    Private Function makeChannelNode(ByVal title As String, ByVal innerText As String) As XmlNode
        Dim retVal As XmlNode = XmlDoc.CreateElement(title)
        retVal.InnerText = innerText
        Return retVal
    End Function

    Private Function makeItemNode(ByVal rowWBTSTYLD As DataRow) As XmlNode
        Dim rowARTCUST1 As DataRow = Nothing
        Dim BASE As New ASFBASE0
        Dim webDesc As String = GetWEB_DESC(rowWBTSTYLD)
        Dim Discounts As List(Of DISCOUNTS)
        Discounts = NTP_Price_Discounts(rowWBTSTYLD.Item("STYLE_CODE"))

        Dim itemElement As System.Xml.XmlElement = XmlDoc.CreateElement("item")
        itemElement.AppendChild(makeChannelNode("id", rowWBTSTYLD.Item("STYLE_CODE") & "-" & rowWBTSTYLD.Item("COLOR_CODE")))
        itemElement.AppendChild(makeChannelNode("title", webDesc))
        itemElement.AppendChild(makeChannelNode("link", _HTTPAddress & "/" & rowWBTSTYLD.Item("STYLE_CODE") & "-" & rowWBTSTYLD.Item("COLOR_CODE") & ".html"))
        itemElement.AppendChild(makeChannelNode("image_link", _HTTPAddress & "/media/product/" & Replace(rowWBTSTYLD.Item("DEFAULT_IMAGE"), ".JPG", ".jpg")))
        itemElement.AppendChild(makeChannelNode("price", Discounts.Item(3).DISCOUNT_PRICE))
        itemElement.AppendChild(makeChannelNode("availability", Val(rowWBTSTYLD.Item("CURR_ON_HAND").ToString & String.Empty)))
        itemElement.AppendChild(makeChannelNode("description", webDesc))
        itemElement.AppendChild(makeChannelNode("product_type", "Home > " & makeCategory(rowWBTSTYLD.Item("STYLE_CLASS_CODE"))))
        itemElement.AppendChild(makeAttributesNode(rowWBTSTYLD.Item("STYLE_CODE"), True))
        itemElement.AppendChild(makeColorsNode(rowWBTSTYLD.Item("STYLE_CODE"), rowWBTSTYLD.Item("COLOR_CODE")))
        itemElement.AppendChild(makeChannelNode("category", makeCategory(rowWBTSTYLD.Item("STYLE_CLASS_CODE"))))
        itemElement.AppendChild(makeAttributesNode(rowWBTSTYLD.Item("STYLE_CODE"), False))
        Return itemElement
    End Function

    Private Function NTP_Price_Discounts(STYLE_CODE As String) As List(Of DISCOUNTS)
        Dim retval As New List(Of DISCOUNTS)
        Dim DiscPromoFound As Boolean = False
        Const DiscPromoPct As Double = 70
        Dim DiscPromoDesc As String = ""
        Dim rowICTSTYL1 As DataRow = _Styles.Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE)).FirstOrDefault
        Dim STYLE_STATUS As String = rowICTSTYL1.Item("STYLE_STATUS") & ""
        Dim STYLE_CLASS_CODE As String = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        If STYLE_CLASS_CODE = "" Then 'We have to protect against this somehow.
            STYLE_CLASS_CODE = "PVC"
        End If
        Dim STYLE_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
        Dim STYLE_PROMO_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PROMO_PRICE") & "")
        Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
        Dim INNER_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")
        Dim MSOQ As Int64 = Val(rowICTSTYL1.Item("STYLE_SO_QTY_MIN") & "")
        Dim HALFCASE As Integer = 0
        If INNER_PACK_QTY > 0 Then
            HALFCASE = Math.Ceiling((CARTON_PACK_QTY / 2) / INNER_PACK_QTY) * INNER_PACK_QTY
        Else
            HALFCASE = (CARTON_PACK_QTY / 2)
        End If
        Dim rowICTCLAS1 As DataRow = _Classes.Select(String.Format("STYLE_CLASS_CODE = '{0}'", STYLE_CLASS_CODE)).FirstOrDefault
        Dim IsPVC As Boolean = rowICTCLAS1.Item("DISC_CODE").ToString = "PVC"

        If STYLE_STATUS = "D" Then
            DiscPromoFound = True
            DiscPromoDesc = "Disc"
        ElseIf STYLE_PROMO_PRICE <> 0 Then
            DiscPromoFound = True
            DiscPromoDesc = "Promo"
        End If

        If CARTON_PACK_QTY = 0 And INNER_PACK_QTY = 0 Then
            MsgBox("Box & Carton Qty Set To Zero", vbOKOnly, "Style Attributes Problem")
            CARTON_PACK_QTY = 1
            INNER_PACK_QTY = 1
        End If

        If rowICTCLAS1 IsNot Nothing Then
            Dim DISC_CODE As String = rowICTCLAS1.Item("DISC_CODE") & ""
            Dim rowICTDISC1 As DataRow = _Discounts.Select(String.Format("DISC_CODE = '{0}'", DISC_CODE)).FirstOrDefault
            Dim HAlfCaseIsOne As Boolean = False
            If rowICTDISC1 IsNot Nothing Then
                For I As Integer = 1 To 4
                    Dim DISC As New DISCOUNTS
                    Dim CASES As Decimal = Val(rowICTDISC1.Item(String.Format("DISC{0}_CASES", CStr(I))) & "")
                    Dim PCT As Decimal = Val(rowICTDISC1.Item(String.Format("DISC{0}_PCT", CStr(I))) & "")
                    If IsPVC Then
                        If I = 4 And HAlfCaseIsOne Then
                            DISC.DISCOUNT_QTY = 0
                        Else
                            If CASES = 0 And CARTON_PACK_QTY > 1 Then
                                If I = 4 And INNER_PACK_QTY > 0 Then
                                    DISC.DISCOUNT_QTY = INNER_PACK_QTY
                                    If DISC.DISCOUNT_QTY < MSOQ Then
                                        DISC.DISCOUNT_QTY = MSOQ
                                    End If
                                Else
                                    If CARTON_PACK_QTY <> MSOQ And CARTON_PACK_QTY > 1 Then
                                        DISC.DISCOUNT_QTY = 1
                                        If DISC.DISCOUNT_QTY < MSOQ Then
                                            DISC.DISCOUNT_QTY = MSOQ
                                        End If
                                    End If
                                End If
                            Else
                                DISC.DISCOUNT_QTY = CARTON_PACK_QTY * CASES
                            End If
                        End If
                    Else
                        If I = 3 And CARTON_PACK_QTY = 2 Then
                            HAlfCaseIsOne = True
                        End If
                        If CARTON_PACK_QTY = 1 And (I = 3 Or I = 4) Then
                            DISC.DISCOUNT_QTY = 0
                        Else
                            If I = 4 And HAlfCaseIsOne Then
                                DISC.DISCOUNT_QTY = 0
                            Else
                                If (CARTON_PACK_QTY * CASES) < 1 Then
                                    If INNER_PACK_QTY > 1 Then
                                        If I = 4 And (INNER_PACK_QTY = HALFCASE) Then
                                            DISC.DISCOUNT_QTY = 0
                                        Else
                                            DISC.DISCOUNT_QTY = INNER_PACK_QTY
                                        End If
                                    Else
                                        If I = 4 And retval(2).DISCOUNT_QTY = 0 Then
                                            DISC.DISCOUNT_QTY = 0
                                        Else
                                            If I = 4 And HALFCASE = MSOQ Then
                                                DISC.DISCOUNT_QTY = 0
                                            Else
                                                DISC.DISCOUNT_QTY = 1
                                                If DISC.DISCOUNT_QTY < MSOQ Then
                                                    DISC.DISCOUNT_QTY = MSOQ
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If I = 3 Then
                                        DISC.DISCOUNT_QTY = HALFCASE
                                        If DISC.DISCOUNT_QTY < MSOQ Then
                                            DISC.DISCOUNT_QTY = 0
                                        End If
                                    Else
                                        DISC.DISCOUNT_QTY = CARTON_PACK_QTY * CASES
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If DiscPromoFound Then
                        If I = 1 Then
                            If STYLE_PROMO_PRICE <> 0 Then
                                DISC.DISCOUNT_PRICE = STYLE_PROMO_PRICE
                                DISC.DISCOUNT_PCT = String.Format("{0}", DiscPromoDesc)
                            Else
                                DISC.DISCOUNT_PRICE = STYLE_PRICE * (100 - DiscPromoPct) / 100
                                DISC.DISCOUNT_PCT = String.Format("{0}->{1}%", DiscPromoDesc, DiscPromoPct)
                            End If
                            DISC.DISCOUNT_DESC = DiscPromoDesc
                            DISC.DISCOUNT_QTY = 1
                        Else
                            DISC.DISCOUNT_QTY = 0
                        End If
                    Else
                        DISC.DISCOUNT_PRICE = (STYLE_PRICE * (100 - PCT) / 100)
                        DISC.DISCOUNT_PCT = String.Format("{0}->{1}%", rowICTDISC1.Item("DISC_DESC"), PCT)
                        DISC.DISCOUNT_DESC = rowICTDISC1.Item(String.Format("DISC{0}_DESC", CStr(I))) & ""
                    End If
                    retval.Add(DISC)
                Next
            End If
        Else
            For I As Integer = 1 To 4
                Dim DISC As New DISCOUNTS() With {.DISCOUNT_QTY = 0, .DISCOUNT_PCT = "Problem With Style", .DISCOUNT_PRICE = 99999, .DISCOUNT_DESC = "Problem With Style"}
                retval.Add(DISC)
            Next
        End If
        Return retval
    End Function

    Private Function makeAttributesNode(ByVal STYLE_CODE As String, ByVal OnlyPrimary As Boolean) As XmlNode
        Dim ElementName As String = ""
        Dim rowFilter As String = ""
        If OnlyPrimary Then
            ElementName = "attributes"
            rowFilter = String.Format("STYLE_CODE = '{0}' AND ATT_RANK = 1", STYLE_CODE)
        Else
            ElementName = "sub_attributes"
            rowFilter = String.Format("STYLE_CODE = '{0}' AND ATT_RANK <> 1", STYLE_CODE)
        End If
        Dim retVal As XmlNode = XmlDoc.CreateElement(ElementName)

        For Each rowTABLE_NAME As DataRow In _Atributes.Select(rowFilter)
            retVal.InnerText = retVal.InnerText + "||" + rowTABLE_NAME.Item("ATTR_DESC").ToString.ToLower
        Next
        If retVal.InnerText.Length > 2 Then
            retVal.InnerText = retVal.InnerText.Substring(2, retVal.InnerText.Length - 2)
        End If
        Return retVal
    End Function

    Private Function makeColorsNode(ByVal STYLE_CODE As String, ByVal COLOR_CODE As String) As XmlNode
        Dim retVal As XmlNode = XmlDoc.CreateElement("colors")
        Dim filter As String = String.Format("STYLE_CODE = '{0}' AND COLOR_CODE = '{1}'", STYLE_CODE, COLOR_CODE)
        For Each rowColors As DataRow In _Colors.Select(filter)
            retVal.InnerText = retVal.InnerText + "||" + rowColors.Item("COLOR_DESC").ToString
        Next
        If retVal.InnerText.Length > 2 Then
            retVal.InnerText = retVal.InnerText.Substring(2, retVal.InnerText.Length - 2)
        End If
        Return retVal
    End Function

    Private Function makeCategory(ByVal STYLE_CLASS_CODE As String) As String
        Dim retVal As String = ""
        Dim rowICTCLAS1 As DataRow = _Classes.Select(String.Format("STYLE_CLASS_CODE = '{0}'", STYLE_CLASS_CODE)).FirstOrDefault
        If Not IsNothing(rowICTCLAS1) Then
            retVal = rowICTCLAS1.Item("STYLE_CLASS_DESC").ToString
        End If
        Return retVal
    End Function
#End Region

End Class
