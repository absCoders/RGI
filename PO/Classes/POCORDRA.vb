Public Class POCORDRA
 
    Dim dst As DataSet
    Public PO_PARM_PO_IMG_DIR As String
    Dim poTables() As String = New String() {"pohdr", "pocolor", "pofactory", "posize", "posizedtl", "potrim", "pofabric"}

    Public Function Produce_XLS(frmASFBASE0 As ASFBASE0, VAN_REF As String, wb As SpreadsheetGear.IWorkbook) As SpreadsheetGear.IWorkbook

        dst = New DataSet

        With dst.Tables.Add("Images")
            .Columns.Add("FILENAME")
            .Columns.Add("IMAGE", GetType(System.Drawing.Bitmap))
            .Columns.Add("IMAGE_TYPE")
            .Columns.Add("IMAGE_DESC")
            .Columns.Add("SOURCE")
            .Columns.Add("POKey", GetType(System.Int32))
            .Columns.Add("POTrimKey", GetType(System.Int32))
        End With

        For Each TABLE_NAME As String In poTables
            ASCMAIN1.sql = "Select * from AT." & Chr(34) & TABLE_NAME & Chr(34) & " where VAN_REF = '" & VAN_REF & "'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, TABLE_NAME)
            dst.Tables.Add(tbl)
            dst.Tables(TABLE_NAME).Columns.Add("ADD")
        Next

        ASCMAIN1.sql = "Select * from POTORDRA where VAN_REF = '" & VAN_REF & "'"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDRA"))
        dst.Tables("POTORDRA").PrimaryKey = New DataColumn() {dst.Tables("POTORDRA").Columns("VAN_REF")}

        Dim rowpohdr As DataRow = dst.Tables("pohdr").Select("")(0) ' .Find(VAN_REF)
        Dim rowPOTORDRA As DataRow = dst.Tables("POTORDRA").Rows.Find(VAN_REF)

        Get_Images(rowpohdr, VAN_REF)

        Dim workbook As SpreadsheetGear.IWorkbook
        Dim worksheet As SpreadsheetGear.IWorksheet

        If wb Is Nothing Then
            workbook = SpreadsheetGear.Factory.GetWorkbook() '(FILENAME)
            worksheet = workbook.Worksheets(0)
        Else
            workbook = wb
            worksheet = workbook.Worksheets.Add
        End If

        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePasteTo As SpreadsheetGear.IRange = Nothing


        Dim NoofSize As Int32 = Val(rowpohdr.Item("NoofSize") & "")
        Dim NoofColor As Int32 = Val(rowpohdr.Item("NoofColor") & "")
        Dim NoofPack As Int32 = Val(rowpohdr.Item("NoofPack") & "")

        Dim CX As Integer = 0

        worksheet.Name = rowpohdr.Item("PONo")

        worksheet.Cells(0, 0).Value = "FOB Cost Sheet"
        With worksheet.Cells(0, 0)
            .Font.Bold = True
            .Font.Size = 14
        End With

        CX = 13
        worksheet.Cells(0, CX + 0).Value = "Printed"
        worksheet.Cells(0, CX + 1).Value = "'" & Format(Now, "MM/dd/yyyy HH:mm")

        worksheet.Cells(1, CX + 0).Value = "Received"
        worksheet.Cells(1, CX + 1).Value = "'" & Format(rowPOTORDRA.Item("DATE_RECEIVED"), "MM/dd/yyyy HH:mm")

        worksheet.Cells(2, CX + 0).Value = "Xmit No"
        worksheet.Cells(2, CX + 1).Value = "'" & rowpohdr.Item("VAN_REF")

        worksheet.Cells(3, CX + 0).Value = "PO Key"
        worksheet.Cells(3, CX + 1).Value = "'" & rowpohdr.Item("POKey")

        worksheet.Cells(0, 0).EntireColumn.ColumnWidth = 25
        worksheet.Cells(0, 1).EntireColumn.ColumnWidth = 25
        worksheet.Cells(0, 2).EntireColumn.ColumnWidth = 40

        Dim RX As Integer = 0
        worksheet.Cells(RX + 2, 0).Value = "Maker"
        worksheet.Cells(RX + 3, 0).Value = "Handled By"
        worksheet.Cells(RX + 4, 0).Value = "Delivery"

        worksheet.Cells(RX + 2, 1).Value = rowpohdr.Item("Factory")
        worksheet.Cells(RX + 3, 1).Value = rowpohdr.Item("FollowBy")
        worksheet.Cells(RX + 4, 1).Value = "'" & Format(rowpohdr.Item("VandaleShipDate"), "MM/dd/yyyy")


        Dim SZ() As String = Nothing
        Dim QZ() As String = Nothing

        Dim SQLVR As String = "VAN_REF = '" & VAN_REF & "'"

        If dst.Tables("posizedtl").Select(SQLVR).Length > 0 Then
            Dim SZL As New List(Of String)
            Dim QZL As New List(Of String)

            For Each row As DataRow In dst.Tables("posizedtl").Select(SQLVR, "POSizeKey")
                If Not SZL.Contains(row.Item("Size") & "") Then
                    SZL.Add(row.Item("Size") & "")
                    QZL.Add(row.Item("Qty") & "")
                End If
            Next

            SZ = SZL.ToArray
            QZ = QZL.ToArray

        Else
            For Each rowposize As DataRow In dst.Tables("posize").Select(SQLVR)
                Dim Color As String = rowposize.Item("Color") & ""
                Dim Size As String = rowposize.Item("Size") & ""
                If Trim(Color).Contains("Size") Then
                    SZ = Size.Split("|")
                ElseIf Color = "" Or Size.Contains("|") Then
                    QZ = Size.Split("|")
                End If
            Next
        End If
         

        RX = 3
        CX = 3
        worksheet.Cells(RX, CX).Value = "Sizes:"
        worksheet.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        worksheet.Cells(RX + 1, CX).Value = "Qtys:"
        worksheet.Cells(RX + 1, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        If SZ Is Nothing OrElse SZ.Length < NoofSize Then
            ReDim Preserve SZ(NoofSize)
            ReDim Preserve QZ(NoofSize)
        End If
        For s As Integer = 1 To NoofSize
            worksheet.Cells(RX, CX + s).Value = SZ(s - 1)
            worksheet.Cells(RX + 1, CX + s).Value = QZ(s - 1)
        Next
        With worksheet.Cells(RX, CX + 1, RX + 1, CX + NoofSize)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With
        With worksheet.Cells(RX + 1, CX + 1, RX + 1, CX + NoofSize)
            .NumberFormat = "#,##0"
        End With

        With worksheet.Cells(RX, CX, RX + 1, CX + NoofSize)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
        End With

        RX = 1

        worksheet.Cells(RX, 0).Value = "Style"
        worksheet.Cells(RX, 1).Value = rowpohdr.Item("StyleNo")
        worksheet.Cells(RX, 2).WrapText = False
        worksheet.Cells(RX, 2).Value = rowpohdr.Item("Style")
        With worksheet.Cells(RX, 1, RX, 2)
            .Font.Bold = True
            .Font.Color = SpreadsheetGear.Colors.Purple
            .WrapText = False
        End With

        If dst.Tables("pofabric").Rows.Count > 0 Then
            Dim rowpofabric = dst.Tables("pofabric").Rows(0)
            worksheet.Cells(RX + 1, 2).Value = rowpofabric.Item("Description")
            With worksheet.Cells(RX + 1, 2)
                .WrapText = False
            End With
        End If

        RX = 5

        RX += 1

        CX = 1
        CX += 1 : worksheet.Cells(RX, CX).Value = "Description"
        CX += 1 : worksheet.Cells(RX, CX).Value = "FOB"
        worksheet.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        worksheet.Cells(RX, CX).EntireColumn.NumberFormat = "#,##0.00"
        CX += 1 : worksheet.Cells(RX, CX).Value = "Total Qty"
        worksheet.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        CX += 1 : worksheet.Cells(RX, CX).Value = "Hangers"
        worksheet.Cells(RX, CX).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        CX += 1 : worksheet.Cells(RX, CX).Value = "Col#"
        worksheet.Cells(RX, CX).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(RX, CX).EntireColumn.NumberFormat = "@"

        Dim CX0 As Integer = CX   ' this is the column of "Col#"
        CX = CX0 + 1

        For s As Integer = 1 To NoofPack
            worksheet.Cells(RX, CX + 0).Value = "Color " & CStr(s)
            worksheet.Cells(RX, CX + 1).Value = "Qty"
            worksheet.Cells(RX, CX + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Right
            worksheet.Cells(RX, CX + 1).EntireColumn.NumberFormat = "#,##0"
            worksheet.Cells(RX, CX + 2).Value = "UM"
            ' worksheet.Cells(RX, CX + 2).EntireColumn.ColumnWidth = 4
            CX += 3
        Next

        Dim CXE As Integer = CX - 1 ' LAST CELL OF COLOR BLOCK
        With worksheet.Cells(RX, 0, RX, CXE)
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With

        Dim RX0 As Integer = RX + 1 ' this is the line under the headings

        Dim RX1 As Integer = RX

        RX = 5
        worksheet.Cells(RX, 0).Value = "Supplier PO Ref"
        worksheet.Cells(RX, 1).Value = rowpohdr.Item("PONo")
        worksheet.Cells(RX, 1).Font.Color = SpreadsheetGear.Colors.Blue
        worksheet.Cells(RX, 1).Font.Bold = True

        worksheet.Cells(RX, 2).Value = rowpohdr.Item("StyleRef")

        RX += 1
        RX += 1


        Dim C0 As Integer = 0
        CX = CX0

        RX = RX0
        worksheet.Cells(RX + 1, 2).Value = "Factory Cost"
        worksheet.Cells(RX + 1, 3).Value = Val(rowpohdr.Item("FactoryCost") & "")
        worksheet.Cells(RX + 1, 3).Font.Bold = True


        Dim XLT As String = ""
        Dim T As Int32 = 0
        Dim ColorCode As String = ""
        Dim ColorCodes As New List(Of String)
        For Each rowpocolor As DataRow In dst.Tables("pocolor").Select(SQLVR, "ColorCode")
            With rowpocolor
                If ColorCode <> .Item("ColorCode") & "" Then
                    If XLT <> "" Then
                        worksheet.Cells(RX, 4).Formula = "=" & Mid(XLT, 2)
                        worksheet.Cells(RX, 5).Formula = "=" & frmASFBASE0.Excel_Cell0(RX, 4) & " * 12 / " & CStr(NoofPack)
                    End If
                    XLT = ""
                    RX += 1
                    C0 = 0
                    ColorCode = .Item("ColorCode")
                    worksheet.Cells(RX, CX + C0 + 0).Value = ColorCode
                    ColorCodes.Add(ColorCode)
                End If
                worksheet.Cells(RX, CX + C0 + 1).Value = .Item("ColorName")
                worksheet.Cells(RX, CX + C0 + 2).Value = .Item("OrderQty")
                worksheet.Cells(RX, CX + C0 + 3).Value = .Item("OrderUnit")
                T += Val(.Item("OrderUnit"))
                XLT &= "+" & frmASFBASE0.Excel_Cell0(RX, CX + C0 + 2)
                C0 += 3
            End With
        Next
        worksheet.Cells(RX, 4).Formula = "=" & Mid(XLT, 2)
        worksheet.Cells(RX, 5).Formula = "=" & frmASFBASE0.Excel_Cell0(RX, 4) & " * 12 / " & CStr(NoofPack)

        RX = 10
        Dim R0 As Integer = 0
        'For Each rowpotrim As DataRow In dst.Tables("potrim").Select(SQLVR)
        '    With rowpotrim
        '        Dim ItemDesc As String = .Item("ItemDesc") & ""
        '        If ItemDesc.EndsWith(vbCr) Then
        '            ItemDesc = Mid(ItemDesc, 1, ItemDesc.Length - 1)
        '        End If
        '        If ItemDesc.EndsWith(vbCrLf) Then
        '            ItemDesc = Mid(ItemDesc, 1, ItemDesc.Length - 2)
        '        End If

        '        worksheet.Cells(RX + R0, 2).Value = ItemDesc
        '        worksheet.Cells(RX + R0, 3).Value = .Item("Price")
        '        R0 += 1
        '    End With
        'Next

        For Each rowfactory As DataRow In dst.Tables("pofactory").Select(SQLVR)
            With rowfactory
                Dim Item As String = .Item("Item") & ""
                If Item.EndsWith(vbCr) Then
                    Item = Mid(Item, 1, Item.Length - 1)
                End If
                If Item.EndsWith(vbCrLf) Then
                    Item = Mid(Item, 1, Item.Length - 2)
                End If

                worksheet.Cells(RX + R0, 2).Value = Item

                Dim MakerCost As String = .Item("MakerCost") & ""
                worksheet.Cells(RX + R0, 3).Value = Val(MakerCost)
                R0 += 1
            End With
        Next

        Dim XC1 As String = ""
        Dim XC2 As String = ""

        XC1 = frmASFBASE0.Excel_Cell0(RX + 0, 3)
        XC2 = frmASFBASE0.Excel_Cell0(RX + R0 - 1, 3)
        worksheet.Cells(RX + R0, 2).Value = "Total Trim Cost"

        worksheet.Cells(RX + R0, 3).Formula = String.Format("=sum({0}..{1})", XC1, XC2)
        worksheet.Cells(RX + R0, 3).Font.Bold = True

        With worksheet.Cells(RX0, 2, RX + R0, 3)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
        End With

        With worksheet.Cells(RX0, 4, RX + R0, 5)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
        End With

        With worksheet.Cells(RX0, 6, RX + R0, 6)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
        End With

        For cc As Integer = 1 To NoofPack
            With worksheet.Cells(RX0, CX0 + (cc - 1) * 3 + 1, RX + R0, CX0 + (cc - 1) * 3 + 3)
                .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continous
                .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continous
            End With
        Next


        RX = R0 + RX
        R0 = 0
        If RX < 20 Then
            RX = 20
        End If
        worksheet.Cells(RX, 0).Value = "Total Amount"
        worksheet.Cells(RX, 1).Value = Format(Val(rowpohdr.Item("TotalVandaleAmount") & ""), "$#,##0.00")
        worksheet.Cells(RX, 1).Font.Bold = True

        XC1 = frmASFBASE0.Excel_Cell0(RX0, 3)
        XC2 = frmASFBASE0.Excel_Cell0(RX - 1, 3)
        worksheet.Cells(RX, 2).Value = "Vandale Cost"
        worksheet.Cells(RX, 3).Value = String.Format("=sum({0}..{1})", XC1, XC2)
        worksheet.Cells(RX, 3).Font.Bold = True

        XC1 = frmASFBASE0.Excel_Cell0(RX0, 4)
        XC2 = frmASFBASE0.Excel_Cell0(RX - 1, 4)
        worksheet.Cells(RX, 4).Value = String.Format("=sum({0}..{1})", XC1, XC2)
        worksheet.Cells(RX, 4).Font.Bold = True

        XC1 = frmASFBASE0.Excel_Cell0(RX0, 5)
        XC2 = frmASFBASE0.Excel_Cell0(RX - 1, 5)
        worksheet.Cells(RX, 5).Value = String.Format("=sum({0}..{1})", XC1, XC2)
        worksheet.Cells(RX, 5).Font.Bold = True

        With worksheet.Cells(RX, 0, RX, 5)
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With

        RX += 1
        worksheet.Cells(RX, 0).Value = "Total Qty"
        worksheet.Cells(RX, 1).Value = Format(Val(rowpohdr.Item("TotalQty") & ""), "#,##0")
        worksheet.Cells(RX, 1).Font.Bold = True


        RX += 0
        R0 += 2

        Dim RR As Int32 = RX + R0 ' STARTING ROW FOR REVISION HISTORY

        worksheet.Cells(RX + R0, 0).Value = "Trim Type"
        worksheet.Cells(RX + R0, 1).Value = "Item No"
        worksheet.Cells(RX + R0, 2).Value = "Description"
        worksheet.Cells(RX + R0, 3).Value = "Price"
        worksheet.Cells(RX + R0, 3).HorizontalAlignment = SpreadsheetGear.HAlign.Right
        worksheet.Cells(RX + R0, 4).Value = "Maker"
        worksheet.Cells(RX + R0, 5).Value = "Supplier"

        With worksheet.Cells(RX + R0, 0, RX + R0, 5)
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With

        RX += 1

        For Each rowpotrim As DataRow In dst.Tables("potrim").Select(SQLVR)
            With rowpotrim
                Dim ItemNo As String = .Item("ItemNo") & ""
                Dim ItemSDesc As String = .Item("ItemSDesc") & ""
                Dim ItemDesc As String = .Item("ItemDesc") & ""
                If ItemDesc.EndsWith(vbCr) Then
                    ItemDesc = Mid(ItemDesc, 1, ItemDesc.Length - 1)
                End If
                If ItemDesc.EndsWith(vbCrLf) Then
                    ItemDesc = Mid(ItemDesc, 1, ItemDesc.Length - 2)
                End If

                worksheet.Cells(RX + R0, 0).Value = ItemSDesc
                worksheet.Cells(RX + R0, 1).Value = ItemNo
                worksheet.Cells(RX + R0, 2).Value = ItemDesc
                worksheet.Cells(RX + R0, 3).Value = .Item("Price")
                worksheet.Cells(RX + R0, 4).Value = .Item("Maker")
                worksheet.Cells(RX + R0, 5).Value = .Item("Supplier")
                R0 += 1
            End With
        Next


        Dim RRC As Integer = 13

        worksheet.Cells(RR, RRC + 0).Value = "Rev No"
        worksheet.Cells(RR, RRC + 0).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(RR, RRC + 1).Value = "Date"
        worksheet.Cells(RR, RRC + 1).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        worksheet.Cells(RR, RRC + 3).Value = "Description"

        With worksheet.Cells(RR, RRC + 0, RR, RRC + 3 + 1)
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With

        RR += 1

        ASCMAIN1.sql = "Select VAN_REF, `SentSeq`, `SendTime`, `SendRemarks` from AT.`pohdr`" _
            & " where `POKey` = " & rowPOTORDRA.Item("POKey")
        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", ChrW(34))

        For Each rowpohdrRev As DataRow In ASCDATA1.GetDataTable.Select("", "VAN_REF")
            With rowpohdrRev
                Dim SentSeq As Int32 = Val(.Item("SentSeq") & "")
                Dim SendTime As DateTime = .Item("SendTime") & ""
                Dim SendRemarks As String = .Item("SendRemarks") & ""

                worksheet.Cells(RR, RRC + 0).Value = SentSeq
                worksheet.Cells(RR, RRC + 0).HorizontalAlignment = SpreadsheetGear.HAlign.Center
                worksheet.Cells(RR, RRC + 1).Value = "'" & Format(SendTime, "MM/dd/yyyy HH:mm")
                worksheet.Cells(RR, RRC + 2).Value = SendRemarks

                RR += 1
            End With
        Next

        ' Main Picture

        Dim rowImages() As DataRow = dst.Tables("Images").Select("IMAGE_TYPE = 'Picture1'")
        If rowImages.Length = 1 Then
            Dim IMAGE_NAME As String = rowImages(0).Item("FILENAME") ' "0\VCO51345.png"
            Dim imageFileStyle As String = PO_PARM_PO_IMG_DIR & "\" & IMAGE_NAME ' & ".jpg"

            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution
                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution
                Finally
                    imageStyle.Dispose()
                End Try

                ' Calculate the left and top placement of the picture by converting 
                ' row and column coordinates to points.  Use fractional values to 
                ' get coordinates anywhere in between row and column boundaries.
                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0)
                Dim topStyle As Double = windowInfoStyle.RowToPoints(RX0)
                Dim rightStyle As Double = windowInfoStyle.ColumnToPoints(2)
                Dim bottomStyle As Double = windowInfoStyle.RowToPoints(RX0 + 13)

                Dim Fw As Double = (rightStyle - leftStyle) / widthStyle
                Dim Fh As Double = (bottomStyle - topStyle) / heightStyle
                Dim F As Double = Fw
                If Fh < Fw Then F = Fh

                widthStyle = widthStyle * F
                heightStyle = heightStyle * F


                '   ImageRows = windowInfoStyle.PointsToRow(heightStyle)

                ' Add the picture from file.
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
            End If
        End If


        RX += R0 + 2

        Dim R_Other As Int32 = RX

        ' Other Pictures in Header

        For Each row As DataRow In dst.Tables("Images").Select("IMAGE_TYPE <> 'Picture1' and SOURCE = 'pohdr'")

            Dim IMAGE_NAME As String = row.Item("FILENAME")
            Dim IMAGE_TYPE As String = row.Item("IMAGE_TYPE")

            worksheet.Cells(RX, 2).Value = row.Item("IMAGE_DESC")
            Dim imageFileStyle As String = PO_PARM_PO_IMG_DIR & "\" & IMAGE_NAME ' & ".jpg"

            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution
                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution
                Finally
                    imageStyle.Dispose()
                End Try

                ' Calculate the left and top placement of the picture by converting 
                ' row and column coordinates to points.  Use fractional values to 
                ' get coordinates anywhere in between row and column boundaries.
                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo
                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(0)
                Dim topStyle As Double = windowInfoStyle.RowToPoints(RX)
                Dim rightStyle As Double = windowInfoStyle.ColumnToPoints(2)
                '  Dim bottomStyle As Double = windowInfoStyle.RowToPoints(RX0 + 13)

                Dim Fw As Double = (rightStyle - leftStyle) / widthStyle
                Dim F As Double = Fw
                widthStyle = widthStyle * F
                heightStyle = heightStyle * F
                Dim RXX As Integer = 0
                Do
                    RXX += 1
                Loop Until windowInfoStyle.RowToPoints(RX + RXX) - topStyle > heightStyle
                RX += RXX
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
            End If
        Next

        ' Block Pictures across the bottom (trim pictures)

        Dim Cpi As Int32 = 0

        For Each row As DataRow In dst.Tables("Images").Select("IMAGE_TYPE <> 'Picture1' and SOURCE <> 'pohdr'")

            Cpi += 1
            RX = R_Other + 8 * Math.Truncate((Cpi - 1) / 3)
            CX = 4 + 6 * ((Cpi - 1) Mod 3)

            Dim IMAGE_NAME As String = row.Item("FILENAME")
            Dim IMAGE_TYPE As String = row.Item("IMAGE_TYPE")

            worksheet.Cells(RX, CX).Value = row.Item("IMAGE_DESC")

            Dim imageFileStyle As String = PO_PARM_PO_IMG_DIR & "\" & IMAGE_NAME ' & ".jpg"

            If My.Computer.FileSystem.FileExists(imageFileStyle) Then
                Dim widthStyle As Double
                Dim heightStyle As Double

                Dim imageStyle As System.Drawing.Image = System.Drawing.Image.FromFile(imageFileStyle)
                Try
                    widthStyle = imageStyle.Width * 72.0 / imageStyle.HorizontalResolution
                    heightStyle = imageStyle.Height * 72.0 / imageStyle.VerticalResolution
                Finally
                    imageStyle.Dispose()
                End Try

                ' Calculate the left and top placement of the picture by converting 
                ' row and column coordinates to points.  Use fractional values to 
                ' get coordinates anywhere in between row and column boundaries.
                Dim windowInfoStyle As SpreadsheetGear.IWorksheetWindowInfo = worksheet.WindowInfo

                Dim leftStyle As Double = windowInfoStyle.ColumnToPoints(CX)
                Dim topStyle As Double = windowInfoStyle.RowToPoints(RX) + 20
                Dim rightStyle As Double = windowInfoStyle.ColumnToPoints(CX + 5)

                Dim Fw As Double = (rightStyle - leftStyle) / widthStyle
                Dim F As Double = Fw

                widthStyle = widthStyle * F
                heightStyle = heightStyle * F

                Dim RXX As Integer = 0
                Do
                    RXX += 1
                Loop Until windowInfoStyle.RowToPoints(RX + RXX) - topStyle > heightStyle
                RX += RXX
                worksheet.Shapes.AddPicture(imageFileStyle, leftStyle, topStyle, widthStyle, heightStyle)
            End If
        Next


        With worksheet.PageSetup
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .FitToPages = True
            .Orientation = SpreadsheetGear.PageOrientation.Landscape
        End With



        Return workbook

    End Function


    Sub Get_Images(rowpohdr As DataRow, VAN_REF As String)

        dst.Tables("Images").Rows.Clear()

        Get_Image(rowpohdr, "PolyBag", "Poly Bag", "PolyBagImg")
        Get_Image(rowpohdr, "ShippingMark", "Shipping Mark", "ShippingMarkImg")
        Get_Image(rowpohdr, "Packing", "Packing", "PackingImg")
        Get_Image(rowpohdr, "Sample", "Sample", "SampleImg")
        Get_Image(rowpohdr, "", "Picture1", "PictureName1")
        Get_Image(rowpohdr, "", "Picture2", "PictureName2")

        For Each rowpotrim As DataRow In dst.Tables("potrim").Select("VAN_REF = '" & VAN_REF & "'")
            Get_Image(rowpotrim, "ItemDesc", "Trim Item", "PictureName")
        Next


    End Sub

    Sub Get_Image(row As DataRow, COLUMN_NAME_DESC As String, IMAGE_TYPE As String, COLUMN_NAME_FILENAME As String)

        Dim FILENAME As String = row.Item(COLUMN_NAME_FILENAME) & ""
        If FILENAME = "" Then Return
        'C:\Ashley-System\UploadFolder\190\IZ-BRA(SML) POLYBAG.jpg
        Dim f As Integer = FILENAME.IndexOf("\UploadFolder\")
        If f = 0 Then Return
        FILENAME = FILENAME.Substring(f + "\UploadFolder\".Length)


        Dim IMAGE_DESC As String = ""
        If COLUMN_NAME_DESC <> "" Then
            IMAGE_DESC = row.Item(COLUMN_NAME_DESC) & ""
        End If

        'Dim FOLDER_FROM As String = "\\192.168.160.100\UploadFolder\"
        'Dim FOLDER_TO As String = PO_PARM_PO_IMG_DIR & "\"
        'My.Computer.FileSystem.CopyFile(FOLDER_FROM & FILENAME, FOLDER_TO & FILENAME, True)

        Dim IMAGE As System.Drawing.Bitmap = ASCMAIN1.Get_Image(PO_PARM_PO_IMG_DIR, FILENAME, True, , , )

        Dim rowImages As DataRow = dst.Tables("Images").NewRow
        rowImages.Item("FILENAME") = FILENAME
        rowImages.Item("IMAGE_TYPE") = IMAGE_TYPE
        rowImages.Item("IMAGE_DESC") = IMAGE_DESC
        rowImages.Item("SOURCE") = row.Table.TableName
        rowImages.Item("IMAGE") = IMAGE
        If IMAGE_TYPE = "Trim Item" Then
            rowImages.Item("POKey") = row.Item("POKey")
            rowImages.Item("POTrimKey") = row.Item("POTrimKey")
        End If

        dst.Tables("Images").Rows.Add(rowImages)
    End Sub

End Class
