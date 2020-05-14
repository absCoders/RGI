Public Class SORSOXP1

    Dim SO_ORDER_NO As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R" ' Y MEANS REPORT WITH UPDATE AN N IS REPORT ONLY A 'U' IS UPDATE ONLY 
        '  FROM NEEDS RGO. TO PREFIX SOTORDR1 & SOTORDR2, SOTORDR5 - POSSIBLY ALL TABLES 
        ' **** REMEMBER RGO. TO UPDATE ROUTINE 
        ' ALSO REPLACE WHERE CLAUSE TO GET ORDR_STATUS = 'L' INSTEAD OF ORDR_NO & REMOVE ALL OTHER STUFF 
        ' & "WHERE  SOTORDR1.ORDR_NO = '0000362015'  " & vbCrLf _
        ' REMOVE STU


        ASCMAIN1.sql = "Select SOTORDR1_L.*, ARTCUST1.CUST_ADDR1, ARTCUST1.CUST_ADDR2, ARTCUST1.CUST_ADDR3, ARTCUST1.CUST_CITY, " & vbCrLf _
            & " ARTCUST1.CUST_STATE, ARTCUST1.CUST_ZIP_CODE, ARTCUST1.CUST_COUNTRY, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_FAX, ARTCUST1.CUST_EMAIL,  " & vbCrLf _
            & " SOTORDR5_L.CUST_NAME SHIP_TO_NAME, SOTORDR5_L.CUST_ADDR1 SHIP_TO_ADDR1, SOTORDR5_L.CUST_ADDR2 SHIP_TO_ADDR2, ARTCUST1.CUST_CITY SHIP_TO_CITY, " & vbCrLf _
            & " SOTORDR5_L.CUST_STATE SHIP_TO_STATE, SOTORDR5_L.CUST_ZIP_CODE  SHIP_TO_ZIP_CODE, SOTORDR5_L.CUST_COUNTRY SHIP_TO_COUNTRY," & vbCrLf _
            & " SOTORDR5_L.CUST_COUNTRY SHIP_TO_COUNTRY, SOTORDR5_L.CUST_PHONE SHIP_TO_PHONE, SOTORDR5_L.CUST_FAX SHIP_TO_FAX " & vbCrLf _
            & " from SOTORDR1_L, ARTCUST1, SOTORDR5_L " & vbCrLf _
            & "WHERE  SOTORDR1_L. ORDR_STATUS = 'O'  " & vbCrLf _
            & "AND SOTORDR1_L.CUST_CODE = ARTCUST1.CUST_CODE " & vbCrLf _
            & "AND SOTORDR1_L.ORDR_NO  = SOTORDR5_L.ORDR_NO (+) " & vbCrLf _
            & "AND SOTORDR5_L.CUST_ADDR_TYPE  (+)  = 'ST'" & vbCrLf
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR1", 1))

        ASCMAIN1.sql = "Select SOTORDR2_L.*, ICTSTYC1.UPC_CODE from SOTORDR2_L , ICTSTYC1 where  SOTORDR2_L.ORDR_NO = :PARM1  AND SOTORDR2_L.STYLE_CODE = ICTSTYC1.STYLE_CODE AND SOTORDR2_L.COLOR_CODE = ICTSTYC1.COLOR_CODE"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR2", 2))
        Create_TDA(dst.Tables.Add(), "SOTORDR2", "**", 0, False, "V", 2)



        'ASCMAIN1.sql = "Select ARTCUST2.*, from ARTCUST2 , SOTORDR1  where SOTORDR1.ORDR_STATUS = 'O' AND ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE AND ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_STORE_NO  "
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2", 3))

        'ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1 WHERE STYLE_CODE IN (SELECT DISTINCT(STYLE_CODE) FROM SOTORDR2, SOTORDR1 where SOTORDR2.PO_STATUS = 'O' AND SOTORDR1.PO_STATUS = 'O' AND SOTORDR2.PO_ORDER_NO = SOTORDR1.PO_ORDER_NO)"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 2))


    End Sub

    Public Overrides Sub Print_Report()
        create_export_file(SO_ORDER_NO)
        ' Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub
    Private Sub create_export_file(ByVal SO_ORDER_NO As String)

        Dim ORDER As String = "C:\VS\VDI\FLAT\" & "ORDER"
        'Dim FLATFOLDER As String = "\\192.168.110.100\FLAT"

        Dim FLATFOLDER As String = "\\192.168.110.100\test\FLAT"
        Dim HCNT As Integer = 0
        Dim DCNT As Integer = 0
        Dim ORDR_NOs As New List(Of String)
        Using swh As System.IO.StreamWriter = _
        New System.IO.StreamWriter(ORDER)

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select
                ORDR_NOs.Add(rowSOTORDR1.Item("ORDR_NO"))
                Dim SOHEAD As String = "".PadLeft(293)
                Mid(SOHEAD, 1, 1) = "H"                                         ' Set header Record
                Mid(SOHEAD, 2, 6) = Mid(rowSOTORDR1.Item("ORDR_NO"), 5, 6)      ' SO Number
                Mid(SOHEAD, 8, 6) = rowSOTORDR1.Item("CUST_CODE") & ""          ' CUST CODE                             ' ALTERNATE-PO-KEY
                Mid(SOHEAD, 14, 30) = rowSOTORDR1.Item("CUST_NAME")             ' CUST_NAME
                Mid(SOHEAD, 44, 25) = rowSOTORDR1.Item("CUST_ADDR1") & ""           ' CUST ADDR1 
                Mid(SOHEAD, 69, 25) = rowSOTORDR1.Item("CUST_ADDR2") & ""       ' CUST ADDR2  
                Mid(SOHEAD, 94, 25) = rowSOTORDR1.Item("CUST_ADDR3") & ""       ' CUST ADDR3
                Mid(SOHEAD, 119, 17) = rowSOTORDR1.Item("CUST_CITY") & ""           ' CUST CITY
                Mid(SOHEAD, 136, 2) = rowSOTORDR1.Item("CUST_STATE") & ""           ' CUST STATE
                Mid(SOHEAD, 138, 9) = rowSOTORDR1.Item("CUST_ZIP_CODE") & ""       ' CUST ZIP CODE
                Mid(SOHEAD, 147, 25) = rowSOTORDR1.Item("CUST_COUNTRY") & ""        ' CUST COUNTRY
                Mid(SOHEAD, 172, 12) = rowSOTORDR1.Item("CUST_PHONE") & ""         ' CUST PHONE
                Mid(SOHEAD, 184, 12) = rowSOTORDR1.Item("CUST_FAX") & ""           ' CUST FAX
                Mid(SOHEAD, 196, 15) = ""                                       ' CUST BUYER 1 
                Mid(SOHEAD, 211, 15) = ""                                       ' CUST BUYER 2
                Mid(SOHEAD, 226, 5) = rowSOTORDR1.Item("ORDR_DEPT") & ""          ' CUST DEPT ORDR_DEPT
                Mid(SOHEAD, 231, 4) = rowSOTORDR1.Item("TERM_CODE") & ""            ' TERM CODE
                Mid(SOHEAD, 235, 1) = Mid(rowSOTORDR1.Item("FRT_TERMS") & "", 1, 1)  ' Freight Terms 
                Mid(SOHEAD, 236, 10) = rowSOTORDR1.Item("ORDR_CUST_PO") & ""       ' cust po 
                Mid(SOHEAD, 246, 8) = Format(CDate(rowSOTORDR1.Item("ORDR_DATE") & ""), "MMddyyyy")           ' Order Date
                'Mid(SOHEAD, 246, 8) = Mid(rowSOTORDR1.Item("ORDR_SHIP_DATE"), 1, 2) & Mid(rowSOTORDR1.Item("ORDR_SHIP_DATE"), 4, 2) & Mid(rowSOTORDR1.Item("ORDR_SHIP_DATE"), 7, 4) & ""           ' Order Date
                Mid(SOHEAD, 254, 39) = rowSOTORDR1.Item("CUST_EMAIL") & ""      ' cust email CUST_EMAIL_TO
                swh.Write(SOHEAD & vbCrLf)                                      ' WRITE 'H' REC

                HCNT = HCNT + 1                                                 ' TOTAL ORDERS IN FILE
                DCNT = 0

                Dim SOADDR As String = "".PadLeft(254)                          ' SHIPPING RECS 
                Mid(SOADDR, 1, 1) = "S"                                         ' set rec type 
                Mid(SOADDR, 2, 6) = Mid(rowSOTORDR1.Item("ORDR_NO"), 5, 6)
                Mid(SOADDR, 8, 3) = rowSOTORDR1.Item("CUST_STORE_NO") & ""   ' CUST STORE NO 
                Mid(SOADDR, 11, 30) = rowSOTORDR1.Item("SHIP_TO_NAME") & ""   ' SHIP TO NAME
                Mid(SOADDR, 41, 25) = rowSOTORDR1.Item("SHIP_TO_ADDR1") & "" ' SHIP TO ADDR1
                Mid(SOADDR, 66, 25) = rowSOTORDR1.Item("SHIP_TO_ADDR2") & ""   ' SHIP TO ADDR2 
                Mid(SOADDR, 91, 24) = ""    ' SHIP TO ADDR3 
                Mid(SOADDR, 114, 18) = rowSOTORDR1.Item("SHIP_TO_CITY") & ""  ' SHIP TO CITY 
                Mid(SOADDR, 133, 2) = rowSOTORDR1.Item("SHIP_TO_STATE") & ""  ' SHIP TO STATE
                Mid(SOADDR, 135, 9) = rowSOTORDR1.Item("SHIP_TO_ZIP_CODE") & ""  ' SHIP TO ZIP CODE  
                Mid(SOADDR, 144, 25) = rowSOTORDR1.Item("SHIP_TO_COUNTRY") & ""  ' SHIP TO COUNTRY 
                Mid(SOADDR, 169, 12) = rowSOTORDR1.Item("SHIP_TO_PHONE") & ""  ' SHIP TO PHONE   
                Mid(SOADDR, 181, 12) = rowSOTORDR1.Item("SHIP_TO_FAX") & ""   ' SHIP TO FAX 
                Mid(SOADDR, 193, 3) = rowSOTORDR1.Item("SHIP_VIA_CODE") & ""   ' SHIP VIA 
                Mid(SOADDR, 196, 2) = rowSOTORDR1.Item("SREP_CODE") & ""    ' SLSMN 1 
                Mid(SOADDR, 198, 2) = ""    ' SLSMN 2
                Mid(SOADDR, 200, 8) = Format(CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE") & ""), "MMddyyyy")   ' SHIP DATE 
                Mid(SOADDR, 208, 8) = ""    ' AS OF DATE 
                Mid(SOADDR, 216, 8) = Format(CDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & ""), "MMddyyyy")   ' CANCEL DATE
                Mid(SOADDR, 224, 1) = ""    ' SHIP AIR OR SEA
                Mid(SOADDR, 225, 2) = rowSOTORDR1.Item("ORDR_FOB") & ""   ' FOB
                Mid(SOADDR, 227, 2) = ""    ' VINYL FOB
                Mid(SOADDR, 229, 1) = ""    ' BAR CODE 
                Mid(SOADDR, 230, 1) = ""    ' PRINT UPC
                Mid(SOADDR, 231, 1) = ""    ' PRINT EACHES
                Mid(SOADDR, 232, 1) = "Y"    ' NET PRICE - AS PER RICH NET PRICES 
                Mid(SOADDR, 233, 5) = ""    ' BOX DISC 
                Mid(SOADDR, 238, 5) = ""    ' CARTON DISC
                Mid(SOADDR, 243, 5) = ""    ' VINYL DISC
                Mid(SOADDR, 248, 2) = ""    ' WRITTEN BY 
                Mid(SOADDR, 250, 5) = ""    ' CUBE

                swh.Write(SOADDR & vbCrLf)

                Dim SOMSG1 As String = "".PadLeft(157)                          ' Order Message 
                Mid(SOMSG1, 1, 1) = "M"                                         ' set rec type  'M'
                Mid(SOMSG1, 2, 6) = Mid(rowSOTORDR1.Item("ORDR_NO"), 5, 6) & "" ' SO Number
                Mid(SOMSG1, 8, 6) = rowSOTORDR1.Item("CUST_CODE") & ""          ' CUST CODE                              
                Mid(SOMSG1, 8, 6) = rowSOTORDR1.Item("ORDR_INV_COMMENT") & ""       ' Message  
                swh.Write(SOMSG1 & vbCrLf)

                Dim SOMSG2 As String = "".PadLeft(157)                          ' Order Message 
                Mid(SOMSG2, 1, 1) = "M"                                         ' set rec type  'N'
                Mid(SOMSG2, 2, 6) = Mid(rowSOTORDR1.Item("ORDR_NO"), 5, 6) & ""     ' SO Number
                Mid(SOMSG2, 8, 6) = Mid(rowSOTORDR1.Item("CUST_CODE"), 5, 6) & ""   ' CUST CODE                              
                Mid(SOMSG2, 8, 6) = rowSOTORDR1.Item("ORDR_INV_COMMENT") & ""      ' Message                          
                swh.Write(SOMSG2 & vbCrLf)
                Fill_Records("SOTORDR2", rowSOTORDR1.Item("ORDR_NO"))
                For Each rowstylecode As DataRow In ASCDATA1.SelectDistinct( _
                    dst.Tables("SOTORDR2"), New String() {"STYLE_CODE"}).Select("")
                    Dim CCNT As Int32 = 0
                    Dim STYLE_CODE As String = rowstylecode.Item("STYLE_CODE")
                    Dim SQLW As String = "STYLE_CODE = '" & STYLE_CODE & "'"
                    For Each rowsotordr2 As DataRow In dst.Tables("SOTORDR2").Select(SQLW, "COLOR_CODE")
                        CCNT = CCNT + 1
                        If CCNT = 1 Then
                            Dim SOITEM As String = "".PadLeft(94)
                            Mid(SOITEM, 1, 1) = "D"                                    ' REC TYPE 'D' 
                            Mid(SOITEM, 2, 6) = Mid(rowsotordr2.Item("ORDR_NO"), 5, 6) & "" ' ORDER NO
                            Mid(SOITEM, 8, 9) = rowsotordr2.Item("STYLE_CODE") & ""        ' STYLE CODE
                            Mid(SOITEM, 17, 3) = "   "      ' cat code  
                            Mid(SOITEM, 20, 25) = rowsotordr2.Item("STYLE_DESC") & ""      ' STYLE DESC
                            Mid(SOITEM, 45, 7) = rowsotordr2.Item("STYLE_PRICE") & "" ' LIST PRICE STYLE_PRICE
                            Mid(SOITEM, 52, 7) = rowsotordr2.Item("ORDR_UNIT_PRICE") & "" ' LIST PRICE STYLE_PRICE
                            Mid(SOITEM, 59, 2) = rowsotordr2.Item("STYLE_UOM") & "" ' style uom

                            Mid(SOITEM, 81, 9) = rowsotordr2.Item("ORDR_UNIT_PRICE") & ""
                            swh.Write(SOITEM & vbCrLf)
                            DCNT = DCNT + 1                                         ' TOTAL DETAILS FOR ORDER
                        End If
                        Dim SOCOLOR As String = "".PadLeft(61)
                        Mid(SOCOLOR, 1, 1) = "C"                                                ' REC TYPE 'D' 
                        Mid(SOCOLOR, 2, 6) = Mid(rowsotordr2.Item("ORDR_NO"), 5, 6) & ""             ' ORDER NO
                        Mid(SOCOLOR, 8, 9) = rowsotordr2.Item("STYLE_CODE") & ""                    ' STYLE CODE
                        Mid(SOCOLOR, 17, 4) = rowsotordr2.Item("COLOR_CODE") & ""                   ' COLOR CODE 
                        Mid(SOCOLOR, 21, 5) = ""                                                 ' DISCOUNT
                        Mid(SOCOLOR, 26, 2) = rowsotordr2.Item("STYLE_UOM") & ""                     ' style uom
                        Dim CF As Integer = IIf(rowsotordr2.Item("STYLE_UOM") = "GR", 144, IIf(rowsotordr2.Item("STYLE_UOM") = "DZ", 12, 1))
                        Mid(SOCOLOR, 28, 5) = Val(rowsotordr2.Item("ORDR_QTY") & "") * CF         ' style uom  
                        Mid(SOCOLOR, 33, 7) = rowsotordr2.Item("STYLE_PRICE") & ""                   ' style uom 
                        Mid(SOCOLOR, 40, 9) = rowsotordr2.Item("ORDR_UNIT_PRICE") & ""       ' SELLING PRICE
                        Mid(SOCOLOR, 49, 12) = rowsotordr2.Item("UPC_CODE") & ""                    ' UPC CODE
                        Mid(SOCOLOR, 61, 1) = " "
                        swh.Write(SOCOLOR & vbCrLf)
                    Next
                Next
                Dim TREC As String = "T" & CStr(DCNT)
                swh.Write(TREC & vbCrLf)
            Next
            Dim XREC As String = "X" & CStr(HCNT)
            swh.Write(XREC & vbCrLf)
            swh.Close()
        End Using

        ' COPY FILE FROM TEMP FOLDER TO LIVE FOLDER

        My.Computer.FileSystem.CopyFile(ORDER, FLATFOLDER & "\" & "ORDERABS", True)

        ' UPDATE STATUS OF ALL ORDERS IN FILE 

        ' REMEMBER TO CHANGE UPDATE TO PREFIX WITH RGO.SOTORDR1
        For Each ORDR_NO As String In ORDR_NOs
            ASCMAIN1.sql = "UPDATE SOTORDR1 SET ORDR_STATUS = 'O' WHERE ORDR_NO = '" & ORDR_NO & "' AND ORDR_STATUS = 'L'"
            ASCDATA1.ExecuteSQL()
        Next

    End Sub



    Overrides Sub Update_Record()
        create_export_file(SO_ORDER_NO)
    End Sub

End Class