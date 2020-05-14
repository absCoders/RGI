Public Class PORCDIF1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R" ' Y MEANS REPORT WITH UPDATE AN N IS REPORT ONLY A 'U' IS UPDATE ONLY 

        ASCMAIN1.sql = "Select POTORDR1.*, APTVEND1.VEND_NAME VENDOR_NAME, APTVEND1.VEND_COUNTRY, ICTWHSE1.WHSE_CITY, ICTWHSE1.WHSE_STATE, ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_NAME from POTORDR1, APTVEND1, ICTWHSE1, ARTCUST1 where PO_STATUS = 'O' AND POTORDR1.VEND_CODE = APTVEND1.VEND_CODE (+) AND POTORDR1.CUST_CODE = ARTCUST1.CUST_CODE (+) AND POTORDR1.WHSE_CODE = ICTWHSE1.WHSE_CODE (+) "
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDR1", 1))

        ' ASCMAIN1.sql = "Select POTORDR2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_UOM from POTORDR2 , POTORDR1 , ICTSTYL1 where POTORDR2.PO_STATUS = 'O' AND POTORDR1.PO_STATUS = 'O' AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+) AND POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+)  and POTORDR2.PO_QTY_OPN > 0 "
        ' dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDR2", 2))

        ASCMAIN1.sql = "Select POTORDR2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.DUTY_RATE_CODE, ICTSTYL1.STYLE_UOM, STYLEXREF.STYLE_CODE_ORIG from POTORDR2 , POTORDR1 , ICTSTYL1, STYLEXREF where  POTORDR1.PO_STATUS = 'O' AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+) AND POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+) AND POTORDR2.STYLE_CODE = STYLEXREF.STYLE_CODE (+)  "
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDR2", 2))



        'ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1 WHERE STYLE_CODE IN (SELECT DISTINCT(STYLE_CODE) FROM POTORDR2, POTORDR1 where POTORDR2.PO_STATUS = 'O' AND POTORDR1.PO_STATUS = 'O' AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO)"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 2))


    End Sub

    Public Overrides Sub Print_Report()
        create_export_file()
        'Generate_Report("PORCDIF1", "Purchase Orders Transmitted to Century")
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub
    Private Sub create_export_file()


        'Dim POHEADER As String = "g:\Century\" & "REGENCYH.txt"
        'Dim PODETAIL As String = "g:\Century\" & "REGENCYD.txt"
        Dim POHEADER As String = "C:\VS\VDI\Century\" & "REGENCYH.txt"
        Dim PODETAIL As String = "C:\VS\VDI\Century\" & "REGENCYD.txt"

        Dim HEADHEADER As String = Chr(34) & "PH-COMP-CODE" & Chr(34) & "~" & Chr(34) & "PH-PO-NO" & Chr(34) & "~" & Chr(34) & "PH-ALTERNATE-PO-KEY" & Chr(34) & "~" & Chr(34) _
                                  & "PH-PO-DESCRIPTION-TX" & Chr(34) & "~" & Chr(34) & "PH-VENDOR-CD" & Chr(34) & "~" & Chr(34) & "TR-VENDOR-NAME" & Chr(34) & "~" & Chr(34) _
                                  & "PH-VENDOR-CTRY-CD" & Chr(34) & "~" & Chr(34) & "PH-SHIP-ORG-CTRY" & Chr(34) & "~" & Chr(34) & "PH-SHIP-ORG-CITY" & Chr(34) & "~" & Chr(34) _
                                  & "PH-EXPORT-CITY-CD" & Chr(34) & "~" & Chr(34) & "PH-FOB-PORT" & Chr(34) & "~" & Chr(34) & "PH-FOB-PORT-DESC" & Chr(34) & "~" & Chr(34) _
                                  & "PH-FOB-PORT-CD-TYPE" & Chr(34) & "~" & Chr(34) & "PH-ORDER-DT" & Chr(34) & "~" & Chr(34) & "PH-FIRST-SHIP-DT" & Chr(34) & "~" & Chr(34) _
                                  & "PH-LAST-SHIP-DT" & Chr(34) & "~" & Chr(34) & "PH-CANCEL-DT" & Chr(34) & "~" & Chr(34) & "PH-IN-WHSE-DT" & Chr(34) & "~" & Chr(34) _
                                  & "PH-FINAL-DEST-TX" & Chr(34) & "~" & Chr(34) & "PH-DELIVER-TO-TX" & Chr(34) & "~" & Chr(34) & "PH-SPEC-INSTRUC-TX" & Chr(34) & "~" & Chr(34) _
                                  & "PH-DIVISION" & Chr(34) & "~" & Chr(34) & "PH-DEPARTMENT" & Chr(34) & "~" & Chr(34) & "PH-CUST-NAME" & Chr(34) & "~" & Chr(34) & "PH-CUST-PO-NO" & Chr(34) & "~" & Chr(34) _
                                  & "PH-BUYER-CD" & Chr(34) & "~" & Chr(34) & "PH-BUYER-NAME" & Chr(34) & "~" & Chr(34) & "PH-AGENT-CD" & Chr(34) & "~" & Chr(34) _
                                  & "PH-AGENT-NAME" & Chr(34) & "~" & Chr(34) & "PH-ORDER-STATUS-CD" & Chr(34) & "~" & Chr(34) & "PH-SHIP-METHOD-CD" & Chr(34) & "~" & Chr(34) & "PH-LC-NO" & Chr(34) & "~" & Chr(34) _
                                  & "PH-LC-DT" & Chr(34) & "~" & Chr(34) & "PH-PRODUCT-CD" & Chr(34) & "~" & Chr(34) & "PH-CURRENCY-CD" & Chr(34) & "~" & Chr(34) & "PH-PAYMENT-CD" & Chr(34) & "~" & Chr(34) _
                                  & "PH-TRANS-TERMS-CD" & Chr(34) & "~" & Chr(34) & "PH-REVISED-SHIP-DT" & Chr(34) & "~" & Chr(34) & "PH-INSPECT" & Chr(34) & "~" & Chr(34) & "PH-MISC1" & Chr(34) & "~" & Chr(34) _
                                  & "PH-MISC2" & Chr(34) & "~" & Chr(34) & "PH-MISC3" & Chr(34) & "~" & Chr(34) & "PH-MISC4" & Chr(34)

        Dim DETHEADER As String = Chr(34) & "PD-COMP-CD" & Chr(34) & "~" & Chr(34) & "PD-PO-NO" & Chr(34) & "~" & Chr(34) & "PD-ITEM-NO" & Chr(34) & "~" & Chr(34) & "PD-COLOR" _
                                & Chr(34) & "~" & Chr(34) & "PD-SIZE" & Chr(34) & "~" & Chr(34) & "PD-ITEM-DESC" & Chr(34) & "~" & Chr(34) & "PD-UNITS-ORDERED-AT" & Chr(34) _
                                & "~" & Chr(34) & "PD-UNITS-UM" & Chr(34) & "~" & Chr(34) & "PD-CARTONS-ORDER-AT" & Chr(34) & "~" & Chr(34) & "PD-CARTONS-UM" & Chr(34) _
                                & "~" & Chr(34) & "PD-UNITS-PER-CARTON" & Chr(34) & "~" & Chr(34) & "PD-COMMODITY-CD" & Chr(34) & "~" & Chr(34) & "PD-ORIGIN-CTRY-CD" & Chr(34) _
                                & "~" & Chr(34) & "PD-ITEM-FIRST-SHIP-D" & Chr(34) & "~" & Chr(34) & "PD-ITEM-LAST-SHIP-DT" & Chr(34) & "~" & Chr(34) & "PD-ITEM-DIVISION" & Chr(34) _
                                & "~" & Chr(34) & "PD-CUST-ITEM-NO" & Chr(34) & "~" & Chr(34) & "PD-MANUFAC-ITEM-NO" & Chr(34) & "~" & Chr(34) & "PD-UNIT-1ST-COST-AT" & Chr(34) _
                                & "~" & Chr(34) & "PD-UNIT-EST-LAND-CST" & Chr(34) & "~" & Chr(34) & "PD-UNIT-SELL-PRICE" & Chr(34) & "~" & Chr(34) & "INSTORE-DATE" & Chr(34) _
                                & "~" & Chr(34) & "HTSUS" & Chr(34) & "~" & Chr(34) & "QUOTA_CATAGORY" & Chr(34) & "~" & Chr(34) & "STYLE" & Chr(34) & "~" & Chr(34) & "PD-MISC1" & Chr(34)
        Dim COMP As String = "134"

        Using swh As System.IO.StreamWriter = _
        New System.IO.StreamWriter(POHEADER)

            swh.Write(HEADHEADER & vbCrLf)

            For Each rowPOTORDR1 As DataRow In dst.Tables("POTORDR1").Select
                Dim POHEAD As String = ""
                POHEAD = COMP                                               ' Company Code
                POHEAD = POHEAD & "~" & rowPOTORDR1.Item("PO_ORDER_NO")   ' PO Number
                POHEAD = POHEAD & "~" & ""                               ' ALTERNATE-PO-KEY
                POHEAD = POHEAD & "~" & ""                              ' PO-DESCRIPTION-TX")
                POHEAD = POHEAD & "~" & Replace(rowPOTORDR1.Item("VEND_CODE") & "", "&", "AND")  ' VEND_CODE
                POHEAD = POHEAD & "~" & Mid(rowPOTORDR1.Item("VENDOR_NAME") & "", 1, 45)  ' Vendor Name 
                POHEAD = POHEAD & "~" & " "                                 ' Vendor Country Code
                POHEAD = POHEAD & "~" & " "                                 ' SHIP-ORG-CTRY
                POHEAD = POHEAD & "~" & " "                                 ' SHIP-ORG-CITY
                POHEAD = POHEAD & "~" & " "                                   ' EXPORT-CITY-CD"
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT-DESC
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT-CD-TYPE
                POHEAD = POHEAD & "~" & FormatDATE(rowPOTORDR1.Item("PO_DATE_ORDERED") & "") ' ORDER-DT
                POHEAD = POHEAD & "~" & FormatDATE(rowPOTORDR1.Item("PO_DATE_SHIP_BY") & "") ' PH-FIRST-SHIP-DT           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-LAST-SHIP-DT        
                POHEAD = POHEAD & "~" & FormatDATE(rowPOTORDR1.Item("PO_DATE_CANCEL") & "") ' PH-CANCEL-DT              
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-IN-WHSE-DT  
                If rowPOTORDR1.Item("WHSE_code") & "" = "MS" Or rowPOTORDR1.Item("WHSE_code") & "" = "NY" Then
                    POHEAD = POHEAD & "~" & Mid(rowPOTORDR1.Item("WHSE_code") & " " & rowPOTORDR1.Item("WHSE_CITY") & " " & rowPOTORDR1.Item("WHSE_STATE") & "", 1, 24)  ' PH-FINAL-DEST-TX            
                Else
                    POHEAD = POHEAD & "~" & Mid(rowPOTORDR1.Item("WHSE_code") & " " & rowPOTORDR1.Item("CUST_CITY") & " " & rowPOTORDR1.Item("CUST_STATE") & "", 1, 24)  ' PH-FINAL-DEST-TX            
                End If
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DELIVER-TO-TX              
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-SPEC-INSTRUC-TX           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DIVISION            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DEPARTMENT            
                POHEAD = POHEAD & "~" & "" & ""                              ' Mid(rowPOTORDR1.Item("CUST_NAME") & "", 1, 45)  '  PH-CUST-NAME           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-CUST-PO-NO            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-BUYER-CD         
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-BUYER-NAME         
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-AGENT-CD          
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-AGENT-NAME        
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-ORDER-STATUS-CD           
                POHEAD = POHEAD & "~" & "" & ""                             ' pH-SHIP-METHOD-CD            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-LC-NO        
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-LC-DT         
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-PRODUCT-CD           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-CURRENCY-CD            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-PAYMENT-CD            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-TRANS-TERMS-CD           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-REVISED-SHIP-DT          
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-INSPECT          
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-MISC1      
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-MISC2       
                POHEAD = POHEAD & "~" & "" & "~"                            ' PH-MISC3
                POHEAD = POHEAD & "~" & "" & "~"                            ' PH-MISC4

                swh.Write(POHEAD & vbCrLf)
            Next
            swh.Close()
        End Using


        Using swd As System.IO.StreamWriter = _
        New System.IO.StreamWriter(PODETAIL)

            swd.Write(DETHEADER & vbCrLf)

            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select

                Dim POHEAD As String = ""
                PODETAIL = COMP                                               ' PD-COMP-CD
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_ORDER_NO") & ""    ' PD-PO-NO
                If rowPOTORDR2.Item("PO_ORDER_NO") & "" <= "148585" Then
                    PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_CODE_ORIG") & ""     ' PD-ITEM-NO
                Else
                    PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_CODE") & ""     ' PD-ITEM-NO
                End If
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("COLOR_CODE") & ""     ' PD-COLOR
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_ORDER_LNO") & ""   ' PD-SIZE - (ACTUALLY PO LINE NUMBER) '
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_DESC") & ""     ' PD-ITEM-DESC
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_QTY_ORD") & ""     ' PD-UNITS-ORDERED-AT
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_UOM") & ""      ' PD-UNITS-UM
                PODETAIL = PODETAIL & "~" & "0"                                     ' PD-CARTONS-ORDER-AT
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-CARTONS-UM
                PODETAIL = PODETAIL & "~" & "0" & ""                                ' PD-UNITS-PER-CARTON
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-COMMODITY-CD
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-ORIGIN-CTRY-CD
                PODETAIL = PODETAIL & "~" & FormatDATE(rowPOTORDR2.Item("PO_DATE_SHIP_BY") & "")     ' PD-ITEM-FIRST-SHIP-D
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-ITEM-LAST-SHIP-DT           
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-ITEM-DIVISION        
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-CUST-ITEM-NO              
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-MANUFAC-ITEM-NO            
                PODETAIL = PODETAIL & "~" & "0"                                     ' PD-UNIT-1ST-COST-AT           
                PODETAIL = PODETAIL & "~" & "0"                                     ' PD-UNIT-EST-LAND-CST              
                PODETAIL = PODETAIL & "~" & "0"                                     ' PD-UNIT-SELL-PRICE         
                PODETAIL = PODETAIL & "~" & ""                                      ' INSTORE-DATE           
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("DUTY_RATE_CODE") & "" ' HTSUS           
                PODETAIL = PODETAIL & "~" & ""                                      ' QUOTA_CATAGORY         
                PODETAIL = PODETAIL & "~" & ""                                      ' STYLE            
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-MISC1        

                swd.Write(PODETAIL & vbCrLf)

            Next
            swd.Close()
        End Using


    End Sub



    Overrides Sub Update_Record()
        create_export_file()
    End Sub




    Public Shared Function FormatDATE(ByVal PDATE As String) As String

        If Trim(PDATE) <> "" Then
            Dim FDATE As String() = Split(PDATE, "/")
            FormatDATE = Format(Convert.ToInt32(Mid(FDATE(2), 1, 4)), "0000") & "-" & Format(Convert.ToInt32(FDATE(0)), "00") & "-" & Format(Convert.ToInt32(FDATE(1)), "00")
        Else
            FormatDATE = Nothing
        End If


    End Function

End Class