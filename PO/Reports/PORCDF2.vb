Public Class PORCDIF1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R" ' Y MEANS REPORT WITH UPDATE AN N IS REPORT ONLY A 'U' IS UPDATE ONLY 

        ASCMAIN1.sql = "Select POTORDR1.*, APTVEND1.VEND_NAME VENDOR_NAME, APTVEND1.VEND_COUNTRY, ICTWHSE1.WHSE_CITY, ICTWHSE1.WHSE_STATE from POTORDR1, APTVEND1, ICTWHSE1 where PO_STATUS = 'O' AND POTORDR1.VEND_CODE = APTVEND1.VEND_CODE (+)  AND POTORDR1.WHSE_CODE = ICTWHSE1.WHSE_CODE (+) "
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDR1", 1))

        ASCMAIN1.sql = "Select POTORDR2.*, ICTSTYL1.STYLE_DESC, ICTSTYL1.DUTY_RATE_CODE from POTORDR2 , POTORDR1 , ICTSTYL1 where POTORDR2.PO_STATUS = 'O' AND POTORDR1.PO_STATUS = 'O' AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+) AND POTORDR2.STYLE_CODE = ICTSTYL1.STYLE_CODE (+) "
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTORDR2", 2))

        'ASCMAIN1.sql = "Select ICTSTYL1.* from ICTSTYL1 WHERE STYLE_CODE IN (SELECT DISTINCT(STYLE_CODE) FROM POTORDR2, POTORDR1 where POTORDR2.PO_STATUS = 'O' AND POTORDR1.PO_STATUS = 'O' AND POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO)"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTSTYL1", 2))


    End Sub

    Public Overrides Sub Print_Report()
        create_export_file()
        ' Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub
    Private Sub create_export_file()

        'Dim FILENAME As String = "C:\VS\HCG\DAL\" & "hillcrest_ice.txt"
        'Dim POHEADER As String = "g:\Century\" & "REGENCYH.txt"
        'Dim PODETAIL As String = "g:\Century\" & "REGENCYD.txt"
        Dim POHEADER As String = "C:\VS\VDI\Century\" & "REGENCYH.txt"
        Dim PODETAIL As String = "C:\VS\VDI\Century\" & "REGENCYD.txt"

        Dim COMP As String = "134"

        Using swh As System.IO.StreamWriter = _
        New System.IO.StreamWriter(POHEADER)

            For Each rowPOTORDR1 As DataRow In dst.Tables("POTORDR1").Select
                Dim POHEAD As String = ""
                POHEAD = "~" & COMP                                       ' Company Code
                POHEAD = POHEAD & "~" & rowPOTORDR1.Item("PO_ORDER_NO")   ' PO Number
                POHEAD = POHEAD & "~" & "~"                               ' ALTERNATE-PO-KEY
                POHEAD = POHEAD & "~" & rowPOTORDR1.Item("WHSE_CITY") & " " & rowPOTORDR1.Item("WHSE_STATE") & ""  ' PO-DESCRIPTION-TX")
                POHEAD = POHEAD & "~" & Replace(rowPOTORDR1.Item("VEND_CODE") & "", "&", "AND")  ' VEND_CODE
                POHEAD = POHEAD & "~" & rowPOTORDR1.Item("VENDOR_NAME") & ""   ' Vendor Name 
                POHEAD = POHEAD & "~" & rowPOTORDR1.Item("VEND_COUNTRY")     ' Vendor Country Code
                POHEAD = POHEAD & "~" & " "                                 ' SHIP-ORG-CTRY
                POHEAD = POHEAD & "~" & " "                                 ' SHIP-ORG-CITY
                POHEAD = POHEAD & "~" & " "                                   ' EXPORT-CITY-CD"
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT-DESC
                POHEAD = POHEAD & "~" & " "                                 ' FOB-PORT-CD-TYPE
                POHEAD = POHEAD & "~" & " "                                 ' ORDER-DT
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-FIRST-SHIP-DT           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-LAST-SHIP-DT        
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-CANCEL-DT              
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-IN-WHSE-DT             
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-FINAL-DEST-TX            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DELIVER-TO-TX              
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-SPEC-INSTRUC-TX           
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DIVISION            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-DEPARTMENT            
                POHEAD = POHEAD & "~" & "" & ""                             ' PH-CUST-NAME           
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
                swh.Write(POHEAD & vbCrLf)
            Next
            swh.Close()
        End Using


        Using swd As System.IO.StreamWriter = _
        New System.IO.StreamWriter(PODETAIL)

            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select

                Dim POHEAD As String = ""
                PODETAIL = "~" & COMP                                               ' PD-COMP-CD
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_ORDER_NO") & ""    ' PD-PO-NO
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_CODE") & ""     ' PD-ITEM-NO
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("COLOR_CODE") & ""     ' PD-COLOR
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_ORDER_LNO") & ""   ' PD-SIZE - (ACTUALLY PO LINE NUMBER) '
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("STYLE_DESC") & ""     ' PD-ITEM-DESC
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_QTY_OPN") & ""     ' PD-UNITS-ORDERED-AT
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_QTY_UOM") & ""     ' PD-UNITS-UM
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-CARTONS-ORDER-AT
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-CARTONS-UM
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-UNITS-PER-CARTON
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-COMMODITY-CD
                PODETAIL = PODETAIL & "~" & "" & ""                                 ' PD-ORIGIN-CTRY-CD
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("PO_DATE_ETA") & ""     ' PD-ITEM-FIRST-SHIP-D
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-ITEM-LAST-SHIP-DT           
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-ITEM-DIVISION        
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-CUST-ITEM-NO              
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-MANUFAC-ITEM-NO            
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-UNIT-1ST-COST-AT           
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-UNIT-EST-LAND-CST              
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-UNIT-SELL-PRICE         
                PODETAIL = PODETAIL & "~" & ""                                      ' INSTORE-DATE           
                PODETAIL = PODETAIL & "~" & ""                                      ' HTSUS           
                PODETAIL = PODETAIL & "~" & ""                                      ' QUOTA_CATAGORY         
                PODETAIL = PODETAIL & "~" & rowPOTORDR2.Item("DUTY_RATE_CODE") & ""          ' STYLE            
                PODETAIL = PODETAIL & "~" & ""                                      ' PD-MISC1        

                swd.Write(PODETAIL & vbCrLf)

            Next
            swd.Close()
        End Using


    End Sub



    Overrides Sub Update_Record()
        create_export_file()
    End Sub

End Class