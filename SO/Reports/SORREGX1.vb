Public Class SORREGX1
    Dim SOTPICK1 As String


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R" ' Y MEANS REPORT WITH UPDATE AN N IS REPORT ONLY A 'U' IS UPDATE ONLY 
        Dim reg_date As Date = Now

        SOTPICK1 = ASCMAIN1.Temp_Table("Select SOTPICK1.* , SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.CUST_STORE_NAME from SOTPICK1, SOTORDR1 where SOTPICK1.PICK_STATUS = 'P' AND SOTPICK1.PICK_PRINTED IS NOT NULL AND SOTPICK1.PICK_PICKER IS NULL AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO")

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTPICK1, "SOTPICK1", 1))

        'ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.*, ARTCUST1.CUST_NAME CUST_NAME, ICTWHSE1.WHSE_CITY, ICTWHSE1.WHSE_STATE from SOTPICK1, SOTORDR1, ARTCUST1, ICTWHSE1 where ORDR_STATUS = 'O' AND SOTORDR1.CUST_CODE = ARTCUST1.CUST_CODE (+)  AND SOTORDR1.WHSE_CODE = ICTWHSE1.WHSE_CODE (+) "
        'ASCMAIN1.sql = "Select SOTPICK1.* , SOTORDR1.* from SOTPICK1, SOTORDR1 where SOTPICK1.PICK_STATUS = 'P' AND SOTPICK1.PICK_PRINTED IS NOT NULL AND SOTPICK1.PICK_PICKER IS NULL AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", 1))

        ' ASCMAIN1.sql = "Select SOTPICK1.* , SOTORDR1.* from SOTPICK1, SOTORDR1 where WHERE PICK_WHSE_BATCH = '" & XNO & "'  AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO'"
        ' dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", 1))


        ' ASCMAIN1.sql = "Select SOTPICK1.* , SOTORDR1.* from SOTPICK1, SOTORDR1 where SOTPICK1.PICK_STATUS = 'P' AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO AND SOTPICK1.PICK_PRINTED > '17-jun-2013'"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", 1))

        'grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")

        create_export_file()

    End Sub

    Public Overrides Sub Print_Report()

        Generate_Report("SORREGX1", "Pick Tickets Printed")
       
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

        Dim REGDATE As String = FormatDATE(DATETIME_STAMP.Date)
        Dim regHOUR As Int32 = DATETIME_STAMP.Hour
        Dim REGMINUTE As Int32 = DATETIME_STAMP.Minute
        Dim TSTAMP As String = regHOUR & REGMINUTE
        Dim REGFILE As String = "S:\WAREHOUSE\FILEAMIGOUPLOAD\" & "REG" & REGDATE & TSTAMP & ".txt"
        'Dim REGFILE As String = "C:\VS\VDI\WAREHOUSE\" & "REG" & REGDATE & ".txt"
        'Dim REGFILE As String = "S:\WAREHOUSE\TEST\" & "REG" & REGDATE & ".txt"

        Dim SHIPLINK As String = "\\njsrvr\inst\"

        Dim EMPLOYEE As String = "UA"
        Dim TASK As String = "ORDER"
        Dim D_DATE As Date = Now

        Using swh As System.IO.StreamWriter = _
        New System.IO.StreamWriter(REGFILE)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select
                Dim REGISTER As String = ""
                Dim SPACER As String = Chr(9)

                Dim DANA As Integer = REGISTER.Length()

                REGISTER = EMPLOYEE & SPACER                                    ' EMPLOYEE
                REGISTER = REGISTER & TASK & SPACER                                         ' TASK
                REGISTER = REGISTER & rowSOTPICK1.Item("CUST_CODE") & SPACER                ' CUST_CODE
                REGISTER = REGISTER & rowSOTPICK1.Item("CUST_NAME") & SPACER                ' CUST_NAME  
                REGISTER = REGISTER & rowSOTPICK1.Item("CUST_STORE_NAME") & SPACER         ' 
                REGISTER = REGISTER & rowSOTPICK1.Item("PICK_NO") & SPACER                 ' REG-REGISTER-NUM
                REGISTER = REGISTER & SPACER                                         'LINES   
                REGISTER = REGISTER & SHIPLINK & rowSOTPICK1.Item("CUST_CODE") & SPACER    ' SHIP-INST  
                REGISTER = REGISTER & "N" & SPACER                                          ' C-C 
                REGISTER = REGISTER & "N" & SPACER                                         ' BLACK-LIST     
                REGISTER = REGISTER & SPACER                                         ' ASSIGMENT  
                REGISTER = REGISTER & SPACER                                         ' START-TIME 
                REGISTER = REGISTER & SPACER                                        ' END-TIME  
                REGISTER = REGISTER & SPACER                                        ' CARRIER   
                REGISTER = REGISTER & SPACER                                          ' TRACKING-NO 
                REGISTER = REGISTER & "N" & SPACER                                         ' CANCEL  
                REGISTER = REGISTER & SPACER                                          ' CXL-RETURN-DATE 
                REGISTER = REGISTER & "N" & SPACER                                         ' NO-STOCK 
                REGISTER = REGISTER & SPACER                                          ' CARTONS  
                REGISTER = REGISTER & SPACER                                         ' NO-PREPAID 
                REGISTER = REGISTER & SPACER                                           ' NO-COLLECT   
                REGISTER = REGISTER & SPACER                                          ' AMT-TO-SANCHEZ  
                REGISTER = REGISTER & SPACER                                          ' SHIP-DATE 
                REGISTER = REGISTER & SPACER                                          ' FREIGHT    
                REGISTER = REGISTER & SPACER                                          ' GROUP 
                REGISTER = REGISTER & SPACER                                          ' END-WEEK  
                REGISTER = REGISTER & "N" & SPACER                                          ' 20-PCT    
                REGISTER = REGISTER & SPACER                                           ' LBS   
                REGISTER = REGISTER & "N" & SPACER                                         ' 1ST-ATEMPT    
                REGISTER = REGISTER & "N" & SPACER                                         ' 2ND-ATEMPT
                REGISTER = REGISTER & SPACER                                          ' COMMENTS     
                REGISTER = REGISTER & D_DATE & SPACER                                    ' D-D  
                REGISTER = REGISTER & rowSOTPICK1.Item("ORDR_SHIP_DATE") & "" & SPACER    ' ORD-SHIP-DATE    
                REGISTER = REGISTER & rowSOTPICK1.Item("ORDR_CANCEL_DATE") & "" & SPACER   ' CANCEL-DATE 
                REGISTER = REGISTER & SPACER                                          ' REG-SHIP-CODE    
                REGISTER = REGISTER & SPACER                                         ' REG-ORDER-NUM  
                REGISTER = REGISTER & Mid(Replace(Replace(rowSOTPICK1.Item("PICK_NO"), Chr(13), ""), Chr(10), ""), 1, 60) & SPACER ' REG-PACK-MSG     
                REGISTER = REGISTER & Mid(Replace(Replace(rowSOTPICK1.Item("PICK_NO"), Chr(13), ""), Chr(10), ""), 61, 60) & SPACER ' REG-VESSEL  
                REGISTER = REGISTER & vbCr
                'REGISTER = REGISTER & "~" & " "                                 ' SHIP-ORG-CTRY
                'REGISTER = REGISTER & "~" & " "                                 ' SHIP-ORG-CITY


                swh.Write(REGISTER & vbCrLf)

            Next
            swh.Close()
        End Using


        'ASCMAIN1.sql = "Update SOTPICK1 " _
        '& " Set PICK_WHSE_BATCH = '" & XNO & "' WHERE SOTPICK1.PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'"
        'ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub create_export_file_itemlocs()

        Dim filename As String = "C:\dmp\ITEMSLOCS.txt"
        Dim data As String = ""
        Dim STYLE_CODE As String = ""
        Dim COLOR_CODE As String = ""
        Dim BIN As String = ""

        Get_PARM("ICTPARMR")


        Using sr As New System.IO.StreamReader(filename)
            data = sr.ReadToEnd
            Dim datarec() As String = Split(data, vbCrLf)

            For i As Integer = 0 To UBound(datarec)
                If datarec(i) <> "" Then
                    Dim datastr() As String = Split(datarec(i), vbTab)
                    STYLE_CODE = Replace(Replace(Replace(datastr(0), Chr(34), ""), "'", ""), " ", "")
                    COLOR_CODE = Replace(Replace(datastr(1), Chr(34), ""), " ", "")
                    BIN = Replace(Replace(datastr(2), Chr(34), ""), " ", "")
                    ASCMAIN1.sql = "UPDATE ICTSTYC1 SET STYLE_LOCATION = '" & BIN & "' where STYLE_CODE = '" & STYLE_CODE & "' and COLOR_CODE = '" & COLOR_CODE & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            Next
        End Using

    End Sub

    Overrides Sub Update_Record()

        Dim SQL As String = "Update SOTPICK1 " _
         & " Set PICK_PICKER = '" & XNO & "' WHERE SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & SOTPICK1 & ")"

        ASCDATA1.ExecuteSQL(SQL)
    End Sub



    Public Shared Function FormatDATE(ByVal PDATE As String)

        Dim datestr() As String = Split(PDATE, "/")

        If Trim(PDATE) <> "" Then
            FormatDATE = Replace(PDATE, "/", "_") ' CDate((Mid(PDATE, 6, 2) & "/" & Mid(PDATE, 9, 2) & "/" & Mid(PDATE, 1, 4)))
            FormatDATE = datestr(2) & "_" & datestr(1) & "_" & datestr(0)
        Else
            FormatDATE = Nothing
        End If


    End Function

End Class