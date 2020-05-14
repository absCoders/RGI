Public Class TARPEND0
    Dim tblTATPENDX As DataTable

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)

        tblTATPENDX = New DataTable
        With tblTATPENDX
            .Columns.Add("MODULE_ID")
            .Columns.Add("TABLE_NAME")
            .Columns.Add("CONDITION")
            .Columns.Add("SQL")
            .Columns.Add("RECORDS", GetType(System.Int16))
        End With

        grdTATPENDX.DataSource = tblTATPENDX
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "U"

    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" = "1" Then
                EMsg = EMsg & vbCr & "Period-End Already Initialized"
            Else
                Dim YYYY As String
                For i As Integer = 0 To 1
                    YYYY = Format$(Val(Mid$(ASCMAIN1.CYP, 1, 4)) + i, "0000")
                    ASCMAIN1.sql = "Select Count (*) from GLTPARM2 where OPS_YYYYPP like '" & YYYY & "%'"
                    If Val(ASCDATA1.GetDataValue & "") <> 12 Then
                        EMsg = EMsg & vbCr & "Please Check Operations Calendar for " & YYYY
                    End If
                Next
            End If

            If EMsg = "" Then

                tblTATPENDX.Rows.Clear()
                Check_Conditions()

                If tblTATPENDX.Rows.Count <> 0 Then
                    EMsg &= vbCr & "Cannot Proceed because of Conditions Listed"
                End If
            End If

            If EMsg <> "" Then
                EMsg = "Cannot Proceed because a Clean Cut-off has not been established as follows:" & vbCr & EMsg
            End If
        End If

    End Sub

    Sub Check_Conditions()

        Check_for_Records("APTPYMT1", "AP Payment Selection Batch")
        'Check_for_Records("ARTPYMT1", "AR Payment Application Journal", "STATUS = '1'")
        Check_for_Records("ARTPYMT1", "AR Payment Application Journal", "NVL(STATUS,'0') <> '2'")
        Check_for_Records("ARTPYMT2", "AR Payment Receipts Journal", "NVL(PYMT_STATUS,'0') <> '2'")
        Check_for_Records("APTINVH1", "AP Vendor Invoice Register", "REGISTER_IND = '0'")
        Check_for_Records("APTCHCK1", "AP Check Register", "REGISTER_IND  = '0' or (OPS_YYYYPP_F is Not Null and REGISTER_IND_F = '0')")

        If ASCMAIN1.CLIENT = "VAN" Then
            Check_for_Records("SOTINVH1", "Sales Journal", "ORDR_YYYYPP_UPDATED IS NULL")
        Else
            Check_for_Records("SOTINVH1", "Sales Journal", "ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "' and NVL(REGISTER_IND,'0') = '0'")
        End If

        ' Check_for_Records("ICTIREC1", "PO Receipts Journal", "REGISTER_IND = '0'")
        ' Check_for_Records("ICTIADJ1", "Inventory Adjustments Journal", "REGISTER_IND = '0'")
        ' Check_for_Records("ICTIXFR1", "Warehouse Transfer Journal", "REGISTER_IND = '0'")

        ' CHECK FOR ICTWHSE1 WITH PI IN PROCESS

        If ASCMAIN1.CLIENT = "NYA" Then
            Check_Open_AP_NYAG_Canada()
        End If

    End Sub

    Sub Check_for_Records( _
    ByVal TABLE_NAME As String, _
    ByVal TABLE_DESC As String, _
    Optional ByVal where_clause As String = "", _
    Optional ByVal custom_sql As String = "")

        If custom_sql <> "" Then
            ASCMAIN1.sql = custom_sql
        Else
            ASCMAIN1.sql = "Select count (*) from " & TABLE_NAME
            If where_clause <> "" Then
                ASCMAIN1.sql &= " where " & where_clause
            End If
        End If

        Dim sql As String = ASCMAIN1.sql
        Dim r As Long = Val(ASCDATA1.GetDataValue() & "")
        If r <> 0 Then
            Dim row As DataRow = tblTATPENDX.NewRow
            row.Item("MODULE_ID") = MODULE_ID
            row.Item("TABLE_NAME") = TABLE_NAME
            row.Item("CONDITION") = TABLE_DESC
            row.Item("SQL") = where_clause
            row.Item("RECORDS") = r
            tblTATPENDX.Rows.Add(row)
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(r) & " Records"
        End If
    End Sub

    Sub Check_Open_AP_NYAG_Canada()

        ' THESE SQLS WERE TAKEN RIGHT OUT OF GLFTBALC

        Dim TABLE_NAME As String
        Dim TABLE_DESC As String

        Dim RECS As Integer = 0

        ASCMAIN1.sql = "Select COUNT (Distinct APTINVH1.VOUCHER_NO) RECS" & vbCrLf _
            & " from APTINVH5,APTINVH1,POTSHIP2,ICTIREC1,ICTIREC2,ICTSTYL1" & vbCrLf _
            & " where APTINVH5.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
            & "   and APTINVH1.OPS_YYYYPP >= '201801'" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_NO = APTINVH5.PO_SHIPMENT_NO" & vbCrLf _
            & "   and POTSHIP2.PO_SHIPMENT_LNO = APTINVH5.PO_SHIPMENT_LNO" & vbCrLf _
            & "   and ICTIREC1.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
            & "   and ICTIREC2.RECEIPT_LNO = APTINVH5.RECEIPT_LNO" _
            & "   and ICTSTYL1.STYLE_CODE = ICTIREC2.STYLE_CODE" & vbCrLf _
            & "   and (ICTIREC1.WHSE_CODE IN (Select WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001')" & vbCrLf _
            & "     or ICTSTYL1.SALES_DIVISION_CODE IN (Select SALES_DIVISION_CODE FROM SOTSDIV1 WHERE SEG4_CODE = '001'))" & vbCrLf _
            & "   and (APTINVH1.INV_STATUS = 'O' or APTINVH1.INV_STATUS = 'H')"
        RECS = Val(ASCDATA1.GetDataValue())

        If RECS <> 0 Then

            If MsgBox(CStr(RECS) & " NYAG Canada Supplier Invoices were not paid.  Skip this block?", MsgBoxStyle.YesNo, "Robin - Answer Y to skip this block") = MsgBoxResult.Yes Then
                RECS = 0
            End If

        End If

        If RECS <> 0 Then
            TABLE_NAME = "NYAGCANP"
            TABLE_DESC = "NYAG Canada Purchases"
            Dim row As DataRow = tblTATPENDX.NewRow
            row.Item("MODULE_ID") = MODULE_ID
            row.Item("TABLE_NAME") = TABLE_NAME
            row.Item("CONDITION") = TABLE_DESC
            row.Item("SQL") = ""
            row.Item("RECORDS") = RECS
            tblTATPENDX.Rows.Add(row)
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(RECS) & " Records"
        End If

        ASCMAIN1.sql = "Select COUNT (Distinct APTINVH1.VOUCHER_NO) RECS" & vbCrLf _
            & " from APTINVH7,APTINVH1,POTSHIP2,POTSHIP1,POTLCST1" & vbCrLf _
            & "WHERE APTINVH7.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
            & "AND APTINVH1.OPS_YYYYPP >= '201801'" & vbCrLf _
            & "AND POTSHIP1.PO_SHIPMENT_NO = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
            & "AND POTSHIP2.PO_SHIPMENT_NO (+) = POTLCST1.PO_SHIPMENT_NO" & vbCrLf _
            & "AND POTSHIP2.PO_SHIPMENT_LNO (+) = POTLCST1.PO_SHIPMENT_LNO" & vbCrLf _
            & "AND POTLCST1.CTL_NO = APTINVH7.CTL_NO " & vbCrLf _
            & "   AND (POTSHIP1.WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE SEG4_CODE = '001') )" _
            & "   and (APTINVH1.INV_STATUS = 'O' or APTINVH1.INV_STATUS = 'H')"
        RECS = Val(ASCDATA1.GetDataValue())

        If RECS <> 0 Then

            If MsgBox(CStr(RECS) & " NYAG Canada Landing Cost Invoices were not paid.  Skip this block?", MsgBoxStyle.YesNo, "Robin - Answer Y to skip this block") = MsgBoxResult.Yes Then
                RECS = 0
            End If

        End If

        If RECS <> 0 Then
            TABLE_NAME = "NYAGCANO"
            TABLE_DESC = "NYAG Canada Landing Costs"
            Dim row As DataRow = tblTATPENDX.NewRow
            row.Item("MODULE_ID") = MODULE_ID
            row.Item("TABLE_NAME") = TABLE_NAME
            row.Item("CONDITION") = TABLE_DESC
            row.Item("SQL") = ""
            row.Item("RECORDS") = RECS
            tblTATPENDX.Rows.Add(row)
            EMsg &= vbCr & TABLE_DESC & " (" & TABLE_NAME & ") " & CStr(RECS) & " Records"
        End If


    End Sub

    Overrides Sub Update_Record()

        'ASCMAIN1.sql = "Select INV_TYPE, INV_NO from SOTINVH1 where ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "' and NVL(REGISTER_IND,'0') = '0'"

        ASCMAIN1.sql = "Update ASTPCTL1 set PRD_CLOSE_IND = '1'"
        ASCDATA1.ExecuteSQL()

        Create_Monthly_Backup_Script()
    End Sub

    Sub Create_Monthly_Backup_Script()
        Dim FOLDER As String = "S:\" & ASCMAIN1.DBS_COMPANY & "\MONTHLY\"
        If ASCMAIN1.CLIENT = "VAN" Then FOLDER = "R:\" & ASCMAIN1.DBS_COMPANY & "\MONTHLY\"
        ' FOLDER SHOULD USE ROOT

        If ASCMAIN1.Running_in_VS Then
            FOLDER = ASCMAIN1.Folders("Temp") & ASCMAIN1.DBS_COMPANY & "\MONTHLY\"
            Stop
        End If
        If Not My.Computer.FileSystem.DirectoryExists(FOLDER) Then
            My.Computer.FileSystem.CreateDirectory(FOLDER)
        End If

        Dim FILENAME As String = FOLDER & "MONTHLY.BAT"
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            My.Computer.FileSystem.DeleteFile(FILENAME)
        End If

        'after bat file exported we could not run bat file when it was pointing to e drive until
        Dim MMM As String = Mid(Format(Now, "MMMM"), 1, 3).ToUpper
        MMM = ASCMAIN1.CYP
        Dim F As String = "e:\Backups\monthend\" & ASCMAIN1.DBS_COMPANY & "\" & MMM
        Using SW As New System.IO.StreamWriter(FILENAME)
            SW.WriteLine("exp " & ASCMAIN1.DBS_COMPANY & "/" & ASCMAIN1.DBS_COMPANY & " file=" & F & ".dmp log=" & F & ".log")
            SW.WriteLine(Chr(34) & "C:\Program Files\WinZip\winzip64" & Chr(34) & " -m " & F & ".zip " & F & ".dmp " & F & ".log")
        End Using
    End Sub

End Class