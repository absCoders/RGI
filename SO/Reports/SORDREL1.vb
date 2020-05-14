Public Class SORDREL1
    'Dim sql_pick As String          ' where clause for all SOTPICK1 in scope of the De-Release with PICK_STATUS of P or C
    'Dim sql_pick_all As String      ' where clause for all SOTPICK1 in scope of the De-Release

    Dim SOTPICK1 As String
    Dim SOTSHIP1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        Dim sql_pick As String = Set_filter_for_Pick_Tickets_to_De_Release()

        ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1" & ASCMAIN1.SQL_Add_WHERE(sql_pick)
        SOTPICK1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPICK1 & " Add Primary Key (PICK_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTPICK1 & "_1 on " & SOTPICK1 & " (SHIP_BOL_NO)")

        ASCMAIN1.sql = "Select * from " & SOTPICK1
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTPICK1", 1))

        ASCMAIN1.sql = "Select SOTSHIP1.*,SOTORDR0.CUST_CODE from SOTSHIP1,SOTORDR0" _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" _
            & " and SOTSHIP1.SHIP_BOL_NO in " _
            & " (Select DISTINCT SHIP_BOL_NO from " & SOTPICK1 & ")"
        SOTSHIP1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP1 & " Add Primary Key (SHIP_BOL_NO)")

        '& " (Select DISTINCT SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")"

        ASCMAIN1.sql = "Select * from " & SOTSHIP1
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIP1", 1))

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        dst.Tables("SOTSHIP1").Columns.Add("PICK_CNT", GetType(System.Int32), "COUNT(CHILD.PICK_NO)")

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
            If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                RWU = "N"
                RWU &= "0"
                xErrMsg = "Could Not Lock All Shipments"
            End If
        Next
        ASCMAIN1.sql = "Select Count (*) from " & SOTPICK1
        Dim C As Int64 = Val(ASCDATA1.GetDataValue)
        If C > 100 Then
            If MsgBox("An Excessive number of Pick Tickets are queued up to be De-Released" _
                      & vbCrLf & vbCrLf & "Are you sure that you want to Continue with the De-Release process?", _
                      MsgBoxStyle.YesNo, "Verfication - Over 100 Pick Tickets will be De-Released") = MsgBoxResult.No Then
                RWU = "N"
                MsgBox("Report will print, but Update will be Disabled", MsgBoxStyle.OkOnly, "Verification")
            End If
        End If

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim C As Integer = tblASTDSQLA.Select("ISNULL(CODE_VALUES,'') <> ''").Length ' tblASTDSQLA.Select("CODE_VALUES IS NOT NULL").Length
            If C < 1 Then
                EMsg &= vbCr & "You Must Specify at least 1 Pick Batch, Ship BOL No, or Pick Ticket"
            ElseIf C > 1 Then
                EMsg &= vbCr & "You Cannot Mix Pick Batches, Ship BOL No's and Pick Tickets in a Single Execution"
            End If

            If tblASTDSQLA.Select("EXCLUDE = '1'").Length <> 0 Then
                EMsg &= vbCr & "You may not use Exclusion on any Filter for De-Release"
            End If

            Dim sql_pick As String = Set_filter_for_Pick_Tickets_to_De_Release()

            If SQLA("PICK_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick)
                If Val(ASCDATA1.GetDataValue) = 0 Then
                    EMsg &= vbCr & "No Pick Tickets to De-Release"
                End If
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" _
                    & " and SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")"
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected are Not In Pick"
                End If
            End If

            If SQLA("SHIP_BOL_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" & SQL_in("SHIP_BOL_NO")
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected are Not In Pick"
                End If
            End If

            If SQLA("PICK_BATCH_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTPICK0 where NVL(PICK_BATCH_STATUS,'?') <> 'O'" & SQL_in("PICK_BATCH_NO")
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Pick Batches Selected are Not In Open"
                End If
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_STATUS,'?') <> 'P'" & SQL_in("PICK_BATCH_NO")
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments in the Pick Batch Selected are Not In Pick"
                End If
            End If

            ' Check LP_STATUS on any SHIP_BOL_NO that is touched by this De-Release

            Dim sqlWHSE_CODE_3PL As String = " and WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where LP_CODE is Not Null)"
            If ASCMAIN1.CLIENT = "RGI" Then
                'Allow 'CG' to de-release
                sqlWHSE_CODE_3PL = " and WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where LP_CODE is Not Null and ICTWHSE1.WHSE_CODE <>'CG')"
            End If
            Dim sqlWHSE_CODE_LOC As String = " and WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_LOCATOR = '1' and WHSE_CTN_CTL = 'C')"

            If SQLA("PICK_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" & vbCrLf _
                    & " and SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")" & vbCrLf _
                    & sqlWHSE_CODE_3PL
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected have been Transmitted to the 3PL"
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_WAVE_STATUS,'0') <> '0'" & vbCrLf _
                    & " and SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from SOTPICK1 " & ASCMAIN1.SQL_Add_WHERE(sql_pick) & ")" & vbCrLf _
                    & sqlWHSE_CODE_LOC
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected have been Waved"
                End If

            End If

            If SQLA("SHIP_BOL_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" & SQL_in("SHIP_BOL_NO") & vbCrLf _
                    & sqlWHSE_CODE_3PL
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected have been Transmitted to the 3PL"
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_WAVE_STATUS,'0') <> '0'" & SQL_in("SHIP_BOL_NO") & vbCrLf _
                    & sqlWHSE_CODE_LOC
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments Selected have been Waved"
                End If
            End If

            If SQLA("PICK_BATCH_NO") <> "" Then
                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(LP_STATUS,'?') = '1'" & SQL_in("PICK_BATCH_NO") & vbCrLf _
                    & sqlWHSE_CODE_3PL
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments in the Pick Batch Selected have been Transmitted to the 3PL"
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTSHIP1 where NVL(SHIP_WAVE_STATUS,'0') <> '0'" & SQL_in("PICK_BATCH_NO") & vbCrLf _
                    & sqlWHSE_CODE_LOC
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Some Shipments in the Pick Batch Selected have been Waved"
                End If
            End If

            End If
    End Sub

    Overrides Sub Update_Record()
        DeRelease()
    End Sub

    Sub DeRelease()

        If ASCMAIN1.Running_in_VS Then Stop

        ' THIS PROCESS ALLOWS THE USER TO DERELEASE BY PICK BATCH NO, PICK NO, SHIPMENT_NO (MAYBE GROUP?)
        ' THE RESULT IS THAT ALL PICK TICKETS TO BE DE-RELEASED ARE LOADED INTO TEMP TABLE SOTPICK1

        'BeginTrans() - BECAUSE BEGINTRANS IS CALLED IN ASTSRPTM

        ASCMAIN1.Progress("Now De-Releasing Pick Tickets", "")

        If SQLA("PICK_BATCH_NO") <> "" Then
            ASCMAIN1.sql = "Update SOTPICK0 Set " & vbCrLf _
                 & "  PICK_BATCH_STATUS = 'D'" & vbCrLf _
                 & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                 & ", LAST_DATE = SYSDATE" & vbCrLf _
                 & " where PICK_BATCH_NO in (" & SQLA("PICK_BATCH_NO", , True) & ")"
            ASCDATA1.ExecuteSQL()
        End If

        TAC.SOCMAIN1.DeRelease(SOTPICK1)

        If ASCMAIN1.Running_in_VS Then Stop

    End Sub

    Function Set_filter_for_Pick_Tickets_to_De_Release() As String

        Dim sql_pick As String = ""
        sql_pick &= SQL_in("PICK_BATCH_NO", "SOTPICK1.PICK_BATCH_NO")
        sql_pick &= SQL_in("SHIP_BOL_NO", "SOTPICK1.SHIP_BOL_NO")
        sql_pick &= SQL_in("PICK_NO", "SOTPICK1.PICK_NO")

        'sql_pick_all = sql_pick
        sql_pick = sql_pick & " and (SOTPICK1.PICK_STATUS = 'P' OR SOTPICK1.PICK_STATUS = 'C')"

        Return sql_pick
    End Function


    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "ORDR_GROUP_NO"
                sqlw = " SOTORDR0.ORDR_CNT_PICK > 0"
            Case "PICK_NO"
                sqlw = " SOTPICK1.PICK_STATUS = 'P'"
            Case "SHIP_BOL_NO"
                sqlw = " SOTSHIP1.SHIP_STATUS = 'P'"

        End Select
        Return sqlw
    End Function
End Class