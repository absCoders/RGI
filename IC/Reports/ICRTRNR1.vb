Public Class ICRTRNR1
    'Dim RYP As String
    'Dim RYPLegend As String
    Dim chkType As String
    Dim sqlw As String
    Dim optINV As String
    Dim CHKFINV_DATE As String
    Dim CHKTINV_DATE As String
    Dim SSDFINV_DATE As String
    Dim SSDTINV_DATE As String
    Dim SUBTITLE As String
    Dim CHKUSEMASTER As String
    Dim CHKAIR As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Range_Events(grpTRAN_DATE)
        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, -1)
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        With dst
            'Dim i As Integer
            'Dim z As String
            'Dim item As String

            ' Get Codes selected from tabs

            ASCMAIN1.Progress("Run-Time Options", "")

            ' Get Run-Time options


            'RYP = Absx1.cbeFor("RYP").Value
            'RYPLEGEND = ASCMAIN1.Get_Legend(RYP)

            CHKUSEMASTER = chkUseMasterFile.Checked
            CHKAIR = chkAirShipments.Checked
            'chkType = Get_chk("TYPE", "SCRAT", "Y")

            chkType = ""
            If chkShipments.Checked Then chkType &= ",'S'"
            If chkReturns.Checked Then chkType &= ",'C'"
            If chkReceipts.Checked Then chkType &= ",'R'"
            If chkAdjustments.Checked Then chkType &= ",'A'"
            If chkTransfers.Checked Then chkType &= ",'T'"
            chkType = Mid(chkType, 2)

            Select Case optPRofAllTransactions.CheckedItem.DataValue
                Case "Period"
                    RYP = Absx1.cbeFor("RYP").Value
                    RYPLEGEND = ASCMAIN1.Get_Legend(RYP)

                Case "Not Uploaded"
                    RYP = ASCMAIN1.CYP
                    RYPLEGEND = ASCMAIN1.Get_Legend(RYP)

                Case "Date"
                    'CHKFINV_DATE = SRead(opts, "CHKFINV_DATE", 2)
                    'CHKTINV_DATE = SRead(opts, "CHKTINV_DATE", 2)
                    'SSDFINV_DATE = SRead(opts, "SSDFINV_DATE", 2)
                    'SSDTINV_DATE = SRead(opts, "SSDTINV_DATE", 2)
            End Select


            ' Set up Work File Definition using X's and 0's and 0.01's as required

            ASCMAIN1.Progress("Initialize Work Tables", "")

            ' Prepare Work File with Data from Server

            ASCMAIN1.Progress("Preparing Result Set on Server", "")

            sql = ""
            Select Case optPRofAllTransactions.CheckedItem.DataValue
                Case "Not Uploaded"
                    sql &= " and ICTTRAN1.TRAN_STATUS_PRT is Null"
                Case "Period"
                    sql &= "   and ICTTRAN1.OPS_YYYYPP = '" & RYP & "'"
                Case "Date"
                    sql &= Get_Dates()
            End Select

            If chkType <> "" Then
                sql &= "   and ICTTRAN1.TRAN_TYPE in (" & chkType & ")"
            End If


            Dim sqlw As String
            sqlw = ASCMAIN1.SQL_Add_WHERE(sql)

            ASCMAIN1.Progress("Transactions Header", "")
            ASCMAIN1.sql = "Select ICTTRAN1.* from ICTTRAN1 " & sqlw
            Create_TDA(.Tables.Add, "ICTTRAN1", "**", 0, True, "", 3)
            Fill_Records("ICTTRAN1")

            ASCMAIN1.Progress("Transactions Detail", "")
            ASCMAIN1.sql = "" _
            & "Select ICTTRAN2.*,POTORDR1.PO_REFERENCE,POTORDR1.PO_SPEC_ORDR_NO" & vbCrLf _
            & ", ICTSTYL1.SUB_UNIT_PACK_QTY, ICTSTYL1.STYLE_COST STYLE_COST_MASTER" & vbCrLf _
            & " from ICTTRAN2, ICTTRAN1, POTORDR1, ICTSTYL1 " & sqlw & vbCrLf _
            & " and ICTTRAN2.OPS_YYYYPP = ICTTRAN1.OPS_YYYYPP " & vbCrLf _
            & " and ICTTRAN2.TRAN_TYPE = ICTTRAN1.TRAN_TYPE " & vbCrLf _
            & " and ICTTRAN2.TRAN_NO = ICTTRAN1.TRAN_NO " & vbCrLf _
            & " and POTORDR1.PO_ORDER_NO = ICTTRAN2.PO_ORDER_NO " & vbCrLf _
            & " and ICTTRAN2.STYLE_CODE = ICTSTYL1.STYLE_CODE "
            Create_TDA(.Tables.Add, "ICTTRAN2", "**", 0, False, "", 4)
            Fill_Records("ICTTRAN2")

            Dim tranRelation As DataRelation = Create_Relation("ICTTRAN1", "ICTTRAN2", "OPS_YYYYPP,TRAN_TYPE,TRAN_NO")

            ASCMAIN1.sql = "Select * from POTSHIP1 where PO_SHIPMENT_NO in (" & vbCrLf _
                & " SELECT TRAN_NO_ORIG from ICTTRAN1 " & vbCrLf & sqlw & ")"
            Create_TDA(.Tables.Add, "POTSHIP1", "**", 0, False, "", 1)
            Fill_Records("POTSHIP1")

            ASCMAIN1.sql = "Select * from POTSHIP2 where PO_SHIPMENT_NO in (" & vbCrLf _
                & " SELECT TRAN_NO_ORIG from ICTTRAN1 " & vbCrLf & sqlw & ")"
            Create_TDA(.Tables.Add, "POTSHIP2", "**", 0, False, "", 2)
            Fill_Records("POTSHIP2")

            Dim shipRelation As DataRelation = Create_Relation("POTSHIP1", "POTSHIP2", "PO_SHIPMENT_NO")


            'Remove all but Air Shipments if user selects this.
            If chkAirShipments.Checked Then

                For Each rowPOTSHIP1 As DataRow In .Tables("POTSHIP1").Select("AIR_SHIP <> '1' OR AIR_SHIP IS NULL")
                    For Each rowPOTSHIP2 As DataRow In rowPOTSHIP1.GetChildRows(shipRelation)
                        rowPOTSHIP2.Delete()
                    Next
                    rowPOTSHIP1.Delete()
                Next
                .Tables("POTSHIP1").AcceptChanges()
                .Tables("POTSHIP2").AcceptChanges()

                Dim validTranNumbers As DataTable = ASCDATA1.SelectDistinct("POTSHIP2", "TRAN_NO")
                validTranNumbers.PrimaryKey = New DataColumn() {validTranNumbers.Columns("TRAN_NO")}
                For Each rowICTTRAN1 As DataRow In .Tables("ICTTRAN1").Select
                    If validTranNumbers.Select("TRAN_NO = " & rowICTTRAN1.Item("TRAN_NO")).Length = 0 Then
                        For Each rowICTTRAN2 As DataRow In rowICTTRAN1.GetChildRows(tranRelation)
                            rowICTTRAN2.Delete()
                        Next
                        rowICTTRAN1.Delete()
                    End If
                Next

                .Tables("ICTTRAN1").AcceptChanges()
                .Tables("ICTTRAN2").AcceptChanges()

            End If
        End With
    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"TRAN_DATE"}
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and ICTTRAN1." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("chk" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and ICTTRAN1." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        Return sql
    End Function

    Public Overrides Sub Print_Report()
        Dim P As String

        If Absx1.chkFor("AIR_SHIP").Checked Then
            SUBT = SUBT & " - Air Shipments Only "
        End If
        P = "Y"

        Select Case optPRofAllTransactions.CheckedItem.DataValue
            Case "Not Uploaded"
                SUBT = "New Transactions"
            Case "Date"
                SUBT = "Invoice Date Range from "

                If Absx1.chkFor("CHKTRAN_DATE_F").Checked Then
                    SUBT &= "First"
                Else
                    SUBT &= Format$(SSDFINV_DATE, "mm/dd/yy")
                End If
                SUBT &= " to "
                If Absx1.chkFor("CHKTRAN_DATE_L").Checked Then
                    SUBT &= "Last"
                Else
                    SUBT &= Format$(SSDTINV_DATE, "mm/dd/yy")
                End If
            Case "Period"
                SUBT = "Transactions Posted in " & RYPLEGEND
        End Select

        Dim master As String = Absx1.chkFor("USE_MASTER").CheckedValue

            CR_params.Add("PRE", P)
        CR_params.Add("MASTER", master)

            Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        For Each rowICTTRAN1 As DataRow In dst.Tables("").Select("")
            rowICTTRAN1.Item("TRAN_STATUS_PRT") = "1"
        Next
        Update_Record_TDA("ICTTRAN1")

        'Dim dynICTTRAN1 As OraDynaset
        'sql = "Select * from ICTTRAN1 where OPS_YYYYPP = :PARM1"
        'sql = sql & " and TRAN_TYPE := PARM2"
        'sql = sql & " and TRAN_NO   := PARM3"
        'dynICTTRAN1 = OraD.CreateDynaset(sql, 0&)

        'Dim dynICWTRAN1 As Recordset
        'dynICWTRAN1 = AccD.OpenRecordset("ICWTRAN1", dbOpenForwardOnly)
        'Do While Not dynICWTRAN1.EOF
        '    OraD.Parameters("OPS_YYYYPP").Value = dynICWTRAN1.Fields("OPS_YYYYPP").Value
        '    OraD.Parameters("CODE1").Value = dynICWTRAN1.Fields("TRAN_TYPE").Value
        '    OraD.Parameters("CODE2").Value = dynICWTRAN1.Fields("TRAN_NO").Value
        '    dynICTTRAN1.Refresh()
        '    dynICTTRAN1.Edit()
        '    dynICTTRAN1.Fields("TRAN_STATUS_PRT").Value = "1"
        '    dynICTTRAN1.Update()
        '    dynICWTRAN1.MoveNext()
        'Loop
        'dynICWTRAN1.Close()
        'dynICTTRAN1.Close()

        'OraS.CommitTrans()
        'Call Done()
    End Sub

    Private Sub optPRofAllTransactions_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPRofAllTransactions.ValueChanged
        Select Case optPRofAllTransactions.CheckedItem.DataValue
            Case "Not Uploaded"
                grpOPS_YYYPP.Visible = False
                grpTRAN_DATE.Visible = False
            Case "Date"
                grpOPS_YYYPP.Visible = False
                grpTRAN_DATE.Visible = True
            Case "Period"
                grpOPS_YYYPP.Visible = True
                grpTRAN_DATE.Visible = False
        End Select
    End Sub
End Class