Imports System.Text

Public Class SORSCOMD

    'Dim SOTINVH1 As String
    'Dim SOTINVH2 As String
    Dim S As New StringBuilder With {.Length = 0}
    Dim DISC_DATE As Date = Now()
    Dim DISC_PCT As Double = 10

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0), -60, 0, 0)

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "N"

        'Dim WR As String = sqlw
        Dim RYP As String = Mid(cboRYP.Text, 1, 4) & Mid(cboRYP.Text, 6, 2)
        Dim SREPS As String = ""
        Dim SREPS_IN As String = ""
        Dim WHSE As String = ""
        Dim WHSE_LIST As String() = {}
        Dim WHSE_NOT_IN As Boolean = False
        If tblASTDSQLA.Rows.Count > 0 Then
            For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select()
                Select Case rowASTDSQLA.Item("COLUMN_NAME")
                    Case "SREP_CODE"
                        SREPS = rowASTDSQLA.Item("CODE_VALUES").ToString & String.Empty
                        If SREPS.Length > 0 Then
                            SREPS = SREPS.Replace(",", "','")
                            SREPS = "'" & SREPS & "'"
                            If rowASTDSQLA.Item("EXCLUDE") = "1" Then
                                SREPS_IN = " NOT "
                            End If
                        End If
                    Case "WHSE_CODE"
                        WHSE = rowASTDSQLA.Item("CODE_VALUES").ToString & String.Empty
                        If WHSE.Length > 0 Then
                            WHSE_LIST = WHSE.Split(",")
                            WHSE = WHSE.Replace(",", "','")
                            WHSE = "'" & WHSE & "'"
                            If rowASTDSQLA.Item("EXCLUDE") = "1" Then
                                WHSE_NOT_IN = True
                            End If
                        End If
                End Select
            Next

        End If

        S.Length = 0
        S.AppendLine("SELECT")
        S.AppendLine("I1.INV_NO,")
        S.AppendLine("NVL(O1.ORDR_DATE_RECD,'30-JUN-2021') AS ORDR_DATE_RECD,")
        S.AppendLine("NVL(I1.INV_DATE,'30-JUN-2021') AS INV_DATE,")
        S.AppendLine("NVL(O1.ORDR_NO,'CREDIT') AS ORDR_NO,")
        S.AppendLine("NVL(O1.CUST_CODE,'CREDIT') AS CUST_CODE,")
        S.AppendLine("NVL(O1.CUST_NAME,'CREDIT') AS CUST_NAME,")
        S.AppendLine("I1.SREP_CODE,")
        S.AppendLine("I1.WHSE_CODE,")
        S.AppendLine("SUM(I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) AS TOTAL_SALES,")
        S.AppendLine("SUM((I2.ORDR_UNIT_PRICE * I2.ORDR_QTY_SHIP) * (COMM_RATE/100)) AS COMMISSION,")
        S.AppendLine("0.00 AS COMM_PCT,")
        S.AppendLine("0.00 AS COMM_CALC,")
        S.AppendLine("1 AS PRIOR_GROUP")
        S.AppendLine("FROM SOTINVH1 I1, SOTINVH2 I2, SOTORDR1 O1")
        S.AppendLine("WHERE I1.INV_TYPE = I2.INV_TYPE")
        S.AppendLine("AND I1.INV_NO = I2.INV_NO")
        S.AppendLine("AND I1.ORDR_NO = O1.ORDR_NO (+)")
        S.AppendLine($"AND I1.ORDR_YYYYPP_UPDATED = '{RYP}'")
        If SREPS.Length > 0 Then
            S.AppendLine($"AND I1.SREP_CODE {SREPS_IN} IN ({SREPS})")
        End If
        'If WHSE.Length > 0 Then
        '    If WHSE_NOT_IN Then
        '        S.AppendLine($"AND I1.WHSE_CODE NOT IN ({WHSE})")
        '    Else
        '        S.AppendLine($"AND I1.WHSE_CODE IN ({WHSE})")
        '    End If
        'End If
        S.AppendLine("GROUP BY I1.INV_NO,")
        S.AppendLine("O1.ORDR_NO,")
        S.AppendLine("O1.ORDR_DATE_RECD,")
        S.AppendLine("I1.INV_DATE,")
        S.AppendLine("O1.CUST_CODE,")
        S.AppendLine("O1.CUST_NAME,")
        S.AppendLine("I1.SREP_CODE,")
        S.AppendLine("I1.WHSE_CODE")
        ASCMAIN1.sql = S.ToString
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTRCOMD", 1))
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTRCOMD", 1))
        dst.Tables.Item("SOTRCOMD").Columns.Item("COMM_PCT").ReadOnly = False
        dst.Tables.Item("SOTRCOMD").Columns.Item("COMM_CALC").ReadOnly = False
        dst.Tables.Item("SOTRCOMD").Columns.Item("PRIOR_GROUP").ReadOnly = False

        DISC_DATE = CDate(dteDiscountDate.Text.ToString)
        DISC_PCT = Val(numDiscount.Value)

        For Each rowSOTRCOMD As DataRow In dst.Tables("SOTRCOMD").Select()
            Dim ORDR_DATE_RECD As DateTime = CDate(rowSOTRCOMD.Item("ORDR_DATE_RECD").ToString & String.Empty)
            Dim WHSE_CODE As String = rowSOTRCOMD.Item("WHSE_CODE").ToString & String.Empty
            Dim WHSE_WHSE_EXCLD As Boolean = False
            If WHSE_NOT_IN = True Then
                If WHSE_LIST.Contains(WHSE_CODE) Then
                    WHSE_WHSE_EXCLD = True
                End If
            End If
            If ORDR_DATE_RECD < DISC_DATE Then
                If WHSE_WHSE_EXCLD Then
                    rowSOTRCOMD.Item("PRIOR_GROUP") = 0
                    rowSOTRCOMD.Item("COMM_PCT") = 0
                    rowSOTRCOMD.Item("COMM_CALC") = Val(rowSOTRCOMD.Item("COMMISSION").ToString & String.Empty)
                Else
                    rowSOTRCOMD.Item("PRIOR_GROUP") = 0
                    rowSOTRCOMD.Item("COMM_PCT") = DISC_PCT
                    rowSOTRCOMD.Item("COMM_CALC") = Val(rowSOTRCOMD.Item("COMMISSION").ToString & String.Empty) * (1 - (DISC_PCT / 100))
                End If
            Else
                rowSOTRCOMD.Item("PRIOR_GROUP") = 1
                rowSOTRCOMD.Item("COMM_PCT") = 0
                rowSOTRCOMD.Item("COMM_CALC") = Val(rowSOTRCOMD.Item("COMMISSION").ToString & String.Empty)
            End If
        Next

        'With dst.Tables("SOTRCOMD").Columns
        '    .Add("COMM_PCT", GetType(System.Decimal))
        '    .Add("COMM_CALC", GetType(System.Decimal))
        'End With


        'Create_TDA(dst.Tables.Add("SOTINVH1"), SOTINVH1, "**", 0, False, "", 2)
        'Fill_Records("SOTRCOMD")

        ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1))
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("DISC_DATE", DISC_DATE)
        CR_params.Add("DISC_PCT", DISC_PCT)
        CR_params.Add("SUBT", $"Discount Orders Prior To {DISC_DATE.ToShortDateString()} With {DISC_PCT}% Discount.")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class