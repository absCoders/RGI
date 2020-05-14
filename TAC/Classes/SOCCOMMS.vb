Public Class SOCCOMMS

    Public Enum NYACommissionCalcTypes
        Customer = 1
        CustomerAndDivision = 2
        StyleGroupOverride = 4
        CustomerOverride = 5
        CustomerAndStyleGroupOverride = 6
    End Enum

    Private Shared NYAGTemptable As String = String.Empty

    ''' <summary>
    ''' Query used to get the commision data for NYAG
    ''' Can be used for Creating a table data adapter
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared NYAGCommissionsQuery As String = "SELECT * FROM SOTCOMMS WHERE OPS_YYYYPP = :PARM1"


#Region "Class Properties"

    ''' <summary>
    ''' Returns the name of the temp work table for the NYAG Commissions
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared ReadOnly Property NYAGCommissionsWorktable As String
        Get
            If NYAGTemptable.Length = 0 Then
                ' Create the temp table shell.
                NYAGTemptable = ASCMAIN1.Temp_Table(NYAGCommissionsQuery.Replace(":PARM1", "'******' and Rownum < 1"))
            End If

            Return NYAGTemptable
        End Get
    End Property

#End Region

#Region "Class Procedures"

    ''' <summary>
    ''' Returns a temp table holding the commissions for the provided period.
    ''' The fields are SOTCOMMS.*
    ''' </summary>
    ''' <param name="Period"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNYAGCommissions(ByVal Period As String)
        Return GetNYAGCommissions(Period, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty)
    End Function

    ''' <summary>
    ''' Returns a temp table holding the commissions for the provided period.
    ''' The fields are SOTCOMMS.*
    ''' </summary>
    ''' <param name="Period"></param>
    ''' <param name="SOTCOMH1">temp table name</param>
    ''' <param name="SOTCOMH4">temp table name</param>
    ''' <param name="SOTCOMH5">temp table name</param>
    ''' <param name="SOTCOMH6">temp table name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNYAGCommissions(ByVal Period As String, ByVal SOTCOMH1 As String, _
                                              ByVal SOTCOMH4 As String, ByVal SOTCOMH5 As String, _
                                              ByVal SOTCOMH6 As String, ByVal SOTINVHS As String) As Boolean
        Dim SOTCOMMS As String = String.Empty

        Try

            If NYAGTemptable.Length = 0 Then
                NYAGTemptable = NYAGCommissionsWorktable
            End If

            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & NYAGTemptable)

            Dim endPeriod As String = String.Empty
            If Period.Contains("-") Then
                endPeriod = Period.Split("-")(1).Trim
                Period = Period.Split("-")(0).Trim
            End If
            If endPeriod.Length = 0 Then
                endPeriod = Period
            End If

            ASCMAIN1.sql = "Insert into " & NYAGTemptable & " " & NYAGCommissionsQuery.Replace("= :PARM1", " BETWEEN '" & Period & "' AND '" & endPeriod & "'")
            ASCDATA1.ExecuteSQL()

            ' If we get records (previous saved period) then pass back those records
            If Val(ASCDATA1.GetDataValue("select count(*) from " & NYAGTemptable) & String.Empty) > 0 Then
                Return True
            End If

            ' Customer / Sales Division Code
            ASCMAIN1.sql = "Insert into " & NYAGTemptable _
                & " Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE" _
                & ", CUST_CODE, ORDR_QTY_SHP, ORDR_AMT_SHP, SREP_COMM_AMT, OPS_YYYYPP, COMM_CALC_BY" _
                & " from " _
                & " ( " _
                & " Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY " _
                & " , SUM(ORDR_QTY_SHP) ORDR_QTY_SHP, SUM(ORDR_AMT_SHP) ORDR_AMT_SHP, SUM(SREP_COMM_AMT) SREP_COMM_AMT " _
                & " from " _
                & " ( " _
                & " SELECT " _
                & " SOTINVHS.INV_TYPE" _
                & ", SOTINVHS.INV_NO" _
                & ", SOTINVHS.SALES_DIVISION_CODE" _
                & ", ICTSTYL1.STYLE_GROUP_CODE" _
                & ", SOTINVHS.SREP_CODE" _
                & ", NVL(SOTINVHS.SREP_COMM_RATE, 0) SREP_COMM_RATE" _
                & ", SOTINVH1.CUST_CODE" _
                & ", nvl(SOTINVH2.ORDR_QTY_SHIP, 0) ORDR_QTY_SHP" _
                & ", nvl(SOTINVH2.ORDR_UNIT_PRICE, 0) * nvl(SOTINVH2.ORDR_QTY_SHIP, 0) ORDR_AMT_SHP" _
                & ", (nvl(SOTINVH2.ORDR_QTY_SHIP, 0)  * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0)) * NVL(SOTINVHS.SREP_COMM_RATE, 0)/100 SREP_COMM_AMT" _
                & ", SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" _
                & ", '" & NYACommissionCalcTypes.CustomerAndDivision & "' COMM_CALC_BY" _
                & " FROM SOTINVH1, SOTINVH2, " & SOTINVHS & " SOTINVHS, ICTSTYL1" _
                & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " AND SOTINVH1.INV_TYPE = SOTINVHS.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVHS.INV_NO" _
                & " AND SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN '" & Period & "' AND '" & endPeriod & "'" _
                & " AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & " AND ICTSTYL1.SALES_DIVISION_CODE = SOTINVHS.SALES_DIVISION_CODE " _
                & " ) " _
                & " GROUP BY INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY" _
                & ")"
            ASCDATA1.ExecuteSQL()


            ' SOTCOMH6 - Customer / Style Group
            ASCMAIN1.sql = "Insert into " & NYAGTemptable _
                & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE" _
                & " , CUST_CODE, ORDR_QTY_SHP, ORDR_AMT_SHP, SREP_COMM_AMT, OPS_YYYYPP, COMM_CALC_BY" _
                & "  from " _
                & "  ( " _
                & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY " _
                & "  , SUM(ORDR_QTY_SHP) ORDR_QTY_SHP, SUM(ORDR_AMT_SHP) ORDR_AMT_SHP, SUM(SREP_COMM_AMT) SREP_COMM_AMT " _
                & "  from " _
                & "  ( " _
                & "  SELECT " _
                & "  SOTINVH1.INV_TYPE" _
                & " , SOTINVH1.INV_NO" _
                & " , ICTSTYL1.SALES_DIVISION_CODE" _
                & " , ICTSTYL1.STYLE_GROUP_CODE" _
                & " , SOTCOMH6.SREP_CODE" _
                & " , NVL(SOTCOMH6.SREP_COMM_RATE, 0) SREP_COMM_RATE" _
                & " , SOTINVH1.CUST_CODE" _
                & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) ORDR_QTY_SHP" _
                & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0) ORDR_AMT_SHP" _
                & " , (nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0)) * NVL(SOTCOMH6.SREP_COMM_RATE, 0)/100 SREP_COMM_AMT" _
                & " , SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" _
                & " , '" & NYACommissionCalcTypes.CustomerAndStyleGroupOverride & "' COMM_CALC_BY" _
                & "  FROM SOTINVH1, SOTINVH2, " & SOTCOMH6 & " SOTCOMH6, ICTSTYL1" _
                & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " AND SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN '" & Period & "' AND '" & endPeriod & "'" _
                & " AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & " AND SOTINVH1.ORDR_YYYYPP_UPDATED = SOTCOMH6.OPS_YYYYPP" _
                & " AND SOTINVH1.CUST_CODE = SOTCOMH6.CUST_CODE" _
                & " AND ICTSTYL1.STYLE_GROUP_CODE = SOTCOMH6.STYLE_GROUP_CODE " _
                & " AND (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH2.STYLE_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED, SOTCOMH6.SREP_CODE) " _
                & " NOT IN (SELECT INV_TYPE, INV_NO, STYLE_GROUP_CODE, OPS_YYYYPP, SREP_CODE FROM " & NYAGTemptable & ")" _
                & " )" _
                & " GROUP BY INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY" _
                & " ) "
            ASCDATA1.ExecuteSQL()

            ' SOTCOMH5 - Customer
            ASCMAIN1.sql = "Insert into " & NYAGTemptable _
               & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE" _
               & " , CUST_CODE, ORDR_QTY_SHP, ORDR_AMT_SHP, SREP_COMM_AMT, OPS_YYYYPP, COMM_CALC_BY" _
               & "  from " _
               & "  ( " _
               & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY " _
               & "  , SUM(ORDR_QTY_SHP) ORDR_QTY_SHP, SUM(ORDR_AMT_SHP) ORDR_AMT_SHP, SUM(SREP_COMM_AMT) SREP_COMM_AMT " _
               & "  from " _
               & "  ( " _
               & "  SELECT " _
               & "  SOTINVH1.INV_TYPE" _
               & " , SOTINVH1.INV_NO" _
               & " , ICTSTYL1.SALES_DIVISION_CODE" _
               & " , ICTSTYL1.STYLE_GROUP_CODE" _
               & " , SOTCOMH5.SREP_CODE" _
               & ", DECODE(NVL(SOTCOMH5.SREP_COMM_USE_STD, '0'), '1', NVL(SOTCOMH1.SREP_COMM_RATE, 0), NVL(SOTCOMH5.SREP_COMM_RATE, 0)) SREP_COMM_RATE " _
               & " , SOTINVH1.CUST_CODE" _
               & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) ORDR_QTY_SHP" _
               & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0) ORDR_AMT_SHP" _
               & " , (nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0)) * NVL(SOTCOMH5.SREP_COMM_RATE, 0)/100 SREP_COMM_AMT" _
               & " , SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" _
               & " , '" & NYACommissionCalcTypes.CustomerOverride & "' COMM_CALC_BY" _
               & "  FROM SOTINVH1, SOTINVH2, " & SOTCOMH1 & " SOTCOMH1, " & SOTCOMH5 & " SOTCOMH5, ICTSTYL1" _
               & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
               & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
               & " AND SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN '" & Period & "' AND '" & endPeriod & "'" _
               & " AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
               & " AND SOTINVH1.ORDR_YYYYPP_UPDATED = SOTCOMH5.OPS_YYYYPP" _
               & " AND SOTINVH1.CUST_CODE = SOTCOMH5.CUST_CODE" _
               & " AND SOTCOMH1.SREP_CODE = SOTCOMH5.SREP_CODE" _
               & " AND SOTCOMH1.OPS_YYYYPP = SOTCOMH5.OPS_YYYYPP" _
               & " AND (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, ICTSTYL1.STYLE_GROUP_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED, SOTCOMH5.SREP_CODE) " _
               & " NOT IN (SELECT INV_TYPE, INV_NO, STYLE_GROUP_CODE, OPS_YYYYPP, SREP_CODE FROM " & NYAGTemptable & ")" _
               & " )" _
               & " GROUP BY INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY" _
               & " ) "
            ASCDATA1.ExecuteSQL()

            ' SOTCOMH4 - Style Group
            ASCMAIN1.sql = "Insert into " & NYAGTemptable _
                & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE" _
                & " , CUST_CODE, ORDR_QTY_SHP, ORDR_AMT_SHP, SREP_COMM_AMT, OPS_YYYYPP, COMM_CALC_BY" _
                & "  from " _
                & "  ( " _
                & "  Select INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY " _
                & "  , SUM(ORDR_QTY_SHP) ORDR_QTY_SHP, SUM(ORDR_AMT_SHP) ORDR_AMT_SHP, SUM(SREP_COMM_AMT) SREP_COMM_AMT " _
                & "  from " _
                & "  ( " _
                & "  SELECT " _
                & "  SOTINVH1.INV_TYPE" _
                & " , SOTINVH1.INV_NO" _
                & " , ICTSTYL1.SALES_DIVISION_CODE" _
                & " , ICTSTYL1.STYLE_GROUP_CODE" _
                & " , SOTCOMH4.SREP_CODE" _
                & " , NVL(SOTCOMH4.SREP_COMM_RATE, 0) SREP_COMM_RATE" _
                & " , SOTINVH1.CUST_CODE" _
                & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) ORDR_QTY_SHP" _
                & " , nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0) ORDR_AMT_SHP" _
                & " , (nvl(SOTINVH2.ORDR_QTY_SHIP, 0) * nvl(SOTINVH2.ORDR_UNIT_PRICE, 0)) * NVL(SOTCOMH4.SREP_COMM_RATE, 0)/100 SREP_COMM_AMT" _
                & " , SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" _
                & " , '" & NYACommissionCalcTypes.StyleGroupOverride & "' COMM_CALC_BY" _
                & "  FROM SOTINVH1, SOTINVH2, " & SOTCOMH4 & " SOTCOMH4, ICTSTYL1" _
                & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " AND SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN '" & Period & "' AND '" & endPeriod & "'" _
                & " AND SOTINVH2.STYLE_CODE = ICTSTYL1.STYLE_CODE" _
                & " AND SOTINVH1.ORDR_YYYYPP_UPDATED = SOTCOMH4.OPS_YYYYPP" _
                & " AND ICTSTYL1.STYLE_GROUP_CODE = SOTCOMH4.STYLE_GROUP_CODE" _
                & " AND (SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, ICTSTYL1.STYLE_GROUP_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED, SOTCOMH4.SREP_CODE) " _
                & " NOT IN (SELECT INV_TYPE, INV_NO, STYLE_GROUP_CODE, OPS_YYYYPP, SREP_CODE FROM " & NYAGTemptable & ")" _
                & " )" _
                & " GROUP BY INV_TYPE, INV_NO, SALES_DIVISION_CODE, STYLE_GROUP_CODE, SREP_CODE, SREP_COMM_RATE, CUST_CODE, OPS_YYYYPP, COMM_CALC_BY" _
                & " ) "
            ASCDATA1.ExecuteSQL()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        Finally

        End Try

        Return True
    End Function

#End Region

End Class
