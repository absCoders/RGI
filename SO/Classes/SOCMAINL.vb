Public Class SOCMAINL

    Public Shared Function IsValidTerms(ByVal CUST_CODE As String, ByVal SEL_TERM_CODE As String) As Boolean
        Dim RetVal As Boolean = False
        Dim sql As New Text.StringBuilder
        sql.Length = 0
        sql.AppendLine("SELECT NVL(TERM_CODE,'CRED') AS TERM_CODE")
        sql.AppendLine("FROM ARTCUST1")
        sql.AppendLine(String.Format("WHERE CUST_CODE = '{0}'", CUST_CODE))
        ASCMAIN1.sql = sql.ToString
        Dim CUST_TERM_CODE As String = ASCDATA1.GetDataValue
        If SEL_TERM_CODE = CUST_TERM_CODE Then
            RetVal = True
        Else
            Select Case CUST_TERM_CODE
                Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D"
                    Select Case SEL_TERM_CODE
                        Case Is = "N30", "N30D", "N30ROG", "N45D", "N60", "N90", "N90D", "COD", "CBD", "CRED", "XMAS", "FALL"
                            RetVal = True
                        Case Else
                            MsgBox("Invalid Terms Code For This Customer", MsgBoxStyle.OkOnly, "Invalid Terms")
                            RetVal = False
                    End Select
                Case Is = "CRED", "COD", "CBD"
                    Select Case SEL_TERM_CODE
                        Case Is = "CRED", "COD", "CBD"
                            RetVal = True
                        Case Else
                            MsgBox(String.Format("Customers With {0} Terms Code Must Select CRED, COD or CBD Terms", CUST_TERM_CODE), MsgBoxStyle.OkOnly, "Invalid Terms")
                            RetVal = False
                    End Select
                Case Is = "AMEX"
                    MsgBox("AMEX Is No Longer Supported As A Terms Code" & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
                    RetVal = False
                Case Else
                    MsgBox("Customer Terms Code Is " & CUST_TERM_CODE, MsgBoxStyle.OkOnly, "Invalid Terms")
                    RetVal = False
            End Select
        End If
        Return RetVal
    End Function

    Private Function Pop_Control_No(ByVal TABLE_NAME As String, ByVal COLUMN_NAME As String) As String
        Dim RetVal As String = ""
        Dim VERBTYPE As String = ""
        Dim MSG As String = ""
        Dim RecsRemaining As Integer = 0
        Dim MAXREC As Integer = 0
        Dim MINREC As Integer = 0
        Dim CTLLENGTH As Integer = 0
        Select Case TABLE_NAME
            Case "SOTORDR1"
                VERBTYPE = "Order Numbers"
            Case "ARTCUST1"
                VERBTYPE = "Customer Numbers"
        End Select

        Dim sql As New Text.StringBuilder
        sql.Length = 0
        sql.AppendLine("SELECT MAX(CTL_NO_NEXT) AS MAXREC")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        MAXREC = Val(ASCDATA1.GetDataValue)
        sql.Length = 0
        sql.AppendLine("SELECT MIN(CTL_NO_NEXT) AS MINREC")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        MINREC = Val(ASCDATA1.GetDataValue)
        sql.Length = 0
        sql.AppendLine("SELECT MAX(CTL_NO_LENGTH) as CTLLENGTH")
        sql.AppendLine(" FROM TATCTLNL")
        sql.AppendLine(String.Format(" WHERE TABLE_NAME = '{0}'", TABLE_NAME))
        sql.AppendLine(String.Format(" AND COLUMN_NAME = '{0}'", COLUMN_NAME))
        ASCMAIN1.sql = sql.ToString
        CTLLENGTH = Val(ASCDATA1.GetDataValue)
        RecsRemaining = MAXREC - MINREC
        If RecsRemaining = 0 Then
            MSG = String.Format("You Have No More {0} Remaining.", VERBTYPE)
            MSG = MSG & vbCrLf & "Please Fetch More Using The Button Available In The Transfer Screen"
            RetVal = "NONE"
        Else
            If RecsRemaining < 10 And RecsRemaining > 0 Then
                MSG = String.Format("You Only Have {0} {1} Remaining.", RecsRemaining, VERBTYPE)
                MSG = MSG & vbCrLf & "Please Fetch More Using The Button Available In The Transfer Screen"
            End If
            RetVal = Format$(MINREC, "".PadLeft(CTLLENGTH, "0"))

        End If
        Return RetVal
    End Function

    Public Shared Function SalesReportCanRun(ByVal START_DATE As Date,
                                      ByVal END_DATE As Date,
                                      ByVal CK_UPDATES As Boolean,
                                      ByVal CK_TARIFFS As Boolean) As String
        Dim RetVal As New System.Text.StringBuilder With {.Length = 0}
        Dim S As New System.Text.StringBuilder With {.Length = 0}
        Dim START_DATE_ORA As String = Format(START_DATE, "dd-MMM-yyyy")
        Dim END_DATE_ORA As String = Format(END_DATE, "dd-MMM-yyyy")
        If CK_UPDATES Then
            S.Length = 0
            S.AppendLine("SELECT DISTINCT INV_DATE")
            S.AppendLine("FROM SOTINVH1")
            S.AppendLine("WHERE ORDR_YYYYPP_UPDATED is Null")
            S.AppendLine(String.Format("AND INV_DATE >= '{0}'", START_DATE_ORA))
            S.AppendLine(String.Format("AND INV_DATE <= '{0}'", END_DATE_ORA))
            ASCMAIN1.sql = S.ToString()
            Dim INV_DATE As String = ASCDATA1.GetDataValue
            If IsDate(INV_DATE) Then
                RetVal.AppendLine("There Are Billed Orders That Need To Be Run Through Sales Journal.")
            End If
        End If
        If CK_TARIFFS Then
            S.Length = 0
            S.AppendLine("SELECT MAX(INIT_DATE) AS INIT_DATE")
            S.AppendLine("FROM ICTCOSTP")
            ASCMAIN1.sql = S.ToString()
            Dim INIT_DATE As String = ASCDATA1.GetDataValue
            If IsDate(INIT_DATE) Then
                If END_DATE > CDate(CDate(INIT_DATE).ToShortDateString) Then
                    RetVal.AppendLine(String.Format("Tariffs Costs Have Only Been Run Through {0}", Format(CDate(INIT_DATE), "MM/dd/yyyy")))
                End If
            End If
        End If
        Return RetVal.ToString
    End Function
End Class
