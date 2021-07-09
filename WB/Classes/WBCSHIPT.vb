
Imports System.Xml
Imports System.IO

Public Class WBCSHIPT
    Public ErrMsg As New Text.StringBuilder With {.Length = 0}
    Public FileNameCSV As String = "shipAddresses.csv"
    Public RemoteFolder As String = ""

    Private _DT As DataTable

#Region "Contructor"
    Public Sub New()
        createDataTable()
        ErrMsg.Length = 0
    End Sub

    Private Sub createDataTable()
        ErrMsg.Length = 0
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("W.EMAIL AS " & Chr(34) & "Email Address" & Chr(34) & ",")
        sql.AppendLine("A1.CUST_NAME AS " & Chr(34) & "Company Name" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_ADDR1 AS " & Chr(34) & "Shipping Address 1" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_ADDR2 AS " & Chr(34) & "Shipping Address 2" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_ADDR3 AS " & Chr(34) & "Shipping Address 3" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_CITY AS " & Chr(34) & "City" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_STATE AS " & Chr(34) & "State" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_ZIP_CODE AS " & Chr(34) & "Zip Code" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_COUNTRY AS " & Chr(34) & "Country" & Chr(34) & ",")
        sql.AppendLine("A2.CUST_ADDR_CODE AS " & Chr(34) & "Customer Address Code" & Chr(34))
        sql.AppendLine("FROM WBTCUST1 W, ARTCUST1 A1, ARTCUST2 A2")
        sql.AppendLine("WHERE W.CUST_CODE_ACTUAL = A1.CUST_CODE")
        sql.AppendLine("AND A1.CUST_CODE = A2.CUST_CODE")
        sql.AppendLine("AND NVL(W.CUST_CODE_ACTUAL,'NULL') <> 'NULL'")
        sql.AppendLine("AND NVL(CUST_ADDR_STATUS,'A') = 'A'")
        _DT = ASCDATA1.GetDataTable(sql.ToString(), String.Empty)
    End Sub

#End Region

#Region "Custom Methods"
    Public Function MakeFile(ByVal FolderName As String) As String
        If Not FolderName.EndsWith("\") Then
            FolderName = FolderName & "\"
        End If
        Dim RetVal As String = FolderName & FileNameCSV
        If IO.File.Exists(RetVal) Then
            IO.File.Delete(RetVal)
        End If

        Dim str As New Text.StringBuilder With {.Length = 0}

        If ASCMAIN1.Running_in_VS Then Stop

        ErrMsg.Length = 0

        For Each dc As DataColumn In _DT.Columns
            str.Append(Chr(34) & dc.ColumnName.ToString & Chr(34) & ",")
        Next
        str.Replace(",", vbNewLine, str.Length - 1, 1)

        For Each rowSHIPTO As DataRow In _DT.Rows
            For Each field As Object In rowSHIPTO.ItemArray
                str.Append(Chr(34) & field.ToString & Chr(34) & ",")
            Next
            str.Replace(",", vbNewLine, str.Length - 1, 1)
        Next

        Try
            My.Computer.FileSystem.WriteAllText(RetVal, str.ToString, False)
        Catch ex As Exception
            ErrMsg.AppendLine("Error Creating ShipTo File")
        End Try

        Return RetVal
    End Function
#End Region

End Class
