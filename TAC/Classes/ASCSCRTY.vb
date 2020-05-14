Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

''' <summary>
''' Security Encryption / Decryption Class
''' </summary>
''' <remarks></remarks>
Public Class ASCSCRTY

    Private Const keyWord As String = "P@$$w0rd"
    Private aesException As Exception = New Exception

    Public Enum Encryption
        Encrypt
        Decrypt
    End Enum

    ''' <summary>
    ''' Returns the last execption which occurred when Encrypt and Decrypt are called
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetLastException() As Exception
        Get
            Return aesException
        End Get
    End Property

    ''' <summary>
    ''' Encrypts a string
    ''' </summary>
    ''' <param name="str2encrypt">String to Encrypt</param>
    ''' <returns>Encrypted value</returns>
    ''' <remarks></remarks>
    Public Function Encrypt_AES(ByVal str2encrypt As String) As String

        Try

            aesException = New Exception

            Dim hexEncryption As String = Nothing
            Dim encrypted() As Byte
            Dim IV() As Byte
            Dim key(keyWord.Length) As Byte
            Dim myRijndael As New RijndaelManaged()
            Dim iter = 0
            Dim toEncrypt() As Byte
            Dim textConverter As New ASCIIEncoding()

            If str2encrypt.Length = 0 Then
                Return String.Empty
            End If

            For Each var As Char In keyWord
                key(iter) = CByte(Asc(var))
                iter += 1
            Next

            IV = myRijndael.IV

            Dim encryptor As ICryptoTransform = myRijndael.CreateEncryptor(key, IV)

            Dim msEncrypt As New MemoryStream()
            Dim csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)

            toEncrypt = textConverter.GetBytes(str2encrypt)

            csEncrypt.Write(toEncrypt, 0, toEncrypt.Length)
            csEncrypt.FlushFinalBlock()
            encrypted = msEncrypt.ToArray()

            hexEncryption = ByteArrayToString(IV) & ByteArrayToString(encrypted)

            Return hexEncryption.ToUpper
            iter = 0

        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Decrypts a string
    ''' </summary>
    ''' <param name="Str2decrypt">String to decrypt</param>
    ''' <returns>Decrypted String</returns>
    ''' <remarks></remarks>
    Public Function Decrypt_AES(ByVal Str2decrypt As String) As String

        Try
            aesException = New Exception

            Dim myRijndael As New RijndaelManaged()
            Dim fromEncrypt() As Byte
            Dim encrypted(15) As Byte
            Dim convertedStr As String
            Dim textConverter As New ASCIIEncoding()
            Dim key(keyWord.Length) As Byte
            Dim IV(15) As Byte
            Dim iter = 0
            Dim counter As Decimal = -1
            Dim hexbit As String = Nothing
            Dim pswByte As Integer

            If Str2decrypt.Length = 0 Then
                Return String.Empty
            End If

            'Make key
            For Each var As Char In keyWord
                key(iter) = CByte(Asc(var))
                iter += 1
            Next

            iter = 0
            'Seperate IV from Encryption
            For Each letter As Char In Str2decrypt
                If counter < 15 Then
                    If iter = 1 Then
                        hexbit = hexbit & letter
                        pswByte = Convert.ToInt16(hexbit, 16)
                        counter += 0.5
                        IV(counter) = pswByte
                        hexbit = Nothing
                        iter = 0
                    Else
                        hexbit = hexbit & letter
                        iter += 1
                        counter += 0.5
                    End If
                Else
                    If iter = 1 Then
                        hexbit = hexbit & letter
                        pswByte = Convert.ToInt16(hexbit, 16)
                        counter += 0.5
                        ReDim Preserve encrypted(counter - 16)
                        encrypted(counter - 16) = pswByte
                        hexbit = Nothing
                        iter = 0
                    Else
                        hexbit = hexbit & letter
                        iter += 1
                        counter += 0.5
                    End If
                End If
            Next

            Dim decryptor As ICryptoTransform = myRijndael.CreateDecryptor(key, IV)

            Dim msdecrypt As New MemoryStream(encrypted)
            Dim csDecrypt As New CryptoStream(msdecrypt, decryptor, CryptoStreamMode.Read)

            fromEncrypt = New Byte(encrypted.Length) {}
            csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length)
            convertedStr = textConverter.GetString(fromEncrypt)

            ' Need this here sice result has lots of chr(0) at end of string.
            convertedStr = convertedStr.Replace(Chr(0), String.Empty)

            Return convertedStr
        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Converts a Byte Array to a String
    ''' </summary>
    ''' <param name="ba"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ByteArrayToString(ByVal ba As Byte()) As String
        Try
            Dim hex As New StringBuilder(ba.Length * 2)
            For Each b As Byte In ba
                hex.AppendFormat("{0:x2}", b)
            Next
            Return hex.ToString()

        Catch ex As Exception
            aesException = ex
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Encrypts / Decrypts strings using procedure ENC_DEC in Oracle.
    ''' </summary>
    ''' <param name="stringToProcess">String to process. If providing a datatable and it is not ARTCUSTC, ARTCCPA1, ARTCCPDA then it should
    ''' be a comma seperated list of column names that appear in the datatable. If the column name does not exist in the table then it is skipped.</param>
    ''' <param name="procType">Set to Encrypt or Decrypt</param>
    ''' <param name="encrytpionKey">Key used to perform the Encryption/Decryption</param>
    ''' <param name="dTable">Datatable to process. It the datatable is ARTCUSTC, ARTCCPA1, ARTCCPDA make sure the tablename property of the datatable is set</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EncryptDecrypt(ByVal stringToProcess As String, _
                                    ByVal procType As Encryption, _
                                    ByVal encrytpionKey As String, _
                                    ByRef dTable As DataTable) As String

        Dim temp As String = String.Empty

        If encrytpionKey.Length = 0 Then
            Return temp
        End If

        If 1 = 1 Then
            Return temp
        End If

        If dTable IsNot Nothing Then
            Select Case dTable.TableName
                Case "ARTCUSTC"
                    For Each rowARTCUSTC As DataRow In dTable.Rows
                        For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}
                            If procType = Encryption.Encrypt Then
                                rowARTCUSTC.Item(field & "_E") = EncryptDecrypt(rowARTCUSTC.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                                rowARTCUSTC.Item(field) = DBNull.Value
                            Else
                                rowARTCUSTC.Item(field) = EncryptDecrypt(rowARTCUSTC.Item(field & "_E") & String.Empty, procType, encrytpionKey, Nothing)
                                rowARTCUSTC.Item(field & "_E") = DBNull.Value
                            End If
                        Next
                    Next

                Case "ARTCCPA1"
                    For Each rowARTCCPA1 As DataRow In dTable.Rows
                        For Each field As String In New String() {"CUST_CREDIT_CARD_NO", "CUST_CREDIT_CARD_EXP_DATE", "CUST_CREDIT_CARD_VER_CODE"}

                            If procType = Encryption.Encrypt Then
                                rowARTCCPA1.Item(field & "_E") = EncryptDecrypt(rowARTCCPA1.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                                rowARTCCPA1.Item(field) = DBNull.Value
                            Else
                                rowARTCCPA1.Item(field) = EncryptDecrypt(rowARTCCPA1.Item(field & "_E") & String.Empty, procType, encrytpionKey, Nothing)
                                rowARTCCPA1.Item(field & "_E") = DBNull.Value
                            End If
                        Next
                    Next

                Case "ARTCCPDA"
                    For Each rowARTCCPDA As DataRow In dTable.Rows
                        For Each field As String In New String() {"DETAIL_AGGREGATE"}

                            If procType = Encryption.Encrypt Then
                                rowARTCCPDA.Item(field) = EncryptDecrypt(rowARTCCPDA.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                            Else
                                rowARTCCPDA.Item(field) = EncryptDecrypt(rowARTCCPDA.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                            End If
                        Next
                    Next

                Case Else
                    For Each rowData As DataRow In dTable.Rows
                        For Each field As String In stringToProcess.Split(",")
                            field = field.Trim
                            If Not dTable.Columns.Contains(field) Then
                                Continue For
                            End If

                            If procType = Encryption.Encrypt Then
                                rowData.Item(field) = EncryptDecrypt(rowData.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                                rowData.Item(field) = DBNull.Value
                            Else
                                rowData.Item(field) = EncryptDecrypt(rowData.Item(field) & String.Empty, procType, encrytpionKey, Nothing)
                            End If
                        Next
                    Next


            End Select

        Else
            Select Case procType
                Case Encryption.Encrypt
                    ASCMAIN1.sql = "Select ENC_DEC.ENCRYPT('" & stringToProcess & "', '" & encrytpionKey & "' ) from DUAL"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    Dim b() As Byte = row.Item(0)
                    temp = BitConverter.ToString(b).Replace("-", String.Empty)

                Case Else
                    ASCMAIN1.sql = "Select ENC_DEC.DECRYPT('" & stringToProcess & "', '" & encrytpionKey & "' ) from DUAL"
                    Dim row As DataRow = ASCDATA1.GetDataRow
                    temp = row.Item(0) & String.Empty
            End Select

        End If

        Return temp

    End Function

End Class
