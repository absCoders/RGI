Imports nsoftware.IPWorksEncrypt

Public Class ASCENCRY
    Implements IDisposable

    ' Requires a reference to nsoftware.IPWorksEncrypt.dll

#Region "Class Variables"

    Private encDecType As EncrytpionTypes = EncrytpionTypes.AdvancedEncryptionStandard_AES
    Private iPaddingMode As Int16 = nsoftware.IPWorksEncrypt.EzcryptPaddingModes.pmNone
    Private iCipherMode As Int16 = nsoftware.IPWorksEncrypt.EzcryptCipherModes.cmCFB
    Private sKey As String = String.Empty
    Private sIteratorVector As String = String.Empty

    Private slastError As String = String.Empty

    Private sInputFile As String = String.Empty
    Private sOutputFile As String = String.Empty

    Private sDecryptedString As String = String.Empty
    Private sEncryptedString As String = String.Empty

    Private Aes As nsoftware.IPWorksEncrypt.Aes
    Private Blowfish As nsoftware.IPWorksEncrypt.Blowfish
    Private Cast As nsoftware.IPWorksEncrypt.Cast
    Private Des As nsoftware.IPWorksEncrypt.Des
    Private Idea As nsoftware.IPWorksEncrypt.Idea
    Private Rc2 As nsoftware.IPWorksEncrypt.Rc2
    Private Rc4 As nsoftware.IPWorksEncrypt.Rc4
    Private Tripledes As nsoftware.IPWorksEncrypt.Tripledes
    Private Twofish As nsoftware.IPWorksEncrypt.Twofish

    Private Ezcrypt1 As New nsoftware.IPWorksEncrypt.Ezcrypt
    Private ezcrypt1RuntimeLicense As String = "31454E394141315355425241533154453345383933333331580000000000000000000000000000004D554637445A525A0000465A4E454647504E433944460000"

    Public UseEncryption As Boolean = False
    Private Const defaultKey As String = "0fficeABS"

    Public Enum EncrytpionTypes
        AdvancedEncryptionStandard_AES
        Cast
        DataEncryptionStandard_DES
        InternationalDataEncryptionAlgorithm_IDEA
        RC2
        RC4
        TripleDES
        BlowFish
        TwoFish
    End Enum

    Public Enum PaddingTypes
        PKCS7 = 0
        Zeros = 1
        None = 2
        ANSIX23 = 3
        ISO10126 = 4
    End Enum

    Public Enum CipherTypes
        CBC = 0
        ECB = 1
        OFB = 2
        CFB = 3
    End Enum

#End Region

#Region "Instantiate Class"

    Public Sub New()
        InitializeVariables()
    End Sub

    Public Sub New(ByVal EncrytpionType As EncrytpionTypes)
        InitializeVariables()
        sKey = String.Empty
        encDecType = EncrytpionType
        sIteratorVector = String.Empty
    End Sub

    Public Sub New(ByVal EncrytpionType As EncrytpionTypes, ByVal Password As String, ByVal IteratorVector As String)
        InitializeVariables()
        sKey = Password
        encDecType = EncrytpionType
        sIteratorVector = IteratorVector
    End Sub

    Private Sub InitializeVariables()

        Aes = New nsoftware.IPWorksEncrypt.Aes
        Aes.RuntimeLicense = ezcrypt1RuntimeLicense

        Blowfish = New nsoftware.IPWorksEncrypt.Blowfish
        Blowfish.RuntimeLicense = ezcrypt1RuntimeLicense

        Cast = New nsoftware.IPWorksEncrypt.Cast
        Cast.RuntimeLicense = ezcrypt1RuntimeLicense

        Des = New nsoftware.IPWorksEncrypt.Des
        Des.RuntimeLicense = ezcrypt1RuntimeLicense

        Idea = New nsoftware.IPWorksEncrypt.Idea
        Idea.RuntimeLicense = ezcrypt1RuntimeLicense

        Rc2 = New nsoftware.IPWorksEncrypt.Rc2
        Rc2.RuntimeLicense = ezcrypt1RuntimeLicense

        Rc4 = New nsoftware.IPWorksEncrypt.Rc4
        Rc4.RuntimeLicense = ezcrypt1RuntimeLicense

        Tripledes = New nsoftware.IPWorksEncrypt.Tripledes
        Tripledes.RuntimeLicense = ezcrypt1RuntimeLicense

        Twofish = New nsoftware.IPWorksEncrypt.Twofish
        Twofish.RuntimeLicense = ezcrypt1RuntimeLicense

        DefaultEncryption()
    End Sub

    Private Sub DefaultEncryption()
        encDecType = EncrytpionTypes.AdvancedEncryptionStandard_AES
        sKey = defaultKey
        iPaddingMode = nsoftware.IPWorksEncrypt.EzcryptPaddingModes.pmNone
        iCipherMode = nsoftware.IPWorksEncrypt.EzcryptCipherModes.cmCFB
        sIteratorVector = String.Empty
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' The cipher mode of operation.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CipherMode As Int16
        Get
            Return iCipherMode
        End Get
        Set(value As Int16)
            iCipherMode = value
        End Set
    End Property

    ''' <summary>
    ''' The padding mode of operation.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PaddingMode As Int16
        Get
            Return iPaddingMode
        End Get
        Set(value As Int16)
            iPaddingMode = value
        End Set
    End Property

    ''' <summary>
    ''' Get set the password used for the ecnryption / decryption
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Key() As String
        Get
            Return sKey
        End Get
        Set(value As String)
            sKey = value
        End Set
    End Property

    ''' <summary>
    ''' Get set the password used for the Iteration Vector ecnryption / decryption
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IteratorVector() As String
        Get
            Return sIteratorVector
        End Get
        Set(value As String)
            sIteratorVector = value
        End Set
    End Property


    ''' <summary>
    ''' Get the Error generated by the last procedure / function call.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Last error generated by function / procedure call</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LastError As String
        Get
            Return slastError
        End Get
    End Property

    ''' <summary>
    ''' Get / Set Input File
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property InputFile As String
        Get
            Return sInputFile
        End Get
        Set(value As String)
            sInputFile = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set Output File
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OutputFile As String
        Get
            Return sOutputFile
        End Get
        Set(value As String)
            sOutputFile = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set String to be / get Decrypted
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DecryptedString() As String
        Get
            Return sDecryptedString
        End Get
        Set(value As String)
            sDecryptedString = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set String to be / get Encrypted
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EncryptedString() As String
        Get
            Return sEncryptedString
        End Get
        Set(value As String)
            sEncryptedString = value
        End Set
    End Property

#End Region

#Region "Encryption / Decryption"

    ''' <summary>
    ''' Decrypt file
    ''' </summary>
    ''' <returns>True is successful; otherwise false</returns>
    ''' <remarks></remarks>
    Public Function DecryptFile(ByVal OverwriteOutputFile As Boolean) As Boolean

        slastError = String.Empty
        If Not UseEncryption Then
            Return True
        End If

        Try
            Select Case encDecType
                Case EncrytpionTypes.AdvancedEncryptionStandard_AES
                    Aes.Overwrite = OverwriteOutputFile
                    Aes.InputFile = sInputFile
                    Aes.OutputFile = sOutputFile
                    Aes.Key = sKey
                    Aes.IV = sIteratorVector
                    Aes.PaddingMode = iPaddingMode
                    Aes.CipherMode = iCipherMode
                    Aes.Decrypt()
                    Aes.Dispose()

                Case EncrytpionTypes.Cast
                    Cast.Overwrite = OverwriteOutputFile
                    Cast.InputFile = sInputFile
                    Cast.OutputFile = sOutputFile
                    Cast.Key = sKey
                    Cast.IV = sIteratorVector
                    Cast.PaddingMode = iPaddingMode
                    Cast.CipherMode = iCipherMode
                    Cast.Decrypt()
                    Cast.Dispose()

                Case EncrytpionTypes.DataEncryptionStandard_DES
                    Des.Overwrite = OverwriteOutputFile
                    Des.InputFile = sInputFile
                    Des.OutputFile = sOutputFile
                    Des.Key = sKey
                    Des.IV = sIteratorVector
                    Des.PaddingMode = iPaddingMode
                    Des.CipherMode = iCipherMode
                    Des.Decrypt()
                    Des.Dispose()

                Case EncrytpionTypes.InternationalDataEncryptionAlgorithm_IDEA
                    Idea.Overwrite = OverwriteOutputFile
                    Idea.InputFile = sInputFile
                    Idea.OutputFile = sOutputFile
                    Idea.Key = sKey
                    Idea.IV = sIteratorVector
                    Idea.PaddingMode = iPaddingMode
                    Idea.CipherMode = iCipherMode
                    Idea.Decrypt()
                    Idea.Dispose()

                Case EncrytpionTypes.RC2
                    Rc2.Overwrite = OverwriteOutputFile
                    Rc2.InputFile = sInputFile
                    Rc2.OutputFile = sOutputFile
                    Rc2.Key = sKey
                    Rc2.IV = sIteratorVector
                    Rc2.PaddingMode = iPaddingMode
                    Rc2.CipherMode = iCipherMode
                    Rc2.Decrypt()
                    Rc2.Dispose()

                Case EncrytpionTypes.RC4
                    Rc4.Overwrite = OverwriteOutputFile
                    Rc4.InputFile = sInputFile
                    Rc4.OutputFile = sOutputFile
                    Rc4.Key = sKey
                    Rc4.IV = sIteratorVector
                    'Rc4.PaddingMode = iPaddingMode
                    'Rc4.CipherMode = iCipherMode
                    Rc4.Decrypt()
                    Rc4.Dispose()

                Case EncrytpionTypes.TripleDES
                    Tripledes.Overwrite = OverwriteOutputFile
                    Tripledes.InputFile = sInputFile
                    Tripledes.OutputFile = sOutputFile
                    Tripledes.Key = sKey
                    Tripledes.IV = sIteratorVector
                    Tripledes.PaddingMode = iPaddingMode
                    Tripledes.CipherMode = iCipherMode
                    Tripledes.Decrypt()
                    Tripledes.Dispose()

                Case EncrytpionTypes.BlowFish
                    Blowfish.Overwrite = OverwriteOutputFile
                    Blowfish.InputFile = sInputFile
                    Blowfish.OutputFile = sOutputFile
                    Blowfish.Key = sKey
                    Blowfish.IV = sIteratorVector
                    Blowfish.PaddingMode = iPaddingMode
                    Blowfish.CipherMode = iCipherMode
                    Blowfish.Decrypt()
                    Blowfish.Dispose()

                Case EncrytpionTypes.TwoFish
                    Twofish.Overwrite = OverwriteOutputFile
                    Twofish.InputFile = sInputFile
                    Twofish.OutputFile = sOutputFile
                    Twofish.Key = sKey
                    Twofish.IV = sIteratorVector
                    Twofish.PaddingMode = iPaddingMode
                    Twofish.CipherMode = iCipherMode
                    Twofish.Decrypt()
                    Twofish.Dispose()

                Case Else
                    slastError = "Invalid selection."
                    Return False
            End Select

            Return True
        Catch ex As Exception
            slastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Encrypt file 
    ''' </summary>
    ''' <returns>True is successful; otherwise false</returns>
    ''' <remarks></remarks>
    Public Function EncryptFile(ByVal OverwriteOutputFile As Boolean) As Boolean

        slastError = String.Empty
        If Not UseEncryption Then
            Return True
        End If

        Try
            Select Case encDecType
                Case EncrytpionTypes.AdvancedEncryptionStandard_AES
                    Aes.Overwrite = OverwriteOutputFile
                    Aes.InputFile = sInputFile
                    Aes.OutputFile = sOutputFile
                    Aes.Key = sKey
                    Aes.IV = sIteratorVector
                    Aes.PaddingMode = iPaddingMode
                    Aes.CipherMode = iCipherMode
                    Aes.Encrypt()
                    Aes.Dispose()

                Case EncrytpionTypes.Cast
                    Cast.Overwrite = OverwriteOutputFile
                    Cast.InputFile = sInputFile
                    Cast.OutputFile = sOutputFile
                    Cast.Key = sKey
                    Cast.IV = sIteratorVector
                    Cast.PaddingMode = iPaddingMode
                    Cast.CipherMode = iCipherMode
                    Cast.Encrypt()
                    Cast.Dispose()

                Case EncrytpionTypes.DataEncryptionStandard_DES
                    Des.Overwrite = OverwriteOutputFile
                    Des.InputFile = sInputFile
                    Des.OutputFile = sOutputFile
                    Des.Key = sKey
                    Des.IV = sIteratorVector
                    Des.PaddingMode = iPaddingMode
                    Des.CipherMode = iCipherMode
                    Des.Encrypt()
                    Des.Dispose()

                Case EncrytpionTypes.InternationalDataEncryptionAlgorithm_IDEA
                    Idea.Overwrite = OverwriteOutputFile
                    Idea.InputFile = sInputFile
                    Idea.OutputFile = sOutputFile
                    Idea.Key = sKey
                    Idea.IV = sIteratorVector
                    Idea.PaddingMode = iPaddingMode
                    Idea.CipherMode = iCipherMode
                    Idea.Encrypt()
                    Idea.Dispose()

                Case EncrytpionTypes.RC2
                    Rc2.Overwrite = OverwriteOutputFile
                    Rc2.InputFile = sInputFile
                    Rc2.OutputFile = sOutputFile
                    Rc2.Key = sKey
                    Rc2.IV = sIteratorVector
                    Rc2.PaddingMode = iPaddingMode
                    Rc2.CipherMode = iCipherMode
                    Rc2.Encrypt()
                    Rc2.Dispose()

                Case EncrytpionTypes.RC2
                    Rc4.Overwrite = OverwriteOutputFile
                    Rc4.InputFile = sInputFile
                    Rc4.OutputFile = sOutputFile
                    Rc4.Key = sKey
                    Rc4.IV = sIteratorVector
                    'Rc4.PaddingMode = iPaddingMode
                    'Rc4.CipherMode = iCipherMode
                    Rc4.Encrypt()
                    Rc4.Dispose()

                Case EncrytpionTypes.TripleDES
                    Tripledes.Overwrite = OverwriteOutputFile
                    Tripledes.InputFile = sInputFile
                    Tripledes.OutputFile = sOutputFile
                    Tripledes.Key = sKey
                    Tripledes.IV = sIteratorVector
                    Tripledes.PaddingMode = iPaddingMode
                    Tripledes.CipherMode = iCipherMode
                    Tripledes.Encrypt()
                    Tripledes.Dispose()

                Case EncrytpionTypes.BlowFish
                    Blowfish.Overwrite = OverwriteOutputFile
                    Blowfish.InputFile = sInputFile
                    Blowfish.OutputFile = sOutputFile
                    Blowfish.Key = sKey
                    Blowfish.IV = sIteratorVector
                    Blowfish.PaddingMode = iPaddingMode
                    Blowfish.CipherMode = iCipherMode
                    Blowfish.Encrypt()
                    Blowfish.Dispose()

                Case EncrytpionTypes.TwoFish
                    Twofish.Overwrite = OverwriteOutputFile
                    Twofish.InputFile = sInputFile
                    Twofish.OutputFile = sOutputFile
                    Twofish.Key = sKey
                    Twofish.IV = sIteratorVector
                    Twofish.PaddingMode = iPaddingMode
                    Twofish.CipherMode = iCipherMode
                    Twofish.Encrypt()
                    Twofish.Dispose()

                Case Else
                    slastError = "Invalid selection."
                    Return False
            End Select
            Return True
        Catch ex As Exception
            slastError = ex.Message
            Return False
        End Try
    End Function

    Public Function DecryptString(ByVal strToBeDecrypted As String) As String

        Return ASCMAIN1.DecryptAES(strToBeDecrypted)

        'slastError = String.Empty
        'If Not UseEncryption Then
        '    Return strToBeDecrypted
        'End If

        'If strToBeDecrypted.Length = 0 Then
        '    Return String.Empty
        'End If

        'Ezcrypt1.RuntimeLicense = ezcrypt1RuntimeLicense
        'Ezcrypt1.Reset()
        'Ezcrypt1.RuntimeLicense = ezcrypt1RuntimeLicense
        'selectAlgorithm()
        'If LastError.Length > 0 Then
        '    Return String.Empty
        'End If

        'Ezcrypt1.UseHex = True
        'Ezcrypt1.InputMessage = strToBeDecrypted
        'Ezcrypt1.KeyPassword = Key
        'Ezcrypt1.Decrypt()
        'Return Ezcrypt1.OutputMessage
    End Function

    Public Function EncryptString(ByVal strToBeEncrypted As String) As String

        Return ASCMAIN1.EncryptAES(strToBeEncrypted)

        'slastError = String.Empty
        'If Not UseEncryption Then
        '    Return strToBeEncrypted
        'End If

        'If strToBeEncrypted.Length = 0 Then
        '    Return String.Empty
        'End If

        'Ezcrypt1.RuntimeLicense = ezcrypt1RuntimeLicense
        'Ezcrypt1.Reset()
        'Ezcrypt1.RuntimeLicense = ezcrypt1RuntimeLicense
        'selectAlgorithm()
        'If LastError.Length > 0 Then
        '    Return String.Empty
        'End If

        'Ezcrypt1.UseHex = True
        'Ezcrypt1.InputMessage = strToBeEncrypted
        'Ezcrypt1.KeyPassword = Key
        'Ezcrypt1.Encrypt()
        'Return Ezcrypt1.OutputMessage
    End Function

    Private Sub selectAlgorithm()
        Select Case encDecType
            Case EncrytpionTypes.AdvancedEncryptionStandard_AES
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezAES
            Case EncrytpionTypes.Cast
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezCAST
            Case EncrytpionTypes.DataEncryptionStandard_DES
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezDES
             Case EncrytpionTypes.InternationalDataEncryptionAlgorithm_IDEA
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezIDEA
            Case EncrytpionTypes.RC2
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezRC2
            Case EncrytpionTypes.RC4
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezRC4
            Case EncrytpionTypes.TripleDES
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezTripleDES
            Case EncrytpionTypes.BlowFish
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezBlowfish
            Case EncrytpionTypes.TwoFish
                Ezcrypt1.Algorithm = EzcryptAlgorithms.ezTwofish
            Case Else
                slastError = "Invalid selection."
                Exit Select
        End Select
    End Sub

    ''' <summary>
    ''' Decrypt string 
    ''' </summary>
    ''' <returns>True is successful; otherwise false</returns>
    ''' <remarks></remarks>
    Private Function DecryptString_old() As Boolean

        Try
            slastError = String.Empty
            sDecryptedString = String.Empty

            Select Case encDecType
                Case EncrytpionTypes.AdvancedEncryptionStandard_AES
                    Aes.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Aes.Key = sKey
                    Aes.IV = sIteratorVector
                    Aes.PaddingMode = iPaddingMode
                    Aes.CipherMode = iCipherMode
                    Aes.Decrypt()
                    sDecryptedString = Aes.OutputMessage
                    Aes.Dispose()

                Case EncrytpionTypes.Cast
                    Cast.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Cast.Key = sKey
                    Cast.IV = sIteratorVector
                    Cast.PaddingMode = iPaddingMode
                    Cast.CipherMode = iCipherMode
                    Cast.Decrypt()
                    sDecryptedString = Cast.OutputMessage
                    Cast.Dispose()

                Case EncrytpionTypes.DataEncryptionStandard_DES
                    Des.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Des.Key = sKey
                    Des.IV = sIteratorVector
                    Des.PaddingMode = iPaddingMode
                    Des.CipherMode = iCipherMode
                    Des.Decrypt()
                    sDecryptedString = Des.OutputMessage
                    Des.Dispose()

                Case EncrytpionTypes.InternationalDataEncryptionAlgorithm_IDEA
                    Idea.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Idea.Key = sKey
                    Idea.IV = sIteratorVector
                    Idea.PaddingMode = iPaddingMode
                    Idea.CipherMode = iCipherMode
                    Idea.Decrypt()
                    sDecryptedString = Idea.OutputMessage
                    Idea.Dispose()

                Case EncrytpionTypes.RC2
                    Rc2.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Rc2.Key = sKey
                    Rc2.IV = sIteratorVector
                    Rc2.PaddingMode = iPaddingMode
                    Rc2.CipherMode = iCipherMode
                    Rc2.Decrypt()
                    sDecryptedString = Rc2.OutputMessage
                    Rc2.Dispose()

                Case EncrytpionTypes.RC4
                    Rc4.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Rc4.Key = sKey
                    Rc4.IV = sIteratorVector
                    'Rc4.PaddingMode = iPaddingMode
                    'Rc4.CipherMode = iCipherMode
                    Rc4.Decrypt()
                    sDecryptedString = Rc4.OutputMessage
                    Rc4.Dispose()

                Case EncrytpionTypes.TripleDES
                    Tripledes.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Tripledes.Key = sKey
                    Tripledes.IV = sIteratorVector
                    Tripledes.PaddingMode = iPaddingMode
                    Tripledes.CipherMode = iCipherMode
                    Tripledes.Decrypt()
                    sDecryptedString = Tripledes.OutputMessage
                    Tripledes.Dispose()

                Case EncrytpionTypes.BlowFish
                    Blowfish.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Blowfish.Key = sKey
                    Blowfish.IV = sIteratorVector
                    Blowfish.PaddingMode = iPaddingMode
                    Blowfish.CipherMode = iCipherMode
                    Blowfish.Decrypt()
                    sDecryptedString = Blowfish.OutputMessage
                    Blowfish.Dispose()

                Case EncrytpionTypes.TwoFish
                    Twofish.InputMessageB = Convert.FromBase64String(sEncryptedString)
                    Twofish.Key = sKey
                    Twofish.IV = sIteratorVector
                    Twofish.PaddingMode = iPaddingMode
                    Twofish.CipherMode = iCipherMode
                    Twofish.Decrypt()
                    sDecryptedString = Twofish.OutputMessage
                    Twofish.Dispose()

                Case Else
                    slastError = "Invalid selection."
                    Return False
            End Select

            Return True
        Catch ex As Exception
            slastError = ex.Message
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Encrypt string 
    ''' </summary>
    ''' <returns>True is successful; otherwise false</returns>
    ''' <remarks></remarks>
    Private Function EncryptString_Old() As Boolean

        Try
            slastError = String.Empty
            sEncryptedString = String.Empty

            Select Case encDecType
                Case EncrytpionTypes.AdvancedEncryptionStandard_AES
                    Aes.InputMessage = sDecryptedString
                    Aes.Key = sKey
                    Aes.IV = sIteratorVector
                    Aes.PaddingMode = iPaddingMode
                    Aes.CipherMode = iCipherMode
                    Aes.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Aes.OutputMessageB)
                    Aes.Dispose()

                Case EncrytpionTypes.Cast
                    Cast.InputMessage = sDecryptedString
                    Cast.Key = sKey
                    Cast.IV = sIteratorVector
                    Cast.PaddingMode = iPaddingMode
                    Cast.CipherMode = iCipherMode
                    Cast.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Cast.OutputMessageB)
                    Cast.Dispose()

                Case EncrytpionTypes.DataEncryptionStandard_DES
                    Des.InputMessage = sDecryptedString
                    Des.Key = sKey
                    Des.IV = sIteratorVector
                    Des.PaddingMode = iPaddingMode
                    Des.CipherMode = iCipherMode
                    Des.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Des.OutputMessageB)
                    Des.Dispose()

                Case EncrytpionTypes.InternationalDataEncryptionAlgorithm_IDEA
                    Idea.InputMessage = sDecryptedString
                    Idea.Key = sKey
                    Idea.IV = sIteratorVector
                    Idea.PaddingMode = iPaddingMode
                    Idea.CipherMode = iCipherMode
                    Idea.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Idea.OutputMessageB)
                    Idea.Dispose()

                Case EncrytpionTypes.RC2
                    Rc2.InputMessage = sDecryptedString
                    Rc2.Key = sKey
                    Rc2.IV = sIteratorVector
                    Rc2.PaddingMode = iPaddingMode
                    Rc2.CipherMode = iCipherMode
                    Rc2.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Rc2.OutputMessageB)
                    Rc2.Dispose()

                Case EncrytpionTypes.RC4
                    Rc4.InputMessage = sDecryptedString
                    Rc4.Key = sKey
                    Rc4.IV = sIteratorVector
                    'Rc4.PaddingMode = iPaddingMode
                    'Rc4.CipherMode = iCipherMode
                    Rc4.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Rc4.OutputMessageB)
                    Rc4.Dispose()

                Case EncrytpionTypes.TripleDES
                    Tripledes.InputMessage = sDecryptedString
                    Tripledes.Key = sKey
                    Tripledes.IV = sIteratorVector
                    Tripledes.PaddingMode = iPaddingMode
                    Tripledes.CipherMode = iCipherMode
                    Tripledes.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Tripledes.OutputMessageB)
                    Tripledes.Dispose()

                Case EncrytpionTypes.BlowFish
                    Blowfish.InputMessage = sDecryptedString
                    Blowfish.Key = sKey
                    Blowfish.IV = sIteratorVector
                    Blowfish.PaddingMode = iPaddingMode
                    Blowfish.CipherMode = iCipherMode
                    Blowfish.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Blowfish.OutputMessageB)
                    Blowfish.Dispose()

                Case EncrytpionTypes.TwoFish
                    Twofish.InputMessage = sDecryptedString
                    Twofish.Key = sKey
                    Twofish.IV = sIteratorVector
                    Twofish.PaddingMode = iPaddingMode
                    Twofish.CipherMode = iCipherMode
                    Twofish.Encrypt()
                    sEncryptedString = Convert.ToBase64String(Twofish.OutputMessageB)
                    Twofish.Dispose()

                Case Else
                    slastError = "Invalid selection."
                    Return False
            End Select

            Return True
        Catch ex As Exception
            slastError = ex.Message
            Return False
        End Try

    End Function

#End Region

#Region "IDisposable Support"

    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

#End Region

End Class
