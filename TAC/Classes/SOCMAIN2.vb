Imports System.IO
Public Class SOCMAIN2
    Public Shared Function Price_Discounts(frm As ASFBASE0,
           CUST_CODE As String, rowARTCUST1 As DataRow,
           STYLE_CODE As String,
           Optional UseCustomer As Boolean = True,
           Optional UseDiscounts As Boolean = True,
           Optional SupressMsg As Boolean = False,
           Optional STYLE_PRICE_NEW As Double = 0,
           Optional UseLocalDST As Boolean = False,
           Optional STYLE_COLOR_STATUS As String = "") As List(Of DISCOUNTS)

        'If Not frm.dst.Tables.Contains("ICTCLAS1") Then
        '    ASCMAIN1.sql = "Select * from ICTCLAS1"
        '    frm.Create_TDA(frm.dst.Tables.Add, "ICTCLAS1", "**", 0, False)
        '    frm.Fill_Records("ICTCLAS1")
        'End If
        'If Not frm.dst.Tables.Contains("ICTDISC1") Then
        '    ASCMAIN1.sql = "Select * from ICTDISC1"
        '    frm.Create_TDA(frm.dst.Tables.Add, "ICTDISC1", "**", 0, False)
        '    frm.Fill_Records("ICTDISC1")
        'End If

        Dim retval As New List(Of DISCOUNTS)
        Dim DiscPromoFound As Boolean = False
        Const DiscPromoPct As Double = 70
        Dim DiscPromoDesc As String = ""
        Dim rowICTSTYL1 As DataRow
        If UseLocalDST Then
            rowICTSTYL1 = frm.dst.Tables.Item("ICTTFLST").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE)).FirstOrDefault
        Else
            rowICTSTYL1 = frm.LookUp("ICTSTYL1", New String() {STYLE_CODE})
        End If
        Dim STYLE_STATUS As String = rowICTSTYL1.Item("STYLE_STATUS") & ""
        Dim STYLE_CLASS_CODE As String = rowICTSTYL1.Item("STYLE_CLASS_CODE") & ""
        If STYLE_CLASS_CODE = "" Then 'We have to protect against this somehow.
            STYLE_CLASS_CODE = "PVC"
        End If
        Dim STYLE_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PRICE") & "")
        If STYLE_PRICE_NEW <> 0 Then
            STYLE_PRICE = STYLE_PRICE_NEW
        End If

        Dim STYLE_PROMO_PRICE As Decimal = Val(rowICTSTYL1.Item("STYLE_PROMO_PRICE") & "")
        '----- Begin New Promo System ------
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT")
        sql.AppendLine("P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE,")
        sql.AppendLine("MAX(P2.PROMO_UNIT_PRICE) PROMO_UNIT_PRICE")
        sql.AppendLine("FROM ICTPROM1 P1, ICTPROM2 P2")
        sql.AppendLine("WHERE P1.PROMO_CTL_NO = P2.PROMO_CTL_NO")
        sql.AppendLine("AND P2.STYLE_CODE = :PARM1")
        sql.AppendLine("GROUP BY P1.PROMO_START_DATE,")
        sql.AppendLine("P1.PROMO_END_DATE")
        Dim tblICTPROMX As DataTable = ASCDATA1.GetDataTable(sql.ToString(), String.Empty, "V", STYLE_CODE)
        For Each rowICTPROMX As DataRow In tblICTPROMX.Select("", "PROMO_START_DATE")
            Dim PROMO_START_DATE As DateTime = CDate(rowICTPROMX.Item("PROMO_START_DATE").ToString & String.Empty)
            Dim PROMO_END_DATE As DateTime = CDate(rowICTPROMX.Item("PROMO_END_DATE").ToString & String.Empty)
            If PROMO_START_DATE <= Now() And PROMO_END_DATE >= Now() Then
                STYLE_PROMO_PRICE = Val(rowICTPROMX.Item("PROMO_UNIT_PRICE").ToString & String.Empty)
            End If
        Next
        '----- End New Promo System -------

        Dim CARTON_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("CARTON_PACK_QTY") & "")
        Dim INNER_PACK_QTY As Int32 = Val(rowICTSTYL1.Item("INNER_PACK_QTY") & "")
        Dim MSOQ As Int64 = Val(rowICTSTYL1.Item("STYLE_SO_QTY_MIN") & "")
        Dim HALFCASE As Integer = 0
        'If INNER_PACK_QTY > 0 Then
        '    HALFCASE = Math.Ceiling((CARTON_PACK_QTY / 2) / INNER_PACK_QTY) * INNER_PACK_QTY
        'Else
        '    HALFCASE = (CARTON_PACK_QTY / 2)
        'End If
        If MSOQ > 0 Then
            HALFCASE = Math.Ceiling((CARTON_PACK_QTY / 2) / MSOQ) * MSOQ
        Else
            HALFCASE = (CARTON_PACK_QTY / 2)
        End If
        Dim rowICTCLAS1 As DataRow
        If UseLocalDST Then
            rowICTCLAS1 = frm.dst.Tables.Item("ICTCLAS1").Select(String.Format("STYLE_CLASS_CODE = '{0}'", STYLE_CLASS_CODE)).FirstOrDefault
        Else
            rowICTCLAS1 = frm.LookUp("ICTCLAS1", STYLE_CLASS_CODE)
        End If
        Dim IsPVC As Boolean = rowICTCLAS1.Item("DISC_CODE").ToString = "PVC"
        Dim CUST_PRICE_TIER As String = ""
        Dim CUST_DISC_PCT_EXTRA As Double = 1
        Dim CUST_DISC_PCT As Boolean = False
        Dim CUST_PRICE_TIER_PVC As String = ""

        If STYLE_COLOR_STATUS <> "" Then
            STYLE_STATUS = STYLE_COLOR_STATUS
        End If

        If UseDiscounts Then
            If STYLE_STATUS = "D" Then
                DiscPromoFound = True
                DiscPromoDesc = "Disc"
            ElseIf STYLE_PROMO_PRICE <> 0 Then
                DiscPromoFound = True
                DiscPromoDesc = "Promo"
            End If
        End If

        If UseCustomer Then
            If Not IsNothing(rowARTCUST1) Then
                If IsPVC Then
                    CUST_PRICE_TIER_PVC = rowARTCUST1.Item("CUST_PRICE_TIER_PVC") & ""
                Else
                    CUST_PRICE_TIER = rowARTCUST1.Item("CUST_PRICE_TIER") & ""
                    If Not IsPVC Then
                        Select Case rowARTCUST1.Item("CUST_DISC_PCT_EXTRA") & ""
                            Case "2"
                                CUST_DISC_PCT_EXTRA = 0.9
                            Case "1"
                                CUST_DISC_PCT_EXTRA = 0.95
                            Case Else
                                CUST_DISC_PCT_EXTRA = 1
                        End Select
                        'Customer Can Not Have An Extra Discount If they Are Set At A Specific Discount.
                        'Per Rich 6/2/14
                        If Val(rowARTCUST1.Item("CUST_DISC_PCT") & "") > 0 Then
                            CUST_DISC_PCT_EXTRA = 1
                        End If
                    End If
                    If Val(rowARTCUST1.Item("CUST_DISC_PCT") & "") <> 0 Then
                        CUST_DISC_PCT = Not IsPVC
                    Else
                        CUST_DISC_PCT = False
                    End If
                End If
            End If
        End If

        If CARTON_PACK_QTY = 0 And INNER_PACK_QTY = 0 Then
            If Not SupressMsg Then
                MsgBox("Box & Carton Qty Set To Zero", vbOKOnly, "Style Attributes Problem")
            End If
            CARTON_PACK_QTY = 1
            INNER_PACK_QTY = 1
        End If

        If rowICTCLAS1 IsNot Nothing Then
            Dim DISC_CODE As String = rowICTCLAS1.Item("DISC_CODE") & ""
            'Dim rowICTDISC1 As DataRow = frm.dst.Tables("ICTDISC1").Rows.Find(DISC_CODE)
            Dim rowICTDISC1 As DataRow
            If UseLocalDST Then
                rowICTDISC1 = frm.dst.Tables.Item("ICTDISC1").Select(String.Format("DISC_CODE = '{0}'", DISC_CODE)).FirstOrDefault
            Else
                rowICTDISC1 = frm.LookUp("ICTDISC1", DISC_CODE)
            End If
            Dim HAlfCaseIsOne As Boolean = False
            If rowICTDISC1 IsNot Nothing Then
                For I As Integer = 1 To 4
                    Dim DISC As New DISCOUNTS
                    Dim CASES As Decimal = Val(rowICTDISC1.Item(String.Format("DISC{0}_CASES", CStr(I))) & "")
                    Dim PCT As Decimal = Val(rowICTDISC1.Item(String.Format("DISC{0}_PCT", CStr(I))) & "")
                    If IsPVC Then
                        If I = 4 And HAlfCaseIsOne Then
                            DISC.DISCOUNT_QTY = 0
                        Else
                            If CASES = 0 And CARTON_PACK_QTY > 1 Then
                                If I = 4 And INNER_PACK_QTY > 0 Then
                                    DISC.DISCOUNT_QTY = INNER_PACK_QTY
                                    If DISC.DISCOUNT_QTY < MSOQ Then
                                        DISC.DISCOUNT_QTY = MSOQ
                                    End If
                                Else
                                    If CARTON_PACK_QTY <> MSOQ And CARTON_PACK_QTY > 1 Then
                                        DISC.DISCOUNT_QTY = 1
                                        If DISC.DISCOUNT_QTY < MSOQ Then
                                            DISC.DISCOUNT_QTY = MSOQ
                                        End If
                                    End If
                                End If
                            Else
                                DISC.DISCOUNT_QTY = CARTON_PACK_QTY * CASES
                            End If
                        End If
                    Else
                        If I = 3 And CARTON_PACK_QTY = 2 Then
                            HAlfCaseIsOne = True
                        End If
                        If CARTON_PACK_QTY = 1 And (I = 3 Or I = 4) Then
                            DISC.DISCOUNT_QTY = 0
                        Else
                            If I = 4 And HAlfCaseIsOne Then
                                DISC.DISCOUNT_QTY = 0
                            Else
                                If (CARTON_PACK_QTY * CASES) < 1 Then
                                    If INNER_PACK_QTY > 1 Then
                                        If I = 4 And (INNER_PACK_QTY = HALFCASE) Then
                                            DISC.DISCOUNT_QTY = 0
                                        Else
                                            If INNER_PACK_QTY <= MSOQ Then
                                                DISC.DISCOUNT_QTY = MSOQ
                                            Else
                                                DISC.DISCOUNT_QTY = INNER_PACK_QTY
                                            End If
                                        End If
                                    Else
                                        If retval.Count >= 3 Then
                                            If I = 4 And retval(2).DISCOUNT_QTY = 0 Then
                                                DISC.DISCOUNT_QTY = 0
                                            Else
                                                If I = 4 And HALFCASE = MSOQ Then
                                                    DISC.DISCOUNT_QTY = 0
                                                Else
                                                    DISC.DISCOUNT_QTY = 1
                                                    If DISC.DISCOUNT_QTY < MSOQ Then
                                                        DISC.DISCOUNT_QTY = MSOQ
                                                    End If
                                                End If
                                            End If
                                        Else
                                            DISC.DISCOUNT_QTY = 0
                                        End If
                                    End If
                                Else
                                    If I = 3 Then
                                        If MSOQ = CARTON_PACK_QTY Then
                                            DISC.DISCOUNT_QTY = 0
                                        Else
                                            DISC.DISCOUNT_QTY = HALFCASE
                                            If DISC.DISCOUNT_QTY < MSOQ Then
                                                DISC.DISCOUNT_QTY = 0
                                            End If
                                        End If
                                    Else
                                        DISC.DISCOUNT_QTY = CARTON_PACK_QTY * CASES
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If DiscPromoFound Then
                        If I = 1 Then
                            If STYLE_PROMO_PRICE <> 0 Then
                                DISC.DISCOUNT_PRICE = STYLE_PROMO_PRICE
                                DISC.DISCOUNT_PCT = String.Format("{0}", DiscPromoDesc)
                            Else
                                DISC.DISCOUNT_PRICE = STYLE_PRICE * (100 - DiscPromoPct) / 100
                                DISC.DISCOUNT_PCT = String.Format("{0}->{1}%", DiscPromoDesc, DiscPromoPct)
                            End If
                            DISC.DISCOUNT_DESC = DiscPromoDesc
                            DISC.DISCOUNT_QTY = 1
                        Else
                            DISC.DISCOUNT_QTY = 0
                        End If
                    Else
                        If CUST_DISC_PCT Then
                            Dim Calc_Price As Double = (STYLE_PRICE * (100 - PCT) / 100) * CUST_DISC_PCT_EXTRA
                            Dim Disc_Price As Double = (STYLE_PRICE * (100 - Val(rowARTCUST1.Item("CUST_DISC_PCT") & "")) / 100)
                            'If Calc_Price < Disc_Price Then
                            '    DISC.DISCOUNT_PRICE = Calc_Price
                            'Else
                            DISC.DISCOUNT_PRICE = Disc_Price
                            'End If
                            'DISC.DISCOUNT_PRICE = (STYLE_PRICE * (100 - Val(rowARTCUST1.Item("CUST_DISC_PCT") & "")) / 100)
                        Else
                            DISC.DISCOUNT_PRICE = (STYLE_PRICE * (100 - PCT) / 100) * CUST_DISC_PCT_EXTRA
                        End If
                        DISC.DISCOUNT_PCT = String.Format("{0}->{1}%", rowICTDISC1.Item("DISC_DESC"), PCT)
                        DISC.DISCOUNT_DESC = rowICTDISC1.Item(String.Format("DISC{0}_DESC", CStr(I))) & ""
                    End If
                    retval.Add(DISC)
                Next
            End If
        Else
            For I As Integer = 1 To 4
                Dim DISC As New DISCOUNTS() With {.DISCOUNT_QTY = 0, .DISCOUNT_PCT = "Problem With Style", .DISCOUNT_PRICE = 99999, .DISCOUNT_DESC = "Problem With Style"}
                retval.Add(DISC)
            Next
        End If
        Return retval
    End Function

    Public Shared Function getTimePassword(ByVal MinutesAhead As Integer) As String
        Dim Retval As String
        Dim ThisDay As Date = Now().ToUniversalTime
        Dim MIN As Integer = Now.Minute
        Dim Factor As Integer = 0
        If MIN < 60 Then
            Factor = 5
        End If
        If MIN < 50 Then
            Factor = 7
        End If
        If MIN < 40 Then
            Factor = 3
        End If
        If MIN < 30 Then
            Factor = 8
        End If
        If MIN < 20 Then
            Factor = 6
        End If
        If MIN < 10 Then
            Factor = 2
        End If
        Dim MM As String = Format(ThisDay.AddMonths(7 * Factor).Month, "00").ToString
        Dim DD As String = Format(ThisDay.AddDays(-18 * Factor).Day, "00").ToString
        Dim YY As String = Format(ThisDay.AddYears(12 * Factor).Year, "0000").ToString.Substring(2, 2)
        Dim HH As String = Format(ThisDay.AddHours(-4 * Factor).Hour, "00").ToString
        Dim MN As String = Format(ThisDay.Minute, "00").ToString
        Retval = YY + DD + MN + HH + MM

        Return Retval
    End Function

    Public Shared Function useTimePassword() As Boolean
        Dim UTime As Date = Now.ToUniversalTime
        Dim Retval As Boolean = False
        Dim iResult As String
        Dim iTitle As String = "Enter Timed Password"
        Dim iMSG As New System.Text.StringBuilder
        iMSG.AppendLine("Please Enter Your Timed Password:")
        iResult = InputBox(iMSG.ToString, iTitle)
        If iResult.Length <> 10 Then
            MsgBox("Password Must Be 10 Digits.", MsgBoxStyle.Critical, "Invalid Password")
            Return False
            Exit Function
        End If

        Dim MN As String = iResult.Substring(4, 2)
        Dim MIN As Integer = Val(MN)
        Dim Factor As Integer = 0
        If MIN < 60 Then
            Factor = 5
        End If
        If MIN < 50 Then
            Factor = 7
        End If
        If MIN < 40 Then
            Factor = 3
        End If
        If MIN < 30 Then
            Factor = 8
        End If
        If MIN < 20 Then
            Factor = 6
        End If
        If MIN < 10 Then
            Factor = 2
        End If

        Dim MM As String = Format(UTime.AddMonths(7 * Factor).Month, "00").ToString
        Dim DD As String = Format(UTime.AddDays(-18 * Factor).Day, "00").ToString
        Dim YY As String = Format(UTime.AddYears(12 * Factor).Year, "0000").ToString.Substring(2, 2)
        Dim HH As String = Format(UTime.AddHours(-4 * Factor).Hour, "00").ToString
        Dim HH2 As String = Format(UTime.AddHours((-4 * Factor) - 1).Hour, "00").ToString

        Dim H1 As String = YY + DD + MN + HH + MM
        Dim H2 As String = YY + DD + MN + HH2 + MM

        If iResult = H1 Or iResult = H2 Then
            Retval = True
        End If
        Return Retval
    End Function

    Public Shared Function TodaysOverRide(ByVal OverrideType As String) As String
        Dim Factor As Integer = 1
        If OverrideType = "F" Then
            Factor = 2
        End If
        Dim iResult As String
        Dim ThisDay As Date = Now.Date.AddDays(18)
        Dim MONTH As Integer = ThisDay.Month
        Dim DAY As Integer = ThisDay.Day
        Dim YEAR As Integer = ThisDay.Year
        Dim SeedNo As Integer = Val((YEAR + (DAY * Factor) + MONTH).ToString().Substring(3, 1))
        Dim U1 As String = ""
        If (DAY + SeedNo) > 24 Then
            U1 = Chr((14 + SeedNo) + 65).ToString
        Else
            U1 = Chr((DAY + SeedNo) + 65)
        End If
        Dim U2 As String = ""
        If (DAY + SeedNo + 10) > 24 Then
            U2 = Chr((SeedNo + 10) + 65).ToString
        Else
            U2 = Chr((DAY + SeedNo + 10) + 65)
        End If
        Dim U3 As String = ""
        If (DAY + SeedNo + 4) > 24 Then
            U3 = Chr((SeedNo + 4) + 97).ToString
        Else
            U3 = Chr((DAY + SeedNo + 4) + 97)
        End If
        Dim L1 As String = ""
        If (DAY + SeedNo + 6) > 24 Then
            L1 = Chr((SeedNo + 6) + 97).ToString
        Else
            L1 = Chr((DAY + SeedNo + 6) + 97)
        End If
        Dim L2 As String = ""
        If (DAY + SeedNo + 2) > 24 Then
            L2 = Chr((SeedNo + 2) + 97).ToString
        Else
            L2 = Chr((DAY + SeedNo + 2) + 97)
        End If
        Dim L3 As String = ""
        If (DAY + SeedNo + 1) > 24 Then
            L3 = Chr((SeedNo + 1) + 97).ToString
        Else
            L3 = Chr((DAY + SeedNo + 1) + 97)
        End If
        Dim N1 As String = SeedNo
        Dim N2 As String = 9 - SeedNo
        If SeedNo Mod 2 = 0 Then
            iResult = U1 & U2 & N1 & L1 & N2 & U3 & L2 & L3
        Else
            iResult = L1 & U1 & N1 & N2 & L2 & U2 & L3 & U3
        End If
        Return iResult
    End Function

    Public Shared Function GetSavedOverRide(ByVal OverrideType As String) As String
        Dim iResult As String = ""
        Dim FetchColumn As String = ""
        Select Case OverrideType
            Case "F"
                FetchColumn = "RO_PARM_LAST_FEPASS"
            Case "C"
                FetchColumn = "RO_PARM_LAST_CUSTPASS"
        End Select
        If FetchColumn <> "" Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            SQLS.AppendLine(String.Format("SELECT NVL({0},'NULL') AS RSLT FROM SOTPARMR WHERE SO_PARM_KEY = 'Z'", FetchColumn))
            ASCMAIN1.sql = SQLS.ToString()
            iResult = ASCDATA1.GetDataValue
            If iResult = "NULL" Then
                iResult = ""
            End If
        End If
        Return iResult
    End Function

    Public Shared Function SaveOverRide(ByVal OverrideType As String, ByVal PassWord As String) As Boolean
        Dim iResult As Boolean = False
        Dim UpdateColumn As String = ""
        Select Case OverrideType
            Case "F"
                UpdateColumn = "RO_PARM_LAST_FEPASS"
            Case "C"
                UpdateColumn = "RO_PARM_LAST_CUSTPASS"
        End Select
        If UpdateColumn <> "" Then
            Dim SQLS As New System.Text.StringBuilder
            SQLS.Length = 0
            PassWord = PassWord.Replace("'", "")
            PassWord = PassWord.Replace("Drop", "")
            PassWord = PassWord.Replace("Alter", "")
            SQLS.AppendLine(String.Format("UPDATE SOTPARMR SET {0} = '{1}' WHERE SO_PARM_KEY = 'Z'", UpdateColumn, PassWord))
            ASCMAIN1.sql = SQLS.ToString
            ASCDATA1.ExecuteSQL()
        End If
        Return iResult
    End Function

    Public Shared Function BackUpLaptop(frm As ASFBASE0) As Boolean
        Dim Zip1 As New nsoftware.IPWorksZip.Zip
        Zip1.RuntimeLicense = frm.nSoftwareKeys("nSoftwareZipkey")
        Dim RetVal As Boolean = True
        Dim BackUpFolder As String = "C:\Shared\RGO\Backup" 'This Should be Paramterized SomeDay
        Dim FileName As String = Format(Now.Year, "00").ToString & Format(Now.Month, "00").ToString & Format(Now.Day, "00").ToString
        Dim FileNum As Integer = 1
        Dim FILEDMP As String = String.Format("{0}\{1}_{2}.DMP", BackUpFolder, FileName, FileNum)
        Dim FILELOG As String = String.Format("{0}\{1}_{2}.LOG", BackUpFolder, FileName, FileNum)
        Dim FILEZIP As String = String.Format("{0}\{1}_{2}.ZIP", BackUpFolder, FileName, FileNum)
        Dim BatFile As String = BackUpFolder & "\RGOBackup.bat"
        Try
            If Not Directory.Exists(BackUpFolder) Then
                Directory.CreateDirectory(BackUpFolder)
            End If
            If File.Exists(BatFile) Then
                File.Delete(BatFile)
            End If

            Do While IO.File.Exists(FILEZIP)
                FileNum += 1
                FILEDMP = String.Format("{0}\{1}_{2}.DMP", BackUpFolder, FileName, FileNum)
                FILELOG = String.Format("{0}\{1}_{2}.LOG", BackUpFolder, FileName, FileNum)
                FILEZIP = String.Format("{0}\{1}_{2}.ZIP", BackUpFolder, FileName, FileNum)
            Loop
            Dim p As New System.Diagnostics.ProcessStartInfo()
            File.WriteAllText(BatFile, String.Format("EXP RGO/RGO FILE={0} LOG={1}", FILEDMP, FILELOG))
            With p
                .WindowStyle = ProcessWindowStyle.Minimized
                .WorkingDirectory = BackUpFolder
                .FileName = String.Format(BatFile)
                .UseShellExecute = True
            End With
            ASCMAIN1.Progress("Now Backing Up Database")
            Dim Proc As System.Diagnostics.Process = System.Diagnostics.Process.Start(p)
            Do While Not Proc.HasExited
            Loop
            Zip1.ArchiveFile = FILEZIP
            Zip1.IncludeFiles(FILEDMP)
            Zip1.IncludeFiles(FILELOG)
            Zip1.OverwriteFiles = True
            Zip1.Compress()
            Zip1.Dispose()
            File.Delete(FILEDMP)
            File.Delete(FILELOG)
            ASCMAIN1.Progress("")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Back Up Error")
            RetVal = False
        End Try
    End Function
End Class

Public Class REMOTE
    Public IsUserSuper As Boolean = False
    Public SREP_CODE As String = ""
    Public SQLWhere As String = ""
    Public Sub New(ByVal FF As ASFBASE1)
        ASCMAIN1.sql = String.Format("SELECT COUNT(*) AS RECCNT FROM ASTUSER2 WHERE USER_ID = '{0}' AND SECURITY_CODE = 'X6'", ASCMAIN1.USER_ID)
        If Val(ASCDATA1.GetDataValue) > 0 Then
            IsUserSuper = True
            SetSREP_CODE()
            SQLWhere = ""
        Else
            IsUserSuper = False
            SetSREP_CODE()
            If SREP_CODE <> "" Then
                'SQLWhere = String.Format("SREP_CODE = '{0}'", SREP_CODE)
                'Special Code to marry MD & JD together.
                'If SREP_CODE = "MD" Or SREP_CODE = "JD" Or SREP_CODE = "JE" Then
                '    SQLWhere = "(SREP_CODE IN ('MD','JD','JE') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('MD','JD','JE')))"
                'Else
                '    SQLWhere = String.Format("(SREP_CODE = '{0}' OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE = '{0}'))", SREP_CODE)
                'End If

                'New Code to Marry Certain Sales Reps Together - WR 03/07/19
                'Case "CB"
                '    SQLWhere = "(SREP_CODE IN ('CB','ST') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('CB','ST')))"
                'Case "JB"
                '    SQLWhere = "(SREP_CODE IN ('JB','ST') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('JB','ST')))"
                Select Case SREP_CODE
                    Case "DA"
                        SQLWhere = "(SREP_CODE IN ('DA','CH','CC','CD','CK') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('DA','CH','CC','CD','CK')))"
                    Case "MD", "JD", "JE"
                        SQLWhere = "(SREP_CODE IN ('MD','JD','JE') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('MD','JD','JE')))"
                    Case "TN"
                        If CDate(Now().ToShortDateString) < CDate("11/01/2020") Then
                            SQLWhere = "(SREP_CODE IN ('TN','CB','JB') OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE IN ('TN','CB','JB')))"
                        Else
                            SQLWhere = String.Format("(SREP_CODE = '{0}' OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE = '{0}'))", SREP_CODE)
                        End If
                    Case Else
                        SQLWhere = String.Format("(SREP_CODE = '{0}' OR CUST_CODE IN (SELECT CUST_CODE FROM ARTCUST1 WHERE SREP_CODE = '{0}'))", SREP_CODE)
                End Select
                'Added so Sales Reps can see all orders for their customers regardless of who wrote it. - w.r. - 8/11/14
            Else
                SQLWhere = "SREP_CODE = 'NONE'"
            End If
        End If
    End Sub
    Private Sub SetSREP_CODE()
        ASCMAIN1.sql = String.Format("SELECT SREP_CODE FROM TATUSER1 WHERE USER_ID = '{0}'", ASCMAIN1.USER_ID)
        SREP_CODE = ASCDATA1.GetDataValue
    End Sub
End Class

Public Class HANGTAG
    Public ErrMsg As String
    Public LabelContinued As Boolean = False
    Public LabelContinued3 As Boolean = False
    Private BackColors As String = ""
    Private frm As ASFBASE1
    Private Discounts As List(Of DISCOUNTS)
    Private ColorsFront As String
    Private ColorsBack As String
    Private Colors3 As String
    Private Printer As String
    Private EXCLUSIVE_STYLE As String = ""
    'You will need the following datatable made in the calling form before making this
    'ASCMAIN1.sql = "SELECT * FROM ICTSTYL1"
    'Create_TDA(.Tables.Add, "HANGTAG1", "**", 0, False)
    '.Tables("HANGTAG1").Columns.Add("DATEPRINTED", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("BOXQTY", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("CARTQTY", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("COLORS", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("Price1_LBL", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("Price1_AMT", GetType(System.Double))
    '.Tables("HANGTAG1").Columns.Add("Price2_LBL", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("Price2_AMT", GetType(System.Double))
    '.Tables("HANGTAG1").Columns.Add("Price3_LBL", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("Price3_AMT", GetType(System.Double))
    '.Tables("HANGTAG1").Columns.Add("Price4_LBL", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("Price4_AMT", GetType(System.Double))
    '.Tables("HANGTAG1").Columns.Add("COLORSDESC", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("VEND_SUPPLIER_ID", GetType(System.String))
    '.Tables("HANGTAG1").Columns.Add("PORT_CODE_ORIG", GetType(System.String))

    Public Sub New(ByVal FF As ASFBASE1, ByVal STYLE_CODE As String, Dis As List(Of DISCOUNTS), ByVal P As String)
        ErrMsg = ""
        frm = FF
        Discounts = Dis
        Printer = P
        frm.Fill_Records("HANGTAG1", STYLE_CODE, True)
        If frm.dst.Tables("HANGTAG1").Rows.Count = 1 Then
            EXCLUSIVE_STYLE = frm.dst.Tables("HANGTAG1").Rows(0).Item("EXCLUSIVE_STYLE").ToString & ""
        End If
        If frm.dst.Tables("HANGTAG1").Rows.Count > 0 Then
            FillExtraFields()
        End If
    End Sub

    Private Function GetVendorData(ByVal VEND_CODE As String, ByVal COLUMN As String) As String
        Dim RetVal As String = ""
        If VEND_CODE.Length > 0 And COLUMN.Length > 0 Then
            ASCMAIN1.sql = String.Format("SELECT {0} FROM APTVEND1 WHERE VEND_CODE = '{1}'", COLUMN, VEND_CODE)
            RetVal = ASCDATA1.GetDataValue
        End If
        Return RetVal
    End Function
    Private Sub FillExtraFields()
        Dim rowHANGTAG1 As DataRow = frm.dst.Tables("HANGTAG1").Rows(0)
        Dim CARTON_PACK_QTY As Integer = Val(rowHANGTAG1.Item("CARTON_PACK_QTY"))
        Dim INNER_PACK_QTY As Integer = Val(rowHANGTAG1.Item("INNER_PACK_QTY"))
        Dim SUB_UNIT_PACK_QTY As Integer = Val(rowHANGTAG1.Item("SUB_UNIT_PACK_QTY") & "")
        Dim STYLE_UOM As String = rowHANGTAG1.Item("STYLE_UOM")
        CalculateColors(rowHANGTAG1.Item("STYLE_CODE"))
        rowHANGTAG1.Item("DATEPRINTED") = String.Format("{0}/{1}/{2}", Now.Date.Month, Now.Day, Now.Year.ToString.Substring(2, 2))
        rowHANGTAG1.Item("BOXQTY") = String.Format("BOX:{0}", (INNER_PACK_QTY * SUB_UNIT_PACK_QTY))
        rowHANGTAG1.Item("CARTQTY") = String.Format("CART:{0}", CARTON_PACK_QTY)
        rowHANGTAG1.Item("COLORS") = ColorsBack
        rowHANGTAG1.Item("COLORS3") = Colors3
        rowHANGTAG1.Item("Price1_LBL") = String.Format("{0}{1} ", Discounts(3).DISCOUNT_QTY, STYLE_UOM)
        rowHANGTAG1.Item("Price1_AMT") = Discounts(3).DISCOUNT_PRICE
        rowHANGTAG1.Item("Price2_LBL") = String.Format("{0}{1} ", Discounts(2).DISCOUNT_QTY, STYLE_UOM)
        rowHANGTAG1.Item("Price2_AMT") = Discounts(2).DISCOUNT_PRICE
        rowHANGTAG1.Item("Price3_LBL") = String.Format("{0}{1} ", Discounts(1).DISCOUNT_QTY, STYLE_UOM)
        rowHANGTAG1.Item("Price3_AMT") = Discounts(1).DISCOUNT_PRICE
        rowHANGTAG1.Item("Price4_LBL") = String.Format("{0}{1} ", Discounts(0).DISCOUNT_QTY, STYLE_UOM)
        rowHANGTAG1.Item("Price4_AMT") = Discounts(0).DISCOUNT_PRICE
        rowHANGTAG1.Item("COLORSDESC") = ColorsFront
        rowHANGTAG1.Item("VEND_SUPPLIER_ID") = GetVendorData(rowHANGTAG1.Item("VEND_CODE"), "VEND_SUPPLIER_ID")
        rowHANGTAG1.Item("PORT_CODE_ORIG") = GetVendorData(rowHANGTAG1.Item("VEND_CODE"), "PORT_CODE")
    End Sub

    Private Sub CalculateColors(ByVal STYLE_CODE As String)
        Dim ColorLines As String()
        Dim RowNum As Integer = 0
        ReDim ColorLines(RowNum)
        Dim NewLine As Boolean = True
        For Each rowICTSTYC1 As DataRow In frm.dst.Tables("ICTSTYC1").Select(String.Format("STYLE_CODE = '{0}'", STYLE_CODE), "COLOR_CODE")
            Dim ColorAdd As String = rowICTSTYC1.Item("COLOR_CODE_LONG").ToString
            Dim STATUS As String = rowICTSTYC1.Item("STYLE_COLOR_STATUS").ToString
            Dim MSQT As Integer = Val(rowICTSTYC1.Item("MSOH").ToString) + Val(rowICTSTYC1.Item("MSFT").ToString)
            If STATUS = "N" Then
                ColorAdd = ColorAdd & "^"
            End If
            If STATUS = "D" Then
                ColorAdd = ColorAdd & "*"
            End If


            If (String.Format("{0}, {1}", ColorLines(RowNum), ColorAdd)).Length < 37 Then
                If NewLine Then
                    ColorLines(RowNum) = ColorAdd
                Else
                    ColorLines(RowNum) = (String.Format("{0}, {1}", ColorLines(RowNum), ColorAdd))
                End If

                NewLine = False
            Else
                RowNum += 1
                'NewLine = True
                ReDim Preserve ColorLines(RowNum)
                ColorLines(RowNum) = ColorAdd
                If RowNum = 3 Then
                    RowNum += 1
                    ReDim Preserve ColorLines(RowNum)
                    ColorLines(RowNum) = ColorLines(RowNum - 1)
                    ColorLines(RowNum - 1) = ColorLines(RowNum - 2)
                    ColorLines(RowNum - 2) = "* Additional Colors On Back *"
                    LabelContinued = True
                End If
                If RowNum = 8 Then
                    LabelContinued3 = True
                End If
            End If
        Next
        NewLine = True
        For i As Integer = 0 To RowNum
            If i <= 2 Then
                If NewLine Then
                    ColorsFront = ColorLines(i)
                    NewLine = False
                Else
                    ColorsFront = ColorsFront & vbCrLf & ColorLines(i)
                End If
            Else
                If i <= 7 Then
                    If NewLine = False Then
                        ColorsBack = ColorLines(i)
                        NewLine = True
                    Else
                        ColorsBack = ColorsBack & vbCrLf & ColorLines(i)
                    End If
                    If i = 7 Then
                        NewLine = False
                    End If
                Else
                    If NewLine = False Then
                        Colors3 = ColorLines(i)
                        NewLine = True
                    Else
                        Colors3 = Colors3 & vbCrLf & ColorLines(i)
                    End If
                End If
            End If
        Next
        'ColorsFront = "COLOR1, COLOR2, COLOR3, COLOR4" & vbCrLf & "* Colors On Back *"
        'ColorsBack = "GOLD,SILVER,RED,WHITE,COPPER,GREEN,SAGE" & vbCrLf & "BURGUNDY,ROSEWOOD,CHOCOLATE,HOTPINK" & vbCrLf & "LIGHTBLUE,BLACK,BLUE,TURQUOISE" & vbCrLf & "APPLE/GREEN DARK GOLD"
    End Sub

    Public Sub Print()
        frm.Print_Report_Begin()
        'frm.CR_params.Add("SUBT", "")
        'frm.CR_params.Add("BLDG", Absx1.txtFor("BUILDING_DESC").Text)
        'frm.Generate_Report("HANGTG1")
        '        Dim Printer As string = GetPrinterName()
        If Printer.Length > 0 Then
            frm.CR_params.Add("EXCLUSIVE_STYLE", EXCLUSIVE_STYLE)

            frm.Generate_Report("HANGTG1")
            If LabelContinued Then
                frm.Generate_Report("HANGTG2")
            End If
            If LabelContinued3 Then
                frm.Generate_Report("HANGTG3")
            End If
            If ASCMAIN1.USER_ID = "wayne" Or ASCMAIN1.USER_ID = "mariog" Then
                frm.Print_Report_End()
            Else
                frm.Print_Report_End(True, False, Printer, 1)
            End If
        Else
            MsgBox("Default Printer Not Defined", MsgBoxStyle.OkOnly, "Printer Set-up")
        End If
    End Sub

    Private Function GetPrinterName() As String
        Dim RetVal As String = ""
        Return RetVal
    End Function
End Class

Public Class DISCOUNTS
    Public DISCOUNT_DESC As String
    Public DISCOUNT_QTY As Integer
    Public DISCOUNT_PRICE As Double
    Public DISCOUNT_PCT As String
End Class

Public Class FEFDPrice
    Public FEPrice As Double
    Public FEMixPrice As Double
    Public FDMixPrice As Double
    Public FDPrice As Double
    Public ErrorMsg As String

    Public Sub New(ByVal FF As ASFBASE1, ByVal STYLE_CODE As String, Optional FACTOR As Double = 1, Optional ByVal SHOW_FD_CALC As Boolean = False)
        Dim SQLS As New System.Text.StringBuilder
        Dim rowSOTPARM2 As DataRow = FF.LookUp("SOTPARM2", "Z")
        If IsNothing(rowSOTPARM2) Then
            ThrowError("Parameter Table Has No Record", STYLE_CODE)
            Exit Sub
        Else
            Dim rowICTSTYL1 As DataRow = FF.LookUp("ICTSTYL1", STYLE_CODE)
            If IsNothing(rowICTSTYL1) Then
                ThrowError("Style Not Found", STYLE_CODE)
                Exit Sub
            Else
                Dim PackCUFeet As Double = 0
                If IsDBNull(rowICTSTYL1.Item("CASE_CUBE")) Then
                    ThrowError("Cube Not Set In MF", STYLE_CODE)
                    Exit Sub
                Else
                    PackCUFeet = rowICTSTYL1.Item("CASE_CUBE")
                End If

                Dim CartonQty As Double = 0
                If IsDBNull(rowICTSTYL1.Item("CARTON_PACK_QTY")) Then
                    ThrowError("Carton Pack Not Set In MF", STYLE_CODE)
                    Exit Sub
                Else
                    CartonQty = rowICTSTYL1.Item("CARTON_PACK_QTY")
                End If
                SQLS.Length = 0
                SQLS.AppendLine("SELECT NVL(PO_COST,0) AS PO_COST")
                SQLS.AppendLine(" FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND PO_COST_DATE = (")
                SQLS.AppendLine(" SELECT MAX(PO_COST_DATE) FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND NVL(PO_COST,0) <> 0")
                SQLS.AppendLine(" AND TO_DATE(PO_COST_DATE, 'DD/MM/YYYY') <= sysdate)")
                ASCMAIN1.sql = SQLS.ToString()
                Dim PO_COST As Double = Val(ASCDATA1.GetDataValue)
                SQLS.Length = 0
                SQLS.AppendLine("SELECT NVL(NEW_PO_COST,0) AS NEW_PO_COST")
                SQLS.AppendLine(" FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND NEW_PO_COST_DATE = (")
                SQLS.AppendLine(" SELECT MAX(NEW_PO_COST_DATE) FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND NVL(NEW_PO_COST,0) <> 0")
                SQLS.AppendLine(" AND NEW_PO_COST_DATE <= sysdate)")
                ASCMAIN1.sql = SQLS.ToString()
                Dim NEW_PO_COST As Double = Val(ASCDATA1.GetDataValue)
                SQLS.Length = 0
                SQLS.AppendLine("SELECT NVL(NEW_PO_COST_DATE,'01-JAN-1900') AS NEW_PO_COST_DATE")
                SQLS.AppendLine(" FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND NEW_PO_COST_DATE = (")
                SQLS.AppendLine(" SELECT MAX(NEW_PO_COST_DATE) FROM ICTSTYV1")
                SQLS.AppendLine(String.Format(" WHERE STYLE_CODE = '{0}'", STYLE_CODE))
                SQLS.AppendLine(" AND NVL(NEW_PO_COST,0) <> 0")
                SQLS.AppendLine(" AND NEW_PO_COST_DATE <= sysdate)")
                ASCMAIN1.sql = SQLS.ToString()
                Dim NEW_PO_COST_DATE As String = CStr(ASCDATA1.GetDataValue)
                If PO_COST = 0 And NEW_PO_COST = 0 Then
                    ThrowError("No Cost Found", STYLE_CODE)
                    Exit Sub
                Else
                    If NEW_PO_COST <> 0 Then
                        If IsDate(NEW_PO_COST_DATE) Then
                            If CDate(NEW_PO_COST_DATE) < Now() Then
                                PO_COST = NEW_PO_COST
                            End If
                        End If
                    End If
                End If
                Dim FACTNUM As Double
                If FACTOR = 0 Or FACTOR = 1 Then
                    FACTNUM = 1
                Else
                    FACTNUM = (1 + (((100 + FACTOR) - 100) / 100))
                End If

                Dim DUTY_RATE_CODE As String = rowICTSTYL1.Item("DUTY_RATE_CODE").ToString & String.Empty
                Dim COUNTRY_CODE As String = rowICTSTYL1.Item("COUNTRY_CODE").ToString & String.Empty
                Dim DUTY_RATE As Double = 0

                If COUNTRY_CODE.Length > 0 And DUTY_RATE_CODE.Length > 0 Then
                    Dim sql As New System.Text.StringBuilder With {.Length = 0}
                    sql.AppendLine("SELECT *")
                    sql.AppendLine("FROM ICTDUTY4")
                    sql.AppendLine(String.Format("WHERE DUTY_RATE_CODE = '{0}'", DUTY_RATE_CODE))
                    sql.AppendLine(String.Format("AND COUNTRY_CODE = '{0}'", COUNTRY_CODE))

                    Dim tblICTDUTY4 As DataTable = ASCDATA1.GetDataTable(sql.ToString())
                    For Each rowICTDUTY4 As DataRow In tblICTDUTY4.Select("", "DUTY_RATE_BEGIN")
                        Dim DUTY_RATE_END As String = rowICTDUTY4.Item("DUTY_RATE_END").ToString & String.Empty
                        If IsDate(DUTY_RATE_END) Then
                            Dim TODAY_BEG As DateTime = CDate(Now().ToShortDateString)
                            If CDate(DUTY_RATE_END) > TODAY_BEG Then
                                DUTY_RATE = (Val(rowICTDUTY4.Item("DUTY_RATE").ToString & String.Empty) * 0.01)
                            End If
                        Else
                            DUTY_RATE = (Val(rowICTDUTY4.Item("DUTY_RATE").ToString & String.Empty) * 0.01)
                        End If
                    Next
                End If
                Dim DUTY_WITH_TARRIF As Double = DUTY_RATE + Val(rowSOTPARM2.Item("SO_PARM_DUTY"))
                FEPrice = (PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT"))) * FACTNUM
                FEMixPrice = (((Val(rowSOTPARM2.Item("SO_PARM_CONCOST")) * PackCUFeet) / Val(CartonQty)) + (PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT")))) * FACTNUM
                'FDMixPrice = ((PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT"))) * Val(rowSOTPARM2.Item("SO_PARM_DUTY"))) + (Val(rowSOTPARM2.Item("SO_PARM_INLANDFRT")) / Val(CartonQty)) + ((Val(rowSOTPARM2.Item("SO_PARM_OCEANFRTCONS")) * PackCUFeet) / Val(CartonQty)) * FACTNUM
                'FDPrice = ((PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT"))) * Val(rowSOTPARM2.Item("SO_PARM_DUTY"))) + (Val(rowSOTPARM2.Item("SO_PARM_INLANDFRT")) / Val(CartonQty)) + ((Val(rowSOTPARM2.Item("SO_PARM_OCEANFRT")) * PackCUFeet) / Val(CartonQty)) * FACTNUM
                FDMixPrice = ((PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT"))) * DUTY_WITH_TARRIF) + (Val(rowSOTPARM2.Item("SO_PARM_INLANDFRT")) / Val(CartonQty)) + ((Val(rowSOTPARM2.Item("SO_PARM_OCEANFRTCONS")) * PackCUFeet) / Val(CartonQty)) * FACTNUM
                FDPrice = ((PO_COST * Val(rowSOTPARM2.Item("SO_PARM_FEFACT"))) * DUTY_WITH_TARRIF) + (Val(rowSOTPARM2.Item("SO_PARM_INLANDFRT")) / Val(CartonQty)) + ((Val(rowSOTPARM2.Item("SO_PARM_OCEANFRT")) * PackCUFeet) / Val(CartonQty)) * FACTNUM
                FEPrice = TAC.ICCMAIN1.Calculate_Style_Royalty_Markup(FF, STYLE_CODE, FEPrice)
                FEMixPrice = TAC.ICCMAIN1.Calculate_Style_Royalty_Markup(FF, STYLE_CODE, FEMixPrice)
                FDMixPrice = TAC.ICCMAIN1.Calculate_Style_Royalty_Markup(FF, STYLE_CODE, FDMixPrice)
                FDPrice = TAC.ICCMAIN1.Calculate_Style_Royalty_Markup(FF, STYLE_CODE, FDPrice)
                If SHOW_FD_CALC Then
                    Dim Msg As New System.Text.StringBuilder With {.Length = 0}
                    Msg.AppendLine(String.Format("DUTY_RATE_CODE: {0}", DUTY_RATE_CODE))
                    Msg.AppendLine("")
                    Msg.AppendLine(String.Format("PO_COST: {0}", PO_COST))
                    Msg.AppendLine(String.Format("DUTY_RATE: {0}", DUTY_RATE))
                    Msg.AppendLine(String.Format("CART_QTY: {0}", CartonQty))
                    Msg.AppendLine(String.Format("CUFEET: {0}", PackCUFeet))
                    Msg.AppendLine("")
                    Msg.AppendLine(String.Format("Results: {0}", FDMixPrice))
                    MsgBox(Msg.ToString, vbOKOnly, String.Format("Calculation For {0}", STYLE_CODE))
                End If

            End If
        End If
    End Sub

    Private Sub SetZeros()
        FEPrice = 0
        FEMixPrice = 0
        FDMixPrice = 0
        FDPrice = 0
    End Sub

    Private Sub ThrowError(ByVal ErrMsg As String, ByVal STYLE_CODE As String)
        ErrorMsg = String.Format("{0} For Style {1}", ErrMsg, STYLE_CODE)
        SetZeros()
    End Sub
End Class

Public Class TATCTLN3
    Private _CTL_NO_TYPE As String
    Public ErrMsg As String
    Public NumbersRemaining As Integer
    Public Next_ctl_no As String
    Public Sub New(ByVal CTL_NO_TYPE As String, ByVal FF As ASFBASE1)
        _CTL_NO_TYPE = CTL_NO_TYPE
        ASCMAIN1.sql = "SELECT COUNT(*)" _
        & " FROM TATCTLN3" _
        & " WHERE CTL_NO_TYPE = '" & _CTL_NO_TYPE & "'"
        NumbersRemaining = Val(ASCDATA1.GetDataValue)
        If NumbersRemaining = 0 Then
            ErrMsg = _CTL_NO_TYPE & "Count = 0"
        Else
            ASCMAIN1.sql = "SELECT min(ctl_no) as Next_ctl_no" _
            & " FROM TATCTLN3" _
            & " WHERE CTL_NO_TYPE = '" & CTL_NO_TYPE & "'"
            Next_ctl_no = ASCDATA1.GetDataValue
            ASCMAIN1.sql = "DELETE FROM TATCTLN3" _
            & " WHERE CTL_NO_TYPE = '" & CTL_NO_TYPE & "'" _
            & " AND ctl_no = '" & Next_ctl_no & "'"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub
End Class