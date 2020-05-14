Imports System.Windows.Forms
Imports System.Math

Public Class ASFPWDC1

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        ASCMAIN1.Message = ""
        Me.Close()
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Dim Emsg As String = ""
        Dim newPassword As String = ""

        txtUSER_ID.Text = txtUSER_ID.Text.Trim
        If txtUSER_ID.Text <> ASCMAIN1.USER_ID Then
            Emsg = vbCr & "Invalid User ID" & Emsg
        End If

        txtUSER_PASSWORD.Text = txtUSER_PASSWORD.Text.Trim
        If txtUSER_PASSWORD.Text.Trim <> ASCMAIN1.USER_PASSWORD Then
            Emsg = vbCr & "Invalid User Password" & Emsg
        End If

        txtNewPass.Text = txtNewPass.Text.Trim
        txtNewPassVer.Text = txtNewPass.Text.Trim

        If txtNewPass.Text <> txtNewPassVer.Text Then
            Emsg = vbCr & "New Password entries are not equal." & Emsg
        ElseIf txtNewPass.Text.Length = 0 Then
            Emsg = vbCr & "New Password entries are missing." & Emsg
        End If

        If txtNewPass.Text.ToUpper = txtUSER_PASSWORD.Text.ToUpper Then
            Emsg = vbCr & "New Password may not be the same as your current password." & Emsg
        End If

        If Emsg.Length > 0 Then
            Emsg = Emsg.Substring(1)
            MsgBox(Emsg, MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If

        Dim tblASTPARMP As DataTable = ASCDATA1.GetDataTable("Select * from ASTPARMP where AS_PARM_KEY = 'Z'", "ASTPARMP")
        If tblASTPARMP.Rows.Count = 0 Then
            MsgBox("Table ASTPARMP is missing parameters.", MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If

        newPassword = txtNewPass.Text

        Dim ErrorMsgs() As UserPasswordError = Validate_User_Password(txtUSER_ID.Text, newPassword)

        Emsg = ""
        For eCount As Integer = LBound(ErrorMsgs) To UBound(ErrorMsgs)
            If (ErrorMsgs(eCount).EMsg & "").Trim.Length > 0 Then
                Emsg = vbCr & ErrorMsgs(eCount).EMsg.Trim & Emsg
            End If
        Next

        If Emsg.Length > 0 Then
            Emsg = Emsg.Substring(1)
            MsgBox(Emsg, MsgBoxStyle.OkOnly, "Change Password Error")
            Exit Sub
        End If


        If tblASTPARMP.Rows(0).Item("AS_PARM_PWD_ENCRYPTED").ToString = "1" Then
            Dim MD5 As New ASCSCMD5
            'Stop ' SHOULD PROBABLY BE USING System.Security.Cryptography 
            newPassword = MD5.DigestStrToHexStr(newPassword)
        End If

        ASCMAIN1.sql = "Update ASTUSER1 Set USER_PASSWORD = '" & newPassword & "'" & _
            " , USER_PASSWORD_LAST_DATE = '" & Format(Now, "dd-MMM-yyyy") & "'" & _
            " where USER_ID = '" & txtUSER_ID.Text & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Insert Into ASTUSERP (USER_ID, INIT_DATE, USER_PASSWORD)" & _
            " VALUES ('" & ASCMAIN1.USER_ID & "', " & _
            "'" & Format(Now, "dd-MMM-yyyy") & "'," & _
            "'" & newPassword & "')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        MsgBox("Password successfully changed.", MsgBoxStyle.OkOnly, "Change Password")
        ASCMAIN1.Message = "X"

        Me.Close()

    End Sub

    ' Move to a centralized location so user maintenance cmay use it
#Region "Validate User Passwords"

    Public Const ABSPWKEY = "Pm6#9&LG%<?"

    Public Structure UserPasswordError
        Dim UserID As String
        Dim UserName As String
        Dim EMsg As String
    End Structure

    ' Make Public when this is moved to centralized location
    Private Function Validate_User_Password(ByVal UserID As String, ByVal UserPassword As String, _
        Optional ByVal Active_Only As Boolean = True, Optional ByVal TABLE_NAME As String = "") As UserPasswordError()
        ' Parameters
        ' DecryptPassword - If True then Decrypt the Password
        ' UserID - If provided then only process this user
        ' User_Password - If Provided then only validate this Password
        ' Active_Only - Only evaluate Active Users. Should be set to True when UserID and Password = ""
        ' TABLE_NAME - Identifes a table other than ASTPARMP to get settings from
        '   Table must have same layout as ASTPARMP

        Dim EMsg() As UserPasswordError

        Dim USER_ID As String
        Dim USER_NAME As String
        Dim User_Password As String
        Dim sql As String = ""
        Dim sqlwhere As String = ""

        If TABLE_NAME = "" Then TABLE_NAME = "ASTPARMP"

        Dim tblASTPARMP As DataTable = ASCDATA1.GetDataTable("Select * from " & TABLE_NAME & " where AS_PARM_KEY = 'Z'", "ASTPARMP")

        If tblASTPARMP.Rows.Count = 0 Then
            MsgBox("Table ASTPARMP is missing parameters.", MsgBoxStyle.OkOnly, "Change Password Error")
            ReDim EMsg(1)
            EMsg(1).UserID = "All"
            EMsg(1).UserName = "All"
            EMsg(1).EMsg = "Cannot locate parameter record in ASTPARMP"
            Return EMsg
        End If

        Dim rowASTPARMP As DataRow = tblASTPARMP.Rows(0)

        ReDim EMsg(1)

        If UserPassword <> "" Then
            USER_ID = UserID
            User_Password = UserPassword
            Call Evaluate_Password(EMsg, USER_ID, USER_ID, User_Password)
        Else
            sqlwhere = ""
            sql = "Select * from ASTUSER1 "
            If Active_Only Then
                sqlwhere = sqlwhere & "And USER_STATUS = 'A' "
            End If
            If UserID <> "" Then
                sqlwhere = sqlwhere & "And USER_ID = '" & UserID & "' "
            End If
            sqlwhere = sqlwhere.ToUpper.Trim
            If sqlwhere.Length > 0 Then
                If sqlwhere.Substring(0, 3) = "AND" Then
                    sqlwhere = " Where " & sqlwhere.Substring(3)
                End If
            End If
            sql = sql & sqlwhere & " Order By USER_ID"

            Dim tblASTUSER1 As DataTable = ASCDATA1.GetDataTable(sql, "ASTUSER1")

            For Each dr As DataRow In tblASTUSER1.Rows
                USER_ID = dr.Item("USER_ID") & ""
                USER_NAME = dr.Item("USER_NAME") & ""
                If tblASTPARMP.Rows(0).Item("AS_PARM_PWD_ENCRYPTED").ToString = "1" Then
                    User_Password = RndCrypt(dr.Item("USER_PASSWORD") & "", ABSPWKEY)
                Else
                    User_Password = dr.Item("USER_PASSWORD") & ""
                End If
                Call Evaluate_Password(EMsg, USER_ID, USER_NAME, User_Password)
            Next
        End If

        Validate_User_Password = EMsg
        Exit Function

    End Function

    Sub Evaluate_Password(ByRef EMsg() As UserPasswordError, ByVal User_ID As String, ByVal User_Name As String, _
        ByVal User_Password As String, Optional ByVal TABLE_NAME As String = "")

        Dim AS_PARM_PWD_MIN_LEN As Integer
        Dim AS_PARM_PWD_REQ_MIX_AN As String
        Dim AS_PARM_PWD_REQ_MIX_CASE As String
        Dim AS_PARM_PWD_REQ_MIX_NON_AN As String
        Dim AS_PARM_PWD_NO_USER_ID As String
        Dim AS_PARM_PWD_NO_USER_ID_PERM As String
        Dim AS_PARM_PWD_REUSE As Long = 0

        Dim has_alpha As Boolean
        Dim has_numeric As Boolean
        Dim has_upper As Boolean
        Dim has_lower As Boolean
        Dim has_non_an As Boolean

        If TABLE_NAME = "" Then TABLE_NAME = "ASTPARMP"

        Dim tblASTPARMP As DataTable = ASCDATA1.GetDataTable("Select * from " & TABLE_NAME & " where AS_PARM_KEY = 'Z'", "ASTPARMP")

        If tblASTPARMP.Rows.Count = 0 Then
            MsgBox("Table ASTPARMP is missing parameters.", MsgBoxStyle.OkOnly, "Change Password Error")
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = "All"
            EMsg(UBound(EMsg)).UserName = "All"
            EMsg(UBound(EMsg)).EMsg = "Cannot locate parameter record in ASTPARMP."
            Exit Sub
        End If

        Dim rowASTPARMP As DataRow = tblASTPARMP.Rows(0)

        AS_PARM_PWD_MIN_LEN = Val(rowASTPARMP.Item("AS_PARM_PWD_MIN_LEN") & "")
        AS_PARM_PWD_REQ_MIX_AN = Val(rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_AN") & "")
        AS_PARM_PWD_REQ_MIX_CASE = Val(rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_CASE") & "")
        AS_PARM_PWD_REQ_MIX_NON_AN = Val(rowASTPARMP.Item("AS_PARM_PWD_REQ_MIX_NON_AN") & "")
        AS_PARM_PWD_NO_USER_ID = Val(rowASTPARMP.Item("AS_PARM_PWD_NO_USER_ID") & "")
        AS_PARM_PWD_NO_USER_ID_PERM = Val(rowASTPARMP.Item("AS_PARM_PWD_NO_USER_ID_PERM") & "")
        AS_PARM_PWD_REUSE = Val(rowASTPARMP.Item("AS_PARM_PWD_REUSE") & "")

        Dim Usorted As String
        Dim Psorted As String
        ReDim EMsg(1)

        If User_Password.Length < AS_PARM_PWD_MIN_LEN Then
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = User_ID
            EMsg(UBound(EMsg)).UserName = User_Name
            EMsg(UBound(EMsg)).EMsg = "Pasword minimum length is " & AS_PARM_PWD_MIN_LEN & " characters."
        End If

        Usorted = X_Sort(User_ID, has_alpha, has_numeric, has_upper, has_lower, has_non_an)
        Psorted = X_Sort(User_Password, has_alpha, has_numeric, has_upper, has_lower, has_non_an)

        If Abs(has_alpha + has_numeric + has_non_an) < 2 And AS_PARM_PWD_REQ_MIX_AN = "1" Then
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = User_ID
            EMsg(UBound(EMsg)).UserName = User_Name
            EMsg(UBound(EMsg)).EMsg = "Password may not be pure alpha or pure numeric."
        End If

        If Abs(has_upper + has_lower) < 2 And AS_PARM_PWD_REQ_MIX_CASE = "1" Then
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = User_ID
            EMsg(UBound(EMsg)).UserName = User_Name
            EMsg(UBound(EMsg)).EMsg = "Password requires mix case."
        End If

        If Not has_non_an And AS_PARM_PWD_REQ_MIX_NON_AN = "1" Then
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = User_ID
            EMsg(UBound(EMsg)).UserName = User_Name
            EMsg(UBound(EMsg)).EMsg = "Password requires special characters."
        End If

        If InStr(1, User_Password, "'") > 0 Then
            ReDim Preserve EMsg(UBound(EMsg) + 1)
            EMsg(UBound(EMsg)).UserID = User_ID
            EMsg(UBound(EMsg)).UserName = User_Name
            EMsg(UBound(EMsg)).EMsg = "Password may not use the apostrophe special character."
        End If

        If User_ID = User_Password Then
            If AS_PARM_PWD_NO_USER_ID = "1" Then
                ReDim Preserve EMsg(UBound(EMsg) + 1)
                EMsg(UBound(EMsg)).UserID = User_ID
                EMsg(UBound(EMsg)).UserName = User_Name
                EMsg(UBound(EMsg)).EMsg = "Password may not be the same as the User ID."
            End If
        Else
            If (Usorted.ToUpper = Psorted.ToUpper) And AS_PARM_PWD_NO_USER_ID_PERM = "1" Then
                ReDim Preserve EMsg(UBound(EMsg) + 1)
                EMsg(UBound(EMsg)).UserID = User_ID
                EMsg(UBound(EMsg)).UserName = User_Name
                EMsg(UBound(EMsg)).EMsg = "Password may not be a permutation of User ID."
            End If
        End If

        If AS_PARM_PWD_REUSE > 0 Then
            Dim sqlx As String = "Select * From ASTUSERP " & _
                " Where USER_ID = '" & User_ID & "'" & _
                " And UPPER(USER_PASSWORD) = '" & User_Password.ToUpper & "'" & _
                " And INIT_DATE > '" & Format(DateAdd(DateInterval.Day, -1 * AS_PARM_PWD_REUSE, Now), "dd-MMM-yyyy") & "'"
            'Dim tblASTUSERP As DataTable = ASCDATA1.GetDataTable(sqlx, "ASTUSERP")
            Dim dr As DataRow = ASCDATA1.GetDataRow(sqlx)
            'If tblASTUSERP.Rows.Count > 0 Then
            If Not dr Is Nothing Then
                ReDim Preserve EMsg(UBound(EMsg) + 1)
                EMsg(UBound(EMsg)).UserID = User_ID
                EMsg(UBound(EMsg)).UserName = User_Name
                EMsg(UBound(EMsg)).EMsg = "Password has already been used in the past."
            End If
        End If

    End Sub

    Public Function X_Sort(ByVal a As String, ByRef has_alpha As Boolean, ByRef has_numeric As Boolean, _
        ByRef has_upper As Boolean, ByRef has_lower As Boolean, ByRef has_non_an As Boolean) As String

        Dim z As String
        Dim i As Integer
        Dim j As Integer

        Dim b As String
        b = ""

        has_alpha = False
        has_numeric = False
        has_upper = False
        has_lower = False
        has_non_an = False

        For i = 1 To a.Length
            z = Mid$(a, i, 1)

            If UCase$(z) >= "A" And UCase$(z) <= "Z" Then
                has_alpha = True
            End If

            If z >= "0" And z <= "9" Then
                has_numeric = True
            End If

            If z >= "A" And z <= "Z" Then
                has_upper = True
            End If

            If z >= "a" And z <= "z" Then
                has_lower = True
            End If

            If (UCase$(z) < "A" Or UCase$(z) > "Z") And (z < "0" Or z > "9") Then
                has_non_an = True
            End If

            If b = "" Then
                b = z
            Else
                If z >= Mid(b, b.Length, 1) Then
                    b = b & z
                ElseIf z <= Mid(b, 1, 1) Then
                    b = z & b
                Else
                    For j = 1 To Len(b)
                        If z <= Mid$(b, j, 1) Then
                            b = Mid(b, 1, j - 1) & z & Mid$(b, j)
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i

        X_Sort = b
    End Function

    Public Function RndCrypt(ByVal Str As String, ByVal password As String) As String

        Dim SK As Long, k As Long

        ' init randomizer for password
        Rnd(-1)
        Randomize(Len(password))
        ' (((K Mod 256) Xor Asc(Mid$(Password, K, 1))) Xor Fix(256 * Rnd)) -> makes sure that a
        ' password like "pass12" does NOT give the same result as the password "sspa12" or "12pass"
        ' or "1pass2" etc. (or any combination of the same letters)

        For k = 1 To Len(password)
            SK = SK + (((k Mod 256) Xor ASC(Mid$(password, k, 1))) Xor Fix(256 * Rnd))
        Next k

        ' init randomizer for encryption/decryption
        Rnd(-1)
        Randomize(SK)

        ' encrypt/decrypt every character using the randomizer
        For k = 1 To Len(Str)
            Mid$(Str, k, 1) = Chr(Fix(256 * Rnd) Xor ASC(Mid$(Str, k, 1)))
        Next k

        RndCrypt = Str
    End Function

#End Region

    Private Sub ASFPWDC1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & ASCMAIN1.CLIENT_CODE & ".bmp")
        UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & "ABS" & ".bmp")
        txtUSER_ID.Text = ASCMAIN1.USER_ID
    End Sub
End Class
