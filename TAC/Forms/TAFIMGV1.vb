
Imports System.Text

Public Class TAFIMGV1
    Public _mode As String = "M" 'M = Main ABS, "L" = Laptop ABS
    Private _FF As ASFBASE1
    Private _dst As New DataSet
    Private _ISLOADING As Boolean = True
    Private _IMAGES_FOLDER_HIGH As String = ""
    Private _IMAGES_FOLDER_LOW As String = ""
    Private _STYLE_CODE As String = ""
    Private _COLOR_CODE As String = ""
    Private _CURR_IMAGE As String = ""
    Private _datICTIMAGT As New List(Of String)
    Private sql As New Text.StringBuilder With {.Length = 0}
#Region "Standard Methods"
    Public Sub New(ByVal FORM_BASE As ASFBASE1, ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal mode As String)
        _FF = FORM_BASE
        _STYLE_CODE = STYLE_CODE
        _COLOR_CODE = COLOR_CODE
        _mode = mode
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitializeComponent()

        Dim STYLE_SEARCH As String = String.Format("{0}-{1}*.*", _STYLE_CODE, _COLOR_CODE)

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTPARMI")
        sql.AppendLine("WHERE IC_PARM_KEY = 'Z'")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTPARMI"))

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM SOTPARM3")
        sql.AppendLine("WHERE RO_PARM_KEY = 'Z'")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "SOTPARM3"))

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTIMAGT")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTIMAGT"))

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM WBTIMGL1")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "WBTIMGL1"))

        If _mode = "L" Then
            If _dst.Tables.Item("SOTPARM3").Rows.Count = 1 Then
                _IMAGES_FOLDER_HIGH = "C:\Regency"
                _IMAGES_FOLDER_LOW = _dst.Tables.Item("SOTPARM3").Rows(0).Item("RO_PARM_STYLE_IMG_DIR").ToString & String.Empty

                If _IMAGES_FOLDER_HIGH.Length = 0 Or _IMAGES_FOLDER_LOW.Length = 0 Then
                    RaiseError("Invalid Image Directory In Paramters")
                Else
                    AddSlash(_IMAGES_FOLDER_HIGH)
                    AddSlash(_IMAGES_FOLDER_LOW)
                End If
            Else
                RaiseError("Invalid Parameter In SO3")
            End If
        Else
            If _dst.Tables.Item("ICTPARMI").Rows.Count = 1 Then
                If (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
                    _IMAGES_FOLDER_HIGH = "S:\RGI\Images\High\"
                    _IMAGES_FOLDER_LOW = "S:\RGI\Images\Low\"
                Else
                    _IMAGES_FOLDER_HIGH = _dst.Tables.Item("ICTPARMI").Rows(0).Item("IMAGES_FOLDER_HIGH").ToString & String.Empty
                    _IMAGES_FOLDER_LOW = _dst.Tables.Item("ICTPARMI").Rows(0).Item("IMAGES_FOLDER_LOW").ToString & String.Empty
                End If

                If _IMAGES_FOLDER_HIGH.Length = 0 Or _IMAGES_FOLDER_LOW.Length = 0 Then
                    RaiseError("Invalid Image Directory In Paramters")
                Else
                    AddSlash(_IMAGES_FOLDER_HIGH)
                    AddSlash(_IMAGES_FOLDER_LOW)
                End If
            Else
                RaiseError("Invalid Parameter In IC")
            End If
        End If

        Dim FILES_HIGH As String() = IO.Directory.GetFiles(_IMAGES_FOLDER_HIGH, STYLE_SEARCH)
        Dim FILES_LOW As String() = IO.Directory.GetFiles(_IMAGES_FOLDER_LOW, STYLE_SEARCH)

        For Each FILENAME As String In FILES_LOW
            FILENAME = FILENAME.ToUpper
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim IMAGE_SUFFIX As String = ""
            TAC.TACMAIN1.PARSE_IMAGE(FILENAME, STYLE_CODE, COLOR_CODE, IMAGE_SUFFIX)
            If STYLE_CODE.Length > 0 And COLOR_CODE.Length > 0 Then
                Dim rowWBTIMGL1 As DataRow = _dst.Tables.Item("WBTIMGL1").NewRow
                rowWBTIMGL1.Item("FILE_NAME") = FILENAME
                rowWBTIMGL1.Item("FILE_SOURCE") = "L"
                rowWBTIMGL1.Item("MATCHED") = "1"
                rowWBTIMGL1.Item("STYLE_CODE") = STYLE_CODE
                rowWBTIMGL1.Item("COLOR_CODE") = COLOR_CODE
                rowWBTIMGL1.Item("IMAGE_SUFFIX") = IMAGE_SUFFIX
                _dst.Tables.Item("WBTIMGL1").Rows.Add(rowWBTIMGL1)
            End If
        Next

        For Each FILENAME As String In FILES_HIGH
            FILENAME = FILENAME.ToUpper
            Dim STYLE_CODE As String = ""
            Dim COLOR_CODE As String = ""
            Dim IMAGE_SUFFIX As String = ""
            TAC.TACMAIN1.PARSE_IMAGE(FILENAME, STYLE_CODE, COLOR_CODE, IMAGE_SUFFIX)
            If STYLE_CODE.Length > 0 And COLOR_CODE.Length > 0 Then
                Dim rowWBTIMGL1 As DataRow = _dst.Tables.Item("WBTIMGL1").NewRow
                rowWBTIMGL1.Item("FILE_NAME") = FILENAME
                rowWBTIMGL1.Item("FILE_SOURCE") = "H"
                rowWBTIMGL1.Item("MATCHED") = "1"
                rowWBTIMGL1.Item("STYLE_CODE") = STYLE_CODE
                rowWBTIMGL1.Item("COLOR_CODE") = COLOR_CODE
                rowWBTIMGL1.Item("IMAGE_SUFFIX") = IMAGE_SUFFIX
                _dst.Tables.Item("WBTIMGL1").Rows.Add(rowWBTIMGL1)
            End If
        Next

        _datICTIMAGT.Clear()

        If _dst.Tables.Item("WBTIMGL1").Select("FILE_SOURCE = 'H'").Count = 0 Then
            rdoRezL.Checked = True
            rdoRezH.Enabled = False
            _datICTIMAGT.Add("Low Res Only")
            cboICTIMAGT.Enabled = False
        Else
            rdoRezH.Checked = True
            _datICTIMAGT.Add(_dst.Tables("ICTIMAGT").Select("IMAGE_DEFAULT = '1'").FirstOrDefault.Item("IMAGE_DESC").ToString & String.Empty)
            For Each rowICTIMAGT As DataRow In _dst.Tables("ICTIMAGT").Select("IMAGE_DEFAULT <> '1'", "IMAGE_DEFAULT DESC , IMAGE_CODE")
                Dim IMAGE_SUFFIX As String = rowICTIMAGT.Item("IMAGE_SUFFIX").ToString & String.Empty
                If IMAGE_SUFFIX.Length > 0 Then
                    If _dst.Tables.Item("WBTIMGL1").Select(String.Format("IMAGE_SUFFIX = '{0}'", IMAGE_SUFFIX)).Count > 0 Then
                        _datICTIMAGT.Add(rowICTIMAGT.Item("IMAGE_DESC").ToString & String.Empty)
                    End If
                End If
            Next
        End If

        cboICTIMAGT.DataSource = _datICTIMAGT
        cboICTIMAGT.SelectedIndex = 0

        _ISLOADING = False
        SetImage()
    End Sub
#End Region

#Region "Form Controls"
    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        Me.Close()
    End Sub
#End Region

#Region "Custom Methods"

    Private Sub AddSlash(ByRef Slasher As String)
        If Not Slasher.EndsWith("\") Then
            Slasher = Slasher & "\"
        End If
    End Sub


    Private Sub RaiseError(ByVal eMsg As String)
        MsgBox(eMsg.ToString(), MsgBoxStyle.Exclamation, "Error On Form")
        Me.Close()
    End Sub

    Private Sub cboICTIMAGT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboICTIMAGT.SelectedIndexChanged
        If Not _ISLOADING Then
            SetImage()
        End If
    End Sub

    Private Sub SetImage()

        imgSTYLE.ImageLocation = ""

        Dim FileName As String = _IMAGES_FOLDER_HIGH
        If rdoRezL.Checked Then
            FileName = _IMAGES_FOLDER_LOW
        End If
        Dim SelFolder As String = FileName
        FileName = String.Format("{0}{1}-{2}", FileName, _STYLE_CODE, _COLOR_CODE)

        Dim SFilter As String = String.Format("IMAGE_DESC = '{0}'", cboICTIMAGT.Text)

        Dim IMAGE_SUFFIX As String = ""
        Dim rowICTIMAGT As DataRow = _dst.Tables("ICTIMAGT").Select(SFilter).FirstOrDefault
        If Not IsNothing(rowICTIMAGT) Then
            IMAGE_SUFFIX = rowICTIMAGT.Item("IMAGE_SUFFIX").ToString & String.Empty
        End If
        If IMAGE_SUFFIX.Length > 0 Then
            FileName = String.Format("{0}-{1}", FileName, IMAGE_SUFFIX)
        End If
        FileName = FileName & ".jpg"
        _CURR_IMAGE = FileName.Replace(SelFolder, "")

        If IO.File.Exists(FileName) Then
            imgSTYLE.ImageLocation = FileName
        Else
            If rdoRezH.Checked Then
                cboICTIMAGT.SelectedIndex = 0
                rdoRezL.Checked = True
            Else
                MsgBox("Selected Image Not In File Sysytem")
                _CURR_IMAGE = ""
            End If
        End If
    End Sub

    Private Sub rdoRezL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRezL.CheckedChanged
        If Not _ISLOADING Then
            cboICTIMAGT.SelectedIndex = 0
            cboICTIMAGT.Enabled = False
            SetImage()
        End If
    End Sub

    Private Sub rdoRezH_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRezH.CheckedChanged
        If Not _ISLOADING Then
            cboICTIMAGT.SelectedIndex = 0
            cboICTIMAGT.Enabled = True
            SetImage()
        End If
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        If imgSTYLE.ImageLocation.ToString.Length > 0 And _CURR_IMAGE.Length > 0 Then
            Dim fList As New List(Of String)
            Dim fDialog As New FolderBrowserDialog
            fDialog.Description = "Please Select The Folder To Save To."
            fDialog.ShowDialog()
            If fDialog.SelectedPath.Length > 0 Then
                imgSTYLE.Image.Save(fDialog.SelectedPath & "\" & _CURR_IMAGE)
            End If
        End If
    End Sub
#End Region
End Class