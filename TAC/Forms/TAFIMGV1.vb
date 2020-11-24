
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
#Region "Standard Methods"
    Public Sub New(ByVal FORM_BASE As ASFBASE1, ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal mode As String)
        _FF = FORM_BASE
        _STYLE_CODE = STYLE_CODE
        _COLOR_CODE = COLOR_CODE
        If mode = "L" Then
            _mode = "L"
            Stop 'Not supported yet.
        End If
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitializeComponent()
        GatherDataRequired()
        InitializeVariables()
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
    Private Sub GatherDataRequired()
        Dim sql As New Text.StringBuilder With {.Length = 0}
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTPARMI")
        sql.AppendLine("WHERE IC_PARM_KEY = 'Z'")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTPARMI"))

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTIMAGT")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString(), "ICTIMAGT"))
    End Sub
    Private Sub InitializeVariables()
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
        _datICTIMAGT.Clear()
        For Each rowICTIMAGT As DataRow In _dst.Tables("ICTIMAGT").Select("", "IMAGE_DEFAULT DESC ,IMAGE_CODE")
            _datICTIMAGT.Add(rowICTIMAGT.Item("IMAGE_DESC").ToString & String.Empty)
        Next
        cboICTIMAGT.DataSource = _datICTIMAGT
        cboICTIMAGT.SelectedIndex = 0
    End Sub

    Private Sub RaiseError(ByVal eMsg As String)
        MsgBox(eMsg.ToString(), MsgBoxStyle.Exclamation, "Error On Form")
        Me.Close()
    End Sub

    Private Sub cboICTIMAGT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboICTIMAGT.SelectedIndexChanged
        SetImage()
    End Sub

    Private Sub SetImage()
        If Not _ISLOADING Then
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
        End If
    End Sub

    Private Sub rdoRezL_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRezL.CheckedChanged
        SetImage()
    End Sub

    Private Sub rdoRezH_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRezH.CheckedChanged
        SetImage()
    End Sub

    Private Sub cmdSave_Click(sender As Object, e As EventArgs) Handles cmdSave.Click
        If imgSTYLE.ImageLocation.ToString.Length > 0 And _CURR_IMAGE.Length > 0 Then
            Dim fList As New List(Of String)
            Dim fDialog As New FolderBrowserDialog
            Dim dInfo As IO.DirectoryInfo
            fDialog.Description = "Please Select The Folder To Save To."
            fDialog.ShowDialog()
            imgSTYLE.Image.Save(fDialog.SelectedPath & "\" & _CURR_IMAGE)
        End If
    End Sub
#End Region
End Class