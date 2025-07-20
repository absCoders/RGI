
Imports System.Text

Public Class TAFIMGV1
    Public _mode As String = "M" 'M = Main ABS, "L" = Laptop ABS
    Public _viewFTP As Boolean = False
    Private _FF As ASFBASE1
    Private _dst As New DataSet
    Private _ISLOADING As Boolean = True
    'Private _IMAGES_FOLDER_HIGH As String = ""
    'Private _IMAGES_FOLDER_LOW As String = ""
    Private _REMOTE_FOLDER As String = ""
    Private _LOCAL_FOLDER As String = ""
    Private _FILE_EXT As String = ".jpg"
    Private _STYLE_CODE As String = ""
    Private _COLOR_CODE As String = ""
    Private _CURR_IMAGE As String = ""
    Private _datICTIMAGT As New List(Of String)
    Private sql As New Text.StringBuilder With {.Length = 0}


#Region "Standard Methods"
    Public Sub New(ByVal FORM_BASE As ASFBASE1, ByVal STYLE_CODE As String, ByVal COLOR_CODE As String, ByVal mode As String, Optional ByVal viewFTP As Boolean = False)
        _FF = FORM_BASE
        _STYLE_CODE = STYLE_CODE
        _COLOR_CODE = COLOR_CODE
        _mode = mode
        _viewFTP = viewFTP
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        InitializeComponent()
        SetFileLocations()

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

        Dim FILES_LOW As String() = IO.Directory.GetFiles(_LOCAL_FOLDER, STYLE_SEARCH)

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

        _datICTIMAGT.Clear()

        _datICTIMAGT.Add("Low Res Only")

        _ISLOADING = False
        SetImage()
    End Sub

    Private Sub SetFileLocations()
        _REMOTE_FOLDER = "https://www.regency-rib.com/media/product/"
        _LOCAL_FOLDER = "S:\Images\"
        _FILE_EXT = "*.jpg"
        If ASCMAIN1.useUNCPath Then
            _LOCAL_FOLDER = $"{ASCMAIN1.Folders("SharedRoot")}\Images\"
        End If
        If Not (ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "whr" Or ASCMAIN1.USER_ID = "wayne")) Then
            Dim rowSOTPARM3 As DataRow = LookUp("SOTPARM3", "Z")
            If Not IsNothing(rowSOTPARM3) Then
                _LOCAL_FOLDER = rowSOTPARM3.Item("RO_PARM_STYLE_IMG_DIR").ToString & String.Empty
            End If
            AddSlash(_LOCAL_FOLDER)
        End If
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

    Private Sub SetImage()
        Dim SRC As String = "Local"
        imgSTYLE.ImageLocation = ""
        If _viewFTP Then
            Dim FileName As String = String.Format("{0}{1}-{2}.jpg", _REMOTE_FOLDER, _STYLE_CODE, _COLOR_CODE)
            imgSTYLE.ImageLocation = FileName
        Else
            Dim FileName As String = _LOCAL_FOLDER
            Dim SelFolder As String = FileName
            FileName = String.Format("{0}{1}-{2}", FileName, _STYLE_CODE, _COLOR_CODE)

            FileName = FileName & ".jpg"
            _CURR_IMAGE = FileName.Replace(SelFolder, "")

            If IO.File.Exists(FileName) Then
                imgSTYLE.ImageLocation = FileName
            Else
                MsgBox("Selected Image Not In File Sysytem")
                _CURR_IMAGE = ""
            End If
        End If
        Me.Text = $"Image Of {_STYLE_CODE}-{_COLOR_CODE} From {SRC}."
    End Sub
#End Region
End Class