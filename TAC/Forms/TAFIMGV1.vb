
Imports System.Text

Public Class TAFIMGV1
    Public mode As String = "" ' N = New, "" = Update Next Step
    Private _IC_PARM_STYLE_IMG_DIR As String
    Private _FF As ASFBASE1
    Private _dst As New DataSet
    Private _ISLIVE As Boolean = False
    Private _HiRezDir As String
    Private _LowRezDir As String
#Region "Standard Methods"
    Public Sub New(ByVal FF As ASFBASE1, ByVal STYLE_CODE As String)
        frmASFBASE1 = FF
        GatherDataRequired()
        InitializeVariables()
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'imgSTYLE.ImageLocation = ImageLocation
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
        sql.AppendLine("FROM ICTPARM1")
        sql.AppendLine("WHERE IC_PARM_KEY = 'Z'")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString()))

        sql.Length = 0
        sql.AppendLine("SELECT *")
        sql.AppendLine("FROM ICTIMAGT")
        _dst.Tables.Add(ASCDATA1.GetDataTable(sql.ToString()))
    End Sub
    Private Sub InitializeVariables()
        If _dst.Tables.Item("ICTPARM1").Rows.Count = 1 Then
            _IC_PARM_STYLE_IMG_DIR = _dst.Tables.Item("ICTPARM1").Rows(0).Item("_IC_PARM_STYLE_IMG_DIR").ToString & String.Empty
            If _ISLIVE Then
                _HiRezDir = "HiRez\All images\"
                _LowRezDir = "LowRez\Master\"
                If _IC_PARM_STYLE_IMG_DIR.Length = 0 Then
                    RaiseError("Invalid Image Directory In Paramters")
                End If
            Else
                _IC_PARM_STYLE_IMG_DIR = "\\192.168.110.236\Media\Pictures\ABS\"
                _HiRezDir = "HiRez\All images\"
                _LowRezDir = "LowRez\Master\"
            End If
            AddSlash(_IC_PARM_STYLE_IMG_DIR)
            AddSlash(_HiRezDir)
            AddSlash(_LowRezDir)
        Else
            RaiseError("Invalid Parameter In IC")
        End If
    End Sub

    Private Sub RaiseError(ByVal eMsg As String)
        MsgBox(eMsg.ToString(), MsgBoxStyle.Exclamation, "Error On Form")
        Me.Close()
    End Sub
#End Region
End Class