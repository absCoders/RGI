
Imports Infragistics.Win.UltraWinGrid

Public Class SOFORDRB
    Public ORDR_BUYER_NAME As String = ""
    Public ORDR_BUYER_EMAIL As String = ""
    Public ORDR_BUYER_CONTACT_NO As Integer
    Public HAS_CONTACT_CHANGES As Boolean = False
    Private FF As ASFBASE1
    Private CUST_CODE As String

#Region "Standard Methods"
    Public Sub New(ByVal frmASFBASE1 As ASFBASE1, ByVal in_CUST_CODE As String)
        FF = frmASFBASE1
        CUST_CODE = in_CUST_CODE
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CUST_CODE.Length = 0 Then
            Me.Close()
        End If
        ASCMAIN1.Add_Value_List(grdARTCUSTD, "CONTACT_TYPE")
        FF.Fill_Records("ARTCUSTD", CUST_CODE)
        grdARTCUSTD.DataSource = FF.dst.Tables("ARTCUSTD")
        With grdARTCUSTD.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.True
        End With
    End Sub

#End Region

#Region "Form Controls"
    Private Sub btnCancel_Click_1(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim Emsg As New Text.StringBuilder With {.Length = 0}
        Emsg.AppendLine("Cancelling Will Discard Any ")
        Emsg.AppendLine("Additions Or Changes Made")
        Emsg.AppendLine("To Contacts!")
        Emsg.AppendLine("")
        Emsg.AppendLine("Is That OK With You?")
        Dim iResult As MsgBoxResult = MsgBox(Emsg.ToString, vbYesNo, "Discarding?")
        If iResult = vbYes Then
            FF.dst.Tables.Item("ARTCUSTD").Clear()
            ORDR_BUYER_NAME = ""
            ORDR_BUYER_EMAIL = ""
            ORDR_BUYER_CONTACT_NO = Nothing
            Me.Close()
        End If
    End Sub

    Private Sub btnDone_Click(sender As Object, e As EventArgs) Handles btnDone.Click
        Dim EMsg As String = ""
        For Each rowARTCUSTD As DataRow In FF.dst.Tables("ARTCUSTD").Select()
            If Not rowARTCUSTD.RowState = DataRowState.Unchanged Then
                HAS_CONTACT_CHANGES = True
            End If
        Next
        For Each ROW As DataRow In ASCDATA1.SelectDistinct(FF.dst.Tables("ARTCUSTD"), New String() {"CONTACT_TYPE"}).Rows

            Dim CONTACT_TYPE As String = ROW.Item("CONTACT_TYPE") & ""
            Dim sqlw As String = "CONTACT_TYPE = '" & CONTACT_TYPE & "'"
            Dim c As Integer = Val(FF.dst.Tables("ARTCUSTD").Compute("COUNT(CONTACT_NO)", sqlw & " and CONTACT_PRIMARY = '1'") & "")
            If c > 1 Then
                EMsg &= vbCr & "Cannot Have > 1 Primary Contact Of Any Type (see Type " & CONTACT_TYPE & ")"
            ElseIf c = 0 Then
                Dim rows() As DataRow = FF.dst.Tables("ARTCUSTD").Select(sqlw)
                If rows.Length = 1 Then
                    rows(0).Item("CONTACT_PRIMARY") = "1"
                Else
                    EMsg &= vbCr & "You Must Select A Primary Contact For Each Type Of Contact (See Type " & CONTACT_TYPE & ")"
                End If
            End If
        Next
        If EMsg.Length = 0 Then
            If grdARTCUSTD.Selected.Rows.Count = 1 Then
                ORDR_BUYER_NAME = grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_NAME").Text
                ORDR_BUYER_EMAIL = grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_EMAIL").Text
                ORDR_BUYER_CONTACT_NO = Val(grdARTCUSTD.Selected.Rows(0).Cells.Item("CONTACT_NO").Text)
            Else
                EMsg &= vbCr & "You Must Select One Row From The Contacts As The Buyer."
            End If
        End If

        If EMsg.Length > 0 Then
            MsgBox(EMsg.ToString, vbOKOnly, "Contacts")
        Else
            Me.Close()
        End If
    End Sub
#End Region

#Region "grdARTCUSTD"
    Private Sub grdARTCUSTD_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdARTCUSTD.AfterRowInsert
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ISNULL(CONTACT_NOTE,'NULL') <> 'DELETED'"
    End Sub

    Private Sub grdARTCUSTD_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdARTCUSTD.BeforeRowsDeleted
        For Each rowARTCUSTD As UltraGridRow In e.Rows
            rowARTCUSTD.Cells.Item("CONTACT_NOTE").Value = "DELETED"
        Next
        Dim dvw As DataView = DirectCast(grdARTCUSTD.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ISNULL(CONTACT_NOTE,'NULL') <> 'DELETED'"
        e.Cancel = True
    End Sub

    Private Sub grdARTCUSTD_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdARTCUSTD.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = CUST_CODE
            e.Row.Cells("CONTACT_NO").Value = Val(FF.dst.Tables("ARTCUSTD").Compute("MAX(CONTACT_NO)", "") & "") + 1
        End If
    End Sub

    Private Sub grdARTCUSTD_Error(sender As Object, e As ErrorEventArgs) Handles grdARTCUSTD.[Error]
        grdARTCUSTD.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub


#End Region
End Class