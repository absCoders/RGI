Imports System.IO
Imports System.Drawing


Public Class WHFWHSEC

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Get_PARM("POTPARM1")
        With dst
            ASCMAIN1.sql = " Select * from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, True, "V", 2)
        End With
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Update"

            Case "Cancel"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Cancel"

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Visible = ScreenMode
            End With
        End If

        If ScreenMode Then
        Else
            tabMISC.Visible = False
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False


        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")


        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub


#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "MEMBER_SS_NO"
        End Select
    End Sub

#End Region

#Region "Form Routines"


#End Region


    Private Sub cmdCreate_Locations_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCreate_Locations.Click
        dst.Tables("WHTLOCM1").Clear()

        Create_Bays(5, 200)
        For i As Integer = 10 To 36
            Select Case i
                Case 10, 11, 12, 13
                    'Create_Bays(i, 13) was told todouble the bay count
                    Create_Bays(i, 26)
                Case 14
                    'Create_Bays(i, 9)  was told todouble the bay count
                    Create_Bays(i, 18)
                Case Else
                    'Create_Bays(i, 11)  was told todouble the bay count
                    Create_Bays(i, 22)
            End Select
        Next

        Update_Record_TDA("WHTLOCM1", "Delete from WHTLOCM1")
        MsgBox("Complete", MsgBoxStyle.OkOnly, "Success")
    End Sub

    Sub Create_Bays(ByVal Aisle_No As Integer, ByVal bay_Count As Integer)
        For i As Integer = 1 To bay_Count
            For Each Level As String In New String() {"A", "B", "C", "D"}
                Dim rowWHTLOCM1 As DataRow = dst.Tables("WHTLOCM1").NewRow
                With rowWHTLOCM1
                    .Item("WHSE_CODE") = "NJE"
                    .Item("LOCATION_CODE") = Format$(Aisle_No, "00") & Format(i, "-000-") & Level 'Format$(Aisle_No, "00") & Format(i, "000") & Level
                    .Item("LOCATION_DESC") = "" 'Format$(Aisle_No, "00") & Format(i, "-000-") & Level
                    .Item("PHY_TYPE") = "S"
                    .Item("LOCATION_SINGLE_LOAD") = "1"
                End With
                dst.Tables("WHTLOCM1").Rows.Add(rowWHTLOCM1)
            Next
        Next
    End Sub

End Class

