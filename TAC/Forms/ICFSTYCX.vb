Public Class ICFSTYCX
    Public STYLE_CODE As String
    Public STYLE_DESC As String
    Public WHSE_CODE As String
    Public select_only As Boolean = False
    Public Price_Caption As String = ""
    Public PRICE As Decimal = 0

    Private Sub ICFSTYCX_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & ", ICTSTDQ2.DATE_1, ICTSTDQ2.QTY_1" _
                & ", ICTSTDQ2.DATE_2, ICTSTDQ2.QTY_2" _
                & ", ICTSTDQ2.DATE_3, ICTSTDQ2.QTY_3" _
                & ", ICTSTDQ2.DATE_4, ICTSTDQ2.QTY_4" _
                & " from ICTCOLR1,ICTSTYC1,ICTSTDQ2" _
                & " where ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" _
                & "   and ICTSTYC1.STYLE_CODE = :PARM1" _
                & "   and ICTSTDQ2.WHSE_CODE (+) = '" & Me.WHSE_CODE & "'" _
                & "   and ICTSTDQ2.STYLE_CODE (+) = ICTSTYC1.STYLE_CODE" _
                & "   and ICTSTDQ2.COLOR_CODE (+) = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRM", "**", 0, False, "V", 1)
            .Tables("ICTCOLRM").Columns.Add("QTY", GetType(System.Int32))
            .Tables("ICTCOLRM").Columns.Add("SEL", GetType(System.String))
        End With

        grdICTCOLRM.DataSource = dst.Tables("ICTCOLRM")
        grdICTCOLRM.DisplayLayout.Bands(0).Columns("QTY").Hidden = select_only
        grdICTCOLRM.DisplayLayout.Bands(0).Columns("SEL").Hidden = Not select_only

        Create_Summary(grdICTCOLRM, "COLOR_CODE", "Count")
        Create_Summary(grdICTCOLRM, New String() {"QTY"})

        With grdICTCOLRM.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With
        With grdICTCOLRM.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "QTY" Or gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For I As Integer = 1 To 4
                .Columns("DATE_" & CStr(I)).Format = "MM/dd"
            Next
        End With

        lblPrice.Text = Price_Caption
        lblPrice.Visible = Price_Caption <> "" And Not select_only
        numPrice.Visible = Price_Caption <> "" And Not select_only

        If STYLE_CODE <> "" Then
            Absx1.txtFor("STYLE_CODE").Text = STYLE_CODE
            Prepare_ICTCOLRM()
        End If

        If WHSE_CODE = "" Then
            For I As Integer = 1 To 4
                grdICTCOLRM.DisplayLayout.Bands(0).Columns("DATE_" & CStr(I)).Hidden = True
                grdICTCOLRM.DisplayLayout.Bands(0).Columns("QTY_" & CStr(I)).Hidden = True
            Next
            Me.Width = Me.Width - (8 * 65)
        End If

    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STYLE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Prepare_ICTCOLRM()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"
                Prepare_ICTCOLRM()
        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click
        STYLE_CODE = Absx1.txtFor("STYLE_CODE").Text
        Dim row As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
        If row Is Nothing Then
            MsgBox("Invalid Style Code Specified", MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If
        PRICE = Val(numPrice.Value & "")
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        STYLE_CODE = ""
        Me.Close()
    End Sub

    Sub Prepare_ICTCOLRM()
        Dim STYLE_CODE As String = Absx1.txtFor("STYLE_CODE").Text
        Dim rowICTSTYL1 As DataRow = Nothing
        If STYLE_CODE <> "" Then rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE)
        If rowICTSTYL1 IsNot Nothing Then
            Fill_Records("ICTCOLRM", STYLE_CODE)
            Sort_grdColumns(grdICTCOLRM, "COLOR_CODE", True)
            grdICTCOLRM.Text = rowICTSTYL1.Item("STYLE_DESC") & ""
            Absx1.numFor("CARTON_PACK_QTY").Value = rowICTSTYL1.Item("CARTON_PACK_QTY")
            Absx1.numFor("SUB_UNIT_PACK_QTY").Value = rowICTSTYL1.Item("SUB_UNIT_PACK_QTY")
        Else
            grdICTCOLRM.Text = ""
            Absx1.numFor("CARTON_PACK_QTY").Value = 0
            Absx1.numFor("SUB_UNIT_PACK_QTY").Value = 0
            dst.Tables("ICTCOLRM").Rows.Clear()
        End If

    End Sub
End Class