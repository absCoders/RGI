Imports System.Text.RegularExpressions
Public Class SOFMLAB1
    Public ORDR_NOs As New List(Of String)
    Public ORDR_GROUP_NOs As New List(Of String)
    Public CARTONS_MAX As Int32 = 100
    Public UNITS_MAX As Int32 = 1000
    Public printed As Boolean = False
    Public CUST_CODE As String
    Public CUST_NAME As String
    Public ORDR_GROUP_NO As String


    Private Sub SOFBINV1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst


            ASCMAIN1.sql = "Select " & vbCrLf _
                & "SOTORDR1.ORDR_NO," & vbCrLf _
                & "SOTORDR1.ORDR_GROUP_NO," & vbCrLf _
                & "SOTORDR1.ORDR_CUST_PO," & vbCrLf _
                & "SOTORDR1.ORDR_SHIP_DATE," & vbCrLf _
                & "SOTORDR1.ORDR_CANCEL_DATE," & vbCrLf _
                & "SOTORDR1.CUST_DC_NO," & vbCrLf _
                & "SUBSTR(SOTORDR1.CUST_DC_NO,1,5) CUST_DC_NO5 ," & vbCrLf _
                & "SOTORDR1.ORDR_DEPT," & vbCrLf _
                & "SOTORDR1.WHSE_CODE," & vbCrLf _
                & "SOTORDR5.CUST_NAME," & vbCrLf _
                & "SOTORDR5.CUST_ADDR1," & vbCrLf _
                & "SOTORDR5.CUST_ADDR2," & vbCrLf _
                & "SOTORDR5.CUST_ADDR3," & vbCrLf _
                & "SOTORDR5.CUST_CITY," & vbCrLf _
                & "SOTORDR5.CUST_STATE," & vbCrLf _
                & "SOTORDR5.CUST_ZIP_CODE" & vbCrLf _
                & " from SOTORDR1,SOTORDR5 where SOTORDR1.ORDR_GROUP_NO = :PARN1" & vbCrLf _
                & " and SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & " and SOTORDR5.CUST_ADDR_TYPE = 'ST'"
            Create_TDA(.Tables.Add, "SOTMLAB1", "**", 0, False, "V", 1)

            With .Tables("SOTMLAB1")
                .Columns.Add("CARTONS", GetType(System.Int32))
                .Columns.Add("UNITS", GetType(System.Int32))
            End With

            .Tables.Add("SOTMLAB2")
            With .Tables("SOTMLAB2")
                .Columns.Add("ORDR_NO")
                .Columns.Add("CARTON_NO", GetType(System.Int32))
                .Columns.Add("UNITS", GetType(System.Int32))
                .PrimaryKey = New DataColumn() {.Columns("ORDR_NO"), .Columns("CARTON_NO")}
            End With

        End With

        grdSOTMLAB1.DataSource = dst.Tables("SOTMLAB1")
        grdSOTMLAB2.DataSource = dst.Tables("SOTMLAB2")

        With grdSOTMLAB1.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.False
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
        End With

        With grdSOTMLAB2.DisplayLayout.Bands(0)
            .Override.AllowUpdate = DefaultableBoolean.True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "UNITS" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        Create_Summary(grdSOTMLAB1, "ORDR_NO", "Count")
        Create_Summary(grdSOTMLAB2, "CARTON_NO", "Count")
        Create_Summary(grdSOTMLAB2, "UNITS")

        Create_New_Batch()

        Dim ZebraPrinters As New List(Of String)
        ASCMAIN1.sql = "Select * from ASTPRNT1"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim PRINTER_CODE As String = row.Item("PRINTER_CODE")
            Dim PRINTER_NAME As String = row.Item("PRINTER_NAME")
            Dim PRINTER_PORT As String = row.Item("PRINTER_PORT")

            Dim ZebraPrinter As String = PRINTER_CODE & "|" & PRINTER_NAME & "|" & PRINTER_PORT
            ZebraPrinters.Add(ZebraPrinter)
        Next
        cboZebraPrinter.DataSource = ZebraPrinters


    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "STYLE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then

            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STYLE_CODE"

        End Select
    End Sub
#End Region

    Private Sub cmdPrint_Click(sender As System.Object, e As System.EventArgs) Handles cmdPrint.Click
        EMsg = ""

        If dst.Tables("SOTMLAB2").Rows.Count = 0 Then
            EMsg &= vbCr & "No Labels Created"
        Else
            For Each rowSOTMLAB1 As DataRow In dst.Tables("SOTMLAB1").Select("")
                Dim ORDR_NO As String = rowSOTMLAB1.Item("ORDR_NO")
                If dst.Tables("SOTMLAB2").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                    EMsg &= vbCr & "No Labels Defined for Order Group " & ORDR_NO
                End If
            Next
        End If

        If dst.Tables("SOTMLAB2").Select("ISNULL(UNITS,0) < 1 OR ISNULL(UNITS,0) > " & CStr(UNITS_MAX)).Length <> 0 Then
            EMsg &= vbCr & "Some Labels have an invalid number of Units"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If


        ' PRINT THE CARTON LABELS


        Dim PrinterName As String = ""

        If ASCMAIN1.CLIENT = "VAN" Then
            Dim ZebraPrinter As String = cboZebraPrinter.SelectedValue
            Dim PRINTER_PORT As String = ZebraPrinter.Split("|")(2)
            PrinterName = PRINTER_PORT
        End If

        Dim LABEL_CODE As String = "KOHLSMCL"
        Dim cartonLabel As New TestLabel(LABEL_CODE, "")


        'Dim labelTemplate As String = "" '  cartonLabel.GetLabelTemplate()

        Dim labelTemplate As String = ASCDATA1.GetDataValue( _
            "SELECT UCC128_COMMANDS FROM " & _
            " SOTUCCL1 U1 " & _
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {LABEL_CODE}) & ""
        If labelTemplate = "" Then Throw New Exception("Label Template '" & LABEL_CODE & "' not found")

        For Each rowSOTMLAB1 As DataRow In dst.Tables("SOTMLAB1").Select("")
            Dim ORDR_NO As String = rowSOTMLAB1.Item("ORDR_NO")
            Dim WHSE_CODE As String = rowSOTMLAB1.Item("WHSE_CODE")
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            For Each rowSOTMLAB2 As DataRow In dst.Tables("SOTMLAB2").Select("ORDR_NO = '" & ORDR_NO & "'")

                Dim labelData As New Dictionary(Of String, DataRow)
                labelData.Add("SOTMLAB1", rowSOTMLAB1)
                labelData.Add("SOTMLAB2", rowSOTMLAB2)
                labelData.Add("ICTWHSE1", rowICTWHSE1)


                Dim labeltoPrint As String = FillLabelTemplateWithData(labelTemplate, labelData)
                ShippingLabel.SendToLabelPrinter(labeltoPrint, PrinterName)

                '  cartonLabel.PrintLabel(1, PrinterName)
            Next
        Next

        'BeginTrans()
        'Update_Record_TDA("SOTBINV1")
        'Update_Record_TDA("SOTBINV2", "BATCH_NO = '" & BATCH_NO & "'")
        'CommitTrans()

        printed = True

        Me.Close()
    End Sub


    Private Function FillLabelTemplateWithData(labelTemplate As String, labelData As Dictionary(Of String, DataRow)) As String
        'Matches <<TABLE.COLUMN>...>, and if the value of TABLE.COLUMN is null, it omits this line from the ZPL
        'Used for hiding a section of label if the data is unavailable
        labelTemplate = Regex.Replace(labelTemplate, "\<\<(?<table>[\w_]+)\.(?<column>[\w_]+)\>(?<command>.*)\>", _
                        Function(m) If(labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "" = "",
                                       "", m.Groups("command").Value))


        'Regex matches {TABLE.COLUMN} and replaces with values from rowUCC128
        labelTemplate = Regex.Replace(labelTemplate, "\{(?<table>[\w_]+)\.(?<column>[\w_]+)\}", _
                        Function(m) labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "")


        'If labelTemplate.Contains("{CARTONDETAILS}") Then
        '    labelTemplate = WriteCartonDetails(labelTemplate, labelData("SOTCART1").Item("CART_NO") & String.Empty)
        'End If

        Return labelTemplate
    End Function

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Sub Create_New_Batch()
        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            'ASCMAIN1.sql = "Select ORDR_GROUP_NO, ORDR_CUST_PO" _
            '    & " from SOTORDR0 where SOTORDR0.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            'Fill_Records("SOTMLAB1", , False, ASCMAIN1.sql)
            Fill_Records("SOTMLAB1", ORDR_GROUP_NO, False)
        Next


        Absx1.txtFor("CUST_CODE").Text = CUST_CODE
        Absx1.txtFor("CUST_NAME").Text = CUST_NAME
        Absx1.txtFor("ORDR_GROUP_NO").Text = ORDR_GROUP_NO
    End Sub

    Private Sub cmdMakeCartons_Click(sender As Object, e As EventArgs) Handles cmdMakeCartons.Click

        If grdSOTMLAB1.ActiveRow Is Nothing Then Exit Sub

        '    Dim ORDR_GROUP_NO As String = grdSOTMLAB1.ActiveRow.Cells("ORDR_GROUP_NO").Value
        Dim ORDR_NO As String = grdSOTMLAB1.ActiveRow.Cells("ORDR_NO").Value
        Dim CARTONS As Int32 = Val(numCARTONS.Value & "")
        Dim UNITS As Int32 = Val(numUNITS.Value & "")

        If CARTONS < 1 Or CARTONS > CARTONS_MAX Then
            MsgBox("Invalid Value specified for number of Cartons", MsgBoxStyle.OkOnly, "Cannot Produce Master Carton Labels")
            Exit Sub
        End If

        If UNITS < 1 Or UNITS > UNITS_MAX Then
            MsgBox("Invalid Value specified for number of Units per Carton", MsgBoxStyle.OkOnly, "Cannot Produce Master Carton Labels")
            Exit Sub
        End If

        Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "'"
        If dst.Tables("SOTMLAB2").Select(sqlw).Length > 0 Then
            If MsgBox("OK to Delete & Re-Create Cartons for Order No " & ORDR_NO, MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub
        End If

        '    dst.Tables("SOTMLAB2").Rows.Clear()

        ASCDATA1.DeleteRows(dst.Tables("SOTMLAB2"), sqlw)

        For CARTON_NO As Int32 = 1 To CARTONS
            Dim rowSOTLABM2 As DataRow = dst.Tables("SOTMLAB2").NewRow
            With rowSOTLABM2
                .Item("ORDR_NO") = ORDR_NO
                .Item("CARTON_NO") = CARTON_NO
                .Item("UNITS") = UNITS
            End With
            dst.Tables("SOTMLAB2").Rows.Add(rowSOTLABM2)
        Next

        Dim rowSOTMLAB1 As DataRow = dst.Tables("SOTMLAB1").Rows.Find(ORDR_NO)
        rowSOTMLAB1.Item("CARTONS") = CARTONS
        rowSOTMLAB1.Item("UNITS") = UNITS

    End Sub

    Private Sub grdSOTMLAB1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTMLAB1.AfterRowActivate
        Dim ORDR_NO As String = grdSOTMLAB1.ActiveRow.Cells("ORDR_NO").Value
        grdSOTMLAB2.Text = "Cartons for Order No " & ORDR_NO

        numCARTONS.Value = Val(grdSOTMLAB1.ActiveRow.Cells("CARTONS").Value & "")
        numUNITS.Value = Val(grdSOTMLAB1.ActiveRow.Cells("UNITS").Value & "")

        Dim dvw As DataView = DirectCast(grdSOTMLAB2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "ORDR_NO = '" & ORDR_NO & "'"
        Sort_grdColumns(grdSOTMLAB2, "CARTON_NO")

    End Sub

    Private Sub UltraLabel3_Click(sender As Object, e As EventArgs) Handles UltraLabel3.Click

    End Sub

    Private Sub grdSOTMLAB2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTMLAB2.InitializeLayout

    End Sub
End Class