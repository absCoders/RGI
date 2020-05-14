Imports ABSolution
Imports System.Text.RegularExpressions

Public Class WHFLABEL


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sql As String = String.Empty

        With dst
            .Tables.Add("LABEL")

            With .Tables("LABEL")
                .Columns.Add("Sequence", GetType(System.Int16))
                .Columns.Add("Caption", GetType(System.String))

                .Columns.Add("Value", GetType(System.String))
                .Columns("Value").MaxLength = 40

                .Columns.Add("FieldName", GetType(System.String))

            End With

        End With

        grdLabel.DataSource = dst.Tables("LABEL")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Clear"
                If MessageBox.Show("Do you want to Clear the Values provided for the label?", "Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Print"

                grdLabel.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                If Not Regex.IsMatch(dst.Tables("LABEL").Select("FieldName = 'FIELD5'")(0).Item("Value") & String.Empty, "^\d{4,4}$") Then
                    EMsg &= "PM (Pack Method) is required and must be 4 digits" & vbCrLf
                End If

                If Not Regex.IsMatch(dst.Tables("LABEL").Select("FieldName = 'FIELD13'")(0).Item("Value") & String.Empty, "^\d\d$") Then
                    EMsg &= "Label Code is required and must be 2 digits" & vbCrLf
                End If



                Dim numCartons As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD2'")(0).Item("Value") & String.Empty)
                Dim numSets As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD100'")(0).Item("Value") & String.Empty)

                If numCartons <= 1 Then numCartons = 1
                If numSets <= 1 Then numSets = 1

                If EMsg = "" AndAlso MessageBox.Show("Do you want to print " & numSets & " set(s) for " & numCartons & " carton(s)?", "Print", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If


        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Clear"
                Clear_Record()

            Case "Print"
                Printlabel()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Print").Settings.Enabled = DefaultableBoolean.True
                .Groups("Screen Control").Items("Clear").Settings.Enabled = DefaultableBoolean.True
            End With
        End If

        'MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If


    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        With dst.Tables("LABEL")
            .Rows.Clear()
            .Rows.Add(New Object() {5, "Number of Sets", "1", "FIELD100"})
            .Rows.Add(New Object() {10, "Number of Cartons", "1", "FIELD2"})
            .Rows.Add(New Object() {11, "Start Carton", "", "FIELD98"})
            .Rows.Add(New Object() {12, "End Carton", "", "FIELD99"})
            .Rows.Add(New Object() {30, "P.O", "", "FIELD1"})
            .Rows.Add(New Object() {35, "Label", "", "FIELD15"})
            .Rows.Add(New Object() {40, "Item Code", "", "FIELD3"})
            .Rows.Add(New Object() {45, "Style", "", "FIELD14"})
            .Rows.Add(New Object() {60, "Color", "", "FIELD4"})
            '.Rows.Add(New Object() {80, "SSA", "", "FIELD6"})
            .Rows.Add(New Object() {90, "Ratio", "", "FIELD7"})
            .Rows.Add(New Object() {91, "Size", "", "FIELD8"})
            .Rows.Add(New Object() {92, "Pre-Pack #", "", "FIELD9"})
            .Rows.Add(New Object() {110, "Total Pcs.", "", "FIELD10"})
            .Rows.Add(New Object() {115, "PM", "", "FIELD5"})
            .Rows.Add(New Object() {120, "Description", "", "FIELD11"})
            .Rows.Add(New Object() {130, "Made In", "", "FIELD12"})
            .Rows.Add(New Object() {140, "Label Code", "", "FIELD13"})
        End With

        Sort_grdColumns(grdLabel, "Sequence", True)

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("")

        EnforceConstraints(False)

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Private Sub Printlabel()

        Try
            Dim customLabel As New CustomLabel("NYA_F21", "")

            Dim numCartons As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD2'")(0).Item("Value") & String.Empty)
            Dim numSets As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD100'")(0).Item("Value") & String.Empty)

            Dim startCarton As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD98'")(0).Item("Value") & String.Empty)
            If startCarton = 0 Then startCarton = 1

            Dim endCarton As Integer = Val(dst.Tables("LABEL").Select("FieldName = 'FIELD99'")(0).Item("Value") & String.Empty)
            If endCarton = 0 Then endCarton = numCartons


            If numCartons <= 1 Then numCartons = 1
            If numSets <= 1 Then numSets = 1

            For cartCount As Int16 = startCarton To endCarton
                Dim cartNo As String = ASCMAIN1.Next_Control_No("SOTCART1.CART_NO")

                customLabel.tblLabelData.Rows.Clear()
                Dim rowLabelData As DataRow = customLabel.tblLabelData.NewRow
                With rowLabelData
                    For Each rowLabel As DataRow In dst.Tables("LABEL").Select("")
                        .Item(rowLabel.Item("Fieldname")) = rowLabel.Item("Value") & String.Empty
                    Next
                End With

                rowLabelData.Item("FIELD100") = cartCount & " of " & numCartons

                Dim cartonBarCode As String = dst.Tables("LABEL").Select("FieldName = 'FIELD1'")(0).Item("Value") & String.Empty 'PO
                'cartonBarCode &= dst.Tables("LABEL").Select("FieldName = 'FIELD3'")(0).Item("Value") & String.Empty ' Item Code
                cartonBarCode &= dst.Tables("LABEL").Select("FieldName = 'FIELD5'")(0).Item("Value") & String.Empty ' Pack Method
                'cartonBarCode &= dst.Tables("LABEL").Select("FieldName = 'FIELD6'")(0).Item("Value") & String.Empty ' SSA
                cartonBarCode &= dst.Tables("LABEL").Select("FieldName = 'FIELD13'")(0).Item("Value") & String.Empty ' Label Code
                cartonBarCode &= StrReverse(StrReverse(cartNo).Substring(0, 5)) ' last 5 of a carton NO
                rowLabelData.Item("FIELD13") = cartonBarCode

                customLabel.tblLabelData.Rows.Add(rowLabelData)
                Try
                    customLabel.PrintLabel(numSets)
                Catch ex As Exception
                    MessageBox.Show("The following error occurred: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message)
        End Try

    End Sub

#End Region


    Private Sub grdLabel_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdLabel.InitializeRow

        If e.Row.Cells("Sequence").Value <= 15 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
        End If
    End Sub
End Class