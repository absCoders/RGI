Public Class POTTRFF1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        dteTARIFF_DATE_START.MinDate = CDate("01/01/2024")

        dteINVOICE_START_DATE.MinDate = CDate("01/01/2024")
        dteINVOICE_ENDING_DATE.MinDate = CDate("01/01/2024")

        dteORDER_START_DATE.MinDate = CDate("01/01/2024")
        dteORDER_ENDING_DATE.MinDate = CDate("01/01/2024")

        dteTARIFF_DATE_START.MaxDate = DateAdd(DateInterval.Year, 2, DateTime.Now)

        dteINVOICE_START_DATE.MaxDate = DateAdd(DateInterval.Year, 2, DateTime.Now)
        dteINVOICE_ENDING_DATE.MaxDate = DateAdd(DateInterval.Year, 2, DateTime.Now)

        dteORDER_START_DATE.MaxDate = DateAdd(DateInterval.Year, 2, DateTime.Now)
        dteORDER_ENDING_DATE.MaxDate = DateAdd(DateInterval.Year, 2, DateTime.Now)

        With dst
            Create_TDA(.Tables.Add, "SOTMISC1", "*")
            Fill_Records("SOTMISC1", String.Empty, True, "SELECT * FROM SOTMISC1")
        End With

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

    End Sub

    Overrides Sub Show_Record_Special()
    End Sub

    Overrides Sub Clear_Record_Special()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                Dim TARIFF_PERC As Double = Val(Absx1.numFor("TARIFF_PERC").Value & String.Empty)
                Dim SURCHARGE_PERC As Double = Val(Absx1.numFor("SURCHARGE_PERC").Value & String.Empty)

                ' Validate Purchasing options only when checked
                If Absx1.chkFor("TARIFF_ACTIVE").Checked Then
                    If TARIFF_PERC > 20 Then
                        If MessageBox.Show("Tariff Percent is greater than 20%. Do you want to Proceed?", "Update ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            EMsg &= vbCr & "User cancelled Update"
                            Exit Sub
                        End If
                    End If

                    If optTARIFF_DATE_FIELD.CheckedIndex = -1 Then
                        EMsg &= vbCr & "You must select a 'Date to use' in the purchasiing section."
                        Exit Sub
                    End If
                End If

                ' Validate Sales Order options only when checked
                If Absx1.chkFor("SURCHARGE_ACTIVE").Checked Then

                    If IsDate(dteORDER_START_DATE.Value) Then
                        If IsDate(dteORDER_ENDING_DATE.Value) Then
                            Select Case DateTime.Compare(dteORDER_START_DATE.Value, dteORDER_ENDING_DATE.Value)
                                Case 0
                                ' Dates arfe equal
                                Case -1
                                ' End date is after Start date
                                Case 1
                                    ' End date is before Start Date
                                    EMsg &= vbCr & "Order Start Date must be before Order End Date"
                            End Select
                        End If
                    Else
                        If IsDate(dteORDER_ENDING_DATE.Value) Then
                            EMsg &= vbCr & "Order Start Date is required when providing an Order End Date"
                        End If
                    End If

                    If IsDate(dteINVOICE_START_DATE.Value) Then
                        If IsDate(dteINVOICE_ENDING_DATE.Value) Then
                            Select Case DateTime.Compare(dteINVOICE_START_DATE.Value, dteINVOICE_ENDING_DATE.Value)
                                Case 0
                                ' Dates arfe equal
                                Case -1
                                ' End date is after Start date
                                Case 1
                                    ' End date is before Start Date
                                    EMsg &= vbCr & "Invoice Start Date must be before Invoice End Date"
                            End Select
                        End If
                    Else
                        If IsDate(dteINVOICE_ENDING_DATE.Value) Then
                            EMsg &= vbCr & "Invoice Start Date is required when providing an Invoice End Date"
                        End If
                    End If

                    Dim MISC_CHG_CODE As String = Absx1.txtFor("MISC_CHG_CODE").Text
                    Dim rowSOTMISC1 As DataRow = dst.Tables("SOTMISC1").Rows.Find(MISC_CHG_CODE)
                    If rowSOTMISC1 Is Nothing Then
                        EMsg &= vbCr & "Invalid or missing Misc Charge Code"
                    End If

                    If EMsg.Length = 0 Then
                        If SURCHARGE_PERC > 20 Then
                            If MessageBox.Show("Surcharge Percent is greater than 20%. Do you want to Proceed?", "Update ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                EMsg &= vbCr & "User cancelled Update"
                                Exit Sub
                            End If
                        End If

                        If SURCHARGE_PERC > TARIFF_PERC Then
                            If MessageBox.Show("Surcharge Percent is greater than the Tariff Percent. Do you want to Proceed?", "Update ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                EMsg &= vbCr & "User cancelled Update"
                                Exit Sub
                            End If
                        End If
                    End If
                End If
        End Select

    End Sub
#End Region

End Class