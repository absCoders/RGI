Public Class ASCAPPT1

    ' Variables 

    ''' <summary>
    '''  Appointment Start Time 
    ''' </summary>
    ''' <remarks></remarks>
    Public StartTime As Date

    ''' <summary>
    ''' Appointment End Time 
    ''' </summary>
    ''' <remarks></remarks>
    Public EndTime As Date

    ''' <summary>
    '''  	Gets/sets the frequency at which the recurrence occurs: daily, weekly, monthly, or yearly.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternFrequency As Infragistics.Win.UltraWinSchedule.RecurrencePatternFrequency

    ''' <summary>
    ''' Gets/sets the limiting factor of the recurrence (i.e., whether it never ends, ends after a certain number of occurrences, or ends on or before a certain date).
    ''' </summary>
    ''' <remarks></remarks>
    Public RangeLimit As Infragistics.Win.UltraWinSchedule.RecurrenceRangeLimit

    ''' <summary>
    ''' Gets/sets the maximum number of occurrences for the recurrence. Applicable only when the RangeLimit property is set to "LimitByNumberOfOccurrences".
    ''' </summary>
    ''' <remarks></remarks>
    Public RangeMaxOccurrences As Int16

    ''' <summary>
    '''  Gets/sets the start time for each non-modified occurrence in the recurrence. 
    '''  The default value is DateTime.MinValue; unless specifically set, the actual value is obtained from the Appointment object referenced by the RootAppointment property.
    ''' </summary>
    ''' <remarks></remarks>
    Public RangeEndDate As Date

    ''' <summary>
    ''' Gets/sets whether the recurrence should occur on the first, second, third, fourth or last occurrence of the day of the week corresponding to the 
    ''' PatternDaysOfWeek property in its respective month. Applicable only when the PatternType property is set to 
    ''' "Calculated" (which is only applicable when the PatternFrequency property is set to "Monthly" or "Yearly"). 
    ''' The default value is None; unless specifically set, the actual value is obtained from the RangeStartDate property.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternOccurrenceOfDayInMonth As Infragistics.Win.UltraWinSchedule.RecurrencePatternOccurrenceOfDayInMonth

    ''' <summary>
    ''' Gets/sets the interval between occurrences of the recurrence. Works in conjunction with the PatternFrequency property. Not applicable when the PatternFrequency property is set to "Yearly".
    ''' Example: If the PatternFrequency property is set to "Weekly", and the PatternInterval property is set to 2, the appointment occurs every other week.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternInterval As Int16

    ''' <summary>
    '''  	Gets/sets the number of the day in its respective month on which each occurrence will occur. Applicable only when the PatternFrequency property is set to "Monthly" or "Yearly".
    ''' The default value is 0; unless specifically set, the actual value is obtained from the Appointment object referenced by the RootAppointment property.
    ''' Note: If the PatternDayOfMonth property is set to a value that exceeds the number of days in any of the months that the recurrence spans, the occurrence for those months will fall on the last day of that month.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternDayOfMonth As Int16

    ''' <summary>
    '''  Gets/sets whether the recurrence pattern is based on a specific day of the month and/or month of the year, or if it is calculated based on other criteria. 
    '''  Only applicable when the PatternFrequency property is set to "Monthly" or "Yearly". 
    '''  If the PatternFrequency property is set to "Monthly", and the PatternType property is set to "Explicit", the PatternDayOfMonth and PatternInterval 
    '''  properties are used to determine the recurrence pattern. If the PatternFrequency property is set to "Monthly", and the PatternType is set to "Calculated", 
    '''  the OccurrenceOfDayInMonth, PatternDaysOfWeek, and PatternInterval properties are used to determine the recurrence pattern. 
    '''  If the PatternFrequency property is set to "Yearly", and the PatternType is set to "Explicit", the recurrence occurs once per year, in the month 
    '''  specified by the PatternMonthOfYear property, and on the day specified by the PatternDayOfMonth property. If the PatternFrequency property is set to "Yearly", 
    '''  and the PatternType is set to "Calculated", the recurrence occurs once per year, on a day which is determined by a combination of the PatternOccurrenceOfDayInMonth, 
    '''  PatternDaysOfWeek, and PatternMonthOfYear properties.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternType As Infragistics.Win.UltraWinSchedule.RecurrencePatternType

    ''' <summary>
    ''' Gets/sets the day(s) of the week on which each occurrence occurs. The default value is None; unless specifically set, the actual value is obtained from the Appointment object referenced by the RootAppointment property.
    ''' Note: The DayOfWeek property is expressed as bit flags, so that multiple days can be represented by the property.
    ''' Example: To specify that a recurrence should occur on Tuesdays and Thursdays, assign a value of: (4 OR 16 ) = 20'''
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternDaysOfWeek As Infragistics.Win.UltraWinSchedule.RecurrencePatternDaysOfWeek

    ''' <summary>
    ''' Gets/sets the month of the year in which the recurrence will occur. Applicable only when the PatternFrequency property is set to "Yearly". The default value is based on the month of year that coincides with the RangeStartDate property.
    ''' </summary>
    ''' <remarks></remarks>
    Public PatternMonthOfYear As Int16

    ''' <summary>
    ''' Last Error generated bu the class
    ''' </summary>
    ''' <remarks></remarks>
    Public LastError As String

    ''' <summary>
    ''' Caption on Pop-Up Form
    ''' </summary>
    ''' <remarks></remarks>
    Public FormCaption As String

    ' Default to 5 years when 'No End Date' is selected on a Recurring Event
    ' This can be overwritten when Instantiating the Class
    Private NoEndDateDays As Int16 = 365 * 5

    Public Sub New()
        Initialize()
    End Sub

    Public Sub New(ByRef NoEndDateYears As Int16)
        Initialize()
        NoEndDateDays = 365 * NoEndDateYears
        If NoEndDateYears > 0 Then
            Try
                NoEndDateDays = 365 * NoEndDateYears
            Catch ex As Exception
                NoEndDateDays = 5000
            End Try
        End If
    End Sub

    Private Sub Initialize()
        StartTime = DateTime.Now
        EndTime = DateTime.Now
        PatternFrequency = UltraWinSchedule.RecurrencePatternFrequency.Daily
        RangeLimit = UltraWinSchedule.RecurrenceRangeLimit.NoLimit
        RangeMaxOccurrences = 10
        RangeEndDate = DateAdd(DateInterval.Day, 30, DateTime.Now)
        PatternOccurrenceOfDayInMonth = UltraWinSchedule.RecurrencePatternOccurrenceOfDayInMonth.First
        PatternInterval = 1
        PatternDayOfMonth = 1
        PatternType = UltraWinSchedule.RecurrencePatternType.Calculated
        PatternDaysOfWeek = UltraWinSchedule.RecurrencePatternDaysOfWeek.All
        PatternMonthOfYear = 1
        LastError = String.Empty
        FormCaption = "Appointment Recurrence"
    End Sub

    ''' <summary>
    ''' Displays the Appointment Recurrence Setup Screen and Updates the Properties
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DisplayAppointment() As Boolean

        Dim ListOfDates As New List(Of DateTime)
        DisplayAppointment(ListOfDates)

    End Function

    ''' <summary>
    ''' Displays the Appointment Recurrence Setup Screen and Updates the Properties
    ''' </summary>
    ''' <param name="ListOfDates">Returns the dates the Recurrence is to occur</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DisplayAppointment(ByRef ListOfDates As List(Of DateTime)) As Boolean

        Try

            LastError = String.Empty
            ' Standard Attributes
            Using calApptRecur As New Infragistics.Win.UltraWinSchedule.UltraCalendarInfo

                calApptRecur.AllowRecurringAppointments = True
                calApptRecur.MinDate = DateAdd(DateInterval.Day, -1, DateTime.Now)
                calApptRecur.MaxDate = DateAdd(DateInterval.Day, NoEndDateDays, DateTime.Now)

                Dim Appointment As New UltraWinSchedule.Appointment(StartTime, EndTime)
                Appointment.Recurrence = New UltraWinSchedule.AppointmentRecurrence

                Appointment.Recurrence.PatternFrequency = PatternFrequency
                Appointment.Recurrence.RangeLimit = RangeLimit
                Appointment.Recurrence.RangeMaxOccurrences = RangeMaxOccurrences
                Appointment.Recurrence.RangeEndDate = RangeEndDate

                Appointment.Recurrence.PatternOccurrenceOfDayInMonth = PatternOccurrenceOfDayInMonth
                Appointment.Recurrence.PatternInterval = PatternInterval
                Appointment.Recurrence.PatternDayOfMonth = PatternDayOfMonth
                Appointment.Recurrence.PatternType = PatternType
                Appointment.Recurrence.PatternDaysOfWeek = PatternDaysOfWeek
                Appointment.Recurrence.PatternMonthOfYear = PatternMonthOfYear

                Using RecurrenceDialog As New Infragistics.Win.UltraWinSchedule.RecurrenceDialog(Appointment, Appointment.Recurrence, True, False, False)
                    FormCaption = FormCaption.Trim
                    RecurrenceDialog.Text = FormCaption
                    RecurrenceDialog.ShowDialog()

                    Select Case RecurrenceDialog.Result
                        Case UltraWinSchedule.RecurrenceDialogResult.Cancel
                            DisplayAppointment = False

                        Case UltraWinSchedule.RecurrenceDialogResult.Ok
                            ' Update the datatable
                            StartTime = Appointment.StartDateTime
                            PatternFrequency = DirectCast(RecurrenceDialog.DateRecurrence.PatternFrequency, Integer)
                            RangeLimit = DirectCast(RecurrenceDialog.Recurrence.RangeLimit, Integer)
                            RangeMaxOccurrences = RecurrenceDialog.Recurrence.RangeMaxOccurrences
                            PatternInterval = RecurrenceDialog.DateRecurrence.PatternInterval
                            RangeEndDate = RecurrenceDialog.DateRecurrence.RangeEndDate
                            PatternDayOfMonth = RecurrenceDialog.DateRecurrence.PatternDayOfMonth
                            PatternDaysOfWeek = DirectCast(RecurrenceDialog.DateRecurrence.PatternDaysOfWeek, Integer)
                            PatternMonthOfYear = RecurrenceDialog.DateRecurrence.PatternMonthOfYear
                            PatternOccurrenceOfDayInMonth = DirectCast(RecurrenceDialog.DateRecurrence.PatternOccurrenceOfDayInMonth, Integer)
                            PatternType = DirectCast(RecurrenceDialog.DateRecurrence.PatternType, Integer)
                            DisplayAppointment = True

                            ' Update Appointment so we can Extract the Recurring Dates
                            Appointment.Recurrence.PatternFrequency = PatternFrequency
                            Appointment.Recurrence.RangeLimit = RangeLimit
                            Appointment.Recurrence.RangeMaxOccurrences = RangeMaxOccurrences
                            Appointment.Recurrence.PatternInterval = PatternInterval
                            Appointment.Recurrence.RangeEndDate = RangeEndDate

                            Appointment.Recurrence.PatternDayOfMonth = PatternDayOfMonth
                            Appointment.Recurrence.PatternDaysOfWeek = PatternDaysOfWeek
                            Appointment.Recurrence.PatternMonthOfYear = PatternMonthOfYear
                            Appointment.Recurrence.PatternOccurrenceOfDayInMonth = PatternOccurrenceOfDayInMonth
                            Appointment.Recurrence.PatternType = PatternType

                    End Select

                    calApptRecur.Appointments.Clear()
                    calApptRecur.Appointments.Add(Appointment)
                    Dim apps As UltraWinSchedule.AppointmentsSubsetCollection = calApptRecur.GetAppointmentsInRange(DateAdd(DateInterval.Day, -1, DateTime.Now), DateAdd(DateInterval.Day, NoEndDateDays, DateTime.Now))
                    For Each appt As UltraWinSchedule.Appointment In apps
                        If Not ListOfDates.Contains(appt.Start) Then
                            ListOfDates.Add(appt.Start)
                        End If
                    Next

                    RecurrenceDialog.Dispose()
                    Appointment.Dispose()
                End Using
                calApptRecur.Dispose()
            End Using

        Catch ex As Exception
            LastError = ex.Message
            Return False
        End Try

    End Function

End Class
