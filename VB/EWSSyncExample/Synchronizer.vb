Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.Exchange.WebServices.Data

Imports ExchangeAppointment = Microsoft.Exchange.WebServices.Data.Appointment
'INSTANT VB NOTE: VB does not allow aliasing interfaces:
'using SchedulerAppointment = DevExpress.XtraScheduler.Appointment
Imports ExchangeWeekOfMonth = Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex
Imports SchedulerWeekOfMonth = DevExpress.XtraScheduler.WeekOfMonth
Imports ExchangeWeekDay = Microsoft.Exchange.WebServices.Data.DayOfTheWeek
Imports SchedulerWeekDay = DevExpress.XtraScheduler.WeekDays
Imports Recurrence = Microsoft.Exchange.WebServices.Data.Recurrence

Namespace EWSSyncExample
	Public Class Synchronizer
		Protected service As ExchangeService

		Public Sub New(ByVal service As ExchangeService)
			Me.service = service
		End Sub

		Private recurrentPropertySet As PropertySet
		Private normalPropertySet As PropertySet

		Protected Function GetPropertySet(ByVal recurrent As Boolean) As PropertySet
			If recurrent Then
				If recurrentPropertySet Is Nothing Then
					recurrentPropertySet = New PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType, AppointmentSchema.Subject, AppointmentSchema.TextBody, AppointmentSchema.FirstOccurrence, AppointmentSchema.LastOccurrence, AppointmentSchema.ModifiedOccurrences, AppointmentSchema.DeletedOccurrences, AppointmentSchema.Recurrence, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.ReminderMinutesBeforeStart, AppointmentSchema.IsAllDayEvent, AppointmentSchema.IsReminderSet, AppointmentSchema.IsCancelled)
				End If
				Return recurrentPropertySet
			Else
				If normalPropertySet Is Nothing Then
					normalPropertySet = New PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType, AppointmentSchema.Subject, AppointmentSchema.TextBody, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.ReminderMinutesBeforeStart, AppointmentSchema.IsAllDayEvent, AppointmentSchema.IsReminderSet, AppointmentSchema.IsCancelled)
				End If
				Return normalPropertySet
			End If
		End Function

		Public Function GetExchangeAppointments(ByVal startSearchDate As Date, ByVal endSearchDate As Date) As List(Of ExchangeAppointment)
			Dim res As New List(Of ExchangeAppointment)()
			Dim recurrentIds As New HashSet(Of String)()
			Dim calView As New CalendarView(startSearchDate, endSearchDate)

			calView.PropertySet = New PropertySet(BasePropertySet.IdOnly, AppointmentSchema.AppointmentType, AppointmentSchema.IsCancelled, AppointmentSchema.ICalUid)

			Dim findResults As FindItemsResults(Of ExchangeAppointment) = service.FindAppointments(WellKnownFolderName.Calendar, calView)
			Dim singleAptProps = GetPropertySet(False)

			For Each appt As ExchangeAppointment In findResults.Items
				If appt.AppointmentType = AppointmentType.Occurrence OrElse appt.AppointmentType = AppointmentType.Exception Then
					If recurrentIds.Add(appt.ICalUid) Then
						res.Add(GetRecurrentTemplate(appt.Id))
						Continue For
					End If
				End If

				If appt.AppointmentType = AppointmentType.Single Then
					res.Add(ExchangeAppointment.Bind(service, appt.Id, singleAptProps))
				End If
			Next appt

			Return res.Where(Function(n) Not n.IsCancelled).ToList()
		End Function

		Protected Function GetRecurrentTemplate(ByVal itemId As ItemId) As ExchangeAppointment
			Dim calendarItem As ExchangeAppointment = Appointment.Bind(service, itemId, New PropertySet(AppointmentSchema.AppointmentType))
			Dim props As PropertySet = GetPropertySet(True)
			Select Case calendarItem.AppointmentType
				Case AppointmentType.RecurringMaster
					Return ExchangeAppointment.Bind(service, itemId, props)
				Case AppointmentType.Occurrence, AppointmentType.Exception
					Return ExchangeAppointment.BindToRecurringMaster(service, itemId, props)
				Case Else
					Return Nothing
			End Select
		End Function

		Private Sub ConvertRecurrenceMSToDX(ByVal exchangeAppointment As ExchangeAppointment, ByVal schedulerAppointment As DevExpress.XtraScheduler.Appointment)

			'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
			'ORIGINAL LINE: if(exchangeAppointment.Recurrence is Microsoft.Exchange.WebServices.Data.Recurrence.YearlyPattern yearlyPattern)
			If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.YearlyPattern) Then
				Dim yearlyPattern As Recurrence.YearlyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.YearlyPattern)
				schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Yearly
				schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0
				schedulerAppointment.RecurrenceInfo.DayNumber = yearlyPattern.DayOfMonth
				schedulerAppointment.RecurrenceInfo.Month = CInt(Math.Truncate(yearlyPattern.Month))
			End If

			'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
			'ORIGINAL LINE: if(exchangeAppointment.Recurrence is Microsoft.Exchange.WebServices.Data.Recurrence.RelativeYearlyPattern relativeYearlyPattern)

			If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.RelativeYearlyPattern) Then
				Dim relativeYearlyPattern As Recurrence.RelativeYearlyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.RelativeYearlyPattern)
				schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Yearly
					schedulerAppointment.RecurrenceInfo.WeekOfMonth = relativeYearlyPattern.DayOfTheWeekIndex.ToSchedulerWeekOfMonth()
					schedulerAppointment.RecurrenceInfo.WeekDays = relativeYearlyPattern.DayOfTheWeek.ToSchedulerWeekDay()
					schedulerAppointment.RecurrenceInfo.Month = CInt(Math.Truncate(relativeYearlyPattern.Month))
				End If

			'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
			'ORIGINAL LINE: if(exchangeAppointment.Recurrence is Microsoft.Exchange.WebServices.Data.Recurrence.MonthlyPattern monthlyPattern)

			If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.MonthlyPattern) Then
					Dim monthlyPattern As Recurrence.MonthlyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.MonthlyPattern)
					schedulerAppointment.RecurrenceInfo.Periodicity = monthlyPattern.Interval
					schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Monthly
					schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0
					schedulerAppointment.RecurrenceInfo.DayNumber = monthlyPattern.DayOfMonth
				End If

			'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
			'ORIGINAL LINE: if(exchangeAppointment.Recurrence is Microsoft.Exchange.WebServices.Data.Recurrence.RelativeMonthlyPattern relativeMonthlyPattern)
			If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.RelativeMonthlyPattern) Then
					Dim relativeMonthlyPattern As Recurrence.RelativeMonthlyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.RelativeMonthlyPattern)
					schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Monthly
					schedulerAppointment.RecurrenceInfo.WeekOfMonth = relativeMonthlyPattern.DayOfTheWeekIndex.ToSchedulerWeekOfMonth()
					schedulerAppointment.RecurrenceInfo.WeekDays = relativeMonthlyPattern.DayOfTheWeek.ToSchedulerWeekDay()
				End If

			'INSTANT VB TODO TASK: VB has no equivalent to C# pattern variables in 'is' expressions:
			'ORIGINAL LINE: if(exchangeAppointment.Recurrence is Microsoft.Exchange.WebServices.Data.Recurrence.WeeklyPattern weeklyPattern)
			If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.WeeklyPattern) Then
				Dim weeklyPattern As Recurrence.WeeklyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.WeeklyPattern)
				If exchangeAppointment.Recurrence.GetType() = GetType(Recurrence.MonthlyPattern) Then
					Dim monthlyPattern As Recurrence.MonthlyPattern = TryCast(exchangeAppointment.Recurrence, Recurrence.MonthlyPattern)
					schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Weekly
					schedulerAppointment.RecurrenceInfo.WeekDays = weeklyPattern.DaysOfTheWeek.ToSchedulerWeekDays()
				End If
			End If

			If TypeOf exchangeAppointment.Recurrence Is Recurrence.DailyPattern Then
				schedulerAppointment.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily
			End If

			If TypeOf exchangeAppointment.Recurrence Is Recurrence.IntervalPattern Then
				schedulerAppointment.RecurrenceInfo.Periodicity = (TryCast(exchangeAppointment.Recurrence, Recurrence.IntervalPattern)).Interval
			End If

			schedulerAppointment.RecurrenceInfo.Start = exchangeAppointment.Recurrence.StartDate.Add(exchangeAppointment.Start.TimeOfDay)

			If exchangeAppointment.Recurrence.HasEnd Then
				If exchangeAppointment.Recurrence.NumberOfOccurrences.HasValue Then
					schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.OccurrenceCount
					schedulerAppointment.RecurrenceInfo.OccurrenceCount = exchangeAppointment.Recurrence.NumberOfOccurrences.Value
				End If
				If exchangeAppointment.Recurrence.EndDate.HasValue Then
					schedulerAppointment.RecurrenceInfo.End = exchangeAppointment.Recurrence.EndDate.Value
					schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.EndByDate
				End If
			Else
				schedulerAppointment.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.NoEndDate
			End If

			If (exchangeAppointment.DeletedOccurrences?.Count > 0) OrElse (exchangeAppointment.ModifiedOccurrences?.Count > 0) Then
				Dim calculator = DevExpress.XtraScheduler.OccurrenceCalculator.CreateInstance(schedulerAppointment.RecurrenceInfo)

				If exchangeAppointment.DeletedOccurrences?.Count > 0 Then
					For Each deleted As DeletedOccurrenceInfo In exchangeAppointment.DeletedOccurrences
						Dim occurrenceIndex As Integer = calculator.FindOccurrenceIndex(deleted.OriginalStart, schedulerAppointment)
						schedulerAppointment.CreateException(DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence, occurrenceIndex)
					Next deleted
				End If

				If exchangeAppointment.ModifiedOccurrences?.Count > 0 Then
					For Each modified As OccurrenceInfo In exchangeAppointment.ModifiedOccurrences
						Dim occurrenceIndex As Integer = calculator.FindOccurrenceIndex(modified.OriginalStart, schedulerAppointment)
						Dim modifiedSchedulerAppointment As DevExpress.XtraScheduler.Appointment = schedulerAppointment.CreateException(DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence, occurrenceIndex)
						modifiedSchedulerAppointment.Start = modified.Start
						modifiedSchedulerAppointment.End = modified.End
					Next modified
				End If
			End If
		End Sub

		Private Sub ConvertRecurrenceDXtoMS(ByVal schedulerAppointment As DevExpress.XtraScheduler.Appointment, ByVal exchangeAppointment As ExchangeAppointment)

			Select Case schedulerAppointment.RecurrenceInfo.Type
				Case DevExpress.XtraScheduler.RecurrenceType.Yearly
					If schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0 Then
						Dim pattern As New Recurrence.YearlyPattern()
						pattern.DayOfMonth = schedulerAppointment.RecurrenceInfo.DayNumber
						pattern.Month = CType(schedulerAppointment.RecurrenceInfo.Month, Month)
						exchangeAppointment.Recurrence = pattern
					Else
						Dim pattern As New Recurrence.RelativeYearlyPattern()
						pattern.DayOfTheWeekIndex = schedulerAppointment.RecurrenceInfo.WeekOfMonth.ToExchangeWeekOfMonth()
						pattern.DayOfTheWeek = schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays().First()
						pattern.Month = CType(schedulerAppointment.RecurrenceInfo.Month, Month)
						exchangeAppointment.Recurrence = pattern
					End If
				Case DevExpress.XtraScheduler.RecurrenceType.Monthly
					If schedulerAppointment.RecurrenceInfo.WeekOfMonth = 0 Then
						Dim pattern As New Recurrence.MonthlyPattern()
						pattern.DayOfMonth = schedulerAppointment.RecurrenceInfo.DayNumber
						exchangeAppointment.Recurrence = pattern
					Else
						Dim pattern As New Recurrence.RelativeMonthlyPattern()
						pattern.DayOfTheWeekIndex = schedulerAppointment.RecurrenceInfo.WeekOfMonth.ToExchangeWeekOfMonth()
						pattern.DayOfTheWeek = schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays().First()
						exchangeAppointment.Recurrence = pattern
					End If
				Case DevExpress.XtraScheduler.RecurrenceType.Weekly
					exchangeAppointment.Recurrence = New Recurrence.WeeklyPattern(schedulerAppointment.RecurrenceInfo.Start, schedulerAppointment.RecurrenceInfo.Periodicity, schedulerAppointment.RecurrenceInfo.WeekDays.ToExchangeWeekDays())
				Case DevExpress.XtraScheduler.RecurrenceType.Daily
					exchangeAppointment.Recurrence = New Recurrence.DailyPattern(schedulerAppointment.RecurrenceInfo.Start, schedulerAppointment.RecurrenceInfo.Periodicity)
				Case Else
					Throw New NotImplementedException("Recurrence type " & schedulerAppointment.RecurrenceInfo.Type & " is not supported.")
			End Select

			If TypeOf exchangeAppointment.Recurrence Is Recurrence.IntervalPattern Then
				TryCast(exchangeAppointment.Recurrence, Recurrence.IntervalPattern).Interval = schedulerAppointment.RecurrenceInfo.Periodicity
			End If

			exchangeAppointment.Recurrence.StartDate = schedulerAppointment.RecurrenceInfo.Start.Date

			Select Case schedulerAppointment.RecurrenceInfo.Range
				Case DevExpress.XtraScheduler.RecurrenceRange.OccurrenceCount
					exchangeAppointment.Recurrence.NumberOfOccurrences = schedulerAppointment.RecurrenceInfo.OccurrenceCount
				Case DevExpress.XtraScheduler.RecurrenceRange.EndByDate
					exchangeAppointment.Recurrence.EndDate = schedulerAppointment.RecurrenceInfo.End
				Case DevExpress.XtraScheduler.RecurrenceRange.NoEndDate
				Case Else
					Throw New NotImplementedException("Recurrence range type " & schedulerAppointment.RecurrenceInfo.Range.ToString() & " is not supported.")
			End Select
		End Sub

		Public Sub SchedulerToExchange(ByVal schedulerAppointment As DevExpress.XtraScheduler.Appointment, Optional ByVal addReminders As Boolean = False)
			If (schedulerAppointment.Type <> DevExpress.XtraScheduler.AppointmentType.Pattern) AndAlso (schedulerAppointment.Type <> DevExpress.XtraScheduler.AppointmentType.Normal) Then
				Throw New Exception("Appointment type " & schedulerAppointment.Type & " is not support")
			End If

			Dim exchangeAppointment As New ExchangeAppointment(service)
			exchangeAppointment.Subject = schedulerAppointment.Subject
			exchangeAppointment.Body = schedulerAppointment.Description
			exchangeAppointment.Start = schedulerAppointment.Start
			exchangeAppointment.End = schedulerAppointment.End

			If (addReminders) AndAlso (schedulerAppointment.HasReminder) Then
				exchangeAppointment.ReminderMinutesBeforeStart = schedulerAppointment.Reminder.TimeBeforeStart.Minutes
			End If

			If schedulerAppointment.Type = DevExpress.XtraScheduler.AppointmentType.Pattern Then
				ConvertRecurrenceDXtoMS(schedulerAppointment, exchangeAppointment)
			End If

			exchangeAppointment.Save(SendInvitationsMode.SendToAllAndSaveCopy)

			If schedulerAppointment.HasExceptions Then
				For Each exception As DevExpress.XtraScheduler.Appointment In schedulerAppointment.GetExceptions()
					If exception.Type = DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence Then
						Dim appointment As ExchangeAppointment = ExchangeAppointment.BindToOccurrence(service, exchangeAppointment.Id, exception.RecurrenceIndex + 1, New PropertySet())
						appointment.Delete(DeleteMode.MoveToDeletedItems)
						Continue For
					End If

					If exception.Type = DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence Then
						Dim appointment As ExchangeAppointment = ExchangeAppointment.BindToOccurrence(service, exchangeAppointment.Id, exception.RecurrenceIndex + 1, New PropertySet(AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.Subject, AppointmentSchema.OriginalStart))
						appointment.Start = exception.Start
						appointment.End = exception.End
						appointment.Subject = exception.Subject
						appointment.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy)
						Continue For
					End If
				Next exception
			End If
		End Sub

		Public Sub ExchangeToScheduler(ByVal appointmentStorage As DevExpress.XtraScheduler.IAppointmentStorageBase, ByVal exchangeAppointment As ExchangeAppointment, Optional ByVal addReminders As Boolean = False)

			Dim schedulerAppointment As DevExpress.XtraScheduler.Appointment = Nothing
			If exchangeAppointment.AppointmentType = AppointmentType.Single Then
				schedulerAppointment = appointmentStorage.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Normal)
			Else
				schedulerAppointment = appointmentStorage.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Pattern)
			End If

			schedulerAppointment.AllDay = exchangeAppointment.IsAllDayEvent
			schedulerAppointment.Start = exchangeAppointment.Start
			schedulerAppointment.End = exchangeAppointment.End
			schedulerAppointment.Subject = exchangeAppointment.Subject
			schedulerAppointment.Description = exchangeAppointment.TextBody

			If exchangeAppointment.AppointmentType <> AppointmentType.Single Then
				ConvertRecurrenceMSToDX(exchangeAppointment, schedulerAppointment)
			End If

			If addReminders Then
				Dim reminder As DevExpress.XtraScheduler.Reminder = schedulerAppointment.CreateNewReminder()
				reminder.TimeBeforeStart = TimeSpan.FromMinutes(exchangeAppointment.ReminderMinutesBeforeStart)
				schedulerAppointment.Reminders.Add(reminder)
			End If

			appointmentStorage.Add(schedulerAppointment)
		End Sub
	End Class
End Namespace
