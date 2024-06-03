Imports System
Imports System.Collections.Generic

Imports ExchangeWeekOfMonth = Microsoft.Exchange.WebServices.Data.DayOfTheWeekIndex
Imports SchedulerWeekOfMonth = DevExpress.XtraScheduler.WeekOfMonth
Imports ExchangeWeekDay = Microsoft.Exchange.WebServices.Data.DayOfTheWeek
Imports SchedulerWeekDay = DevExpress.XtraScheduler.WeekDays

Imports Microsoft.Exchange.WebServices.Data

Namespace EWSSyncExample
	Public Module ExchangeSchedulerExtConverters
		<System.Runtime.CompilerServices.Extension> _
		Public Function ToSchedulerWeekOfMonth(ByVal weekIndex As ExchangeWeekOfMonth) As SchedulerWeekOfMonth
			Select Case weekIndex
				Case ExchangeWeekOfMonth.First
					Return SchedulerWeekOfMonth.First
				Case ExchangeWeekOfMonth.Second
					Return SchedulerWeekOfMonth.Second
				Case ExchangeWeekOfMonth.Third
					Return SchedulerWeekOfMonth.Third
				Case ExchangeWeekOfMonth.Fourth
					Return SchedulerWeekOfMonth.Fourth
				Case ExchangeWeekOfMonth.Last
					Return SchedulerWeekOfMonth.Last
			End Select
			Throw New Exception("Incorrect WeekIndex: " & weekIndex.ToString())
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Function ToExchangeWeekOfMonth(ByVal weekIndex As SchedulerWeekOfMonth) As ExchangeWeekOfMonth
			Select Case weekIndex
				Case SchedulerWeekOfMonth.First
					Return ExchangeWeekOfMonth.First
				Case SchedulerWeekOfMonth.Second
					Return ExchangeWeekOfMonth.Second
				Case SchedulerWeekOfMonth.Third
					Return ExchangeWeekOfMonth.Third
				Case SchedulerWeekOfMonth.Fourth
					Return ExchangeWeekOfMonth.Fourth
				Case SchedulerWeekOfMonth.Last
					Return ExchangeWeekOfMonth.Last
			End Select
			Throw New Exception("Incorrect WeekIndex: " & weekIndex.ToString())
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Function ToExchangeWeekDays(ByVal dw As SchedulerWeekDay) As ExchangeWeekDay()
			Dim result As New List(Of ExchangeWeekDay)()
			If (dw And SchedulerWeekDay.Sunday) <> 0 Then
				result.Add(ExchangeWeekDay.Sunday)
			End If
			If (dw And SchedulerWeekDay.Monday) <> 0 Then
				result.Add(ExchangeWeekDay.Monday)
			End If
			If (dw And SchedulerWeekDay.Tuesday) <> 0 Then
				result.Add(ExchangeWeekDay.Tuesday)
			End If
			If (dw And SchedulerWeekDay.Wednesday) <> 0 Then
				result.Add(ExchangeWeekDay.Wednesday)
			End If
			If (dw And SchedulerWeekDay.Thursday) <> 0 Then
				result.Add(ExchangeWeekDay.Thursday)
			End If
			If (dw And SchedulerWeekDay.Friday) <> 0 Then
				result.Add(ExchangeWeekDay.Friday)
			End If
			If (dw And SchedulerWeekDay.Saturday) <> 0 Then
				result.Add(ExchangeWeekDay.Saturday)
			End If
			If (dw And SchedulerWeekDay.EveryDay) <> 0 Then
				result.Add(ExchangeWeekDay.Day)
			End If
			If (dw And SchedulerWeekDay.WeekendDays) <> 0 Then
				result.Add(ExchangeWeekDay.WeekendDay)
			End If
			If (dw And SchedulerWeekDay.WorkDays) <> 0 Then
				result.Add(ExchangeWeekDay.Weekday)
			End If
			Return result.ToArray()
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Function ToSchedulerWeekDay(ByVal dw As ExchangeWeekDay) As SchedulerWeekDay
			Select Case dw
				Case ExchangeWeekDay.Sunday
					Return SchedulerWeekDay.Sunday
				Case ExchangeWeekDay.Monday
					Return SchedulerWeekDay.Monday
				Case ExchangeWeekDay.Tuesday
					Return SchedulerWeekDay.Tuesday
				Case ExchangeWeekDay.Wednesday
					Return SchedulerWeekDay.Wednesday
				Case ExchangeWeekDay.Thursday
					Return SchedulerWeekDay.Thursday
				Case ExchangeWeekDay.Friday
					Return SchedulerWeekDay.Friday
				Case ExchangeWeekDay.Saturday
					Return SchedulerWeekDay.Saturday
				Case ExchangeWeekDay.Day
					Return SchedulerWeekDay.EveryDay
				Case ExchangeWeekDay.WeekendDay
					Return SchedulerWeekDay.WeekendDays
				Case ExchangeWeekDay.Weekday
					Return SchedulerWeekDay.WorkDays
			End Select
			Throw New Exception("Incorrect day of the week: " & dw.ToString())
		End Function
		<System.Runtime.CompilerServices.Extension> _
		Public Function ToSchedulerWeekDays(ByVal dws As DayOfTheWeekCollection) As SchedulerWeekDay
			Dim result As SchedulerWeekDay = 0
			For Each dw As ExchangeWeekDay In dws
				result = result Or dw.ToSchedulerWeekDay()
			Next dw
			Return result
		End Function
	End Module
End Namespace
