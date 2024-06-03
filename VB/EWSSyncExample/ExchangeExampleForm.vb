Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports DevExpress.XtraEditors
Imports Microsoft.Exchange.WebServices.Data

Namespace EWSSyncExample
	Partial Public Class ExchangeExampleForm
		Inherits XtraForm

		Private synchronizer As Synchronizer

		Public Sub New()
			InitializeComponent()
			schedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Month

			Dim now As Date = Date.Now.Date.AddHours(18)
			' Create a recurring Appointment.
			Dim apt = schedulerDataStorage1.Appointments.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Pattern, now, now.AddHours(2), "TEST_TEST")

			apt.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily
			apt.RecurrenceInfo.Periodicity = 1
			apt.RecurrenceInfo.Start = now
			apt.RecurrenceInfo.End = now.AddDays(6)
			apt.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.EndByDate
			apt.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily
			schedulerDataStorage1.Appointments.Add(apt)
			apt.CreateException(DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence, 4)
			Dim apt2 = apt.CreateException(DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence, 3)
			apt2.Start = now.AddDays(3).AddHours(2)

			' Create a regular Appointment
			apt = schedulerDataStorage1.Appointments.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Normal, now.AddDays(8), now.AddDays(8).AddHours(2), "TEST_SINGLE")
			schedulerDataStorage1.Appointments.Add(apt)
		End Sub

		''' <summary>
		''' Import data from Exchange to the DevExpress WinForms Scheduler.
		''' </summary>
		Private Sub btn_sync_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_sync.Click
			' Receive from Exchange all appointments starting from the current date and a month in advance.
			Dim apts As List(Of Microsoft.Exchange.WebServices.Data.Appointment) = synchronizer.GetExchangeAppointments(Date.Now, Date.Now.AddMonths(1))
			For Each apt In apts
				synchronizer.ExchangeToScheduler(schedulerControl1.DataStorage.Appointments, apt, True)
			Next apt
			XtraMessageBox.Show("Imported")
		End Sub

		Private Sub btn_export_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_export.Click
			For Each schedulerAppointment In schedulerDataStorage1.Appointments.Items
				If (schedulerAppointment.Type = DevExpress.XtraScheduler.AppointmentType.Normal) OrElse (schedulerAppointment.Type = DevExpress.XtraScheduler.AppointmentType.Pattern) Then
					' Export Patern and Normal Appointments to Exchange.
					synchronizer.SchedulerToExchange(schedulerAppointment, True)
				End If
			Next schedulerAppointment
			XtraMessageBox.Show("Exported")
		End Sub

		''' <summary>
		''' Exchange authorization
		''' </summary>
		Private Async Sub btn_authorize_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btn_authorize.Click
			Dim service As ExchangeService = Nothing
			Try
				service = Await Authorization.GetExchangeWebService()
			Catch ex As Exception
				XtraMessageBox.Show("Authorization error: " & ex.Message)
			End Try

			If service IsNot Nothing Then
				Me.synchronizer = New Synchronizer(service)
				btn_export.Enabled = True
				btn_sync.Enabled = True
				btn_authorize.Enabled = False
			End If
		End Sub
	End Class
End Namespace
