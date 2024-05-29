using System;
using System.Collections.Generic;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Microsoft.Exchange.WebServices.Data;

namespace EWSSyncExample {
    public partial class ExchangeExampleForm : XtraForm {
        Synchronizer synchronizer;        

        public ExchangeExampleForm() {
            InitializeComponent();
            schedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Month;

            DateTime now = DateTime.Now.Date.AddHours(18);
            // Create a recurring Appointment.
            var apt = schedulerDataStorage1.Appointments.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Pattern, now, now.AddHours(2), "TEST_TEST");

            apt.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily;
            apt.RecurrenceInfo.Periodicity = 1;
            apt.RecurrenceInfo.Start = now;
            apt.RecurrenceInfo.End = now.AddDays(6);
            apt.RecurrenceInfo.Range = DevExpress.XtraScheduler.RecurrenceRange.EndByDate;
            apt.RecurrenceInfo.Type = DevExpress.XtraScheduler.RecurrenceType.Daily;
            schedulerDataStorage1.Appointments.Add(apt);
            apt.CreateException(DevExpress.XtraScheduler.AppointmentType.DeletedOccurrence, 4);
            var apt2 = apt.CreateException(DevExpress.XtraScheduler.AppointmentType.ChangedOccurrence, 3);
            apt2.Start = now.AddDays(3).AddHours(2); 

            // Create a regular Appointment
            apt = schedulerDataStorage1.Appointments.CreateAppointment(DevExpress.XtraScheduler.AppointmentType.Normal, now.AddDays(8), now.AddDays(8).AddHours(2), "TEST_SINGLE");
            schedulerDataStorage1.Appointments.Add(apt);
        }

        /// <summary>
        /// Import data from Exchange to the DevExpress WinForms Scheduler.
        /// </summary>
        private void btn_sync_Click(object sender, EventArgs e) {
            // Receive from Exchange all appointments starting from the current date and a month in advance.
            List<Microsoft.Exchange.WebServices.Data.Appointment> apts = synchronizer.GetExchangeAppointments(DateTime.Now, DateTime.Now.AddMonths(1));
            foreach(var apt in apts)
                synchronizer.ExchangeToScheduler(schedulerControl1.DataStorage.Appointments, apt, true);
            XtraMessageBox.Show("Imported");
        }

        private void btn_export_Click(object sender, EventArgs e) {            
            foreach(var schedulerAppointment in schedulerDataStorage1.Appointments.Items) 
                if((schedulerAppointment.Type == DevExpress.XtraScheduler.AppointmentType.Normal) || (schedulerAppointment.Type == DevExpress.XtraScheduler.AppointmentType.Pattern))
                    // Export Patern and Normal Appointments to Exchange.
                    synchronizer.SchedulerToExchange(schedulerAppointment, true);
            XtraMessageBox.Show("Exported");
        }

        /// <summary>
        /// Exchange authorization
        /// </summary>
        private async void btn_authorize_Click(object sender, EventArgs e) {
            ExchangeService service = null;
            try {
                service = await Authorization.GetExchangeWebService();
            }
            catch(Exception ex) {
                XtraMessageBox.Show("Authorization error: " + ex.Message);
            }

            if(service != null) {
                this.synchronizer = new Synchronizer(service);
                btn_export.Enabled = true;
                btn_sync.Enabled = true;
                btn_authorize.Enabled = false;
            }
        }
    }
}
