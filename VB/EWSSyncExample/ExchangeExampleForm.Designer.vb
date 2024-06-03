Namespace EWSSyncExample
	Partial Public Class ExchangeExampleForm
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.components = New System.ComponentModel.Container()
			Dim timeRuler1 As New DevExpress.XtraScheduler.TimeRuler()
			Dim timeRuler2 As New DevExpress.XtraScheduler.TimeRuler()
			Dim timeRuler3 As New DevExpress.XtraScheduler.TimeRuler()
			Me.schedulerControl1 = New DevExpress.XtraScheduler.SchedulerControl()
			Me.schedulerDataStorage1 = New DevExpress.XtraScheduler.SchedulerDataStorage(Me.components)
			Me.btn_sync = New System.Windows.Forms.Button()
			Me.btn_authorize = New System.Windows.Forms.Button()
			Me.btn_export = New System.Windows.Forms.Button()
			CType(Me.schedulerControl1, System.ComponentModel.ISupportInitialize).BeginInit()
			CType(Me.schedulerDataStorage1, System.ComponentModel.ISupportInitialize).BeginInit()
			Me.SuspendLayout()
			' 
			' schedulerControl1
			' 
			Me.schedulerControl1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.schedulerControl1.DataStorage = Me.schedulerDataStorage1
			Me.schedulerControl1.Location = New System.Drawing.Point(24, 115)
			Me.schedulerControl1.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.schedulerControl1.Name = "schedulerControl1"
			Me.schedulerControl1.Size = New System.Drawing.Size(1552, 727)
			Me.schedulerControl1.Start = New Date(2023, 11, 30, 0, 0, 0, 0)
			Me.schedulerControl1.TabIndex = 0
			Me.schedulerControl1.Text = "schedulerControl1"
			Me.schedulerControl1.Views.DayView.TimeRulers.Add(timeRuler1)
			Me.schedulerControl1.Views.FullWeekView.Enabled = True
			Me.schedulerControl1.Views.FullWeekView.TimeRulers.Add(timeRuler2)
			Me.schedulerControl1.Views.WeekView.Enabled = False
			Me.schedulerControl1.Views.WorkWeekView.TimeRulers.Add(timeRuler3)
			Me.schedulerControl1.Views.YearView.UseOptimizedScrolling = False
			' 
			' schedulerDataStorage1
			' 
			' 
			' 
			' 
			Me.schedulerDataStorage1.AppointmentDependencies.AutoReload = False
			' 
			' 
			' 
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(0, "None", "&None", System.Drawing.SystemColors.Window)
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(1, "Important", "&Important", System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(194)))), (CInt((CByte(190))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(2, "Business", "&Business", System.Drawing.Color.FromArgb((CInt((CByte(168)))), (CInt((CByte(213)))), (CInt((CByte(255))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(3, "Personal", "&Personal", System.Drawing.Color.FromArgb((CInt((CByte(193)))), (CInt((CByte(244)))), (CInt((CByte(156))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(4, "Vacation", "&Vacation", System.Drawing.Color.FromArgb((CInt((CByte(243)))), (CInt((CByte(228)))), (CInt((CByte(199))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(5, "Must Attend", "Must &Attend", System.Drawing.Color.FromArgb((CInt((CByte(244)))), (CInt((CByte(206)))), (CInt((CByte(147))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(6, "Travel Required", "&Travel Required", System.Drawing.Color.FromArgb((CInt((CByte(199)))), (CInt((CByte(244)))), (CInt((CByte(255))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(7, "Needs Preparation", "&Needs Preparation", System.Drawing.Color.FromArgb((CInt((CByte(207)))), (CInt((CByte(219)))), (CInt((CByte(152))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(8, "Birthday", "&Birthday", System.Drawing.Color.FromArgb((CInt((CByte(224)))), (CInt((CByte(207)))), (CInt((CByte(233))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(9, "Anniversary", "&Anniversary", System.Drawing.Color.FromArgb((CInt((CByte(141)))), (CInt((CByte(233)))), (CInt((CByte(223))))))
			Me.schedulerDataStorage1.Appointments.Labels.CreateNewLabel(10, "Phone Call", "Phone &Call", System.Drawing.Color.FromArgb((CInt((CByte(255)))), (CInt((CByte(247)))), (CInt((CByte(165))))))
			' 
			' btn_sync
			' 
			Me.btn_sync.Enabled = False
			Me.btn_sync.Location = New System.Drawing.Point(256, 23)
			Me.btn_sync.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btn_sync.Name = "btn_sync"
			Me.btn_sync.Size = New System.Drawing.Size(200, 81)
			Me.btn_sync.TabIndex = 1
			Me.btn_sync.Text = "Import from Exchange"
			Me.btn_sync.UseVisualStyleBackColor = True
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.btn_sync.Click += new System.EventHandler(this.btn_sync_Click);
			' 
			' btn_authorize
			' 
			Me.btn_authorize.Location = New System.Drawing.Point(24, 23)
			Me.btn_authorize.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btn_authorize.Name = "btn_authorize"
			Me.btn_authorize.Size = New System.Drawing.Size(200, 81)
			Me.btn_authorize.TabIndex = 2
			Me.btn_authorize.Text = "Authorize"
			Me.btn_authorize.UseVisualStyleBackColor = True
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.btn_authorize.Click += new System.EventHandler(this.btn_authorize_Click);
			' 
			' btn_export
			' 
			Me.btn_export.Enabled = False
			Me.btn_export.Location = New System.Drawing.Point(468, 23)
			Me.btn_export.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.btn_export.Name = "btn_export"
			Me.btn_export.Size = New System.Drawing.Size(200, 81)
			Me.btn_export.TabIndex = 3
			Me.btn_export.Text = "Export to Exchange"
			Me.btn_export.UseVisualStyleBackColor = True
'INSTANT VB NOTE: The following InitializeComponent event wireup was converted to a 'Handles' clause:
'ORIGINAL LINE: this.btn_export.Click += new System.EventHandler(this.btn_export_Click);
			' 
			' ExchangeExampleForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(12F, 25F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(1600, 865)
			Me.Controls.Add(Me.btn_export)
			Me.Controls.Add(Me.btn_authorize)
			Me.Controls.Add(Me.btn_sync)
			Me.Controls.Add(Me.schedulerControl1)
			Me.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
			Me.Name = "ExchangeExampleForm"
			Me.Text = "Form1"
			CType(Me.schedulerControl1, System.ComponentModel.ISupportInitialize).EndInit()
			CType(Me.schedulerDataStorage1, System.ComponentModel.ISupportInitialize).EndInit()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private schedulerControl1 As DevExpress.XtraScheduler.SchedulerControl
		Private schedulerDataStorage1 As DevExpress.XtraScheduler.SchedulerDataStorage
		Private WithEvents btn_sync As System.Windows.Forms.Button
		Private WithEvents btn_authorize As System.Windows.Forms.Button
		Private WithEvents btn_export As System.Windows.Forms.Button
	End Class
End Namespace

