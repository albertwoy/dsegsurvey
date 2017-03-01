<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmsurveyreport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.cboryear = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dg1 = New System.Windows.Forms.DataGridView()
        Me.cbocol = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbosy = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblgtot = New System.Windows.Forms.Label()
        Me.btnexport = New System.Windows.Forms.Button()
        Me.btncancelthread = New System.Windows.Forms.Button()
        Me.lblbarep = New System.Windows.Forms.Label()
        Me.pbar1 = New System.Windows.Forms.ProgressBar()
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboryear
        '
        Me.cboryear.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cboryear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboryear.FormattingEnabled = True
        Me.cboryear.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.cboryear.Location = New System.Drawing.Point(408, 7)
        Me.cboryear.Name = "cboryear"
        Me.cboryear.Size = New System.Drawing.Size(141, 21)
        Me.cboryear.TabIndex = 35
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(340, 13)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Select Year"
        '
        'dg1
        '
        Me.dg1.AllowUserToAddRows = False
        Me.dg1.AllowUserToDeleteRows = False
        Me.dg1.AllowUserToResizeColumns = False
        Me.dg1.AllowUserToResizeRows = False
        Me.dg1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dg1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg1.Cursor = System.Windows.Forms.Cursors.Hand
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dg1.DefaultCellStyle = DataGridViewCellStyle4
        Me.dg1.Location = New System.Drawing.Point(11, 34)
        Me.dg1.Name = "dg1"
        Me.dg1.RowHeadersVisible = False
        Me.dg1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dg1.RowTemplate.Height = 42
        Me.dg1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dg1.Size = New System.Drawing.Size(947, 322)
        Me.dg1.TabIndex = 33
        '
        'cbocol
        '
        Me.cbocol.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbocol.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbocol.FormattingEnabled = True
        Me.cbocol.Location = New System.Drawing.Point(649, 7)
        Me.cbocol.Name = "cbocol"
        Me.cbocol.Size = New System.Drawing.Size(282, 21)
        Me.cbocol.TabIndex = 32
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(568, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Select College"
        '
        'cbosy
        '
        Me.cbosy.Cursor = System.Windows.Forms.Cursors.Hand
        Me.cbosy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbosy.FormattingEnabled = True
        Me.cbosy.Location = New System.Drawing.Point(125, 7)
        Me.cbosy.Name = "cbosy"
        Me.cbosy.Size = New System.Drawing.Size(195, 21)
        Me.cbosy.TabIndex = 30
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(98, 13)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Select School Year"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 365)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(153, 13)
        Me.Label4.TabIndex = 36
        Me.Label4.Text = "Total Number of Respondent : "
        '
        'lblgtot
        '
        Me.lblgtot.AutoSize = True
        Me.lblgtot.Location = New System.Drawing.Point(171, 365)
        Me.lblgtot.Name = "lblgtot"
        Me.lblgtot.Size = New System.Drawing.Size(0, 13)
        Me.lblgtot.TabIndex = 37
        '
        'btnexport
        '
        Me.btnexport.Location = New System.Drawing.Point(12, 409)
        Me.btnexport.Name = "btnexport"
        Me.btnexport.Size = New System.Drawing.Size(947, 26)
        Me.btnexport.TabIndex = 47
        Me.btnexport.Text = "Export to Excel"
        Me.btnexport.UseVisualStyleBackColor = True
        '
        'btncancelthread
        '
        Me.btncancelthread.Location = New System.Drawing.Point(815, 358)
        Me.btncancelthread.Name = "btncancelthread"
        Me.btncancelthread.Size = New System.Drawing.Size(144, 26)
        Me.btncancelthread.TabIndex = 46
        Me.btncancelthread.Text = "Cancel Thread 1"
        Me.btncancelthread.UseVisualStyleBackColor = True
        Me.btncancelthread.Visible = False
        '
        'lblbarep
        '
        Me.lblbarep.AutoSize = True
        Me.lblbarep.Location = New System.Drawing.Point(506, 369)
        Me.lblbarep.Name = "lblbarep"
        Me.lblbarep.Size = New System.Drawing.Size(44, 13)
        Me.lblbarep.TabIndex = 45
        Me.lblbarep.Text = "lblbarep"
        Me.lblbarep.Visible = False
        '
        'pbar1
        '
        Me.pbar1.Location = New System.Drawing.Point(12, 385)
        Me.pbar1.Name = "pbar1"
        Me.pbar1.Size = New System.Drawing.Size(947, 21)
        Me.pbar1.TabIndex = 44
        Me.pbar1.Visible = False
        '
        'frmsurveyreport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(971, 435)
        Me.Controls.Add(Me.btnexport)
        Me.Controls.Add(Me.btncancelthread)
        Me.Controls.Add(Me.lblbarep)
        Me.Controls.Add(Me.pbar1)
        Me.Controls.Add(Me.lblgtot)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboryear)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dg1)
        Me.Controls.Add(Me.cbocol)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbosy)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmsurveyreport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Survey Report"
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboryear As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dg1 As System.Windows.Forms.DataGridView
    Friend WithEvents cbocol As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbosy As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lblgtot As System.Windows.Forms.Label
    Friend WithEvents btnexport As System.Windows.Forms.Button
    Friend WithEvents btncancelthread As System.Windows.Forms.Button
    Friend WithEvents lblbarep As System.Windows.Forms.Label
    Friend WithEvents pbar1 As System.Windows.Forms.ProgressBar
End Class
