<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmsurveyform
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
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmsurveyform))
        Me.cboyr1 = New System.Windows.Forms.ComboBox()
        Me.cbocrs1 = New System.Windows.Forms.ComboBox()
        Me.lblcrow = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgq1 = New System.Windows.Forms.DataGridView()
        Me.Category = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.answers = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.CATNO = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblsurveyid = New System.Windows.Forms.Label()
        Me.tsot1 = New System.Windows.Forms.ToolStrip()
        Me.saveme = New System.Windows.Forms.ToolStripButton()
        Me.tsreport = New System.Windows.Forms.ToolStripButton()
        Me.tsreset = New System.Windows.Forms.ToolStripButton()
        Me.tsrespondent = New System.Windows.Forms.ToolStripButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtage = New System.Windows.Forms.TextBox()
        Me.txthome = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tresno = New System.Windows.Forms.Label()
        Me.cbogender = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.lblrno = New System.Windows.Forms.Label()
        Me.lblrn = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        CType(Me.dgq1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tsot1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboyr1
        '
        Me.cboyr1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboyr1.FormattingEnabled = True
        Me.cboyr1.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6"})
        Me.cboyr1.Location = New System.Drawing.Point(299, 19)
        Me.cboyr1.Name = "cboyr1"
        Me.cboyr1.Size = New System.Drawing.Size(80, 21)
        Me.cboyr1.TabIndex = 1
        '
        'cbocrs1
        '
        Me.cbocrs1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbocrs1.FormattingEnabled = True
        Me.cbocrs1.Location = New System.Drawing.Point(67, 19)
        Me.cbocrs1.Name = "cbocrs1"
        Me.cbocrs1.Size = New System.Drawing.Size(177, 21)
        Me.cbocrs1.TabIndex = 0
        '
        'lblcrow
        '
        Me.lblcrow.AutoSize = True
        Me.lblcrow.Location = New System.Drawing.Point(948, 78)
        Me.lblcrow.Name = "lblcrow"
        Me.lblcrow.Size = New System.Drawing.Size(30, 13)
        Me.lblcrow.TabIndex = 24
        Me.lblcrow.Text = "crow"
        Me.lblcrow.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(267, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 13)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Year"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 22
        Me.Label1.Text = "Course"
        '
        'dgq1
        '
        Me.dgq1.AllowUserToAddRows = False
        Me.dgq1.AllowUserToDeleteRows = False
        Me.dgq1.AllowUserToResizeColumns = False
        Me.dgq1.AllowUserToResizeRows = False
        Me.dgq1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgq1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        Me.dgq1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgq1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Category, Me.answers, Me.CATNO})
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgq1.DefaultCellStyle = DataGridViewCellStyle3
        Me.dgq1.Location = New System.Drawing.Point(24, 102)
        Me.dgq1.Name = "dgq1"
        Me.dgq1.RowHeadersVisible = False
        Me.dgq1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgq1.RowTemplate.Height = 42
        Me.dgq1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgq1.Size = New System.Drawing.Size(1002, 503)
        Me.dgq1.TabIndex = 5
        '
        'Category
        '
        Me.Category.HeaderText = "QUESTION"
        Me.Category.Name = "Category"
        Me.Category.Width = 250
        '
        'answers
        '
        Me.answers.DataPropertyName = "answers"
        Me.answers.DisplayStyleForCurrentCellOnly = True
        Me.answers.HeaderText = "ANSWER"
        Me.answers.Name = "answers"
        Me.answers.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.answers.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.answers.ToolTipText = "click cell to choose"
        Me.answers.Width = 200
        '
        'CATNO
        '
        Me.CATNO.HeaderText = "CATNO"
        Me.CATNO.Name = "CATNO"
        Me.CATNO.Visible = False
        '
        'lblsurveyid
        '
        Me.lblsurveyid.AutoSize = True
        Me.lblsurveyid.Location = New System.Drawing.Point(857, 78)
        Me.lblsurveyid.Name = "lblsurveyid"
        Me.lblsurveyid.Size = New System.Drawing.Size(56, 13)
        Me.lblsurveyid.TabIndex = 17
        Me.lblsurveyid.Text = "lblsurveyid"
        Me.lblsurveyid.Visible = False
        '
        'tsot1
        '
        Me.tsot1.AutoSize = False
        Me.tsot1.BackColor = System.Drawing.Color.Khaki
        Me.tsot1.Dock = System.Windows.Forms.DockStyle.None
        Me.tsot1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.saveme, Me.tsreport, Me.tsreset, Me.tsrespondent})
        Me.tsot1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.Flow
        Me.tsot1.Location = New System.Drawing.Point(1, 608)
        Me.tsot1.Name = "tsot1"
        Me.tsot1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.tsot1.Size = New System.Drawing.Size(1036, 46)
        Me.tsot1.TabIndex = 7
        Me.tsot1.Text = "ToolStrip1"
        '
        'saveme
        '
        Me.saveme.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.saveme.AutoSize = False
        Me.saveme.Image = CType(resources.GetObject("saveme.Image"), System.Drawing.Image)
        Me.saveme.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.saveme.Name = "saveme"
        Me.saveme.Padding = New System.Windows.Forms.Padding(9)
        Me.saveme.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.saveme.Size = New System.Drawing.Size(140, 45)
        Me.saveme.Text = "Save Record Ctrl + S"
        '
        'tsreport
        '
        Me.tsreport.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsreport.AutoSize = False
        Me.tsreport.Image = Global.DSEGSURVEY.My.Resources.Resources._001_44
        Me.tsreport.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsreport.Name = "tsreport"
        Me.tsreport.Padding = New System.Windows.Forms.Padding(9)
        Me.tsreport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tsreport.Size = New System.Drawing.Size(140, 45)
        Me.tsreport.Text = "Go To Report"
        Me.tsreport.ToolTipText = "Reset Ctrl + R"
        '
        'tsreset
        '
        Me.tsreset.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsreset.AutoSize = False
        Me.tsreset.Image = CType(resources.GetObject("tsreset.Image"), System.Drawing.Image)
        Me.tsreset.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsreset.Name = "tsreset"
        Me.tsreset.Padding = New System.Windows.Forms.Padding(9)
        Me.tsreset.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tsreset.Size = New System.Drawing.Size(140, 45)
        Me.tsreset.Text = "Reset/Clear Ctrl + R"
        Me.tsreset.ToolTipText = "Reset Ctrl + R"
        '
        'tsrespondent
        '
        Me.tsrespondent.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsrespondent.AutoSize = False
        Me.tsrespondent.Image = CType(resources.GetObject("tsrespondent.Image"), System.Drawing.Image)
        Me.tsrespondent.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsrespondent.Name = "tsrespondent"
        Me.tsrespondent.Padding = New System.Windows.Forms.Padding(9)
        Me.tsrespondent.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tsrespondent.Size = New System.Drawing.Size(140, 45)
        Me.tsrespondent.Text = "View Respondents"
        Me.tsrespondent.ToolTipText = "Reset Ctrl + R"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(35, 52)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(26, 13)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Age"
        '
        'txtage
        '
        Me.txtage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtage.Location = New System.Drawing.Point(67, 49)
        Me.txtage.Name = "txtage"
        Me.txtage.Size = New System.Drawing.Size(44, 20)
        Me.txtage.TabIndex = 3
        '
        'txthome
        '
        Me.txthome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txthome.Location = New System.Drawing.Point(199, 49)
        Me.txthome.Name = "txthome"
        Me.txthome.Size = New System.Drawing.Size(562, 20)
        Me.txthome.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(117, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 13)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Home Address"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'tresno
        '
        Me.tresno.AutoSize = True
        Me.tresno.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tresno.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tresno.Location = New System.Drawing.Point(870, 10)
        Me.tresno.Name = "tresno"
        Me.tresno.Size = New System.Drawing.Size(2, 22)
        Me.tresno.TabIndex = 45
        '
        'cbogender
        '
        Me.cbogender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbogender.FormattingEnabled = True
        Me.cbogender.Items.AddRange(New Object() {"Female", "Male"})
        Me.cbogender.Location = New System.Drawing.Point(454, 19)
        Me.cbogender.Name = "cbogender"
        Me.cbogender.Size = New System.Drawing.Size(134, 21)
        Me.cbogender.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(410, 23)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(42, 13)
        Me.Label5.TabIndex = 46
        Me.Label5.Text = "Gender"
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(702, 15)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(162, 13)
        Me.Label21.TabIndex = 47
        Me.Label21.Text = "Temporary Respondent No."
        '
        'lblrno
        '
        Me.lblrno.AutoSize = True
        Me.lblrno.Location = New System.Drawing.Point(867, 36)
        Me.lblrno.Name = "lblrno"
        Me.lblrno.Size = New System.Drawing.Size(32, 13)
        Me.lblrno.TabIndex = 48
        Me.lblrno.Text = "lblrno"
        Me.lblrno.Visible = False
        '
        'lblrn
        '
        Me.lblrn.AutoSize = True
        Me.lblrn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblrn.Location = New System.Drawing.Point(762, 36)
        Me.lblrn.Name = "lblrn"
        Me.lblrn.Size = New System.Drawing.Size(99, 13)
        Me.lblrn.TabIndex = 51
        Me.lblrn.Text = "Respondent No."
        Me.lblrn.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(21, 86)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(248, 13)
        Me.Label6.TabIndex = 52
        Me.Label6.Text = "NOTE :: Please press enter key after entry"
        '
        'frmsurveyform
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1046, 655)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.lblrn)
        Me.Controls.Add(Me.lblrno)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.cbogender)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tresno)
        Me.Controls.Add(Me.txthome)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtage)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboyr1)
        Me.Controls.Add(Me.cbocrs1)
        Me.Controls.Add(Me.lblcrow)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgq1)
        Me.Controls.Add(Me.lblsurveyid)
        Me.Controls.Add(Me.tsot1)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmsurveyform"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SURVEY FORM "
        CType(Me.dgq1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tsot1.ResumeLayout(False)
        Me.tsot1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboyr1 As System.Windows.Forms.ComboBox
    Friend WithEvents cbocrs1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblcrow As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents dgq1 As System.Windows.Forms.DataGridView
    Friend WithEvents lblsurveyid As System.Windows.Forms.Label
    Friend WithEvents tsot1 As System.Windows.Forms.ToolStrip
    Friend WithEvents saveme As System.Windows.Forms.ToolStripButton
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtage As System.Windows.Forms.TextBox
    Friend WithEvents txthome As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tsreset As System.Windows.Forms.ToolStripButton
    Friend WithEvents tresno As System.Windows.Forms.Label
    Friend WithEvents cbogender As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Category As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents answers As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents CATNO As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents tsreport As System.Windows.Forms.ToolStripButton
    Friend WithEvents tsrespondent As System.Windows.Forms.ToolStripButton
    Friend WithEvents lblrno As System.Windows.Forms.Label
    Friend WithEvents lblrn As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label

End Class
