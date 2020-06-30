<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.cboEmailName = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TxtPrompt2 = New System.Windows.Forms.TextBox()
        Me.Date2 = New System.Windows.Forms.DateTimePicker()
        Me.lblPrompt2 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.TxtPrompt1 = New System.Windows.Forms.TextBox()
        Me.Date1 = New System.Windows.Forms.DateTimePicker()
        Me.lblPrompt1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cboEmailName
        '
        Me.cboEmailName.FormattingEnabled = True
        Me.cboEmailName.Location = New System.Drawing.Point(155, 29)
        Me.cboEmailName.Name = "cboEmailName"
        Me.cboEmailName.Size = New System.Drawing.Size(340, 21)
        Me.cboEmailName.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(59, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select Email"
        '
        'TxtPrompt2
        '
        Me.TxtPrompt2.Location = New System.Drawing.Point(258, 116)
        Me.TxtPrompt2.Name = "TxtPrompt2"
        Me.TxtPrompt2.Size = New System.Drawing.Size(176, 20)
        Me.TxtPrompt2.TabIndex = 15
        Me.TxtPrompt2.Visible = False
        '
        'Date2
        '
        Me.Date2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.Date2.Location = New System.Drawing.Point(258, 116)
        Me.Date2.Name = "Date2"
        Me.Date2.Size = New System.Drawing.Size(200, 20)
        Me.Date2.TabIndex = 14
        Me.Date2.Visible = False
        '
        'lblPrompt2
        '
        Me.lblPrompt2.AutoSize = True
        Me.lblPrompt2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrompt2.Location = New System.Drawing.Point(59, 117)
        Me.lblPrompt2.Name = "lblPrompt2"
        Me.lblPrompt2.Size = New System.Drawing.Size(49, 16)
        Me.lblPrompt2.TabIndex = 13
        Me.lblPrompt2.Text = "Label1"
        Me.lblPrompt2.Visible = False
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(318, 184)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(86, 47)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(102, 184)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(86, 47)
        Me.btnOK.TabIndex = 11
        Me.btnOK.Text = "Send Email"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'TxtPrompt1
        '
        Me.TxtPrompt1.Location = New System.Drawing.Point(258, 77)
        Me.TxtPrompt1.Name = "TxtPrompt1"
        Me.TxtPrompt1.Size = New System.Drawing.Size(176, 20)
        Me.TxtPrompt1.TabIndex = 10
        Me.TxtPrompt1.Visible = False
        '
        'Date1
        '
        Me.Date1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.Date1.Location = New System.Drawing.Point(258, 77)
        Me.Date1.Name = "Date1"
        Me.Date1.Size = New System.Drawing.Size(200, 20)
        Me.Date1.TabIndex = 9
        Me.Date1.Visible = False
        '
        'lblPrompt1
        '
        Me.lblPrompt1.AutoSize = True
        Me.lblPrompt1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrompt1.Location = New System.Drawing.Point(59, 78)
        Me.lblPrompt1.Name = "lblPrompt1"
        Me.lblPrompt1.Size = New System.Drawing.Size(49, 16)
        Me.lblPrompt1.TabIndex = 8
        Me.lblPrompt1.Text = "Label1"
        Me.lblPrompt1.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 261)
        Me.Controls.Add(Me.TxtPrompt2)
        Me.Controls.Add(Me.Date2)
        Me.Controls.Add(Me.lblPrompt2)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.TxtPrompt1)
        Me.Controls.Add(Me.Date1)
        Me.Controls.Add(Me.lblPrompt1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboEmailName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Send Emails"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cboEmailName As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TxtPrompt2 As System.Windows.Forms.TextBox
    Friend WithEvents Date2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPrompt2 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents TxtPrompt1 As System.Windows.Forms.TextBox
    Friend WithEvents Date1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPrompt1 As System.Windows.Forms.Label

End Class
