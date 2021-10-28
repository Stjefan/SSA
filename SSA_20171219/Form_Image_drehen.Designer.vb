<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Image_drehen
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.PNL_Lageplan_drehen = New System.Windows.Forms.Panel
        Me.BT_Abbruch_Uhrzeigersinn = New System.Windows.Forms.Button
        Me.BT_gegen_Uhrzeigersinn = New System.Windows.Forms.Button
        Me.BT_im_Uhrzeigersinn = New System.Windows.Forms.Button
        Me.LBL_Lageplan_drehen = New System.Windows.Forms.Label
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PNL_Lageplan_drehen.SuspendLayout()
        Me.SuspendLayout()
        '
        'PNL_Lageplan_drehen
        '
        Me.PNL_Lageplan_drehen.BackColor = System.Drawing.SystemColors.Control
        Me.PNL_Lageplan_drehen.Controls.Add(Me.BT_Abbruch_Uhrzeigersinn)
        Me.PNL_Lageplan_drehen.Controls.Add(Me.BT_gegen_Uhrzeigersinn)
        Me.PNL_Lageplan_drehen.Controls.Add(Me.BT_im_Uhrzeigersinn)
        Me.PNL_Lageplan_drehen.Controls.Add(Me.LBL_Lageplan_drehen)
        Me.PNL_Lageplan_drehen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PNL_Lageplan_drehen.Location = New System.Drawing.Point(0, 0)
        Me.PNL_Lageplan_drehen.Name = "PNL_Lageplan_drehen"
        Me.PNL_Lageplan_drehen.Padding = New System.Windows.Forms.Padding(5)
        Me.PNL_Lageplan_drehen.Size = New System.Drawing.Size(292, 216)
        Me.PNL_Lageplan_drehen.TabIndex = 10
        '
        'BT_Abbruch_Uhrzeigersinn
        '
        Me.BT_Abbruch_Uhrzeigersinn.BackColor = System.Drawing.SystemColors.Control
        Me.BT_Abbruch_Uhrzeigersinn.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.BT_Abbruch_Uhrzeigersinn.Location = New System.Drawing.Point(110, 163)
        Me.BT_Abbruch_Uhrzeigersinn.Name = "BT_Abbruch_Uhrzeigersinn"
        Me.BT_Abbruch_Uhrzeigersinn.Size = New System.Drawing.Size(75, 23)
        Me.BT_Abbruch_Uhrzeigersinn.TabIndex = 3
        Me.BT_Abbruch_Uhrzeigersinn.Text = "OK"
        Me.ToolTip1.SetToolTip(Me.BT_Abbruch_Uhrzeigersinn, "Der Grundriss wird nicht gedreht")
        Me.BT_Abbruch_Uhrzeigersinn.UseVisualStyleBackColor = False
        '
        'BT_gegen_Uhrzeigersinn
        '
        Me.BT_gegen_Uhrzeigersinn.BackColor = System.Drawing.SystemColors.Control
        Me.BT_gegen_Uhrzeigersinn.Location = New System.Drawing.Point(159, 106)
        Me.BT_gegen_Uhrzeigersinn.Name = "BT_gegen_Uhrzeigersinn"
        Me.BT_gegen_Uhrzeigersinn.Size = New System.Drawing.Size(88, 37)
        Me.BT_gegen_Uhrzeigersinn.TabIndex = 2
        Me.BT_gegen_Uhrzeigersinn.Text = "Drehen gegen Uhrzeigersinn"
        Me.ToolTip1.SetToolTip(Me.BT_gegen_Uhrzeigersinn, "Der Lageplan wir um 90° gegen den Uhrzeigersinn gedreht")
        Me.BT_gegen_Uhrzeigersinn.UseVisualStyleBackColor = False
        '
        'BT_im_Uhrzeigersinn
        '
        Me.BT_im_Uhrzeigersinn.BackColor = System.Drawing.SystemColors.Control
        Me.BT_im_Uhrzeigersinn.Location = New System.Drawing.Point(48, 106)
        Me.BT_im_Uhrzeigersinn.Name = "BT_im_Uhrzeigersinn"
        Me.BT_im_Uhrzeigersinn.Size = New System.Drawing.Size(88, 37)
        Me.BT_im_Uhrzeigersinn.TabIndex = 1
        Me.BT_im_Uhrzeigersinn.Text = "Drehen im Uhrzeigersinn"
        Me.ToolTip1.SetToolTip(Me.BT_im_Uhrzeigersinn, "Der Lageplan wir um 90° im Uhrzeigersinn gedreht")
        Me.BT_im_Uhrzeigersinn.UseVisualStyleBackColor = False
        '
        'LBL_Lageplan_drehen
        '
        Me.LBL_Lageplan_drehen.AutoSize = True
        Me.LBL_Lageplan_drehen.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBL_Lageplan_drehen.Location = New System.Drawing.Point(73, 38)
        Me.LBL_Lageplan_drehen.Name = "LBL_Lageplan_drehen"
        Me.LBL_Lageplan_drehen.Size = New System.Drawing.Size(153, 40)
        Me.LBL_Lageplan_drehen.TabIndex = 0
        Me.LBL_Lageplan_drehen.Text = "Soll der Grundriss" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "gedreht werden?"
        Me.LBL_Lageplan_drehen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form_Image_drehen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(292, 216)
        Me.Controls.Add(Me.PNL_Lageplan_drehen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "Form_Image_drehen"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Grundriss drehen"
        Me.PNL_Lageplan_drehen.ResumeLayout(False)
        Me.PNL_Lageplan_drehen.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents PNL_Lageplan_drehen As System.Windows.Forms.Panel
    Friend WithEvents BT_Abbruch_Uhrzeigersinn As System.Windows.Forms.Button
    Friend WithEvents BT_gegen_Uhrzeigersinn As System.Windows.Forms.Button
    Friend WithEvents BT_im_Uhrzeigersinn As System.Windows.Forms.Button
    Friend WithEvents LBL_Lageplan_drehen As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
