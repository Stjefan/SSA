<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Help
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
        Me.Panel_Help = New System.Windows.Forms.Panel
        Me.PB_3 = New System.Windows.Forms.PictureBox
        Me.PB_2 = New System.Windows.Forms.PictureBox
        Me.PB = New System.Windows.Forms.PictureBox
        Me.Panel_Help.SuspendLayout()
        CType(Me.PB_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB_2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_Help
        '
        Me.Panel_Help.AutoScroll = True
        Me.Panel_Help.BackColor = System.Drawing.SystemColors.Desktop
        Me.Panel_Help.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel_Help.Controls.Add(Me.PB_3)
        Me.Panel_Help.Controls.Add(Me.PB_2)
        Me.Panel_Help.Controls.Add(Me.PB)
        Me.Panel_Help.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Help.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Help.Name = "Panel_Help"
        Me.Panel_Help.Size = New System.Drawing.Size(329, 217)
        Me.Panel_Help.TabIndex = 6
        '
        'PB_3
        '
        Me.PB_3.BackColor = System.Drawing.Color.White
        Me.PB_3.Location = New System.Drawing.Point(3, 155)
        Me.PB_3.Name = "PB_3"
        Me.PB_3.Size = New System.Drawing.Size(238, 23)
        Me.PB_3.TabIndex = 2
        Me.PB_3.TabStop = False
        '
        'PB_2
        '
        Me.PB_2.BackColor = System.Drawing.Color.White
        Me.PB_2.Location = New System.Drawing.Point(3, 109)
        Me.PB_2.Name = "PB_2"
        Me.PB_2.Size = New System.Drawing.Size(238, 26)
        Me.PB_2.TabIndex = 1
        Me.PB_2.TabStop = False
        '
        'PB
        '
        Me.PB.BackColor = System.Drawing.Color.White
        Me.PB.Location = New System.Drawing.Point(3, 3)
        Me.PB.Name = "PB"
        Me.PB.Size = New System.Drawing.Size(214, 84)
        Me.PB.TabIndex = 0
        Me.PB.TabStop = False
        '
        'Form_Help
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(329, 217)
        Me.Controls.Add(Me.Panel_Help)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "Form_Help"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Direkthilfe"
        Me.Panel_Help.ResumeLayout(False)
        CType(Me.PB_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB_2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel_Help As System.Windows.Forms.Panel
    Friend WithEvents PB_3 As System.Windows.Forms.PictureBox
    Friend WithEvents PB_2 As System.Windows.Forms.PictureBox
    Friend WithEvents PB As System.Windows.Forms.PictureBox
End Class
