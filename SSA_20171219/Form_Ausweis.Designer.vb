<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Ausweis
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
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage_Deckblatt = New System.Windows.Forms.TabPage
        Me.TabPage_detailliert = New System.Windows.Forms.TabPage
        Me.TabPage_einfach = New System.Windows.Forms.TabPage
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage_Deckblatt)
        Me.TabControl1.Controls.Add(Me.TabPage_detailliert)
        Me.TabControl1.Controls.Add(Me.TabPage_einfach)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1196, 782)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage_Deckblatt
        '
        Me.TabPage_Deckblatt.AutoScroll = True
        Me.TabPage_Deckblatt.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_Deckblatt.Name = "TabPage_Deckblatt"
        Me.TabPage_Deckblatt.Size = New System.Drawing.Size(1188, 756)
        Me.TabPage_Deckblatt.TabIndex = 2
        Me.TabPage_Deckblatt.Text = "Deckblatt"
        Me.TabPage_Deckblatt.UseVisualStyleBackColor = True
        '
        'TabPage_detailliert
        '
        Me.TabPage_detailliert.AutoScroll = True
        Me.TabPage_detailliert.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.TabPage_detailliert.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_detailliert.Name = "TabPage_detailliert"
        Me.TabPage_detailliert.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_detailliert.Size = New System.Drawing.Size(1188, 756)
        Me.TabPage_detailliert.TabIndex = 0
        Me.TabPage_detailliert.Text = "Detaillierter Schallschutzausweis"
        Me.TabPage_detailliert.UseVisualStyleBackColor = True
        '
        'TabPage_einfach
        '
        Me.TabPage_einfach.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.TabPage_einfach.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_einfach.Name = "TabPage_einfach"
        Me.TabPage_einfach.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_einfach.Size = New System.Drawing.Size(1188, 756)
        Me.TabPage_einfach.TabIndex = 1
        Me.TabPage_einfach.Text = "einfacher Schallschutzausweis"
        Me.TabPage_einfach.UseVisualStyleBackColor = True
        '
        'Form_Ausweis
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1196, 782)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "Form_Ausweis"
        Me.Text = "Schallschutzausweis"
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_detailliert As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_einfach As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Deckblatt As System.Windows.Forms.TabPage
End Class
