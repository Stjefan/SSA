<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Projekt_einrichten
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TB_Projektname = New System.Windows.Forms.TextBox
        Me.LBL_Projektname = New System.Windows.Forms.Label
        Me.LBL_Projektverzeichnis = New System.Windows.Forms.Label
        Me.TB_Projektverzeichnis = New System.Windows.Forms.TextBox
        Me.BT_Verzeichnis_durchsuchen = New System.Windows.Forms.Button
        Me.BT_Ok = New System.Windows.Forms.Button
        Me.BT_Abbruch = New System.Windows.Forms.Button
        Me.FBD_Projekt_einrichten = New System.Windows.Forms.FolderBrowserDialog
        Me.SuspendLayout()
        '
        'TB_Projektname
        '
        Me.TB_Projektname.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_Projektname.Location = New System.Drawing.Point(263, 21)
        Me.TB_Projektname.Name = "TB_Projektname"
        Me.TB_Projektname.Size = New System.Drawing.Size(521, 23)
        Me.TB_Projektname.TabIndex = 1
        '
        'LBL_Projektname
        '
        Me.LBL_Projektname.AutoSize = True
        Me.LBL_Projektname.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBL_Projektname.Location = New System.Drawing.Point(12, 24)
        Me.LBL_Projektname.Name = "LBL_Projektname"
        Me.LBL_Projektname.Size = New System.Drawing.Size(245, 16)
        Me.LBL_Projektname.TabIndex = 1
        Me.LBL_Projektname.Text = "Bezeichung Projekt / Projektname"
        Me.LBL_Projektname.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'LBL_Projektverzeichnis
        '
        Me.LBL_Projektverzeichnis.AutoSize = True
        Me.LBL_Projektverzeichnis.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LBL_Projektverzeichnis.Location = New System.Drawing.Point(53, 88)
        Me.LBL_Projektverzeichnis.Name = "LBL_Projektverzeichnis"
        Me.LBL_Projektverzeichnis.Size = New System.Drawing.Size(204, 16)
        Me.LBL_Projektverzeichnis.TabIndex = 2
        Me.LBL_Projektverzeichnis.Text = "Dateiverzeichnis für Projekt"
        Me.LBL_Projektverzeichnis.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'TB_Projektverzeichnis
        '
        Me.TB_Projektverzeichnis.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_Projektverzeichnis.Location = New System.Drawing.Point(263, 85)
        Me.TB_Projektverzeichnis.Name = "TB_Projektverzeichnis"
        Me.TB_Projektverzeichnis.Size = New System.Drawing.Size(521, 23)
        Me.TB_Projektverzeichnis.TabIndex = 2
        '
        'BT_Verzeichnis_durchsuchen
        '
        Me.BT_Verzeichnis_durchsuchen.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Verzeichnis_durchsuchen.Location = New System.Drawing.Point(263, 114)
        Me.BT_Verzeichnis_durchsuchen.Name = "BT_Verzeichnis_durchsuchen"
        Me.BT_Verzeichnis_durchsuchen.Size = New System.Drawing.Size(100, 25)
        Me.BT_Verzeichnis_durchsuchen.TabIndex = 3
        Me.BT_Verzeichnis_durchsuchen.Text = "Durchsuchen"
        Me.BT_Verzeichnis_durchsuchen.UseVisualStyleBackColor = True
        '
        'BT_Ok
        '
        Me.BT_Ok.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.BT_Ok.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Ok.Location = New System.Drawing.Point(583, 130)
        Me.BT_Ok.Name = "BT_Ok"
        Me.BT_Ok.Size = New System.Drawing.Size(88, 29)
        Me.BT_Ok.TabIndex = 4
        Me.BT_Ok.Text = "OK"
        Me.BT_Ok.UseVisualStyleBackColor = True
        '
        'BT_Abbruch
        '
        Me.BT_Abbruch.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BT_Abbruch.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Abbruch.Location = New System.Drawing.Point(686, 130)
        Me.BT_Abbruch.Name = "BT_Abbruch"
        Me.BT_Abbruch.Size = New System.Drawing.Size(100, 29)
        Me.BT_Abbruch.TabIndex = 5
        Me.BT_Abbruch.Text = "Abbrechen"
        Me.BT_Abbruch.UseVisualStyleBackColor = True
        '
        'FBD_Projekt_einrichten
        '
        Me.FBD_Projekt_einrichten.SelectedPath = "C:\"
        '
        'Form_Projekt_einrichten
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(796, 171)
        Me.Controls.Add(Me.BT_Abbruch)
        Me.Controls.Add(Me.BT_Ok)
        Me.Controls.Add(Me.BT_Verzeichnis_durchsuchen)
        Me.Controls.Add(Me.TB_Projektverzeichnis)
        Me.Controls.Add(Me.LBL_Projektverzeichnis)
        Me.Controls.Add(Me.LBL_Projektname)
        Me.Controls.Add(Me.TB_Projektname)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form_Projekt_einrichten"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "neues Projekt einrichten"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TB_Projektname As System.Windows.Forms.TextBox
    Friend WithEvents LBL_Projektname As System.Windows.Forms.Label
    Friend WithEvents LBL_Projektverzeichnis As System.Windows.Forms.Label
    Friend WithEvents TB_Projektverzeichnis As System.Windows.Forms.TextBox
    Friend WithEvents BT_Verzeichnis_durchsuchen As System.Windows.Forms.Button
    Friend WithEvents BT_Ok As System.Windows.Forms.Button
    Friend WithEvents BT_Abbruch As System.Windows.Forms.Button
    Friend WithEvents FBD_Projekt_einrichten As System.Windows.Forms.FolderBrowserDialog
End Class
