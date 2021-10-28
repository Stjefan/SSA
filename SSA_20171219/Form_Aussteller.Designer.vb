<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Aussteller
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form_Aussteller))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TB_Ort = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TB_PLZ = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TB_Nr = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TB_Strasse = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TB_Zusatz = New System.Windows.Forms.TextBox
        Me.TB_Name = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button_Logo_gegen = New System.Windows.Forms.Button
        Me.ImageList_Buttons = New System.Windows.Forms.ImageList(Me.components)
        Me.Button_Logo_entfernen = New System.Windows.Forms.Button
        Me.Button_Logo_mit = New System.Windows.Forms.Button
        Me.Button_Logo_hinzufuegen = New System.Windows.Forms.Button
        Me.PB_Logo = New System.Windows.Forms.PictureBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button_Signatur_gegen = New System.Windows.Forms.Button
        Me.Button_Signatur_entfernen = New System.Windows.Forms.Button
        Me.Button_Signatur_mit = New System.Windows.Forms.Button
        Me.Button_Signatur_hinzufuegen = New System.Windows.Forms.Button
        Me.PB_Signatur = New System.Windows.Forms.PictureBox
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.Button_OK = New System.Windows.Forms.Button
        Me.OFD_Logo = New System.Windows.Forms.OpenFileDialog
        Me.OFD_Signatur = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PB_Logo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        CType(Me.PB_Signatur, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TB_Ort)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TB_PLZ)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TB_Nr)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.TB_Strasse)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TB_Zusatz)
        Me.GroupBox1.Controls.Add(Me.TB_Name)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(388, 107)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Adresse"
        '
        'TB_Ort
        '
        Me.TB_Ort.Location = New System.Drawing.Point(182, 80)
        Me.TB_Ort.Name = "TB_Ort"
        Me.TB_Ort.Size = New System.Drawing.Size(197, 20)
        Me.TB_Ort.TabIndex = 24
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Location = New System.Drawing.Point(152, 84)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 13)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "Ort:"
        '
        'TB_PLZ
        '
        Me.TB_PLZ.Location = New System.Drawing.Point(78, 80)
        Me.TB_PLZ.Name = "TB_PLZ"
        Me.TB_PLZ.Size = New System.Drawing.Size(70, 20)
        Me.TB_PLZ.TabIndex = 22
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Location = New System.Drawing.Point(3, 84)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "PLZ:"
        '
        'TB_Nr
        '
        Me.TB_Nr.Location = New System.Drawing.Point(309, 59)
        Me.TB_Nr.Name = "TB_Nr"
        Me.TB_Nr.Size = New System.Drawing.Size(70, 20)
        Me.TB_Nr.TabIndex = 20
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Location = New System.Drawing.Point(282, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(21, 13)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Nr:"
        '
        'TB_Strasse
        '
        Me.TB_Strasse.Location = New System.Drawing.Point(78, 59)
        Me.TB_Strasse.Name = "TB_Strasse"
        Me.TB_Strasse.Size = New System.Drawing.Size(200, 20)
        Me.TB_Strasse.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Location = New System.Drawing.Point(3, 65)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(41, 13)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Straße:"
        '
        'TB_Zusatz
        '
        Me.TB_Zusatz.Location = New System.Drawing.Point(78, 38)
        Me.TB_Zusatz.Name = "TB_Zusatz"
        Me.TB_Zusatz.Size = New System.Drawing.Size(301, 20)
        Me.TB_Zusatz.TabIndex = 16
        '
        'TB_Name
        '
        Me.TB_Name.Location = New System.Drawing.Point(78, 17)
        Me.TB_Name.Name = "TB_Name"
        Me.TB_Name.Size = New System.Drawing.Size(301, 20)
        Me.TB_Name.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(3, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 13)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Zusatz:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Location = New System.Drawing.Point(3, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Name:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button_Logo_gegen)
        Me.GroupBox2.Controls.Add(Me.Button_Logo_entfernen)
        Me.GroupBox2.Controls.Add(Me.Button_Logo_mit)
        Me.GroupBox2.Controls.Add(Me.Button_Logo_hinzufuegen)
        Me.GroupBox2.Controls.Add(Me.PB_Logo)
        Me.GroupBox2.Location = New System.Drawing.Point(4, 113)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(388, 123)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Logo"
        '
        'Button_Logo_gegen
        '
        Me.Button_Logo_gegen.ImageKey = "gegenUhr.bmp"
        Me.Button_Logo_gegen.ImageList = Me.ImageList_Buttons
        Me.Button_Logo_gegen.Location = New System.Drawing.Point(342, 56)
        Me.Button_Logo_gegen.Name = "Button_Logo_gegen"
        Me.Button_Logo_gegen.Size = New System.Drawing.Size(37, 23)
        Me.Button_Logo_gegen.TabIndex = 4
        Me.Button_Logo_gegen.UseVisualStyleBackColor = True
        '
        'ImageList_Buttons
        '
        Me.ImageList_Buttons.ImageStream = CType(resources.GetObject("ImageList_Buttons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList_Buttons.TransparentColor = System.Drawing.Color.Magenta
        Me.ImageList_Buttons.Images.SetKeyName(0, "gegenUhr.bmp")
        Me.ImageList_Buttons.Images.SetKeyName(1, "mitUhr.bmp")
        '
        'Button_Logo_entfernen
        '
        Me.Button_Logo_entfernen.Location = New System.Drawing.Point(303, 85)
        Me.Button_Logo_entfernen.Name = "Button_Logo_entfernen"
        Me.Button_Logo_entfernen.Size = New System.Drawing.Size(75, 23)
        Me.Button_Logo_entfernen.TabIndex = 3
        Me.Button_Logo_entfernen.Text = "Entfernen"
        Me.Button_Logo_entfernen.UseVisualStyleBackColor = True
        '
        'Button_Logo_mit
        '
        Me.Button_Logo_mit.ImageKey = "mitUhr.bmp"
        Me.Button_Logo_mit.ImageList = Me.ImageList_Buttons
        Me.Button_Logo_mit.Location = New System.Drawing.Point(303, 56)
        Me.Button_Logo_mit.Name = "Button_Logo_mit"
        Me.Button_Logo_mit.Size = New System.Drawing.Size(37, 23)
        Me.Button_Logo_mit.TabIndex = 2
        Me.Button_Logo_mit.UseVisualStyleBackColor = True
        '
        'Button_Logo_hinzufuegen
        '
        Me.Button_Logo_hinzufuegen.Location = New System.Drawing.Point(303, 27)
        Me.Button_Logo_hinzufuegen.Name = "Button_Logo_hinzufuegen"
        Me.Button_Logo_hinzufuegen.Size = New System.Drawing.Size(75, 23)
        Me.Button_Logo_hinzufuegen.TabIndex = 1
        Me.Button_Logo_hinzufuegen.Text = "Hinzufügen"
        Me.Button_Logo_hinzufuegen.UseVisualStyleBackColor = True
        '
        'PB_Logo
        '
        Me.PB_Logo.BackColor = System.Drawing.Color.White
        Me.PB_Logo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PB_Logo.Location = New System.Drawing.Point(9, 16)
        Me.PB_Logo.Name = "PB_Logo"
        Me.PB_Logo.Size = New System.Drawing.Size(275, 101)
        Me.PB_Logo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_Logo.TabIndex = 0
        Me.PB_Logo.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button_Signatur_gegen)
        Me.GroupBox3.Controls.Add(Me.Button_Signatur_entfernen)
        Me.GroupBox3.Controls.Add(Me.Button_Signatur_mit)
        Me.GroupBox3.Controls.Add(Me.Button_Signatur_hinzufuegen)
        Me.GroupBox3.Controls.Add(Me.PB_Signatur)
        Me.GroupBox3.Location = New System.Drawing.Point(4, 242)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(388, 124)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Signatur"
        '
        'Button_Signatur_gegen
        '
        Me.Button_Signatur_gegen.ImageKey = "gegenUhr.bmp"
        Me.Button_Signatur_gegen.ImageList = Me.ImageList_Buttons
        Me.Button_Signatur_gegen.Location = New System.Drawing.Point(341, 53)
        Me.Button_Signatur_gegen.Name = "Button_Signatur_gegen"
        Me.Button_Signatur_gegen.Size = New System.Drawing.Size(37, 23)
        Me.Button_Signatur_gegen.TabIndex = 8
        Me.Button_Signatur_gegen.UseVisualStyleBackColor = True
        '
        'Button_Signatur_entfernen
        '
        Me.Button_Signatur_entfernen.Location = New System.Drawing.Point(303, 82)
        Me.Button_Signatur_entfernen.Name = "Button_Signatur_entfernen"
        Me.Button_Signatur_entfernen.Size = New System.Drawing.Size(75, 23)
        Me.Button_Signatur_entfernen.TabIndex = 7
        Me.Button_Signatur_entfernen.Text = "Entfernen"
        Me.Button_Signatur_entfernen.UseVisualStyleBackColor = True
        '
        'Button_Signatur_mit
        '
        Me.Button_Signatur_mit.ImageKey = "mitUhr.bmp"
        Me.Button_Signatur_mit.ImageList = Me.ImageList_Buttons
        Me.Button_Signatur_mit.Location = New System.Drawing.Point(303, 53)
        Me.Button_Signatur_mit.Name = "Button_Signatur_mit"
        Me.Button_Signatur_mit.Size = New System.Drawing.Size(37, 23)
        Me.Button_Signatur_mit.TabIndex = 6
        Me.Button_Signatur_mit.UseVisualStyleBackColor = True
        '
        'Button_Signatur_hinzufuegen
        '
        Me.Button_Signatur_hinzufuegen.Location = New System.Drawing.Point(303, 24)
        Me.Button_Signatur_hinzufuegen.Name = "Button_Signatur_hinzufuegen"
        Me.Button_Signatur_hinzufuegen.Size = New System.Drawing.Size(75, 23)
        Me.Button_Signatur_hinzufuegen.TabIndex = 5
        Me.Button_Signatur_hinzufuegen.Text = "Hinzufügen"
        Me.Button_Signatur_hinzufuegen.UseVisualStyleBackColor = True
        '
        'PB_Signatur
        '
        Me.PB_Signatur.BackColor = System.Drawing.Color.White
        Me.PB_Signatur.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.PB_Signatur.Location = New System.Drawing.Point(9, 14)
        Me.PB_Signatur.Name = "PB_Signatur"
        Me.PB_Signatur.Size = New System.Drawing.Size(275, 101)
        Me.PB_Signatur.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_Signatur.TabIndex = 4
        Me.PB_Signatur.TabStop = False
        '
        'Button_Cancel
        '
        Me.Button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Cancel.Location = New System.Drawing.Point(209, 372)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.Button_Cancel.TabIndex = 4
        Me.Button_Cancel.Text = "Abbrechen"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'Button_OK
        '
        Me.Button_OK.Location = New System.Drawing.Point(128, 372)
        Me.Button_OK.Name = "Button_OK"
        Me.Button_OK.Size = New System.Drawing.Size(75, 23)
        Me.Button_OK.TabIndex = 3
        Me.Button_OK.Text = "OK"
        Me.Button_OK.UseVisualStyleBackColor = True
        '
        'OFD_Logo
        '
        Me.OFD_Logo.Filter = "*.bmp; *.jpg;*.jpeg;*.gif;*.tif|*.bmp; *.jpg;*.jpeg;*.gif;*.tif"
        Me.OFD_Logo.Title = "Logo einlesen"
        '
        'OFD_Signatur
        '
        Me.OFD_Signatur.Filter = "*.bmp; *.jpg;*.jpeg;*.gif;*.tif|*.bmp; *.jpg;*.jpeg;*.gif;*.tif"
        Me.OFD_Signatur.Title = "Signatur einlesen"
        '
        'Form_Aussteller
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(397, 403)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_OK)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "Form_Aussteller"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Aussteller"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.PB_Logo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.PB_Signatur, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TB_Ort As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TB_PLZ As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TB_Nr As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TB_Strasse As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TB_Zusatz As System.Windows.Forms.TextBox
    Friend WithEvents TB_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Logo_entfernen As System.Windows.Forms.Button
    Friend WithEvents Button_Logo_mit As System.Windows.Forms.Button
    Friend WithEvents Button_Logo_hinzufuegen As System.Windows.Forms.Button
    Friend WithEvents PB_Logo As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Signatur_entfernen As System.Windows.Forms.Button
    Friend WithEvents Button_Signatur_mit As System.Windows.Forms.Button
    Friend WithEvents Button_Signatur_hinzufuegen As System.Windows.Forms.Button
    Friend WithEvents PB_Signatur As System.Windows.Forms.PictureBox
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_OK As System.Windows.Forms.Button
    Friend WithEvents OFD_Logo As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OFD_Signatur As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button_Logo_gegen As System.Windows.Forms.Button
    Friend WithEvents ImageList_Buttons As System.Windows.Forms.ImageList
    Friend WithEvents Button_Signatur_gegen As System.Windows.Forms.Button
End Class
