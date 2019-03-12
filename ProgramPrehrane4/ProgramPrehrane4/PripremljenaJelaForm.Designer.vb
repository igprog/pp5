<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PripremljenaJelaForm
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
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NazivJelaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MasaPripremljenogJelaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PripremaJelaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NazivNamirniceDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.MasagDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TermickaObradaDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PripremljenaJelaBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.PP5DataSet = New ProgramPrehrane4.PP5DataSet()
        Me.PripremljenaJelaTableAdapter = New ProgramPrehrane4.PP5DataSetTableAdapters.PripremljenaJelaTableAdapter()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PripremljenaJelaBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PP5DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.NazivJelaDataGridViewTextBoxColumn, Me.MasaPripremljenogJelaDataGridViewTextBoxColumn, Me.PripremaJelaDataGridViewTextBoxColumn, Me.NazivNamirniceDataGridViewTextBoxColumn, Me.MasagDataGridViewTextBoxColumn, Me.TermickaObradaDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.PripremljenaJelaBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(30, 26)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(311, 160)
        Me.DataGridView1.TabIndex = 12
        Me.DataGridView1.Visible = False
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        '
        'NazivJelaDataGridViewTextBoxColumn
        '
        Me.NazivJelaDataGridViewTextBoxColumn.DataPropertyName = "NazivJela"
        Me.NazivJelaDataGridViewTextBoxColumn.HeaderText = "NazivJela"
        Me.NazivJelaDataGridViewTextBoxColumn.Name = "NazivJelaDataGridViewTextBoxColumn"
        '
        'MasaPripremljenogJelaDataGridViewTextBoxColumn
        '
        Me.MasaPripremljenogJelaDataGridViewTextBoxColumn.DataPropertyName = "MasaPripremljenogJela"
        Me.MasaPripremljenogJelaDataGridViewTextBoxColumn.HeaderText = "MasaPripremljenogJela"
        Me.MasaPripremljenogJelaDataGridViewTextBoxColumn.Name = "MasaPripremljenogJelaDataGridViewTextBoxColumn"
        '
        'PripremaJelaDataGridViewTextBoxColumn
        '
        Me.PripremaJelaDataGridViewTextBoxColumn.DataPropertyName = "PripremaJela"
        Me.PripremaJelaDataGridViewTextBoxColumn.HeaderText = "PripremaJela"
        Me.PripremaJelaDataGridViewTextBoxColumn.Name = "PripremaJelaDataGridViewTextBoxColumn"
        '
        'NazivNamirniceDataGridViewTextBoxColumn
        '
        Me.NazivNamirniceDataGridViewTextBoxColumn.DataPropertyName = "NazivNamirnice"
        Me.NazivNamirniceDataGridViewTextBoxColumn.HeaderText = "NazivNamirnice"
        Me.NazivNamirniceDataGridViewTextBoxColumn.Name = "NazivNamirniceDataGridViewTextBoxColumn"
        '
        'MasagDataGridViewTextBoxColumn
        '
        Me.MasagDataGridViewTextBoxColumn.DataPropertyName = "Masa_g"
        Me.MasagDataGridViewTextBoxColumn.HeaderText = "Masa_g"
        Me.MasagDataGridViewTextBoxColumn.Name = "MasagDataGridViewTextBoxColumn"
        '
        'TermickaObradaDataGridViewTextBoxColumn
        '
        Me.TermickaObradaDataGridViewTextBoxColumn.DataPropertyName = "TermickaObrada"
        Me.TermickaObradaDataGridViewTextBoxColumn.HeaderText = "TermickaObrada"
        Me.TermickaObradaDataGridViewTextBoxColumn.Name = "TermickaObradaDataGridViewTextBoxColumn"
        '
        'PripremljenaJelaBindingSource
        '
        Me.PripremljenaJelaBindingSource.DataMember = "PripremljenaJela"
        Me.PripremljenaJelaBindingSource.DataSource = Me.PP5DataSet
        '
        'PP5DataSet
        '
        Me.PP5DataSet.DataSetName = "PP5DataSet"
        Me.PP5DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'PripremljenaJelaTableAdapter
        '
        Me.PripremljenaJelaTableAdapter.ClearBeforeFill = True
        '
        'PripremljenaJelaForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(377, 215)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "PripremljenaJelaForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Pripremljeno jelo"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PripremljenaJelaBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PP5DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents PP5DataSet As ProgramPrehrane4.PP5DataSet
    Friend WithEvents PripremljenaJelaBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents PripremljenaJelaTableAdapter As ProgramPrehrane4.PP5DataSetTableAdapters.PripremljenaJelaTableAdapter
    Friend WithEvents IDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NazivJelaDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MasaPripremljenogJelaDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PripremaJelaDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NazivNamirniceDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents MasagDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TermickaObradaDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
