<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewDetail
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
        Dim DataGridViewCellStyle25 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle26 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle27 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle28 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle29 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle30 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ProductDetailGrid = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colActualBranch = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCategoryResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSubCategoryResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCollectionResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colProductResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSizeResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colColorResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCostResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSaleDateResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSaleQtyResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSaleAmountResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1.SuspendLayout()
        CType(Me.ProductDetailGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ProductDetailGrid)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1364, 751)
        Me.Panel1.TabIndex = 0
        '
        'ProductDetailGrid
        '
        Me.ProductDetailGrid.AllowUserToAddRows = False
        Me.ProductDetailGrid.AllowUserToDeleteRows = False
        Me.ProductDetailGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ProductDetailGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.colActualBranch, Me.colCategoryResult, Me.colSubCategoryResult, Me.colCollectionResult, Me.colProductResult, Me.colSizeResult, Me.colColorResult, Me.colCostResult, Me.colSaleDateResult, Me.colSaleQtyResult, Me.colSaleAmountResult})
        Me.ProductDetailGrid.Location = New System.Drawing.Point(0, 0)
        Me.ProductDetailGrid.Name = "ProductDetailGrid"
        Me.ProductDetailGrid.ReadOnly = True
        Me.ProductDetailGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.ProductDetailGrid.Size = New System.Drawing.Size(1363, 751)
        Me.ProductDetailGrid.TabIndex = 2
        '
        'Column1
        '
        Me.Column1.Frozen = True
        Me.Column1.HeaderText = "No"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 40
        '
        'colActualBranch
        '
        Me.colActualBranch.Frozen = True
        Me.colActualBranch.HeaderText = "Actual Branch"
        Me.colActualBranch.Name = "colActualBranch"
        Me.colActualBranch.ReadOnly = True
        Me.colActualBranch.Width = 120
        '
        'colCategoryResult
        '
        Me.colCategoryResult.Frozen = True
        Me.colCategoryResult.HeaderText = "Category"
        Me.colCategoryResult.Name = "colCategoryResult"
        Me.colCategoryResult.ReadOnly = True
        Me.colCategoryResult.Width = 120
        '
        'colSubCategoryResult
        '
        Me.colSubCategoryResult.Frozen = True
        Me.colSubCategoryResult.HeaderText = "Sub Category"
        Me.colSubCategoryResult.Name = "colSubCategoryResult"
        Me.colSubCategoryResult.ReadOnly = True
        Me.colSubCategoryResult.Width = 120
        '
        'colCollectionResult
        '
        Me.colCollectionResult.Frozen = True
        Me.colCollectionResult.HeaderText = "Collection"
        Me.colCollectionResult.Name = "colCollectionResult"
        Me.colCollectionResult.ReadOnly = True
        Me.colCollectionResult.Width = 120
        '
        'colProductResult
        '
        Me.colProductResult.Frozen = True
        Me.colProductResult.HeaderText = "Product"
        Me.colProductResult.Name = "colProductResult"
        Me.colProductResult.ReadOnly = True
        Me.colProductResult.Width = 300
        '
        'colSizeResult
        '
        DataGridViewCellStyle25.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.colSizeResult.DefaultCellStyle = DataGridViewCellStyle25
        Me.colSizeResult.HeaderText = "Size"
        Me.colSizeResult.Name = "colSizeResult"
        Me.colSizeResult.ReadOnly = True
        Me.colSizeResult.Width = 60
        '
        'colColorResult
        '
        DataGridViewCellStyle26.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.colColorResult.DefaultCellStyle = DataGridViewCellStyle26
        Me.colColorResult.HeaderText = "Color"
        Me.colColorResult.Name = "colColorResult"
        Me.colColorResult.ReadOnly = True
        Me.colColorResult.Width = 60
        '
        'colCostResult
        '
        DataGridViewCellStyle27.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colCostResult.DefaultCellStyle = DataGridViewCellStyle27
        Me.colCostResult.HeaderText = "Cost"
        Me.colCostResult.Name = "colCostResult"
        Me.colCostResult.ReadOnly = True
        Me.colCostResult.Width = 60
        '
        'colSaleDateResult
        '
        DataGridViewCellStyle28.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.colSaleDateResult.DefaultCellStyle = DataGridViewCellStyle28
        Me.colSaleDateResult.HeaderText = "Sale Date"
        Me.colSaleDateResult.Name = "colSaleDateResult"
        Me.colSaleDateResult.ReadOnly = True
        '
        'colSaleQtyResult
        '
        DataGridViewCellStyle29.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSaleQtyResult.DefaultCellStyle = DataGridViewCellStyle29
        Me.colSaleQtyResult.HeaderText = "Sale Qty"
        Me.colSaleQtyResult.Name = "colSaleQtyResult"
        Me.colSaleQtyResult.ReadOnly = True
        '
        'colSaleAmountResult
        '
        DataGridViewCellStyle30.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSaleAmountResult.DefaultCellStyle = DataGridViewCellStyle30
        Me.colSaleAmountResult.HeaderText = "Sale Amount"
        Me.colSaleAmountResult.Name = "colSaleAmountResult"
        Me.colSaleAmountResult.ReadOnly = True
        Me.colSaleAmountResult.Width = 110
        '
        'frmViewDetail
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1364, 751)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmViewDetail"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "View detail for product"
        Me.Panel1.ResumeLayout(False)
        CType(Me.ProductDetailGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents ProductDetailGrid As DataGridView
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents colActualBranch As DataGridViewTextBoxColumn
    Friend WithEvents colCategoryResult As DataGridViewTextBoxColumn
    Friend WithEvents colSubCategoryResult As DataGridViewTextBoxColumn
    Friend WithEvents colCollectionResult As DataGridViewTextBoxColumn
    Friend WithEvents colProductResult As DataGridViewTextBoxColumn
    Friend WithEvents colSizeResult As DataGridViewTextBoxColumn
    Friend WithEvents colColorResult As DataGridViewTextBoxColumn
    Friend WithEvents colCostResult As DataGridViewTextBoxColumn
    Friend WithEvents colSaleDateResult As DataGridViewTextBoxColumn
    Friend WithEvents colSaleQtyResult As DataGridViewTextBoxColumn
    Friend WithEvents colSaleAmountResult As DataGridViewTextBoxColumn
End Class
