<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOpenFile
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
        Dim DataGridViewCellStyle63 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle55 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle56 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle57 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle58 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle59 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle60 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle61 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle62 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Online")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("In Store")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Retail", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2})
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Wholesale")
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOpenFile))
        Dim DataGridViewCellStyle72 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle64 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle65 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle66 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle67 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle68 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle69 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle70 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle71 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle81 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle73 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle74 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle75 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle76 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle77 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle78 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle79 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle80 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker()
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker()
        Me.pgLoading = New System.Windows.Forms.ProgressBar()
        Me.btnExportPdf = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.cboCollection = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cboColor = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cboSize = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cboProduct = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cboSubCategory = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboMainCategory = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cboActualBranch = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tabBody = New System.Windows.Forms.TabControl()
        Me.tabMain = New System.Windows.Forms.TabPage()
        Me.mainGridView = New System.Windows.Forms.DataGridView()
        Me.tabMainReport = New System.Windows.Forms.TabPage()
        Me.btnGenerateReport = New System.Windows.Forms.Button()
        Me.ReportGrid = New System.Windows.Forms.DataGridView()
        Me.colNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colBranchType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCategory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSubCategory = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCollection = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSOH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colAmount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colUnit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCost = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colRRP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colGp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colDiscount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colGpPercent = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.treeMain = New System.Windows.Forms.TreeView()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.tabProductReport = New System.Windows.Forms.TabPage()
        Me.ProductReportGrid = New System.Windows.Forms.DataGridView()
        Me.colPNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colProduct = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPSOH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPAmount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPUnit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPCost = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPRRP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPGp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn12 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.tabSearchResult = New System.Windows.Forms.TabPage()
        Me.SearchResultGrid = New System.Windows.Forms.DataGridView()
        Me.colSNo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn8 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn9 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn10 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSSOH = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSAmount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSSaleQty = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSCost = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSRRP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colSGP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn19 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn20 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.butBrowse = New System.Windows.Forms.Button()
        Me.butOpen = New System.Windows.Forms.Button()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.colNoResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colActualBranchResult = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.chkRefund = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.tabBody.SuspendLayout()
        Me.tabMain.SuspendLayout()
        CType(Me.mainGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabMainReport.SuspendLayout()
        CType(Me.ReportGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabProductReport.SuspendLayout()
        CType(Me.ProductReportGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabSearchResult.SuspendLayout()
        CType(Me.SearchResultGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.chkRefund)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.dtpToDate)
        Me.Panel1.Controls.Add(Me.dtpFromDate)
        Me.Panel1.Controls.Add(Me.pgLoading)
        Me.Panel1.Controls.Add(Me.btnExportPdf)
        Me.Panel1.Controls.Add(Me.btnExportExcel)
        Me.Panel1.Controls.Add(Me.btnSearch)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.tabBody)
        Me.Panel1.Controls.Add(Me.butBrowse)
        Me.Panel1.Controls.Add(Me.butOpen)
        Me.Panel1.Controls.Add(Me.txtPath)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1683, 848)
        Me.Panel1.TabIndex = 0
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(1002, 33)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(46, 13)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "To Date"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(775, 31)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 13)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "From Date"
        '
        'dtpToDate
        '
        Me.dtpToDate.CustomFormat = "MM/dd/yyyy"
        Me.dtpToDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpToDate.Location = New System.Drawing.Point(1060, 28)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(134, 22)
        Me.dtpToDate.TabIndex = 25
        '
        'dtpFromDate
        '
        Me.dtpFromDate.CustomFormat = "MM/dd/yyyy"
        Me.dtpFromDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFromDate.Location = New System.Drawing.Point(840, 28)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(152, 22)
        Me.dtpFromDate.TabIndex = 24
        '
        'pgLoading
        '
        Me.pgLoading.Location = New System.Drawing.Point(0, 838)
        Me.pgLoading.Name = "pgLoading"
        Me.pgLoading.Size = New System.Drawing.Size(1683, 10)
        Me.pgLoading.TabIndex = 23
        Me.pgLoading.Value = 20
        Me.pgLoading.Visible = False
        '
        'btnExportPdf
        '
        Me.btnExportPdf.Location = New System.Drawing.Point(1576, 22)
        Me.btnExportPdf.Name = "btnExportPdf"
        Me.btnExportPdf.Size = New System.Drawing.Size(88, 31)
        Me.btnExportPdf.TabIndex = 21
        Me.btnExportPdf.Text = "Export to PDF"
        Me.btnExportPdf.UseVisualStyleBackColor = True
        '
        'btnExportExcel
        '
        Me.btnExportExcel.Location = New System.Drawing.Point(1474, 23)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(95, 31)
        Me.btnExportExcel.TabIndex = 20
        Me.btnExportExcel.Text = "Export to Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = True
        '
        'btnSearch
        '
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(1589, 66)
        Me.btnSearch.Margin = New System.Windows.Forms.Padding(3, 3, 3, 1)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 49)
        Me.btnSearch.TabIndex = 18
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cboCollection)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.cboColor)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.cboSize)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.cboProduct)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.cboSubCategory)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.cboMainCategory)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.cboActualBranch)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(19, 60)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1564, 55)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter"
        '
        'cboCollection
        '
        Me.cboCollection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCollection.FormattingEnabled = True
        Me.cboCollection.Location = New System.Drawing.Point(820, 21)
        Me.cboCollection.Name = "cboCollection"
        Me.cboCollection.Size = New System.Drawing.Size(153, 24)
        Me.cboCollection.TabIndex = 26
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(745, 27)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(67, 16)
        Me.Label8.TabIndex = 25
        Me.Label8.Text = "Collection"
        '
        'cboColor
        '
        Me.cboColor.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboColor.FormattingEnabled = True
        Me.cboColor.Location = New System.Drawing.Point(1447, 20)
        Me.cboColor.Name = "cboColor"
        Me.cboColor.Size = New System.Drawing.Size(98, 24)
        Me.cboColor.TabIndex = 24
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(1401, 26)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 23
        Me.Label7.Text = "Color"
        '
        'cboSize
        '
        Me.cboSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSize.FormattingEnabled = True
        Me.cboSize.Location = New System.Drawing.Point(1304, 20)
        Me.cboSize.Name = "cboSize"
        Me.cboSize.Size = New System.Drawing.Size(80, 24)
        Me.cboSize.TabIndex = 22
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(1264, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(34, 16)
        Me.Label6.TabIndex = 21
        Me.Label6.Text = "Size"
        '
        'cboProduct
        '
        Me.cboProduct.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboProduct.FormattingEnabled = True
        Me.cboProduct.Location = New System.Drawing.Point(1041, 20)
        Me.cboProduct.Name = "cboProduct"
        Me.cboProduct.Size = New System.Drawing.Size(210, 24)
        Me.cboProduct.TabIndex = 20
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(983, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(54, 16)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Product"
        '
        'cboSubCategory
        '
        Me.cboSubCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSubCategory.FormattingEnabled = True
        Me.cboSubCategory.Location = New System.Drawing.Point(583, 20)
        Me.cboSubCategory.Name = "cboSubCategory"
        Me.cboSubCategory.Size = New System.Drawing.Size(153, 24)
        Me.cboSubCategory.TabIndex = 18
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(490, 26)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(87, 16)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "SubCategory"
        '
        'cboMainCategory
        '
        Me.cboMainCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboMainCategory.FormattingEnabled = True
        Me.cboMainCategory.Location = New System.Drawing.Point(335, 20)
        Me.cboMainCategory.Name = "cboMainCategory"
        Me.cboMainCategory.Size = New System.Drawing.Size(141, 24)
        Me.cboMainCategory.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(275, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 16)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Catgory"
        '
        'cboActualBranch
        '
        Me.cboActualBranch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboActualBranch.FormattingEnabled = True
        Me.cboActualBranch.Location = New System.Drawing.Point(113, 20)
        Me.cboActualBranch.Name = "cboActualBranch"
        Me.cboActualBranch.Size = New System.Drawing.Size(142, 24)
        Me.cboActualBranch.TabIndex = 14
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(17, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Actual Branch: "
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(63, 16)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = " File Path"
        '
        'tabBody
        '
        Me.tabBody.Controls.Add(Me.tabMain)
        Me.tabBody.Controls.Add(Me.tabMainReport)
        Me.tabBody.Controls.Add(Me.tabProductReport)
        Me.tabBody.Controls.Add(Me.tabSearchResult)
        Me.tabBody.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabBody.Location = New System.Drawing.Point(19, 121)
        Me.tabBody.Name = "tabBody"
        Me.tabBody.SelectedIndex = 0
        Me.tabBody.Size = New System.Drawing.Size(1652, 715)
        Me.tabBody.TabIndex = 14
        '
        'tabMain
        '
        Me.tabMain.Controls.Add(Me.mainGridView)
        Me.tabMain.Location = New System.Drawing.Point(4, 25)
        Me.tabMain.Name = "tabMain"
        Me.tabMain.Padding = New System.Windows.Forms.Padding(3)
        Me.tabMain.Size = New System.Drawing.Size(1644, 686)
        Me.tabMain.TabIndex = 0
        Me.tabMain.Text = "Main Sheet"
        Me.tabMain.UseVisualStyleBackColor = True
        '
        'mainGridView
        '
        Me.mainGridView.AllowUserToAddRows = False
        Me.mainGridView.AllowUserToDeleteRows = False
        Me.mainGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.mainGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.mainGridView.Location = New System.Drawing.Point(3, 3)
        Me.mainGridView.Name = "mainGridView"
        Me.mainGridView.ReadOnly = True
        Me.mainGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.mainGridView.Size = New System.Drawing.Size(1638, 680)
        Me.mainGridView.TabIndex = 4
        '
        'tabMainReport
        '
        Me.tabMainReport.Controls.Add(Me.btnGenerateReport)
        Me.tabMainReport.Controls.Add(Me.ReportGrid)
        Me.tabMainReport.Controls.Add(Me.treeMain)
        Me.tabMainReport.Location = New System.Drawing.Point(4, 25)
        Me.tabMainReport.Name = "tabMainReport"
        Me.tabMainReport.Size = New System.Drawing.Size(1644, 686)
        Me.tabMainReport.TabIndex = 2
        Me.tabMainReport.Text = "Main Report"
        Me.tabMainReport.UseVisualStyleBackColor = True
        '
        'btnGenerateReport
        '
        Me.btnGenerateReport.Location = New System.Drawing.Point(57, 646)
        Me.btnGenerateReport.Name = "btnGenerateReport"
        Me.btnGenerateReport.Size = New System.Drawing.Size(140, 37)
        Me.btnGenerateReport.TabIndex = 2
        Me.btnGenerateReport.Text = "Generate Report"
        Me.btnGenerateReport.UseVisualStyleBackColor = True
        '
        'ReportGrid
        '
        Me.ReportGrid.AllowUserToAddRows = False
        Me.ReportGrid.AllowUserToDeleteRows = False
        Me.ReportGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ReportGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colNo, Me.colBranchType, Me.colCategory, Me.colSubCategory, Me.colCollection, Me.colSOH, Me.colAmount, Me.colUnit, Me.colCost, Me.colRRP, Me.colGp, Me.colDiscount, Me.colGpPercent})
        Me.ReportGrid.Location = New System.Drawing.Point(252, 3)
        Me.ReportGrid.Name = "ReportGrid"
        Me.ReportGrid.ReadOnly = True
        Me.ReportGrid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        DataGridViewCellStyle63.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle63.BackColor = System.Drawing.SystemColors.ScrollBar
        DataGridViewCellStyle63.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle63.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle63.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle63.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle63.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ReportGrid.RowHeadersDefaultCellStyle = DataGridViewCellStyle63
        Me.ReportGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.ReportGrid.Size = New System.Drawing.Size(1389, 637)
        Me.ReportGrid.TabIndex = 1
        '
        'colNo
        '
        Me.colNo.Frozen = True
        Me.colNo.HeaderText = "No"
        Me.colNo.Name = "colNo"
        Me.colNo.ReadOnly = True
        Me.colNo.Width = 50
        '
        'colBranchType
        '
        Me.colBranchType.Frozen = True
        Me.colBranchType.HeaderText = "Actual Branch"
        Me.colBranchType.Name = "colBranchType"
        Me.colBranchType.ReadOnly = True
        Me.colBranchType.Width = 120
        '
        'colCategory
        '
        Me.colCategory.Frozen = True
        Me.colCategory.HeaderText = "Category"
        Me.colCategory.Name = "colCategory"
        Me.colCategory.ReadOnly = True
        Me.colCategory.Width = 120
        '
        'colSubCategory
        '
        Me.colSubCategory.Frozen = True
        Me.colSubCategory.HeaderText = "Sub Category"
        Me.colSubCategory.Name = "colSubCategory"
        Me.colSubCategory.ReadOnly = True
        Me.colSubCategory.Width = 120
        '
        'colCollection
        '
        Me.colCollection.Frozen = True
        Me.colCollection.HeaderText = "Collection"
        Me.colCollection.Name = "colCollection"
        Me.colCollection.ReadOnly = True
        Me.colCollection.Width = 120
        '
        'colSOH
        '
        DataGridViewCellStyle55.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSOH.DefaultCellStyle = DataGridViewCellStyle55
        Me.colSOH.HeaderText = "SOH"
        Me.colSOH.Name = "colSOH"
        Me.colSOH.ReadOnly = True
        Me.colSOH.Width = 70
        '
        'colAmount
        '
        DataGridViewCellStyle56.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colAmount.DefaultCellStyle = DataGridViewCellStyle56
        Me.colAmount.HeaderText = "Total Sales for last 30 days AUD"
        Me.colAmount.Name = "colAmount"
        Me.colAmount.ReadOnly = True
        Me.colAmount.Width = 140
        '
        'colUnit
        '
        DataGridViewCellStyle57.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colUnit.DefaultCellStyle = DataGridViewCellStyle57
        Me.colUnit.HeaderText = "Total Sales for last 30 days Units"
        Me.colUnit.Name = "colUnit"
        Me.colUnit.ReadOnly = True
        Me.colUnit.Width = 140
        '
        'colCost
        '
        DataGridViewCellStyle58.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colCost.DefaultCellStyle = DataGridViewCellStyle58
        Me.colCost.HeaderText = "Cost of sale AUD"
        Me.colCost.Name = "colCost"
        Me.colCost.ReadOnly = True
        Me.colCost.Width = 90
        '
        'colRRP
        '
        DataGridViewCellStyle59.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colRRP.DefaultCellStyle = DataGridViewCellStyle59
        Me.colRRP.HeaderText = "RRP Sales Value AUD"
        Me.colRRP.Name = "colRRP"
        Me.colRRP.ReadOnly = True
        '
        'colGp
        '
        DataGridViewCellStyle60.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colGp.DefaultCellStyle = DataGridViewCellStyle60
        Me.colGp.HeaderText = "GP AUD"
        Me.colGp.Name = "colGp"
        Me.colGp.ReadOnly = True
        Me.colGp.Width = 70
        '
        'colDiscount
        '
        DataGridViewCellStyle61.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colDiscount.DefaultCellStyle = DataGridViewCellStyle61
        Me.colDiscount.HeaderText = "Discount ( % )"
        Me.colDiscount.Name = "colDiscount"
        Me.colDiscount.ReadOnly = True
        Me.colDiscount.Width = 115
        '
        'colGpPercent
        '
        DataGridViewCellStyle62.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colGpPercent.DefaultCellStyle = DataGridViewCellStyle62
        Me.colGpPercent.HeaderText = "GP ( % )"
        Me.colGpPercent.Name = "colGpPercent"
        Me.colGpPercent.ReadOnly = True
        Me.colGpPercent.Width = 90
        '
        'treeMain
        '
        Me.treeMain.ImageIndex = 0
        Me.treeMain.ImageList = Me.ImageList1
        Me.treeMain.Location = New System.Drawing.Point(0, 3)
        Me.treeMain.Name = "treeMain"
        TreeNode1.Name = "treeOnline"
        TreeNode1.SelectedImageIndex = 1
        TreeNode1.Text = "Online"
        TreeNode2.Name = "treeStore"
        TreeNode2.SelectedImageIndex = 1
        TreeNode2.Text = "In Store"
        TreeNode3.Name = "treeRetail"
        TreeNode3.SelectedImageIndex = 1
        TreeNode3.Text = "Retail"
        TreeNode4.Name = "treeWholesale"
        TreeNode4.SelectedImageIndex = 1
        TreeNode4.Text = "Wholesale"
        Me.treeMain.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode3, TreeNode4})
        Me.treeMain.SelectedImageIndex = 0
        Me.treeMain.Size = New System.Drawing.Size(251, 637)
        Me.treeMain.TabIndex = 0
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "folder-colose.png")
        Me.ImageList1.Images.SetKeyName(1, "foler-open.png")
        '
        'tabProductReport
        '
        Me.tabProductReport.Controls.Add(Me.ProductReportGrid)
        Me.tabProductReport.Location = New System.Drawing.Point(4, 25)
        Me.tabProductReport.Name = "tabProductReport"
        Me.tabProductReport.Padding = New System.Windows.Forms.Padding(3)
        Me.tabProductReport.Size = New System.Drawing.Size(1644, 686)
        Me.tabProductReport.TabIndex = 1
        Me.tabProductReport.Text = "Product Report"
        Me.tabProductReport.UseVisualStyleBackColor = True
        '
        'ProductReportGrid
        '
        Me.ProductReportGrid.AllowUserToAddRows = False
        Me.ProductReportGrid.AllowUserToDeleteRows = False
        Me.ProductReportGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.ProductReportGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colPNo, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.colProduct, Me.colPSOH, Me.colPAmount, Me.colPUnit, Me.colPCost, Me.colPRRP, Me.colPGp, Me.DataGridViewTextBoxColumn12, Me.DataGridViewTextBoxColumn13})
        Me.ProductReportGrid.Location = New System.Drawing.Point(0, 0)
        Me.ProductReportGrid.Name = "ProductReportGrid"
        Me.ProductReportGrid.ReadOnly = True
        Me.ProductReportGrid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        DataGridViewCellStyle72.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle72.BackColor = System.Drawing.SystemColors.ScrollBar
        DataGridViewCellStyle72.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle72.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle72.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle72.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle72.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.ProductReportGrid.RowHeadersDefaultCellStyle = DataGridViewCellStyle72
        Me.ProductReportGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.ProductReportGrid.Size = New System.Drawing.Size(1644, 686)
        Me.ProductReportGrid.TabIndex = 2
        '
        'colPNo
        '
        Me.colPNo.Frozen = True
        Me.colPNo.HeaderText = "No"
        Me.colPNo.Name = "colPNo"
        Me.colPNo.ReadOnly = True
        Me.colPNo.Width = 50
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.Frozen = True
        Me.DataGridViewTextBoxColumn2.HeaderText = "Actual Branch"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.ReadOnly = True
        Me.DataGridViewTextBoxColumn2.Width = 120
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.Frozen = True
        Me.DataGridViewTextBoxColumn3.HeaderText = "Category"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.ReadOnly = True
        Me.DataGridViewTextBoxColumn3.Width = 120
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.Frozen = True
        Me.DataGridViewTextBoxColumn4.HeaderText = "Sub Category"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        Me.DataGridViewTextBoxColumn4.ReadOnly = True
        Me.DataGridViewTextBoxColumn4.Width = 120
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.Frozen = True
        Me.DataGridViewTextBoxColumn5.HeaderText = "Collection"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.ReadOnly = True
        Me.DataGridViewTextBoxColumn5.Width = 120
        '
        'colProduct
        '
        Me.colProduct.HeaderText = "Product"
        Me.colProduct.Name = "colProduct"
        Me.colProduct.ReadOnly = True
        Me.colProduct.Width = 250
        '
        'colPSOH
        '
        DataGridViewCellStyle64.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPSOH.DefaultCellStyle = DataGridViewCellStyle64
        Me.colPSOH.HeaderText = "SOH"
        Me.colPSOH.Name = "colPSOH"
        Me.colPSOH.ReadOnly = True
        Me.colPSOH.Width = 70
        '
        'colPAmount
        '
        DataGridViewCellStyle65.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPAmount.DefaultCellStyle = DataGridViewCellStyle65
        Me.colPAmount.HeaderText = "Total Sales for last 30 days AUD"
        Me.colPAmount.Name = "colPAmount"
        Me.colPAmount.ReadOnly = True
        Me.colPAmount.Width = 140
        '
        'colPUnit
        '
        DataGridViewCellStyle66.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPUnit.DefaultCellStyle = DataGridViewCellStyle66
        Me.colPUnit.HeaderText = "Total Sales for last 30 days Units"
        Me.colPUnit.Name = "colPUnit"
        Me.colPUnit.ReadOnly = True
        Me.colPUnit.Width = 140
        '
        'colPCost
        '
        DataGridViewCellStyle67.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPCost.DefaultCellStyle = DataGridViewCellStyle67
        Me.colPCost.HeaderText = "Cost of sale AUD"
        Me.colPCost.Name = "colPCost"
        Me.colPCost.ReadOnly = True
        Me.colPCost.Width = 90
        '
        'colPRRP
        '
        DataGridViewCellStyle68.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPRRP.DefaultCellStyle = DataGridViewCellStyle68
        Me.colPRRP.HeaderText = "RRP Sales Value AUD"
        Me.colPRRP.Name = "colPRRP"
        Me.colPRRP.ReadOnly = True
        '
        'colPGp
        '
        DataGridViewCellStyle69.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colPGp.DefaultCellStyle = DataGridViewCellStyle69
        Me.colPGp.HeaderText = "GP AUD"
        Me.colPGp.Name = "colPGp"
        Me.colPGp.ReadOnly = True
        Me.colPGp.Width = 70
        '
        'DataGridViewTextBoxColumn12
        '
        DataGridViewCellStyle70.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn12.DefaultCellStyle = DataGridViewCellStyle70
        Me.DataGridViewTextBoxColumn12.HeaderText = "Discount ( % )"
        Me.DataGridViewTextBoxColumn12.Name = "DataGridViewTextBoxColumn12"
        Me.DataGridViewTextBoxColumn12.ReadOnly = True
        Me.DataGridViewTextBoxColumn12.Width = 115
        '
        'DataGridViewTextBoxColumn13
        '
        DataGridViewCellStyle71.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn13.DefaultCellStyle = DataGridViewCellStyle71
        Me.DataGridViewTextBoxColumn13.HeaderText = "GP ( % )"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        Me.DataGridViewTextBoxColumn13.Width = 90
        '
        'tabSearchResult
        '
        Me.tabSearchResult.Controls.Add(Me.SearchResultGrid)
        Me.tabSearchResult.Location = New System.Drawing.Point(4, 25)
        Me.tabSearchResult.Name = "tabSearchResult"
        Me.tabSearchResult.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSearchResult.Size = New System.Drawing.Size(1644, 686)
        Me.tabSearchResult.TabIndex = 3
        Me.tabSearchResult.Text = "Search Result"
        Me.tabSearchResult.UseVisualStyleBackColor = True
        '
        'SearchResultGrid
        '
        Me.SearchResultGrid.AllowUserToAddRows = False
        Me.SearchResultGrid.AllowUserToDeleteRows = False
        Me.SearchResultGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.SearchResultGrid.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colSNo, Me.DataGridViewTextBoxColumn6, Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn8, Me.DataGridViewTextBoxColumn9, Me.DataGridViewTextBoxColumn10, Me.colSSOH, Me.colSAmount, Me.colSSaleQty, Me.colSCost, Me.colSRRP, Me.colSGP, Me.DataGridViewTextBoxColumn19, Me.DataGridViewTextBoxColumn20})
        Me.SearchResultGrid.Location = New System.Drawing.Point(8, 8)
        Me.SearchResultGrid.Name = "SearchResultGrid"
        Me.SearchResultGrid.ReadOnly = True
        Me.SearchResultGrid.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        DataGridViewCellStyle81.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle81.BackColor = System.Drawing.SystemColors.ScrollBar
        DataGridViewCellStyle81.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle81.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle81.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle81.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle81.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.SearchResultGrid.RowHeadersDefaultCellStyle = DataGridViewCellStyle81
        Me.SearchResultGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.SearchResultGrid.Size = New System.Drawing.Size(1644, 686)
        Me.SearchResultGrid.TabIndex = 3
        '
        'colSNo
        '
        Me.colSNo.Frozen = True
        Me.colSNo.HeaderText = "No"
        Me.colSNo.Name = "colSNo"
        Me.colSNo.ReadOnly = True
        Me.colSNo.Width = 50
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.Frozen = True
        Me.DataGridViewTextBoxColumn6.HeaderText = "Actual Branch"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.ReadOnly = True
        Me.DataGridViewTextBoxColumn6.Width = 120
        '
        'DataGridViewTextBoxColumn7
        '
        Me.DataGridViewTextBoxColumn7.Frozen = True
        Me.DataGridViewTextBoxColumn7.HeaderText = "Category"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        Me.DataGridViewTextBoxColumn7.Width = 120
        '
        'DataGridViewTextBoxColumn8
        '
        Me.DataGridViewTextBoxColumn8.Frozen = True
        Me.DataGridViewTextBoxColumn8.HeaderText = "Sub Category"
        Me.DataGridViewTextBoxColumn8.Name = "DataGridViewTextBoxColumn8"
        Me.DataGridViewTextBoxColumn8.ReadOnly = True
        Me.DataGridViewTextBoxColumn8.Width = 120
        '
        'DataGridViewTextBoxColumn9
        '
        Me.DataGridViewTextBoxColumn9.Frozen = True
        Me.DataGridViewTextBoxColumn9.HeaderText = "Collection"
        Me.DataGridViewTextBoxColumn9.Name = "DataGridViewTextBoxColumn9"
        Me.DataGridViewTextBoxColumn9.ReadOnly = True
        Me.DataGridViewTextBoxColumn9.Width = 120
        '
        'DataGridViewTextBoxColumn10
        '
        Me.DataGridViewTextBoxColumn10.HeaderText = "Product"
        Me.DataGridViewTextBoxColumn10.Name = "DataGridViewTextBoxColumn10"
        Me.DataGridViewTextBoxColumn10.ReadOnly = True
        Me.DataGridViewTextBoxColumn10.Width = 250
        '
        'colSSOH
        '
        DataGridViewCellStyle73.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSSOH.DefaultCellStyle = DataGridViewCellStyle73
        Me.colSSOH.HeaderText = "SOH"
        Me.colSSOH.Name = "colSSOH"
        Me.colSSOH.ReadOnly = True
        Me.colSSOH.Width = 70
        '
        'colSAmount
        '
        DataGridViewCellStyle74.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSAmount.DefaultCellStyle = DataGridViewCellStyle74
        Me.colSAmount.HeaderText = "Total Sales for last 30 days AUD"
        Me.colSAmount.Name = "colSAmount"
        Me.colSAmount.ReadOnly = True
        Me.colSAmount.Width = 140
        '
        'colSSaleQty
        '
        DataGridViewCellStyle75.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSSaleQty.DefaultCellStyle = DataGridViewCellStyle75
        Me.colSSaleQty.HeaderText = "Total Sales for last 30 days Units"
        Me.colSSaleQty.Name = "colSSaleQty"
        Me.colSSaleQty.ReadOnly = True
        Me.colSSaleQty.Width = 140
        '
        'colSCost
        '
        DataGridViewCellStyle76.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSCost.DefaultCellStyle = DataGridViewCellStyle76
        Me.colSCost.HeaderText = "Cost of sale AUD"
        Me.colSCost.Name = "colSCost"
        Me.colSCost.ReadOnly = True
        Me.colSCost.Width = 90
        '
        'colSRRP
        '
        DataGridViewCellStyle77.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSRRP.DefaultCellStyle = DataGridViewCellStyle77
        Me.colSRRP.HeaderText = "RRP Sales Value AUD"
        Me.colSRRP.Name = "colSRRP"
        Me.colSRRP.ReadOnly = True
        '
        'colSGP
        '
        DataGridViewCellStyle78.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.colSGP.DefaultCellStyle = DataGridViewCellStyle78
        Me.colSGP.HeaderText = "GP AUD"
        Me.colSGP.Name = "colSGP"
        Me.colSGP.ReadOnly = True
        Me.colSGP.Width = 70
        '
        'DataGridViewTextBoxColumn19
        '
        DataGridViewCellStyle79.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn19.DefaultCellStyle = DataGridViewCellStyle79
        Me.DataGridViewTextBoxColumn19.HeaderText = "Discount ( % )"
        Me.DataGridViewTextBoxColumn19.Name = "DataGridViewTextBoxColumn19"
        Me.DataGridViewTextBoxColumn19.ReadOnly = True
        Me.DataGridViewTextBoxColumn19.Width = 115
        '
        'DataGridViewTextBoxColumn20
        '
        DataGridViewCellStyle80.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.DataGridViewTextBoxColumn20.DefaultCellStyle = DataGridViewCellStyle80
        Me.DataGridViewTextBoxColumn20.HeaderText = "GP ( % )"
        Me.DataGridViewTextBoxColumn20.Name = "DataGridViewTextBoxColumn20"
        Me.DataGridViewTextBoxColumn20.ReadOnly = True
        Me.DataGridViewTextBoxColumn20.Width = 90
        '
        'butBrowse
        '
        Me.butBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butBrowse.Location = New System.Drawing.Point(728, 26)
        Me.butBrowse.Name = "butBrowse"
        Me.butBrowse.Size = New System.Drawing.Size(28, 24)
        Me.butBrowse.TabIndex = 2
        Me.butBrowse.Text = "..."
        Me.butBrowse.UseVisualStyleBackColor = True
        '
        'butOpen
        '
        Me.butOpen.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.butOpen.Location = New System.Drawing.Point(1214, 28)
        Me.butOpen.Name = "butOpen"
        Me.butOpen.Size = New System.Drawing.Size(75, 24)
        Me.butOpen.TabIndex = 1
        Me.butOpen.Text = "Refresh"
        Me.butOpen.UseVisualStyleBackColor = True
        '
        'txtPath
        '
        Me.txtPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPath.Location = New System.Drawing.Point(80, 27)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.ReadOnly = True
        Me.txtPath.Size = New System.Drawing.Size(643, 22)
        Me.txtPath.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'colNoResult
        '
        Me.colNoResult.HeaderText = "No"
        Me.colNoResult.Name = "colNoResult"
        Me.colNoResult.ReadOnly = True
        '
        'colActualBranchResult
        '
        Me.colActualBranchResult.HeaderText = "Actual Branch"
        Me.colActualBranchResult.Name = "colActualBranchResult"
        Me.colActualBranchResult.ReadOnly = True
        '
        'chkRefund
        '
        Me.chkRefund.AutoSize = True
        Me.chkRefund.Location = New System.Drawing.Point(1323, 35)
        Me.chkRefund.Name = "chkRefund"
        Me.chkRefund.Size = New System.Drawing.Size(132, 17)
        Me.chkRefund.TabIndex = 28
        Me.chkRefund.Text = "Show only refund data"
        Me.chkRefund.UseVisualStyleBackColor = True
        '
        'frmOpenFile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1683, 848)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmOpenFile"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Data Analyzer"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.tabBody.ResumeLayout(False)
        Me.tabMain.ResumeLayout(False)
        CType(Me.mainGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabMainReport.ResumeLayout(False)
        CType(Me.ReportGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabProductReport.ResumeLayout(False)
        CType(Me.ProductReportGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabSearchResult.ResumeLayout(False)
        CType(Me.SearchResultGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents butBrowse As Button
    Friend WithEvents butOpen As Button
    Friend WithEvents txtPath As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents tabBody As TabControl
    Friend WithEvents tabMain As TabPage
    Friend WithEvents mainGridView As DataGridView
    Friend WithEvents tabProductReport As TabPage
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents cboSubCategory As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents cboMainCategory As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents cboActualBranch As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents btnSearch As Button
    Friend WithEvents cboColor As ComboBox
    Friend WithEvents Label7 As Label
    Friend WithEvents cboSize As ComboBox
    Friend WithEvents Label6 As Label
    Friend WithEvents cboProduct As ComboBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tabMainReport As TabPage
    Friend WithEvents cboCollection As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents treeMain As TreeView
    Friend WithEvents ImageList1 As ImageList
    Friend WithEvents ReportGrid As DataGridView
    Friend WithEvents btnGenerateReport As Button
    Friend WithEvents colNoResult As DataGridViewTextBoxColumn
    Friend WithEvents colActualBranchResult As DataGridViewTextBoxColumn
    Friend WithEvents btnExportPdf As Button
    Friend WithEvents btnExportExcel As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents colNo As DataGridViewTextBoxColumn
    Friend WithEvents colBranchType As DataGridViewTextBoxColumn
    Friend WithEvents colCategory As DataGridViewTextBoxColumn
    Friend WithEvents colSubCategory As DataGridViewTextBoxColumn
    Friend WithEvents colCollection As DataGridViewTextBoxColumn
    Friend WithEvents colSOH As DataGridViewTextBoxColumn
    Friend WithEvents colAmount As DataGridViewTextBoxColumn
    Friend WithEvents colUnit As DataGridViewTextBoxColumn
    Friend WithEvents colCost As DataGridViewTextBoxColumn
    Friend WithEvents colRRP As DataGridViewTextBoxColumn
    Friend WithEvents colGp As DataGridViewTextBoxColumn
    Friend WithEvents colDiscount As DataGridViewTextBoxColumn
    Friend WithEvents colGpPercent As DataGridViewTextBoxColumn
    Friend WithEvents ProductReportGrid As DataGridView
    Friend WithEvents pgLoading As ProgressBar
    Friend WithEvents tabSearchResult As TabPage
    Friend WithEvents colPNo As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As DataGridViewTextBoxColumn
    Friend WithEvents colProduct As DataGridViewTextBoxColumn
    Friend WithEvents colPSOH As DataGridViewTextBoxColumn
    Friend WithEvents colPAmount As DataGridViewTextBoxColumn
    Friend WithEvents colPUnit As DataGridViewTextBoxColumn
    Friend WithEvents colPCost As DataGridViewTextBoxColumn
    Friend WithEvents colPRRP As DataGridViewTextBoxColumn
    Friend WithEvents colPGp As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn12 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As DataGridViewTextBoxColumn
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents dtpToDate As DateTimePicker
    Friend WithEvents dtpFromDate As DateTimePicker
    Friend WithEvents SearchResultGrid As DataGridView
    Friend WithEvents colSNo As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn7 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn8 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn9 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn10 As DataGridViewTextBoxColumn
    Friend WithEvents colSSOH As DataGridViewTextBoxColumn
    Friend WithEvents colSAmount As DataGridViewTextBoxColumn
    Friend WithEvents colSSaleQty As DataGridViewTextBoxColumn
    Friend WithEvents colSCost As DataGridViewTextBoxColumn
    Friend WithEvents colSRRP As DataGridViewTextBoxColumn
    Friend WithEvents colSGP As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn19 As DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn20 As DataGridViewTextBoxColumn
    Friend WithEvents chkRefund As CheckBox
End Class
