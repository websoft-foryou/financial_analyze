Option Explicit On
Imports System.IO
Imports ExcelDataReader
Imports Excel = Microsoft.Office.Interop.Excel

Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class frmOpenFile
    Dim tables As DataTableCollection

    Private Sub butBrowse_Click(sender As Object, e As EventArgs) Handles butBrowse.Click
        OpenFileDialog1.Filter = "Excel 97-2003|*.xls|Excel Workbook|*.xlsx|Excel Macro-Enabled Workbook|*.xlsm|All Files|*.*"
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.Title = "Open Excel File"

        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            Dim path As String = OpenFileDialog1.FileName
            Try
                txtPath.Text = path
                butOpen_Click(sender, e)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub butOpen_Click(sender As Object, e As EventArgs) Handles butOpen.Click
        Dim filepath As String = txtPath.Text
        Dim i As Integer

        mainGridView.DataSource = Nothing
        ReportGrid.Rows.Clear()
        ProductReportGrid.Rows.Clear()
        SearchResultGrid.Rows.Clear()

        If filepath = "" Then
            MsgBox("Please select excel file for analysis.", vbExclamation)
            butBrowse.Focus()
            Exit Sub
        End If

        Using Stream = File.Open(filepath, FileMode.Open, FileAccess.Read)
            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(Stream)
                Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                         .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                         .UseHeaderRow = True}})
                tables = result.Tables
                mainGridView.DataSource = tables("Sheet")
            End Using
        End Using

        cboActualBranch.Items.Clear()
        cboActualBranch.Items.Add("[All Branch]")

        cboMainCategory.Items.Clear()
        cboMainCategory.Items.Add("[All Categories]")

        cboSubCategory.Items.Clear()
        cboSubCategory.Items.Add("[All Sub Categories]")

        cboCollection.Items.Clear()
        cboCollection.Items.Add("[All Collections]")

        cboProduct.Items.Clear()
        cboProduct.Items.Add("[All Units]")

        cboSize.Items.Clear()
        cboSize.Items.Add("[All Sizes]")

        cboColor.Items.Clear()
        cboColor.Items.Add("[All Colors]")

        pgLoading.Visible = True
        pgLoading.Value = 0
        pgLoading.Maximum = mainGridView.Rows.Count - 1

        Dim startDate As String = dtpFromDate.Value.ToString("yyyy-MM-dd")
        Dim endDate As String = dtpToDate.Value.ToString("yyyy-MM-dd")

        For i = mainGridView.Rows.Count - 1 To 1 Step -1

            Try
                If mainGridView.Rows(i).Cells(14).Value Is Nothing Or mainGridView.Rows(i).Cells(14).Value.ToString = "" Then
                    mainGridView.Rows.Remove(mainGridView.Rows(i))
                Else
                    Dim saleDate As String = Date.ParseExact(mainGridView.Rows(i).Cells(14).Value, "dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo).ToString("yyyy-MM-dd")
                    If saleDate < startDate Or saleDate > endDate Then
                        mainGridView.Rows.Remove(mainGridView.Rows(i))
                    End If
                End If

                pgLoading.Value = pgLoading.Maximum - i
            Catch ex As Exception
                mainGridView.Rows.Remove(mainGridView.Rows(i))
                Continue For
            End Try
        Next

        pgLoading.Visible = False

        For i = 2 To mainGridView.Rows.Count - 1
            Try
                Dim strActualBranch As String = mainGridView.Rows(i).Cells(0).Value.ToString
                Dim strMainCategoryName As String = mainGridView.Rows(i).Cells(1).Value.ToString
                Dim strSubCategoryName As String = mainGridView.Rows(i).Cells(2).Value.ToString
                Dim strCollectionName As String = mainGridView.Rows(i).Cells(3).Value.ToString

                Dim strProductName As String = mainGridView.Rows(i).Cells(5).Value.ToString
                Dim strSizeName As String = mainGridView.Rows(i).Cells(8).Value.ToString
                Dim strColorName As String = mainGridView.Rows(i).Cells(9).Value.ToString
                Dim nSaleQty As Double = Val(mainGridView.Rows(i).Cells(15).Value.ToString)
                Dim nSalePrice As Double = Val(mainGridView.Rows(i).Cells(22).Value.ToString)

                'If nSaleQty = 0 And nSalePrice = 0 Then Continue For

                If IsExistItemList(cboActualBranch, strActualBranch) = False Then
                    cboActualBranch.Items.Add(strActualBranch)
                End If
                If IsExistItemList(cboMainCategory, strMainCategoryName) = False Then
                    cboMainCategory.Items.Add(strMainCategoryName)
                End If
                If IsExistItemList(cboSubCategory, strSubCategoryName) = False Then
                    cboSubCategory.Items.Add(strSubCategoryName)
                End If
                If IsExistItemList(cboCollection, strCollectionName) = False Then
                    cboCollection.Items.Add(strCollectionName)
                End If
                If IsExistItemList(cboProduct, strProductName) = False Then
                    cboProduct.Items.Add(strProductName)
                End If
                If IsExistItemList(cboSize, strSizeName) = False Then
                    cboSize.Items.Add(strSizeName)
                End If
                If IsExistItemList(cboColor, strColorName) = False Then
                    cboColor.Items.Add(strColorName)
                End If

                MakeCategoryTree(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName)

            Catch ex As Exception
                MsgBox("add category failure! =>" + i.ToString + Chr(13) + ex.Message)
            End Try
        Next

        cboActualBranch.SelectedIndex = 0
        cboMainCategory.SelectedIndex = 0
        cboSubCategory.SelectedIndex = 0
        cboCollection.SelectedIndex = 0
        cboProduct.SelectedIndex = 0
        cboSize.SelectedIndex = 0
        cboColor.SelectedIndex = 0

        tabBody.SelectedTab = tabMain

    End Sub

    Private Function IsExistItemList(ByVal objCategoryList As ComboBox, ByVal strItemName As String) As Boolean
        Dim i As Integer
        For i = 0 To objCategoryList.Items.Count - 1
            If strItemName = vbNull.ToString Then
                IsExistItemList = True
                Exit Function
            End If
            If strItemName = objCategoryList.Items.Item(i).ToString Then
                IsExistItemList = True
                Exit Function
            End If
        Next

        IsExistItemList = False
    End Function

    Private Sub MakeCategoryTree(ByVal strBranchName As String, ByVal strCategoryText As String, ByVal strSubCategoryText As String, ByVal strCollectionText As String)
        Dim treeBranchNode As TreeNode, mainCategoryNode As TreeNode, subCategoryNode As TreeNode, collectionNode As TreeNode

        If strBranchName = "Adrift HQ, QLD" Then                                                ' Retail online
            treeBranchNode = treeMain.Nodes("treeRetail").Nodes("treeOnline")
        ElseIf strBranchName = "USA Wholesale, QLD" Or strBranchName = "Wholesale, QLD" Then    ' Wholesale
            treeBranchNode = treeMain.Nodes("treeWholesale")
        Else                                                                                    ' Retail in store
            treeBranchNode = treeMain.Nodes("treeRetail").Nodes("treeStore")
        End If

        If strCategoryText = "" Then strCategoryText = "(Empty Main Category)"
        mainCategoryNode = GetChildNode(treeBranchNode, strCategoryText)
        If IsNothing(mainCategoryNode) Then
            mainCategoryNode = treeBranchNode.Nodes.Add("mainCategory_node", strCategoryText, 0)
            mainCategoryNode.SelectedImageIndex = 1
        End If

        If strSubCategoryText = "" Then strSubCategoryText = "(Empty Sub Category)"
        subCategoryNode = GetChildNode(mainCategoryNode, strSubCategoryText)
        If IsNothing(subCategoryNode) Then
            subCategoryNode = mainCategoryNode.Nodes.Add("subCategory_node", strSubCategoryText, 0)
            subCategoryNode.SelectedImageIndex = 1
        End If

        If strCollectionText = "" Then strCollectionText = "(Empty Collection)"
        collectionNode = GetChildNode(subCategoryNode, strCollectionText)
        If IsNothing(collectionNode) Then
            collectionNode = subCategoryNode.Nodes.Add("collection_node", strCollectionText, 0)
            collectionNode.SelectedImageIndex = 1
        End If
    End Sub

    Private Function GetChildNode(ByVal treeSearchNode As TreeNode, ByVal strNodeText As String) As TreeNode
        For Each nodeItem As TreeNode In treeSearchNode.Nodes
            If nodeItem.Text = strNodeText Then
                GetChildNode = nodeItem
                Exit Function
            End If
        Next
        GetChildNode = Nothing
    End Function

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        If IsNothing(cboActualBranch.SelectedItem) Then
            MsgBox("Please select the appropriate item for [Actual Branch].", vbExclamation)
            cboActualBranch.Focus()
            Exit Sub
        End If

        If IsNothing(cboMainCategory.SelectedItem) Then
            MsgBox("Please select the appropriate item for [Category].", vbExclamation)
            cboMainCategory.Focus()
            Exit Sub
        End If

        If IsNothing(cboSubCategory.SelectedItem) Then
            MsgBox("Please select the appropriate item for [Sub Category].", vbExclamation)
            cboSubCategory.Focus()
            Exit Sub
        End If

        If IsNothing(cboCollection.SelectedItem) Then
            cboCollection.Focus()
            MsgBox("Please select the appropriate item for [Collection].", vbExclamation)
            Exit Sub
        End If

        If IsNothing(cboProduct.SelectedItem) Then
            cboProduct.Focus()
            MsgBox("Please select the appropriate item for [Product].", vbExclamation)
            Exit Sub
        End If

        If IsNothing(cboSize.SelectedItem) Then
            cboSize.Focus()
            MsgBox("Please select the appropriate item for [Size].", vbExclamation)
            Exit Sub
        End If

        If IsNothing(cboColor.SelectedItem) Then
            cboColor.Focus()
            MsgBox("Please select the appropriate item for [Color].", vbExclamation)
            Exit Sub
        End If

        Dim strActualBranch As String = cboActualBranch.SelectedItem.ToString
        Dim strMainCategoryName As String = cboMainCategory.SelectedItem.ToString
        Dim strSubCategoryName As String = cboSubCategory.SelectedItem.ToString
        Dim strCollectionName As String = cboCollection.SelectedItem.ToString
        Dim strProductName As String = cboProduct.SelectedItem.ToString
        Dim strSizeName As String = cboSize.SelectedItem.ToString
        Dim strColorName As String = cboColor.SelectedItem.ToString

        If strActualBranch = "[All Branch]" Then strActualBranch = "NULLSTRING"
        If strMainCategoryName = "[All Categories]" Then strMainCategoryName = "NULLSTRING"
        If strSubCategoryName = "[All Sub Categories]" Then strSubCategoryName = "NULLSTRING"
        If strCollectionName = "[All Collections]" Then strCollectionName = "NULLSTRING"
        If strProductName = "[All Units]" Then strProductName = "NULLSTRING"
        If strSizeName = "[All Sizes]" Then strSizeName = "NULLSTRING"
        If strColorName = "[All Colors]" Then strColorName = "NULLSTRING"

        ShowSearchResult(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, strSizeName, strColorName)
    End Sub

    Private Sub frmOpenFile_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dtpFromDate.Value = Str(Month(Now())) + "/01/" + Str(Year(Now()))
        dtpToDate.Value = Now()

    End Sub

    Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
        Dim selectedNode As TreeNode = treeMain.SelectedNode

        Dim searchBranch As String = "NULLSTRING", searchMainCategory As String = "NULLSTRING", searchSubCategory As String = "NULLSTRING", searchCollection As String = "NULLSTRING"
        If IsNothing(selectedNode) Then
            MsgBox("Select the proper item for the report you want to generate.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If


        If selectedNode.Name = "collection_node" Then
            searchCollection = selectedNode.Text
            If searchCollection = "(Empty Collection)" Then searchCollection = ""
            selectedNode = selectedNode.Parent

        End If

        If selectedNode.Name = "subCategory_node" Then
            searchSubCategory = selectedNode.Text
            If searchSubCategory = "(Empty Sub Category)" Then searchSubCategory = ""
            selectedNode = selectedNode.Parent
        End If


        If selectedNode.Name = "mainCategory_node" Then
            searchMainCategory = selectedNode.Text
            If searchMainCategory = "(Empty Main Category)" Then searchMainCategory = ""
            selectedNode = selectedNode.Parent
        End If

        If selectedNode.Name = "treeRetail" Then searchBranch = "retail"
        If selectedNode.Name = "treeStore" Then searchBranch = "store"
        If selectedNode.Name = "treeOnline" Then searchBranch = "online"
        If selectedNode.Name = "treeWholesale" Then searchBranch = "wholesale"

        Dim i As Integer
        Dim prevActualBranch As String = "EmptyString"
        Dim prevMainCategoryName As String = "EmptyString"
        Dim prevSubCategoryName As String = "EmptyString"
        Dim prevCollectName As String = "EmptyString"
        Dim prevProductName As String = "EmptyString"
        Dim objProduct As ProductSaleInfo = Nothing
        Dim aryProductList As New List(Of ProductSaleInfo)

        Dim strActualBranch As String = ""
        Dim strMainCategoryName As String = ""
        Dim strSubCategoryName As String = ""
        Dim strCollectionName As String = ""
        Dim strProductName As String = ""
        Dim strOrderRef As String = ""
        Dim nSohQty As Double = 0                        ' SOH : stock availiable
        Dim nCost As Double = 0                          ' Cost
        Dim nSalePrice As Double = 0                     ' RRPexGST
        Dim nSaleQty As Double = 0                       ' Sales Qty Last 30 Days

        ReportGrid.Rows.Clear()

        For i = 2 To mainGridView.RowCount - 1
            strActualBranch = mainGridView.Rows(i).Cells(0).Value.ToString
            strMainCategoryName = mainGridView.Rows(i).Cells(1).Value.ToString
            strSubCategoryName = mainGridView.Rows(i).Cells(2).Value.ToString
            strCollectionName = mainGridView.Rows(i).Cells(3).Value.ToString
            strProductName = mainGridView.Rows(i).Cells(5).Value.ToString
            strOrderRef = mainGridView.Rows(i).Cells(12).Value.ToString                         ' OrderRef
            nSohQty = Val(mainGridView.Rows(i).Cells(16).Value.ToString)                        ' stock availiable: SOH
            nCost = Val(mainGridView.Rows(i).Cells(11).Value.ToString)                          ' Cost
            nSalePrice = Val(mainGridView.Rows(i).Cells(22).Value.ToString)                     ' RRPexGST
            nSaleQty = Val(mainGridView.Rows(i).Cells(15).Value.ToString)                       ' Sales Qty Last 30 Days

            If searchBranch = "retail" And (strActualBranch = "USA Wholesale, QLD" Or strActualBranch = "Wholesale, QLD") Then Continue For
            If searchBranch = "online" And strActualBranch <> "Adrift HQ, QLD" Then Continue For
            If searchBranch = "wholesale" And strActualBranch <> "USA Wholesale, QLD" And strActualBranch <> "Wholesale, QLD" Then Continue For
            If searchBranch = "store" And (strActualBranch = "Adrift HQ, QLD" Or strActualBranch = "USA Wholesale, QLD" Or strActualBranch = "Wholesale, QLD") Then Continue For

            If searchMainCategory <> "NULLSTRING" And strMainCategoryName <> searchMainCategory Then Continue For
            If searchSubCategory <> "NULLSTRING" And strSubCategoryName <> searchSubCategory Then Continue For
            If searchCollection <> "NULLSTRING" And strCollectionName <> searchCollection Then Continue For
            If chkRefund.Checked = True And Mid(strOrderRef, 1, 3) <> "CRN" Then Continue For
            'If nSaleQty = 0 Then Continue For

            If (prevActualBranch <> "EmptyString" And prevActualBranch <> strActualBranch) Or
                (prevMainCategoryName <> "EmptyString" And prevMainCategoryName <> strMainCategoryName) Or
                (prevSubCategoryName <> "EmptyString" And prevSubCategoryName <> strSubCategoryName) Or
                (prevCollectName <> "EmptyString" And prevCollectName <> strCollectionName) Then

                ShowReportResult(ReportGrid, aryProductList, prevActualBranch, prevMainCategoryName, prevSubCategoryName, prevCollectName)
                aryProductList = New List(Of ProductSaleInfo)
            End If

            Try
                If prevActualBranch <> strActualBranch Or prevMainCategoryName <> strMainCategoryName Or
                prevSubCategoryName <> strSubCategoryName Or prevCollectName <> strCollectionName Or prevProductName <> strProductName Then
                    objProduct = New ProductSaleInfo(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, nCost, nSohQty, nSaleQty, nSalePrice)

                    aryProductList.Add(objProduct)
                Else
                    objProduct.AddCost(nCost)
                    objProduct.AddPrice(nSalePrice)
                    objProduct.AddSaleQty(nSaleQty)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            prevActualBranch = strActualBranch
            prevMainCategoryName = strMainCategoryName
            prevSubCategoryName = strSubCategoryName
            prevCollectName = strCollectionName
            prevProductName = strProductName
        Next

        ShowReportResult(ReportGrid, aryProductList, prevActualBranch, prevMainCategoryName, prevSubCategoryName, prevCollectName)

        ' calculate total
        Dim nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0, nTotalSohQty As Double = 0
        Dim nTotalSaleAmountByFullPrice As Double = 0, nTotalSaleAmouontByDiscountPrice As Double = 0, nTotalProfit As Double = 0
        Dim nDiscountPercent As Double = 0, nProfitPercent As Double = 0

        For i = 0 To ReportGrid.Rows.Count - 1
            ReportGrid.Rows(i).Cells("colNo").Value = i + 1
            nTotalSohQty += ReportGrid.Rows(i).Cells("colSOH").Value
            nTotalSaleAmount += ReportGrid.Rows(i).Cells("colAmount").Value
            nTotalSaleQty += ReportGrid.Rows(i).Cells("colUnit").Value
            nTotalCost += ReportGrid.Rows(i).Cells("colCost").Value
            nTotalSaleAmountByFullPrice += ReportGrid.Rows(i).Cells("colRRP").Value
            nTotalProfit += ReportGrid.Rows(i).Cells("colGp").Value
        Next

        nTotalSaleAmouontByDiscountPrice = nTotalSaleAmount - nTotalSaleAmountByFullPrice

        nDiscountPercent = IIf(nTotalSaleAmount = 0, 0, nTotalSaleAmouontByDiscountPrice / nTotalSaleAmount * 100)
        nProfitPercent = IIf(nTotalSaleAmount = 0, 0, nTotalProfit / nTotalSaleAmount * 100)

        ReportGrid.Rows.Add("", "Total", "", "", "", FormatNumber(nTotalSohQty, 2), FormatNumber(nTotalSaleAmount, 2), FormatNumber(nTotalSaleQty, 2), FormatNumber(nTotalCost, 2),
            FormatNumber(nTotalSaleAmountByFullPrice, 2), FormatNumber(nTotalProfit, 2), FormatNumber(nDiscountPercent, 2) & "%", FormatNumber(nProfitPercent, 2) & "%")
    End Sub

    Private Sub ShowReportResult(ByVal gridObj As DataGridView, ByVal ProductList As List(Of ProductSaleInfo), ByVal strActualBranch As String, ByVal strMainCategoryName As String,
                           ByVal strSubCategoryName As String, ByVal strCollectionName As String)

        Dim nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0, nTotalSohQty As Double = 0
        Dim nTotalSaleAmountByFullPrice As Double = 0, nTotalSaleAmouontByDiscountPrice As Double = 0, nTotalProfit As Double = 0
        Dim nDiscountPercent As Double = 0, nProfitPercent As Double = 0
        Dim nCost As Double = 0, nSaleAmountByFullPrice As Double = 0, nSaleAmountByDiscountPrice As Double = 0
        Dim i = 0
        Try
            For Each product As ProductSaleInfo In ProductList
                nTotalSohQty += product.GeTotalSohQty()
                nTotalSaleQty += product.GeTotalSaleQty()
                nCost = product.GetTotalCost()
                nSaleAmountByFullPrice = product.GetSaleAmountByFullPrice()
                nSaleAmountByDiscountPrice = product.GetSaleAmountByDiscountPrice()
                nTotalCost += nCost
                nTotalSaleAmountByFullPrice += nSaleAmountByFullPrice
                nTotalSaleAmouontByDiscountPrice += nSaleAmountByDiscountPrice
                nTotalSaleAmount += (nSaleAmountByFullPrice + nSaleAmountByDiscountPrice)
                nTotalProfit += (nSaleAmountByFullPrice + nSaleAmountByDiscountPrice - nCost)
                i = i + 1
            Next

            nDiscountPercent = FormatNumber(IIf(nTotalSaleAmount = 0, 0, nTotalSaleAmouontByDiscountPrice / nTotalSaleAmount * 100), 2)
            nProfitPercent = FormatNumber(IIf(nTotalSaleAmount = 0, 0, nTotalProfit / nTotalSaleAmount * 100), 2)

            gridObj.Rows.Add("", strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, FormatNumber(nTotalSohQty, 2),
                FormatNumber(nTotalSaleAmount, 2), FormatNumber(nTotalSaleQty, 2), FormatNumber(nTotalCost, 2),
                FormatNumber(nTotalSaleAmountByFullPrice, 2), FormatNumber(nTotalProfit, 2),
                FormatNumber(nDiscountPercent, 2) & "%", FormatNumber(nProfitPercent, 2) & "%")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub treeMain_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles treeMain.NodeMouseDoubleClick
        btnGenerateReport_Click(sender, e)
    End Sub


    Private Sub ReportGrid_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles ReportGrid.CellMouseDoubleClick
        Dim no As Integer = ReportGrid.Rows(0).Cells(0).Value
        If no = 0 Then
            MsgBox("Select the data you want to see in detail.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim strActualBranch As String = ReportGrid.SelectedRows(0).Cells(1).Value.ToString
        Dim strMainCategoryName As String = ReportGrid.SelectedRows(0).Cells(2).Value.ToString
        Dim strSubCategoryName As String = ReportGrid.SelectedRows(0).Cells(3).Value.ToString
        Dim strCollectionName As String = ReportGrid.SelectedRows(0).Cells(4).Value.ToString

        GetProductReport(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName)
    End Sub

    Private Sub ProductReportGrid_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles ProductReportGrid.CellMouseDoubleClick
        Dim no As Integer = ProductReportGrid.Rows(0).Cells(0).Value
        If no = 0 Then
            MsgBox("Select the data you want to see in detail.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim strActualBranch As String = ProductReportGrid.SelectedRows(0).Cells(1).Value.ToString
        Dim strMainCategoryName As String = ProductReportGrid.SelectedRows(0).Cells(2).Value.ToString
        Dim strSubCategoryName As String = ProductReportGrid.SelectedRows(0).Cells(3).Value.ToString
        Dim strCollectionName As String = ProductReportGrid.SelectedRows(0).Cells(4).Value.ToString
        Dim strProductName As String = ProductReportGrid.SelectedRows(0).Cells(5).Value.ToString

        Dim viewDetailForm As New frmViewDetail
        viewDetailForm.ViewProductDetail(Me, strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName)
        viewDetailForm.ShowDialog()


    End Sub

    Private Sub GetProductReport(ByVal searchActualBranch As String, ByVal searchMainCategory As String,
                               ByVal searchSubCategory As String, ByVal searchCollection As String)
        Dim i As Integer
        Dim prevActualBranch As String = "EmptyString"
        Dim prevMainCategoryName As String = "EmptyString"
        Dim prevSubCategoryName As String = "EmptyString"
        Dim prevCollectName As String = "EmptyString"
        Dim prevProductName As String = "EmptyString"
        Dim objProduct As ProductSaleInfo = Nothing
        Dim aryProductList As New List(Of ProductSaleInfo)

        Dim strActualBranch As String = ""
        Dim strMainCategoryName As String = ""
        Dim strSubCategoryName As String = ""
        Dim strCollectionName As String = ""
        Dim strProductName As String = ""
        Dim strOrderRef As String = ""
        Dim nSohQty As Double = 0                        ' SOH : stock availiable
        Dim nCost As Double = 0                          ' Cost
        Dim nSalePrice As Double = 0                     ' RRPexGST
        Dim nSaleQty As Double = 0                       ' Sales Qty Last 30 Days

        ProductReportGrid.Rows.Clear()

        For i = 2 To mainGridView.RowCount - 1
            strActualBranch = mainGridView.Rows(i).Cells(0).Value.ToString
            strMainCategoryName = mainGridView.Rows(i).Cells(1).Value.ToString
            strSubCategoryName = mainGridView.Rows(i).Cells(2).Value.ToString
            strCollectionName = mainGridView.Rows(i).Cells(3).Value.ToString
            strProductName = mainGridView.Rows(i).Cells(5).Value.ToString
            strOrderRef = mainGridView.Rows(i).Cells(12).Value.ToString                         ' order ref
            nSohQty = Val(mainGridView.Rows(i).Cells(16).Value.ToString)                        ' stock availiable: SOH
            nCost = Val(mainGridView.Rows(i).Cells(11).Value.ToString)                          ' Cost
            nSalePrice = Val(mainGridView.Rows(i).Cells(22).Value.ToString)                     ' RRPexGST
            nSaleQty = Val(mainGridView.Rows(i).Cells(15).Value.ToString)                       ' Sales Qty Last 30 Days


            If searchActualBranch <> "NULLSTRING" And strActualBranch <> searchActualBranch Then Continue For
            If searchMainCategory <> "NULLSTRING" And strMainCategoryName <> searchMainCategory Then Continue For
            If searchSubCategory <> "NULLSTRING" And strSubCategoryName <> searchSubCategory Then Continue For
            If searchCollection <> "NULLSTRING" And strCollectionName <> searchCollection Then Continue For
            If chkRefund.Checked = True And Mid(strOrderRef, 1, 3) <> "CRN" Then Continue For
            'If nSaleQty = 0 Then Continue For

            If (prevActualBranch <> "EmptyString" And prevActualBranch <> strActualBranch) Or
                (prevMainCategoryName <> "EmptyString" And prevMainCategoryName <> strMainCategoryName) Or
                (prevSubCategoryName <> "EmptyString" And prevSubCategoryName <> strSubCategoryName) Or
                (prevCollectName <> "EmptyString" And prevCollectName <> strCollectionName) Or
                (prevProductName <> "EmptyString" And prevProductName <> strProductName) Then

                ShowProductResult(ProductReportGrid, aryProductList, prevActualBranch, prevMainCategoryName, prevSubCategoryName, prevCollectName, prevProductName)
                aryProductList = New List(Of ProductSaleInfo)
            End If

            Try
                If prevActualBranch <> strActualBranch Or prevMainCategoryName <> strMainCategoryName Or prevSubCategoryName <> strSubCategoryName Or
                    prevCollectName <> strCollectionName Or prevProductName <> strProductName Then
                    objProduct = New ProductSaleInfo(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, nCost, nSohQty, nSaleQty, nSalePrice)

                    aryProductList.Add(objProduct)
                Else
                    objProduct.AddCost(nCost)
                    objProduct.AddPrice(nSalePrice)
                    objProduct.AddSaleQty(nSaleQty)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            prevActualBranch = strActualBranch
            prevMainCategoryName = strMainCategoryName
            prevSubCategoryName = strSubCategoryName
            prevCollectName = strCollectionName
            prevProductName = strProductName
        Next

        ShowProductResult(ProductReportGrid, aryProductList, prevActualBranch, prevMainCategoryName, prevSubCategoryName, prevCollectName, prevProductName)

        ' calculate total
        Dim nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0, nTotalSohQty As Double = 0
        Dim nTotalSaleAmountByFullPrice As Double = 0, nTotalSaleAmouontByDiscountPrice As Double = 0, nTotalProfit As Double = 0
        Dim nDiscountPercent As Double = 0, nProfitPercent As Double = 0

        For i = 0 To ProductReportGrid.Rows.Count - 1
            ProductReportGrid.Rows(i).Cells("colPNo").Value = i + 1
            nTotalSohQty += ProductReportGrid.Rows(i).Cells("colPSOH").Value
            nTotalSaleAmount += ProductReportGrid.Rows(i).Cells("colPAmount").Value
            nTotalSaleQty += ProductReportGrid.Rows(i).Cells("colPUnit").Value
            nTotalCost += ProductReportGrid.Rows(i).Cells("colPCost").Value
            nTotalSaleAmountByFullPrice += ProductReportGrid.Rows(i).Cells("colPRRP").Value
            nTotalProfit += ProductReportGrid.Rows(i).Cells("colPGp").Value
        Next

        nTotalSaleAmouontByDiscountPrice = nTotalSaleAmount - nTotalSaleAmountByFullPrice

        nDiscountPercent = IIf(nTotalSaleAmount = 0, 0, nTotalSaleAmouontByDiscountPrice / nTotalSaleAmount * 100)
        nProfitPercent = IIf(nTotalSaleAmount = 0, 0, nTotalProfit / nTotalSaleAmount * 100)

        ProductReportGrid.Rows.Add("", "Total", "", "", "", "", FormatNumber(nTotalSohQty, 2), FormatNumber(nTotalSaleAmount, 2), FormatNumber(nTotalSaleQty, 2), FormatNumber(nTotalCost, 2),
            FormatNumber(nTotalSaleAmountByFullPrice, 2), FormatNumber(nTotalProfit, 2), FormatNumber(nDiscountPercent, 2) & "%", FormatNumber(nProfitPercent, 2) & "%")

        tabBody.SelectedTab = tabProductReport
    End Sub

    Private Sub ShowProductResult(ByVal gridObj As DataGridView, ByVal ProductList As List(Of ProductSaleInfo), ByVal strActualBranch As String, ByVal strMainCategoryName As String,
                           ByVal strSubCategoryName As String, ByVal strCollectionName As String, ByVal strProductName As String)

        Dim nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0, nTotalSohQty As Double = 0
        Dim nTotalSaleAmountByFullPrice As Double = 0, nTotalSaleAmouontByDiscountPrice As Double = 0, nTotalProfit As Double = 0
        Dim nDiscountPercent As Double = 0, nProfitPercent As Double = 0
        Dim nCost As Double = 0, nSaleAmountByFullPrice As Double = 0, nSaleAmountByDiscountPrice As Double = 0
        Dim i = 0
        Try
            For Each product As ProductSaleInfo In ProductList
                nTotalSohQty += product.GeTotalSohQty()
                nTotalSaleQty += product.GeTotalSaleQty()
                nCost = product.GetTotalCost()
                nSaleAmountByFullPrice = product.GetSaleAmountByFullPrice()
                nSaleAmountByDiscountPrice = product.GetSaleAmountByDiscountPrice()
                nTotalCost += nCost
                nTotalSaleAmountByFullPrice += nSaleAmountByFullPrice
                nTotalSaleAmouontByDiscountPrice += nSaleAmountByDiscountPrice
                nTotalSaleAmount += (nSaleAmountByFullPrice + nSaleAmountByDiscountPrice)
                nTotalProfit += (nSaleAmountByFullPrice + nSaleAmountByDiscountPrice - nCost)
                i = i + 1
            Next

            nDiscountPercent = FormatNumber(IIf(nTotalSaleAmount = 0, 0, nTotalSaleAmouontByDiscountPrice / nTotalSaleAmount * 100), 2)
            nProfitPercent = FormatNumber(IIf(nTotalSaleAmount = 0, 0, nTotalProfit / nTotalSaleAmount * 100), 2)

            gridObj.Rows.Add("", strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, FormatNumber(nTotalSohQty, 2),
                FormatNumber(nTotalSaleAmount, 2), FormatNumber(nTotalSaleQty, 2), FormatNumber(nTotalCost, 2),
                FormatNumber(nTotalSaleAmountByFullPrice, 2), FormatNumber(nTotalProfit, 2),
                FormatNumber(nDiscountPercent, 2) & "%", FormatNumber(nProfitPercent, 2) & "%")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ShowSearchResult(Optional ByVal searchActualBranch As String = "NULLSTRING", Optional ByVal searchMainCategory As String = "NULLSTRING",
                               Optional ByVal searchSubCategory As String = "NULLSTRING", Optional ByVal searchCollection As String = "NULLSTRING",
                               Optional ByVal searchProduct As String = "NULLSTRING", Optional ByVal searchSize As String = "NULLSTRING",
                               Optional ByVal searchColor As String = "NULLSTRING")
        Dim i As Integer
        Dim prevActualBranch As String = "EmptyString"
        Dim prevMainCategoryName As String = "EmptyString"
        Dim prevSubCategoryName As String = "EmptyString"
        Dim prevCollectName As String = "EmptyString"
        Dim prevProductName As String = "EmptyString"
        Dim objProduct As ProductSaleInfo = Nothing
        Dim aryProductList As New List(Of ProductSaleInfo)

        Dim strActualBranch As String = ""
        Dim strMainCategoryName As String = ""
        Dim strSubCategoryName As String = ""
        Dim strCollectionName As String = ""
        Dim strProductName As String = ""
        Dim strSizeName As String = ""
        Dim strColorName As String = ""
        Dim strOrderRef As String = ""
        Dim nSohQty As Double = 0                        ' SOH : stock availiable
        Dim nCost As Double = 0                          ' Cost
        Dim nSalePrice As Double = 0                     ' RRPexGST
        Dim nSaleQty As Double = 0                       ' Sales Qty Last 30 Days

        SearchResultGrid.Rows.Clear()

        For i = 2 To mainGridView.RowCount - 1
            strActualBranch = mainGridView.Rows(i).Cells(0).Value.ToString
            strMainCategoryName = mainGridView.Rows(i).Cells(1).Value.ToString
            strSubCategoryName = mainGridView.Rows(i).Cells(2).Value.ToString
            strCollectionName = mainGridView.Rows(i).Cells(3).Value.ToString
            strProductName = mainGridView.Rows(i).Cells(5).Value.ToString
            strOrderRef = mainGridView.Rows(i).Cells(12).Value.ToString                         ' order ref
            nSohQty = Val(mainGridView.Rows(i).Cells(16).Value.ToString)                        ' stock availiable: SOH
            nCost = Val(mainGridView.Rows(i).Cells(11).Value.ToString)                          ' Cost
            nSalePrice = Val(mainGridView.Rows(i).Cells(22).Value.ToString)                     ' RRPexGST
            nSaleQty = Val(mainGridView.Rows(i).Cells(15).Value.ToString)                       ' Sales Qty Last 30 Days


            If searchActualBranch <> "NULLSTRING" And strActualBranch <> searchActualBranch Then Continue For
            If searchMainCategory <> "NULLSTRING" And strMainCategoryName <> searchMainCategory Then Continue For
            If searchSubCategory <> "NULLSTRING" And strSubCategoryName <> searchSubCategory Then Continue For
            If searchCollection <> "NULLSTRING" And strCollectionName <> searchCollection Then Continue For
            If searchProduct <> "NULLSTRING" And strProductName <> searchProduct Then Continue For
            If searchSize <> "NULLSTRING" And strSizeName <> searchSize Then Continue For
            If searchColor <> "NULLSTRING" And strColorName <> searchColor Then Continue For
            If chkRefund.Checked = True And Mid(strOrderRef, 1, 3) <> "CRN" Then Continue For
            'If nSaleQty = 0 Then Continue For

            If (prevActualBranch <> "EmptyString" And prevActualBranch <> strActualBranch) Or
                (searchMainCategory <> "NULLSTRING" And prevMainCategoryName <> "EmptyString" And prevMainCategoryName <> strMainCategoryName) Or
                (searchSubCategory <> "NULLSTRING" And prevSubCategoryName <> "EmptyString" And prevSubCategoryName <> strSubCategoryName) Or
                (searchCollection <> "NULLSTRING" And prevCollectName <> "EmptyString" And prevCollectName <> strCollectionName) Or
                (searchProduct <> "NULLSTRING" And prevProductName <> "EmptyString" And prevProductName <> strProductName) Then

                ShowProductResult(SearchResultGrid, aryProductList, prevActualBranch,
                                  IIf(searchMainCategory = "NULLSTRING", "", prevMainCategoryName),
                                  IIf(searchSubCategory = "NULLSTRING", "", prevSubCategoryName),
                                  IIf(searchCollection = "NULLSTRING", "", prevCollectName),
                                  IIf(searchProduct = "NULLSTRING", "", prevProductName))
                aryProductList = New List(Of ProductSaleInfo)
            End If

            Try
                If prevActualBranch <> strActualBranch Or prevMainCategoryName <> strMainCategoryName Or prevSubCategoryName <> strSubCategoryName Or
                    prevCollectName <> strCollectionName Or prevProductName <> strProductName Then
                    objProduct = New ProductSaleInfo(strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, nCost, nSohQty, nSaleQty, nSalePrice)

                    aryProductList.Add(objProduct)
                Else
                    objProduct.AddCost(nCost)
                    objProduct.AddPrice(nSalePrice)
                    objProduct.AddSaleQty(nSaleQty)
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            prevActualBranch = strActualBranch
            prevMainCategoryName = strMainCategoryName
            prevSubCategoryName = strSubCategoryName
            prevCollectName = strCollectionName
            prevProductName = strProductName
        Next

        ShowProductResult(SearchResultGrid, aryProductList, prevActualBranch,
                                  IIf(searchMainCategory = "NULLSTRING", "", prevMainCategoryName),
                                  IIf(searchSubCategory = "NULLSTRING", "", prevSubCategoryName),
                                  IIf(searchCollection = "NULLSTRING", "", prevCollectName),
                                  IIf(searchProduct = "NULLSTRING", "", prevProductName))

        ' calculate total
        Dim nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0, nTotalSohQty As Double = 0
        Dim nTotalSaleAmountByFullPrice As Double = 0, nTotalSaleAmouontByDiscountPrice As Double = 0, nTotalProfit As Double = 0
        Dim nDiscountPercent As Double = 0, nProfitPercent As Double = 0

        For i = 0 To SearchResultGrid.Rows.Count - 1
            SearchResultGrid.Rows(i).Cells("colSNo").Value = i + 1
            nTotalSohQty += SearchResultGrid.Rows(i).Cells("colSSOH").Value
            nTotalSaleAmount += SearchResultGrid.Rows(i).Cells("colSAmount").Value
            nTotalSaleQty += SearchResultGrid.Rows(i).Cells("colSSaleQty").Value
            nTotalCost += SearchResultGrid.Rows(i).Cells("colSCost").Value
            nTotalSaleAmountByFullPrice += SearchResultGrid.Rows(i).Cells("colSRRP").Value
            nTotalProfit += SearchResultGrid.Rows(i).Cells("colSGp").Value
        Next

        nTotalSaleAmouontByDiscountPrice = nTotalSaleAmount - nTotalSaleAmountByFullPrice

        nDiscountPercent = IIf(nTotalSaleAmount = 0, 0, nTotalSaleAmouontByDiscountPrice / nTotalSaleAmount * 100)
        nProfitPercent = IIf(nTotalSaleAmount = 0, 0, nTotalProfit / nTotalSaleAmount * 100)

        SearchResultGrid.Rows.Add("", "Total", "", "", "", "", FormatNumber(nTotalSohQty, 2), FormatNumber(nTotalSaleAmount, 2), FormatNumber(nTotalSaleQty, 2), FormatNumber(nTotalCost, 2),
            FormatNumber(nTotalSaleAmountByFullPrice, 2), FormatNumber(nTotalProfit, 2), FormatNumber(nDiscountPercent, 2) & "%", FormatNumber(nProfitPercent, 2) & "%")

        tabBody.SelectedTab = tabSearchResult

    End Sub

    Private Sub SearchResultGrid_CellMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles SearchResultGrid.CellMouseDoubleClick
        Dim no As Integer = SearchResultGrid.Rows(0).Cells(0).Value
        If no = 0 Then
            MsgBox("Select the data you want to see in detail.", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim strActualBranch As String = SearchResultGrid.SelectedRows(0).Cells(1).Value.ToString
        Dim strMainCategoryName As String = SearchResultGrid.SelectedRows(0).Cells(2).Value.ToString
        Dim strSubCategoryName As String = SearchResultGrid.SelectedRows(0).Cells(3).Value.ToString
        Dim strCollectionName As String = SearchResultGrid.SelectedRows(0).Cells(4).Value.ToString
        Dim strProductName As String = SearchResultGrid.SelectedRows(0).Cells(5).Value.ToString

        If strMainCategoryName = "" Then strMainCategoryName = "NULLSTRING"
        If strSubCategoryName = "" Then strSubCategoryName = "NULLSTRING"
        If strCollectionName = "" Then strCollectionName = "NULLSTRING"
        If strProductName = "" Then strProductName = "NULLSTRING"

        Dim viewDetailForm As New frmViewDetail
        viewDetailForm.ViewProductDetail(Me, strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName)
        viewDetailForm.ShowDialog()


    End Sub

    Private Sub butExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        Dim saveFilePath As String = ""
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim SelectedGrid As DataGridView
        Dim nSelectedTab As Integer = tabBody.SelectedIndex

        If nSelectedTab = 1 Then
            SelectedGrid = ReportGrid
        ElseIf nSelectedTab = 2 Then
            SelectedGrid = ProductReportGrid
        ElseIf nSelectedTab = 3 Then
            SelectedGrid = SearchResultGrid
        Else
            MsgBox("Can not download the data.", vbCritical)
            Exit Sub
        End If

        If SelectedGrid.Rows.Count = 0 Then
            MsgBox("There is no download data.", vbCritical)
            Exit Sub
        End If

        SaveFileDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003|*.xls|Excel Macro-Enabled Workbook|*.xlsm"
        SaveFileDialog1.Title = "Save Excel File"
        SaveFileDialog1.FileName = "DataReport.xlsx"
        SaveFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory

        Dim result As DialogResult = SaveFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            saveFilePath = SaveFileDialog1.FileName
        Else
            Exit Sub
        End If

        xlApp = DirectCast(CreateObject("Excel.Application"), Excel.Application)
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = CType(xlWorkBook.Sheets("sheet1"), Excel.Worksheet)

        Dim widths(0 To SelectedGrid.Columns.Count - 1) As Single
        Dim aligns(0 To SelectedGrid.Columns.Count - 1) As Integer
        If nSelectedTab = 1 Then
            widths(0) = 5 : widths(1) = 16 : widths(2) = 16 : widths(3) = 16 : widths(4) = 16 : widths(5) = 8 : widths(6) = 12
            widths(7) = 12 : widths(8) = 12 : widths(9) = 12 : widths(10) = 12 : widths(11) = 12 : widths(12) = 12

            aligns(0) = Excel.Constants.xlCenter : aligns(1) = Excel.Constants.xlLeft : aligns(2) = Excel.Constants.xlLeft
            aligns(3) = Excel.Constants.xlLeft : aligns(4) = Excel.Constants.xlLeft : aligns(5) = Excel.Constants.xlRight
            aligns(6) = Excel.Constants.xlRight : aligns(7) = Excel.Constants.xlRight : aligns(8) = Excel.Constants.xlRight
            aligns(9) = Excel.Constants.xlRight : aligns(10) = Excel.Constants.xlRight : aligns(11) = Excel.Constants.xlRight
            aligns(12) = Excel.Constants.xlRight
        End If
        If nSelectedTab = 2 Then
            widths(0) = 4 : widths(1) = 15 : widths(2) = 15 : widths(3) = 15 : widths(4) = 15 : widths(5) = 25 : widths(6) = 7 : widths(7) = 11
            widths(8) = 11 : widths(9) = 11 : widths(10) = 11 : widths(11) = 11 : widths(12) = 11 : widths(13) = 11

            aligns(0) = Excel.Constants.xlCenter : aligns(1) = Excel.Constants.xlLeft : aligns(2) = Excel.Constants.xlLeft
            aligns(3) = Excel.Constants.xlLeft : aligns(4) = Excel.Constants.xlLeft : aligns(5) = Excel.Constants.xlLeft : aligns(6) = Excel.Constants.xlRight
            aligns(7) = Excel.Constants.xlRight : aligns(8) = Excel.Constants.xlRight : aligns(9) = Excel.Constants.xlRight
            aligns(10) = Excel.Constants.xlRight : aligns(11) = Excel.Constants.xlRight : aligns(12) = Excel.Constants.xlRight
            aligns(13) = Excel.Constants.xlRight
        End If
        If nSelectedTab = 3 Then
            widths(0) = 4 : widths(1) = 15 : widths(2) = 15 : widths(3) = 15 : widths(4) = 15 : widths(5) = 25 : widths(6) = 7 : widths(7) = 11
            widths(8) = 11 : widths(9) = 11 : widths(10) = 11 : widths(11) = 11 : widths(12) = 11 : widths(13) = 11

            aligns(0) = Excel.Constants.xlCenter : aligns(1) = Excel.Constants.xlLeft : aligns(2) = Excel.Constants.xlLeft
            aligns(3) = Excel.Constants.xlLeft : aligns(4) = Excel.Constants.xlLeft : aligns(5) = Excel.Constants.xlLeft : aligns(6) = Excel.Constants.xlRight
            aligns(7) = Excel.Constants.xlRight : aligns(8) = Excel.Constants.xlRight : aligns(9) = Excel.Constants.xlRight
            aligns(10) = Excel.Constants.xlRight : aligns(11) = Excel.Constants.xlRight : aligns(12) = Excel.Constants.xlRight
            aligns(13) = Excel.Constants.xlRight
        End If

        xlWorkSheet.Rows(1).RowHeight = 20
        For col As Integer = 0 To SelectedGrid.Columns.Count - 1
            xlWorkSheet.Cells(1, col + 1) = SelectedGrid.Columns(col).HeaderText()
            xlWorkSheet.Cells(1, col + 1).WrapText = True
            xlWorkSheet.Cells(1, col + 1).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Cells(1, col + 1).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Cells(1, col + 1).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            xlWorkSheet.Cells(1, col + 1).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

            xlWorkSheet.Columns(col + 1).ColumnWidth = widths(col)
            xlWorkSheet.Columns(col + 1).HorizontalAlignment = Excel.Constants.xlCenter
        Next

        pgLoading.Visible = True
        pgLoading.Value = 0
        pgLoading.Maximum = SelectedGrid.Rows.Count - 1

        For row As Integer = 0 To SelectedGrid.Rows.Count - 1
            xlWorkSheet.Rows(row + 2).RowHeight = 20
            For col As Integer = 0 To SelectedGrid.ColumnCount - 1
                xlWorkSheet.Cells(row + 2, col + 1) = SelectedGrid.Rows(row).Cells(SelectedGrid.Columns(col).Name).Value.ToString
                xlWorkSheet.Cells(row + 2, col + 1).HorizontalAlignment = aligns(col)
                xlWorkSheet.Cells(row + 2, col + 1).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                xlWorkSheet.Cells(row + 2, col + 1).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                xlWorkSheet.Cells(row + 2, col + 1).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                xlWorkSheet.Cells(row + 2, col + 1).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            Next
            pgLoading.Value = row
        Next
        pgLoading.Visible = False


        Try
            With xlWorkBook
                .SaveAs(saveFilePath)
                .Saved = True
            End With

            MsgBox("Excel format success exported.", vbInformation)
        Catch ex As Exception
            MsgBox("An error occurred while exporting." + vbCrLf + ex.Message, vbExclamation)
        Finally

            GC.Collect()
        End Try

        'xlApp.Visible = True
        'xlApp.UserControl = True

        xlWorkBook.Close()
        xlApp.Quit()

        xlWorkSheet = Nothing
        xlWorkBook = Nothing
        xlApp = Nothing
    End Sub

    Private Sub butExportPdf_Click(sender As Object, e As EventArgs) Handles btnExportPdf.Click
        Dim saveFilePath As String = ""
        Dim SelectedGrid As DataGridView
        Dim nSelectedTab As Integer = tabBody.SelectedIndex


        If nSelectedTab = 1 Then
            SelectedGrid = ReportGrid
        ElseIf nSelectedTab = 2 Then
            SelectedGrid = ProductReportGrid
        ElseIf nSelectedTab = 3 Then
            SelectedGrid = SearchResultGrid
        Else
            MsgBox("Can not export the data.", vbCritical)
            Exit Sub
        End If

        If SelectedGrid.Rows.Count = 0 Then
            MsgBox("There is no download data.", vbCritical)
            Exit Sub
        End If

        SaveFileDialog1.Filter = "PDF files|*.pdf|All files|*.*"
        SaveFileDialog1.Title = "Save Excel File"
        SaveFileDialog1.FileName = "DataReport.pdf"
        SaveFileDialog1.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory

        Dim result As DialogResult = SaveFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            saveFilePath = SaveFileDialog1.FileName
        Else
            Exit Sub
        End If

        Dim Paragraph As New Paragraph  ' delclaration for new paragraph
        Dim pdfFile As New Document(PageSize.A4.Rotate(), 40, 40, 40, 20) ' set pdf page size
        pdfFile.AddTitle("Export to Data") ' set our pdf title

        Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(pdfFile, New FileStream(saveFilePath, FileMode.Create))
        pdfFile.Open()

        ' declaration font type
        Dim pTitle As New Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
        Dim pTable As New Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)

        ' insert title into pdf file
        Paragraph = New Paragraph(New Chunk("Data Analysize Report", pTitle))
        Paragraph.Alignment = Element.ALIGN_CENTER
        Paragraph.SpacingAfter = 5.0F

        ' set and add page with current settings
        pdfFile.Add(Paragraph)

        ' create data into table
        Dim pdfTable As New PdfPTable(SelectedGrid.Columns.Count)

        ' setting width of table
        pdfTable.TotalWidth = 750.0F ' A4 Landscape size, If A4 Portrait, 500.0F
        pdfTable.LockedWidth = True

        Dim widths(0 To SelectedGrid.Columns.Count - 1) As Single
        Dim aligns(0 To SelectedGrid.Columns.Count - 1) As Integer
        If nSelectedTab = 1 Then
            widths(0) = 0.7F : widths(1) = 1.9F : widths(2) = 1.9F : widths(3) = 1.9F : widths(4) = 2.0F : widths(5) = 1.0F : widths(6) = 1.5F
            widths(7) = 1.4F : widths(8) = 1.4F : widths(9) = 1.4F : widths(10) = 1.4F : widths(11) = 1.4F : widths(12) = 1.4F
            aligns(0) = PdfPCell.ALIGN_CENTER : aligns(1) = PdfPCell.ALIGN_LEFT : aligns(2) = PdfPCell.ALIGN_LEFT : aligns(3) = PdfPCell.ALIGN_LEFT : aligns(4) = PdfPCell.ALIGN_LEFT : aligns(5) = PdfPCell.ALIGN_RIGHT
            aligns(6) = PdfPCell.ALIGN_RIGHT : aligns(7) = PdfPCell.ALIGN_RIGHT : aligns(8) = PdfPCell.ALIGN_RIGHT : aligns(9) = PdfPCell.ALIGN_RIGHT : aligns(10) = PdfPCell.ALIGN_RIGHT : aligns(11) = PdfPCell.ALIGN_RIGHT
            aligns(12) = PdfPCell.ALIGN_RIGHT
        End If
        If nSelectedTab = 2 Then
            widths(0) = 0.6F : widths(1) = 1.8F : widths(2) = 1.8F : widths(3) = 1.8F : widths(4) = 1.9F : widths(5) = 2.0F : widths(6) = 0.9F
            widths(7) = 1.3F : widths(8) = 1.2F : widths(9) = 1.2F : widths(10) = 1.2F : widths(11) = 1.2F : widths(12) = 1.2F : widths(13) = 1.2F
            aligns(0) = PdfPCell.ALIGN_CENTER : aligns(1) = PdfPCell.ALIGN_LEFT : aligns(2) = PdfPCell.ALIGN_LEFT : aligns(3) = PdfPCell.ALIGN_LEFT
            aligns(4) = PdfPCell.ALIGN_LEFT : aligns(5) = PdfPCell.ALIGN_LEFT : aligns(6) = PdfPCell.ALIGN_RIGHT
            aligns(7) = PdfPCell.ALIGN_RIGHT : aligns(8) = PdfPCell.ALIGN_RIGHT : aligns(9) = PdfPCell.ALIGN_RIGHT
            aligns(10) = PdfPCell.ALIGN_RIGHT : aligns(11) = PdfPCell.ALIGN_RIGHT : aligns(12) = PdfPCell.ALIGN_RIGHT
            aligns(13) = PdfPCell.ALIGN_RIGHT
        End If
        If nSelectedTab = 3 Then
            widths(0) = 0.6F : widths(1) = 1.8F : widths(2) = 1.8F : widths(3) = 1.8F : widths(4) = 1.9F : widths(5) = 2.0F : widths(6) = 0.9F
            widths(7) = 1.3F : widths(8) = 1.2F : widths(9) = 1.2F : widths(10) = 1.2F : widths(11) = 1.2F : widths(12) = 1.2F : widths(13) = 1.2F
            aligns(0) = PdfPCell.ALIGN_CENTER : aligns(1) = PdfPCell.ALIGN_LEFT : aligns(2) = PdfPCell.ALIGN_LEFT : aligns(3) = PdfPCell.ALIGN_LEFT
            aligns(4) = PdfPCell.ALIGN_LEFT : aligns(5) = PdfPCell.ALIGN_LEFT : aligns(6) = PdfPCell.ALIGN_RIGHT
            aligns(7) = PdfPCell.ALIGN_RIGHT : aligns(8) = PdfPCell.ALIGN_RIGHT : aligns(9) = PdfPCell.ALIGN_RIGHT
            aligns(10) = PdfPCell.ALIGN_RIGHT : aligns(11) = PdfPCell.ALIGN_RIGHT : aligns(12) = PdfPCell.ALIGN_RIGHT
            aligns(13) = PdfPCell.ALIGN_RIGHT
        End If


        pdfTable.SetWidths(widths)
        pdfTable.HorizontalAlignment = 0
        pdfTable.SpacingBefore = 5.0F

        ' declaration pdf cells
        Dim pdfCell As PdfPCell = New PdfPCell

        ' create pdf header
        For i As Integer = 0 To SelectedGrid.Columns.Count - 1
            pdfCell = New PdfPCell(New Phrase(New Chunk(SelectedGrid.Columns(i).HeaderText, pTable)))
            ' alignment header table
            pdfCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER
            ' add cells into pdf table
            pdfTable.AddCell(pdfCell)
        Next

        ' add data into pdf table
        pgLoading.Visible = True
        pgLoading.Value = 0
        pgLoading.Maximum = SelectedGrid.Rows.Count - 1

        For i As Integer = 0 To SelectedGrid.Rows.Count - 1
            For j As Integer = 0 To SelectedGrid.Columns.Count - 1
                pdfCell = New PdfPCell(New Phrase(SelectedGrid(j, i).Value.ToString(), pTable))
                pdfCell.HorizontalAlignment = aligns(j)
                pdfTable.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                pdfTable.AddCell(pdfCell)
            Next
            pgLoading.Value = i
        Next
        pgLoading.Visible = False

        ' add pdf table into pdf document
        pdfFile.Add(pdfTable)
        pdfFile.Close() ' close all sessions

        ' show message if has been exported
        MessageBox.Show("PDF format success exported !", "Informations", MessageBoxButtons.OK, MessageBoxIcon.Information)




        'Dim pdfTable As New PdfPTable(SelectedGrid.ColumnCount)
        'pdfTable.DefaultCell.Padding = 3
        'pdfTable.WidthPercentage = 30
        'pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
        'pdfTable.DefaultCell.BorderWidth = 1

        ''Adding Header row
        'For Each column As DataGridViewColumn In productReportGrid.Columns
        '    Dim cell As New PdfPCell(New Phrase(column.HeaderText))
        '    cell.BackgroundColor = New iTextSharp.text.BaseColor(240, 240, 240)
        '    pdfTable.AddCell(cell)
        'Next

        ''Adding DataRow
        'Dim cellvalue As String = ""
        'Dim i As Integer = 0
        'For Each row As DataGridViewRow In productReportGrid.Rows
        '    For Each cell As DataGridViewCell In row.Cells
        '        cellvalue = cell.FormattedValue
        '        pdfTable.AddCell(Convert.ToString(cellvalue))
        '    Next
        'Next

        ''Exporting to PDF
        'Dim folderPath As String = "D:\"
        'If Not Directory.Exists(folderPath) Then
        '    Directory.CreateDirectory(folderPath)
        'End If
        'Using stream As New FileStream(folderPath & "DataGridViewExport.pdf", FileMode.Create)
        '    Dim pdfDoc As New Document(PageSize.A2, 10.0F, 10.0F, 10.0F, 0.0F)
        '    PdfWriter.GetInstance(pdfDoc, stream)
        '    pdfDoc.Open()
        '    pdfDoc.Add(pdfTable)
        '    pdfDoc.Close()
        '    stream.Close()
        'End Using
    End Sub

End Class
