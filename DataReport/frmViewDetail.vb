Public Class frmViewDetail
    Public Sub ViewProductDetail(ByVal parentForm As frmOpenFile, Optional ByVal searchActualBranch As String = "NULLSTRING", Optional ByVal searchMainCategory As String = "NULLSTRING",
                               Optional ByVal searchSubCategory As String = "NULLSTRING", Optional ByVal searchCollection As String = "NULLSTRING",
                               Optional ByVal searchProduct As String = "NULLSTRING", Optional ByVal searchSize As String = "NULLSTRING",
                               Optional ByVal searchColor As String = "NULLSTRING")
        Dim no As Integer = 0
        Dim nTotalSohQty As Double = 0, nTotalSaleAmount As Double = 0, nTotalSaleQty As Double = 0, nTotalCost As Double = 0


        ProductDetailGrid.Rows.Clear()
        For i = 2 To parentForm.mainGridView.RowCount - 2
            Dim strActualBranch As String = parentForm.mainGridView.Rows(i).Cells(0).Value.ToString
            Dim strMainCategoryName As String = parentForm.mainGridView.Rows(i).Cells(1).Value.ToString
            Dim strSubCategoryName As String = parentForm.mainGridView.Rows(i).Cells(2).Value.ToString
            Dim strCollectionName As String = parentForm.mainGridView.Rows(i).Cells(3).Value.ToString
            Dim strProductName As String = parentForm.mainGridView.Rows(i).Cells(5).Value.ToString
            Dim strSizeName As String = parentForm.mainGridView.Rows(i).Cells(8).Value.ToString
            Dim strColorName As String = parentForm.mainGridView.Rows(i).Cells(9).Value.ToString
            Dim strSaleDate As String = parentForm.mainGridView.Rows(i).Cells(13).Value.ToString
            Dim strOrderRef As String = parentForm.mainGridView.Rows(i).Cells(12).Value.ToString
            Dim nSaleQty As Integer = Val(parentForm.mainGridView.Rows(i).Cells(15).Value)
            Dim nCost As Double = Val(parentForm.mainGridView.Rows(i).Cells(11).Value)                           ' Cost
            Dim nSaleAmount As Double = Val(parentForm.mainGridView.Rows(i).Cells(22).Value)                     ' RRPexGST

            If searchActualBranch <> "NULLSTRING" And strActualBranch <> searchActualBranch Then Continue For
            If searchMainCategory <> "NULLSTRING" And strMainCategoryName <> searchMainCategory Then Continue For
            If searchSubCategory <> "NULLSTRING" And strSubCategoryName <> searchSubCategory Then Continue For
            If searchCollection <> "NULLSTRING" And strCollectionName <> searchCollection Then Continue For
            If searchProduct <> "NULLSTRING" And strProductName <> searchProduct Then Continue For
            If searchSize <> "NULLSTRING" And strSizeName <> searchSize Then Continue For
            If searchColor <> "NULLSTRING" And strColorName <> searchColor Then Continue For

            If parentForm.chkRefund.Checked = True And Mid(strOrderRef, 1, 3) <> "CRN" Then Continue For
            'If nSaleQty = 0 Then Continue For

            nTotalSaleAmount += nSaleAmount
            nTotalSaleQty += nSaleQty
            nTotalCost += nCost

            no += 1
            ProductDetailGrid.Rows.Add(no, strActualBranch, strMainCategoryName, strSubCategoryName, strCollectionName, strProductName, strSizeName,
                strColorName, FormatNumber(nCost, 2), strSaleDate, FormatNumber(nSaleQty, 2), FormatNumber(nSaleAmount, 2))
        Next

        ProductDetailGrid.Rows.Add("", "Total: ", "", "", "", "", "", "", FormatNumber(nTotalCost, 2), "", FormatNumber(nTotalSaleQty, 2),
                            FormatNumber(nTotalSaleAmount, 2))

    End Sub
End Class