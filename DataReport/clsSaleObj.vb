Public Class ProductSaleInfo
    Dim _strBranchName As String
    Dim _strMainCategoryName As String
    Dim _strSubCategoryNamee As String
    Dim _strCollectionName As String
    Dim _strProductName As String

    Dim arySohQtyList() As Double
    Dim aryCostList() As Double
    Dim aryPriceList() As Double
    Dim arySaleQtyList() As Double

    Public Property strBranchName As Integer
        Get
            Return _strBranchName
        End Get
        Set(ByVal value As Integer)
            _strBranchName = value
        End Set
    End Property

    Public Property strMainCategoryName As Integer
        Get
            Return _strMainCategoryName
        End Get
        Set(ByVal value As Integer)
            _strMainCategoryName = value
        End Set
    End Property

    Public Property strSubCategoryNamee As Integer
        Get
            Return _strSubCategoryNamee
        End Get
        Set(ByVal value As Integer)
            _strSubCategoryNamee = value
        End Set
    End Property

    Public Property strCollectionName As Integer
        Get
            Return _strCollectionName
        End Get
        Set(ByVal value As Integer)
            _strCollectionName = value
        End Set
    End Property

    Public Property strProductName As Integer
        Get
            Return _strProductName
        End Get
        Set(ByVal value As Integer)
            _strProductName = value
        End Set
    End Property



    Public Sub New(branch As String, maincategory As String, subcategory As String, collect As String,
                    product As String, cost As Double, soh As Double, saleqty As Double, price As Double)
        _strBranchName = branch
        _strMainCategoryName = maincategory
        _strSubCategoryNamee = subcategory
        _strCollectionName = collect
        _strProductName = product

        AddSoh(soh)
        AddCost(cost)
        AddSaleQty(saleqty)
        AddPrice(price)
    End Sub

    Public Sub AddSoh(soh As Double)

        Dim n As Integer = -1

        If IsNothing(arySohQtyList) = False Then n = arySohQtyList.GetUpperBound(0)

        ReDim Preserve arySohQtyList(n + 1)
        arySohQtyList(n + 1) = soh
    End Sub

    Public Sub AddCost(cost As Double)

        Dim n As Integer = -1

        If IsNothing(aryCostList) = False Then n = aryCostList.GetUpperBound(0)

        ReDim Preserve aryCostList(n + 1)
        aryCostList(n + 1) = cost
    End Sub


    Public Sub AddSaleQty(saleqty As Double)
        Dim n As Integer = -1

        If IsNothing(arySaleQtyList) = False Then n = arySaleQtyList.GetUpperBound(0)

        ReDim Preserve arySaleQtyList(n + 1)
        arySaleQtyList(n + 1) = saleqty
    End Sub

    Public Sub AddPrice(price As Double)
        Dim n As Integer = -1

        If IsNothing(aryPriceList) = False Then n = aryPriceList.GetUpperBound(0)

        ReDim Preserve aryPriceList(n + 1)
        aryPriceList(n + 1) = price
    End Sub

    Public Function GeTotalSohQty() As Integer
        Dim i As Integer
        Dim nTotalSohQty As Integer = 0

        For i = 0 To arySohQtyList.GetUpperBound(0)
            nTotalSohQty += arySohQtyList(i)
        Next
        Return nTotalSohQty

    End Function

    Public Function GeTotalSaleQty() As Integer
        Dim i As Integer
        Dim nTotalSaleQty As Integer = 0

        For i = 0 To arySaleQtyList.GetUpperBound(0)
            nTotalSaleQty += arySaleQtyList(i)
        Next
        Return nTotalSaleQty

    End Function

    Public Function GetTotalCost() As Double
        Dim i As Integer
        Dim nTotalCost As Double = 0

        For i = 0 To aryCostList.GetUpperBound(0)
            nTotalCost += (aryCostList(i) * arySaleQtyList(i))
        Next
        Return nTotalCost

    End Function

    Public Function GetSaleAmountByFullPrice() As Double
        Dim i As Integer
        Dim nFullPrice As Double = 0, nTotalAmt As Double = 0

        nFullPrice = GetFullPrice()
        For i = 0 To aryPriceList.GetUpperBound(0)
            If aryPriceList(i) = nFullPrice Then
                nTotalAmt += aryPriceList(i)
            End If
        Next
        Return nTotalAmt

    End Function

    Public Function GetSaleAmountByDiscountPrice() As Double
        Dim i As Integer
        Dim nFullPrice As Double = 0, nTotalAmt As Double = 0

        nFullPrice = GetFullPrice()
        For i = 0 To aryPriceList.GetUpperBound(0)
            If aryPriceList(i) <> nFullPrice Then
                nTotalAmt += aryPriceList(i)
            End If
        Next
        Return nTotalAmt

    End Function

    Private Function GetFullPrice() As Double
        Dim nPrice As Double = 0

        If aryPriceList.GetUpperBound(0) = 0 Then
            GetFullPrice = 0
            Exit Function
        End If

        For i = 0 To aryPriceList.GetUpperBound(0)
            If aryPriceList(i) > nPrice Then nPrice = aryPriceList(i)
        Next

        GetFullPrice = nPrice
    End Function
End Class
