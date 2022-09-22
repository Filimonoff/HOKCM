Function ArrTransp(ByVal arr As Variant)
Dim tempArray As Variant, clearArr As Variant
Dim r As Long, rr As Long, X As Long, y As Long

     ReDim tempArray(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1))
     For X = LBound(arr, 2) To UBound(arr, 2)
         For y = LBound(arr, 1) To UBound(arr, 1)
             tempArray(X, y) = arr(y, X)
         Next y
     Next X
     y = 0
     X = 0
     
     

r = UBound(tempArray, 2) + 1
rr = UBound(tempArray, 1) + 1




Cells.Clear

Range(Cells(2, 1), Cells(rr + 1, r)).Value = tempArray
End Function

Sub GetBuyerData()


Dim Buyer As String, skDateIdFrom As String, skDateIdTo As String, selects As String, declales As String, fildsNames As Variant, joins As String
 
fildsNames = Array("IDD", "IDO", "BuyerOrderNumber", "SK_Product_ID", "DocTTNNumber", "ProductName", "DeliveryAddress", "ChainName", "BuyerName", "PlanOrderAmount", "FactOrderAmount", "PlanRealAmount", "FactRealAmount", "DiffPlanAmount", "DiffFactAmount", "OrderDate", "SalesDate", "Reasons", "OrdersInDateAmount")

Dim lastCS As Long
lastCS = UBound(fildsNames) + 1
 Dim RS As New ADODB.Recordset
 Dim Cnn As New ADODB.Connection

Dim cCube As String, cCAKE_WH As String, cCubeTest As String
cCubeTest = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Data Source=olapbkk2;Initial Catalog=OLAP_TEST"
cCube = "Provider=MSOLAP.5;Integrated Security=SSPI;Persist Security Info=True;Data Source=olapbkk2;Initial Catalog=OLAP; Cube=SALES"
cCAKE_WH = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=olapbkk2;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=BKK00VDI33;Use Encryption for Data=False;Tag with column collation when possible=False;Initial Catalog=CAKE_WH"
 Cnn.ConnectionString = cCAKE_WH
 Cnn.Open
 
Set RS = Cnn.Execute("SELECT DISTINCT(ChainName) FROM CAKE_WH.dim.qry_Outlets WHERE SalesChannelMkiPdhName = 'ВИП'")
ch = RS.GetRows
ReDim chains(LBound(ch, 1) To UBound(ch, 2))
For i = LBound(ch, 1) To UBound(ch, 2)
chains(i) = ch(0, i)
Next i

Set RS = Nothing



SetChainNameForm.ComboBox1.List = chains
SetChainNameForm.Show
Buyer = SetChainNameForm.ComboBox1
skDateIdFrom = SetPeriodOfSales.TextBoxFrom
skDateIdTo = SetPeriodOfSales.TextBoxTo


Application.ScreenUpdating = False
Set RS = Cnn.Execute("DECLARE @ChN nvarchar(100) ='" & Buyer & _
"' DECLARE @SK_SalesDateFrom int =" & skDateIdFrom & _
" DECLARE @SK_SalesDateTo int =" & skDateIdTo & _
" DECLARE @ChainID int SELECT @ChainID = Chain_ID FROM CAKE_WH.dim.Chains WHERE ChainName = @ChN " & _
"SELECT CONCAT(LossesMax.SK_SalesDate_ID, '_', LossesMax.SK_Product_ID, LossesMax.SK_Outlet_ID) AS IDD, CONCAT(LossesMax.BuyerOrderNumber, LossesMax.SK_Product_ID, LossesMax.SK_Outlet_ID, '_') AS IDO, LossesMax.BuyerOrderNumber, LossesMax.SK_Product_ID, LossesMax.DocTTNNumber, LossesMax.ProductName, LossesMax.DeliveryAddress, qry_Outlets.ChainName, LossesMax.BuyerName, LossesMax.PlanOrderAmount, LossesMax.FactOrderAmount, LossesMax.PlanRealAmount, LossesMax.FactRealAmount, LossesMax.DiffPlanAmount, LossesMax.DiffFactAmount, LossesMax.OrderDate, LossesMax.SalesDate," & _
" CASE" & _
" WHEN CONCAT(LossesMax.ReasonForLosses, LossesMax.ReasonForReturn) ='' AND LossesMax.DocTTNNumber = '' AND LossesMax.DiffFactAmount > 0 THEN 'Замовлення не відпрацьовано. Коментар відсутній.'" & _
" WHEN CONCAT(LossesMax.ReasonForLosses, LossesMax.ReasonForReturn) ='' AND LossesMax.DocTTNNumber <> '' AND LossesMax.FactRealAmount - LossesMax.PlanRealAmount < 0 THEN 'Повернення якістної продукції. Коментар відсутній.'" & _
" WHEN LossesMax.PlanOrderAmount = LossesMax.FactOrderAmount AND LossesMax.DiffPlanAmount = LossesMax.DiffFactAmount AND LossesMax.DiffFactAmount = 0 AND LossesMax.DocTTNNumber <> '' THEN CONCAT('Доставлене ', FORMAT(LossesMax.SalesDate, 'dd.MM.yyyy'), ' замовл. № ', LossesMax.BuyerOrderNumber, ' в кількості: ', LossesMax.FactRealAmount, 'шт. ТТН №: ', LossesMax.DocTTNNumber, ' ціна ', CASE WHEN pp.SellingPriceWithVATWithDiscount IS NOT NULL THEN FORMAT(pp.SellingPriceWithVATWithDiscount, '0.00') ELSE 'не акційна' END)" & _
" WHEN LossesMax.ReasonForLosses = 'Мала кількість' AND LossesMax.FactOrderWeight - LossesMax.PlanOrderWeight = 0 THEN 'Замовлення меньше 3 кг.'" & _
" WHEN LossesMax.ReasonForLosses = 'Мала кількість' AND LossesMax.FactOrderWeight - LossesMax.PlanOrderWeight <> 0 THEN 'Замовлення після корегування меньше 3 кг.'" & _
" ELSE CONCAT(LossesMax.ReasonForLosses, LossesMax.ReasonForReturn)" & _
" END Reasons,co.CountOrders" & _
" FROM CAKE_WH.fact.LossesMax WITH (NOLOCK)" & _
" LEFT JOIN CAKE_WH.dim.qry_Outlets ON qry_Outlets.SK_Outlet_ID = LossesMax.SK_Outlet_ID" & _
" Left Join(SELECT pp.SK_Date_Purchase_From, pp.SK_Date_Purchase_To, pp.SK_Outlet_ID, pp.SK_Product_ID, pp.SK_Promo_ID, pp.SellingPriceWithVATWithDiscount" & _
" FROM CAKE_WH.dim.qry_PromoPriceByTT AS pp WITH (NOLOCK)" & _
" WHERE pp.Chain_ID = @ChainID AND (pp.SK_Date_Purchase_From BETWEEN @SK_SalesDateFrom AND @SK_SalesDateTo" & _
" OR pp.SK_Date_Purchase_To BETWEEN @SK_SalesDateFrom AND @SK_SalesDateTo" & _
" OR pp.SK_Date_Purchase_From < @SK_SalesDateFrom AND pp.SK_Date_Purchase_To > @SK_SalesDateTo)) AS pp" & _
" ON LossesMax.SK_SalesDate_ID BETWEEN pp.SK_Date_Purchase_From AND pp.SK_Date_Purchase_To AND LossesMax.SK_Outlet_ID = pp.SK_Outlet_ID AND LossesMax.SK_Product_ID = pp.SK_Product_ID" & _
" LEFT JOIN(SELECT LossesMax.SK_SalesDate_ID, LossesMax.SK_Outlet_ID, LossesMax.SK_Product_ID, COUNT(DISTINCT LossesMax.DocOrderNumber) AS CountOrders FROM CAKE_WH.fact.LossesMax WITH (NOLOCK)" & _
" INNER JOIN CAKE_WH.dim.qry_Outlets WITH (NOLOCK) ON LossesMax.SK_Outlet_ID = qry_Outlets.SK_Outlet_ID WHERE LossesMax.SK_SalesDate_ID BETWEEN @SK_SalesDateFrom AND @SK_SalesDateTo AND qry_Outlets.Chain_ID = @ChainID GROUP BY LossesMax.SK_SalesDate_ID, LossesMax.SK_Outlet_ID, LossesMax.SK_Product_ID) AS co" & _
" ON LossesMax.SK_SalesDate_ID = co.SK_SalesDate_ID AND LossesMax.SK_Outlet_ID = co.SK_Outlet_ID AND LossesMax.SK_Product_ID = co.SK_Product_ID" & _
" WHERE LossesMax.SK_SalesDate_ID BETWEEN @SK_SalesDateFrom AND @SK_SalesDateTo AND qry_Outlets.ChainName = @ChN")


 arr = RS.GetRows

Sheets.Add.Name = "Отгрузки"

ArrTransp (arr)
Range(Cells(1, 1), Cells(1, UBound(fildsNames) + 1)).Value = fildsNames


Set RS = Cnn.Execute("SELECT ChainName, BuyerOutletAddress, BuyerOutletCode, SK_Outlet_ID, TransportCode, DeliveryAddress FROM dim.OutletsCollation WHERE ChainName = '" & Buyer & "'")
 arr = RS.GetRows

Sheets.Add.Name = "OutletsCollation"

ArrTransp (arr)

fildsNames = Array("ChainName", "BuyerOutletAddress", "BuyerOutletCode", "SK_Outlet_ID", "TransportCode", "DeliveryAddress")

Range(Cells(1, 1), Cells(1, UBound(fildsNames) + 1)).Value = fildsNames

Set RS = Cnn.Execute("SELECT ChainName,   BuyerProductCode,   BuyerProductName,   SK_Product_ID,  ProductCode, ProductName FROM dim.ProductsCollation WHERE ChainName = '" & Buyer & "'")
 arr = RS.GetRows

Sheets.Add.Name = "ProductsCollation"

ArrTransp (arr)

fildsNames = Array("ChainName", "BuyerProductCode", "BuyerProductName", "SK_Product_ID", "ProductCode", "ProductName")
Range(Cells(1, 1), Cells(1, UBound(fildsNames) + 1)).Value = fildsNames

SetComments Buyer, lastCS
Application.ScreenUpdating = True
End Sub
