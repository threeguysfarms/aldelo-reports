Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adClipString = 2
Const ForWriting = 2
Const ForAppending = 8

Const db = "C:\ProgramData\Aldelo\Aldelo For Restaurants\Databases\MonteVilla\MonteVilla.mdb"

Dim theDate
theDate = Replace(CStr(Date() - 1), "/", "-")

Dim TextExportFile 
TextExportFile = "C:\ProgramData\Aldelo\Aldelo For Restaurants\Export\Reports\transactions-export-" & theDate & ".csv"

Set cn = CreateObject("ADODB.Connection")
Set rs = CreateObject("ADODB.Recordset")

cn.Open _
   "Provider = Microsoft.Jet.OLEDB.4.0; " & _
   "Data Source =" & db

strSQL = "SELECT o.OrderId, o.OrderDateTime, o.AmountDue, o.SubTotal,"  & _
  "p.OrderPaymentID, p.PaymentDateTime, " & _
  "p.PaymentMethod, " & _
  "Switch(p.PaymentMethod=""1"",""CASH"",p.PaymentMethod=""3"",""VISA"",p.PaymentMethod=""4"",""MASTERCARD"",p.PaymentMethod=""5"",""AMEX"",p.PaymentMethod=""6"",""CC-DISCOVER"",True,""?"") AS PaymentMethodName, " & _
  "p.AmountTendered, p.AmountPaid, p.EDCTransID," & _
  "t.MenuItemId, t.ExtendedPrice," & _
  "i.MenuItemText," & _
  "c.MenuCategoryText " & _
  "FROM OrderHeaders o, OrderPayments p, OrderTransactions t, MenuItems i, MenuCategories c " & _
  "WHERE o.OrderId = p.OrderId " & _
  "AND o.OrderId = t.OrderId " & _
  "AND t.MenuItemId = i.MenuItemId " & _
  "AND i.MenuCategoryId = c.MenuCategoryId " & _
  "AND o.OrderDateTime >= Date() - 1;"

rs.Open strSQL, cn, 3, 3

Set objCSV = CreateObject("Scripting.FileSystemObject").OpenTextFile(TextExportFile, ForWriting, True)

Set objLog = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\Elo\Desktop\export-daily-transactions.vbs.log", ForAppending, True)

objLog.Write "---start processing for: " & theDate & Chr(13) & Chr(10)

If (rs.RecordCount > 0) Then
  objCSV.WriteLine "order_id, order_datetime, order_amount_due, order_subtotal," & _
    "payment_id, payment_datetime, payment_method_id, payment_method_name, payment_amount_tendered, payment_amount_paid, payment_edc_trans_id," & _
    "transaction_menuitem_id, transaction_extended_price, menu_item_name, menu_item_category_name" 
  objCSV.Write rs.GetString(adClipString,,",",CRLF)

  objCSV.Close
  Set objCSV = Nothing
  rs.Close
  Set rs = Nothing
  cn.Close
  Set cn = Nothing

  Dim objShell
  Set objShell = WScript.CreateObject ("WScript.shell")
  Dim aRet: Set aRet = objShell.exec("C:\work\bin\gdrive.exe upload --parent ""1z-PmgneqidqFG109SqD_ohw7f-yaMbvB"" """ & TextExportFile & """")

  Do
    Dim strFromProc
    strFromProc = aRet.StdOut.ReadLine()
    objLog.Write strFromProc & Chr(13) & Chr(10)
  Loop While Not aRet.Stdout.atEndOfStream
Else
  objLog.Write "---There were no records to export for: " & theDate & Chr(13) & Chr(10)
End If

objLog.Write "---end processing for: " & theDate & Chr(13) & Chr(10)

Set objShell = Nothing
