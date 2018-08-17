ipmo -force ProgisticsPowerShell, TervisProgisticsPowerShell, ShipWarranty, TervisWarrantyRequest
Set-TervisFreshDeskEnvironment
Set-TervisProgisticsEnvironment -Name delta


$FreshDeskWarrantyParentTicketID = 81263
$WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID -WeightInLB .1 -PrinterName "ClevelandState"
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID -WeightInLB 1.1 -PrinterName "ClevelandState"
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID -WeightInLB 10.1 -PrinterName "ClevelandState"
Invoke-UnShipWarrantyOrder -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID

$result = Find-ProgisticsPackage -TrackingNumber 445622761400 -carrier TANDATA_FEDEXFSMS.FEDEX  |
Select-Object -ExpandProperty ResultData |
Select-Object -ExpandProperty ResultData |
Select-Object -ExpandProperty Service

$Result = Find-ProgisticsPackage -MSN 101320200 -carrier TANDATA_FEDEXFSMS.FEDEX
$Result.resultData