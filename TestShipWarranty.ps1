ipmo -force ProgisticsPowerShell, TervisProgisticsPowerShell, ShipWarranty, TervisWarrantyRequest
Set-TervisFreshDeskEnvironment
Set-TervisProgisticsEnvironment -Name delta


$FreshDeskWarrantyParentTicketID = 68942
$WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID 68942 -WeightInLB .1 -PrinterName "cheers"
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID 68942 -WeightInLB 1.1 -PrinterName "cheers"
Invoke-ShipAndPrintWarrantyOrder -FreshDeskWarrantyParentTicketID 68942 -WeightInLB 10.1 -PrinterName "cheers"
Invoke-UnShipWarrantyOrder -FreshDeskWarrantyParentTicketID 68942

$result = Find-ProgisticsPackage -TrackingNumber 445622761400 -carrier TANDATA_FEDEXFSMS.FEDEX  |
Select-Object -ExpandProperty ResultData |
Select-Object -ExpandProperty ResultData |
Select-Object -ExpandProperty Service

$Result = Find-ProgisticsPackage -MSN 101320200 -carrier TANDATA_FEDEXFSMS.FEDEX
$Result.resultData