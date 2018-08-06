function Invoke-ShipWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID,
        $WeightInLB,
        $PrinterName
    )
    $WarrantyRequest = Get-FreshDeskTicket -ID $FreshDeskWarrantyParentTicketID |
    Where-Object {-Not $_.Deleted} |
    ConvertFrom-FreshDeskTicketToWarrantyRequest
    
    if (-not $WarrantyRequest.ShippingMSN) {
        $ShipParameters = @{
            Company = "$($WarrantyRequest.FirstName) $($WarrantyRequest.LastName)"
            Address1 = $WarrantyRequest.Address1
            Address2 = $WarrantyRequest.Address2
            City = $WarrantyRequest.City        
            StateProvince = $WarrantyRequest.State
            PostalCode = $WarrantyRequest.PostalCode
            Phone = $WarrantyRequest.PhoneNumber
            WeightInLB = $WeightInLB
        } | Remove-HashtableKeysWithEmptyOrNullValues
        
        $ShipmentResult = Invoke-TervisProgisticsReturnsShip @ShipParameters
        if ($ShipmentResult.code -eq 0 -and $ShipmentResult.packageResults.code -eq 0) {
            $CarrierParts = $ShipmentResult.service.symbol -split "\." | Select-Object -First 2
            $Carrier = $CarrierParts -join "."

            Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -custom_fields @{
                cf_shipping_msn = $ShipmentResult.packageResults.resultdata.msn
                cf_tracking_number = $ShipmentResult.packageResults.resultdata.trackingNumber
                cf_shipping_service = $Carrier
            }
        }
        $MSN = $ShipmentResult.packageResults.resultdata.msn
    } else {
        $MSN = $WarrantyRequest.ShippingMSN
        $Carrier = $WarrantyRequest.Carrier
    }

    $Response = Invoke-TervisProgisticsPrint -Carrier $Carrier -MSN $MSN -Output Zebra.Zebra110XiIIIPlus
    $Data = [System.Text.Encoding]::ASCII.GetString($Response.resultdata.output.binaryOutput)
    Send-PrinterData -Data $Data -ComputerName $PrinterName
}