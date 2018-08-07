function Invoke-ShipAndPrintWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID,
        $WeightInLB,
        $PrinterName
    )
    $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID
    
    if (-not $WarrantyRequest.ShippingMSN) {
        Invoke-ShipWarrantyOrder -WarrantyRequest $WarrantyRequest -WeightInLB $WeightInLB
        $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID
    }

    if ($PrintParameters) {
        Invoke-PrintWarrantyOrder -WarrantyRequest $WarrantyRequest -PrinterName $PrinterName
    }
}

function Invoke-ShipWarrantyOrder {
    param (
        $WarrantyRequest,
        $WeightInLB
    )
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
        } | Out-Null
    } else {
        throw "$($ShipmentResult.code) $($ShipmentResult.Message)"
    }
}

function Invoke-PrintWarrantyOrder {
    param (
        [Parameter(Mandatory, ParameterSetName="WarrantyRequest")]$WarrantyRequest,
        [Parameter(Mandatory, ParameterSetName="FreshDeskWarrantyParentTicketID")]$FreshDeskWarrantyParentTicketID,
        [Parameter(Mandatory)]$PrinterName
    )
    if (-not $WarrantyRequest) {
        $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID
    }
    $Response = Invoke-TervisProgisticsPrint -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN -Output Zebra.Zebra110XiIIIPlus
    $Data = [System.Text.Encoding]::ASCII.GetString($Response.resultdata.output.binaryOutput)
    Send-PrinterData -Data $Data -ComputerName $PrinterName
}

function Invoke-UnShipWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID
    )
    $WarrantyRequest = Get-FreshDeskTicket -ID $FreshDeskWarrantyParentTicketID |
    Where-Object {-Not $_.Deleted} |
    ConvertFrom-FreshDeskTicketToWarrantyRequest

    $Response = Remove-ProgisticsShip -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN
    if ($Response.code -eq 0) {
        Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -custom_fields @{
            cf_shipping_msn = $null
            cf_tracking_number = ""
            cf_shipping_service = $null
        }
    }
}