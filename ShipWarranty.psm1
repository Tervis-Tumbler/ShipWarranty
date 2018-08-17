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
    Invoke-PrintWarrantyOrder -WarrantyRequest $WarrantyRequest -PrinterName $PrinterName
}

function ConvertFrom-WarrantyRequestToShipmentParameters {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]$WarrantyRequest,
        $WeightInLB
    )
    process {
        @{
            Company = "$($WarrantyRequest.FirstName) $($WarrantyRequest.LastName)"
            Contact = $WarrantyRequest.BusinessName
            Address1 = $WarrantyRequest.Address1
            Address2 = $WarrantyRequest.Address2
            City = $WarrantyRequest.City        
            StateProvince = $WarrantyRequest.State
            PostalCode = $WarrantyRequest.PostalCode
            Phone = $WarrantyRequest.PhoneNumber
            WeightInLB = $WeightInLB
        } | Remove-HashtableKeysWithEmptyOrNullValues
    }
}

function Invoke-ShipWarrantyOrder {
    param (
        $WarrantyRequest,
        $WeightInLB
    )
    $ShipmentParameters = $WarrantyRequest | ConvertFrom-WarrantyRequestToShipmentParameters -WeightInLB $WeightInLB
    $ShipmentResult = New-TervisProgisticsPackageShipmentWarrantyOrder @ShipmentParameters

    if ($ShipmentResult.code -eq 0 -and $ShipmentResult.packageResults.code -eq 0) {
        $CarrierParts = $ShipmentResult.service.symbol -split "\." | Select-Object -First 2
        $Carrier = $CarrierParts -join "."
        try {
            Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -status 5 -custom_fields @{
                cf_shipping_msn = $ShipmentResult.packageResults.resultdata.msn
                cf_tracking_number = $ShipmentResult.packageResults.resultdata.trackingNumber
                cf_shipping_service = $Carrier
            }
        } catch {
            Invoke-UnShipWarrantyOrder -FreshDeskWarrantyParentTicketID $WarrantyRequest.ID
            throw "Check to confirm all children are closed. Unable to close ticket and set properties. Shipment has been voided."            
        }
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
    $Response = Invoke-TervisProgisticsPackagePrintWarrantyOrder -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN -Output Zebra.Zebra110XiIIIPlus -TrackingNumber $WarrantyRequest.TrackingNumber
    $Data = [System.Text.Encoding]::ASCII.GetString($Response.resultdata.output.binaryOutput)
    Send-PrinterData -Data $Data -ComputerName $PrinterName
}

function Invoke-UnShipWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID
    )
    $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID

    if ($WarrantyRequest.Carrier) {
        $Response = Remove-ProgisticsPackage -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN
        if ($Response.code -eq 0) {
            Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -status 2 -custom_fields @{
                cf_shipping_msn = $null
                cf_tracking_number = ""
                cf_shipping_service = $null
            }
        } else {
            Throw "$($Response.code) $($Response.Message) $($Response.resultdata.code) $($Response.resultdata.message)"
        }
    }
}