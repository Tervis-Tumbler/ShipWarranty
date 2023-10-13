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
            Name = "$($WarrantyRequest.FirstName) $($WarrantyRequest.LastName)"
            CompanyName = $WarrantyRequest.BusinessName
            AddressLine1 = $WarrantyRequest.Address1
            AddressLine2 = $WarrantyRequest.Address2
            CityLocality = $WarrantyRequest.City        
            StateProvince = $WarrantyRequest.State
            PostalCode = $WarrantyRequest.PostalCode
            CountryCode = "US"
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
    $ShipmentResult = Invoke-TervisShipEngineShipWarrantyOrder @ShipmentParameters

    if ($ShipmentResult.Status -eq 200) {
        try {
            Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -status 5 -custom_fields @{
                cf_shipping_msn = $ShipmentResult.Content.label_id.substring(3) / 1
                cf_tracking_number = $ShipmentResult.Content.tracking_number
                cf_shipping_service = $ShipmentResult.Content.service_code
            }
        } catch {
            # Invoke-UnShipWarrantyOrder -FreshDeskWarrantyParentTicketID $WarrantyRequest.ID
            $Response = Remove-TervisShipEngineLabel -LabelId $ShipmentResult.Content.label_id
            throw "Check to confirm all children are closed. Unable to close ticket and set properties. $($Response.Content.message)"
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
    $Data = Get-TervisShipEngineLabel -LabelId $WarrantyRequest.ShippingMSN
    Send-PrinterData -Data $Data -ComputerName $PrinterName
}

function Invoke-UnShipWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID
    )
    $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID

    if ($WarrantyRequest.Carrier) {
        $Response = Remove-TervisShipEngineLabel -LabelId "se-$($WarrantyRequest.ShippingMSN)"
        if ($Response.Status -eq 200) {
            Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -status 2 -custom_fields @{
                cf_shipping_msn = $null
                cf_tracking_number = ""
                cf_shipping_service = $null
            }
        } else {
            Throw "$($Response.Status) $($Response.Content.errors.message) [$($Response.Content.errors.error_type)][$($Response.Content.errors.error_code)][$($Response.Content.errors.path)]"
        }
    }
}