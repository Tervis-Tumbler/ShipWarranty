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
        [Parameter(Mandatory,ValueFromPipeline)]$WarrantyRequest
    )
    process {
        @{
            Company = "$($WarrantyRequest.FirstName) $($WarrantyRequest.LastName)"
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
    $ShipmentParameters = $WarrantyRequest | ConvertFrom-WarrantyRequestToShipmentParameters
    $ShipmentResult = New-TervisProgisticsPackageShipmentWarrantyOrder @ShipmentParameters

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
    $Response = Invoke-TervisProgisticsPackagePrintWarrantyOrder -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN -Output Zebra.Zebra110XiIIIPlus
    $Data = [System.Text.Encoding]::ASCII.GetString($Response.resultdata.output.binaryOutput)
    Send-PrinterData -Data $Data -ComputerName $PrinterName
}

function Invoke-UnShipWarrantyOrder {
    param (
        $FreshDeskWarrantyParentTicketID
    )
    $WarrantyRequest = Get-WarrantyRequest -FreshDeskWarrantyParentTicketID $FreshDeskWarrantyParentTicketID

    $Response = Remove-ProgisticsPackage -Carrier $WarrantyRequest.Carrier -MSN $WarrantyRequest.ShippingMSN
    if ($Response.code -eq 0) {
        Set-FreshDeskTicket -id $FreshDeskWarrantyParentTicketID -custom_fields @{
            cf_shipping_msn = $null
            cf_tracking_number = ""
            cf_shipping_service = $null
        }
    }
}