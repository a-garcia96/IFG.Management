function New-IFGEmailPurge {
    [cmdletbinding()]
    
    param(
        [Parameter(Mandatory = $true)]
        [string]$ifgUpn,
        [Parameter(Mandatory = $true)]
        [string]$searchName,
        [Parameter(Mandatory = $true)]
        [string]$From,
        [Parameter(Mandatory = $true)]
        [string]$Subject
    )

    $SearchStatus = ""

    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
        Write-Host "Exchange Online Management Module installed. Proceeding with purge." -ForegroundColor Green

        Connect-IPPSSession -UserPrincipalName $ifgUpn

        $Query = "From:" + $From + " AND Subject:" + $Subject

        New-ComplianceSearch -Name $searchName -ExchangeLocation All -ContentMatchQuery $Query

        Start-ComplianceSearch -Identity $searchName

        while($SearchStatus -ne "Completed"){
            $searchStatus = Get-ComplianceSearch -Identity $searchName | Select-Object -ExpandProperty Status
            Write-Host "Search Status: $searchStatus" -ForegroundColor Yellow
            Start-Sleep -s 3
        }

        Write-Host "Search Complete. Reviewing items found." -ForegroundColor Green

        Start-Sleep -s 3

        $ItemsFound = Get-ComplianceSearch -Identity $searchName | Select-Object -ExpandProperty Items

        if($ItemsFound -eq 0) {
            Write-Host "No items found." -ForegroundColor Yellow
            Break
        } else {
            Write-Host "Found a total of $ItemsFound items. Building Preview for review..." -ForegroundColor Green
            
            New-ComplianceSearchAction -SearchName $searchName -Preview

            $PreviewStatus = ""
            $PreviewSearchName = $searchName + "_Preview"

            while($PreviewStatus -ne "Completed"){
                $PreviewStatus = Get-ComplianceSearchAction -Identity $PreviewSearchName | Select-Object -ExpandProperty Status
                Write-Host "Building Preview..." -ForegroundColor Yellow
                Start-Sleep -s 3
            }

            Write-Host "Preview Ready!" -ForegroundColor Green

            $Preview = (Get-ComplianceSearchAction $PreviewSearchName | Select-Object -ExpandProperty Results) -split ","

            Write-Host $Preview -ForegroundColor Green

            Write-Warning "Review results and confirm to continue." -WarningAction Inquire
        }


        New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete


        $PurgeStatus = ""
        $PurgeName = $searchName + "_Purge"

        while($PurgeStatus -ne "Completed"){
            $PurgeStatus = Get-ComplianceSearchAction -Identity $PurgeName | Select-Object -ExpandProperty Status
            Write-Host "Purge in Progress" -ForegroundColor Yellow
            Start-Sleep -s 3
        }

        Write-Host "All items successfully purged. Closing Program..."

        Start-Sleep -s 5

        Break
    }
    else {
        Write-Error "Exchange Online Management Module is not installed. Please install the module first then re-run the script." -Category NotInstalled -ErrorAction Stop
    }
}

Export-ModuleMember -Function New-IFGEmailPurge