<#
.SYNOPSIS
    Builds an SSRS-style catalog of Power BI workspaces, reports, datasets,
    and datasources using the Power BI Admin Scanner API.

.DESCRIPTION
    - Authenticates interactively
    - Discovers all workspaces (modified in the last 30 days by default)
    - Submits scans in batches of 100 (API limit)
    - Polls until scans complete
    - Flattens results into CSVs:
        Workspaces.csv
        Datasets.csv
        Reports.csv
        Datasources.csv
        Report_To_Datasource.csv  <-- the SSRS-style breakdown

.PREREQUISITES
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser
    Tenant settings enabled:
      - "Enhance admin APIs responses with detailed metadata"
    User must be Fabric Admin / Power BI Admin / Global Admin.
#>

# ---------- CONFIG ----------
$OutputFolder   = "C:\PBICatalog"          # change as needed
$LookbackDays   = 30                        # 1..30; how far back to find modified workspaces. Use 0 for all.
$BatchSize      = 100                       # API max
$PollSeconds    = 5
# ----------------------------

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

Write-Host "Signing in to Power BI..." -ForegroundColor Cyan
Login-PowerBI | Out-Null

# 1. Get modified workspaces
$modifiedUrl = "admin/workspaces/modified?excludePersonalWorkspaces=True&excludeInActiveWorkspaces=True"
if ($LookbackDays -gt 0) {
    $sinceUtc = (Get-Date).ToUniversalTime().AddDays(-$LookbackDays).ToString("yyyy-MM-ddTHH:mm:ss")
    $modifiedUrl += "&modifiedSince=${sinceUtc}.000Z"
}

Write-Host "Fetching workspace list..." -ForegroundColor Cyan
$workspaces = (Invoke-PowerBIRestMethod -Url $modifiedUrl -Method Get | ConvertFrom-Json)
Write-Host "Found $($workspaces.Count) workspaces." -ForegroundColor Green

# 2. Submit scans in batches
$scanIds = @()
for ($i = 0; $i -lt $workspaces.Count; $i += $BatchSize) {
    $batch = $workspaces[$i..([Math]::Min($i + $BatchSize - 1, $workspaces.Count - 1))]
    $body = @{ workspaces = @($batch.id) } | ConvertTo-Json

    $scanUrl = "admin/workspaces/getInfo?lineage=True&datasourceDetails=True&getArtifactUsers=False&datasetSchema=False&datasetExpressions=False"
    $resp = Invoke-PowerBIRestMethod -Url $scanUrl -Method Post -Body $body | ConvertFrom-Json
    $scanIds += $resp.id
    Write-Host "Submitted scan $($resp.id) for $($batch.Count) workspaces." -ForegroundColor Yellow
}

# 3. Poll until all scans succeed
Write-Host "Waiting for scans to complete..." -ForegroundColor Cyan
foreach ($sid in $scanIds) {
    do {
        Start-Sleep -Seconds $PollSeconds
        $status = (Invoke-PowerBIRestMethod -Url "admin/workspaces/scanStatus/$sid" -Method Get | ConvertFrom-Json).status
        Write-Host "  Scan $sid : $status"
    } until ($status -eq "Succeeded")
}

# 4. Pull results and flatten
$wsRows   = New-Object System.Collections.Generic.List[Object]
$dsRows   = New-Object System.Collections.Generic.List[Object]
$rptRows  = New-Object System.Collections.Generic.List[Object]
$srcRows  = New-Object System.Collections.Generic.List[Object]
$linkRows = New-Object System.Collections.Generic.List[Object]

foreach ($sid in $scanIds) {
    $result = Invoke-PowerBIRestMethod -Url "admin/workspaces/scanResult/$sid" -Method Get | ConvertFrom-Json

    foreach ($ws in $result.workspaces) {
        $wsRows.Add([pscustomobject]@{
            WorkspaceId   = $ws.id
            WorkspaceName = $ws.name
            Type          = $ws.type
            State         = $ws.state
            IsOnDedicatedCapacity = $ws.isOnDedicatedCapacity
            CapacityId    = $ws.capacityId
        })

        foreach ($d in $ws.datasets) {
            $dsRows.Add([pscustomobject]@{
                WorkspaceId       = $ws.id
                WorkspaceName     = $ws.name
                DatasetId         = $d.id
                DatasetName       = $d.name
                ConfiguredBy      = $d.configuredBy
                CreatedDate       = $d.createdDate
                ContentProviderType = $d.contentProviderType
            })

            foreach ($src in $d.datasources) {
                $srcRows.Add([pscustomobject]@{
                    WorkspaceId        = $ws.id
                    DatasetId          = $d.id
                    DatasetName        = $d.name
                    DatasourceId       = $src.datasourceId
                    DatasourceType     = $src.datasourceType
                    GatewayId          = $src.gatewayId
                    Server             = $src.connectionDetails.server
                    Database           = $src.connectionDetails.database
                    Url                = $src.connectionDetails.url
                    Path               = $src.connectionDetails.path
                    ConnectionString   = ($src.connectionDetails | ConvertTo-Json -Compress)
                })
            }
        }

        foreach ($r in $ws.reports) {
            $rptRows.Add([pscustomobject]@{
                WorkspaceId   = $ws.id
                WorkspaceName = $ws.name
                ReportId      = $r.id
                ReportName    = $r.name
                ReportType    = $r.reportType
                DatasetId     = $r.datasetId
                CreatedBy     = $r.createdBy
                ModifiedBy    = $r.modifiedBy
                ModifiedDateTime = $r.modifiedDateTime
            })

            # Build report-to-datasource link rows by joining via datasetId
            $matchedDs = $ws.datasets | Where-Object { $_.id -eq $r.datasetId }
            foreach ($d in $matchedDs) {
                foreach ($src in $d.datasources) {
                    $linkRows.Add([pscustomobject]@{
                        WorkspaceName  = $ws.name
                        ReportName     = $r.name
                        ReportId       = $r.id
                        DatasetName    = $d.name
                        DatasetId      = $d.id
                        DatasourceType = $src.datasourceType
                        Server         = $src.connectionDetails.server
                        Database       = $src.connectionDetails.database
                        Url            = $src.connectionDetails.url
                        GatewayId      = $src.gatewayId
                    })
                }
            }
        }
    }
}

# 5. Export
$wsRows   | Export-Csv "$OutputFolder\Workspaces.csv"          -NoTypeInformation -Encoding UTF8
$dsRows   | Export-Csv "$OutputFolder\Datasets.csv"            -NoTypeInformation -Encoding UTF8
$rptRows  | Export-Csv "$OutputFolder\Reports.csv"             -NoTypeInformation -Encoding UTF8
$srcRows  | Export-Csv "$OutputFolder\Datasources.csv"         -NoTypeInformation -Encoding UTF8
$linkRows | Export-Csv "$OutputFolder\Report_To_Datasource.csv" -NoTypeInformation -Encoding UTF8

Write-Host "`nDone. CSVs written to $OutputFolder" -ForegroundColor Green
Write-Host "  Workspaces: $($wsRows.Count)"
Write-Host "  Datasets:   $($dsRows.Count)"
Write-Host "  Reports:    $($rptRows.Count)"
Write-Host "  Sources:    $($srcRows.Count)"
Write-Host "  Links:      $($linkRows.Count)"
