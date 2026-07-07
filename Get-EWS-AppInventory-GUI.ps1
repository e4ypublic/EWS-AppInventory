<#
.SYNOPSIS
    Moderne GUI-Auswertung für App Registrations und Enterprise Applications mit EWS-Bezug.

.DESCRIPTION
    Gesamtauswertung nach AppId.

.OUTPUT
    Type, Name, AppId, ObjectIds, Permissions, PermissionTypes, Status, Enabled

.NOTES
    Start:
    powershell.exe -STA -ExecutionPolicy Bypass -File .\Get-EWS-AppInventory-GUI.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ------------------------------------------------------------
# Constants
# ------------------------------------------------------------

$Script:ExchangeAppId = "00000002-0000-0ff1-ce00-000000000000"

$Script:EwsPermissionMap = @{
    "dc890d15-9560-4a4c-9b7f-a736ec74ec40" = @{
        Name = "full_access_as_app"
        Type = "Application"
    }
    "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5" = @{
        Name = "EWS.AccessAsUser.All"
        Type = "Delegated"
    }
}

$Script:EwsPermissionIds = @(
    "dc890d15-9560-4a4c-9b7f-a736ec74ec40",
    "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5"
)

$Script:RawResults = New-Object System.Collections.Generic.List[object]
$Script:DisplayResults = @()
$Script:SpCacheByObjectId = @{}
$Script:SpCacheByAppId = @{}

# ------------------------------------------------------------
# Theme
# ------------------------------------------------------------

$ColorBackground  = [System.Drawing.Color]::FromArgb(245,247,250)
$ColorPanel       = [System.Drawing.Color]::White
$ColorPrimary     = [System.Drawing.Color]::FromArgb(0,120,212)
$ColorTextDark    = [System.Drawing.Color]::FromArgb(32,32,32)
$ColorSubText     = [System.Drawing.Color]::FromArgb(90,90,90)
$ColorGridAlt     = [System.Drawing.Color]::FromArgb(248,250,252)
$ColorBorder      = [System.Drawing.Color]::FromArgb(220,223,228)

# ------------------------------------------------------------
# Helper
# ------------------------------------------------------------

function Fix-Umlaut {
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }

    return [System.Text.Encoding]::UTF8.GetString(
        [System.Text.Encoding]::Default.GetBytes($Text)
    )
}

function Write-GuiLog {
    param([string]$Message)

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    $fixedMessage = Fix-Umlaut $Message

    if ($null -ne $txtLog) {
        $txtLog.AppendText("[$timestamp] $fixedMessage`r`n")
        $txtLog.SelectionStart = $txtLog.Text.Length
        $txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Get-PermissionName {
    param([string]$PermissionId)

    if ($Script:EwsPermissionMap.ContainsKey($PermissionId)) {
        return $Script:EwsPermissionMap[$PermissionId].Name
    }

    return $PermissionId
}

function Get-PermissionType {
    param([string]$PermissionId)

    if ($Script:EwsPermissionMap.ContainsKey($PermissionId)) {
        return $Script:EwsPermissionMap[$PermissionId].Type
    }

    return ""
}

function Add-RawResult {
    param(
        [string]$Type,
        [string]$Name,
        [string]$AppId,
        [string]$ObjectId,
        [string]$Permission,
        [string]$PermissionType,
        [string]$Status,
        [object]$Enabled
    )

    if ([string]::IsNullOrWhiteSpace($AppId)) {
        return
    }

    $Script:RawResults.Add([PSCustomObject]@{
        Type           = $Type
        Name           = $Name
        AppId          = $AppId
        ObjectId       = $ObjectId
        Permission     = $Permission
        PermissionType = $PermissionType
        Status         = $Status
        Enabled        = $Enabled
    })
}

function Ensure-MgModules {
    Write-GuiLog "Prüfe Microsoft Graph PowerShell Module..."

    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Applications"
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            throw "Microsoft Graph Modul '$module' fehlt. Installation: Install-Module Microsoft.Graph -Scope CurrentUser"
        }
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Write-GuiLog "Microsoft Graph Module geladen."
}

function Connect-Graph {
    Ensure-MgModules

    $requiredScopes = @(
        "Application.Read.All",
        "Directory.Read.All"
    )

    $ctx = Get-MgContext
    $needsConnect = $true

    if ($null -ne $ctx) {
        $missingScopes = @()

        foreach ($scope in $requiredScopes) {
            if ($ctx.Scopes -notcontains $scope) {
                $missingScopes += $scope
            }
        }

        if ($missingScopes.Count -eq 0) {
            $needsConnect = $false
            Write-GuiLog "Microsoft Graph ist bereits verbunden."
        }
        else {
            Write-GuiLog "Graph verbunden, aber Scopes fehlen: $($missingScopes -join ', ')"
            Disconnect-MgGraph | Out-Null
            $needsConnect = $true
        }
    }

    if ($needsConnect) {
        Write-GuiLog "Verbinde zu Microsoft Graph..."
        Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    }

    $ctx = Get-MgContext

    if ($null -eq $ctx) {
        throw "Microsoft Graph Verbindung konnte nicht hergestellt werden."
    }

    Write-GuiLog "Microsoft Graph verbunden. TenantId: $($ctx.TenantId)"
}

function Resolve-ServicePrincipalByAppId {
    param([string]$AppId)

    if ([string]::IsNullOrWhiteSpace($AppId)) {
        return $null
    }

    if ($Script:SpCacheByAppId.ContainsKey($AppId)) {
        return $Script:SpCacheByAppId[$AppId]
    }

    try {
        $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" -ErrorAction Stop

        if ($null -ne $sp) {
            $Script:SpCacheByAppId[$AppId] = $sp
            $Script:SpCacheByObjectId[$sp.Id] = $sp
        }

        return $sp
    }
    catch {
        return $null
    }
}

function Resolve-ServicePrincipalByObjectId {
    param([string]$ObjectId)

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $null
    }

    if ($Script:SpCacheByObjectId.ContainsKey($ObjectId)) {
        return $Script:SpCacheByObjectId[$ObjectId]
    }

    try {
        $sp = Get-MgServicePrincipal -ServicePrincipalId $ObjectId -ErrorAction Stop

        if ($null -ne $sp) {
            $Script:SpCacheByObjectId[$ObjectId] = $sp
            $Script:SpCacheByAppId[$sp.AppId] = $sp
        }

        return $sp
    }
    catch {
        Write-GuiLog "WARNUNG: Service Principal konnte nicht aufgelöst werden: $ObjectId"
        return $null
    }
}

function Invoke-GraphGetAll {
    param([string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri

    while (-not [string]::IsNullOrWhiteSpace($next)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop

        if ($response.value) {
            foreach ($item in $response.value) {
                $items.Add($item)
            }
        }

        if ($response.'@odata.nextLink') {
            $next = $response.'@odata.nextLink'
        }
        else {
            $next = $null
        }
    }

    return $items
}

function Get-TenantInfo {
    try {
        $org = Get-MgOrganization -ErrorAction Stop

        if ($org) {
            return [PSCustomObject]@{
                Name = $org.DisplayName
                TenantId = $org.Id
            }
        }
    }
    catch {
        return $null
    }
}

# ------------------------------------------------------------
# Inventory Logic
# ------------------------------------------------------------

function Get-EwsAppRegistrations {
    Write-GuiLog "Prüfe App Registrations auf EWS RequiredResourceAccess..."

    $apps = Get-MgApplication -All
    $count = 0

    foreach ($app in $apps) {
        if ($null -eq $app.RequiredResourceAccess) {
            continue
        }

        foreach ($rra in $app.RequiredResourceAccess) {
            if ($rra.ResourceAppId -ne $Script:ExchangeAppId) {
                continue
            }

            foreach ($ra in $rra.ResourceAccess) {
                $permissionId = $ra.Id.ToString()

                if ($permissionId -notin $Script:EwsPermissionIds) {
                    continue
                }

                $permissionName = Get-PermissionName -PermissionId $permissionId
                $permissionType = Get-PermissionType -PermissionId $permissionId

                $relatedSp = Resolve-ServicePrincipalByAppId -AppId $app.AppId

                Add-RawResult `
                    -Type "AppRegistration" `
                    -Name $app.DisplayName `
                    -AppId $app.AppId `
                    -ObjectId $app.Id `
                    -Permission $permissionName `
                    -PermissionType $permissionType `
                    -Status "Configured" `
                    -Enabled $(if ($relatedSp) { $relatedSp.AccountEnabled } else { "" })

                $count++
            }
        }
    }

    Write-GuiLog "App Registration Treffer gefunden: $count"
}

function Get-EwsServicePrincipalDefinitions {
    Write-GuiLog "Prüfe Service Principals auf EWS Permission Definitions..."

    $servicePrincipals = Get-MgServicePrincipal -All
    $count = 0

    foreach ($sp in $servicePrincipals) {
        if ($null -eq $sp.AppId) {
            continue
        }

        foreach ($appRole in $sp.AppRoles) {
            if ($appRole.Value -eq "full_access_as_app") {
                Add-RawResult `
                    -Type "ServicePrincipalDefinition" `
                    -Name $sp.DisplayName `
                    -AppId $sp.AppId `
                    -ObjectId $sp.Id `
                    -Permission "full_access_as_app" `
                    -PermissionType "Application" `
                    -Status "Permission definition" `
                    -Enabled $sp.AccountEnabled

                $count++
            }
        }

        foreach ($scope in $sp.Oauth2PermissionScopes) {
            if ($scope.Value -eq "EWS.AccessAsUser.All") {
                Add-RawResult `
                    -Type "ServicePrincipalDefinition" `
                    -Name $sp.DisplayName `
                    -AppId $sp.AppId `
                    -ObjectId $sp.Id `
                    -Permission "EWS.AccessAsUser.All" `
                    -PermissionType "Delegated" `
                    -Status "Permission definition" `
                    -Enabled $sp.AccountEnabled

                $count++
            }
        }
    }

    Write-GuiLog "Service Principal Definition Treffer gefunden: $count"
}

function Get-ExchangeOnlineServicePrincipal {
    Write-GuiLog "Suche Exchange Online Service Principal..."

    $exchangeSp = Get-MgServicePrincipal -Filter "appId eq '$($Script:ExchangeAppId)'" -ErrorAction Stop

    if ($null -eq $exchangeSp) {
        throw "Exchange Online Service Principal wurde nicht gefunden."
    }

    Write-GuiLog "Exchange Online Service Principal gefunden: $($exchangeSp.DisplayName)"
    return $exchangeSp
}

function Get-EwsApplicationConsent {
    param([object]$ExchangeSp)

    Write-GuiLog "Prüfe Enterprise Apps auf Application Consent full_access_as_app..."

    $ewsApplicationPermissionId = "dc890d15-9560-4a4c-9b7f-a736ec74ec40"
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($ExchangeSp.Id)/appRoleAssignedTo?`$top=999"

    $assignments = Invoke-GraphGetAll -Uri $uri
    $count = 0

    foreach ($assignment in $assignments) {
        $appRoleId = $assignment.appRoleId.ToString()

        if ($appRoleId -ne $ewsApplicationPermissionId) {
            continue
        }

        $clientSp = Resolve-ServicePrincipalByObjectId -ObjectId $assignment.principalId

        Add-RawResult `
            -Type "EnterpriseApp" `
            -Name $(if ($clientSp) { $clientSp.DisplayName } else { $assignment.principalDisplayName }) `
            -AppId $(if ($clientSp) { $clientSp.AppId } else { "" }) `
            -ObjectId $assignment.principalId `
            -Permission "full_access_as_app" `
            -PermissionType "Application" `
            -Status "Application consent granted" `
            -Enabled $(if ($clientSp) { $clientSp.AccountEnabled } else { "" })

        $count++
    }

    Write-GuiLog "Enterprise App Application Consent Treffer gefunden: $count"
}

function Get-EwsDelegatedConsent {
    param([object]$ExchangeSp)

    Write-GuiLog "Prüfe Enterprise Apps auf Delegated Consent EWS.AccessAsUser.All..."

    $ewsDelegatedScopeName = "EWS.AccessAsUser.All"
    $filter = [System.Uri]::EscapeDataString("resourceId eq '$($ExchangeSp.Id)'")
    $uri = "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=$filter&`$top=999"

    $grants = Invoke-GraphGetAll -Uri $uri
    $count = 0

    foreach ($grant in $grants) {
        if ([string]::IsNullOrWhiteSpace($grant.scope)) {
            continue
        }

        $scopes = $grant.scope -split " "

        if ($scopes -notcontains $ewsDelegatedScopeName) {
            continue
        }

        $clientSp = Resolve-ServicePrincipalByObjectId -ObjectId $grant.clientId

        Add-RawResult `
            -Type "EnterpriseApp" `
            -Name $(if ($clientSp) { $clientSp.DisplayName } else { $grant.clientId }) `
            -AppId $(if ($clientSp) { $clientSp.AppId } else { "" }) `
            -ObjectId $grant.clientId `
            -Permission "EWS.AccessAsUser.All" `
            -PermissionType "Delegated" `
            -Status "Delegated consent granted" `
            -Enabled $(if ($clientSp) { $clientSp.AccountEnabled } else { "" })

        $count++
    }

    Write-GuiLog "Enterprise App Delegated Consent Treffer gefunden: $count"
}

# ------------------------------------------------------------
# Deduplication / Display Logic
# ------------------------------------------------------------

function Build-DeduplicatedResults {
    Write-GuiLog "Dedupliziere Ergebnisse nach AppId..."

    $groups = $Script:RawResults |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.AppId) } |
        Group-Object AppId

    $deduped = foreach ($group in $groups) {
        $items = $group.Group

        $preferredName = (
            $items |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.Name) } |
            Select-Object -First 1
        ).Name

        $types = (
            $items.Type |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        ) -join " + "

        $objectIds = (
            $items.ObjectId |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        ) -join "; "

        $permissions = (
            $items.Permission |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        ) -join "; "

        $permissionTypes = (
            $items.PermissionType |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        ) -join "; "

        $statuses = (
            $items.Status |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
        ) -join "; "

        $enabledValues = $items.Enabled | Where-Object { $_ -ne $null -and $_ -ne "" }

        if ($enabledValues -contains $true) {
            $enabled = $true
        }
        elseif ($enabledValues -contains $false) {
            $enabled = $false
        }
        else {
            $enabled = ""
        }

        [PSCustomObject]@{
            Type            = $types
            Name            = $preferredName
            AppId           = $group.Name
            ObjectIds       = $objectIds
            Permissions     = $permissions
            PermissionTypes = $permissionTypes
            Status          = $statuses
            Enabled         = $enabled
        }
    }

    $Script:DisplayResults = $deduped | Sort-Object Name, AppId

    Write-GuiLog "Deduplizierte Apps: $($Script:DisplayResults.Count)"
}

# ------------------------------------------------------------
# GUI Actions
# ------------------------------------------------------------

function Update-SummaryCards {
    $total = $Script:DisplayResults.Count

    $appRegs = (
        $Script:DisplayResults |
        Where-Object { $_.Type -like "*AppRegistration*" }
    ).Count

    $enterprise = (
        $Script:DisplayResults |
        Where-Object { $_.Type -like "*EnterpriseApp*" }
    ).Count

    $definitions = (
        $Script:DisplayResults |
        Where-Object { $_.Type -like "*ServicePrincipalDefinition*" }
    ).Count

    $lblTotalValue.Text = $total.ToString()
    $lblAppRegValue.Text = $appRegs.ToString()
    $lblEnterpriseValue.Text = $enterprise.ToString()
    $lblDefinitionValue.Text = $definitions.ToString()
}

function Refresh-Grid {
    $filter = $txtFilter.Text

    $data = $Script:DisplayResults

    if (-not [string]::IsNullOrWhiteSpace($filter)) {
        $data = $data | Where-Object {
            $_.Name -like "*$filter*" -or
            $_.AppId -like "*$filter*" -or
            $_.Permissions -like "*$filter*" -or
            $_.Status -like "*$filter*" -or
            $_.Type -like "*$filter*"
        }
    }

    $bindingSource = New-Object System.Windows.Forms.BindingSource
    $bindingSource.DataSource = [System.Collections.ArrayList]@($data)

    $grid.DataSource = $bindingSource

    if ($grid.Columns.Count -gt 0) {
        $grid.Columns["Type"].Width = 260
        $grid.Columns["Name"].Width = 330
        $grid.Columns["AppId"].Width = 275
        $grid.Columns["ObjectIds"].Width = 370
        $grid.Columns["Permissions"].Width = 260
        $grid.Columns["PermissionTypes"].Width = 160
        $grid.Columns["Status"].Width = 330
        $grid.Columns["Enabled"].Width = 80
    }

    Update-SummaryCards
}

function Start-EwsInventory {
    try {
        $btnStart.Enabled = $false
        $btnExport.Enabled = $false
        $txtLog.Clear()
        $grid.DataSource = $null

        $Script:RawResults.Clear()
        $Script:DisplayResults = @()
        $Script:SpCacheByObjectId.Clear()
        $Script:SpCacheByAppId.Clear()

        Connect-Graph

        $tenantInfo = Get-TenantInfo

        if ($tenantInfo) {
            $lblTenant.Text = "Tenant: $($tenantInfo.Name) ($($tenantInfo.TenantId))"
        }
        else {
            $lblTenant.Text = "Tenant: unbekannt"
        }

        if ($chkAppRegistrations.Checked) {
            Get-EwsAppRegistrations
        }

        if ($chkServicePrincipalDefinitions.Checked) {
            Get-EwsServicePrincipalDefinitions
        }

        $exchangeSp = $null

        if ($chkApplicationConsent.Checked -or $chkDelegatedConsent.Checked) {
            $exchangeSp = Get-ExchangeOnlineServicePrincipal
        }

        if ($chkApplicationConsent.Checked) {
            Get-EwsApplicationConsent -ExchangeSp $exchangeSp
        }

        if ($chkDelegatedConsent.Checked) {
            Get-EwsDelegatedConsent -ExchangeSp $exchangeSp
        }

        Write-GuiLog "Roh-Treffer vor Deduplizierung: $($Script:RawResults.Count)"

        Build-DeduplicatedResults
        Refresh-Grid

        Write-GuiLog "Fertig. Eindeutige Apps: $($Script:DisplayResults.Count)"
        $btnExport.Enabled = $true
    }
    catch {
        Write-GuiLog "FEHLER: $($_.Exception.Message)"

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Fehler",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    finally {
        $btnStart.Enabled = $true
    }
}

function Export-EwsInventory {
    if ($Script:DisplayResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Es sind keine Ergebnisse vorhanden.",
            "CSV Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null

        return
    }

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Datei (*.csv)|*.csv"
    $saveDialog.FileName = "EWS_AppInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

    if ($saveDialog.ShowDialog() -eq "OK") {
        $Script:DisplayResults |
            Sort-Object Name, AppId |
            Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

        Write-GuiLog "CSV exportiert: $($saveDialog.FileName)"

        [System.Windows.Forms.MessageBox]::Show(
            "CSV Export abgeschlossen:`r`n$($saveDialog.FileName)",
            "CSV Export",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
}

# ------------------------------------------------------------
# GUI Helper
# ------------------------------------------------------------

function New-Card {
    param(
        [string]$Title,
        [int]$X,
        [int]$Y,
        [ref]$ValueLabelRef
    )

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point($X, $Y)
    $panel.Size = New-Object System.Drawing.Size(230, 78)
    $panel.BackColor = $ColorPanel
    $panel.BorderStyle = "FixedSingle"

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = $Title
    $lblTitle.Location = New-Object System.Drawing.Point(14, 10)
    $lblTitle.Size = New-Object System.Drawing.Size(200, 20)
    $lblTitle.ForeColor = $ColorSubText
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $panel.Controls.Add($lblTitle)

    $lblValue = New-Object System.Windows.Forms.Label
    $lblValue.Text = "0"
    $lblValue.Location = New-Object System.Drawing.Point(14, 33)
    $lblValue.Size = New-Object System.Drawing.Size(200, 34)
    $lblValue.ForeColor = $ColorPrimary
    $lblValue.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($lblValue)

    $ValueLabelRef.Value = $lblValue

    return $panel
}

function Disconnect-GraphSafe {

    try {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue

        if ($null -eq $ctx) {
            Write-GuiLog "Keine aktive Graph Sitzung vorhanden."
            return
        }

        Write-GuiLog "Trenne Microsoft Graph Verbindung..."

        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
        }
        catch {
            Write-GuiLog "WARNUNG: Disconnect-MgGraph fehlgeschlagen: $($_.Exception.Message)"
        }

        try {
            Remove-MgGraphContext -ErrorAction SilentlyContinue
        }
        catch {
            Write-GuiLog "WARNUNG: Remove-MgGraphContext fehlgeschlagen (nicht kritisch)"
        }

        Write-GuiLog "Microsoft Graph Sitzung wurde beendet."
    }
    catch {
        Write-GuiLog "FEHLER beim Logout: $($_.Exception.Message)"
    }
}



# ------------------------------------------------------------
# GUI Definition
# ------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "EWS App Inventory"
$form.Size = New-Object System.Drawing.Size(1450, 880)
$form.StartPosition = "CenterScreen"
$form.BackColor = $ColorBackground
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Header
$header = New-Object System.Windows.Forms.Panel
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.Size = New-Object System.Drawing.Size(1450, 88)
$header.BackColor = $ColorPrimary
$header.Anchor = "Top,Left,Right"
$form.Controls.Add($header)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "EWS App Inventory"
$lblTitle.Location = New-Object System.Drawing.Point(24, 16)
$lblTitle.Size = New-Object System.Drawing.Size(700, 30)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$header.Controls.Add($lblTitle)

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = "Exchange Web Services Zugriffe nach AppId"
$lblSubtitle.Location = New-Object System.Drawing.Point(26, 50)
$lblSubtitle.Size = New-Object System.Drawing.Size(900, 22)
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(230,240,255)
$lblSubtitle.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$header.Controls.Add($lblSubtitle)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = "Tenant: nicht verbunden"
$lblTenant.Location = New-Object System.Drawing.Point(26, 68)
$lblTenant.Size = New-Object System.Drawing.Size(1200, 18)
$lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(220,230,255)
$lblTenant.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$header.Controls.Add($lblTenant)

# Options Panel
$optionsPanel = New-Object System.Windows.Forms.Panel
$optionsPanel.Location = New-Object System.Drawing.Point(20, 105)
$optionsPanel.Size = New-Object System.Drawing.Size(1390, 82)
$optionsPanel.BackColor = $ColorPanel
$optionsPanel.BorderStyle = "FixedSingle"
$form.Controls.Add($optionsPanel)

$btnLogout = New-Object System.Windows.Forms.Button
$btnLogout.Text = "Logout"
$btnLogout.Location = New-Object System.Drawing.Point(1220, 52)
$btnLogout.Size = New-Object System.Drawing.Size(140, 28)
$btnLogout.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 80)
$btnLogout.ForeColor = [System.Drawing.Color]::White
$btnLogout.FlatStyle = "Flat"
$btnLogout.FlatAppearance.BorderSize = 0

$btnLogout.Add_Click({
    Disconnect-GraphSafe
$lblTenant.Text = "Tenant: nicht verbunden"
})

$optionsPanel.Controls.Add($btnLogout)

$chkAppRegistrations = New-Object System.Windows.Forms.CheckBox
$chkAppRegistrations.Text = "App Registrations"
$chkAppRegistrations.Location = New-Object System.Drawing.Point(18, 16)
$chkAppRegistrations.Size = New-Object System.Drawing.Size(170, 24)
$chkAppRegistrations.Checked = $true
$chkAppRegistrations.BackColor = $ColorPanel
$optionsPanel.Controls.Add($chkAppRegistrations)

$chkServicePrincipalDefinitions = New-Object System.Windows.Forms.CheckBox
$chkServicePrincipalDefinitions.Text = "Service Principal Definitions"
$chkServicePrincipalDefinitions.Location = New-Object System.Drawing.Point(205, 16)
$chkServicePrincipalDefinitions.Size = New-Object System.Drawing.Size(230, 24)
$chkServicePrincipalDefinitions.Checked = $true
$chkServicePrincipalDefinitions.BackColor = $ColorPanel
$optionsPanel.Controls.Add($chkServicePrincipalDefinitions)

$chkApplicationConsent = New-Object System.Windows.Forms.CheckBox
$chkApplicationConsent.Text = "Application Consent"
$chkApplicationConsent.Location = New-Object System.Drawing.Point(455, 16)
$chkApplicationConsent.Size = New-Object System.Drawing.Size(180, 24)
$chkApplicationConsent.Checked = $true
$chkApplicationConsent.BackColor = $ColorPanel
$optionsPanel.Controls.Add($chkApplicationConsent)

$chkDelegatedConsent = New-Object System.Windows.Forms.CheckBox
$chkDelegatedConsent.Text = "Delegated Consent"
$chkDelegatedConsent.Location = New-Object System.Drawing.Point(655, 16)
$chkDelegatedConsent.Size = New-Object System.Drawing.Size(180, 24)
$chkDelegatedConsent.Checked = $true
$chkDelegatedConsent.BackColor = $ColorPanel
$optionsPanel.Controls.Add($chkDelegatedConsent)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Auswertung starten"
$btnStart.Location = New-Object System.Drawing.Point(1050, 14)
$btnStart.Size = New-Object System.Drawing.Size(155, 34)
$btnStart.BackColor = $ColorPrimary
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = "Flat"
$btnStart.FlatAppearance.BorderSize = 0
$btnStart.Add_Click({ Start-EwsInventory })
$optionsPanel.Controls.Add($btnStart)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "CSV exportieren"
$btnExport.Location = New-Object System.Drawing.Point(1220, 14)
$btnExport.Size = New-Object System.Drawing.Size(140, 34)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(88,88,88)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.FlatStyle = "Flat"
$btnExport.FlatAppearance.BorderSize = 0
$btnExport.Enabled = $false
$btnExport.Add_Click({ Export-EwsInventory })
$optionsPanel.Controls.Add($btnExport)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Filter:"
$lblFilter.Location = New-Object System.Drawing.Point(18, 52)
$lblFilter.Size = New-Object System.Drawing.Size(50, 20)
$lblFilter.ForeColor = $ColorSubText
$optionsPanel.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(70, 49)
$txtFilter.Size = New-Object System.Drawing.Size(430, 24)
$txtFilter.Add_TextChanged({ Refresh-Grid })
$optionsPanel.Controls.Add($txtFilter)

# Summary Cards
$lblTotalValue = $null
$lblAppRegValue = $null
$lblEnterpriseValue = $null
$lblDefinitionValue = $null

$form.Controls.Add((New-Card -Title "Eindeutige Apps gesamt" -X 20  -Y 202 -ValueLabelRef ([ref]$lblTotalValue)))
$form.Controls.Add((New-Card -Title "mit App Registration"   -X 270 -Y 202 -ValueLabelRef ([ref]$lblAppRegValue)))
$form.Controls.Add((New-Card -Title "mit Enterprise App"     -X 520 -Y 202 -ValueLabelRef ([ref]$lblEnterpriseValue)))
$form.Controls.Add((New-Card -Title "mit SP Definition"      -X 770 -Y 202 -ValueLabelRef ([ref]$lblDefinitionValue)))


# Grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 295)
$grid.Size = New-Object System.Drawing.Size(1390, 390)
$grid.Anchor = "Top,Left,Right,Bottom"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "None"
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = "None"
$grid.GridColor = $ColorBorder
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(32,32,32)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(220,235,250)
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
$grid.AlternatingRowsDefaultCellStyle.BackColor = $ColorGridAlt
$form.Controls.Add($grid)

# Log
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 705)
$txtLog.Size = New-Object System.Drawing.Size(1390, 120)
$txtLog.Anchor = "Left,Right,Bottom"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$txtLog.ForeColor = [System.Drawing.Color]::FromArgb(230,230,230)
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtLog)

$form.Add_FormClosing({

    try {
        Write-GuiLog "GUI wird geschlossen - Logout..."

        # Logout durchführen
        Disconnect-GraphSafe
	$lblTenant.Text = "Tenant: nicht verbunden"
    }
    catch {
        # Falls GUI schon disposed ist → kein Logging mehr möglich
    }

})

[void]$form.ShowDialog()

