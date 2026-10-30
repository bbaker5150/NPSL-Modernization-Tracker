# Run in an authenticated PnP.PowerShell session as a site owner.
# Example: ./scripts/setup-feedback.ps1 -SubmitterGroups 'Metrology App User','Tracker Project Engineers','Tracker Managers'
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string[]]$SubmitterGroups,
  [string]$ListPrefix = 'Modernization'
)
$ErrorActionPreference = 'Stop'
$ctx = Get-PnPContext
$web = $ctx.Web
$ctx.Load($web.AssociatedOwnerGroup)
$ctx.Load($web.CurrentUser)
$ctx.ExecuteQuery()
$owners = $web.AssociatedOwnerGroup
if ($owners.ServerObjectIsNull -or $owners.Id -le 0) { throw 'Configure the associated site Owners group first.' }
$ctx.Load($owners.Users)
$ctx.ExecuteQuery()
if (-not $web.CurrentUser.IsSiteAdmin -and $web.CurrentUser.Id -notin $owners.Users.Id) { throw 'Only a site owner can configure feedback.' }
# Resolve every intended group before making changes. Include every group through
# which users access this deployment (including legacy Members, if applicable).
$groups = @($SubmitterGroups | ForEach-Object { Get-PnPGroup -Identity $_ })
$namespace = $ListPrefix -replace '[^A-Za-z0-9]', ''
if (-not $namespace) { $namespace = 'Modernization' }
$appName = if ($namespace -eq 'Modernization') { 'Modernization-Tracker' } else { "$namespace Tracker" }
$title = "$appName - Feedback"
$aliases = @($title, "${namespace}Feedback")
if ($namespace -eq 'Modernization') { $aliases += 'NPSL Tracker - Feedback' }
$list = @(Get-PnPList | Where-Object { $_.Title -in $aliases })
if ($list.Count) { throw "List $title already exists. No changes made. Audit its permissions before reusing it; this script deliberately does not reset existing feedback." }
New-PnPList -Title $title -Template GenericList -OnQuickLaunch:$false | Out-Null
$list = Get-PnPList -Identity $title
# Break inheritance BEFORE adding fields or granting submission access.
$list.BreakRoleInheritance($false, $true)
$list.NoCrawl = $true
$list.EnableAttachments = $false
$list.Update()
$ctx.ExecuteQuery()
Add-PnPField -List $title -InternalName FeedbackKind -DisplayName 'Feedback kind' -Type Choice -Choices 'Bug','Feature' -Required | Out-Null
Add-PnPField -List $title -InternalName FeedbackDetails -DisplayName 'Feedback details' -Type Note -Required | Out-Null
# Exact permission mask. Do not clone Contribute: it includes ViewListItems.
$roleName = "${namespace}Feedback Submit Only"
$roles = $web.RoleDefinitions
$ctx.Load($roles)
$ctx.ExecuteQuery()
if (@($roles | Where-Object Name -eq $roleName).Count) { throw "Role $roleName already exists. Audit it before retrying. The new list remains restricted." }
$mask = New-Object Microsoft.SharePoint.Client.BasePermissions
@('AddListItems','Open','ViewPages','ViewFormPages','UseRemoteAPIs') | ForEach-Object { $mask.Set([Microsoft.SharePoint.Client.PermissionKind]::$_) }
$info = New-Object Microsoft.SharePoint.Client.RoleDefinitionCreationInformation
$info.Name = $roleName
$info.Description = 'Submit feedback without reading, editing, or deleting any submissions.'
$info.BasePermissions = $mask
$submitRole = $roles.Add($info)
$fullControl = $roles.GetByType([Microsoft.SharePoint.Client.RoleType]::Administrator)
$ctx.ExecuteQuery()
$binding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
$binding.Add($fullControl)
$list.RoleAssignments.Add($owners, $binding) | Out-Null
$ctx.ExecuteQuery()
# Remove the direct creator grant automatically added when inheritance is broken.
$ctx.Load($list.RoleAssignments)
$ctx.ExecuteQuery()
foreach ($assignment in @($list.RoleAssignments)) {
  $ctx.Load($assignment.Member)
  $ctx.ExecuteQuery()
  if ($assignment.Member.Id -ne $owners.Id) { $assignment.DeleteObject(); $ctx.ExecuteQuery() }
}
foreach ($group in $groups) {
  if ($group.Id -eq $owners.Id) { continue }
  $binding = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
  $binding.Add($submitRole)
  $list.RoleAssignments.Add($group, $binding) | Out-Null
}
$ctx.ExecuteQuery()
Write-Host "Configured $title. Verify POST succeeds and GET is denied with a non-owner account before releasing the feature."
