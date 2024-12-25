<#
.SYNOPSIS
  SharePoint Online Data Crawler
.DESCRIPTION
  Script uses PnP.PowerShell and a simple configuration xml to crawl a Microsoft 365 tenant for SharePoint sites returning SiteCollection Admins, Owners, Document Inventories, Recursive Unique Permissions, Inherited Permissions, and SubSites. AzureAD App-Only
  access is required for the script to function.
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  <Inputs if any, otherwise state None>
.OUTPUTS
  Output files stored in C:\temp\SharePointCrawler
.NOTES
  Version:        1.0
  Author:         David Collopy
  Creation Date:  August 15, 2021
  Purpose/Change: Initial script development
.EXAMPLE
  From a PowerShell prompt, navigate to the same directory as this script and run it using .\SharePointCrawler.ps1.
  This script does not require administrative privileges.
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[System.Xml.XmlDocument]$Xml = Get-Content -Path .\config.xml
$ConnectionXml = $Xml.crawler.connection
$SitesXml = $Xml.crawler.sites
$WebsXml = $Xml.crawler.webs
$ListsXml = $Xml.crawler.lists
#-----------------------------------------------------------[Functions]------------------------------------------------------------
function EstablishConnection ([System.String]$Url = "$($ConnectionXml.AdminUrl)") {
    Connect-PnPOnline -Url $Url -ReturnConnection -ClientId $ConnectionXml.ClientId -Thumbprint $ConnectionXml.Thumbprint -Tenant $ConnectionXml.Tenant -TenantAdminUrl $ConnectionXml.AdminUrl
    Write-Host -ForegroundColor Green "`t Connection established to" $Url
}
function CrawlSharePoint {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeLine = $true)]
        [System.Object[]]$Site
    )

    begin {

    }

    process {
        EstablishConnection "$($Site.Url)" | Out-Null
        $Web = Get-PnPWeb -Includes $WebsXml.properties.property
        
        
        $PermissionsOutput = "$($Web.ServerRelativeUrl.Replace('/','.').Replace(' ','').Replace('\','-')).csv"
        [System.IO.Directory]::CreateDirectory("$PermissionsDirectory") | Out-Null

        if ($Web.SiteUsers.Count -ne 0) {
            Write-Host -f Yellow "Processing SiteUsers..."
            $SiteUsersOutput = Join-Path "$OutputDirectory\SiteUsers" "$($Web.Url.Replace('https://','').Replace('/','.').Replace(' ','.')).csv"
            Write-Host -f Yellow "Output set to:" $SiteUsersOutput
            
            $SiteUsers = $Web.SiteUsers
            $UserData = @()
            foreach ($User in $SiteUsers) {
                $UserData += New-Object PSObject -Property ([ordered] @{
                        Title             = $User.Title
                        LoginName         = $User.LoginName
                        UserPrincipalName = $User.UserPrincipalName
                        IsSiteAdmin       = $IsSiteAdmin
                    })
            }
            $UserData | Export-Csv -Path $SiteUsersOutput -NoTypeInformation -Append
        }
        
        $Lists = Get-PnPList -Includes $ListsXml.properties.property | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and $_.Title -notin $ListsXml.exclusions.title -and $_.ItemCount -gt 0 }
        if ($Lists.Count -ne 0) {
            $InventoryOutput = "$($Web.ServerRelativeUrl.Replace('/','.').Replace(' ','').Replace('\','-')).csv"
            [System.IO.Directory]::CreateDirectory("$InventoryDirectory") | Out-Null
            foreach ($List in $Lists) {
                CrawlInventory -List $List -InventoryDirectory $InventoryDirectory
            }
        } else { $InventoryDirectory = "" }

        $SiteOutput = Join-Path $Directory "HMH.SharePoint.Index.csv"
        Write-Host -f Yellow "Output set to:" $SiteOutput

        $WebData = @()
        $WebData += New-Object PSObject -Property ([ordered] @{
                Title         = $Web.Title
                RootWeb       = $Web.RootWeb
                TotalSubSites = $Web.Webs.Count
                SubSitesData  = $SubSitesDirectory
                Permissions   = $PermissionsDirectory
                Inventory     = $InventoryDirectory
                SiteUsers     = $SiteUsersOutput
            })
        $WebData | Export-Csv -Path "$Directory\SiteIndex.csv" -NoTypeInformation -Append
        
        Write-Host -f Yellow "CrawlingPermissions"
        CrawlPermissions -Web $Web -PermissionsDirectory $PermissionsDirectory

        if ($Web.Webs.Count -ne 0) {
            Write-Host -f Yellow "Processing SubSites..."
            $SubWebs = Get-PnPSubWeb -Recurse | Where-Object -Property Template -NotIn $SitesXml.exclusions.template
            $SubSitesDirectory = Join-Path "$OutputDirectory\SubSites" "$($Web.Url.Replace('https://','').Replace('/','.').Replace(' ','.'))"
            [System.IO.Directory]::CreateDirectory("$SubSitesDirectory") | Out-Null
            $SubWebs | CrawlSharePoint
        } else { Continue }
    }
}
function CrawlSubSites {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeLine = $true)]
        [System.Object[]]$Site,
        [Parameter(Mandatory = $true)]
        [System.String]$SubSitesDirectory,
        [Parameter(Mandatory = $true)]
        [System.String]$PermissionsDirectory,
        [Parameter(Mandatory = $true)]
        [System.String]$InventoryDirectory
    )

    process {
        EstablishConnection "$($Site.Url)" | Out-Null
        $Web = Get-PnPWeb -Includes $WebsXml.properties.property
        
        if ($Web.SiteUsers.Count -ne 0) {
            Write-Host -f Yellow "Processing SiteUsers..."
            $SiteUsersOutput = Join-Path "$Directory\SiteUsers" "$($Web.Url.Replace('https://','').Replace('/','.').Replace(' ','.')).csv"
            Write-Host -f Yellow "Output set to:" $SiteUsersOutput
            
            $SiteUsers = $Web.SiteUsers
            foreach ($User in $SiteUsers) {
                $UserData = New-Object psobject
                $UserData | Add-Member NoteProperty Title($User.Title)
                $UserData | Add-Member NoteProperty LoginName($User.LoginName)
                $UserData | Add-Member NoteProperty UserPrincipalName($User.UserPrincipalName)
                $UserData | Add-Member NoteProperty IsSiteAdmin($User.IsSiteAdmin)
                $UserData | Export-Csv -Path $SiteUersOutput -NoTypeInformation -Append
            }
        }
        
        $Lists = Get-PnPList -Includes $ListsXml.properties.property | Where-Object { $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $false -and $_.Title -notin $ListsXml.exclusions.title -and $_.ItemCount -gt 0 }
        if ($Lists.Count -ne 0) {
            [System.IO.Directory]::CreateDirectory("$InventoryDirectory") | Out-Null
            foreach ($List in $Lists) {
                CrawlInventory -List $List -InventoryDirectory $InventoryDirectory
            }
        } else { $InventoryDirectory = "" }

        if ($Web.Webs.Count -ne 0) {
            $SubSiteCount = $Web.Webs.Count
            $SubSitesDirectory = Join-Path "$Directory\SubSites" "$($Web.Url.Replace('https://','').Replace('/','.').Replace(' ','.'))"
            [System.IO.Directory]::CreateDirectory("$SubSitesDirectory") | Out-Null
            Get-PnPSubWeb -Recurse | CrawlSubSites -SubSitesDirectory $SubSitesDirectory -PermissionsDirectory $PermissionsDirectory -InventoryDirectory $InventoryDirectory
        } else {
            $SubSitesDirectory = ""
            $SubSiteCount = 0 
        }

        $Output = Join-Path $SubSitesDirectory "$($Web.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-SubSite.csv"
        Write-Host -f Yellow "Output set to:" $Output

        $Index = New-Object PSObject
        $Index | Add-Member NoteProperty Title($Web.Title)
        $Index | Add-Member NoteProperty Url($Web.Url)
        $Index | Add-Member NoteProperty SubSites($SubSiteCount)
        $Index | Add-Member NoteProperty Webs("$SubSitesDirectory")
        $Index | Add-Member NoteProperty HasUniquePermissions("$PermissionsDirectory")
        $Index | Add-Member NoteProperty Inventory("$InventoryDirectory")
        $Index | Add-Member NoteProperty SiteUsers("$SiteUsersOutput")
        $Index | Export-Csv -Path $Output -NoTypeInformation -Append

        Write-Host -f Yellow "CrawlingPermissions"
        CrawlPermissions -Web $Web -PermissionsDirectory $PermissionsDirectory
    }
}
function CrawlPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.SecurableObject]$Web,
        [Parameter(Mandatory = $true)]
        [System.String]$PermissionsDirectory,
        [Parameter(Mandatory = $false)]
        [switch]$Recursive,
        [Parameter(Mandatory = $false)]
        [switch]$ScanItemLevel,
        [Parameter(Mandatory = $false)]
        [switch]$IncludeInherited
    )

    try {
        $Web = Get-PnPWeb -Includes $WebsXml.properties.property
        $Output = Join-Path $PermissionsDirectory "$($Web.Url.Replace('https://','').Replace('/','.').Replace(' ','.'))-SiteCollectionAdmins.csv"
        Write-Host -f Yellow "Output set to:" $Output

        Write-Host -f Yellow "Getting Site Collection Administrators..."
        $SiteAdmins = Get-PnPSiteCollectionAdmin

        $SiteCollectionAdmins = ($SiteAdmins | Select-Object -ExpandProperty Title) -join " | "
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.Url)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
        $Permissions | Export-Csv $Output -NoTypeInformation
        function CrawlItemPermissions([Microsoft.SharePoint.Client.List]$List) {
            Write-Host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title
            $ListItems = Get-PnPListItem -List $List -PageSize 2000
            $ItemCounter = 0
            foreach ($ListItem in $ListItems) {
                if ($IncludeInherited) {
                    $Output = Join-Path $PermissionsDirectory "$($List.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-Items-InheritedPermissions.csv"
                    Write-Host -f Yellow "Output set to:" $Output
                    Get-Permissions -Object $ListItem -Output $Output
                } else {
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
                    if ($HasUniquePermissions -eq $true) {
                        
                        $Output = Join-Path $PermissionsDirectory "$($List.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-Items-HasUniquePermissions.csv"
                        Write-Host -f Yellow "Output set to:" $Output
                        Get-Permissions -Object $ListItem -Output $Output
                    }
                }
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
            }
        }
        function CrawlListPermissions([Microsoft.SharePoint.Client.Web]$Web) {
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
            $ListsCounter = 0
            foreach ($List in $Lists) {
                if ($List.Hidden -eq $false -and $ListsXml.exclusions.title -notcontains $List.Title) {
                    $ListsCounter++
                    Write-Progress -PercentComplete ($ListsCounter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.Url)" -Status "Processing Lists $ListsCounter of $($Lists.Count)"

                    if ($ScanItemLevel) {
                        CrawlItemPermissions -List $List
                    }
                    if ($IncludeInherited) {
                        $Output = Join-Path $PermissionsDirectory "$($List.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-InheritedPermissions.csv"
                        Write-Host -f Yellow "Output set to:" $Output
                        Get-Permissions -Object $List -Output $Output
                    } else {
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        if ($HasUniquePermissions -eq $true) {
                            $Output = Join-Path $PermissionsDirectory "$($List.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-HasUniquePermissions.csv"
                            Write-Host -f Yellow "Output set to:" $Output
                            Get-Permissions -Object $List -Output $Output
                        }
                    }
                }
            }
        }
        function CrawlWebPermissions([Microsoft.SharePoint.Client.Web]$Web) {
            Write-Host -f Yellow "Getting Permissions of the Web: $($Web.Url)..."
            $Output = Join-Path $PermissionsDirectory "$($Web.Url.Replace('https://','').Replace('/','.'))-WebPermissions.csv"
            Write-Host -f Yellow "Output set to:" $Output
            Get-Permissions -Object $Web -Output $Output
            Write-Host -f Yellow "`t Getting Permissions of Lists and Libraries..."
            CrawlListPermissions($Web)
            
            if ($Recursive) {
                $SubWebs = Get-PnPProperty -ClientObject $Web -Property Webs
                foreach ($SubWeb in $SubWebs) {
                    if ($IncludeInherited) {
                        CrawlWebPermissions($SubWeb)
                    } else {
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
                        if ($HasUniquePermissions -eq $true) {
                            CrawlWebPermissions($SubWeb)
                        }
                    }
                }
            }
        }

        CrawlWebPermissions $Web
        Write-Host -f Green "`n*** Site Permission Report Generated Successfully!***"
    } catch {
        Write-Host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
    }
}
function Get-Permissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.SecurableObject]$Object,
        [Parameter(Mandatory = $true)]
        [System.String]$Output
    )

    switch ($Object.TypedObject.ToString()) {
        "Microsoft.SharePoint.Client.Web" { $ObjectType = "Site" ; $ObjectUrl = $Object.Url; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem" {
            if ($Object.FileSystemObjectType -eq "Folder") {
                $ObjectType = "Folder"
                $Folder = Get-PnPProperty -ClientObject $Object -Property Folder
                $ObjectTitle = $Object.Folder.Name
                $ObjectUrl = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.Folder.ServerRelativeUrl)
            } else {
                Get-PnPProperty -ClientObject $Object -Property File, ParentList
                if ($null -ne $Object.File.Name) {
                    $ObjectType = "File"
                    $ObjectTitle = $Object.File.Name
                    $ObjectUrl = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $Object.File.ServerRelativeUrl)
                } else {
                    $ObjectType = "List Item"
                    $ObjectTitle = $Object["Title"]
                    $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl
                    $ObjectUrl = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $DefaultDisplayFormUrl, $Object.Id)
                }
            }
        }
        default {
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder
            $ObjectUrl = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl, ''), $RootFolder.ServerRelativeUrl)
        }
    }

    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments

    $PermissionCollection = @()
    foreach ($RoleAssignment in $Object.RoleAssignments) {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

        $PermissionType = $RoleAssignment.Member.PrincipalType
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select-Object -ExpandProperty Name
        $PermissionLevels = ($PermissionLevels | Where-Object { $_ -ne "Limited Access" }) -join " | "
        
        if ($PermissionLevels.Length -eq 0) { Continue }

        if ($PermissionType -eq "SharePointGroup") {
            $GroupMembers = Get-PnPGroupMember -Group $RoleAssignment.Member.LoginName

            if ($GroupMembers.Count -eq 0) { Continue }
            $GroupUsers = ($GroupMembers | Select-Object -ExpandProperty Email) -join ";"

            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectUrl)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($GroupUsers)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("SharePoint Group: $($RoleAssignment.Member.LoginName)")
            $PermissionCollection += $Permissions
        } else {
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectUrl)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Email)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    $PermissionCollection | Export-Csv $Output -NoTypeInformation -Append
}
function CrawlInventory {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.SharePoint.Client.SecurableObject]$List,
        [Parameter(Mandatory = $true)]
        [System.String]$InventoryDirectory
    )

    $Output = Join-Path $InventoryDirectory "$($List.Title.Replace('\','.').Replace('/','.').Replace(' ',''))-Inventory.csv"
    Write-Host -f Yellow "Output set to:" $Output
    $ItemsCounter = 0
    $ListItems = Get-PnPListItem -List $List -PageSize 1000 -Fields Author, Created -ScriptBlock {
        Param($Items)
        $ItemsCounter += $Items.Count;
        Write-Progress -PercentComplete ($ItemsCounter / ($List.ItemCount) * 100) -Activity "Getting Inventory from: $($List.Title)" -Status "Processing Items $ItemsCounter to $($List.ItemCount)";
    }
    
    Write-Progress -Activity "Completed Retrieving Inventory from Library $($List.Title)" -Completed
    Write-Host -f Green "`t Completed Retrieving Inventory from Library: $($List.Title)"
    Write-Host -f Green "`t Total Items:" $List.ItemCount
    $Inventory = @()

    foreach ($ListItem in $ListItems) {
        $Inventory += New-Object PSObject -Property ([ordered] @{
                LibraryName     = $List.Title
                Type            = $ListItem.FileSystemObjectType
                ItemRelativeURL = $ListItem.FieldValues.FileRef
                CreatedBy       = $ListItem.FieldValues.Author.Email
                CreatedAt       = $ListItem.FieldValues.Created
                ModifiedBy      = $ListItem.FieldValues.Editor.Email
                ModifiedAt      = $ListItem.FieldValues.Modified
            })
    }
    $Inventory | Export-Csv -Path $Output -NoTypeInformation -Append
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
#Log-Start -LogPath $sLogPath -LogName $sLogName -ScriptVersion $sScriptVersion
Write-Host -ForegroundColor Yellow "Creating output directory..."
$OutputDirectory = Split-Path $MyInvocation.MyCommand.Path -Parent | Join-Path -ChildPath $ConnectionXml.Tenant.Replace('.onmicrosoft.com', '')
$PermissionsDirectory = Join-Path $OutputDirectory "Permissions"
$InventoryDirectory = Join-Path $OutputDirectory "Inventory"
$SubSitesDirectory = Join-Path $OutputDirectory "SubSites"

[System.IO.Directory]::CreateDirectory("$PermissionsDirectory") | Out-Null
[System.IO.Directory]::CreateDirectory("$InventoryDirectory") | Out-Null

EstablishConnection  | Out-Null
Get-PnPTenantSite -Filter "Url -like '$($ConnectionXml.Url)' -and Url -notlike '-my.sharepoint.com/' -and Url -notlike '/portals/'" | Where-Object -Property Template -NotIn $SitesXml.exclusions.template | CrawlSharePoint
#Log-Finish -LogPath $sLogFile
