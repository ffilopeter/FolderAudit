Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $adAvailable = $true
} catch { $adAvailable = $false }

$global:PermissionModel = @()

# =======================
# GUI
# =======================

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Tree with Permissions Viewer"
$form.Size = New-Object System.Drawing.Size(950, 500)
$form.StartPosition = "CenterScreen"

# TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(400, 400)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.Anchor = 'Top, Bottom, Left'

# ListView
$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Size = New-Object System.Drawing.Size(510, 400)
$listView.Location = New-Object System.Drawing.Point(420, 50)
$listView.Anchor = 'Top, Bottom, Left, Right'
$listView.Font = New-Object System.Drawing.Font("Segoe UI", 9)

[void]$listView.Columns.Add("Identity", 200)
[void]$listView.Columns.Add("Rights", 200)
[void]$listView.Columns.Add("Access", 100)
[void]$listView.Columns.Add("Inherited", 100)

# Browse Button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse Folder"
$browseButton.Size = New-Object System.Drawing.Size(120, 30)
$browseButton.Location = New-Object System.Drawing.Point(10, 10)
$browseButton.Anchor = 'Top, Left'

# Loading label
$loadingLabel = New-Object System.Windows.Forms.Label
$loadingLabel.Text = "Loading..."
$loadingLabel.Size = New-Object System.Drawing.Size(100, 20)
$loadingLabel.Location = New-Object System.Drawing.Point(140, 20)
$loadingLabel.Visible = $false

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# =======================
# Event handlers
# =======================

# Click to Browse button
# load folder structure to tree view
$browseButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq "OK") {

        # Clear treeView and listView
        $treeView.Nodes.Clear()
        $listView.Items.Clear()

        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = [System.IO.Path]::GetFileName($folderBrowser.SelectedPath)
        $rootNode.Tag = $folderBrowser.SelectedPath
        $rootNode.ForeColor = [System.Drawing.Color]::Blue

        $treeView.Nodes.Add($rootNode)

        # $loadingLabel.Visible = $true
        # $form.Refresh()

        # Load the whole tree
        Load-TreeView -parentNode $rootNode -path $folderBrowser.SelectedPath
        
        $rootNode.Expand()

        Load-Permissions -path $folderBrowser.SelectedPath
        Rebuild-ListView $listView

        # $loadingLabel.Visible = $false
        # $form.Refresh()
    }
})

$treeView.Add_AfterSelect({
    $selectedPath = $treeView.SelectedNode.Tag
    if (Test-Path $selectedPath) {
        Write-Host "Path exists"
        Load-Permissions -path $selectedPath
        Rebuild-ListView $listView
    } else {
        Write-Host "Path doesnt exist"
    }
})

# Expand / Collapse group
$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -eq 1) {
        $selected = $listView.SelectedItems[0]
        $tag = $selected.Tag

        if ($tag.Type -eq "User") {
    
        } elseif ($tag.Type -eq "Group") {
            Toggle-ExpandProperty $tag.Id $global:PermissionModel
            Rebuild-ListView $listView
        }
    }
})

# =======================
# Functions
# =======================

# Simplify rights
function Format-Rights {
    param ([System.Security.AccessControl.FileSystemRights]$rights)
    $rights = $rights -band (-bnot [System.Security.AccessControl.FileSystemRights]::Synchronize)
    $friendly = @()
    if ($rights -band [System.Security.AccessControl.FileSystemRights]::FullControl) {
        $friendly += "Full Control"
    } else {
        if ($rights -band [System.Security.AccessControl.FileSystemRights]::Modify) { $friendly += "Modify" }
        if ($rights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) {
            $friendly += "Read & Execute"
        } elseif ($rights -band [System.Security.AccessControl.FileSystemRights]::Read) {
            $friendly += "Read"
        }
        if ($rights -band [System.Security.AccessControl.FileSystemRights]::Write) {
            $friendly += "Write"
        }
        if ($friendly.Count -eq 0) { $friendly += $rights.ToString() }
    }
    return ($friendly -join ", ")
}

# Load TreeView recursively
function Load-TreeView {
    param([System.Windows.Forms.TreeNode]$parentNode, [string]$path)

    try {
        
        # Get parentNode subfolders
        Get-ChildItem -Path $path -Directory -ErrorAction Stop | ForEach-Object {

            # Add to treeView (parentNode)
            $dirNode = New-Object System.Windows.Forms.TreeNode
            $dirNode.Text = $_.Name
            $dirNode.Tag = $_.FullName
            $dirNode.ForeColor = [System.Drawing.Color]::Blue
            $parentNode.Nodes.Add($dirNode)

            # Repeat
            Load-TreeView -parentNode $dirNode -path $_.FullName
        }

        # Get parentNode files
        Get-ChildItem -Path $path -File -ErrorAction Stop | ForEach-Object {

            # Add to treeView (parentNode)
            $fileNode = New-Object System.Windows.Forms.TreeNode
            $fileNode.Text = $_.Name
            $fileNode.Tag = $_.FullName
            $fileNode.ForeColor = [System.Drawing.Color]::Black
            $parentNode.Nodes.Add($fileNode)
        }
    } catch {}
}

# Check if identity is a group or user
function Is-Group {
    param([string]$identity)

    $result = [PSCustomObject]@{
        Type    = "Undefined"
        Group   = $false
    }

    if ($identity -Match '\\') {
        $domain, $name = $identity.Split('\', 2)

        if ($domain -Match '^(BUILTIN|NT AUTHORITY)$' -Or $domain -eq $env:COMPUTERNAME) {

            # Identity is local
            try {
                $group = [ADSI]"WinNT://./$name,group"
                if ($group.Path) {
                    $result.Type = "Local"
                    $result.Group = $true
                }
            } catch {}
        } else {

            # Identity is domain
            if ($adAvailable) {
                try {
                    $adObject = Get-ADObject -Filter { SamAccountName -eq $name } -ErrorAction Stop
                    if ($adObject.ObjectClass -eq 'group') {
                        $result.Type = "Domain"
                        $result.Group = $true
                    }
                } catch {}
            }
        }
    }

    return $result
}

function Get-GroupMembers {
    param([string]$identity)

    $members = @()
    $type = (Is-Group $identity).Type
    $domain, $name = $identity.Split('\', 2)

    if ($type -eq 'Domain') {

        # Get domain group members
        if ($adAvailable) {
            try {
                Get-ADGroupMember -Identity $name -ErrorAction Stop | ForEach-Object {
                    $sam = $_.SamAccountName
                    $name = $_.Name
                    $members += "$domain\$sam"
                }
            } catch {}
        }

    } elseif ($type -eq 'Local') {

        # Get local group members
        try {
            $output = net localgroup "$name" 2>&1
            $start = ($output | Select-String -SimpleMatch '---').LineNumber
            $end = ($output | Select-String -SimpleMatch 'The command completed successfully.').LineNumber
            $result = $output[$start..($end - 2)] | Where-Object { $_.Trim() -ne '' }

            foreach ($member in $result) {
                if ($member -Match '\\') {
                    $members += $member
                } else {
                    $members += $env:COMPUTERNAME +"\$member"
                }
            }
        } catch {}
    } else {}

    return $members
}

function Resolve-Group {
    param([string]$identity, [int]$indent)
    $indentation = $indent + 1
    $members = Get-GroupMembers $identity

    $result = @()
    $members | ForEach-Object {
        $id = $_
        $isGroup = (Is-Group $_).Group
        $name = $id.Split('\')[-1]

        $result += [PSCustomObject]@{
            Id              = [guid]::NewGuid().ToString()
            Identity        = $id
            DisplayName     = $(if ($isGroup) { $name } else { $id })
            Type            = $(if ($isGroup) { "Group" } else { "User" })
            Indent          = $indentation
            Expanded        = $false

            Members         = $(if ($isGroup) { Resolve-Group $id $indentation } else { @() })

            Rights          = ""
            AccessType      = ""
            Inherited       = ""
        }
    }

    return $result
}

function Load-Permissions {
    param([string]$path)

    $global:PermissionModel = @()
    $listView.Items.Clear()

    try {
        $acl = Get-Acl -Path $path
        foreach ($access in $acl.Access) {
            
            # name with domain portion prepended
            $identity = $access.IdentityReference.ToString()

            $inherited = $access.IsInherited
            
            $isGroup = (Is-Group $identity).Group

            # identity name (without domain portion)
            $name = $identity.Split('\')[-1]

            # add to PermissionModel array
            $global:PermissionModel += [PSCustomObject]@{
                Id              = [guid]::NewGuid().ToString()
                Identity        = $identity
                DisplayName     = $(if ($isGroup) { $name } else { $identity })
                Type            = $(if ($isGroup) { "Group" } else { "User" })
                Indent          = 0
                Expanded        = $false

                # If identity is group, resolve it's members
                Members         = $(if ($isGroup) { Resolve-Group $identity 0 } else { @() })

                Rights          = Format-Rights $access.FileSystemRights
                AccessType      = $access.AccessControlType.ToString()
                Inherited       = $(if ($inherited) { "Inherited" } else { "Not inherited" })
            }
        }
    } catch {
        $global:PermissionModel += [PSCustomObject]@{
            Id              = 0
            Identity        = "Access Denied"
            DisplayName     = "Access Denied"
            Type            = "Error"
            Indent          = 0
            Expanded        = $false
            Members         = @()
            Rights          = ""
            AccessType      = ""
        }
    }
}

function Rebuild-ListView {
    param($listView)

    $listView.Items.Clear()

    function Add-Visible($obj) {
        $groupPrefix    = ""
        $indentSize     = 4

        if ($obj.Type -eq "Group") {
            $groupPrefix = $(if ($obj.Expanded) { "- " } else { "+ " })
        }

        $prefix = " " * $indentSize * $obj.Indent
        $item = New-Object System.Windows.Forms.ListViewItem($prefix + $groupPrefix + $obj.DisplayName)

        [void]$item.SubItems.Add($obj.Rights)
        [void]$item.SubItems.Add($obj.AccessType)
        [void]$item.SubItems.Add($obj.Inherited)

        $item.Tag = @{ Id = $obj.Id; Type = $obj.Type }
        $item.ForeColor = $(if ($obj.Type -eq "Group") { [System.Drawing.Color]::Blue } else { [System.Drawing.Color]::Black })

        $listView.Items.Add($item)

        # Expand group
        if ($obj.Type -eq "Group" -and $obj.Expanded) {
            $obj.Members | ForEach-Object {
                Add-Visible $_
            }
        }
    }

    $global:PermissionModel | ForEach-Object {
        Add-Visible $_
    }
}

# Find identity by GUID and toggle Expand property
# search recursively
function Toggle-ExpandProperty {
    param($guid, $permissionModel)

    foreach ($identity in $permissionModel) {
        if ($identity.Id -eq $guid) {
            $identity.Expanded = -not $identity.Expanded
            return $true
        }

        if ($identity.Members) {
            $toggle = Toggle-ExpandProperty $guid $identity.Members
            if ($toggle) { return $true }
        }
    }

    return $false
}

# =======================
# Add controls and render
# =======================

$form.Controls.Add($treeView)
$form.Controls.Add($listView)
$form.Controls.Add($browseButton)
$form.Controls.Add($loadingLabel)
[void]$form.ShowDialog()
