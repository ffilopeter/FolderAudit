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

# Browse Button
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse Folder"
$browseButton.Size = New-Object System.Drawing.Size(120, 30)
$browseButton.Location = New-Object System.Drawing.Point(10, 10)
$browseButton.Anchor = 'Top, Left'

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

        # Load the whole tree
        Load-TreeView -parentNode $rootNode -path $folderBrowser.SelectedPath
        
        $rootNode.Expand()

        Load-Permissions -path $folderBrowser.SelectedPath
        Rebuild-ListView $listView
    }
})

$treeView.Add_AfterSelect({
    $selectedPath = $treeView.SelectedNode.Tag
    if (Test-Path $selectedPath) {
        Load-Permissions -path $selectedPath
        Rebuild-ListView $listView
    }
})

# =======================
# Functions
# =======================

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
                    $members += "$domain\$sam"
                }
            } catch {}
        }

    } elseif ($type -eq 'Local') {

        # Get local group members
        try {
            $output = net localgroup "$name" 2>&1
            $start = ($output | Select-String -SimpleMatch '---').LineNumber + 1
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
            Id              = $result.Count
            Identity        = $id
            DisplayName     = $(if ($isGroup) { "+ $name" } else { $id })
            Type            = $(if ($isGroup) { "Group" } else { "User" })
            Indent          = $indentation
            Expanded        = $false

            Members         = $(if ($isGroup) { Resolve-Group $id $indentation } else { @() })

            Rights          = ""
            AccessType      = ""
        }
    }

    return $result
}

function Load-Permissions {
    param([string]$path)

    $global:PermissionModel = @()

    try {
        $acl = Get-Acl -Path $path
        foreach ($access in $acl.Access) {
            
            # name with domain portion prepended
            $identity = $access.IdentityReference.ToString()
            
            $isGroup = (Is-Group $identity).Group

            # identity name (without domain portion)
            $name = $identity.Split('\')[-1]

            # add to PermissionModel array
            $global:PermissionModel += [PSCustomObject]@{
                Id              = $global:PermissionModel.Count
                Identity        = $identity
                DisplayName     = $(if ($isGroup) { "+ $name" } else { $identity })
                Type            = $(if ($isGroup) { "Group" } else { "User" })
                Indent          = 0
                Expanded        = $false

                # If identity is group, resolve it's members
                Members         = $(if ($isGroup) { Resolve-Group $identity 0 } else { @() })

                Rights          = $access.FileSystemRights.ToString()
                AccessType      = $access.AccessControlType.ToString()
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
    $itemTag = 0

    function Add-Visible($obj) {
        $id             = $obj.Id
        $identity       = $obj.Identity
        $displayName    = $obj.DisplayName
        $type           = $obj.Type
        $indent         = $obj.Indent
        $expanded       = $obj.Expanded
        $members        = $obj.Members
        $rights         = $obj.Rights
        $accessType     = $obj.AccessType

        $prefix = "    " * $indent
        $item = New-Object System.Windows.Forms.ListViewItem($prefix + $displayName)
        [void]$item.SubItems.Add($rights)
        [void]$item.SubItems.Add($accessType)

        if ($type -eq "User") {
            $item.ForeColor = [System.Drawing.Color]::Black
        } elseif ($type -eq "Group") {
            $item.ForeColor = [System.Drawing.Color]::Blue
        }

        $listView.Items.Add($item)

        if ($expanded) {
            $members | ForEach-Object {
                Add-Visible $_
            }
        }
    }

    $global:PermissionModel | ForEach-Object {
        Add-Visible $_
    }
}

# =======================
# Add controls and render
# =======================

$form.Controls.Add($treeView)
$form.Controls.Add($listView)
$form.Controls.Add($browseButton)
[void]$form.ShowDialog()
