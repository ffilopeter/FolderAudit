Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $adAvailable = $true
} catch {
    $adAvailable = $false
}

$Global:PermissionModel = @()

function Format-Rights {
    param ([System.Security.AccessControl.FileSystemRights]$rights)
    $rights = $rights -band (-bnot [System.Security.AccessControl.FileSystemRights]::Synchronize)
    $friendly = @()
    if ($rights -band [System.Security.AccessControl.FileSystemRights]::FullControl) { $friendly += "Full Control" }
    if ($rights -band [System.Security.AccessControl.FileSystemRights]::Modify) { $friendly += "Modify" }
    if ($rights -band [System.Security.AccessControl.FileSystemRights]::ReadAndExecute) { $friendly += "Read & Execute" }
    elseif ($rights -band [System.Security.AccessControl.FileSystemRights]::Read) { $friendly += "Read" }
    if ($rights -band [System.Security.AccessControl.FileSystemRights]::Write) { $friendly += "Write" }
    if ($friendly.Count -eq 0) { $friendly += $rights.ToString() }
    return ($friendly -join ", ")
}

function Is-Group {
    param ([string]$name)
    if ($name -match '\\') {
        $sam = $name.Split('\')[-1]
        if ($adAvailable) {
            try { Get-ADGroup -Identity $sam -ErrorAction Stop | Out-Null; return $true } catch {}
        }
    } else {
        try {
            $group = [ADSI]"WinNT://./$name,group"
            if ($group.Path) { return $true }
        } catch {}
    }
    return $false
}

function Get-LocalGroupMembers {
    param ([string]$groupName)
    try {
        $output = net localgroup "$groupName" 2>&1
        $start = ($output | Select-String -SimpleMatch '---').LineNumber + 1
        $end = ($output | Select-String -SimpleMatch 'The command completed successfully.').LineNumber
        return $output[$start..($end - 2)] | Where-Object { $_.Trim() -ne '' }
    } catch { return @() }
}

function Load-Permissions {
    param ($path)
    $Global:PermissionModel = @()
    try {
        $acl = Get-Acl -Path $path
        foreach ($access in $acl.Access) {
            $id = $access.IdentityReference.ToString()
            $isGroup = Is-Group $id
            $name = $id.Split('\')[-1]
            $Global:PermissionModel += [PSCustomObject]@{
                Identity = $id
                Display  = $(if ($isGroup) { "+ $name" } else { $id })
                Type     = $(if ($isGroup) { "Group" } else { "User" })
                Parent   = ""
                Indent   = 0
                Expanded = $false
                Rights   = Format-Rights $access.FileSystemRights
                Access   = $access.AccessControlType.ToString()
            }
        }
    } catch {
        $Global:PermissionModel += [PSCustomObject]@{
            Identity = "Access Denied"
            Display  = "Access Denied"
            Type     = "Error"
            Parent   = ""
            Indent   = 0
            Expanded = $false
            Rights   = ""
            Access   = ""
        }
    }
}

function Expand-Group {
    param ($item)

    $name = $item.Identity.Split('\')[-1]
    $indent = $item.Indent + 1
    $parentId = $item.Identity

    $existingIds = @{}

    function Normalize($x) {
        if ($x -match '\\') { return $x.Trim().ToLower() }
        return "$env:USERDOMAIN\$x".ToLower()
    }

    $members = Get-LocalGroupMembers $name
    foreach ($member in $members) {
        $id = $member.Trim()
        $nid = Normalize $id
        if ($existingIds.ContainsKey($nid)) { continue }
        $existingIds[$nid] = $true

        $isGroup = Is-Group $id
        $short = $id.Split('\')[-1]
        $Global:PermissionModel += [PSCustomObject]@{
            Identity = $id
            Display  = $(if ($isGroup) { (' ' * ($indent * 4 - 2)) + "+ $short" } else { (' ' * ($indent * 4)) + $id })
            Type     = $(if ($isGroup) { "Group" } else { "User" })
            Parent   = $parentId
            Indent   = $indent
            Expanded = $false
            Rights   = "(member of $name)"
            Access   = ""
        }
    }

    if ($adAvailable) {
        try {
            Get-ADGroupMember -Identity $name -ErrorAction Stop | ForEach-Object {
                $id = $_.SamAccountName
                $nid = Normalize $id
                if ($existingIds.ContainsKey($nid)) { return }
                $existingIds[$nid] = $true

                $isGroup = ($_.ObjectClass -eq 'group')
                $Global:PermissionModel += [PSCustomObject]@{
                    Identity = $id
                    Display  = $(if ($isGroup) { (' ' * ($indent * 4 - 2)) + "+ $id" } else { (' ' * ($indent * 4)) + $id })
                    Type     = $(if ($isGroup) { "Group" } else { "User" })
                    Parent   = $parentId
                    Indent   = $indent
                    Expanded = $false
                    Rights   = "(member of $name)"
                    Access   = ""
                }
            }
        } catch {}
    }
}

function Rebuild-ListView {
    param ($listView)
    $listView.Items.Clear()

    function Add-Visible ($parent) {
        $items = $Global:PermissionModel | Where-Object { $_.Parent -eq $parent }
        foreach ($item in $items) {
            $lv = New-Object System.Windows.Forms.ListViewItem($item.Display)
            [void]$lv.SubItems.Add($item.Rights)
            [void]$lv.SubItems.Add($item.Access)
            $lv.ForeColor = if ($item.Type -eq 'Group') { [System.Drawing.Color]::Blue } elseif ($item.Type -eq 'Error') { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::Black }
            $lv.Tag = $item
            $listView.Items.Add($lv)
            if ($item.Type -eq 'Group' -and $item.Expanded) {
                Add-Visible $item.Identity
            }
        }
    }

    Add-Visible ""
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Permissions Viewer"
$form.Size = New-Object System.Drawing.Size(1000, 550)
$form.StartPosition = "CenterScreen"

$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(400, 450)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.Anchor = 'Top, Bottom, Left'

$listView = New-Object System.Windows.Forms.ListView
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.Size = New-Object System.Drawing.Size(560, 450)
$listView.Location = New-Object System.Drawing.Point(420, 50)
$listView.Anchor = 'Top, Bottom, Left, Right'

[void]$listView.Columns.Add("Identity", 300)
[void]$listView.Columns.Add("Rights", 200)
[void]$listView.Columns.Add("Access", 100)

$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse Folder"
$browseButton.Size = New-Object System.Drawing.Size(120, 30)
$browseButton.Location = New-Object System.Drawing.Point(10, 10)

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$form.Controls.AddRange(@($treeView, $listView, $browseButton))

$treeView.Add_AfterSelect({
    $path = $treeView.SelectedNode.Tag
    if (Test-Path $path) {
        Load-Permissions -path $path
        Rebuild-ListView -listView $listView
    }
})

$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -ne 1) { return }
    $selectedItem = $listView.SelectedItems[0]
    $obj = $selectedItem.Tag
    if ($obj.Type -ne 'Group') { return }

    $obj.Expanded = -not $obj.Expanded
    if ($obj.Expanded) {
        Expand-Group -item $obj
    } else {
        $Global:PermissionModel = $Global:PermissionModel | Where-Object {
            $_.Identity -eq $obj.Identity -or $_.Parent -notlike "$($obj.Identity)*"
        }
    }

    Rebuild-ListView -listView $listView
})

$browseButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $treeView.Nodes.Clear()
        $listView.Items.Clear()
        $root = $folderBrowser.SelectedPath
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = [System.IO.Path]::GetFileName($root)
        $rootNode.Tag = $root
        $rootNode.ForeColor = [System.Drawing.Color]::Blue
        $treeView.Nodes.Add($rootNode)

        function Load-TreeView {
            param ($node, $path)
            Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $child = New-Object System.Windows.Forms.TreeNode
                $child.Text = $_.Name
                $child.Tag = $_.FullName
                $child.ForeColor = [System.Drawing.Color]::Blue
                $node.Nodes.Add($child)
                Load-TreeView -node $child -path $_.FullName
            }
            Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | ForEach-Object {
                $fileNode = New-Object System.Windows.Forms.TreeNode
                $fileNode.Text = $_.Name
                $fileNode.Tag = $_.FullName
                $fileNode.ForeColor = [System.Drawing.Color]::Black
                $node.Nodes.Add($fileNode)
            }
        }

        Load-TreeView -node $rootNode -path $root
        $rootNode.Expand()
        Load-Permissions -path $root
        Rebuild-ListView -listView $listView
    }
})

[void]$form.ShowDialog()
