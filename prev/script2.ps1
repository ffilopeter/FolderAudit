Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Try to import AD module
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $adAvailable = $true
} catch {
    $adAvailable = $false
}

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

# Helper: Rights formatting
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

# Helper: Local group members
function Get-LocalGroupMembers {
    param ([string]$groupName)
    try {
        $output = net localgroup "$groupName" 2>&1
        $startIndexObj = $output | Select-String -SimpleMatch '----------'
        $endIndexObj = $output | Select-String -SimpleMatch 'The command completed successfully.'
        if ($startIndexObj -and $endIndexObj) {
            $startIndex = $startIndexObj[0].LineNumber
            $endIndex = $endIndexObj[0].LineNumber
            if ($endIndex -gt $startIndex) {
                $members = $output[($startIndex)..($endIndex - 2)] | Where-Object { $_.Trim() -ne '' }
                return $members
            }
        }
        return @()
    } catch { return @() }
}

# Load permissions
function Show-Permissions {
    param ([string]$path)
    $listView.Items.Clear()
    try {
        $acl = Get-Acl -Path $path
        foreach ($access in $acl.Access) {
            $identity = $access.IdentityReference.ToString()
            $rights = Format-Rights $access.FileSystemRights
            $accessType = $access.AccessControlType.ToString()

            $item = New-Object System.Windows.Forms.ListViewItem($identity)
            [void]$item.SubItems.Add($rights)
            [void]$item.SubItems.Add($accessType)
            $item.ForeColor = [System.Drawing.Color]::Black

            if ($identity -match '\\') {
                $domain, $name = $identity.Split('\', 2)
                if ($domain -match '^(BUILTIN|NT AUTHORITY)$') {
                    $item.Tag = @{ Type = "LocalGroup"; Name = $name; Expanded = $false }
                    $item.ForeColor = [System.Drawing.Color]::Blue
                } elseif ($adAvailable) {
                    try {
                        $adObject = Get-ADObject -Filter { SamAccountName -eq $name } -ErrorAction Stop
                        if ($adObject.ObjectClass -eq 'group') {
                            $item.Tag = @{ Type = "ADGroup"; Name = $name; Expanded = $false }
                            $item.ForeColor = [System.Drawing.Color]::Blue
                        }
                    } catch { }
                }
            }

            $listView.Items.Add($item)
        }
    } catch {
        $item = New-Object System.Windows.Forms.ListViewItem("Access Denied")
        [void]$item.SubItems.Add("")
        [void]$item.SubItems.Add("")
        $item.ForeColor = [System.Drawing.Color]::Red
        $listView.Items.Add($item)
    }
}

# Expand or collapse on double-click
$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -eq 1) {
        $selected = $listView.SelectedItems[0]
        $tag = $selected.Tag
        if ($tag -and $tag.ContainsKey("Type")) {
            $index = $listView.Items.IndexOf($selected) + 1

            if ($tag["Expanded"]) {
                # Collapse: remove all child entries starting with "  -> "
                while ($index -lt $listView.Items.Count -and $listView.Items[$index].Text -like "  ->*") {
                    $listView.Items.RemoveAt($index)
                }
                $tag["Expanded"] = $false
            } else {
                # Expand
                $groupType = $tag["Type"]
                $groupName = $tag["Name"]
                if ($groupType -eq "LocalGroup") {
                    $members = Get-LocalGroupMembers $groupName
                    foreach ($member in $members) {
                        $subItem = New-Object System.Windows.Forms.ListViewItem("  -> $member")
                        [void]$subItem.SubItems.Add("(local member of $groupName)")
                        [void]$subItem.SubItems.Add("")
                        $subItem.ForeColor = [System.Drawing.Color]::Black
                        $listView.Items.Insert($index, $subItem)
                        $index++
                    }
                } elseif ($groupType -eq "ADGroup" -and $adAvailable) {
                    try {
                        $members = Get-ADGroupMember -Identity $groupName -Recursive -ErrorAction Stop
                        foreach ($member in $members) {
                            $subItem = New-Object System.Windows.Forms.ListViewItem("  -> $($member.SamAccountName)")
                            [void]$subItem.SubItems.Add("(AD member of $groupName)")
                            [void]$subItem.SubItems.Add("")
                            $subItem.ForeColor = [System.Drawing.Color]::Black
                            $listView.Items.Insert($index, $subItem)
                            $index++
                        }
                    } catch { }
                }
                $tag["Expanded"] = $true
            }
        }
    }
})

$treeView.Add_AfterSelect({
    $selectedPath = $treeView.SelectedNode.Tag
    if (Test-Path $selectedPath) {
        Show-Permissions -path $selectedPath
    }
})

$browseButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $treeView.Nodes.Clear()
        $listView.Items.Clear()
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = [System.IO.Path]::GetFileName($folderBrowser.SelectedPath)
        $rootNode.Tag = $folderBrowser.SelectedPath
        $rootNode.ForeColor = [System.Drawing.Color]::Blue
        $treeView.Nodes.Add($rootNode)
        Load-TreeView -parentNode $rootNode -path $folderBrowser.SelectedPath
        $rootNode.Expand()
        Show-Permissions -path $folderBrowser.SelectedPath
    }
})

function Load-TreeView {
    param ([System.Windows.Forms.TreeNode]$parentNode, [string]$path)
    try {
        Get-ChildItem -Path $path -Directory -ErrorAction Stop | ForEach-Object {
            $dirNode = New-Object System.Windows.Forms.TreeNode
            $dirNode.Text = $_.Name
            $dirNode.Tag = $_.FullName
            $dirNode.ForeColor = [System.Drawing.Color]::Blue
            $parentNode.Nodes.Add($dirNode)
            Load-TreeView -parentNode $dirNode -path $_.FullName
        }
        Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | ForEach-Object {
            $fileNode = New-Object System.Windows.Forms.TreeNode
            $fileNode.Text = $_.Name
            $fileNode.Tag = $_.FullName
            $fileNode.ForeColor = [System.Drawing.Color]::Black
            $parentNode.Nodes.Add($fileNode)
        }
    } catch { }
}

$form.Controls.Add($treeView)
$form.Controls.Add($listView)
$form.Controls.Add($browseButton)
[void]$form.ShowDialog()
