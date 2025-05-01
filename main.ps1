Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

# =======================
# Add controls and render
# =======================

$form.Controls.Add($treeView)
$form.Controls.Add($listView)
$form.Controls.Add($browseButton)
[void]$form.ShowDialog()
