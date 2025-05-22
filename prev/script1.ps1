Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Folder Tree Viewer"
$form.Size = New-Object System.Drawing.Size(600, 500)
$form.StartPosition = "CenterScreen"

# Create the TreeView
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Size = New-Object System.Drawing.Size(560, 400)
$treeView.Location = New-Object System.Drawing.Point(10, 50)
$treeView.Anchor = "Top, Bottom, Left, Right"

# Create the Button to Browse for Folder
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "Browse Folder"
$browseButton.Size = New-Object System.Drawing.Size(120, 30)
$browseButton.Location = New-Object System.Drawing.Point(10, 10)

# Folder Browser Dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# Function to Load Directory Tree
function Load-TreeView {
    param (
        [System.Windows.Forms.TreeNode]$parentNode,
        [string]$path
    )

    try {
        Get-ChildItem -Path $path -Directory -ErrorAction Stop | ForEach-Object {
            $node = New-Object System.Windows.Forms.TreeNode
            $node.Text = $_.Name
            $node.Tag = $_.FullName
            $parentNode.Nodes.Add($node)

            # Recursively add subdirectories
            Load-TreeView -parentNode $node -path $_.FullName
        }
    } catch {
        # Skip inaccessible folders
    }
}

# Button click event
$browseButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $treeView.Nodes.Clear()
        $rootNode = New-Object System.Windows.Forms.TreeNode
        $rootNode.Text = [System.IO.Path]::GetFileName($folderBrowser.SelectedPath)
        $rootNode.Tag = $folderBrowser.SelectedPath
        $treeView.Nodes.Add($rootNode)

        Load-TreeView -parentNode $rootNode -path $folderBrowser.SelectedPath
        $rootNode.Expand()
    }
})

# Add controls to form
$form.Controls.Add($treeView)
$form.Controls.Add($browseButton)

# Show the form
[void]$form.ShowDialog()
