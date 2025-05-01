Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

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

$form.Controls.Add($treeView)
$form.Controls.Add($listView)
$form.Controls.Add($browseButton)
[void]$form.ShowDialog()
