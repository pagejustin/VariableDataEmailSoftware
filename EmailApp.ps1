Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ExcelFile = ""
$SheetName = "Sheet1"
$Recipients = @()
$SMTPServer = "smtp.gmail.com"
$SMTPPort = 587
$SMTPUsername = "youremail@gmail.com"
$SMTPPassword = "yourpassword"
$From = "youremail@gmail.com"
$Subject = "Email Subject"
$BodyTemplate = Get-Content -Path "C:\Users\user\Documents\EmailBodyTemplate.html" -Raw

$form = New-Object System.Windows.Forms.Form
$form.Text = "Email Sender"
$form.Size = New-Object System.Drawing.Size(400,400)
$form.StartPosition = "CenterScreen"

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(280,20)
$label1.Text = "Enter the email subject:"
$form.Controls.Add($label1)

$textbox1 = New-Object System.Windows.Forms.TextBox
$textbox1.Location = New-Object System.Drawing.Point(10,40)
$textbox1.Size = New-Object System.Drawing.Size(280,20)
$form.Controls.Add($textbox1)

$listbox1 = New-Object System.Windows.Forms.ListBox
$listbox1.Location = New-Object System.Drawing.Point(10,70)
$listbox1.Size = New-Object System.Drawing.Size(280,200)
foreach ($Recipient in $Recipients) {
    $listbox1.Items.Add($Recipient.EmailAddress)
}
$form.Controls.Add($listbox1)

$button3 = New-Object System.Windows.Forms.Button
$button3.Location = New-Object System.Drawing.Point(10, 280)
$button3.Size = New-Object System.Drawing.Size(75, 23)
$button3.Text = "Select Excel File"
$form.Controls.Add($button3)

$button3.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
    $openFileDialog.Title = "Select an Excel File"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $ExcelFile=$openFileDialog.FileName
        $Recipients=Import-Excel -Path $ExcelFile -WorksheetName $SheetName
        foreach ($Recipient in $Recipients) {
            $listbox1.Items.Add($Recipient.EmailAddress)
        }
    }
})

$warningLabel=New-Object System.Windows.Forms.Label
$warningLabel.Location=New-Object System.Drawing.Point(10,310)
$warningLabel.Size=New-Object System.Drawing.Size(280,40)
$warningLabel.Text="Warning: This application can be used to send emails to multiple recipients. Be careful when using this application and ensure that you have permission to send emails to the selected recipients."
$form.Controls.Add($warningLabel)

$button1 = New-Object System.Windows.Forms.Button
$button1.Location = New-Object System.Drawing.Point(100,340)
$button1.Size = New-Object System.Drawing.Size(75,23)
$button1.Text = "Send"
$button1.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $button1
$form.Controls.Add($button1)

$button2 = New-Object System.Windows.Forms.Button
$button2.Location = New-Object System.Drawing.Point(190,340)
$button2.Size = New-Object System.Drawing.Size(75,23)
$button2.Text = "Cancel"
$button2.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $button2
$form.Controls.Add($button2)

$form.Topmost=$true

$result=$form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    foreach ($EmailAddress in $listbox1.SelectedItems) {
        $To=$EmailAddress.ToString()
        $Body=$BodyTemplate -replace "__FirstName__", $Recipient.FirstName -replace "__LastName__", $Recipient.LastName -replace "__Email__", $Recipient.EmailAddress
        Send-MailMessage -To $To -From $From -Subject $textbox1.Text -Body $Body -SmtpServer $SMTPServer -Port $SMTPPort
    }
}

$form.Dispose()
