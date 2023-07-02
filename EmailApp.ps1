# Load assembly
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Define the form
$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '400,400'
$Form.text = "Form"

# Add labels and text boxes for necessary fields
$Fields = 'Excel File Path', 'HTML File Path', 'SMTP Server', 'Port', 'Username', 'Password', 'From', 'Subject'
$Top = 10
foreach ($Field in $Fields) {
    $Label = New-Object system.Windows.Forms.Label
    $Label.text = "$Field:"
    $Label.AutoSize = $true
    $Label.width = 25
    $Label.height = 10
    $Label.location = New-Object System.Drawing.Point(10,$Top)
    $Form.Controls.Add($Label)

    $TextBox = New-Object system.Windows.Forms.TextBox
    $TextBox.multiline = $false
    $TextBox.width = 250
    $TextBox.height = 20
    $TextBox.location = New-Object System.Drawing.Point(100,$Top)
    $TextBox.Name = $Field -replace ' ', ''
    $Form.Controls.Add($TextBox)

    $Top += 30
}

# Add button to start the email process
$Button = New-Object system.Windows.Forms.Button
$Button.text = "Start"
$Button.width = 60
$Button.height = 30
$Button.location = New-Object System.Drawing.Point(170,$Top)
$Button.Font = 'Microsoft Sans Serif,10'

# Actions to perform on button click
$Button.Add_Click({
    # Import the Excel module
    Install-Module -Name ImportExcel -Force -AllowClobber

    # Get the details from the form
    $ExcelFilePath = $Form.Controls['ExcelFilePath'].Text
    $HtmlFilePath = $Form.Controls['HTMLFilePath'].Text
    $SmtpServer = $Form.Controls['SMTPServer'].Text
    $Port = $Form.Controls['Port'].Text
    $Username = $Form.Controls['Username'].Text
    $Password = $Form.Controls['Password'].Text
    $From = $Form.Controls['From'].Text
    $Subject = $Form.Controls['Subject'].Text

    # Import the Excel file
    $EmailList = Import-Excel -Path $ExcelFilePath

    # Read the contents of the HTML file
    $HtmlTemplate = Get-Content -Path $HtmlFilePath -Raw

    # Loop over each recipient in the Excel file
    foreach ($Recipient in $EmailList) {
        # Replace placeholders in the HTML template with actual values from the Excel file
        $HtmlBody = $HtmlTemplate.Replace("{Name}", $Recipient.Name).Replace("{Address}", $Recipient.Address) # continue with other replacements as needed

        # Specify the details for the email
        $EmailDetails = @{
            SmtpServer = $SmtpServer       # your SMTP server
            Port = $Port                   # your SMTP server port
            UseSsl = $true                 # depending on your server, this may need to be $false
            Credential = New-Object System.Management.Automation.PSCredential ($Username, ($Password | ConvertTo-SecureString -AsPlainText -Force)) # your SMTP server username and password
            From = $From                   # your email address
            To = $Recipient.Email          # the recipient's email address
            Subject = $Subject             # the email's subject
            Body = $HtmlBody               # the email's body, as HTML
            BodyAsHtml = $true             # specify that the body is HTML
        }

        # Send the email
        Send-MailMessage @EmailDetails
    }
})
$Form.Controls.Add($Button)

# Show the form
[void]$Form.ShowDialog()
