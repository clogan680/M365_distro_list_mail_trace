Import-Module ExchangeOnlineManagement

# Connect to Microsoft 365
Connect-ExchangeOnline

# Get all distribution groups in Microsoft 365
$groups = Get-DistributionGroup

# Initialize an array to store the results
$results = @()

$getdate = Get-Date
Write-Host('start date - ',$getdate.AddDays(-10))
Write-Host('end date - ',$getdate)

# Loop through each group
foreach ($group in $groups) {
   Write-Host($Group.PrimarySmtpAddress)
   $traceresults = Get-MessageTrace -RecipientAddress $group.PrimarySmtpAddress -StartDate $getdate.AddDays(-10) -EndDate $getdate
   if ($traceresults.length -eq 0) {
   $results += New-Object PSObject -Property @{
            Group = $group.DisplayName
            Sender = "No Results"
            Recipient = "No Results"
            Subject = "No Results"
            Date = "No Results"
        }
   }else {
   foreach ($traceresult in $traceresults) {
        # Store the relevant information in the results array
        $results += New-Object PSObject -Property @{
            Group = $group.DisplayName
            Sender = $traceresult.SenderAddress
            Recipient = $traceresult.RecipientAddress
            Subject = $traceresult.Subject
            Date = $traceresult.Received
        }
    }
   }
}

# Export
$results | Export-Csv -Path ".\MessageTraceResults.csv" -NoTypeInformation
