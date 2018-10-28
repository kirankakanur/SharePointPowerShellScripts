$targetSite = "https://kiran.sharepoint.com/"   # your SPO site URL here
$outputFile = "C:\temp\SPSearchResults2.csv"    # path to the CSV file that will store the SharePoint search API results
$userName = "kirank@microsoft.com"              # your SPO user name here
$password = "pass@word1"                        # your password here

$encPassword = ConvertTo-SecureString -String $password -AsPlainText -Force
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $encPassword

Connect-PnPOnline $targetSite -Credentials $creds
$results = Submit-PnPSearchQuery -Query "ContentTypeId:0x010100DC221BCE2C5C654E8F0FBD8BE07D329F*" -SelectProperties "Title, OriginalPath" -MaxResults 20


#Array to Hold Results Collection - PSObjects
$ResultsCollection = @()   

 $results.ResultRows | foreach {

  $ExportItem = New-Object PSObject 
  $ExportItem | Add-Member -MemberType NoteProperty -name "Title" -value $_.Title
  $ExportItem | Add-Member -MemberType NoteProperty -name "Path" -value $_.OriginalPath

#Add the object with property to an Array
$ResultsCollection += $ExportItem

}    

#Export the result Array to CSV file
$ResultsCollection | Export-CSV $outputFile -NoTypeInformation

Write-Host "Done!"
