# Connect to SharePoint Online
 $targetSite = "https://kirank.sharepoint.com/"
 $outputFile = "C:\temp\SPSearchResults2.csv" 
 $targetSiteUri = [System.Uri]$targetSite

Connect-SPOnline $targetSite

# Retrieve the client credentials and the related Authentication Cookies
 $context = (Get-PnPWeb).Context #(Get-SPOWeb).Context
 $credentials = $context.Credentials
 $authenticationCookies = $credentials.GetAuthenticationCookie($targetSiteUri, $true)

# Set the Authentication Cookies and the Accept HTTP Header
 $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
 $webSession.Cookies.SetCookies($targetSiteUri, $authenticationCookies)
 $webSession.Headers.Add("Accept", "application/json;odata=verbose")

# Set request variables
$apiUrl = $targetSite + "_api/search/query?querytext='ContentTypeId:0x010100DC221BCE2C5C654E8F0FBD8BE07D329F*'&selectproperties='URL,Title'&rowlimit=10&trimuduplicates=true"

# Make the REST request
 $webRequest = Invoke-WebRequest -Uri $apiUrl -Method Get -WebSession $webSession

# Consume the JSON result
 $jsonLibrary = $webRequest.Content | ConvertFrom-Json 

 #Array to Hold Results Collection - PSObjects
$ResultsCollection = @()   

$results = $jsonLibrary.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results


$results | foreach {        
        $ExportItem = New-Object PSObject 
        $_.Cells.Results | foreach {                  

              if($_.Key -eq "Title")
              {
                 $ExportItem | Add-Member -MemberType NoteProperty -name "Title" -value $_.Value                              
              }

               if($_.Key -eq "URL")
              {
                 $ExportItem | Add-Member -MemberType NoteProperty -name "URL" -value $_.Value                              
              }                               
        }
        $ResultsCollection += $ExportItem         
    }

#Export the result Array to CSV file
$ResultsCollection | Export-CSV $outputFile -NoTypeInformation

Write-Host "Done!"
 