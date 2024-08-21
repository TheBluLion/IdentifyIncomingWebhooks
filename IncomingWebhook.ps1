# Description: This script will list all Teams, along with its owners, that have an incoming webhook installed.
<###########################################
This code is provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code,
provided that You agree:  to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded;
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees,
that arise or result from the use or distribution of the this code
###########################################>

# Define the application credentials
$clientId = "*******CLIENT ID*********"
$clientSecret = "*******CLIENT SECRET*********"
$tenantId = "*******TENANT ID*********"

# The incoming webhook app id to filter the installed apps for each Team
$incomingWebhookAppId = "MjAzYTFlMmMtMjZjYy00N2NhLTgzYWUtYmU5OGY5NjBiNmIyIyMxLjAjI1B1Ymxpc2hlZA=="

# Get Access Token using OAuth 2.0, client-credentials flow
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$Body = @{    
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientId
    Client_Secret = $clientSecret
} 
$authToken = Invoke-RestMethod -Uri $uri -Method POST -Body $Body
$accessTokenValue = $authToken.access_token

# Use the access token for further operations
# Get the list of Teams using the incoming webhooks
$path = read-host "Enter Path to report location ending in .csv"

function Get-GraphResult($url,$token){
    try {
      $Headers = @{Authorization = "Bearer $($token)"}
      Invoke-RestMethod $url -Method Get -Headers $Headers
    } 
    catch {
      if($($_.Exception) -like "*429*"){
        start-sleep -Milliseconds 500
        $Headers = @{Authorization = "Bearer $($token)"}
        Invoke-RestMethod $url -Method Get -Headers $Headers
      }
      else{$_.Exception}
    }
}

function Get-GraphPost($url,$token, $body){
    try {
      $Headers = @{Authorization = "Bearer $($token)";"Content-Type"="application/json"}
      Invoke-RestMethod $url -Method POST -Headers $Headers -Body $body
    } 
    catch {
      if($($_.Exception) -like "*429*"){
        start-sleep -Milliseconds 500
        $Headers = @{Authorization = "Bearer $($token)";"Content-Type"="application/json"}
        Invoke-RestMethod $url -Method Get -Headers $Headers -body $body
      }
      else{"error: $($_.Exception)"}
    }
}

#check for temp file to see this is a resumed attempt
$url=get-content "$env:TEMP\webhook.tmp" -ea 0 


#this if clause looks to see if the saved url is a valid skiptoken request, if not, just start over
if (($url -eq $null) -or ($url -notlike "*skiptoken*")){
    $url = "https://graph.microsoft.com/beta/groups?`$top=100&`$filter=resourceProvisioningOptions/any(c:c eq 'Team')&`$select=id"
}

while($url -notlike ""){
    #write url to temp file for persistence
    try{$url | out-file "$env:TEMP\webhook.tmp"}
    catch{Write-Warning "Unable to save temp information for safe resume on crash"}
    #get page of teams
    $result=Get-GraphResult -url $url -token $accessTokenValue
    #create batch request
    if ($result -notlike "Error:*"){
        $i=1
        $array=foreach($line in $result.value){
            @{"id"=$i;"method"="GET";"url"="/teams/$($line.id)/installedApps?`$expand=teamsAppDefinition&`$filter=teamsAppDefinition/id eq '$($incomingWebhookAppId)'"} #"MjAzYTFlMmMtMjZjYy00N2NhLTgzYWUtYmU5OGY5NjBiNmIyIyMxLjAjI1B1Ymxpc2hlZA=="
            $i++
        }
        $body=@{"requests"=$array}|ConvertTo-Json   
    }
    else {
        write-error "Unexpected error $result"
        break
    }

    #submit batch request
    $appidresults=Get-GraphPost -url "https://graph.microsoft.com/v1.0/`$batch" -token $accessTokenValue -body $body
    #filter results with 0 instances of the webhook
    $appIdResults = $appidresults.responses| ?{$_.body.'@odata.count' -gt 0}
    
    $getownerbatchindex=1    
    $getownerbatcharray=@()
    #loop through identified teams, get owners and write outputs
    foreach($appIdIndex in $appIdresults){

        #get owners
            try {               
            
                $groupid = $result.value[$appIdIndex.id-1].id  
                if($null -ne $groupid){
                    #add request to batch request
                    #$getownerbatcharray.add(@{"id"=$getownerbatchindex;"method"="GET";"url"="/groups/$($groupid)?`$expand=owners&`$select=id,displayName,owners"})
                    $getownerbatcharray += @{"id"=$getownerbatchindex;"method"="GET";"url"="/groups/$($groupid)?`$expand=owners&`$select=id,displayName,owners"};
                    $getownerbatchindex++

                    if($getownerbatchindex -gt 15){
                        $body=@{"requests"=$getownerbatcharray}|ConvertTo-Json
                        $ownerresults=Get-GraphPost -url "https://graph.microsoft.com/v1.0/`$batch" -token $accessTokenValue -body $body
                        $getownerbatchindex=1    
                        $getownerbatcharray=@()
                        #write to csv
                        $ownerresults.responses.body|Select id,displayName,@{n='hasIncomingWebhook';e={"Yes"}},@{n='owners';e={($_.owners).mail -join ";"}}  |export-csv $path -Append -NoTypeInformation
                    }
                }
            }
            catch {

                    Write-Host "Exception occurred: $($_.Exception)"
                
            }
                  
        }


    #update nextlink
    $url=$result.'@odata.nextLink'
}

#do one last export of getownerbatcharray as this is may be less than 15

if($getownerbatcharray.count -gt 0){
    $body=@{"requests"=$getownerbatcharray}|ConvertTo-Json
    $ownerresults=Get-GraphPost -url "https://graph.microsoft.com/v1.0/`$batch" -token $accessTokenValue -body $body
    #write to csv
    $ownerresults.responses.body|Select id,displayName,@{n='hasIncomingWebhook';e={"Yes"}},@{n='owners';e={($_.owners).mail -join ";"}}  |export-csv $path -Append -NoTypeInformation
}



#cleanupstuff

Write-Host "Report completed. File saved to $($path)"
try{del "$env:TEMP\webhook.tmp" }
catch{Write-Warning "Unable to clean up temp information for safe resume on crash"}
write-host "Please close this powershell window to terminate the token refresh process"