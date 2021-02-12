#requires -modules MicrosoftTeams

import-module MicrosoftTeams
$username = Connect-MicrosoftTeams 
$path = "c:\users\euloh\incomingwebhookreport.csv"

#this hash table syncs the access token between multiple threads
$hash = [hashtable]::Synchronized(@{})
$hash.path = $path
$hash.username=$username.Account.id
$hash.token="init"

#create runspace to refresh token
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()
$runspace.SessionStateProxy.SetVariable('Hash',$hash)
$powershell = [powershell]::Create()
$powershell.Runspace = $runspace

$powershell.AddScript({
    while($true){
        #this is the song that never ends
        $provider = [Microsoft.Open.Teams.CommonLibrary.TeamsPowerShellSession]::SessionProvider
        $endpoint = [Microsoft.Open.Teams.CommonLibrary.Endpoint]::MsGraphEndpointResourceId
        $tenantId = $provider.AccessTokens['AccessToken'].TenantId
        $PromptBehavior = [Microsoft.Open.Teams.CommonLibrary.ShowDialog]::Auto
        $account = [Microsoft.Open.Teams.CommonLibrary.AzureAccount]::new()
        $account.Type = [Microsoft.Open.Teams.CommonLibrary.AzureAccount+AccountType]::User
        $account.Id = $hash.username
        $hash.token = $provider.AuthenticationFactory.Authenticate($account, $provider.AzureEnvironment, $tenantId, $null, $PromptBehavior, $null, $provider.TokenCache, $endpoint).AccessToken
     start-sleep -Seconds 3300
    }
}) 

$handle = $powershell.BeginInvoke()

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

#check for temp file to see this is a resumed attempte

try{
    $url=get-content "$env:TEMP\webhook.tmp"
}
catch{
    $url="failed to read temp file"
}

#this if clause looks to see if the saved url is a valid skiptoken request, if not, just start over
if ($url -notlike "*skiptoken"){
    $url = "https://graph.microsoft.com/beta/groups?`$top=15&`$filter=resourceProvisioningOptions/any(c:c eq 'Team')&`$select=id,displayName"
}


while($url -notlike ""){
    #write url to temp file for persistence
    try{$url | out-file "$env:TEMP\webhook.tmp"}
    catch{Write-Warning "Unable to save temp information for safe resume on crash"}
    #get page of teams
    $result=Get-GraphResult -url $url -token $hash.token
    #create batch request
    if ($result -notlike "Error:*"){
        $i=1
        [System.Collections.ArrayList]$array
        $array=foreach($line in $result.value){
            @{"id"=$i;"method"="GET";"url"="/teams/$($line.id)/installedApps?`$expand=teamsAppDefinition&`$filter=teamsAppDefinition/id eq 'MjAzYTFlMmMtMjZjYy00N2NhLTgzYWUtYmU5OGY5NjBiNmIyIyMxLjA='"}
            $i++
        }
        $body=@{"requests"=$array}|ConvertTo-Json   
    }
    else {
        write-error "Unexpected error $result"
        break
    }

    #submit batch request
    $appidresults=Get-GraphPost -url "https://graph.microsoft.com/v1.0/`$batch" -token $hash.token -body $body
    #filter results with 0 instances of the webhook
    $appIdResults = $appidresults.responses| ?{$_.body.'@odata.count' -gt 0}
    
    #loop through identified teams, get owners and write outputs
    foreach($appIdIndex in $appIdresults){

        #get owners

        #write to csv
        $result.value[$appIdIndex.id-1] |Select id,displayName,@{n='hasIncomingWebhook';e={"Yes"}}  |export-csv $path -Append -NoTypeInformation
    }


    #update nextlink
    $url=$result.'@odata.nextLink'
    #save
}

#cleanupstuff

Write-Host "Report completed. File saved to $path"
try{remove-file "$env:TEMP\webhook.tmp"}
catch{Write-Warning "Unable to clean up temp information for safe resume on crash"}

#clean up runspace that refreshes token
$powershell.EndInvoke($handle)
$runspace.Close()
$powershell.Dispose()