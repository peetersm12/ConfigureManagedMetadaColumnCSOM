function update-site ($web, $action, $SPOCredentials, $title){
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($web)   
	$context.Credentials = $SPOCredentials
	$web = $context.Web 
	$context.Load($web)
	$context.Load($web.Webs)
	$context.Load($web.Lists)

	#send the request containing all operations to the server
	try{
		$context.executeQuery()
		
		# Change any subweb if present
		if ($web.Webs.Count -ne 0){
			foreach ($subweb in $web.Webs){
				update-site -web $subweb.url -action $action -SPOCredentials $SPOCredentials -title $title
			}
		}

        write-host "URL: $($web.url)" -foregroundcolor Cyan

		foreach($list in $web.lists){
            write-host "List: $($list.title)" -foregroundcolor yellow
            $context.Load($list)
            $field = $list.fields.getbytitle($title)

            $context.load($field)
            try{
                $context.ExecuteQuery()

                if($action -eq "Preview"){
                    $taxField = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($context, $field)
                    $result = $taxField.open
                    write-host "Found: $($title) found and allow fill in is $($result)" -foregroundcolor green
                }
                if($action -eq "Update"){
                    $taxField = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.Taxonomy.TaxonomyField]).Invoke($context, $field)
                    $result = $taxField.open
                    $taxfield.open = $true
                    $taxfield.update()

                    try{
                        $context.executequery()
                    }
                    catch{
                        write-host "Error: $($_.Exception.Message)" -foregroundcolor red
                    }
                    write-host "Found: $($title) found and allow fill in was $($result) and has been set to true" -foregroundcolor green
                }
                
            }
            catch{
                write-host "Not found: No field with title $($title) found"
            }
        }
	}
	catch{
		write-host "Error: $($_.Exception.Message)" -foregroundcolor red
	}
}

function set-mmcolumn{
    param(
    [Parameter(mandatory=$false)]
    [string] $Action
    )
    # Let the user fill in their username and password in the PowerShell window
    $userName = Read-Host "Please enter the username with sufficient permissions"
    $password = Read-Host "Please enter the password for $($userName)" -AsSecureString
    $adminUrl = Read-Host "Please enter the admin URL to verify permissions"
    $siteUrl = Read-Host "Please enter the SharePoint site URL"
    $loglocation = Read-Host "Enter log location for the PowerShell Transcript"
    $ColumnTitle = Read-Host "Enter the title of the column"

    #start transcript
    $date = (get-date).tostring('sshhMMddyyyy')
    Start-Transcript -Path "$($loglocation)\Transcript_$($date).txt" -NoClobber

	# set SharePoint Online and Office 365 credentials
	$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $password)
	$credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $password
	
    #import taxonomy dll
    Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"

	#connect to to Office 365
	try{
		Connect-SPOService -Url $adminUrl -Credential $credentials
		write-host "Info: Connected succesfully to Office 365" -foregroundcolor green
	}
	catch{
		write-host "Error: $($_.Exception.Message)" -foregroundcolor red
		Break Change-SPOWebs
	}
    
    update-site -web $siteUrl -action $Action -SPOCredentials $SPOCredentials -title $ColumnTItle

    #stop transcript
    stop-transcript
}

UhH13AJa7Jq
