# This script will replace links in the Quick Launch, Links Lists, and default
# page for each SPWeb object to ECM cloud.
#
# Basically converts links from the legacy format (http://dewey-das-v:8080...) into the new format (https://gno-test.t1cloud.com...)
#
# Author:Juan Villanueva
######################## Start Variables ########################
$siteURL = "http://localhost" #URL to any site in the web application.
#$siteURL = "http://intranet" #URL to any site in the web application.
$filePath = "C:\Users\MOSSAdmin\Documents\Links.csv"
$PublishingFeatureGUID = "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb"
######################## End Variables ########################
if(Test-Path $filePath)
{
 Remove-Item $filePath
}
Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Publishing")
[System.Reflection.Assembly]::LoadWithPartialName("System.Net.WebClient")

# Creates an object that represents an SPWeb's Title and URL
function CreateNewWebObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $web.Title
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $web.URL

 return $linkObject
}
# Creates an object that represents the header links of the Quick Launch
function CreateNewLinkHeaderObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $prevWebTitle
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $prevWebURL  
 $linkObject | Add-Member -type NoteProperty -Name QLHeaderTitle -Value $node.Title
 $linkObject | Add-Member -type NoteProperty -Name QLHeaderLink -Value $node.Url
 return $linkObject
}
# Creates an object that represents to the links in the Top Link bar
function CreateNewTopLinkObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $prevWebTitle
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $prevWebURL  
 $linkObject | Add-Member -type NoteProperty -Name TopLinkTitle -Value $node.Title
 $linkObject | Add-Member -type NoteProperty -Name TopLinkURL -Value $node.Url
 $linkObject | Add-Member -type NoteProperty -Name TopNavLink -Value $true
 return $linkObject
}
# Creates an object that represents the links of in the Quick Launch (underneath the headers)
function CreateNewLinkChildObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $prevWebTitle
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $prevWebURL
 $linkObject | Add-Member -type NoteProperty -Name QLHeaderTitle -Value $prevHeaderTitle
 $linkObject | Add-Member -type NoteProperty -Name QLHeaderLink -Value $prevHeaderLink
 $linkObject | Add-Member -type NoteProperty -Name QLChildLinkTitle -Value $childNode.Title
 $linkObject | Add-Member -type NoteProperty -Name QLChildLink -Value $childNode.URL
 return $linkObject
}
## Creates an object that represents items in a Links list.
function CreateNewLinkItemObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $prevWebTitle
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $prevWebURL
 $linkObject | Add-Member -type NoteProperty -Name ListName -Value $list.Title

 $spFieldURLValue = New-Object microsoft.sharepoint.spfieldurlvalue($item["URL"])

 $linkObject | Add-Member -type NoteProperty -Name ItemTitle -Value $spFieldURLValue.Description
 $linkObject | Add-Member -type NoteProperty -Name ItemURL -Value $spFieldURLValue.Url
 return $linkObject
}
# Determines whether or not the passed in Feature is activated on the site or not.
function FeatureIsActivated
{param($FeatureID, $Web)
 return $web.Features[$FeatureID] -ne $null
}
# Creates an object that represents a link within the body of a content page.
function CreateNewPageContentLinkObject
{
 $linkObject = New-Object system.Object
 $linkObject | Add-Member -type NoteProperty -Name WebTitle -Value $prevWebTitle
 $linkObject | Add-Member -type NoteProperty -Name WebURL -Value $prevWebURL
 $linkObject | Add-Member -type NoteProperty -Name PageContentLink -Value $link

 return $linkObject
}
$wc = New-Object System.Net.WebClient
$wc.UseDefaultCredentials = $true
$pattern = "(((f|ht){1}tp://)[-a-zA-Z0-9@:%_\+.~#?&//=]+)"
$site = new-object microsoft.sharepoint.spsite($siteURL)
$webApp = $site.webapplication
$allSites = $webApp.sites
$customLinkObjects =@()
$ecmLinkObjects =@()
$int = 1
foreach ($site in $allSites)
{
 $allWebs = $site.AllWebs

 foreach ($web in $allWebs)
 {
  ## If the web has the publishing feature turned OFF, use this method
  if((FeatureIsActivated $PublishingFeatureGUID $web) -ne $true)
  {
   $quickLaunch = $web.Navigation.QuickLaunch
   $customLinkObject = CreateNewWebObject
   $customLinkObjects += $customLinkObject

   $prevWebTitle = $customLinkObject.WebTitle
   $prevWebURL = $customLinkObject.WebURL

  }

  ## If the web has the publishing feature turned ON, use this method
  else
  {
   $publishingWeb = [Microsoft.SharePoint.Publishing.PublishingWeb]::GetPublishingWeb($web)
   $quickLaunch = $publishingWeb.CurrentNavigationNodes
   $customLinkObject = CreateNewWebObject
   $customLinkObjects += $customLinkObject

   $prevWebTitle = $customLinkObject.WebTitle
   $prevWebURL = $customLinkObject.WebURL
 

  }

  #Looking for lists of type Links
  $lists = $web.Lists
  foreach ($list in $lists)
  {
   if($list.BaseTemplate -eq "Links")
   {
    $prevWebTitle = $customLinkObject.WebTitle
    $prevWebURL = $customLinkObject.WebURL

    # Going through all the links in a Links List
    foreach ($item in $list.Items)
    {
     $customLinkObject = CreateNewLinkItemObject
     $customLinkObjects += $customLinkObject     
	 
	 
		if($item -and  -Not ([string]::IsNullOrEmpty($item["URL"])))
		{ 
	 
		 	$value = new-object Microsoft.SharePoint.SPFieldUrlValue($item["URL"].ToString())
			$theURL = $value.Url;
			
			if(-Not ([string]::IsNullOrEmpty($theURL)))
			{
				if ($theURL.StartsWith("http://dewey-das-v:8080") -Or $theURL.StartsWith("http://ecm:8080"))
				{
					
					$linkBy = $null
					$startIndex = $null
					$requiredPadding = $null
					$newLinkFormat = $null
					
					$startIndex = $theURL.IndexOf("/docsetid/")
		
					if ($startIndex -gt 0)
					{
						# We are dealing with a link to ECM that uses docsetid and the $startIndex variable is pointing to the index of the first / before "/docsetid/#####". 
						# To point to the index of the first character after the last / in "/docsetid/", we add 10 characters. Now we are pointing to the start of the docsetid number
						$linkBy = "docsetid"
						$requiredPadding = "10"
						$newLinkFormat = "https://gno-test.t1cloud.com/T1Default/CiAnywhere/Web/GNO-TEST/LogOn/NA?returnUrl=https://gno-test.t1cloud.com/T1Default/CiAnywhere/Web/GNO-TEST/Api/CMIS/T1/content/?id=folder-"
					}
					else {
						$startIndex = $theURL.IndexOf("/docid/")
						if ($startIndex -gt 0)
						{
							# We are dealing with a link to ECM that uses docid and the $startIndex variable is pointing to the index of the first / before "/docid/#####". 
							# To point to the index of the first character after the last / in "/docid/", we add 7 characters. Now we are pointing to the start of the docid number
							$linkBy = "docid"
							$requiredPadding = "7"
							$newLinkFormat = "https://gno-test.t1cloud.com/T1Default/CiAnywhere/Web/GNO-TEST/LogOn/NA?returnUrl=https://gno-test.t1cloud.com/T1Default/CiAnywhere/Web/GNO-TEST/Api/CMIS/T1/content/?id=document-"
						}
					}
					
					if(-Not ([string]::IsNullOrEmpty($linkBy)))
					{
						
						$startIndex = $startIndex + $requiredPadding
						
						Write-Host $theURL.Substring($startIndex)
						
						#Lets now get the index of the first / after $startIndex in $theURL
						
						$endIndex = $theURL.Substring($startIndex).IndexOf("/")
						
						# Now we can find the docsetid number as a substring between the $startIndex and the $endIndex
						
						$theID = $theURL.Substring($startIndex, $endIndex)
						
						Write-Host $theID
						
						# Now, we are ready to replace the link with the new link
						
						$newLink = ($newLinkFormat + $theID)
						
						Write-Host $newLink
					
						# Replace just the first link
						if ($int -le 2)
						{
							Write-Host "Press any key to continue ..."
							$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
						
							$newSPFieldUrlValue = new-object Microsoft.SharePoint.SPFieldUrlValue($Item["URL"])
							$newSPFieldUrlValue.URL = $newLink
							$newSPFieldUrlValue.Description = $newSPFieldUrlValue.Description
							$item["URL"] = $newSPFieldUrlValue
							$item.Update()
							$int++
						
						}
					}
						
				}
			
			} 
	 
	 	}
    }

Write-Host $list.Title
   }  
  }

  #Looking at the default page for each web for links embedded within the content areas
  $htmlContent = $wc.DownloadString($web.URL)
  $result = $htmlContent | Select-String -Pattern $pattern -AllMatches
  $links = $result.Matches | ForEach-Object {$_.Groups[1].Value}
  foreach ($link in $links)
  {
   $customLinkObject = CreateNewPageContentLinkObject
   $customLinkObjects += $customLinkObject
  }

Write-Host $web.Title
  $web.Dispose()
 }
$site.dispose()
}

#Cleaning customLinkObject to leave only the ECM links
forEach ($myCustomLinkObject in $customLinkObjects)
{
	
	if(-Not ([string]::IsNullOrEmpty($myCustomLinkObject.ItemURL)))
	{
		if ($myCustomLinkObject.ItemURL.StartsWith("http://dewey-das-v:8080") -Or $myCustomLinkObject.ItemURL.StartsWith("http://ecm:8080"))
		{
			$ecmLinkObjects += $myCustomLinkObject
		}
		
	}
	
	if(-Not ([string]::IsNullOrEmpty($myCustomLinkObject.PageContentLink)))
	{
		if ($myCustomLinkObject.PageContentLink.StartsWith("http://dewey-das-v:8080") -Or $myCustomLinkObject.PageContentLink.StartsWith("http://ecm:8080"))
		{
			$ecmLinkObjects += $myCustomLinkObject
		}
		
	}
}

# Exporting the data to a CSV file
$ecmLinkObjects | Select-Object WebTitle,WebURL,TopNavLink,TopLinkTitle,TopLinkURL,QLHeaderTitle,QLHeaderLink,QLChildLinkTitle,QLChildLink,ListName,ItemTitle,ItemURL,PageContentLink | Export-Csv $filePath
write-host "Done"