param(
	 [Parameter(Mandatory=$True)][string]$url
	,[Parameter(Mandatory=$False)][string]$username
	,[Parameter(Mandatory=$False)][string]$password
    ,[Parameter(Mandatory=$False)][System.Management.Automation.PSCredential]$credential)
$ErrorActionPreference = 'Stop';

if([string]::IsNullOrEmpty($url)) {
   Write-Error 'Missing url.'
}

$securePassword = $null;
if((-not [string]::IsNullOrEmpty($username)) -and (-not [string]::IsNullOrEmpty($password))) {
   $securePassword = $password | ConvertTo-SecureString -AsPlainText -Force
}
 
$scriptLocation = Get-Location;
Install-Module SharePointPnPPowerShellOnline -Scope CurrentUser;
#Import-Module 'OfficeDevPnP.PowerShell.V16.Commands' -DisableNameChecking;

# connect/authenticate to SharePoint Online and get ClientContext object.. 
if ($securePassword) {
	$credential = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $securePassword
}
if ($credential) {
	Connect-SPOnline -Url $url -Credentials $credential
} else {
	Connect-SPOnline -Url $url
}
$clientContext = Get-SPOContext
$clientContext.Load($clientContext.Web);
$clientContext.ExecuteQuery();

$SiteCollectionUrl = $clientContext.Web.Url;
$extensionFolder = '{ExtensionFolder}';

#### Create Folders ####################
Write-Host 'Create Folders'
function AddSPOFolder {
	param (
         $Name
        ,$Folder
    )
	process {
		$clientContext = Get-SPOContext;
		$result = $clientContext.Web.GetFolderByServerRelativeUrl($clientContext.Web.Url + $Folder).Folders.Add($Name);
		$clientContext.ExecuteQuery();
	}
}
$parentFolder = '';
$extensionFolder -split '/' | ForEach {
    $folderName = $_;
    if ($folderName.Length -gt 0) {
        if ($parentFolder.Length -eq 0) {
            #Do not create root folder.
        } else {
            AddSPOFolder -Name $folderName -Folder $parentFolder
        }
        $parentFolder += '/' + $folderName;
    }
}

#### Upload Content Files ####################
Write-Host 'Upload Files'
function UploadFiles{
    param (
         $item
        ,$folder
    )
    process {
        if ($item.Name -match '.dwp' -or $item.Name -match '.js.map') {
            #Do nothing. These files are not uploaded
        } else { 
            If ($item -is [System.IO.DirectoryInfo]){
                Write-Host "$folder/$item/"
                AddSPOFolder -Name $item.Name -Folder $folder
                $subFolder = $folder + '/' + $item.Name;
                Get-ChildItem $item.FullName | ForEach {
                    $item2 = $_;
                    UploadFiles -item $item2 -folder $subFolder    
                }
            } ElseIf ($item.Name -match '.html' ) {
                Write-Host "$folder/$item";
                (Get-Content $item.FullName).replace('{SiteCollectionUrl}', $SiteCollectionUrl) | Set-Content $item.FullName
                $fileToUpload = $item.FullName;
                $result = Add-SPOFile -Path $fileToUpload -Folder $folder
            } Else {
                Write-Host "$folder/$item";
                $fileToUpload = $item.FullName;
                $result = Add-SPOFile -Path $fileToUpload -Folder $folder
            }
        }
    }
}
Get-ChildItem ($scriptLocation.ToString() + '\..\content') | ForEach {
    $item = $_;
    UploadFiles -item $item -folder $extensionFolder
}
$clientContext.ExecuteQuery();

#### Upload web part Files ####################
Write-Host 'Upload web part Files'
Get-ChildItem ($scriptLocation.ToString() + '\..\content') | ForEach {
    $item = $_;
    if ($item.Name -match '.dwp') {
        $folder = '_catalogs/wp';
        Write-Host "$folder/$item";
        (Get-Content $item.FullName).replace('{SiteCollectionUrl}', $SiteCollectionUrl) | Set-Content $item.FullName
        $fileToUpload = $item.FullName;
        $result = Add-SPOFile -Path $fileToUpload -Folder $folder
        $fullFileUrl = $clientContext.Web.ServerRelativeUrl + $folder + '/' + $item.Name;
        $file = $clientContext.Web.GetFileByServerRelativeUrl($fullFileUrl);
        $item = $file.ListItemAllFields
        $item['Group'] = 'Sensei';
        $item.Update();
    }
}
$clientContext.ExecuteQuery();

Write-Host 'Done.'
