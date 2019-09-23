param([string[]]$dir);

$psVersion = $PSVersionTable.PSVersion

if($psVersion.Major -lt 4){
    Add-Type –AssemblyName System.Windows.Forms
    $message = @"
Your System is not meeting the requirements of the Application. 
Try installing .Net Framework 4.5.
After that try installing Windows Management Framework 4.0.
"@
    [System.Windows.Forms.MessageBox]::Show($message , "Information")
    EXIT
}

####Get script location
if(!($dir)){
    $scriptpath = $MyInvocation.MyCommand.Path
    $dir = Split-Path $scriptpath
}

#$dir = "C:\Desktop\Temp\SP_Copy_Content_DEV\17_07\SP_Copy_Content_PROTECT"
$dirFormURLs = "file:///" + $dir -replace "\\", "/"

####Add CSOM DLLs
Add-Type -Path "$dir\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$dir\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "$dir\Microsoft.SharePoint.Client.Taxonomy.dll" 

####Global Variables. Some globals are not kept in Global Variables but in .tag in buttons.
$global:selectedItemsForCopy = @()
$global:itemsToCopy = @()
$global:userData = @()
$global:comboSelection = @()
$global:fieldValuesArrayForCSV = @()
$global:errorsAll = @()
$global:metadataAll = @()
$global:copyMode = ""
$global:sourceSPTypeCheck = ""
$global:destSPTypeCheck = ""
$global:fileShareRoot = ""
$global:modeSwitch = ""
$global:sourceURLGlobal = ""
$global:destURLGlobal = ""
$global:showOnlyOnce = ""
$global:preciseLocation = ""
$global:prevClickable = $false
$global:nextClickable = $false
[Microsoft.SharePoint.Client.FieldUserValue[]]$global:ownerValueCollection = @()

####Call Hide PowerShell Console
#[Void]$(Hide-PowerShellWindow)

####Fix Source Paths of images embeded in buttons
(Get-Content "$dir\MainWindow.xaml").Replace("CHANGELATERIMAGEPATH",$dirFormURLs) | Set-Content "$dir\MainWindow.xaml"

####NOT NEEDED. JUST FOR REFERENCE
#Add-Type -Path "C:\Desktop\Temp\PowerShell_FORMS\WpfAnimatedGif.dll"

####Form XAML
$inputXML = get-content "$dir\MainWindow.xaml"  
 
####Make XAML formed in Visual Studio readable
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
####Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
################################################
## Store Form Objects In PowerShell
################################################
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables

####Set Image Sources
#$WPFimage_Initial.source = $dirFormURLs+"/Initial.png"

####Set Window Icon
$base64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes("$dir\Logo.ico"))
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()
$bitmap.Freeze()
$Form.icon = $bitmap
#$Form.TaskbarItemInfo.Overlay = $bitmap

####Remember Fields
function Remember {
     $stringRemember = $WPFtextBox_URL.text +","+$WPFtextBox_User.text+","+$WPFtextBox_Pass.password+","+$WPFtextBox_URL_Dest.text+","+$WPFtextBox_User_Dest.text+","+$WPFtextBox_Pass_Dest.password
     $stringRemember | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "$dir\RC.txt"
}

####Recall Fields
function Recall {
    $checkForPrevRem = Test-Path "$dir\RC.txt"
    write-host "Do exist????"$checkForPrevRem -ForegroundColor Green
    if($checkForPrevRem -eq $true){
        $PrevRem = Get-Content "$dir\RC.txt" | ConvertTo-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($PrevRem)
        $PrevRemClear = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
        $PrevRemArray = $PrevRemClear.split(",")

        $WPFtextBox_URL.text = $PrevRemArray[0]
        $WPFtextBox_User.text = $PrevRemArray[1]
        $WPFtextBox_Pass.password = $PrevRemArray[2]
        $WPFtextBox_URL_Dest.text = $PrevRemArray[3]
        $WPFtextBox_User_Dest.text = $PrevRemArray[4]
        $WPFtextBox_Pass_Dest.password = $PrevRemArray[5]
    }
    ####Set Variable used to check Change with FileShare in Source URL field
    $global:sourceURLGlobal = $WPFtextBox_URL.text
}
Recall

################################################
## UI Related Functions
################################################

####Logging Function
function logEverything($relatedItemURL, $exceptionToLog){
    $global:errorsCount++

    write-host $relatedItemURL -foreground DarkYellow
    write-host $exceptionToLog -foreground DarkYellow


    $obj = new-object psobject -Property @{
        'FileRef' = $relatedItemURL
        'Exception' = $exceptionToLog
    }

    $global:errorsAll += $obj
    $errorsCount = $global:errorsAll.Count

    $WPFlabel_Error_Status.content = "Errors: $errorsCount"
}

####Progress Bar down left control
function progressbar($state, $all, $action){
    if($state -eq "Start"){
        $wpfProgressBar.Value = 0
        $wpfProgressBar.Maximum = $All
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
    if($state -eq "Plus"){
        $wpfProgressBar.Value++
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
    if($state -eq "Minus"){
        $wpfProgressBar.Value= $wpfProgressBar.Value--
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
    if($state -eq "Stop"){
        $wpfProgressBar.Value = 0
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
}

####Fade In/Fade Out Animation. Used mainly for WPF Grids.
function opacityAnimation ($grid, $action){
    if($action -eq "Pre"){
        $grid.Opacity = 0
        $grid.Visibility = "Visible"
    }
    if($action -eq "Open"){
        $grid.Visibility = "Visible"
        $i=0
        while ($i -lt 11){
            if($i -lt 10){
                $grid.Opacity = "0."+$i
            } else {
                $grid.Opacity = "1"
            }
            Start-Sleep -m 10 
            $Form.Dispatcher.Invoke("Background", [action]{})
            $i++
        }
    }
    if($action -eq "Close"){
        $i=9
        while ($i -gt -1){
            $grid.Opacity = "0."+$i
            Start-Sleep -m 10 
            $Form.Dispatcher.Invoke("Background", [action]{})
            $i--
        }
        $grid.Visibility = "Hidden"
    }
}

####Creates Credentials
function createCredentials($User, $Password, $spType){
    $Pass = $Password | ConvertTo-SecureString -AsPlainText -Force
    if($spType -eq "True"){
        try{
            $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Pass)
        } catch {
            ####If SP Type is wrong for On-Premise the Creds creation will fail and PS will try to login to SP with the User running PS. So we create right type Creds to fail the Auth.
            $user = "FailCreds"
            $Creds = New-Object System.Management.Automation.PSCredential($User,$Pass)
        }
    }
    if($spType -like "*False*"){
        $Creds = New-Object System.Net.NetworkCredential($User,$Pass)
    }
    if($spType -like "*PSCred*"){
        $Creds = New-Object System.Management.Automation.PSCredential($User,$Pass)
    } 
    return $Creds
}

####Gets Source and Destination Lists
function getLists ($User, $Password, $SiteURL, $listView,$spType,$SourceOrTarget){
    $listGetSuccess = $true
    progressbar -state "Start" -all 2
    progressbar -state "Plus" 
    $Creds = createCredentials -User $User -Password $Password -spType $spType
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Context.Credentials = $Creds
    $Form.Dispatcher.Invoke("Background", [action]{})
    ####Get List
    $Lists = $Context.Web.Lists
    $Context.Load($Lists)
    progressbar -state "Plus"
    try{
        $Context.ExecuteQuery()
    } catch {
        $WPFlabel_Status.Content = "Status: Connection Failed!"
        $Form.Dispatcher.Invoke("Background", [action]{})
        write-host $_.exception.message -foreground yellow
        $listGetSuccess = $false
    }
    if($listGetSuccess -eq $true){
        $allListsArray = @()
        write-host "Type"$global:sourceSPTypeCheck
                write-host "Tar"$SourceOrTarget
        if(($SourceOrTarget -eq "Target") -and (($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta"))){
            foreach ($list in $lists | where {$_.BaseType -eq "DocumentLibrary"}){
                $obj = new-object psobject -Property @{
                    'Title' = $list.Title
                    'Type' = $list.BaseType
                }
                $allListsArray+=$obj
                #$listView.AddChild($obj)
            }
        } else {
            foreach ($list in $lists){
                $obj = new-object psobject -Property @{
                    'Title' = $list.Title
                    'Type' = $list.BaseType
                }
                $allListsArray+=$obj
                #$listView.AddChild($obj)
            }
        }
        $listView.ItemsSource = $allListsArray
        #$Context.dispose()

        ####Pass Contexts to Global Variables
        if($SourceOrTarget -eq "Source"){
        ####If Context is Online
            $global:Context = $Context
        }
        if($SourceOrTarget -eq "Target"){
            $global:destContext = $Context
        }
        progressbar -state "Plus"
        progressbar -state "Stop" 

            return $listGetSuccess
    } else {
        progressbar -state "Stop" 
    }
 }

 ####Gets Fields in Source and Destination List
function getFields ($User, $Password, $SiteURL, $docLibName, $listView, $spType, $SourceOrTarget){
    $Creds = createCredentials -User $User -Password $Password -spType $spType
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Context.Credentials = $Creds

    ####Get Source List
    $List = $Context.Web.Lists.GetByTitle($DocLibName)
    $fields = $list.Fields
    $Context.Load($List)
    $Context.Load($fields)
    $Context.ExecuteQuery()
    $allFieldsArray = @()

    $permitedTrueBaseTypeFields = "Author", "Editor", "Modified", "Created", "_ModerationStatus", "_ModerationComments", "Title"
    foreach ($field in $fields | where {$_.Hidden -eq $False}){       
        $obj = new-object psobject -Property @{
        'Name' = $field.InternalName
        'Type' = $field.TypeDisplayName 
        'BaseType' = $field.FromBaseType
        } 

        ####Fiter out BaseType True fields while keeping those who are known to be editable
        if($obj.BaseType -eq "True"){
            if($permitedTrueBaseTypeFields -contains $obj.Name){
                $allFieldsArray += $obj
            }
        } else {
            $allFieldsArray += $obj       
        }
    }
    $listView.ItemsSource = $allFieldsArray 
    return $allFieldsArray
 }

################################################
## Copy Related Functions
################################################

####Called by copyMetadata. Ensures User on Source, takes Login Name and tries to ensure on Destination. Returns an array of ensured users for a field. Supports Multiuser fields. If there is user mapping will apply that.
function ensureUser($sourceFieldUV){
                        $multiUserArray = @()
                        foreach ($user in $sourceFieldUV){
                            if($user){
                                    write-host "Ensuring User on Source"

                                    ####If SP is On-Premise ensure by Lookup Value. If it is Online ensure by Mail.
                                    if(($global:sourceSPTypeCheck -eq "False") -or ($global:sourceSPTypeCheck -eq "False 2010")){
                                        write-host "Type Check On-Prem"
                                        $userSourceEnsured = $Web.EnsureUser($user.LookupValue)
                                    }
                                    if($global:sourceSPTypeCheck -eq "True"){
                                        write-host "Type Check Online"
                                        $userSourceEnsured = $Web.EnsureUser($user.Email)
                                    }
                                    $Context.load($userSourceEnsured)
                                    try{
                                        $Context.ExecuteQuery()
                                    } catch {
                                        Write-host "Ensure Source User Failed" -ForegroundColor Yellow
                                    }
                                    write-host "What are we ensuring on Destination" $userSourceEnsured.LoginName -ForegroundColor Green

                                    ####If no User Mapping CSV is loaded try to ensure Source Login on Destination
                                    if(!($global:userData)){
                                    $userToEnsureOnDest = $userSourceEnsured.LoginName
                                        ####Stripping down the user Login name to only the Name without pipe. 2010 gets confused with multiple users found if you search with the pipe
                                        if($global:destSPTypeCheck -eq "False 2010"){        
                                            $userToEnsureOnDest = $userSourceEnsured.LoginName
                                            $userToEnsureOnDest = $userToEnsureOnDest.split("|")[-1]
                                            write-host "User Split" $userToEnsureOnDest
                                        }
                                    ####If User Mapping CSV is loaded map user and ensure
                                    } else {
                                        $userToEnsureOnDest = $userSourceEnsured.LoginName
                                        write-host "What is the Login Name" $userToEnsureOnDest
                                        $userToEnsureOnDest = $userToEnsureOnDest.split("|")[-1] 
                                        write-host "User split for Mapping" $userToEnsureOnDest -ForegroundColor Magenta
                                        $userToEnsureOnDest = $global:userData | where {$_.Source_User -eq $userToEnsureOnDest}    
                                        write-host "What we got" $userToEnsureOnDest -ForegroundColor Green 
                                        $userToEnsureOnDest = $userToEnsureOnDest.Destination_User
                                    }
                                    $userDestEnsured = $destWeb.EnsureUser($userToEnsureOnDest)
                                    $destContext.load($userDestEnsured)
                                    try{
                                        $destContext.ExecuteQuery()
                                    } catch{
                                        Write-host "Ensure Destination User Failed" -ForegroundColor Yellow
                                    }
                                    if($userDestEnsured.id ){
                                        write-host "Dest User ID Found" -ForegroundColor Green
                                        $spuserValue = New-Object Microsoft.SharePoint.Client.FieldUserValue
                                        $spuserValue.LookupId = $userDestEnsured.id 
                                        $spuserValue | Add-Member -type NoteProperty -Name 'BelongingTo' -Value $field.SourceName
                                        $multiUserArray += $spuserValue
                                    }
                            }
                    }

                            return $multiUserArray
}

function ensureUserMetaDataFile($sourceFieldUV){
    if($sourceFieldUV -contains ","){
        $sourceFieldUV = $sourceFieldUV.split(",")
    }

    foreach ($user in $sourceFieldUV) {
        if($global:userData){
            write-host "Searching for User in CSV" -ForegroundColor DarkYellow
            $findUser = $user.split("|")[-1] 
            $userValue = $global:userData | where {$_.Source_User -eq $findUser}
            $user = $userValue.Destination_User
        } else {
            if($global:destSPTypeCheck -eq "False 2010"){
                $user = $user.split("|")[-1]
            }
        }
        write-host "Ensuring User" $user -ForegroundColor DarkYellow
        $userDestEnsured = $destWeb.EnsureUser($user)
        $destContext.load($userDestEnsured)
        try{
            $destContext.ExecuteQuery()
        } catch{
            Write-host "Ensure Destination User Failed" -ForegroundColor Yellow
        }

        if($userDestEnsured.id){
            write-host "Dest User ID Found" -ForegroundColor Green
            $spuserValue = New-Object Microsoft.SharePoint.Client.FieldUserValue
            $spuserValue.LookupId = $userDestEnsured.id 
            $spuserValue | Add-Member -type NoteProperty -Name 'BelongingTo' -Value $field.SourceName
            $multiUserArray += $spuserValue
        }
    }
    return $multiUserArray
}

####Ensures Destination Login User. We fall back on this user if ensure of user fails
function ensureOwner{
        ####Prepare owner as ensured user
        write-host "Ensuring Owner"$WPFtextBox_User_Dest.text
        $userOwnerEnsured = $destWeb.EnsureUser($WPFtextBox_User_Dest.text)
        $destContext.load($userOwnerEnsured)
        try{
            $destContext.ExecuteQuery()
            } catch{
            Write-host "Ensure Destination User Failed" -ForegroundColor Green
            write-host $_.exception.message -foreground yellow
            }
        if($userOwnerEnsured.id ){
            write-host "Dest User ID Found" -ForegroundColor Green
            $spOwnerValue = New-Object Microsoft.SharePoint.Client.FieldUserValue
            $spOwnerValue.LookupId = $userOwnerEnsured.id
            $global:ownerValueCollection = [Microsoft.SharePoint.Client.FieldUserValue[]]$spOwnerValue
            ####return $ownerValueCollection
        }
}

####Called by iterateEverything. Gets Item with CAML query.
function getSubItem($theItem, $targetLocation){
    write-host "Doing CAML Query for Folder" -ForegroundColor Cyan
    ####Get Dir of Item
    $sourceRefForCAML = $theItem
    write-host "THE ITEM" $sourceRefForCAML -ForegroundColor Green
    $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $qry.viewXML = @"
<View Scope="RecursiveAll">
    <Query>
		<Where>
			<Eq>
				<FieldRef Name='FileRef' /><Value Type='Text'>$sourceRefForCAML</Value>
			</Eq>
		</Where>
    </Query>
</View>
"@

    if($targetLocation -eq "Source"){
        [Microsoft.SharePoint.Client.ListItemCollection]$sourceItemFolder = $list.GetItems($qry)
        $Context.Load($sourceItemFolder)
        $Context.ExecuteQuery()
    }
    if($targetLocation -eq "Dest"){
        [Microsoft.SharePoint.Client.ListItemCollection]$sourceItemFolder = $destList.GetItems($qry)
        $destContext.Load($sourceItemFolder)
        $destContext.ExecuteQuery()
    }


    #$itemSingled = $sourceItemFolder
    foreach($folder in $sourceItemFolder){
        if($folder.id){
            $itemSingled = $folder
            write-host "Item is Singled!!!!!!"$itemSingled["FileRef"]
        }
    }

   write-host "SourceQ!!!!!!!!!!"$itemSingled["FileRef"]
    return $itemSingled
}

####Called by iterateEverything. Copies Metadata
function copyMetadata($targetitem, $sourceItem, $fieldsToUpdate, $whatIsCreated){
    write-host "Copying MetaData"$targetitem["FSObjType"] -ForegroundColor Cyan

    ####Approve Items in List. Files are approved below with a new update because we can't approve and update other properties at the same time.
    if($WPFcheckBox_Approve.isChecked -eq $true){
        if($list.BaseType -eq "GenericList"){
            approveContent -item $targetitem -type "List"
        }
    }
    if($WPFtextBox_URL.text -ne $WPFtextBox_URL_Dest.text){
        
        ####Get all Lookups and Try to Copy them as Text 
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -ne "Lookup"))}){
            if(($field.DestinationType -eq "Single line of text") -or ($field.DestinationType -eq "Multiple lines of text")){
                $lookupArray=$sourceItem[$field.SourceName].LookupValue   
                if($lookupArray){
                    write-host "Joining LookupString"
                    $lookupToString = [system.String]::Join(", ",$lookupArray)
                }
                Write-Host "Look Srt" $lookupToString  -ForegroundColor Magenta 
                $targetitem[$field.DestinationName]=$lookupToString
            }
        }

        ####Get all Lookups and Try to Copy them as Text 
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -eq "Lookup"))}){
            #foreach ($fieldValue in $sourceItem[$field.SourceName]){
            #    [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupFieldsArray += $fieldValue
            #}
            #$targetitem[$field.DestinationName] = $lookupFieldsArray
            $targetitem[$field.DestinationName] = $sourceItem[$field.SourceName]
        }

        ####If Fields are not User Fields or Lookup Fields just copy
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -ne "Person or Group") -and ($_.SourceType -ne "Lookup") -and ($_.SourceType -ne "Choice") -and ($_.SourceType -ne "Managed Metadata"))}){
            $targetitem[$field.DestinationName] = $sourceItem[$field.SourceName]
            #$targetitem.update()
        }

        ####if Fields are User Fields get them from Ensured User Array and create Fields User Array to Pass to the Destination Field
        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Person or Group"}){ 
                            $filterItOut = $allEnsuredArray | where {$_.BelongingTo -eq $field.SourceName}
                            if($filterItOut){
                                write-host "Updating User Field with users other then Owner"
                                $userValueCollection = [Microsoft.SharePoint.Client.FieldUserValue[]]$filterItOut
                                $targetitem[$field.DestinationName]=$userValueCollection
                            } else { 
                                ####Comented it out! Now not updating Field with Owner if there is not an ensured User for the field. Otherwise even fields that don't originally have users filled in get the Owner.  
                                #write-host "Updating User Field with Owner"                              
                                #$targetitem[$field.DestinationName] = $global:ownerValueCollection
                            }
                            #$targetitem.update()
        } 

        ####If Fields are Choice
        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Choice"}){
            $choiceArray = @()            
            foreach ($choice in $sourceItem[$field.SourceName]){
                write-host "CHOICE !!!" $choice
                $choiceArray += $choice  
            }       
            $targetitem[$field.DestinationName] = $choiceArray
        }

        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -eq "Managed Metadata"))}){
            $metaValue = ""
            foreach($meta in $sourceItem[$field.SourceName]){
                write-host "Managed Meta" -ForegroundColor Green
                $labelMeta = $meta.Label
                $guidMeta = $meta.TermGuid
                $idMeta = $meta.$item.WssId
                if(!($metaValue)){
                    $metaValue = "-1;#$labelMeta|$guidMeta"
                } else {
                    $metaValue += ";#-1;#$labelMeta|$guidMeta"
                }
            }         
            $targetitem[$field.DestinationName] = $metaValue
        }

        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -ne "Managed Metadata"))}){
            if(($field.DestinationType -eq "Single line of text") -or ($field.DestinationType -eq "Multiple lines of text")){
                $metaArray = @()
                if($sourceItem[$field.SourceName]){
                    foreach($meta in $sourceItem[$field.SourceName]){
                        $metaArray += $meta.Label   
                    }
                    $metaString = [system.String]::Join(", ",$metaArray)
                    $targetitem[$field.DestinationName] = $metaString
                }
            }
        }



    } else {
        foreach ($field in $fieldsToUpdate| where {($_.SourceType -ne "Managed Metadata")}){
                $targetitem[$field.DestinationName] = $sourceItem[$field.SourceName]
        }

        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Managed Metadata"}){
            $metaValue = ""
            foreach($meta in $sourceItem[$field.SourceName]){
                write-host "Managed Meta Single!!!!" -ForegroundColor Green
                $labelMeta = $meta.Label
                $guidMeta = $meta.TermGuid
                $idMeta = $meta.$item.WssId
                if(!($metaValue)){
                    $metaValue = "-1;#$labelMeta|$guidMeta"
                } else {
                    $metaValue += ";#-1;#$labelMeta|$guidMeta"
                }
            }         
            $targetitem[$field.DestinationName] = $metaValue
        }
    }
    ####TRYTRY
    $targetitem["ContentTypeId"] = $sourceItem["ContentTypeId"]

    ####Finaly Update
    $targetitem.update()

    ####Approve if Library
    if($WPFcheckBox_Approve.isChecked -eq $true){
        if($list.BaseType -eq "DocumentLibrary"){     
            approveContent -item $targetitem -type "Library"     
        }
    }

}

####Called by iterateEverything. Creates folder in Library and calls Copy Metadata
function createFolderLibrary ($destinationFileURL, $sourceItemSub){

                        write-host "Destination "$destinationFileURL -ForegroundColor DarkCyan

                        ####Upload
                        $upload = $destList.RootFolder.folders.Add($destinationFileURL)
                        
                        write-host "Copying Metadata" -ForegroundColor DarkCyan
                        if($global:destSPTypeCheck -eq "False 2010"){
                            write-host "SP2010!"
                            $targetMeta = getSubItem -theItem $destinationFileURL -targetLocation "Dest"
                        } else {
                            ####Get Target Item Fields
                            $targetMeta = $upload.ListItemAllFields
                        }
                        copyMetadata -targetitem $targetMeta -sourceItem $sourceItemSub -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                        ####Commit
                        $destContext.Load($upload)
                        $destContext.ExecuteQuery()

                        write-host "Folder Copied" -ForegroundColor Green
                        progressbar -state "Plus" 
                        }

####Called by iterateEverything. Creates folder in List and calls Copy Metadata.
function createFolderList ($destinationFileURL, $sourceItemSub){
                        write-host "Destination "$destinationFileURL -ForegroundColor DarkCyan

                        $destinationFolderURLTrim = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('/'))
                        $destinationFileNameTrim = $destinationFileURL.split("/")[-1]
                        ####Upload
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation

                        write-host "Destination Folder Name"$destinationFolderURLTrim
                        $FileCreationInfo.FolderURL = $destinationFolderURLTrim

                        write-host "Destination leaf Name"$destinationFileNameTrim
                        $FileCreationInfo.LeafName = $destinationFileNameTrim
                        $FileCreationInfo.UnderlyingObjectType = [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder
                        $Upload = $destList.AddItem($FileCreationInfo)

                        write-host "Copying Metadata" -ForegroundColor DarkCyan

                        copyMetadata -targetitem $Upload -sourceItem $sourceItemSub -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                        ####Commit
                        $destContext.Load($Upload)
                        $destContext.ExecuteQuery()
                     
                        write-host "Folder Copied" -ForegroundColor Cyan
                        progressbar -state "Plus" 
                        }

####Called by iterateEverything. Creates File in a Library and calls Copy Metadata. Has an option for 2010 since files are created diferently in 2010.
function createFile {

                            ####Open source
                            $sourceFile = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $sourceItem["FileRef"])

                            write-host "Destination"$destinationFileURL -ForegroundColor DarkCyan
                            ####Create IO Stream from Net Connection Stream
                            $memoryStream = New-Object System.IO.MemoryStream
                            $sourceFile.stream.copyTo($memoryStream)
                            $memoryStream.Seek(0, [System.IO.SeekOrigin]::Begin)

                            ####If SharePoint is everything else but 2010
                            if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                                ####Create Creation Info and Upload
                                $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                                $FileCreationInfo.Overwrite = $true
                                $FileCreationInfo.ContentStream = $memoryStream
                                $FileCreationInfo.URL = $destinationFileURL
                                $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                            }

                            ####If SharePoint is 2010
                            if($global:destSPTypeCheck -eq "False 2010"){

                                ####Convert Stream to Byte Array
                                [byte[]]$fileCreationContent = $memoryStream.ToArray()

                                ####Create Creation Info and Upload
                                $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                                $FileCreationInfo.Overwrite = $true
                                ####$FileCreationInfo.ContentStream = $memoryStream
                                $FileCreationInfo.Content = $fileCreationContent
                                $FileCreationInfo.URL = $destinationFileURL
                                $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                            }

                            ####Get Target Item Fields
                            
                            write-host "Copying Metadata" -ForegroundColor DarkCyan
                            $targetMeta = $Upload.ListItemAllFields

                            #Handle additional Version created if Versioning is Turned On
                            $modifyVersioning = $false
                            $minorVer = $false
                            if($destList.EnableVersioning -eq $true){
                                if($destList.EnableMinorVersions -eq $true){
                                    $minorVer = $true
                                }
                            write-host "MODIFIYNG VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $false
                                $destList.update()
                                $destContext.ExecuteQuery()
                                $modifyVersioning = $true
                            }

                            copyMetadata -targetitem $targetMeta -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "File"

                            if($modifyVersioning -eq $true){
                                write-host "TURNING ON VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $true

                                if($minorVer -eq $true){
                                    $destList.EnableMinorVersions = $true
                                }
                                $destList.update()
                            }
                            ####Will "Approve" but "Modified" field data will ber lost
                            #approveContent -item $Upload -type "File"

                            ####Commit
                            $destContext.Load($Upload)
                            $destContext.ExecuteQuery()

                            write-host "File Copied" -ForegroundColor Green
                            progressbar -state "Plus" 
                        }

####Called by iterateEverything.Creates File in a Library and calls Copy Metadata. Copies Attachments if present.
function createItem {
                            ####Create Creation Info and Create
                            $FileCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                            $FileCreationInfo.FolderURL = $destinationFolderURL
                            $FileCreationInfo.LeafName = $destinationFileName
                            $Upload = $destList.AddItem($FileCreationInfo)

                            #Handle additional Version created if Versioning is Turned On
                            $modifyVersioning = $false
                            $minorVer = $false
                            if($destList.EnableVersioning -eq $true){
                                if($destList.EnableMinorVersions -eq $true){
                                    $minorVer = $true
                                }
                            write-host "MODIFIYNG VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $false
                                $destList.update()
                                $destContext.ExecuteQuery()
                                $modifyVersioning = $true
                            }

                            write-host "Copying Metadata" -ForegroundColor DarkCyan
                            copyMetadata -targetitem $Upload -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Item"

                            if($modifyVersioning -eq $true){
                                write-host "TURNING ON  VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $true

                                if($minorVer -eq $true){
                                    $destList.EnableMinorVersions = $true
                                }
                                $destList.update()
                            }

                            ####Commit
                            $destContext.Load($Upload)
                            $destContext.ExecuteQuery()

                            if($sourceItem["Attachments"] -eq $true){

                            write-host "List Rel URL:"$listRelativeURL -ForegroundColor Green
                            write-host "Source ID:"$sourceItem["ID"] -ForegroundColor Green

                            $sourceAttachmentFolderURL = $listRelativeURL+"/Attachments/"+$sourceItem["ID"]

                            write-host "Full URL"$sourceAttachmentFolderURL -ForegroundColor Cyan

                            $sourceAttachmentFolder = $Context.Web.GetFolderByServerRelativeUrl("$sourceAttachmentFolderURL")
                            [Microsoft.SharePoint.Client.FileCollection]$sourceAttachmentFiles =  $sourceAttachmentFolder.files

                            ####Commit
                            $Context.Load($sourceAttachmentFiles)
                            $Context.ExecuteQuery()

                            foreach($file in $sourceAttachmentFiles){
                                ####Open source
                                write-host "FileURL:"$file.ServerRelativeUrl -ForegroundColor Green
                                $sourceAttFile = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $file.ServerRelativeUrl) 

                                ####Create IO Stream from Net Connection Stream
                                $memoryStream = New-Object System.IO.MemoryStream
                                $sourceAttFile.stream.copyTo($memoryStream)
                                $memoryStream.Seek(0, [System.IO.SeekOrigin]::Begin)

                                ####If SharePoint is everything else but 2010
                                if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                                    ####Get File Name
                                    $fileRelativeURL = $file.ServerRelativeUrl
                                    $fileTitleRegex =  $fileRelativeURL.split("/")[-1]
                                    write-host "File Title" $fileTitleRegex -ForegroundColor Cyan

                                    ####Create Attachment Creation Info
                                    $attInfo = New-Object Microsoft.SharePoint.Client.AttachmentCreationInformation
                                    $attInfo.FileName = $fileTitleRegex
                                    $attInfo.ContentStream = $memoryStream
                                    $uploadAtt = $Upload.AttachmentFiles.Add($attInfo)
                                    
                                    ####Commit
                                    $destContext.Load($uploadAtt)
                                    $destContext.ExecuteQuery()
                                }

                                ####If SharePoint is 2010
                                if($global:destSPTypeCheck -eq "False 2010"){
                                    [byte[]]$fileCreationContent = $memoryStream.ToArray()

                                    $fileRelativeURL = $file.ServerRelativeUrl
                                    $fileTitleRegex =  $fileRelativeURL.split("/")[-1]
                                    $destinationAttURL = $destListRelativeURL+"/Attachments/"+$Upload["ID"]+"/"+$fileTitleRegex                                   

                                    if($WPFtextBox_URL_Dest.text[-1] -ne "/"){
                                        $fullURLDest = $WPFtextBox_URL_Dest.text+"/"
                                    } else {
                                        $fullURLDest = $WPFtextBox_URL_Dest.text
                                    }

                                    $WebSrvUrl=$fullURLDest+"_vti_bin/lists.asmx"
                                    $userPS = $WPFtextBox_User_Dest.text
                                    $passPS = $WPFtextBox_Pass_Dest.password | ConvertTo-SecureString -AsPlainText -Force
                                    $CredsPS = New-Object System.Management.Automation.PSCredential($userPS,$passPS)

                                    #$Creds = createCredentials -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.password -spType "PSCred"
                                    try {
                                        $proxy=New-WebServiceProxy -Uri $WebSrvUrl -Credential $CredsPS
                                    } catch {
                                        write-host $_.Exception.Message -ForegroundColor Yellow
                                    }
                                    #$proxy.timeout=1000000
                                    $proxy.URL=$WebSrvUrl
                                    $proxy.AddAttachment($destList.Title, $Upload["ID"], $fileTitleRegex, $fileCreationContent)
                                }
                            }
                        }
                            write-host "Item Copied" -ForegroundColor Green
                            progressbar -state "Plus" 
                    }

####CheckIn, Publish, Approve Items/Files. Not Used as of now. Kept for reference in the future.
function approveContent ($item, $type) {
    if ($item["_ModerationStatus"] -ne '0'){
            ####File is not approved, approval process is applied
            Write-Host "File:" $item["FileLeafRef"] "needs approval" -ForegroundColor Cyan

            $item["_ModerationStatus"] = 0
            Write-Host "File Approved" -ForegroundColor Green
    }  

    if($type -eq "Library"){
        $item.update()
    }
}

####Called by iterateEverything, calls ensureUser in a loop for ech fields. Returns all users for all fields for an item in an array with an additional value for which field it belongs to.
function preEnsureUsers ($sourceItemFuncEnsure, $whoIsAsking){
                ####If we are not copying in the same Site Collection
                if($WPFtextBox_URL.text -ne $WPFtextBox_URL_Dest.text){

                    ####Get Fields to update for User Ensuring below
                    $fieldsToUpdate = $WPFlistView_Fields_Final.items

                    ####Prepare all user fields. Get all User Fields, ensure them and get them into array. This is necessary because Context Executes during Metadata copy screw up the Modified, Editor and Author fields.
                    $allEnsuredArray=@()
                    foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Person or Group"}){
                        write-host "Field Entered"$field
                        write-host "Pre-Ensuring User"
                        if($whoIsAsking -eq "iterateEverything"){
                            $userValueCollection = ensureUser -sourceFieldUV $sourceItemFuncEnsure[$field.SourceName]
                        }
                        if($whoIsAsking -eq "iterateFileShareSourceMetaData"){
                        write-host "FUNC ENSURE" $sourceItemFuncEnsure.($field.SourceName) -ForegroundColor Green
                            $userValueCollection = ensureUserMetaDataFile -sourceFieldUV $sourceItemFuncEnsure.($field.SourceName) 
                        }
                        if($userValueCollection){
                            write-host "Adding User to Array"
                            $allEnsuredArray+=$userValueCollection
                        }
                    }
                    return $allEnsuredArray
                }
}

function getFoldersThatShouldBeCreated ($checkFolderTrim){
                             $folderToCreateArray = @()    
                            while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                            $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                            $obj = new-object psobject -Property @{
                                'URL' = $checkFolderTrim
                                'Count' = $folderURLCount
                            }
                                $folderToCreateArray += $obj
                                $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('/'))
                            }
                            $folderToCreateArray = $folderToCreateArray| sort-object Count
                            return $folderToCreateArray
 }
 
function iterateEverything ($whatTo){
            ####Store All created Folders
            $fileDirRefs = @()

            $listRelativeURL = $list.RootFolder.ServerRelativeUrl
            $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl
                       
            ####Add List Root in the created Folders List
            $fileDirRefs += $destList.RootFolder.serverrelativeurl

            ####Handle User Fields
            write-host "Pre-Ensuring Owner"

            ####Ensure Owner. If users in user fields cannot be ensured the User wich is used for Destination logging will be used.
            try {
                ensureOwner
            } catch {
                logEverything -relatedItemURL "Ensure Owner" -exceptionToLog $_.exception.message
            }
            ####Start Iteration
            foreach ($sourceItem in $whatTo){

                ####if Folder
                if($sourceItem["FSObjType"] -eq 1){
                    $WPFlabel_Status.Content = "Status: Copying Folder "+$sourceItem["FileRef"]
                    $Form.Dispatcher.Invoke("Background", [action]{})

                    write-host "Folder to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan
              
                    ####Construct File Name and URL
                    $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
                    $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $destListRelativeURL
                    $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('/'))

                    ####if Folder in Document Library
                    if($list.BaseType -eq "DocumentLibrary"){
                        if($fileDirRefs -contains $destinationFolderURL){
                            
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"

                            try {
                                createFolderLibrary -destinationFileURL $destinationFileURL -sourceItemSub $sourceItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                            $fileDirRefs += $destinationFileURL
                        } else {
                            $checkFolderTrim = $destinationFolderURL
                            $folderToCreateArray = getFoldersThatShouldBeCreated -checkFolderTrim $checkFolderTrim

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace $destListRelativeURL, $listRelativeURL
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"
                                $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemSub -whoIsAsking "iterateEverything"
                                try {
                                    createFolderLibrary -destinationFileURL $one.URL -sourceItemSub $sourceItemSub
                                } catch {
                                    logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                                }
                                $fileDirRefs += $one.URL                        
                            }
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                createFolderLibrary -destinationFileURL $destinationFileURL -sourceItemSub $sourceItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                            $fileDirRefs += $destinationFileURL
                        }
                    }

                    ####if Folder in List 
                    if($list.BaseType -eq "GenericList"){
                    $WPFlabel_Status.Content = "Status: Copying Folder "+$sourceItem["FileRef"]
                    $Form.Dispatcher.Invoke("Background", [action]{})
                        if($fileDirRefs -contains $destinationFolderURL){

                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"

                            try {
                                createFolderList -destinationFileURL $destinationFileURL -sourceItemSub $sourceItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                            $fileDirRefs += $destinationFileURL
                        } else {
                            $checkFolderTrim = $destinationFolderURL
                            
                            $folderToCreateArray = getFoldersThatShouldBeCreated -checkFolderTrim $checkFolderTrim

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace $destListRelativeURL, $listRelativeURL
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"

                                $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemSub -whoIsAsking "iterateEverything"
                                try {
                                    createFolderList -destinationFileURL $one.URL -sourceItemSub $sourceItemSub
                                } catch {
                                    logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                                }
                                $fileDirRefs += $one.URL                        
                            }
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                createFolderList -destinationFileURL $destinationFileURL -sourceItemSub $sourceItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                            $fileDirRefs += $destinationFolderURL
                        }
                    }
                }

                ####if Item
                if($sourceItem["FSObjType"] -eq 0){

                    ####If File in Document Library (Copy File)
                    if($list.BaseType -eq "DocumentLibrary"){
                        $WPFlabel_Status.Content = "Status: Copying File "+$sourceItem["FileRef"]
                        $Form.Dispatcher.Invoke("Background", [action]{})

                        write-host "File to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan

                        ####Construct File Name and URL
                        $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
                        $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $destListRelativeURL
                        $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('/'))

                        if($fileDirRefs -contains $destinationFolderURL){

                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                CreateFile 
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                        } else {
                               Write-Host "The Trim" $destinationFolderURL

                            $checkFolderTrim = $destinationFolderURL
                            
                            $folderToCreateArray = getFoldersThatShouldBeCreated -checkFolderTrim $checkFolderTrim

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace $destListRelativeURL, $listRelativeURL
                                write-host "Source URL of the Trim" $sourceURLOfTheTrim -ForegroundColor Green
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"

                                $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemSub -whoIsAsking "iterateEverything"
                                try {
                                    createFolderLibrary -destinationFileURL $one.URL -sourceItemSub $sourceItemSub
                                } catch {
                                    logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                                }                                
                                $fileDirRefs += $one.URL                        
                            }
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                createFile
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                        }
                    }

                    ####If Item in List (Create Item)
                    if($list.BaseType -eq "GenericList"){
                        $WPFlabel_Status.Content = "Status: Copying Item "+$sourceItem["FileRef"]
                        $Form.Dispatcher.Invoke("Background", [action]{})

                        write-host "File to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan

                        ####Construct File Name and URL
                        $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
                        $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $destListRelativeURL
                        $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('/'))
                        if($fileDirRefs -contains $destinationFolderURL){

                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                createItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                        } else {
                            $checkFolderTrim = $destinationFolderURL
                            
                            $folderToCreateArray = getFoldersThatShouldBeCreated -checkFolderTrim $checkFolderTrim

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace $destListRelativeURL, $listRelativeURL
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"
                                $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemSub -whoIsAsking "iterateEverything"
                                try {
                                    createFolderList  -destinationFileURL $one.URL -sourceItemSub $sourceItemSub
                                } catch {
                                    logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                                }
                                $fileDirRefs += $one.URL                        
                            }
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItem -whoIsAsking "iterateEverything"
                            try {
                                createItem
                            } catch {
                                logEverything -relatedItemURL $destinationFileURL -exceptionToLog $_.exception.message
                            }
                        }
                    }
                }
            }
        }

####used to copy Full Library/List. Calls iterateEverything.
function copyLibrary ($siteURL,$docLibName,$sourceSPType,$destSiteURL,$destDocLibName,$destSPType){
    ####Check if Lists are Input
    if((!($docLibName)) -or (!($destDocLibName))){
        $WPFlabel_Status.Content = "Status: Lists/Libraries are not selected! Select Lists/Libraries and try again!"
        $Form.Dispatcher.Invoke("Background", [action]{})
    } else {
        ####Get Source List
        $web = $Context.Web
        $List = $Context.Web.Lists.GetByTitle($DocLibName)
        $Context.Load($Web)
        $Context.Load($List)
        $context.Load($List.RootFolder)

        ####Get Destination List
        $destList = $destContext.Web.Lists.GetByTitle($destDocLibName)
        $destWeb = $destContext.Web
        $destContext.Load($destList)
        $destContext.Load($destWeb)
        $destContext.Load($destList.RootFolder)

        ####Check if Applying Filter and if not go on and copy all


        if ($global:copyMode -eq "NoFilter"){
            ####Create CAML Query
            $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        }

        if ($global:copyMode -eq "Filter"){
            ####Create CAML Query
            $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
            $camlQueryConstruct = createCAMLFilter
            $qryItems.viewXML = $camlQueryConstruct
        }
        ####Load
        [Microsoft.SharePoint.Client.ListItemCollection]$items = $list.GetItems($qryItems)
        $Context.Load($items)

        ####Commit
        try{
            $Context.ExecuteQuery()
        } catch {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button

            $exceptionMessage = $_.exception.message
            $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

            [System.Windows.MessageBox]::Show($message,"Warning", $Button)

            $WPFlabel_Status.Content = "Status: Idle"
        }
        try{
            $destContext.ExecuteQuery()
        } catch {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button

            $exceptionMessage = $_.exception.message
            $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

            [System.Windows.MessageBox]::Show($message,"Warning", $Button)

            $WPFlabel_Status.Content = "Status: Idle"
        }
        ####Add a Property reflecting URL deepness to avoid creating nested Items and Folders before the nesting Folder is created 
        foreach ($item in $items){
            $urlCount = ($item["FileRef"].ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            $item | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount
        }

        ####Sort items by URL Deepness. Create items with the least amaount of "/" first
        $itemsSorted = $items | sort-object URLLengthCustom

        ####Get Lists Relative URLs
        $listRelativeURL = $list.RootFolder.ServerRelativeUrl
        $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl

        foreach($item in $itemsSorted){
        write-host $item["FileRef"] -ForegroundColor Cyan
        }
        progressbar -state "Start" -all $itemsSorted.count
        iterateEverything -whatTo $itemsSorted

        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
        progressbar -state "Stop"

        #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
        #$global:errorsAll | Out-File -FilePath $errorLogPath
        #Invoke-Item  $errorLogPath
        #$global:errorsAll | Out-GridView
    }
}

################################################
## Button Actions
################################################


################################################
## New Buttons
################################################


###Previous Button
$global:prevClicked = $false
$WPFbutton_Prev.Add_MouseDown({
    if($global:prevClickable -eq $true){
        $WPFbutton_Prev.source = "$dir\Button_Prev_S_Click.png"
        $global:prevClicked = $true
    }
})
$WPFbutton_Prev.Add_MouseLeave({
    if($global:prevClickable -eq $true){
        if($global:prevClicked -eq $true){
            $WPFbutton_Prev.source = "$dir\Button_Prev_S.png"
        }
    }
})
$WPFbutton_Prev.Add_MouseUp({
    if($global:prevClickable -eq $true){
        $WPFbutton_Prev.source = "$dir\Button_Prev_S.png"
        $global:prevClicked = $false
        if(($global:preciseLocation -eq "Destination") -or ($global:preciseLocation -eq "destPre")){
            opacityAnimation -grid $WPFdestSiteCol -action "Close"  
            $WPFtop_One.source = "$dir\Top_One.png" 
            opacityAnimation -grid $WPFSourceSiteCol -action "Open"
            $global:preciseLocation = "Source"
            $WPFbutton_Prev.source = "$dir\Button_Prev_S_Click.png" 
            $global:prevClickable = $false                     
        }
        if(($global:preciseLocation -eq "Metadata")){
            opacityAnimation -grid $WPFfield_Controls -action "Close"  
            $WPFtop_One.source = "$dir\Top_Two.png"
            opacityAnimation -grid $WPFdestSiteCol -action "Open"  
 
            $global:preciseLocation = "Destination"    
        }
        if(($global:preciseLocation -eq "Copy")){
            $global:preciseLocation = "Metadata" 
            opacityAnimation -grid $WPFFilterChoice -action "Close"
            $WPFtop_One.source = "$dir\Top_Three.png"  
            opacityAnimation -grid $WPFfield_Controls -action "Open"

            ####Nullify Filters
            $WPFtextBox_Filter_Value.Text = ""
            $WPFcomboBox_Field_To_Filter.itemsSource = @()
            $WPFcomboBox_Condition.items.clear()
            $WPFlistView_Filters.items.clear()

            ####Nullify Selected Items for Copy
            $addItemsForViewArray = @()
            $WPFlistView_Items.itemsSource = @()
            $global:itemsToCopy = @()
            $WPFradioButton_Browser.isChecked = $false
            $WPFradioButton_No_Filter.isChecked = $false
            $WPFradioButton_Filter.isChecked = $false

            ####Hide Sections
            $WPFbrowserGrid.Visibility = "Hidden"
            $WPFallGrid.Visibility = "Hidden"
            $WPFgrid_Filters.Visibility = "Hidden"
        }
    }
})

###Next Button
$global:nextClicked = $false
$WPFbutton_Next.Add_MouseDown({
    if($global:nextClickable -eq $true){
        $WPFbutton_Next.source = "$dir\Button_Next_S_Click.png"
        $global:nextClicked = $true
    }
})
$WPFbutton_Next.Add_MouseLeave({
    if($global:nextClickable -eq $true){
        if( $global:nextClicked -eq $true){
            $WPFbutton_Next.source = "$dir\Button_Next_S.png"
        }
    }
})
$WPFbutton_Next.Add_MouseUp({

    if($global:nextClickable -eq $true){
        $WPFbutton_Next.source = "$dir\Button_Next_S.png"
        $global:nextClicked = $false
        if($global:preciseLocation -eq "Copy") {
        if(($WPFradioButton_No_Filter.isChecked -eq $true) -or ($WPFradioButton_Filter.isChecked -eq $true) -or ($WPFradioButton_Browser.isChecked -eq $true)){
                 if((($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta")) -or ($global:destSPTypeCheck -eq "File")){
                    copyFunctionFileShare
                } else {
                    copyFunction
                } 
            } else {
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button
                [System.Windows.MessageBox]::Show("Select filter mode first.","Warning", $Button) 
            }    
        }

        if($global:preciseLocation -eq "Metadata"){
            write-host $global:copyMode
            if(($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta")) {
                if($global:sourceSPTypeCheck -eq "Meta"){
                    $WPFradioButton_Browser.isEnabled = $true
                    $WPFradioButton_Filter.isEnabled = $false
                } else {
                    $WPFradioButton_Browser.isEnabled = $false
                    $WPFradioButton_Filter.isEnabled = $false        
                }
        } else {
            $WPFradioButton_Browser.isEnabled = $true
            $WPFradioButton_Filter.isEnabled = $true    
        }


        $WPFradioButton_No_Filter.isChecked = $false
        $WPFradioButton_Filter.isChecked = $false
        $WPFradioButton_Browser.isChecked = $false

        $global:preciseLocation = "Copy" 
        opacityAnimation -grid $WPFfield_Controls -action "Close"
        $WPFtop_One.source = "$dir\Top_Four.png"
        opacityAnimation -grid $WPFfilterChoice -action "Open"
        }

        if($global:preciseLocation -eq "Destination"){
            if(($WPFlistView.SelectedItem.Title) -and ($WPFlistView_Dest.SelectedItem.Title)){

                if(($WPFlistView.SelectedItem.Type -eq $WPFlistView_Dest.SelectedItem.Type) -or (($WPFlistView.SelectedItem.Type -eq "FileShare") -and ($WPFlistView_Dest.SelectedItem.Type -eq "DocumentLibrary")) -or (($WPFlistView.SelectedItem.Type -eq "DocumentLibrary") -and ($WPFlistView_Dest.SelectedItem.Type -eq "FileShare"))){
                    ####Remember Input Fields for Next Program opening
                    Remember

                    ####Show/Hide bitton to manage Lookup Fields auto
                    if($WPFtextBox_URL.text -ne $WPFtextBox_URL_Dest.text){
                        write-host "URLs not equal"
                    if($global:sourceSPTypeCheck -eq "Meta"){
                        $WPFcheckBox_Manage_Lookup_Auto.content = "Restoring to Different Site Collection than Original"
                        $WPFcheckBox_Manage_Lookup_Auto.isEnabled = $true 
                        $WPFcheckBox_Manage_Meta_Auto.content = "Restoring to a Different Farm than the Original"
                        $WPFcheckBox_Manage_Meta_Auto.isEnabled = $true 
                        $WPFcheckBox_Approve.isEnabled = $true 
                        write-host "Source Meta"
                    } elseif(($global:sourceSPTypeCheck -eq "File") -or ($global:destSPTypeCheck -eq "File")) {
                        write-host "Source Dest File"
                        $WPFcheckBox_Manage_Lookup_Auto.isEnabled = $false  
                        $WPFcheckBox_Manage_Meta_Auto.isEnabled = $false   
                        $WPFcheckBox_Approve.isEnabled = $true  
                        if($global:destSPTypeCheck -eq "File"){
                        write-host "Additional Dest File"
                            $WPFcheckBox_Approve.isEnabled = $false                
                        }
                    } else {
                        write-host "Else on All"
                        $WPFcheckBox_Manage_Lookup_Auto.isEnabled = $true 
                        $WPFcheckBox_Manage_Meta_Auto.isEnabled = $true 
                        $WPFcheckBox_Approve.isEnabled = $true 
                    }

                    } else {
                            $WPFcheckBox_Manage_Lookup_Auto.isEnabled = $false      
                            $WPFcheckBox_Manage_Meta_Auto.isEnabled = $false 
                            $WPFcheckBox_Approve.isEnabled = $true 
                    }

                    $WPFlabel_Status.Content = "Status: Getting Lists/Libraries"
                    opacityAnimation -grid $WPFField_Controls -action "Pre"
                    opacityAnimation -grid $WPFField_Advanced -action "Pre"
                    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
                    [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
                    progressbar -state "Start" -all 2
                    write-host "Source Type" $global:sourceSPTypeCheck
                    if(($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta")){
                        if($global:sourceSPTypeCheck -eq "File"){
                            $obj = new-object psobject -Property @{
                                'Name' = "No Available Fields"
                                'Type' = "FileShare"
                                'BaseType' = ""
                            }
                            $fileShareFieldsArray = @()
                            $fileShareFieldsArray+=$obj
                            $WPFlistView_Fields_Source.ItemsSource = $fileShareFieldsArray
                        } else {
                            $metadataFileFields = @()
                            $global:metadataAll | get-member -type NoteProperty | foreach-object {
                            $obj = new-object psobject -Property @{
                                 'Name' = $_.Name
                                 'Type' = $global:metadataAll.($_.Name)[0]
                                 'BaseType'= "NULL"
                             }
                            $metadataFileFields+=$obj
                            }
                        $WPFlistView_Fields_Source.ItemsSource = $metadataFileFields 
                        }
                    }else{
                        $sourceFieldsArray = getFields -User $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -listView $WPFlistView_Fields_Source -spType $global:sourceSPTypeCheck -SourceOrTarget "Source"
                    }
                    progressbar -state "Plus" 
                    if($global:destSPTypeCheck -eq "File"){
                        $WPFlistView_Fields_Dest.ItemsSource = $sourceFieldsArray
                    }else{
                        $destFieldsArray = getFields -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.password -SiteURL $WPFtextBox_URL_Dest.text -docLibName $destListTitle -listView $WPFlistView_Fields_Dest -spType $global:destSPTypeCheck -SourceOrTarget "Target"
                    }
                    progressbar -state "Plus" 
                    $WPFlabel_Status.Content = "Status: Idle"

                    $WPFlistView_Fields_Final.ItemsSource = @()
                    preAddFields
                    $WPFField_Advanced.Visibility = "Hidden"
                    opacityAnimation -grid $WPFdestSiteCol -action "Close"
                    opacityAnimation -grid $WPFfield_Controls -action "Open"
                    $global:preciseLocation = "Metadata"
                    $WPFtop_One.source = "$dir\Top_Three.png"
                    progressbar -state "Stop" 
                } else {
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button
                    [System.Windows.MessageBox]::Show("This combination of Types is not supported.","Warning", $Button)   
                }
            } else {
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button
                [System.Windows.MessageBox]::Show("Login and select Library/List first.","Warning", $Button)
            }                
        }       

        if($global:preciseLocation -eq "Source"){
            if($WPFlistView.SelectedItem.Title) {
                opacityAnimation -grid $WPFSourceSiteCol -action "Close"
                $WPFtop_One.source = "$dir\Top_Two.png"
                opacityAnimation -grid $WPFdestSiteCol -action "Open"
                $global:preciseLocation = "Destination" 
                $global:prevClickable = $true
                $WPFbutton_Prev.source = "$dir\Button_Prev_S.png"              
            } else {
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button
                [System.Windows.MessageBox]::Show("Select Source List/Library first.","Warning", $Button)
                $WPFradioButton_SourceSPVer_FileShare.isChecked = $false            
            }     
        }
    }
})

###Close Button
$global:closeClicked = $false
$WPFClose_Adv.Add_MouseDown({
        $WPFClose_Adv.source = "$dir\Close_Clicked.png"
        $global:closeClicked = $true
        write-host "Clicked"
})
$WPFClose_Adv.Add_MouseLeave({
        if($global:closeClicked -eq $true){
            $WPFClose_Adv.source = "$dir\Close.png"
        }
})
$WPFClose_Adv.Add_MouseUp({
        $global:closeClicked = $false
        $WPFClose_Adv.source = "$dir\Close.png"
        opacityAnimation -grid $WPFField_Advanced -action "Close"
})


####Get Source Lists Button
$WPFbutton_GetSource.Add_Click({
    ####If SharePoint
    if($WPFradioButton_SourceSPVer_Premise.isChecked -eq $True){
        if(($WPFtextBox_User.text) -and ($WPFtextBox_Pass.password) -and ($WPFtextBox_URL.text)){
            ####If SharePoint Online
            if($WPFtextBox_URL -like "*sharepoint.com*"){
                ####Making Source List Grid Visible with 0 opacity so when it is filled the columns pickup the right width
                opacityAnimation -grid $WPFsourceListsAll -action "Pre"
                ####Get Lists
                $listGetSuccess = getLists -User $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -listView $WPFlistView -spType "True" -SourceOrTarget "Source"
                if($listGetSuccess -eq $true){
                    $global:sourceSPTypeCheck = "True"
                    $WPFlabel_Status.Content = "Status: Idle"
                    $infoSourceColType = "SharePoint Online"
                    $infoSourceColURL = $WPFtextBox_URL.text 
                    $WPFsourceConInfo.Text = @"
Connected to Site Collection: $infoSourceColURL
Type: $infoSourceColType
"@

                    opacityAnimation -grid $WPFLoginSource -action "Close"
                    opacityAnimation -grid $WPFsourceListsAll -action "Open"
                    $global:preciseLocation = "Source"
                    $WPFbutton_Next.source = "$dir\Button_Next_S.png"
                    $global:nextClickable = $true    
                }  else {
                    $WPFsourceListsAll.Visibility = "Hidden"
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button
                    [System.Windows.MessageBox]::Show("Error while connecting. Check connection info.","Warning", $Button)
                }
            ####If not SharePoint Online
            } else {
                ####Making Source List Grid Visible with 0 opacity so when it is populated the columns pickup the right width
                opacityAnimation -grid $WPFsourceListsAll -action "Pre"
                ####Get Lists
                $listGetSuccess = getLists -User $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -listView $WPFlistView -spType "False" -SourceOrTarget "Source"
                if($listGetSuccess -eq $true){
                    ####Get SharePoint Version
                    $spVersion = $context.ServerVersion.ToString()
                    $spVersion = $spVersion.split(".")[0]
                    if($spVersion -eq "14"){
                        $infoSourceColType = "SharePoint 2010"
                        $global:sourceSPTypeCheck = "False 2010"
                    }
                    if($spVersion -eq "15"){
                        $infoSourceColType = "SharePoint 2013"
                        $global:sourceSPTypeCheck = "False"
                    }
                    if($spVersion -eq "16"){
                        $infoSourceColType = "SharePoint 2016"
                        $global:sourceSPTypeCheck = "False"
                    }
                    $infoSourceColURL = $WPFtextBox_URL.text 
                    $WPFsourceConInfo.Text = @"
Connected to Site Collection: $infoSourceColURL
Type: $infoSourceColType
"@
                    $WPFlabel_Status.Content = "Status: Idle"
                    opacityAnimation -grid $WPFLoginSource -action "Close"
                    opacityAnimation -grid $WPFsourceListsAll -action "Open"
                    $global:preciseLocation = "Source"
                    $WPFbutton_Next.source = "$dir\Button_Next_S.png"
                    $global:nextClickable = $true
                }  else {
                    $WPFsourceListsAll.Visibility = "Hidden"
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button
                    [System.Windows.MessageBox]::Show("Error while connecting. Check connection info.","Warning", $Button)
                }
            }
        } else {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            [System.Windows.MessageBox]::Show("Fill all connection information","Warning", $Button)   
        }
    }
    ####If FileShare
    if($WPFradioButton_SourceSPVer_FileShare.isChecked -eq $True){
        opacityAnimation -grid $WPFsourceListsAll -action "Pre"
        $shouldContinue = getFileShareRoot -Location $WPFlistView -target "Source"
        if($shouldContinue -eq $true){ 
            $WPFlabel_Status.Content = "Status: Idle"
            opacityAnimation -grid $WPFLoginSource -action "Close"
            opacityAnimation -grid $WPFsourceListsAll -action "Open"
            $global:preciseLocation = "Source"
            $WPFbutton_Next.source = "$dir\Button_Next_S.png"
            $global:nextClickable = $true
        } else {
            $WPFsourceListsAll.Visibility = "Hidden"
        }
    }
})

$WPFbutton_GetDest.Add_Click({
    ####If SharePoint
    if($WPFradioButton_DestSPVer_Premise.isChecked -eq $True){
        write-host "User" $WPFtextBox_User_Dest.text
        write-host "Pass" $WPFtextBox_Pass_Dest.password
        write-host "URL" $WPFtextBox_URL_Dest.text

        if(($WPFtextBox_User_Dest.text) -and ($WPFtextBox_Pass_Dest.password) -and ($WPFtextBox_URL_Dest.text)){
            ####If SharePoint Online
            if($WPFtextBox_URL_Dest -like "*sharepoint.com*"){
                ####Making Source List Grid Visible with 0 opacity so when it is filled the columns pickup the right width
                opacityAnimation -grid $WPFDestListsAll -action "Pre"
                ####Get Lists
                $listGetSuccess = getLists -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.password -SiteURL $WPFtextBox_URL_Dest.text -listView $WPFlistView_Dest -spType "True" -SourceOrTarget "Target"
                if($listGetSuccess -eq $true){
                write-host "blaaaaa"
                    $global:destSPTypeCheck = "True"
                    $WPFlabel_Status.Content = "Status: Idle"
                    $infoDestColType = "SharePoint Online"
                    $infoDestColURL = $WPFtextBox_URL_Dest.text 
                    $WPFdestConInfo.Text = @"
Connected to Site Collection: $infoDestColURL
Type: $infoDestColType
"@

                    opacityAnimation -grid $WPFLoginDest -action "Close"
                    opacityAnimation -grid $WPFdestListsAll -action "Open"
                    $global:preciseLocation = "Destination"
                    $WPFbutton_Next.source = "$dir\Button_Next_S.png"
                    $global:nextClickable = $true    
                } else {
                    $WPFdestListsAll.Visibility = "Hidden"
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button
                    [System.Windows.MessageBox]::Show("Error while connecting. Check connection info.","Warning", $Button)
                }
            ####If not SharePoint Online
            } else {
                ####Making Source List Grid Visible with 0 opacity so when it is populated the columns pickup the right width
                opacityAnimation -grid $WPFdestListsAll -action "Pre"
                ####Get Lists
                $listGetSuccess = getLists -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.password -SiteURL $WPFtextBox_URL_Dest.text -listView $WPFlistView_Dest -spType "False" -SourceOrTarget "Target"
                if($listGetSuccess -eq $true){
                    ####Get SharePoint Version
                    $spVersion = $destContext.ServerVersion.ToString()
                    $spVersion = $spVersion.split(".")[0]
                    if($spVersion -eq "14"){
                        $infoDestColType = "SharePoint 2010"
                        $global:destSPTypeCheck = "False 2010"
                    }
                    if($spVersion -eq "15"){
                        $infoDestColType = "SharePoint 2013"
                        $global:destSPTypeCheck = "False"
                    }
                    if($spVersion -eq "16"){
                        $infoDestColType = "SharePoint 2016"
                        $global:destSPTypeCheck = "False"
                    }
                    $infoDestColURL = $WPFtextBox_URL_Dest.text 
                    $WPFdestConInfo.Text = @"
Connected to Site Collection: $infoDestColURL
Type: $infoDestColType
"@
                    $WPFlabel_Status.Content = "Status: Idle"
                    opacityAnimation -grid $WPFLoginDest -action "Close"
                    opacityAnimation -grid $WPFdestListsAll -action "Open"
                    $global:preciseLocation = "Destination"
                    $WPFbutton_Next.source = "$dir\Button_Next_S.png"
                    $global:nextClickable = $true
                } else {
                    $WPFdestListsAll.Visibility = "Hidden"
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button
                    [System.Windows.MessageBox]::Show("Error while connecting. Check connection info.","Warning", $Button)
                }
            }
        } else {
            write-host "Not All Fields are Input"
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            [System.Windows.MessageBox]::Show("Fill all connection information","Warning", $Button)   
        }
    }
    ####If FileShare
    if($WPFradioButton_DestSPVer_FileShare.isChecked -eq $True){
        opacityAnimation -grid $WPFdestListsAll -action "Pre"
        $shouldContinue = getFileShareRoot -Location $WPFlistView_Dest -target "Dest"
        if($shouldContinue -eq $true){ 
            $WPFlabel_Status.Content = "Status: Idle"
            opacityAnimation -grid $WPFLoginDest -action "Close"
            opacityAnimation -grid $WPFdestListsAll -action "Open"
            $global:preciseLocation = "Destination"
            $WPFbutton_Next.source = "$dir\Button_Next_S.png"
            $global:nextClickable = $true
        } else {
            $WPFdestListsAll.Visibility = "Hidden"
        }
    }
})

####Back to Login
$WPFbutton_changeSourceCol.Add_Click({
    opacityAnimation -grid $WPFsourceListsAll -action "Close"
    opacityAnimation -grid $WPFLoginSource -action "Open"
    $WPFbutton_Next.source = "$dir\Button_Next_S_Click.png"
    $global:nextClickable = $false    
})

$WPFbutton_changeDestCol.Add_Click({
    opacityAnimation -grid $WPFdestListsAll -action "Close"
    opacityAnimation -grid $WPFLoginDest -action "Open"
    $WPFbutton_Next.source = "$dir\Button_Next_S_Click.png"
    $global:nextClickable = $false    
})

################################################Source and estination SharePoint Version Radion Buttons
####Source SP Version

$WPFradioButton_SourceSPVer_FileShare.Add_Click({
    $WPFtextBox_URL.isEnabled = $false
    $WPFtextBox_User.isEnabled = $false
    $WPFtextBox_Pass.isEnabled = $false

    $WPFbutton_GetSource.content = "Browse..."

    $global:sourceURLGlobal = $WPFtextBox_URL.text
    $WPFtextBox_URL.text = "FileShare"
})

$WPFradioButton_SourceSPVer_Premise.Add_Click({
    $WPFtextBox_URL.isEnabled = $true
    $WPFtextBox_User.isEnabled = $true
    $WPFtextBox_Pass.isEnabled = $true

    $WPFbutton_GetSource.content = "Connect..."

    $WPFtextBox_URL.text = $global:sourceURLGlobal
})

$WPFradioButton_DestSPVer_FileShare.Add_Click({
    write-host $global:sourceSPTypeCheck
    if(($global:sourceSPTypeCheck -ne "File") -and ($global:sourceSPTypeCheck -ne "Meta")){
        $WPFtextBox_URL_Dest.isEnabled = $false
        $WPFtextBox_User_Dest.isEnabled = $false
        $WPFtextBox_Pass_Dest.isEnabled = $false

        $WPFbutton_GetDest.content = "Browse..."

        $global:destURLGlobal = $WPFtextBox_URL_Dest.text
        $WPFtextBox_URL_Dest.text = "FileShare"

    } else {
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button
        [System.Windows.MessageBox]::Show("Source and Destination can't be both FileShare. Use Windows Explorer for that.","Warning", $Button)
        $WPFradioButton_DestSPVer_FileShare.isChecked = $false
    }
})

$WPFradioButton_DestSPVer_Premise.Add_Click({
    $WPFtextBox_URL_Dest.isEnabled = $true
    $WPFtextBox_User_Dest.isEnabled = $true
    $WPFtextBox_Pass_Dest.isEnabled = $true

    $WPFbutton_GetDest.content = "Connect..."

    $WPFtextBox_URL_Dest.text = $global:destURLGlobal
})
################################################Get Source and Destination Library Lists and Fields Lists Buttons


$WPFbutton_Map_Manually.Add_Click({
    opacityAnimation -grid $WPFField_Advanced -action "Open"

})

function getFileShareRoot ($Location, $target){
        $shouldContinue = $true
        $thanksButNoThanks = $false
        ####Browse For CSV
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $templatepathfull = New-Object System.Windows.Forms.FolderBrowserDialog
        $templatepathfull.ShowDialog() | Out-Null
        $global:fileShareRoot = $templatepathfull.SelectedPath
        if($global:fileShareRoot){
            ####Check for Metadata CVS
            $contentsCheck = get-childitem $global:fileShareRoot
            if (($contentsCheck.name -eq "Metadata.csv") -and $target -eq "Source"){
                $popupMessage = @"
The Application found a Metadata file on the location. 

This probably means that the files in this Directory were previously backed up from SharePoint by the Application.

Would you like to use this file to get the files Metadata?
"@

                $popup = new-object -comobject wscript.shell
                $intAnswer = $popup.popup($popupMessage, 0,"Metadata File",4)
                
                if($intAnswer -eq 6){
                    $shouldContinue = getMetadataFile -Location $WPFlistView -templatepath "$global:fileShareRoot\Metadata.csv"

                    if($shouldContinue -eq $true){
                        $global:sourceSPTypeCheck = "Meta"
                        $WPFsourceConInfo.Text = @"
Source: File Directory with preserved Metadata
"@
                    }
                } else {
                    $thanksButNoThanks = $true
                } 
            } 
            
            if((!($contentsCheck.name -eq "Metadata.csv")) -or ($thanksButNoThanks -eq $true)) {
                $obj = new-object psobject -Property @{
                    'Title' = $global:fileShareRoot
                    'Type' = "FileShare"
                }
                $fileShareRootArray = @()
                $fileShareRootArray+=$obj
                $Location.ItemsSource = $fileShareRootArray
                $Location.selectedItems.Add($Location.ItemsSource[0])
                if($target -eq "Source"){
                    write-host "Source"
                    $global:sourceSPTypeCheck = "File"
                    $WPFsourceConInfo.Text = @"
Source: File Directory
"@
                }
                if($target -eq "Dest"){
                    write-host "Dest"
                    $global:destSPTypeCheck = "File"
                    $WPFdestConInfo.Text = @"
Destination: File Directory
"@
                }
            }
        } else {
            $shouldContinue = $false 
        }
        return $shouldContinue
}

function getMetadataFile ($Location, $templatepath){
        $shouldContinue = $true
        try {
            $global:metadataAll = import-csv $templatepath
        } catch {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            [System.Windows.MessageBox]::Show("Error while importing CSV file","Warning", $Button)
            $shouldContinue = $false
        }
        if(!($global:metadataAll.FileRef[0])){
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            [System.Windows.MessageBox]::Show("Error while importing CSV file","Warning", $Button)
            $shouldContinue = $false       
        }

        $fileShareRootArray = @()
        $obj = new-object psobject -Property @{
            'Title' = $global:metadataAll.FileRef[0]
            'Type' = "DocumentLibrary"
        }
        $fileShareRootArray += $obj
        $Location.ItemsSource = $fileShareRootArray
        $Location.selectedItems.Add($Location.ItemsSource[0])
        return $shouldContinue
}

function preAddFields {
    $MapListArray = @()
    $normalMatch = @()
    $funkyMatch = @()
    $noLuckNoMatch=0
    foreach($sourceOne in $WPFlistView_Fields_Source.ItemsSource){
        $destOne = $WPFlistView_Fields_Dest.ItemsSource | where {$_.Name -eq $sourceOne.Name}
        if($destOne){
            if($sourceOne.type -eq $destOne.type){
            write-host $sourceOne -ForegroundColor Green
                if($WPFtextBox_URL.text -ne $WPFtextBox_URL_Dest.text){

                if(($global:sourceSPTypeCheck -ne "Meta") -and ($global:destSPTypeCheck -ne "File")){
                    $quickCheck = "Normal"
                    if(($sourceOne.type -ne "Lookup") -and ($sourceOne.type -ne "Managed Metadata") -and ($sourceOne.type -ne "Calculated") -and ($sourceOne.type -ne "External Data")){
                        $obj = new-object psobject -Property @{
                            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                            'SourceName' = $sourceOne.Name
                            'SourceType' = $sourceOne.Type
                            'DestinationName' = $sourceOne.Name
                            'DestinationType' = $sourceOne.Type
                        }     
                        $normalMatch+=$obj               
                    } else {
                        $obj = new-object psobject -Property @{
                            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                            'SourceName' = $sourceOne.Name
                            'SourceType' = $sourceOne.Type
                            'DestinationName' = $sourceOne.Name
                            'DestinationType' = $sourceOne.Type
                        } 
                        $funkyMatch+=$obj                                          
                    }
                    $lookupsMatch = $funkyMatch | where {$_.SourceType -eq "Lookup"}
                    $metaMatch = $funkyMatch | where {$_.SourceType -eq "Managed Metadata"}
                    $extMatch = $funkyMatch | where {$_.SourceType -eq "External Data"}   
}

                if($global:destSPTypeCheck -eq "File"){
                    $quickCheck = "DestFile"
                    $obj = new-object psobject -Property @{
                            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                            'SourceName' = $sourceOne.Name
                            'SourceType' = $sourceOne.Type
                            'DestinationName' = $sourceOne.Name
                            'DestinationType' = $sourceOne.Type
                        }     
                        $normalMatch+=$obj               
                }

                if(($global:sourceSPTypeCheck -eq "Meta")){
                    $quickCheck = "SourceMeta"
                    if(($sourceOne.type -ne "Calculated") -and ($sourceOne.type -ne "External Data")){
                        $obj = new-object psobject -Property @{
                            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                            'SourceName' = $sourceOne.Name
                            'SourceType' = $sourceOne.Type
                            'DestinationName' = $sourceOne.Name
                            'DestinationType' = $sourceOne.Type
                        }     
                        $normalMatch+=$obj               
                    } else {
                        $obj = new-object psobject -Property @{
                            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                            'SourceName' = $sourceOne.Name
                            'SourceType' = $sourceOne.Type
                            'DestinationName' = $sourceOne.Name
                            'DestinationType' = $sourceOne.Type
                        } 
                        $funkyMatch+=$obj                                          
                    }
                    $lookupsMatch = $funkyMatch | where {$_.SourceType -eq "Lookup"}
                    $metaMatch = $funkyMatch | where {$_.SourceType -eq "Managed Metadata"}
                    $extMatch = $funkyMatch | where {$_.SourceType -eq "External Data"}   
}

                } else {
                     $quickCheck = "Same"
                     if($sourceOne.type -ne "Calculated"){
                        write-host "SO Type"$sourceOne.type
                        write-host "SO Name"$sourceOne.name
                                                $obj = new-object psobject -Property @{
                        ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                        'SourceName' = $sourceOne.Name
                        'SourceType' = $sourceOne.Type
                        'DestinationName' = $sourceOne.Name
                        'DestinationType' = $sourceOne.Type
                    }     
                        $normalMatch+=$obj               
                    } else {
                                                $obj = new-object psobject -Property @{
                        ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
                        'SourceName' = $sourceOne.Name
                        'SourceType' = $sourceOne.Type
                        'DestinationName' = $sourceOne.Name
                        'DestinationType' = $sourceOne.Type
                    } 
                        $funkyMatch+=$obj                                          
                    }               
                    $lookupsMatch = $funkyMatch | where {$_.SourceType -eq "Lookup"}
                    $metaMatch = $funkyMatch | where {$_.SourceType -eq "Managed Metadata"}
                    $extMatch = $funkyMatch | where {$_.SourceType -eq "External Data"}                   
                }
            }        
        } else {
        $noLuckNoMatch++
        }
    }


$changeBetween = $WPFlistView_Fields_Source.ItemsSource.count - $WPFlistView_Fields_Dest.ItemsSource.count

    if($quickCheck -eq "Normal") {
    $WPFtextBlock_Fields.text = @"
$($normalMatch.count) mathing fields Automatically Mapped.
$($funkyMatch.count) matching fields not mapped.  
$noLuckNoMatch fields on the Source have no match.  
"@ 
}

    if($quickCheck -eq "SourceMeta") {
    $WPFtextBlock_Fields.text = @"
$($normalMatch.count) mathing fields Automatically Mapped.
$($funkyMatch.count) matching fields not mapped.  
$noLuckNoMatch fields on the Source have no match.       
"@ 
}

    if($quickCheck -eq "DestFile") {
    $WPFtextBlock_Fields.text = @"
The app will preserve all metadata in a CSV file called "Metadata.csv" on the root of the target folder.  
"@ 
}

    if($quickCheck -eq "Same") {
    $WPFtextBlock_Fields.text = @"
$($normalMatch.count) mathing fields Automatically Mapped.
$($funkyMatch.count) matching fields not mapped.  
$noLuckNoMatch fields on the Source have no match.           
"@ 
}
    foreach ($item in $normalMatch){
            $MapListArray += $item
    }

    $WPFlistView_Fields_Final.ItemsSource += $MapListArray
}
####Get Fields Button

####Select All Matching Fields in Fields ListViews Button
$WPFbutton_Map_Matching.Add_Click({

####Clear Selections
$WPFlistView_Fields_Source.SelectedItems.Clear()
$WPFlistView_Fields_Dest.SelectedItems.Clear()

    foreach ($item in $WPFlistView_Fields_Source.Items){
        if ($WPFlistView_Fields_Dest.Items.name -contains $item.name){
            Write-host $item
            $WPFlistView_Fields_Source.selecteditems.Add($item)


            $destItem = $WPFlistView_Fields_Dest.Items | where {$_.Name -eq $item.name}
            $WPFlistView_Fields_Dest.selecteditems.Add($destItem)
        }
    }
})

####Disselect all fields in Fields ListViews Button
$WPFbutton_Map_Disselect.Add_Click({

####Clear Selections
$WPFlistView_Fields_Source.SelectedItems.Clear()
$WPFlistView_Fields_Dest.SelectedItems.Clear()
})

$WPFbutton_Import_User_CSV.Add_Click({
        $global:userData = @()
        $WPFlabel_Imported_Status.Content = ""  
        $Form.Dispatcher.Invoke("Background", [action]{})

        ####Browse For CSV
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $templatepathfull = New-Object System.Windows.Forms.OpenFileDialog
        $templatepathfull.initialDirectory = "%HOMEPATH%"
        $templatepathfull.filter = "All files (*.*)| *.*"
        $templatepathfull.ShowDialog() | Out-Null
        $templatepath = $templatepathfull.filename
        $fileNameCSV = split-path  $templatepath -leaf -resolve
        if($fileNameCSV.split(".")[-1] -eq "CSV"){
            try{
                $global:userData=import-csv -path $templatepath
            } catch {
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button
                [System.Windows.MessageBox]::Show("Couldn't import CSV. Check file.","Warning", $Button)
            }
            if(($global:userData.Source_User[0]) -and ($global:userData.Destination_User[0])){
              $WPFlabel_Imported_Status.Content = "User Map: $fileNameCSV"  
              $Form.Dispatcher.Invoke("Background", [action]{})
            } else {
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button
                [System.Windows.MessageBox]::Show("CSV is not containing expected data. Make sure the CSV has two columns - Source_User and Destination_User which contain respectively the Source Users and the Destination Users", $Button)
            }
        } else {
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button
        [System.Windows.MessageBox]::Show("Filetype is not CSV","Warning", $Button)
        }
})

################################################Fields Mapping Buttons

####Map all selected Fields in Fields ListViews to Final Mapping. Also Populate Filter List View!
$WPFbutton_Map_Map.Add_Click({
    ####Map
    $i=0
    $MapListArray = @()
    $shouldCopy = $true
    $oneCheck = $false
    $twoCheck = $false

    if(!($WPFlistView_Fields_Source.selecteditems)){
        $shouldCopy = $false
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button
        $infoText = @"
Please make sure every selected Source Field has a corresponding Destination Field and vice versa.
"@
        [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)    
    }

    foreach ($item in $WPFlistView_Fields_Source.selecteditems){
         ####Create Final Mapping List Object

        $obj = new-object psobject -Property @{
            ####This below works because no matter how the items are shown in the WPF ListView, in the Array they are ordered
            'SourceName' = $item.Name
            'SourceType' = $item.Type
            'DestinationName' = $WPFlistView_Fields_Dest.selecteditems[$i].Name
            'DestinationType' = $WPFlistView_Fields_Dest.selecteditems[$i].Type
        }
        #$obj | Add-Member -type NoteProperty -Name 'Action' -Value "Copy"
        $MapListArray += $obj
        $i++
    }



    if($WPFtextBox_URL.text -ne $WPFtextBox_URL_Dest.text){
        if($WPFcheckBox_Manage_Lookup_Auto.isChecked -eq $false){
                $checkForLookupToLookup = $false
                foreach($field in $MapListArray){
                    if($field.DestinationType -eq "Lookup"){
                        $checkForLookupToLookup = $true
                    }
                }
                if($checkForLookupToLookup -eq $true){
                    if($WPFcheckBox_Manage_Lookup_Auto.content -eq "Copy Lookup Fields as Text Fields"){
                        $shouldCopy = $false
                        [String]$Button="OK"
                        $Button = [System.Windows.MessageBoxButton]::$Button
                        $infoText = @"
To copy Lookup Fields between Site Collections select 'Manage Mapped Lookup Fields Automatically'. 

This will automatically create a new Text Field to store Lookup Value on the Destination when you map Lookup fields. 

You can also leave Destination Field blank and a Text Field with similar name will be created on the Destination. 

Yet another option is to create Text fields on the Destination yourself and then map the Source Lookups to them.
"@
                        [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)
                    }
                }
            } else {
                $oneCheck = $true
                foreach($field in $MapListArray){
                        if($field.DestinationType -eq "Lookup"){
                            $field.DestinationName = $field.DestinationName+"_Value_Text"
                            $field.DestinationType = "Single line of text"
                        }

                        if(($field.SourceType -eq "Lookup") -and (!($field.DestinationType))){
                            $field.DestinationName = $field.SourceName+"_Value_Text"
                            $field.DestinationType = "Single line of text"
                        }
                    }      
                }
            if($WPFcheckBox_Manage_Meta_Auto.isChecked -eq $false){
                $checkForMetaToMeta = $false            
                foreach($field in $MapListArray){
                    if($field.DestinationType -eq "Managed Metadata"){
                        $checkForMetaToMeta = $true
                    }
                }           
            
                if($checkForMetaToMeta -eq $true){
                    if($WPFcheckBox_Manage_Lookup_Auto.content -eq "Copy Lookup Fields as Text Fields"){
                        if(!($global:showOnlyOnce)){
                            $global:showOnlyOnce = "Yes"
                            [String]$Button="OK"
                            $Button = [System.Windows.MessageBoxButton]::$Button
                            $infoText = @"
You are going to copy a Metadata field. This is only going to work if your Source and Destination are in the same Farm or Tennant and your Destination Metadata Field is configured the same as your Source Metadata Field. 

If either of these is false you have the option of copying the Metadata Field as a Text Field by checking "Copy Metadata Fields as Text Fields"
"@   
                            [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)
                        }
                    }
                }
            } else {
                $twoCheck = $true
                foreach($field in $MapListArray){
                        if($field.DestinationType -eq "Managed Metadata"){
                            $field.DestinationName = $field.DestinationName+"_Value_Text"
                            $field.DestinationType = "Single line of text"
                        }

                        if(($field.SourceType -eq "Managed Metadata") -and (!($field.DestinationType))){
                            $field.DestinationName = $field.SourceName+"_Value_Text"
                            $field.DestinationType = "Single line of text"
                        }
                }
            }
        }
        if(($oneCheck -eq $true) -and ($twoCheck -eq $true)){
            $shouldCopy = $true
        }
        foreach ($row in $MapListArray){

            if(($row.SourceName) -and ($row.DestinationName)){

            if($row.SourceType -ne $row.DestinationType){
                if(($row.DestinationType -ne "Single line of text") -and ($row.DestinationType -ne "Multiple lines of text")){
                $shouldCopy = $false
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button

                $infoText = @"
You cannot map fields of different types.!
"@
                [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)               
                }
        }
            } else {
                $shouldCopy = $false
                [String]$Button="OK"
                $Button = [System.Windows.MessageBoxButton]::$Button

                $infoText = @"
Please make sure every selected Source Field has a corresponding Destination Field and vice versa.
"@
                [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)               
            }
    }
    if($shouldCopy -eq $true){

    write-host "I Source" $WPFlistView_Fields_Final.ItemsSource

        foreach ($item in $MapListArray){
            if (!(($WPFlistView_Fields_Final.ItemsSource.SourceName -contains $item.SourceName) -and ($WPFlistView_Fields_Final.ItemsSource.DestinationName -contains $item.DestinationName) -and ($WPFlistView_Fields_Final.ItemsSource.SourceType -contains $item.SourceType)-and ($WPFlistView_Fields_Final.ItemsSource.DestinationType -contains $item.DestinationType))){
                #$WPFlistView_Fields_Final.ItemsSource += $MapListArray
                $WPFlistView_Fields_Final.ItemsSource += $item
                $WPFlistView_Fields_Final.selectedItems.Add($item)
            } 
        }
        opacityAnimation -grid $WPFField_Advanced -action "Close"
    }
})

####Warning for Library and Approve on Checking the Checkbox
$WPFcheckBox_Approve.Add_Click({
    if($WPFcheckBox_Approve.isChecked -eq $true){
        if($list.BaseType -eq "DocumentLibrary"){  
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            $infoText = @"
Approving the Files in the Destination Library will change the "Modified" field value to the Current Time. 

If you need to preserve these Field values turn off Approval on the Destination Library and then copy the items.
"@   
            [System.Windows.MessageBox]::Show($infoText,"Warning", $Button)
        }
    }
})

####Remove Selected Mappings in Final Mapping List View
$WPFbutton_Map_Remove_Selected.Add_Click({

    $mapAll = @()
    foreach($item in $WPFlistView_Fields_Final.ItemsSource){
    $mapAll+=$item
    }

    foreach($item in $WPFlistView_Fields_Final.selecteditems){
        $mapAll = $mapAll | where {$_ -ne $item}
    }
write-host "couny"$mapAll.count
write-host "sel" $WPFlistView_Fields_Final.selecteditems
    if($mapAll.count){
        $WPFlistView_Fields_Final.ItemsSource = $mapAll
    } else {
        $WPFlistView_Fields_Final.ItemsSource = @()
        $mapAll = @()
    }  
})

$WPFbutton_Map_Remove_All.Add_Click({
    $WPFlistView_Fields_Final.ItemsSource = @()
})

function listBrowser ($itemsForView, $listViewToDo){
        $browserArray = @()
        foreach ($item in $itemsForView){
            $obj = new-object psobject -Property @{
                'path' = $item["FileRef"]
                'name' = $item["FileRef"].split("/")[-1]
            }
            if($item["FSObjType"] -eq 1){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            if($item["FSObjType"] -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            $browserArray += $obj
            #$wpflistView_Browser.items.Add($obj)
        }
        $listViewToDo.ItemsSource = $browserArray
    }

function listBrowserMetaDataFile ($itemsForView, $listViewToDo, $currentLoc){
        $browserArray = @()
        foreach ($item in $itemsForView){
            $obj = new-object psobject -Property @{
                'path' = $item.FileRef
                'name' = $item.FileRef.split("/")[-1]
            }
            if($item.FSObjType -eq 1){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item.FSObjType
            }
            if($item.FSObjType -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item.FSObjType
            }
            $browserArray += $obj
            #$wpflistView_Browser.items.Add($obj)
        }
        $listViewToDo.ItemsSource = $browserArray

        ####Store Current Location in Button
        $WPFradioButton_Browser.tag = $global:metadataAll.FileDirRef[0]
        $wpflistView_Browser.tag = $currentLoc

        $global:selectedItemsForCopy = $itemsForView
    }

####Enter a Folder in Explorer
function getFolderCAML($user,$password,$SiteURL,$docLibName,$folderPath){
    ####Get Source List
    $List = $Context.Web.Lists.GetByTitle($DocLibName)
    $Context.Load($List)
    $context.Load($List.RootFolder)
    $Context.ExecuteQuery()
    if(!($folderPath)){
        $folderPath = $List.RootFolder.serverrelativeurl

        ####Create CAML Query For Items
        $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        $qryItems.viewXML = @"
<View Scope="All">
    <Query>
    </Query>
</View>
"@
} else {
        $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        $qryItems.FolderServerRelativeUrl = $folderPath
        $qryItems.viewXML = @"
<View Scope="All">
    <Query>
    </Query>
</View>
"@
}
    [Microsoft.SharePoint.Client.ListItemCollection]$items = $List.GetItems($qryItems)
    ####Commit
    $context.Load($items)

    try {
        $context.ExecuteQuery()
    } catch {
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button

        $exceptionMessage = $_.exception.message
        $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

        [System.Windows.MessageBox]::Show($message,"Warning", $Button)

        $WPFlabel_Status.Content = "Status: Idle"
        }
    ####Store Current Location in Button
    $WPFradioButton_Browser.tag = $List.RootFolder.serverrelativeurl
    $wpflistView_Browser.tag = $folderPath

    $global:selectedItemsForCopy = $items

    ####Return
    return $items
}

####Browser Radio
$WPFradioButton_Browser.Add_Click({    
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
    opacityAnimation -grid $WPFbrowserGrid -action "Pre"


    if($global:sourceSPTypeCheck -eq "Meta"){
        $itemsForView = $global:metadataAll | where {($_.FileDirRef -eq $global:metadataAll.FileDirRef[0]) -and ($_.FSObjType -ne "Type")}
        listBrowserMetaDataFile -itemsForView $itemsForView -listViewToDo $wpflistView_Browser -currentLoc $global:metadataAll.FileDirRef[0]
    } else {
        $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath ""
        listBrowser -itemsForView $itemsForView -listViewToDo $wpflistView_Browser
    }

    addToGlobalCopymode -mode "Browser"

    if($WPFallGrid.Visibility -eq "Visible"){      
        opacityAnimation -grid $WPFallGrid -action "Close"
    }
        if($WPFgrid_Filters.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFgrid_Filters -action "Close"
        ####Nullify
        $WPFtextBox_Filter_Value.Text = ""
        $WPFcomboBox_Field_To_Filter.itemsSource = @()
        $WPFcomboBox_Condition.items.clear()
        $WPFlistView_Filters.items.clear()
        if($WPFbutton_DateValidator.visibility -eq "Visible"){
            opacityAnimation -grid $WPFbutton_DateValidator -action "Close"
        }
    }
    opacityAnimation -grid $WPFbrowserGrid -action "Open"
    $WPFimage_Tips.source = $dirFormURLs+"/Explorer.png"

})

function CopyAllRadio {
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
    $global:copyMode = "NoFilter"

    if($WPFbrowserGrid.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFbrowserGrid -action "Close"
    }
    if($WPFgrid_Filters.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFgrid_Filters -action "Close"
    }

        if($global:sourceSPTypeCheck -eq "Meta"){
            write-host "Meta"
            $itemsForView = $global:metadataAll | where {($_.FileDirRef -eq $global:metadataAll.FileDirRef[0]) -and ($_.FSObjType -ne "Type")}
            listBrowserMetaDataFile -itemsForView $itemsForView -listViewToDo $wpflistView_all -currentLoc $global:metadataAll.FileDirRef[0]

        } elseif ($global:sourceSPTypeCheck -eq "File"){
            $contentsFileRoot = get-childitem $global:fileShareRoot 
               
            $rootListFiles = @()
            foreach ($item in $contentsFileRoot){
                write-host "File"
                $obj = new-object psobject -Property @{
                    'path' = $item.fullname
                    'name' = $item.name
                }
                if($item.PSIsContainer){
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value "Folder"
                } else {
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value "File"
                }
                $rootListFiles += $obj
                #$wpflistView_Browser.items.Add($obj)
            }
            $wpflistView_all.ItemsSource = $rootListFiles
        } else {
            write-host "Else"

            $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath ""
            listBrowser -itemsForView $itemsForView -listViewToDo $wpflistView_all
        }
    opacityAnimation -grid $WPFallGrid -action "Open"
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
}

####Copy All Radio
$WPFradioButton_No_Filter.Add_Click({
    CopyAllRadio
})

####Copy with Filter Radio
$WPFradioButton_Filter.Add_Click({  
    $global:copyMode = "Filter"  
    if($WPFbrowserGrid.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFbrowserGrid -action "Close"
        ####Nullify Selected Items for Copy
        $addItemsForViewArray = @()
        $WPFlistView_Items.itemsSource = @()
        $global:itemsToCopy = @()
    }
    if($WPFallGrid.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFallGrid -action "Close"
    }
    $WPFimage_Tips.source = $dirFormURLs+"/Filters.png"
    opacityAnimation -grid $WPFgrid_Filters -action "Open"
    $fieldsToShowFilter = $WPFlistView_Fields_Source.itemssource | where {(($_.Type -eq "Single line of Text") -or ($_.Type -eq "Date and Time") -or ($_.Type -eq "Lookup"))}
    $WPFcomboBox_Field_To_Filter.itemsSource = $fieldsToShowFilter.Name
})

####Selection Changed in Field to Filter
$WPFcomboBox_Field_To_Filter.add_SelectionChanged({
    $global:comboSelection = $WPFlistView_Fields_Source.itemssource | where {$_.Name -eq $WPFcomboBox_Field_To_Filter.SelectedValue}
    if ($global:comboSelection.Type -eq "Date and Time"){
        $WPFcomboBox_Condition.items.clear()
        $WPFcomboBox_Condition.Items.Add("Equal")
        $WPFcomboBox_Condition.Items.Add("Greater")
        $WPFcomboBox_Condition.Items.Add("Lesser")
        if($WPFbutton_DateValidator.visibility -eq "Hidden"){
            opacityAnimation -grid $WPFbutton_DateValidator -action "Open"
        }       
    } elseif ($global:comboSelection.Type -eq "Lookup") {
        $WPFcomboBox_Condition.items.clear()
        $WPFcomboBox_Condition.Items.Add("Equal")
        $WPFcomboBox_Condition.Items.Add("Greater")
        $WPFcomboBox_Condition.Items.Add("Lesser")
        if($WPFbutton_DateValidator.visibility -eq "Visible"){
            opacityAnimation -grid $WPFbutton_DateValidator -action "Close"
        }
    } else {
        $WPFcomboBox_Condition.items.clear()
        $WPFcomboBox_Condition.Items.Add("Equal")
        if($WPFbutton_DateValidator.visibility -eq "Visible"){
            opacityAnimation -grid $WPFbutton_DateValidator -action "Close"
        }
    }

})

$WPFbutton_DateValidator.Add_click({
    try{
        [datetime]$validateTime = $WPFtextBox_Filter_Value.Text
        $WPFlabel_Status.Content = "Status: Date Validated"
        $Form.Dispatcher.Invoke("Background", [action]{})
    } catch {
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button
        [System.Windows.MessageBox]::Show("Date can't be validated. Try 'MM.DD.YYYY' or 'MM.DD.YYYY hh:mm'.","Warning", $Button)
    }
    $WPFtextBox_Filter_Value.Text = $validateTime.ToString('yyyy-MM-dd hh:mm:ss')
})

function createCAMLFilter {
    if($WPFlistView_Filters.items.count -eq 0){       
        $camlQueryConstruct = @"
<View Scope="RecursiveAll">
    <Query>
    </Query>
</View>
"@
    }
    if($WPFlistView_Filters.items.count -eq 1){       
        $camlQueryConstruct = @"
<View Scope="RecursiveAll">
    <Query>
		<Where>
			<{0}>
				<FieldRef Name='{1}' /><Value Type='{2}'>{3}</Value>
			</{0}>
		</Where>
    </Query>
</View>
"@ -f $WPFlistView_Filters.items[0].Condition, $WPFlistView_Filters.items[0].Field, $WPFlistView_Filters.items[0].Type, $WPFlistView_Filters.items[0].Value
    }
    if($WPFlistView_Filters.items.count -eq 2){       
        $camlQueryConstruct = @"
<View Scope="RecursiveAll">
    <Query>
		<Where>
            <And>
			    <{0}>
				    <FieldRef Name='{1}' /><Value Type='{2}'>{3}</Value>
			    </{0}>
			    <{4}>
				    <FieldRef Name='{5}' /><Value Type='{6}'>{7}</Value>
			    </{4}>
            </And>
		</Where>
    </Query>
</View>
"@ -f $WPFlistView_Filters.items[0].Condition, $WPFlistView_Filters.items[0].Field, $WPFlistView_Filters.items[0].Type, $WPFlistView_Filters.items[0].Value, $WPFlistView_Filters.items[1].Condition, $WPFlistView_Filters.items[1].Field, $WPFlistView_Filters.items[1].Type, $WPFlistView_Filters.items[1].Value
    }
    if($WPFlistView_Filters.items.count -eq 3){       
        $camlQueryConstruct = @"
<View Scope="RecursiveAll">
    <Query>
		<Where>
            <And>
			    <{0}>
				    <FieldRef Name='{1}'/><Value Type='{2}'>{3}</Value>
			    </{0}>
                <And>
			        <{4}>
				        <FieldRef Name='{5}' /><Value Type='{6}'>{7}</Value>
			        </{4}>
			        <{8}>
				        <FieldRef Name='{9}' /><Value Type='{10}'>{11}</Value>
			        </{8}>
                </And>
            </And>
		</Where>
    </Query>
</View>
"@ -f $WPFlistView_Filters.items[0].Condition, $WPFlistView_Filters.items[0].Field, $WPFlistView_Filters.items[0].Type, $WPFlistView_Filters.items[0].Value, $WPFlistView_Filters.items[1].Condition, $WPFlistView_Filters.items[1].Field, $WPFlistView_Filters.items[1].Type, $WPFlistView_Filters.items[1].Value, $WPFlistView_Filters.items[2].Condition, $WPFlistView_Filters.items[2].Field, $WPFlistView_Filters.items[2].Type, $WPFlistView_Filters.items[2].Value
    }

    return $camlQueryConstruct
}

$WPFbutton_Test_Filter.Add_click({
    progressbar -state "Start" -all 3
    $WPFlabel_Status.Content = "Status: Testing Filter Query!"
    $camlQueryConstruct = createCAMLFilter
    write-host $camlQueryConstruct 

    progressbar -state "Plus" 
    [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

    ####Get Source List
    $web = $Context.Web
    $List = $Context.Web.Lists.GetByTitle($sourceListTitle)
    $Context.Load($Web)
    $Context.Load($List)
    $context.Load($List.RootFolder)
    progressbar -state "Plus" 
    $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $qry.viewXML = $camlQueryConstruct

    [Microsoft.SharePoint.Client.ListItemCollection]$itemsConstructedQuery = $list.GetItems($qry)
    $Context.Load($itemsConstructedQuery)
    try {
        $Context.ExecuteQuery()
    } catch {
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button

        $exceptionMessage = $_.exception.message
        $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

        [System.Windows.MessageBox]::Show($message,"Warning", $Button)

        $WPFlabel_Status.Content = "Status: Idle"
        }
    progressbar -state "Plus" 
    listBrowser -itemsForView $itemsConstructedQuery -listViewToDo $WPFlistView_Filter_Query_Test
    $WPFlabel_Status.Content = "Status: Idle!"
    progressbar -state "Stop"

})

$WPFbutton_AddFilter.Add_Click({
    if($WPFlistView_Filters.items.count -lt 3){
    write-host $WPFlistView_Filters.items.count
        $obj = new-object psobject -Property @{
            'Field' = $global:comboSelection.Name
            'Condition' = ""
            'Value' = $WPFtextBox_Filter_Value.Text
            'Type' = ""
        }

        if($WPFcomboBox_Condition.SelectedValue -eq "Equal"){
            $obj.Condition = "Eq"
        }
        if($WPFcomboBox_Condition.SelectedValue -eq "Greater"){
            $obj.Condition = "Gt"
        }
        if($WPFcomboBox_Condition.SelectedValue -eq "Lesser"){
            $obj.Condition = "Lt"
        }
        if($global:comboSelection.Type -eq "Single line of Text"){
            $obj.Type = "Text"
        }
        if($global:comboSelection.Type -eq "Lookup"){
            $obj.Type = "Lookup"
        }
        if($global:comboSelection.Type -eq "Date and Time"){
            $obj.Type = "Datetime"
        }

        if(($obj.Field) -and ($obj.Value) -and ($obj.Condition)){
            $WPFlistView_Filters.items.add($obj)
        } else {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button
            [System.Windows.MessageBox]::Show("Not all fields are filled.","Warning", $Button)
        }
    } else {
        #try {[System.Windows.Window]} 
        #catch{Add-Type -AssemblyName PresentationCore,PresentationFramework} 
        [String]$Button="OK"
        $Button = [System.Windows.MessageBoxButton]::$Button
        [System.Windows.MessageBox]::Show("Up to three filters supported.","Warning", $Button)
    }
})

$WPFbutton_All_Filters_Remove_Selected.Add_Click({
    $itemsToRemove = @()
    foreach($item in $WPFlistView_Filters.SelectedItems){
        $itemsToRemove+=$item
    }
    foreach($item in  $itemsToRemove){
        $WPFlistView_Filters.items.remove($item)
    }
})

####Get into browser
$wpflistView_Browser.add_MouseDoubleClick({
    $selectedItem = $wpflistView_Browser.SelectedItem
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

    write-host "Selected Item Path"$selectedItem.path -ForegroundColor Green

    if($selectedItem.type -eq 1){

        if($global:sourceSPTypeCheck -eq "Meta"){
            $itemsForView = $global:metadataAll | where {($_.FileDirRef -eq $selectedItem.path) -and ($_.FSObjType -ne "Type")}
            listBrowserMetaDataFile -itemsForView $itemsForView -listViewToDo $wpflistView_Browser -currentLoc $selectedItem.path
        } else {
            $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath $selectedItem.path
            listBrowser -itemsForView $itemsForView -listViewToDo $wpflistView_Browser
        }
    }
})

$WPFbutton_BrowserUp.Add_Click({
    ####If current Folder is not Root
    if($wpflistView_Browser.tag -ne $WPFradioButton_Browser.tag){
        if($global:sourceSPTypeCheck -eq "Meta"){
            $currentLocationTrim = $wpflistView_Browser.tag
            $currentLocationTrim = $currentLocationTrim.Substring(0, $currentLocationTrim.lastIndexOf('/'))

            $itemsForView = $global:metadataAll | where {($_.FileDirRef -eq $currentLocationTrim) -and ($_.FSObjType -ne "Type")}
            listBrowserMetaDataFile -itemsForView $itemsForView -listViewToDo $wpflistView_Browser -currentLoc $currentLocationTrim
        } else {
            [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
            ###Trim Current Folder with one level
            $currentLocationTrim = $wpflistView_Browser.tag
            $currentLocationTrim = $currentLocationTrim.Substring(0, $currentLocationTrim.lastIndexOf('/'))

            $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.password -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath $currentLocationTrim
            listBrowser -itemsForView $itemsForView -listViewToDo $wpflistView_Browser
        }
    }
})

function addToGlobalCopymode($mode){
$global:copyMode = $mode
}

$WPFbutton_Add_For_Copy.Add_Click({
    if($global:sourceSPTypeCheck -eq "Meta"){

        foreach ($item in $global:selectedItemsForCopy){
            if($wpflistView_Browser.selecteditems.Path -contains $item.FileRef){
                if($global:itemsToCopy.count -eq 1){
                    $tempItemArray = @()
                    $tempItemArray += $global:itemsToCopy
                    $tempItemArray += $item
                    $global:itemsToCopy = $tempItemArray
                } else {
                    $global:itemsToCopy = $global:itemsToCopy + $item 
                }    
            } 
             $listViewForItemsArray = @()
            foreach ($item in $global:itemsToCopy){
                $obj = new-object psobject -Property @{
                    'path' = $item.FileRef
                    'name' = $item.FileRef.split("/")[-1]
                }
                if($item.FSObjType -eq 1){
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
                }
                if($item.FSObjType -eq 0){
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
                }
                $listViewForItemsArray += $obj   
            }
            $WPFlistView_Items.itemsSource = $listViewForItemsArray                  
        }

    } else {
        foreach ($item in $global:selectedItemsForCopy){
            if($wpflistView_Browser.selecteditems.Path -contains $item["FileRef"]){
                ####Add to Items for Copy
                if($global:itemsToCopy.count -eq 1){
                    $tempItemArray = @()
                    $tempItemArray += $global:itemsToCopy
                    $tempItemArray += $item
                    $global:itemsToCopy = $tempItemArray
                } else {
                    $global:itemsToCopy = $global:itemsToCopy + $item 
                }    
            }

            $listViewForItemsArray = @()
            foreach ($item in $global:itemsToCopy){
                $obj = new-object psobject -Property @{
                    'path' = $item["FileRef"]
                    'name' = $item["FileRef"].split("/")[-1]
                }
                if($item["FSObjType"] -eq 1){
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
                }
                if($item["FSObjType"] -eq 0){
                    $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                    $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
                }
                $listViewForItemsArray += $obj   
            }
            $WPFlistView_Items.itemsSource = $listViewForItemsArray
        }
    }
})

$WPFbutton_ItemsListView_RemoveAll.Add_Click({
    $global:itemsToCopy = @()
    $WPFlistView_Items.Clear()
    $WPFlistView_Items.itemsSource = @()
})

$WPFbutton_ItemsListView_RemoveSelected.Add_Click({
    ####Remove selected from Array for Copy 
    foreach($item in $WPFlistView_Items.SelectedItems){
        $global:itemsToCopy = $global:itemsToCopy | where {$_["FileRef"] -ne $item.path}
    }
        write-host "Count"
    if($global:itemsToCopy.count){

        $listViewForItemsArray = @()
        foreach ($item in $global:itemsToCopy){
            $obj = new-object psobject -Property @{
                'path' = $item["FileRef"]
                'name' = $item["FileRef"].split("/")[-1]
            }
            if($item["FSObjType"] -eq 1){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            if($item["FSObjType"] -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "$dir\file-icon-28038.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            $listViewForItemsArray += $obj 
              
        }
        $WPFlistView_Items.itemsSource = $listViewForItemsArray
    } else {
        $global:itemsToCopy = @()        
        $WPFlistView_Items.itemsSource = @()
    }
})

function copyBrowsedItems($browsedItems) {
        [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

        ####Get Source List
        $web = $Context.Web
        $List = $Context.Web.Lists.GetByTitle($sourceListTitle)
        $Context.Load($Web)
        $Context.Load($List)
        $context.Load($List.RootFolder)

        ####Get Destination List
        $destWeb = $destContext.Web
        $destList = $destContext.Web.Lists.GetByTitle($destListTitle)

        $destContext.Load($destList)
        $destContext.Load($destWeb)
        $destContext.Load($destList.RootFolder)

        $Context.ExecuteQuery()
        $destContext.ExecuteQuery()

        iterateEverything -whatTo $browsedItems

        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
        progressbar -state "Stop"

        #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
        #$global:errorsAll | Out-File -FilePath $errorLogPath
        #Invoke-Item  $errorLogPath

        $global:errorsAll | Out-GridView -Title "Error Log"
        $global:errorsAll = @()
}

################################################Copy Lists Button

function copyFunction {
   $WPFlabel_Status.Content = "Status: Starting Copy"
   [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
   [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

    ####This line is for Status Update. Otherwise the Label won't Update. More Info: https://powershell.org/forums/topic/refresh-wpa-label-multiple-times-with-one-click/
    $Form.Dispatcher.Invoke("Background", [action]{})

        ####Get Source List
        $web = $Context.Web
        $List = $Context.Web.Lists.GetByTitle($sourceListTitle)
        $Context.Load($Web)
        $Context.Load($List)
        $context.Load($List.RootFolder)

        ####Get Destination List
        $destWeb = $destContext.Web
        $destList = $destContext.Web.Lists.GetByTitle($destListTitle)

        $destContext.Load($destList)
        $destContext.Load($destWeb)
        $destContext.Load($destList.RootFolder)
        $destContext.Load($destList.Fields)
        $destContext.Load($destList.Views)

        $Context.ExecuteQuery()
        $destContext.ExecuteQuery()
        write-host "Starting" -ForegroundColor Cyan
        if($WPFcheckBox_Manage_Lookup_Auto.isChecked -eq $true){
        write-host "Lookup Creation Func" -ForegroundColor Cyan
            ####Create Text Fields for Lookup Columns
            foreach ($field in $WPFlistView_Fields_Final.items | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -eq "Single line of text"))}){
                write-host "Lookup Field"$field.SourceName 
                #$fieldToCreateRaw = $field.SourceName + "_Raw_Text"
                $fieldToCreateValue =  $field.DestinationName
                if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                }
                ####Field for ID
                #if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                #    write-host "Creating Field"$fieldToCreateValue
                #    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                #}

                foreach ($view in $destList.Views){
                    if($view.defaultView -eq $true){
                        if($view.ViewFields[$field.SourceName]){
                            $view.ViewFields.Remove($field.SourceName);
                            $view.Update();
                        }
                    }
                }

                $destList.Update()
                $destContext.ExecuteQuery()    
            }

        }

        if($WPFcheckBox_Manage_Meta_Auto.isChecked -eq $true){
            ####Create Text Fields for Metadata Columns
            foreach ($field in $WPFlistView_Fields_Final.items | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -eq "Single line of text"))}){
                write-host "Lookup Field"$field.SourceName 
                #$fieldToCreateRaw = $field.SourceName + "_Raw_Text"
                $fieldToCreateValue =  $field.DestinationName
                if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                }
                ####Field for ID
                #if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                #    write-host "Creating Field"$fieldToCreateValue
                #    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                #}

                foreach ($view in $destList.Views){
                    if($view.defaultView -eq $true){
                        if($view.ViewFields[$field.SourceName]){
                            $view.ViewFields.Remove($field.SourceName);
                            $view.Update();
                        }
                    }
                }

                $destList.Update()
                $destContext.ExecuteQuery()    
            }                
        }


    if ($global:copyMode -eq "Browser"){
        Write-Host "File Ref of Folder" $sourceListTitle -ForegroundColor Cyan

        $itemsToCopyFinal = @()
        $foldersToCopyFinal = @()
        $allFinal = @()
        foreach ($item in $global:itemsToCopy){
            if($item["FSObjType"] -eq 0){
                $itemsToCopyFinal+=$item        
            }
            if($item["FSObjType"] -eq 1){
                ####Get Dir of Item
                Write-Host "File Ref of Folder" $item["FileRef"] -ForegroundColor Green
                $sourceRefFolder = $item["FileRef"]


                ####Create Query
                $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
                $qry.FolderServerRelativeUrl = $sourceRefFolder
                [Microsoft.SharePoint.Client.ListItemCollection]$items = $List.GetItems($qry)

                ####Commit
                $context.Load($items)

                try {
                    $context.ExecuteQuery()
                } catch {
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button

                    $exceptionMessage = $_.exception.message
                    $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

                    [System.Windows.MessageBox]::Show($message,"Warning", $Button)

                    $WPFlabel_Status.Content = "Status: Idle"
        }
                ####Add to Array
                $foldersToCopyFinal += $items
            }
        }

        ####Order Folder CAMLs by URL deepness 
        foreach ($item in $foldersToCopyFinal){
            $urlCount = ($item["FileRef"].ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            $item | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount
        }

        ####Sort items by URL Deepness. Create items with the least amaount of "/" first
        $foldersToCopyFinal = $foldersToCopyFinal | sort-object URLLengthCustom

        $allFinal += $foldersToCopyFinal 
        $allFinal += $itemsToCopyFinal

        progressbar -state "Start" -all $allFinal.count
        copyBrowsedItems -browsedItems $allFinal
    }

    if (($global:copyMode -eq "NoFilter") -or ($global:copyMode -eq "Filter")){

        copyLibrary -siteURL $WPFtextBox_URL.text -docLibName $WPFlistView.SelectedItem.Title -sourceSPType $global:sourceSPTypeCheck -destSiteURL $WPFtextBox_URL_Dest.text -destDocLibName $WPFlistView_Dest.SelectedItem.Title -destSPType $global:destSPTypeCheck
        
        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})

        #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
        #$global:errorsAll | Out-File -FilePath $errorLogPath
        #Invoke-Item  $errorLogPath
        $global:errorsAll | Out-GridView -Title "Error Log"
        $global:errorsAll = @()
        }
    }

function copyFunctionFileShare {
    $WPFlabel_Status.Content = "Status: Starting Copy"

    if(($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta")) {
    [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title

        if(($global:sourceSPTypeCheck -eq "Meta") -and (($WPFcheckBox_Manage_Lookup_Auto.isChecked -eq $true) -or ($WPFcheckBox_Manage_Meta_Auto.isChecked -eq $true))){
            ####Get Destination List
            $destWeb = $destContext.Web
            $destList = $destContext.Web.Lists.GetByTitle($destListTitle)

            $destContext.Load($destList)
            $destContext.Load($destWeb)
            $destContext.Load($destList.RootFolder)
            $destContext.Load($destList.Fields)
            $destContext.Load($destList.Views)    
            $destContext.ExecuteQuery()

        write-host "Starting" -ForegroundColor Cyan
        if($WPFcheckBox_Manage_Lookup_Auto.isChecked -eq $true){
        write-host "Lookup Creation Func" -ForegroundColor Cyan
            ####Create Text Fields for Lookup Columns
            foreach ($field in $WPFlistView_Fields_Final.items | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -eq "Single line of text"))}){
                write-host "Lookup Field"$field.SourceName 
                #$fieldToCreateRaw = $field.SourceName + "_Raw_Text"
                $fieldToCreateValue =  $field.DestinationName
                if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                }
                ####Field for ID
                #if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                #    write-host "Creating Field"$fieldToCreateValue
                #    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                #}

                foreach ($view in $destList.Views){
                    if($view.defaultView -eq $true){
                        if($view.ViewFields[$field.SourceName]){
                            $view.ViewFields.Remove($field.SourceName);
                            $view.Update();
                        }
                    }
                }

                $destList.Update()
                $destContext.ExecuteQuery()    
            }

        }

        if($WPFcheckBox_Manage_Meta_Auto.isChecked -eq $true){
            ####Create Text Fields for Metadata Columns
            foreach ($field in $WPFlistView_Fields_Final.items | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -eq "Single line of text"))}){
                write-host "Lookup Field"$field.SourceName 
                #$fieldToCreateRaw = $field.SourceName + "_Raw_Text"
                $fieldToCreateValue =  $field.DestinationName
                if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                }
                ####Field for ID
                #if(!($destList.Fields.InternalName -contains $fieldToCreateValue)){
                #    write-host "Creating Field"$fieldToCreateValue
                #    $destList.Fields.AddFieldAsXml("<Field Type='Text' DisplayName='$fieldToCreateValue'/>",$true,[Microsoft.SharePoint.Client.AddFieldOptions]::AddFieldToDefaultView)
                #}

                foreach ($view in $destList.Views){
                    if($view.defaultView -eq $true){
                        if($view.ViewFields[$field.SourceName]){
                            $view.ViewFields.Remove($field.SourceName);
                            $view.Update();
                        }
                    }
                }

                $destList.Update()
                $destContext.ExecuteQuery()    
            }                
        }                    
        }

        if($global:sourceSPTypeCheck -eq "File"){


        write-host "Getting FileShare!!" -ForegroundColor Green
            $allFileShareFiles = get-childitem -path $global:fileShareRoot -recurse

            foreach ($file in $allFileShareFiles){
                $urlCount = ($file.FullName.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
                $file | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount      
            }
        
            $filesSorted = $allFileShareFiles | sort-object URLLengthCustom 
            iterateFileShareSourceBasic -whatTo $filesSorted

            progressbar -state "Stop"
            ####Update Form Status
            $WPFlabel_Status.Content = "Status: Finished Copying!"

            #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
            #$global:errorsAll |  -FilePath $errorLogPath
            #Invoke-Item  $errorLogPath

            $global:errorsAll | Out-GridView -Title "Error Log"
            $global:errorsAll = @()
        } elseif(($global:sourceSPTypeCheck -eq "Meta") -and ($global:copyMode -eq "Browser")) {

            $itemsToCopyFinal = @()
            $foldersToCopyFinal = @()
            $allFinal = @()

            foreach ($item in $global:itemsToCopy){
                if($item.FSObjType -eq 0){
                    $itemsToCopyFinal+=$item        
                }   
                if($item.FSObjType -eq 1){       
                    $itemsToCopyFinal += $global:metadataAll | where {($_.FileDirRef -like "*"+$item.FileRef+"*") -and ($_.FSObjType -ne "Type")}
                    $itemsToCopyFinal += $global:metadataAll | where {($_.FileRef -eq $item.FileRef) -and ($_.FSObjType -ne "Type")}
                }  
                #$itemsToCopyFinal | Out-GridView
                }
                foreach ($item in $itemsToCopyFinal){
                    $urlCount = ($item.FileRef.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                    $item | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount
                }
                ####Sort items by URL Deepness. Create items with the least amaount of "/" first 
                $itemsToCopyFinal = $itemsToCopyFinal | sort-object URLLengthCustom
                ##$itemsToCopyFinal | Out-GridView
                iterateFileShareSourceMetaData -whatTo $itemsToCopyFinal
                progressbar -state "Stop"
                $WPFlabel_Status.Content = "Status: Finished Copying!"
                $global:errorsAll | Out-GridView -Title "Error Log" 
                $global:errorsAll = @()
            
        } else {
            foreach ($row in $global:metadataAll){
            $urlCount = ($row.FileRefLocal.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
            $row | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount 
            }
            $global:metadataAll = $global:metadataAll | sort-object URLLengthCustom 
            iterateFileShareSourceMetaData -whatTo $global:metadataAll

            progressbar -state "Stop"
            ####Update Form Status
            $WPFlabel_Status.Content = "Status: Finished Copying!"

            #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
            #$global:errorsAll |  -FilePath $errorLogPath
            #Invoke-Item  $errorLogPath

            $global:errorsAll | Out-GridView -Title "Error Log"
            $global:errorsAll = @()
            }
        }

    if($global:destSPTypeCheck -eq "File"){
        $WPFlabel_Status.Content = "Status: Starting Copy"
        [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

        ####This line is for Status Update. Otherwise the Label won't Update. More Info: https://powershell.org/forums/topic/refresh-wpa-label-multiple-times-with-one-click/
        $Form.Dispatcher.Invoke("Background", [action]{})

        ####Get Source List
        $web = $Context.Web
        $List = $Context.Web.Lists.GetByTitle($sourceListTitle)
        $Context.Load($Web)
        $Context.Load($List)
        $context.Load($List.RootFolder)

        $Context.ExecuteQuery()

    if ($global:copyMode -eq "Browser"){
        Write-Host "File Ref of Folder" $sourceListTitle -ForegroundColor Cyan

        $itemsToCopyFinal = @()
        $foldersToCopyFinal = @()
        $allFinal = @()
        foreach ($item in $global:itemsToCopy){
            if($item["FSObjType"] -eq 0){
                $itemsToCopyFinal+=$item        
            }
            if($item["FSObjType"] -eq 1){
                ####Get Dir of Item
                Write-Host "File Ref of Folder" $item["FileRef"] -ForegroundColor Green
                $sourceRefFolder = $item["FileRef"]

                ####Create Query
                $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
                $qry.FolderServerRelativeUrl = $sourceRefFolder
                [Microsoft.SharePoint.Client.ListItemCollection]$items = $List.GetItems($qry)

                ####Commit
                $context.Load($items)

                try {
                    $context.ExecuteQuery()
                } catch {
                    [String]$Button="OK"
                    $Button = [System.Windows.MessageBoxButton]::$Button

                    $exceptionMessage = $_.exception.message
                    $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

                    [System.Windows.MessageBox]::Show($message,"Warning", $Button)

                    $WPFlabel_Status.Content = "Status: Idle"
        }
                ####Add to Array
                $foldersToCopyFinal += $items
            }
        }

        ####Order Folder CAMLs by URL deepness 
        foreach ($item in $foldersToCopyFinal){
            $urlCount = ($item["FileRef"].ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            $item | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount
        }

        ####Sort items by URL Deepness. Create items with the least amaount of "/" first
        $foldersToCopyFinal = $foldersToCopyFinal | sort-object URLLengthCustom

        $allFinal += $foldersToCopyFinal 
        $allFinal += $itemsToCopyFinal

        progressbar -state "Start" -all $allFinal.count
        iterateFileShareDestination -whatTo $allFinal

        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
        progressbar -state "Stop"

        $global:errorsAll | Out-GridView -Title "Error Log"
        $global:errorsAll = @()
    }

    if (($global:copyMode -eq "NoFilter") -or ($global:copyMode -eq "Filter")){

        ####Check if Applying Filter and if not go on and copy all

        if ($global:copyMode -eq "NoFilter"){
            ####Create CAML Query
            $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
        }

        if ($global:copyMode -eq "Filter"){
            ####Create CAML Query
            $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
            $camlQueryConstruct = createCAMLFilter
            $qryItems.viewXML = $camlQueryConstruct
        }
        ####Load
        [Microsoft.SharePoint.Client.ListItemCollection]$items = $list.GetItems($qryItems)
        $Context.Load($items)

        ####Commit
        try {
            $Context.ExecuteQuery()
        } catch {
            [String]$Button="OK"
            $Button = [System.Windows.MessageBoxButton]::$Button

            $exceptionMessage = $_.exception.message
            $message = @"
We encountered the following Error while getting the Items

$exceptionMessage
"@

            [System.Windows.MessageBox]::Show($message,"Warning", $Button)

            $WPFlabel_Status.Content = "Status: Idle"
        }        
        #$destContext.ExecuteQuery()

        ####Add a Property reflecting URL deepness to avoid creating nested Items and Folders before the nesting Folder is created 
        foreach ($item in $items){
            $urlCount = ($item["FileRef"].ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
            $item | Add-Member -type NoteProperty -Name 'URLLengthCustom' -Value $urlCount
        }

        ####Sort items by URL Deepness. Create items with the least amaount of "/" first
        $itemsSorted = $items | sort-object URLLengthCustom

        ####Get Lists Relative URLs
        $listRelativeURL = $list.RootFolder.ServerRelativeUrl
        $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl

        foreach($item in $itemsSorted){
        write-host $item["FileRef"] -ForegroundColor Cyan
        }
        progressbar -state "Start" -all $itemsSorted.count
        iterateFileShareDestination -whatTo $itemsSorted

        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
        progressbar -state "Stop"

        #$errorLogPath = $env:TEMP+"\"+"SPCopyErrorLog_$((Get-Date).ToString('dd-MM-yyyy_hh_mm')).txt"
        #$global:errorsAll | Out-File -FilePath $errorLogPath
        #Invoke-Item  $errorLogPath
        $global:errorsAll | Out-GridView -Title "Error Log"
        $global:errorsAll = @()
        }
    }

}

function copyMetadataFromFile($targetitem, $sourceItem, $fieldsToUpdate, $whatIsCreated){
    write-host "Copying MetaData"$targetitem.FSObjType -ForegroundColor Cyan
    if($WPFcheckBox_Approve.isChecked -eq $true){
        if($destList.BaseType -eq "GenericList"){
            approveContent -item $targetitem -type "List"
        }
    }
        ####Get all Lookups and Try to Copy them as Text 
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -ne "Lookup"))}){
            if(($field.DestinationType -eq "Single line of text") -or ($field.DestinationType -eq "Multiple lines of text")){
                $lookupString=$sourceItem.($field.SourceName)
                if($lookupString){
                    $lookupString = $lookupString.split("|")[0]
                }
                Write-Host "Look Srt" $lookupToString  -ForegroundColor Magenta 
                $targetitem[$field.DestinationName]=$lookupToString
            }
        }

        ####Get all Lookups to Lookups and give them LookupID values
        $lookupFieldsArray = @()
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Lookup") -and ($_.DestinationType -eq "Lookup"))}){
            $lookupArray=$sourceItem.($field.SourceName)
            if($lookupArray){
                $lookupArray = $lookupArray.split("|")[1]
                $lookupArray = $lookupArray.split(" ")
                foreach($value in $lookupArray){
                #$stringLookup = $stringLookup+"#;"+$value

                $lookupField = new-object Microsoft.SharePoint.Client.FieldLookupValue
                $lookupField.lookupID = $value

                [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupFieldsArray += $lookupField
                }

                $targetitem[$field.DestinationName]=$lookupFieldsArray
            }
        }

        ####If Fields are not User Fields or Lookup Fields just copy
        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -ne "Person or Group") -and ($_.SourceType -ne "Lookup") -and ($_.SourceType -ne "Date and Time") -and ($_.SourceType -ne "Choice") -and ($_.SourceType -ne "Managed Metadata"))}){
            $targetitem[$field.DestinationName] = $sourceItem.($field.SourceName)
            #$targetitem.update()
        }

        ####if Fields are User Fields get them from Ensured User Array and create Fields User Array to Pass to the Destination Field
        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Person or Group"}){ 
                            $filterItOut = $allEnsuredArray | where {$_.BelongingTo -eq $field.SourceName}
                            if($filterItOut){
                                write-host "Updating User Field with users other then Owner"
                                $userValueCollection = [Microsoft.SharePoint.Client.FieldUserValue[]]$filterItOut
                                $targetitem[$field.DestinationName]=$userValueCollection
                            } else {  
                                ####Comented it out! Now not updating Field with Owner if there is not an ensured User for the field. Otherwise even fields that don't originally have users filled in get the Owner.  
                                #write-host "Updating User Field with Owner"                              
                                #$targetitem[$field.DestinationName] = $global:ownerValueCollection
                            }
                            #$targetitem.update()
        } 

        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Choice"}){
            $choiceArray = @() 
            if($sourceItem.($field.SourceName) -like "*|*"){
                $choices = $sourceItem.($field.SourceName).split("|")  
            } else {
                $choices = $sourceItem.($field.SourceName)
            }         
            foreach ($choice in $choices){
                write-host "CHOICE" $choice
                $choiceArray += $choice  
            }       
            $targetitem[$field.DestinationName] = $choiceArray
        }

        foreach ($field in $fieldsToUpdate | where {$_.SourceType -eq "Date and Time"}){
                    [datetime]$validateTime = $sourceItem.($field.SourceName)
                    write-host "Time Before" $validateTime

                    $timeSP = [System.TimeZoneInfo]::ConvertTimeFromUtc($validateTime, [System.TimeZoneInfo]::Local)

                    #$validateTime = [System.TimeZoneInfo]::ConvertTimetoUtc($validateTime)

                    write-host "Time After" $timeSP

                    write-host "field name" $field.SourceName

                    #$timeSP = $validateTime.ToString('yyyy-MM-ddThh:mm:ssZ')

                    $targetitem[$field.DestinationName] = $timeSP
        }

        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -ne "Managed Metadata"))}){
            if(($field.DestinationType -eq "Single line of text") -or ($field.DestinationType -eq "Multiple lines of text")){
                $metaString=$sourceItem.($field.SourceName)
                if($metaString){
                    $metaRegex = [regex]::Matches($metaString, '(?<=;#)(.*?)(?=\|)') |ForEach-Object { $_.Groups[1].Value }
                    $i=0
                    foreach($value in $metaRegex){
                        if($value -like "*;#*"){
                            $value = $value -replace ";#", "^"
                            $value = $value.split("^")[1]
                            $metaRegex[$i] = $value
                        }
                        $i++
                    }
                    $metaToStringAgain = [system.String]::Join(", ",$metaRegex)
                    $targetitem[$field.DestinationName] = $metaToStringAgain
                }
            }
        }

        foreach ($field in $fieldsToUpdate | where {(($_.SourceType -eq "Managed Metadata") -and ($_.DestinationType -eq "Managed Metadata"))}){
            $targetitem[$field.DestinationName] = $sourceItem.($field.SourceName)
        }
    ####Finaly Update
    $targetitem.update()

    ####Approve if Library
    if($WPFcheckBox_Approve.isChecked -eq $true){
        if($destList.BaseType -eq "DocumentLibrary"){     
            approveContent -item $targetitem -type "Library"     
        }
    }
}

function iterateFileShareSourceMetaData ($whatTo){
        progressbar -state "Start" -all $whatTo.count
        ####Get Destination List
        $destWeb = $destContext.Web
        $destList = $destContext.Web.Lists.GetByTitle($WPFlistView_Dest.SelectedItem.Title)

        $destContext.Load($destList)
        $destContext.Load($destWeb)
        $destContext.Load($destList.RootFolder)
        $destContext.Load($destList.Fields)
        $destContext.Load($destList.Views)

        $destContext.ExecuteQuery()

        ####Store All created Folders
        $fileDirRefs = @()
        $fileDirRefs += $destList.RootFolder.serverrelativeurl

        $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl

        $global:fileShareRoot = $global:metadataAll.FileRefLocal[0]
        try {
            ensureOwner
        } catch {
            logEverything -relatedItemURL "Ensure Owner" -exceptionToLog $_.exception.message
        }
        foreach ($item in $whatTo){

        progressbar -state "Plus" 

        $urlRelative = $item.FileRefLocal -replace ([RegEx]::Escape($global:fileShareRoot)) , ""
        $urlRelative = $urlRelative -replace ([RegEx]::Escape("\")), "/"
        $fileURL = $destListRelativeURL + $urlRelative
        $folderURL = $destListRelativeURL + $urlRelative  
        write-host "URL" $folderURL -ForegroundColor Green
            $destinationFolderURL = $folderURL.Substring(0, $folderURL.lastIndexOf('/'))
                
            $WPFlabel_Status.Content = "Status: Copying Folder "+$folderURL
            $Form.Dispatcher.Invoke("Background", [action]{})   
           

           write-host "ITEM FSOBJECT TYPE" $Item.FSObjType -ForegroundColor Magenta

            ####If Folder
            if($Item.FSObjType -eq 1){
                write-host "IS FOLDER" -ForegroundColor Magenta

                if($fileDirRefs -contains $destinationFolderURL){

                    $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $Item -whoIsAsking "iterateFileShareSourceMetaData"

                    write-host  "Array" $allEnsuredArray 
                    ####Upload
                    write-host "URL" $folderURL
                    write-host "Root"$destListRelativeURL
                    $upload = $destList.RootFolder.folders.Add($folderURL) 

                    write-host "Copying Metadata" -ForegroundColor DarkCyan
                    if($global:destSPTypeCheck -eq "False 2010"){
                        write-host "SP2010!"
                        $targetMeta = getSubItem -theItem $folderURL -targetLocation "Dest"
                    } else {
                        ####Get Target Item Fields
                        $targetMeta = $upload.ListItemAllFields
                    }
                    copyMetadataFromFile -targetitem $targetMeta -sourceItem $Item -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                    ####Commit
                    $destContext.Load($upload)
                    try {
                        $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $folderURL -exceptionToLog $_.exception.message
                    }
                    $fileDirRefs += $folderURL
                } else {
                    $checkFolderTrim = $destinationFolderURL
                    $folderToCreateArray = @()  

                    while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                        $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                        $obj = new-object psobject -Property @{
                            'URL' = $checkFolderTrim
                            'Count' = $folderURLCount
                        }
                        $folderToCreateArray += $obj
                        $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('/'))
                    }
                    $folderToCreateArray = $folderToCreateArray| sort-object Count

                     foreach ($one in $folderToCreateArray){
                        write-host "Creating Folder for Folder SUB" $one.URL

                        $reverseURL = $one.URL -replace $destListRelativeURL, ""
                        $reverseURL = $reverseURL -replace "/", "\"
                        $reverseURL = $global:fileShareRoot + $reverseURL
                        write-host "Reverse URL" $reverseURL
                        $sourceItemGot = $global:metadataAll | where {$_.FileRefLocal -eq $reverseURL}
                        write-host "DID WE GET IT!!??"$sourceItemGot
                        if($sourceItemGot){
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemGot -whoIsAsking "iterateFileShareSourceMetaData"
                        }
                        $upload = $destList.RootFolder.folders.Add($one.URL) 

                    if($global:destSPTypeCheck -eq "False 2010"){
                        $targetMeta = getSubItem -theItem $one.URL -targetLocation "Dest"
                    } else {
                        ####Get Target Item Fields
                        $targetMeta = $upload.ListItemAllFields
                    }
                        copyMetadataFromFile -targetitem $targetMeta -sourceItem $sourceItemGot -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"
                        ####Commit
                        $destContext.Load($upload)

                        try {
                            $destContext.ExecuteQuery()
                        } catch {
                            logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                        }
                        $fileDirRefs += $one.URL
                     }
                    $upload = $destList.RootFolder.folders.Add($folderURL)

                    ####Commit
                    $destContext.Load($upload)
                    try {
                        $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $folderURL -exceptionToLog $_.exception.message
                    }
                    $fileDirRefs += $folderURL
                }
            }
                       
            if($Item.FSObjType -eq 0){
                if($fileDirRefs -contains $destinationFolderURL){
                    write-host "File URL" $fileURL


                    $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $Item -whoIsAsking "iterateFileShareSourceMetaData"


                    ####If SharePoint is everything else but 2010
                    if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                        ####File Creation
                        $FileStream = New-Object IO.FileStream($item.FileRefLocal,[System.IO.FileMode]::Open)
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.ContentStream = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                    }
                    ####If SharePoint is 2010
                    if($global:destSPTypeCheck -eq "False 2010"){
                        ####File Creation
                        $FileStream = Get-Content $item.FileRefLocal -Encoding Byte -ReadCount 0
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.Content = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)               
                    }

                    write-host "Copying Metadata" -ForegroundColor DarkCyan
                    $targetMeta = $Upload.ListItemAllFields

                            #Handle additional Version created if Versioning is Turned On
                            $modifyVersioning = $false
                            $minorVer = $false
                            if($destList.EnableVersioning -eq $true){
                                if($destList.EnableMinorVersions -eq $true){
                                    $minorVer = $true
                                }
                            write-host "MODIFIYNG VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $false
                                $destList.update()
                                $destContext.ExecuteQuery()
                                $modifyVersioning = $true
                            }


                    copyMetadataFromFile -targetitem $targetMeta -sourceItem $Item -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Item"

                            if($modifyVersioning -eq $true){
                                write-host "TURNING ON VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $true

                                if($minorVer -eq $true){
                                    $destList.EnableMinorVersions = $true
                                }
                                $destList.update()
                            }

                    $destContext.Load($Upload)

                    try {
                            $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $fileURL -exceptionToLog $_.exception.message
                    }
                } else {
                    $checkFolderTrim = $destinationFolderURL
                    $folderToCreateArray = @()  

                    while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                        $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                        $obj = new-object psobject -Property @{
                            'URL' = $checkFolderTrim
                            'Count' = $folderURLCount
                        }
                        $folderToCreateArray += $obj
                        $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('/'))
                    }
                    $folderToCreateArray = $folderToCreateArray| sort-object Count

                     foreach ($one in $folderToCreateArray){
                        write-host "Creating Folder for Folder SUB" $one.URL

                        $reverseURL = $one.URL -replace $destListRelativeURL, ""
                        $reverseURL = $reverseURL -replace "/", "\"
                        $reverseURL = $global:fileShareRoot + $reverseURL
                        write-host "Reverse URL" $reverseURL
                        $sourceItemGot = $global:metadataAll | where {$_.FileRefLocal -eq $reverseURL}
                        write-host "DID WE GET IT!!??"$sourceItemGot
                        if($sourceItemGot){
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemGot -whoIsAsking "iterateFileShareSourceMetaData"
                        }
                        $upload = $destList.RootFolder.folders.Add($one.URL) 

                    if($global:destSPTypeCheck -eq "False 2010"){
                        $targetMeta = getSubItem -theItem $one.URL -targetLocation "Dest"
                    } else {
                        ####Get Target Item Fields
                        $targetMeta = $upload.ListItemAllFields
                    }
                        copyMetadataFromFile -targetitem $targetMeta -sourceItem $sourceItemGot -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                        ####Commit
                        $destContext.Load($upload)
                        try {
                            $destContext.ExecuteQuery()
                        } catch {
                            logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                        }

                        $fileDirRefs += $one.URL
                     }   
                    ####If SharePoint is everything else but 2010
                    if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                        ####File Creation
                        $FileStream = New-Object IO.FileStream($item.FileRefLocal,[System.IO.FileMode]::Open)
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.ContentStream = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                    }
                    ####If SharePoint is 2010
                    if($global:destSPTypeCheck -eq "False 2010"){
                        ####File Creation
                        $FileStream = Get-Content $item.FileRefLocal -Encoding Byte -ReadCount 0
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.Content = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)               
                    }

                    write-host "Copying Metadata" -ForegroundColor DarkCyan
                    $targetMeta = $Upload.ListItemAllFields

                            #Handle additional Version created if Versioning is Turned On
                            $modifyVersioning = $false
                            $minorVer = $false
                            if($destList.EnableVersioning -eq $true){
                                if($destList.EnableMinorVersions -eq $true){
                                    $minorVer = $true
                                }
                            write-host "MODIFIYNG VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $false
                                $destList.update()
                                $destContext.ExecuteQuery()
                                $modifyVersioning = $true
                            }

                    copyMetadataFromFile -targetitem $targetMeta -sourceItem $Item -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Item"

                            if($modifyVersioning -eq $true){
                                write-host "TURNING ON VERSIONING" -ForegroundColor Magenta
                                $destList.EnableVersioning = $true

                                if($minorVer -eq $true){
                                    $destList.EnableMinorVersions = $true
                                }
                                $destList.update()
                            }

                    $destContext.Load($Upload)

                    try {
                        $destContext.ExecuteQuery()   
                    } catch {
                        logEverything -relatedItemURL $fileURL -exceptionToLog $_.exception.message
                    }                    
                }
            }
        }
}

function iterateFileShareSourceBasic ($whatTo){
        progressbar -state "Start" -all $whatTo.count
        ####Get Destination List
        $destWeb = $destContext.Web
        $destList = $destContext.Web.Lists.GetByTitle($WPFlistView_Dest.SelectedItem.Title)

        $destContext.Load($destList)
        $destContext.Load($destWeb)
        $destContext.Load($destList.RootFolder)
        $destContext.Load($destList.Fields)
        $destContext.Load($destList.Views)

        $destContext.ExecuteQuery()

        ####Store All created Folders
        $fileDirRefs = @()
        $fileDirRefs += $destList.RootFolder.serverrelativeurl

        $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl

        foreach ($item in $whatTo){
        progressbar -state "Plus" 
        $urlRelative = $item.FullName -replace ([RegEx]::Escape($global:fileShareRoot)) , ""
        $urlRelative = $urlRelative -replace ([RegEx]::Escape("\")), "/"
        $item | Add-Member -type NoteProperty -Name 'URLlistRelative' -Value $urlRelative 
        $fileURL = $destListRelativeURL + $item.URLlistRelative
        $folderURL = $destListRelativeURL + $item.URLlistRelative   
        write-host "URL" $folderURL -ForegroundColor Green
            $destinationFolderURL = $folderURL.Substring(0, $folderURL.lastIndexOf('/'))
                
            $WPFlabel_Status.Content = "Status: Copying Folder "+$folderURL
            $Form.Dispatcher.Invoke("Background", [action]{})   

            if($item.PSIsContainer){
                if($fileDirRefs -contains $destinationFolderURL){
                    ####Upload
                    write-host "URL" $folderURL
                    write-host "Root"$destListRelativeURL
                    $upload = $destList.RootFolder.folders.Add($folderURL) 
                    ####Commit
                    $destContext.Load($upload)
                    try {
                        $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $folderURL -exceptionToLog $_.exception.message
                    }
                    $fileDirRefs += $folderURL
                } else {
                    $checkFolderTrim = $destinationFolderURL
                    $folderToCreateArray = @()  

                    while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                        $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                        $obj = new-object psobject -Property @{
                            'URL' = $checkFolderTrim
                            'Count' = $folderURLCount
                        }
                        $folderToCreateArray += $obj
                        $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('/'))
                    }
                    $folderToCreateArray = $folderToCreateArray| sort-object Count

                     foreach ($one in $folderToCreateArray){
                        write-host "Creating Folder for Folder SUB" $one.URL
                        $upload = $destList.RootFolder.folders.Add($one.URL) 

                        ####Commit
                        $destContext.Load($upload)

                        try {
                            $destContext.ExecuteQuery()
                        } catch {
                            logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                        }
                        $fileDirRefs += $one.URL
                     }
                    $upload = $destList.RootFolder.folders.Add($folderURL) 
                    ####Commit
                    $destContext.Load($upload)
                    try {
                        $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $folderURL -exceptionToLog $_.exception.message
                    }
                    $fileDirRefs += $folderURL
                }
            }else{
                if($fileDirRefs -contains $destinationFolderURL){
                    write-host "File URL" $fileURL
                    if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                        ####File Creation
                        $FileStream = New-Object IO.FileStream($item.FullName,[System.IO.FileMode]::Open)
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.ContentStream = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                    }
                    ####If SharePoint is 2010
                    if($global:destSPTypeCheck -eq "False 2010"){
                        ####File Creation
                        $FileStream = Get-Content $item.FullName -Encoding Byte -ReadCount 0
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.Content = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)               
                    }
                    if($WPFcheckBox_Approve.isChecked -eq $true){
                        if($destList.BaseType -eq "DocumentLibrary"){  
                            $targetMeta = $Upload.ListItemAllFields     
                            approveContent -item $targetMeta -type "Library"     
                        }
                    }
                    $destContext.Load($Upload)

                    try {
                            $destContext.ExecuteQuery()
                    } catch {
                        logEverything -relatedItemURL $fileURL -exceptionToLog $_.exception.message
                    }
                } else {
                    $checkFolderTrim = $destinationFolderURL
                    $folderToCreateArray = @()  

                    while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                        $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '/'} | Measure-Object).Count
                        $obj = new-object psobject -Property @{
                            'URL' = $checkFolderTrim
                            'Count' = $folderURLCount
                        }
                        $folderToCreateArray += $obj
                        $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('/'))
                    }
                    $folderToCreateArray = $folderToCreateArray| sort-object Count

                     foreach ($one in $folderToCreateArray){
                        write-host "Creating Folder for Folder SUB" $one.URL

                        $reverseURL = $one.URL -replace $destListRelativeURL, ""
                        $reverseURL = $reverseURL -replace "/", "\"
                        $reverseURL = $global:fileShareRoot + $reverseURL
                        write-host "Reverse URL" $reverseURL
                        $sourceItemGot = $global:metadataAll | where {$_.FileRefLocal -eq $reverseURL}
                        write-host "DID WE GET IT!!??"$sourceItemGot
                        if($sourceItem){
                            $allEnsuredArray = preEnsureUsers -sourceItemFuncEnsure $sourceItemGot -whoIsAsking "iterateFileShareSourceMetaData"
                        }
                        $upload = $destList.RootFolder.folders.Add($one.URL) 

                    if($global:destSPTypeCheck -eq "False 2010"){
                        $targetMeta = getSubItem -theItem $one.URL -targetLocation "Dest"
                    } else {
                        ####Get Target Item Fields
                        $targetMeta = $upload.ListItemAllFields
                    }
                        copyMetadataFromFile -targetitem $targetMeta -sourceItem $sourceItemGot -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"


                        ####Commit
                        $destContext.Load($upload)
                        try {
                            $destContext.ExecuteQuery()
                        } catch {
                            logEverything -relatedItemURL $one.URL -exceptionToLog $_.exception.message
                        }

                        $fileDirRefs += $one.URL
                     }   
                    ####File Creation

                    write-host "URL"$fileURL -ForegroundColor Green
                    if(($global:destSPTypeCheck -eq "False") -or ($global:destSPTypeCheck -eq "True")){
                        ####File Creation
                        $FileStream = New-Object IO.FileStream($item.FullName,[System.IO.FileMode]::Open)
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.ContentStream = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)
                    }
                    ####If SharePoint is 2010
                    if($global:destSPTypeCheck -eq "False 2010"){
                        ####File Creation
                        $FileStream = Get-Content $item.FullName -Encoding Byte -ReadCount 0
                        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                        $FileCreationInfo.Overwrite = $true
                        $FileCreationInfo.Content = $FileStream
                        $FileCreationInfo.URL = $fileURL
                        $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)               
                    }      
                    if($WPFcheckBox_Approve.isChecked -eq $true){
                        if($destList.BaseType -eq "DocumentLibrary"){   
                            $targetMeta = $Upload.ListItemAllFields  
                            approveContent -item $targetMeta -type "Library"     
                        }
                    }
                    $destContext.Load($Upload)

                    try {
                        $destContext.ExecuteQuery()   
                    } catch {
                        logEverything -relatedItemURL $fileURL -exceptionToLog $_.exception.message
                    }                    
                }
            }
        }
}

function exportMetaData ($theItem, $theDestinationURL){
                    $objFields = new-object psobject
                    foreach($field in $WPFlistView_Fields_Final.ItemsSource) {
                        if($field.SourceType -eq "Lookup"){ 
                            $lookupArray=$theItem[$field.SourceName].LookupValue   
                            if($lookupArray){
                                $lookupToString = [system.String]::Join(", ",$lookupArray)
                                [string]$stringValue = $lookupToString + "|" + $theItem[$field.SourceName].LookupID
                                $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $stringValue
                            }
                        }

                        if($field.SourceType -eq "Person or Group"){
                            write-host "In Person" -ForegroundColor Green   
                            $allUsersEnsured = @()
                            foreach($user in $theItem[$field.SourceName]){
                                if($user){
                                    ####If SP is On-Premise ensure by Lookup Value. If it is Online ensure by Mail.
                                    if(($global:sourceSPTypeCheck -eq "False") -or ($global:sourceSPTypeCheck -eq "False 2010")){
                                        write-host "Type Check On-Prem"
                                        $userSourceEnsured = $Web.EnsureUser($user.LookupValue)
                                    }
                                    if($global:sourceSPTypeCheck -eq "True"){
                                        write-host "Type Check Online"
                                        $userSourceEnsured = $Web.EnsureUser($user.Email)
                                    }
                                    $Context.load($userSourceEnsured)
                                    $Context.ExecuteQuery()
                                    $userRegex = $userSourceEnsured.LoginName

                                        #$userRegex = $userRegex.split("|")[-1]

                                    $allUsersEnsured+=$userRegex
                                }
                            }
                            $usersToString = [system.String]::Join(", ",$allUsersEnsured)
                            write-host "Users String"  $usersToString
                            $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $usersToString
                        }

                        if($field.SourceType -eq "Date and Time"){                         
                            #[datetime]$validateTime = $theItem[$field.SourceName]
                            #$time = $validateTime.ToString('yyyy-MM-dd hh:mm:ss')    
                            #$objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $time  
                            $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $theItem[$field.SourceName]                
                        }

                        if($field.SourceType -eq "Choice"){
                            $choiceArray = @()
                            foreach($choice in $theItem[$field.SourceName]){
                                $choiceArray += $choice   
                            }
                            $choiceArray = [system.String]::Join("|",$choiceArray)

                            $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $choiceArray   
                        }

                        if($field.SourceType -eq "Managed Metadata"){
                            $metaValue = ""
                            foreach($meta in $theItem[$field.SourceName]){
                                write-host "Managed Meta Single!!!!" -ForegroundColor Green
                                $labelMeta = $meta.Label
                                $guidMeta = $meta.TermGuid
                                $idMeta = $meta.$item.WssId
                                if(!($metaValue)){
                                    $metaValue = "-1;#$labelMeta|$guidMeta"
                                } else {
                                    $metaValue += ";#-1;#$labelMeta|$guidMeta"
                                }
                            }  
                             $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $metaValue                                  
                        }

                        if(($field.SourceType -ne "Person or Group") -and ($field.SourceType -ne "Lookup") -and ($field.SourceType -ne "Date and Time") -and ($field.SourceType -ne "Choice") -and ($field.SourceType -ne "Managed Metadata")){
                            $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $theItem[$field.SourceName]
                        }
                        
                    }
                    $objFields | Add-Member -type NoteProperty -Name "FileRef" -Value $theItem["FileRef"]
                    $objFields | Add-Member -type NoteProperty -Name "FileRefLocal" -Value $theDestinationURL
                    $objFields | Add-Member -type NoteProperty -Name "FileDirRef" -Value $theItem["FileDirRef"]
                    $objFields | Add-Member -type NoteProperty -Name "FSObjType" -Value $theItem["FSObjType"]
                    $objFields | Add-Member -type NoteProperty -Name "BaseType" -Value "BaseType"

                    $global:fieldValuesArrayForCSV += $objFields
}

function iterateFileShareDestination ($whatTo){

            ####Store Field Values
            $global:fieldValuesArrayForCSV = @()

            $objFields = new-object psobject
            foreach($field in $WPFlistView_Fields_Final.ItemsSource) {
                $objFields | Add-Member -type NoteProperty -Name $field.SourceName -Value $field.SourceType
            }
            $objFields | Add-Member -type NoteProperty -Name "FileRef" -Value $sourceListTitle
            $objFields | Add-Member -type NoteProperty -Name "FileDirRef" -Value $list.RootFolder.serverrelativeurl
            $objFields | Add-Member -type NoteProperty -Name "FileRefLocal" -Value $global:fileShareRoot
            $objFields | Add-Member -type NoteProperty -Name "FSObjType" -Value "Type"

            $global:fieldValuesArrayForCSV += $objFields

            ####Store All created Folders
            $fileDirRefs = @()

            $listRelativeURL = $list.RootFolder.ServerRelativeUrl
                       
            ####Add List Root in the created Folders List
            $fileDirRefs += $global:fileShareRoot

            ####Handle User Fields
            write-host "Pre-Ensuring Owner"

            ####Ensure Owner. If users in user fields cannot be ensured the User wich is used for Destination logging will be used.
            #ensureOwner

            ####Start Iteration
            foreach ($sourceItem in $whatTo){
                progressbar -state "Plus"
                ####if Folder
                if($sourceItem["FSObjType"] -eq 1){
                    $WPFlabel_Status.Content = "Status: Copying Folder "+$sourceItem["FileRef"]
                    $Form.Dispatcher.Invoke("Background", [action]{})

                    write-host "Folder to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan

                    ####Construct File Name and URL
                    $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
                    $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $global:fileShareRoot
                    $destinationFileURL = $destinationFileURL -replace "/", "\"
                    $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('\'))

                    exportMetaData -theItem $sourceItem -theDestinationURL $destinationFileURL

                        if($fileDirRefs -contains $destinationFolderURL){
                            try {
                                New-Item -ItemType directory -Path $destinationFileURL
                            } catch {
                                write-host "Dir already exists."
                            }
                            $fileDirRefs += $destinationFileURL
                        } else {
                            $checkFolderTrim = $destinationFolderURL
                            write-host "CHK FLD TRM"  $checkFolderTrim
                            $folderToCreateArray = @()    
                            while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $global:fileShareRoot)){
                            $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
                            $obj = new-object psobject -Property @{
                                'URL' = $checkFolderTrim
                                'Count' = $folderURLCount
                            }
                                $folderToCreateArray += $obj
                                $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('\'))
                            }
                            $folderToCreateArray = $folderToCreateArray| sort-object Count

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace ([RegEx]::Escape($global:fileShareRoot)), $listRelativeURL
                                $sourceURLOfTheTrim = $sourceURLOfTheTrim -replace ([RegEx]::Escape("\")), "/"
                                write-host "Reverse URL Folder" $sourceURLOfTheTrim
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"
                                exportMetaData -theItem $sourceItemSub -theDestinationURL $one.URL

                                New-Item -ItemType directory -Path $one.URL
                                $fileDirRefs += $one.URL                        
                            }

                                New-Item -ItemType directory -Path $destinationFileURL
                                $fileDirRefs += $destinationFileURL                                
                    }
                }

                ####if File
                if($sourceItem["FSObjType"] -eq 0){
                    $WPFlabel_Status.Content = "Status: Copying File "+$sourceItem["FileRef"]
                    $Form.Dispatcher.Invoke("Background", [action]{})

                    write-host "File to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan
              
                    ####Construct File Name and URL
                    $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
 
                    $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $global:fileShareRoot
                    $destinationFileURL = $destinationFileURL -replace "/", "\"

                    $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('\'))

                    exportMetaData -theItem $sourceItem -theDestinationURL $destinationFileURL

                    if($list.BaseType -eq "DocumentLibrary"){                   
                        if($fileDirRefs -contains $destinationFolderURL){
                            [Microsoft.SharePoint.Client.FileInformation]$fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$sourceItem["FileRef"]);
                            [System.IO.FileStream]$writeStream = [System.IO.File]::Open($destinationFileURL,[System.IO.FileMode]::Create);
                            $fileInfo.Stream.CopyTo($writeStream);
                            $writeStream.Close();
                        } else {
                            $checkFolderTrim = $destinationFolderURL
                            $folderToCreateArray = @()    
                            while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $global:fileShareRoot)){
                            $folderURLCount = ($checkFolderTrim.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
                            $obj = new-object psobject -Property @{
                                'URL' = $checkFolderTrim
                                'Count' = $folderURLCount
                            }
                                $folderToCreateArray += $obj
                                $checkFolderTrim = $checkFolderTrim.Substring(0, $checkFolderTrim.lastIndexOf('\'))
                            }
                            $folderToCreateArray = $folderToCreateArray| sort-object Count

                            foreach ($one in $folderToCreateArray){
                                $sourceURLOfTheTrim = $one.URL -replace ([RegEx]::Escape($global:fileShareRoot)), $listRelativeURL
                                $sourceURLOfTheTrim = $sourceURLOfTheTrim -replace ([RegEx]::Escape("\")), "/"
                                write-host "Reverse URL Item" $sourceURLOfTheTrim
                                $sourceItemSub = getSubItem -theItem $sourceURLOfTheTrim -targetLocation "Source"
                                exportMetaData -theItem $sourceItemSub -theDestinationURL $one.URL
                                New-Item -ItemType directory -Path $one.URL
                                $fileDirRefs += $one.URL                        
                            }
                                [Microsoft.SharePoint.Client.FileInformation]$fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context,$sourceItem["FileRef"]);
                                [System.IO.FileStream]$writeStream = [System.IO.File]::Open($destinationFileURL,[System.IO.FileMode]::Create);
                                $fileInfo.Stream.CopyTo($writeStream);
                                $writeStream.Close();
                        }
                    }
                }
            }

            #$global:fieldValuesArrayForCSV | Out-GridView 

            $csvPath = $global:fileShareRoot + "\"+"Metadata.csv"
            $global:fieldValuesArrayForCSV | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }

$WPFbutton_Copy_NoFilter.Add_Click({
    if((($global:sourceSPTypeCheck -eq "File") -or ($global:sourceSPTypeCheck -eq "Meta")) -or ($global:destSPTypeCheck -eq "File")){
        copyFunctionFileShare
    } else {
        copyFunction
    }
})

####Routed Event Handlers for ListView Sorting.These refuse to work inside a function in any way, so it's a copy/paste ordeal for each one.
$RoutedSourceLists= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc1 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc1 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc1 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc1)
        }
    }
$eventSourceLists = [Windows.RoutedEventHandler]$RoutedSourceLists
$WPFlistView.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventSourceLists)

$RoutedDestLists= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_Dest.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc2 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc2 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc2 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc2)
            }
    }
$eventDestLists = [Windows.RoutedEventHandler]$RoutedDestLists
$WPFlistView_Dest.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventDestLists)

$RoutedSourceFields= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_Fields_Source.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc3 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc3 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc3 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc3)
        }
    }
$eventSourceFields = [Windows.RoutedEventHandler]$RoutedSourceFields
$WPFlistView_Fields_Source.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventSourceFields)

$RoutedDestFields= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_Fields_Dest.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc4 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc4 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc4 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc4)
            }
    }
$eventDestFields = [Windows.RoutedEventHandler]$RoutedDestFields
$WPFlistView_Fields_Dest.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventDestFields)

$RoutedFieldsFinal= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_Fields_Final.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc5 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc5 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc5 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc5)
        }
    }
$eventFieldsFinal = [Windows.RoutedEventHandler]$RoutedFieldsFinal
$WPFlistView_Fields_Final.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventFieldsFinal)

$RoutedBrowser= {
	    $col = $_.OriginalSource.Column.Header
        if($col){
	        $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFlistView_Browser.ItemsSource)
	        $view.SortDescriptions.Clear()

            if ($script:sort -eq 'descending'){
                $desc6 = New-Object System.ComponentModel.SortDescription($col,'Ascending')
                $script:sort = 'ascending'
            }
            ElseIf ($script:sort -eq 'ascending')  {
                $desc6 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending'
            }
            Else {
                $desc6 = New-Object System.ComponentModel.SortDescription($col,'Descending')
                $script:sort = 'descending' 
            }  
            $view.SortDescriptions.Add($desc6)
        }
    }
$eventBrowser = [Windows.RoutedEventHandler]$RoutedBrowser
$WPFlistView_Browser.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $eventBrowser)


####Initiate Tips
#$WPFimage_Tips.source = $dirFormURLs+"/Site_Col.png"


function InitialTip {
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("Form",$form)
    $Runspace.SessionStateProxy.SetVariable("TargetBox",$WPFimage_Initial)
 
    $code = {
        Start-Sleep -Seconds 3
        $form.Dispatcher.invoke(
        [action]{$TargetBox.Visibility = "Hidden" })
    }

    $PSinstance = [powershell]::Create().AddScript($Code)
    $PSinstance.Runspace = $Runspace
    $job = $PSinstance.BeginInvoke()
}
$global:check = "Run"
$WPFmain.add_mouseEnter({
    if(($global:check -eq "Run")){
        InitialTip
        $global:check = "Passed"
        write-host "Triggered!"
    } 
})
################################################
## Show the form
################################################
write-host "To show the form, run the following" -ForegroundColor Cyan

$Form.ShowDialog() | out-null

#[void]$Form.Dispatcher.InvokeAsync{$Form.ShowDialog()}.Wait()

EXIT
