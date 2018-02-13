####Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted; Get-ExecutionPolicy
####Hide Powershell Console.
Function Hide-PowerShellWindow()
{
[CmdletBinding()]
param (
[IntPtr]$Handle=$(Get-Process -id $PID).MainWindowHandle
)
$WindowDisplay = @"
using System;
using System.Runtime.InteropServices;

namespace Window
{
public class Display
{
[DllImport("user32.dll")]
private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

public static bool Hide(IntPtr hWnd)
{
return ShowWindowAsync(hWnd, 0);
}
}
}
"@
Try
{
Add-Type -TypeDefinition $WindowDisplay
[Window.Display]::Hide($Handle)
}
Catch
{
}
}

####Global Variables. Some globals are not kept in Global Variables but in .tag in buttons.
$global:selectedItemsForCopy = @()
$global:itemsToCopy = @()
$global:copyMode = ""
$global:sourceSPTypeCheck = ""
$global:destSPTypeCheck = ""

####Call Hide PowerShell Console
#[Void]$(Hide-PowerShellWindow)

####Get script location
#$scriptpath = $MyInvocation.MyCommand.Path
#$dir = Split-Path $scriptpath

####Add CSOM DLLs
Add-Type -Path "C:\Desktop\Temp\PowerShell_FORMS\New_Exe\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Desktop\Temp\PowerShell_FORMS\New_Exe\Microsoft.SharePoint.Client.Runtime.dll" 

####NOT NEEDED. JUST FOR REFERENCE
#Add-Type -Path "C:\Desktop\Temp\PowerShell_FORMS\WpfAnimatedGif.dll"

####Form XAML
$inputXML = get-content "C:\Desktop\Temp\PowerShell_FORMS\New_Exe\MainWindow.xaml"  
 

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

################################################
## UI Related Functions
################################################

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
        while ($i -lt 10){
            $grid.Opacity = "0."+$i
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
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User,$Pass)
    }
    if($spType -eq "False"){
        $Creds = New-Object System.Net.NetworkCredential($User,$Pass)
    }
    return $Creds
}

####Gets Source and Destination Lists
function getLists ($User, $Password, $SiteURL, $listView,$spType,$SourceOrTarget){
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
    }

    $allListsArray = @()
    foreach ($list in $lists){
        $obj = new-object psobject -Property @{
            'Title' = $list.Title
            'Type' = $list.BaseType
        }
        $allListsArray+=$obj
        #$listView.AddChild($obj)
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
 }

 ####Gets Fields in Source and Destination List
function getFields ($User, $Password, $SiteURL, $docLibName, $listView, $spType, $SourceOrTarget){
    $Creds = createCredentials -User $User -Password $Password -spType $spType
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $Context.Credentials = $Creds

    progressbar -state "Start" -all 2
    progressbar -state "Plus" 

    ####Get Source List
    $List = $Context.Web.Lists.GetByTitle($DocLibName)
    $fields = $list.Fields
    $Context.Load($List)
    $Context.Load($fields)
    $Context.ExecuteQuery()

    $allFieldsArray = @()

    foreach ($field in $fields | where {$_.Hidden -eq $False}){       
        $obj = new-object psobject -Property @{
        'Name' = $field.InternalName
        'Type' = $field.TypeDisplayName 
        'BaseType' = $field.FromBaseType
        } 
        #$listView.AddChild($obj)
        $allFieldsArray += $obj
    }
    $listView.ItemsSource = $allFieldsArray
    progressbar -state "Plus" 
    progressbar -state "Stop" 
    return $allFieldsArray
 }

################################################
## Copy Related Functions
################################################

####Called by iterateEverything
function copyMetadata($targetitem, $sourceItem, $fieldsToUpdate, $whatIsCreated){
write-host "Target FSOType"$targetitem["FSObjType"] -ForegroundColor Cyan
if(($sourceItem["FSObjType"] -eq 0) -and ($whatIsCreated -eq "Folder")){
    write-host "It's working" -ForegroundColor Green

    ####Get Dir of Item
    $sourceRefForCAML = $sourceItem["FileDirRef"]
    $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
    $qry.viewXML = @"
<View Scope="All">
    <Query>
		<Where>
			<Eq>
				<FieldRef Name='FileRef' /><Value Type='Text'>$sourceRefForCAML</Value>
			</Eq>
		</Where>
    </Query>
</View>
"@

    [Microsoft.SharePoint.Client.ListItemCollection]$sourceItemFolder = $list.GetItems($qry)
    $Context.Load($sourceItemFolder)
    $Context.ExecuteQuery()

    foreach ($field in $fieldsToUpdate){
        try{
            $targetitem[$field.SourceName] = $sourceItemFolder[0][$field.DestinationName]
        }catch{
            write-host $_.exception.message -foreground yellow
        }
    }
    $targetitem.update()
} else {
        foreach ($field in $fieldsToUpdate){
            try{
                $targetitem[$field.SourceName] = $sourceItem[$field.DestinationName]
            }catch{
                write-host $_.exception.message -foreground yellow
            }
        }
        $targetitem.update()
    }
}

####Called by iterateEverything
function createFolderLibrary ($destinationFileURL){

                        write-host "Destination "$destinationFileURL -ForegroundColor DarkCyan

                        ####Upload
                        $upload = $destList.RootFolder.folders.Add($destinationFileURL)

                        write-host "Copying Metadata" -ForegroundColor DarkCyan
                    
                        ####Get Target Item Fields
                        $targetMeta = $upload.ListItemAllFields

                        copyMetadata -targetitem $targetMeta -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                        ####Commit
                        $destContext.Load($upload)
                        $destContext.ExecuteQuery()

                        write-host "Folder Copied" -ForegroundColor Green
                        }

####Called by iterateEverything
function createFolderList ($destinationFileURL){
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

                        copyMetadata -targetitem $Upload -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Folder"

                        ####Commit
                        $destContext.Load($Upload)
                        $destContext.ExecuteQuery()
                     
                        write-host "Folder Didn't Exist. Folder Copied" -ForegroundColor Cyan
                        }

####Called by iterateEverything
function createFile {

                            ####Open source
                            $sourceFile = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Context, $sourceItem["FileRef"])

                            write-host "Destination"$destinationFileURL -ForegroundColor DarkCyan
                            ####Create IO Stream from Net Connection Stream
                            $memoryStream = New-Object System.IO.MemoryStream
                            $sourceFile.stream.copyTo($memoryStream)
                            $memoryStream.Seek(0, [System.IO.SeekOrigin]::Begin)

                            ####Create Creation Info and Upload
                            $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
                            $FileCreationInfo.Overwrite = $true
                            $FileCreationInfo.ContentStream = $memoryStream
                            $FileCreationInfo.URL = $destinationFileURL
                            $Upload = $destList.RootFolder.Files.Add($FileCreationInfo)

                            ####Get Target Item Fields
                            $targetMeta = $Upload.ListItemAllFields

                            copyMetadata -targetitem $targetMeta -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "File"

                            ####Commit
                            $destContext.Load($Upload)
                            $destContext.ExecuteQuery()

                            write-host "File Copied" -ForegroundColor Green
                        }

####Called by iterateEverything
function createItem {
                            ####Create Creation Info and Create
                            $FileCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
                            $FileCreationInfo.FolderURL = $destinationFolderURL
                            $FileCreationInfo.LeafName = $destinationFileName
                            $Upload = $destList.AddItem($FileCreationInfo)

                            write-host "Copying Metadata" -ForegroundColor DarkCyan

                            copyMetadata -targetitem $Upload -sourceItem $sourceItem -fieldsToUpdate $WPFlistView_Fields_Final.items -whatIsCreated "Item"

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
                        }
                            write-host "File Copied" -ForegroundColor Green
                        }

 ####Core Copy Function. 
function iterateEverything ($whatTo){
            ####Store All created Folders
            $fileDirRefs = @()

            $listRelativeURL = $list.RootFolder.ServerRelativeUrl
            $destListRelativeURL = $destList.RootFolder.ServerRelativeUrl

            ####Add List Root in the created Folders List
            $fileDirRefs += $destList.RootFolder.serverrelativeurl
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
                            createFolderLibrary $destinationFileURL
                            $fileDirRefs += $destinationFileURL
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
                                createFolderLibrary $one.URL
                                $fileDirRefs += $one.URL                        
                            }

                            createFolderLibrary $destinationFileURL
                            $fileDirRefs += $destinationFileURL
                        }
                    }

                    ####if Folder in List 
                    if($list.BaseType -eq "GenericList"){
                        if($fileDirRefs -contains $destinationFolderURL){
                            createFolderList $destinationFileURL 
                            $fileDirRefs += $destinationFileURL
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
                                createFolderList $one.URL
                                $fileDirRefs += $one.URL                        
                            }

                            createFolderList $destinationFileURL 
                            $fileDirRefs += $destinationFolderURL
                        }
                    }
                }

                ####if Item
                if($sourceItem["FSObjType"] -eq 0){

                    ####If File in Document Library (Copy File)
                    if($list.BaseType -eq "DocumentLibrary"){
                        $WPFlabel_Status.Content = "Status: Copying Item "+$sourceItem["FileRef"]
                        $Form.Dispatcher.Invoke("Background", [action]{})

                        write-host "File to be copied"$sourceItem["FileRef"] -ForegroundColor DarkCyan

                        ####Construct File Name and URL
                        $destinationFileName = $sourceItem["FileRef"].split("/")[-1]
                        $destinationFileURL = $sourceItem["FileRef"] -replace $listRelativeURL, $destListRelativeURL
                        $destinationFolderURL = $destinationFileURL.Substring(0, $destinationFileURL.lastIndexOf('/'))

                        if($fileDirRefs -contains $destinationFolderURL){
                            CreateFile 
                        } else {
                               Write-Host "The Trim" $destinationFolderURL

                            $checkFolderTrim = $destinationFolderURL
                            $folderToCreateArray = @()    
                            while ((!($fileDirRefs -contains $checkFolderTrim)) -and ($checkFolderTrim -ne $destList.RootFolder.serverrelativeurl)){
                             Write-Host "File Dir Refs" $fileDirRefs
                             Write-Host "checkFolderTrim" $checkFolderTrim


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
                                createFolderLibrary $one.URL
                                $fileDirRefs += $one.URL                        
                            }
                            createFile
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
                        write-host "FITEDIR REFS" $fileDirRefs -ForegroundColor Green
                        if($fileDirRefs -contains $destinationFolderURL){
                            createItem
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
                                createFolderList $one.URL
                                $fileDirRefs += $one.URL                        
                            }
                            createItem
                        }
                    }
                }
            }
        }

####used to copy Full Library/List. Calls iterateEverything.
function copyLibrary ($siteURL,$docLibName,$sourceSPType,$destSiteURL,$destDocLibName,$destSPType){
    ####Check if Lists are Input
    if((!($docLibName)) -or (!($destDocLibName))){
        $WPFlabel_Status.Content = "Status: Lists are not selected! Select Lists and try again!"
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
        $destContext.Load($destList)
        $destContext.Load($destList.RootFolder)

        ####Check if Applying Filter and if not go on and copy all

        ####Create CAML Query For Items
        $qryItems = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()

        ####Load
        [Microsoft.SharePoint.Client.ListItemCollection]$items = $list.GetItems($qryItems)
        $Context.Load($items)

        ####Commit
        $Context.ExecuteQuery()
        $destContext.ExecuteQuery()

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

        iterateEverything -whatTo $itemsSorted

        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
}

################################################
## Button Actions
################################################

################################################Source and estination SharePoint Version Radion Buttons
####Source SP Version Online Radio
$WPFradioButton_SourceSPVer_Online.Add_Click({
    $global:sourceSPTypeCheck = "True"
})

####Source SP Version On-Premise Radio
$WPFradioButton_SourceSPVer_Premise.Add_Click({
    $global:sourceSPTypeCheck = "False"
})

####Destination SP Version Online Radio
$WPFradioButton_DestSPVer_Online.Add_Click({
    $global:destSPTypeCheck = "True"
})

####Destination SP Version On-Premise Radio
$WPFradioButton_DestSPVer_Premise.Add_Click({
    $global:destSPTypeCheck = "False"
})

################################################Get Source and Destination Library Lists and Fields Lists Buttons

####Back Button
$WPFBack_Button_One.Add_Click({
    if($WPFsourceListsAll.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFsourceListsAll -action "Close"
        opacityAnimation -grid $WPFSourceSiteCol -action "Open"
        opacityAnimation -grid $WPFBack_Button_One -action "Close"
    }
    if($WPFlistFieldsSource.Visibility -eq "Visible"){
        if($WPFlistFieldsCombine.Visibility -eq "Visible"){
            opacityAnimation -grid $WPFlistFieldsCombine -action "Close"
            $WPFlistView_Fields_Final.ItemsSource = @()
        }
        opacityAnimation -grid $WPFlistFieldsSource -action "Close"
        opacityAnimation -grid $WPFsourceListsAll -action "Open"
    }
})

####Back Button
$WPFBack_Button_Two.Add_Click({
    if($WPFdestListsAll.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFdestListsAll -action "Close"
        opacityAnimation -grid $WPFdestSiteCol -action "Open"
        opacityAnimation -grid $WPFBack_Button_Two -action "Close"
    }
    if($WPFlistFieldsDest.Visibility -eq "Visible"){
        if($WPFlistFieldsCombine.Visibility -eq "Visible"){
            opacityAnimation -grid $WPFlistFieldsCombine -action "Close"
            $WPFlistView_Fields_Final.ItemsSource = @()
        }
        opacityAnimation -grid $WPFlistFieldsDest -action "Close"
        opacityAnimation -grid $WPFdestListsAll -action "Open"

    }
})

####Get Source List Button
$WPFbutton_GetSource.Add_Click({
if(($WPFtextBox_User.text) -and ($WPFtextBox_Pass.text) -and ($WPFtextBox_URL.text) -and ($global:sourceSPTypeCheck)){
    $WPFlabel_Status.Content = "Status: Getting All Source Lists and Libraries"
    opacityAnimation -grid $WPFsourceListsAll -action "Pre"
    getLists -User $WPFtextBox_User.text -Password $WPFtextBox_Pass.text -SiteURL $WPFtextBox_URL.text -listView $WPFlistView -spType $global:sourceSPTypeCheck -SourceOrTarget "Source"
    $WPFlabel_Status.Content = "Status: Idle"
    ####Close and Open
    opacityAnimation -grid $WPFsourceSiteCol -action "Close"
    opacityAnimation -grid $WPFsourceListsAll -action "Open"
    if($WPFBack_Button_One.Visibility -eq "Hidden"){
        opacityAnimation -grid $WPFBack_Button_One -action "Open"
    }
} else {
    $WPFlabel_Status.Content = "Status: Fill all Fields!"
    $Form.Dispatcher.Invoke("Background", [action]{})
    }
})

####Get Dest List Button
$WPFbutton_GetDest.Add_Click({
if(($WPFtextBox_User_Dest.text) -and ($WPFtextBox_Pass_Dest.text) -and ($WPFtextBox_URL_Dest.text) -and ($global:destSPTypeCheck)){
    $WPFlabel_Status.Content = "Status: Getting All Destination Lists and Libraries"
    opacityAnimation -grid $WPFdestListsAll -action "Pre"
    getLists -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.text -SiteURL $WPFtextBox_URL_Dest.text -listView $WPFlistView_Dest -spType $global:destSPTypeCheck -SourceOrTarget "Target"
    $WPFlabel_Status.Content = "Status: Idle"
    opacityAnimation -grid $WPFdestSiteCol -action "Close"
    opacityAnimation -grid $WPFdestListsAll -action "Open"
    if($WPFBack_Button_Two.Visibility -eq "Hidden"){
        opacityAnimation -grid $WPFBack_Button_Two -action "Open"
    } 
} else {
    $WPFlabel_Status.Content = "Status: Fill all Fields!"
    $Form.Dispatcher.Invoke("Background", [action]{})
    }
})

####Get Source Fields Button
$WPFbutton_GetSourceFields.Add_Click({
    if($WPFlistView.SelectedItem.Title){
        $WPFlabel_Status.Content = "Status: Getting Source List/Library"
        opacityAnimation -grid $WPFlistFieldsSource -action "Pre"

        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
        [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
        $sourceFieldsArray = getFields -User $WPFtextBox_User.text -Password $WPFtextBox_Pass.text -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -listView $WPFlistView_Fields_Source -spType $global:sourceSPTypeCheck -SourceOrTarget "Source"
        $WPFbutton_GetSourceFields.tag = $sourceFieldsArray
        $WPFlabel_Status.Content = "Status: Idle"
        opacityAnimation -grid $WPFsourceListsAll -action "Close"
        opacityAnimation -grid $WPFlistFieldsSource -action "Open"

        if($WPFlistFieldsDest.Visibility -eq "Visible"){
            opacityAnimation -grid $WPFlistFieldsCombine -action "Open"
        }
    } else {
        $WPFlabel_Status.Content = "Status: Choose Source List!"
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
})

####Get Dest Fields Button
$WPFbutton_GetDestFields.Add_Click({
    if($WPFlistView_Dest.SelectedItem.Title){
        $WPFlabel_Status.Content = "Status: Status: Getting All Destination Lists and Libraries"

        ####This line is for Status Update. Otherwise the Label won't Update. More Info: https://powershell.org/forums/topic/refresh-wpa-label-multiple-times-with-one-click/
        $Form.Dispatcher.Invoke("Background", [action]{})
        opacityAnimation -grid $WPFlistFieldsDest -action "Pre"

        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
        [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title
        $destFieldsArray = getFields -User $WPFtextBox_User_Dest.text -Password $WPFtextBox_Pass_Dest.text -SiteURL $WPFtextBox_URL_Dest.text -docLibName $destListTitle -listView $WPFlistView_Fields_Dest -spType $global:destSPTypeCheck -SourceOrTarget "Target"
        $WPFbutton_GetDestFields.tag = $destFieldsArray
        $WPFlabel_Status.Content = "Status: Idle"
        opacityAnimation -grid $WPFdestListsAll -action "Close"
        opacityAnimation -grid $WPFlistFieldsDest -action "Open"
        #Start-Sleep -m 400
        if($WPFlistFieldsSource.Visibility -eq "Visible"){
            opacityAnimation -grid $WPFlistFieldsCombine -action "Open"
        }
    } else {
        $WPFlabel_Status.Content = "Status: Choose Source List!"
        $Form.Dispatcher.Invoke("Background", [action]{})
    }

})

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

################################################Fields Mapping Buttons
####Map all selected Fields in Fields ListViews to Final Mapping. Also Populate Filter List View!
$WPFbutton_Map_Map.Add_Click({
    ####Map
    $i=0
    $MapListArray = @()
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
    $WPFlistView_Fields_Final.ItemsSource += $MapListArray
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

$WPFbutton_Map_Continue.Add_Click({
    opacityAnimation -grid $WPFfield_Controls -action "Close"
    opacityAnimation -grid $WPFfilterChoice -action "Open"
    opacityAnimation -grid $WPFBack_Button_Three -action "Open"
})


function listBrowser ($itemsForView){
        $browserArray = @()
        foreach ($item in $itemsForView){
            $obj = new-object psobject -Property @{
                'path' = $item["FileRef"]
                'name' = $item["FileRef"].split("/")[-1]
            }
            if($item["FSObjType"] -eq 1){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            if($item["FSObjType"] -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\file-icon-28038.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            $browserArray += $obj
            #$wpflistView_Browser.items.Add($obj)
        }
        $wpflistView_Browser.ItemsSource = $browserArray
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
    $context.ExecuteQuery()

    ####Store Current Location in Button
    $WPFradioButton_Browser.tag = $List.RootFolder.serverrelativeurl
    $wpflistView_Browser.tag = $folderPath

    $global:selectedItemsForCopy = $items

    ####Return
    return $items
}

$WPFBack_Button_Three.Add_Click({   
    if($WPFbrowserGrid.visibility -eq "Visible"){
        opacityAnimation -grid $WPFbrowserGrid -action "Close"        
    }
    opacityAnimation -grid $WPFBack_Button_Three -action "Close" 
    opacityAnimation -grid $WPFfilterChoice -action "Close"
    opacityAnimation -grid $WPFfield_Controls -action "Open"

    ####Nullify Selected Items for Copy
    $addItemsForViewArray = @()
    $WPFlistView_Items.itemsSource = @()
    $global:itemsToCopy = @()
    $WPFradioButton_Browser.isChecked = $false
})

####Browser Radio
$WPFradioButton_Browser.Add_Click({    
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
    opacityAnimation -grid $WPFbrowserGrid -action "Pre"
    $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.text -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath ""
    listBrowser -itemsForView $itemsForView
    addToGlobalCopymode -mode "Browser"

    if($WPFbutton_Copy_NoFilter.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFbutton_Copy_NoFilter -action "Close"
    }
    opacityAnimation -grid $WPFbrowserGrid -action "Open"

})

####Copy without Filter Radio
$WPFradioButton_No_Filter.Add_Click({
    addToGlobalCopymode -mode "NoFilter"
    if($WPFbrowserGrid.Visibility -eq "Visible"){
        opacityAnimation -grid $WPFbrowserGrid -action "Close"
    }
    opacityAnimation -grid $WPFbutton_Copy_NoFilter -action "Open"
})

####Get into browser
$wpflistView_Browser.add_MouseDoubleClick({
    $selectedItem = $wpflistView_Browser.SelectedItem
    [string]$sourceListTitle = $WPFlistView.SelectedItem.Title

    write-host "Selected Item Path"$selectedItem.path -ForegroundColor Green

    if($selectedItem.type -eq 1){
        $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.text -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath $selectedItem.path
        listBrowser -itemsForView $itemsForView
    }
})

$WPFbutton_BrowserUp.Add_Click({
    ####If current Folder is not Root
    if($wpflistView_Browser.tag -ne $WPFradioButton_Browser.tag){
        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
        ###Trim Current Folder with one level
        $currentLocationTrim = $wpflistView_Browser.tag
        $currentLocationTrim = $currentLocationTrim.Substring(0, $currentLocationTrim.lastIndexOf('/'))

        $itemsForView = getFolderCAML -user $WPFtextBox_User.text -Password $WPFtextBox_Pass.text -SiteURL $WPFtextBox_URL.text -docLibName $sourceListTitle -folderPath $currentLocationTrim
        listBrowser -itemsForView $itemsForView
   
    }
})

function addToGlobalCopymode($mode){
$global:copyMode = $mode
}

#function addToGlobalItemArray($array){
#    $global:itemsToCopy += $array
#}

$WPFbutton_Add_For_Copy.Add_Click({
    foreach ($item in $global:selectedItemsForCopy){
        if($wpflistView_Browser.selecteditems.Path -contains $item["FileRef"]){
            ####Add to Items for Copy
            $global:itemsToCopy += $item      
        }

        $listViewForItemsArray = @()
        foreach ($item in $global:itemsToCopy){
            $obj = new-object psobject -Property @{
                'path' = $item["FileRef"]
                'name' = $item["FileRef"].split("/")[-1]
            }
            if($item["FSObjType"] -eq 1){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            if($item["FSObjType"] -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\file-icon-28038.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            $listViewForItemsArray += $obj   
        }
        $WPFlistView_Items.itemsSource = $listViewForItemsArray
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
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\icon-folder-128.png"
                $obj | Add-Member -type NoteProperty -Name 'Type' -Value $item["FSObjType"]
            }
            if($item["FSObjType"] -eq 0){
                $obj | Add-Member -type NoteProperty -Name 'imagepath' -Value "C:\Desktop\Temp\PowerShell_FORMS\file-icon-28038.png"
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
        $destList = $destContext.Web.Lists.GetByTitle($destListTitle)
        $destContext.Load($destList)
        $destContext.Load($destList.RootFolder)

        $Context.ExecuteQuery()
        $destContext.ExecuteQuery()

        iterateEverything -whatTo $browsedItems
}

################################################Copy Lists Button

function copyFunction {
    $WPFlabel_Status.Content = "Status: Starting Copy"

    ####This line is for Status Update. Otherwise the Label won't Update. More Info: https://powershell.org/forums/topic/refresh-wpa-label-multiple-times-with-one-click/
    $Form.Dispatcher.Invoke("Background", [action]{})


    if ($global:copyMode -eq "Browser"){
        [string]$sourceListTitle = $WPFlistView.SelectedItem.Title
        [string]$destListTitle = $WPFlistView_Dest.SelectedItem.Title

        Write-Host "File Ref of Folder" $sourceListTitle -ForegroundColor Cyan
        ####Get Source List
        $web = $Context.Web
        $List = $Context.Web.Lists.GetByTitle($sourceListTitle)
        $Context.Load($Web)
        $Context.Load($List)
        $context.Load($List.RootFolder)

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
                $context.ExecuteQuery()

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

        $allFinal = $foldersToCopyFinal + $itemsToCopyFinal
        copyBrowsedItems -browsedItems $allFinal
    }

    if ($global:copyMode -eq "NoFilter"){
        copyLibrary -siteURL $WPFtextBox_URL.text -docLibName $WPFlistView.SelectedItem.Title -sourceSPType $global:sourceSPTypeCheck -destSiteURL $WPFtextBox_URL_Dest.text -destDocLibName $WPFlistView_Dest.SelectedItem.Title -destSPType $global:destSPTypeCheck
        
        ####Update Form Status
        $WPFlabel_Status.Content = "Status: Finished Copying!"
        $Form.Dispatcher.Invoke("Background", [action]{})
    }
    }

####Copy Lists Button
$WPFbutton_Copy.Add_Click({
    copyFunction
})

$WPFbutton_Copy_NoFilter.Add_Click({
    copyFunction
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
################################################
## Show the form
################################################
write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null
