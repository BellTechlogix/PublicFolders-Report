<#
	PublicFolders-Report.ps1
	Created By - Steve Halligan
	Original Source - https://gallery.technet.microsoft.com/office/Snapshot-report-of-Public-21235573
	Created On - 22 Feb 2013
	Modified By - Kristopher Roy
	Modified On - 14 Feb 20202
#>
<#
	Note from Source:
	This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.   
	THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,  
	INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.   
	We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object  
	code form of the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market  
	Your software product in which the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product  
	in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against  
	any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.

#>

Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn;

 
Write-Progress -Activity "Finding Public Folders" -Status "running get-publicfolders -recurse" 
$folders = get-publicfolder -recurse -resultsize unlimited 
$arFolderData = @() 
$totalfolders = $folders.count 
$i = 1 
foreach ($folder in $folders)  
{ 
    $statusstring = "$i of $totalfolders" 
    write-Progress -Activity "Gathering Public Folder Information" -Status $statusstring -PercentComplete ($i/$totalfolders*100) 
    $folderstats = get-publicfolderstatistics $folder.identity 
    $folderdata = new-object Object 
    $folderdata | add-member -type NoteProperty -name FolderName $folder.name 
    $folderdata | add-member -type NoteProperty -name FolderPath $folder.identity 
    $folderdata | add-member -type NoteProperty -name LastAccessed $folderstats.LastAccessTime 
    $folderdata | add-member -type NoteProperty -name LastModified $folderstats.LastModificationTime 
    $folderdata | add-member -type NoteProperty -name Created $folderstats.CreationTime 
    $folderdata | add-member -type NoteProperty -name ItemCount $folderstats.ItemCount 
    $folderdata | add-member -type NoteProperty -name Size $folderstats.TotalItemSize 
    $folderdata | Add-Member -type NoteProperty -Name Mailenabled $folder.mailenabled 
 
    if ($folder.mailenabled) 
    { 
        #since there is no guarentee that a public folder has a unique name we need to compare the PF's entry ID to the recipient objects external email address 
        $entryid = $folder.entryid.tostring().substring(76,12) 
        $primaryemail = (get-recipient -filter "recipienttype -eq 'PublicFolder'" -resultsize unlimited | where {$_.externalemailaddress -like "*$entryid"}).primarysmtpaddress 
        $folderdata | add-member -type NoteProperty -name PrimaryEmailAddress $primaryemail 
    } else  
    { 
        $folderdata | add-member -type NoteProperty -name PrimaryEmailAddress "Not Mail Enabled" 
    } 
 
    if ($folderstats.ownercount -gt 0) 
    { 
        $owners = get-publicfolderclientpermission $folder.identity | where {$_.accessrights -like "*owner*"} 
        $ownerstr = "" 
        foreach ($owner in $owners)  
        { 
            $ownerstr += $owner.user.exchangeaddressbookdisplayname + "," 
        } 
     } else { 
        $ownerstr = "" 
     } 
     $folderdata | add-member -type NoteProperty -name Owners $ownerstr 
     $arFolderData += $folderdata 
     $i++ 
 } 
 $arFolderData | export-csv -path c:\Reports\PublicFolderData.csv -notypeinformation

 $recipients = @("Kristopher <kroy@belltechlogix.com>","Tim <twheeler@belltechlogix.com>","Jared <jglenn@belltechlogix.com>")

 Send-MailMessage -from ExchangeReports@crowley.com -to $recipients -subject "Crowley Public Folder Report" -smtpserver mail.crowley.com -Attachments "C:\Reports\PublicFolderData.csv"
