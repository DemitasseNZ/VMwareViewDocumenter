#If this is a View hosts then gather data. If there is data & Word is installed generate document
$DidWork = $False
Add-PSSnapin VMware.View.Broker -ErrorAction SilentlyContinue
if (Get-PSSnapin -Name VMware.View.Broker -ErrorAction SilentlyContinue) {
    Write-Host "Gathering Data"
    get-connectionbroker | export-csv Connectionbroker.csv
    get-ViewVC | export-csv ViewVC.csv
    get-pool | export-csv Pool.csv
    get-poolentitlement | export-csv PoolEntitlement.csv
    Get-GlobalSetting | export-csv GlobalSetting.csv
    Get-License | export-csv License.csv
    Get-ComposerDomain | export-csv ComposerDomain.csv
    $DidWork = $True
}
# Check Documents exist
$FoundAll = $True
If (!(Test-Path .\Connectionbroker.csv)) {$FoundAll = $False}
If (!(Test-Path .\ViewVC.csv)) {$FoundAll = $False}
If (!(Test-Path .\ViewVC.csv)) {$FoundAll = $False}
If (!(Test-Path .\GlobalSetting.csv)) {$FoundAll = $False}
If (!(Test-Path .\License.csv)) {$FoundAll = $False}
If (!(Test-Path .\ComposerDomain.csv)) {$FoundAll = $False}
If (!(Test-Path .\PoolEntitlement.csv)) {$FoundAll = $False}
If (!($FoundALL)) {
	Write-Host "Did not find View data files in current folder"
	Write-Host "Run this script on the connection server and copy the csv files here."
	Read-Host "Then rerun this script on a computer with Word installed"
} Else {
    #Found Data files
    #Check Word is installed.
    $FoundWord = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where {($_.Publisher -eq "Microsoft Corporation" -and ($_.DisplayName -Like "*Office*" -or $_.DisplayName -Like "*Word*"))}
    If ($FoundWord) {
    	Write-Host "Word appears to be installed and we have input data, generating documentation"
        #Read Documents in
        $ConnectionBroker = import-csv .\Connectionbroker.csv
        $ViewVC = import-csv .\ViewVC.csv
        $Pool = import-csv .\Pool.csv
        $PoolEntitlement = import-csv .\PoolEntitlement.csv
        $GlobalSetting = import-csv .\GlobalSetting.csv
        $License = import-csv .\License.csv
        $ComposerDomain = import-csv .\ComposerDomain.csv
        $GlobalSetting = import-csv .\GlobalSetting.csv

        #Create new Word Document
        $word = New-Object -ComObject word.application
        $word.visible = $true
        $doc = $word.documents.add()
        $selection = $word.selection
        
        #Add Connection Global Information"
        $Selection.typeText("Global Information")
        $selection.Style = "Heading 3"
        $Selection.TypeParagraph()
        foreach ($aGlobalSetting in $GlobalSetting) { 
        	$selection.Font.Size=12
        	$paragraph = $doc.Content.Paragraphs.Add()
        	$range = $paragraph.Range
        	$rows = 12; $columns = 2
        	$table = $doc.Tables.add($range,$rows,$columns)
        	$table.cell(1,1).range.text = "Setting"
        	$table.cell(1,2).range.text = "Value"
        	$table.cell(2,1).range.text = "Session Timeout"
        	$table.cell(2,2).range.text = ($aGlobalSetting.SessionTimeout / 3600)
        	$table.cell(3,1).range.text = "User Secure Tunnel"
        	$table.cell(3,2).range.text = $aGlobalSetting.UseSslClient
        	$table.cell(4,1).range.text = "Pre-Logon Message"
            If ($aGlobalSetting.DisplayPreLogin -eq $True) {
        	   $table.cell(4,2).range.text = $aGlobalSetting.PreLoginMessage
            } Else {
                $table.cell(4,2).range.text = "Disabled"
            }
        	$table.cell(5,1).range.text = "Forced Logoff Warning Time"
        	$table.cell(6,1).range.text = "Forced Logoff Message"
            If ($aGlobalSetting.DisplayLogoffWarning -eq $True) {
        	   $table.cell(5,2).range.text = $aGlobalSetting.ForcedLogoffAfter + " Minutes"
               $table.cell(6,2).range.text = $aGlobalSetting.ForcedLogoffMessage
            } Else {
               $table.cell(5,2).range.text = "Disabled"
               $table.cell(6,2).range.text = "Disabled"
            }
        	
        	$table.cell(7,1).range.text = "License Expires"
        	$table.cell(7,2).range.text = $License.'Expiry Date'
            
        	$table.cell(8,1).range.text = "Event Database Server Type"
        	$table.cell(8,2).range.text = " "
        	$table.cell(9,1).range.text = "Event Database Server"
        	$table.cell(9,2).range.text = " "
        	$table.cell(10,1).range.text = "Event Database Server Port"
        	$table.cell(10,2).range.text = " "
        	$table.cell(11,1).range.text = "Event Database User"
        	$table.cell(11,2).range.text = " "
        	$table.cell(12,1).range.text = "Event Database Table Prefix"
        	$table.cell(12,2).range.text = " "

            $table.UpdateAutoFormat()
        	$Table.Style = "Medium List 1 - Accent 1"
        	$a = $Selection.EndKey(6) 
        	$Selection.TypeParagraph() 
        }
        
        #Add Connection Server information
        $Selection.typeText("View Connection Servers")
        $selection.Style = "Heading 3"
        $Selection.TypeParagraph()
        foreach ($aConnectionBroker in $ConnectionBroker) { 
        	$selection.Font.Size=12
        	$paragraph = $doc.Content.Paragraphs.Add()
        	$range = $paragraph.Range
        	$rows = 6; $columns = 2
        	$table = $doc.Tables.add($range,$rows,$columns)
        	$table.cell(1,1).range.text = "Setting"
        	$table.cell(1,2).range.text = "Value"
        	$table.cell(2,1).range.text = "Computer Name"
        	$table.cell(2,2).range.text = $aConnectionBroker.broker_id
        	$table.cell(3,1).range.text = "Role"
        	$table.cell(3,2).range.text = $aConnectionBroker.type
        	$table.cell(4,1).range.text = "Tags"
        	$table.cell(4,2).range.text = $aConnectionBroker.tags
        	$table.cell(5,1).range.text = "External URL"
        	$table.cell(5,2).range.text = $aConnectionBroker.externalURL
        	$table.cell(6,1).range.text = "External PCoIP URL"
        	$table.cell(6,2).range.text = $aConnectionBroker.externalPCoIPURL
        	
        	$table.UpdateAutoFormat()
        	$Table.Style = "Medium List 1 - Accent 1"
        	$a = $Selection.EndKey(6) 
        	$Selection.TypeParagraph() 
        }
        #Add vCentre Server information
        $Selection.typeText("vCentre Servers")
        $selection.Style = "Heading 3"
        $Selection.TypeParagraph()
        foreach ($aViewVC in $ViewVC) { 
        	$selection.Font.Size=12
        	$paragraph = $doc.Content.Paragraphs.Add()
        	$range = $paragraph.Range
        	$rows = 5; $columns = 2
        	$table = $doc.Tables.add($range,$rows,$columns)
        	$table.cell(1,1).range.text = "Setting"
        	$table.cell(1,2).range.text = "Value"
        	$table.cell(2,1).range.text = "Computer Name"
        	$table.cell(2,2).range.text = $aViewVC.serverName
        	$table.cell(3,1).range.text = "Service User Name"
        	$table.cell(3,2).range.text = $aViewVC.username
        	$table.cell(4,1).range.text = "View composer URL"
        	$table.cell(4,2).range.text = $aViewVC.composerUrl
        	$table.cell(5,1).range.text = "View Composer User Name"
        	$table.cell(5,2).range.text = $aViewVC.composerUsername
        	
        	$table.UpdateAutoFormat()
        	$Table.Style = "Medium List 1 - Accent 1"
        	$a = $Selection.EndKey(6) 
        	$Selection.TypeParagraph() 
        }
        #Add Pool information
        $Selection.typeText("View Pools")
        $selection.Style = "Heading 3"
        $Selection.TypeParagraph()
        foreach ($aPool in $Pool) { 
        	$selection.Font.Size=12
        	$paragraph = $doc.Content.Paragraphs.Add()
        	$range = $paragraph.Range
        	$rows = 13; $columns = 2
        	$table = $doc.Tables.add($range,$rows,$columns)
        	$table.cell(1,1).range.text = "Setting"
        	$table.cell(1,2).range.text = "Value"
        	$table.cell(2,1).range.text = "Pool Name"
        	$table.cell(2,2).range.text = $aPool.DisplayName
        	$table.cell(3,1).range.text = "Pool Identifier"
        	$table.cell(3,2).range.text = $aPool.Pool_ID
        	$table.cell(4,1).range.text = "Pool Description"
        	$table.cell(4,2).range.text = $aPool.Description
        	$table.cell(5,1).range.text = "Delivery Model"
        	$table.cell(5,2).range.text = $aPool.deliveryModel
        	$table.cell(6,1).range.text = "Desktop Source"
        	$table.cell(6,2).range.text = $aPool.desktopSource
        	$table.cell(7,1).range.text = "Persistence"
        	$table.cell(7,2).range.text = $aPool.persistence
        	$table.cell(8,1).range.text = "Pool Type"
        	$table.cell(8,2).range.text = $aPool.poolType
        	$table.cell(9,1).range.text = "Default Protocol"
        	$table.cell(9,2).range.text = $aPool.Protocol
        	$table.cell(10,1).range.text = "User Protocol Override"
        	$table.cell(10,2).range.text = $aPool.allowProtocolOverride
        	$table.cell(11,1).range.text = "Allow User to Reset"
        	$table.cell(11,2).range.text = $aPool.userResetAllowed
        	$table.cell(12,1).range.text = "Inventory Folder"
        	$table.cell(12,2).range.text = $aPool.folderId
        	$table.cell(13,1).range.text = "Entitled Users"
        	$EntitledUsers =""
        	foreach ($aPoolEntitlement in $PoolEntitlement) {
        		If ($aPoolEntitlement.pool_id -eq $aPool.Pool_ID) {
        			If ($EntitledUsers -ne "") { $EntitledUsers += ", "}
        			$EntitledUsers += $aPoolEntitlement.displayName
        			
        		}
        	}
        	$table.cell(13,2).range.text = $EntitledUsers
        	$table.UpdateAutoFormat()
        	$Table.Style = "Medium List 1 - Accent 1"
        	$a = $Selection.EndKey(6) 
        	$Selection.TypeParagraph() 
        }
        $DidWork = $True
    }
}
If (!($DidWork)) {
	Write-Host "Did not find View, Word or Office installed on this computer"
	Write-Host "First run this script on the connection server to generate csv files."
    Write-Host "Then run this script in a folder with the gathered csv files on a computer with Word installed"
    Write-Host "If Word is installed on a connection server then both actions will be taken"
}