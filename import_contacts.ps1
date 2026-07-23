$ErrorActionPreference = "Stop"

try {
    # 1. Initialize Outlook COM
    Write-Output "[1/8] Initializing Outlook COM Object..."
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $defaultContactsFolder = $namespace.GetDefaultFolder(10) # olFolderContacts
    
    # 2. Get current items
    $oldItems = $defaultContactsFolder.Items
    $oldItemsCount = $oldItems.Count
    Write-Output "Current contacts and groups count: $oldItemsCount"
    
    # 3. Create backup folder
    $backupFolderName = "連絡人_Backup_20260702"
    $backupFolder = $null
    try {
        $backupFolder = $defaultContactsFolder.Folders.Item($backupFolderName)
        Write-Output "Existing backup folder found. Reusing it."
    } catch {
        Write-Output "[2/8] Creating backup folder: $backupFolderName"
        $backupFolder = $defaultContactsFolder.Folders.Add($backupFolderName)
    }
    
    # 4. Perform Backup
    Write-Output "[3/8] Copying all existing items to backup folder..."
    $backupCount = 0
    # Copy items one by one
    for ($i = 1; $i -le $oldItemsCount; $i++) {
        $item = $oldItems.Item($i)
        $copiedItem = $item.Copy()
        $copiedItem.Move($backupFolder) | Out-Null
        $backupCount++
        if ($backupCount % 100 -eq 0) {
            Write-Output "  Backed up $backupCount / $oldItemsCount items..."
        }
    }
    Write-Output "Backup completed. Items in backup folder: $($backupFolder.Items.Count) (Expected: $oldItemsCount)"
    
    # Double check backup count
    if ($backupFolder.Items.Count -ne $oldItemsCount) {
        throw "Backup count verification failed! Backup count: $($backupFolder.Items.Count), Original count: $oldItemsCount. Aborting operation."
    }
    
    # 5. Clear Old Contacts (Delete original items)
    Write-Output "[4/8] Clearing old contacts and groups from default folder..."
    # We must delete items in reverse order because the collection indices change upon deletion
    $deletedCount = 0
    for ($i = $oldItems.Count; $i -ge 1; $i--) {
        $item = $oldItems.Item($i)
        $item.Delete()
        $deletedCount++
        if ($deletedCount % 100 -eq 0) {
            Write-Output "  Deleted $deletedCount items..."
        }
    }
    Write-Output "Cleared all items. Current items count in default folder: $($defaultContactsFolder.Items.Count)"
    
    if ($defaultContactsFolder.Items.Count -ne 0) {
        throw "Failed to clear all default contacts! Remaining: $($defaultContactsFolder.Items.Count)"
    }
    
    # 6. Extract and Load the New PST Store
    $pstPath = "C:\Users\richard_zhang\.gemini\antigravity\scratch\260702address.pst"
    Write-Output "[5/8] Adding PST store: $pstPath"
    $namespace.AddStore($pstPath)
    
    # Find the added PST Store
    $pstStore = $null
    foreach ($store in $namespace.Stores) {
        if ($store.FilePath -eq $pstPath) {
            $pstStore = $store
            break
        }
    }
    
    if ($pstStore -eq $null) {
        throw "Could not load the PST file: $pstPath"
    }
    
    # Find the Contacts folder in the PST
    $pstRoot = $pstStore.GetRootFolder()
    $pstContactsFolder = $null
    foreach ($folder in $pstRoot.Folders) {
        if ($folder.DefaultItemType -eq 2) { # olContactItem
            $pstContactsFolder = $folder
            break
        }
    }
    
    if ($pstContactsFolder -eq $null) {
        throw "Contacts folder not found in the PST file!"
    }
    
    $newItems = $pstContactsFolder.Items
    $newItemsCount = $newItems.Count
    Write-Output "Found $newItemsCount items to import from PST."
    
    # 7. Import Items
    Write-Output "[6/8] Importing new contacts and groups..."
    $importedCount = 0
    for ($i = 1; $i -le $newItemsCount; $i++) {
        $item = $newItems.Item($i)
        $copiedItem = $item.Copy()
        $copiedItem.Move($defaultContactsFolder) | Out-Null
        $importedCount++
        if ($importedCount % 50 -eq 0) {
            Write-Output "  Imported $importedCount / $newItemsCount items..."
        }
    }
    
    # Verify import count
    Write-Output "Import completed. Default folder items count: $($defaultContactsFolder.Items.Count) (Expected: $newItemsCount)"
    
    # 8. Duplicate Check
    Write-Output "[7/8] Checking for duplicate contacts by email..."
    $emailMap = @{}
    $duplicateCount = 0
    $duplicatesList = @()
    $importedItems = $defaultContactsFolder.Items
    for ($i = 1; $i -le $importedItems.Count; $i++) {
        $item = $importedItems.Item($i)
        if ($item.MessageClass -eq "IPM.Contact") {
            $email = $item.Email1Address
            if ($email) {
                $emailKey = $email.ToLower().Trim()
                if ($emailMap.ContainsKey($emailKey)) {
                    $duplicateCount++
                    $duplicatesList += [PSCustomObject]@{
                        Name = $item.FullName
                        Email = $email
                        ExistingName = $emailMap[$emailKey]
                    }
                } else {
                    $emailMap[$emailKey] = $item.FullName
                }
            }
        }
    }
    Write-Output "Duplicate check finished. Found $duplicateCount duplicate(s) by email."
    if ($duplicateCount -gt 0) {
        Write-Output "Duplicate contact details:"
        foreach ($dup in $duplicatesList) {
            Write-Output "  - Email: $($dup.Email) | Name: $($dup.Name) vs Existing Name: $($dup.ExistingName)"
        }
    }
    
    # Remove store
    Write-Output "[8/8] Detaching PST store..."
    $namespace.RemoveStore($pstRoot)
    
    Write-Output "SUCCESS: Outlook contacts have been successfully updated!"
    Write-Output "Summary:"
    Write-Output "- Backup folder: $($backupFolder.FolderPath) with $oldItemsCount items."
    Write-Output "- New contacts imported: $importedCount items."
    Write-Output "- Duplicate contacts found: $duplicateCount items."

} catch {
    Write-Error $_
    Write-Output "ERROR occurred during import process. Backup remains intact in the Outlook folder '連絡人_Backup_20260702'."
}
