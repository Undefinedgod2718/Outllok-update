import os
import sys
import datetime
import pythoncom
import win32com.client
from mcp.server.fastmcp import FastMCP

# Initialize FastMCP server
mcp = FastMCP("Outlook")

def get_outlook_namespace():
    # Helper to get Namespace in thread-safe way
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")

@mcp.tool()
def get_contacts_count() -> str:
    """Get the count of contacts in the default folder and list backup folders."""
    pythoncom.CoInitialize()
    try:
        ns = get_outlook_namespace()
        default_folder = ns.GetDefaultFolder(10) # olFolderContacts
        count = default_folder.Items.Count
        
        backup_folders = []
        for i in range(1, default_folder.Folders.Count + 1):
            sub = default_folder.Folders.Item(i)
            backup_folders.append(f"{sub.Name} ({sub.Items.Count} items)")
            
        result = f"Default Contacts Folder: {default_folder.Name}\n"
        result += f"Total Contacts/Groups: {count}\n"
        if backup_folders:
            result += f"Subfolders/Backups:\n" + "\n".join([f"  - {bf}" for bf in backup_folders])
        else:
            result += "No backup folders found."
        return result
    except Exception as e:
        return f"Error: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

@mcp.tool()
def search_contacts(query: str) -> str:
    """Search default Outlook contacts by name or email (case-insensitive)."""
    pythoncom.CoInitialize()
    try:
        ns = get_outlook_namespace()
        default_folder = ns.GetDefaultFolder(10)
        items = default_folder.Items
        
        matches = []
        query_lower = query.lower().strip()
        
        for i in range(1, items.Count + 1):
            item = items.Item(i)
            # Check message class
            msg_class = item.MessageClass
            if msg_class == "IPM.Contact":
                name = item.FullName or ""
                email = item.Email1Address or ""
                if query_lower in name.lower() or query_lower in email.lower():
                    matches.append(f"Contact: {name} <{email}>")
            elif msg_class == "IPM.DistList":
                name = item.DLName or ""
                if query_lower in name.lower():
                    matches.append(f"Group: {name} ({item.MemberCount} members)")
                    
        if not matches:
            return f"No contacts or groups found matching query '{query}'."
        
        return f"Found {len(matches)} match(es) for '{query}':\n" + "\n".join(matches)
    except Exception as e:
        return f"Error: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

@mcp.tool()
def list_backups() -> str:
    """List all contact backup folders in Outlook."""
    pythoncom.CoInitialize()
    try:
        ns = get_outlook_namespace()
        default_folder = ns.GetDefaultFolder(10)
        
        backups = []
        for i in range(1, default_folder.Folders.Count + 1):
            sub = default_folder.Folders.Item(i)
            backups.append(f"Backup folder: {sub.Name} | Items count: {sub.Items.Count}")
            
        if not backups:
            return "No backup folders found."
        return "\n".join(backups)
    except Exception as e:
        return f"Error: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

@mcp.tool()
def import_new_pst(pst_path: str) -> str:
    """Backup current contacts, clear them, and import new ones from a PST file path.
    Also executes a duplicate check by email address.
    """
    pythoncom.CoInitialize()
    try:
        if not os.path.exists(pst_path):
            return f"Error: PST file not found at {pst_path}"
            
        ns = get_outlook_namespace()
        default_folder = ns.GetDefaultFolder(10)
        old_items = default_folder.Items
        old_items_count = old_items.Count
        
        # 1. Create Backup Folder
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"連絡人_Backup_{timestamp}"
        backup_folder = default_folder.Folders.Add(backup_name)
        
        # 2. Copy items to backup
        for i in range(1, old_items_count + 1):
            item = old_items.Item(i)
            copied = item.Copy()
            copied.Move(backup_folder)
            
        # Verify backup count
        if backup_folder.Items.Count != old_items_count:
            return f"Backup verification failed! Copied: {backup_folder.Items.Count}, Original: {old_items_count}. Operation aborted."
            
        # 3. Clear default folder
        # Must delete in reverse order
        for i in range(old_items.Count, 0, -1):
            old_items.Item(i).Delete()
            
        if default_folder.Items.Count != 0:
            return f"Failed to clear all items. Remaining: {default_folder.Items.Count}. Manual check required!"
            
        # 4. Mount PST store
        ns.AddStore(pst_path)
        
        # Find store
        pst_store = None
        for i in range(1, ns.Stores.Count + 1):
            store = ns.Stores.Item(i)
            # Compare file path
            if store.FilePath and os.path.normpath(store.FilePath) == os.path.normpath(pst_path):
                pst_store = store
                break
                
        if not pst_store:
            # Try matching by filename in case path format differs
            filename = os.path.basename(pst_path).lower()
            for i in range(1, ns.Stores.Count + 1):
                store = ns.Stores.Item(i)
                if store.FilePath and os.path.basename(store.FilePath).lower() == filename:
                    pst_store = store
                    break
                    
        if not pst_store:
            return f"Error: PST store was added but could not be resolved by path: {pst_path}"
            
        # Find Contacts folder in PST
        pst_root = pst_store.GetRootFolder()
        pst_contacts_folder = None
        for i in range(1, pst_root.Folders.Count + 1):
            f = pst_root.Folders.Item(i)
            if f.DefaultItemType == 2: # olContactItem
                pst_contacts_folder = f
                break
                
        if not pst_contacts_folder:
            # Clean up store
            ns.RemoveStore(pst_root)
            return "Error: Contacts folder not found in the PST file."
            
        # 5. Import items
        pst_items = pst_contacts_folder.Items
        pst_items_count = pst_items.Count
        for i in range(1, pst_items_count + 1):
            item = pst_items.Item(i)
            copied = item.Copy()
            copied.Move(default_folder)
            
        # 6. Check duplicates
        email_map = {}
        duplicates = []
        imported_items = default_folder.Items
        for i in range(1, imported_items.Count + 1):
            item = imported_items.Item(i)
            if item.MessageClass == "IPM.Contact":
                email = item.Email1Address
                if email:
                    email_key = email.lower().strip()
                    if email_key in email_map:
                        duplicates.append(f"Email: {email} | Name: {item.FullName} vs {email_map[email_key]}")
                    else:
                        email_map[email_key] = item.FullName
                        
        # 7. Detach PST store
        ns.RemoveStore(pst_root)
        
        summary = f"SUCCESS: Outlook contacts updated!\n"
        summary += f"- Backup folder created: {backup_folder.Name} with {old_items_count} items.\n"
        summary += f"- Imported: {pst_items_count} items.\n"
        summary += f"- Duplicates found by email: {len(duplicates)}\n"
        if duplicates:
            summary += "Duplicate details:\n" + "\n".join([f"  - {d}" for d in duplicates])
            
        return summary
    except Exception as e:
        return f"Error during import: {str(e)}"
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    mcp.run()
