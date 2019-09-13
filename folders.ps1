#Check if user uses Outlook. 
if (!(Test-Path $env:localappdata\Microsoft\Outlook\*.*st))
{
    Write-Host "outlook has not been started or configured on client yet";
    exit
}

# Garbage collection for COM objects
function ReleaseRef  
{
    param($ref)
    ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

# Add folder to Outlook under Inbox
function AddFolder
{
    param($name)

    Add-Type -AssemblyName microsoft.office.interop.outlook
    # $olFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
    $olFolderInbox = 6
    $outlook = new-object -comobject outlook.application
    $namespace = $outlook.GetNamespace("MAPI")

    # root: Parent of Inbox
    $root = $namespace.GetDefaultFolder($olFolderInbox).Parent
    # $root = $namespace.GetDefaultFolder($olFolders::olFolderInbox).Parent

    $exists = $root.Folders | where-object { $_.name -eq $name }

    if (!$exists)
    {
        $nf = $root.Folders.Add($name)

        Write-Host "Folder created: $name"

        ReleaseRef($nf)
    }
    else
    {
        Write-Host "Folder exists: $name"
    }

    ReleaseRef($root)
    ReleaseRef($namespace)
    ReleaseRef($outlook)
}

function AddOutlookFolderRule
{
    param([string]$RuleName, [string]$FromEmail, [string]$FolderName)

    Add-Type -AssemblyName microsoft.office.interop.outlook
    $olFolders = “Microsoft.Office.Interop.Outlook.OlDefaultFolders” -as [type]
    $olRuleType = “Microsoft.Office.Interop.Outlook.OlRuleType” -as [type]
    $outlook = New-Object -ComObject outlook.application
    $namespace  = $outlook.GetNameSpace(“mapi”)
    $root = $namespace.GetDefaultFolder($olFolders::olFolderInbox).Parent
    # $MoveTarget = $root.Folders.item($FolderName)
    $id = [System.__ComObject].InvokeMember(
        "EntryID",
        [System.Reflection.BindingFlags]::GetProperty,
        $null,
        $root.Folders.Item($FolderName),
        $null)
    $MoveTarget = $namespace.getFolderFromID($id)
    $rules = $outlook.session.DefaultStore.GetRules()

    foreach($r in $rules)
    {
        # Write-Host($r | Format-Table | Out-String)
        if($r.Name -eq $RuleName){
            Write-Host "Rule already exists: $RuleName"
            return
        }
    }

    $rule = $rules.Create($RuleName,$olRuleType::OlRuleReceive)
    $FromCondition = $rule.Conditions.From
    $FromCondition.Enabled = $true
    $FromCondition.Recipients.Add($FromEmail)
    $FromCondition.Recipients.ResolveAll()
    $MoveRuleAction = $rule.Actions.MoveToFolder
    # $MoveRuleAction.Folder = $MoveTarget
    [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
        "Folder",
        [System.Reflection.BindingFlags]::SetProperty,
        $null,
        $MoveRuleAction,
        $MoveTarget)
    $MoveRuleAction.Enabled = $true
    $rules.Save()

    Write-Host "Rule added: $RuleName"
}

AddFolder "Receipts"
AddFolder "Contracts"
AddOutlookFolderRule "Receipts" "receipts@r1.rpost.net" "Receipts"
AddOutlookFolderRule "Contracts" "contracts@r1.rpost.net" "Contracts"
