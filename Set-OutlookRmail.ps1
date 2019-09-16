<#
.SYNOPSIS
    This script configures Rmail folders and rules in Outlook desktop clients.
.DESCRIPTION
    The script will add Receipts and Contracts folders.
    It will then add rules to route messages:
        From: receipts@r1.rpost.net => Receipts folder
        From: contracts@r1.rpost.net => Contracts folder
    It checks for the existence for these folders and rules before attempting to create them.
.NOTES
  Version:        0.1.1
  Author:         Tim Jenks <tjenks@rpost.com>
  Creation Date:  09/12/2019
  Purpose/Change: Initial script development
#>

[CmdletBinding()]
param (
    [String] $ReceiptFolderName = "Receipts",
    [String] $ReceiptSender = "receipts@r1.rpost.net",
    [String] $ReceiptRule = "Receipts",
    [String] $ContractsFolderName = "Contracts",
    [String] $ContractsSender = "contracts@r1.rpost.net",
    [String] $ContractsRule = "Contracts"
)

begin {

}

process {
    # Verify Outlook is installed and set up
    function CheckOutlook {
        if (!(Test-Path $env:localappdata\Microsoft\Outlook\*.*st)) {
            Write-Host "outlook has not been started or configured on client yet";
            exit
        }
    }

    # Garbage collection for COM objects
    function ReleaseRef {
        param($ref)
        ([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # Add folder to Outlook under Inbox
    function AddOutlookFolder {
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

        if (!$exists) {
            try {
                $root.Folders.Add($name) | Out-Null
            }
            catch {
                return Write-Host $_.Exception.Message`n
            }

            Write-Host "Folder created: $name"
        }
        else {
            Write-Host "Folder exists: $name"
        }
    }

    function AddOutlookFolderRule {
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

        foreach($r in $rules) {
            # Write-Host($r | Format-Table | Out-String)
            if($r.Name -eq $RuleName){
                return Write-Host "Rule already exists: $RuleName"
            }
        }

        try {
            $rule = $rules.Create($RuleName,$olRuleType::OlRuleReceive)
            $FromCondition = $rule.Conditions.From
            $FromCondition.Enabled = $true
            $FromCondition.Recipients.Add($FromEmail) | Out-Null
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
        }

        catch {
            return Write-Host $_.Exception.Message`n
        }

        Write-Host "Rule added: $RuleName"
    }

    function Set-OutlookRmail {
        param (
            [String] $ReceiptFolderName,
            [String] $ReceiptSender,
            [String] $ReceiptRule,
            [String] $ContractsFolderName,
            [String] $ContractsSender,
            [String] $ContractsRule
        )
    
        AddOutlookFolder $ReceiptFolderName
        AddOutlookFolder $ContractsFolderName
        AddOutlookFolderRule $ReceiptRule $ReceiptSender $ReceiptFolderName
        AddOutlookFolderRule $ContractsRule $ContractsSender $ContractsFolderName
    }

    Set-OutlookRmail $ReceiptFolderName $ReceiptSender $ReceiptRule $ContractsFolderName $ContractsSender $ContractsRule

}

end {

}









