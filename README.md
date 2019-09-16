# Set-OutlookRmail

## Summary

Powershell script to configure Outlook desktop clients with Rmail folders and rules

## Description

The script will add Receipts and Contracts folders.
It will then add rules to route messages:

- From: receipts@r1.rpost.net => Receipts folder
- From: contracts@r1.rpost.net => Contracts folder

It checks for the existence for these folders and rules before attempting to create them.

## Parameters

- `$ReceiptFolderName`
    + Type: String
    + Description: Name of the folder for Rmail Receipts
    + Default: "Receipts"
- `$ReceiptSender`
    + Type: String
    + Description: Sender email address for Rmail Receipts
    + Default: "receipts@r1.rpost.net"
- `$ReceiptRule`
    + Type: String
    + Description: Name receipt folder routing rule
    + Default: "Receipts"
- `$ContractFolderName`
    + Type: String
    + Description: Name of the folder for Rmail contracts
    + Default: "Contracts"
- `$ContractSender`
    + Type: String
    + Description: Sender email address for Rmail contracts
    + Default: "contracts@r1.rpost.net"
- `$ContractRule`
    + Type: String
    + Description: Name contract folder routing rule
    + Default: "Contracts"

## Example Usage

    Set-OutlookRmail 

    Set-OutlookRmail -ReceiptRule RmailReceipts -ContractRule RmailContracts