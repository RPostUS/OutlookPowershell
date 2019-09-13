# Set-OutlookRmail

## Summary

Powershell script to configure Outlook desktop clients with Rmail folders and rules

## Description

The script will add Receipts and Contracts folders.
It will then add rules to route messages:

- From: receipts@r1.rpost.net => Receipts folder
- From: contracts@r1.rpost.net => Contracts folder

It checks for the existence for these folders and rules before attempting to create them.

