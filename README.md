# PowerShell-scripts
This repository contains a collection of useful PowerShell-scripts for interacting with Microsoft Windows which we use for our software-installer (built with [Advanced Installer](https://www.advancedinstaller.com/)).

The scripts are used to improve the user-expirence of the installation/update-process, e.g. to prevent the installation from commencing before the required database-access has been confirmed, which would otherwise lead to a failure later in the installation-process if applying a software-update containing one or more database-upgrades.
