# This script connects to a specified SQL Server-instance using the specified
# authentication method (Windows Authentication or SQL Server Authentication)
# to verify that the privilege requirements are met.
#
# Question: WTF(lip) is this, and how do I run it?
#
# Answer: This is a Windows PowerShell-script. To run it:
#
# - Save the contents as a `.ps1` file, e.g. `TestSqlServerConnection.ps1`, e.g. to your desktop.
# - Run `cmd`
# - Run `powershell`
# - `cd` into the folder where you saved the file
# - Run the script by referencing the file `.\TestSqlServerConnection.ps1`, and adding each script parameter as a `parameterName "value"`, e.g.:
#
# If you get an error stating that the script is not digitally signed, execute the following command in PowerShell (without the quotes):
# `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`
#
# Examples:
# "Windows Authentication"
# `.\TestSqlServerConnection.ps1 -server ".\SQLEXPRESS" -port "" -sqlAuthorization "Trusted_Connection=yes;" -database "CloudPortal"`
#
# "SQL Server Authentication"
# `.\TestSqlServerConnection.ps1 -server ".\SQLEXPRESS" -port "" -sqlAuthorization "Uid=dbcreatorUser2;Pwd=Password123;" -database "CloudPortal"`


# Block for declaring the script parameters.
Param($server, $port, $sqlAuthorization, $database)

# Get params from Advanced Installer as we're apparently not able to use
# external PowerShell-scripts with parameters. All references to "AI" cmdlets, i.e.
#
# - AI_GetMsiProperty
# - AI_SetMsiProperty
#
# ... are contained in try/catch-blocks to enable error messages during development
# in a non-Advanced Installer environment where the said cmdlets aren't available

try {
    $server = AI_GetMsiProperty SERVER_PROP;
    $port = AI_GetMsiProperty PORT_PROP;
    $sqlAuthorization = AI_GetMsiProperty SQL_AUTHORIZATION;
    $database = AI_GetMsiProperty DATABASE_PROP;

    $sqlAuthorizationStyle = AI_GetMsiProperty SQL_AUTHORIZATION_STYLE;
    $sqlUser = AI_GetMsiProperty SQL_USER;
    $installationIdentity = AI_GetMsiProperty USER_NAME;

    # Reset MSI property-values to support multiple, consecutive executions
    AI_SetMsiProperty TestSqlServerConnection $false;

    AI_SetMsiProperty SqlConnectionOK $false;
    AI_SetMsiProperty SqlServerConnection $false;
    AI_SetMsiProperty SqlUserPrivilegesOK $false;
    AI_SetMsiProperty TargetPermissions $false;

    AI_SetMsiProperty SqlErrorMesssageHeader $false;
    AI_SetMsiProperty SqlErrorMessageBody $false;
} catch {}

Write-Host "CUSTOM ACTION SCRIPT: TestSqlServerConnection ($server, $port, $sqlAuthorization, $database)";
Write-Host "";

$timeoutSeconds = 10;


###############################################################################
# Private functions
###############################################################################

function VerifySqlClientAccess($connStr) {
    Write-host "";
    Write-host "VerifySqlClientAccess($connStr)";

    try {
        $sqlConnection = New-Object System.Data.SqlClient.SqlConnection ($connStr);
        $sqlConnection.Open();
        return $true;
    } catch {
        return $false;
    } finally {
        $sqlConnection.Close();
    }
}

function DatabaseExists($connStr, $database) {
    Write-host ""
    Write-host "DatabaseExists($connStr, $database)"

    $sql = "SELECT name FROM master.sys.databases WHERE name LIKE '" + $database + "';";

    $da = new-object System.Data.SqlClient.SqlDataAdapter ($sql, $connStr);
    $da.SelectCommand.CommandTimeout = $timeoutSeconds;
    $dt = new-object System.Data.DataTable;
    $da.fill($dt) | out-null;

    $dt | format-table | out-host;

    $dv = new-object System.Data.DataView($dt);

    if ($dv.Count -eq 0) {
        return $false;
    } else {
        return $true;
    }
}

function SQLResultsExist($connStr, $sql, $permissions) {
    #Write-host "SQLResultsExist($connStr, $sql)";

    $da = new-object System.Data.SqlClient.SqlDataAdapter ($sql, $connStr);

    try {
        $da.SelectCommand.CommandTimeout = $timeoutSeconds;
        $dt = new-object System.Data.DataTable;
        $da.fill($dt) | out-null;

        # Write the result to the console as a table
        Write-host "User permissions:";
        $dt | format-table | out-host;

        # Check for missing privileges
        [System.Collections.ArrayList]$missingPermissions = @();
        $dv = new-object System.Data.DataView($dt);

        for($i=0; $i -lt $permissions.length; $i++) {
            $dv.RowFilter = "permission_name = '" + $permissions[$i] + "'";

            if ($dv.Count -gt 0) {
                Write-Host "Found: " $permissions[$i];
            } else {
                Write-Host "Missing: " $permissions[$i];

                # As a match has been found, add the item to the list of missing privileges
                $missingPermissions.Add($permissions[$i]) | Out-Null;
            }
        }

        # Write-host "missingPermissions: $missingPermissions";

        # Merely returning $matchCount leads to an array being returned, so we're being explicit instead
        if ($missingPermissions.count -gt 0) {
            # Not all required persmissions were found, so return the list of missing permissions
            #Write-host "False"
            return $missingPermissions -join ', ';
        } else {
            # All required permissions were found in the query, so return true
            #Write-host "True";
            return $true;
        }
    } catch {
        return $false;
    }
}

function RowCount($connStrWithDatabase, $sql) {
    $da = new-object system.data.sqlclient.sqldataadapter ($sql, $connStrWithDatabase)
    $da.selectcommand.commandtimeout = 10
    $dt = new-object system.data.datatable
    $da.fill($dt) | out-null

    # Write the result to the console as a table
    # Write-host "Results:"
    $dt | format-table | out-host

    $dv = new-object System.Data.DataView($dt)
    
    $rowCount = $dv.Count;

    return $rowCount;
}

function CurrUserHasPermission($connStr, $permissions, $securableClass) {
    #Write-host "CurrUserHasPermission($connStr, $permissions, $securableClass)"
    Write-host "";

    #Build the permissions list
    $permissionsList = "";

    for($i=0; $i -lt $permissions.length; $i++) {
        $permissionsList = $permissionsList + "'" + $permissions[$i] + "'";

        # When not dealing with the last element, add a trailing comma
        if ($i -ne ($permissions.length - 1)) {
            $permissionsList = $permissionsList + ",";
        }
    }

    #Write-host "permissionsList: $permissionsList";

    $sql = "SELECT permission_name
            FROM fn_my_permissions(NULL, '$securableClass')
            WHERE permission_name in ($permissionsList)";


    $hasPermission = SQLResultsExist $connStr $sql $permissions;
    Write-host "hasPermission: $hasPermission";

    if ($hasPermission -eq $true) {
        return $hasPermission;
    } else {
        #If the user doesn't have all of the necessary permissions, we'll end up here
        #with a list of permissions
        return $hasPermission;
    }

    # Try {
    #   # True if any records exist
    #   $hasPermission = SQLResultsExist $connStr $sql $permissions

    #   #Write-host "hasPermission: $hasPermission"
    #   return $hasPermission
    # } Catch {
    #   # If the user doesn't have all of the necessary permissions, we'll end up here
    #   return $hasPermission
    # }
}

function FailProcess($log, $header, $body) {
    # For some reason, this function doesn't appear to be executed when called
    Write-host "FailProcess($log, $header, $body)"

    # destroy form
        $objForm.Close() | Out-Null;

    Write-host $log

    try {
        # Set a property indicating whether the SQL connection satisfies the requirements
        AI_SetMsiProperty SqlUserPrivilegesOK "False";

        # Specify an error text
        AI_SetMsiProperty SqlErrorMesssageHeader $header;
        AI_SetMsiProperty SqlErrorMessageBody $body;

        # Mark the process as failure
        AI_SetMsiProperty TestSqlServerConnection "False";
    } catch {}

    # Inspiration: https://stackoverflow.com/a/63718769/8229998
    Exit
}


###############################################################################
# Init
###############################################################################

#------------------------------------------------------------------------------
# MAKE A FORM TO VISUALIZE THE PROCESS
#------------------------------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms;

$formWidth = 300;
$formHeight = 60;
$padding = 10;
$elementWidth = $formWidth - ($padding * 2);
$backColor = "white"; # white; "#101010"; # Confirm dialog background # "#092233"; # Installer background-color
$backColorContrast = "black"; # black; "#d9d7d4"; # light gray

# Build the form
$objForm = New-Object System.Windows.Forms.Form;
$objForm.StartPosition ="CenterScreen";
$objForm.Text = "Verifying database connection";
$objForm.Size = New-Object System.Drawing.Size($formWidth, $formHeight);
$objForm.BackColor = $backColor;
$objForm.FormBorderStyle = "None"; # Remove titlebar 

# Add a label
$objLabel = New-Object System.Windows.Forms.Label;
$objLabel.Location = New-Object System.Drawing.Size(10, 10);
$objLabel.Width = $elementWidth;
$objLabel.Height = 20;
$objLabel.BackColor = "transparent";
$objLabel.ForeColor = $backColorContrast;
$objLabel.Text = "Verifying database connection and etc.";
$objForm.Controls.Add($objLabel);

# Add a progressbar
$Progressbar = New-Object System.Windows.Forms.ProgressBar;
$Progressbar.Location = New-Object System.Drawing.Point(10, (30 + $padding));
$Progressbar.Width = $elementWidth;
$Progressbar.Height = 10;
$Progressbar.Style  = [System.Windows.Forms.ProgressBarStyle]::Continuous;
$Progressbar.BackColor = "#2d2d2d";
$Progressbar.ForeColor = "#1e8000";
$Progressbar.Maximum = 3; # initial number of steps to perform
$Progressbar.Step = 1;
$Progressbar.Value = 0;
$objForm.Controls.Add($Progressbar);

# Show the form
#$objForm.ShowDialog() | Out-Null # Shows dialog and halts script execution until the dialog is closed
$objForm.Show(); # Shows dialog and continues script execution

# Refresh the form to display the changed label. This action needs to be performed after
# all changes (and initialization) to the form when using `form.Show()` or the label
# will not display correctly
$objForm.Refresh(); 


#------------------------------------------------------------------------------
# IDENTIFY THE SQL SERVER USER
#------------------------------------------------------------------------------

if ($sqlAuthorizationStyle -eq "TrustedConnection") {
    $sqlUser = $installationIdentity;
} else {
    $sqlUser = $sqlUser;
}

# Write-Host "sqlUser: $sqlUser";


#------------------------------------------------------------------------------
# CREATE THE CONNECTION STRING
#------------------------------------------------------------------------------

$connStr;

# Generate the connection-string which differs depending on whether a `port` is specified
if ($port) {
    #Write-Host "Port is specified"
    $connStr = $server + "," + $port;
} else {
    #Write-Host "Port is unspecified"
    $connStr = $server;
}

$connStr = "Server=" + $connStr + ";";
#Write-Host "connStr (1): $connStr";

# Add the SQL authentication method
$connStr = $connStr + " " + $sqlAuthorization;
#Write-Host "connStr (2): $connStr";

# Add various required info
$connStr = $connStr + " Persist Security Info=True; MultipleActiveResultSets=True; App=EntityFramework;";
#Write-Host "connStr (3): $connStr";

# Trust the server certificate, i.e. don't break if the server certificate is invalid. This will allow connections
# despite an invalid server certificate.
$connStr = $connStr + " TrustServerCertificate=yes;";
# Write-Host "connStr (final): $connStr";


#------------------------------------------------------------------------------
# CHECK SQL CLIENT ACCESS FOR SQL SERVER-INSTANCE
#------------------------------------------------------------------------------

$objLabel.Text = "Connecting to database";
$objForm.Refresh(); # Refresh the form to display the changed label
$Progressbar.PerformStep();

$sqlConnectionOk = VerifySqlClientAccess $connStr;
Write-Host "sqlConnectionOk: $sqlConnectionOk";

if ($sqlConnectionOk -eq $true) {
    Write-host "The software-installer can access the SQL Server";

    try {
        # Set a property indicating whether the SQL connection satisfies the requirements
        AI_SetMsiProperty SqlConnectionOK "True";
    } catch {}
} else {
    # destroy form
    $objForm.Close() | Out-Null;

    AI_SetMsiProperty SqlConnectionOK "False";

    $log = "The software-installer cannot access the SQL Server"
    $messageHeader = "Unable to access SQL Server-instance"
    $messageBody = "The software-installer is unable to access the SQL Server-instance using the specified values"

    FailProcess $log $messageHeader $messageBody
}


#------------------------------------------------------------------------------
# CHECK USER PERMISSIONS FOR SQL SERVER-INSTANCE
#------------------------------------------------------------------------------

$objLabel.Text = "Verifying user access privileges for 'msdb' database";
$objForm.Refresh(); # Refresh the form to display the changed label
$Progressbar.PerformStep();

if ($sqlConnectionOk -eq $true) {
    $hasPermissions;

    # CHECK MSDB PERMISSIONS
    #$connStr = "data source=" + $connStr + ";initial catalog=msdb; Persist Security Info=True;" + $sqlAuthorization + "; MultipleActiveResultSets=True; App=EntityFramework";
    $connStrWithDatabase = $connStr + " Database=msdb;";
    Write-Host "connStrWithDatabase: $connStrWithDatabase";
    
    $permissions = @("SELECT","INSERT","UPDATE");
    $hasPermissions = CurrUserHasPermission $connStrWithDatabase $permissions 'DATABASE';
    Write-Host "hasPermissions: $hasPermissions";

    Write-Host "";

    if ($hasPermissions -eq $true) {
        Write-host "SQL Server user permissions are OK";

        try {
            # Set a property indicating whether the SQL connection satisfies the requirements
            AI_SetMsiProperty SqlUserPrivilegesOK "True";
        } catch {}

        # CHECK FOR INVALID 'msdb' ENTRY
        $sql = "SELECT * FROM sysdac_instances WHERE database_name = '$database';"
        Write-host "Prope 'msdb' for '$database': $SQL"

        $result = RowCount $connStrWithDatabase $sql
        Write-host "Database-instances named '$database': $result"
        Write-host ""

        if ($result -gt 0) {
            # CHECK FOR A VALID `instance_id`
            Write-host "msdb/sysdac_instances contains one or more entries for '$database'."
            Write-host ""
            
            $sql = "SELECT * FROM sysdac_instances WHERE database_name = '$database' AND instance_id IS NOT NULL;"
            Write-host "Check for valid 'instance_id': $SQL"

            $result = RowCount $connStrWithDatabase $sql
            Write-host "Database-instances with valid 'instance_id': $result"

            if ($result -gt 0) {
                # The database instance(s) presumably has a valid 'instance_id'. All is well 
            } else {
                # destroy form
                $objForm.Close() | Out-Null;

                $log = "msdb/sysdac_instances appears to be corrupt as the entry for '$database' appears to have an invalid 'instance_id'. Please ensure that the installation identity (Windows Authentication) or username (SQL Server Authentication) is either a member of the 'sysadmin' fixed server role or is the database owner (DBO) for the '$database' database and has the necessary privileges for the 'msdb' database (refer to the administrator's manual)."
                $messageHeader = "Invalid entry in 'msdb' database"
                # $messageBody = "The entry for '$database' in: 'SQL Server > Databases > System Databases > msdb > sysdac_instances' appears to have an invalid value for 'instance_id'. Please ensure that '$sqlUser' has the necessary privileges for both 'msdb' and '$database' (refer to the administrator's manual)."
                $messageBody = "Invalid 'instance_id' for '$database' in: 'SQL Server > Databases > System Databases > msdb > sysdac_instances'. Ensure that '$sqlUser' has the necessary privileges for both 'msdb' and '$database'."

                FailProcess $log $messageHeader $messageBody            
            }
        } else {
            Write-host "msdb/sysdac_instances doesn't contain an entry for '$database'. A new entry will be created"
        }
    } else {
        # destroy form
        $objForm.Close() | Out-Null;

        $log = "SQL Server user permissions are insufficient"
        $messageHeader = "Insufficient SQL user privileges for database: msdb"
        $messageBody = "The SQL Server user '$sqlUser' does not have the necessary privileges for the database: 'SQL Server > Databases > System Databases > msdb'. Missing privileges: $hasPermissions"

        FailProcess $log $messageHeader $messageBody
    }
}


#------------------------------------------------------------------------------
# CHECK FOR PRESENCE OF TARGET DATABASE
#------------------------------------------------------------------------------

$objLabel.Text = "Checking for presence of target database";
$objForm.Refresh(); # Refresh the form to display the changed label
$Progressbar.PerformStep();

if ($hasPermissions -eq $true) {
    $hasTargetDatabase = DatabaseExists $connStr $database;
    Write-Host "hasTargetDatabase: $hasTargetDatabase";
}


#------------------------------------------------------------------------------
# CHECK USER PERMISSIONS FOR TARGET DATABASE
#------------------------------------------------------------------------------

if ($hasTargetDatabase -eq $true) {
    # Increase the number of steps as we now have an additional step to perform
    $progressbar.Maximum = $progressbar.Maximum + 1;
    # Write-Host "progressbar: $progressbar";

    $objLabel.Text = "Checking user permissions for target database";
    $objForm.Refresh(); # Refresh the form to display the changed label
    $Progressbar.PerformStep();

    $connStrWithDatabase = $connStr + " Database=$database;"
    Write-Host "connStrWithDatabase: $connStrWithDatabase";
    
    $permissions = @("SELECT","INSERT","UPDATE");
    $targetPermissions = CurrUserHasPermission $connStrWithDatabase $permissions 'DATABASE';

    Write-Host "targetPermissions: $targetPermissions";

    if ($targetPermissions -eq $true) {
        Write-host "User permissions are OK for target database";

        try {
            # Set a property indicating whether the SQL connection satisfies the requirements
            AI_SetMsiProperty TargetPermissions "True";
        } catch {}
    } else {
        # destroy form
        $objForm.Close() | Out-Null;

        $log = "User permissions are insufficient for target database"

        # Specify an error text
        if ($targetPermissions -eq $false) {
            $messageHeader = "Unable to access database: $database"
            $messageBody = "The SQL Server user '$sqlUser' does not have the necessary privileges for the database: 'SQL Server > Databases > $database' or database is unavailable. 'Database owner' ('DBO') or (temporary) 'sysadmin' server role required"
        } else {
            $messageHeader = "Insufficient SQL user privileges for database: $database";
            $messageBody = "The SQL Server user '$sqlUser' does not have the necessary privileges for the database: 'SQL Server > Databases > $database'. Missing privileges: $targetPermissions";
        }

        FailProcess $log $messageHeader $messageBody        
    }   
}


#------------------------------------------------------------------------------
# CLEAN UP
#------------------------------------------------------------------------------

# destroy form
$objForm.Close() | Out-Null;

try {
    # If running under "Advanced Installer"
    # Mark the process as success
    AI_SetMsiProperty TestSqlServerConnection "True";
} catch {}

return $true;

# DEBUG: Invalidate the installer to prevent it from continuing during
#AI_SetMsiProperty SqlConnectionOK = "False";
