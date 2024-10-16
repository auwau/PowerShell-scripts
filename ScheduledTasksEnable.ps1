# This script enables any scheduled task which name matches the
# `scheduledTaskKeyword` parameter value which may currently be disabled.
#
# Question: WTF(lip) is this, and how do I run it?
#
# Answer: This is a Windows PowerShell-script. To run it:
#
# - Save the contents as a `.ps1` file, e.g. `ScheduledTasksEnable.ps1`, e.g. to your desktop.
# - Run `cmd`
# - Execute the following command: `powershell`
# - `cd` into the folder where you saved the file
# - Execute the script by referencing the file, i.e.: `.\ScheduledTasksEnable.ps1 -scheduledTaskKeyword "Cloutility*"`
#
# 	> Asterisks (i.e. '*') in the `scheduledTaskKeyword` function as wildcards.
#
# If you get an error stating that the script is not digitally signed, execute the following command in PowerShell (without the quotes):
# `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`


# Block for declaring the script parameters.
Param($scheduledTaskKeyword)

# Get params from Advanced Installer as we're apparently not able to use
# external PowerShell-scripts with parameters. All references to "AI" cmdlets, i.e.
#
# - AI_GetMsiProperty
# - AI_SetMsiProperty
#
# ... are contained in try/catch-blocks to enable error messages during development
# in a non-Advanced Installer environment where the said cmdlets aren't available

try {
	$scheduledTaskKeyword = AI_GetMsiProperty ProductName;
	
	# Reset MSI property-values to support multiple, consecutive executions
	AI_SetMsiProperty ScheduledTasksEnable $false;
} catch {}

Write-Host "CUSTOM ACTION SCRIPT: ScheduledTasksEnable ($scheduledTaskKeyword)";
Write-Host "";


###############################################################################
# Init
###############################################################################

# Make a list of the relevant scheduled tasks
$tasks = Get-ScheduledTask | Select TaskName | Where {($_.TaskName -like $scheduledTaskKeyword)}
Write-Host "The following scheduled tasks have been found:" $tasks;

Write-Host "Stop and disable the scheduled tasks";
ForEach($task in $tasks){
	# Write-Host "task: $task";
	
	# Enable the scheduled task to allow it to run again
	Enable-ScheduledTask -TaskName $task.TaskName
}


#------------------------------------------------------------------------------
# CLEAN UP
#------------------------------------------------------------------------------

# As we want the software-installer to keep running regardless of how the
# script execution has proceeded, we always return `$true` for this script

try {
	# If running under "Advanced Installer"
	# Mark the process as success
	AI_SetMsiProperty ScheduledTasksEnable "True";
} catch {}

return $true;
