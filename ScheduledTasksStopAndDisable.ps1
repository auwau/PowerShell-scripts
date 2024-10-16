# This script stops and disables any scheduled task which name matches the
# `scheduledTaskKeyword` parameter value which can potentially be running
# (including program instances instantiated by previous/obsolete scheduled
# tasks) and thus potentially cause deadlocks ("race conditions") when applying
# database updates.
#
# Question: WTF(lip) is this, and how do I run it?
#
# Answer: This is a Windows PowerShell-script. To run it:
#
# - Save the contents as a `.ps1` file, e.g. `ScheduledTasksStopAndDisable.ps1`, e.g. to your desktop.
# - Run `cmd`
# - Execute the following command: `powershell`
# - `cd` into the folder where you saved the file
# - Execute the script by referencing the file, i.e.: `.\ScheduledTasksStopAndDisable.ps1 -scheduledTaskKeyword "Cloutility*"`
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
	AI_SetMsiProperty ScheduledTasksStopAndDisable $false;
} catch {}

Write-Host "CUSTOM ACTION SCRIPT: ScheduledTasksStopAndDisable ($scheduledTaskKeyword)";
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
	
	# Stop the scheduled task (if it's running)
    Stop-ScheduledTask -TaskName $task.TaskName;
	
	# Disable the scheduled task to prevent it starting again
	Disable-ScheduledTask -TaskName $task.TaskName;

	
	# Scheduled tasks are run by the task scheduler starting an instance of its
	# target program, which keeps on running until the process completes. If you
	# `end` a scheduled task then the related process will be killed off as well,
	# however if you `delete` a scheduled task which has a running process (as the
	# software-installer can do, depending on when it's run) then the process from
	# the previous/obsolete scheduled task (with the same name) will continue
	# running regardless of what you do with its replacement. Hence, as we have now
	# (above) made sure that no new scheduled tasks start, we now have to ensure
	# that no instances of previous/obsolete scheduled tasks are running. As the
	# previous/obsolete task references the same program, we can use the current
	# task's `Actions` to obtain the program name.

	# Identify the scheduled task's action, i.e. the program which is run
	$actions = (Get-ScheduledTask -TaskName $task.TaskName).Actions.Execute; # Full path
	# Write-Host "Actions: $actions"; 

	# Extract the name of the program, as this is what is referenced in Windows'
	# "task manager", i.e. we're referencing the program name and not the name of
	# the scheduled taks responsible for initiating the process.
	$programName = Split-Path $actions -leaf; # Last item in the list
	# Write-Host "programName: $programName"; 

	# Kill the target process (if it's running)
	taskkill /F /IM $programName;
}


#------------------------------------------------------------------------------
# CLEAN UP
#------------------------------------------------------------------------------

# As we want the software-installer to keep running regardless of how the
# script execution has proceeded, we always return `$true` for this script

try {
	# If running under "Advanced Installer"
	# Mark the process as success
	AI_SetMsiProperty ScheduledTasksStopAndDisable "True";
} catch {}

return $true;
