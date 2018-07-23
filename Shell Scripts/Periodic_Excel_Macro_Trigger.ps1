# Create a Task Scheduler, so that it will call a PowerShell script, which in turn calls
# a VBA macro residing in Excel at periodic intervals

# The below code is the ".ps1" script which opens Batch file testing file (xlsm/xlsb) and runs the "timestamp_macro" VBA subroutine
# Go to Task Scheduler and create a new task
#   1. In General tab, provide a name for the job
#   2. In Triggers tab, provide a time trigger
#   3. In Actions tab, for the "Start a program" action
#         a. Provide PowerShell.exe location "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
#         b. In the Add arguments box, provide this "-ExecutionPolicy Bypass -File E:\Periodic_Run_PowerShell.ps1"
#         c. In the Start in box, provide this "E:\" {path of the powershell script}

$excel = new-object -comobject excel.application
$workbook = $excel.workbooks.open("E:\Batch file testing.xlsm")
$excel.Run("timestamp_macro")
$workbook.save()
$workbook.close()
$excel.quit()
