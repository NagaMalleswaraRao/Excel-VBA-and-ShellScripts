#Create a Task

$excel = new-object -comobject excel.application
$workbook = $excel.workbooks.open("E:\Batch file testing.xlsm")
$excel.Run("timestamp_macro")
$workbook.save()
$workbook.close()
$excel.quit()
