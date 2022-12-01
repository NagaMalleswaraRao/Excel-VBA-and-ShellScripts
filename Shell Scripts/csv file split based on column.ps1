$InputFilePath = "C:\Zepto - Work Files\Inventory Dashboard\Inventory Report\Inventory File Split\Inventory Instock Dump.csv"
$SplitByColumnName = "category_name" #Enter ColumnName here on basis of which you want to split.

$data = Import-Csv $InputFilePath | Select -ExpandProperty $SplitByColumnName -Unique

$a = $data | select 

ForEach ($i in $a)
{  
  $FinalFileNamePath = "C:\Zepto - Work Files\Inventory Dashboard\Inventory Report\Inventory File Split\" + $i + ".CSV" #This is where you would keep the splitted files.

  Import-Csv $InputFilePath | where {$_.$SplitByColumnName -eq $i } | Export-Csv $FinalFileNamePath -NoTypeInformation  
}
