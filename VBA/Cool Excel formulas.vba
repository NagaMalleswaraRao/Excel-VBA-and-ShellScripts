'''To get the Datasource (named range) dynamically to a Pivot in another workbook
'Enter this formula ="'"&SUBSTITUTE( LEFT(CELL("filename",A1),FIND("]",CELL("filename",A1))-1),"[","")&"'!CPU_DATA"
