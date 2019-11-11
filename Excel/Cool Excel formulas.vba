'''To get the Datasource (named range- here it is 'CPU_DATA' ) dynamically to a Pivot in another workbook
'Enter this formula ="'"&SUBSTITUTE( LEFT(CELL("filename",A1),FIND("]",CELL("filename",A1))-1),"[","")&"'!CPU_DATA"

'''Remove unnecessary rows (FY=0)
'=IF(AND(T2=0,U2=0,V2=0,W2=0,X2=0,Y2=0,Z2=0,AA2=0,AB2=0,AC2=0,AD2=0,AE2=0),"Remove","Keep")



