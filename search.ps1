#insert full path for excel files here, separated by comma
$myFile = ()

$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $false

Foreach($file in $myFile)
{
	$Workbook = $Excel.workbooks.open($file)
	echo "`nFilename: $file"
	Foreach($sheet in 1..$Workbook.$Worksheets.count)
	{
		$worksheetname = $Workbook.sheets.item($sheet)
		echo "`nWorksheet : $($worksheetname.name)"
		$Worksheet = $Workbook.Worksheets.Item($sheet)
		#insert string to be search here
		$SearchString = ""
		$Range = $Worksheet.UsedRange
		$Search = $Range.find($SearchString)
		echo "Found string: $($Search.text)"
	}
$Excel.workbooks.close()
