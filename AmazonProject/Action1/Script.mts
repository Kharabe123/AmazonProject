
SystemUtil.CloseProcessByName"chrome.exe"
SystemUtil.Run"chrome.exe","www.amazon.in"
'On Error Resume Next
DataTable.AddSheet "TestData1"
DataTable.ImportSheet "C:\Users\user239\Documents\AmazonProject\TestData1\TestData1.xlsx","AmazonData","TestData1"

rowCount=DataTable.GetSheet("TestData1").GetRowCount

For rows=1 To rowCount

DataTable.SetCurrentRow rows

If DataTable.Value("Execution_flag","TestData1")="Y" Then

executeTest(DataTable.Value("TestCaseID","TestData1"))	
DataTable.Value("Result","TestData1")= Environment.Value("Result")
End If

Next
DataTable.ExportSheet "C:\Users\user239\Documents\AmazonProject\TestData1\TestData1.xlsx","TestData1","AmazonData"






