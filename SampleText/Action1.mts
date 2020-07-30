'On Error Resume Next 

aSubfolder=split(Environment("TestDir"),Environment("TestName"))
sSubfolderpath=aSubfolder(0)



TestCaseList=sSubfolderpath&"TestCaseList.xlsx"
GlobalDataTxt=sSubfolderpath&"\TestData\GlobalData.txt"

DataTable.ImportSheet TestCaseList, 1, "Global"



stotalTestCase=Datatable.GlobalSheet.GetRowCount

loadGlobalData(GlobalDataTxt)




For globalRow = 1 To stotalTestCase

Environment("GlobalTestRow")=globalRow

DataTable.GlobalSheet.SetCurrentRow(Environment("GlobalTestRow"))

sTestCaseName=DataTable("TestCaseName","Global")

If trim(sTestCaseName)="" Then
	Exit for
End If


If DataTable("Execution","Global")="Y" Then
	

Environment("CurrentTestCaseName")=sTestCaseName

sTestDatapath=sSubfolderpath & "TestData\" &sTestCaseName & ".xls"

sTestResultPath= sSubfolderpath & "Result\" & sTestCaseName & "_result.xls"

'check data file
if isfileexit(sTestDatapath) then


DataTable.ImportSheet sTestDatapath, 1, "Action1"

stotalstep=	DataTable.GetSheet("Action1").GetRowCount

Set IEObject=createobject("InternetExplorer.Application")


IEObject.Visible=true
IEHNWD=IEObject.HWND

window("hwnd:="&IEHNWD).Maximize


Environment("IEObject")=IEObject


TestPassFlag=true

'DataTable.ImportSheet sTestDatapath, 1, "Action1"
'test step running
For testCaseRow = 1 To stotalstep



Environment("testRow")=testCaseRow

DataTable.GetSheet("Action1").SetCurrentRow(Environment("testRow"))




sAction=DataTable("Page","Action1")

sActionDataRow=DataTable("DataReferenceRow","Action1")


If trim(sAction)="" Then
	Exit for
End If

Set dict = CreateObject("Scripting.Dictionary")

If isnumeric(sActionDataRow)=0 Then
	
	sActionDataRow=1
End If 'action row





'IEObject.Navigate "https://intra.qa.apps.rus.mto.gov.on.ca/iss/login.jsp"
'IESync(IEObject)

If left(sAction,3)="TC_" Then
	
sCurrentStepRow=Environment("testRow")


sReturn=RunAction ("TCAction", oneIteration,sAction)
DataTable.ImportSheet sTestDatapath, 1, "Action1"
Environment("testRow")=sCurrentStepRow


else

set dict=readDatafromExcel(sTestDatapath,sAction,sActionDataRow+1)

Environment("TestDict")=dict

sReturn=RunAction( sAction&" ["& sAction & "]", oneIteration,sActionDataRow)

	
End If



'IESync(IEObject)


DataTable.GetSheet("Action1").SetCurrentRow(Environment("testRow"))

If sReturn<>-1 Then
DataTable("Status","Action1")= "Passed"

else
DataTable("Status","Action1")= "Failed"


timestap=replace(now(),"/","_")

timestap=replace(timestap,":","_")

failedpicturename=sSubfolderpath&"Result\Failed\"&Environment("CurrentTestCaseName")&"_"&timestap&".png"

Desktop.CaptureBitmap failedpicturename

TestPassFlag=false
IEClose(IEObject)

wait(1)
cleanupBrower()
Set IEObject=nothing

Exit for

		
End If 'test case report



Call sendKeyESC()

Next 'test step run



DataTable.ExportSheet sTestResultPath, "Action1", "Action1"

IEClose(IEObject)
wait(1)
cleanupBrower()
Set IEObject=nothing

If TestPassFlag Then
	reporter.ReportEvent micPass, "Test Case " & sTestCaseName, "All Steps are passed status"
	
	call writeResultReport(TestCaseList,Environment("GlobalTestRow")+1,"Passed",sTestResultPath)
	
	else
	
	reporter.ReportEvent micFail, "Test Case " & sTestCaseName, "Failed in some steps, please check the detail",failedpicturename
	
	call writeResultReport(TestCaseList,Environment("GlobalTestRow")+1,"Failed",sTestResultPath)



End If ' test pass flag

else

reporter.ReportEvent micFail, "Test Case " & sTestCaseName, "Test Case data file not in folder, please check"
call writeResultReport(TestCaseList,Environment("GlobalTestRow")+1,"Missing Data","")



End  if	'data file


else 'execution control

call writeResultReport(TestCaseList,Environment("GlobalTestRow")+1,"Manual Skip","")

End If	'execution control



Datatable.GlobalSheet.SetCurrentRow(Environment("GlobalTestRow"))

	
Next ' test case run




exitaction (0) 'end testing










Function readDatafromExcel(spath,sSheetName,sRow)

Set datadic=CreateObject("Scripting.Dictionary")
Set OExcel=CreateObject("Excel.Application")
OExcel.Visible=false
set Oworkbook=OExcel.Workbooks.Open(spath)

sNumberSheet=Oworkbook.Worksheets.Count

SheetFlag=false
For Iterator = 1 To sNumberSheet
	
	If sSheetName= Oworkbook.Worksheets(Iterator).name then
		
		SheetFlag=true
		Exit for
	End If
	
Next

If SheetFlag=0 Then

OExcel.Quit

Set Oworkbook=nothing
Set OExcel=nothing
	Exit function
	
End If

'add the data to the dictionary
j=1
Do  while(trim(Oworkbook.Worksheets(sSheetName).cells(1,j))<>"")

sKey=Oworkbook.Worksheets(sSheetName).cells(1,j)
sValue=Oworkbook.Worksheets(sSheetName).cells(sRow,j)

datadic.Add sKey,sValue

j=j+1
loop

Oworkbook.close False

OExcel.Quit

Set Oworkbook=nothing
Set OExcel=nothing


	
	set readDatafromExcel=datadic
	
	
End Function


Function isfileexit(sfilepath)
	
	Set fso = CreateObject("Scripting.FileSystemObject")  
	
	If (fso.FileExists(sfilepath)) Then  
		isfileexit=true
		else
		
	isfileexit=false
	
	End  if


	
End Function

Function writeResultReport(spath,sRow,sStatus,resultpath)
	
	sStatusCol = 4
	sResultFileCol = 5
	Set OExcel=CreateObject("Excel.Application")
OExcel.Visible=false
set Oworkbook=OExcel.Workbooks.Open(spath)


Oworkbook.Worksheets(1).cells(sRow,sStatusCol)=sStatus


If resultpath<>"" Then
Oworkbook.Worksheets(1).Hyperlinks.add Oworkbook.Worksheets(1).cells(sRow,sResultFileCol),resultpath,,,"Detail Step Result"
	
End If

Oworkbook.Save
Oworkbook.Close
OExcel.Quit

Set Oworkbook=nothing
Set OExcel=nothing
	
	
End Function

Function loadGlobalData(datafilepath)
	
	Set objFSO=CreateObject("Scripting.FileSystemObject")	
	
Set file = objFSO.OpenTextFile(datafilepath, 1)
row = 0
Do Until file.AtEndOfStream
  line = file.Readline
  If trim(line)<>"" Then
  	
  	uLine=split(line,"=")
  	
  	sBound=Ubound(uLine)
  	
  	If sBound>0 Then
  		
  		EParaName=uLine(0)
  		sParaValue=""
  		For Iterator = 1 To sBound Step 1
  			
  			sParaValue= sParaValue & uLine(Iterator)
  			
  		Next
  		Environment(EParaName)=sParaValue
  		
  	End If
  	
  	
  End If
  
  
Loop
	

Set file=nothing
Set objFSO=nothing

	
	
End Function
