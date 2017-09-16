Dim QTPObj,QTPTest
Set QTPObj=CreateObject("QuickTest.Application")

'Check if the application is not already Launched
If Not QTPObj.Launched then
	QTPObj.Launch
end if

QTPObj.Visible=True
QTPObj.Open "D:\UFT\FormTest" 'name of the start up script
Set QTPTest=QTPObj.Test
QTPTest.Run 'Run the Test
'QTPTest.Close 'Close the Test
'QTPObj.Quit 'Quit the QTP Application

