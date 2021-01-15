'===========================================================
'20201007 - DJ: Initial credation
'20210107 - DJ: Added clicking the logout link and a synchronization of the proposal creation status message.
'20210115 - DJ: Disabled smart identification for runs
'===========================================================

'===========================================================
'Function to Create a Random Number with DateTime Stamp
'===========================================================
Function fnRandomNumberWithDateTimeStamp()

'Find out the current date and time
Dim sDate : sDate = Day(Now)
Dim sMonth : sMonth = Month(Now)
Dim sYear : sYear = Year(Now)
Dim sHour : sHour = Hour(Now)
Dim sMinute : sMinute = Minute(Now)
Dim sSecond : sSecond = Second(Now)

'Create Random Number
fnRandomNumberWithDateTimeStamp = Int(sDate & sMonth & sYear & sHour & sMinute & sSecond)

'======================== End Function =====================
End Function

Dim BrowserExecutable, ProposalName, ExecutiveOverview, rc

While Browser("CreationTime:=0").Exist(0)   												'Loop to close all open browsers
	Browser("CreationTime:=0").Close 
Wend
BrowserExecutable = DataTable.Value("BrowserName") & ".exe"
SystemUtil.Run BrowserExecutable,"","","",3													'launch the browser specified in the data table
Set AppContext=Browser("CreationTime:=0")													'Set the variable for what application (in this case the browser) we are acting upon

'===========================================================================================
'BP:  Navigate to the PPM Launch Pages
'===========================================================================================

AppContext.ClearCache																		'Clear the browser cache to ensure you're getting the latest forms from the application
AppContext.Navigate DataTable.Value("URL")													'Navigate to the application URL
AppContext.Maximize																			'Maximize the application to give the best chance that the fields will be visible on the screen
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Executive Overview link
'===========================================================================================
Browser("Browser").Page("PPM Launch Page").Image("Strategic Portfolio Image").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Tina Fry (Business User) link to log in as Tina Fry
'===========================================================================================
Browser("Browser").Page("Portfolio Management").WebArea("Tina Fry Link").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Create menu item
'===========================================================================================
Browser("Browser").Page("PPM Page").Link("CREATE").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Click the Proposal text
'===========================================================================================
Browser("Browser").Page("Dashboard - Request Status").Link("Proposal").Click
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Select the "Corporate" value in the Business Unit combobox
'===========================================================================================
Browser("Browser").Page("PPM Page").WebList("Business Unit Combobox").Select "Corporate"

'===========================================================================================
'BP:  Type a unique proposal name into the Proposal Name field
'===========================================================================================
ProposalName = "Proposal Name " & fnRandomNumberWithDateTimeStamp
Browser("Browser").Page("PPM Page").WebEdit("Project Name").Set ProposalName

'===========================================================================================
'BP:  Enter unique text into the Executive Overview: field
'===========================================================================================
ExecutiveOverview = DataTable.Value("dtExecutiveOverview") & ProposalName
Browser("Browser").Page("PPM Page").WebEdit("ExecutiveOverview").Set ExecutiveOverview

'===========================================================================================
'BP:  Enter the Business Objective value
'===========================================================================================
Browser("Browser").Page("PPM Page").WebEdit("Business Objective Combobox").Set DataTable.Value("BusinessObjective")

'===========================================================================================
'BP:  Click the Submit text
'===========================================================================================
Browser("Browser").Page("PPM Page").Link("Submit").Click
rc = Browser("Browser").Page("PPM Menu").WebElement("Your Request is Created").Exist(30)
AppContext.Sync																				'Wait for the browser to stop spinning

'===========================================================================================
'BP:  Logout
'===========================================================================================
Browser("Browser").Page("PPM Menu").WebElement("menuUserIcon").Click
AppContext.Sync																				'Wait for the browser to stop spinning
Browser("Browser").Page("PPM Menu").Link("Sign Out").Click

AppContext.Close																			'Close the application at the end of your script

