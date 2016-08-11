# Chorus-UFT

Option Explicit
Dim Iterator
Public WFC_Count_Temp
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

' Option Explicit ensures that all variables used withing functions are declared before being used.  This avoids potential problems.

' *******************************************************************************************
' FUNCTION LIBRARY NAME:
' 
' SSP_Functions

' FUNCTION LIBRARY DESCRIPTION:
' 
' This function library contains functions specific to the BPT (Business Process Testing)
' automation of the Chorus SSP (Self Service Portal) application.
' SSP is a browser-based web application used to provision UFB products.
' These functions enhance the out-of-the-box BPT functionality.
' 
' *******************************************************************************************

' List of SSP functions:
' RJ updated the list of functions at 28/07/2016 
' CloseAllBrowserExceptALM() ' Closes all browsers exept ALM (QC)
' SSP_Check_Session_Timed_out()
' SSP_Launch_Browser()
' SSP_Clear_Security_Certificate()
' SSP_Enter_Login_Details(UserID)
' SSP_Check_Login_Page()
' SSP_Nav_Create_Order()
' SSP_Nav_Work_Queue()
' SSP_Nav_Search_Orders()
' SSP_Location_Search_TLC(TLC)
' SSP_Location_Search_SAM(SAM_ID)
' SSP_Product_Search_Product_ID(Product_ID)
' SSP_Location_Search_Exchange_ID(Exchange_ID)
' SSP_Logout()
' ChorusUtilities_AttachFileToALMTestRun(FileName)
' Chorus_
'Select_RSP (RSP)
' SSP_Location_Search_StreetNo_StreetName_Town(StreetNo,StreetName,Town)
' SSP_Location_Search_StreetName(StreetName)
' SSP_Select_Product_Offer(prodoffer)
' SSP_SelectAim(Passedaim)
' SSP_SelectCSE(PassedCSE)
' SSP_Check_Order_Type()
' SSP_Capture_Button()
' SSP_Check_Feasibility_Message()
' SSP_Check_Feasibility_Message_Manualinvestigation()
' SSP_Get_Order_ID()
' SSP_Get_Location_ID()
' SSP_Check_Feasibility_Message_ManualFeasibility()
' SSP_Check_Feasibility_Message_SF_Fails()
' SSP_Click_Validate_Location
' SSP_Validate_Req_Enter_Customer_Details (Name,Email)
' SSP_Validate_Req_Select_NBAP_Type
' SSP_Validate_Req_NBAP_Location_Details (HouseNo,Coordinate)
' SSP_Validate_Req_NBAP_Click_Send
' SSP_Validate_Get_RequestID
' SSP_Validate_Select_Address_Type
' SSP_Validate_Attach_Map_And_PropertyType (NewLocationAddress,CustomerLetterBox,BuildingLocation,AddressTenancyAgreement,NearestNeighboursAddress,ASID,FilePath)
' SSP_wait_mins(timeinMis)
' SSP_Validate_Location_Request_Creation_And_Assignment(ReqID,Taskassinee)
' SSP_Select_Order_Type()
' SSP_SelectAim_Layer1(Passedaim)
' SSP_Enter_Site_Information_Details(Input_Name,Input_Phone_No)
' SSP_Submit_Feasibility()
' SSP_GET_Feasibility_ORDER_ID()' returns the order id
' SSP_Order_Nav_Order_Task()
' SSP_Order_Nav_Order_Order_Details()
' SSP_Order_Nav_Order_Schedule()
' SSP_Nav_Summary_Verify_Order_Sub_Status(Sub_status)
' SSP_Task_Perform_Feasibility(Task_Outcome)
' SSP_Order_Nav_Order_Summary()
' SSP_Appointment_Go_To_Next_Month()
' SSP_Appointment_Select_First_Available_Date()
' SSP_Appointment_Select_Scope_Date()
' SSP_Appointment_Select_Install_Date()
' SSP_Appointment_Select_PM_SLOT()
' SSP_Appointment_Select_AM_SLOT()
' SSP_Appointment_Reserve_Appointment()
' SSP_Offer_Characteristics_Set_Default_Values()
' SSP_Order_Check_Scope_Calender()
' SSP_Select_Product_Instance()
' SSP_Select_product_Instance_Secondary()
' SSP_Click_on_Next_Button_to_Get_Aim()
' SSP_Order_Nav_Offer_Characteristics()
' SSP_Error_message_Appointment_Not_Reserved()
' SSP_Submit_Order()
' SSP_Appointment_check_SLA(min_no_days)
' SSP_ManageConsent_Verify_AppointmentPage
' SSP_ManageConsent_Nav_Task_ManageConsentTaskCreated
' SSP_ManageConsent_ManualTask_Complete(Consentstatus,NextStep,SendNotification)
' SSP_ManageConsent_ManualTask_Save(ConsentStatus,SendNotification)
' SSP_ManageConsent_ManualTask_StatusNotificationSenttoRSP
' SSP_MDUBuild_ManualTask_Complete(BuildStatus,NextStep,SendNotification)
' SSP_MDUDesign_ManualTask_Complete(DesignStatus,NextStep,SendNotification)
' SSP_Appointment_Check_Appointment_Types(Expected_Scope,Expected_Install,Expected_CSE,Expected_RFS)
' SSP_ManageConsent_ConsentQueue_ConsentTaskIsAssociated
' SSP_Check_OrderSubmittedSuccessfully
' SSP_Nav_Summary_Verify_OrderStatus(Status)
' SSP_click_on_Cancel_Button()
' SSP_Appointment_Non_Standard_Install_Pending_Install()
' SSP_ManageConsent_ManualTask_DisassociateOrder(Action)
' SSP_Appointment_Book_First_Date_Intact_Connect()
' SSP_Task_Initiate_Cancel_Order()
' SSP_Task_Chorus_Approve_Cancel_Order()
' SSP_Chorus_Complete_Order_Cancellation()
' SSP_Appointment_Check_After_cancellation()
' SSP_Prod_Inv_Administartion()
' SSP_Prod_Inv_Nav()
' SSP_Prod_Inv_Nav_Create_New_Record()
' SSP_Prod_Inv_Nav_Create_New_Record_Service_Provider(service_Provider)
' SSP_Prod_Inv_Nav_Create_New_Record_Loc_ID(Loc_ID)
' SSP_Prod_Inv_Nav_Create_New_Record_Prod_ID(Prod_ID)
' SSP_Prod_Inv_Nav_Create_New_Record_Product(Product)
' SSP_Prod_Inv_Create_Record_Loc_Reference()
' SSP_Prod_Inv_Update_Record()
' SSP_Prod_Inv_Product_Status()
' SSP_Prod_Inv_Offer_Characteristics_Set_Default_Values()
' SSP_Prod_Inv_Submit_Record()
' SSP_Prod_Inv_Search_Record(service_Provider_2,Prod_ID_2)
' SSP_Prod_Inv_Display_Product_Characteristics()
' SSP_Prod_Inv_Query_Terminate_Date_Blank()
' SSP_Prod_Inv_Terminate_Product_Instance()
' SSP_Appointment_Update_Date_Other_than_First_date() 
' SSP_Appointment_Book_Other_Than_First_Date_Intact_Connect()
' SSP_Appointment_Reschedule_Appointment()
' SSP_Order_History_tab()
' SSP_ManualTask_Manage_Complex_Design(Next_Step)
' SSP_ManualTask_Manage_Complex_Build(Next_Step)
' SSP_ManualTask_Complete_Button()
' SSP_ManualTask_Continue_Button()
' SSP_Manage_Design_Build_Exception_Task()
' SSP_Manage_Design_Build_Cancel()
' SSP_Order_Nav_Summary_Status()
' SSP_ManageConsent_Select_OfferType(Offer_Type)
' SSP_Offer_Char_Get_RFS_Date()
' SSP_Select_OfferType(Offer_Type)
' SSP_ManualTask_MDU_Design(Design_status,Next_Step)
' SSP_ManualTask_MDU_Build(Design_status,Next_Step)
' SSP_Summary_Verify_Order_Status(status)
' SSP_Check_Design_Build_Message()
' SSP_Manage_Design_Build_OrderID_TaskIsAssociated
' SSP_MDU_Build_Task_OrderID_TaskIsAssociated
' SSP_Manual_Fibre_Feasibility_Message()
' SSP_Manage_Design_Build_Associate_Order()
' SSP_Manage_Design_Build_Manual_Task_Cancelled()
' SSP_Appointment_Chorus_Amending_On_Behalf_Of_RSP(RSP)
' SSP_Task_Perform_Billing()
' SSP_Offer_Charac_set_Handover_connection(Handover)
' Get_OrderDetails_into_Excel(FileName,Product,WFC_Count)
' Create_Excel(FileName,Product)
' FormatDatein_YYYY_MM_DD_HH_MM()
' Get_ProductID(Product)
' OpenExcel_Fill_FirstColumn(FileName)
' OpenExcel_Fill_OtherColumns(FileName,Product)
' SSP_Modify_Switch_Handover_connection(Handover1,Handover2)
' SSP_Modify_Switch_VLAN_connection()
' SSP_SelectLosing_Service_ProviderName(Losing_RSP)
' SSP_Move_Enter_ProductID(ProductID)
' SSP_Offer_Characteristics_Save_Button()
' SSP_Offer_Characteristics_Change_Voice_Details(Voice_Handover_ID,Voice_SVID,Voice_CVID)
' SSP_Offer_Characteristics_Change_Data_Details(Data_Handover_ID_1,Data_SVID_1,Data_CVID_1, Data_Handover_ID_2,Data_SVID_2,Data_CVID_2, Data_Handover_ID_3,Data_SVID_3,Data_CVID_3)
' SSP_Offer_Characteristics_Change_VLAN_SVID(VLAN_SVID)
' Verify_SSP_Order_SubStatus_Not_Received()
' SSP_ICMS_FILE_Creation_Scope(Order_ID)
' SSP_ICMS_FILE_Creation_Install_Service_Given(Order_ID)
' SSP_ICMS_FILE_Creation_Install_Completed(Order_ID)
' GetFile(FileName)
' WriteFile(FileName_change, Contents)
' GetFileName()
' GetLine(FileName)
' GetLastLine(FileName)
' SSP_Task_Create_Scope_Work_Order(ICMS_SO_Nos)
' SSP_Task_Create_Install_Work_Order(ICMS_SO_Nos)
' SSP_Task_Create_CSE_Work_Order(ICMS_SO_Nos)
' SSP_ICMS_Validate_Scope_Completed()
' SSP_ICMS_Validate_Install_CSE_Completed()
' SSP_ICMS_Validate_Install_CSE_Service_Given()
' ICMS_Update_And_Upload_Winscp_BatchFile(FilePath)
' ICMS_Invoke_Winscp_Batch_File()
' SSP_Driver_To_Call_ICMS_Functions(Order_ID,Appointment_Check_Results)
' SSP_Manual_Intervention()
' SSP_Check_If_Cancel_Button_Present()
' SSP_SearchOrder_Search_Order_ID_1(OrderID)
' SSP_SearchOrder_Click_on_Searched_Order_2(OrderID)
' SSP_Utilities_To_Cancel_All_Orders_Present_in_Cancellation_Database()
' SSP_Summary_Verify_Order_Status_SubStatus(status,SubStatus)
' SSP_Nav_Notifications()
' SearchAndActionNotificationAtTLC(TLC,Action)
' SSP_Chorus_Add_Interaction()
' SSP_Chorus_Add_Interaction_Bulk(NumInteractions)
' SSP_Launch_Browser_Brendon(Environment)
' SSP_Enter_Login_Details_Brendon(UserID, Environment)
' SSP_Charges_Add_Quote()
' SSP_Amend_Amending_on_behalf_of()
' SSP_Amend_Click_on_Update_Order_Task()
' SSP_Amend_Complete_Amend_Task()
' SSP_Amend_Contact_Details()
' SSP_Amend_Customer_Interaction()
' SSP_Amend_Click_on_Manage_Customer_Interaction()
' SSP_Amend_Complete_Manage_Customer_Interaction_Task()
' SSP_Amend_Escalate_Order()
' SSP_Amend_Click_on_Process_Escalation_Task()
' SSP_Amend_Task_Complete_Process_Escalation(EscalationDecision)
' SSP_Select_RescheduleReason()
' SSP_Reschedule_Confirm_Charges()
' SSP_Select_Order_Task(TaskName)
' SSP_Validate_Default_CSE_Date()
' SSP_Validate_Default_RFS_Date()
' SSP_Validate_Default_Scope_Install_Dates()

'**********************************************************************************************
Function CloseAllBrowserExceptALM() ' Closes all browsers exept ALM (QC)

Dim ObjBrow,ObjName,ObjBrowName,brCnt,ic

Set ObjBrow = Description.Create()
ObjBrow("micClass").Value = "Browser"
Set ObjName = Desktop.ChildObjects(ObjBrow)
brCnt=ObjName.Count

For ic = 0 To  brCnt-1
	ObjBrowName = ObjName(ic).getROproperty("Name")
		If InStr("HP Application", ObjBrowName) <> 0 Then
				' do Nothing"
			Else 
				ObjName(ic).Close
		End If
Next

End Function


Function SSP_Check_Session_Timed_out()
	DIM SSP_Home_Desc
	Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	
	If SSP_Home_Desc.Link("innerhtml:=Login to Chorus Self Service Portal").Exist(4) = True Then
		
		SSP_Home_Desc.Link("innerhtml:=Login to Chorus Self Service Portal").Click
		
	End If
	SSP_Clear_Security_Certificate
	
End Function


'@Description Launches the Browser
'@Documentation Launches the Browser with the URL given in the parameter
Function SSP_Launch_Browser()
	' Current desktop policy means that the Screensaver kicks in after 10 minutes and locks the desktop.  This prevents scheduled tests from running.
	' To get around this, a utility called MouseJiggle.exe can be used to simulate the mouse being moved every 5 seconds.  However, when the test is running
	' it's best to shutdown MouseJiggle using the following CloseProcessByName.
systemutil.CloseProcessByName "MouseJiggle.exe"
' 'Close all of the open browsers to ensure that no other browser is open
	'CloseAllBrowserExceptALM
CloseAllBrowserExceptALM
'***************************
Dim URL_Name : URL_Name = "SSP_URL"
Dim i, URLAddress

Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_SIT",  "URLs" , dtLocalSheet
dim Rowcount : Rowcount= DataTable.GetSheet(dtLocalSheet).GetRowCount
For i =1 To Rowcount

  
	If strcomp(trim(DataTable("coffiiURLs",dtLocalSheet)),trim(URL_Name),1) = 0 Then
		URLAddress = DataTable("URLAddress",dtLocalSheet)
		Systemutil.Run "iexplore.exe" ,URLAddress
		SSP_Check_Session_Timed_out
		Exit for
	else
		DataTable.setnextrow   
	End If	
Next


If i > Rowcount Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Login", "Please check the URL Details in Resources"
End If



End Function

'@Description Before the login page , there is an issue of website’s security certificate. We need to bypass it
'@Documentation Bypass the Internet Explorer warnings about the website’s security certificate.
Function SSP_Clear_Security_Certificate()

Dim Odesc,i
Set Odesc = Browser("name:=Certificate Error: Navigation Blocked").Page("title:=Certificate Error: Navigation Blocked")
'wait(6)
' Clear Security
For i = 0 To 1
'wait(10)
	If Odesc.Webtable("Name:=Shield icon").Link("innerhtml:=Continue to this website \(not recommended\)\.").Exist(10) = True Then
		oDesc.Webtable("Name:=Shield icon").Link("innerhtml:=Continue to this website \(not recommended\)\.").Click	
		wait(1)
				If Browser("name:=Sign In").Page("Title:=Sign In").Exist(1) =True Then
							Exit for
							
					End If
		
	else
	 	Reporter.ReportEvent micDone ,"Clear Security" , "SSP bypassed the security check"
	End IF
Next
End Function


'@Description Enter login details
'@Documentation Enter login details
Function SSP_Enter_Login_Details(UserID)

Dim SSPPassword,i
Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_SIT",  "SSP_login_Info" , dtLocalSheet 
dim Rowcount : Rowcount= DataTable.GetSheet(dtLocalSheet).GetRowCount
For i =1 To Rowcount

	If strcomp(trim(DataTable("UserID",dtLocalSheet)),trim(UserID),1) = 0 Then
		SSPPassword = DataTable("Password",dtLocalSheet)
		Exit for
	else
		DataTable.setnextrow   
	End If	
Next


If i > Rowcount Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Login", "Please check the login Details in Resources"
End If
' Set the condition of Exiting test case in Business components

If Browser("name:=Sign In").Page("Title:=Sign In").Exist(1) =True Then
	Browser("name:=Sign In").Page("Title:=Sign In").WebEdit("Class Name:=WebEdit","html id:=ContentPlaceHolder1_UsernameTextBox").Set UserID
	Browser("name:=Sign In").Page("Title:=Sign In").WebEdit("Class Name:=WebEdit","html id:=ContentPlaceHolder1_PasswordTextBox").Set SSPPassword
	wait(0.5) ' To sync up with the Page
	Browser("name:=Sign In").Page("Title:=Sign In").WebButton("name:=Sign In").Click
	Reporter.ReportEvent micDone, "Login", "Submitted login credentials from the Sign In page"
Else
	Reporter.ReportEvent micDone, "Login", "SSP bypassed the login credentials"
End If

End Function


'@Description Check the landing page
'@Documentation Check the landing page
Function SSP_Check_Login_Page()

Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebElement("class:=banner small navigation-banner-link").Exist(4) = True Then
		Reporter.ReportEvent micPass, "Login" , "Success"
		'msgbox "logged in"
Else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Login" , "Failure - could not find SSP banner after logging in"
		
End If
End Function

'@Description Click on create order button
'@Documentation Click on create order button
Function SSP_Nav_Create_Order()

'Dim SSP_Home_Desc
'Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

wait(5)

Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("outerhtml:=<a href=""/chorus-ssp-web/pages/OrderManagement/InitiateOrder"" s-key=""c"" s-key-type=""1"">Create Order</a>").Click
'SSP_Home_Desc.Link("outerhtml:=<a href=""/chorus-ssp-web/pages/OrderManagement/InitiateOrder"" s-key=""c"" s-key-type=""1"">Create Order</a>").Click
wait(1)
'SSP_Home_Desc.Sync
End Function

Function SSP_Nav_Work_Queue()

'Dim SSP_Home_Desc
'Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

wait(1)

Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("outerhtml:=<a href=""/chorus-ssp-web/pages/WorkQueue/WorkQueue"" s-key=""w"" s-key-type=""1"">Work Queue</a>").Click
wait(1)
'SSP_Home_Desc.Sync
End Function

'@Description Function to Navigate to Notification Tab
'@Documentation Function to Navigate to Notification Tab
Function SSP_Nav_Search_Orders()

Dim SSP_Home_Desc

set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")

SSP_HOME_Desc.Link("outerhtml:=<a href=""/chorus-ssp-web/pages/OrderManagement/Search"" s-key=""s"" s-key-type=""1"">Search Orders</a>").Click

SSP_Home_Desc.Sync

End Function


'@Description Enter the TLC ID and search
'@Documentation Enter the TLC ID and search
Function SSP_Location_Search_TLC(TLC)


Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebEdit("html id:=locationSearchId").Set TLC
wait(1)

Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebButton("html id:=locationSearchButton").Click
wait(1) ' to sync up if the search is taking up time


' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Location ID: "&TLC&" " ' this is the property of TLC pane

if Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Location Search", "TLC Search functionality is working"
else
	Reporter.ReportEvent  micFail, "Location Search", "TLC Search functionality is not working or the TLC is wrong"
	
End If

End Function


'@Description Enter the SAM ID and search
'@Documentation Enter the SAM ID and search
Function SSP_Location_Search_SAM(SAM_ID)

Browser("SSP").Page("SSP Home Page").WebEdit("locationSearchId").Set SAM_ID
wait(1)
Browser("SSP").Page("SSP Home Page").WebList("Location Search Dropdown").Select("SAM ID")
Browser("SSP").Page("SSP Home Page").WebButton("Location Search").Click
wait(1) ' to sync up if the search is taking up time


' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "SAM ID: "&SAM_ID&" " ' this is the property of SAM ID pane . ex : SAM ID: 1244439 

if Browser("SSP").Page("SSP Home Page").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Location Search", "SAM Search functionality is working"
else
	Reporter.ReportEvent  micFail, "Location Search", "SAM Search functionality is not working or the SAM ID is wrong"
	
End If

End Function

'@Description Enter the Product ID and search
'@Documentation Enter the Product ID and search
Function SSP_Product_Search_Product_ID(Product_ID)

Browser("SSP").Page("SSP Home Page").WebEdit("productSearchId").Set Product_ID
wait(1)
Browser("SSP").Page("SSP Home Page").WebList("Product Search Dropdown").Select("Product ID")
Browser("SSP").Page("SSP Home Page").WebButton("html id:=productSearchButton","name:=Search").Click
wait(1) ' to sync up if the search is taking up time


' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
'Dim oDesc 'Declare an object variable
'Set oDesc = Description.Create 'Create an empty description
'oDesc("Class Name").value= "WebElement"
'oDesc("innertext").value= "Product ID: "&Product_ID&" " ' this is the property of SAM ID pane . ex : SAM ID: 1244439 
'
'if Browser("SSP").Page("SSP Home Page").Webelement(oDesc).Exist = True then
'	Reporter.ReportEvent micPass, "Location Search", "Product Search functionality is working"
'else
'	Reporter.ReportEvent  micFail, "Product Search", "PRODUCT Search functionality is not working or the Product ID is wrong"
'	
'End If

End Function

'@Description Enter the Exchange ID and search
'@Documentation Enter the Exchange ID and search
Function SSP_Location_Search_Exchange_ID(Exchange_ID)

Browser("SSP").Page("SSP Home Page").WebEdit("locationSearchId").Set Exchange_ID
Browser("SSP").Page("SSP Home Page").WebList("Location Search Dropdown").Select("Exchange Code")
Browser("SSP").Page("SSP Home Page").WebButton("Location Search").Click
wait(1) ' to sync up if the search is taking up time

' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Central Office Code: "&Exchange_ID' this is the property of SAM ID pane . ex : Central Office Code: TWL

if Browser("SSP").Page("SSP Home Page").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Location Search", "Exchange Code Search functionality is working"
else
	Reporter.ReportEvent  micFail, "Location Search", "Exchange Code Search functionality is not working or the Exchange ID is wrong"
	
End If

End Function



'@Description Takes a screenshot of the full screen and stores it both locally and within ALM Test Run.
'@Documentation Take a screenshot and store it somewhere.
Public Function  ChorusUtilities_TakeScreenshot()
	Dim objFSO, fileName, screenshotPath, fullScreenshotPathAndFileName

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists("E:\Test_Automation\Screenshots\") Then
		'Test tools server
 		screenshotPath = "E:\Test_Automation\Screenshots\"
 	Else
 		'VDI
 		screenshotPath = "D:\"
	End If
	
	fileName =  "Timestamp=" & ChorusUtilities_GetCurrentDateTimestamp & " TestName=" & Environment("TestName")
	fullScreenshotPathAndFileName = screenshotPath & fileName & ".png"
	Browser("name:=Chorus Self Service Portal").CaptureBitmap fullScreenshotPathAndFileName
	'Window("regexpwndtitle:=Program Manager", "regexpwndclass:=Progman").CaptureBitmap fullScreenshotPathAndFileName
	ChorusUtilities_AttachFileToALMTestRun(fullScreenshotPathAndFileName)
	Reporter.ReportEvent micDone, "", "Screenshot has been taken and attached to ALM test run. To view it, double-click on test in Test Lab's Execution Grid, click on Runs in left-hand menu of Test Instance Details window, then click on Attachments paperclip on left-hand side of test run row. Screenshot also stored on execution pc at following location: '" & fullScreenshotPathAndFileName & "'. "

End Function


'@Description Get the current date and time
'@Documentation Get the current date and time and format it in yymmddhhmmss order.
Public function ChorusUtilities_GetCurrentDateTimestamp ()
	Dim timestampArray
	Dim timestampString, timestampTime, timestampDate
	
	timestampArray = split (Now, ":")
	timestampString = join (timestampArray, "")
	'above results in "24/06/2009 160855"
	timestampArray = split (timestampString, " ")
	timestampTime = timestampArray(1)
	timestampDate = timestampArray(0)
	timestampArray = split (timestampDate, "/")
	timestampString = right(timestampArray(2), 2) & timestampArray(1) & timestampArray(0) & timestampTime
	'above results in yymmddhhmmss, e.g. "090624160855".
'	Reporter.ReportEvent micDone, "ChorusUtilities_GetCurrentDateTimestamp", "Timestamp = {" & timestampString & "}"
	ChorusUtilities_GetCurrentDateTimestamp = timestampString
	
End Function

'@Description Logout and close the browser
'@Documentation Logout and close the browser
Function SSP_Logout()
Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
SSP_Home_Desc.Link("innerhtml:=Logout").Click
SSP_Clear_Security_Certificate
	If SSP_Home_Desc.WebElement("innerhtml:=You have been successfully logged out from the Chorus Self Service Portal.").Exist(5) = True Then	
			Reporter.ReportEvent micPass, "Logout" , "Success"
	Else
			ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Logout" ,"Failure"
			
	End If
'Close the Browser
	Browser("name:=Chorus Self Service Portal").Close
End Function



' The following function should not be selectable within BPT Components. Therefore there is no @Description or @Documentation.
' Function Description:
' 				This function attaches a file to the Attachments tab of the current test run within ALM.
'				It's used to attach screenshots in the event of fatal errors in order to assist test failure analysis.
Function ChorusUtilities_AttachFileToALMTestRun(FileName)
	Dim currentTestRun, testRunAttachmentsFactory, theAttachment
	
	If QCUtil.IsConnected Then
		' Get the current test run object
		Set currentTestRun = QCUtil.CurrentRun
		
		If currentTestRun is Nothing Then
			Reporter.ReportEvent micDone,"ChorusUtilities_AttachFileToALMTestRun","Test is not running from an ALM Test Set in the Test Lab (it's likely to be running from the ALM Test Plan module), so not attaching file '" & FileName & "' to test run."
		Else
			' Get the Attachments Factory of the current test run
			Set testRunAttachmentsFactory = currentTestRun.Attachments
			
			' Attach the file to the current test run.  Attachment will display in Test Set > Execution Details > Test Instance Details > Runs > Run Details > Attachments
			Set theAttachment = testRunAttachmentsFactory.AddItem(Null)
			theAttachment.FileName = FileName
			theAttachment.Type = 1
			theAttachment.Post
			theAttachment.Refresh
		End If
	Else
		Reporter.ReportEvent micDone,"ChorusUtilities_AttachFileToALMTestRun","UFT is not connected to ALM, so not attaching file '" & FileName & "' to test run."
	End If

End Function



'@Description Chorus User needs to select the RSP
'@Documentation Chorus User selected the RSP
Function Chorus_Select_RSP (RSP)

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebList("html id:=customer").Exist(5) = True Then

		SSP_Home_Desc.WebList("html id:=customer").Select RSP	
' Keeping the condition as there is some issue with the logout feature, WIll remove the IF condition at the end.
End If

Chorus_Select_RSP = RSP

End Function



'@Description Enter the Physical address (Street No, Street Name , Town) and Search
'@Documentation Enter the Physical address (Street No, Street Name , Town) and Search
Function SSP_Location_Search_StreetNo_StreetName_Town(StreetNo,StreetName,Town)

Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

SSP_Home_Desc.WebEdit("html id:=streetNumQck").Set StreetNo
SSP_Home_Desc.WebEdit("html id:=streetNameQck").Set StreetName
SSP_Home_Desc.WebEdit("html id:=townQck").Set Town

SSP_Home_Desc.WebButton("html id:=quickSearchButton").Click

wait(1) ' to sync up if the search is taking up time

If SSP_Home_Desc.Webtable("Class:=ssp-data-table keyboard-table").Exist(0) = True Then
	Reporter.ReportEvent micPass, "Location Search", "Quick Search functionality is working"
else
	Reporter.ReportEvent  micFail, "Location Search", "Quick Code Search functionality is not working or Location details are wrong"
	
End If

End Function

'@Description Search location with Street Name
'@Documentation Search location with Street Name
Function SSP_Location_Search_StreetName(StreetName)

Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

SSP_Home_Desc.WebEdit("html id:=streetNameQck").Set StreetName

SSP_Home_Desc.WebButton("html id:=quickSearchButton").Click

wait(1) ' to sync up if the search is taking up time

If SSP_Home_Desc.Webtable("Class:=ssp-data-table keyboard-table").Exist(0) = True Then
	Reporter.ReportEvent micPass, "Location Search", "Quick Search functionality is working"
else
	Reporter.ReportEvent  micFail, "Location Search", "Quick Code Search functionality is not working or Location details are wrong"
	
End If

End Function

'@Description Select Product Offer
'@Documentation Select Product Offer
Function SSP_Select_Product_Offer(prodoffer)

Dim oDesc, Linkobjs,i,x,ButtObjs
set oDesc = Description.Create
oDesc("micclass").value = "WebElement"
oDesc("class").value = "checkbox-label productOfferingCheckboxLabel"

'Find all the Product Offers
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)
'obj.Count value has the number of Webelements in the page
For i = 0 to Linkobjs.Count - 1 
	'get the name of all the links in the page 
	x = Linkobjs(i).GetROProperty("innerhtml")
	
	If Instr(1,x,prodoffer) <> False Then
		Linkobjs(i).click
		Exit for
	End If
Next
SSP_Select_Product_Offer = prodoffer
wait(0.5)


' check if the correct product has been selected. Whenever we select the corect button , The "NEXT" Webbutton gets enabled.

set oDesc = Description.Create
oDesc("micclass").value = "WebButton"
oDesc("name").value = "Next"
oDesc("class").value = "button ssp-action-btn"

Set ButtObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)

For i = 0 To ButtObjs.count -1 
	If ButtObjs(i).GetROProperty("disabled") = 0 Then
		Reporter.ReportEvent micPass,"Product Offer got Selected","Product offer got selected properly"
		ButtObjs(i).Click
		Exit for
		else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Product Offer was wrong","Product offer wasnt selected properly"
		
		Exit for
	End If 
Next

End Function

'@Description Select Aim
'@Documentation Select Aim
Function SSP_SelectAim(Passedaim)
	
Dim AllConnectoptions ' This will check if the Aim is present in the dropdown or not
AllConnectoptions = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=aimselect","html tag:=SELECT").GetROProperty("all items")

If instr(1,AllConnectoptions,Passedaim) <> 0 Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=aimselect","html tag:=SELECT").select(PassedAim)
	Reporter.ReportEvent micPass ,"AimSelected" ," Connection Aim was selected"
Else
     ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Wrong aim Selected"," Connection Aim was not selected"
     
End If

SSP_SelectAim = Passedaim
End Function

'@Description Select CSE
'@Documentation Select CSE
Function SSP_SelectCSE(PassedCSE)
	
Dim AllCSEoptions ' This will check if the CSE is present in the dropdown or not
AllCSEoptions = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=cseOffer","html tag:=SELECT").GetROProperty("all items")

If instr(1,AllCSEoptions,PassedCSE) <> 0 Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=cseOffer","html tag:=SELECT").select(PassedCSE)
	Reporter.ReportEvent micPass ,"CSESelected" ," Passed CSE was selected"

ElseIf (PassedCSE = "") Then
	
	Reporter.ReportEvent micPass ,"CSE Not  Passed" ," CSE Was not passed.. hence not selected"

Else
     Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=cseOffer","html tag:=SELECT").select(1)
     Reporter.ReportEvent micwarning ,"Passed CSE was not Present"," Passed CSE was not selected So keeping the default CSE"
     
End If

End Function

'@Description Check Order Type
'@Documentation Check Order Type
Function SSP_Check_Order_Type()

Dim Order_Type
Order_Type = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("html id:=orderTypeSpan").GetROProperty("innertext")

Reporter.ReportEvent micDone,Order_Type,"Done"

End Function

'@Description Capture Order
'@Documentation Capture Order
Function SSP_Capture_Button()
	
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webbutton("html id:=captureOrderButton").Click
	
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
	
End Function

'@Description Check Feasibility by validaing Reserve Button is enabled
'@Documentation Check Feasibility by validaing reserve button is enabled
Function SSP_Check_Feasibility_Message()
'Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
'Dim Temp
'If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert alert-success").Exist(1) Then
'	Temp = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert alert-success").GetROProperty("innerhtml")
'	
'Else 
'	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Order Capture","Order Not Captured Successfully"
'	
'End If
'
'
'If  (instr(1,Temp,"Feasible",1) > 0) Then
'	Reporter.ReportEvent micpass ,"Feasibility","Feasible to Provide Service at Location"
'	
'Else
'	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Feasibility","NOT Feasible to Provide Service at Location"
'	
'End If

DIM oDescWebButton,i,isDisabled
isDisabled = 1
wait(2)

If  Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*").WebButton("outertext:=Reserve").Exist(2) Then

	Set oDescWebButton = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*").WebButton("outertext:=Reserve")
	isDisabled = oDescWebButton.GetROProperty("disabled")
	
	If isDisabled = 0  Then
		Reporter.ReportEvent micPass,"Order Capture","Order Captured Successfully"		
	else
		ChorusUtilities_TakeScreenshot
		Reporter.ReportEvent micFail,"Order Capture","Order Captured but either ISL issue or Manual Investigation is required"
	
	End If
Else
	ChorusUtilities_TakeScreenshot
	Reporter.ReportEvent micFail,"Order Capture","Order Captured Manual Investigation may be  required"
End If

End Function


'@Description Check Feasibility Message Manual Investigation if Fibre Status is not ready
'@Documentation Check Feasibility Message Manual Investigation if Fibre Status is not ready
Function SSP_Check_Feasibility_Message_Manualinvestigation()
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim Temp
Temp = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert warning").GetROProperty("innerhtml")

If  Instr(Temp,"Feasibility request requires manual investigation") <> 0 Then
	Reporter.ReportEvent micpass ,"Feasibility","Request Manual investigation"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Feasibility","Wrong set of TLC has been used ot the system is not working properly"
	
End If

End Function

'@Description Get Order ID from Order Summary page or any other Order tab
'@Documentation Get Order ID from Order Summary page or any other Order tab
Function SSP_Get_Order_ID()
Dim OrderID
OrderID = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebElement("xpath:=//label[text()='Order ID:']/following-sibling::span").GetROProperty("outertext")
'msgbox OrderID
SSP_Get_Order_ID = OrderID
End Function

'@Description Get Location ID from Order Summary page
'@Documentation Get Location ID from Order Summary page
Function SSP_Get_Location_ID()
Dim LocationID
LocationID = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebElement("xpath:=//label[text()='Location ID:']/..").GetROProperty("outertext")
'msgbox LocationID
SSP_Get_Location_ID = LocationID
End Function



'@Description Check Feasibility Message if Fibre Status is not ready
'@Documentation Check Feasibility Message if Fibre Status is not ready
Function SSP_Check_Feasibility_Message_ManualFeasibility()
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim Temp
Temp = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert warning").GetROProperty("innerhtml")

If  Instr(Temp,"Request Manual Feasibility") <> 0 Then
	Reporter.ReportEvent micpass ,"Feasibility","Request Manual Feasibility"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Feasibility","Wrong set of TLC has been used ot the system is not working properly"
	
End If

End Function

'@Description Check Feasibility Message if SF Fails
'@Documentation  Check Feasibility Message if SF Fails
Function SSP_Check_Feasibility_Message_SF_Fails()

Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim Temp
Temp = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert warning").GetROProperty("innerhtml")

If  Instr(Temp,"Feasibility request requires manual investigation.") <> 0 Then
	Reporter.ReportEvent micpass ,"Feasibility","Feasibility request requires manual investigation."
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Feasibility","Wrong set of TLC has been used ot the system is not working properly"
	
End If

End Function


'@Description Click on Validate Location
'@Documentation  Click on Validate Location
Function SSP_Click_Validate_Location
'Author Swetha
'Validate_Location_Submit_Address_Request_NBAP
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebButton("outertext:=Validate Location").Exist(0) = True Then
'Click on Validate Location
A.WebButton("outerhtml:=<button class=""button"" id=""updateLocationButton"" type=""button"" value=""Submit"">        <span>Validate Location</span>       </button>").Click
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Validate Location","Validate Location button not Exist"
    
End If  

End Function


'@Description Enter Validate Location  Details
'@Documentation Enter Validate Location  Details
Function SSP_Validate_Req_Enter_Customer_Details (Name,Email)

Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebElement("html id:=createUpdateLocationPopup","html tag:=DIV").Exist(0) = True Then
    'Enter End Customer Name
    A.WebEdit("type:=text","name:=endCustomerName","html tag:=INPUT").Set Name
    'Enter Notification Email
    A.WebEdit("type:=text","name:=notificationEmail","html tag:=INPUT").Set Email
        
else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Location Validation", "Page doesnt exist"
    
End If

End Function

'@Description Select NBAP Request Type
'@Documentation Select NBAP Request Type
Function SSP_Validate_Req_Select_NBAP_Type
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebElement("outertext:=Request Type  Address  NBAP ").Exist(0) = True Then
   A.WebRadioGroup("name:=locationRequestType").Select "NBAP"
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"NBAP Type","NBAP not exist"
    
End If
End Function

'@Description Enter NBAP Details
'@Documentation Enter NBAP Details
Function SSP_Validate_Req_NBAP_Location_Details (HouseNo,Coordinate)
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

If A.WebElement("outertext:=NBAP Location Information").Exist(3) = True Then
    A.WebRadioGroup("name:=nbapDetails\.isNewBuilding","value:=true").Click
    wait(3)
    A.WebRadioGroup("name:=nbapDetails\.isNumberGivenByCouncil").Click
    Wait(1)
    A.WebEdit("type:=text","name:=nbapDetails.councilNumber").Set HouseNo
    Wait(2)
    A.WebEdit("type:=text","name:=nbapDetails.coordinates").Set Coordinate

'    A.WebRadioGroup("name:=nbapDetails.isNewBuilding").Select "false" 
'    Reporter.ReportEvent micDone,"No","Selected No"
    
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Address Location Information","Address Location Info not exist"
    
End If
End Function

'@Description Click on send Button
'@Documentation Click on send Button
Function SSP_Validate_Req_NBAP_Click_Send
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

If A.WebButton("outertext:=Send").Exist(0) = True Then
    A.WebButton("outertext:=Send").Click
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Send","Send does Not exist"
    
End If
    
End Function

'@Description Capture Request ID
'@Documentation Capture Request ID
Function SSP_Validate_Get_RequestID
Dim Temp,Order_ID
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebElement("class:=alert alert-success").Exist = True Then
  'Get the Request ID
    Order_ID = Replace(Replace(Trim(A.WebElement("class:=alert alert-success").GetROProperty("innertext")),"Thank you for submitting location request # ",""),".","")
 
    SSP_Validate_Get_RequestID =  Order_ID 
  Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Location Request #","Location Request Message not exist"
    
End If
End Function

'@Description Select Address Request type
'@Documentation Select Address Request type
Function SSP_Validate_Select_Address_Type
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

IF A.WebElement("html id:=addressDiv","html tag:=DIV").Exist(0)= True Then
	A.WebRadioGroup("outerhtml:=<input name=""locationRequestType"" class=""with-label"" id=""locationRequestType1"" type=""radio"" value=""ADDRESS"" for-label=""lblLocationRequestTypeAddress"">").Click
    Reporter.ReportEvent micDone,"Address Location Information","Address Location Information opens"
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Address Location Information","Address Location Inforamtion does not exist"
    
End If
End Function

'@Description Enter Property type and Attach a map
'@Documentation Enter Property type and Attach a map
Function SSP_Validate_Attach_Map_And_PropertyType (NewLocationAddress,CustomerLetterBox,BuildingLocation,AddressTenancyAgreement,NearestNeighboursAddress,ASID,FilePath)
'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

A.WebEdit("name:=addressDetails.newLocationAddress").Set NewLocationAddress
A.WebEdit("name:=addressDetails.letterboxOrBusinessSign").Set CustomerLetterBox
A.WebRadioGroup("name:=addressDetails.isNewBuilding").Select "true"
A.WebRadioGroup("name:=addressDetails.isMultipleBuildings").Select "true"
A.WebEdit("name:=addressDetails.customersProperty").Set BuildingLocation
If A.WebElement("outertext:=Property Type:\*  Tenanted  Owner Occupied ").Exist = True Then
	A.WebRadioGroup("name:=addressDetails\.propertyType","value:=TENANTED").Click 
 	A.WebEdit("name:=addressDetails\.tenancyAgreementAddress","type:=textarea").Set AddressTenancyAgreement
 Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Property Type","Property Type Not Found"
	
End If
A.WebEdit("name:=addressDetails.nearestNeighboursAddress").Set NearestNeighboursAddress
A.WebEdit("name:=addressDetails.previousPhoneNumberOrASID").Set ASID
If A.WebElement("outertext:=AttachmentsOptional: Attach a map with the location marked or other supporting document  Browse Selected File Path:   ").Exist = True Then

     A.WebFile("name:=documentData").Set FilePath
Wait(2)
   
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Attachments","Attachments does not found"
	
End If
End Function

'@Description Wait for passed Minutes
'@Documentation  Wait for Passed Minutes
Function SSP_wait_mins(timeinMis)
	'this function just waits for the given amount of seconds.
	Dim timeinsecs
	timeinsecs = timeinMis * 60
	wait(timeinsecs)
End Function

'@Description Assign the task to other user
'@Documentation  Assign the task to other user
Function SSP_Validate_Location_Request_Creation_And_Assignment(ReqID,Taskassinee)
'Author Swetha , RJ 
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

If A.Link("outertext:=Work Queue").Exist = True Then
	A.Link("outerhtml:=<a href=""/chorus-ssp-web/pages/WorkQueue/WorkQueue"" s-key=""w"" s-key-type=""1"">Work Queue</a>").Click    
	A.WebList("outertext:=Complex ProvisioningConsent and BuildException ManagementLocation RequestsNGA Provisioning ").Select "Location Requests" 
Dim oDesc, Linkobjs,i,x,y,Start_string,End_string,zz
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all of the links of Request
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","class:=task-results-table has-rollover ssp-data-table keyboard-table").ChildObjects(oDesc)
'obj.Count value has the number of Webelements in the page 

'msgbox Linkobjs.Count
For i = 0 to Linkobjs.Count - 1 
	'get the name of all the links in the table
	x = trim(Linkobjs(i).GetROProperty("outertext"))
	
	If Instr(1,x,ReqID) <> False Then
		y = Linkobjs(i).GetROProperty("url")
		Start_string = Instr(1,y,"htmTaskId=") + 10
		End_string = Instr(1,y,"redirectOnExit")
		zz =  mid(y,Start_string,End_string-start_string-1)
		' Now we need to find the corresponding Checkbox
		Exit for
	End If
Next

Dim oiDesc, chkobjs,j
set oiDesc = Description.Create
oiDesc("micclass").value = "WebCheckBox"
Set chkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","class:=task-results-table has-rollover ssp-data-table keyboard-table").ChildObjects(oiDesc)

					'Find all of the webCheckboxes
							For j = 0 To chkobjs.count - 1
					 				y = trim(chkobjs(j).getROProperty("Value"))
					 				
							 				If Instr(y,zz) <> 0 Then

							 				chkobjs(j).Click
							 				Exit For
					 			End If
					 	Next
	
	A.WebButton("outerhtml:=<button name=""assignToOther"" class=""button assign-tasks-action"" id=""assignToOther"" type=""button""><span>Assign To Other</span></button>").Click   
	A.WebList("html id:=popupAssignToUserId","html tag:=SELECT","name:=popupAssignToUserId").Select Taskassinee
	A.WebButton("outerhtml:=<button class=""button continue float-right ssp-action-btn"" id=""popupAssignButton"" type=""button""><span>Assign</span></button>").Click   
	
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Work Queue","Work Queue does not exist"
	
End If

End Function

'@Description Select Order Type
'@Documentation  Select Order type
Function SSP_Select_Order_Type()
Dim oDesc,ButtObjs , i
set oDesc = Description.Create
oDesc("micclass").value = "WebButton"
oDesc("name").value = "Select Order Type"
oDesc("class").value = "button ssp-action-btn"
wait(2)
Set ButtObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)

For i = 0 To ButtObjs.count -1 
	If ButtObjs(i).GetROProperty("disabled") = 0 Then
		Reporter.ReportEvent micPass,"Location 2 is valid","Location 2 is valid"
		ButtObjs(i).Click
		Exit for
		else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Location 2 is invalid","Location 2 is invalid"
		
		Exit for
	End If 
Next

End Function

'@Description Select Aim for Layer 1 
'@Documentation Select Aim for Layer 1 
Function SSP_SelectAim_Layer1(Passedaim)
	
Dim AllConnectoptions ' This will check if the Aim is present in the dropdown or not
AllConnectoptions = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=aimSelectLayer1","html tag:=SELECT").GetROProperty("all items")

If instr(1,AllConnectoptions,Passedaim) <> 0 Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=aimSelectLayer1","html tag:=SELECT").select(PassedAim)
	Reporter.ReportEvent micPass ,"AimSelected" ," Connection Aim was selected"
Else
     ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Wrong aim Selected"," Connection Aim was not selected"
     
End If

End Function


'@Description Enter Site Information Details 
'@Documentation Enter Site Information Details 
Function SSP_Enter_Site_Information_Details(Input_Name,Input_Phone_No)
	' THis will enter for all Contact Name 
Dim oDescName,Contact_Names,i,Prov_Ref
set oDescName = Description.Create
oDescName("micclass").value = "WebEdit"
oDescName("class").value = "input-large contact-name alphabet-only"

Set Contact_Names = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDescName)

For i = 0 To Contact_Names.count -1 
	Contact_Names(i).Set Input_Name &" "&i+1
Next

' This will enter for all Preferred Phone and Alternate Phone
Dim oDescNo,Contact_Phone,j
set oDescNo = Description.Create
oDescNo("micclass").value = "WebEdit"
oDescNo("class").value = "input-large"

Set Contact_Phone = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDescNo)

For j = 0 To Contact_Phone.count -1 
	Contact_Phone(j).Set Input_Phone_No &j
Next

' This will enter for all of the other details 
Dim oDescOtherDetails,k,Other_Details
set oDescOtherDetails = Description.Create
oDescOtherDetails("micclass").value = "WebEdit"
oDescOtherDetails("class").value = "input-xlarge"

Set Other_Details = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDescOtherDetails)

For k = 0 To Other_Details.count -1 
	Other_Details(k).Set Input_Name &" Notes:" & Other_Details(k).GetROProperty("InnerText") 
Next

'ENd Customer Name
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webedit("html id:=customerDetails.endCustomerName").Set Input_Name
Prov_Ref = Input_Name & Replace(date,"/","") & right("0" & hour(time),2) & right("0" & minute(time),2)
SSP_Enter_Site_Information_Details = Prov_Ref
'Provider Reference
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webedit("html id:=customerDetails.customerOrderNumber").set Prov_Ref

End Function

'@Description Submit Feasibility
'@Documentation Submit Feasibility
Function SSP_Submit_Feasibility()
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("Innerhtml:=Submit Feasibility").Click
End Function

'@Description Note down the Order No( Feasibility ID ) For Layer 1
'@Documentation Note down the Order No( Feasibility ID ) For Layer 1
Function SSP_GET_Feasibility_ORDER_ID()' returns the order id
Dim OrderID, Alert_Text

Alert_Text = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").webelement("class:=alert alert-success").GetROproperty("innertext")

OrderID = Mid(Alert_Text,Instr(Alert_Text,"#")+1,9)
SSP_GET_Feasibility_ORDER_ID = OrderID
End Function

'@Description Navigate to Work Queue
'@Documentation  Navigate to Work Queue
Function SSP_Nav_Work_Queue()
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("outertext:=Work Queue").Exist = True Then
		A.Link("outerhtml:=<a href=""/chorus-ssp-web/pages/WorkQueue/WorkQueue"" s-key=""w"" s-key-type=""1"">Work Queue</a>").Click   
	Else
		Reporter.ReportEvent micWarning, "WOrk queue doesnt Exist"
	End iF
End Function


'@Description Navigate In the order to Tasks
'@Documentation  Navigate In the order to Tasks
Function SSP_Order_Nav_Order_Task()
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("name:=Tasks ").Exist = True Then
		A.Link("name:=Tasks ").Click   
		A.Sync
	Else
		Reporter.ReportEvent micWarning, "Tasks tab", "Tasks tab could not be found"
	End iF
End Function

'@Description Navigate In the order to Order Details
'@Documentation  Navigate In the order to Order Details
Function SSP_Order_Nav_Order_Order_Details()
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("name:= Order Details ").Exist = True Then
		A.Link("name:= Order Details ").Click   
	Else
		Reporter.ReportEvent micWarning, "Order Details Page Doesnt Exist"
	End iF
End Function

'@Description Navigate In the order to  Schedule 
'@Documentation  Navigate In the order to  Schedule 
Function SSP_Order_Nav_Order_Schedule()
' This will navigate to the schedule page
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("name:= Schedule ").Exist = True Then
		A.Link("name:= Schedule ").Click   
	Else
		Reporter.ReportEvent micWarning, " Schedule  Page Doesnt Exist"
	End iF
End Function

'@Description Order Summary Sub-Status
'@Documentation Order Summary  Sub-Status 
Function SSP_Nav_Summary_Verify_Order_Sub_Status(Sub_status)

' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
DIM A
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Sub-status: "&Sub_status&" "  

if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Order Summary", "Order Summary Sub-status updated"
else
	Reporter.ReportEvent  micFail, "Order Summary", "Order Summary is not Sub-status updated"
	
End If

End Function

'@Description Complete the Feasibility Task
'@Documentation Complete the Feasibility Task
Function SSP_Task_Perform_Feasibility(Task_Outcome)
	'If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Perform Layer One Feasibility").Exist(0) = True Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Perform Layer One Feasibility").Click
    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").weblist("all items:=Feasible;Not Feasible").select Task_Outcome
    
    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").weblist("html id:=notFeasibleReason").Select(1)
    
    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webbutton("html id:=completeButton").Click
    wait(2)
    
    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webbutton("html id:=complete-popupContinueButton").Click
    
    
'else
	'ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, " Perform Layer One Feasibility"," Perform Layer One Feasibility Task not present" 
'End If

End Function

'@Description Navigate In the order to the Order Summary tab
'@Documentation  Navigate In the order to the Order Summary tab
Function SSP_Order_Nav_Order_Summary()
' This will navigate to the Summary page
	DIM A
	
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("name:=Summary ").Exist = True Then
		A.Link("name:=Summary ").Click   
	Else
		Reporter.ReportEvent micWarning,"Order Summary", " Summary Page Doesnt Exist"
	End iF
End Function

'@Description In Appointment Page, Go the the next Month
'@Documentation In Appointment Page, Go the the next Month
Function SSP_Appointment_Go_To_Next_Month()
' To go To next Month
				DIM SSP_Home_Desc
				Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
				Dim Go_to_Next_Month
				Set Go_to_Next_Month = SSP_Home_Desc.WebElement("class:=ui-datepicker-next ui-corner-all")
				
				Go_to_Next_Month.Click
End Function


'@Description In Appointment Page, Select First Appointment Date
'@Documentation In Appointment Page, Select First Appointment Date
Function SSP_Appointment_Select_First_Available_Date()
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*Appointment.*")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
End Function



'@Description In Appointment Page, Select Select Scope Date
'@Documentation In Appointment Page, Select Scope Date
' needs to be updated
Function SSP_Appointment_Select_Scope_Date()
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
End Function

'@Description In Appointment Page, Select Install Date
'@Documentation In Appointment Page, Select Install Date
' needs to be updated
Function SSP_Appointment_Select_Install_Date()
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
End Function



'@Description In Appointment Page, Select PM Slot
'@Documentation In Appointment Page, Select PM Slot
' needs to be updated
Function SSP_Appointment_Select_PM_SLOT()
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
End Function



'@Description In Appointment Page, Select AM Slot
'@Documentation In Appointment Page, Select AM Slot
Function SSP_Appointment_Select_AM_SLOT()
' needs to be updated
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
End Function



'@Description In Appointment Page, Reserve Appointment
'@Documentation In Appointment Page,  Reserve Appointment
Function SSP_Appointment_Reserve_Appointment()
'DIM SSP_Home_Desc,i
'        
'    Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*")
'    SSP_Home_Desc.WebButton("outertext:=Reserve").Click
'    
'    wait(2)
'    Browser("name:=Chorus Self Service Portal").Sync
DIM oDescWebButton,i,isDisabled
isDisabled = 1
wait(2)

If  Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*").WebButton("outertext:=Reserve").Exist(2) Then

	Set oDescWebButton = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*").WebButton("outertext:=Reserve")
	isDisabled = oDescWebButton.GetROProperty("disabled")
	
	If isDisabled = 0  Then
		Reporter.ReportEvent micPass,"Order Capture","Order Captured Successfully"	
		oDescWebButton.Click
		Browser("name:=Chorus Self Service Portal").Sync
	else
		ChorusUtilities_TakeScreenshot
		Reporter.ReportEvent micFail,"Order Capture","Order Captured but either ISL issue or Manual Investigation is required"
	
	End If
Else
	ChorusUtilities_TakeScreenshot
	Reporter.ReportEvent micFail,"Order Capture","Order Captured Manual Investigation may be  required"
End If

        
End Function




'@Description In Offer Characteristics  Page, Select Default Values
'@Documentation In Offer Characteristics  Page, Select Default Values
Function SSP_Offer_Characteristics_Set_Default_Values()
On Error Resume Next ' the properties of Webedits and Weblists are wrong, so have set up all webedits, and igrnored the errors

        
        DIM SSP_Home_Desc,i,i2,testvar
        Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

' Checking if we on current Page
If SSP_Home_Desc.Link("innertext:= Offer Characteristics ").Exist(0) = True Then
        SSP_Home_Desc.Link("innertext:= Offer Characteristics ").Click
    ' Get the list of all Webedits
            Dim oDesc,oDesc2,WebeditObjs,WeblistObjs
            
            set oDesc = Description.Create
            oDesc("micclass").value = "WebEdit"
            
            Set WebeditObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)
            
            For i = 0 To WebeditObjs.count -1  ' to edit webedits

testVar = WebeditObjs(i).getRoProperty ("readonly")

if testVar = 0 then
			If WebeditObjs(i).getRoProperty ("default value") = "" Then
				WebeditObjs(i).Set "12"
			End If


end if

                    
                
            Next
            
            set oDesc2 = Description.Create
            oDesc2("micclass").value = "WebList"
            
            Set WeblistObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc2)
            For i2 = 0 To WeblistObjs.count -1  ' to set defult values for list
            
testVar = WeblistObjs(i2).getRoProperty ("selection")

'msgbox testVar
if instr(testVar,"#0") <> 0 then
WeblistObjs(i2).Select 1
'msgbox WeblistObjs(i2).getRoProperty ("html id")
wait(0.25)

end if
  
            Next
            
          Reporter.ReportEvent micpass ,"Offer Char" ," Offer char. was updated"   
            
else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Offer Char. Page was not displayed" ," Check appointment Page" 
  ' 
    
End If
SSP_Home_Desc.WebButton("name:=Save").Click
    
SSP_Home_Desc.sync

End Function

'@Description Check if there is Scope Calender.
'@Documentation  Check if there is Scope Calender.
Function SSP_Order_Check_Scope_Calender()




End Function

'@Description Select Product Instance
'@Documentation  Select Product Instance
Function SSP_Select_Product_Instance()

Dim RC,A,i,CC,PrimaryInstance,Active,PID
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
If A.WebTable("html id:=searchResultTable").Exist(1) = True Then
    RC = A.WebTable("html id:=searchResultTable").RowCount
    
For i = 1 To RC Step 1
	CC = A.WebTable("html id:=searchResultTable").ColumnCount(i)
	
	If CC > 1 Then
		
		PrimaryInstance =trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,2))
		Active = trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,8))
	
		
		If (PrimaryInstance = "Primary") And (Active = "Active") Then
			
			A.WebTable("name:=selectedProductInstanceAtLocationIndex").ChildItem(i,1,"WebCheckBox",0).Set "ON"
			wait(1)
			PID = trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,3))
			Reporter.ReportEvent micPass,"Active Product Instance","Active Product Instance Selected"
			Exit For
			
		End If
	End If
    
    
Next
Else

    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micFail,"Active Product Instance","Active Product Instance does not exist"
    
End If

SSP_Select_Product_Instance = PID
End Function

'@Description SSP_Select_product_Instance_Secondary
'@Documentation SSP_Select_product_Instance_Secondary
Function SSP_Select_product_Instance_Secondary()

Dim RC,A,i,CC,SecondaryInstance,Active,PID
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
If A.WebTable("html id:=searchResultTable").Exist(1) = True Then
    RC = A.WebTable("html id:=searchResultTable").RowCount
    
For i = 1 To RC Step 1
	CC = A.WebTable("html id:=searchResultTable").ColumnCount(i)
	
	If CC > 1 Then
		
		SecondaryInstance =trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,2))
		Active = trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,8))
	
		
		If (SecondaryInstance = "Secondary") And (Active = "Active") Then
			
			A.WebTable("name:=selectedProductInstanceAtLocationIndex").ChildItem(i,1,"WebCheckBox",0).Set "ON"
			wait(1)
			PID = trim(A.WebTable("name:=selectedProductInstanceAtLocationIndex").GetCellData(i,3))
			Reporter.ReportEvent micPass,"Active Product Instance","Active Product Instance Selected"
			Exit For
			
		End If
	End If
    
    
Next
Else

    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micFail,"Active Product Instance","Active Product Instance does not exist"
    
End If

SSP_Select_product_Instance_Secondary = PID

End Function

'@Description Click on Next Button
'@Documentation  Click on Next Button
Function SSP_Click_on_Next_Button_to_Get_Aim()
Dim oDesc,ButtObjs,i
set oDesc = Description.Create
oDesc("micclass").value = "WebButton"
oDesc("name").value = "Next"
oDesc("class").value = "button ssp-action-btn"

Set ButtObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)

For i = 0 To ButtObjs.count -1 
	If ButtObjs(i).GetROProperty("disabled") = 0 Then
		Reporter.ReportEvent micPass,"Next Button is enabled","Next Button is enabled"
		ButtObjs(i).Click
		Exit for
		else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Next Button is disabled","Next Button is disabled"
		
		Exit for
	End If 
Next
	
End Function



'@Description Navigate In the order to Offer Characteristics 
'@Documentation  Navigate In the order to Offer Characteristics 
Function SSP_Order_Nav_Offer_Characteristics()
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	If A.Link("name:= Offer Characteristics ").Exist = True Then
		A.Link("name:= Offer Characteristics ").Click   
	Else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, " Offer Characteristics  Doesnt Exist"
		
	End iF
End Function




'@Description Check Error Messages for appointment not reserved
'@Documentation  Check Error Messages for appointment not reserved
Function SSP_Error_message_Appointment_Not_Reserved()
	DIM A
	Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
	wait(5)
	
	If A.Webelement("innertext:=Your order cannot be submitted due to some errors. Please check highlighted tabs. ").Exist(0) = True Then
		REporter.ReportEvent micPass,"Appointment was expired","Appointment was expired"
	
	Else
	
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Appointment was not expired","Appointment was not expired"
		
	End If

End Function



'@Description Submit Order
'@Documentation Submit Order
Function SSP_Submit_Order()
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("Innertext:=Submit Order").Click
End Function

'@Description Check the Min SLA required for the order type
'@Documentation Check the Min SLA required for the order type

Function SSP_Appointment_check_SLA(min_no_days)

DIM SSP_Home_Desc,i,month_count, Iterator, First_Available_Date,Available_Month,Available_Year,Booked_date , ODesc_List_Of_Dates,List_Of_Dates
month_count = 0
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
        Set ODesc_List_Of_Dates = Description.Create()
        ODesc_List_Of_Dates("micclass").Value = "Link"
        
        Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)    

For Iterator = 0 To 3 Step 1 ' We will check for three months
    
    If List_Of_Dates.Count = 0  Then
        SSP_Home_Desc.WebElement("Innertext:=Next","class:=ui-datepicker-next ui-corner-all").Click
        month_count = month_count + 1
        Set ODesc_List_Of_Dates = Description.Create()
        ODesc_List_Of_Dates("micclass").Value = "Link"
        Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)    
        
    Else     
    
                For i = 0 To List_Of_Dates.Count() - 1 Step 1
                        
                        First_Available_Date = Cint(List_Of_Dates(i).GETROProperty("innerhtml"))
                        
                        List_Of_Dates(i).Click
                        
                        Exit For    
                Next    
        Exit for
        
    End If    
    
Next

If Iterator = 4 Then
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Appointment not available","Appointment not available"
    
Else
    ' Check first available day was after how many days

Available_Month = SSP_Home_Desc.WebElement("class:=ui-datepicker-month").GetROProperty("innerhtml")
Available_Year = SSP_Home_Desc.WebElement("class:=ui-datepicker-year").GetROProperty("innerhtml")


Booked_date = cdate(DateValue(First_Available_Date &"-" & Available_Month &"-" & Available_Year))

    If (Datediff("d",date,Booked_date) > min_no_days)  Then
            reporter.ReportEvent micPass,"SLA ","SLA Passed"
            
            
    else
            ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"SLA ","SLA Failed"
            
           
        End If

End If    

	
End Function


'@Description Appointment Page should not allow RSP to book an Appointment  
'@Documentation Appointment Page should not allow RSP to book an Appointment
Function SSP_ManageConsent_Verify_AppointmentPage
	'Appointment page
	'Author Swetha
Dim A
'Bookable Appointments
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebTable("column names:=Bookable Appointments").exist(1) = True then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Schedule page","Appointment Page allow RSP to book an Appointment"
	
Else
	Reporter.ReportEvent micpass,"Schedule page","Appointment Page does not allow RSP to book an Appointment"
End if
End Function

'@Description In the consent queue, Manage Consent task is created  
'@Documentation In the consent queue, Manage Consent task is created
Function SSP_ManageConsent_Nav_Task_ManageConsentTaskCreated
	'Author Swetha
	'Consent Queue
SSP_Order_Nav_Order_Task()
Dim OrderIDCreated
OrderIDCreated = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("outertext:=Order ID:1........ ").GetROProperty("outertext"),10))
'Print OrderIDCreated
Reporter.ReportEvent micDone,OrderIDCreated,"Order ID created"

SSP_Select_Order_Task("Manage Consent")

Dim oDesc,Linkobjs,i,OrderIDPresent
set oDesc = Description.Create
oDesc("micclass").value = "Link"
Dim childobjs
'Find all the links
Set childobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)

For i = 0 to childobjs.Count-1
    'get the name of all the links in the page 
    OrderIDPresent = Trim(Right(childobjs(i).GetROProperty("innerhtml"),10))
    'print OrderIDPresent
    If Instr(1,OrderIDPresent,OrderIDCreated) <> False Then
       Reporter.ReportEvent micDone,"Order Id","Order ID present in table"
'    else
'    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Order Id","Order ID not Assoicated with Existed Order"
	Exit for
 	End If 
   
Next

'SSP_ManageConsent_Nav_Task_ManageConsentTaskCreated = OrderIDCreated
End Function

'@Description Click on the Manage Consent Task, set the Consent Status and click Complete
'@Documentation Click on the Manage Consent Task, set the Consent Status and click Complete
Function SSP_ManageConsent_ManualTask_Complete(Consentstatus,NextStep,SendNotification)

SSP_Select_Order_Task("Manage Consent")

Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")	
	A.WebList("name:=consentAndBuild\.consentBuildRequestSubState").Select Consentstatus
	A.WebList("name:=data\.consentOutcome").Select NextStep
	A.WebList("name:=data\.consentSendNotification").Select SendNotification
	A.WebButton("outertext:=Complete").Click
'Confirm Page
If A.WebElement("html id:=complete-dialog-confirm").Exist(1)=True Then
	A.WebButton("name:=Continue").Click
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Confirm Dialog","Confirmation dialog does not exist"
	
End IF
End Function


'@Description Click on the Manage Consent Task, set the Consent Status and click Save
'@Documentation Click on the Manage Consent Task, set the Consent Status and click Save
Function SSP_ManageConsent_ManualTask_Save(ConsentStatus,SendNotification)

'Capture the current Order ID before navigating to the Consent Task
Dim OrderID
'OrderID = SSP_Get_Order_ID()
OrderID = SSP_Get_Order_ID()
SSP_Select_Order_Task("Manage Consent")

'Update the Consent Status and Save
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")	
	A.Sync
	A.WebList("name:=consentAndBuild\.consentBuildRequestSubState").Select ConsentStatus
	A.WebList("name:=data\.consentSendNotification").Select SendNotification
	A.WebEdit("name:=data\.notificationText").Set("Automation")
	A.WebButton("outertext:=Save").Click

'Find the link to navigate back to the Order
Dim oDesc,childobjs,i,OrderIDPresent
set oDesc = Description.Create
oDesc("micclass").value = "Link"
'Find all the links in the Supplied Data ELements table
Set childobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)
For i = 0 to childobjs.Count-1
    'get the name of all the links in the table and find the one matching our Order ID
    OrderIDPresent = Trim(Right(childobjs(i).GetROProperty("innerhtml"),10))
    If Instr(1,OrderIDPresent,OrderID) <> False Then
       childobjs(i).Click
       A.Sync
	Exit for
 	End If    
Next
	
End Function

'@Description A STATUS Notification is Sent to the RSP Indicating Consent has been Gained  
'@Documentation A STATUS Notification is Sent to the RSP Indicating Consent has been Gained
Function SSP_ManageConsent_ManualTask_StatusNotificationSenttoRSP
	'Author Swetha
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.Link("name:=History ").Exist(0) = True Then
wait(1)
A.Link("name:=History ").Click
End If
Dim RC,i
RC = A.WebTable("class:=expandable-table").RowCount
    'MsgBox RC
For i = 0 To RC Step 1
	If Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=expandable-table").WebElement("outerhtml:=<td>         <i class=""toggle-icon  icon-chorus-down-circle""></i>         Consent Gained        </td>").exist(1) = true Then
	Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=expandable-table").WebElement("outerhtml:=<td>         <i class=""toggle-icon  icon-chorus-down-circle""></i>         Consent Gained        </td>").WebElement("outerhtml:=<i class=""toggle-icon  icon-chorus-down-circle""></i>").Click
	Reporter.ReportEvent micDone,"Consent Gained","A STATUS Notification is Sent to the RSP Indicating Consent has been Gained"
	Wait(1)
	' ()
Exit For
End IF
A.Link("outertext:=Next").Click
Next

End Function

'@Description Navigate to the Tasks tab and click the Manage MDU Build task, then complete the task
'@Documentation Navigate to the Tasks tab and click the Manage MDU Build task, then complete the task
Function SSP_MDUBuild_ManualTask_Complete(BuildStatus,NextStep,SendNotification)

Dim A : Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

'Go to Tasks tab and click on the Manage MDU Build task
SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage MDU Build")

'Required Data
A.WebList("name:=consentAndBuild\.consentBuildRequestSubState").Select BuildStatus
A.WebList("name:=data\.consentOutcome").Select NextStep
A.WebList("name:=data\.consentSendNotification").Select SendNotification
If A.WebButton("outertext:=Complete").Exist(1)=True Then
	A.WebButton("outertext:=Complete").Click
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Complete","Complete button not exist"
	
End If

'Confirmation Page

If A.WebElement("html id:=complete-dialog-confirm").Exist(1)=True Then
	A.WebButton("name:=Continue").Click
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Confirm Dialog","Confirmation dialog not found"
	
End If	
	
End Function

'@Description Navigate to the Tasks tab and click the Manage MDU Design task, then complete the task
'@Documentation Navigate to the Tasks tab and click the Manage MDU Design task, then complete the task
Function SSP_MDUDesign_ManualTask_Complete(DesignStatus,NextStep,SendNotification)

Dim A : Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

'Go to Tasks tab and click on the Manage MDU Design task
SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage MDU Design")

'Required Fields
A.WebList("name:=consentAndBuild\.consentBuildRequestSubState").Select DesignStatus
A.WebList("name:=data\.consentOutcome").Select NextStep
A.WebList("name:=data\.consentSendNotification").Select SendNotification
A.WebButton("name:=Complete").Click
'Confirm

If A.WebElement("html id:=complete-dialog-confirm").Exist(1)=True Then
	A.WebButton("name:=Continue").Click
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Confirm Dialog","Confirmation dialog not found"
	
End IF
End Function

'@Description Complete the Order by following the Manage MDU Design Process 
'@Documentation Complete the Order by following the Manage MDU Design Process
Function SSP_Appointment_Check_Appointment_Types(Expected_Scope,Expected_Install,Expected_CSE,Expected_RFS)
		
		DIM SSP_Home_Desc,i,iFlag
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*Appointment.*")
		Dim ODesc_List_Of_Appts,List_Of_Appts,IsScopePresent,IsInstallPresent,IsCSEPresent,ISRFSPresent,Chk_Appt_Type
		
		iFlag = 0
		IsScopePresent = "no"
		IsInstallPresent = "no"
		IsCSEPresent = "no"
		ISRFSPresent = "yes"
		
		
		Set ODesc_List_Of_Appts = Description.Create()
		ODesc_List_Of_Appts("micclass").Value = "WebElement"
		'ODesc_List_Of_Appts("class").Value = "ui-datepicker-inline ui-datepicker ui-widget ui-widget-content ui-helper-clearfix ui-corner-all"
		ODesc_List_Of_Appts("html tag").Value = "H3"
		
		Set List_Of_Appts = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Appts)	
'msgbox List_Of_Appts.Count		
		For i = 0 To List_Of_Appts.Count() - 1 Step 1
		
			Chk_Appt_Type = List_Of_Appts(i).GETROProperty("innertext")
			'msgbox Chk_Appt_Type
			'msgbox InStr("CSE", Chk_Appt_Type)
			
			If InStr(Chk_Appt_Type,"CSE") <> 0 Then
			
				IsCSEPresent = "yes"
				ISRFSPresent = "no"
				
			End If
			'msgbox InStr("Install", Chk_Appt_Type)
			
			If InStr(Chk_Appt_Type,"Install") <> 0 Then
				IsInstallPresent = "yes"
				ISRFSPresent = "no"
				
				
			End If
			'msgbox InStr("Scope", Chk_Appt_Type)
			
			If InStr(Chk_Appt_Type,"Scope") <> 0 Then
				IsScopePresent = "yes"
				ISRFSPresent = "no"
				
				
			End If

		Next	
		
		
'		If SSP_Home_Desc.Webelement("innerhtml:=No site visits are required for this order.").exist(0) = True Then
'			ISRFSPresent = "yes"
'		End If

If (IsScopePresent = "yes" And IsInstallPresent = "yes") Then
	
	SSP_Validate_Default_Scope_Install_Dates()
	
End If

If (IsInstallPresent = "no" And IsCSEPresent = "yes") Then
	
	'SSP_Validate_Default_CSE_Date()
	
End If

If (ISRFSPresent = "yes") Then
	
'	SSP_Validate_Default_RFS_Date()
	
End If




If 	(IsScopePresent = Expected_Scope and IsInstallPresent = Expected_Install and IsCSEPresent = Expected_CSE and ISRFSPresent = Expected_RFS) Then
	
	Reporter.ReportEvent micPass ,"Appointment type","Appointment type present as per requirements"
else
	ChorusUtilities_TakeScreenshot
	Reporter.ReportEvent micFail ,"Appointment type","Appointment type is not correct as per requirements or check the test data and parameters"
	
End If

SSP_Appointment_Check_Appointment_Types =  "Scope"&IsScopePresent&"Install"&IsInstallPresent&"CSE"&IsCSEPresent
	
End Function


'@Description In the Consent Queue Existing Consent Task is associated with this order
'@Documentation In the Consent Queue Existing Consent Task is associated with this order
Function SSP_ManageConsent_ConsentQueue_ConsentTaskIsAssociated
'Author Swetha
SSP_Order_Nav_Order_Task()
Dim OrderIDCreated
OrderIDCreated = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("outertext:=Order ID:1........ ").GetROProperty("outertext"),10))
'Print OrderIDCreated
SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage Consent")

'Verifying for existed order
Dim rowcount,ExistedOrder
rowcount = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").RowCount

ExistedOrder = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").ChildItem(rowcount,0,"webElement",2).GetRoProperty("innertext"),10))
'Print ExistedOrder
Wait(1)
Dim oDesc,Linkobjs,i,AssociatedOrder
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the links
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)

For i = 1 to Linkobjs.Count-1
    'get the name of all the links in the page 
    AssociatedOrder = Trim(Right(Linkobjs(i).GetROProperty("innerhtml"),10))
    'print AssociatedOrder

    If Instr(1,AssociatedOrder,ExistedOrder) <> False Then
        Reporter.ReportEvent micDone,"Order Id","Order ID Assoicated with Existed Order"
    else
    	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Order Id","Order ID not Assoicated with Existed Order"
    	
    Exit for
    End If
Next

End Function

'@Description Order Submitted Successful message
'@Documentation Order Submitted Successfully
Function SSP_Check_OrderSubmittedSuccessfully
'Author Swetha
DIM A
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

Dim oDesc ,ListObjs,temp , i 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("micclass").value= "WebElement"
oDesc("class").value= "alert alert-success"

Set ListObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)

'msgbox ListObjs.count
For i = 0 To ListObjs.count - 1
		temp =  ListObjs(i).GetROProperty("innertext")
	If instr(temp,"Thank you for submitting order") <> 0 Then
		Reporter.ReportEvent micPass,"Order","Order submitted successfully"
	Exit for
	End If
Exit for

If i = ListObjs.count Then
			ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Order","Order was not submitted successfully"
			
End If

Next 
End Function

'@Description Order Status
'@Documentation Order Status 
Function SSP_Nav_Summary_Verify_OrderStatus(Status)
'Author swetha
' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
DIM A
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
Dim oDesc 'Declare an object variable
A.Link("outertext:=Summary ").Click
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Status: "&Status&" "  

if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist(2) = True then
	Reporter.ReportEvent micPass, "Location Search", "order status searched"
else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Location Search", "order status not searched"
	
End If

End Function

'@Description Click on Cancel button
'@Documentation Click on Cancel button
Function SSP_click_on_Cancel_Button()
'Author Swetha
	Dim A
	Set A = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
	A.WebButton("outerhtml:=<button name=""cancel"" class=""button"" id=""cancelButton"" type=""button""><span>Cancel</span></button>").Click
End Function

'@Description Check that in appointment Page, Install should be in Pending
'@Documentation Check that in appointment Page for Non Standard Install, Install should be in Pending
Function SSP_Appointment_Non_Standard_Install_Pending_Install()
wait(10)
		If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Pending AppointmentsInstall appointment pending").Exist(0) = True Then
			
			Reporter.ReportEvent micPass ,"Appointment type","Install appointment pending for consent scenario . Appointment type present as per requirements"
		else
			ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Appointment type","Appointment type is not correct as per requirements or check the test data and parameters"
			
		End If
End Function

'@Description In the Consent Queue, Disassociate the Order from the Consent Task by selecting 'Consent' option
'@Documentation In the Consent Queue, Disassociate the Order from the Consent Task by selecting 'Consent' option
Function SSP_ManageConsent_ManualTask_DisassociateOrder(Action)
	'Author Swetha
DIM A
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")	
Dim oDesc,Linkobjs,i,TaskCreated
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the links
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)
For i = 0 to Linkobjs.Count-1 
    'get the name of all the links in the page 
    
    TaskCreated = Linkobjs(i).GetROProperty("Y")
    'print TaskCreated
Exit for
Next    
Dim oDesc1,Linkobjs1,i1,Disassociate,AO
set oDesc1 = Description.Create
oDesc1("micclass").value = "WebButton"

'Find all the links
Set Linkobjs1 = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc1)
For i = 0 to Linkobjs1.Count-1 
    'get the name of all the links in the page 
    
     AO = Linkobjs1(i).GetROProperty("Y")
     Disassociate = AO+3
    'print Disassociate
 If Instr(1,Disassociate,TaskCreated) <> False Then
     Linkobjs1(i).Click   
   Reporter.ReportEvent micDone,"Disassociate Order","Select disassociate"
 End if      
Exit For
Next 
'Dialogue box opens
If Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("html id:=dialog-disassociate-order").Exist(1) = True Then
        'Dim Action 
        'Action = "Revisit Consenting"
    	Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebList("html id:=actionSelectDropDown").Select Action
        Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebButton("class:=popupConfirmButton button float-right").Click
        'reporter.ReportEvent micPass,ConsentTaskCreated,"Order disassociated from the Consent Task"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Disassociate Order Doalogue","Disassociate Order Doalogue not exist"
	
End If
wait(1)
'A.Webbutton("outerhtml:=<button name=""cancel"" class=""button"" id=""cancelButton"" type=""button""><span>Cancel</span></button>").click

End Function


'@Description Book First appointment For Intact Cases
'@Documentation Book First appointment for Intact cases
Function SSP_Appointment_Book_First_Date_Intact_Connect()
		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")
		SSP_Home_Desc.WebEdit("class:=hasDatepicker").Click
		wait(2)
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","class:=ui-datepicker-calendar")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 0 To List_Of_Dates.Count() - 1 Step 1
		
			First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next	
If i  = List_Of_Dates.Count() Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"RFS Date  not selected","RFS Date not Selected"
Else
	Reporter.ReportEvent micPass,"RFS Date selected","RFS Date Selected"
	
End If		
End Function


'@Description User Initiates Order Cancellation
'@Documentation User Initiates Order Cancellation
Function SSP_Task_Initiate_Cancel_Order()
SSP_Order_Nav_Order_Summary()
DIM SSP_Home_Desc,i
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
wait(5)
If SSP_Home_Desc.WebButton("html id:=cancelOrder").Exist(0) = True Then
	SSP_Home_Desc.WebButton("html id:=cancelOrder").Click
	
	SSP_Home_Desc.WebList("html id:=cancelReason").Select(3)
	SSP_Home_Desc.WebButton("html id:=popupCancelSaveButton").Click
	SSP_Home_Desc.Sync
	wait(10)
	SSP_Order_Nav_Order_Summary()
	
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Cancel Was not Initiated", " Cancel was not initiated "	
	
End If
	
End Function

'@Description Chorus User Approves Cancel Order
'@Documentation Chorus User Approves Cancel Order
Function SSP_Task_Chorus_Approve_Cancel_Order()
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
SSP_Order_Nav_Order_Summary()
If SSP_Home_Desc.WebButton("html id:=cancelOrderApproval").Exist(0) = True Then
    SSP_Home_Desc.WebButton("html id:=cancelOrderApproval").Click

    
    SSP_Home_Desc.WebList("innertext:= Accept Reject ").Select(1)
    SSP_Home_Desc.Webbutton("html id:=popupCancelOrderApprovalSaveButton").Click
    SSP_Home_Desc.Sync
    wait(3)
    SSP_Order_Nav_Order_Summary()
    wait(40) ' Time required to task to come there
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Cancel was not Initiated" ,"Order Cancellation was not initiated"
    
End If


End Function

'@Description Chorus User Completes Order Cancellation
'@Documentation Chorus User Completes Order Cancellation
Function SSP_Chorus_Complete_Order_Cancellation()

DIM SSP_Home_Desc,i
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Complete Order Cancellation")
SSP_Home_Desc.WebButton("html id:=completeButton","name:=Complete").Click								
SSP_Home_Desc.WebButton("class:=button continue float-right","html id:=complete-popupContinueButton").Click																
Reporter.ReportEvent micdone , "Chorus action completed the order cancellation ", "Chorus action completed the order cancellation"

SSP_Order_Nav_Order_Summary()
If SSP_Home_Desc.Webelement("innertext:=Sub-status: Cancelled ").Exist = True Then		
	Reporter.ReportEvent micPass , "Order Cancelled" ,"Order Cancelled" 									
else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail , "Order was not Cancelled" ,"Order was not Cancelled" 									
End if

End Function


'@Description check Appointment page after cancelling the order
'@Documentation check Appointment page after cancelling the order
Function SSP_Appointment_Check_After_cancellation()

SSP_Order_Nav_Order_Schedule()
wait(2)

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Column names:=Schedule appointments")

if SSP_Home_Desc.Webelement("innerhtml:=No appointment is required.").exist(0) = True  then
		Reporter.ReportEvent micPass, "Appointment", "Appointment not available after cancelling the order"		
else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Appointment", "Appointment available even after cancelling the order"
End if


End Function

'@Documentation navigate to Administartion
'@Description navigate to Administartion
Function SSP_Prod_Inv_Administartion()

Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.link("outertext:=Administration ").Click	
End Function


'@Documentation navigate to Product Inventory 
'@Description navigate to Product Inventory
Function SSP_Prod_Inv_Nav()
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.link("name:=Product Inventory").click
End Function

'@Documentation Select the Create Record
'@Description Select the Create Record
Function SSP_Prod_Inv_Nav_Create_New_Record()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.webbutton("name:=Create Record").Click
End Function

'@Documentation Select the Service Provider
'@Description Select the Service Provider
Function SSP_Prod_Inv_Nav_Create_New_Record_Service_Provider(service_Provider)
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.weblist("name:=customerReference").Select service_Provider
SSP_Prod_Inv_Nav_Create_New_Record_Service_Provider = service_Provider
End Function

'@Documentation Select the Location ID
'@Description Select the Location ID
Function SSP_Prod_Inv_Nav_Create_New_Record_Loc_ID(Loc_ID)
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
	SSP_Home_Desc.webedit("name:=locationReference").set Loc_ID
End Function

'@Documentation Select the product ID
'@Description Select the product ID
Function SSP_Prod_Inv_Nav_Create_New_Record_Prod_ID(Prod_ID)
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.webedit("name:=externalReference").Set Prod_ID
SSP_Prod_Inv_Nav_Create_New_Record_Prod_ID=Prod_ID
End Function

'@Documentation Select the product
'@Description Select the product
Function SSP_Prod_Inv_Nav_Create_New_Record_Product(Product)

Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
'SSP_Home_Desc.WebList("name:=productOfferingId").Select Product
Dim odesc,All_Weblists,i

Set Odesc = Description.Create
Odesc("micclass").value = "WebList"
set All_Weblists = SSP_Home_Desc.Childobjects(Odesc)
For i =0 to All_weblists.count - 1
'msgbox all_weblists(i).GETROPROPERTY("all items")
	If instr(all_weblists(i).GETROPROPERTY("all items"),Product) <> 0 Then
		all_weblists(i).select Product
		Exit for
End if
Next
End Function

'@Documentation Verify the Location Reference
'@Description Verify the Location Reference
Function SSP_Prod_Inv_Create_Record_Loc_Reference()
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
SSP_Home_Desc.WebEdit("name:=locationReferenceB").Set ""
End Function

'@Documentation Update the Record
'@Description Update the record 
Function SSP_Prod_Inv_Update_Record()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
	SSP_Home_Desc.WebButton("name:=Update").Click
End Function

'@Documentation Verify the Product Status 
'@Description Verify the Product Status
Function SSP_Prod_Inv_Product_Status()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
	SSP_Home_Desc.WebRadioGroup("name:=showActiveOnly").Select "false"
End Function

'@Description Create New Record Offer Characteristics  Page, Select Default Values
'@Documentation Create New Record  Offer Characteristics  Page, Select Default Values
Function SSP_Prod_Inv_Offer_Characteristics_Set_Default_Values()
On Error Resume Next ' the properties of Webedits and Weblists are wrong, so have set up all webedits, and igrnored the errors

        
        DIM SSP_Home_Desc,i,i2,testvar,testvar2
        Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
        
        
        
' Checking if we on current Page

If SSP_Home_Desc.WebElement("innertext:= Offer Characteristics ").Exist(0) = True Then
        SSP_Home_Desc.WebElement("innertext:= Offer Characteristics ").Click
    ' Get the list of all Webedits
            Dim oDesc,oDesc2,WebeditObjs,WeblistObjs
            
            set oDesc = Description.Create
            oDesc("micclass").value = "WebEdit"
            
            Set WebeditObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)
            
            For i = 0 To WebeditObjs.count -1  ' to edit webedits

testVar = WebeditObjs(i).getRoProperty ("readonly")
testvar2 = WebeditObjs(i).getRoProperty ("disabled")
if testVar = 0 then
'msgbox WebeditObjs(i).getRoProperty ("html id")


           If testvar2 <> 1 Then
                 WebeditObjs(i).Set "13"	
           End If


end if

                    
                
            Next
            
            set oDesc2 = Description.Create
            oDesc2("micclass").value = "WebList"
            
            Set WeblistObjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc2)
            For i2 = 0 To WeblistObjs.count -1  ' to set defult values for list
            
testVar = WeblistObjs(i2).getRoProperty ("selection")

'msgbox testVar
if instr(testVar,"#0") <> 0 then
WeblistObjs(i2).Select 1
'msgbox WeblistObjs(i2).getRoProperty ("html id")
wait(0.25)

end if
 
            Next
            
          Reporter.ReportEvent micpass ,"Offer Char" ," Offer char. was updated"   
            
else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Offer Char. Page was not displayed" ," Check appointment Page" 
  ' 
    
End If
End Function

'@Description Submit the product Inventory order 
'@Documentation Submit the product Inventory order 
Function SSP_Prod_Inv_Submit_Record()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
	SSP_Home_Desc.webbutton("name:=Submit").Click
	wait(1)
End Function

'@Description Search the Record 
'@Documentation  Search the Record 
Function SSP_Prod_Inv_Search_Record(service_Provider_2,Prod_ID_2)
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
'If SSP_Home_Desc.WebElement("innertext:=The Product ID you entered is not unique, please verify the ID and try again\. ").Exist(1)=true Then
	'ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Prod_ID already exist", "Please provide Unique Prod_ID"
'End If

'SSP_Home_Desc.webbutton("outertext:=Search").Click
SSP_Home_Desc.weblist("name:=customerReference").Select service_Provider_2
SSP_Home_Desc.webedit("name:=externalReference").Set Prod_ID_2
'If SSP_Home_Desc.WebButton("name:=Search").Exist(0)=true Then

	SSP_Home_Desc.WebButton("name:=Search").Click
	
	'reporter.ReportEvent micpass, "Search button","Search button is enabled and selected"
	'else
	'ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Search Button","Search Button is disabled"
'End If

Dim oDescc,Linkobjs,i,x
set oDescc = Description.Create
oDescc("micclass").value = "Link"

'Find all the links
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("html id:=searchResultTable").ChildObjects(oDescc)
'obj.Count value has the number of Webelements in the page 

For i = 0 to Linkobjs.Count - 1 
	'get the name of all the links in the page 
	x = Linkobjs(i).GetROProperty("innerhtml")
	
	If Instr(1,x,Prod_ID_2) <> False Then
		Linkobjs(i).click
		Exit for
	End If
Next
wait(1)
End Function

'@Description Display the Product Charateristics
'@Documentation Display the Product Charateristics
Function SSP_Prod_Inv_Display_Product_Characteristics()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
	If SSP_Home_Desc.webelement("outertext:=Offer Characteristics").Exist(1) = true Then
	
wait(1)

Reporter.ReportEvent micpass, "Offer characteristics", "SSP Record Offer characteristics displayed" 
	    
 else
    
ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Offer characteristics", "SSP Record Offer characteristics not displayed" 
End If
End Function

'@Descrption Verify the Teminate Date is blank
'@Documentation Verify the Teminate Date is blank
Function SSP_Prod_Inv_Query_Terminate_Date_Blank()
	Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
if SSP_Home_Desc.webbutton("outertext:=Terminate").Exist(0)=true then

reporter.ReportEvent micpass, "Terminate button","Terminate date is blank"
else
ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Terminate button not found","Terminate date is not blank"
End if
End Function

'@Description Terminate the prod_ID
'@Documentation Terminate the prod_ID
Function SSP_Prod_Inv_Terminate_Product_Instance()
Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
If SSP_Home_Desc.webbutton("name:=Terminate").Exist(1)=true  Then
	SSP_Home_Desc.webbutton("name:=Terminate").Click
	Else 
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Product ID is terminated already", "Product ID is terminated already"
End If
SSP_Home_Desc.webbutton("name:=Yes").Click
End Function

''@Description Verify Product ID has been Terminated Successfully 
''@Documentation Verify Product ID has been Terminated successfully 
'Function SSP_Prod_Inv_Termination_Changes_Implemented(prod_ID_2)
'Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
'Dim ResultDesc,Webobjs,k,z
'set ResultDesc = Description.Create
'ResultDesc("micclass").value = "Webelement"
''prod_ID =1621215885
'
''Find all the Webelements
'Set Webobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(ResultDesc)
''obj.Count value has the number of Webelements in the page 
'
'For k = 0 to Webobjs.Count - 1 
'	'get the name of all the links in the page 
'	z = Webobjs(k).GetROProperty("outerhtml")
'	
'	If Instr(1,z,prod_ID_2) <> False Then
'		Reporter.ReportEvent micPass,"Termination  is successful", "Termination  is successful"
'		Exit for
'	End If
'Next
'
'
'If k > Webobjs.Count Then
'	
'	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Termination  is not successful", "Termination  is not successful.Please refer to the screenshots"
'	
'End If
'
'wait(0.5)
'End Function


'@Description Appointment, Select a date other than first available date, Take 2nd Day
'@Documentation Appointment, Select a date other than first available date, Take 2nd Day
Function SSP_Appointment_Update_Date_Other_than_First_date() 
DIM SSP_Home_Desc,i,month_count, Iterator, First_Available_Date,Available_Month,Available_Year,Booked_date , ODesc_List_Of_Dates,List_Of_Dates
month_count = 0
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
      
      Set ODesc_List_Of_Dates = Description.Create()
        ODesc_List_Of_Dates("micclass").Value = "Link"
        ODesc_List_Of_Dates("class").Value ="ui-state-default"
        
        Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)    

For Iterator = 0 To 3 Step 1 ' We will check for three months
    
    If List_Of_Dates.Count < 1  Then
        SSP_Home_Desc.WebElement("Innertext:=Next","class:=ui-datepicker-next ui-corner-all").Click
        month_count = month_count + 1
        Set ODesc_List_Of_Dates = Description.Create()
        ODesc_List_Of_Dates("micclass").Value = "Link"
        Set List_Of_Dates = SSP_Home_Desc.WebTable("name:=24").ChildObjects(ODesc_List_Of_Dates)    
        
    Else     
    
                For i = 1 To List_Of_Dates.Count() - 1 Step 1
                        
                        First_Available_Date = Cint(List_Of_Dates(i).GETROProperty("innerhtml"))
                        
                        List_Of_Dates(i).Click
                        
                        Exit For    
                Next    
        Exit for
        
    End If    
    
Next 
End Function



'@Description Book appointment Intact Cases other than first availabe date
'@Documentation Book appointment Intact Cases other than first availabe date
Function SSP_Appointment_Book_Other_Than_First_Date_Intact_Connect()

		DIM SSP_Home_Desc,i
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Schedule appointments")
		SSP_Home_Desc.WebEdit("class:=hasDatepicker").Click
		wait(2)
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","class:=ui-datepicker-calendar")
		Dim First_Available_Date
		Dim ODesc_List_Of_Dates,List_Of_Dates
		Set ODesc_List_Of_Dates = Description.Create()
		ODesc_List_Of_Dates("micclass").Value = "Link"
		
		Set List_Of_Dates = SSP_Home_Desc.ChildObjects(ODesc_List_Of_Dates)			
		For i = 1 To List_Of_Dates.Count() - 1 Step 1
		
		First_Available_Date = List_Of_Dates(i).GETROProperty("innerhtml")
			List_Of_Dates(i).Click
			Exit For
		Next		
		
End Function


'@Description In Appointment Page, Reschedule Appointment
'@Documentation In Appointment Page,  Reschedule Appointment
Function SSP_Appointment_Reschedule_Appointment()
		DIM SSP_Home_Desc
		Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=.*appointment.*")
		SSP_Home_Desc.WebButton("outertext:=Reschedule").Click
		SSP_Select_RescheduleReason
End Function


'@Description Verify the History tab to confirm the notification sent
'@Documentation Verify the History tab to confirm the notification sent
Function SSP_Order_History_tab()

Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Link("name:= Order Details ").Click
Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Link("name:=History").Click

if Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Link("innerhtml:=Expand All").Exist(0)=true then
Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Link("innerhtml:=Expand All").Click
reporter.ReportEvent micPass, "Order History", "Order Status notification sent successfully"
else
ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Order History", "Order Status notification not sent successfully"
End if

End Function

'@Description Order Summary Sub-Status
'@Documentation Order Summary  Sub-Status 
Function SSP_Nav_Summary_Verify_Order_Sub_Status(Sub_status)

' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Sub-status: "&Sub_status&" "  

if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist(2) = True then
	Reporter.ReportEvent micPass, "Order Summary", "Order Summary Sub-status updated"
else
	Reporter.ReportEvent  micDone, "Order Summary", "Order Summary is not Sub-status updated"
	
End If

End Function

'@Description Verify the Manual Task for Manage Design Build 
'@Documentation Verify the Manual Task for Manage Design Build 
Function SSP_ManualTask_Manage_Complex_Design(Next_Step)
Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
Browser("name:=Chorus Self Service Portal").Sync

SSP_Select_Order_Task("Manage Complex Design")
wait(0.5)
SSP_Home_Desc.WebList("name:=data\.nextStep").select Next_Step
	
End Function

'@Description Verify the Manual Task for Manage Complex Build 
'@Documentation Verify the Manual Task for Manage Complex Build 
Function SSP_ManualTask_Manage_Complex_Build(Next_Step)

SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage Complex Build")
	
Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebList("name:=data\.nextStep").select Next_Step	
End Function

'@Description Verify the Complete Button
'@Documentation Verify the Complete Button
Function SSP_ManualTask_Complete_Button()

Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If SSP_Home_Desc.WebButton("name:=Complete").Exist(0)=true Then
	SSP_Home_Desc.WebButton("name:=Complete").Click
End If

End Function

'@Description Verify the Continue Button
'@Documentation Verify the Continue Button
Function SSP_ManualTask_Continue_Button()

Set SSP_Home_Desc=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
wait(0.5)
If SSP_Home_Desc.WebButton("name:=Continue").Exist(0)=true Then
	SSP_Home_Desc.WebButton("name:=Continue").Click
End If

End Function

'@Description Verify the Exception Task has been created
'@Documentation Verify the Exception Task has been created
Function SSP_Manage_Design_Build_Exception_Task()

Dim SSP_Home_Desc
SSP_Select_Order_Task("Build Exception")

End Function

'@Description Cancel button
'@Documentation Cancel button
Function SSP_Manage_Design_Build_Cancel()

	set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
		SSP_Home_Desc.WebButton("html id:=cancelButton","outertext:=Cancel").Click
	
End Function

'@Description Order Summary Status
'@Documentation Order Summary Status
Function SSP_Order_Nav_Summary_Status()

set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
wait(1)
Browser("name:=Chorus Self Service Portal").Sync
SSP_Home_Desc.Link("outertext:=Tasks ").Click
SSP_Home_Desc.Link("name:=Summary ").Click
Wait(2)
if SSP_Home_Desc.WebElement("innertext:=Status: In Progress ").Exist(0)=true then
reporter.ReportEvent micPass, "Order Status", "Order Status in Progress"
else
ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Order Status", "Order Status is not in Progress"
End if 

End Function

'@Description Order Summary Sub-Status
'@Documentation Order Summary  Sub-Status 
Function SSP_Nav_Summary_Verify_Order_Sub_Status(Sub_status)

' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Sub-status: "&Sub_status&" "  

if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Order Summary", "Order Summary Sub-status updated"
else
	Reporter.ReportEvent  micdone, "Order Summary is either Schedule scoping or Schedule Install", "Order Summary is not Sub-status updated"
	
End If

End Function

'@Description Select Offer type  
'@Documentation Select Offer type secondary 
Function SSP_ManageConsent_Select_OfferType(Offer_Type)
'Author swetha
'Dim Offer_Type
'Offer_Type = "Secondary" without CSE
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebList("outertext:=All Not Applicable Primary Secondary ").Exist(1) = True Then
	A.WebList("outertext:=All Not Applicable Primary Secondary ").Select Offer_Type
	Reporter.ReportEvent micPass,"Offer Type","Offer Type selected valid"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Offer Type","Offer Type selected unknown"
	
End If
End Function

'@Description SSP Offer CHar page, Get RFS Date
'@Documentation  SSP Offer CHar page, Get RFS Date
Function SSP_Offer_Char_Get_RFS_Date()

        DIM SSP_Home_Desc,i,Odesc,RFS_Date_Value,RFS
        Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

        Set Odesc = Description.Create()
        ODesc("micclass").value = "WebElement"
        ODesc("Outertext").Value = "RFS Date:.*"
        ODesc("html tag").Value = "P"
        
        
        set RFS_Date_Value = SSP_Home_Desc.ChildObjects(Odesc)
        
        'msgbox RFS_Date_Value.COunt

        For i = 0 To RFS_Date_Value.COunt - 1 Step 1
            RFS = Left(Right(RFS_Date_Value(i).GETROPROPERTy("outertext"),9),8)
        Next
        SSP_Offer_Char_Get_RFS_Date = RFS
        
End Function


'@Description Select Offer Type  
'@Documentation Select Offer Type
Function SSP_Select_OfferType(Offer_Type)
'Author swetha
'Dim Offer_Type
'Offer_Type = "Secondary" without CSE
Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
If A.WebList("outertext:=All Not Applicable Primary Secondary ").Exist(1) = True Then
	A.WebList("outertext:=All Not Applicable Primary Secondary ").Select Offer_Type
	Reporter.ReportEvent micPass,"Offer Type","Offer Type selected valid"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Offer Type","Offer Type selected unknown"
	
End If
End Function

'@Description  Manual  MDU Design Task 
'@Documentation  Manual  MDU Design Task 
Function SSP_ManualTask_MDU_Design(Design_status,Next_Step)

SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage MDU Design")

Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Weblist("html id:=consentResult","name:=consentAndBuild\.consentBuildRequestSubState").Select Design_status

Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebList("name:=data\.consentOutcome").select Next_Step
	
End Function

'@Description  Manual  MDU Build Task 
'@Documentation  Manual  MDU Build Task
Function SSP_ManualTask_MDU_Build(Design_status,Next_Step)

SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage MDU Build")
Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Weblist("name:=consentAndBuild\.consentBuildRequestSubState").Select Design_status
Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebList("name:=data\.consentOutcome").select Next_Step
	
End Function


'@Description Order Summary Status
'@Documentation Order Summary  Status 
Function SSP_Summary_Verify_Order_Status(status)

' Writing Descriptive Programming as the location details and Netwrok details would be dynamic in nature
set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
wait(0.4)
SSP_Home_Desc.Link("outertext:=Tasks ").Click

Browser("name:=Chorus Self Service Portal").Sync
wait(0.5)
SSP_Home_Desc.Link("name:=Summary ").Click
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
Dim oDesc 'Declare an object variable
Set oDesc = Description.Create 'Create an empty description
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "status: "&Status&" "  

if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist = True then
	Reporter.ReportEvent micPass, "Order Summary", "Order Summary status updated"
else
	Reporter.ReportEvent  micFail, "Order Summary", "Order Summary is not status updated"
	
End If

End Function


'@Description Check Design and Build is required
'@Documentation Check Design and Build is required
Function SSP_Check_Design_Build_Message()
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
'Dim Temp
'Temp = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement("class:=alert warning").GetROProperty("innerhtml")

If  Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebElement("outerhtml:=<div class=""alert warning"">  Design and Build is required if Customer Order is submitted for this location\. </div>").Exist(0)=true Then
	Reporter.ReportEvent micpass ,"Design and Build is required ","Design and Build is required if Customer Order is submitted for this location"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Feasibility","Wrong set of TLC has been used ot the system is not working properly"
	
End If

End Function

'@Description Manage Design and Build OrderID verification
'@Documentation Manage Design and Build OrderID verification
Function SSP_Manage_Design_Build_OrderID_TaskIsAssociated

SSP_Order_Nav_Order_Task()
Dim OrderIDCreated
OrderIDCreated = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("outertext:=Order ID:1........ ").GetROProperty("outertext"),10))
SSP_Order_Nav_Order_Task()
SSP_Select_Order_Task("Manage MDU Design")

'Verifying for existed order
Dim rowcount,ExistedOrder
rowcount = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").RowCount

ExistedOrder = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").ChildItem(rowcount,0,"webElement",2).GetRoProperty("innertext"),10))
'Print ExistedOrder
Wait(1)
Dim oDesc,Linkobjs,i,AssociatedOrder
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the links
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)

For i = 0 to Linkobjs.Count-1
    'get the name of all the links in the page 
    AssociatedOrder = Trim(Right(Linkobjs(i).GetROProperty("innerhtml"),10))
    print AssociatedOrder

    If Instr(1,AssociatedOrder,ExistedOrder) <> False Then
        Reporter.ReportEvent micDone,"Order Id","Order ID Assoicated with Existed Order"
    else
    	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Order Id","Order ID not Assoicated with Existed Order"
    	
    Exit for
    End If
Next

End Function

'@Description MDU Build Task OrderID verification
'@Documentation MDU Build Task OrderID verification
Function SSP_MDU_Build_Task_OrderID_TaskIsAssociated

SSP_Order_Nav_Order_Task()
Dim OrderIDCreated
OrderIDCreated = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("outertext:=Order ID:1........ ").GetROProperty("outertext"),10))
'Print OrderIDCreated
SSP_Select_Order_Task("Manage MDU Build")

'Verifying for existed order
Dim rowcount,ExistedOrder
rowcount = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").RowCount

ExistedOrder = Trim(Right(Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").webtable("html id:=searchResultTable").ChildItem(rowcount,0,"webElement",2).GetRoProperty("innertext"),10))
'Print ExistedOrder
Wait(1)
Dim oDesc,Linkobjs,i,AssociatedOrder
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the links
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebTable("class:=ssp-data-table").ChildObjects(oDesc)

For i = 0 to Linkobjs.Count-1
    'get the name of all the links in the page 
    AssociatedOrder = Trim(Right(Linkobjs(i).GetROProperty("innerhtml"),10))
    print AssociatedOrder

    If Instr(1,AssociatedOrder,ExistedOrder) <> False Then
        Reporter.ReportEvent micDone,"Order Id","Order ID Assoicated with Existed Order"
    else
    	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Order Id","Order ID not Assoicated with Existed Order"
    	
    Exit for
    End If
Next

End Function

'@Description Manual Fibre Feasibility Message
'@Documentation Manual Fibre Feasibility Message
Function SSP_Manual_Fibre_Feasibility_Message()
	if Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebElement("innerhtml:=  Fibre is not available in the location\. Request Manual Feasibility to proceed if required\. ").Exist(0)=true then
	reporter.ReportEvent micPass, "Feasibility", "Fibre is not available in the location.Request Manual Feasibility to proceed if required"
	else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Feasibility Error", "User used Wrong TLC"
	End if
	
End Function

'@Documentation Manage Design and Build Associate Order
'@Description Manage Design and Build Associate Order
Function SSP_Manage_Design_Build_Associate_Order()
If Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").link("text:=Network Availability ").Exist(0)=true  Then
	Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").link("text:=Network Availability ").Click
End If
wait(0.5)
	if Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebButton("name:=Associate").Exist(0)=true then
	wait(0.5)
	Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebButton("name:=Associate").Click
	reporter.ReportEvent micPass, "Associate Button", "Associate Button is present and Enabled"
	else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Associate Button", "Associate Button is present but  disabled"
	End if
	
End Function

'@Description verify Manage Complex Build task is cancelled
'@Documentation verify Manage Complex Build task is cancelled
Function SSP_Manage_Design_Build_Manual_Task_Cancelled()

	SSP_Order_Nav_Order_Task()
	If Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").Webtable("name:=Manage Consent").link("name:=CANCELLED").Exist(0)=true Then
		Reporter.ReportEvent micPass, "Manage Complex Build", "Manage Complex Build Task is cancelled"
		else
		Reporter.ReportEvent micPass, "Manage Complex Build", "Manage Complex Build Task is not cancelled"
	End If
	
End Function


'@Description Amend Schedule order on behalf of RSP
'@Description Amend Schedule on behalf of RSP
Function SSP_Appointment_Chorus_Amending_On_Behalf_Of_RSP(RSP)

	If 	Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebEdit("html id:=amendOnBehalfInput").Exist(0) = True Then
		Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebEdit("html id:=amendOnBehalfInput").set RSP
		Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal").WebButton("innertext:=Amend Order").Click
	End IF
	
End Function

'@Description Complete the Billing Task
'@Documentation Complete the Billing Task
Function SSP_Task_Perform_Billing()
	
	SSP_Order_Nav_Order_Task()
	SSP_Select_Order_Task("Perform Billing")
  	Browser("name:=Chorus Self Service Portal").Sync
   	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("html id:=completeButton").Click
   	wait(1)
   	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("class:=button continue float-right").Click
   	wait(1)
	SSP_Order_Nav_Order_Summary

End Function


'@Description In SSP Page,Change the handover connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
'@Documentation In SSP Page,Change the handover connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
Function SSP_Offer_Charac_set_Handover_connection(Handover)
	
SSP_Order_Nav_Offer_Characteristics
Dim ODesc, All_WebLists, i , Flag
flag = 0

' If Handover is not provided, select the value from the DataTable Resources
If Handover = "" Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_SIT",  "B2B_RSP" , dtLocalSheet 
	DataTable.SetCurrentRow(1)
	Handover = trim(DataTable.Value("B2B_RSP_HandOver_Detail",dtLocalSheet))
End If



set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

Set Odesc = Description.Create

Odesc("micclass").value = "WebList"

set All_Weblists = SSP_Home_Desc.Childobjects(Odesc)

	For i =0 to All_weblists.count - 1
	'msgbox all_weblists(i).GETROPROPERTY("all items")
		If instr(all_weblists(i).GETROPROPERTY("all items"),Handover) <> 0 Then
				'If flag = 0 Then
					all_weblists(i).select Handover
					'flag = 1					
				'Else
					'all_weblists(i).select ""
				'End If
			
		End If
	
	next


SSP_Home_Desc.WebButton("name:=Save").Click
SSP_Order_Nav_Order_Order_Details
End Function

'@Description read's Order details from Summary page and write's into Excel
'@Documentation read's Order details from Summary page and write's into Excel
Function Get_OrderDetails_into_Excel(FileName,Product,WFC_Count)
'Author Swetha
WFC_Count_Temp = WFC_Count
If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("outertext:=Summary ").Exist(0) = True then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("outertext:=Summary ").Click
	Call Create_Excel(FileName,Product) 
	else
	Reporter.ReportEvent micWarning,"Summary Page","summary page not exists"
End If
End Function

Function Create_Excel(FileName,Product)
'Author Swetha
'Create an Microsoft Excel Object using VBScript
Dim objWorkbook,objExcel
'FileName = "D:\test7.xls"
Set objExcel = CreateObject("Scripting.FileSystemobject")
'check's file exists or not,if exists open excel and enter data, else creates new excel file and enters order details
If objExcel.FileExists(FileName) Then
	'msgbox a
	Call OpenExcel_Fill_OtherColumns(FileName,Product)
else
   	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = False
	Set objWorkbook = objExcel.Workbooks.Add()
	objworkbook.SaveAs(FileName) 
		objExcel.Quit
		'Calling function to enter first column
	OpenExcel_Fill_FirstColumn(FileName)
	'calling function to enter order details
	Call OpenExcel_Fill_OtherColumns(FileName,Product)
End If
End Function



'Get current date details
Function FormatDatein_YYYY_MM_DD_HH_MM()
	'YYYY-MM-DD HH:MM
Dim y,m,d,Dateinyyyymmdd
y= year(date)
m = month(date)

if Len(m) = 1 then
	m = "0" & m
End if

d= day(date)


if Len(d) = 1 then
	d = "0" & d
End if

FormatDatein_YYYY_MM_DD_HH_MM = y&"-"&m&"-"&d&" 00:00"
'msgbox FormatDatein_YYYY_MM_DD_HH_MM
End Function

'Get ProductID from summary page
Function Get_ProductID(Product)
'Product = "Business"
Dim oDesc,Linkobjs,Start_string,x,i,zz
set oDesc = Description.Create
oDesc("micclass").value = "WebElement"
oDesc("class").value = "multi-line-field"
'Find all the Product Offers
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)
'obj.Count value has the number of Webelements in the page
For i = 0 to Linkobjs.Count - 1 
			'get the name of all the links in the page 
			x = Linkobjs(i).GetROProperty("innertext")
			'msgbox x
			wait(2)
			If Instr(1,x,Product) <> False Then
			wait(2)
			Start_string = Instr(1,x,"16")
			'msgbox Start_string
			wait(1)
		 	zz =  trim(mid(x,Start_string,Start_string + 9))
			'msgbox zz	
			Get_ProductID = zz
		Exit for
			End If
Next

End Function

'Open file and fill the first column details
Function OpenExcel_Fill_FirstColumn(FileName)
'Author Swetha
'Open Microsoft file using VBScript
Dim objworkbook,objExcel,MySheet
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open FileName
objExcel.Application.Visible = True
Set MySheet = objExcel.ActiveWorkbook.Worksheets("Sheet1")
'Get the Number of Columns used in the Excel Sheet
'Col = MySheet.UsedRange.Columns.Count
'column names 
MySheet.Cells(1,1) = "SerialNo"
MySheet.Cells(1,2) = "OrderID"
MySheet.Cells(1,3) = "LocationID"
MySheet.Cells(1,4) = "SubmittedDate"
MySheet.Cells(1,5) = "PrdouctID"
MySheet.Cells(1,6) = "WFCCount"	
MySheet.Cells(1,7) = "TestRunName"	' RJ adding as feedback after Demo

'Save the Excel Workbook
objExcel.ActiveWorkbook.Save
objExcel.Quit
End Function

'open file and fill all other details
Function OpenExcel_Fill_OtherColumns(FileName,Product)
'Author Swetha
'Open Microsoft file using VBScript
Dim objExcel,MySheet,i,j,k,A,L,Row,Z
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Open FileName
objExcel.Application.Visible = True
Set MySheet = objExcel.ActiveWorkbook.Worksheets("Sheet1")	
'Get the Number of Rows used in the Excel sheet
Row = MySheet.UsedRange.Rows.Count
'msgbox Row
For i = 0 To Row+1 Step 1
	i =Row+1
   	MySheet.Cells(i,1).Value = Row
Next
For j = 0 To Row+1 Step 1
	j=Row+1
	MySheet.Cells(j,2).Value = "'" &SSP_Get_Order_ID 
	'msgbox j
Next 
For k = 0 To Row+1 Step 1
	k=Row+1
	MySheet.Cells(k,3).Value = "'" &SSP_Get_Location_ID
	'msgbox k
Next

For A = 0 To Row+1 Step 1
	A=Row+1
	MySheet.Cells(A,4).Value = "'" &FormatDatein_YYYY_MM_DD_HH_MM
Next
For L = 0 To Row+1 Step 1
	L=Row+1
	MySheet.Cells(L,5).Value = "'" &Get_ProductID(Product)
Next
For Z = 0 To Row+1 Step 1
	Z=Row+1
	MySheet.Cells(Z,6).Value = "'" & WFC_Count_Temp
Next

' RJ adding this for TestRUN Name after feedback from Demo

Dim currentTestRun, testRunAttachmentsFactory
	
If QCUtil.IsConnected Then
		' Get the current test run object
	Set currentTestRun = QCUtil.CurrentRun

	If currentTestRun is Nothing Then
		MySheet.Cells(Row+1,7).Value = "Test is not running from an ALM Test Set in the Test Lab (it's likely to be running from the ALM Test Plan module), so no Test RUN" 
			
	else
		MySheet.Cells(Row+1,7).Value = currentTestRun.Name
	End If

End If



'Save the Excel Workbook
objExcel.ActiveWorkbook.Save
'Close the Excel workbook
objExcel.Quit	

End Function


'@Description In SSP Page,switch the handover connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
'@Documentation In SSP Page,switch the handover connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
Function SSP_Modify_Switch_Handover_connection(Handover1,Handover2)
    
SSP_Order_Nav_Offer_Characteristics
Dim ODesc, All_WebLists, i , Flag
flag = 0
set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")

Set Odesc = Description.Create

Odesc("micclass").value = "WebList"

set All_Weblists = SSP_Home_Desc.Childobjects(Odesc)

    For i =0 to All_weblists.count - 1
    'msgbox all_weblists(i).GETROPROPERTY("all items")
        If instr(all_weblists(i).GETROPROPERTY("all items"),Handover1) <> 0 Then
                    If instr(all_weblists(i).GETROPROPERTY("default value"),Handover1) = 1 Then
                    
                        all_weblists(i).select Handover2
                    else
                        all_weblists(i).select Handover1
                    End If
            
        End If
    
    next


SSP_Home_Desc.WebButton("name:=Save").Click
SSP_Order_Nav_Order_Order_Details
End Function


'@Description In SSP Page,switch the VLAN connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
'@Documentation In SSP Page,switch the VLAN connection which is being passed in Parameter. Needs to call after setting default values in Offer charac page
Function SSP_Modify_Switch_VLAN_connection()
    
'SSP_Order_Nav_Offer_Characteristics
Dim ODesc, All_WebEdit, i , Flag,testVar,SSP_Home_Desc
flag = 0



set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")


if SSP_Home_Desc.webedit("html id:=orderItem0.productOrder0.characteristic8.value").exist(3) = True then

			If SSP_Home_Desc.webedit("html id:=orderItem0.productOrder0.characteristic8.value").getRoProperty ("default value") = "12" Then
	                    
	                  SSP_Home_Desc.webedit("html id:=orderItem0.productOrder0.characteristic8.value").set "13"
	        else
	                  SSP_Home_Desc.webedit("html id:=orderItem0.productOrder0.characteristic8.value").set "12"
	         End If
	         
	         
else

	Reporter.ReportEvent micwarning, "Field Not editable", "check the TLc supplied"
            
End If


SSP_Home_Desc.WebButton("name:=Save").Click
SSP_Order_Nav_Order_Order_Details
End Function


'@Description Chorus User selects Losing Service Provider Name
'@Documentation Chorus User selects Losing Service Provider Name
Function SSP_SelectLosing_Service_ProviderName(Losing_RSP)

Dim Odesc, All_Weblists, i

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

Set Odesc = Description.Create
Odesc("micclass").value = "weblist"
Odesc("html id").value = "losingServiceProviderName"
'Odesc("visible").value ="-1"
set All_Weblists = SSP_Home_Desc.ChildObjects(Odesc)

For i = 0 to All_Weblists.Count -1


	If instr(All_Weblists(i).GetROProperty("all items"),Losing_RSP) <> 0  Then
		All_Weblists(i).Select Losing_RSP
		Reporter.ReportEvent micPass , "Losing_RSP was selected", " Passed Losing_RSP was selected"
	
	else
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail , "Losing_RSP was not selected", " Passed Losing_RSP was not selected. please check the passed parameter or Losing Service Provider doesnt exist. Please check the passed aim, TLC"
	End If
	
	
Next

End Function

'@Description To Enter the Product Instance to be Moved
'@Documentation To Enter the Product Instance to be Moved
Function SSP_Move_Enter_ProductID(ProductID)


if Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebEdit("html id:=originatingLocationExternalProductInstanceId").Exist(3) Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebEdit("html id:=originatingLocationExternalProductInstanceId").set(ProductID)
	wait(1)
	Reporter.ReportEvent micPass , "Product Instance ID", " Product Instance ID entered successfully"
else
	ChorusUtilities_TakeScreenshot
	Reporter.ReportEvent micfail , "Product Instance ID", " Please check the Create Order page and verify correct AIM and product offer has been selected"
End if

End Function

'@Description offer characteristics save button
'@Documentation offer characteristics save button
Function SSP_Offer_Characteristics_Save_Button()
	Dim A
Set A=Browser("name:=Chorus Self Service Portal").Page("title:=Chorus Self Service Portal")
if A.Webbutton("name:=Save").Exist(0)=true then 
A.Webbutton("name:=Save").Click
reporter.ReportEvent micPass, "Save button", "Save button is enabled"
else
reporter.ReportEvent micdone, "Save button", "Save button is disabled"
End If
End Function

'@Description offer characteristics Edit Voice Details using xpath
'@Documentation offer characteristics Edit Voice Details using xpath
Function SSP_Offer_Characteristics_Change_Voice_Details(Voice_Handover_ID,Voice_SVID,Voice_CVID)
SSP_Order_Nav_Offer_Characteristics
If Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Voice Handover Connection')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Voice Handover Connection')]/following-sibling::*").select Voice_Handover_ID 
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Voice VLAN SVID')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Voice VLAN SVID')]/following-sibling::*").set Voice_SVID
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Voice VLAN CVID')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Voice VLAN CVID')]/following-sibling::*").set Voice_CVID
End If

SSP_Order_Nav_Order_Order_Details
End Function

'@Description offer characteristics Edit Kordia Handover Details using xpath - placeholder until detailed specs are avilable with actual field names
'@Documentation offer characteristics Edit Kordia Handover Details using xpath - placeholder until detailed specs are avilable with actual field names
Function SSP_Offer_Characteristics_Change_Data_Details(Data_Handover_ID_1,Data_SVID_1,Data_CVID_1, Data_Handover_ID_2,Data_SVID_2,Data_CVID_2, Data_Handover_ID_3,Data_SVID_3,Data_CVID_3)
msgbox Err.Number
SSP_Order_Nav_Offer_Characteristics
If Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection')]/following-sibling::*").select Data_Handover_ID_1 
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID')]/following-sibling::*").set Data_SVID_1
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID')]/following-sibling::*").set Data_CVID_1
End If
	
If Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 2')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 2')]/following-sibling::*").select Data_Handover_ID_2
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 2')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 2')]/following-sibling::*").set Data_SVID_2
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 2')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 2')]/following-sibling::*").set Data_CVID_2
End If

If Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 3')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 3')]/following-sibling::*").select Data_Handover_ID_3 
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 3')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 3')]/following-sibling::*").set Data_SVID_3
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 3')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 3')]/following-sibling::*").set Data_CVID_3
End If

If Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 4')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebList("xpath:=//label[contains(text(), 'Data Handover Connection 4')]/following-sibling::*").select Data_Handover_ID_4 
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 4')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID 4')]/following-sibling::*").set Data_SVID_4
End If

If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 4')]/following-sibling::*").exist(1) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN CVID 4')]/following-sibling::*").set Data_CVID_4
End If
SSP_Offer_Characteristics_Save_Button
SSP_Order_Nav_Order_Order_Details
End Function


'@Description offer characteristics Change VLAN SVID
'@Documentation offer characteristics Change VLAN SVID
Function SSP_Offer_Characteristics_Change_VLAN_SVID(VLAN_SVID)

SSP_Order_Nav_Offer_Characteristics


If Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID')]/following-sibling::*").exist(5) = True Then
	Browser("name:=Chorus Self Service Portal").WebEdit("xpath:=//label[contains(text(), 'Data VLAN SVID')]/following-sibling::*").set VLAN_SVID
End If
SSP_Offer_Characteristics_Save_Button
SSP_Amend_Amending_on_behalf_of
SSP_Order_Nav_Order_Order_Details

End Function


'@Description Verify Order Staus is not Received before proceeding to SOM Order Processing
'@Documentation Verify Order Staus is not Received before proceeding to SOM Order Processing
'Author : Atul Gupta(600363)
'Date : 11 May 2016
Function Verify_SSP_Order_SubStatus_Not_Received()
DIM A,Flag
Set A = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
Dim oDesc 
A.Link("outertext:=Summary ").Click
Set oDesc = Description.Create 
oDesc("Class Name").value= "WebElement"
oDesc("innertext").value= "Sub-status: Received "

For Iterator = 1 To 20 Step 1
	If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Webelement(oDesc).Exist(2) = True then
		A.Link("outertext:=Summary ").Click
		A.Sync
		wait(3)
	Else
		Flag = 1
		Reporter.ReportEvent micPass, "Order Status", "Order created in SOM.. Proceed to SOM Processing"
		Exit for
	End IF
Next

If Flag <> 1 Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Order Status", "Order still in Acknowledged-Recieved State"
End If

End Function



'******************************************************************************************************************'
'Then following functions are added as per ICMS upload requirements to enable other test cases, Please do not touch
'******************************************************************************************************************'


'Function to Create an ICMS Scope Event as Completed
Function SSP_ICMS_FILE_Creation_Scope(Order_ID)


	
Dim FileName,Find,ReplaceWith,FileContents,dFileContents,Temp,wFileName,dFileContents1,strInput
strInput = Order_ID


FileName     = "C:\Dev\COM\B2B\Atul\ICMS\Scripts\Scope_Completed_ADG_DAT.txt"
Temp = GetFileName()
wFileName = "C:\Dev\COM\B2B\Atul\ICMS\Scope_Completed_ADG_DAT_" & strInput & "_" & GetFileName() & ".txt"

''msgbox wFileName
Find = split(GetLine(FileName),",")

ReplaceWith = mid(Temp,3,8)

'Read source text file
FileContents = GetFile(FileName)

'replace all string In the source file
dFileContents = replace(FileContents, Find(0), ReplaceWith, 1, -1, 1)
dFileContents1 = replace(dFileContents,100029449, strInput, 1, -1, 1)

WriteFile wFileName, dFileContents1
SSP_ICMS_FILE_Creation_Scope = wFileName
End Function

'Function to Create an ICMS Install Event as Servie Given
Function SSP_ICMS_FILE_Creation_Install_Service_Given(Order_ID)

	
Dim FileName,Find,ReplaceWith,FileContents,dFileContents,Temp,wFileName,dFileContents1,strInput
strInput = Order_ID


FileName     = "C:\Dev\COM\B2B\Atul\ICMS\Scripts\Install_ServiceGiven_ADG_DAT.txt"
Temp = GetFileName()
wFileName = "C:\Dev\COM\B2B\Atul\ICMS\Install_ServiceGiven_ADG_DAT_" & strInput & "_" & GetFileName() & ".txt"

'msgbox wFileName
Find = split(GetLine(FileName),",")


ReplaceWith = mid(Temp,3,8)

'Read source text file
FileContents = GetFile(FileName)

'replace all string In the source file
dFileContents = replace(FileContents, Find(0), ReplaceWith, 1, -1, 1)
dFileContents1 = replace(dFileContents,100029449, strInput&"1", 1, -1, 1)

WriteFile wFileName, dFileContents1

SSP_ICMS_FILE_Creation_Install_Service_Given = wFileName

End Function

''Function to Create an ICMS Install Event as Completed
Function SSP_ICMS_FILE_Creation_Install_Completed(Order_ID)

	
Dim FileName,Find,ReplaceWith,FileContents,dFileContents,Temp,wFileName,dFileContents1,strInput
strInput = Order_ID


FileName     = "C:\Dev\COM\B2B\Atul\ICMS\Scripts\Install_Completed_ADG_DAT.txt"
Temp = GetFileName()
wFileName = "C:\Dev\COM\B2B\Atul\ICMS\Install_Completed_ADG_DAT_" & strInput & "_" & GetFileName() & ".txt"

'msgbox wFileName
Find = split(GetLine(FileName),",")

ReplaceWith = mid(Temp,3,8)

'Read source text file
FileContents = GetFile(FileName)

'replace all string In the source file
dFileContents = replace(FileContents, Find(0), ReplaceWith, 1, -1, 1)
dFileContents1 = replace(dFileContents,100029449, strInput&"1", 1, -1, 1)

WriteFile wFileName, dFileContents1

SSP_ICMS_FILE_Creation_Install_Completed = wFileName

End Function


'Read text file
function GetFile(FileName)
 Dim FS, FileStream
  
  If FileName<>"" Then
    Set FS = CreateObject("Scripting.FileSystemObject")
      Set FileStream = FS.OpenTextFile(FileName)
      GetFile = FileStream.ReadAll
  End If
  
End Function

'Write string As a text file.
function WriteFile(FileName_change, Contents)
  Dim OutStream, FS
'msgbox "check" & FileName_change
  ''on error resume Next
  Set FS = CreateObject("Scripting.FileSystemObject")
    Set OutStream = FS.OpenTextFile(FileName_change, 2, True)
    OutStream.Write Contents
End Function

'Function to create a unique File Name
Function GetFileName()
	Dim timestampArray
	Dim timestampString, timestampTime, timestampDate               
	timestampArray = split (Now, ":")
	timestampString = join (timestampArray, "")
'above results in "24/06/2009 160855"
	timestampArray = split (timestampString, " ")
	timestampTime = timestampArray(1)
	timestampDate = timestampArray(0)
	timestampArray = split (timestampDate, "/")
	timestampString = right(timestampArray(2), 2) & timestampArray(1) & timestampArray(0) & timestampTime
	GetFileName = timestampString
End Function

'Read 2nd Line from Static Event Files like Scope/Install
function GetLine(FileName)

  If FileName<>"" Then
    Dim FS, FileStream,abc
    Set FS = CreateObject("Scripting.FileSystemObject")
    
      Set FileStream = FS.OpenTextFile(FileName)
      Do While Not FileStream.AtEndofStream
	abc = FileStream.ReadLine
      Loop
      GetLine = abc
  End If
End Function

'Get Last Line from the File
function GetLastLine(FileName)

  If FileName<>"" Then
    Dim FS, FileStream,abc
    Set FS = CreateObject("Scripting.FileSystemObject")
      
      Set FileStream = FS.OpenTextFile(FileName)
      Do While Not FileStream.AtEndofStream
		abc = FileStream.ReadLine
      Loop
      
      GetLastLine = abc
      
  End If
End Function



'@Description Complete the Create Scope Work Order Task
'@Documentation Complete the Create Scope Work Order Task
Function SSP_Task_Create_Scope_Work_Order(ICMS_SO_Nos)

SSP_Order_Nav_Order_Task()
Browser("name:=Chorus Self Service Portal").Sync
SSP_Select_Order_Task("Create Scope Work Order")
Browser("name:=Chorus Self Service Portal").Sync
	    
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebEdit("html id:=data\.icmsSoId").Set(ICMS_SO_Nos)
wait(1)
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("html id:=completeButton").Click
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("class:=button continue float-right").Click
Browser("name:=Chorus Self Service Portal").Sync
	    
Reporter.ReportEvent micPass, " Create Scope Work Order", " Completed"	    
	    
End Function

'@Description Complete the Create Install Work Order Task
'@Documentation Complete the Create Install Work Order Task
Function SSP_Task_Create_Install_Work_Order(ICMS_SO_Nos)

SSP_Order_Nav_Order_Task()

Browser("name:=Chorus Self Service Portal").Sync
SSP_Select_Order_Task("Create Install.*")
Browser("name:=Chorus Self Service Portal").Sync
	    
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebEdit("html id:=data\.icmsSoId").Set(ICMS_SO_Nos&"1")
	    wait(1)
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("html id:=completeButton").Click
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("class:=button continue float-right").Click
	    Browser("name:=Chorus Self Service Portal").Sync
	    
	    Reporter.ReportEvent micPass, " Create Install Work Order", " Completed"

End Function

'@Description Complete the Create CSE Work Order Task
'@Documentation Complete the Create CSE Work Order Task
Function SSP_Task_Create_CSE_Work_Order(ICMS_SO_Nos)
SSP_Order_Nav_Order_Task()
Browser("name:=Chorus Self Service Portal").Sync
SSP_Select_Order_Task("Create CSE Work Order")
Browser("name:=Chorus Self Service Portal").Sync
	    
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebEdit("html id:=data\.icmsSoId").Set(ICMS_SO_Nos&"1")
	    wait(1)
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("html id:=completeButton").Click
	    Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("class:=button continue float-right").Click
	    Browser("name:=Chorus Self Service Portal").Sync
	    
	    Reporter.ReportEvent micPass, " Create CSE Work Order", " Completed"

End Function


'@Description To Validate that Scope Work Order is completed
'@Documentation To Validate that Scope Work Order is completed
Function SSP_ICMS_Validate_Scope_Completed()

Dim Iterator,SSP_WebTable,Scope_Row_Num
If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("name:=Work Order ").Exist(0) = True Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
	Browser("name:=Chorus Self Service Portal").Sync
		    
	For Iterator = 0 To 30 Step 1
		
		set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
		Scope_Row_Num = SSP_WebTable.GetRowWithCellText("Scope")
		
		' If Scope exists it will always be in the 3rd row
		
		If Scope_Row_Num = 3 Then
		'msgbox SSP_WebTable.GetCellData(Scope_row_num,7)
	    'msgbox strcomp(trim(SSP_WebTable.GetCellData(Scope_row_num,7)),"Completed")
			If strcomp(trim(SSP_WebTable.GetCellData(Scope_Row_Num,7)),"Completed") = 0 Then
				Reporter.ReportEvent micPass, "Scope Work Order", "Scope Work Order Completed"
				Exit for
			else
				Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
				Browser("name:=Chorus Self Service Portal").Sync
				wait(20)
			End If	
			
		Else
			ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, " Create Scope Work Order","  doesn't exist" 
		    Exit for 
		End If
	
	Next
	
	
	If Iterator = 31 Then
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Scope Work Order"," Scope Work Order was not Completed" 
	End If
	
Else 
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Link to Work Order"," Could not navigate to work order, Please check the user ID or User Permissions" 	
End if

End Function



'@Description To Validate that Instal  Work Order is completed
'@Documentation To Validate that Install Work Order is completed
Function SSP_ICMS_Validate_Install_CSE_Completed()

Dim Iterator,Install_Row_Num,Total_Rows,SSP_WebTable,Cell_Data

Install_Row_Num = 0

If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("name:=Work Order ").Exist(0) = True Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
	Browser("name:=Chorus Self Service Portal").Sync
	set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
	Total_Rows = SSP_WebTable.RowCount
	If Total_Rows=4 Then
		Install_Row_Num = Total_Rows
	Else
		For Iterator = 1 To Total_Rows Step 1
			set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
			Cell_Data = SSP_WebTable.GetCellData(Iterator,1)
			If ((Instr(Cell_Data,"Install") > 0) OR (Instr(Cell_Data,"CSE") > 0)) Then
				Install_Row_Num = Iterator
				Exit for
			End If
			

		Next
	End If
	
	If Install_Row_Num = 0 Then
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Install/CSE Work Order", "Install/CSE Work Order does not exist"
		
	End If
	
	For Iterator = 0 To 30 Step 1
		set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
		If strcomp(trim(SSP_WebTable.GetCellData(Install_Row_Num,7)),"Completed") = 0 Then
			Reporter.ReportEvent micPass, "Install/ CSE Work Order", "Install/ CSE Work Order Completed"
			Exit for
		else
			Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
			Browser("name:=Chorus Self Service Portal").Sync
			wait(20)
		End If	
	
	Next
	
	
	If Iterator = 31 Then
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Install Work Order"," Install Work Order was not Completed" 
	End If
	
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Work Order Tab", "Navigation to work order page failed"
End if

End Function

'@Description To Validate that Instal  Work Order is Service given 
'@Documentation To Validate that Install Work Order is Service given

Function SSP_ICMS_Validate_Install_CSE_Service_Given()

Dim Iterator,Install_Row_Num,Total_Rows,SSP_WebTable

Install_Row_Num = 0

If Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("name:=Work Order ").Exist(0) = True Then
	Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
	Browser("name:=Chorus Self Service Portal").Sync
	set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
	Total_Rows = SSP_WebTable.RowCount
	If Total_Rows=4 Then
		Install_Row_Num = Total_Rows
	Else
		For Iterator = 1 To Total_Rows Step 1
			set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
			Cell_Data = SSP_WebTable.GetCellData(Iterator,7)
			If ((Instr(Cell_Data,"Install") > 0) OR (Instr(Cell_Data,"CSE") > 0)) Then
				Install_Row_Num = Total_Rows
				Exit for
			End If
			

		Next
	End If
	
	If Install_Row_Num = 0 Then
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Install Work Order", "Install Work Order does not exist"
		
	End If
	
	For Iterator = 0 To 30 Step 1
		set SSP_WebTable = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("column names:=Work Order List;Last Update")
		If strcomp(trim(SSP_WebTable.GetCellData(Install_Row_Num,7)),"Service Given") = 0 Then
			Reporter.ReportEvent micPass, "Install/ CSE Work Order", "Install/ CSE Work Order Service Given"
			Exit for
		else
			Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Link("Innerhtml:=Work Order ").Click
			Browser("name:=Chorus Self Service Portal").Sync
			wait(20)
		End If	
	
	Next
	
	
	If Iterator = 31 Then
		ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Install Work Order"," Install Work Order was not Completed" 
	End If
	
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Work Order Tab", "Navigation to work order page failed"
End if

End Function


'@Description This will update ICMS WINSCP Batch File
Function ICMS_Update_And_Upload_Winscp_BatchFile(FilePath)
	
dim firstpart,thirdpart,newline,FileName,FileContents,oldLine,dFileContents

'Location of the batch file
FileName = "C:\Dev\COM\B2B\Atul\ICMS\winscp.bat"

'Read source text file
FileContents = GetFile(FileName)

'Fetch the line to Replace
oldLine = GetLastLine(FileName)

firstpart = left(oldLine ,92)
'msgbox firstpart

thirdpart = right(oldLine ,44)
'msgbox thirdpart

newline = firstpart & FilePath & thirdpart
'msgbox newline 

'Replace the oldline with the new line
dFileContents = replace(FileContents,oldLine,newline, 1, -1, 1)

'Call the Write File Function
Call WriteFile(FileName, dFileContents)

wait(4)
ICMS_Invoke_Winscp_Batch_File

End Function

' Invoke Batch file which will upload ICMS files to winscp server
Function ICMS_Invoke_Winscp_Batch_File()
    Dim objWshShell
	Set objWshShell = CreateObject("Wscript.Shell")
	objWshShell.run "C:\Dev\COM\B2B\Atul\ICMS\winscp.bat", 1 ' static name
	Set objWshShell = Nothing
End Function

'@Description This function will call ICMS functions and will upload the ICMS spool files
'@Documentation This function will call ICMS functions and will upload the ICMS spool files
Function SSP_Driver_To_Call_ICMS_Functions(Order_ID,Appointment_Check_Results)

Dim ScopeFlag, InstallFlag, CSEFlag
ScopeFlag = 0
InstallFlag = 0 
CSEFlag = 0

' Here Install is commonly used for Install as well as CSE
' InstalyesScopeyesCSeno

'*******************'***************************************************
'Check which type of appointment is present
'*******************'***************************************************

	If instr(1,Appointment_Check_Results,"Scopeyes",1) <> 0 Then
	
    	ScopeFlag = 1
    	
	End If

	If (instr(1,Appointment_Check_Results,"Installyes",1) <> 0) Then
	
	    InstallFlag = 1
	    
	End If
	
	If (instr(1,Appointment_Check_Results,"Installyes",1) = 0 And instr(1,Appointment_Check_Results,"CSEyes",1) <> 0) Then
	
	    CSEFlag = 1
	    
	End If

'*******************'***************************************************
' Create Work Orders in SSP
'*******************'***************************************************
If ScopeFlag = 1 Then
	Call SSP_Task_Create_Scope_Work_Order(Order_ID)
End If	

If InstallFlag = 1 Then
	Call SSP_Task_Create_Install_Work_Order(Order_ID)
End If	


If CSEFlag = 1 Then
	Call SSP_Task_Create_CSE_Work_Order(Order_ID)
End If	

'*******************'***************************************************
' Create ICMS Files and get the File Path which we need to upload, FIle Path would be the output, Upload the file and wait
'*******************'***************************************************

Dim FilePath_Scope, FilePath_Install_Service_Given, FilePath_Install_Completed

If ScopeFlag = 1 Then
	FilePath_Scope =  SSP_ICMS_FILE_Creation_Scope(Order_ID)
	call ICMS_Update_And_Upload_Winscp_BatchFile(FilePath_Scope)
	Call SSP_ICMS_Validate_Scope_Completed()
	
End if

If InstallFlag = 1 or CSEFlag = 1 Then

	 	'FilePath_Install_Service_Given = SSP_ICMS_FILE_Creation_Install_Service_Given(Order_ID)

		FilePath_Install_Completed = SSP_ICMS_FILE_Creation_Install_Completed(Order_ID)
		
		' Upload and then wait
		'call ICMS_Update_And_Upload_Winscp_BatchFile(FilePath_Install_Service_Given)
'		Call SSP_ICMS_Validate_Install_CSE_Service_Given()
		' wait for 1 sec
'		wait(3)
		call ICMS_Update_And_Upload_Winscp_BatchFile(FilePath_Install_Completed)
		SSP_ICMS_Validate_Install_CSE_Completed()
End if

'*******************'************************************************************

' Add a functionality to move these files from Current Folder to Archieve folder
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")
objFSO.MoveFile "C:\Dev\COM\B2B\Atul\ICMS\*.txt" , "C:\Dev\COM\B2B\Atul\ICMS\Archieved\"

End Function


Function SSP_Manual_Intervention()
	msgbox("Please complete the manual activity, Then Press OK")
	msgbox SSP_Home_Desc.Exist
	
End Function


Function SSP_Check_If_Cancel_Button_Present()

SSP_Order_Nav_Order_Summary   ' This will navigate to the Summary page
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")


If     SSP_Home_Desc.WebButton("HTML id:=CancelOrder").Exist(10)= true Then
		Reporter.reportevent micpass, "Order ", "Order Submitted Properly"
	
Else
		Reporter.reportevent micfail, "Order ", "Order not Submitted Properly"
		

End IF

End Function



'@Description This Function will search Order ID 
'@Documentation This Function will search Order ID 
Function SSP_SearchOrder_Search_Order_ID_1(OrderID)

'Navigate to Search Order Page
SSP_Nav_Search_Orders()
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If (SSP_Home_Desc.Webedit("name:=searchValue").Exist(10) = True) then
	SSP_Home_Desc.Webedit("name:=searchValue").set OrderID
	SSP_Home_Desc.WebButton("html id:=button-quick-search").Click
Else
	Reporter.ReportEvent micFail ,"Search Order", "Search Order Page Failed"
End If
SSP_SearchOrder_Search_Order_ID_1 = OrderID
End Function

'@Description This Function will Click on Searched Order ID 
'@Documentation This Function will Click on Searched Order ID 
Function SSP_SearchOrder_Click_on_Searched_Order_2(OrderID)

Dim oDesc, Linkobjs,i,x
set oDesc = Description.Create
oDesc("micclass").value = "Link"

'Find all the Product Offers
Set Linkobjs = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDesc)

For i = 0 to Linkobjs.Count - 1 
	'get the name of all the links in the page 
	x = Linkobjs(i).GetROProperty("innerhtml")
	If Instr(1,x,OrderID) <> False Then
		Linkobjs(i).click
		Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
		Reporter.ReportEvent micpass ,"Order Present","Clicked on Order"
		Exit for
	End If
Next

If i = Linkobjs.Count Then	
	Reporter.ReportEvent micFail ,"Order Not Present","Please check the order Number"
End If


End Function

'@Description Cancel all of the orders which are present in MYSQL Database, Table Orders_To_Be_Cancelled and mark cancelled as Yes
'@Documentation Cancel all of the orders which are present in MYSQL Database, Table Orders_To_Be_Cancelled and mark cancelled as Yes
Function SSP_Utilities_To_Cancel_All_Orders_Present_in_Cancellation_Database()
Dim oCn,ConnectionString,Sql,updateSQL
Dim OrderNumbers(200) ' Max 200 orders can be cancelled in one go to preserve Memory of test tool box
Set oCn = CreateObject("ADODB.Connection")
sql = "SELECT OrderNumber FROM Orders_To_Be_Cancelled WHERE Cancelled is NULL"

ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=localhost;DATABASE=b2b_automation_71327;USER=root;PASSWORD=;OPTION=11;"
oCn.open(ConnectionString)    'Open your connection

Dim rs
Set rs = CreateObject("ADODB.Recordset")
rs.open Sql,oCn,3,3
Dim i,a : i  = 0

if rs.EOF then
  Reporter.ReportEvent micdone,"Database No Orders","There are no Orders to be deleted.Please verify if Connect Order was successful"
else
    Do while NOT rs.EOF
        OrderNumbers (i) =rs.Fields(0).Value
        rs.MoveNext
        i= i +1
    Loop
end if

For a = 0 to i-1
'msgbox OrderNumbers(a)
    Call SSP_Nav_Search_Orders()
	Call SSP_SearchOrder_Search_Order_ID_1(OrderNumbers(a))
    
    Call SSP_SearchOrder_Click_on_Searched_Order_2(OrderNumbers(a))
    Call SSP_Check_If_Cancel_Button_Present()
    Call SSP_Task_Initiate_Cancel_Order()
    wait(20)
    Call SSP_Order_Nav_Order_Summary()
    Call SSP_Task_Chorus_Approve_Cancel_Order()
    wait(30)
    Call SSP_Chorus_Complete_Order_Cancellation()
    
Next
Dim Cancelled_Order_No
updateSQL = "Update Orders_To_Be_Cancelled set Cancelled = 'Yes' Where OrderNumber = '"+Cancelled_Order_No+"'"
'"UPDATE datainputs SET OrderNo = '"+Order_ID +"' WHERE TESTCASENAME = '"+trim(Test_Case)+"'"

For a = 0 to i-1

    Cancelled_Order_No = trim(OrderNumbers(a))
    updateSQL = "Update Orders_To_Be_Cancelled set Cancelled = 'Yes' Where OrderNumber = '"+Cancelled_Order_No+"'"
'    'msgbox updateSQL
     oCN.Execute updateSQL

Next

oCn.close()

    
End Function

'@Description Function to Validate Order Status Order Status and Sub Status
'@Documentation Function to Validate Order Status Order Status and Sub Status

Function SSP_Summary_Verify_Order_Status_SubStatus(status,SubStatus)

Dim oDesc 

set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")

if SSP_Home_Desc.Webelement("innertext:=Status: "&Status& ".*").Exist = True then
	
	Reporter.ReportEvent micPass, "Order Status", "Order Status is: "&status
else

	Reporter.ReportEvent micFail, "Order Status", "Order Status is not: "&status
	
End If


if SSP_Home_Desc.Webelement("innertext:=Sub-status: "&SubStatus& ".*").Exist = True then

	Reporter.ReportEvent micPass, "Order Sub-Status", "Order Sub Status is: "&SubStatus
	
else

	Reporter.ReportEvent micFail, "Order Sub-Status", "Order Sub Status is not : "&SubStatus
	
End If

End Function


'@Description Function to Navigate to Notification Tab
'@Documentation Function to Navigate to Notification Tab
Function SSP_Nav_Notifications()

Dim SSP_Home_Desc

set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")

SSP_Home_Desc.Link("outerhtml:=<a href=""/chorus-ssp-web/pages/Notification/Search"" s-key=""n"" s-key-type=""1"">Notifications</a>").click

SSP_Home_Desc.Sync

End Function

'@Description Function to Search a Notification and Take Action
'@Documentation Function to Search a Notification and Take Action
Function SearchAndActionNotificationAtTLC(TLC,Action)

Dim SSP_Home_Desc,RowNo

set SSP_Home_Desc = browser("name:=Chorus Self Service Portal").page("title:=Chorus Self Service Portal")

SSP_Home_Desc.WebEdit("html id:=tlc").Set TLC

SSP_Home_Desc.WebButton("name:=Search").Click

SSP_Home_Desc.Sync


if SSP_Home_Desc.WebElement("innertext:=No Matching Notification Instances found. ").Exist = True then
	
	Reporter.ReportEvent micFail, "Notification", "Notification is not generated at TLC : "&TLC
	
else
	RowNo = SSP_Home_Desc.WebTable("html id:=searchResultTable").GetRowWithCellText("New")
	If RowNo = -1 Then
		Reporter.ReportEvent micFail, "Notification", "No Notification in New State at TLC: "&TLC
	Else
		'Select the Notification with New State
		SSP_Home_Desc.WebTable("html id:=searchResultTable").ChildItem(RowNo,1,"WebCheckBox",0).Set "ON"
		
		'Click on Accept/Reject/Acknowledge
		SSP_Home_Desc.WebButton("innertext:="&Action).Click
		
		SSP_Home_Desc.Sync
		
		
	End If
	
End If

End Function


'@Description Chorus User Adds Interaction
'@Documentation Chorus User Adds Interaction
Function SSP_Chorus_Add_Interaction()
Dim SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

SSP_Order_Nav_Order_Summary()

If SSP_Home_Desc.WebButton("html id:=addInteraction").Exist(1) = True Then
    SSP_Home_Desc.WebButton("html id:=addInteraction").Click
	'
	SSP_Home_Desc.WebList("html id:=interactionType").WaitProperty "items count", micGreaterThan(10), 5000
	SSP_Home_Desc.WebList("html id:=interactionType").Select("Order Status Update and Note")
    SSP_Home_Desc.Webbutton("html id:=popupSaveButton").Click
    SSP_Home_Desc.Sync
Else
    ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail, "Add Interaction was not Initiated" ,"Add Interaction was not initiated"
End If
End Function

Function SSP_Chorus_Add_Interaction_Bulk(NumInteractions)

For Iterator = 1 To NumInteractions Step 1
	Call SSP_Chorus_Add_Interaction()
Next

End Function


'@Description Launches the Browser
'@Documentation Launches the Browser with the URL given in the parameter
Function SSP_Launch_Browser_Brendon(Environment)
	' Current desktop policy means that the Screensaver kicks in after 10 minutes and locks the desktop.  This prevents scheduled tests from running.
	' To get around this, a utility called MouseJiggle.exe can be used to simulate the mouse being moved every 5 seconds.  However, when the test is running
	' it's best to shutdown MouseJiggle using the following CloseProcessByName.
systemutil.CloseProcessByName "MouseJiggle.exe"
' 'Close all of the open browsers to ensure that no other browser is open
	'CloseAllBrowserExceptALM
CloseAllBrowserExceptALM
'***************************
Dim URL_Name : URL_Name = "SSP_URL"
Dim i, URLAddress

If (Environment = "EMMA SIT") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_EMMA_SIT",  "URLs" , dtLocalSheet 
ElseIf (Environment = "EMMA Prod") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_EMMA_Prod",  "URLs" , dtLocalSheet 
ElseIf (Environment = "Pre Prod") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_Pre_Prod",  "URLs" , dtLocalSheet 	
Else
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_SIT",  "URLs" , dtLocalSheet 
End If

dim Rowcount : Rowcount= DataTable.GetSheet(dtLocalSheet).GetRowCount
For i =1 To Rowcount
	If strcomp(trim(DataTable("coffiiURLs",dtLocalSheet)),trim(URL_Name),1) = 0 Then
		URLAddress = DataTable("URLAddress",dtLocalSheet)
		Systemutil.Run "iexplore.exe" ,URLAddress
		SSP_Check_Session_Timed_out
		Exit for
	else
		DataTable.setnextrow   
	End If	
Next

If i > Rowcount Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Login", "Please check the URL Details in Resources"
End If

End Function

'@Description Enter login details
'@Documentation Enter login details
Function SSP_Enter_Login_Details_Brendon(UserID, Environment)

Dim SSPPassword,i

If (Environment = "EMMA SIT") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_EMMA_SIT",  "SSP_Login_Info" , dtLocalSheet 
ElseIf (Environment = "EMMA Prod") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_EMMA_Prod",  "SSP_Login_Info" , dtLocalSheet 
ElseIf (Environment = "Pre Prod") Then
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_Pre_Prod",  "SSP_Login_Info" , dtLocalSheet 
Else
	Datatable.Importsheet "[QualityCenter\Resources] Resources\BPT Resources\Data\Environment_Data_SIT",  "SSP_Login_Info" , dtLocalSheet 
End If 

dim Rowcount : Rowcount= DataTable.GetSheet(dtLocalSheet).GetRowCount
For i =1 To Rowcount
	If strcomp(trim(DataTable("UserID",dtLocalSheet)),trim(UserID),1) = 0 Then
		SSPPassword = DataTable("Password",dtLocalSheet)
		Exit for
	else
		DataTable.setnextrow   
	End If	
Next

If i > Rowcount Then
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail ,"Login", "Cannot find the specified user ID (" & UserID & "). Please check the login Details in Test Resources"
End If

If Browser("name:=Sign In").Page("Title:=Sign In").Exist(10) =True Then
	Browser("name:=Sign In").Page("Title:=Sign In").WebEdit("Class Name:=WebEdit","html id:=ContentPlaceHolder1_UsernameTextBox").Set UserID
	Browser("name:=Sign In").Page("Title:=Sign In").WebEdit("Class Name:=WebEdit","html id:=ContentPlaceHolder1_PasswordTextBox").Set SSPPassword
	wait(0.5) ' To sync up with the Page
	Browser("name:=Sign In").Page("Title:=Sign In").WebButton("name:=Sign In").Click
	
Else
	wait(2) 'Just wait for the page to load
	Reporter.ReportEvent micDone, "Login", "SSP bypassed the login credentials"
End If

End Function


'@Description Chorus User will add quotes
'@Documentation Chorus User will add quotes
Function SSP_Charges_Add_Quote()
	
	DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.link("name:=Charges ").Exist(5) = True Then
	SSP_Home_Desc.link("name:=Charges ").Click
	
	' Add Quote here
	If SSP_Home_Desc.webEdit("html id:=quoteNumber").GetROProperty("readonly") <> 1 Then
		 SSP_Home_Desc.webEdit("html id:=quoteNumber").Set("12")
		 SSP_Home_Desc.webEdit("html id:=quoteDescription").Set("This is quote added by COFFII Automation Team")
		SSP_Home_Desc.webEdit("html id:=quoteValue").Set("20")
		SSP_Home_Desc.WebButton("html id:=saveQuoteButton").Click
		SSP_Order_Nav_Order_Summary
	else
		Reporter.ReportEvent micFail,"Quote","not able to add quote . Check if there is already quote present"	
	End If
	
Else	
	Reporter.ReportEvent micFail,"Charges" , "Charges Tab not available, Please check your Login Credenitals"	
End If


End Function

'@Description If the popup comes , Enter the details for Amending on Behalf of
'@Documentation If the popup comes , Enter the details for Amending on Behalf of
Function SSP_Amend_Amending_on_behalf_of()
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebEdit("html id:=amendOnBehalfInput").exist(10) = True Then
	SSP_Home_Desc.WebEdit("html id:=amendOnBehalfInput").Set "B2BAutoRSP"
	SSP_Home_Desc.WebButton("html id:=popupAmendSaveButton").click
End If
End Function

'@Description Click the Update Order Task
'@Documentation Click the Update Order Task
Function SSP_Amend_Click_on_Update_Order_Task()

SSP_Select_Order_Task("Update Order")

End Function

'@Description Complete the Amend Task
'@Documentation Complete the Amend Task
Function SSP_Amend_Complete_Amend_Task()
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
if SSP_Home_Desc.WebButton("html id:=completeButton").Exist(5) = True Then
	SSP_Home_Desc.WebButton("html id:=completeButton").Click
	
	if SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Exist(5) = True Then
		SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Click
	End If
Else
	Reporter.ReportEvent micfail ,"Amend Order", "Task Not Completed"
End if 

End Function



'@Description Amend to change contact details
'@Documentation Amend to change contact details
Function SSP_Amend_Contact_Details()

SSP_Order_Nav_Order_Order_Details

Dim oDescName,Contact_Names,i
set oDescName = Description.Create
oDescName("micclass").value = "WebEdit"
oDescName("class").value = "input-large contact-name alphabet-only"

Set Contact_Names = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDescName)

For i = 0 To Contact_Names.count -1 
	Contact_Names(i).Set "AutomationAmend " &i+1
Next

' This will enter for all Preferred Phone and Alternate Phone
Dim oDescNo,Contact_Phone,j
set oDescNo = Description.Create
oDescNo("micclass").value = "WebEdit"
oDescNo("class").value = "input-large"

Set Contact_Phone = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").ChildObjects(oDescNo)

For j = 0 To Contact_Phone.count -1 
	Contact_Phone(j).Set "5675875" &j
Next

Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebButton("class:=button amend ssp-action-btn").Click

SSP_Amend_Amending_on_behalf_of
End Function

'@Description Amend to add customer Interaction (via Interaction button - not to be confused with Add Interaction which is a Chorus only function)
'@Documentation Amend to add customer Interaction (via Interaction button - not to be confused with Add Interaction which is a Chorus only function)
Function SSP_Amend_Customer_Interaction()
SSP_Order_Nav_Order_Summary

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebButton("html id:=interaction").Exist(4) = True Then
	SSP_Home_Desc.WebButton("html id:=interaction").click
	SSP_Home_Desc.Sync
	' This will pop uo a webedit with html id amendTemplate
	If SSP_Home_Desc.WebEdit("html id:=amendTemplate").Exist(4) = True  Then
		SSP_Home_Desc.WebEdit("html id:=amendTemplate").Set "This is Interaction Added by B2B Automation Team"
		
		SSP_Home_Desc.WebButton("html id:=popupInteractionSaveButton").click
		
	Else
		Reporter.ReportEvent micFail,"Interaction","Interaction Button Doesnt exist"
	End If
	
Else
	Reporter.ReportEvent micFail,"Interaction","Interaction Button Doesnt exist"
End If
	
End Function



'@Description Click the Manage Customer Interaction
'@Documentation Click the Manage Customer Interaction
Function SSP_Amend_Click_on_Manage_Customer_Interaction()

SSP_Select_Order_Task("Manage Customer Interaction")

End Function


'@Description Complete the Manage Customer Interaction  Task
'@Documentation Complete the Manage Customer Interaction  Task
Function SSP_Amend_Complete_Manage_Customer_Interaction_Task()
DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
If SSP_Home_Desc.WebEdit("html id:=interactionResponse").Exist(5) = True Then
	SSP_Home_Desc.WebEdit("html id:=interactionResponse").Set "B2B Automation Team is looking into your interaction."
End If


if SSP_Home_Desc.WebButton("html id:=completeButton").Exist(5) = True Then
	SSP_Home_Desc.WebButton("html id:=completeButton").Click
	
	if SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Exist(5) = True Then
		SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Click
	End If
Else
	Reporter.ReportEvent micfail ,"Amend Order", "Task Not Completed"
End if 

End Function


'@Description Amend to Escalate Order
'@Documentation Amend to Escalate Order
Function SSP_Amend_Escalate_Order()
SSP_Order_Nav_Order_Summary

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebButton("html id:=escalateOrder").Exist(4) = True Then
	SSP_Home_Desc.WebButton("html id:=escalateOrder").click
	SSP_Home_Desc.Sync
	' This will pop uo a webedit with html id amendTemplate
	If SSP_Home_Desc.WebEdit("html id:=escalateText").Exist(4) = True  Then
		SSP_Home_Desc.WebEdit("html id:=escalateText").Set " This is an escalation raised by B2B Automation Team"
		SSP_Home_Desc.weblist("html id:=escalateReason").select(1)
		SSP_Home_Desc.WebButton("html id:=popupEscalateSaveButton").click
		
	Else
		Reporter.ReportEvent micFail,"Escalation ","Escalation Button Doesnt exist"
	End If
	
Else
	Reporter.ReportEvent micFail,"Escalation ","Escalation Button Doesnt exist"
End If


	
End Function	




'@Description Click on Process Escalation Task
'@Documentation Click on Process Escalation Task
Function SSP_Amend_Click_on_Process_Escalation_Task()

SSP_Select_Order_Task("Process Escalation")

End Function


'@Description Complete Process Escalation Task. EscalationDecision ("Accepted" or "Declined") must be provided.
'@Documentation Complete Process Escalation Task. EscalationDecision ("Accepted" or "Declined") must be provided.
Function SSP_Amend_Task_Complete_Process_Escalation(EscalationDecision)

DIM SSP_Home_Desc
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

If SSP_Home_Desc.WebEdit("html id:=notes").Exist(5) = True Then
	SSP_Home_Desc.WebEdit("html id:=notes").Set "B2B Automation Team has escalated your order."
End If


if SSP_Home_Desc.WebButton("html id:=completeButton").Exist(5) = True Then
	SSP_Home_Desc.Weblist("html id:=escalationResult").Select(EscalationDecision) 'Accepted or Declined
	SSP_Home_Desc.WebButton("html id:=completeButton").Click
	
	if SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Exist(5) = True Then
		SSP_Home_Desc.WebButton("html id:=complete-popupContinueButton").Click
	End If
Else
	Reporter.ReportEvent micfail ,"Amend Order", "Task Not Completed"
End if 

	
	
	
End Function


Function SSP_Select_RescheduleReason()
DIM SSP_Home_Desc,ODesc1,ArrOKButtons
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
If SSP_Home_Desc.weblist("html id:=rescheduleReason").Exist(1) = true Then
	SSP_Home_Desc.weblist("html id:=rescheduleReason").Select(2)
	Set ODesc1 = Description.Create()
	ODesc1("html tag").value = "BUTTON"
	ODesc1("name").value = "OK"
	ODesc1("class").value = "button float-right spaceAbove"
	ODesc1("micclass").value = "WebButton"
	set ArrOKButtons = SSP_Home_Desc.ChildObjects(ODesc1)
	ArrOKButtons(0).Click
	SSP_Home_Desc.Sync
End If
SSP_Reschedule_Confirm_Charges
End Function

Function SSP_Reschedule_Confirm_Charges()
DIM SSP_Home_Desc,ODesc1,ArrOKButtons,i
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

Dim Odesc2,ReschElement
Set Odesc2 = Description.Create()
	Odesc2("outertext").value = " Confirm RescheduleRescheduling this appointment now might incur charges, do you want to proceed\? Cancel OK "
	Odesc2("class").value = "ui-dialog ui-widget ui-widget-content ui-corner-all ui-front"
	Odesc2("micclass").value = "WebElement"
	
	set ReschElement = SSP_Home_Desc.ChildObjects(Odesc2)
	If ReschElement.count = 1 Then
		Set ODesc1 = Description.Create()
		ODesc1("html tag").value = "BUTTON"
		ODesc1("innertext").value = "OK"
		ODesc1("class").value = "button float-right spaceAbove"
		ODesc1("micclass").value = "WebButton"
		set ArrOKButtons = SSP_Home_Desc.ChildObjects(ODesc1)
		ArrOKButtons(0).Click	
		SSP_Home_Desc.Sync
End If
End Function

'@Description Click on the specified Task in the Order Tasks table
'@Documentation Click on the specified Task in the Order Tasks table
Function SSP_Select_Order_Task(TaskName)

DIM SSP_Home_Desc,oDesc, Linkobjs, Iterator, LinkFound
LinkFound = false
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")

SSP_Order_Nav_Order_Task()

'Try to find the task with the specified name. Loop and retry as it can take a little time for COM to process and create new tasks.
For Iterator = 1 To 10 Step 1
	'Reload the page
	'Find the link with the specified name
	set oDesc = Description.Create
	oDesc("micclass").value = "Link"
	oDesc("innerhtml").value = TaskName
	'MsgBox TaskName
	Set Linkobjs = SSP_Home_Desc.ChildObjects(oDesc)
	If Linkobjs.Count > 0 Then
		'MsgBox CStr(Linkobjs.Count) + " " + TaskName + " links found"
		LinkFound = true
		Linkobjs(0).click 'click on the first matching Task, in case there is more than one with the same name
		SSP_Home_Desc.Sync
		Exit For
	Else
		Reporter.ReportEvent micDone,"Select Order Task", "Order Task '" & TaskName & "' not found. Reloading the Tasks page and retrying (10 attempts will be made before giving up)."
		SSP_Order_Nav_Order_Task()
	End If
Next

If LinkFound = true Then
	Reporter.ReportEvent micDone,"Select Order Task","Order Task '" & TaskName & "' found and clicked"
Else
	ChorusUtilities_TakeScreenshot : Reporter.ReportEvent micfail,"Select Order Task","Order Task '" & TaskName & "' was not found"
End If

End Function

'@Description : Function To Validate the CSE SLA Check
'@Documentation : Function To Validate the CSE SLA Check
Function SSP_Validate_Default_CSE_Date()

Dim SSP_Home_Desc,ODescTable,ListOfWebTables,ODesc_List_Of_Dates,List_Of_Dates,Temp
Dim Selected_CSE_Date,Selected_CSE_Month,Selected_CSE_Year,Booked_CSE_Date,Date1,CurrentDay

'Main Outer Table
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")


set ODescTable = Description.Create()
ODescTable("micclass").value = "WebTable"
ODescTable("column names").value = "Su;Mo;Tu;We;Th;Fr;Sa"


Set ListOfWebTables = SSP_Home_Desc.ChildObjects(ODescTable)


'Set Object Description for Links in both the Table
Set ODesc_List_Of_Dates = Description.Create()
ODesc_List_Of_Dates("micclass").Value = "Link"

'Scope
Set List_Of_Dates = ListOfWebTables(0).ChildObjects(ODesc_List_Of_Dates)


Temp =0
If (List_Of_Dates.Count > 0) Then
	
	Selected_CSE_Date = cint(trim(List_Of_Dates(0).GetROProperty("innertext")))
	Selected_CSE_Month = SSP_Home_Desc.WebElement("class:=ui-datepicker-month","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[1]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[1]").GetROProperty("innerhtml")
	Selected_CSE_Year = SSP_Home_Desc.WebElement("class:=ui-datepicker-year","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[1]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[2]").GetROProperty("innerhtml")
	
	Booked_CSE_Date = cdate(DateValue(Selected_CSE_Date &"-" & Selected_CSE_Month &"-" & Selected_CSE_Year))
	Date1 = Date
	CurrentDay = Weekday(Date1)
	'Satuday Logic
	If CurrentDay =  7 Then
		Temp = 1
	
	'Sunday Logic
	ElseIf CurrentDay = 1 Then
		Temp = 0
	'Other Week Days Logic (Mon - Friday)
	Else
	
		While (Date1 <= Booked_CSE_Date)
			CurrentDay = Weekday(Date1)
			If (CurrentDay = 1 or CurrentDay = 7 )Then
				Temp = Temp + 1
			End If
			Date1=DateAdd("d",1,Date1)
		
		Wend
		
	End If
	

	
	If ((DateDiff("d",Date,Booked_CSE_Date) - Temp) > 2) then
		Reporter.ReportEvent micWarning, "Default CSE Day is more than 2 days Apart","Check if this is due to Slot unavailability"
		
	ElseIf ((DateDiff("d",Date,Booked_CSE_Date) - Temp) = 2) Then
		Reporter.ReportEvent micPass, "Default CSE Day is 2 days from Today","CSE SLA Check Passed"
	Else
		Reporter.ReportEvent micFail, "Default CSE  Day cannot be on or before Day1","CSE SLA Check Failed"
	End if
	
Else

	Reporter.ReportEvent micFail, "CSE Slots not Available","Please updload AF Slots for this Region"
	
End IF

End Function

'@Description : Function To Validate the RFS SLA Check
'@Documentation : Function To Validate the RFS SLA Check
Function SSP_Validate_Default_RFS_Date()

Dim SSP_Home_Desc,Selected_RFS_Date,Temp,DayofWeek
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").WebTable("Class Name:=WebTable","column names:=Bookable Appointments")

Selected_RFS_Date = Cdate(SSP_Home_Desc.WebEdit("Class Name:=WebEdit","class :=hasDatepicker").GetROProperty("value"))

Temp = 0
DayofWeek = WeekDay(Date)
	
'IF Current Day is Friday Scope Would be on Monday
If DayofWeek = 7 Then
	Temp = 1
End If

If ((DateDiff("d",Date,Selected_RFS_Date) - Temp) > 1 ) then
	Reporter.ReportEvent micWarning, "RFS Day Should Ideally be next Day from Today","SLA Check failed.. Please check the Schedule Page of your order"
ElseIf ((DateDiff("d",Date,Selected_RFS_Date) - Temp ) = 1) Then
	Reporter.ReportEvent micPass, "RFS is on the next Business Day","SLA Check Passed"
	
Else
	Reporter.ReportEvent micFail, "RFS Can't be before or on Day0 ","SLA Check Failed"
End IF

End Function

'@Description : Function To Validate the Scope and Install  SLA Check
'@Documentation : Function To Validate the Scope and Install SLA Check
Function SSP_Validate_Default_Scope_Install_Dates()

DIM SSP_Home_Desc,ODescTable,ListOfWebTables,ODesc_List_Of_Dates,Temp,List_Of_Dates
Dim Selected_Scope_Date,Selected_Scope_Month,Selected_Scope_Year,Booked_Scope_Date,DayofWeek,Date1
Dim Selected_Install_Date,Selected_Install_Month,Selected_Install_Year,Booked_Install_Date,Day2
Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Sync
'Main Outer Table
Set SSP_Home_Desc = Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal")
'.WebTable("Class Name:=WebTable","column names:=Schedule appointments")
'Scope and Install Table
set ODescTable = Description.Create()
ODescTable("micclass").value = "WebTable"
ODescTable("class").value = "ui-datepicker-calendar"
Set ListOfWebTables = SSP_Home_Desc.ChildObjects(ODescTable)
msgbox ListOfWebTables.count
'Set Object Description for Links in both the Table
Set ODesc_List_Of_Dates = Description.Create()
ODesc_List_Of_Dates("micclass").Value = "Link"

'Scope
Set List_Of_Dates = ListOfWebTables(0).ChildObjects(ODesc_List_Of_Dates)
msgbox List_Of_Dates.Count
If (List_Of_Dates.Count > 0) Then
	
	Selected_Scope_Date = cint(trim(List_Of_Dates(0).GetROProperty("innertext")))
	Selected_Scope_Month = SSP_Home_Desc.WebElement("class:=ui-datepicker-month","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[1]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[1]").GetROProperty("innerhtml")
	Selected_Scope_Year = SSP_Home_Desc.WebElement("class:=ui-datepicker-year","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[1]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[2]").GetROProperty("innerhtml")
	
	Booked_Scope_Date = cdate(DateValue(Selected_Scope_Date &"-" & Selected_Scope_Month &"-" & Selected_Scope_Year))
	msgbox Booked_Scope_Date
	'Set Temporary Variable to 0
	Temp = 0
	DayofWeek = WeekDay(Date)
	
	'IF Current Day is Friday Scope Would be on Monday
	If DayofWeek = 6 Then
		Temp = 2
	End If
	
	If DayofWeek = 7 Then
		Temp = 1
	End If
	
	If ((DateDiff("d",Date,Booked_Scope_Date) - Temp) > 1 ) then
		Reporter.ReportEvent micWarning, "Scope Should Ideally be next Day from Today","Check if this is due to Slot unavailability"
	ElseIf ((DateDiff("d",Date,Booked_Scope_Date) - Temp ) = 1) Then
		Reporter.ReportEvent micPass, "Scope follows the minimum SLA of 1 Day from Day 0","Scope SLA Check Passed"
		
	Else
		Reporter.ReportEvent micFail, "Scope Can't be before or on Day0 ","Scope SLA Check Failed"
	End IF
	
	'Assign Selected Scoep Date to a Temporary Variable to avoid loosing Selected Scope Date
	Date1 = Booked_Scope_Date
Else

	Reporter.ReportEvent micFail, "Scope Slots not Available","Please updload SC Slots for this Region"
	
End IF

'Fetch Selected Install Date
Set List_Of_Dates = ListOfWebTables(1).ChildObjects(ODesc_List_Of_Dates)

If (List_Of_Dates.Count > 0) Then
	
	Selected_Install_Date = cint(trim(List_Of_Dates(0).GetROProperty("innertext")))
	Selected_Install_Month = SSP_Home_Desc.WebElement("class:=ui-datepicker-month","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[2]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[1]").GetROProperty("innerhtml")
	Selected_Install_Year = SSP_Home_Desc.WebElement("class:=ui-datepicker-year","xpath:=//FORM[@id='appointmentForm']/TABLE[1]/TBODY[1]/TR[5]/TD[2]/DIV[1]/DIV[1]/DIV[1]/DIV[1]/SPAN[2]").GetROProperty("innerhtml")

	Booked_Install_Date = cdate(DateValue(Selected_Install_Date &"-" & Selected_Install_Month &"-" & Selected_Install_Year))
	
	'Logic To Exclude Saturday and Sunday between First Available Scope and Install Date
	Temp = 0
	While (Date1 <= Booked_Install_Date)
	
		Day2=Weekday(Date1)
		If (Day2 = 1 or Day2 = 7) Then
			Temp = Temp + 1
		End If
		Date1=DateAdd("d",1,Date1)
	Wend

	If ((DateDiff("d",Booked_Scope_Date,Booked_Install_Date) - Temp) > 3) then
		Reporter.ReportEvent micWarning, "Scope And Install Appointment are more than 3 days Apart","Check if this is due to Slot unavailability"
		
	ElseIf ((DateDiff("d",Booked_Scope_Date,Booked_Install_Date) - Temp) = 3) Then
		Reporter.ReportEvent micPass, "Scope And Install Appointment follows minimum SLA of 3 days","Scope And Install SLA Check Passed"
		
	Else
		Reporter.ReportEvent micFail, "Scope And Install Appointment does not follows minimum SLA of 3 days","Scope And Install SLA Check Failed"
	End if

Else

	Reporter.ReportEvent micFail, "Install Slots not Available","Please updload AN Slots for this Region"
	
End If

End Function

'@Description   : Function to Compare Default CSE from DB and Actual Default CSE in Portal
'@Documentation : Function to Compare Default CSE from DB and Actual Default CSE in Portal
Function ValidateDefaultCSE(CustomerName,sAIM,ProductOffer)

Dim strSQLQuery,CustomerReferenceQuery,CustomerReference,DefaultCSE,DefaultSelectedCSE
strSQLQuery = getCSEQuery()


'Replace Required Customer Reference
CustomerReferenceQuery = getCustomerReference(CustomerName)
CustomerReference = DB_COM_FetchData(CustomerReferenceQuery)
strSQLQuery = Replace(strSQLQuery,"54765460",CustomerReference)


'Replace Aim 
strSQLQuery = Replace(strSQLQuery,"sAIM",sAIM)

'Fetch Default CSE from DB
DefaultCSE = trim(DB_COM_FetchData(strSQLQuery))

'Fetch Selected CSE from the Portal
DefaultSelectedCSE =  trim(Browser("name:=Chorus Self Service Portal").Page("Title:=Chorus Self Service Portal").Weblist("html id:=cseOffer","html tag:=SELECT").GetROProperty("selection"))


If DefaultCSE = "" Then

	If (instr(1,DefaultSelectedCSE,"No CSE required",1) > 0)  Then
	
		Reporter.ReportEvent micPass, "Defautl CSE","No Default CSE Associated..Hence No CSE Required is Displayed as default Selection"
		
	Else
		
		Reporter.ReportEvent micFail, "Defautl CSE","No Default CSE Associated.. But No CSE Required is not Displayed as default Selection"
		
	End If
	
Else

	If (instr(1,DefaultSelectedCSE,DefaultCSE,1) > 0)  Then
	
		Reporter.ReportEvent micPass, "Defautl CSE","Default CSE Selected"
		
	Else
		
		Reporter.ReportEvent micFail, "Defautl CSE","Default CSE Not Selected"
		
	End If

End If

End Function
