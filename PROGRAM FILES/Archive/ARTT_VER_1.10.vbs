Rem ##########################################################################################################################################################
Rem Script/Tool Name : ARTT
Rem Version - 1.9 (This version has the implementation of REFER~TC functionality implemented and CSID validations (MHS, CE ACTION).
Rem Previous Version - 1.2 (This version has the implementation of CE real time run).
Rem Version Creation Date: 10/20/2018
Rem Version Reviewer: Mohammad Sarwar
Rem ##########################################################################################################################################################


Dim tdExcelFileLoc, tdNotePadLoc, notePadLogFolder,createTDSummaryLog,tcSummaryLogFolder,envInfoExcelPath,envInfoExcelSheetName
Dim defaultAge,defaultGender,defaultSSN,defaultADD1,defaultCITY,defaultSTATE,defaultZIP,defaultMEMBERTYPE,defaultDEPTYPE
Dim dbConnGbl,currUserGbl,logFileDirGbl,dbConnStrGbl,runProductGbl,systemNameGbl,runCEGbl,wsUrlGbl,excelLogDirGbl,authUserGbl,memberInfoType,exitArtt

Rem fileLoc is the location of the config file for this tool to run. This location is hard coded.
Rem ==========================================================================================================================================================
fileLoc = "C:\ARTT\PROGRAM FILES\tdConfig.cfg"
Call load_config_file (fileLoc) 'Call function to load all the config value from the config file ('tdConfig.cfg') to execute the following scripts.
Call create_excel_application_object (objXl)'Call function to create an instance of excel before the driver scripts is run.
loadTdExcelFile = load_specified_excel (tdExcelFileLoc,objXl,1,bookXl,objXlSheet, True) 'Call function to load the excel sheet (driver script).
Set tdExlSheet = objXlSheet
tdExcelRows = tdExlSheet.UsedRange.Rows.Count
tdExcelCols = tdExlSheet.UsedRange.Columns.Count
executeCounter = 0 'set the number of rows to be executed to 0

'Create TD Summary Log if the flag (createTDSummaryLog) is set to 'TRUE in the config file.
If CBool(createTDSummaryLog) = True Then 'C.0
	tdSummaryFileName = "TD_SUMMARY_"&get_time_date_stamp()
	Call create_text_file (notePadLogFolder,tdSummaryLogFolder,tdSummaryFileName)
End If 'C.0
executeCounter = 0
'The following 'For-loop' is used to propogate through the testDriver which the following columns.
' SEQUENCE	EXECUTE	SHEET_ID_NAME	RULE_CATEGORY	TEST_ENV	SUPPLIER_ID	RUN_PRODUCT	RUN_CE_RT	MEMBERSET	RESULTS
' Any row has values ("Y"/"YES"/"y"/"yes") will be run with the related information

Rem ==========================================================================================================================================================
For n= 1 To tdExcelRows	
	curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"EXECUTE")
	executeFlag = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId) 	
	If UCase(executeFlag) = "YES" Or UCase(executeFlag) = "Y" Then 'C.1 - If any row has value "YES" for the 'EXECUTE', then read the remaining column values
		executeCounter = executeCounter +1
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"SHEET_ID_NAME") ' call function to get the column id for this given column name (SHEET_ID_NAME)
		ruleID = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId) 'call function to get the column value in the row that is set to execute = YES
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"RULE_CATEGORY")
		ruleCat = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"TEST_ENV")
		testEnv = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"SUPPLIER_ID")
		supplierID = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"RUN_PRODUCT")
		runProduct = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		Call get_operational_product_and_system (runProduct,runProductGbl, systemNameGbl)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"RUN_CE_RT")
		runCERealTime = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		If (CBool (runCERealTime) = True) Or (IsEmpty (runCERealTime) = True) Then 'C.1.a - If flag is set to TRUE/Empty to run CE real time for all test cases for a given rule.
			runCEGbl = True
			Else
			runCEGbl = False
		End If 'C.1.a
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"MEMBERSET")
		memberSet = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		'Translate values in this column (MEMBERSET)
		If InStr(memberSet,"MEMBER")>0  Then 'C.a.1 - If the column has this string ("MEMBER") Rem Previous condition - And IsEmpty(memberSet) = False
			Call get_number_of_membersets (memberSet,memberSetArr,useTdMember) 	
		End If 'C.a.1
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"TC_FILE_LOC")
		tcFileLoc = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"TC_RANGE")
		tcExecuteRange = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"RESULTS")
		executeResults = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
		curColId = get_column_id_from_excel_matching_a_column_name (tdExlSheet,1,"AUTH_USER_ID")
		authUserGbl = get_cell_value_given_rowid_columnid (tdExlSheet,n,curColId)
			
		If executeCounter >=1 And (UCase(executeFlag) = "YES" Or UCase(executeFlag) = "Y")Then 'C.1.0
			'Log each entry with 'YES' to the TDSummaryLog if createTDSummaryLog = True
			If CBool(createTDSummaryLog) = True Then 'C.1.1
				currAppendText = "# "&executeCounter&") "& ruleCat&Space(1)&"for RULE ID:"&ruleID&Space(1)&"is set to (YES) to be executed by ARTT."&VbCrlf
				Call append_text_to_notepad_file (notePadLogFolder&tdSummaryLogFolder,tdSummaryFileName&".txt",currAppendText)'Call function to log in the notepad log
				'Call function to create a summary log in excel format.
			End If 'C.1.1
			currFileName = "TC_LOG_RULE_ID_"&ruleID&"_"&ruleCat&"_"&"_"&Replace(Date,"/","_")&Space(1)& Replace(Time,":","_")
			Call create_text_file (notePadLogFolder,tcSummaryLogFolder,currFileName)
			notePadLogDir = notePadLogFolder&tcSummaryLogFolder&currFileName&".txt"
			logFileDirGbl = notePadLogDir
			currXlOutputFile = ruleID&"_"&ruleCat&"_"&"_"&Replace(Date,"/","_")&Space(1)& Replace(Time,":","_")&".xlsx"
			excelLogDir = excelOutputFile&currXlOutputFile 'to be worked on later	
			excelLogDirGbl = excelOutputFile 'Set the excel log directory to this global variable (excelLogDirGbl).	
			
			Rem : Section-A1 : This section is used for creating the excel output for each rule-id (the single entry from the TD/Driver Script
			If CBool(createExcelOutput) = True Then 'C.b.1 - If the config file flag to create excel output file for TCs is set to 'True'
				Call create_excel_output_file_for_rule_id (excelOutputFile,currXlOutputFile,tcExcelOutputType)
			End If 'C.b.1
			
			Rem Section-A2: This section is to create the excel output Sheet if the flag is set in config file.			
			If CBool(createMemberInfoExcel) = True Then 'C.c1 - If the flag is set in config file to 'True, then create the member info output excel
				Call create_member_info_excel (ruleID,ruleCat,notePadLogFolder,tcSummaryLogFolder,excelLogDirGbl,memberInfoExcelGbl,memberInfoLogFileGbl)
			End If 'C.c1
			
			'Call function to get database info for the given environment
			dbInfo = get_db_connection_information (testEnv,envInfoExcelPath,dbHost,dbSid,dbPort,dbUser,dbPassword,webServiceEndpoint)
			wsUrlGbl = webServiceEndpoint 'set the web service url to the global variable (wsUrlGbl)
			dbConnectSuccess = connect_to_a_database_for_a_given_env (dbInfo,dbHost,dbSid,dbPort,dbUser,dbPassword,dbConnGbl,currUserGbl)'Call function to establish database connection	
						
			Rem : If exitArtt is set to True
			If exitArtt <> True Then 'C.e - If exitArtt is not set to True
			If useTdMember = True Then 'C.a - If memberSet is specified in TD.
				'load the membersets from notepad file
				Call load_config_file (memberSetsLoc)
				numOfMemSets = UBound(memberSetArr)+1
				For d=1 To numOfMemSets
					currVarName = memberSetArr(d-1) '"MEMBERSET"&d
					currMemSet = Eval(currVarName)
					If InStr(currMemSet,",") Then 'C.b - If the current memberset has more than 1 member (separated by delimeter)
						currMemSetArr = Split(currMemSet,",")
						appendText = "ARTT will execute TCs (range):"&tcExecuteRange&" with all members in "&currVarName&" which has memberIDs:"&currMemSet
						Call append_text_to_notepad_file (notePadLogDir, "",appendText)
						totalMemInCurrSet = UBound (currMemSetArr)
						memSetHasMembers = True
						createTDMember = False
						Else												
						totalMemInCurrSet = currMemSet-1 '-1 because the total would 0 to (n-1), ie, 200, would be 0-199, where n is the number in MEMBERSET file.
						appendText = "ARTT will execute TCs (range):"&tcExecuteRange&" with "&currVarName&" needing "&totalMemInCurrSet+1&" new members which is created by ARTT."
						Call append_text_to_notepad_file (notePadLogDir,"",appendText)		
						Call create_members_for_memberset (currMemSetArr,totalMemInCurrSet,supplierID,testEnv,memberInfoLogFileGbl)
						memSetStr = array_elements_to_string_conversion (currMemSetArr,",")
						appendTxt = currVarName&" = "&Chr(34)& memSetStr&Chr(34)
						Call append_text_to_notepad_file_without_borders_timestamp (memberSetsLoc,"",appendTxt)	'Log the new members to the memberset notepad file				
					End If 'C.b
					
					For c=0 To totalMemInCurrSet 'totalMemInCurrSet = totalMemInCurrSet+1
						memberFromSet = Trim(currMemSetArr(c))
						myVal = execute_test_cases_for_a_given_rule_id (ruleCat,ruleID,testEnv,tcFileLoc,tcExecuteRange,memberFromSet,supplierID,notePadLogDir,excelLogDirGbl,currXlOutputFile)				
					Next
					c=0'Reset the value of c
				Next 
				Else 'C.a - If member is not in memberSet
				allTcsExecuted = execute_test_cases_for_a_given_rule_id (ruleCat,ruleID,testEnv,tcFileLoc,tcExecuteRange,memberFromSet,supplierID,notePadLogDir,excelLogDirGbl,currXlOutputFile)				
			End If 'C.a
			
			If allTcsExecuted = False Then 'C.d
				appendTxt = "All TCs for Rule-ID"&ruleID&" for "&ruleCat&" was not successful."
			Else
			End If 'C.d
		End If 'C.1.0
		curColId = 0
		previousYesRow = executeYesRow
		End If 'C.e
	End If 'C.1
Next

'exitArtt = True
'MsgBox IsEmpty(exitArtt)= False &"-"&exitArtt
If (exitArtt = False) Or (IsEmpty(exitArtt)= False) Then 'C.z1 - If the flag was set to True to exit ARTT.
	If dbConnGbl.State = 1 Then 'C.a1 - If the data base connection is still open, close the connection.
		dbConnGbl.Close
	End If 'C.a1
	
	objXl.Quit'Close the opened excel book
	On Error Resume Next
	If Err.Number = 0 Then 'C.b1 - If error occurred with Quit/Closing excel
		bookXl.Close 
	End If 'C.b1
	Set bookXl = Nothing
End If 'C.z1
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Rem Fuction name: execute_test_cases_for_a_given_rule_id (ruleCat,ruleId,testEnv,tcFileLoc,tcRange,memberFromTd,tcDetailedLog,tcExcelOutput)
Rem Fuction Arguments: ruleCat (type of rules, ie:CONDVAL,MKVAL),ruleId (sheet name or ID in the excel),testEnv (the environment that TCs are intended to be executed
Rem ,tcFileLoc (location of the excel file containing the TCs),tcRange (range of the test cases passed from the driver script,memberFromTd (memberID passed from the driver script, if any)
Rem ,tcDetailedLog (Notepad log directory),tcExcelOutput (exceloutput log directory),memberExistsTD (True = Create a member, False = don't create a member)
Rem Fuction tasks: This is the main function of ARTT which executes a all test cases for a given ruleID.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_test_cases_for_a_given_rule_id (ruleCat,ruleId,testEnv,tcFileLoc,tcRange,memberFromTd,supplierID,tcDetailedLog,tcExcelOutputDir,tcExcelOutputSheetName)
	If exitArtt = True Then 'C.0- If the driver script an error and exitArtt is set to True
		Exit Function 
		Else
		Rem =================BEGINNING of SECTION-A - it is used to identify which sets ot TC should be run as specified in the driver script column (TC_RANGE)=======================
		Call find_test_case_range (tcRange,tcRangeArr,tcLowerLim,tcUpperLim,randomTcSelection,randSelectSingleTC)
	
		If useTdMember <> True Then 'C.e - If member is not coming for TD (False will execute else, then don't update tcUpperLim
			currAppendTxt = "ARTT WILL EXECUTE - "&vbTab&"TEST CASES ("&tcRange&") in ENVIRONMENT ("&testEnv&") with RULE_CATEGORY ("&ruleCat&") for RULE_ID ("&ruleId&")."
			Call append_text_to_notepad_file (tcDetailedLog,"",currAppendTxt)
		End If 'C.e		
			Rem ================END of SECTION-A===============================================================================================================
			Rem ================BEGINNING of SECTION-B - this scetion reads the excel with the test cases for a given rule-id and also creates an excel output file.
			loadTCExcel = load_specified_excel_by_sheet_name (tcFileLoc,objXl,ruleId,bookXl,objTCXlSheet, False)'call fucntion to load excel with TCs	
			tcExcelRows = objTCXlSheet.UsedRange.Rows.Count
			If useTdMember = True Then 'C.0 - If member is not coming for TD (False will execute else, then don't update tcUpperLim
				tcExecuteFlag = "YES"
				Else
				If randSelectSingleTC = True Then 'C.0.2
					tcUpperLim = tcLowerLim
					Else 'C.0.2
					If CInt(tcExcelRows) >= CInt(tcUpperLim) Then 'C.0.1 - If the last excel rows (= # of total test cases) in that excel is not bigger or equal to the highest range provided in driver script for TC range.
						tcUpperLim = CInt(tcExcelRows)
						Else
						errMsg =  "TC-ID#"&tcUpperLim&" is bigger than the last row ("&tcExcelRows&") in excel, hence either TC excel does not have enough rows (each row = one test case) or TC-ID provided in the driver/controller is wrong.ARTT will abort."
						MsgBox errMsg
						Call append_text_to_notepad_file (tcDetailedLog,"",errMsg)
						execute_test_cases_for_a_given_rule_id = False
						Exit Function
					End If 'C.0.1
				End If 'C.0.2
			End If 'C.0
			
			tcCounterForCurrRuleID = 0
			tcExecuteCounterForCurrRuleID = 0 
			
			Rem EXECUTE	TCID	RELATED_TCID	MEMBERID	MEMBER_DEMOGRAPHICS	TC_DESCRIPTION	EVALUATION	TC_EVENTS	DML	SAVEDATAFORMONTHS
			
			Rem ================BEGINNING of SECTION-C - this section is looping through an excel to execute all the TC from that excel (for a given rule)
			For k = tcLowerLim+1 To tcUpperLim+1 'Loop-A.1 - Loop runs from 1 to tcExcelRows if TC range is not specified in DriverScript, else runs from the lowere range of TC range.
				tcCounterForCurrRuleID = tcCounterForCurrRuleID+1
				If randomTcSelection = True And randSelectSingleTC <> True Then 'C.a				
					currTCFlagged = verify_number_exist_in_container (tcRangeArr,k-1) 'Passing k-1 because the first row value of k, is the header column and ignored by setting k to begin with tcLowerLim+1
					If currTCFlagged = True Then
						randSeclecMultiTC = True 'If there are multiple TC from TD.
						tcExecuteFlag = "YES"
						Else
						tcExecuteFlag = "NO"
					End If
					ElseIf randSelectSingleTC = True Then 'C.a' Then
						currTCFlagged = True
					Else 'C.a''
					curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"EXECUTE")
					tcExecuteFlag = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
				End If 'C.a
				
				If useTdMember = True Or randSelectSingleTC = True Then 'C.b-If member is passed from Test Driver
					If randSeclecMultiTC <> True Then 'C.c - 
						tcExecuteFlag = "YES"
						Else 'C.c
					End If 
					tcExecuteCounterForCurrRuleID = tcExecuteCounterForCurrRuleID+1
				End If 'C.b
				
	 			If UCase(tcExecuteFlag) = "YES" Or UCase(tcExecuteFlag) = "Y" Then 'C.3 - If the execution for the TC is set to "YES"/"Y"
	 				tcExecuteCounterForCurrRuleID = tcExecuteCounterForCurrRuleID+1
					'Get all related columns for this TC
					curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"TCID")
					tcId = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
					curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"RELATED_TCID")
					relatedTcId = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)					
					
					logMsg =VbTab&VbTab&VbTab&VBtab&"BEGINNING of TEST CASE: "&tcId 'This will be logged in log file
					If relatedTcId <> Empty Or relatedTcId <> "" Then 'C.3.1 Then
						currAppendTxt = logMsg & " which is related to "&relatedTcId
						Else
						currAppendTxt = logMsg
					End If '3.1
					logMsgFinal = currAppendTxt&", with MemberID ("&memberFromTd&")."
					Call log_header_footer (logMsgFinal,tcDetailedLog,"=","=",140)
					currAppendTxt = VbTab&VbTab&VbTab&"TC DESCRIPTION (Copied from Input Excel for this test case)"
					Call append_text_to_notepad_file_without_borders_timestamp (tcDetailedLog,"",currAppendTxt)
					currAppendTxt = VbTab&VbTab&VbTab&create_a_line_of_repeated_characters ("-",60)
					Call append_text_to_notepad_file_without_borders_timestamp (tcDetailedLog,"",currAppendTxt)
					curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"TC_DESCRIPTION")				
					tcDesc = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
					If tcDesc = Empty Or tcDesc = "" Then 'C.3.2
						tcDesc = "No description is given in the excel for this Test Case."
					End If 'C.3.2
					Call append_text_to_notepad_file_without_borders_timestamp (tcDetailedLog,"",tcDesc)
					Rem =======================SECTION-C.1 - member creation if no existing member in TC=========================================			
					If	useTdMember = True  Then 'Or memberSetHasMember = False Then 'C.1.a	- Meaning the member should be used from TC file
						currTcMemberId = memberFromTd
						Else 'C.1.a'
						curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"MEMBERID")
						currTcMemberId = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)				
						curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"MEMBER_DEMOGRAPHICS")
						currMemberDemo = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)	
					End If 'C.1.a			

					If IsEmpty(currTcMemberId) Then 'C.1.c - If the memberID column is empty then a member with details in 'Member_Demographics' will be used to create the member.
						memberIDCreated = create_a_member_for_tc (currMemberDemo,supplierID,testEnv,tcDetailedLog,aggMemberID)
						newMemberFlag = True 'Set this to be used for TDM log 
						memberIDToUse = memberIDCreated
						'Get members (both Regular and Aggregate) to add to TDM tracker table						
						memberLogMsg = "newly created member "
						ElseIf IsEmpty (currTcMemberId) = False Then 
						memberIDToUse = currTcMemberId
						memberLogMsg = "existing Member, provided in TC/TD"
					End If 'C.1.c

					currAppendTxt = "This TC will be executed using "&memberLogMsg&", (ID:"&memberIDToUse&") for supplier (ID:"&supplierID&") in ("&testEnv&") environment."
					Call append_text_to_notepad_file (tcDetailedLog,"",currAppendTxt)
					ReDim tcMemInfoArr (1,9)
					Call get_member_info_from_database (memberIDToUse,dbConnGbl,tcMemInfoArr)
					memberFetchedFromDB = Trim(tcMemInfoArr(1,1))
					If Trim(memberIDToUse) = memberFetchedFromDB Then 'C.3.3.1 - If the member provided in TC does not exist in database.
						Call print_member_information_to_the_log (currLogFile,"",tcMemInfoArr,"REGULAR") 'Captures the member information in the log file
						Rem The following logs member info in the memberInfo excel
						If createMemberInfoExcel = True Then 'C.c1 - If the member info output file is created
							fileAlreadyExists = verify_if_file_exist (False,excelLogDirGbl,memberInfoExcelGbl,"")
							If fileAlreadyExists = True Then 'C.c2
								memberInfoOutputSheetLoc = excelLogDirGbl&memberInfoExcelGbl
								loadMemInfoExcelFile =  load_specified_excel_by_sheet_name  (memberInfoOutputSheetLoc,objXl,"MEMBER_INFO",bookXl,memberInfoOutputSheet, False)
																	  '(excelLogDirGbl&tcExcelOutputSheetName,objXl,1,bookXl,objXlSheet, False)
								'Set memberInfoOutputSheet = objXlSheet
								rowTot = memberInfoOutputSheet.UsedRange.Rows.Count
								If UCase(memberInfoType) = "LONG" Then 'C.c3
									strExcelInfo = rowTot+1&",1;"&tcMemInfoArr(1,0)&"|"&rowTot+1&",2;"&tcMemInfoArr(1,1)&"|"&rowTot+1&",3;"&tcMemInfoArr(1,2)&"|"&rowTot+1&",4;"&tcMemInfoArr(1,3)&"|"&rowTot+1&",5;"&tcMemInfoArr(1,4)&"|"&rowTot+1&",6;"&tcMemInfoArr(1,5)&"|"&rowTot+1&",7;"&tcMemInfoArr(1,6)&"|"&rowTot+1&",8;"&tcMemInfoArr(1,7)&"|"&rowTot+1&",9;"&tcMemInfoArr(1,8)&"|"&rowTot+1&",10;"&tcMemInfoArr(1,9)&"|"&rowTot+1&",11;"&tcId
									numOfCols = 11
									ElseIf UCase(memberInfoType) = "SHORT" Then 'C.c3
									strExcelInfo = rowTot+1&",1;"&tcMemInfoArr(1,0)&"|"&rowTot+1&",2;"&tcMemInfoArr(1,1)&"|"&rowTot+1&",3;"&tcMemInfoArr(1,2)&"|"&rowTot+1&",4;"&tcId
									numOfCols = 4
								End If 'C.c3
								'Call function to write to the excel sheet.
								Call write_to_excel_output_log (memberInfoOutputSheet,strExcelInfo,"|",";",numOfCols)
								bookXl.Save
								bookXl.Close
							End If 	'C.c2
						End If 'C.c1
						Else
						appendTxt = "Member ("&memberIDToUse&") does NOT exist in Database, hence TC cannot be executed. ARTT is aborted"
						Call append_text_to_notepad_file (tcDetailedLog,"",appendTxt)
						functionRetMsg = appendTxt
						currAppendTxt = VbTab&VbTab&VbTab&VBtab&"END of TEST CASE: "&tcId & ", MEMBERID ("&memberIDToUse&")."
						Call log_header_footer (currAppendTxt,tcDetailedLog,"=","=",140)
						Exit Function
					End If 'C.3.3.1 
	'					memberIDToUse = currTcMemberId
	'				End If 'C.3.3
					Rem ==============END of SECTION-C.1==============================================================================================				
					Rem ==============END of SECTION-C================================================================================================				
					Rem ==============BEGINNING of SECTION-D - this section is to collect all the events requirements for TC (as provided in columns (TC_EVENTS and DML)
					Rem =================================================================================================================================================
					Rem Collect the TC_EVENTS from a related TC if any.
					If IsEmpty (relatedTcId) = False Then 'C.3.e - If the related_tcid has a value 
						relatedTCEvents = get_a_value_from_excel_column_matching_a_key (tcFileLoc,2,relatedTcId,objTCXlSheet,"TC_EVENTS")
						ReDim relatedTCEventsDMLArr (50)
						Call collect_and_translate_test_case_events_into_dmls (relatedTCEvents,relatedTCEventsArr,relatedTCEventsDMLArr,memberIDToUse)	
						relatedTCDmls = get_a_value_from_excel_column_matching_a_key (tcFileLoc,2,relatedTcId,objTCXlSheet,"DML")
						appendTxt = "/* 'TC_EVENTS' used from the related test case ("&relatedTcId&") logged below. */"
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						Call execute_dml_from_an_array_of_dmls (relatedTCEventsArr,relatedTCEventsDMLArr,False)'True = DMLs created from TC_EVENTS
						If IsEmpty(relatedTCDmls) = False Then 'C.3.f - If there's DML provided in the DML column of TC sheet
							appendTxt = "/* DML(s) used from the related test case ("&relatedTcId&") logged below. */"
							Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
							'The DML column has more than 1 DML, delimeted by ";".
							If InStr(relatedTCDmls,";")>=1 Then 'C.3.g - If more than 1 DML
								relatedTCDmls = Replace(relatedTCDmls,"GMemberid",memberIDToUse)
								reledTcDMLArr = Split(relatedTCDmls,";")
								Else
								ReDim reledTcDMLArr (0)
								reledTcDMLArr (0) = Replace(relatedTCDmls,"GMemberid",memberIDToUse)
							End If 'C.3.g
							Call execute_dml_from_an_array_of_dmls ("",reledTcDMLArr,True)'False = DMLs copied from the DML column
						End If 'C.3.f
					End If 'C.3.e
					Rem Collect the TC_EVENTS from the current TC
					curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"TC_EVENTS")
					currTcEvents = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
					If InStr(currTcEvents,"NOEVENTS")>0 Or InStr(currTcEvents,"NOEVENT")>0 Or InStr(currTcEvents,"NONE")>0 Then 'C.3.a - If the TC_EVENTS column has no events specified
						appendTxt = "There's no event specified in the TC_EVENTS column, hence no data will be seeded in ODS for this TC."
						Call append_text_to_notepad_file (logFileDirGbl, "",appendTxt)
						ElseIf InStr(UCase(currTcEvents),"REF")>0  Then 'C.3.5' - If TC has reference to other sheet for TC Events.
						appendTxt = "The current TC is referred to and external sheet within TC file"
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						Call collect_tc_events_from_referred_sheet (currTcEvents,tcFileLoc,objXl,memberIDToUse) 						
						ElseIf InStr(UCase(currTcEvents),"#")>0 And InStr(UCase(currTcEvents),"REF")=0 Then 'C.3.5 - If the TC_EVENTS column has no reference to other TC (using REFER keyword).
						'Declare an array for DMLs 
						ReDim tcEventsDMLArr (50)
						dmlCounter = 0
						appendTxt = "/* 'TC_EVENTS' used from the current test case ("&tcId&") logged below. */"
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						Call collect_and_translate_test_case_events_into_dmls (currTcEvents,tcEventsArr,tcEventsDMLArr,memberIDToUse)	
						Call execute_dml_from_an_array_of_dmls (tcEventsArr,tcEventsDMLArr,False)			
					End If 'C.3.a
					'Collect additional DMLs from the DML input column
					curColId = 9 'get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"DML")
					currTCDml = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
					If IsEmpty(currTCDml) = False Then 'C.3.4 - If there's DML provided in the DML column of TC sheet
						'The DML column has more than 1 DML, delimeted by ";".
						If InStr(currTCDml,";")>=1 Then 'C.3.4.1
							currTCDML = Replace(currTCDml,"GMemberid",memberIDToUse)
							tcAdditionalDMLArr = Split(currTCDml,";")
							Else
							Dim tcAdditionalDMLArr (0)
							tcAdditionalDMLArr (0) = Replace(currTCDml,"GMemberid",memberIDToUse)
						End If 'C.3.4.1
						appendTxt = "/* DML(s) used from the current test case ("&tcId&") logged below. */"
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						Call execute_dml_from_an_array_of_dmls ("",tcAdditionalDMLArr,True)
					End If 'C.3.4
					Rem ==============END of SECTION-D ================================================================================================
					
					Rem ==============BEGINNING of SECTION-E - this section is for running CE real time ===============================================			
					If runCEGbl = True Then 'C.3.5- The CE real time service will be invoked if the flag is set to TRUE/Empty in Driver script.					
						ceRTrun = invoke_ce_realtime_service (memberIDToUse,supplierID,runProductGbl,systemNameGbl,startTimeWS,endTimeWS)
						If CBool (ceRTrun) = True Then 'C.3.5.a
							continueCSIDVal = True
							appendTxt = "Member ("&memberIDToUse&") was run successfully via CE REAL TIME web service. The service was initiated @ "&startTimeWS&" (-2 minutes), ARTT will now validate CSID tables for "&_
							"Test Case EVALUATION."
							commentGbl = appendTxt
							Else
							appendTxt = "Member ("&memberIDToUse&") was NOT run successfully via CE REAL TIME web service. Hence ARTT will NOT validate CSID tables for "&_
							"Test Case EVALUATION."
							commentGbl = appendTxt
							'Update Excel output file with 'FAILED' message. - to worked on later.
						End If 'C.3.5.a
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						Else 'C.3.5
						appendTxt = "MEMBER is not set to RUN real time as the flag (RUN_CE) in driver script is set to 'False'."
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						currAppendTxt = VbTab&VbTab&VbTab&VBtab&"END of TEST CASE: "&tcId & ", MEMBERID ("&memberIDToUse&")."
						Call log_header_footer (currAppendTxt,tcDetailedLog,"=","=",140)
						abortArtt = True 'Set this variable to True if ARTT was aborted.
						commentGbl = appendTxt
					End If 'C.3.5
					Rem ==============END of SECTION-E ================================================================================================				
					
					Rem ==============BEGINNING of SECTION-F - this section is database validations of the TC==========================================
					If continueCSIDVal = True Then 'C.f.1 - If CE was run and CSID validtion is expected.
						curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"EVALUATION")
						currTcEval = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
						currTcEvalArr = Split(currTcEval,"$")
						evalRuleType = UCase(Trim(currTcEvalArr(0)))
						mfID = Trim(currTcEvalArr(1))
						expectedDF = get_df_information_for_medical_finding (mfID, mfType)
						mrrID = get_member_recommend_runid_from_mrr_table (memberIDToUse,startTimeWS,endTimeWS)
						actualDF = retrieve_all_derived_fact_ids_for_the_current_run (mrrID,True,expectedDF,"") 'Call function to fetch the expected derived Fact fired or not.
						If IsEmpty(expectedDF) = False And IsEmpty(actualDF) = False And UCase(mfType) = UCase(evalRuleType) And actualDF Then 'C.d
							appendTxt = "Since he expected DF-ID ("&expectedDF&") for MF-ID ("&mfID&"), of TYPE - "&mfType&" is evaluated SUCCESSFULLY , additional validations may follow."
							commentGbl = appendTxt
							tcStatus = rule_category_csid_validation (ruleCat,memberIDToUse,mrrID,expectedDF,runProductGbl)
'							tcStatus = abc (ruleCategory,memberId,memberRunId,stateComponentId,productCode)
							Else
							ReDim dfArr (1000)
							Call retrieve_all_derived_fact_ids_for_the_current_run (mrrID,False,expectedDF,dfArr)
							appendTxt = "TC FAILURE REASON - The expected DF-ID ("&expectedDF&") for MF-ID ("&mfID&"), of TYPE - "&mfType&" is evaluated and the evaluation is NEGATIVE, "&_
							" the derived facts that actually triggered are followed --> ("
							allFiredDFs = ""
							For b=0 To UBound(dfArr)
								allFiredDFs = allFiredDFs&dfArr(b)&","
							Next
							appendTxt = appendTxt&get_rid_off_chars (allFiredDFs,"Left",1,1)&")."
							tcStatus = "FAILED"
						End If 'C.d
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
					End If 'C.f.1			
					
					Rem ==============BEGINNING of SECTION-G - this section to purge data if not required for future usage===============================				
					If CBool(tdmTrackerEnabled) = True And newMemberFlag = True Then 'C.g1	-If the flag (tdmTrackerEnabled) in config is set to true and a member was created new (newMemberFlag)		
						curColId = get_column_id_from_excel_matching_a_column_name (objTCXlSheet,1,"PURGE_DATA")
						saveDataMonths = get_cell_value_given_rowid_columnid (objTCXlSheet,k,curColId)
						Call insert_members_into_tdm_tracker_table (memberIDToUse,aggMemberID,saveDataMonths)
					End If 'C.g1
					Rem ==============END of SECTION-G - this section to purge data if not required for future usage=====================================
					Rem ==============BEGINNING of SECTION-H - this section to log info to the excel output log ===============================	
					fileAlreadyExists = verify_if_file_exist (False,excelLogDirGbl,tcExcelOutputSheetName,"")
					If fileAlreadyExists = True Then 'C.h.1
						loadTdExcelFile = load_specified_excel (excelLogDirGbl&tcExcelOutputSheetName,objXl,1,bookXl,objXlSheet, False)
						Set tcExcelOutputSheet = objXlSheet
						rowTotXlTc = tcExcelOutputSheet.UsedRange.Rows.Count
						strExcelInfo = rowTotXlTc+1&",1;"&tcID&"|"&rowTotXlTc+1&",2;"&MemberIDToUse&"|"&rowTotXlTc+1&",3;"&currTcEval&"|"&rowTotXlTc+1&",4;"&tcStatus&"|"&rowTotXlTc+1&",5;"&commentGbl
						'Call function to write to the excel sheet.
						Call write_to_excel_output_log (tcExcelOutputSheet,strExcelInfo,"|",";",5)
	'					tcStatus = "PASSED" 'To be removed.
						If UCase(tcStatus) = "PASS" Or UCase(tcStatus) = "PASSED" Then 'C.h.2-Choose the color coding for output file in case of PASS/FAIL.
							cellColor = 4
							ElseIf UCase(tcStatus) = "FAIL" Or UCase(tcStatus) = "FAILED" Then
							cellColor = 3
							Else
							cellColor = 5
						End If 'C.h.2
						Call color_code_excel_cell (tcExcelOutputSheet,rowTotXlTc+1,1,cellColor,1)'
						Call color_code_excel_cell (tcExcelOutputSheet,rowTotXlTc+1,4,cellColor,1)'
						bookXl.Save
						bookXl.Close
					End If 	'C.h.1
				End If 'C.3	-If the row is set to "YES/Y" in the TC file 'EXECUTE Column.				
			Next 'Loop-A.1
		
		Set objTCXlSheet = Nothing
		
		If abortArtt <> True Then 'C.z - If ARTT was not halted earlier.
			currAppendTxt = VbTab&VbTab&VbTab&VBtab&"END of TEST CASE: "&tcId & ", MEMBERID ("&memberIDToUse&")."
			Call log_header_footer (currAppendTxt,tcDetailedLog,"=","=",140)
		End If 'C.z
	
		execute_test_cases_for_a_given_rule_id = functionRetMsg
	End If 'C.0
'	bookXl.Close
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Rem Fuction name: log_header_footer (textToAppend,logDir,headerChar,footerChar,numOfChar)
Rem Fuction Arguments: textToAppend (the text to be added to the file),logDir,headerChar (ie,'+',footerChar (ie,'=',numOfChar (120)
Rem Fuction tasks: Function creates appends text (passed in 'textToAppend') within a header/footer design to a log file
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function log_header_footer (textToAppend,logDir,headerChar,footerChar,numOfChar)
	Call append_text_to_notepad_file_without_borders_timestamp (logDir,"",create_a_line_of_repeated_characters (headerChar,numOfChar))
	Call append_text_to_notepad_file_without_borders_timestamp (logDir,"",textToAppend)
	Call append_text_to_notepad_file_without_borders_timestamp (logDir,"",create_a_line_of_repeated_characters (footerChar,numOfChar))
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_a_member_for_tc ()
Rem Fuction Arguments: memberDemo (this variable should be passed in with a format (AGE:18;GENDER:M;SSN:1112226666;TYPE:D),memberSupplier (supplier ID for member to be associated with)
Rem ,testEnv (the environment in which member will be created, ie: QA1),aggMember (the aggregate member id if created).
Rem Fuction tasks: Function reads the config file to be used in the caller script
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_a_member_for_tc (memberDemo,memberSupplier,testEnv,currLogFile, ByRef aggMember)
	If memberDemo <> Empty Or memberDemo <> "" Then 'C.1 - If the member demographics are given in the test case
		If InStr (memberDemo,"/")>0 Then 'C.1.1 - If only age/gender is given in TC, ie 19/F
			memberDemoArr = Split (memberDemo,"/")
			memberAge = memberDemoArr (0)
			memberGender = memberDemoArr (1)
			ElseIf InStr (memberDemo,";")>0 Then 'C.1.2 - If additional member demographics are provided in TC in the form (AGE:19;GEN:F)
			memberDemoArr = Split (memberDemo,";")
			memberDemoSize = UBound(memberDemoArr)	
			For u = 0 To memberDemoSize
				currElement = memberDemoArr (u)
				currVarVal = member_demo_extractor (currElement,currVariable,":")
				Select Case UCase(currVariable)
					Case "AGE"
					memberAge = currVarVal
					memberDOB = create_a_date_of_birth_for_given_age (memberAge)
					Case "GENDER","GEN"
					memberGender = currVarVal					
					Case "DOB","DTOFBIRTH","DATEOFBIRTH"
					memberDOB = currVarVal
					Case "SSN"
					memberSSN = currVarVal
					Case "TYPE","MTYPE","MEMBERYPE","MBRTYPE","MEMTYPE"
					memberSSN = currVarVal
				End Select
			Next
		End If 'C.1.1
		Else 'Use the default values for the member demographics.
		memberAge = defaultAge
		memberGender = defaultGender
		memberSSN = defaultSSN
		memberType = defaultMEMBERTYPE
	End If 'C.1
		
	Rem If the database connection is created, then use the connection to create a member in Database
	If dbConnGbl.State = 1 Then 'C.4 - need to set to 1
	'If the TC is not provided with any of the following required demographic information, then use the ones as default from the config file
		If memberAge = Empty Then 'C.4.1
			If memberDOB = Empty Then 'C.4.1.1
				memberDOB = create_a_date_of_birth_for_given_age (defaultAge)
			End If 'C.4.1.1	
			Else 
			memberDOB = create_a_date_of_birth_for_given_age (memberAge)	
		End If 'C.4.1
		If memberGender = Empty Then 'C.4.2
			memberGender = defaultGender
		End If 'C.4.2
		If memberSSN = Empty Then 'C.4.3
			memberSSN = defaultSSN
		End If 'C.4.3
		If memberType = Empty Then 'C.4.4
			memberType = defaultMEMBERTYPE
		End If 'C.4.4
		
		Call create_member_personal_info (memberDOB,memberFirstName,memberLastName,memberMiddleInitial,memberFullName,memberEmailAddr)	'Call function to get personal details created in random
		
		dbConnGbl.BeginTrans

		'Call function to create a member in Database.
		tcMemberID = create_a_member_in_database (memberDOB,memberFirstName,memberLastName,memberMiddleInitial,memberFullName,memberEmailAddr,memberSupplier,memberGender,memberSSN,memberType,currUserGbl,dbConnGbl,currLogFile)
		'MsgBox IsEmpty (tcMemberID)&"-"&IsNull(tcMemberID) 
		If IsEmpty (tcMemberID) = False And IsNull(tcMemberID)= False Then 'C.4.5
'			Call get_member_info_from_database (createdMemberID,dbConnGbl,tcMemInfoArr)
'			Call print_member_information_to_the_log (currLogFile,"",tcMemInfoArr,"REGULAR") 'Captures the member information in the log file
			regMemCreated = True
		End If 'C.4.5
		If pvTurnedOn = True Then 'C.4.6 - If the PersonView flag is turned on (TRUE), then create person aggregation
			agMemberID = create_a_member_in_database (memberDOB,memberFirstName,memberLastName,memberMiddleInitial,memberFullName,memberEmailAddr,personAggSupplier,memberGender,memberSSN,memberType,currUserGbl,dbConnGbl,currLogFile)
			'MsgBox "Null="&IsNull(agMemberID) &"and Empty="& IsEmpty (agMemberID)
			If IsNull(agMemberID)= False And IsEmpty (agMemberID) = False Then 'C.4.6.1 - create the person aggregation in '' table by calling the following function.
				aggMemCreated = True
				'Call print_member_information_to_the_log (currLogFile,"",agMemInfoArr,"AGGREGATE") 'Captures the member information in the log file
			End If 'C.4.6.1
		End If 'C.4.6
		ReDim tcMemInfoArr (1,9)
		ReDim agMemInfoArr (1,9)
		If regMemCreated = True And CBool(pvTurnedOn) = False Then 'C.4.7 - Member data was successfully inserted in 12 ODS tables, then commit the transactions and print to the log.
			dbConnGbl.CommitTrans
			Call get_member_info_from_database (tcMemberID,dbConnGbl,tcMemInfoArr)
			Call print_member_information_to_the_log (currLogFile,"",tcMemInfoArr,"REGULAR") 'Captures the member information in the log file
			ElseIf regMemCreated = True And CBool(pvTurnedOn) = True And aggMemCreated = True Then
'			dbConnGbl.CommitTrans
			Call get_member_info_from_database (tcMemberID,dbConnGbl,tcMemInfoArr)
			Call get_member_info_from_database (agMemberID,dbConnGbl,agMemInfoArr)
			'Call function to create member aggregation in ODS.PERSONAGGREGATION
			memberAggregated = execute_person_aggregation_table_insert_dml (dbConnGbl,tcMemberID,agMemberID,currUserGbl)
			If memberAggregated = True Then 'C.4.7.1 Then
				dbConnGbl.CommitTrans
				aggMember = agMemberID 'aggregated member id is save to the ByRef variable 'aggMember'
				Else
				appendTxt = "PERSON AGGREGATION FOR MEMBERS, REGULAR MEMBER ("&tcMemInfoArr(1,1)&") AND AGGREGATE MEMBER ("&agMemInfoArr(1,1)&") FAILED."
				Call append_text_to_notepad_file (currLogFile,"",appendTxt)
			End If 'C.4.7.1
			'Call get_member_info_from_database (tcMemberID,dbConn,tcMemInfoArr)
			Call print_member_information_to_the_log (currLogFile,"",tcMemInfoArr,"REGULAR") 'Captures the member information in the log file
'			Call get_member_info_from_database (agMemberID,dbConn,agMemInfoArr)
			Call print_member_information_to_the_log (currLogFile,"",agMemInfoArr,"AGGREGATE") 'Captures the member information in the log file
		Else 'Rollback all the transactions and record failure messageprint to the log.
			appendTxt = "MEMBER CREATION was not successful. ARTT WILL ABORT NOW."
			Call append_text_to_notepad_file (currLogFile,"",appendTxt)
			dbConnGbl.RollbackTrans 
		End If 'C.4.7
	End If 'C.4

	create_a_member_for_tc = tcMemberID
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_a_member_in_database ()
Rem Fuction Arguments: memberDemo (this variable should be passed in with a format (AGE:18;GENDER:M;SSN:1112226666;TYPE:D),
Rem memberSupplier (supplier ID for member to be associated with)
Rem ,testEnv (the environment in which member will be created, ie: QA1)
Rem Fuction tasks: Function creates a member in data base based on the information passed in, and returns a memberID after creating member successfully,
Rem otherwise returns NULL.
Rem Created By: Mohammad Sarwar
Rem Creation Date: 09/01/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_a_member_in_database (memberDOB,memberFirstName,memberLastName,memberMiddleInitial,memberFullName,memberEmailAddr,currMemberSupplier,currMemberGender,currMemberSSN,currMemberType,currArttUser,dbConnGbl,currLogFile)
	currProviderID = verify_provider_is_valid_if_not_query_db_for_one (dbConnGbl,defaultProviderID) 'Call function to verify that the defaultProvider is valid, if not query database to get a valid provider id.
	Call get_supplier_accouunt_info (dbConnGbl,currMemberSupplier,supplierAccountName,supplierAccountID) 'Call function to get supplier Info

	partyIDSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ods.ODS_PARTY_SEQ") 'Get ODS.party table PK seq (PartyID). is used in Query#2.1
	partyAddrSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ods.ods_partyaddr_seq") 'Get ODS.partyaddr PK seq .
    memberIDSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ods.ods_member_seq")''Get ODS.member PK seq.
    memberPlanIDSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ods.ods_memberplan_seq")'Get ODS.member PK seq.
    memberPatIDSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ods.ODS_TEST") 'Get SOURCE PATIENT ID PK seq.
	memberPatID = "AUTOGEN_MEM"&memberPatIDSeq
    providerIDSeq = get_sequence_key_for_a_given_table (dbConnGbl,"ODS.ODS_MBRPROV_SEQ") 'Get ODS.MEMBERPROVIDER PK seq.
  	
  	If activateReporting = True Then 'C.a - If the flag (activateReporting) is set to add data to additional table, ods.memberreportinggroup
	  	dmlArrSize = 12
	  	Else
	  	dmlArrSize = 11
  	End If 'C.a
  	
  	ReDim dmlArr (dmlArrSize,1)
  	Dim dmlDtlArr (0,1)
  	
  	'Query#2.1 - Insert the data in  in ODS.PARTY table.  	
  	Call execute_party_table_insert_dml (dbConnGbl,partyIDSeq, dmlDtlArr) 'insert data into ODS.PARTY table
  	dmlArr(0,0) = dmlDtlArr (0,0)
  	dmlArr(0,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(0,0)&"-"&dmlArr(0,1))
	'Query#2.2 - Insert the data in  in ODS.PARTYADDRESS table.
	Call execute_party_address_table_insert_dml (dbConnGbl,partyAddrSeq,partyIDSeq,defaultADD1,defaultCITY,defaultSTATE,defaultZIP,currArttUser,dmlDtlArr) 'Query#2.1 - Call function to execute the DML (for Address) and return the error code if any error occurs. 
	dmlArr(1,0) = dmlDtlArr (0,0)
  	dmlArr(1,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(1,0)&"-"&dmlArr(1,1))
  	'Query#2.3 - Insert the data into ODS.PERSONXREF table. 
  	Call execute_personxref_table_insert_dml (dbConnGbl,partyIDSeq,currMemberSupplier,memberPatIDSeq,currArttUser,memberIDSeq,dmlDtlArr)
	dmlArr(2,0) = dmlDtlArr (0,0)
  	dmlArr(2,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(2,0)&"-"&dmlArr(2,1))
  	'Query#2.4 - Insert the data in to ODS.PERSON table.
  	Call execute_person_table_insert_dml (dbConnGbl,partyIDSeq,memberFirstName,memberMiddleInitial,memberLastName,memberFullName,currMemberGender,currMemberSSN,memberDOB,currArttUser,dmlDtlArr)
	dmlArr(3,0) = dmlDtlArr (0,0)
  	dmlArr(3,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(3,0)&"-"&dmlArr(3,1))
	'Query#2.5 - Insert the data in to ODS.PERSONFACT table.
	Call execute_person_fact_table_insert_dml (dbConnGbl,supplierAccountID,currMemberSupplier,memberIDSeq,memberPlanIDSeq,partyIDSeq,memberFirstName,memberLastName,memberDOB,memberGender,defaultCITY,supplierAccountName,memberFullName,dmlDtlArr)
	dmlArr(4,0) = dmlDtlArr (0,0)
  	dmlArr(4,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(4,0)&"-"&dmlArr(4,1))
	Rem Query#2.6 - Insert data into ODS.MEMBER table.
	Call execute_member_table_insert_dml (dbConnGbl,memberIDSeq,currMemberSupplier,memberPatID,partyIDSeq,currMemberType,memberDOB,currArttUser,memberPlanIDSeq,dmlDtlArr)
	dmlArr(5,0) = dmlDtlArr (0,0)
  	dmlArr(5,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(5,0)&"-"&dmlArr(5,1))
  	'Query#2.7 - Insert data into ODS.MEMBERMEMBERRELATION table.
  	Call execute_member_member_relation_table_insert_dml (dbConnGbl,memberIDSeq,currMemberType,currArttUser,dmlDtlArr)
	dmlArr(6,0) = dmlDtlArr (0,0)
  	dmlArr(6,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(6,0)&"-"&dmlArr(6,1))
	'Query#2.8 - Insert data into ODS.UATMEMBER table.
	Call execute_uat_member_table_insert_dml (dbConnGbl,memberIDSeq,currArttUser,dmlDtlArr)
	dmlArr(7,0) = dmlDtlArr (0,0)
  	dmlArr(7,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(7,0)&"-"&dmlArr(7,1))
	'Query#2.9 - Insert data into ODS.CAREENGINEMEMBERPROCESSSTATUS table.
	Call execute_ce_member_process_table_insert_dml (dbConnGbl,memberIDSeq,currArttUser,dmlDtlArr)
	dmlArr(8,0) = dmlDtlArr (0,0)
  	dmlArr(8,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(8,0)&"-"&dmlArr(8,1))		
	'Query#2.10 - Insert data into ODS.MEMBERPROVIDERRELATIONSHIP table.
	Call execute_member_provider_relation_table_insert_dml (dbConnGbl,providerIDSeq,memberIDSeq,currProviderID,currArttUser,supplierAccountID,dmlDtlArr)
	dmlArr(9,0) = dmlDtlArr (0,0)
  	dmlArr(9,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(9,0)&"-"&dmlArr(9,1))
	'Query#2.11 - Insert data into ODS.Memberpcprelationshiphist table.
	Call execute_member_provider_relation_hist_table_insert_dml (dbConnGbl,providerIDSeq,memberIDSeq,currProviderID,currArttUser,dmlDtlArr)
	dmlArr(10,0) = dmlDtlArr (0,0)
  	dmlArr(10,1) = dmlDtlArr (0,1)
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(10,0)&"-"&dmlArr(10,1))
	'Query#2.12 - Insert data into ODS.PARTYEMAILADDRESS table.
	Call execute_member_email_table_insert_dml (dbConnGbl,partyIDSeq,memberEmailAddr,currArttUser,dmlDtlArr)
	dmlArr(11,0) = dmlDtlArr (0,0)
  	dmlArr(11,1) = dmlDtlArr (0,1)
  	If activateReporting = True Then 
	  	Call execute_member_reporting_table_insert_dml (dbConnGbl,memberIDSeq,currArttUser,dmlDtlArr)
		dmlArr(12,0) = dmlDtlArr (0,0)
	  	dmlArr(12,1) = dmlDtlArr (0,1)
  	End If
'  	Call append_text_to_notepad_file (currLogFile,"",dmlArr(11,0)&"-"&dmlArr(11,1))
	'Query#2.13 - to track member in TDM schema
	If tdmTrackerEnabled = True Then 'C.0-Need to work on this later
'		tdmMemberSQL = "INSERT INTO TDM.TDMMEMBER (ADDRLINE1,ADDRLINE2,CITY,STATE,ZIPCODE,RECORDINSERTDT,RECORDUPDTDT,INSERTEDBY,UPDTDBY,PERSONID,AHMSUPPLIERID,MEMBERID,FIRSTNM,LASTNM,MIDINITAL,GENDER,SSN,DTOFBIRTH,MEMBERPLANID,EMAILADDR,PHONEFAXDISPLAYNUMBER,SAVEDATAFORMONTHS) "&_
'		" VALUES ('" & strADD1 & "' ,null,'" & strCITY & "' , '" & strSTATE & "' , '" & strZIP & "' ,SYSDATE,SYSDATE,'"&currArttUser&"','"&currArttUser&"'," & strPartyID & "," & strSupplier_Hts & "," & strMemID & ",'" & Ucase(strFNAME) & "' , '" & Ucase(strLNAME)& "' ,null,'" & strGENDER & "' ," & strSSN & ",TO_DATE('" & dteDOB & "','MM/DD/YYYY')," & strMemberPlanID & ",null,null," & StrSAVEDATAFORMONTHS & ")"
'		currErrCode = execute_dml_in_database (dbConnGbl,tdmMemberSQL)
	End If 'C.0
	Rem The following FOR loop scans the array (dmlArr) to find if any query retuned actual error code upon insertion.
	errorFreeDmlCount = 0
	For h = 0 To dmlArrSize 'this variable (dmlArrSize) is the first dimension of the DML array.
		currAppendTxt = h&") "&dmlArr(h,0)&"----------------------------------------"&dmlArr(h,1)
		If dmlArr(h,1) <> 0 Or InStr(UCase(dmlArr(h,1)),"ERROR")>0 Then 'C.1 - If there were any error while executing the INSERT DMLs.
			currAppendTxt = h+1&") Member Creation Query#2."&h+1&") "&dmlArr(h,0)&VbCrLf&_
			" had ERROR ,( "&dmlArr(h,1)&" ) while inserting into database, hence member creation for TC WAS NOT SUCCESSFUL!"
			Call append_text_to_notepad_file_without_borders_timestamp (currLogFile,"",currAppendTxt)
			Exit For
			Else 
			errorFreeDmlCount = errorFreeDmlCount+1
		End If 'C.1
	Next
	If errorFreeDmlCount >= dmlArrSize Then	'C.1 - If the total number of error free DMLs is >=12, meaning all DMLs were successful, then commit the transactions and retrieve member information in the array (memberInfoArr).	
		'dbConnGbl.CommitTrans 'commit the transactions in database.
		'Dim memberInfoArr (1,6)
		createdMemberID = memberIDSeq 'Use this createdMemberID (=memberIDSeq) to refer to the member that is created in database.
		
'		Call get_member_info_from_database (createdMemberID,dbConnGbl,memberInfoArr)
'		memberCreated = memberInfoArr (1,0)
		'MsgBox memberInfoArr (1,0)&"-"&memberInfoArr (1,1)
		If IsNull(createdMemberID) = False Then 'C.1.1 - If the member creation was successful and a member id is retrieved from data base.
			create_a_member_in_database = createdMemberID
			Else
			create_a_member_in_database = Null
		End If 'C.1.1
		'Call print_member_information_to_the_log (currLogFile,"",memberInfoArr) 'Captures the member information in the log file
		Else
		'Rollback the transactions
		'dbConnGbl.RollbackTrans
		create_a_member_in_database = Null 'return Null if member was not created due to DB error 
	End If
End Function

Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_member_personal_info ()
Rem Fuction Arguments: ByRef perDob,ByRef perFN, ByRef perLN, ByRef perMI, ByRef perFullNm, ByRef perEmail
Rem Fuction tasks: Function creates member personal information (ie, names) with random characters and numbers.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_member_personal_info (ByRef perDob,ByRef perFN, ByRef perLN, ByRef perMI, ByRef perFullNm, ByRef perEmail)
'Generate all member details, like names, address etc and use given demographics from TC
	If IsDate (perDob) Then 
		Else
		perDob = create_a_date_of_birth_for_given_age (currMemberAge)
	End If
	perFN = "OLE"&rand_num_gen (2,99,10)&rand_str_gen(2)
	perLN = "MAN"&rand_num_gen (2,99,10)&rand_str_gen(2)
	perMI = UCase(rand_str_gen(1))
	perFullNm = perFN&Space(1)&perMI&Space(1)& perLN
	perEmail = perFN&perLN&memberEmailExtension
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: verify_provider_is_valid_if_not_query_db_for_one ()
Rem Fuction Arguments: dbConnGbl (active DB connection), careProviderID (a provider ID is passed in from the caller)
Rem Fuction tasks: This Function is to verify that the defaultProvider (careProviderID) is valid, if not query database to get a valid provider id.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function verify_provider_is_valid_if_not_query_db_for_one (dbConnGbl,careProviderID)
	columnName = "CAREPROVIDERID"
	If careProviderID <> Empty Then 'C.1 - If the provider is not available in the config file, then Query the DB to get a valid providerID for member.
		verifyProviderSQL = "select * from ods.careprovider cp where cp.careproviderid ="&defaultProviderID
		currProviderID = get_column_value_from_a_tupple (dbConnGbl,Empty,verifyProviderSQL,columnName)
		Else
		currProviderID = careProviderID		
	End If 'C.1
	If currProviderID = Empty Then 'C.1.1
			providerSQL = "select * from ods.careprovider cp where cp.careprovidertype = 'PHY' and cp.providerfilterflag = 'N' and cp.provideroptoutflag = 'N' "&_
			"and cp.exclusioncode is null and rownum <=1"		
			currProviderID = get_column_value_from_a_tupple (dbConnGbl,Empty,providerSQL,columnName)
	End If 'C.1.1
	verify_provider_is_valid_if_not_query_db_for_one = currProviderID
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_supplier_accouunt_info ()
Rem Fuction Arguments: dbConnGbl,memSupp (a given supplier for the member) ,ByRef supplierAcctNm, ByRef supplierAcctID
Rem Fuction tasks: This Function find the account name and account id for a given supplier.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_supplier_accouunt_info (dbConnGbl,memSupp,ByRef supplierAcctNm, ByRef supplierAcctID)
	supplierInfoSQL = "SELECT so.supplierorgid,so.orgnm accountName,so.ahmsupplierid,io.INSURANCEORGID accountID FROM ods.supplierorganization so,ods.mastersuppliersupplierrelation mssr,"&_
	"ods.mastersupplierorganization mso,ods.insuranceorgsupplierrelation iosr,ods.insuranceorganization io,ods.projectsupplierrelation psr"&_
	" WHERE so.supplierorgid = mssr.supplierid AND mso.MASTERSUPPLIERORGID = mssr.mastersupplierid AND mssr.mastersupplierid = mso.mastersupplierorgid "&_
	"AND iosr.supplierid = mssr.supplierid AND io.insuranceorgid = iosr.insuranceorgid And Psr.Ahmsupplierid = Mso.Ahmsupplierid AND "&_
	"so.ahmsupplierid In ("&memSupp&")"
	'Set acctRs = get_recordset_from_db_table (dbConnGbl,supplierInfoSQL)
	supplierAcctNm = get_column_value_from_a_tupple (dbConnGbl,Empty,supplierInfoSQL,"ACCOUNTNAME")
	supplierAcctID = get_column_value_from_a_tupple (dbConnGbl,Empty,supplierInfoSQL,"ACCOUNTID")
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: execute_dml_in_database (dbConn,currSQL)
Rem Fuction Arguments: dbConn (and active DB connection),currSQL (SQL statement to be executed)
Rem Fuction tasks: Function executes the DML (currSQL) and return the error code if any error occurs, returns 0 if no error occured.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_dml_in_database (dbConn,currSQL)
On Error Resume Next
	dbConn.Execute currSQL
	If Err.Number <> 0 Then
	  'MsgBox Err.Number
	  dbConn.RollbackTrans
	  execute_dml_in_database = Err.Number
	  Exit Function
	  Else
	  execute_dml_in_database = 0
	End If 
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_a_date_of_birth_for_given_age ()
Rem Fuction Arguments: currMemberAge (an integer value)
Rem Fuction tasks: Function creates a date of birth based on a given age in the form ("01-JAN-2001")
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_a_date_of_birth_for_given_age (currMemberAge)
	If InStr(UCase(currMemberAge),"M")>0 Or InStr(UCase(currMemberAge),"Y") Or InStr(UCase(currMemberAge),"D") Then 'C.1-If the age is in terms of Y (ie, 18Y) or months (ie, 24M)
		ageInterval = Right (currMemberAge,1)
		ageNumber = Left(currMemberAge,Len(currMemberAge)-Len(ageInterval))
		Select Case UCase(ageInterval)
		Case "Y"
		calcDate = DateAdd("YYYY",(-1*ageNumber),Date)
		Case "M"
		calcDate = DateAdd("m",(-1*ageNumber),Date)
		Case "D"
		calcDate = DateAdd("d",(-1*ageNumber),Date)
		End Select
		ElseIf IsNumeric(currMemberAge) Then
			calcDate = DateAdd("YYYY",(-1*currMemberAge),Date)
	End If 'C.1
	
	If IsEmpty (calcDate) = False Then 'C.2 - If the date is calculated correctly
		formattedDate = Day(calcDate)&"-"&MonthName (Month(calcDate),True)&"-"&Year(calcDate)
		create_a_date_of_birth_for_given_age = UCase(formattedDate)
		Else
		create_a_date_of_birth_for_given_age = Empty 'Return empty if no date is calculated
	End If 'C.2	
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_column_value_from_a_tupple ()//NOT IMPLEMENTED YET.
Rem Fuction Arguments: dbConnGbl (Active DB Connection),strSql (the query with the sequence)
Rem Fuction tasks: Function returns the sequence key from a database table given the query with the sequence.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++========================================================= 
Function get_column_value_from_a_tupple (dbConnGbl,currRS,queryStatement,columnName)
	On Error Resume Next
	Err.Clear
	If currRS = Empty Then 
		Set currRS = createobject("ADODB.Recordset") 
		currRS.Open queryStatement,dbConnGbl
	End If 
	If Err.Number = 0 Then 'C.1 - If there was no error while retrieving the record set from db using the sql statement (strSql)
		currValue = currRS.Fields (columnName).Value
		get_column_value_from_a_tupple = currValue
		Set currRS = Nothing
		Else
		Set currRS = Nothing
		get_column_value_from_a_tupple = "ERROR (#"&Err.Number&") - "&Err.Description
	End If 'C.1
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_sequence_key_for_a_given_table ()
Rem Fuction Arguments: dbConnGbl (Active DB Connection),seqName (the name of the sequence from database)
Rem Fuction tasks: Function returns the sequence key from a database table given the query with the sequence.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_sequence_key_for_a_given_table (dbConnGbl,seqName)
	On Error Resume Next
	Err.Clear
	seqSQL = "SELECT "&seqName&".NEXTVAL FROM DUAL"
	'Select ods.ODS_PARTY_SEQ.NEXTVAL FROM DUAL
	Set currRS = createobject("ADODB.Recordset") 
	currRS.Open seqSQL,dbConnGbl
	If Err.Number = 0 Then 'C.1 - If there was no error while retrieving the record set from db using the sql statement (strSql)
		currSeq = currRS.Fields ("NEXTVAL").Value
		'nextSeq = currSeq+1
		get_sequence_key_for_a_given_table = currSeq
		Set currRS = Nothing
		Else
		Set currRS = Nothing
		get_sequence_key_for_a_given_table = "ERROR (#"&Err.Number&") - "&Err.Description
	End If 'C.1
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_recordset_from_db_table ()
Rem Fuction Arguments: dbConnGbl (active DB connection),strSql (sql statements to be executed).
Rem Fuction tasks: Function returns the records set from a database after executing a given sql statements and returns error message if fails.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_recordset_from_db_table (dbConnGbl,strSql)
On Error Resume Next
	Set currRS = createobject("ADODB.Recordset") 
	currRS.Open strSql,dbConnGbl
	If Err.Number = 0 Then 'C.1 - If there was no error while retrieving the record set from db using the sql statement (strSql)
		Set get_recordset_from_db_table = currRS
		Set currRS = Nothing
		Else
		get_recordset_from_db_table = "ERROR (#"&Err.Number&") - "&Err.Description
	End If 'C.1
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_database_connection_string ()
Rem Fuction Arguments: DBHost (DB host address),DBSid (Database SID),DBPort,DBUser (valid user),DBPassword (valid password)
Rem Fuction tasks: Function creates a string that has the ADO string for data base connection and returns the string to the caller.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_database_connection_string (DBHost,DBSid,DBPort,DBUser,DBPassword)
	strConnect =  "Provider=OraOLEDB.Oracle; Data Source=" & _
	"(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST="&DBHost&")(PORT="&DBPort&")))(CONNECT_DATA=(SID="&DBSid&")(SERVER=DEDICATED)));" & _
	"User Id="&DBUser&";Password="&DBPassword&";" 'SERVICE_NAME (insted of SID)
	create_database_connection_string = strConnect 'return the connection string to the caller
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_database_connection_string_ver2 ()
Rem Fuction Arguments: DBHost (DB host address),DBSid (Database SID),DBPort,DBUser (valid user),DBPassword (valid password)
Rem Fuction tasks: Function creates a string that has the ADO string for data base connection and returns the string to the caller (this has a different
Rem driver requirement for Oracle DB).
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_database_connection_string_ver2 (DBHost,DBSid,DBPort,DBUser,DBPassword)
create_database_connection_string_ver2 = "Driver={Microsoft ODBC for Oracle};" & _
                     "CONNECTSTRING=(DESCRIPTION=" & _
        			 "(ADDRESS=(PROTOCOL=TCP)" & _
        			 "(HOST="&DBHost&")(PORT="&DBPort&"))" & _
        			 "(CONNECT_DATA=(SERVER=dedicated)(SID="&DBSid&")));uid="&DBUser&";pwd="&DBPassword&";"
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_db_connection_information ()
Rem Fuction Arguments: (testEnvironment,envInfoDriver,ByRef currDbHost,ByRef currDbSid, ByRef currDbPort,ByRef currDbUser, ByRef currDbPassword, ByRef currWebServiceEndpoint)
Rem Fuction tasks: returns true and hold respected values for each field of the DB connection as specified in the excel (envInfoDriver) 
Rem in the ByRef variables, returns a message if the information cannot be retrieved.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_db_connection_information (testEnvironment,envInfoDriver,ByRef currDbHost,ByRef currDbSid, ByRef currDbPort,ByRef currDbUser, ByRef currDbPassword, ByRef currWebServiceEndpoint)
	openEnvExcel =  load_specified_excel_by_sheet_name ( envInfoDriver,objXl,envInfoExcelSheetName,eBookObj,envInfoXlSheet,1)
	If openEnvExcel = True Then
		envInfoRowNum = get_row_num_from_excel_column_matching_a_key_value (envInfoDriver,1,testEnvironment,envInfoXlSheet,"ENVIRONMENT")
		currDbHost = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,2)
		currDbSid = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,3)
		currDbPort = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,4)
		currDbUser = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,6)
		currDbPassword = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,7)
		currWebServiceEndpoint = get_cell_value_given_rowid_columnid (envInfoXlSheet,envInfoRowNum,8)
		get_db_connection_information = True
		Else
		get_db_connection_information = "The environment information for the given environment ("&testEnvironment&" cannot be retrieved from the given excel file (located @ "&envInfoDriver
	End If 
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: member_demo_extractor ()
Rem Fuction Arguments: currElement (a string with a delimeter, ie-AGE:19,GENDER:MALE,ByRef currVariable (the variable name, ie AGE,delimeter (:, in this case)
Rem Fuction tasks: returns the value of the parameter, 19 from (AGE:19) and also sets the current variable (AGE, in this case).
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function member_demo_extractor (currElement,ByRef currVariable,delimeter)
	memberDemoArr = Split(currElement,delimeter)	
	currVariable = memberDemoArr (0)	
	Select Case UCase(currVariable)
	Case "AGE","A"
	currVariableValue = memberDemoArr (1)
	Case "DOB","DTOFBIRTH","DATEOFBIRTH"
	currVariableValue = memberDemoArr (1)
	Case "GEN","GENDER"
	currVariableValue = memberDemoArr (1)
	Case "TYPE","MTYPE","MEMBERYPE","MBRTYPE","MEMTYPE"
	currVariableValue = memberDemoArr (1)
	Case "SSN","SOCIAL"
	currVariableValue = memberDemoArr (1)
	End Select
	member_demo_extractor = currVariableValue
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: load_config_file ()
Rem Fuction Arguments: fileLoc (location of the config file, needs to be hardcoded in the caller script). 
Rem Fuction tasks: Function reads the config file to be used in the caller script
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function load_config_file (fileLoc)
	Dim fileObj,txtObj 
	Set fileObj = CreateObject("Scripting.FileSystemObject")	
	If fileObj.FileExists (fileLoc) Then
		Set txtObj = fileObj.OpenTextFile (fileLoc,1,True)
		Do While Not txtObj.AtEndOfStream
			curLine = txtObj.ReadLine ()
			Execute curLine				
		Loop
	End If
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_a_value_from_excel_column_matching_a_key ()
Rem Fuction Arguments: fileLoc (location of the excel file opened/loaded),searchColumnNum (column id of the key to be searched with), 
Rem searchKey (a row value matching the searchColumn),currSheet (sheet that needs to be scanned,
Rem columnName (value that will be searched in this column)
Rem Fuction tasks: Function searches for a value that matches the row of the values (searchKey) that is search with the column denoted by searchColumnNum
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_a_value_from_excel_column_matching_a_key (fileLoc,searchColumnNum,searchKey,currSheet,columnName)
	totalRows = currSheet.UsedRange.Rows.Count
	totalColumns = currSheet.UsedRange.Columns.Count
	For i=1 To totalRows 'F.1
	curRow_column1Val = currSheet.Cells(i,searchColumnNum)
		If UCase(curRow_column1Val) = UCase(searchKey) Then 'C.1 - If the first column value match for the environment (QA1=environment)
			For j = 1 To totalColumns 'F.1.a
				curColumn = currSheet.Cells(1,j)
				curRowColumnJVal = currSheet.Cells(i,j)
				If UCase(curColumn) = UCase(columnName) Then 'C.2 - If the specified column () has a matching non-null value
					get_a_value_from_excel_column_matching_a_key = Trim(curRowColumnJVal)
					Exit Function
					Else
				End If 'C.2
				If j = totalColumns Then 'C.3 - If no column value found or empty
					get_a_value_from_excel_column_matching_a_key ="No match found"
				End If 'C.3
			Next 'F.1.a
		End If 'C.1
	Next 'F.1
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: load_specified_excel ()
Rem Fuction Arguments: fileLoc (location of the excel file to be opened/loaded),objXl,sheetNumber,ByRef objXlSheet, 
Rem readwrite (Read=True,Write=False)
Rem Fuction tasks: Function opens an excel file based on a given location, returns 'True' if successfully opened the file othewise returns 'False'.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function load_specified_excel (fileLoc,objXl,sheetNumber,ByRef bookXl,ByRef objXlSheet, readwrite)
	On Error Resume Next
	Set bookXl = objXl.Workbooks.Open (fileLoc,,readwrite)
	Set objXlSheet = bookXl.Sheets(sheetNumber)
	
	If Err.Number <> 0 Then 'If there were no error in creating the excel file/sheet/header
		load_specified_excel =False
		MsgBox Err.Number&"-"&Err.Description
		Else
		load_specified_excel = True 
	End If 
End Function 
Function load_specified_excel_by_sheet_name (fileLoc,objXl,sheetName,ByRef bookXl,ByRef objTCXlSheet, readwrite)
	On Error Resume Next
	Set bookXl = objXl.Workbooks.Open (fileLoc,,readwrite)
	sheetCounter = 0
	For Each objWorksheet in objXl.Worksheets
    	currSheetName = objWorksheet.Name
    	sheetCounter = sheetCounter+1
    	If UCase(sheetName) = UCase(currSheetName) Then 
    	sheetId = sheetCounter
    	Set objTCXlSheet = bookXl.Worksheets(sheetId)
   		Exit For
   		End If
	Next
	
	If Err.Number <> 0 Then 'If there were no error in creating the excel file/sheet/header
		load_specified_excel_by_sheet_name =False
		MsgBox Err.Number&"-"&Err.Description
		Else
		load_specified_excel_by_sheet_name = True 
	End If 
End Function 

Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem function name: create_excel_application_object
Rem function tasks: Function creates objects for excel workbook and worksheet, renames the first work sheet and creates the header for the sheet
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_excel_application_object (ByRef objXl)
	Set oExcel = CreateObject("Excel.Application")
		
	Set objXl = oExcel
	
	If Err.Number <> 0 Then 'If there were no error in creating the excel file/sheet/header
		create_excel_application_object =False
		Else
		create_excel_application_object = True 
	End If 
	
	Set oExcel = Nothing
End Function 

Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_row_num_from_excel_column_matching_a_key_value ()
Rem Fuction Arguments: fileLoc (location of the excel file opened/loaded),searchColumnNum (column id of the key to be searched with), searchKey (a row value matching the searchColumn) 
Rem ,currSheet (sheet that needs to be scanned,columnName (value that will be searched in this column (name of searchColumnNum = columnName)
Rem Fuction tasks: Function searches for a matching values (searchKey) in a given column (searchColumnNum) and return the row number of the matching value
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_row_num_from_excel_column_matching_a_key_value (fileLoc,searchColumnNum,searchKey,currSheet,columnName)
	totalRows = currSheet.UsedRange.Rows.Count
	totalColumns = currSheet.UsedRange.Columns.Count
	For i=1 To totalRows 'F.1
	curRow_column1Val = currSheet.Cells(i,searchColumnNum)
		If UCase(curRow_column1Val) = UCase(searchKey) Then 'C.1 - If the first column value match for the environment (QA1=environment)
			For j = 1 To totalColumns 'F.1.a
				curColumn = currSheet.Cells(1,j)
				curRowColumnJVal = currSheet.Cells(i,j)
				If UCase(curColumn) = UCase(columnName) Then 'C.2 - If the specified column () has a matching non-null value
					get_row_num_from_excel_column_matching_a_key_value = i ' return the i'th value which is the row that matches the passed in value (searchKey)
					Exit Function
					Else
				End If 'C.2
				If j = totalColumns Then 'C.3 - If no column value found or empty
					get_row_num_from_excel_column_matching_a_key_value ="No match found"
				End If 'C.3
			Next 'F.1.a
		End If 'C.1
	Next 'F.1
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_column_id_from_excel_matching_a_column_name (currSheet,rowNum,columnName)
Rem Fuction Arguments: currSheet (opened excel sheet),rowNum,columnName
Rem Fuction tasks: Function searches for a value that matches the row of the values (searchKey) that is search with the column denoted by searchColumnNum and return the row number of the matching value
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_column_id_from_excel_matching_a_column_name (currExcelSheet,rowNum,columnName)
	'totalRows = currSheet.UsedRange.Rows.Count
	totalColumns = currExcelSheet.UsedRange.Columns.Count
	For i= 1 To totalColumns 'F.1
	curRow_colVal = currExcelSheet.Cells(rowNum,i)
		If UCase(curRow_colVal) = UCase(columnName) Then 'C.1 - If the first column value match for the environment (QA1=environment)
			get_column_id_from_excel_matching_a_column_name = i
			Exit Function 
			Else
			get_column_id_from_excel_matching_a_column_name = "Not Found"
		End If
	Next 'F.1
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_cell_value_given_rowid_columnid ()
Rem Fuction Arguments: currSheet (the opened excel sheet,rowId (row number of the excel,colId (column number of the excel)
Rem Fuction tasks: Function returns the value from an opened excel cell, given the coordinates, ie (3,7) will return the non-null/empty value 
Rem that is in cell (3X7), 3rd row and 7th column.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_cell_value_given_rowid_columnid (currSheetLcl,rowId,colId)
	curCellValue = currSheetLcl.Cells(rowId,colId).Value
	If Trim (curCellValue) <> Empty Or Trim(curCellValue) <> "" Or Trim (curCellValue) <> Null Then
		get_cell_value_given_rowid_columnid = curCellValue
		Else
		get_cell_value_given_rowid_columnid = Empty
	End If
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_text_file ()
Rem Fuction Arguments: fileLocation (the location of the file where it should be created,folderName (provide a name if a new folder should be created (OPTIONAL) ,fileName (the name of the file)
Rem Fuction tasks: Function creates a text file in a given location (with/out a new folder name)
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_text_file (fileLocation,folderName,fileName)
	Dim filesys, demofolder, filetxt 
	fileLoc = 	fileLocation&folderName&"\"
	Set filesys = CreateObject("Scripting.FileSystemObject") 
	Set demofolder = filesys.GetFolder(fileLoc) 
	Set filetxt = demofolder.CreateTextFile(fileName&".txt", True) 
	headerLine = create_a_line_of_repeated_characters ("=",30)&create_a_line_of_repeated_characters ("+",40)&create_a_line_of_repeated_characters ("=",30)
	footerLine = create_a_line_of_repeated_characters ("=",30)&create_a_line_of_repeated_characters ("+",40)&create_a_line_of_repeated_characters ("=",30)
	filetxt.WriteLine headerLine
	filetxt.WriteLine("File created @") &Time()&" on "&Date()
	filetxt.WriteLine footerLine
	filetxt.Close 
	Set demofolder = Nothing
	Set filetxt = Nothing
	Set filesys = Nothing 
'	create_text_file = filetxt
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_a_folder ()
Rem Fuction Arguments: folderLocation (location of the folder) ,folderName (a unique name for the folder to be created)
Rem Fuction tasks: Function creates a text file in a given location (with/out a new folder name) and returns a 'true' if the folder is created and return 'false' if the folder already exsits.
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_a_folder (folderLocation,folderName)
	Dim filesys, newfolder, newfolderpath 
	newfolderpath = folderLocation&folderName
	set objFSO=CreateObject("Scripting.FileSystemObject") 
	If objFSO.FolderExists(newfolderpath) = False Then
		objFSO.CreateFolder newfolderpath
		create_a_folder = True
		Else
		create_a_folder = False
	End If 
	Set objFSO = Nothing
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================

Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem FunctionName: create_a_line_of_repeated_characters
Rem FunctionArguments: givenCharacters (ie, "*"),repeatNbr (100)
Rem FunctionTasks: Function creates a line with a given character by repeating it n (repeatNbr) number of times.
Rem CreationDate:12/22/2017
Rem CreatedBy: Mohammad Sarwar
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_a_line_of_repeated_characters (givenCharacters,repeatNbr)
	formedLine = givenCharacters
	For i = 0 To repeatNbr
		formedLine = formedLine&givenCharacters
	Next 
	create_a_line_of_repeated_characters = formedLine
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem FunctionName: append_text_to_notepad_file
Rem FunctionArguments: fileLocation (file directory), fileName (name of the file, ie myFile.txt),appendText (the actual text to be appened, ie 'this text will be appened to the file).
Rem FunctionTasks: Function appends a given text to an existing text file, if file does not exist, then file is created in the same directory.
Rem CreationDate:12/22/2017
Rem CreatedBy: Mohammad Sarwar
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function append_text_to_notepad_file (fileLocation, fileName,appendText)
	Dim fileSys, fileDir 
	fileDir = fileLocation&fileName
	set objFSO=CreateObject("Scripting.FileSystemObject") 
	If objFSO.FileExists(fileDir) = False Then
		Call create_text_file (fileLocation,"",fileName)		
	End If 
	On Error Resume Next 'if any error occurred during file creation, then continue
	If Err.Number <> 0 Then 
		MsgBox "Error - "&Err.Number&" ("&Err.Description&") occured."
		Exit Function 
	End If
	Set objTextFile = objFSO.OpenTextFile (fileDir, 8, True)
	borderLine = create_a_line_of_repeated_characters ("-",104)
	objTextFile.WriteLine(borderLine)
	objTextFile.WriteLine ("/* @"&Replace(Timer(),".",":"))&" */"
	objTextFile.WriteLine(create_a_line_of_repeated_characters ("-",12))
	objTextFile.WriteLine(appendText)
	'objTextFile.WriteLine(borderLine)
	objTextFile.Close
	
	Set objFSO = Nothing
	Set objTextFile = Nothing
End Function 

Function append_text_to_notepad_file_without_borders_timestamp (fileLocation, fileName,appendText)
	Dim fileSys, fileDir 
	fileDir = fileLocation&fileName
	set objFSO=CreateObject("Scripting.FileSystemObject") 
	If objFSO.FileExists(fileDir) = False Then
		Call create_text_file (fileLocation,"",fileName)		
	End If 
	On Error Resume Next 'if any error occurred during file creation, then continue
	If Err.Number <> 0 Then 
		MsgBox "Error - "&Err.Number&" ("&Err.Description&") occured."
		Exit Function 
	End If
	Set objTextFile = objFSO.OpenTextFile (fileDir, 8, True)
	'borderLine = create_a_line_of_repeated_characters ("-",104)
'	objTextFile.WriteLine(borderLine)
'	objTextFile.WriteLine ("@"&Time())&" on "&Date()& VbCrlf
	objTextFile.WriteLine(appendText)
	'objTextFile.WriteLine(borderLine)
	objTextFile.Close
	
	Set objFSO = Nothing
	Set objTextFile = Nothing
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem FunctionName: print_member_information_to_the_log
Rem FunctionArguments: (memberId,dbConn,memInfoArr,additonalTxt)
Rem FunctionTasks: Function prints the information in the notepad log from an array with the member information that are retrieved from database.
Rem CreationDate:2/27/2018
Rem CreatedBy: Mohammad Sarwar
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function print_member_information_to_the_log (fileLocation,fileName,memInfoArr,additonalTxt)
	Dim fileSys, fileDir 
	fileDir = fileLocation&fileName
	set objFSO=CreateObject("Scripting.FileSystemObject") 
	If objFSO.FileExists(fileDir) = False Then
		Call create_text_file (fileLocation,"",fileName)		
	End If 
	On Error Resume Next 'if any error occurred during file creation, then continue
	If Err.Number <> 0 Then 
		MsgBox "Error - "&Err.Number&" ("&Err.Description&") occured."
		Exit Function 
	End If
	Set objTextFile = objFSO.OpenTextFile (fileDir, 8, True)
	
	objTextFile.WriteLine create_a_line_of_repeated_characters ("-",104) 'Call function to create a string with "-" of 100 times 
	objTextFile.WriteLine UCase(additonalTxt)& " MEMBER ( "&memInfoArr(1,1)&" ) is successfully created, member details listed below."
	objTextFile.WriteLine create_a_line_of_repeated_characters ("-",104) 'Call function to create a string with "-" of 100 times 
	'MsgBox UBound(memInfoArr,2)
	For i = 0 To UBound(memInfoArr,1)
		msgToPrint = ""
		For j = 0 To UBound(memInfoArr,2)
			msgToPrint = msgToPrint & UCase(memInfoArr (i,j))&"|"
		Next
		objTextFile.WriteLine msgToPrint
	Next
	objTextFile.WriteLine create_a_line_of_repeated_characters ("-",104) 'Call function to create a string with "-" of 100 times 
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: get_member_info_from_database
Rem **	Task(s): This function gets member information from database and put them in a multi-dimensional array.
Rem **	Parameter: Input parameters are 1) queryWithMember (the memberID to query with),2) dbConnGbl, 3) ByRef memInfoArrLocal, a (1x7) array to be filled
Rem **  with from the result of the query.
Rem **	Date created: 5/4/2015
Rem **	Revision History: Revised on 9/14/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_member_info_from_database (queryWithMember,dbConnGbl,ByRef memInfoArrLocal)
	sqlQuery = "Select mem.ahmsupplierid,mem.memberid,mem.primarymemberplanid,mem.sourcememberpatientid,per.dtofbirth,per.gender,mem.personid,mem.membertypecode,per.firstnm,per.lastnm "&_
	" from ods.person per, ods.member mem where mem.personid = per.personid and mem.memberid = "&queryWithMember 
	Set oRs8 = get_recordset_from_db_table (dbConnGbl,sqlQuery)
	'Dim memInfoArrLocal (1,7)
	memInfoArrLocal (0,0) = "supplierid"
	memInfoArrLocal (0,1) = "memberid"
	memInfoArrLocal (0,2) = "memberplanid"
	memInfoArrLocal (0,3) = "sourcepatientid"
	memInfoArrLocal (0,4) = "dtofbirth"
	memInfoArrLocal (0,5) = "gender"
	memInfoArrLocal (0,6) = "personid"
	memInfoArrLocal (0,7) = "membertypecode"
	memInfoArrLocal (0,8) = "firstname"
	memInfoArrLocal (0,9) = "lastname"
	
	While Not oRs8.EOF
		memInfoArrLocal (1,0) = oRs8.Fields(0).Value
		memInfoArrLocal (1,1) = oRs8.Fields(1).Value
		memInfoArrLocal (1,2) = oRs8.Fields(2).Value
		memInfoArrLocal (1,3) = oRs8.Fields(3).Value
		memInfoArrLocal (1,4) = oRs8.Fields(4).Value
		memInfoArrLocal (1,5) = oRs8.Fields(5).Value
		memInfoArrLocal (1,6) = oRs8.Fields(6).Value
		memInfoArrLocal (1,7) = oRs8.Fields(7).Value
		memInfoArrLocal (1,8) = oRs8.Fields(8).Value
		memInfoArrLocal (1,9) = oRs8.Fields(9).Value
		oRs8.MoveNext
	Wend
	Set oRs8 = Nothing

End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: validate_passkey
Rem **	Task(s): This function validates whether a given passkey is VALID (returns TRUE) or INVALID (returns FALSE).
Rem **	Parameter: Input parameters are 1) 'PassKey' to specify the passkey provided to the user and 2) passwordDecEnc (the actual ecrypted or decrypted password)
Rem	**	Parameter list: nChar
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function validate_passkey (passKey,passwordDecEnc)
	Call load_config_file (authUserInitsDir)
	'passkeyInits = "MS|AK|NC|KA|RR|VP|NA|NK" ' Initials can be added or removed as needed
	passkeyInitsArr = Split (passkeyInits,"|") 'fill in the array with the values in 'passkeyInits' 
	'If InStr (passKey,"-")>0 Then 'C.b
	Rem - passKeyInitPassed = UCase(Left(passKey,2)) 'the first 2 characters are assumed to be the initials of a resouce.
	Rem - passkeyNumber = Mid(passKey,3,Len(passKey)-2) ' the passkeyNumber is the number given by admin
'	End If 'C.b
	Rem - This (C.a) is enhanced on 10/13/2018
	If InStr (passKey,"-")>0 Then 'C.a	
		passKeyArr = Split(passKey,"-")
		passKeyInitPassed = passKeyArr(0)
		passkeyNumber = passKeyArr(1)
		Else
		passKeyInitPassed = UCase(Left(passKey,2)) 'the first 2 characters are assumed to be the initials of a resouce.
		passkeyNumber = Mid(passKey,3,Len(passKey)-2) ' the passkeyNumber is the number given by admin
	End If 'C.a
	
	currPassword = passwordDecEnc 'InputBox ("Enter the current password (decrypted)")' Call the function to get the ASCII for the current password.
	aSampleKeypass = generate_a_key_for_db_access (currPassword,currPasswordASCII,"")
	
	If CLng(passkeyNumber) Mod CLng(currPasswordASCII) = 0 Then 'C1 - If the passkey number is a multiple of the ASCII number of the sum of the password 
		boolPassKeyNum = True 
	End If 'C1 - If the passkey number is a multiple of the ASCII number of the sum of the password 
	For i = 0 To UBound(passkeyInitsArr)
		If passkeyInitsArr(i) = passKeyInitPassed Then 'C2 - If the initials that is passed in with the passkey matches one of the initials in this function.
			boolInit = True 
			Exit For
		End If 'C2
	Next

	If boolPassKeyNum = True And boolInit = True Then 'C3
		validate_passkey = True
		Else
		validate_passkey = False
	End If 'C3
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: encrypt_val
Rem **	Task(s): This function encrypts a password using the formula (Chr(Asc(Mid(enc_var,a+1,1))+2) and revers it)
Rem **	Parameter: Input parameter is the enc_var to specify string to be encrypted by using the formula (Chr(Asc(Mid(enc_var,a+1,1))+2))
Rem	**	Parameter list: enc_var
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
'Encryption function – 
Function encrypt_val (enc_var)
    arrSize = Len(enc_var)-1
    ReDim myArrDec(arrSize)
    
    For a=0 To arrSize
        myArrDec(a)=Chr(Asc(Mid(enc_var,a+1,1))+2)
    Next
    
    dec_var = ""
    
    For i=0 To arrSize
        dec_var = dec_var + myArrDec(arrSize-i)
    Next 
    encrypt_val = dec_var
End function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: decrypt_val
Rem **	Task(s): This function decrypts an encrypted password using the formula (Chr(Asc(Mid(enc_var,a+1,1))-2) and revers it)
Rem **	Parameter: Input parameter is the enc_var to specify string to be encrypted by using the formula (Chr(Asc(Mid(enc_var,a+1,1))-2))
Rem	**	Parameter list: enc_var
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
'Decryption function
Function decrypt_val (enc_var)
    arrSize = Len(enc_var)-1
    ReDim myArrDec(arrSize)
    
    For a=0 To arrSize
        myArrDec(a)=Chr(Asc(Mid(enc_var,a+1,1))-2)
    Next
    
    dec_var = ""
    
    For i=0 To arrSize
        dec_var = dec_var + myArrDec(arrSize-i)
    Next 
    decrypt_val = dec_var
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Function Name: generate_a_key_for_db_access
Rem Function Arguments: valToBeEnc, inits (initials of the person, ie MS for Mohammad Sarwar).
Rem Function tasks: Fucntion to return a multiple of the sum of the ASCII values of the passed string.
Rem Creation date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function generate_a_key_for_db_access (valToBeEnc,ByRef valToBeEncASCII,inits)
	arrSize = Len(valToBeEnc)-1
	
	ReDim myArrDec(arrSize)
	valToBeEncASCII =0  
	For a=0 To arrSize
	    myArrDec(a)=Asc(Mid(valToBeEnc,a+1,1))
	Next
	
	For i=0 To arrSize
	        valToBeEncASCII = valToBeEncASCII + myArrDec(i)'myArrDec(arrSize-i)
	Next 
	generate_a_key_for_db_access = UCase(inits) & valToBeEncASCII * rand_num_gen (2,99,50)
'generate_a_key_for_db_access = enc_val
End Function 

Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: rand_num_gen
Rem **	Task(s): This function provides a random number of n-digit. 
Rem **	Parameter: Input parameter is the 'nDig' to specify the number of digit the random number should have.
Rem		'upperLim' to specify the upper most number expected within nDigit and 'lowerLim' is to specify the lower most number expected within n digits.
Rem	**	Parameter list: nDig
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function rand_num_gen (nDig, upperLim, lowerLim)
Randomize
	If nDig = Empty Then nDig = Len (upperLim) End If
	If Trim (upperLim) = Empty Then upperLim = get_upper_num (nDig)End If
	If Trim(lowerLim) = Empty Then lowerLim = get_lower_num (nDig)End If
	fracNum = (upperLim-lowerLim+1)* Rnd ()
	Rem MsgBox fracNum
	rand_num_gen = Int(fracNum + lowerLim)
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: get_upper_num
Rem **	Task(s): This function creates an upper most possible number of n-digit. 
Rem **	Parameter: Input parameter is the nDig to specify the number of digit the number should have.
Rem	**	Parameter list: nDig
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_upper_num (nDig)
	startNum = 1
	For i=0 To nDig-1
	startNum = Int (startNum & "0")
	Next 
	get_upper_num = startNum - 1
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: get_lower_num
Rem **	Task(s): This function creates a lower most possible number of n-digit. 
Rem **	Parameter: Input parameter is the nDig to specify the number of digit the number should have.
Rem	**	Parameter list: nDig
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_lower_num (nDig)
	startNum = 1
	For i=0 To nDig-1
	If i <> 0 Then startNum = Int (startNum & "0") End If
	Next 
	get_lower_num = startNum 
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: rand_str_gen
Rem **	Task(s): This function creates a random string of n-digit. 
Rem **	Parameter: Input parameter is the nChar to specify the number of characters the string should have.
Rem	**	Parameter list: nChar
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function rand_str_gen (nChar)
	startStr = " "
	For i=1 To nChar
	startStr = Trim(startStr + Chr(rand_num_gen (Empty,90,65)))
	rand_str_gen = startStr
	'MsgBox startStr
	Next
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: get_time_date_stamp
Rem **	Task(s): This function creates a time stamp in this format, 12_15_2018_12_15_30_PM 
Rem **	Parameter: None
Rem **	Date created: 5/4/2015
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function get_time_date_stamp ()
	get_time_date_stamp = Replace(Replace(Date,"/","_")&"_"&Replace(Time,":","_")," ","_")
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: ** Name: The following 12 Queries are used to create a member in database with the INSERT DMLs for the respecitve tables in ODS schema.
Rem **	Task(s): This functions should be called in sequence from the main script so that member creation is smooth. All these functions insert the 
Rem passed in values into the database and returns the error code (0 if no error occurs otherwise returns the actual error number) and the DML statement
Rem it executed, via the ByRef variable (which is a 0x1 array).
Rem **	Parameter: listed in each function parameter list.
Rem **	Date created: 9/4/2018
Rem **	Revision History: None
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#1 - FOR ODS.PARTY table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_party_table_insert_dml (dbConnGbl,partySkey, ByRef partyArr)
	partyIDSQL = "INSERT INTO ODS.PARTY (PARTYID, PARTYTYPECODE) VALUES ("&partySkey& ",'P' )" 
	currErrCode = execute_dml_in_database (dbConnGbl,partyIDSQL)
	partyArr(0,0) = partyIDSQL
	partyArr(0,1) = currErrCode
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#2 - FOR ODS.PARTYADDRESS table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_party_address_table_insert_dml (dbConnGbl,partyAddrSkey,partySkey,currAdd1,currCity,currState,currZip,currUser, ByRef partyAddArr)
	partyAddrSQL =  "INSERT INTO ODS.PARTYADDRESS (PARTYADDRID, PARTYID, ADDRTYPE, ADDRUSAGETYPE, ADDRLINE1, CITY, STATE, ZIPCODE, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, UPDTDATASOURCENM,CREATEDBYDATASOURCENM) "&_
					"VALUES  ("& partyAddrSkey & "," & partySkey & ", null,'HOME','"&currAdd1&"' , '" &currCity&"' , '" & currState & "' , '" &currZip& "' ,SYSDATE , SYSDATE,'"&currUser&"', '"&currUser&"', 'HDMS','HDMS')"
	currErrCode = execute_dml_in_database (dbConnGbl,partyAddrSQL)
	partyAddArr(0,0) = partyAddrSQL
	partyAddArr(0,1) = currErrCode
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#3 - FOR ODS.PERSONXREF table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_personxref_table_insert_dml (dbConnGbl,partySkey,currMemberSupplier,memberPatIDSeq,currArttUser,memberIDSeq,ByRef personXrefArr)
	personXrefSQL = "INSERT INTO ODS.PERSONXREF (PERSONXREFSKEY, DATASOURCENM, PERSONID, SOURCEALTERNATEUNIQUEID, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, AHMSUPPLIERID, MEMBERID, SOURCEMEMBERPATIENTID) "&_
					"VALUES   (ods.ODS_PERSONXREF_SEQ.NEXTVAL, 'CAREENGINE',"&partySkey&",'"& currMemberSupplier&"-*-"&memberPatIDSeq&"',SYSDATE,SYSDATE,'"&currArttUser&"','"&currArttUser&"',"&currMemberSupplier& "," & memberIDSeq & ",'" & memberPatIDSeq & "')" 
	currErrCode = execute_dml_in_database (dbConnGbl,personXrefSQL)
	personXrefArr (0,0) = personXrefSQL
	personXrefArr (0,1) = currErrCode
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#4 - FOR ODS.PERSON table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_person_table_insert_dml (dbConnGbl,partySkey,memberFirstName,memberMiddleInitial,memberLastName,memberFullName,currMemberGender,currMemberSSN,currMemberDOB,currArttUser,ByRef personArr)
	personSQL = "INSERT INTO ODS.PERSON (PERSONID, FIRSTNM, MIDDLEINITIAL,LASTNM, FULLNM, GENDER, SSN, DTOFBIRTH, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, LAST4SSN) "&_
				"VALUES  ("&partySkey& ",'"&memberFirstName&"','"&memberMiddleInitial&"','"&memberLastName&"','"&memberFullName&"','"&currMemberGender&"',"&currMemberSSN&", TO_DATE('"&currMemberDOB&"','DD/MM/YYYY'),SYSDATE,SYSDATE,'"&currArttUser&"','"&currArttUser&"',"&Right(currMemberSSN,4)&")"
	currErrCode = execute_dml_in_database (dbConnGbl,personSQL)
	personArr (0,0) = personSQL
	personArr (0,1) = currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#5 - FOR ODS.PERSONFACT table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_person_fact_table_insert_dml (dbConnGbl,supplierAccountID,currMemberSupplier,memberIDSeq,memberPlanIDSeq,partySkey,memberFirstName,memberLastName,currMemberDOB,memberGender,defaultCITY,supplierAccountName,memberFullName,Byref personFactArr)
	personFactSQL = "INSERT INTO ODS.PERSONFACT (USAGEMNEMONIC, INSURANCEORGID, AHMSUPPLIERID, MEMBERID, PRIMARYMEMBERPLANID, PERSONID, FIRSTNM, LASTNM, DOB, GENDER, CITY, LASTBUSINESSAHMSUPPLIERID, SUPPLIERNAME, EFFECTIVESTARTDT, EFFECTIVEENDDT, ACTELIGIBILITYFLG, DATASOURCENM, ACTSEARCHFLAG, AETNAEMPLFLG, FULLNM) "&_
					"VALUES ('P',"&supplierAccountID&","&currMemberSupplier&", "&memberIDSeq&", "&memberPlanIDSeq&", "&partySkey&", '" &memberFirstName& "' , '" &memberLastName&"', TO_DATE('"&currMemberDOB&"','DD/MM/YYYY'),'" &memberGender&"', '"&defaultCITY&"', null, '"&supplierAccountName&"', TRUNC(SYSDATE) , Null,'Y', 'CAREENGINE', 'Y','N', '"&memberFullName&"')"
	currErrCode = execute_dml_in_database (dbConnGbl,personFactSQL)
	personFactArr (0,0)= personFactSQL
	personFactArr (0,1)= currErrCode	
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#6 - FOR ODS.MEMBER table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_member_table_insert_dml (dbConnGbl,memberSkey,currMemberSupplier,memberPatID,partySkey,currMemberType,currMemberDOB,currArttUser,memberPlanIDSeq,ByRef memArr)
	If activateMember <> True Then
		activateMemberFilter = "TO_DATE('"&currMemberDOB&"','DD/MM/YYYY'), TO_DATE('"&currMemberDOB&"','DD/MM/YYYY')"
		ElseIf activateMember = True Then
		activateMemberFilter = "TO_DATE('"&currMemberDOB&"','DD/MM/YYYY'),NULL"
	End If 
	memberIdSQL = 	"INSERT INTO ODS.MEMBER (MEMBERID, DATASOURCENM, AHMSUPPLIERID, SOURCEMEMBERPATIENTID, PERSONID, MEMBERTYPECODE, EFFECTIVESTARTDT, EFFECTIVEENDDT, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, PRIMARYMEMBERPLANID) "&_
					" VALUES  ( "&memberSkey&" , 'CAREENGINE' , "&currMemberSupplier&" , '"&memberPatID& "' , "&partySkey&", '"&currMemberType&"' , "&activateMemberFilter& ", SYSDATE, SYSDATE, '"&currArttUser&"' , '"&currArttUser&"' , "&memberPlanIDSeq&")" 
	currErrCode = execute_dml_in_database (dbConnGbl,memberIdSQL)
	memArr (0,0)= memberIdSQL
	memArr (0,1)= currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#7 - FOR ODS.MEMBERMEMBERRELATION table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_member_member_relation_table_insert_dml (dbConnGbl,memberSkey,currMemberType,currArttUser,ByRef memMemRelArr)
	memberRelationSQL = "INSERT INTO ODS.MEMBERMEMBERRELATION (MEMBERMEMBERSKEY, MEMBERID, DEPENDENTMEMBERID, DEPENDENTTYPE, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, DEPENDENTSUBTYPECD) "&_
						"VALUES   (ODS.ODS_MBRMBR_SEQ.NEXTVAL,"& memberSkey& "," & memberSkey & ",'"&currMemberType&"',SYSDATE,SYSDATE,'"&currArttUser&"','"&currArttUser&"','"&currMemberType&"')" 'DEPTYPE = MEMBERTYPE
	currErrCode = execute_dml_in_database (dbConnGbl,memberRelationSQL)
	memMemRelArr (0,0)= memberRelationSQL
	memMemRelArr (0,1)= currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#8 - FOR ODS.UATMEMBER table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_uat_member_table_insert_dml (dbConnGbl,memberSkey,currArttUser,ByRef uatMemArr)
	memberUatSQL = "INSERT INTO ODS.UATMEMBER VALUES ("& memberSkey & ", SYSDATE, SYSDATE, '"&currArttUser&"','"&currArttUser&"','PHR')"
	currErrCode = execute_dml_in_database (dbConnGbl,memberUatSQL)
	uatMemArr (0,0)= memberUatSQL
	uatMemArr (0,1)= currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#9 - FOR ODS.CAREENGINEMEMBERPROCESSSTATUS table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_ce_member_process_table_insert_dml (dbConnGbl,memberSkey,currArttUser,ByRef ceMemProcessArr)
	CEprocessBitSQL = 	"INSERT INTO ODS.CAREENGINEMEMBERPROCESSSTATUS (MEMBERID, BATCHID, PROCESSEDFLAG, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, PROCESSEDBITIND) "&_
						"VALUES   ("& memberSkey & ", 1, 'N', SYSDATE, SYSDATE, '"&currArttUser&"', '"&currArttUser&"', 1)"
	currErrCode = execute_dml_in_database (dbConnGbl,CEprocessBitSQL)
	ceMemProcessArr (0,0)= CEprocessBitSQL
	ceMemProcessArr (0,1)= currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#10 - FOR ODS.MEMBERPROVIDERRELATIONSHIP table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_member_provider_relation_table_insert_dml (dbConnGbl,providerIDSeq,memberSkey,currProviderID,currArttUser,supplierAccountID,ByRef memProvRelArr)
	memberProviderRelSQL = 	"INSERT INTO ODS.MEMBERPROVIDERRELATIONSHIP (MEMBERPROVIDERSKEY, MEMBERID, PROVIDERID, DATASOURCENM, PCPFLG, PROVIDERTYPECD, RELATIONSTATUSCD, EXCLUSIONCD, RECORDINSERTDT, RECORDUPDTDT, INSERTEDBY, UPDTDBY, ACCOUNTID, MASTERCAREPROVIDERID, AHMMEMBERID,WINNERPCPFLG) "&_
							"VALUES ( "&providerIDSeq&","&memberSkey&", "&currProviderID&", 'HDMS', 'Y', 'P', 'CR','IN', sysdate, sysdate, '"&currArttUser&"', '"&currArttUser&"',"&supplierAccountID&", NULL," & memberSkey & ", 'Y')"
	currErrCode = execute_dml_in_database (dbConnGbl,memberProviderRelSQL)
	memProvRelArr (0,0)= memberProviderRelSQL
	memProvRelArr (0,1)= currErrCode
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#11 - FOR ODS.MEMBERPCPRELATIONSHIPHIST table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++========================================================= 
Function execute_member_provider_relation_hist_table_insert_dml (dbConnGbl,providerIDSeq,memberSkey,currProviderID,currArttUser,ByRef memProvRelHistArr)
	memberProviderRelHistSQL = 	"INSERT INTO ODS.MEMBERPCPRELATIONSHIPHIST (MEMBERPCPHISTSKEY,MEMBERPROVIDERSKEY,MEMBERID,PROVIDERID,EFFSTARTDT,EFFENDDT,PCPFLG,DATASOURCENM,VENDORSOURCENM,CLINICALDOCTYPEMNEMONIC,RECORDINSERTDT,RECORDUPDATEDT, INSERTEDBY,UPDATEDBY) "&_
								"VALUES (ods.ODS_MBRPROVHIST_SEQ.nextval,"&providerIDSeq&", "&memberSkey&", "&currProviderID&", SYSDATE, SYSDATE, 'Y' ,'HDMS', null,null, SYSDATE,SYSDATE,'"&currArttUser&"','"&currArttUser&"')"
	currErrCode = execute_dml_in_database (dbConnGbl,memberProviderRelHistSQL)
	memProvRelHistArr (0,0) = memberProviderRelHistSQL
	memProvRelHistArr (0,1) = currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem: MEMBER CREATION FUNCTION#12 - FOR ODS.PARTYEMAILADDRESS table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_member_email_table_insert_dml (dbConnGbl,partySkey,memberEmailAddr,currArttUser,ByRef memEmailArr)
	partyEmailSQL = "INSERT INTO ODS.PARTYEMAILADDRESS (Emailid,Partyid,Zdel_Emailpreferenceseq,Emailaddr,Emailtypecode,Permissiontocontact,Effectivestartdt,Effectiveenddt,Recordinsertdt,Recordupdtdt,Insertedby,Updtdby,Exclusioncode,Updtdatasourcenm,Deletedbydatasourcenm,Emailpreferenceflg,Emailformatcd,Createdbydatasourcenm,Preferredflg)"&_
					"VALUES (Ods.Ods_Email_Seq.Nextval,"&partySkey&",1,'"&memberEmailAddr&"',Null,'NA',Null,Null,Sysdate,Sysdate,'"&currArttUser&"','"&currArttUser&"',Null,'PHR_UE',Null,'Y',Null,'HDMS','Y')"
	currErrCode = execute_dml_in_database (dbConnGbl,partyEmailSQL)
	memEmailArr (0,0) = partyEmailSQL
	memEmailArr (0,1) = currErrCode
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem PERSON AGGREGATION FUNCTION#13 - FOR ODS.PERSONAGGREGATION table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_person_aggregation_table_insert_dml (dbConnection,regMember,aggMember,currArttUser)
	aggMemSQL = "INSERT INTO ODS.PERSONAGGREGATION (AGGREGATEMEMBERID,MEMBERID,EFFECTIVESTARTDT,EFFECTIVEENDDT,INSERTEDBY,INSERTEDDT,UPDATEDBY,UPDATEDDT) "&_
	"VALUES ("&aggMember&","&aggMember&",SYSDATE-1,NULL,'"&currArttUser&"',SYSDATE,'"&currArttUser&"',SYSDATE)"
	
	regMemSQL = "INSERT INTO ODS.PERSONAGGREGATION (AGGREGATEMEMBERID,MEMBERID,EFFECTIVESTARTDT,EFFECTIVEENDDT,INSERTEDBY,INSERTEDDT,UPDATEDBY,UPDATEDDT) "&_
	"VALUES ("&aggMember&","&regMember&",SYSDATE-1,NULL,'"&currArttUser&"',SYSDATE,'"&currArttUser&"',SYSDATE)"
	
	errorCode1 = execute_dml_in_database (dbConnection,aggMemSQL)
	errorCode2 = execute_dml_in_database (dbConnection,regMemSQL)
	If errorCode1 = 0 And errorCode2 = 0 Then
		execute_person_aggregation_table_insert_dml = True
		Else
		execute_person_aggregation_table_insert_dml = False
	End If
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem MEMBER REPORTING FUNCTION#14 - FOR ODS.MEMBERREPORTINGGROUP table.
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function execute_member_reporting_table_insert_dml (dbConnGbl,memberIDSeq,currArttUser,memberReportingArr)
	memberReportingSQL = "INSERT INTO ODS.MEMBERREPORTINGGROUP (MEMBERID,REPORTGROUPVALUE1,REPORTGROUPVALUE2,REPORTGROUPVALUE3,REPORTGROUPVALUE4,REPORTGROUPVALUE5,"&_
	"REPORTGROUPVALUE6,REPORTGROUPVALUE7,REPORTGROUPVALUE8,REPORTGROUPVALUE9,REPORTGROUPVALUE10,INSERTEDBY,UPDTDBY,RECORDINSERTDT,RECORDUPDTDT)"&_
	" VALUES ("&memberIDSeq&",'TEST-GROUP-"&currArttUser&"',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'"&currArttUser&"','"&currArttUser&"',SYSDATE,SYSDATE)"
	currErrCode = execute_dml_in_database (dbConnGbl,memberReportingSQL)
	memberReportingArr (0,0) = memberProviderRelHistSQL
	memberReportingArr (0,1) = currErrCode
End Function
Rem ========================================================================================================================================
Rem FunctionName: sort_number_in_ascending_order
Rem FunctionParams: passedStr,delimeter where passedStr is a string containing numbers (ie,4,12 etc.) and the delimeter is character that
Rem separates each number in the string (ie,",","/" etc.)
Rem FunctionTasks: 'Function returns a string that contains the number in the ascending order.
Rem CreationDate: 2/15/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function sort_number_in_ascending_order (passedStr,delimeter)
	arr = Split(passedStr,delimeter)
	For i = LBound(arr) to UBound(arr)
	  For j = LBound(arr) to UBound(arr) - 1
	      If CInt(arr(j)) > CInt(arr(j + 1)) Then
	         TempValue = CInt(arr(j + 1))
	         arr(j + 1) = arr(j)
	         arr(j) = TempValue
	      End If
	  Next
	Next
	 
	s = ""
	For i = LBound(arr) To UBound(arr)
	    s = s & arr(i) & delimeter
	Next 
	
	If CStr(Right(s,1)) = CStr(delimeter) Then 
		sort_number_in_ascending_order = Left(s,Len(s)-1)
		Else
		sort_number_in_ascending_order = s
	End If
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ** 	Name: verify_number_exist_in_container
Rem **	Task(s): This function verifies whether a number is in the container (array), retruns 'True' if exists otherwise returns 'False'
Rem **	Parameter: numContainer (the array containing the numbers,num (the number that to be verified).
Rem **	Date created: 9/14/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function verify_number_exist_in_container (numContainer,num)
	For z = 0 To UBound(numContainer)
	currElement = numContainer(z)
	If CInt(currElement) = CInt(num) Then
		verify_number_exist_in_container = True
		Exit For
	Else
		verify_number_exist_in_container = False
	End If 
	Next
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem function name: create_excel_output_file ()
Rem function tasks: Function creates objects for excel workbook and worksheet, renames the first work sheet and creates the header for the sheet
Rem Creation Date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_excel_output_file (ByRef objXl,ByRef objXlBook, ByRef objXlSheet)
	Set oExcel = CreateObject("Excel.Application")
	Set oWorkBook = oExcel.Workbooks.Add()
	Set oWorkSheet = oWorkBook.Worksheets(1)
		
	Set objXl = oExcel
	Set objXlBook = oWorkBook
	Set objXlSheet = oWorkSheet
	
	If Err.Number <> 0 Then 'If there were no error in creating the excel file/sheet/header
		create_excel_summary_output_file =False
		Else
		create_excel_summary_output_file = True 
	End If 
	
	Set oWorkBook = Nothing
	Set oWorkSheet = Nothing
	Set oExcel = Nothing
End Function 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem function name: create_header_for_excel_file ()
Rem Arguments: filePath (The directory for the excel file),objXl (not needed),objBook,objSheet,sheetId,sheetName,colNamesStr,strDelimeter
Rem function tasks: Function creates objects for excel workbook and worksheet, renames the first work sheet and creates the header for the sheet
Rem Creation Date: 9/20/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_header_for_excel_file (filePath,objBook,objSheet,sheetId,sheetName,colNamesStr,strDelimeter)
	objBook.Sheets(sheetId).Name = sheetName
	colNamesStrArr = Split(colNamesStr,strDelimeter)
	colCount = UBound(colNamesStrArr)+1
	For u=1 To colCount
		objSheet.Cells(1,u) = colNamesStrArr(u-1)
	Next	
	Set rng1 = objSheet.Range(objSheet.Cells(1,1),objSheet.Cells(1,colCount))    
    With rng1  
    .Interior.ColorIndex = 33
    .Borders.LineStyle = xlDouble
    .Borders.ColorIndex = 30 'dark chocolate
    .Font.ColorIndex = 9 'dark blue
    End With     

	objBook.SaveAs filePath
	'Close the current excel book
	objBook.Close
'	objXl.Quit
'	Set objXl = Nothing
'	Set objBook = Nothing
End Function 
Rem ========================================================================================================================================
Rem FunctionName: get_atom_code_and_code_system
Rem FunctionParams: dbConn,elementIDStr (either the elementID (1111) or the elementID with the related atome, ie, 1111~G2305),ByRef elementID, ByRef atomCode, ByRef codeSystem
Rem FunctionTasks: 'Function returns the related atom code (randomly chosen from the query results of database) with the system name (ie, NDC for Drug codes, ICD9CM for Diagnosis codes) and 
Rem returns '0000' code and 'INVALID_ELEMENT' if there are no atom mapped to the element or atom itself is not found with the element.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_event_class_by_event_type (eventTYpe)
	Select Case UCase(eventTYpe) 'S.2 'The types claims that are allowed, ie, DIAGNOSIS
		Case "DIAG","DIAGNOSIS","DGS"
		eventClass = 8
		Case "PROC","PROCEDURE","PDR"
		Case "DRUG","DRG","NDC"
		Case "LAB","LABS","LOINC"
		Case "UTZ","UTILIZATION","UTILIZATIONS"
	End Select
End Function 
Rem ========================================================================================================================================
Rem FunctionName: get_count_for_a_recordset
Rem FunctionParams: recordSet (the query results from database)
Rem FunctionTasks: 'Function returns the count of the rows from a record set.
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_count_for_a_recordset (recordSet)
	Dim rsCount
	rsCount = 0
	While Not recordSet.EOF
		rsCount = rsCount+1
		recordSet.MoveNext
	Wend
	get_count_for_a_recordset = rsCount
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: choose_recordset_values_on_rownum
Rem FunctionParams: recordSet (the query results from database)
Rem FunctionTasks: 'Function returns value of a given column (fieldNum) for a given row (rowNum) from a recordset.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function choose_recordset_values_on_rownum (recordSet,rowNum,fieldNum)
	currRow = 1
	recordSet.MoveFirst
	Do While Not recordSet.EOF
	'MsgBox recordSet.Fields(0).Value&"-"&recordSet.Fields(2).Value
		If currRow = rowNum Then
			currRowVal = recordSet.Fields(fieldNum).Value
			Exit Do
			Else
			currRowVal = ""
		End If 
		currRow = currRow+1
		recordSet.MoveNext
	Loop
	choose_recordset_values_on_rownum = currRowVal
End Function
Rem ========================================================================================================================================
Rem FunctionName: get_atom_code_and_code_system
Rem FunctionParams: dbConn,elementIDStr (either the elementID (1111) or the elementID with the related atome, ie, 1111~G2305),ByRef elementID, ByRef atomCode, ByRef codeSystem
Rem FunctionTasks: 'Function returns the related atom code (randomly chosen from the query results of database) with the system name (ie, NDC for Drug codes, ICD9CM for Diagnosis codes) and 
Rem returns '0000' code and 'INVALID_ELEMENT' if there are no atom mapped to the element or atom itself is not found with the element.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_atom_code_and_code_system (eventType,ByRef eventClass,elementIDStr,ByRef elementID, ByRef atomCode, ByRef codeSystem)
	If InStr(elementIDStr,"~")>0 Then 'C.1 - If the atom id exist with element id (1111~G5678)
		elementIDArr = Split (elementIDStr,"~")
		elementID = elementIDArr (0)
		atomID = elementIDArr (1)
		ElseIf IsNumeric (elementIDStr) Then 
		elementID = elementIDStr
	End If 'C.1
	'eventClass = get_event_class_by_event_type (eventType)
	If IsEmpty (atomID) Then 'C.2 - If the element needs to be verified and a related atom needs to be retrieved from database use this query.
		additionalFilter = "" 
		Else
		additionalFilter = " and atm.atom ='"&atomID&"'"
	End If 'C.2
	
	atomElementSQL = "select atm.atom,atm.elementid,atm.elementclass,atm.cdsystemnm from ods.atom atm, ods.element elm "&_
	"where atm.elementid = elm.elementid and atm.elementid in ("&elementID&")"&additionalFilter&" and rownum<=100 order by atm.atom"
	
	If dbConnGbl.State = 1 Then 'C.3-DB connection is not estblished, connection to DB is required.
		Set atmElmRS = Nothing
		Set atmElmRS = get_recordset_from_db_table (dbConnGbl,atomElementSQL)
		If IsEmpty(atmElmRs) = False Then
			rsCount = get_count_for_a_recordset (atmElmRS)
			currRowNum = rand_num_gen (Len(rsCount),1,rsCount)'get a random row num by generating a number between 1 and the total row count
'			atmElmRs.MoveFirst			
'			atmElmRs.MoveFirst
			atomCode = choose_recordset_values_on_rownum (atmElmRS,currRowNum,0)
			eventClass = choose_recordset_values_on_rownum (atmElmRS,currRowNum,2)
'			atmElmRs.MoveFirst
			codeSystem = choose_recordset_values_on_rownum (atmElmRS,currRowNum,3)
		End If
	End If 'C.3
	If IsEmpty(elementID) Then 'C.4 - If the element ID does not exist in db.
		get_atom_code_and_code_system = False
	ElseIf IsEmpty (atomCode) Then
		get_atom_code_and_code_system = False
	Else 
		get_atom_code_and_code_system = True
	End If	'C.4
'	atomCode ="ABCD12"
'	codeSystem = "ICD10"
End Function 
Rem ===================================================================================================================================================
Rem FunctionName: collect_and_translate_test_case_events_into_dmls
Rem FunctionParams: tcEventsStr (TC_EVENTS as read from the TC excel input),ByRef tcEventsArr,ByRef tcEventsDMLArr,memberID
Rem FunctionTasks: 'Function returns 2 arrays 1) filled with the TC EVENTS as given in TC Excel and 2) the DML queries for each corresponding events.
Rem This function calls 2 other functions, 1)break_events_in_logical_parts () and 2) build_dml_for_an_event()
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ====================================================================================================================================================
Function collect_and_translate_test_case_events_into_dmls (tcEventsStr,ByRef tcEventsArr,ByRef tcEventsDMLArr,memberID)	
	If InStr (tcEventsStr,",")>0 Then 'C.a-If the string has only 1 event (no ',' delimeter)
		tcEventsArr = Split(tcEventsStr,",")
		totalEventsCount = UBound(tcEventsArr)+1
		Else
		ReDim tcEventsArr (0)
		totalEventsCount = 1
		tcEventsArr (0) = tcEventsStr
	End If 'C.a	

	For e=1 To totalEventsCount
		currEvent = tcEventsArr (e-1)
		Call break_events_in_logical_parts (currEvent, eventSource, eventElement,eventType, eventTimeFrame) 'Call function to traslate a given event (CLAIM#1200#DIAGNOSIS#1M)
		currDML = build_dml_for_an_event (eventSource,eventType,eventElement,eventTimeFrame,memberID) 'Call function to create a DML for a given event
		tcEventsDMLArr (e-1) = currDML 'Fill in the DML container (Array)
	Next
	ReDim Preserve tcEventsDMLArr (totalEventsCount-1)
	'ReDim Preserve tcEventsDMLArr (totalEventsCount-1)
End Function 
Rem ===================================================================================================================================================
Rem FunctionName: break_events_in_logical_parts
Rem FunctionParams: currEventsStr, ByRef eventSource, ByRef eventElement, ByRef eventType, ByRef eventTimeFrame
Rem FunctionTasks: 'Function breaks the TCEVENTS which is '#' delimeted, and return them to the calling function via the ByRef parameters.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ====================================================================================================================================================
Function break_events_in_logical_parts (currEventsStr, ByRef eventSource, ByRef eventElement, ByRef eventType, ByRef eventTimeFrame)
	If InStr(currEventsStr,"#")>0 Then 'C.1
		currEventsArr = Split (currEventsStr,"#")
		eventSource = currEventsArr (0)
		eventElement = currEventsArr (1)
'		If InStr(eventElement,"~")>0 Then 'C.2-If the atom code is provided with element (ie,4088~V50.1)
'			eventElementArr = Split(eventElement,"~")
'			eventElement = eventElementArr(0)
'			eventAtom = eventElementArr(1)
'		End If 'C.2
		eventType = currEventsArr (2)
		eventTimeFrame = currEventsArr (3)
		Else 'C.1'
		appendTxt = "The specified events ("&currEventsStr&") is Invalid for DML conversion."
		Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
	End If 'C.1
End Function
Rem ===================================================================================================================================================
Rem FunctionName: build_dml_for_an_event ()
Rem FunctionParams: eventSource (CLAIMS/HIE),eventsStr,elementIDStr (elementID and/or atom for event),eventTimeFrame (timeframe for an event),currMemberID
Rem FunctionTasks: 'Function translates the Keyword formatted EVENTS into DML to be used in the seeding data into database tables.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ====================================================================================================================================================
Function build_dml_for_an_event (eventSource,eventsStr,elementIDStr,eventTimeFrame,currMemberID)	
	eventTimeFrameDate = create_a_date_in_different_formats (eventTimeFrame,"VBS") 'Call function to create a date for the event based on the timeframe value provided in TC.
	todayDate = Date ()
'	eventTimeFrameDate = calculate_days_back (eventTimeFrame,todayDate,"-") 'Call function to create a date for the event based on the timeframe value provided in TC.
	eventDaysBack = todayDate-eventTimeFrameDate
	eventDaysBackDML = "SYSDATE-"&eventDaysBack 'Find the time frame in SYSDATE-100 format which is equal to eventTimeFrame (ie,2M)
	If InStr(eventsStr,"~")>0 Then 'C.1-If additional colums other than the defaults are added in TC 
		eventTypeArr = Split (eventsStr,"~")
		eventType = eventTypeArr (0)
		For g = 1 To UBound(eventTypeArr)
			If InStr(eventTypeArr(g),"-")>0 Then
				additionalColsArr = Split (eventTypeArr(g),"-")
				eventsColsAddition = eventsColsAddition+additionalColsArr(0)&","
				currColValues = "'"&additionalColsArr(1)&"'"
				eventsValsAddition = eventsValsAddition+currColValues&","			
			End If			
		Next
		eventsColsAddition = get_rid_off_chars (eventsColsAddition,"LEFT",1,"")
		eventsValsAddition = get_rid_off_chars (eventsValsAddition,"LEFT",1,"")
		Else
		eventType = eventsStr
	End If 'C1
	
	doesElmAtmExist = get_atom_code_and_code_system (eventType,eventClass,elementIDStr,elementID,atomID,codeSystemName)
	'doesElmAtmExist =tRUE
	If CBool(doesElmAtmExist) = False Then 'C.0 - If the given element does not exist
		errMsg = "The given element ID ("&elementID&") or the atom code ("&atomID&")is not valid."
		build_dml_for_an_event = errMsg
		Exit Function 
	End If 'C.0

	userComments = "'"&codeSystemName&" code ("&atomID&") from "&eventType&" ELEMENT: "&elementID&"'" 'This comment will be added in the DML

	Select Case UCase(eventSource) 'S.1 'Switch case for event Types, CLAIMS,HIE,PDD etc
	Case "CLAIMS","CLAIM"
		defaultColumns = "CODESETTYPE,MEMBERID,SERVICEDT,PAIDDT,RECVDDT,RECORDINSERTDT,RECORDUPDTDT,INSERTEDBY,UPDTDBY,EXCLUSIONFLAG,BATCHID,COMMENTS"
'		defaultColValues = "'"&codeSystemName&"',"&currMemberID&",TO_DATE ('"&eventTimeFrameDate&"','MM/DD/YYYY'),TO_DATE ('"&eventTimeFrameDate&"','MM/DD/YYYY'),TO_DATE ('"&eventTimeFrameDate&"','MM/DD/YYYY'),"&"SYSDATE,SYSDATE,'"&currUserGbl&"','"&currUserGbl&"','IN',1,"&userComments
		defaultColValues = "'"&codeSystemName&"',"&currMemberID&","&eventDaysBackDML&","&eventDaysBackDML&","&eventDaysBackDML&","&"SYSDATE,SYSDATE,'"&currUserGbl&"','"&currUserGbl&"','IN',1,"&userComments
	
		Select Case (eventType) 'S.2 'The types claims that are allowed, ie, DIAGNOSIS
			Case "DIAG","DIAGNOSIS","DGS"
			tableName = "ODS.PATIENTMEDICALDIAGNOSIS"
			seqKey = "ods.ods_patientdiag_seq.nextval"
			extraDefaultColumns = "MEDICALDIAGINSTANCEID,MEDICALDIAGNOSISCODE"',MEMBERID,CODESETTYPE,SERVICEDT,PAIDDT,RECVDDT,RECORDINSERTDT,RECORDUPDTDT,INSERTEDBY,UPDTDBY,EXCLUSIONFLAG,BATCHID,COMMENTS"
			'extraDefaultColValues = seqKey&",'"&atomID&"',"'&currMemberID&",'"&codeSystemName&"',"&eventTimeFrameDate&","&eventTimeFrameDate&","&eventTimeFrameDate&","&"SYSDATE,SYSDATE,"&currUserGbl&","&currUserGbl&",'IN',1,'"&userComments&"'"
			Case "PROC","PROCEDURE","PDR"
			tableName = "ODS.PATIENTMEDICALPROCEDURE"
			seqKey = "ods.ods_patientproc_seq.nextval"
			extraDefaultColumns = "MEDICALPROCINSTANCEID,MEDICALPROCEDURECODE"'MEMBERID,,CODESETTYPE,SERVICEDT,PAIDDT,RECVDDT,RECORDINSERTDT,RECORDUPDTDT,INSERTEDBY,UPDTDBY,EXCLUSIONFLAG,BATCHID,COMMENTS"
			Case "DRUG","DRG","NDC"
			tableName = "ODS.PATIENTDRUGPRESCRIPTION"
			seqKey = "ods.ods_patientdrug_seq.nextval"
			extraDefaultColumns = "PRESCRIPTIONINSTANCEID,NDCCODE"
			Case "LAB","LABS","LOINC"
			defaultColumns = "CODESETTYPE,MEMBERID,SERVICEDT,RECVDDT,RECORDINSERTDT,RECORDUPDTDT,INSERTEDBY,UPDTDBY,EXCLUSIONFLAG,BATCHID,COMMENTS"
			defaultColValues = "'"&codeSystemName&"',"&currMemberID&","&eventDaysBackDML&","&eventDaysBackDML&","&"SYSDATE,SYSDATE,'"&currUserGbl&"','"&currUserGbl&"','IN',1,"&userComments
			tableName = "ODS.PATIENTLABRESULT"
			seqKey = "ods.ods_patientlab_seq.nextval"
			extraDefaultColumns = "LABRESULTINSTANCEID,LOINC"
		End Select 'S.2	
	Case "HIE","HDMS"
	End Select 'S.1
	
	extraDefaultColVals = seqKey&",'"&atomID&"',"
		
	If IsEmpty (eventsValsAddition) Then
	currDML = "INSERT INTO "&tableName&" ("&extraDefaultColumns&","&defaultColumns&") VALUES ("&extraDefaultColVals&defaultColValues&")"
	Else
	currDML = "INSERT INTO "&tableName&" ("&extraDefaultColumns&","&defaultColumns&","&eventsColsAddition&") VALUES ("&extraDefaultColVals&defaultColValues&","&eventsValsAddition&")"
	End If
	build_dml_for_an_event = currDML	
End Function
Rem ===================================================================================================================================================
Rem FunctionName: create_a_date_in_different_formats ()
Rem FunctionParams: currUnformattedVal (ie,24M+45D),dateFormat (ie, Oracle, VB etc).
Rem FunctionTasks: 'Function creates a date (a date in the past based on the timeframe (ie, 24M) in oracle format, ie: '19-JAN-2018'.
Rem This function calls another function, calculate_days_back () to create a date in the past.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ====================================================================================================================================================
Function create_a_date_in_different_formats (currUnformattedVal,dateFormat)
	If InStr(currUnformattedVal,"+")>0 Then 'C.1
		currValArr = Split (currUnformattedVal,"+")
		currValue = currValArr (0)
		additionalValue = currValArr (1)
		mathOperator2ndPart = "+"
		calcDate = calculate_days_back (currValue,Date(),"-")
		calcDateFinal = calculate_days_back (additionalValue,calcDate,mathOperator2ndPart)
		ElseIf InStr(currUnformattedVal,"-")>0 Then
		currValArr = Split (currUnformattedVal,"-")
		currValue = currValArr (0)
		additionalValue = currValArr (1)
		mathOperator2ndPart = "-"
		calcDate = calculate_days_back (currValue,Date(),"-")
		calcDateFinal = calculate_days_back (additionalValue,calcDate,mathOperator2ndPart)
		Else
		currValue = currUnformattedVal
		mathOperator2ndPart = "-"
		calcDateFinal = calculate_days_back (currValue,Date(),mathOperator2ndPart)
	End If 'C.1
	Select Case UCase(dateFormat)
		Case "ORA","ORACLE"
			If IsEmpty (calcDateFinal) = False Then 'C.2 - If the date is calculated correctly
				formattedDate = Day(calcDateFinal)&"-"&MonthName (Month(calcDateFinal),True)&"-"&Year(calcDateFinal)
				returnDate = UCase(formattedDate)
				Else
				returnDate = Empty 'Return empty if no date is calculated
			End If 'C.2	
			create_a_date_in_different_formats = returnDate
		Case "VBS","VB"
		create_a_date_in_different_formats = calcDateFinal 
	End Select
	
End Function
Rem ===================================================================================================================================================
Rem FunctionName: calculate_days_back ()
Rem FunctionParams: currValue (ie, 2M),fromDate (ie, SYSDATE, '11-JAN-2018),mathOperator (+ means subtract from the fromDate along with the currValue
Rem , 2M -> create a date 2 months in the past, 2M with with mathOperator value (-) means create a date 2 months in the future )
Rem FunctionTasks: 'Function creates a date (a date in the past based on the timeframe (ie, 24M) in oracle format, ie: '19-JAN-2018'.
Rem This function calls another function, calculate_days_back () to create a date in the past.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ====================================================================================================================================================
Function calculate_days_back (currValue,fromDate,mathOperator)
	If InStr(UCase(currValue),"M")>0 Or InStr(UCase(currValue),"Y") Or InStr(UCase(currValue),"D") Then 'C.1-If the age is in terms of Y (ie, 18Y) or months (ie, 24M)
		mathUnit = Right (currValue,1)
		unitNumber = Left(currValue,Len(currValue)-Len(mathUnit))
		ElseIf IsNumeric(currValue) Then
		unitNumber = currValue
		'calcDate = DateAdd("d",unitNumber,fromDate)	
		mathUnit = "d"
	End If 'C.1
	
	If mathOperator = "+" Then 'C.2
		unitNumber = unitNumber
		ElseIf mathOperator = "-" Then
		unitNumber = -1*unitNumber
	End If 'C.2
		
	Select Case UCase(mathUnit)
		Case "Y"
		calcDate = DateAdd("YYYY",unitNumber,fromDate)
		Case "M"
		calcDate = DateAdd("m",unitNumber,fromDate)
		Case "D"
		calcDate = DateAdd("d",unitNumber,fromDate)
	End Select

	calculate_days_back = calcDate
End Function

Rem ========================================================================================================================================
Rem FunctionName: get_rid_off_chars
Rem FunctionParams: stringPassed (string with all characters),leftRight(indicator to use Left/Right/Mid methods),numChars,startPos (is optional)
Rem FunctionTasks: 'Function returns a string with certain characters removed from it
Rem CreationDate: 2/16/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_rid_off_chars (stringPassed,leftRight,numChars,startPos)
	Select Case UCase(leftRight)
		Case "LEFT"
		get_rid_off_chars = Left(stringPassed,Len(stringPassed)-numChars)
		Case "RIGHT"
		get_rid_off_chars = Right(stringPassed,Len(stringPassed)-numChars)
		Case "MID"
		get_rid_off_chars = Mid(stringPassed,startPos,Len(stringPassed)-numChars)
	End Select
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: establish_a_database_connection ()
Rem FunctionParams: dbConnStr (the OLEDB connection string),ByRef dbConnLcl
Rem FunctionTasks: 'Function returns 0 if the database connection was succesful (and saves the connection in ByRef variable (dbConnLcl), else it returns the error number
Rem CreationDate: 9/16/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function establish_a_database_connection (dbConnStr,ByRef dbConnLcl)
	On Error Resume Next
	Set dbConn = CreateObject("ADODB.CONNECTION")
	If dbConnStr <> Empty And dbConn.State <> 1 Then 'C.1 - If the detailed DB information is retrieved and there's no active DB connection, then create a db connection using ADODB connection.
		dbConn.Open dbConnStr 'A direct connection to DB is opened		
		'On Error Resume Next
		If Err.Number <> 0 Then 'C.1.1 - If there's no error at DB connection
			returnMsg = "Error (#"&Err.Number&")- "&Err.Description&" while trying to establish a database connection."
			Else 'Set the current DB connection to be used globally
			Set dbConnLcl = dbConn
			returnMsg = "NO ERROR"
		End If 'C.1.1
	End If 'C.1
	establish_a_database_connection = returnMsg
End Function
Rem ========================================================================================================================================
Rem FunctionName: invoke_ce_realtime_service
Rem FunctionParams: strMemberID (MemberID),strSupplier (AHMSupplierID),strProduct ("" for all, or a specified product like DM),
Rem strSystem ("AA" to run a specified product, "CEUI" to run all eligible products).
Rem FunctionTasks: 'Function invokes the realtime CE REST service and returns 'True' if an action was created for the successful run otherwise returns 'False' 
Rem and captures the times (in CErunTimeStampAtStart and CErunTimeStampAtEnd variables) at the beginning and ending of the REST calls for using in DB validations.
Rem CreationDate: 2/8/2018
Rem CreatedBy: Mohammad Sarwar 
Rem UpdatedDate:2/9/2018
Rem ========================================================================================================================================
Function invoke_ce_realtime_service(strMemberID,strSupplier,strProduct,strSystem, ByRef CErunTimeStampAtStart, ByRef CErunTimeStampAtEnd)
	If CBool(runCEGbl) = False Then 'If the driver script was set to 'False' for not running CE service.
		logMsg = "The flag to run CE RT is set to FALSE in the Test Controller, hence CE Realtime service was NOT invoked."
		'Call write_to_notepad_log_file (logMsg)
		Call append_text_to_notepad_file (logFileDirGbl,"",logMsg)
		Exit Function 
	End If

	'Open the web service file from the given location (webServiceFilePath, coming from the config file) 
	Set ofs = CreateObject("scripting.filesystemobject")
    Set ofil = ofs.OpenTextFile(webServiceFilePath,1,true)'realTimeRunRequestPath
    strSoapRequest = Trim(ofil.ReadAll)
    Execute strSoapRequest            
    ofil.Close
            
    CErunTimeStampAtStart = "" 'clear the variable since it's global variables
    CErunTimeStampAtStart = Trim(oracle_format_time_stamp_up_to_min_plus_minus(2,"minus"))'call function to create time stamp in oracle format with 2 minutes back.
    StrXml = send_web_service_request_and_receive_response(wsUrlGbl, ssoaprequest) 'wsUrlGbl
    '*********** Verify the Response XML from Soap UI ***********            
    'Capturing the time when CE real time occured while calling CE RT run web service.
    CErunTimeStampAtEnd = "" 'clear the variable since it's global variables
    CErunTimeStampAtEnd = Trim(oracle_format_time_stamp_up_to_min_plus_minus(2,"plus"))'call function to create time stamp in oracle format with 2 minutes forward.
            'MsgBox CErunTimeStampAtEnd
    If InStr(1,Trim(CStr(StrXml)),Trim("productrun operationalproduct")) >0  Then
       	ceRun = True
    	logMsg = "RT SOAP call was made for  Member:"&strMemberID& ", with Supplier:" & strSupplier & " " & "Product:" & " " & strProduct & " " & "System:" & " " & strSystem 
        Else
        ceRun = False
        logMsg = "RT SOAP call was made and response was NOT successful, for  Member:"&strMemberID& ", with Supplier:" & strSupplier & " " & "Product:" & " " & strProduct & " " & "System:" & " " & strSystem
	End If
    Call append_text_to_notepad_file (logFileDirGbl,"",logMsg)
    invoke_ce_realtime_service = ceRun
End Function
Rem ========================================================================================================================================
Rem FunctionName: send_web_service_request_and_receive_response
Rem FunctionParams: strWSURL (url for the web service, coming from the environment file), strSoapRequest (the actual REST request built from
Rem inputs coming from the driver script (supplier & product) and the body coming from the web service path provided in the CONFIG file.
Rem FunctionTasks: 'Function sends a REST request and returns the response upon a successful response. 
Rem otherwise returns 'False'.
Rem CreationDate: 2/8/2018
Rem CreatedBy: Mohammad Sarwar 
Rem UpdatedDate:2/9/2018
Rem ========================================================================================================================================

Function send_web_service_request_and_receive_response(strWSURL, strSoapRequest)
	Dim i, oWinHttp, oXMlDoc, objNodeList
    Dim  sContentType, sSOAPRequest
    Dim MemberCEID, bValidate
    Dim strResponseStatus, strResponseStatusDesc, strbody
    Dim strGetLabResult, strLabTestName, strLabTestNumericResult
    Dim strServiceDate, strFeedSourceNm
  
  Set oWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")    
  'Web Service Content Type 
  sContentType ="application/soap+xml;charset=UTF-8"   
  'Open HTTP connection  
  oWinHttp.Open "POST", strWSURL, False   
  'Setting request headers  
  oWinHttp.setRequestHeader "Content-Type", sContentType  
  'MemberCEID = "75939889"
   MemberCEID = pstrMemberCEID  
  On error Resume Next
  'Send SOAP request 
  oWinHttp.Send  strSoapRequest 
  If Err.Number <> 0 Then
  	objResultsfile.WriteLine create_a_line_of_repeated_characters ("-",140) 'Call function to create a string with "*" of 100 times
  	logMsg = "Server may be down and the error description is: "& VBcrlf& Err.Description & " the error code is: "& Err.Number
  End If 
  'Get XML Response 
  send_web_service_request_and_receive_response = oWinHttp.ResponseText 
  Set oWinHttp = Nothing 
End Function
 
Rem ========================================================================================================================================
Rem FunctionName: oracle_format_time_stamp_up_to_min_plus_minus
Rem FunctionParams: minuteValue,plusMinus (1, plus to add a minute to time now and 1,Minus to subtract a minute from time now.
Rem FunctionTasks: 'Function to returns the time stamp in oracle format with a standdeviation of 1 minute, plus or minus
Rem CreationDate: 2/8/2018
Rem CreatedBy: Mohammad Sarwar 
Rem UpdatedDate:2/9/2018
Rem ========================================================================================================================================
Function oracle_format_time_stamp_up_to_min_plus_minus (minuteValue,plusMinus)
	Select Case UCase(plusMinus)
		Case "PLUS"
			timeDiff = CInt(minuteValue)
		Case "MINUS"
			timeDiff = (minuteValue)*(-1)
	End Select
	timeNow = DateAdd("n",timeDiff,Now())
	month_part = UCase(MonthName(Month(timeNow),True))
	date_part = Day(timeNow)
	If Len(date_part) =1 Then
		date_part = "0"&date_part
	End If
	year_part = Right(Year(timeNow),2)
	hour_part = Hour(timeNow)
	If Len(hour_part) = 1 Then
	hour_part = "0"&hour_part
	End If
	min_part = Minute(timeNow)
	If Len(min_part) = 1 Then
	min_part = "0"&min_part
	End If
'am_pm = Right(TimeValue(timeNow),2)
oracle_format_time_stamp_up_to_min_plus_minus = date_part&"-"&month_part&"-"&year_part&" "&hour_part&"."&min_part&".00.000000000 "'&am_pm
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: get_the_latest_member_run_id_from_csid_mrr_table
Rem FunctionParams: strMember
Rem FunctionTasks: 'Function returns the latest memberRecommendRunId from csid.memberrecommendrun
Rem CreationDate: 2/15/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_the_latest_member_run_id_from_csid_mrr_table (strMember,CErunTimeStampAtStart,CErunTimeStampAtEnd)
	'Get the latest runID from 'csid.memberrecommendrun' table
   	StrPPSql = "select mrr2.memberrecommendrunid from csid.memberrecommendrun mrr2 where mrr2.memberid = "&strMember&_
   	" and TO_TIMESTAMP(substr(mrr.recordinsertdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') between TO_TIMESTAMP('"&CErunTimeStampAtStart&"','DD-Mon-RR HH24:MI:SS.FF') "&_
	" and TO_TIMESTAMP('"&CErunTimeStampAtEnd&"','DD-Mon-RR HH24:MI:SS.FF')'"  			
  	Set oRs8 = FetchDataFromOracleDB(conn,StrPPSql) 'Call function to execute the query and return the result set
  	While Not oRs8.EOF 
  	memberRunId = oRs8.Fields("memberrecommendrunid").Value
  	oRs8.MoveNext
  	Wend
  	'MsgBox memberRunId
  	If Not IsNull(memberRunId) Then 
	  	get_the_latest_member_run_id_from_csid_mrr_table = memberRunId 
	  	Else get_the_latest_member_run_id_from_csid_mrr_table = "Null" 
  	End If
  	Set oRs8 = Nothing
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: get_df_information_for_medical_finding ()
Rem FunctionParams: mfID (Medical finding ID coming from the TC Excel), ByRef mfType (MedicalFindingTypeCd from database)
Rem FunctionTasks: 'Function returns the derivedfact for a given medicalfinding from database.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_df_information_for_medical_finding (mfID, ByRef mfType)
	dfSql = "SELECT df.DERIVEDFACTID,df.DERIVEDFACTTYPEID,mf.medicalfindingid,mf.medicalfindingnm Title,mf.medicalfindingtypecd,mft.MEDICALFINDINGTYPEDESC,mf.clinical_condition_cod,"&_
	" cond.CLINICAL_CONDITION_NAM ,mf.severitylevelcd FROM ce.derivedfact df,ce.medicalfinding mf,ce.medicalfindingtype mft,ce.clinical_condition cond WHERE df.DERIVEDFACTTYPEID = 1 "&_
	" and df.DERIVEDFACTTRACKINGID  = mf.medicalfindingid and mf.medicalfindingtypecd   = mft.medicalfindingtypecd and mf.clinical_condition_cod = cond.clinical_condition_cod (+)"&_
	" and mf.medicalfindingid in ("&mfID&")"
	
	Set dfRS = get_recordset_from_db_table (dbConnGbl,dfSql)
	totalDFs = get_count_for_a_recordset (dfRS)
	If totalDFs >=1 Then 
		dfRS.MoveFirst
		dfIDs = dfRS.Fields(0).Value
		mfType = dfRS.Fields(4).Value
	End If
	get_df_information_for_medical_finding = dfIDs
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: insert_members_into_tdm_tracker_table ()
Rem FunctionParams: regMemberID,aggMemberID,saveData (how many months data should be saved, ie 6)
Rem FunctionTasks: 'Function enters the records (members, both regular and aggregated) to the TDM table.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function insert_members_into_tdm_tracker_table (regMemberID,aggMemberID,saveData)
	memDtlSQL = "select mem.memberid,mem.ahmsupplierid,mem.primarymemberplanid,mem.personid,per.firstnm,per.middleinitial,per.lastnm,per.gender,per.ssn,per.dtofbirth "&_
	" from ods.member mem, ods.person per where mem.personid = per.personid and mem.memberid in ("&regMemberID&","&aggMemberID&")"
	Set tdmRS = get_recordset_from_db_table (dbConnGbl,memDtlSQL)
	While Not tdmRS.EOF
		memID = tdmRS.Fields(0).Value
		suppID = tdmRS.Fields(1).Value
		planID = tdmRS.Fields(2).Value	
		perID = tdmRS.Fields(3).Value
		fName = tdmRS.Fields(4).Value
		midInit = tdmRS.Fields(5).Value
		lName = tdmRS.Fields(6).Value
		gen = tdmRS.Fields(7).Value
		ssn = tdmRS.Fields(8).Value
		dob = tdmRS.Fields(9).Value
		tdmSQL ="INSERT INTO TDM.TDMMEMBER (MEMBERID,AHMSUPPLIERID,MEMBERPLANID,PERSONID,FIRSTNM,MIDINITAL,LASTNM,GENDER,SSN,DTOFBIRTH,ADDRLINE1,ADDRLINE2,"&_
		"CITY,STATE,ZIPCODE,EMAILADDR,PHONEFAXDISPLAYNUMBER,RECORDINSERTDT,INSERTEDBY,RECORDUPDTDT,UPDTDBY,SAVEDATAFORMONTHS,ODSDELETIONSTATUS,CSIDDELETIONSTATUS)"&_
		" VALUES ("&memID&","&suppID&","&planID&","&perID&",'"&fName&"','"&midInit&"','"&lName&"','"&gen&"',"&ssn&",TO_DATE('"&dob&"','MM/DD/YYYY'),Null,Null,Null,Null,Null,Null,Null,"&_
		"SYSDATE,'"&currUserGbl&"',SYSDATE,'"&currUserGbl&"',"&saveData&",'ACT','ACT')"
		errorCode = execute_dml_in_database (dbConnGbl,tdmSQL)
		If errorCode <> 0 Then 'C.a1 - If the Insert query to TDM.TDMMEMBER table failed.
			appendTxt = "MemberID ("&memID&") is not logged in TDM.TDMMEMBER table"
			Call append_text_to_notepad_file_without_borders_timestamp (logFileDirGbl, "",appendTxt)
		End If 'C.a1
		tdmRS.MoveNext
	Wend	
End Function
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: get_row_col_coordinate
Rem FunctionParams: coordidate (ie,1,3),delimeter(,), ByRef x, ByRef y
Rem FunctionTasks: 'Function returns the coordinate from a given string (ie, 3/5 or 3,5) and returns them in (x,y) as separate integer
Rem CreationDate: 3/2/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_row_col_coordinate (coordidate,delimeter, ByRef x, ByRef y)
coordidateArr = Split(coordidate,delimeter)
		x = CInt(coordidateArr(0))
		y = CInt(coordidateArr(1))
End Function
Rem ========================================================================================================================================
Rem FunctionName: write_to_excel_output_log ()
Rem FunctionParams: excelSheetPassed (excelObject for current sheet),cellValues (the values for multiple cells delimeted by a character delimeter (ie: 1,1;Val1|1,2;Val2|1,3;Val3,delimeter1 (is '|' in this example,
Rem delimeter2 (is ';' in this example) and ,totalColumns (number of collumns in the cellValues to check that the column count matches.
Rem FunctionTasks: 'Function writes the specified value (ie, tcid) into the specified cell (ie, cord1)
Rem CreationDate: 3/2/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function write_to_excel_output_log (excelSheetPassed,cellValues,delimeter1,delimeter2,totalColumns)
On Error Resume Next
	cellValuesArr = Split(cellValues,delimeter1)
	totalColsFromCellValuesArr = CInt(UBound(cellValuesArr))
	If CInt(totalColumns) = totalColsFromCellValuesArr+1 Then
		For i=0 To totalColsFromCellValuesArr
			currCellVal = cellValuesArr(i)
			'MsgBox currCellVal
			cordValuesArr = Split (currCellVal,delimeter2)
			cord = cordValuesArr (0)
			vals = cordValuesArr (1)
			If cord <> Empty Then 
				Call get_row_col_coordinate (cord,",",a,b)'coordinates are assumed to be in "," delimeted, ie:1,1
				'MsgBox i&"-"&j
				excelSheetPassed.Cells(a,b)= vals
			End If
		Next
	End If 	
End Function 
Rem ========================================================================================================================================
Rem ========================================================================================================================================
Rem FunctionName: find_test_case_range ()
Rem FunctionParams: memberFromTd,ByRef memberIDToUse,ByRef tcRange,ByRef tcLowerLim, ByRef tcUpperLim, ByRef randomTcSelection (get sets if TC range is in this form (1,2,3)
Rem FunctionTasks: 'Function is used to identify which sets ot TC should be run as specified in the driver script column (TC_RANGE)
Rem CreationDate: 3/2/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function find_test_case_range (ByRef tcRange,ByRef tcRangeArr,ByRef tcLowerLim, ByRef tcUpperLim, ByRef randomTcSelection, ByRef randSelectSingleTC)
'	If createTDMember = True Then 'C.a-If this flag is set to true, then member needed to be created and used as useTDMember = True
'	If memberSetHasMember = True Then 'Trim(memberFromTd) <> Empty Or Trim(memberFromTd) <> "" Then 'C.1 - If memberset is provided in the driver script instead of TC sheet.
'		memberIDToUse = ""
'		useTdMember = True 'If member is passed from MemberList from Test Driver, set this to 'True'
'		ElseIf memberFromTd <> "" And createTDMember = False Then
'		memberIDToUse = memberFromTd
'		useTdMember = True
'		Else 'C.1
'		useTdMember = False  
'	End If 'C.1	
	If tcRange = Empty Or tcRange = "" Then 'C.2
		tcRange = "ALL" 'to be used for notepad log
		tcLowerLim = 1
		ElseIf InStr(tcRange,",")>0 Then
		tcRange = sort_number_in_ascending_order(tcRange,",")
		randomTcSelection = True 'Set this variable to 'True' so that the execution flag in the TC excel sheet is ignored while running test cases from driver script.
		tcRangeArr = Split(tcRange,",")
		tcTotal = UBound(tcRangeArr)
		tcLowerLim = CInt(tcRangeArr(0))
		tcUpperLim = CInt(tcRangeArr(tcTotal))
		ElseIf InStr(tcRange,"-")>0 Then  
		tcRangeArr = Split(tcRange,"-")
		tcLowerLim = CInt(tcRangeArr(0))
		tcUpperLim = CInt(tcRangeArr(1))
		Else 'If the TCRange has a single number (ie, 2)
		randomTcSelection = True 'Set this variable to 'True' so that the execution flag in the TC excel sheet is ignored while running test cases from driver script.
		tcLowerLim = tcRange
		tcUpperLim = tcRange
		randSelectSingleTC = True 'If the range has a single TC (ie, 2 or 4, not 2,4)
	End If 'C.2
End Function 
Rem ========================================================================================================================================
Rem FunctionName: verify_if_file_exist (folderLoc,fileName,fileExt)
Rem FunctionParams: fileDirectPath (if the folderLoc value is the direct path of the file including the extension (True/False, 
Rem folderLoc (location of the file),fileName (file name),fileExt (file extension, ie, txt, xls etc)
Rem FunctionTasks: 'Function is used to verify a specific file exists in a given locaiton)
Rem CreationDate: 3/2/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function verify_if_file_exist (fileDirectPath,folderLoc,fileName,fileExt)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If fileDirectPath = True Then 'C.a - if the folderLoc value is the direct path of the file including the extension, this value is passed in as True/False otherwise.
		filePath = folderLoc
		Else
		If fileExt <> Empty Then
			filePath = folderLoc&"\"&fileName&"."&fileExt
			Else
			filePath = folderLoc&"\"&fileName
		End If 
	End If 'C.a
	
	Set objFolder = FSO.GetFolder(folderLoc)
	Set objFiles = objFolder.Files
	 
	For i=0 to objFiles.Count
	    If FSO.FileExists(filePath) Then 'C.b-If the file exists, return True
	        verify_if_file_exist = True
	        Exit Function
	        Else
	        verify_if_file_exist = False
	    End If 'C.b
	Next 	
	Set FSO = Nothing
	Set objFolder = Nothing
	Set objFiles = Nothing
End Function 
Rem ========================================================================================================================================
Rem FunctionName: get_member_recommend_runid_from_mrr_table (mrrMemberID,startTime,endTime)
Rem FunctionParams: mrrMemberID,startTime,endTime
Rem folderLoc (location of the file),fileName (file name),fileExt (file extension, ie, txt, xls etc)
Rem FunctionTasks: 'Function returns the current memberrecommendrunid from csid.memberrecommendrun for the current run defined by the timeframe
Rem between 'startTime' and 'endTime'.
Rem CreationDate: 3/2/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function get_member_recommend_runid_from_mrr_table (mrrMemberID,startTime,endTime)
	On Error Resume Next
	mrrSQL = "SELECT MRR.MEMBERRECOMMENDRUNID FROM csid.memberrecommendrun mrr WHERE mrr.memberid = "&mrrMemberID&" And "&_
	"TO_TIMESTAMP(substr(mrr.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') "&_
	"and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"	
	Set mrrRS = get_recordset_from_db_table (dbConnGbl,mrrSQL)
	If IsEmpty (mrrRS) Then 'C.1
		get_member_recommend_runid_from_mrr_table = 0
		appendTxt = "The member ("&mrrMemberID&") was NOT run between "&startTime&" and "&endTime&", as DB record in csid.memberrecommendrun was NOT present."
		Call append_text_to_notepad_file (logFileDirGbl, "",appendTxt)
		Else 'C.1'
		get_member_recommend_runid_from_mrr_table = mrrRS.Fields(0).Value
	End If 'C.1
End Function

Rem ========================================================================================================================================
Rem FunctionName: retrieve_all_derived_fact_ids_for_the_current_run ()
Rem FunctionParams: mrrID (MRR-ID passed from function for a given member CE run),matchDFs (if set to 'True' then will return the matching
Rem Derivedfact if fired, False will fetch all derived facts that fired),expectedDFs (the expected DF),ByRef actualDFsArr (array will be filled
Rem with the fired DFs if 'matchDFs' is set to 'False'
Rem FunctionTasks: 'Function returns the matching DFs (with expected) if DF triggered, returns 0 If the expected DF did not trigger, returns 
Rem an array filled with all DFs that fired if 'matchDFs' set to False.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function retrieve_all_derived_fact_ids_for_the_current_run (mrrID,matchDFs,expectedDFs,ByRef actualDFsArr)
	preventInfiniteRecursion = True
	If matchDFs = True Then 'c.1 - If this value is passed as True then use the passed DFs (expectedDFs) in the query as filter
		dfSQL = "select DERIVEDFACTID from csid.memberderivedfact mdf where mdf.memberrecommendrunid = "&mrrID&" and mdf.derivedfactid in ("&expectedDFs&")"
		Set dfRSLocal = get_recordset_from_db_table (dbConnGbl,dfSQL)
		If IsEmpty (dfRSLocal) = False Then 'C.1.a - If the records has some records
			While Not dfRSLocal.EOF 
				dfFired = dfRSLocal.Fields(0).Value
				dfRSLocal.MoveNext
			Wend
			retrieve_all_derived_fact_ids_for_the_current_run = dfFired
			If IsEmpty (dfFired) = False Then 'C.1.b
				appendTxt = "The expected DERIVED-FACT ("&expectedDFs&") was triggered at this run (MRR-RUNID:"&mrrID&")."
				Else
				appendTxt = "TC FAILURE REASON - The expected DERIVED-FACT ("&expectedDFs&") was NOT triggered at this run (MRR-RUNID:"&mrrID&")."
			End If 'C.1.b						
			Else
			retrieve_all_derived_fact_ids_for_the_current_run = 0
			appendTxt = "TC FAILURE REASON - The expected DERIVED-FACT ("&expectedDFs&") was NOT triggered at this run (MRR-RUNID:"&mrrID&")."
'			ReDim actualDFsArr (1000)  
'			Call retrieve_all_derived_fact_ids_for_the_current_run (mrrID,False,expectedDFs,actualDFsArr) 
		End If 'C.1.a
		Else 'C.1'		                          
		dfSQL = "select DERIVEDFACTID from csid.memberderivedfact mdf where mdf.memberrecommendrunid = "&mrrID
		Set dfRSLocal2 = get_recordset_from_db_table (dbConnGbl,dfSQL)
		recordCount = get_count_for_a_recordset (dfRSLocal2)
		dfRSLocal2.MoveFirst
		ReDim actualDFsArr (recordCount-1)
		Call fill_in_multi_dimensional_array_with_db_records (actualDFsArr,1,dfRSLocal2)
		ReDim Preserve actualDFsArr (recordCount-1)
	End If 'C.1 
	dtlAppendTxt = appendTxt&vbCrLf&"/* DF QUERY */"&vbTab&dfSQL
	Call append_text_to_notepad_file (logFileDirGbl,"",dtlAppendTxt)
End Function 
Rem ========================================================================================================================================
Rem FunctionName: fill_in_multi_dimensional_array_with_db_records ()
Rem FunctionParams: ByRef arrToFill (Predefined array),NumDimensions (number of dimensions that the array has),rsPassed (active record set
Rem that was created a query.
Rem FunctionTasks: 'Function fills an array with the values from the DB recordset in a predefined array.
Rem CreationDate: 9/24/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function fill_in_multi_dimensional_array_with_db_records (ByRef arrToFill,NumDimensions,rsPassed)
	Select Case NumDimensions
		Case 1
			p=0
			While Not rsPassed.EOF 
				arrToFill(p)= rsPassed.Fields(0).Value
				p=p+1
				rsPassed.MoveNext
			Wend
			ReDim Preserve arrToFill (p-1)
'			MsgBox UBound(arrToFill)
		Case 2
'		p=0
'		r=0
'		Set rsFields = rsPassed.Fields
'		fieldsCount = rsFields.Count
'		For p=0 To recordCount
'			If p=0 Then 'fill the first row with column names from record set
'			For r=0 To fieldsCount
'				arrToFill(p,r) = rsFields.Item(r).Name
'			Next
'			Else
			
'			End If
'		Next
	End Select	
End Function
Rem ========================================================================================================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: collect_tc_events_from_referred_sheet ()
Rem Fuction Arguments: currTcEvents (REFER~7~TC1,REFER~57~TC2,CLAIM#1771#DRUG#2M+15) ,tcFileLoc (TC file location),objXl,memberIDToUse
Rem Fuction tasks: Function collects test case events from referred TC (REFER~7~TC1) and converts the TC_EVENTS to DML and inserts them in DB
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
 Sub collect_tc_events_from_referred_sheet (currTcEvents,tcFileLoc,objXl,memberIDToUse) 
	If InStr (currTcEvents,",")= 0 Then 'C.x-If the TC_EVENTS column has more than 1 event specified (',' delimeted)
		Dim referTCArr (0)
		referTCArr (0) = currTcEvents
		totalRuleIDs = 0
		Else 'C.x
		referTCArr = Split(currTcEvents,",")
		totalRuleIDs = UBound(referTCArr)
	End If 'C.x
					
	For g=0 To totalRuleIDs
		ReDim referTcDMLArr (100)
		If InStr(referTCArr(g),"#")>0 Then 'C.g.1 - If the events are specified in the test case
			Call collect_and_translate_test_case_events_into_dmls (referTCArr(g),tcEventsArr,referTcDMLArr,memberIDToUse)
			ElseIf InStr(referTCArr(g),"REF")>0 Then 'C.g.1' - If the 'REFER'/'REF' keyword, then open the corresponding sheet to collect the events for TC
				referTCSectionArr = Split(referTCArr(g),"~")
				excelSheetToOpen = referTCSectionArr (1)
				tcEventsToBeCopiedFromTC = referTCSectionArr (2)
				openExtExcel = load_specified_excel_by_sheet_name (tcFileLoc,objXl,excelSheetToOpen,excelBookRF,excelSheetRF, True)'Call function to load excel
'								MsgBox excelSheetRF.UsedRange.Rows.Count
				tcIDColNum = get_column_id_from_excel_matching_a_column_name (excelSheetRF,1,"TCID")
				tcEventsIDRF = get_column_id_from_excel_matching_a_column_name (excelSheetRF,1,"TC_EVENTS")
				tcDmlIDRF = get_column_id_from_excel_matching_a_column_name (excelSheetRF,1,"DML")								
				rowWithEvents = get_row_num_from_excel_column_matching_a_key_value (tcFileLoc,tcIDColNum,tcEventsToBeCopiedFromTC,excelSheetRF,"TCID")
				tcEventsRF = get_cell_value_given_rowid_columnid (excelSheetRF,rowWithEvents,tcEventsIDRF)
				Call collect_and_translate_test_case_events_into_dmls (tcEventsRF,tcEventsArr,referTcDMLArr,memberIDToUse)
				Call execute_dml_from_an_array_of_dmls (tcEventsArr,referTcDMLArr,False)

				excelBookRF.Close 'Close the excel book if opened.						
				Else 'C.g.1 - If the events are specified in the test case (ELSE)
				referMultiTCArr = Split(referTCArr(g),"~")							
		End If 'C.g.1
	Next	
 End Sub
 Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: execute_dml_from_an_array_of_dmls ()
Rem Fuction Arguments: tcEventsArr (arry containing all TC EVENTS) ,tcEventsDMLArr (tcEventsArr corresponding DMLS in this array)
Rem tcDML (False = DMLs that are built out of TC_EVENTS, True = DMLs that are provided in DML colum ot TC).
Rem Fuction tasks: Function inserts distinct DMLs in DB from an array containing DMLs.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
 Sub execute_dml_from_an_array_of_dmls (tcEventsArr,tcEventsDMLArr,tcDML) 
 	For y = 0 To UBound(tcEventsDMLArr)	
 		If tcDML = True Then 'C.1
	 		logMsg = "/* Additonal"
	 		Else
	 		logMsg = "/* "&tcEventsArr(y)
	 	End If 'C.1				
		insertSuccess =  execute_dml_in_database (dbConnGbl,tcEventsDMLArr(y))
		If insertSuccess = 0 Then 'C.0
			dmlCounter = dmlCounter+1
			appendTxt = logMsg&" - DML is successully executed and data is inserted in DB */"&vbCrLf&tcEventsDMLArr(y)&";"
			Else
			appendTxt = logMsg&" - DML is NOT successully executed and data is not inserted in DB */"&vbCrLf&tcEventsDMLArr(y)&vbCrLf&_
			"The reported DB error is - "&insertSuccess&"."
		End If	 'C.0				
	
		Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)					
	Next
End Sub         
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++========================================================= 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_members_for_memberset ()
Rem Fuction Arguments: ByRef currMemSetArrFull,numberOfMembers,supplierID,testEnv,tcDetailedLog
Rem Fuction tasks: Function creates n (numberOfMembers) number of members and fills array (currMemSetArrFull) with memberID.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Sub create_members_for_memberset (ByRef currMemSetArrFull,numberOfMembers,supplierID,testEnv,memberInfoLogFile)
	ReDim currMemSetArrFull(numberOfMembers)	
	For v=0 To numberOfMembers
		If createMemInfoNotepad = True Then 'C.1 - If the flag (createMemInfoNotepad) is set to True in config file for creating notepad file.
			appendTxt = "Creating new member, #"&v+1 '&vbcrlf
			Call append_text_to_notepad_file_without_borders_timestamp (memberInfoLogFile,"",appendTxt)
		End If 'C.1
		memberCreated = create_a_member_for_tc (currMemberDemo,supplierID,testEnv,memberInfoLogFile,aggMemberID)
		If IsEmpty(memberCreated) = False Then 'C.2-If the member was successfully created.
			currMemSetArrFull(v) = memberCreated
		End If 'C.2
	Next
End Sub	 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: connect_to_a_database_for_a_given_env ()
Rem Fuction Arguments: retrievedDbInfo (boolean value, True = db info collected),dbHost,dbSid,dbPort,dbUser,dbPassword
Rem Fuction tasks: Function establishes a database connection for a given environment with the access credentials passed in to the function.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function connect_to_a_database_for_a_given_env (retrievedDbInfo,dbHost,dbSid,dbPort,dbUser,dbPassword, ByRef dbConnGbl, ByRef currUserGbl)
	If retrievedDbInfo = True Then 'C.a1 - If DB info was successfully retrieved from the given excel location.
		'Check to see if auth key is provided with the DB password
		If InStr(dbPassword,"|") >0 Then 'C.b1 - If the auth-key is provided with the decrypted password (ie, abcd123|MS40405
			dbPasswordArr = Split (dbPassword,"|")
			encryptedPassword = dbPasswordArr (0)
			userAuthKey = dbPasswordArr (1)
			Else 'C.b1
			encryptedPassword = dbPassword
			userAuthKey = authUserGbl
			'Call function to decrypt the password
			decryptedPassword = decrypt_val (encryptedPassword)
			'Call function to authenticate the user passKey
			userAuthResult = validate_passkey (userAuthKey,encryptedPassword)
			If userAuthResult = False Then 'C.1.2.1.1 - If the user authentication failed, then abort the execution.
				returnMsg = "This user auth-key ("&userAuthKey&" or the encrypted password ("&encryptedPassword&" is invalid, hence ARTT execution is aborted."
				currUserGbl = userAuthKey
				exitArtt = True 'Set the variable to true to abort ARTT. 		
				Else 'Continue the execution to establish DB connection.
				currUserGbl = userAuthKey 
				'Set db connection
				dbConnStr = create_database_connection_string (dbHost,dbSid,dbPort,dbUser,decryptedPassword)
				dbConnStrGbl = dbConnStr
				connectionSuccess = establish_a_database_connection (dbConnStrGbl,dbConnLcl)
				If InStr (UCase(connectionSuccess),"NO ERROR")> 0 Then 'C.z
					Set dbConnGbl = dbConnLcl
					returnMsg = connectionSuccess
					exitArtt = False 'Set the variable to False to continue to execute using ARTT.
					Else
					dbConnStr_2 = create_database_connection_string_ver2(dbHost,dbSid,dbPort,dbUser,decryptedPassword)
					dbConnStrGbl = dbConnStr_2
					connectionSuccess = establish_a_database_connection (dbConnStrGbl,dbConnLcl_2)
					If InStr (UCase(connectionSuccess),"NO ERROR")> 0 Then 'C.y - If the DB connection was successful
						Set dbConnGbl = dbConnLcl_2
						exitArtt = False 'Set the variable to False to continue to execute using ARTT.
						returnMsg = connectionSuccess
						Else
						appendTxt = "Database connection has failed, hence ARTT is aborted. The Database Error is :"&vbCrLf&connectionSuccess
						Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
						returnMsg = appendTxt
						exitArtt = True 'Set the variable to true to abort ARTT.
'						Exit Function 'Get out of the main script by exiting the 'For' loop.
					End If 'C.y
				End If 'C.z
			End If 'C.1.2.1.1
		End If 'C.b1
		Else 'C.1.2		
	End If 'C.a1
	connect_to_a_database_for_a_given_env = returnMsg 
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++========================================================= 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: array_elements_to_string_conversion ()
Rem Fuction Arguments: arrPassed (an array (must be one dimensional) containing n number of elements),delimeterToUse (delimeter to use in building a string, ie, ",")
Rem Fuction tasks: Function builds a string where each array element is separted by delimeterToUse and returns it to the caller, returns string ("Empty Array was passed in.")
Rem if passed in array had no element in it.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function array_elements_to_string_conversion (arrPassed,delimeterToUse)	
	If IsEmpty(arrPassed) Then
		strBuiltRefined = "Empty Array was passed in."
		Else
		totalElements = UBound(arrPassed)
		strBuilt = ""
		For i=0 To totalElements
			strBuilt = strBuilt&arrPassed(i)&delimeterToUse
		Next
		strBuiltRefined =  get_rid_off_chars (strBuilt,"LEFT",1,"")
	End If 	
	array_elements_to_string_conversion = strBuiltRefined
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: get_operational_product_and_system ()
Rem Fuction Arguments: runProduct (passed in from the driver script),ByRef runProductGbl (ie, "" if this value (runProduct) is null or "ALL"), ByRef systemNameGbl (ie, AA)
Rem Fuction tasks: Function returns the system name and product to be used in CE real time run.
Rem if passed in array had no element in it.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Sub get_operational_product_and_system (runProduct,ByRef runProductGbl, ByRef systemNameGbl)
	If UCase (runProduct) = "ALL" Or IsEmpty (runProduct) = True Then 'C.1.1 - If the runProduct column in driver script is set to ALL or Empty then set the variable to run all eligible products.
			runProductGbl = ""
			systemNameGbl = "CEUI"
			Else
			runProductGbl = runProduct
			systemNameGbl = "AA"
	End If 'C.1.1
End Sub
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Sub name: get_number_of_membersets ()
Rem Sub Arguments: memberSet (string coming from TD, ie, MEMBERSET1-MEMBERSET5,MEMBERSET5), ByRef memberSetArr, ByRef useTdMember (flag that denotes
Rem whether the member set is used to create/use members, set to True if it is otherwise false).
Rem Sub tasks: Function returns array filled with the number of membersets to used and flag as True if memberset is used otherwise false.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Sub get_number_of_membersets (memberSet, ByRef memberSetArr, ByRef useTdMember )
	If InStr(memberSet,"MEMBER")>0  Then 'C.a.1 - If the column has this string ("MEMBER") Rem Previous condition - And IsEmpty(memberSet) = False
		useTdMember = True
		If InStr(memberSet,",")>0 Then 'C.a.2
			memberSetArr = Split (memberSet,",")
			ElseIf InStr(memberSet,"-")>0 Then 'C.1.1.1 - 
			memberSetRange = Split (memberSet,"-")
			memberSetFirst = Right(memberSetRange (0),Len(memberSetRange (0))-9)
			memberSetLast = Right(memberSetRange (1),Len(memberSetRange (1))-9)
			ReDim memSetArr (memberSetLast-1)
			For a=0 To memberSetLast-1
				memberSetArr (a) = "MEMBERSET"&a+1				
				Next 				
				Else 'C.a.2'
				useTdMember = True
 				ReDim memberSetArr (0)
				memberSetArr (0) = 	memberSet		
		End If 'C.a.2
		Else 'C.a.1 - If the cell is empty
		useTdMember = False
	End If 'C.a.1
End Sub 
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_excel_output_file_for_rule_id ()
Rem Fuction Arguments: excelOutputFile (directory of the excel file location),currXlOutputFile (name of the excel file),tcExcelOutputType (the type of the file, NEW/COPY)
Rem Fuction tasks: Function returns array filled with the number of membersets to used and flag as True if memberset is used otherwise false.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_excel_output_file_for_rule_id (excelOutputFile,currXlOutputFile,tcExcelOutputType)
	fileAlreadyExists = verify_if_file_exist (True,excelOutputFile,currXlOutputFile,"")
	If fileAlreadyExists <> True Then 'C.b.2 - If the output file is not created already, then create with header.
		If CStr(UCase(tcExcelOutputType)) = "NEW" Then 'C.b.1.a - If the excel ouput is a type of new (modified from input)
			columnStr = "TCID/MEMBERID/EVALUATION/RESULTS/COMMENTS"
			ElseIf CStr(UCase(tcExcelOutputType)) = "COPY" Then 'C.b.1.b - If the excel ouput is a copy of input file 
			columnStr = "EXECUTE/TCID/RELATED_TCID/TC_DESCRIPTION/MEMBERID/MEMBER_DEMOGRAPHICS/EVALUATION/TC_EVENTS/DML/PURGE_DATA/RESULTS"
		End If 'C.b.1.a		
		Call create_excel_output_file (exelObject,excelBook,excelSheet)	'Create Excel output file for this rule-ID.		
		Call create_header_for_excel_file (excelLogDir,excelBook,excelSheet,1,"OUTPUT",columnStr,"/")
	End If 'C.b.2
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Fuction name: create_member_info_excel ()
Rem Fuction Arguments: ruleID,ruleCat,notePadLogFolder,tcSummaryLogFolder,excelLogDirGbl,ByRef memberInfoExcelGbl (the directory with file name for the
Rem excel file with member info,ByRef memberInfoLogFileGbl (the directory of the notepad log file)
Rem Fuction tasks: Function creates an excel file with column names hard coded in this function.
Rem Creation Date: 6/1/2018
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function create_member_info_excel (ruleID,ruleCat,notePadLogFolder,tcSummaryLogFolder,excelLogDirGbl,ByRef memberInfoExcelGbl,ByRef memberInfoLogFileGbl)	
	memInfoFileName = "MEMBER_INFO_"&ruleID&"_"&ruleCat&"_"&Replace(Date,"/","_")&Space(1)& Replace(Time,":","_")
	memberInfoExcelGbl = memInfoFileName&".xlsx"
	Call create_text_file (notePadLogFolder,tcSummaryLogFolder,memInfoFileName)
	memberInfoLogFileGbl = notePadLogFolder&tcSummaryLogFolder&memInfoFileName&".txt"
	fileAlreadyExists = verify_if_file_exist (True,excelOutputFile,memberInfoExcelGbl,"")
	If fileAlreadyExists <> True Then 'C.b.2 - If the output file is not created already, then create with header.
		If CStr(UCase(memberInfoType)) = "LONG" Then 'C.b.1.a - If the excel ouput is a type of new (modified from input)
			columnStr = "SUPPLIERID|MEMBERID|MEMBERPLANID|SOURCEPATIENTID|DTOFBIRTH|GENDER|PERSONID|MEMBERTYPECODE|FIRSTNAME|LASTNAME|TCID"
			ElseIf CStr(UCase(memberInfoType)) = "SHORT" Then 'C.b.1.b - If the excel ouput is a copy of input file 
			columnStr = "SUPPLIERID|MEMBERID|MEMBERPLANID|TCID"
		End If 'C.b.1.a					
		Call create_excel_output_file (exelObject,excelBook,excelSheet)	'	Create Excel output file for member info		
		Call create_header_for_excel_file (excelLogDirGbl&memberInfoExcelGbl,excelBook,excelSheet,1,"MEMBER_INFO",columnStr,"|")
	End If 'C.b.2
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Function name: color_code_excel_cell
Rem Function Arguments: objSheet,xVal,yVal,colorNumber (4=Green), fontColor (1=Black)
Rem Function tasks: fucntion to color code a specific excel cell with a given color (colorNumber).
Rem Creation date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function color_code_excel_cell(objSheet,xVal,yVal,colorNumber,fontColor)
	Set rng2 = objSheet.Range(objSheet.Cells(xVal,yVal),objSheet.Cells(xVal,yVal))
	With rng2
    .Interior.ColorIndex = colorNumber
    .Font.Bold = True
    .Font.ColorIndex = fontColor
    End With
End Function
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Rem Function name: rule_category__csid_validation ()
Rem Function Arguments: ruleCategory,memberId,memberRunId,stateComponentId
Rem Function tasks: fucntion to color code a specific excel cell with a given color (colorNumber).
Rem Creation date: 11/25/2017
Rem ===================================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++=========================================================
Function rule_category_csid_validation (ruleCategory,memberId,memberRunId,derivedFactID,productCode)
	stateComponentId = get_statecomponent_id (derivedFactID)
	Select Case UCase(ruleCat)
		Case "MK_VAL","MKVAL","CONDVAL","COND_VAL"
			'Call function to validate MHS
		validationResult = validate_mhs_for_a_given_ce_run (memberId,memberRunId,stateComponentId,productCode,derivedFactID)
		Case "MKSEV","MK_SEV"
		Case "ME_VAL","MEVAL"
		Case "ME_SEV","MESEV"
		Case "PPVAL","PP_VAL"
		Case "PIVAL","PI_VAL"
		Case "PPMOD","PP_MOD"
		Case "CCVAL","CC_VAL"
		Case Else
	End Select
	rule_category_csid_validation = validationResult
End Function 
Function get_statecomponent_id (derivedFact)
	scDfSQL = "select href.statecomponentid,href.derivedfactid,href.statetypecd from ce.healthstatederivedfactxref href "&_
				"where href.derivedfactid = "&derivedFact
	Set scdfRS = get_recordset_from_db_table (dbConnGbl,scDfSQL)
	scdfCount = get_count_for_a_recordset (scdfRS)
	If scdfCount >=1 Then
	scdfRS.MoveFirst
	stateCompID = scdfRS.Fields(0).Value
	Else
	stateCompID = 0
	End If
	get_statecomponent_id = stateCompID
End Function 
Function validate_mhs_for_a_given_ce_run (memberId,memberRunId,stateComponentId,productCode,derivedFactID)
	If productCode <> "" Then 'C.a
		If productCode = "ALL" Then 'C.a1
			productFilter = ""
			Else
			productFilter = "and cerma.productmnemoniccd In('"&productCode&"')"
		End If 'C.a1
		Else 'C.a'
		productFilter = ""
	End If 'C.a
	mhsQuery = "select  mrr.memberid,mrr.memberrecommendrunid,mhs.memberhealthstateskey mhs_skey,cerma.careenginerunmemberactionid actionid,"&_
	"cerma.productmnemoniccd program_name,cerma.recommendflg,cerma.programreferralintensitycd program_intensity, mhs.statecomponentid,"&_
	"cerma.overallscorenbr,mhs.severitylevel mhs_severity,mhs.healthstatestatuscd mhs_status,mhs.healthstatestatuschangedt mhs_status_change_dt,"&_
	"cerma.recordinsertdt action_creation_dt, cerma.recordupdtdt action_update_dt,mhs.lastevaluationdt mhs_last_eval_dt"&_
	" from csid.memberrecommendrun mrr, csid.careenginerunmemberaction cerma, csid.memberhealthstateactionxref actxref, csid.memberhealthstate mhs"&_
	" where cerma.careenginerunmemberactionid = actxref.careenginerunmemberactionid and actxref.memberhealthstateskey=mhs.memberhealthstateskey "&_
	"and mrr.memberrecommendrunid = cerma.memberrecommendrunid and mrr.memberrecommendrunid in ("&memberRunId&")"&_
	"and mhs.statecomponentid in ("&stateComponentId&")and mhs.yearqtr = '"&DatePart("yyyy",Date)&DatePart("q",Date)&"'"&productFilter
	
	Set mhsRS = get_recordset_from_db_table (dbConnGbl,mhsQuery)
	mhsCount = get_count_for_a_recordset (mhsRS)
	If mhsCount >=1 Then 
		fnRetMsg = "PASS"
		retMsg = "MHS - The CSID validation PASSED since the corresponding STATECOMPONENT-ID ("&stateComponentId&") for the given"&_
		" MEDICAL-FINDINGID/DERIVED-FACTID ("&derivedFactID&") has triggered at this run (RUN-ID:"&memberRunId&")."  
		Else
		fnRetMsg = "FAIL"
		retMsg = "TC FAILURE REASON : MHS - The CSID validation FAILED since the given STATECOMPONENTID ("&stateComponentId&") has NOT triggered at this run (RUN-ID:"&memberRunId&")."  
	End If
	appendTxt = retMsg&vbCrLf&"/* The following query was run for MHS validation. */"&mhsQuery
	Call append_text_to_notepad_file (logFileDirGbl,"",appendTxt)
	validate_mhs_for_a_given_ce_run = fnRetMsg
End Function
Rem ========================================================================================================================================
Rem FunctionName: fetch_cerma_validation_query
Rem FunctionParams: strMember (memberID),productCode (product name, ie. DM),memberRunId,startTime,endTime,queryIndicator (whether to use Query#1
Rem Query#2, Query#3 etc.)
Rem FunctionTasks: 'Function returns a string with certain characters removed from it
Rem CreationDate: 2/16/2018
Rem CreatedBy: Mohammad Sarwar 
Rem ========================================================================================================================================
Function fetch_cerma_validation_query (strMember,productCode,memberRunId,startTime,endTime,queryIndicator)
	Select Case Cint(queryIndicator)
	Case 1	'CERMA query#1 is used to validate whether a program placement product (ie, DM) fired with a successful action.
	cermaProductVal=  "select mhs.memberid,cerma.careenginerunmemberactionid,cerma.productmnemoniccd,cerma.recommendflg,"&_
			    "cerma.programreferralintensitycd,mhs.statecomponentid,substr(cerma.recordupdtdt,1,28) actionUpdatedDate "&_
				"from csid.careenginerunmemberaction cerma,csid.memberhealthstateactionxref mhsxref,csid.memberhealthstate mhs "&_
				"where mhs.memberid in ("& strMember &") and cerma.memberrecommendrunid = "& memberRunId & _
				" and cerma.productmnemoniccd in ('"& productCode & "') and cerma.careenginerunmemberactionid=mhsxref.careenginerunmemberactionid"& _        	 
				" and mhsxref.memberhealthstateskey=mhs.memberhealthstateskey and Upper(cerma.recommendflg) ='Y' and mhs.healthstatestatuscd = 'CURR'"&_
				" and TO_TIMESTAMP(substr(cerma.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') "&_
		        "between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"
	Case 2 'CERMA query#2 is used to validate whether a non placement product (ie, MHP) fired with a successful action.
	cermaProductVal= "select mhs.memberid,cerma.careenginerunmemberactionid,cerma.productmnemoniccd,cerma.recommendflg,"&_
			    "cerma.programreferralintensitycd,mhs.statecomponentid,substr(cerma.recordupdtdt,1,28) actionUpdatedDate "&_
				"from csid.careenginerunmemberaction cerma,csid.memberhealthstateactionxref mhsxref,csid.memberhealthstate mhs "&_
				"where mhs.memberid in ("& strMember &") and cerma.memberrecommendrunid = "& memberRunId & _
				" and cerma.productmnemoniccd In('"& productCode & "') and cerma.careenginerunmemberactionid=mhsxref.careenginerunmemberactionid"& _        	 
				" and mhsxref.memberhealthstateskey=mhs.memberhealthstateskey and Upper(cerma.recommendflg) ='N' and mhs.healthstatestatuscd = 'CURR'"&_
				" and TO_TIMESTAMP(substr(cerma.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') "&_
		        "between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"
	Case 3 'CERMA query#3 is used to validate whether a placement product (ie, DM) fired with a successful action where there's no related marker.
	cermaProductVal = "select mrr.memberid,cerma.careenginerunmemberactionid,cerma.productmnemoniccd,cerma.recommendflg,"& _
	        "cerma.programreferralintensitycd,substr(mrr.recordupdtdt,1,28) actionUpdatedDate from csid.memberrecommendrun mrr inner join csid.careenginerunmemberaction cerma on "& _
	        "cerma.memberrecommendrunid = mrr.memberrecommendrunid and mrr.memberid = "&strMember&" and Upper(cerma.recommendflg) = 'Y' and cerma.productmnemoniccd "&_
	        " in ('"&productCode&"')And mrr.memberrecommendrunid = "&memberRunId&" and TO_TIMESTAMP(substr(cerma.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') "&_
	        "between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"	
	Case 4 'CERMA query#3 is used to validate whether a placement product (ie, DM) fired with a successful action where there's no related marker.
	cermaProductVal = "select mrr.memberid,cerma.careenginerunmemberactionid,cerma.productmnemoniccd,cerma.recommendflg,"& _
	        "cerma.programreferralintensitycd,substr(mrr.recordupdtdt,1,28) actionUpdatedDate from csid.memberrecommendrun mrr inner join csid.careenginerunmemberaction cerma on "& _
	        "cerma.memberrecommendrunid = mrr.memberrecommendrunid and mrr.memberid = "&strMember&" and Upper(cerma.recommendflg) = 'N' and cerma.productmnemoniccd "&_
	        " in ('"&productCode&"')And mrr.memberrecommendrunid = "&memberRunId&" and TO_TIMESTAMP(substr(cerma.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') "&_
	        "between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"	
	Case 5	'CERMA query#5 without the timestamp \\DO NOT USE
	cermaProductVal=  "select mhs.memberid,cerma.careenginerunmemberactionid,cerma.productmnemoniccd,cerma.recommendflg,"&_
			    "cerma.programreferralintensitycd,mhs.statecomponentid,substr(cerma.recordupdtdt,1,28) actionUpdatedDate "&_
				"from csid.careenginerunmemberaction cerma,csid.memberhealthstateactionxref mhsxref,csid.memberhealthstate mhs "&_
				"where cerma.memberrecommendrunid = "& memberRunId & _
				" and cerma.productmnemoniccd in ('"& productCode & "') and cerma.careenginerunmemberactionid=mhsxref.careenginerunmemberactionid"& _        	 
				" and mhsxref.memberhealthstateskey=mhs.memberhealthstateskey and Upper(cerma.recommendflg) ='Y' and mhs.healthstatestatuscd = 'CURR'" '&_
				Rem " and TO_TIMESTAMP(substr(cerma.recordupdtdt,1,28),'DD-Mon-RR HH12:MI:SS.FF PM') "&_
		        Rem "between TO_TIMESTAMP('"&startTime&"','DD-Mon-RR HH24:MI:SS.FF') and TO_TIMESTAMP('"&endTime&"','DD-Mon-RR HH24:MI:SS.FF')"
	End Select
	fetch_cerma_validation_query = cermaProductVal
End Function


