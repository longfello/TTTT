<!--#include virtual="/core/classes/aspTemplate.asp" -->
<!--#include virtual="/core/classes/error.asp" -->
<!--#include virtual="/core/modules/common.asp" -->
<%
'---------------------------------------------------------------------
'INITIALIZATIONS
'---------------------------------------------------------------------
set tIn  = new aspTemplate
set tOut = new aspTemplate
set e    = new error

tIn.setTemplatesDir "/admin/templates/events/"
tOut.setTemplatesDir "/core/templates/admin/"
'---------------------------------------------------------------------
'CONTROL STRUCTURE
'---------------------------------------------------------------------
select case request("a")
	case ""
		call content_listEvents()
	case "addEditEvents"
		call content_addEditEvents()
	case "addEditEvents_finish"
		call parse_addEditEvents()
	case "deleteEvents"
		call content_deleteEvents()
	case "deleteEvents_finish"
		call parse_deleteEvents()	
end select
'---------------------------------------------------------------------
'GARBAGE
'---------------------------------------------------------------------
set t = nothing
'---------------------------------------------------------------------
'MODULES
'---------------------------------------------------------------------
sub content_listEvents()
	with tIn
		.setTemplateFile "listEvents.html"
		.updateBlock     "no_data"
		.updateBlock     "data"
		
		db_open(Application("ConnectionString"))
		set Events = db_query("SELECT * FROM events ORDER BY sDate")
		
		if(Events.eof = true) then
			.parseBlock "no_data"
		end if
		
		do until Events.eof
			.setVariable "TITLE"	, Events("title")
			.setVariable "DES"		, Events("desc")
			.setVariable "TIME"		, Events("sTime")
			.setVariable "LOC"		, Events("location")
			
			.setVariable "EDIT_LINK",	request("SCRIPT_NAME") & "?a=addEditEvents&eid=" & Events("eid")
			.setVariable "DELETE_LINK",	request("SCRIPT_NAME") & "?a=deleteEvents&eid=" & Events("eid")
			
			.parseBlock "data"
			Events.movenext
		loop
		db_close
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock     "warning"
		
		if e.hasErrors then
			.setVariable "WARNING",	e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable "FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable "HEADING"		,	"Event Listings"
		.setVariable "COL1"			,	"Event Listings"
		.setVariable "BODY1"		,	tIn.getOutput
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION", "navigation.html"
		
		.parse
	end with
end sub

sub content_addEditEvents()
	with tIn
		.setTemplateFile "addEditEvents.html"
		if(request("eid") <> "") then
			isUpdate = true
		else
			isNew = true
		end if
		
		if isNew then
			title 		= request("title")
			des 		= request("des")
			sDate  		= request("sDate")
			eDate		= request("eDate")
			sTime		= request("sTime")
			eTime 		= request("eTime")
			loc			= request("loc")
		else
			db_open(Application("ConnectionString"))
			set Events = db_query("SELECT * FROM events WHERE eid = " & request("eid"))
			title		= 	Events("title")
			des			= 	Events("desc")
			sDate		= 	Events("sDate")
			eDate		=	Events("eDate")
			sTime		=	Events("sTime")
			eTime		= 	Events("eTime")
			loc			=	Events("location")
			db_close
		end if
		if isNew then buttonvalue = "Add" else buttonvalue = "Edit"
		.setVariable 	"BUTTONVALUE"	,	buttonvalue
		.setVariable 	"TITLE"			,	title
		.setVariable 	"DES"			,	des
		.setVariable	"SDATE"			,	sDate
		.setVariable 	"EDATE"			,	eDate
		.setVariable	"STIME"			,	sTime
		.setVariable	"ETIME"			,	eTime
		.setVariable	"LOC"			,	loc
		.setVariable	"EID"			,	request("eid")
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock     "warning"
		
		if e.hasErrors then
			.setVariable "WARNING", e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable 	"FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable	"COL1"			,	"Event Listings"
		.setVariable	"BODY1"			,	tIn.getOutput
		
		if isUpdate = true then
			.setVariable "HEADING"	,	"Editing Event Listing"
		else
			.setVariable "HEADING"	,	"Adding Event Listings"
		end if
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION", "navigation.html"
		
		.parse
	end with
end sub

sub parse_addEditEvents()
	if(request("eid") <> "") then
		isUpdate = true
	else 
		isNew = true
	end if
	
	'Make sure all required fields are not empty
	if(request("title")		= "") then e.raiseError("Title is a required field.")
	if(request("des")		= "") then e.raiseError("Description is a required field.")
	if(request("sDate")		= "") then e.raiseError("Start date is a required field.")
	if(request("sTime")		= "") then e.raiseError("Start time is a required field.")
	if(request("location")	= "") then e.raiseError("Location is a required field.")
	
	'Make sure all fields are within the valid database length
	if(Len(request("title")) > 50) then e.raiseError("Title is limited to 50 characters.")
	if(Len(request("des")) >  255) then e.raiseError("Description is limited to 255 characters.")
	if(Len(request("sDate")) > 10) then e.raiseError("Start date is limited to 10 characters. dd/mm/yyyy")
	if(Len(request("eDate")) > 10) then e.raiseError("End date is limited to 10 characters. dd/mm/yyyy")
	if(Len(request("sTime")) >  8) then e.raiseError("Start time is limited to 8 characters. HH:MM PM/AM")
	if(Len(request("eTime")) >  8) then e.raiseError("End time is limited to 8 characters. HH:MM PM/AM")
	if(Len(request("location")) > 50) then e.raiseError("Location is limited to 50 characters.")
	
	if e.hasErrors then
		call content_addEditEvents()
	else
		db_open(Application("ConnectionString"))
		set get_nextid = db_query("SELECT MAX(eid) AS [max] FROM events")
		if(isNull(get_nextid("max"))) then
			nextid = 1
		else
			nextid = get_nextid("max") + 1
		end if
		
		if isNew = true then
			db_execute("INSERT INTO events VALUES('" & nextid & "', '" & db_encode(request("title")) & "', '" & db_encode(request("des")) & "', '" & db_encode(request("sDate")) & "', '" & db_encode(request("eDate")) & "', '" & db_encode(request("sTime")) & "', '" & db_encode(request("eTime")) & "', '" & db_encode(request("location")) & "')")
		else
			db_execute("UPDATE events SET title = '" & db_encode(request("title")) & "', [desc] = '" & db_encode(request("des")) & "', sDate = '" & db_encode(request("sDate")) & "', eDate = '" & db_encode(request("eDate")) & "', sTime = '" & db_encode(request("sTime")) & "', eTime = '" & db_encode(request("eTime")) & "', location = '" & db_encode(request("location")) & "' WHERE eid = " & request("eid"))
		end if
		
		db_close
	end if
	call content_listEvents()
end sub

sub content_deleteEvents()
	with tIn
		.setTemplateFile "deleteEvents.html"
		.setVariable     "EID",	request("eid")
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock	 "warning"
		
		if e.hasErrors then
			.setVariable "WARNING", e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable 	"FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable	"HEADING"		,	"Deleting Events"
		.setVariable	"COL1"			,	"Events"
		.setVariable	"BODY1"			,	tIn.getOutput
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION", "navigation.html"
		.parse
	end with
end sub

sub parse_deleteEvents()
	if request("yes") <> "" then
		db_open(Application("ConnectionString"))
		db_execute("DELETE FROM events WHERE eid = " & request("eid"))
		db_close
	end if
	
	call content_listEvents()
end sub
%>
