<!--#include virtual="/core/classes/aspTemplate.asp" -->
<!--#include virtual="/core/classes/error.asp" -->
<!--#include virtual="/core/modules/common.asp" -->
<%
'---------------------------------------------------------------------
'INITIALIZATIONS
'---------------------------------------------------------------------
set tIn  = new aspTemplate
set tOut = new aspTemplate
set e = new error
tIn.setTemplatesDir "/admin/templates/whatsnew/"
tOut.setTemplatesDir "/core/templates/admin/"
'---------------------------------------------------------------------
'CONTROL STRUCTURE
'---------------------------------------------------------------------
select case request("a")
	case ""
		call content_listWhatsnew()
	case "addEditWhatsnew"
		call content_addEditWhatsnew()
	case "addEditWhatsnew_finish"
		call parse_addEditWhatsnew()
	case "deleteWhatsnew"
		call content_deleteWhatsnew()
	case "deleteWhatsnew_finish"
		call parse_deleteWhatsnew()
end select
'---------------------------------------------------------------------
'GARBAGE
'---------------------------------------------------------------------
set tIn  = nothing
set tOUt = nothing
'---------------------------------------------------------------------
'MODULES
'---------------------------------------------------------------------
sub content_listWhatsnew()
	with tIn
		.setTemplateFile "listWhatsnew.html"
		.updateBlock     "no_data"
		.updateBlock     "data"
		
		db_open(Application("ConnectionString"))
		set newListings = db_query("SELECT * FROM whatsnew ORDER BY dateAdded")
	
		if(newListings.eof) then
			.parseBlock "no_data"
		end if
		
		do until newListings.eof
			if(newListings("email") = "") then email = "" else email = "<a class='txtSmall' href=mailto:" & newListings("email") & ">" & newListings("email") & "</a>"
			if(newListings("link") = "") then 
				link = "" 
			else
				if(Left(newListings("link"), 7) <> "http://") then
					link = "<a href='http://" & newListings("link") & "'>" & newListings("link") & "</a>"
				else
					link = "<a href=" & newListings("link") & ">" & newListings("link") & "</a>"
				end if
			end if
			
			.setVariable "TITLE"	,	newListings("title")
			.setVariable "DES"		,	newListings("desc")
			.setVariable "EMAIL"	,	email
			.setVariable "LINK"		,	link
			.setVariable "EDIT_LINK",	request("SCRIPT_NAME") & "?a=addEditWhatsnew&nid=" & newListings("nid")
			.setVariable "DELETE_LINK",	request("SCRIPT_NAME") & "?a=deleteWhatsnew&nid=" & newListings("nid")
			
			.parseBlock "data"
			newListings.movenext 
		loop
		
		db_close	
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock     "warning"
		
		if e.hasErrors = true then
			.setVariable "WARNING", e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable "FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable "HEADING"		,	"What's New Listings"
		.setVariable "COL1"			,	"What's New Listings"
		.setVariable "BODY1"		,	tIn.getOutput
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION",	"navigation.html"
		
		.parse
	end with
end sub

sub content_addEditWhatsnew()
	with tIn
		.setTemplateFile "addEditWhatsnew.html"
		if request("nid") <> "" then
			isUpdate = true
		else 
			isNew = true
		end if
		
		if isNew then
			title  = request("title")
			des    = request("des")
			email  = request("email")
			link   = request("link")
		else
			db_open(Application("ConnectionString"))
			set newListings = db_query("SELECT * FROM Whatsnew WHERE nid = " & request("nid"))
			title  = newListings("title")
			des    = newListings("desc")
			email  = newListings("email")
			link   = newListings("link")
			
			db_close
		end if
		if isNew then buttonvalue = "Add" else buttonvalue = "Edit"
		.setVariable	"BUTTONVALUE"	,	buttonvalue
		.setVariable	"TITLE"			,	title
		.setVariable	"DES"			,	des
		.setVariable	"EMAIL"			,	email
		.setVariable 	"LINK"			,	link
		.setVariable	"NID"			,	request("nid")
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock     "warning"
		
		if e.hasErrors then
			.setVariable "WARNING", e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable "FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable "COL1"			,	"What's New Listings"
		.setVariable "BODY1"		,	tIn.getOutput
		
		if isNew = true then
			.setVariable "HEADING"	,	"Adding What's New Listing"
		else
			.setVariable "HEADING"	,	"Editing What's New Listing"
		end if
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION", "navigation.html"
		
		.parse
	end with
end sub

sub parse_addEditWhatsnew()
	if request("nid") <> "" then
		isUpdate = true
	else
		isNew = true
	end if
	
	if request("title") = "" then e.raiseError("Title is a required field.")
	if request("des")  = "" then e.raiseError("Description is a required field.")
	
	if e.hasErrors then
		call content_addEditWhatsnew()
	else
		db_open(Application("ConnectionString"))
		set next_nid = db_query("SELECT MAX(nid) AS [max] FROM whatsnew")
		if(isNull(next_nid("max"))) then
			nextid = 1
		else
			nextid = next_nid("max") + 1
		end if
		
		if isNew = true then
			db_execute("INSERT INTO whatsnew VALUES('" & nextid & "', '" & db_encode(request("title")) & "', '" & db_encode(request("des")) & "', '" & db_encode(request("email")) & "', '" & db_encode(request("link")) & "', '" & Now() & "')")
		else
			db_execute("UPDATE whatsnew SET title = '" & db_encode(request("title")) & "', [desc] = '" & db_encode(request("des")) & "', email = '" & db_encode(request("email")) & "', link = '" & db_encode(request("link")) & "', dateAdded = '" & Now() & "' WHERE nid = " & request("nid"))
		end if
		db_close
	end if
	call content_listWhatsnew()
end sub

sub content_deleteWhatsnew()
	with tIn
		.setTemplateFile "deleteWhatsnew.html"
		.setVariable     "NID",	request("nid")
	end with
	
	with tOut
		.setTemplateFile "interface_100.html"
		.updateBlock     "warning"
		
		if e.hasErrors then
			.setVariable "WARNING", e.getErrorsFormatted("/core/templates/admin/adminWarning.html")
			.parseBlock  "warning"
		end if
		
		.setVariable "FORMACTION"	,	request("SCRIPT_NAME")
		.setVariable "HEADING"		,	"Deleting What's New Listing"
		.setVariable "COL1"			,	"What's New Listings"
		.setVariable "BODY1"		,	tIn.getOutput
		
		.setTemplatesDir "/admin/templates/"
		.setVariableFile "NAVIGATION", "navigation.html"
		
		.parse
	end with
end sub

sub parse_deleteWhatsnew()
	if request("yes") <> "" then
		db_open(Application("ConnectionString"))
		db_execute("DELETE FROM Whatsnew WHERE nid = " & request("nid"))
		db_close
	end if
	
	call content_listWhatsnew()
end sub
%>
