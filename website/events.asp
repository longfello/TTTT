<!--#include virtual="/core/classes/aspTemplate.asp" -->
<!--#include virtual="/core/modules/common.asp" -->
<!--#include virtual="/core/classes/error.asp" -->
<%
'Details of this file are located in the index.asp file
'----------------------------------------------------------------------
'INITIALIZATIONS
'    Initialize all variables here
'----------------------------------------------------------------------
set t = new aspTemplate
t.SetTemplatesDir "/templates/"
'----------------------------------------------------------------------
'CONTROL STRUCTURE
'----------------------------------------------------------------------
select case request("a")
	case ""
		call content_viewEventsPage()
end select
'----------------------------------------------------------------------
'GARBAGE
'----------------------------------------------------------------------
set t = nothing
'----------------------------------------------------------------------
'MODULES
'----------------------------------------------------------------------
sub content_viewEventsPage()
	with t
		.setTemplateFile "events.html"
		.updateBlock "no_data"
		.updateBlock "data"
		
		db_open(Application("ConnectionString"))
		set listings = db_query("SELECT * FROM events ORDER BY sDate")
		
		if(listings.eof) then
			.parseBlock "no_data"
		end if
		
		do until listings.eof
			.setVariable "TITLE"	,	listings("title")
			.setVariable "DES"		,	listings("desc")
			.setVariable "LOC"		,	listings("location")
			
			if(listings("eDate") = "") then
				.setVariable "DATE"	,	listings("sDate")
			else
				.setVariable "DATE",	listings("sDate") & " To " & listings("eDate")
			end if
			
			if(listings("eTime") = "") then
				.setVariable "TIME"	,	listings("sTime")
			else
				.setVariable "TIME"	,	listings("sTime") & " To " & listings("eTime")
			end if
			
			.parseBlock "data"
			listings.movenext
		loop
		
		db_close
		.parse
	end with
end sub
%>
