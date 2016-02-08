<!--#include virtual="/core/classes/aspTemplate.asp" -->
<!--#include virtual="/core/classes/error.asp" -->
<!--#include virtual="/core/modules/common.asp" -->
<%
'---------------------------------------------------------------------
'INITIALIZATIONS
'---------------------------------------------------------------------
set t = new aspTemplate
t.setTemplatesDir "/admin/templates/"
'---------------------------------------------------------------------
'CONTROL STRUCTURE
'---------------------------------------------------------------------
select case request("a")
	case ""
		call content_view()
end select
'---------------------------------------------------------------------
'GARBAGE
'---------------------------------------------------------------------
set t = nothing
'---------------------------------------------------------------------
'MODULES
'---------------------------------------------------------------------
sub content_view()
	with t
		.setTemplateFile "interface.htm"
		.parse
	end with
end sub
%>
