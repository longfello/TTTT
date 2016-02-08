<!--#include virtual="/core/classes/aspTemplate.asp" -->
<!--#include virtual="/core/modules/common.asp" -->
<!--#include virtual="/core/classes/error.asp" -->
<%
'The above lines include 3 files located in the core directory
'The ASP Template Class
'The Common.asp file which handles all database activity
'The Error class
'----------------------------------------------------------------------
'INITIALIZATIONS
'    Initialize all variables here
'----------------------------------------------------------------------
'initialize your template variable
set t = new aspTemplate
'set the location of the templates to be used
t.SetTemplatesDir "/templates/"
'----------------------------------------------------------------------
'CONTROL STRUCTURE
'----------------------------------------------------------------------
select case request("a")
    case ""
        call content_viewWhatsnewPage()
end select
'----------------------------------------------------------------------
'GARBAGE
'----------------------------------------------------------------------
set t = nothing
'----------------------------------------------------------------------
'MODULES
'----------------------------------------------------------------------
sub content_viewWhatsnewPage()
    with t
        'set the template file to work with
        .setTemplateFile "whatsnew.html"
        'update all blocks in the file
        .updateBlock "no_data"
        .updateBlock "data"
        
        'Open the database connection (this is handled by the db_open
        'function in the common.asp file, located in the core folder
        db_open(Application("ConnectionString"))
        'set the recordset to be returned. (This is handled by the db_query
        'function in the common.asp file, located in the core folder
        'it takes a sqlstatement as a parameter
        set listings = db_query("SELECT * FROM whatsnew ORDER BY dateAdded")
        
        'if no records are returned, then show the no_data block
        'inside the template
        if(listings.eof) then
            .parseBlock "no_data"
        end if
        
        'If there are records 
        do until listings.eof
            'checks to determine the recordset data returned
            if(listings("email") = "") then email = "" else email = "<a class='txtSmallBoldMaroon' href='mailto:" & listings("email") & "'>" & listings("email") & "</a>"
            if(listings("link") = "") then 
                link = "" 
            else
                if(Left(listings("link"), 7) <> "http://") then
                    link = "<a class='txtSmallBoldMaroon' href='http://" & listings("link") & "' target='_blank'>" & listings("link") & "</a>"
                else
                    link = "<a class='txtSmallBoldMaroon' href=" & listings("link") & " target='_blank'>" & listings("link") & "</a>"
                end if
            end if
            
            'set the variables on the template = to the data returned in the recordset
            .setVariable "TITLE"    ,   listings("title")
            .setVariable "DES"      ,   listings("desc")
            .setVariable "EMAIL"    ,   email
            .setVariable "LINK"     ,   link
            
            'parse the block on the template for this record
            .parseBlock "data"
            'move to the next record and repeat until the recordset is empty
            listings.movenext
        loop
        
        'close the database connection (this is handle by the db_close function
        'in the common.asp file, located in the core folder
        db_close
        'parse the page (this is required to display the page to the browser)
        .parse
    end with
end sub
%>