<%
'--------------------------------------------------------------------------------
' DECLARATIONS
'--------------------------------------------------------------------------------

dim objConn


'--------------------------------------------------------------------------------
' HTML/URL ENOCODING AND CRLF TO <BR> FUNCTIONS
'--------------------------------------------------------------------------------

function url(stringPayload)
	url = server.urlEncode(stringPayload)
end function


function html(stringPayload)
	html = server.htmlEncode(stringPayload)
end function


function crlfToBr(stringPayload)
	crlfToBr = replace(stringPayload, chr(13) & chr(10), "<br>")
end function


'--------------------------------------------------------------------------------
' DATABASE OPERATIONS
'--------------------------------------------------------------------------------

function db_type()
	if db_state() then
		c = lCase(objConn.connectionString)
		
		if inStr(c, "pgsql") > 0 then
			t = "PGSQL"
		else
			t = "MSSQL"
		end if
	else
		t = null
	end if

	db_type = t
end function


function db_state()
	if isObject(objConn) then
		if objConn.state = 0 then db_state = false else db_state = true
	else
		db_state = false
	end if
end function


function db_open(connectionStringPayload)
	if isObject(objConn) then
		if objConn.state = 0 then
			set objConn = server.createObject("ADODB.CONNECTION")
			objConn.open connectionStringPayload
			db_open = true
		else
			db_open = false
		end if
	else
		set objConn = server.createObject("ADODB.CONNECTION")
		objConn.open connectionStringPayload
		db_open = true
	end if
end function


function db_close()
	if isObject(objConn) then
		if objConn.state = 1 then
			objConn.close
			set objConn = nothing
			objConn     = null
			db_close = true
		else
			db_close = false
		end if
	else
		db_close = false
	end if
end function


function db_query(sqlPayload)
	if db_state = true then set db_query = objConn.execute(sqlPayload)
end function


sub db_execute(sqlPayload)
	if db_state = true then objConn.execute(sqlPayload)
end sub


function db_encode(value)

	if db_type() = "PGSQL" then
		value = replace(value, "\", "\\")
		value = replace(value, "'", "\'")
		value = replace(value, chr(13), "\r")
		value = replace(value, chr(12), "\f")
		value = replace(value, chr(10), "\n")
		value = replace(value, chr(9),  "\t")
		value = replace(value, chr(8),  "\b")

	else
		value = replace(value, "'", "''")
	end if

	db_encode = value
end function


function getId(id, tableName)
	if db_state = true then
		if db_type() = "PGSQL" then
			set nextId = db_query("SELECT MAX(" & id & ") + 1 AS ""id"" FROM """ & tableName & """")
		else
			set nextId = db_query("SELECT MAX(" & id & ") + 1 AS id FROM " & tableName)
		end if

		if varType(nextId("id")) <> 1 then getId = nextId("id") else getId = 1
	end if
end function	


'--------------------------------------------------------------------------------
' ADMINISTRATION
'--------------------------------------------------------------------------------

sub redirectIfNotAuthenticated()
	if isObject(session("loginInformation")) then
		if session("loginInformation")("loggedIn") = FALSE then response.redirect("/admin")
	else
		response.redirect("/admin")
	end if
end sub
%>