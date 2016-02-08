<%
'    ASP Template
'    Copyright (C) 2001  Valerio Santinelli
'
'    This library is free software; you can redistribute it and/or
'    modify it under the terms of the GNU Lesser General Public
'    License as published by the Free Software Foundation; either
'    version 2.1 of the License, or (at your option) any later version.
'
'    This library is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'    Lesser General Public License for more details.
'
'    You should have received a copy of the GNU Lesser General Public
'    License along with this library; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'   ---------------------------------------------------------------------------
'
' ASP Template main class file
'
' Author: Valerio Santinelli <tanis@mediacom.it>
' $Id: asptemplate.asp,v 1.8 2001/12/19 16:51:48 tanis Exp $
'

'===============================================================================
' Name: ASPTemplate Class
' Purpose: HTML separation class
' Functions:
'     <functions' list in alphabetical order>
' Properties:
'     <properties' list in alphabetical order>
' Methods:
'     <Methods' list in alphabetical order>
' Author: Valerio Santinelli <tanis@mediacom.it>
' Start: 2001/01/01
' Modified: 2001/12/19
'===============================================================================
class ASPTemplate

	' Contains the error objects
	private p_error
	
	' Print error messages?
	private p_print_errors
	
	' Opening delimiter (usually "{{")
	private p_var_tag_o
	
	' Closing delimiter (usually "}}")
	private p_var_tag_c

	'private p_start_block_delimiter_o
	'private p_start_block_delimiter_c
	'private p_end_block_delimiter_o
	'private p_end_block_delimiter_c
	
	'private p_int_block_delimiter
	
	private p_template
	private p_variables_list
	private p_blocks_list
	private p_blocks_name_list
	private	p_regexp
	private p_parsed_blocks_list
	
	' Directory containing HTML templates
	private p_templates_dir
	
	'===============================================================================
	' Name: class_Initialize
	' Purpose: Constructor
	' Remarks: None
	'===============================================================================
	private sub class_Initialize
		p_print_errors = FALSE
		' Remember that opening and closing tags are being used in regular expressions
		' and must be explicitly escaped
		p_var_tag_o = "\{\{"
		p_var_tag_c = "\}\}"
		' Block delimiters are actually disabled and no longer available. Maybe they'll be again
		' in the future.
		'p_start_block_delimiter_o = "<!-- BEGIN "
		'p_start_block_delimiter_c = " -->"
		'p_end_block_delimiter_o = "<!-- END "
		'p_end_block_delimiter_c = " -->"
		'p_int_block_delimiter = "__"
		p_templates_dir = "templates/"
		set p_variables_list = createobject("Scripting.Dictionary")
		set p_blocks_list = createobject("Scripting.Dictionary")
		set p_blocks_name_list = createobject("Scripting.Dictionary")
		set p_parsed_blocks_list = createobject("Scripting.Dictionary")
		p_template = ""
		Set p_regexp = New RegExp   
	end sub
	
	'===============================================================================
	' Name: SetTemplatesDir
	' Input:
	'    dir as Variant Directory
	' Output:
	' Purpose: Sets the directory containing html templates
	' Remarks: None
	'===============================================================================
	public sub SetTemplatesDir(dir)
		p_templates_dir = dir
	end sub

	'===============================================================================
	' Name: SetTemplate
	' Input:
	'    template as Variant String containing the template
	' Output:
	' Purpose: Sets a template passed through a string argument
	' Remarks: None
	'===============================================================================
	public sub SetTemplate(template)
		p_template = template
	end sub
	

	'===============================================================================
	' Name: SetTemplateFile
	' Input:
	'    inFileName as Variant Name of the file to read the template from
	' Output:
	' Purpose: Sets a template given the filename to load the template from
	' Remarks: None
	'===============================================================================
	public sub SetTemplateFile(inFileName)
		if len(inFileName) > 0 then
			dim FSO, oFile
			set FSO = createobject("Scripting.FileSystemObject")
			if FSO.FileExists(server.mappath(p_templates_dir & inFileName)) then
				set oFile = FSO.OpenTextFile(server.mappath(p_templates_dir & inFileName))
				p_template = oFile.ReadAll
				oFile.Close
				set oFile = Nothing
			else
				response.write "<b>ASPTemplate Error: File [" & inFileName & "] does not exists!</b><br>"
			end if
			set FSO = nothing
		else
			response.write "<b>ASPTemplate Error: SetTemplateFile missing filename.</b><br>"
		end if
		
	end sub
	

	'===============================================================================
	' Name: SetVariable
	' Input:
	'    s as Variant - Variable name
	'    v as Variant - Value
	' Output:
	' Purpose: Sets a variable given it's name and value
	' Remarks: None
	'===============================================================================
	public sub SetVariable(s, v)
		if p_variables_list.Exists(s) then
			p_variables_list.Remove s
			p_variables_list.Add s, CStr(v)
		else
			p_variables_list.Add s, CStr(v)
		end if
	end sub


	'===============================================================================
	' Name: Append
	' Input:
	'    s as Variant - Variable name
	'    v as Variant - Value
	' Output:
	' Purpose: Sets a variable appending the new value to the existing one
	' Remarks: None
	'===============================================================================
	public sub Append(s, v)
		if p_variables_list.Exists(s) then
			tmp = p_variables_list.Item(s) & v
			p_variables_list.Remove s
			p_variables_list.Add s, tmp
		else
			p_variables_list.Add s, v
		end if
	end sub
	
	
	'===============================================================================
	' Name: SetVariableFile
	' Input:
	'    s as Variant Variable name
	'    inFileName as Variant Name of the file to read the value from
	' Output:
	' Purpose: Load a file into a variable's value
	' Remarks: None
	'===============================================================================
	public sub SetVariableFile(s, inFileName)
		if len(inFileName) > 0 then
			dim FSO, oFile
			set FSO = createobject("Scripting.FileSystemObject")
			if FSO.FileExists(server.mappath(p_templates_dir & inFileName)) then
				set oFile = FSO.OpenTextFile(server.mappath(p_templates_dir & inFileName))
				ReplaceBlock s, oFile.ReadAll
				oFile.Close
				set oFile = Nothing
			else
				response.write "<b>ASPTemplate Error: File [" & inFileName & "] does not exists!</b><br>"
			end if
			set FSO = nothing
		else
			'Filename was never passed!
		end if
	end sub


	'===============================================================================
	' Name: ReplaceBlock
	' Input:
	'    s as Variant Variable name
	'    inFile as Variant Content of the file to place in the template
	' Output:
	' Purpose: Function used by SetVariableFile to load a file and replace it
	'          into the template in place of a variable
	' Remarks: None
	'===============================================================================
	public sub ReplaceBlock(s, inFile)
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = p_var_tag_o & s & p_var_tag_c
		p_template = p_regexp.Replace(p_template, inFile)   
	end sub

	public property get GetOutput
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		Set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			if p_variables_list.Exists(match.SubMatches(1)) then
				p_regexp.Pattern = match.Value
				p_template = p_regexp.Replace(p_template, p_variables_list.Item(match.SubMatches(1)))
			end if
			'response.write match.Value & "<br>"
		next

		p_regexp.Pattern = "__[_a-z0-9]*__"
		Set Matches = p_regexp.Execute(p_template)   
		for each match in Matches
			'response.write "[[" & match.Value & "]]<br>"
			p_regexp.Pattern = match.Value
			p_template = p_regexp.Replace(p_template, "")
		next

		GetOutput = p_template
	end property

	public sub Parse
		parsed = GetOutput
		response.write parsed
	end sub
	
		
	' TODO: if the block foud contains other blocks, it should recursively update all of them without the needing
	' of doing this by hand.
	public sub UpdateBlock(inBlockName)
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "<!--\s+BEGIN\s+(" & inBlockName & ")\s+-->([\s\S.]*)<!--\s+END\s+\1\s+-->"
		Set Matches = p_regexp.Execute(p_template)
		Set match = Matches
		for each match in Matches
			p_blocks_list.Add inBlockName, match.SubMatches(1)
			p_blocks_name_list.Add inBlockName, inBlockName
			p_template = p_regexp.Replace(p_template, "__" & inBlockName & "__")
			'response.write "[[" & match.SubMatches(1) & "]]<br>"
		next		
	end sub

	public sub ParseBlock(inBlockName)
		w = GetBlock(inBlockName)
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "(__)([_a-z0-9]+)__"
		Set Matches = p_regexp.Execute(w)
		Set match = Matches
		for each match in Matches
			'response.write inBlockName & " - " & match.Value & "<br>"
			'response.write "[[" & match.SubMatches(1) & "]]<br>"
			if p_parsed_blocks_list.Exists(match.SubMatches(1)) then
				w = p_regexp.Replace(w, p_parsed_blocks_list.Item(match.SubMatches(1)) & "__" & match.SubMatches(1) &"__")
				p_parsed_blocks_list.Remove(match.SubMatches(1))
			end if 		
		next
		
		if p_parsed_blocks_list.Exists(inBlockName) then
			tmp = p_parsed_blocks_list.Item(inBlockName) & w
			p_parsed_blocks_list.Remove inBlockName
			p_parsed_blocks_list.Add inBlockName, tmp
		else
			p_parsed_blocks_list.Add inBlockName, w
		end if

		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "__" & inBlockName & "__"
		Set Matches = p_regexp.Execute(p_template)
		Set match = Matches
		for each match in Matches
			w = GetParsedBlock(inBlockName)
			'response.write "w:" & w
			p_regexp.Pattern = "__" & inBlockName & "__"
			p_template = p_regexp.Replace(p_template, w & "__" & inBlockName & "__")
			'response.write "[[" & match.Value & "]]<br>"
			'response.write "[[" & p_regexp.Pattern & "]]<br>"
		next		


	end sub

	private property get GetBlock(inToken)
		'This routine checks the Dictionary for the text passed to it.
		'If it finds a key in the Dictionary it Display the value to the user.
		'If not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_blocks_list.Item(inToken)
			s = ParseBlockVars(tmp)
			GetBlock = s
			'response.write "s: " & s
		else
			GetBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	private property get GetParsedBlock(inToken)
		'This routine checks the Dictionary for the text passed to it.
		'If it finds a key in the Dictionary it Display the value to the user.
		'If not, by default it will display the full Token in the HTML source so that you can debug your templates.
		if p_blocks_list.Exists(inToken) then
			tmp = p_parsed_blocks_list.Item(inToken)
			s = ParseBlockVars(tmp)
			GetParsedBlock = s
			'response.write "s: " & s
			p_parsed_blocks_list.Remove(inToken)
		else
			GetParsedBlock = "<!--__" & inToken & "__-->" & VbCrLf
		end if
	end property


	public property get ParseBlockVars(inText)
		p_regexp.IgnoreCase = True
		p_regexp.Global = True

		p_regexp.Pattern = "(" & p_var_tag_o & ")([^}]+)" & p_var_tag_c
		Set Matches = p_regexp.Execute(inText)   
		for each match in Matches
			if p_variables_list.Exists(match.SubMatches(1)) then
				p_regexp.Pattern = match.Value
				inText = p_regexp.Replace(inText, p_variables_list.Item(match.SubMatches(1)))
			end if
			'response.write match.Value & "<br>"
			'response.write inText & "<br>"
		next
		ParseBlockVars = inText
	end property

end class
%>