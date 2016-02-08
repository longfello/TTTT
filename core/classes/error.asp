<%
class error
	private rx
	private debug
	private errors
	private errNumber

	private sub class_initialize
		debug      = false
		set rx     = new RegExp
		set errors = server.createObject("SCRIPTING.DICTIONARY")
		errNumber  = 999
	end sub

	
	private sub class_terminate
		set rx     = nothing
		set errors = nothing
	end sub


	public function rxDataCheck(pattern, subject)
		rx.pattern = pattern
		if rx.test(subject) then rxDataCheck = true else rxDataCheck = false
	end function


	public sub raiseError(errorPayload)
		if debug then err.raise errNumber, "Class Error", errorPayload
		if NOT errors.exists(errors.count) then errors.add errors.count, errorPayload
	end sub


	public property get getErrors()
		set getErrors = errors
	end property


	public property get hasErrors()
		if errors.count > 0 then hasErrors = true else hasErrors = false
	end property


	public function getErrorsFormatted(templateFilePayload)
		set tErrors = new aspTemplate

		tErrors.setTemplatesDir ""
		tErrors.setTemplateFile templateFilePayload
		tErrors.updateBlock     "row"	

		for a = 0 to errors.count
			if errors.exists(a) then
				tErrors.setVariable "ERROR", html(errors(a))

				tErrors.parseBlock  "row"
			end if
		next
	
		getErrorsFormatted = tErrors.getOutput

		set tErrors = nothing
	end function

end class
%>