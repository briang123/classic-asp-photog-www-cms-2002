<%

function IsNaN(byval n)
	dim d

	on error resume next
	if not isnumeric(n) then
		IsNan = true
		Exit Function
	end if
	d = cdbl(n)
	if err.number <> 0 then isNan = true else isNan = false
	On Error GoTo 0
end function

Response.write("An error occurred while trying to access the web page. If you believe this to be a mistake, please contact the site administrator.")
%>
