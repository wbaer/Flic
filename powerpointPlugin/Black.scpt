tell application "System Events"
	if frontmost of process "Microsoft PowerPoint" then
		tell application "Microsoft PowerPoint"
			if (slide state of slide show view of slide show window of active presentation is slide show state running) then
				set slide state of slide show view of slide show window of active presentation to slide show state black screen
			end if
		end tell
	end if
end tell	