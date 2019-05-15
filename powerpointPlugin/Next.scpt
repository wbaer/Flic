tell application "System Events"
	if frontmost of process "Microsoft PowerPoint" then
		tell application "Microsoft PowerPoint"
			if (slide state of slide show view of slide show window of active presentation is slide show state running) then
				go to next slide (slide show view of slide show window of active presentation)
			else
				set slide state of slide show view of slide show window of active presentation to slide show state running
			end if
		end tell
	end if
end tell