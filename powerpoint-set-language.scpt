tell application "Microsoft PowerPoint"
	activate
	
	set theView to view of document window 1
	
	repeat with slideNumber from 1 to count of slides of active presentation
		
		go to slide theView number slideNumber
		
		tell slide slideNumber of active presentation
			activate
			tell shapes
				select
			end tell
		end tell
		tell application "System Events"
			tell process "Microsoft PowerPoint"
				tell menu bar 1
					tell menu bar item "Tools"
						tell menu "Tools"
							click menu item 5
						end tell
					end tell
				end tell
			end tell
			tell window 1 of process "Microsoft PowerPoint"
				tell table 1 of scroll area 1
					select (first row whose value of static text 1 of UI element 1 is "English (United Kingdom)")
				end tell
				click button "OK"
			end tell
		end tell
		
	end repeat
end tell
