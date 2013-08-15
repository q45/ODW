--
--  PSAppDelegate.applescript
--  ODW
--
--  Created by Quintin Smith on 8/8/13.
--  Copyright (c) 2013 Quintin Smith. All rights reserved.
--

script PSAppDelegate
	property parent : class "NSObject"
	property status_bar : missing value
	property input_text : missing value
	property move_or_copy : missing value
	property output_text : missing value
	property main_window : missing value
	property obs_path : "fas1:Photo_Studio:Photo Studio Library:Award Images"
	property desk_to_path : ""
	property current_move_or_copy : 0
	property ilp_folder_list : null
	property textField : missing value
	property statusTextField : missing value
	
	
	
	on cancel_(sender)
		
		quit
	end cancel_
	
	on start_(sender)
		tell application "Microsoft Excel"
			activate
			set excelFile to "Macintosh HD:Users:quintin.smith:Desktop:studiotest" as string
			log excelFile
			open workbook workbook file name (excelFile) as string
			
			tell active sheet of active workbook
				set cellIndex to 1
				
				repeat
				set oldname to (string value of cell cellIndex of column 1) as text
				set newname to (string value of cell cellIndex of column 2) as text
				set threename to (string value of cell cellIndex of column 3) as text
				
				if oldname is equal to "" then
					tell active sheet of active workbook to quit
					
				else
					set inputFolderPath to folderFinder(oldname as integer) of me
					
				end if
				end repeat
			end tell
		end tell
	end start_
	
	
	on input_(sender)
		status("Choosing input excel file")
		tell application "Finder"
			set input_text to ((choose folder default location (path to desktop))as string)
		end tell
		
		#set currentTextFieldText to input_text's' stringValue() as text
		
	end input_
	
	on output_(sender)
		status("Choosing output folder")
		tell application "Finder"
			if current_move_or_copy = 0 then
				tell application "Finder"
					#set output_text to (choose folder default location(file obs_path as alias)) as string)
					tell myTextField to setStringValue_(output_text)
					#output_text setStringValue_((choose folder default location (file obs_path as alias))as string)
				end tell
				else
				#output_text setStringValue_((choose folder default location(file desk_to_path as alias)) as string)
				tell myTextField to setStringValue_(desk_to_path)
			end if
		end tell
		status("")
	end output_
	
	on folderFinder(mmNumber)
		set return_folder to ""
		tell application "Finder"
			repeat with ailpFolder in ilp_folder_list
				try
					set starting_number to ((first word of((name of ailpFolder)as string))as integer)
					if(starting_number ≤ mmNumber and (return_folder = "" or starting_number > ((first word of ((name of return_folder)as string))as integer))) then
						set return_folder to ailpFolder
					end if
				end try
			end repeat
		end tell
	end folderFinder
		
		on status(text)
			statusTextField's setStringValue_(text) as text
		end status
 	
		
	on applicationWillFinishLaunching_(aNotification)
		-- Insert code here to initialize your application before any files are opened

		textField's setStringValue_("quintin.smith:Desktop:studiotest.xlsx") as text
		#textField's setStringValue_("Macintosh HD:Users:rick.hayward:Desktop:Image Pull.xlsx") as text
		output_text's setStringValue_("Macintosh HD:Users:rick.hayward:Desktop:RENAME") as text
		
		log output_text's stringValue()
		log textField's stringValue()

		log obs_path
		tell application "Finder"
		if not (exists folder "Photo_Studio:") then
			mount volume "smb://fas1:Photo_Studio:PhotoStudio_Libr:Image Library_Print"
		
			else
			set ilp_folder_list to (every file of folder "Photo_Studio:PhotoStudio_Libr:Image Library_Print:")
			
			end if
		end tell
		
		#set desk_to_path to((path to desktop) as string)
		
		
		
		
				
	end applicationWillFinishLaunching_
	
	on applicationShouldTerminate_(sender)
		-- Insert code here to do any housekeeping before your application quits 
		return current application's NSTerminateNow
	end applicationShouldTerminate_
	
end script