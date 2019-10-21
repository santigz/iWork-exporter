#!/usr/bin/osascript


on run argv
	set inputFile to cleanInput(item 1 of argv)
	if length of (inputfile as string) is 0
		log "ERROR reading file: " & inputFile
		return
	end if

	ensureUTF8Encoding()
	processFile(inputFile)

end run

on cleanInput(inputFile)
	-- Get full path
	if not inputFile starts with "/"
		set inputFile to (POSIX path of (POSIX file (do shell script "pwd") as alias)) & inputFile
	end if
	tell application "System Events"
		if not exists file inputFile then
			return "ERROR reading file: " & inputFile
		end if
	end tell
	return POSIX file inputFile as alias
end cleanInput

on processFile(theFile)
	tell application "Finder"
		log POSIX path of theFile

		set fileInfo to (info for (theFile))
		set fileName to name of (fileInfo)
		set fileExtension to name extension of (fileInfo)

		set extensionLength to ((length of fileExtension) + 2)
		set fileName to fileName's text 1 thru (-1 * extensionLength)
		
		-- Create export directory if not exists
		set exportDirName to fileName & "-export"
		set exportDir to (container of theFile as text) & exportDirName
		if not exists exportDir then
			make new folder at (container of theFile) with properties {name:exportDirName}
		end if

		set the CSVpath to exportDir & ":" & fileName & ".csv"
		set the XLSpath to exportDir & ":" & fileName & ".xls"
		set the PDFpath to exportDir & ":" & fileName & ".pdf"
	end tell
	
	if fileExtension is "numbers"
		exportNumbers(theFile, CSVpath, XLSpath, PDFpath)
	end if

	if fileExtension is "pages"
		exportNumbers(theFile, CSVpath, XLSpath, PDFpath)
	end if
	
end processFile

on exportNumbers(theFile, CSVpath, XLSpath, PDFpath)
	tell application "Numbers"
		activate
		set docRef to open theFile
		
		if length of CSVpath is not 0
			log "    " & POSIX path of CSVpath
			export docRef to file CSVpath as CSV
		end if
		if length of XLSpath is not 0
			log "    " & POSIX path of XLSpath
			export docRef to file XLSpath as CSV
		end if
		if length of PDFpath is not 0
			log "    " & POSIX path of PDFpath
			export docRef to file PDFpath as CSV
		end if
		
		close docRef without saving
	end tell
end exportNumbers

(*
 Here are the codes that apply: 4=UTF8, 12=windows latin, 30=MacRoman
 Taken from: https://gist.github.com/idStar/61994506d69595da3d30
 More in : https://discussions.apple.com/thread/4018778?tstart=0 
*)
on ensureUTF8Encoding()
	do shell script "/usr/bin/defaults write com.apple.iWork.Numbers CSVExportEncoding -int 4"
end ensureUTF8Encoding

