#!/usr/bin/osascript
(*
  File: export.applescript

  Export an iWork file (Pages, Numbers, Keynote) to other formats.
  Doesn't do much error handling.

  Example use: ./export.applescript document.pages

  Leaves exported files in document-exported folder.
*)

on run argv
	if (count of argv) < 2
		log "Usage: " & name of (info for (path to me)) & " DOCUMENT DIR"
		log "    DOCUMENT:  Pages, Numbers or Keynote document path"
		log "    DIR:       Directory for exported files"
		return
	end if

	set inputFile to cleanInputFile(item 1 of argv)
	set outputDir to cleanOutputDir(item 2 of argv)

	ensureUTF8Encoding()
	processFile(inputFile, outputDir)
end run

on cleanInputFile(inputFile)
	set inputFile to absolutePath(inputFile)

	-- Test file exists
	try
		set cleanPath to POSIX file inputFile as alias
	on error msg
		error "ERROR: " & msg
	end try

	return POSIX file inputFile as alias
end cleanInputFile

on cleanOutputDir(outputDir)
	set outputDir to cleanInputFile(outputDir)
	-- Check path is a folder
	if kind of (info for outputDir) is not "Folder" then
		error "Not folder: " & outputDir
	end if
	return outputDir
end cleanOutputDir

on absolutePath(thePath)
	if thePath starts with "/"
		return thePath
	else
		return (POSIX path of (POSIX file (do shell script "pwd") as alias)) & thePath
	end if
end absolutePath

on processFile(theFile, outputDir)
	log POSIX path of theFile

	-- File extension
	set fileInfo to (info for (theFile))
	set fileExtension to name extension of (fileInfo)

	-- File name without extension
	set extensionLength to ((length of fileExtension) + 2)
	set fileName to name of (fileInfo)
	set fileName to fileName's text 1 thru (-1 * extensionLength)

	-- Exported file paths
	set outputDir to outputDir as text
	set the CSVpath to outputDir & fileName & ".csv"
	set the XLSpath to outputDir & fileName & ".xlsx"
	set the PDFpath to outputDir & fileName & ".pdf"
	set the TXTpath to outputDir & fileName & ".txt"
	set the DOCpath to outputDir & fileName & ".docx"
	set the PPTpath to outputDir & fileName & ".pptx"
	set the HTMLpath to outputDir & fileName & ".html"
	
	if fileExtension is "numbers"
		exportNumbers(theFile, CSVpath, XLSpath, PDFpath)
	end if

	if fileExtension is "pages"
		exportPages(theFile, TXTpath, XLSpath, PDFpath)
	end if

	if fileExtension is "keynote"
		exportKeynote(theFile, HTMLpath, PPTpath, PDFpath)
	end if
	
end processFile

on exportPages(theFile, DOCpath, TXTpath, PDFpath)
	tell application "Pages"
		activate
		set docRef to open theFile
		
		log "    " & POSIX path of DOCpath
		export docRef to file DOCpath as Microsoft Word

		log "    " & POSIX path of TXTpath
		export docRef to file TXTpath as Unformatted Text

		log "    " & POSIX path of PDFpath
		export docRef to file PDFpath as PDF
		
		close docRef without saving
	end tell
end exportPages

on exportNumbers(theFile, CSVpath, XLSpath, PDFpath)
	tell application "Numbers"
		activate
		set docRef to open theFile
		
		log "    " & POSIX path of CSVpath
		export docRef to file CSVpath as CSV

		log "    " & POSIX path of XLSpath
		export docRef to file XLSpath as Microsoft Excel

		log "    " & POSIX path of PDFpath
		export docRef to file PDFpath as PDF
		
		close docRef without saving
	end tell
end exportNumbers

on exportKeynote(theFile, HTMLpath, PPTpath, PDFpath)
	tell application "Keynote"
		activate
		set docRef to open theFile
		
		log "    " & POSIX path of HTMLpath
		export docRef to file HTMLpath as HTML

		log "    " & POSIX path of PPTpath
		export docRef to file PPTpath as Microsoft PowerPoint

		log "    " & POSIX path of PDFpath
		export docRef to file PDFpath as PDF
		
		close docRef without saving
	end tell
end exportPages

on FileExists(theFile) -- (String) as Boolean
    tell application "System Events"
        if exists file theFile then
            return true
        else
            return false
        end if
    end tell
end FileExists

(*
 Here are the codes that apply: 4=UTF8, 12=windows latin, 30=MacRoman
 Taken from: https://gist.github.com/idStar/61994506d69595da3d30
 More in : https://discussions.apple.com/thread/4018778?tstart=0 
*)
on ensureUTF8Encoding()
	do shell script "/usr/bin/defaults write com.apple.iWork.Numbers CSVExportEncoding -int 4"
end ensureUTF8Encoding

