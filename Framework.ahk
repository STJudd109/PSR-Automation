#SingleInstance 
#NoEnv
SendMode, Input
FormatTime, Now ,, M/dd/yyyy hh:mm tt
FormatTime, Nowdate ,, MMMM/dd/yyyy
FormatTime, todaysFolder ,, dd MMM yy
FormatTime, Year ,, yy
FormatTime, Month ,, MMMM
FormatTime, MonthNumber ,, MM
SetWorkingDir %A_ScriptDir%
FileCreateDir, %A_WorkingDir%\Stamps
Gosub, generator
file := % A_WorkingDir "\Stamps\signature.png" 
ScottMA := % A_WorkingDir "\Stamps\MA.png"
ScottCA := % A_WorkingDir "\Stamps\scottCAstamp_clear.png"
ScottMD := % A_WorkingDir "\Stamps\scottmdwatermark.png"
ScottCT :=  % A_WorkingDir "\Stamps\connstamp.png"
ScottNY :=  % A_WorkingDir "\Stamps\nytestsize4.png"
Gui, Color, 0B787F
Gui, Add, Button, x72 y190 w100 h30 gwindow2 , Start
Gui, Add, Edit, x42 y100 w150 h20 vAccountList ,
Gui, Add, Text, x42 y10 w160 h40 +Center, AHK Automation
Gui, Font, S12 CDefault Bold, Verdana
; Generated using SmartGUI Creator 4.0
Gui, Show, x195 y105 w246 h274, Start
Return


wordClear:
run, taskkill /im winword.exe /f,, hide
return


window1:
Gui, Color, 0B787F
Gui, Add, Button, x72 y190 w100 h30 gwindow3, Start
Gui, Add, Button, x72 y130 w100 h30 gwindow2, Edit Last
Gui, Add, Text, x42 y10 w160 h40 +Center, AHK Automation
Gui, Font, S12 CDefault Bold, Verdana
; Generated using SmartGUI Creator 4.0
Gui, Show, x195 y105 h274 w246, Start
Return

GuiClose:
ExitApp


window3:
Gui, Destroy
Gui, Color, 0B787F
Gui, Add, Button, x72 y190 w100 h30 gwindow2 , Start
Gui, Add, Edit, x42 y100 w150 h20 vAccountList ,
Gui, Add, Text, x42 y10 w160 h40 +Center, AHK Automation
Gui, Font, S12 CDefault Bold, Verdana
; Generated using SmartGUI Creator 4.0
Gui, Show, x195 y105 w246 h274, Start
Return
;Gui, Add, Edit, x42 y100 w150 h20 vAccountList ,


window2:
Gui, Submit
Gui, Destroy
counter = 0;
Gosub, accountParse
if (State = "CA")
{
	vOption = Select State:|CA||CT|MA|MD|NJ|NY 
}
else if (State = "CT")
{
	vOption = Select State:|CA|CT||MA|MD|NJ|NY
}
else if (State = "MA")
{
	vOption = Select State:|CA|CT|MA||MD|NJ|NY
}
else if (State = "MD")
{
	vOption = Select State:|CA|CT|MA|MD||NJ|NY
}
else if (State = "NJ")
{
	vOption = Select State:|CA|CT|MA|MD|NJ||NY
}
else if (State = "NY")
{
	vOption = Select State:|CA|CT|MA|MD|NJ|NY||
}
else
{
	vOption = Select State:||CA|CT|MA|MD|NJ|NY
}

if (vMountType = "MSI")
{
	vMountType = Select Mount Type:|MSI||ECO|ZEP
}
else if (vMountType = "ECO")
{
	vMountType = Select Mount Type:|MSI|ECO||ZEP
}
else if (vMountType = "ZEP")
{
	vMountType = Select Mount Type:|MSI|ECO|ZEP||
}
Else
{
	vMountType = Select Mount Type:||MSI|ECO|ZEP
}

Gui, Font, cffffff
Gui, Add, Text, x32 y17, Snumber:
Gui, Font, c000000
Gui, Add, Edit, x32 y30 w150 h32 vSnum, %Snum%
Gui, Font, cffffff
Gui, Add, Text, x32, Last Name:
Gui, Font, c000000
Gui, Add, Edit, x32 y82 w150 h32 vLastName, %LastName%
Gui, Font, cffffff
Gui, Add, Text, x32, Address:
Gui, Font, c000000
Gui, Add, Edit, x32 y134 w150 h32 vAddress, %Address%
Gui, Font, cffffff
Gui, Add, Text, x32, City:
Gui, Font, c000000
Gui, Add, Edit, x32 y186 w150 h32 vCity, %City%
Gui, Font, cffffff
Gui, Add, Text, x32, kW:
Gui, Font, c000000
Gui, Add, Edit, x32 y238 w150 h32 vKw, %kw%
Gui, Font, cffffff
Gui, Add, Text, x32, First SR Date(M/D/Y):
Gui, Font, c000000
Gui, Add, Edit, x32 y290 w150 h32 vsrDate, %SRDate%
Gui, Add, ListBox, x207 y30 h100 vState, %vOption%
Gui, Add, ListBox, x207 y139 h60 vMountType, %vMountType%
Gui, Font, cffffff
Gui, Add, Text, x207 y206, SR Completed By:
Gui, Add, Radio, Checked x207 y224 vThisOffice, This Office
Gui, Add, Radio, x285 y224 vWC, WC
Gui, Add, Text, x207 y249, Was There an Upgrade?
vUpgradeValueNo = 1
Gui, Add, Radio, Checked x207 y267 gUpgradeNo, No
vUpgradeValueYes = 0
Gui, Add, Radio, x247 y267 gUpgradeYes, Yes
Gui, Add, DropDownList, Disabled%UpgradeValueYes% x207 y287 vUpgradeType, Select:||Knee Wall|Sister Rafters|Purlin

;Gui, Add, MonthCal, x160 y280 yCalendar, Calendar
Gui, Add, Button, x132 y340 w100 h30 gdecideNextWindow, Submit
; Generated using SmartGUI Creator 4.0

Gui, Color, 0B787F
Gui, Show, x343 y198 w360 h400, Account Info
gosub, wordClear
Return


UpgradeNo:
if vUpgradeValueNo = 0
{
	GuiControl, Disable, UpgradeType
	vUpgradeValueNo = 1
	vUpgradeValueYes = 0
	UpgradeYes = 0
}
return


UpgradeYes:
if vUpgradeValueYes = 0
{
	GuiControl, Enable, UpgradeType
	vUpgradeValueYes = 1
	vUpgradeValueNo = 0
	UpgradeYes = 1
}
return


WindowExtraRafterCollar:
Gui, Font, cffffff
Gui, Add, Text, x50 y20 , Rafter Size:
Gui, Font, c000000
Gui, Add, Edit, x50 y35 w60 vRafterSize
Gui, Font, cffffff
Gui, Add, Text, x130 y20 , Rafter Spacing:
Gui, Font, c000000
Gui, Add, Edit, x130 y35 w60 vRafterSpacing
Gui, Font, cffffff
Gui, Add, Text, x50 y65 , Collar Tie Size:
Gui, Font, c000000
Gui, Add, Edit, x50 y80 w60 vCollarTieSize
Gui, Font, cffffff
Gui, Add, Text, x130 y65 , Collar Tie Spacing:
Gui, Font, c000000
Gui, Add, Edit, x130 y80 w60 vCollarTieSpacing
Gui, Add, Button, x72 y120 w100 h30 gwordSend, Submit
Gui, Color, 0B787F
Gui, Show, x343 y298 h174 w246, Extra info
return


WindowExtraSiteVisit:
Gui, Font, cffffff
Gui, Add, Text, x67 y40 , Site Visit Date(M/D/Y):
Gui, Font, c000000
Gui, Add, Edit, x67 y60 w110 vInstallDate
Gui, Add, Button, x72 y120 w100 h30 gwordSend, Submit
Gui, Color, 0B787F
Gui, Show, x343 y298 h174 w246, Extra info
return


decideNextWindow:
Gui, Submit
Gui, Destroy
if (Snum = "" || LastName = "" || Address = "" || City = "" || Kw = "" || srDate = "" || State = "" || MountType = "")
{
	Msgbox Fill out all values.  Thanks.
	Gosub, window2
}
Else
{
	if (City = "Bellport" || City = "Blue Point" || City = "Brookhaven" || City = "Calverton" || City = "Center Moriches" || City = "Centereach" || City = "Coram" || City = "East Patchoque" || City = "East Moriches" || City = "East Setauket" || City = "Eastport" || City = "Farmingville" || City = "Holbrook" || City = "Holtsville" || City = "Lake Grove" || City = "Manorville" || City = "Mastic" || City = "Mastic Beach" || City = "Medford" || City = "Middle Island" || City = "Miller Place" || City = "Moriches" || City = "Mount Sinai" || City = "North Patchogue" || City = "Patchogue" || City = "Port Jefferson" || City = "Port Jefferson Station" || City = "Ridge" || City = "Rocky Point" || City = "Ronkonkoma" || City = "Selden" || City = "Shirley" || City = "Shoreham" || City = "Sound Beach" || City = "South Setauket" || City = "Stony Brook" || City = "Upton" || City = "Wading River" || City = "Yaphank") 
	{
		County := "Brookhaven"
	}

	if (City = "Chelmsford" && State = "MA")
	{
		Gosub, WindowExtraRafterCollar
	}
	else if(County = "Brookhaven" && State = "NY")
	{
		Gosub, WindowExtraSiteVisit
	}
	Else
	{
		Gosub, wordSend
	}
	return
}


newline:
	Word.ActiveWindow.Selection.TypeParagraph
return

#Include CA.ahk
#Include CT.ahk
#Include MA.ahk
#Include MD.ahk
#Include NJ.ahk
#Include NY.ahk


wordSend:
Gui, Submit
Gui, Destroy
run, Chrome.exe https://mercury.vivintsolar.com/#/account/%Snum%

if (State = "CA") {
	gosub, CA
}
else if(State = "CT") {
	gosub, CT
}
else if(State = "MA") {
	gosub, MA
}
else if(State = "MD") {
	gosub, MD
}
else if(State = "NJ") {
	gosub, NJ
}
else if(State = "NY") {
	gosub, NY
}
else {
	Msgbox , "Pick a State next time!"
	ExitApp
}
Return


wordSave:
wdExportFormatPDF := 17
;savePath := % MonthNumber"-" %Month% "\" %todaysFolder%
if (Office = "NJ-01") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-01\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-1\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
if (Office = "NJ-02") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-02\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-2\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
if (Office = "NJ-03") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-03\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-3\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)	
}
if (Office = "NJ-04") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-04\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-4\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)	
}
if (Office = "NJ-05") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-05\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-5\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
if (Office = "NJ-06") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NJ-06\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NJ-6\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
if (Office = "NY-05") {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\NY-05\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\" State "\NY-5\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
else {
Word.ActiveDocument.SaveAs("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\" todaysFolder "\" State "\S-" Snum " " Address "_PSR.doc")
Word.ActiveDocument.ExportAsFixedFormat("\\Media01\solar\Structural Engineering\01 Vivint 2016\" MonthNumber "-" Month "\Completed files\" todaysFolder "\S-" Snum " " Address "_PSR.PDF",wdExportFormatPDF)
}
return


accountParse:
if AccountList = "" 
return
else {
Loop, Parse, AccountList, %A_Tab% 
{

If (A_Index = 1) {
Snum = %A_LoopField% 
}
if (A_Index = 2) {
LastName = %A_LoopField%
}
if (A_Index = 3) {
Address := A_LoopField
StringUpper, Address, Address , T
}
if (A_Index = 4) {
City := A_LoopField
StringUpper, City, City , T
}
if (A_Index = 5) {
State = %A_LoopField%
}
if (A_Index = 6) {
Kw = %A_LoopField%
}
if (A_Index = 7) {
SRDate = %A_LoopField%
}
if (A_Index = 8) {
Office = %A_LoopField%
}
}
}

Loop, Parse, Office, " ", 
{
if (A_Index = 1) {
Office = %A_LoopField%
}
}

Loop, Parse, Snum, "-", 
{
if (A_Index = 2) {
Snum = %A_LoopField%
}
}


Loop, Parse, Lastname, " ", 
{
LastName = %A_LoopField%
}

Loop, Parse, SRDate, /, 
{

if (A_index = 1)
{
if (A_LoopField = 01) 
{
SRDate1 := "January"
}
if (A_LoopField = 02) 
{
SRDate1 := "February"
}
if (A_LoopField = 03) 
{
SRDate1 := "March"
}
if (A_LoopField = 04) 
{
SRDate1 := "April"
}
if (A_LoopField = 05) 
{
SRDate1 := "May"
}
if (A_LoopField = 06) 
{
SRDate1 := "June"
}
if (A_LoopField = 07) 
{
SRDate1 := "July"
}
if (A_LoopField = 08) 
{
SRDate1 := "August"
}
if (A_LoopField = 09) 
{
SRDate1 := "September"
}
if (A_LoopField = 10) 
{
SRDate1 := "October"
}
if (A_LoopField = 11) 
{
SRDate1 := "November"
}
if (A_LoopField = 12) 
{
SRDate1 := "December"
}

}
if (A_Index = 2) 
{
SRDate = %SRDate1% %A_LoopField%
}
if (A_Index = 3) 
{
timesplit = %A_LoopField%

Loop, Parse, timesplit, " ",
{
	if (A_Index = 1) {
		timesplit = %A_LoopField%
		Break
	}
}

SRDate = %SRDate%, %timesplit%
}

}
return


generator:
FileInstall, Template.doc, %A_WorkingDir%\Template.doc ,
FileInstall, Stamps\signature.png, %A_WorkingDir%\Stamps\signature.png ,
FileInstall, Stamps\MA.png, %A_WorkingDir%\Stamps\MA.png ,
FileInstall, Stamps\scottCAstamp_clear.png, %A_WorkingDir%\Stamps\scottCAstamp_clear.png ,
FileInstall, Stamps\scottmdwatermark.png, %A_WorkingDir%\Stamps\scottmdwatermark.png ,
FileInstall, Stamps\connstamp.png, %A_WorkingDir%\Stamps\connstamp.png ,
FileInstall, Stamps\nytestsize4.png, %A_WorkingDir%\Stamps\nytestsize4.png ,
gosub, folders
return


folders:
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\CA  
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\CT
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\MA
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\MD
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\NJ 
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\%todaysFolder%\NY
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\CA
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\CT
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\MA 
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\MD
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\NJ
FileCreateDir, \\Media01\solar\Structural Engineering\01 Vivint 2016\Post Structurals\%MonthNumber%-%Month%\Completed files\%todaysFolder%\NY
return