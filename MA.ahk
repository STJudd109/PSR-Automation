MA:


Word := ComObjCreate("Word.Application")
worddoc := Word.documents.open( "" A_WorkingDir "\template.doc")
Word.Visible := true
wdAlignParagraphCenter := 1
wdAlignParagraphLeft := 0
wdAlignParagraphJustify := 3
wdWrapInline := 7
wdWrapThrough := 2
if (UpgradeYes = true) {
	Upgrade := "true"
}
else {
	Upgrade := "false"
}

if (MountType = "MSI")
{
	MountType := "Mounting System Inc."
}
else if (MountType = "ECO")
{
	MountType := "Ecolibrium Solar"
}
else if (MountType = "ZEP")
{
	MountType := "ZEP Company"
}

TemplateSelect := 0

Word.Selection.InsertDateTime("MMMM dd, yyyy")

Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphCenter
Word.ActiveWindow.Selection.TypeParagraph
Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphLeft

gosub, newline

Word.Selection.TypeText("Mr. Dan Rock, Project Manager")
gosub, newline
Word.Selection.TypeText("Vivint Solar")
gosub, newline
Word.Selection.TypeText("24 Normac Road")
gosub, newline
Word.Selection.TypeText("Woburn, MA 01801")
gosub, newline
gosub, newline

Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent
Word.Selection.Paragraphs.Indent

Word.Selection.TypeText("Re:      Post Structural Certification")
gosub, newline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText("" LastName " Residence" )
gosub, newline
Word.Selection.TypeText("" Address ", " City ", " State )
gosub, newline
Word.Selection.TypeText("S-" Snum "")
gosub, newline
Word.Selection.TypeText("" Kw " kW System" )
gosub, newline

Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent
Word.Selection.Paragraphs.Outdent


;Post Structural Chelmsford Template
 if(City = "Chelmsford") {

 	TemplateSelect = 1
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The panel support locations and spacing are in conformance with the structural assessment. Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
 	;Paragraph 2
 	Word.Selection.TypeText("The existing residence is typical wood framing construction with the roof system consisting of " RafterSize " dimensional lumber at " RafterSpacing " on center with " CollarTieSize " collar ties every " CollarTieSpacing """. Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in our report. No structural changes have been made to the roof structure since the original structural review. This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }
 ;Post Structural NON WC Template
 else if(WC = "true") {
 	TemplateSelect = 2
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  While the original structural review was not performed by this office, we have evaluated the installation in conjunction with the original design drawings provided by your office to verify panel layout and connection spacing only.  The photographs show panel support and spacing locations which are consistent with the drawings provided.  Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
 	;Paragraph 2
 	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to the original design drawings, " MountType " product installation criteria, and the layout plan provided.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }
 ;Post Structural ECO Knee Wall Upgrade Template
 else if(Upgrade = "true" and UpgradeType = "Knee Wall") {
 	TemplateSelect = 3
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", Ecolibrium Solar product installation criteria, and the layout plan as specified in our report.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("We further certify that the installed knee wall was constructed in conformance with our direction and requirements.")
 	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations and connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 5
	Word.Selection.TypeText("This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
 	gosub, newline
	gosub, newline
	;Paragraph 6
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }
 ;Post Structural ECO Template
 else if(Upgrade = "false" and MountType = "ECO") {
 	TemplateSelect = 4
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", Ecolibrium Solar product installation criteria, and the layout plan as specified in our report.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
 	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }
;Post Structural ECO Sister Rafters Upgrade Template
 else if(Upgrade = "true" and UpgradeType = "Sister Rafters") {
 	TemplateSelect = 5
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", Ecolibrium Solar product installation criteria, and the layout plan as specified in our report.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("We further certify that the installed sister rafters are an acceptable structural upgrade method and were constructed in conformance with our direction and requirements.")
 	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations and connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 5
	Word.Selection.TypeText("This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
 	gosub, newline
	gosub, newline
	;Paragraph 6
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }
 ;Post Structural Template
else {
	TemplateSelect = 0
	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
	;Paragraph 1
	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections.  The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText( "Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in our report.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections." )
	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText( "This certification is based on the 8th Edition Residential Code (2009 International Residential Code with Massachusetts Amendments), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only." )
	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText( "Should you have any questions regarding the above or if you require additional information do not hesitate to contact me." )
 }
gosub, newline
gosub, newline
gosub, newline

Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Very truly yours," )
gosub, newline
Word.Selection.Paragraphs.Outdent
if (TemplateSelect = 0) {
	worddoc.Shapes.AddPicture(file,false,true,-10,463,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,440)
}
else if (TemplateSelect = 1) {
	worddoc.Shapes.AddPicture(file,false,true,-10,503,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,490)
}
else if (TemplateSelect = 2) {
	worddoc.Shapes.AddPicture(file,false,true,-10,463,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,440)
}
else if (TemplateSelect = 3) {
	worddoc.Shapes.AddPicture(file,false,true,-10,513,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,490)
}
else if (TemplateSelect = 4) {
	worddoc.Shapes.AddPicture(file,false,true,-10,463,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,440)
}
else if (TemplateSelect = 5) {
	worddoc.Shapes.AddPicture(file,false,true,-10,513,158,87)
	worddoc.Shapes.AddPicture(ScottMA,false,true,300,490)
}

gosub, newline
gosub, newline
;Word.Selection.ShapeRange.WrapFormat.Type := wdWrapInline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Scott Wyssling" )
gosub, newline
Word.Selection.TypeText( "MA License No. 50507" )

gosub, wordSave
;Word.ActiveDocument.SaveAs("" A_WorkingDir "\" Snum "-psr.doc")

gosub, window1
return


;Word.Selection.TypeText("")
 	;gosub, newline
	;gosub, newline