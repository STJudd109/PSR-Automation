NY:


Word := ComObjCreate("Word.Application")
worddoc := Word.documents.open( "" A_WorkingDir "\template.doc")
Word.Visible := true
wdAlignParagraphCenter := 1
wdAlignParagraphLeft := 0
wdAlignParagraphJustify := 3
wdWrapInline := 7
wdWrapThrough := 2

if (City = "Amityville" || City = "Babylon" || City = "Copiague" || City = "Deer Park" || City = "Farmingdale" || City = "Lindenhurst" || City = "North Amityville" || City = "North Babylon" || City = "West Babylon" || City = "Wyandach") {
	County := "Babylon"
}
else if (City = "Bellport" || City = "Blue Point" || City = "Brookhaven" || City = "Calverton" || City = "Center Moriches" || City = "Centereach" || City = "Coram" || City = "East Patchoque" || City = "East Moriches" || City = "East Setauket" || City = "Eastport" || City = "Farmingville" || City = "Holbrook" || City = "Holtsville" || City = "Lake Grove" || City = "Manorville" || City = "Mastic" || City = "Mastic Beach" || City = "Medford" || City = "Middle Island" || City = "Miller Place" || City = "Moriches" || City = "Mount Sinai" || City = "North Patchogue" || City = "Patchogue" || City = "Port Jefferson" || City = "Port Jefferson Station" || City = "Ridge" || City = "Rocky Point" || City = "Ronkonkoma" || City = "Selden" || City = "Shirley" || City = "Shoreham" || City = "Sound Beach" || City = "South Setauket" || City = "Stony Brook" || City = "Upton" || City = "Wading River" || City = "Yaphank") {
	County := "Brookhaven"
}
else if (City = "Bayport" || City = "Bay Shore" || City = "Bohemia" || City = "Brentwood" || City = "Brightwaters" || City = "Central Islip" || City = "East Islip" || City = "Greal Hiver" || City = "Hauppage" || City = "Islandia" || City = "Islip" || City = "Islip Terrace" || City = "Oakdale" || City = "Ocean Beach" || City = "Sayville" || City = "West Sayville") {
	County := "Islip"
}
else if (City = "East Meadow" || City = "North Bellmore" || City = "Seaford") {
	County := "Nassau"
}

if (WC = true) {
	InitialOffice := "Wyssling Consulting"
}
else {
	InitialOffice := "this office"
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
gosub, newline

Word.Selection.TypeText("Mr. Dan Rock, Project Manager")
gosub, newline
Word.Selection.TypeText("Vivint Solar")
gosub, newline
Word.Selection.TypeText("120 Fairchild Avenue")
gosub, newline
Word.Selection.TypeText("Plainview, NY 11803")
gosub, newline
gosub, newline

Word.Selection.Paragraphs.Indent
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


if(County = "Babylon") {
 	TemplateSelect = 1
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, " InitialOffice " initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections.  The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
 	;Paragraph 2
 	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance with our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in our report. This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on applicable building codes, professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }

else if(County = "Brookhaven") {
 	TemplateSelect = 2
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided the post installation photos for the above referenced solar panel installation. The site visit was conducted on " InstallDate ". As you are aware, " InitialOffice " initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The panel support locations and spacing are in conformance with the structural assessment. Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof at roof locations.")
 	gosub, newline
	gosub, newline
 	;Paragraph 2
 	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance with our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in the report. This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on the NYS Residential Code and the provisions of ASCE 7-05, professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information, do not hesitate to contact me.")
 }

else if(County = "Islip") {
 	TemplateSelect = 3
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation. As you are aware, " InitialOffice " initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The panel support locations and spacing are in conformance with the structural assessment. Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance with our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in the report. This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on applicable building codes, professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
 	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information, do not hesitate to contact me.")
 }

else if(WC = true) {
 	TemplateSelect = 4
 	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
 	;Paragraph 1
 	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  While the original structural review was not performed by this office, we have evaluated the installation in conjunction with the original design drawings provided by your office to verify panel layout and connection spacing only.  The photographs show panel support and spacing locations which are consistent with the drawings provided.  Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
 	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText("Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to the original design drawings, " MountType " product installation criteria, and the layout plan provided.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
 	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText("This certification is based on applicable building codes, professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
 	gosub, newline
	gosub, newline
	;Paragraph 4
	Word.Selection.TypeText("Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
 }

else {
	TemplateSelect = 0
	Word.Selection.TypeText( "Dear Mr. Rock:" )
	gosub, newline
	gosub, newline
	gosub, newline
	Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
	;Paragraph 1
	Word.Selection.TypeText("Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, " InitialOffice " initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections.  The photographs show panel support locations and spacing which conform to our structural assessment.  Acceptable minor changes to the layout include panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
	gosub, newline
	gosub, newline
	;Paragraph 2
	Word.Selection.TypeText( "Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in our report.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections." )
	gosub, newline
	gosub, newline
	;Paragraph 3
	Word.Selection.TypeText( "This certification is based on applicable building codes, professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only." )
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
if (TemplateSelect = 1)
{
	worddoc.Shapes.AddPicture(file,false,true,-10,475,158,87)
	worddoc.Shapes.AddPicture(ScottNY,false,true,300,455)
}
else if (TemplateSelect = 2)
{
	worddoc.Shapes.AddPicture(file,false,true,-10,502,158,87)
	worddoc.Shapes.AddPicture(ScottNY,false,true,300,480)
}
else if (TemplateSelect = 3)
{
	worddoc.Shapes.AddPicture(file,false,true,-10,463,158,87)
	worddoc.Shapes.AddPicture(ScottNY,false,true,300,445)
}
else if (TemplateSelect = 4)
{
	worddoc.Shapes.AddPicture(file,false,true,-10,490,158,87)
	worddoc.Shapes.AddPicture(ScottNY,false,true,300,465)
}
else
{
	worddoc.Shapes.AddPicture(file,false,true,-10,475,158,87)
	worddoc.Shapes.AddPicture(ScottNY,false,true,300,455)
}
gosub, newline
gosub, newline
;Word.Selection.ShapeRange.WrapFormat.Type := wdWrapInline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Scott Wyssling" )
gosub, newline
Word.Selection.TypeText( "NY License No. 092303" )
gosub, wordSave
;Word.ActiveDocument.SaveAs("" A_WorkingDir "\" Snum "-psr.doc")

gosub, window1
return
