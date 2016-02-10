CT:


Word := ComObjCreate("Word.Application")
worddoc := Word.documents.open( "" A_WorkingDir "\template.doc")
Word.Visible := true
wdAlignParagraphCenter := 1
wdAlignParagraphLeft := 0
wdAlignParagraphJustify := 3
wdWrapInline := 7
wdWrapThrough := 2

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
Word.Selection.TypeText("3301 North Thanksgiving Way, Suite 500")
gosub, newline
Word.Selection.TypeText("Lehi, UT 84043")
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

Word.Selection.TypeText( "Dear Mr. Rock:" )
gosub, newline
gosub, newline
gosub, newline
Word.Selection.ParagraphFormat.Alignment := wdAlignParagraphJustify
Word.Selection.TypeText( "Pursuant to your request, a representative from our company conducted a post installation site visit under my supervision and provided post installation photos for the above referenced solar panel installation.  As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The photographs show panel support locations and spacing which conform to our structural assessment. Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or additions or deletions of panels at roof locations.")
gosub, newline
gosub, newline
Word.Selection.TypeText( "Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance to our structural assessment report dated " SRDate ", " MountType " product installation criteria, and the layout plan as specified in our report.  This letter pertains only to the panel support attachments to the roof framing and not the engineered photovoltaic panel products, components, panel positioning, or electrical related installations/connections.")
gosub, newline
gosub, newline
Word.Selection.TypeText( "This certification is based on the 8th Edition Residential Code (2009 International Residential), professional engineering assessment and judgment and covers this dwellings assessment for solar panel connections and support only.")
gosub, newline
gosub, newline
Word.Selection.TypeText( "Should you have any questions regarding the above or if you require additional information do not hesitate to contact me.")
gosub, newline
gosub, newline
gosub, newline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Very truly yours," )
gosub, newline
Word.Selection.Paragraphs.Outdent
worddoc.Shapes.AddPicture(file,false,true,-10,490,158,87)
worddoc.Shapes.AddPicture(ScottCT,false,true,300,485)
gosub, newline
gosub, newline
;Word.Selection.ShapeRange.WrapFormat.Type := wdWrapInline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Scott Wyssling" )
gosub, newline
Word.Selection.TypeText( "CT License No. 30023" )
gosub, wordSave
;Word.ActiveDocument.SaveAs("" A_WorkingDir "\" Snum "-psr.doc")

gosub, window1
return
