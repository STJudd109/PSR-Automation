MD:


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
Word.Selection.TypeText("7030 Virginia Manor Road")
gosub, newline
Word.Selection.TypeText("Beltsville, MD 20705")
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
Word.Selection.TypeText( LotNum )
gosub, newline
Word.Selection.TypeText( "Building Permit # " BuildingPermitNum )
gosub, newline
Word.Selection.TypeText( "Eledtrical Permit # " ElectricalPermitNum )
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
;Paragraph 1
Word.Selection.TypeText("Pursuant to your request, a representative from our company (VivintSolar.inc), conducted a post installation site visit under my supervision and provided the post installation photos for the above referenced solar panel installation. The site visit was conducted on " SRDate ". As you are aware, this office initially prepared a structural assessment of the proposed solar panel installation, the adequacy of the connections for this system and identified maximum spacing of the connections. The panel support locations and spacing are in conformance with the structural assessment. Acceptable minor changes to the layout include the panel position, support spacing less than or equal to 64"", and/or deletions of panels at roof, at locations.")
gosub, newline
gosub, newline
;Paragraph 2
Word.Selection.TypeText( "Based upon the post installation site visit, our office certifies the solar panel installation for this roof and that it was in conformance with our structural assessment report dated " SRDate ", Ecolibrium solar product installation criteria, and the layout plan as specified in the report. I certify that the existing roof structure can support all loads by the 2009 IRC, and the additional loads of the solar panels." )
gosub, newline
gosub, newline
;Paragraph 3
Word.Selection.TypeText( "This certification is based on the International Residential Code, professional engineering assessment and judgment and covers this dwellings assessment for solar panels connections and support only. It is also our professional opinion that the structure is capable of safely supporting the design loads as required by the 2015 International Building Code (or the 2015 International Residential Code)." )
gosub, newline
gosub, newline
;Paragraph 4
Word.Selection.TypeText( "Should you have any questions regarding the above or if you require additional information, do not hesitate to contact me." )
gosub, newline
gosub, newline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Very truly yours," )
gosub, newline
Word.Selection.Paragraphs.Outdent
worddoc.Shapes.AddPicture(file,false,true,-10,523,158,87)
worddoc.Shapes.AddPicture(ScottMD,false,true,300,515)
gosub, newline
gosub, newline
;Word.Selection.ShapeRange.WrapFormat.Type := wdWrapInline
Word.Selection.Paragraphs.Indent
Word.Selection.TypeText( "Scott Wyssling, PE" )
gosub, newline
Word.Selection.TypeText( "MD PE License No. 43466" )
gosub, newline
Word.Selection.TypeText( "License Exp; 4/11/2017" )

gosub, wordSave
;Word.ActiveDocument.SaveAs("" A_WorkingDir "\" Snum "-psr.doc")

gosub, window1
return