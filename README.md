# Word-to-HTML
<html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=Generator content="Microsoft Word 15 (filtered)">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
	{font-family:Aptos;}
@font-face
	{font-family:"Aptos Display";}
 /* Style Definitions */
 h1
	{mso-style-link:"Heading 1 Char";
	margin-top:18.0pt;
	margin-right:0cm;
	margin-bottom:4.0pt;
	margin-left:0cm;
	text-align:right;
	line-height:115%;
	page-break-after:avoid;
	direction:rtl;
	unicode-bidi:embed;
	font-size:20.0pt;
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;
	font-weight:normal;}
p.MsoNoSpacing, li.MsoNoSpacing, div.MsoNoSpacing
	{margin:0cm;
	text-align:right;
	direction:rtl;
	unicode-bidi:embed;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-link:"Heading 1";
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;}
.MsoChpDefault
	{font-size:12.0pt;}
.MsoPapDefault
	{margin-bottom:8.0pt;
	line-height:115%;}
@page WordSection1
	{size:595.3pt 841.9pt;
	margin:36.0pt 36.0pt 36.0pt 36.0pt;}
div.WordSection1
	{page:WordSection1;}
 /* List Definitions */
 ol
	{margin-bottom:0cm;}
ul
	{margin-bottom:0cm;}
-->
</style>

</head>

<body lang=EN-US style='word-wrap:break-word'>

<div class=WordSection1 dir=RTL>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'><span class=Heading1Char><span style='font-size:20.0pt;color:black'>This
VBA macro for Microsoft Word is designed to convert the contents of a Word
document into a UTF-8 encoded text file. This is particularly useful for
creating formatted text to be used in environments like Jupyter Notebook</span></span><span
dir=RTL></span><span lang=FA dir=RTL style='font-family:"Arial",sans-serif'><span
dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'>&nbsp;</p>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'><b>Execution Steps</b><span dir=RTL></span><span lang=FA dir=RTL
style='font-family:"Arial",sans-serif'><span dir=RTL></span>:</span></p>

<p class=MsoNoSpacing dir=LTR style='margin-left:36.0pt;text-align:left;
text-indent:-18.0pt;direction:ltr;unicode-bidi:embed'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><span dir=LTR></span>It checks the current document's save path. If the
document is not saved, the system's temporary folder (<span style='background:
yellow'>TEMP</span>) is used<span dir=RTL></span><span lang=FA dir=RTL
style='font-family:"Arial",sans-serif'><span dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='margin-left:36.0pt;text-align:left;
text-indent:-18.0pt;direction:ltr;unicode-bidi:embed'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><span dir=LTR></span>The Word document is saved as Filtered HTML <span
style='background:yellow'>(.html)</span>, preserving its raw text and structure<span
dir=RTL></span><span lang=FA dir=RTL style='font-family:"Arial",sans-serif'><span
dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='margin-left:36.0pt;text-align:left;
text-indent:-18.0pt;direction:ltr;unicode-bidi:embed'>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><span dir=LTR></span>The HTML content is read and stored in a variable<span
dir=RTL></span><span lang=FA dir=RTL style='font-family:"Arial",sans-serif'><span
dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='margin-left:36.0pt;text-align:left;
text-indent:-18.0pt;direction:ltr;unicode-bidi:embed'>4.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><span dir=LTR></span>A new text file with UTF-8 encoding is created, and
the extracted data is written to it<span dir=RTL></span><span lang=FA dir=RTL
style='font-family:"Arial",sans-serif'><span dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='margin-left:36.0pt;text-align:left;
text-indent:-18.0pt;direction:ltr;unicode-bidi:embed'>5.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span><span dir=LTR></span>The temporary HTML file is deleted, and a success
message is displayed<span dir=RTL></span><span lang=FA dir=RTL
style='font-family:"Arial",sans-serif'><span dir=RTL></span>.</span></p>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'><b>&nbsp;</b></p>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'><b>This macro enables you to easily generate formatted text for use in
Jupyter Notebook or other environments while ensuring proper encoding and
formatting.</b></p>

<p class=MsoNoSpacing dir=LTR style='text-align:left;direction:ltr;unicode-bidi:
embed'><b>---------------------------------------------------------------------------------------------------------------------------------------</b></p>

<p class=MsoNoSpacing dir=RTL><b><span lang=FA style='font-family:"Calibri",sans-serif'>&nbsp;</span></b></p>

</div>

</body>

</html>
