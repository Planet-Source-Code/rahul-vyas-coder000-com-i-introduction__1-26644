<div align="center">

## COM \- I \(Introduction\)


</div>

### Description

First tutorial of the series COM, Introduction about the series. The series will contain topics like COM, DCOM, MTS, Pointers, Dynamic Memory Allocation, etc. Plus some basic Data Structures like Collections, Structures, Arrays, Enum, etc.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rahul Vyas \(coder000\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rahul-vyas-coder000.md)
**Level**          |Intermediate
**User Rating**    |4.0 (40 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rahul-vyas-coder000-com-i-introduction__1-26644/archive/master.zip)





### Source Code

```
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./COM%20I%20-%20Introduction_files/filelist.xml">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Author>Rahul Vyas</o:Author>
 <o:LastAuthor>Rahul Vyas</o:LastAuthor>
 <o:Revision>51</o:Revision>
 <o:TotalTime>56</o:TotalTime>
 <o:Created>2001-08-25T18:07:00Z</o:Created>
 <o:LastSaved>2001-08-25T19:04:00Z</o:LastSaved>
 <o:Pages>2</o:Pages>
 <o:Words>494</o:Words>
 <o:Characters>2817</o:Characters>
 <o:Company>ABC</o:Company>
 <o:Lines>23</o:Lines>
 <o:Paragraphs>5</o:Paragraphs>
 <o:CharactersWithSpaces>3459</o:CharactersWithSpaces>
 <o:Version>9.2720</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
h1
	{mso-style-next:Normal;
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	page-break-after:avoid;
	mso-outline-level:1;
	font-size:13.0pt;
	mso-bidi-font-size:12.0pt;
	font-family:"Times New Roman";
	color:blue;
	mso-font-kerning:0pt;}
p.MsoBodyText, li.MsoBodyText, div.MsoBodyText
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	font-style:italic;}
p.MsoBodyText2, li.MsoBodyText2, div.MsoBodyText2
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";
	font-weight:bold;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;
	text-underline:single;}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 1.25in 1.0in 1.25in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
 /* List Definitions */
@list l0
	{mso-list-id:1474831830;
	mso-list-type:hybrid;
	mso-list-template-ids:-1802442776 67698703 67698713 67698715 67698703 67698713 67698715 67698703 67698713 67698715;}
ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>
<!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1027">
 <o:colormenu v:ext="edit" fillcolor="none"/>
 </o:shapedefaults></xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
 <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>
<body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:20.0pt;mso-bidi-font-size:12.0pt;color:#333399'>COM &#8211; I
(INTRODUCTION)<o:p></o:p></span></b></p>
<p class=MsoNormal align=center style='text-align:center'><b><span
style='font-size:20.0pt;mso-bidi-font-size:12.0pt;color:#333399'>[</span></b><b><span
style='font-size:20.0pt;mso-bidi-font-size:12.0pt;color:red'>Visual Basic 6</span></b><b><span
style='font-size:20.0pt;mso-bidi-font-size:12.0pt;color:#333399'>]</span></b><span
style='font-size:20.0pt;mso-bidi-font-size:12.0pt'><o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><i>Friends</i>, I am planning to post a complete series of
Tutorials under the title COM. COM, because basically we will be dealing with <i>components</i>.
The topics that we will cover in the series are:</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<ol style='margin-top:0in' start=1 type=1>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Introduction</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Using DLLs
   - Automation</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Introduction
   to ADO</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>ADO
   (continued)</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Arrays,
   Functions &amp; Sub - Routines</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Files,
   Structures, Enumerations &amp; Collections</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Dynamic
   Memory Allocation &amp; Pointers in VB</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Objects
   &amp; Classes</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Active
   &#8211; X DLL &#8211; Making our own COM</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Interfaces</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Collection
   Classes</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Project
   to show the practical use of Components (covering collection classes,
   interfaces, etc.)</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Active
   &#8211; X EXE</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>DCOM</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>MTS</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Miscellaneous
   &amp; Left Overs</li>
 <li class=MsoNormal style='mso-list:l0 level1 lfo3;tab-stops:list .5in'>Summary</li>
</ol>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>If at any point of time you need help, you may contact me
at: <a href="mailto:rahulreceive@hotmail.com">rahulreceive@hotmail.com</a>.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>The purpose of this series is to make new as well as
experienced programmers, get some in-depth idea of the above mentioned topics.
Many of these topics may not be called as parts of COM. But, I have included
them to make them clear so that we may use them in our components. <b>Data
Structures</b> (at least basic) must be clear! Don&#8217;t you think so?</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Some topic might seem like C there like <b>Pointers</b> and <b>Dynamic
Memory Allocation</b>. These are very important concepts and need to be seen at
least.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>After this point we are beginning with <b>Object Oriented
Concepts</b> and their implementation in VB. Here we will have in-depth
knowledge of <b>Objects</b>, <b>Classes</b>, COM, etc. Then, we will see some very
important concepts like Interfaces and their implementation in VB. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>After all this it is better to go for some Practical example
to show the usage of various concepts.</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Then we will cover <b>Active &#8211; X EXE</b>, <b>DCOM </b>and <b>MTS</b>.
<i>But, we will not be covering asynchronous callbacks here.<o:p></o:p></i></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>All this will greatly depend on your participation and
interest. So, I would feel good, if you give your precious suggestions and
comments. </p>
<div style='border:none;border-bottom:solid windowtext .75pt;padding:0in 0in 1.0pt 0in'>
<p class=MsoNormal style='border:none;mso-border-bottom-alt:solid windowtext .75pt;
padding:0in;mso-padding-alt:0in 0in 1.0pt 0in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal style='border:none;mso-border-bottom-alt:solid windowtext .75pt;
padding:0in;mso-padding-alt:0in 0in 1.0pt 0in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><span style="mso-spacerun: yes"> </span></p>
<p class=MsoNormal><b><u>What is COM?<o:p></o:p></u></b></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText>COM stands for &#8211; Component Object Model. It is basically a
concept which suggests that a larger piece of code must be broken down into
smaller components (specialized for some particular task). These components are
integrated and used as a whole. </p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><b><u>Why use COM?<o:p></o:p></u></b></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><i>Answer to this question is basically based on feeling&#8230;<o:p></o:p></i></p>
<p class=MsoNormal><i>See, Assume that you are given a task of making a
software for school.<o:p></o:p></i></p>
<p class=MsoNormal><i>Now, if you see as a whole a school is a very large
system. And, it is very difficult to implement the whole functionality of a
school.<o:p></o:p></i></p>
<p class=MsoNormal><i>So, if you get this project what will happen:<o:p></o:p></i></p>
<p class=MsoNormal><i>For some days you will continue that project, But soon it
will become unmanageable as you will start feeling bored. Somehow, you will
finish and the project will have many bugs.<o:p></o:p></i></p>
<p class=MsoNormal><i><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></i></p>
<p class=MsoNormal><i>It will be very difficult to remove the bugs as the code
is very unstructured. And, the efficiency of the overall system will be very
poor. The solution is, implement small-small functionalities on at a time
perfectly, and lastly integrate them. This is called COM.<o:p></o:p></i></p>
<p class=MsoNormal><i><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></i></p>
<p class=MsoBodyText>Writing code for smaller things is very easy, and you
constantly feel that yes, something has been done!</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText2>&lt;The exact example and need of COM will become more
clear when we go for Classes&gt;</p>
<div style='border:none;border-bottom:solid windowtext .75pt;padding:0in 0in 1.0pt 0in'>
<p class=MsoNormal style='border:none;mso-border-bottom-alt:solid windowtext .75pt;
padding:0in;mso-padding-alt:0in 0in 1.0pt 0in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoBodyText style='border:none;mso-border-bottom-alt:solid windowtext .75pt;
padding:0in;mso-padding-alt:0in 0in 1.0pt 0in'>There was nothing much today, It
was just introduction. The actual thing will start from next one.</p>
<p class=MsoNormal style='border:none;mso-border-bottom-alt:solid windowtext .75pt;
padding:0in;mso-padding-alt:0in 0in 1.0pt 0in'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><u>NOTE</u><span style='color:red'>: Your <b>feedback</b> is
the fuel by which the series will be moving&#8230;&#8230;<o:p></o:p></span></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal>Bye,</p>
<p class=MsoNormal>Best of Luck,</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<h1><span style='font-size:14.0pt;mso-bidi-font-size:12.0pt;color:#333399'>Rahul
Vyas<o:p></o:p></span></h1>
<p class=MsoNormal>(<i><a href="mailto:rahulreceive@hotmail.com">rahulreceive@hotmail.com</a></i>)</p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
<p class=MsoNormal><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></p>
</div>
</body>
</html>
```

