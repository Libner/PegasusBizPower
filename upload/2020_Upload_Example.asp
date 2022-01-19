<%@ Language="VBScript" ENABLESESSIONSTATE="False" %>
<%
Option Explicit
response.buffer = true
%>
<!--#include file="_cls2020Upload.asp"-->
<%
dim strExamplePath
    strExamplePath = "E:\wwwroot\bgu4u_appeals\upload\" 'Replace(server.MapPath(request.servervariables("SCRIPT_NAME")),Mid(request.servervariables("SCRIPT_NAME"),2),"")

	IF request.querystring("action") = "upload" THEN

	dim objUpload
	set objUpload = New clsUpload
	    objUpload.RestrictFileExtentions = True
	    objUpload.SafeFileExtensions = "txt|rtf|inc|html|htm|jpg|jpeg|gif|bmp|png|tiff|tif|eps|psd|ps|mov|mpg|mpeg|avi|ram|wmf|rm|mp3|aiff|wav|mus|mid|rmf|pd|pdf|doc|pub|xls|mdb|ppt|wmv|wma|dxf|3ds|ac|asf|asx|au|wpd|kmz|asc|ans|tab|cvs"

		IF NOT objUpload.Upload(strExamplePath) THEN
		response.write(err.Number &": "& err.Description)
		response.end
		END IF

		IF objUpload.Files.Count = 0 THEN
		response.write("There were no files, or something went wrong.")
		response.end
		END IF

	dim item
	dim objItem

		FOR EACH item IN objUpload.Files 'there may be multiples

		set objItem = objUpload.Files(item)
		    objItem.Save
		set objItem = nothing

		NEXT

	set objUpload = nothing

	response.redirect "2020_Upload_Example.asp"

	END IF
	%>





<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en">
<head>
	<title>20/20 Applications ReadMe Documentation - 20/20 DataShed Upload Utility</title>
	<meta http-equiv="content-type" content="text/html;charset=utf-8" />

	<style type="text/css">
	<!--
	/* CSS for this setup template */
	/* These styles do not effect the 20/20 DataShed contents */

	body {
		margin:0px;
		padding:0px;
		font-family:Arial,serif;
		font-size:11px;
		background-color:#ebe9e3;
		}

	a,
	a:link,
	a:active,
	a:visited {
		color:#183c63;
		}

	a:hover {
		color:#a0a9bc;
		}

	#datashedHeader {
		width:100%;
		margin:0px;
		margin-top:40px;
		margin-bottom:0px;
		padding:0px;
		padding-top:13px;
		border-style:solid;
		border-color:#a0a9bc;
		border-width:0px;
		border-top-width:1px;
		border-bottom-width:1px;
		font-family:Arial,serif;
		font-size:16px;
		font-weight:bold;
		color:#a0a9bc;
		background-color:#183c63;
		/*filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#FFa0a9bc', EndColorStr='#FF183c63')*/
		}

	#datashedHeader h1 {
		margin:0px;
		font-size:16px;
		padding-left:20px;
		}

	#datashedLeft {
		text-align:center;
		position:absolute;
		top:105px;
		left:10px;
		width:172px;
		}

	#datashedLeft img {
		margin:0px;
		padding:10px;
		margin-bottom:22px;
		background-color:#ffffff;
		border-style:solid;
		border-width:1px;
		border-color:#a0a9bc;
		}

	.datashedInfo {
		width:100%;
		text-align:left;
		border-style:dashed;
		border-color:#a0a9bc;
		border-width:1px;
		padding:3px;
		margin:0px;
		margin-bottom:10px;
		color:#183c63;
		font-size:10px;
		font-family:Arial,serif;
		background-color:#ded4b4;
		}

	.centered {
		text-align:center;
		}

	#datashedContent {
		margin:0px;
		margin-left:200px;
		padding:10px;
		font-size:12px;
		font-family:Verdana,serif;
		border-style:solid;
		border-color:#a0a9bc;
		border-width:0px;
		border-left-width:1px;
		border-bottom-width:1px;
		background-color:#ffffff;
		}

	h2 {
		font-family:Verdana,Arial,serif;
		font-size:20px;
		color:#a0a9bc;
		}

	h3 {
		font-family:Verdana,Arial,serif;
		font-size:16px;
		color:#888888;
		}

	h4 {
		font-family:Verdana,Arial,serif;
		font-size:13px;
		color:#888888;
		}

	dd {
		margin-bottom:12px;
		}

	li {
		margin-bottom:12px;
		line-height:18px;
		}

	.failed {
		color:#ff0000;
		font-weight:bold;
		}

	.passed {
		color:#009900;
		font-weight:bold;
		}


	.divlink {
		color:#183c63;
		font-style:italic;
		text-decoration:none;
		}

	.divlinkhover {
		color:#a0a9bc;
		font-style:italic;
		text-decoration:underline;
		}

	dl.objectreference {
		margin:20px;
		border-style:solid;
		border-width:1px;
		border-color:#a0a9bc;
		}

	dl.objectreference dt {
		line-height:14px;
		font-size:11px;
		padding-left:3px;
		padding-top:8px;
		clear:both;
		font-weight:bold;
		background-color:#ebe9e3;
		zoom:1;
		/*filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#FFffffff', EndColorStr='#FFebe9e3')*/
		filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr='#FFded4b4', EndColorStr='#FFffffff')
		}

	dl.objectreference dd {
		margin-left:0px;
		margin-bottom:22px;
		}

/*	dl.objectreference dl dt {
		float:left;
		width:180px;
		margin-left:0px;
		font-weight:normal;
		}

	dl.objectreference dl span {
		}

	dl.objectreference dl dd {
		margin-left:181px;
		margin-bottom:22px;
		}
*/
	noscript {
		color:#a0a9bc;
		display:block;
		}

	acronym {
		border-style:dashed;
		border-width:0px;
		border-bottom-width:1px;
		}

	form input {
		display:block;
		}
	//-->
	</style>
</head>
<body>

<div id="datashedHeader"><h1>20/20 Applications Example Documentation - 20/20 DataShed Upload Utility</h1></div>

	
	<div id="datashedLeft">

	<a href="http://www.2020applications.com/listings.asp?itemID=6"><img src="http://www.2020applications.com/images/2020DataShed_logo.gif" alt="20/20 DataShed" /></a>
	<a href="http://www.2020applications.com/listings.asp?itemID=2"><img src="http://www.2020applications.com/images/2020realtor_logo.gif" alt="20/20 Realtor" /></a>
	<a href="http://www.2020applications.com/listings.asp?itemID=4"><img src="http://www.2020applications.com/images/2020autogallery_logo.gif" alt="20/20 Auto Gallery" /></a>

	<a href="http://www.hotscripts.com/Detailed/52226.html?RID=6219"><img src="http://images.hotscripts.com/dynamic/rating.gif?52226" alt="Rated at Hotscripts.com" /></a>

		<div class="datashedInfo centered">
		Copyright &#169; 2005 - <%=Year(Now())%><br />
		by <a href="http://www.2020applications.com/">20/20 Applications</a>
		</div>

	</div>

	<div id="datashedContent">

	<h2>2020_Upload_Example.asp</h2>

	<h3>Here is an example upload form...try it out.</h3>

		<form action="2020_Upload_Example.asp?action=upload" enctype="multipart/form-data" method="post">

		<input type="file" name="strUploadImageFileName1" />
		<input type="submit" name="submitbutton" value="Upload" />

		</form>

	<p>Files will be uploaded to <%=strExamplePath%></p>
	<p>Here is a list of files in <%=strExamplePath%>:</p>

		<%
		dim objFSO
		set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		dim objFolder
		set objFolder = objFSO.GetFolder(strExamplePath)
		dim collFiles
		set collFiles = objFolder.Files

		dim objFile
		dim strFileName
			%>
			<ul>
			<% FOR EACH objFile IN collFiles %>
			<li><a href="<%=objFile.Name%>"><%=objFile.Name%></a></li>
			<% NEXT %>
			</ul>
			<%
		set collFiles = nothing		
		set objFolder = nothing
		set objFSO = nothing
		%>

	<p>You'll know that the upload worked if the list above changes or if you review the list of files in your <acronym title="File Transfer Protocol">FTP</acronym> window.</p>


	<h3>How to Implement this Utility</h3>

		<ol>
		<li>First, you can include a proper &lt;form&gt; in any web page.
			<ul>
			<li>The example &lt;form&gt; above looks like this...you can copy-n-paste this code into any web page.
				<blockquote>
				<code>
				&lt;form action=&quot;<span style="color:blue;">whateveryouwant.asp</span>&quot; enctype=&quot;multipart/form-data&quot; method=&quot;post&quot;&gt;<br />
				&lt;input type=&quot;file&quot; name=&quot;strUploadImageFileName1&quot; /&gt;<br />
				&lt;input type=&quot;submit&quot; name=&quot;submitbutton&quot; value=&quot;Upload&quot; /&gt;<br />
				&lt;/form&gt;
				</code>
				</blockquote>
			</li>
			<li>You can edit the <span style="color:blue;">blue text</span> above so that the form will post the data to an <acronym title="Active Server Pages">ASP</acronym> script for processing -- you'll create that file in the next step.</li>
			<li>The &quot;name&quot; attributes of the &lt;input&gt; elements can be modified as you please -- it really doesn't matter.  In this example, the names are &quot;strUploadImageFileName1&quot; and &quot;submitbutton&quot; but you can rename as you please.</li>
			</ul>
		</li>
		<li>Second, you can produce an <acronym title="Active Server Pages">ASP</acronym> script to process the uploads.
			<ul>
			<li>Create a new file and call it &quot;whateveryouwant.asp&quot;.</li>
			<li>Then copy-n-paste this code into that file:
				<blockquote>
				&lt;%@ Language=&quot;VBScript&quot; %&gt;<br />
				&lt;!--#include file=&quot;_clsUpload.asp&quot;--&gt;<br />
				&lt;%<br />
				dim objUpload<br />
				set objUpload = New clsUpload<br />
				&#160;&#160;&#160;&#160;objUpload.Upload &quot;<span style="color:blue;">YOUR FOLDER PATH HERE</span>&quot;<br />
				<br />
				dim item<br />
				dim objItem<br />
				<br />
				&#160;&#160;&#160;&#160;FOR EACH item IN objUpload.Files<br />
				&#160;&#160;&#160;&#160;set objItem = objUpload.Files(item)<br />
				&#160;&#160;&#160;&#160;&#160;&#160;&#160;&#160;objItem.Save<br />
				&#160;&#160;&#160;&#160;set objItem = nothing<br />
				&#160;&#160;&#160;&#160;NEXT<br />
				<br />
				set objUpload = nothing<br />
				<br />
				response.redirect &quot;<span style="color:blue;">SOME OTHER WEB PAGE.html</span>&quot;<br />
				%&gt;
				</blockquote>
			That's it...that's all you need.  However, you'll learn later that there are many other features available and you may choose to rewrite this code as necessary.</li>
			<li>You should alter the blue text:
				<ul>
				<li>The &quot;objUpload.Upload&quot; method expects the full path of a folder on your web server (example: <tt>C:\Inetpub\wwwroot\your_folder\</tt>).  The folder must have permission to read and write files.</li>
				<li>And you can use the &quot;response.redirect&quot; method to return or redirect the user to a different web page when the upload is complete.</li>
				</ul>
			</li>
			</ul>
		</li>
		</ol>

	<h3>Complete Object Reference: clsUpload</h3>

	<p>The clsUpload object contains public functions and methods which enable you to upload and save files to a web server using pure <acronym title="Visual Basic">VB</acronym>Script.</p>

	<h4>Properties</h4>

		<dl class="objectreference">

		<dt>Files</dt>
		<dd>Read only</dd>
		<dd>Data Type: Dictionary Collection</dd>
		<dd>Syntax: <var>object</var>.Files</dd>
		<dd>Returns a Files collection consisting of all the File objects parsed from the <acronym title="HyperText Transfer Protocol">HTTP</acronym> header.</dd>
		<dd>More information about manipulating Dictionary objects can be viewed at <a href="http://www.devguru.com/technologies/vbscript/13992.asp">www.devguru.com</a>.</dd>

		<dt>Form</dt>
		<dd>Read only</dd>
		<dd>Data Type: Dictionary Collection</dd>
		<dd>Syntax: <var>object</var>.Form</dd>
		<dd>Returns a Forms collection consisting of all the Form objects parsed from the <acronym title="HyperText Transfer Protocol">HTTP</acronym> header.</dd>
		<dd>This collection contains all the <acronym title="HyperText Markup Language">HTML</acronym> form elements which are not &quot;Files&quot;</dd>
		<dd>This collection should be used instead of the typical &quot;request.form&quot; method to access data from the other form elements (if you decide to combine &quot;file&quot; upload elements with other form element types, like checkboxes, text boxes, drop-down lists, textareas).  Why? Because the form data is posted using &quot;multipart/form-data&quot; encoding (and needs to be retrieved like binary data instead of text data) and therefore the typical &quot;request.form&quot; method isn't effective.</dd>
		<dd>More information about manipulating Dictionary objects can be viewed at <a href="http://www.devguru.com/technologies/vbscript/13992.asp">www.devguru.com</a>.</dd>

		<dt>RestrictFileExtensions</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: Boolean (True or False).</dd>
		<dd>Syntax: <var>object</var>.RestrictFileExtensions = [ True | False ]</dd>
		<dd>This property sets or returns a boolean value.</dd>
		<dd>True signifies that the object should compare the uploaded files to the &quot;SafeFileExtensions&quot; property (and decline the upload if the file extension is not considered safe).</dd>
		<dd>False signifies that no restrictions should be imposed on file extensions.</dd>

		<dt>SafeFileExtensions</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: An emtpy string or a pipe-delimited list of file extensions (in lower-case).  Note: the &quot;pipe&quot; is the vertical-bar on your keyboard: |.</dd>
		<dd>Syntax: <var>object</var>.SafeFileExtensions = &quot;txt|html|gif|jpg&quot;</dd>
		<dd>This property sets or returns either an empty string or a pipe-delimited list of file extensions (in lower-case) that are considered &quot;safe&quot; to upload if the &quot;RestrictFileExtentions&quot; property is set to True.</dd>
		<dd>If the &quot;RestrictFileExtentions&quot; property is set to False, then the &quot;SafeFileExtensions&quot; property will have no effect (it will be ignored).</dd>

		<dt>TotalBytes</dt>
		<dd>Read only</dd>
		<dd>Data Type: Long</dd>
		<dd>Syntax: <var>lngTotalBytes</var> = <var>object</var>.TotalBytes</dd>
		<dd>Returns a long (number) value representing the total size of the <acronym title="HyperText Transfer Protocol">HTTP</acronym> header.</dd>
		<dd>Defaults to zero (0).</dd>
		</dl>

	<h4>Methods</h4>

		<dl class="objectreference">
		<dt>Upload</dt>
		<dd>Syntax: <var>object</var>.Upload <var>TargetPath</var></dd>
		<dd>Returns a boolean value: True indicates that the upload was successful; False indicates that the upload was not successful.  If False, then the <acronym title="Active Server Pages">ASP</acronym> error object, &quot;Err&quot;, can be used to retrieve the details of the error.</dd>
		</dl>

	<h4>Events</h4>

		<dl class="objectreference">
		<dt>Initialize</dt>
		<dd>Syntax:
			<blockquote>
			<code>
			dim <var>objUpload</var><br />
			set <var>objUpload</var> = New clsUpload
			</code>
			</blockquote>
		</dd>
		<dd>When the object is instantiated, this happens:
			<ul>
			<li>The &quot;Files&quot; collection is created (with a count of zero (0)).</li>
			<li>The &quot;Form&quot; collection is created (with a count of zero (0)).</li>
			<li>RestrictFileExtension is set to False.</li>
			<li>SafeFileExtensions is set to &quot;&quot; (an empty string).</li>
			<li>TotalBytes is set to zero (0).</li>
			<li>The &quot;ScriptTimeout&quot; property of the <acronym title="Active Server Pages">ASP</acronym> &quot;Server&quot; object is set to 600 seconds (10 minutes).  This is to allow sufficient time for the files to upload from the user's web browser -- in most cases this process takes mere seconds.</li>
			<li>The &quot;CodePage&quot; property of the <acronym title="Active Server Pages">ASP</acronym> &quot;Response&quot; object is set to 1252 if the &quot;Response&quot; object supports this property.  This is done to ensure that binary data is converted properly to strings (and vice-versa) -- some mult-lingual or non-English web servers suffer errors in such conversions if the &quot;CodePage&quot; is non-English.  (Hmm...non-<i>Latin</i> is perhaps more accurate).  This feature will fail in <acronym title="Internet Information Server">IIS</acronym> version 5.0 or older -- but it won't cause a problem unless that server is also configured to use non-English language settings.  Upgrading the web server with a recent version of <acronym title="Internet Information Server">IIS</acronym> resolves this issue.</li>
			</ul>
		</dd>
		<dt>Terminate</dt>
		<dd>Syntax:
			<blockquote>
			<code>
			set <var>objUpload</var> = nothing
			</code>
			</blockquote>
		</dd>
		<dd>When the object is terminated, this happens:
			<ul>
			<li>The &quot;Files&quot; collection is terminated.</li>
			<li>The &quot;Form&quot; collection is terminated.</li>
			<li>The &quot;CodePage&quot; property of the <acronym title="Active Server Pages">ASP</acronym> &quot;Response&quot; object is set to 65001 (20/20 DataShed's default setting) if the &quot;Response&quot; object supports this property.  This is done primarily as a housekeeping chore but won't make much difference in most cases -- because the &quot;CodePage&quot; property only survives until the end of the script.  Other scripts on the server will default to the system's &quot;CodePage&quot; or the &quot;CodePage&quot; may be explicitly set in your other <acronym title="Active Server Pages">ASP</acronym> scripts.</li>
			</ul>
		</dd>
		</dl>




	<h3>Complete Object Reference: File</h3>

	<p>Multiple files can be uploaded simultaneously.  A &quot;clsUploadedFile&quot; object is created by clsUpload for each file-blob that is uploaded -- each clsUploadedFile object corresponds to a &quot;File&quot; item in the &quot;Files&quot; collection is accessible through the clsUpload.Files dictionary collection.</p>
	<p>In other words, once your files are uploaded, you can manipulate them with the following properties and methods.</p>
	<p>Note: if &quot;RestrictFileExtensions&quot; is True, then only &quot;Safe&quot; files will be available in the &quot;File&quot; collection.  This means that if you upload a file, then if the <var>UploadObject</var>.Files.Count property is zero (0), then you can conclude that the file extension was rejected.</p>
	<h4>Properties</h4>

		<dl class="objectreference">

		<dt>ContentType</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: String</dd>
		<dd>Syntax: <var>FileObject</var>.ContentType</dd>
		<dd>Sets or returns a string representing the content type of the file blob.  Example: &quot;image/gif&quot; or &quot;text/plain&quot;</dd>
		<dd>This property is set automatically when the clsUpload.Upload method is called but can be altered prior to saving the file blob to the web server's hard drive.</dd>
		<dd>It is usually impractical to alter this property and should be treated as Read-only.</dd>

		<dt>Data</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: Binary</dd>
		<dd>Syntax: <var>FileObject</var>.Data</dd>
		<dd>Sets or returns the binary data contained in each file blob.</dd>
		<dd>This property is set automatically when the clsUpload.Upload method is called but can be altered prior to saving the file blob to the web server's hard drive.</dd>
		<dd>It is usually impractical to alter this property and should be treated as Read-only.</dd>

		<dt>FileExists</dt>
		<dd>Read only</dd>
		<dd>Data Type: Boolean (True or False).</dd>
		<dd>Syntax: <var>FileObject</var>.FileExists</dd>
		<dd>Returns a boolean value.  True signifies that a file of the same filename already exists on the hard drive at this &quot;UploadPath&quot; -- if the script proceeds with the .Save() method, then the existing file will be overwritten.  False signifies that no files exists at this &quot;UploadPath&quot; with this filename.</dd>

		<dt>FileName</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: String</dd>
		<dd>Syntax: <var>FileObject</var>.FileName = <var>strFileName</var></dd>
		<dd>Sets or returns the file name for each file blob.</dd>
		<dd>This property is set automatically when the clsUpload.Upload method is called but can be altered prior to saving the file blob to the web server's hard drive.</dd>

		<dt>FolderExists</dt>
		<dd>Read only</dd>
		<dd>Data Type: Boolean (True or False).</dd>
		<dd>Syntax: <var>FileObject</var>.FolderExists</dd>
		<dd>Returns a boolean value.  True signifies that that &quot;UploadPath&quot; exists on the hard drive.  False signifies that the &quot;UploadPath&quot; does not exist -- if the script proceeds with the .Save() method, then the function will fail.  The folder must exist prior to calling the .Save() method.</dd>

		<dt>Size</dt>
		<dd>Read only</dd>
		<dd>Data Type: Long</dd>
		<dd>Syntax: <var>FileObject</var>.Size</dd>
		<dd>Returns a long (number) representing the total length of the file blob (i.e. the file size).</dd>

		<dt>UploadPath</dt>
		<dd>Read/Write</dd>
		<dd>Data Type: String</dd>
		<dd>Syntax: <var>FileObject</var>.UploadPath = <var>TargetPath</var></dd>
		<dd>Sets or returns the target path for saving each file blob to disk.</dd>
		<dd>This property is set automatically when the clsUpload.Upload method is called but can be altered prior to saving the file blob to the web server's hard drive.</dd>

		</dl>


	<h4>Methods</h4>

		<dl class="objectreference">

		<dt>BuildPath</dt>
		<dd>Syntax: <var>FileObject</var>.BuildPath</dd>
		<dd>Returns a string representing a path.</dd>
		<dd>This method is used internally by the clsUploadedFile object and mimics the behaviour of the FileSystemObject.BuildPath(<var>UploadPath</var>,<var>FileName</var>) method.</dd>
		<dd>To alter this value you must first change either the UploadPath or FileName properties of this File object, then call this method to build a new path string using those new values.</dd>

		<dt>Save</dt>
		<dd>Syntax: <var>FileObject</var>.Save</dd>
		<dd>Saves the file blob to disk at the UploadPath and with the current FileName.  Returns a boolean indicating whether the process was successful.</dd>
		<dd>This method should be called after appropriate measures have been taken to ensure that this file blob carries an appropriate FileName and that the folder exists.</dd>
		<dd>This method must be called for each File object.</dd>
		<dd>No error handling is built-in.  So, you will immediately see an error message if this process fails so that you can take appropriate action (fix permissions, check that the folder exists).  Therefore it's best if you use other methods and properties to ensure that the .Save() method will succeed.</dd>

		</dl>



	<h4>Events</h4>

		<dl class="objectreference">
		<dt>Initialize</dt>
		<dd>New clsUploadedFile objects are created automatically when the clsUpload.Upload method is called -- they do not have to be created explicitly.</dd>
		<dd>When the object is instantiated, the UploadPath property is set to the current folder.</dd>
		<dt>Terminate</dt>
		<dd>The object is destroyed.</dd>
		</dl>


	<h3>A Better Code Example</h3>

	<p><a href="2020_Upload_AdvancedForm.html">Click here to visit a more advanced &lt;form&gt; example.</a></p>
	<p>This advanced form example utilizes the code found in &quot;_2020AdvancedUploadProcessor.asp&quot;.</p>
	<p>You can open these files in a text editor like Notepad to view the source -- this example demonstrates all the properties and methods available in the object.</p>


	</div>

</body>
</html>