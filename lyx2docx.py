#!/usr/bin/env python

#  "lyx2docx" LyX to DOCX Conversion Script
#  Copyright (c) 2014, Adam Rehn
# 
#  Script to convert LyX files to DOCX format using tex4ht and pandoc.
#
#  Requires the following programs in the PATH:
#
#    lyx      (unless using the --latex option to process a LaTeX input file)
#    latex
#    bibtex
#    htlatex  (the MikTeX tex4ht-bin package is insufficient, tex4ht-base is required)
#    iconv    (only used under Windows, where tex4ht UTF-8 output support is broken)
#    pandoc
# 
#  ---
#
#  Permission is hereby granted, free of charge, to any person obtaining a copy
#  of this software and associated documentation files (the "Software"), to deal
#  in the Software without restriction, including without limitation the rights
#  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
#  copies of the Software, and to permit persons to whom the Software is
#  furnished to do so, subject to the following conditions:
#
#  The above copyright notice and this permission notice shall be included in
#  all copies or substantial portions of the Software.
#
#  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
#  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
#  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
#  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
#  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
#  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
#  THE SOFTWARE.

from __future__ import print_function
from xml.dom.minidom import parse, parseString
from shutil import copyfile
import re, os, sys, subprocess, platform, argparse

# Replace the last occurence of a substring
# Adapted from the accepted answer at: <http://stackoverflow.com/questions/2556108/how-to-replace-the-last-occurence-of-an-expression-in-a-string>
def rreplace(s, old, new):
	li = s.rsplit(old, 1)
	return new.join(li)

# Retrieves the contents of a file
def getFileContents(filename):
	f = open(filename, "r")
	data = f.read()
	f.close()
	return data

# Writes the contents of a file
def putFileContents(filename, data):
	f = open(filename, "w")
	f.write(data.encode("utf-8"))
	f.close()

# Removes all files with the specified filename and any of the specified extensions
def removeExtensionVersions(filename, extensions):
	for extension in extensions:
		if os.path.exists(filename + "." + extension):
			os.remove(filename + "." + extension)

# Replaces the file extension component of a path with the specified extension
def replaceExtension(path, newExtension):
	fileName, fileExtension = os.path.splitext(path)
	return rreplace(path, fileExtension, newExtension)

# Replaces the specified subpattern of a regular expression match
def replaceSubpattern(string, pattern, groupNo, replacement):
	p = re.compile(pattern, re.DOTALL)
	matches = p.search(string)
	if matches != None:
		return string.replace(matches.group(groupNo), replacement)
	else:
		return string

# Removes any \usepackage commands from a LaTex source for the specified package
def removePackage(string, packageName):
	return replaceSubpattern(string, "\\\\usepackage(\\[[^\\]]+?\\]){0,1}\\s*?\\{" + packageName + "\\}", 0, "")

# Removes all tags with the specified name, and whose attribute matches the specified value, if supplied
def removeTags(dom, tagName, attribName = None, attribValue = None):
	for node in dom.getElementsByTagName(tagName):
		if attribName == None or node.getAttribute(attribName) == attribValue:
			node.parentNode.removeChild(node)

# Reads XML data from a file, sanitises it, and parses it using minidom
def parseXmlFile(xmlFile):
	
	# Read the XML data from the specified file
	xmlData = getFileContents(xmlFile)
	
	# Convert unsupported HTML entities into XML-compatible character references
	# (See <http://www.dwheeler.com/essays/quotes-in-html.html>)
	xmlData = xmlData.replace("&ldquo;", "&#8220;")
	xmlData = xmlData.replace("&rdquo;", "&#8221;")
	xmlData = xmlData.replace("&lsquo;", "&#8216;")
	xmlData = xmlData.replace("&rsquo;", "&#8217;")
	
	# Parse the XML data
	dom = parseString(xmlData)
	return dom

# Executes a command and, if it fails, prints the return code, stdout, and stderr
def executeCommand(commandArgs, quitOnError = True, redirectStdoutHere = None):
	
	# Redirect the command's stdout to a file if requested
	stdout = subprocess.PIPE
	if redirectStdoutHere != None:
		stdout = open(redirectStdoutHere, "w")
	
	# Execute the command and capture its stdout and stderr
	proc = subprocess.Popen(commandArgs, stdout=stdout, stderr=subprocess.PIPE)
	(stdoutdata, stderrdata) = proc.communicate(None)
	
	# If we were redirecting the command's stdout, close the output file
	if stdout != subprocess.PIPE:
		stdout.close()
	
	# If the command failed, report the error
	if proc.returncode != 0:
		print("Command", commandArgs, "failed with Exit Code", proc.returncode)
		print("Stdout was: \"", stdoutdata, "\"", sep="")
		print("Stderr was: \"", stderrdata, "\"", sep="")
		
		# If requested, terminate execution
		if quitOnError == True:
			sys.exit(1)

# Close stdin to ensure that subprocesses can't hang waiting for input
stdinFileno = sys.stdin.fileno()
sys.stdin.close()
os.close(stdinFileno)

# Adding items to the PATH under Darwin can be a hassle compared to other platforms, so we support auto-adding LyX
if platform.system() == "Darwin":
	os.environ["PATH"] += ":/Applications/LyX.app/Contents/MacOS"

# Parse command-line arguments
parser = argparse.ArgumentParser()
parser.add_argument("--keep-files", action="store_true", help="Don't delete intermediate files")
parser.add_argument("-t", "--template", default="", help="Template DOCX file for pandoc (maps to pandoc --reference-docx argument)")
parser.add_argument("-dir", default="", help="Input file directory (prepended to input file)")
parser.add_argument("-filter", default=None, help="Filter command to process intermediate XHTML file")
parser.add_argument("--latex", action="store_true", help="Treat input file as a LaTeX file (don't process it using LyX)")
parser.add_argument("infile", help="Input file")
args = parser.parse_args()

# If an absolute path to the input file was supplied, change into the file's directory
infile        = args.dir + args.infile
infileDir     = os.path.dirname(infile)
origWorkingDir = os.getcwd()
if infileDir != "" and infileDir != ".":
	os.chdir(infileDir)

# Determine if we are treating the input file as a LyX file or a LaTeX file
infile = os.path.basename(infile)
if args.latex == True:
	
	# Create a copy of the LaTeX file and work with that (allows cleanup to clobber the copy instead of the original)
	texFile = "____" + infile
	copyfile(infile, texFile)
	
else:
	
	# Export a LaTex file using LyX
	texFile = replaceExtension(infile, ".tex")
	executeCommand(["lyx", "--export", "latex", infile])

# Determine the filenames of the generated files
texFileNoSpaces = texFile.replace(" ", "_")
xhtmlFile       = replaceExtension(texFileNoSpaces, ".html")
docxFile        = replaceExtension(infile, ".docx")

# If the LaTeX filename has spaces in it, rename it
if texFile != texFileNoSpaces:
	
	# If the renamed LaTeX file exists (from a previous run), overwrite it
	if os.path.exists(texFileNoSpaces):
		os.remove(texFileNoSpaces)
	
	os.rename(texFile, texFileNoSpaces)
	texFile = texFileNoSpaces

# Read the generated LaTex file
texData = getFileContents(texFile)

# Remove use the hyperref package and packages that rely on it
texData = removePackage(texData, "hyperref")
texData = removePackage(texData, "breakurl")

# Strip the code inserted by LyX containing the \phantomsection command
texData = replaceSubpattern(texData, "\\\\bibliographystyle\\{.+?\\}(.+?)\\\\bibliography\\{", 1, "")

# Write the modified LaTex back to the generated LaTex file
putFileContents(texFile, texData)

# Run latex and bibtex on the LaTeX file
executeCommand(["latex", texFile])
executeCommand(["bibtex", replaceExtension(texFile, ".aux")])

# Determine if we are running under Windows
if platform.system() == "Windows":
	
	# Use tex4ht to generate the XHTML file, avoiding the broken UTF-8 output support under Windows
	executeCommand(["htlatex", texFile, "xhtml, charset=utf-8"])
	
	# Use iconv to transform the XHTML file into UTF-8 (input and output filenames cannot match, older iconv versions don't support "-o")
	xhtmlFileUTF8 = replaceExtension(xhtmlFile, ".utf8.html")
	executeCommand(["iconv", "-t", "utf-8", xhtmlFile], True, xhtmlFileUTF8)
	
	# Delete the original (non-UTF-8) XHTML file and replace it with the UTF-8 one
	os.remove(xhtmlFile)
	os.rename(xhtmlFileUTF8, xhtmlFile)
	
else:
	
	# On non-Windows platforms, simply use the native UTF-8 output support of tex4ht
	executeCommand(["htlatex", texFile, "xhtml, charset=utf-8", " -cunihtf -utf8"])

# Parse the generated XHTML file
dom = parseXmlFile(xhtmlFile)

# Strip the <title> and <meta name="date"> tags to prevent pandoc prepending the title and date to the output
removeTags(dom, "title")
removeTags(dom, "meta", "name", "date")

# Strip the <hr> tags surrounding figures and tables
removeTags(dom, "hr", "class", "figure")
removeTags(dom, "hr", "class", "endfigure")
removeTags(dom, "hr", "class", "float")
removeTags(dom, "hr", "class", "endfloat")

# Strip all <span class="bibsp"> tags
removeTags(dom, "span", "class", "bibsp")

# Strip all <a> tags with a href value starting with "#"
for node in dom.getElementsByTagName("a"):
	
	# Check that the tag has a href value that starts with "#"
	href = node.getAttribute("href")
	if href != None and href[0:1] == "#":
	
		# If the tag has any child nodes, insert them in its place
		for childNode in node.childNodes:
			node.parentNode.insertBefore(childNode, node)
		
		# Remove the tag
		node.parentNode.removeChild(node)

# Find all span tags whose class is in a format resembling "ptmri8t-" and contains the "i" (italic) specifier, inserting <em> tags
for node in dom.getElementsByTagName("span"):
	if node.getAttribute("class").find("i") != -1 and node.getAttribute("class").find("-") != -1:
		
		# Create an <em> node and move all of the span's child nodes
		emNode = dom.createElement("em")
		while node.firstChild != None:
			emNode.appendChild( node.removeChild(node.firstChild) )
		
		# Append the <em> node as the span's new child node
		node.appendChild(emNode)
		
		# If the last character in the newly created <em> is a space, duplicate it after the <span>
		if emNode.lastChild != None and emNode.lastChild.nodeValue.endswith(" "):
			space = dom.createTextNode(" ")
			node.parentNode.insertBefore(space, node.nextSibling)

# Find all empty div tags and remove them, since some implementations of the XML DOM classes write empty-element tags (<div />)
# for empty div elements, breaking layout in some web browsers
for node in dom.getElementsByTagName("div"):
	if node.hasChildNodes() == False and (node.nodeValue == None or len(node.nodeValue) == 0):
		node.parentNode.removeChild(node)

# Retrieve the modified XML
modifiedXML = dom.toxml()

# Write the modified XML back to the XHTML file
putFileContents(xhtmlFile, modifiedXML)

# If a filter step was specified, execute it now
if args.filter != None:
	executeCommand([args.filter, xhtmlFile, xhtmlFile])

# Generate the DOCX file using pandoc (using a custom template if specified)
pandocCommand = ["pandoc", "-o", docxFile, xhtmlFile]
if args.template != "":
	pandocCommand.extend(["--reference-docx", args.template])
executeCommand(pandocCommand)

# Determine if we are removing the intermediate files
if args.keep_files == False:
	
	# Remove any image files generated by tex4ht
	for node in dom.getElementsByTagName("img"):
		if node.getAttribute("src") != None:
			os.remove(node.getAttribute("src"))

	# Cleanup intermediate files
	removeExtensionVersions(replaceExtension(texFileNoSpaces, ""), ["4ct","4tc","aux","bbl","blg","css","dvi","html","idv","lg","log","tex","tmp","xref"])

# Change back to our original working directory
os.chdir(origWorkingDir)
