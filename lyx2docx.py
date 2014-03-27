#!/usr/bin/env python

#  "lyx2docx" LyX to DOCX Conversion Script
#  Copyright (c) 2014, Adam Rehn
# 
#  Script to convert LyX files to DOCX format using tex4ht and pandoc.
#
#  Requires the following programs in the PATH:
#
#    lyx
#    latex
#    bibtex
#    htlatex
#    iconv
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
from subprocess import call
import re, os, sys, argparse

# Retrieves the contents of a file
def file_get_contents(filename):
	f = open(filename, "r")
	data = f.read()
	f.close()
	return data

# Writes the contents of a file
def file_put_contents(filename, data):
	f = open(filename, "w")
	f.write(data.encode("utf-8"))
	f.close()

# Removes all files with the specified filename and any of the specified extensions
def remove_extension_versions(filename, extensions):
	for extension in extensions:
		if os.path.exists(filename + "." + extension):
			os.remove(filename + "." + extension)

# Replaces the specified subpattern of a regular expression match
def replace_subpattern(string, pattern, groupNo, replacement):
	p = re.compile(pattern, re.DOTALL)
	matches = p.search(string)
	if matches != None:
		return string.replace(matches.group(groupNo), replacement)
	else:
		return string

# Removes any \usepackage commands from a LaTex source for the specified package
def remove_package(string, packageName):
	return replace_subpattern(string, "\\\\usepackage(\\[[^\\]]+?\\]){0,1}\\s*?\\{" + packageName + "\\}", 0, "")

# Removes all tags with the specified name, and whose attribute matches the specified value, if supplied
def remove_tags(dom, tagName, attribName = None, attribValue = None):
	for node in dom.getElementsByTagName(tagName):
		if attribName == None or node.getAttribute(attribName) == attribValue:
			node.parentNode.removeChild(node)

# Reads XML data from a file, sanitises it, and parses it using minidom
def parse_xml_file(xmlFile):
	
	# Read the XML data from the specified file
	xmlData = file_get_contents(xmlFile)
	
	# Convert unsupported HTML entities into XML-compatible character references
	# (See <http://www.dwheeler.com/essays/quotes-in-html.html>)
	xmlData = xmlData.replace("&ldquo;", "&#8220;")
	xmlData = xmlData.replace("&rdquo;", "&#8221;")
	xmlData = xmlData.replace("&lsquo;", "&#8216;")
	xmlData = xmlData.replace("&rsquo;", "&#8217;")
	
	# Parse the XML data
	dom = parseString(xmlData)
	return dom


# Ensure the correct number of command-line arguments were supplied
if len(sys.argv) < 2:
	print("Usage syntax:\n\n" + sys.argv[0].replace(".py", "") + " LYXFILE.LYX")
	sys.exit()

# Process command-line arguments
lyxFile         = sys.argv[1]
texFile         = lyxFile.replace(".lyx", ".tex")
texFileNoSpaces = texFile.replace(" ", "_")
xhtmlFile       = texFileNoSpaces.replace(".tex", ".html")
xhtmlFileUTF8   = xhtmlFile.replace(".html", ".utf8.html")
docxFile        = lyxFile.replace(".lyx", ".docx")

# Export the LaTex file using LyX
call(["lyx", "--export", "latex", lyxFile])

# If the LaTeX filename has spaces in it, rename it
if texFile != texFileNoSpaces:
	os.rename(texFile, texFileNoSpaces)
	texFile = texFileNoSpaces

# Read the generated LaTex file
texData = file_get_contents(texFile)

# Remove use the hyperref package and packages that rely on it
texData = remove_package(texData, "hyperref")
texData = remove_package(texData, "breakurl")

# Strip the code inserted by LyX containing the \phantomsection command
texData = replace_subpattern(texData, "\\\\bibliographystyle\\{.+?\\}(.+?)\\\\bibliography\\{", 1, "")

# Write the modified LaTex back to the generated LaTex file
file_put_contents(texFile, texData)

# Run latex on the LaTeX file
call(["latex", texFile])

# Run bibtex on the LaTeX file
call(["bibtex", texFile.replace(".tex", ".aux")])

# Use tex4ht to generate the XHTML file
call(["htlatex", texFile, "xhtml, charset=utf-8"])

# Use iconv to transform the XHTML file into UTF-8 (input and output filenames cannot match, older iconv versions don't support "-o")
utf8out = open(xhtmlFileUTF8, "w")
call(["iconv", "-t", "utf-8", xhtmlFile], stdout=utf8out)
utf8out.close()

# Delete the original (non-UTF-8) XHTML file and replace it with the UTF-8 one
os.remove(xhtmlFile)
os.rename(xhtmlFileUTF8, xhtmlFile)

# Parse the generated XHTML file
dom = parse_xml_file(xhtmlFile)

# Strip the <title> tag to prevent pandoc prepending the title to the output
remove_tags(dom, "title")

# Strip any <meta> tag that supplies a date value, to prevent pandoc prepending the date to the output
remove_tags(dom, "meta", "name", "date")

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

# Strip all <span class="bibsp"> tags
remove_tags(dom, "span", "class", "bibsp")

# Write the modified XML back to the XHTML file
file_put_contents(xhtmlFile, dom.toxml())

# Generate the DOCX file using pandoc
call(["pandoc", "-o", docxFile, xhtmlFile])

# Remove any image files generated by tex4ht
for node in dom.getElementsByTagName("img"):
	if node.getAttribute("src") != None:
		os.remove(node.getAttribute("src"))

# Cleanup intermediate files
remove_extension_versions(texFileNoSpaces.replace(".tex", ""), ["4ct","4tc","aux","bbl","blg","css","dvi","html","idv","lg","log","tex","tmp","xref"])
