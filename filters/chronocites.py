#!/usr/bin/env python

#  "chronocites" Chronological Citation Filter Script
#  Copyright (c) 2014, Adam Rehn
#
#  Filter to chronologically sort the citations in a UTF-8 (X)HTML file.
#
#  Supports citations in the following format:
#    (Author names, Year; Author names, Year)
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

import re
import sys

# Retrieves the contents of a file
def getFileContents(filename):
	f = open(filename, "r")
	data = f.read().decode("utf-8")
	f.close()
	return data

# Writes the contents of a file
def putFileContents(filename, data):
	f = open(filename, "w")
	f.write(data.encode("utf-8"))
	f.close()

def removeAmpersandEntities(s):
	return s.replace('&amp;', '&&')

def restoreAmpersandEntities(s):
	return s.replace('&&', '&amp;')


if (len(sys.argv) > 2):
	
	data = getFileContents(sys.argv[1]);
	
	# Find each of the citations with multiple entries
	def sortCitationChronologically(match):
		
		citationEntries = removeAmpersandEntities(match.group(1)).split(';');
		if len(citationEntries) < 2:
			return restoreAmpersandEntities(match.group(0))
		
		# Iterate over each of the entries and extract the year identifiers, grouping the entries by year
		years = {}
		for entry in citationEntries:
			components = entry.strip().rsplit(',', 1)
			if len(components) == 2:
				year = components[1].strip()
				if year in years:
					years[year].append(components[0])
				else:
					years[year] = [ components[0] ]
		
		# Build the chronological list of citation entries
		chronoEntries = []
		yearsSorted = list(years.keys())
		yearsSorted.sort()
		for year in yearsSorted:
			years[year].sort()
			for entry in years[year]:
				chronoEntries.append(entry + u", " + year)
		
		# Replace the original citation with the chronologically sorted one
		sortedCitation = '(' + restoreAmpersandEntities('; '.join(chronoEntries)) + ')'
		return sortedCitation
		
	
	processedData = re.sub("\((.+?)\)", sortCitationChronologically, data, flags=re.DOTALL)
	putFileContents(sys.argv[2], processedData)

else:
	print "Usage syntax:\n" + sys.argv[0] + " <INFILE> <OUTFILE>\n"
