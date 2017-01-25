// PageExporterUtility5.0.js
// An InDesign CS JavaScript
// 08 NOV 2007
// Copyright (C) 2007  Scott Zanelli. Lonelytree Software. (www.lonelytreesw.com)
// Coming to you from Quincy, MA, USA

// This program is free software; you can redistribute it and/or
// modify it under the terms of the GNU General Public License
// as published by the Free Software Foundation; either version 2
// of the License, or (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program; if not, write to the Free Software
// Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

var peuINFO = new Array();
peuINFO["csVersion"] = parseInt(app.version);
// Save the old interaction level
if(peuINFO.csVersion == 3) { //CS1
	peuINFO["oldInteractionPref"] = app.userInteractionLevel;
	app.userInteractionLevel = UserInteractionLevels.interactWithAll;
}
else { //CS2+
	peuINFO["oldInteractionPref"] = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

// See if a document is open. If not, exit
if((peuINFO["numDocsToExport"] = app.documents.length) == 0){
	byeBye("Please open a document and try again.",true);
}

// Global Variable Initializations
var VERSION_NAME = "Page Exporter Utility 5.0"
var pseudoSingleton = 0;
var progCurrentPage = 0;// Used for progress bar
var progTotalPages = 0; // Used for progress bar
var commonBatchINFO = getNewTempENTRY(); // Used if "Set For All Batch Jobs" is selected
commonBatchINFO["infoLoaded"] = false;
peuInit(); // Initialize commonly needed info

// Store information needed for batch processing (valid for single also)
var printINFO = new Array();

// Get all needed info by looping for each document being exported
for(currentDoc = 0; currentDoc < peuINFO.numDocsToExport; currentDoc++) {
	// Get name of document and set it as the base name
	// minus the extention if it exists
	var tempENTRY = getNewTempENTRY(); // "Struct" for current document info
	tempENTRY.theDoc = app.documents[currentDoc]
	tempENTRY.singlePage = (tempENTRY.theDoc.documentPreferences.pagesPerDocument==1)?true:false;
	tempENTRY.getOut = true;
	var baseName = (tempENTRY.theDoc.name.split(".ind"))[0];

	// Display the dialog box and loop for correct info
	if((!commonBatchINFO.infoLoaded && peuINFO.batchSameForAll) || (!peuINFO.batchSameForAll) ){// get all info
		do{
			tempENTRY.getOut = true; // For handling input errors

			var mainDialog = createMainDialog(tempENTRY.theDoc.name, currentDoc+1, peuINFO.numDocsToExport);

			// Exit if cancel button chosen
			if(!mainDialog.show() ){
				mainDialog.destroy();
				byeBye("Exporting has canceled by user.",peuINFO.sayCancel);
			}

			if(formatTypeRB.selectedButton == 4 - peuINFO.adjustForNoPS){
				changePrefs();
				tempENTRY.getOut = false;
				continue;
			}

			// Read info from dialog items and keep as defaults and place in tempENTRY
			peuINFO.defaultDir = outputDirs.selectedIndex;
			// The index of the selected directory is the same as the
			// index in the array of directories.
			tempENTRY.outDir = peuINFO.dirListArray[peuINFO.defaultDir];
			// Type of renaming to do
			tempENTRY.nameConvType = namingConvention.selectedButton
			// The base name to add page info to
			tempENTRY.baseName = removeColons(removeSpaces(newBaseName.editContents) );
			// The start [and end] page numbers
			var startEndPgs = startEndPgs.editContents;
			// Wether to do spreads or not
			peuINFO.doSpreadsON = (doSpreads.checkedState)?1:0;
			tempENTRY.doSpreadsON = peuINFO.doSpreadsON;
			// Wether to send entire file as one document
			peuINFO.doOneFile = (oneFile.checkedState)?1:0;
			tempENTRY.doOneFile = peuINFO.doOneFile;
			// Export format type
			tempENTRY.formatType = formatTypeRB.selectedButton + peuINFO.adjustForNoPS;
			
			// Set persistence during warnings
			peuINFO.editStartEndPgs = startEndPgs;
			baseName = tempENTRY.baseName;
			peuINFO.exportDefaultType = tempENTRY.formatType;
			
			// Determine if page replacement token exists when the page token option is used
			if(peuINFO.pageNamePlacement == 2){
				var temp = tempENTRY.baseName.indexOf("<#>");
				if(temp == -1){//Token isn't there
					alert("There is no page item token (<#>) in the base name. Please add one or change the page naming placement preference.");
					tempENTRY.getOut = false;
					pseudoSingleton--; // Allow prefs to be accessed again, but just once
					continue;
				}
				else
					pseudoSingleton++;
			}	
			else // Try to remove any <#>s as a precaution
				tempENTRY.baseName = tempENTRY.baseName.replace(/<#>/g,"");	
						
			// Layer Versioning & Batch options
			if(currentDoc < 1){
				peuINFO.layersON = (doLayers.checkedState)?1:0;
				if(peuINFO.numDocsToExport > 1){
					peuINFO.batchSameForAll = (commonBatch.checkedState)?1:0;
					peuINFO.batchON = (doBatch.checkedState)?1:0;
					if(!peuINFO.batchON)
						peuINFO.numDocsToExport = 1;
				}
			}

			//Check if spreads chosen with 'Add ".L"' option as this isn't supported.
			if(peuINFO.doSpreadsON && tempENTRY.nameConvType == 1){
				alert ("Spreads cannot be used with the 'Add \".L\"' option.\nThis combination is not supported. (1.1)");
				tempENTRY.nameConvType = 0;
				tempENTRY.getOut = false;
				continue;
			}
			else if(peuINFO.doSpreadsON && tempENTRY.nameConvType == 4){
				alert ("Spreads cannot be used with the 'Numeric Override' option.\nThis combination is not supported. (1.2)");
				tempENTRY.nameConvType = 0;
				tempENTRY.getOut = false;
				continue;
			}

			// Check if "Send Entire File At Once" is selected with JPG or EPS
			if(peuINFO.doOneFile && tempENTRY.formatType > 1 ){
				alert ("The 'Send Entire File At Once' option can only be used with PostScript or PDF formats. (1.3)");
				tempENTRY.getOut = false;
				continue;
			}

			// Check if: batch printing and using the "Same for all jobs options" and a page range other than "All" was used
			if(peuINFO.doBatch && peuINFO.batchSameForAll && startEndPgs != "All"){
				alert ("The 'Set For All Batch Jobs' option can only be used with a Page Range of 'All'. Page Range has been reset to 'All'. (1.4)");
				startEndPgs = "All";
			}

			// Create page info, skip if doing entire file as one
			var tempPageCount = 0;
			if(tempENTRY.doOneFile)
				tempPageCount = 1;
			else{
				// Get names of all the pages. Needed when pages are named using sectioning
				tempENTRY = recordPgNames(tempENTRY);
				// Check Page Validity and get Page counts of entered section(s)
				var temp = checkPages(tempENTRY, startEndPgs);
				tempENTRY = temp[0];
				tempPageCount = temp[1];
				temp = null; // Free it up
			}
		} while(!tempENTRY.getOut);

		// Remove dialog from memory
		mainDialog.destroy();

		// Determine if tag will fit correctly
		tempENTRY.useTag = usePgInfoTag(tempENTRY.theDoc.viewPreferences.horizontalMeasurementUnits,tempENTRY.theDoc.documentPreferences.pageWidth);

			// Get the format info for this document
		switch(tempENTRY.formatType){
			case 0://PS
			tempENTRY.psINFO = getPSoptions(tempENTRY.theDoc.name.split(".ind")[0]);
			break;
			case 1://PDF
			tempENTRY.pdfPRESET = getPDFoptions(tempENTRY);
			break;
			case 2: // EPS Formatting
			tempENTRY.epsINFO = getEPSoptions(tempENTRY.theDoc.name.split(".ind")[0]);
			peuINFO.origSpread = app.epsExportPreferences.epsSpreads;// Used to reset to original state when done
			app.epsExportPreferences.epsSpreads = peuINFO.doSpreadsON;
			break;
			case 3: // JPEG Formatting
			tempENTRY.jpegINFO = getJPEGoptions(tempENTRY.theDoc.name.split(".ind")[0]);
			break;
		}

		// If Specific Directory was chosen for the output directory, get it now
		if(peuINFO.defaultDir == 0){
			tempENTRY.outDir = getDirectory("Please select the output directory:",peuINFO.startingDirectory);
			if(tempENTRY.outDir != null)
				tempENTRY.outDir += "/";
			else
				byeBye("Exporting has been canceled by user.", peuINFO.sayCancel);
		}
		// Set the common elements for all batch jobs if it was selected
		if(!commonBatchINFO.infoLoaded && peuINFO.batchSameForAll){
			commonBatchINFO.infoLoaded = true;
			commonBatchINFO.pageNamePlacement = tempENTRY.pageNamePlacement;
			commonBatchINFO.outDir = tempENTRY.outDir;
			commonBatchINFO.nameConvType = tempENTRY.nameConvType
			commonBatchINFO.doSpreadsON = tempENTRY.doSpreadsON;
			commonBatchINFO.doOneFile = tempENTRY.doOneFile;
			commonBatchINFO.formatType = tempENTRY.formatType;
			commonBatchINFO.psINFO = tempENTRY.psINFO;
			commonBatchINFO.pdfPRESET = tempENTRY.pdfPRESET
			commonBatchINFO.epsINFO = tempENTRY.epsINFO
			commonBatchINFO.jpegINFO = tempENTRY.jpegINFO;
		}
	} // End each/first of batch
	else{ // Get the base name for other batch jobs
		do{
			tempENTRY.getOut = true;
			var nameDialog = app.dialogs.add({name:(VERSION_NAME + ": Base Name for \"" + tempENTRY.theDoc.name + "\"" + ((peuINFO.numDocsToExport==1)?"":" (" + (currentDoc+1) + " of " + peuINFO.numDocsToExport + " documents)") ), canCancel:true} );
			with (nameDialog){
				with (dialogColumns.add() ){
					with(dialogRows.add() ){
					staticTexts.add({staticLabel:"Enter the Base Name for \"" + tempENTRY.theDoc.name + "\""} );
					var newBaseName = textEditboxes.add({editContents:baseName, minWidth:135} );
					}
					with(dialogRows.add() )
						staticTexts.add({staticLabel:"", minWidth:400} );
				}
			}
			if(!nameDialog.show() ){
				nameDialog.destroy();
				byeBye("User canceled export.",peuINFO.sayCancel);
			}
			else{
				tempENTRY.baseName = removeColons(removeSpaces(newBaseName.editContents) );
				nameDialog.destroy();
				// Determine if page replacement token exists when the page token option is used
				if(peuINFO.pageNamePlacement == 2){
					var temp = tempENTRY.baseName.indexOf("<#>");
					if(temp == -1){//Token isn't there
						alert("There is no page item token (<#>) in the base name. Please add one or click cancel in the next dialog box.");
						tempENTRY.getOut = false;
					}
				}	
				else // Try to remove any <#>s as a precaution
					tempENTRY.baseName = tempENTRY.baseName.replace(/<#>/g,"");	
				}
		}while(!tempENTRY.getOut);

		// Get names of all the pages. Needed when pages are named using sectioning
		tempENTRY = recordPgNames(tempENTRY);

		// Set pgStart and pgEnd, forcing "All" pages to output
		tempENTRY = (checkPages(tempENTRY, "All"))[0];

		// The page count is all pages due to common batching
		tempPageCount = tempENTRY.theDoc.pages.length;

		// This info is common, get it from commonBatchINFO:
		tempENTRY.pageNamePlacement = commonBatchINFO.pageNamePlacement;
		tempENTRY.outDir = commonBatchINFO.outDir;
		tempENTRY.nameConvType = commonBatchINFO.nameConvType
		tempENTRY.doSpreadsON = commonBatchINFO.doSpreadsON;
		tempENTRY.doOneFile = commonBatchINFO.doOneFile;
		tempENTRY.formatType = commonBatchINFO.formatType;
		tempENTRY.psINFO = commonBatchINFO.psINFO;
		tempENTRY.pdfPRESET = commonBatchINFO.pdfPRESET
		tempENTRY.epsINFO = commonBatchINFO.epsINFO
		tempENTRY.jpegINFO = commonBatchINFO.jpegINFO;
	}

	// Get any layering info
	if(peuINFO.layersON){
		tempENTRY.layerINFO = layerManager(tempENTRY.theDoc);
		if (tempENTRY.layerINFO == null) // Only one layer, turn it off for this doc
			tempENTRY.layersON = false;
		else
			tempENTRY.layersON = true;
	}

	// Sum up pages for the grand total for use in progress bar
	var temp = 1;
	if(peuINFO.doProgressBar && tempENTRY.layersON){
		// Figure tally for progress bar to include versions
		for(i=0;i < tempENTRY.layerINFO.verControls.length; i++)
			if (tempENTRY.layerINFO.verControls[i] == 1)
				temp++;
		if(!peuINFO.baseLaersAsVersion)
			temp--;
	}
	progTotalPages += (tempPageCount*temp);

	// All info for this doc is finally gathered, add it to the main printINFO array
	printINFO.push(tempENTRY);

	// Only one chance to change prefs: trigger singleton
	pseudoSingleton++;

}// end of main for loop

savePrefs(); // Record any changes

// Initiallize progress bar if available
if(peuINFO.doProgressBar)
	app.createProgressBar("Exporting Pages...", 0, progTotalPages, true);

// Export by looping through all open documents if using batch option, otherwise just the front document is exported
for(currentDoc = 0; currentDoc < printINFO.length; currentDoc++){
	var currentINFO = printINFO[currentDoc];

	// Set message in progress bar if available
	if(peuINFO.doProgressBar){
			var progCancel = app.setProgress(currentINFO.theDoc.name);
			if(progCancel)
				byeBye("User canceled export.",peuINFO.sayCancel);
	}

	// Set format options here so it's done just once per document
	setExportOption(currentINFO);

	// "Do one file" or PS/PDF with one page:
	if (currentINFO.doOneFile || currentINFO.singlePage){
		// Remove page token if it was entered and this name positioning option is set
		currentINFO.baseName = currentINFO.baseName.replace(/<#>/g,"");

		if(currentINFO.layersON){
			var theLayers = currentINFO.theDoc.layers;
			var baseControls = currentINFO.layerINFO.baseControls;
			var versionControls = currentINFO.layerINFO.verControls;
			var lastVersion = -1;

			// Loop for versioning
 			for(v = 0; v < versionControls.length; v++){
				if(!versionControls[v])
					continue;
				if(lastVersion != -1)// Turn the last layer back off
					theLayers[lastVersion].visible = false;
				lastVersion = v;
				theLayers[v].visible = true;
				currentINFO.outfileName = addPartToName(currentINFO.baseName, theLayers[v].name, peuINFO.layerBeforeON)

				// Export this version
				exportPage(currentINFO, PageRange.allPages);

				// Advance progress bar if available
				if(peuINFO.doProgressBar)
					advanceBar();
			}
			// If Base layer/s is/are to be output as a version, do it now
			if(peuINFO.baseLaersAsVersion){
				lastVersion = -1;
				// Turn off all versioning layers
				for(v = 0; v < currentINFO.layerINFO.baseControls.length; v++){
					if(currentINFO.layerINFO.baseControls[v])// its a base layer, keep track of last base version layer number
						lastVersion = v;
					else
						theLayers[v].visible = false;
				}
				if (!lastVersion == -1){// Only export if there was a base version
					currentINFO.outfileName = addPartToName(currentINFO.baseName, theLayers[lastVersion].name, peuINFO.layerBeforeON);

					// Export the base layer(s)
					exportPage(currentINFO, PageRange.allPages);

					// Advance progress bar if available
					if(peuINFO.doProgressBar)
						advanceBar();
				}
			}
		}
		else{ // No layer versioning, just export
			currentINFO.outfileName = currentINFO.baseName;
			// Export the base layer(s)
			exportPage(currentINFO, PageRange.allPages);

			// Advance progress bar if available
			if(peuINFO.doProgressBar)
				advanceBar();
		}

		if(!peuINFO.batchON)
			byeBye("Done exporting as a single file.",true);
	}
	else{ // Do single pages/spreads
		if (!currentINFO.hasNonContig)
			// Pages are contiguous, can just export
			outputPages(currentINFO.pgStart, currentINFO.pgEnd, currentINFO);
		else{ // Export non-contiguous
			// Loop through array of page sections
			for (ii = 0; ii < currentINFO.nonContigPgs.length; ii++){
				temp = currentINFO.nonContigPgs[ii];
				// Here we handle the start/end pages for any non-contig that has "-"
				if (temp.indexOf("-") != -1){
					temp2 = temp.split("-");
					outputPages(temp2[0],temp2[1], currentINFO);
				}
				else // The non-contiguous page is a single page
					outputPages(temp, temp, currentINFO);
			}
		}
	}

	// Set the spread settings back to what it was originally
	try{
		switch (currentINFO.formatType){
			case 0: // PostScript Formatting
				theDoc.printPreferences.printSpreads = peuINFO.origSpread;
				break;
			case 1: // PDF Formatting
				currentINFO.pdfPRESET.exportReaderSpreads = peuINFO.origSpread;
				break;
			case 2: // EPS Formatting
				app.epsExportPreferences.epsSpreads = peuINFO.origSpread;
				break;
			case 3: // JPEG Formatting
				app.jpegExportPreferences.exportingSpread = peuINFO.origSpread;
				break;
		}
	}
	catch(e){/*Just ignore it*/}

}

byeBye("The requested pages are done being exported.",true); // Last line of script execution

/*******************************************/
/*         Operational Functions           */
/*******************************************/

/*
 * Handle exporting
 */
function outputPages(pgStart, pgEnd, currentINFO){
	var pgRange;
	var layerName = "";
	var numVersions;
	var currentPage;
	var lastVersion = -1;
	var numericallyLastPage;

	if (currentINFO.layersON){
		var theLayers = currentINFO.theDoc.layers;
		var baseControls = currentINFO.layerINFO.baseControls;
		var versionControls = currentINFO.layerINFO.verControls;
		numVersions = versionControls.length;

		// Compensate for base layers as a version
		if(peuINFO.baseLaersAsVersion)
			numVersions++;
	}
	else
		numVersions = 1;
	for (v = 0; v < numVersions; v++){
		if(currentINFO.layersON){
			if(v == (numVersions - 1) && peuINFO.baseLaersAsVersion){
				var currentLayer = -1;
				// Base layer(s) are to be output as a version
				// Turn off all versioning layers
				for(slbm = 0; slbm < baseControls.length; slbm++){
					if(baseControls[slbm])// its a base layer, use its name for page name
						currentLayer = slbm;
					else
						theLayers[slbm].visible = false;
				}
				// Check if there was no base layer at all
				if (currentLayer == -1)
					layerName = "**NO_BASE**"
				else
					layerName = theLayers[currentLayer].name;
			}
			else{
				if(!versionControls[v])
					continue;
				if(lastVersion != -1)// Turn the last layer back off
					theLayers[lastVersion].visible = false;
				lastVersion = v;
				theLayers[v].visible = true;
				layerName = theLayers[v].name;
			}
		}

		if (currentINFO.nameConvType == 4){
			currentPage = pgStart;
			numericallyLastPage = pgEnd;
		}
		else if (currentINFO.doSpreadsON){
			currentPage = pgStart - 1;
			numericallyLastPage = pgEnd;
		}
		else {
			currentPage = getPageOffset(pgStart, currentINFO.pageNameArray , currentINFO.pageRangeArray);
			numericallyLastPage = getPageOffset(pgEnd, currentINFO.pageNameArray, currentINFO.pageRangeArray);
		}
		if(layerName != "**NO_BASE**"){
			do{
				currentINFO.outfileName = getPageName(currentPage, layerName, currentINFO);
				if (currentINFO.doSpreadsON){
					pgRange = currentINFO.pageRangeArray[getPageOffset(currentINFO.theDoc.spreads[currentPage].pages[0].name, currentINFO.pageNameArray, currentINFO.pageRangeArray)];
				}
				else if (currentINFO.nameConvType == 4)
					pgRange = currentINFO.pageRangeArray[currentPage-1];
				else
					pgRange = currentINFO.pageRangeArray[currentPage];

				// Do the actual export:
				exportPage(currentINFO, pgRange);

				// Update progress bar if available
				if(peuINFO.doProgressBar)
					advanceBar();

				currentPage++;
			} while(currentPage <= numericallyLastPage);
		}
	}
}

/*
 * Export the page
 */
function exportPage(currentINFO, pgRange){
	var outFile = currentINFO.outDir + currentINFO.outfileName;
	switch (currentINFO.formatType){
		case 0: // PostScript Formatting
			with(currentINFO.theDoc.printPreferences){
				printFile = new File(outFile + ((currentINFO.psINFO.ext)?".ps":""));
				pageRange = pgRange;
			}
			// Needed to get around blank pages using separations
			try{
				currentINFO.theDoc.print(false);
			}
			catch(e){/*Just skip it*/}
			break;
		case 1: // PDF Formatting
			app.pdfExportPreferences.pageRange = pgRange;
			currentINFO.theDoc.exportFile(ExportFormat.pdfType, (new File(outFile + ".pdf")), false, currentINFO.pdfPRESET);
			break;
		case 2: // EPS Formatting
			app.epsExportPreferences.pageRange = pgRange;
			currentINFO.theDoc.exportFile(ExportFormat.epsType, (new File(outFile + ".eps")), false);
			break;
		case 3: // JPEG Formatting
			if(pgRange == PageRange.allPages){
				app.jpegExportPreferences.jpegExportRange = ExportRangeOrAllPages.exportAll;
			}
			else{
				app.jpegExportPreferences.jpegExportRange = ExportRangeOrAllPages.exportRange;
				app.jpegExportPreferences.pageString = pgRange;
			}
			currentINFO.theDoc.exportFile(ExportFormat.jpg, (new File(outFile + ".jpg")), false);
			break;
	}
}

/*
 * Create a name for the page being exported
 */
function getPageName(currentPage, layerName, currentINFO){
	var pgRename = "";
	if (currentINFO.doSpreadsON)
		currentINFO["currentSpread"] = currentINFO.theDoc.spreads[currentPage].pages;
	switch (currentINFO.nameConvType){
		case 3: // Odd/Even pages/spreads = .LA.F/LA.B, LB.F/LB.B ...
			pgRename = makeLotName(currentPage+1, peuINFO.subType);
			break;
		case 2: // Odd/Even pages/spreads = .F/.B
			pgRename = ((currentPage+1)%2 == 0) ? "B" : "F";
			break;
		case 1: // Add ".L" to the page name
			pgRename = "L" + currentINFO.pageNameArray[currentPage];
			break;
		case 0: case 4:// As is or Numeric Override
						// Optionally add "P" and any zeros if options chosen and is numerically named
						// otherwise, just the "seperatorChar" is added to page name
			if (currentINFO.doSpreadsON){
				// Loops through number of pages per spread
				// and adds each page name to the final name (P08.P01)
				for (j = 0; j < currentINFO.currentSpread.length; j++){
					if(currentINFO.currentSpread[j].appliedSection.includeSectionPrefix)
						var tempPage = currentINFO.pageRangeArray[getPageOffset(currentINFO.currentSpread[j].name, currentINFO.pageNameArray, currentINFO.pageRangeArray)];
					else
						var tempPage = currentINFO.pageNameArray[getPageOffset(currentINFO.currentSpread[j].name, currentINFO.pageNameArray, currentINFO.pageRangeArray)];
					var tempPageNum = parseInt(tempPage,10);
					/* If section name starts with a number, need to compare length of orig vs parsed
					 * to see if the page name is solely a number or a combo num + letter, etc.
					 */
					if (! isNaN(tempPageNum) && ((""+tempPage).length == (""+tempPageNum).length )){
						if (peuINFO.addZeroON)
							tempPage = addLeadingZero(tempPageNum, currentINFO.theDoc.pages.length);
						if (peuINFO.addPon)
							tempPage = "P" + tempPage;
					}
					pgRename = (j==0) ? tempPage : pgRename + peuINFO.charList[peuINFO.seperatorChar] + tempPage;
				}
			}
			else {
				// Create a new name for an individual page
				if (currentINFO.nameConvType == 4)
					pgRename = currentPage;
				else
					pgRename = currentINFO.pageNameArray[currentPage];
				if (! isNaN(parseInt(pgRename,10)) && (""+pgRename).length == (""+parseInt(pgRename,10)).length) {
					if (peuINFO.addZeroON)
						pgRename = addLeadingZero(pgRename, currentINFO.theDoc.pages.length);
					if (peuINFO.addPon)
						pgRename = "P" + pgRename;
				}
			}
			break;
		}

	if(currentINFO.layersON)
		pgRename = addPartToName(pgRename, layerName, peuINFO.layerBeforeON);

	// Add page name to base name based on option selected
	if(peuINFO.pageNamePlacement == 2)
		pgRename = removeColons(currentINFO.baseName.replace(/<#>/g,pgRename) );
	else
		pgRename = addPartToName(currentINFO.baseName, pgRename,peuINFO.pageNamePlacement);
	return pgRename;
}

/*
 * Add a name part before or after a given base string
 */
function addPartToName(theBase, addThis, addBefore){
	//Remove any colons
	theBase = removeColons(theBase);
	addThis = removeColons(addThis);
	return (addBefore) ? (addThis + peuINFO.charList[peuINFO.seperatorChar] + theBase ):(theBase + peuINFO.charList[peuINFO.seperatorChar] + addThis);
}

/*
 * Find the offset page number for a page by its name
 */
function getPageOffset(pgToFind, pageNameArray, pageRangeArray){
	var offset;
	for(offset = 0; offset<pageRangeArray.length;offset++){
		if((""+ pgToFind).toLowerCase() == (("" + pageNameArray[offset]).toLowerCase() ) || (""+ pgToFind).toLowerCase() == (("" + pageRangeArray[offset]).toLowerCase() ) )
			return offset;
	}
	return -1;
}

/*
 * Replace any colons with specialReplaceChar
 */
function removeColons(tempName){
	return tempName.replace(/:/g,peuINFO.charList[peuINFO.specialReplaceChar]);
}

/*
 * Remove spaces from front and end of name
 */
function removeSpaces(theName){
	// Trim any leading or trailing spaces in base name
	var i,j;
	for(i = theName.length-1;i>0 && theName.charAt(i) == " ";i--);// Ignore any spaces on end of name
	for(j = 0; j<theName.length && theName.charAt(j) == " ";j++);// Ignore any spaces at front of name
	theName = theName.substring(j,i+1);
	return theName
}

/*
 * Add leading zero(s)
 */
function addLeadingZero(tempPageNum, pageCount){
	if(peuINFO.zeroPadding == 0){
		// Normal padding
		if((tempPageNum < 10 && pageCount < 100) || (tempPageNum > 9 && pageCount > 99 && tempPageNum < 100))
			return addSingleZero(tempPageNum);
		else if(tempPageNum < 10 && pageCount > 99)
			return addDoubleZero(tempPageNum);
		else
			return ("" + tempPageNum);
	}else if(peuINFO.zeroPadding == 1){
		// Pad to 2 digits
		if(tempPageNum < 10)
			return addSingleZero(tempPageNum);
		else
			return ("" + tempPageNum);
	}else{
		// Pad to 3 digits
		if(tempPageNum < 10)
			return addDoubleZero(tempPageNum);
		else if(tempPageNum < 100)
			return addSingleZero(tempPageNum);
		else
			return ("" + tempPageNum);
	}
}

/*
 * Add leading zero helper for single
 */
function addSingleZero(pgNum){
	return ("0" + pgNum);
}

/*
 * Add leading zero helper for double
 */
function addDoubleZero(pgNum){
	return ("00" + pgNum);
}

/*
 * Create lot name from page number
 */
function makeLotName(thePage, subType){
	var iii = thePage;
	var curr = 0;
	var alphaBet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
	var lotName = "L";
	if(subType == 0){
		while(iii>52){
			curr = Math.floor((iii-1)/52)-1;
			lotName += alphaBet[curr];
			if(curr >= 0)
				iii -= 52*(1+curr);
			else
				iii -= 52;
		}
		lotName += alphaBet[Math.floor((iii-1)/2)%26];
	}
	else
		for(iii=thePage; iii>0; iii-=52)
			lotName += alphaBet[Math.floor((iii-1)/2)%26];

	return lotName += (thePage & 0x1)?".F":".B";
}

/*
 * Advance progress bar one unit
 */
function advanceBar(){
	var progCancel = app.setProgress(++progCurrentPage);
	if(progCancel)
		byeBye("User canceled export.",peuINFO.sayCancel);
}
/*
 * Create an Empty tempENTRY "struct"
 */
function getNewTempENTRY(){
	var newTempENTRY = new Array();
	newTempENTRY["theDoc"] = null;
	newTempENTRY["singlePage"] = null;
	newTempENTRY["getOut"] = null;
	newTempENTRY["outDir"] = null;
	newTempENTRY["outfileName"] = "";
	newTempENTRY["nameConvType"] = null;
	newTempENTRY["baseName"] = null;
	newTempENTRY["doSpreadsON"] = null;
	newTempENTRY["doOneFile"] = null;
	newTempENTRY["formatType"] = null;
	newTempENTRY["layersON"] = null;
	newTempENTRY["hasNonContig"] = false;
	newTempENTRY["nonContigPgs"] = null;
	newTempENTRY["pageNameArray"] = new Array();
	newTempENTRY["pageRangeArray"] = new Array();
	newTempENTRY["psINFO"] = null;
	newTempENTRY["pdfPRESET"] = null;
	newTempENTRY["epsINFO"] = null;
	newTempENTRY["jpegINFO"] = null;
	newTempENTRY["layerINFO"] = null;
	newTempENTRY["useTag"] = null;
	newTempENTRY["pgStart"] = null;
	newTempENTRY["pgEnd"] = null;
	return newTempENTRY;
}

/*
 * Record all the page/spread names
 */
function recordPgNames(tempENTRY){
		// Get names of all the pages. Needed when pages are named using sectioning
	for (i = 0; i < tempENTRY.theDoc.documentPreferences.pagesPerDocument; i++){
		var aPage = tempENTRY.theDoc.pages.item(i);
		tempENTRY.pageNameArray[i] = aPage.name;
		tempENTRY.pageRangeArray[i] = (aPage.appliedSection.includeSectionPrefix)? aPage.name : (aPage.appliedSection.name + aPage.name);
	}
	return tempENTRY;
}

/*
 * Set the export options
 */
function setExportOption(currentINFO){
		// Set any options here instead of with each page
	switch (currentINFO.formatType){
		case 0: // PostScript Formatting
			setPSoptions(currentINFO);
			break;
		case 1: // PDF Formatting
			// Nothing to do
			break;
		case 2: // EPS Formatting
			setEPSoptions(currentINFO.epsINFO);
			break;
		case 3: // JPEG Formatting
			setJPEGoptions(currentINFO.jpegINFO);
			break;
	}
}

/*
 * Get PostScript format options
 */
function getPSoptions(docName){
	var psOptions = new Array();
	psOptions["ignore"] = true;
	var tempGetOut, PSdlog, pgHeight, pgWidth;
	var changeAddPSextention, tempBaseName;
	var printPreset;
	do{
		tempGetOut = true;
		PSdlog = app.dialogs.add({name:"PostScript Options for \"" + docName + "\"", canCancel:true} );
		with (PSdlog)
			with (dialogColumns.add() ){
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Print Presets:"} );
				printPreset = dropdowns.add({stringList:peuINFO.psPrinterNames , minWidth:236, selectedIndex:peuINFO.defaultPrintPreset} );
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Override PS Page Size (" + peuINFO.measureLableArray[peuINFO.measurementUnits] + ")"} );
				with (borderPanels.add() )
					with (dialogColumns.add() )
						with (dialogRows.add() ){
							staticTexts.add({staticLabel:"Width:", minWidth:45} );
							pgWidth = textEditboxes.add({editContents:"0", minWidth:53} );
							staticTexts.add({staticLabel:"Height:", minWidth:45} );
							pgHeight = textEditboxes.add({editContents:"0", minWidth:54} );
						}
				with (dialogRows.add() ){
					staticTexts.add({staticLabel:"Add \".ps\" to end of file name"} );
					changeAddPSextention = dropdowns.add({stringList:["No","Yes"], selectedIndex:peuINFO.addPSextention} )
		}
			}
		if((PSdlog.show()) ){
			// Get the page height + width
			pgHeight = parseFloat(pgHeight.editContents);
			pgWidth = parseFloat(pgWidth.editContents);

			// Check entered H & W for error
			if(isNaN(pgHeight) || isNaN(pgWidth) || pgHeight < 0 || pgWidth < 0 ){
				alert ("Both page height and width must be numeric and greater than zero (3.1).");
				pgHeight = "0";
				pgWidth  = "0";
				tempGetOut = false;
				continue;
			}
			if(pgHeight > 0 && pgWidth > 0) // User changed size, use the new size
				psOptions.ignore = false;

			psOptions["height"] = pgHeight + peuINFO.measureUnitArray[peuINFO.measurementUnits];
			psOptions["width"] = pgWidth + peuINFO.measureUnitArray[peuINFO.measurementUnits];
			psOptions["ext"] = changeAddPSextention.selectedIndex;
			peuINFO.addPSextention = psOptions["ext"];
			psOptions["preset"] = printPreset.selectedIndex
			peuINFO.defaultPrintPreset = psOptions.preset;
			savePrefs();
			PSdlog.destroy();
		}
		else{
			PSdlog.destroy();
			byeBye("Exporting has been canceled by user.",peuINFO.sayCancel);
		}
	} while(!tempGetOut);
	return psOptions;
}

/*
 * Set Postscript options
 */
function setPSoptions(theINFO){
	with(currentINFO.theDoc.printPreferences){
		activePrinterPreset = peuINFO.csPSprinters[theINFO.psINFO.preset];
		peuINFO.origSpread = printSpreads; // Used to reset to original state when done
		printSpreads = theINFO.doSpreadsON;
		if(colorOutput != ColorOutputModes.separations && colorOutput != ColorOutputModes.inripSeparations)
			printBlankPages = true;
		if (theINFO.useTag)
			pageInformationMarks = true;
		else
			pageInformationMarks = false;
		if(!theINFO.psINFO.ignore){
			try{
				paperSize = PaperSizes.custom;
				paperHeight = theINFO.psINFO.height;
				paperWidth = theINFO.psINFO.width;
			}
			catch(Exception){
				alert ("The current PPD doesn't support custom page sizes. The page size from the Print Preset will be used (3.2).");
			}
		}
	}
}

/*
 * Get PDF options
 */
function getPDFoptions(theINFO){
	var PDFdlog = app.dialogs.add({name:"PDF Options for \"" + theINFO.theDoc.name.split(".ind")[0] + "\"", canCancel:true} );
	var temp = new Array();
	for(i=0;i<app.pdfExportPresets.length;i++)
		temp.push(app.pdfExportPresets[i].name);
	// Test if default PDFpreset is greater # than actual list.
	// This occurs if one was deleted and the last preset in the list was the default
	if(peuINFO.defaultPDFpreset > temp.length-1)
		peuINFO.defaultPDFpreset = 0;
	with (PDFdlog)
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"PDF Export Preset:"} );
			pdfPresets = dropdowns.add({stringList: temp, minWidth:50, selectedIndex:peuINFO.defaultPDFpreset} );
		}
	if(PDFdlog.show() ){
		temp = app.pdfExportPresets[pdfPresets.selectedIndex];
		peuINFO.defaultPDFpreset = pdfPresets.selectedIndex;
		peuINFO.origSpread = temp.exportReaderSpreads;
		try{
			temp.exportReaderSpreads = theINFO.doSpreadsON;
			temp.pageInformationMarks = (theINFO.useTag && temp.cropMarks)?true:false;
		}catch(e){/*ignore it*/}
		PDFdlog.destroy();
		return temp;
	}
	else{
		PDFdlog.destroy();
		byeBye("PDF exporting has been canceled by user.", peuINFO.sayCancel);
	}
}

/*
 * Get JPEG options
 */
function getJPEGoptions(docName){
	var temp = new Array();
	var JPEGdlog = app.dialogs.add({name:"JPEG Options for \"" + docName + "\"", canCancel:true} );
	with (JPEGdlog)
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Quality:"} );
			JPEGquality = dropdowns.add({stringList:(new Array("Low","Medium","High","Maximum")) , minWidth:50, selectedIndex:peuINFO.defaultJPEGquality} );
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Encoding Type:"} );
				JPEGrender = dropdowns.add({stringList:["Baseline","Progressive"] , minWidth:50, selectedIndex:peuINFO.defaultJPEGrender } );
		}
	if(JPEGdlog.show() ){
		peuINFO.defaultJPEGquality = JPEGquality.selectedIndex;
		temp["qualityType"] = peuINFO.defaultJPEGquality;
		peuINFO.defaultJPEGrender = JPEGrender.selectedIndex;
		temp["renderType"] = peuINFO.defaultJPEGrender;
	}
	else{
		JPEGdlog.destroy();
		byeBye("JPEG exporting has been canceled by user.",peuINFO.sayCancel);
	}
	JPEGdlog.destroy();
	return temp;
}

/*
 * Set JPEG options
 */
function setJPEGoptions(theINFO){
	with(app.jpegExportPreferences){
		peuINFO.origSpread = exportingSpread; // Used to reset to original state when done
		exportingSpread = currentINFO.doSpreadsON;
		exportingSelection = false; // Export the entire page
		if(peuINFO.csVersion > 3)
			jpegExportRange = ExportRangeOrAllPages.exportRange;
		switch (theINFO.qualityType){
			case 0:
				jpegQuality = JPEGOptionsQuality.low;
				break;
			case 1:
				jpegQuality = JPEGOptionsQuality.medium;
				break;
			case 2:
				jpegQuality = JPEGOptionsQuality.high;
				break;
			case 3:
				jpegQuality = JPEGOptionsQuality.maximum;
				break;
		}
		jpegRenderingStyle = (theINFO.renderType)? JPEGOptionsFormat.baselineEncoding : JPEGOptionsFormat.progressiveEncoding;
	}
}

/*
 * Get EPS options
 */
function getEPSoptions(docName){
	var epsOptions = new Array();
	var epsDialog = app.dialogs.add({name:"EPS Options for \"" + docName + "\"", canCancel:true} );
	var oldBleed = peuINFO.bleed;
	with (epsDialog){
		// Left Column
		with (dialogColumns.add() ){
			with (dialogRows.add() )
			with (borderPanels.add() )
				with (dialogColumns.add() ){
					with (dialogRows.add() )
						staticTexts.add({staticLabel:"Flattener Presets:"} );
					changeFlattenerPreset = dropdowns.add({stringList:peuINFO.flattenerNames , minWidth:180, selectedIndex:peuINFO.defaultFlattenerPreset} );
					with (dialogRows.add() )
						changeIgnoreOverride = checkboxControls.add({staticLabel:"Ignore Overrides", checkedState:peuINFO.ignoreON} );
					with (dialogRows.add() )
						staticTexts.add({staticLabel:"Preview Type:"} );
					changePreviewPreset = dropdowns.add({stringList:peuINFO.previewTypes , minWidth:180, selectedIndex:peuINFO.defaultPreview} );
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Bleed:"} );
						changeBleedVal = realEditboxes.add({editValue:peuINFO.bleed, minWidth:60} );
						staticTexts.add({staticLabel:peuINFO.measureLableArray[peuINFO.measurementUnits]} );
					}
					with (dialogRows.add() )
						staticTexts.add({staticLabel:"OPI Options:"} );
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Omit:"} );
						changeOpiEPS = checkboxControls.add({staticLabel:"EPS", checkedState:peuINFO.epsON} );
						changeOpiPDF = checkboxControls.add({staticLabel:"PDF", checkedState:peuINFO.pdfON} );
						changeOpiBitmap = checkboxControls.add({staticLabel:"Bitmapped", checkedState:peuINFO.bitmapON} );
					}
				}
		}
		// Right column
		with (dialogColumns.add() ){
			with(borderPanels.add() ){
				with(dialogColumns.add() ){
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"PostScript level:"} );
						var changePSlevel = dropdowns.add({stringList:["2","3"] , minWidth:75, selectedIndex:peuINFO.psLevel} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Color mode:"} );
						if(peuINFO.csVersion == 3)
							var changeColorMode = dropdowns.add({stringList:["Unchanged","Grayscale", "RGB", "CMYK"] , minWidth:100, selectedIndex:peuINFO.colorType } );
						else
							var changeColorMode = dropdowns.add({stringList:["Unchanged","Grayscale", "RGB", "CMYK","PostScript Color Management"] , minWidth:100, selectedIndex:peuINFO.colorType } );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Font embedding:"} );
						var changeFontEmbedding = dropdowns.add({stringList:["None","Complete", "Subset"] , minWidth:100, selectedIndex:peuINFO.fontEmbed } );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Type of data to send:"} );
						var changeDataToSend = dropdowns.add({stringList:["All","Proxy"] , minWidth:50, selectedIndex:peuINFO.dataSent } );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Data type:"} );
						var changeDataType = dropdowns.add({stringList:["Binary","ASCII"] , minWidth:50, selectedIndex:peuINFO.dataType } );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Perform OPI replacement:"} );
						var changeOPIreplace = dropdowns.add({stringList:["No","Yes"] , minWidth:50, selectedIndex:peuINFO.opiReplacement} );
					}
				}
			}
		}
	}
	do{
		var getOut = true;
		if((epsDialog.show()) ){
			// Use these to update the prefs file
			peuINFO.defaultFlattenerPreset = changeFlattenerPreset.selectedIndex;
			peuINFO.ignoreON = (changeIgnoreOverride.checkedState)?1:0;
			peuINFO.defaultPreview = changePreviewPreset.selectedIndex;
			peuINFO.bleed = changeBleedVal.editContents;
			peuINFO.epsON = (changeOpiEPS.checkedState)?1:0;
			peuINFO.pdfON = (changeOpiPDF.checkedState)?1:0;
			peuINFO.bitmapON = (changeOpiBitmap.checkedState)?1:0;
			peuINFO.psLevel = changePSlevel.selectedIndex;
			peuINFO.colorType = changeColorMode.selectedIndex;
			peuINFO.fontEmbed = changeFontEmbedding.selectedIndex;
			peuINFO.dataSent = changeDataToSend.selectedIndex;
			peuINFO.dataType = changeDataType.selectedIndex;
			peuINFO.opiReplacement = changeOPIreplace.selectedIndex;

			// Check if bleed value is OK
			peuINFO.bleed = parseFloat(peuINFO.bleed)
			if (isNaN(peuINFO.bleed)){
				alert("Bleed value must be a number (1.1).");
				getOut = false;
				peuINFO.bleed = oldBleed;
			}
			else if (peuINFO.bleed < 0){
				alert("Bleed value must be greater or equal to zero (1.2).");
				getOut = false;
				peuINFO.bleed = oldBleed;
			}
			else {
				// Check if bleed is too big
				try {
					app.epsExportPreferences.bleedBottom = "" + peuINFO.bleed + peuINFO.measureUnitArray[peuINFO.measurementUnits];
				}
				catch (Exception){
					alert("The bleed value must be less than one of the following: 6 in | 152.4 mm | 432 pt | 33c9.384");
					getOut = false;
					peuINFO.bleed = oldBleed;
				}
			}
		}
		else{
			epsDialog.destroy();
			byeBye("EPS Export canceled by user.", peuINFO.sayCancel);
		}
	}while(!getOut);

		// These are used for exporting
		epsOptions["defaultFlattenerPreset"] = changeFlattenerPreset.selectedIndex;
		epsOptions["ignoreON"] = peuINFO.ignoreON;
		epsOptions["defaultPreview"] = changePreviewPreset.selectedIndex;
		epsOptions["bleed"] = peuINFO.bleed;
		epsOptions["epsON"] = peuINFO.epsON;
		epsOptions["pdfON"] = peuINFO.pdfON;
		epsOptions["bitmapON"] = peuINFO.bitmapON;
		epsOptions["psLevel"] = changePSlevel.selectedIndex;
		epsOptions["colorType"] = changeColorMode.selectedIndex;
		epsOptions["fontEmbed"] = changeFontEmbedding.selectedIndex;
		epsOptions["dataSent"] = changeDataToSend.selectedIndex;
		epsOptions["dataType"] = changeDataType.selectedIndex;
		epsOptions["opiReplacement"] = changeOPIreplace.selectedIndex;
		epsDialog.destroy();
		return epsOptions;
}

/*
 * 	Apply chosen settings to the EPS export prefs
 */
function setEPSoptions(theINFO){
	with(app.epsExportPreferences){
		appliedFlattererPreset = peuINFO.flattenerNames[theINFO.defaultFlattenerPreset];
		bleedBottom = "" + theINFO.bleed + peuINFO.measureUnitArray[peuINFO.measurementUnits];
		bleedInside = bleedBottom;
		bleedOutside = bleedBottom;
		bleedTop = bleedBottom;
		epsSpreads = currentINFO.doSpreadsON;
		ignoreSpreadOverrides = theINFO.ignoreON;
		switch (theINFO.dataType){
			case 0:
				dataFormat = DataFormat.binary;
				break;
			case 1:
				dataFormat = DataFormat.ascii;
				break;
		}
		switch (theINFO.colorType){
			case 0:
				epsColor = EPSColorSpace.unchangedColorSpace;
				break;
			case 1:
				epsColor = EPSColorSpace.gray;
				break;
			case 2:
				epsColor = EPSColorSpace.rgb;
				break;
			case 3:
				epsColor = EPSColorSpace.cmyk;
				break;
			case 4:
				epsColor = EPSColorSpace.postscriptColorManagement;
				break;
		}
		switch (theINFO.fontEmbed){
			case 0:
				fontEmbedding = FontEmbedding.none;
				break;
			case 1:
				fontEmbedding = FontEmbedding.complete;
				break;
			case 2:
				fontEmbedding = FontEmbedding.subset;
				break;
		}
		switch (theINFO.dataSent){
			case 0:
				imageData = EPSImageData.allImageData;
				break;
			case 1:
				imageData = EPSImageData.proxyImageData;
				break;
		}
		switch (theINFO.defaultPreview){
			case 0:
				preview = PreviewTypes.none;
				break;
			case 1:
				preview = PreviewTypes.tiffPreview;
				break;
			case 2:
				preview = PreviewTypes.pictPreview;
				break;
		}
		switch (theINFO.psLevel){
			case 0:
				postScriptLevel = PostScriptLevels.level2;
				break;
			case 1:
				postScriptLevel = PostScriptLevels.level3;
				break;
		}

		// Setting these three to false prevents a conflict error when trying to set the opiImageReplacement value
		omitBitmaps = false;
		omitEPS = false;
		omitPDF = false;
		if (theINFO.opiReplacement){
			opiImageReplacement = true;
		}
		else {
			opiImageReplacement = false;
			omitBitmaps = theINFO.bitmapON;
			omitEPS = theINFO.epsON;
			omitPDF = theINFO.pdfON;
		}
	}
}

/*
 * Build the main dialog box
 */
function createMainDialog (docName, thisNum, endNum){
	var theDialog = app.dialogs.add({name:(VERSION_NAME + ": Enter the options for \"" + docName + "\"" + ((endNum==1)?"":" (" + thisNum + " of " + endNum + " documents)") )
, canCancel:true} );
	with (theDialog){
		// Left Column
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Page Naming Options"} );
			with(borderPanels.add() ){
				with(dialogColumns.add() ){
					// Radio butons for renaming convention
					namingConvention = radiobuttonGroups.add();
					with(namingConvention){
						radiobuttonControls.add({staticLabel:"As Is", checkedState:(peuINFO.nameConvType == 0)} );
						radiobuttonControls.add({staticLabel:"Add \".L\"", checkedState:(peuINFO.nameConvType == 1)} );
						radiobuttonControls.add({staticLabel:"Odd/Even = \".F/.B\"", checkedState:(peuINFO.nameConvType == 2)} );
						radiobuttonControls.add({staticLabel:"Odd/Even = \"LA.F/LA.B\"" , checkedState:(peuINFO.nameConvType == 3)});
						radiobuttonControls.add({staticLabel:"Numeric Override" , checkedState:(peuINFO.nameConvType == 4)});
					}
				with (dialogRows.add() )
					staticTexts.add({staticLabel:""} );
					with(dialogRows.add() ){
					staticTexts.add({staticLabel:"Base Name"} );
					newBaseName = textEditboxes.add({editContents:baseName, minWidth:100} );
					}
				}
			}
		}
		// Middle Column
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Page Export Options"} );
			with (borderPanels.add() ){
				with (dialogColumns.add() ){
					with (dialogRows.add() ){
						// Start and end pages text entry box
						staticTexts.add({staticLabel:"Start[-End] Page"} );
						startEndPgs = textEditboxes.add({editContents:peuINFO.editStartEndPgs, minWidth:103} );
					}
					with (dialogRows.add() )
						doSpreads = checkboxControls.add({staticLabel:"As Spreads", checkedState:peuINFO.doSpreadsON} );
					with (dialogRows.add() )
						oneFile = checkboxControls.add({staticLabel:"Send Entire File At Once (PS+PDF only)", checkedState:peuINFO.doOneFile});
				}
			}
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Output Directory"} );
			with (borderPanels.add() )
				outputDirs = dropdowns.add({stringList:peuINFO.outDirNameArray , minWidth:204, selectedIndex:peuINFO.defaultDir} );
		}
		// Right Column 1
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"Output Format"} );
			with (borderPanels.add() ){
				with (dialogColumns.add() ){
					with (dialogRows.add() )
						formatTypeRB = radiobuttonGroups.add();
					with(formatTypeRB){
						// If there are no print presets, don't add radio button for it
						if(peuINFO.adjustForNoPS == 0)
							radiobuttonControls.add({staticLabel:"PostScript", checkedState:(peuINFO.exportDefaultType ==0)} );
						radiobuttonControls.add({staticLabel:"PDF", checkedState:((peuINFO.exportDefaultType==1)||(peuINFO.adjustForNoPS==1 && peuINFO.exportDefaultType==0))} );
						radiobuttonControls.add({staticLabel:"EPS", checkedState:(peuINFO.exportDefaultType==2)} );
						radiobuttonControls.add({staticLabel:"JPEG", checkedState:(peuINFO.exportDefaultType==3), minWidth:140} );
						if(pseudoSingleton < 1)
							radiobuttonControls.add({staticLabel:"Change Preferences", checkedState:false} );
					}
				}
			}
			if(pseudoSingleton < 2){
				pseudoSingleton++;
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Advanced Export Options"} );
				with (borderPanels.add() )
					with (dialogColumns.add() ){
						with (dialogRows.add() )
							doLayers = checkboxControls.add({staticLabel:"Layer Versioning", checkedState:peuINFO.layersON, minWidth:140});
						if(peuINFO.numDocsToExport > 1){
									with (dialogRows.add() )
									doBatch = checkboxControls.add({staticLabel:"Batch Export Enabled", checkedState:peuINFO.batchON, minWidth:130});
									with (dialogRows.add() )
										commonBatch = checkboxControls.add({staticLabel:"Set For All Batch Jobs", checkedState:peuINFO.batchSameForAll, minWidth:140});
						}
					}
			}
		}
	}
	return theDialog;
}

/*
 * Test if doc width can support page info.
 */
function usePgInfoTag(theUnits, docWidth){
	var usePgInfo = true;
	switch (theUnits){
		case MeasurementUnits.points:
			if(docWidth < 337.5)
				usePgInfo = false;
		break;
		case MeasurementUnits.picas:
			if(docWidth < 28.125)
				usePgInfo = false;
		break;
		case MeasurementUnits.inches:
		case MeasurementUnits.inchesDecimal:
			if(docWidth < 4.6875)
				usePgInfo = false;
		break;
		case MeasurementUnits.millimeters:
			if(docWidth < 119.0625)
				usePgInfo = false;
		break;
		case MeasurementUnits.centimeters:
			if(docWidth < 11.90625 )
				usePgInfo = false;
		break;
		case MeasurementUnits.ciceros:
			if(docWidth < 26.3922)
				usePgInfo = false;
		break;
	}
	return usePgInfo;
}

/*
 * Optionally display alert and get out of Dodge
 */
function byeBye(msg,doTell){
	if(peuINFO.doProgressBar)
		app.setProgress(999999);

	if(peuINFO.csVersion == 3)
		app.userInteractionLevel = peuINFO.oldInteractionPref;
	else
		app.scriptPreferences.userInteractionLevel = peuINFO.oldInteractionPref;
	if(doTell)
		alert(msg);
	exit();
}

/*
 * Let user select a directory
 */
function getDirectory(showText, startingDirectory){
	return Folder.selectDialog(showText, startingDirectory);
}

/*
 * The Layer Manager
 */
function layerManager(theDoc){
	var theLayers = theDoc.layers;
	var baseControls = new Array();
	var verControls = new Array();
	var lmi,getOut,numVerChosen;
	if (theLayers.length < 2){
  	alert("Versioning Error:\nYou need at least two layers for versioning support.\nVersioning is turned off for \"" + theDoc.name + "\".");
  	return null;
	}

	do{
		getOut = true;
		var layerDialog = app.dialogs.add({name:"Version Manager for \"" + theDoc.name + "\"", canCancel:true} );
		with (layerDialog){
			with (dialogColumns.add() ){
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Select Base Layer(s)     "} );
				with(borderPanels.add() ){
					with(dialogColumns.add() ){
						for(lmi = 0; lmi<theLayers.length;lmi++){
							baseControls[lmi] = checkboxControls.add({staticLabel:theLayers[lmi].name, checkedState:false, minWidth:104});
						}
						baseControls[lmi-1].checkedState = true;
					}
				}
			}
			with (dialogColumns.add() ){
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Select Version Layer(s)"} );
				with(borderPanels.add() ){
					with(dialogColumns.add() ){
						for(lmi = 0; lmi<theLayers.length;lmi++){
							verControls[lmi] = checkboxControls.add({staticLabel:theLayers[lmi].name, checkedState:true, minWidth:104});
						}
						verControls[lmi-1].checkedState = false;
					}
				}
			}
		}
		if(layerDialog.show() ){
			numVerChosen = 0;
			for(lmi=0; lmi< baseControls.length; lmi++){
				if(baseControls[lmi].checkedState == true && verControls[lmi].checkedState == true){
					alert("A layer cannot be both a base and version layer.\nPlease try again.");
					getOut = false;
					break;
				}
				if(verControls[lmi].checkedState == true){
					numVerChosen++;
					verControls[lmi] = true;
				}
				else
					verControls[lmi] = false;

				// Turn all layers off except base layers
				theLayers[lmi].visible = baseControls[lmi].checkedState;
				baseControls[lmi] = baseControls[lmi].checkedState;
			}
			if(numVerChosen == 0 && getOut){
				alert("At least one version layer needs to be chosen.\nPlease try again.");
				getOut = false;
			}
		}
		else{ // user canceled
		  layerDialog.destroy();
			byeBye("Exporting was canceled by user.",peuINFO.sayCancel);
		}
	} while(!getOut);

	layerDialog.destroy();
	var theControls = new Array();
	theControls["baseControls"] = baseControls;
	theControls["verControls"] = verControls;
	return theControls;
}

/*
 * Check Pages for Validity
 */
function checkPages(tempENTRY, startEndPgs){
	// Set Start and End Pages
	var pgsNoGood = false;
	var spreadsNoGood = false;
	var notNumber = false;
	var allSpreads = tempENTRY.theDoc.spreads;
	var spreadCount = allSpreads.length;
	var tempPageCount = 0;
	if (startEndPgs == "All"){
		if (tempENTRY.doSpreadsON){
			tempENTRY.pgStart = 1;
			tempENTRY.pgEnd = spreadCount;
		}
		else if (tempENTRY.nameConvType == 4){
			tempENTRY.pgStart = 1;
			tempENTRY.pgEnd = tempENTRY.theDoc.documentPreferences.pagesPerDocument;
		}
		else {
			tempENTRY.pgStart = tempENTRY.pageRangeArray[0];
			tempENTRY.pgEnd = tempENTRY.pageRangeArray[tempENTRY.pageRangeArray.length-1];
		}

		// Tally pages for the progress bar
		tempPageCount = (tempENTRY.doSpreadsON)?spreadCount:tempENTRY.theDoc.documentPreferences.pagesPerDocument;
	}
	else if (startEndPgs.indexOf("-") != -1 && startEndPgs.indexOf(",") == -1){
		// Page range contains a "-" but not a "," (ie 11-15)
		var temp = startEndPgs.split("-");
		tempENTRY.pgStart = temp[0];
		tempENTRY.pgEnd = temp[1];
		// Test if any contiguous pages are out of range or not valid page names...
		if (tempENTRY.doSpreadsON){
			if (isNaN(tempENTRY.pgStart) || isNaN(tempENTRY.pgEnd) )
				notNumber = true;
			else if (tempENTRY.pgStart > spreadCount || tempENTRY.pgEnd > spreadCount || tempENTRY.pgStart < 0 || tempENTRY.pgEnd < 0 )
				spreadsNoGood = true;
			tempPageCount = tempENTRY.pgEnd - tempENTRY.pgStart +1;
		}
		else if (tempENTRY.nameConvType != 4 && (getPageOffset(tempENTRY.pgStart,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1 || getPageOffset(tempENTRY.pgEnd,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1) )
			pgsNoGood = true;
		else if (tempENTRY.nameConvType == 4 && (tempENTRY.pgStart < 1 || tempENTRY.pgEnd < 1 || tempENTRY.pgStart > tempENTRY.tempENTRY.pageRangeArray.length || tempENTRY.pgEnd > tempENTRY.tempENTRY.pageRangeArray.length) )
			pgsNoGood = true;
		else// Range OK, find out num of pages
			tempPageCount = getPageOffset(tempENTRY.pgEnd,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) - getPageOffset(tempENTRY.pgStart,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) + 1;
	}
	else if (startEndPgs.indexOf(",") > -1){
		// Page range contains a comma. Deal with any "-" at print time (ie 1-3,4,7-11)
		tempENTRY.nonContigPgs = startEndPgs.split(",");
		tempENTRY.hasNonContig = true;

		// Check page range(s) for anything funky
		// and eventually issue error message if not compliant
		for(i = 0; i < tempENTRY.nonContigPgs.length; i++){
			var temp = tempENTRY.nonContigPgs[i];
			if (temp.indexOf("-") != -1){
				// Test if start or end page is not in the document
				var temp2 = temp.split("-");
				var temp3 = temp2[0];
				var temp4 = temp2[1];
				if (tempENTRY.doSpreadsON){
					temp3 = parseInt(temp3,10);
					temp4 = parseInt(temp4,10);
					if (isNaN(temp3) || isNaN(temp4) )
						notNumber = true;
					else	if (temp3 > spreadCount || temp3 < 1 || temp4 > spreadCount || temp4 < 1)
						spreadsNoGood = true;
					tempPageCount += temp4 - temp3;
				}
				else if (tempENTRY.nameConvType == 4 && (temp3 < 1 || temp4 < 1 || temp3 > tempENTRY.pageRangeArray.length || temp4 > tempENTRY.pageRangeArray.length) )
					pgsNoGood = true;
				else if (tempENTRY.nameConvType != 4 && getPageOffset(temp3,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1 || getPageOffset(temp4,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1)
					pgsNoGood = true;
				else// Range OK, find out num of pages for prog bar
					tempPageCount += getPageOffset(temp4,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) - getPageOffset(temp3,tempENTRY.pageNameArray,tempENTRY.pageRangeArray);
			}
			// Test if a single page (was between a ",") is out of range or isn't a number
			else if (tempENTRY.doSpreadsON && isNaN(parseInt(temp,10) ) )
				notNumber = true;
			else if ((tempENTRY.doSpreadsON && parseInt(temp,10) > spreadCount) || (tempENTRY.doSpreadsON && parseInt(temp,10) < 0) )
				spreadsNoGood = true;
			else if (tempENTRY.nameConvType == 4 && (parseInt(temp,10) < 1 || parseInt(temp,10) > tempENTRY.pageRangeArray.length ) )
				pgsNoGood = true;
			else if (tempENTRY.nameConvType != 4 && getPageOffset(temp,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1)
				pgsNoGood = true;
			tempPageCount++; // Add either a single page between commas or add the extra from "tempPageCount += (temp4 - temp3)" above
		}
	}
	else{ // Page range is single page (ie 5)
		tempENTRY.pgStart = startEndPgs;
		tempENTRY.pgEnd = startEndPgs;
		if (tempENTRY.doSpreadsON){
			if (isNaN(tempENTRY.pgStart) )
				notNumber = true;
			else if (tempENTRY.pgStart > spreadCount || tempENTRY.pgStart < 0)
				spreadsNoGood = true;
		}
		else if (tempENTRY.nameConvType != 4 && getPageOffset(tempENTRY.pgStart,tempENTRY.pageNameArray,tempENTRY.pageRangeArray) == -1)
			pgsNoGood = true;
		else if (tempENTRY.nameConvType == 4 && (tempENTRY.pgStart < 1 || tempENTRY.pgStart > tempENTRY.tempENTRY.pageRangeArray.length) )
			pgsNoGood = true;
		else// Range OK, find out num of pages
			tempPageCount = 1;
	}

	// Check if any problems exist, alert user and reset value as needed
	if (spreadsNoGood||notNumber||pgsNoGood){
		tempENTRY.getOut = false;
		peuINFO.editStartEndPgs = "All";
		tempPageCount = 0; // Ranges are no good, don't add yet
		if (spreadsNoGood)
			alert("You have entered an invalid spread range!\nSpreads must be greater than zero and less than\nor equal to the number of spreads in the document (2.1).");
		else if (notNumber)
			alert("You have entered a non-numeric spread number. Please use numerals and try again (2.2).");
		else if (pgsNoGood)
			alert("You have entered a page name that doesn't exist in this document.\nBe sure to include section names if they are being used.\nPlease check your entered page name(s) and try again (2.3).");
	}
	return [tempENTRY, tempPageCount];
}

/*******************************************/
/* Initialization and preference functions */
/*******************************************/

/*
 * Initiallization routine
 */
function peuInit(){
	// Default Preferences, init peuINFO "dictionary"
	peuINFO["outDirNameArray"] = ["Specific Directory"];
	peuINFO["dirListArray"] = ["Specific Directory"]
	peuINFO["startingDirectory"] = "/c/";
	peuINFO["defaultDir"] = 0;
	peuINFO["addZeroON"] = 1;
	peuINFO["addPon"] = 1;
	peuINFO["pageNamePlacement"] = 0;
	peuINFO["doSpreadsON"] = 0;
	peuINFO["seperatorChar"] = 0;
	peuINFO["specialReplaceChar"] = 1;
	peuINFO["doOneFile"] = 0;
	peuINFO["exportDefaultType"] = 0;
	peuINFO["defaultPrintPreset"] = 0;
	peuINFO["measurementUnits"] = 0;
	peuINFO["addPSextention"] = 1;
	peuINFO["defaultJPEGquality"] = 0;
	peuINFO["defaultJPEGrender"] = 0;
	peuINFO["sayCancel"] = 1;
	peuINFO["subType"] = 1;
	peuINFO["layersON"]  = 0;
	peuINFO["layerBeforeON"] = 0;
	peuINFO["batchON"] = 0;
	peuINFO["origSpread"] = false;
	peuINFO["nameConvType"] = 0;
	peuINFO["defaultPDFpreset"] = 0;
	peuINFO["baseLaersAsVersion"] = 0;
	peuINFO["batchSameForAll"] = 0;
	peuINFO["zeroPadding"] = 0;

	// EPS options
	peuINFO["ignoreON"] = 1;
	peuINFO["bleed"]= 0.125;
	peuINFO["epsON"] = 0;
	peuINFO["pdfON"] = 0;
	peuINFO["bitmapON"] = 0;
	peuINFO["psLevel"] = 1;
	peuINFO["colorType"] = 3;
	peuINFO["fontEmbed"] = 1;
	peuINFO["dataSent"] = 0;
	peuINFO["dataType"] = 1;
	peuINFO["opiReplacement"] = 0;
	peuINFO["defaultPreview"] = 1;
	peuINFO["defaultFlattenerPreset"] = 0;
	peuINFO["previewTypes"] = new Array();
	peuINFO.previewTypes[0] = "None";
	peuINFO.previewTypes[1] = "Tiff";
	// Only add pict if we are on a Mac
	if (File.fs != "Windows")
		peuINFO.previewTypes[2] = "Pict";

	// Check if preferences file exists, if not save one
	peuINFO["prefsFile"] = File((Folder(app.activeScript)).parent + "/PageExporter5Prefs.txt");
	if(!peuINFO.prefsFile.exists)
		savePrefs();
	else
		readPrefs();

	peuINFO["csFlattenerPresets"] = app.flattenerPresets;
	peuINFO["flattenerNames"] = new Array();
	for (i = 0; i < peuINFO.csFlattenerPresets.length; i ++)
		peuINFO.flattenerNames.push(peuINFO.csFlattenerPresets[i].name);
	// Check if the default flattener preset is beyond the actual number of presets.
	// This can occur if the default was the last in the list and was deleted
	if(peuINFO.defaultFlattenerPreset > peuINFO.flattenerNames.length-1)
			peuINFO.defaultFlattenerPreset = 0;

	// Test for 'ActivePageItemRunTime' plugin for progress bar:
	try{
		app.documents[0].labeledPageItems;
		peuINFO["doProgressBar"] = true;
	}
	catch(e){
		peuINFO["doProgressBar"] = false;
	}

	// Various Globals
	peuINFO["editStartEndPgs"] = "All";
	peuINFO["charList"] = [".","_"," ",""];
	peuINFO["measureUnitArray"] = [" in"," cm"," mm"," pt"," p"," c"];
	peuINFO["measureLableArray"] = ["inches","centimeters","millimeters","points","picas","ciceros"];

	// Load and determine which print presets are "PostScript File" for the printer
	// Need to do this before building main dialog to check for PSfile "printers"
	var csPrintPresets = app.printerPresets;
	peuINFO["csPSprinters"] = new Array();
	peuINFO["psPrinterNames"] = new Array();
	for(i = 0; i < csPrintPresets.length; i ++)
		if(csPrintPresets[i].printer == Printer.postscriptFile){
			peuINFO.psPrinterNames.push(csPrintPresets[i].name);
			peuINFO.csPSprinters.push(csPrintPresets[i]);
		}
	// If a print preset was deleted and the default was last in the list,
	// problems resulted. This fixes that
	if(peuINFO.defaultPrintPreset > peuINFO.psPrinterNames.length-1)
		peuINFO.defaultPrintPreset = 0;
	peuINFO["adjustForNoPS"] = (peuINFO.psPrinterNames.length == 0)? 1:0; // Adjusts array offset if there are no PSfile printers

	return peuINFO;
}

/*
 * Read the preferences from a text file
 */
function readPrefs(){
	peuINFO.prefsFile.open("r");
	peuINFO.outDirNameArray = (peuINFO.prefsFile.readln() ).split(",");
	peuINFO.dirListArray = (peuINFO.prefsFile.readln() ).split(",");
	peuINFO.startingDirectory = peuINFO.prefsFile.readln();
	peuINFO.defaultDir = Number(peuINFO.prefsFile.readln() );
	peuINFO.addZeroON = Number(peuINFO.prefsFile.readln() );
	peuINFO.addPon = Number(peuINFO.prefsFile.readln() );
	peuINFO.pageNamePlacement = Number(peuINFO.prefsFile.readln() );
	peuINFO.doSpreadsON = Number(peuINFO.prefsFile.readln() );
	peuINFO.seperatorChar = Number(peuINFO.prefsFile.readln() );
	peuINFO.specialReplaceChar = Number(peuINFO.prefsFile.readln() );
	peuINFO.doOneFile = Number(peuINFO.prefsFile.readln() );
	peuINFO.exportDefaultType = Number((peuINFO.prefsFile.readln() ));
	peuINFO.defaultPrintPreset = Number(peuINFO.prefsFile.readln() );
	peuINFO.measurementUnits = Number(peuINFO.prefsFile.readln() );
	peuINFO.addPSextention = Number(peuINFO.prefsFile.readln() );
	peuINFO.defaultJPEGquality = Number(peuINFO.prefsFile.readln() );
	peuINFO.defaultJPEGrender = Number(peuINFO.prefsFile.readln() );
	peuINFO.sayCancel = Number(peuINFO.prefsFile.readln() );
	peuINFO.subType = Number(peuINFO.prefsFile.readln() );
	peuINFO.layersON = Number(peuINFO.prefsFile.readln() );
	peuINFO.layerBeforeON = Number(peuINFO.prefsFile.readln() );
	peuINFO.batchON = Number(peuINFO.prefsFile.readln() );
	peuINFO.ignoreON = Number(peuINFO.prefsFile.readln() );
	peuINFO.bleed = Number(peuINFO.prefsFile.readln() );
	peuINFO.epsON = Number(peuINFO.prefsFile.readln() );
	peuINFO.pdfON = Number(peuINFO.prefsFile.readln() );
	peuINFO.bitmapON = Number(peuINFO.prefsFile.readln() );
	peuINFO.psLevel = Number(peuINFO.prefsFile.readln() );
	peuINFO.colorType = Number(peuINFO.prefsFile.readln() );
	peuINFO.fontEmbed = Number(peuINFO.prefsFile.readln() );
	peuINFO.dataSent = Number(peuINFO.prefsFile.readln() );
	peuINFO.dataType = Number(peuINFO.prefsFile.readln() );
	peuINFO.opiReplacement = Number(peuINFO.prefsFile.readln() );
	peuINFO.defaultPreview = Number(peuINFO.prefsFile.readln() );
	peuINFO.defaultFlattenerPreset = Number(peuINFO.prefsFile.readln() );
	peuINFO.defaultPDFpreset =  Number(peuINFO.prefsFile.readln() );
	peuINFO.baseLaersAsVersion = Number(peuINFO.prefsFile.readln() );
	peuINFO.batchSameForAll = Number(peuINFO.prefsFile.readln() );
	peuINFO.zeroPadding = Number(peuINFO.prefsFile.readln() );
	peuINFO.prefsFile.close();
}

/*
 * Save the preferences to a text file
 */
function savePrefs(){
	var newPrefs =
	peuINFO.outDirNameArray + "\n" +
	peuINFO.dirListArray + "\n" +
	peuINFO.startingDirectory  + "\n" +
	peuINFO.defaultDir + "\n" +
	peuINFO.addZeroON + "\n" +
	peuINFO.addPon + "\n" +
	peuINFO.pageNamePlacement + "\n" +
	peuINFO.doSpreadsON + "\n" +
	peuINFO.seperatorChar + "\n" +
	peuINFO.specialReplaceChar + "\n" +
	peuINFO.doOneFile + "\n" +
	peuINFO.exportDefaultType + "\n" +
	peuINFO.defaultPrintPreset + "\n" +
	peuINFO.measurementUnits + "\n" +
	peuINFO.addPSextention + "\n" +
	peuINFO.defaultJPEGquality + "\n" +
	peuINFO.defaultJPEGrender + "\n" +
	peuINFO.sayCancel + "\n" +
	peuINFO.subType + "\n" +
	peuINFO.layersON + "\n" +
	peuINFO.layerBeforeON + "\n" +
	peuINFO.batchON + "\n" +
	peuINFO.ignoreON + "\n" +
	peuINFO.bleed + "\n" +
	peuINFO.epsON + "\n" +
	peuINFO.pdfON + "\n" +
	peuINFO.bitmapON + "\n" +
	peuINFO.psLevel + "\n" +
	peuINFO.colorType + "\n" +
	peuINFO.fontEmbed + "\n" +
	peuINFO.dataSent + "\n" +
	peuINFO.dataType + "\n" +
	peuINFO.opiReplacement + "\n" +
	peuINFO.defaultPreview + "\n" +
	peuINFO.defaultFlattenerPreset + "\n" +
	peuINFO.defaultPDFpreset + "\n" +
	peuINFO.baseLaersAsVersion + "\n" +
	peuINFO.batchSameForAll + "\n" +
	peuINFO.zeroPadding + "\n" +
	"0\n0\n0\n0\n0\n0\n0\n0\n0\n0\n";// Add (10) for extra prefs in future
	peuINFO.prefsFile.open("w");
	peuINFO.prefsFile.write(newPrefs);
	peuINFO.prefsFile.close();
}

/*
 * Change Preferences
 */
function changePrefs(){
	var popChoice = 0;
	var prefItems = makePrefsDlog();

	if(prefItems.otherPrefsDlog.show()){
		// Add a default directories
		if(prefItems.addDefDir.checkedState)
			addDirectories();
		// Delete directories
		if(prefItems.delDefDir.checkedState)
			deleteDirectories();
		// Change Start Directory
		if(prefItems.changeStartDirectory.checkedState){
			tempStartDirectory = getDirectory("Please select the new start directory:",peuINFO.startingDirectory);
			if(tempStartDirectory != null)
				peuINFO.startingDirectory = tempStartDirectory;
		}
		// The other prefs = the index of popups selected
		peuINFO.seperatorChar = prefItems.changeSepToken.selectedIndex;
		peuINFO.specialReplaceChar = prefItems.changeRepToken.selectedIndex;
		peuINFO.measurementUnits = prefItems.changeUnits.selectedIndex;
		peuINFO.addZeroON = prefItems.changeAddZero.selectedIndex;
		peuINFO.addPon = prefItems.changeAddP.selectedIndex;
		peuINFO.pageNamePlacement = prefItems.changeBeforeON.selectedIndex;
		peuINFO.layerBeforeON = prefItems.changeLayerON.selectedIndex;
		peuINFO.sayCancel = prefItems.changeShowCancel.selectedIndex;
		peuINFO.subType = prefItems.changeSubType.selectedIndex;
		peuINFO.baseLaersAsVersion = (prefItems.changeBaseExport.checkedState)?1:0;
		peuINFO.zeroPadding = prefItems.changePadding.selectedIndex;
		prefItems.otherPrefsDlog.destroy();
		savePrefs()
	}
}

/*
 * Change Preferences Helper: Build Prefs Dialog Box
 */
function makePrefsDlog(){
	var prefsItems = new Array();
	prefsItems["otherPrefsDlog"] = app.dialogs.add({name:"Change Preferences", canCancel:true} );
	with (prefsItems.otherPrefsDlog){
		// Left column
		with (dialogColumns.add() ){
			with (dialogRows.add() )
				staticTexts.add({staticLabel:"User Defined Preferences"} );
			with(borderPanels.add() ){
				with(dialogColumns.add() ){
					with (dialogRows.add() )
						prefsItems["addDefDir"] = checkboxControls.add({staticLabel:"Add an Output Directory", checkedState:false});
					with (dialogRows.add() )
						prefsItems["delDefDir"] = checkboxControls.add({staticLabel:"Delete an Output Directory", checkedState:false});
					with (dialogRows.add() )
						prefsItems["changeStartDirectory"] = checkboxControls.add({staticLabel:"Change Starting Directory", checkedState:false});
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Separator Character"} );
						prefsItems["changeSepToken"] = dropdowns.add({stringList:[".","_","<space>","nothing"], selectedIndex:peuINFO.seperatorChar});
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Replacement Character"} );
						prefsItems["changeRepToken"] = dropdowns.add({stringList:[".","_","<space>","nothing"], selectedIndex:peuINFO.specialReplaceChar} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Measurement Units"} );
						prefsItems["changeUnits"] = dropdowns.add({stringList:peuINFO.measureLableArray, selectedIndex:peuINFO.measurementUnits} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Add Leading Zero(s)"} );
						prefsItems["changeAddZero"] = dropdowns.add({stringList:["No","Yes"], selectedIndex:peuINFO.addZeroON} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Force Page Numbers to"} );
						prefsItems["changePadding"] = dropdowns.add({stringList:["Normal","2","3"], selectedIndex:peuINFO.zeroPadding} );
						staticTexts.add({staticLabel:"Digits when Adding Leading Zeros"} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Add \"P\" to Numeric Pages"} );
						prefsItems["changeAddP"] = dropdowns.add({stringList:["No","Yes"], selectedIndex:peuINFO.addPon} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Add Page Name"} );
						prefsItems["changeBeforeON"] = dropdowns.add({stringList:["After","Before","by replacing \"<#>\" in"], selectedIndex:peuINFO.pageNamePlacement} );
						staticTexts.add({staticLabel:"Base Name"} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Add Layer Name"} );
						prefsItems["changeLayerON"] = dropdowns.add({stringList:["After","Before"], selectedIndex:peuINFO.layerBeforeON} );
						staticTexts.add({staticLabel:"Page Name"} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Alert on Cancel"} );
						prefsItems["changeShowCancel"] = dropdowns.add({stringList:["No","Yes"], selectedIndex:peuINFO.sayCancel} );
					}
					with (dialogRows.add() ){
						staticTexts.add({staticLabel:"Lot Naming Type"} );
						prefsItems["changeSubType"] = dropdowns.add({stringList:["Consecutive","Redundant"], selectedIndex:peuINFO.subType} );
					}
					with (dialogRows.add() )
						prefsItems["changeBaseExport"] = checkboxControls.add({staticLabel:"Export Base Layer(s) as an Individual Version", checkedState:peuINFO.baseLaersAsVersion, minWidth:140});
				}
			}
		}
	}
	return prefsItems;
}

/*
 * Change Preferences Helper: Add one or more output directories
 */
function addDirectories(){
	var prefGetOut1 = true;
	do{
		tempDir = getDirectory("Select a new directory to add:",peuINFO.startingDirectory);
		if(tempDir != null){
			peuINFO.dirListArray.push(tempDir + "/");
			do{
				var prefGetOut = true;
				var tempDlog = app.dialogs.add({name:"Output Directory Name", canCancel:false} );
				with (tempDlog)
					with (dialogColumns.add() )
						with (dialogRows.add() ){
							staticTexts.add({staticLabel:"Enter the name to use for this new directory:"} );
							var tempChange = textEditboxes.add({editContents:"newName", minWidth:170} );
						}
				if(tempDlog.show()){
					tempChange = tempChange.editContents;
					if(tempChange.length < 1){
						prefGetOut = false;
						alert("The name must have at least one character!");
					}
					else
						peuINFO.outDirNameArray.push(tempChange);
				}
			} while(!prefGetOut);
		}
		if(prefGetOut1){
			tempDlog = app.dialogs.add({name:"", canCancel:true} );
			with (tempDlog)
				with (dialogColumns.add() )
					with (dialogRows.add() )
						staticTexts.add({staticLabel:"Do you want to add another default directory? (OK=yes, Cancel=no)"} );
			prefGetOut1 = tempDlog.show();
		}
	} while(prefGetOut1);
	tempDlog.destroy();
}

/*
 * Change Preferences Helper: Delete directories from the output directories list
 */
function deleteDirectories(){
	do{
		var menuArray = new Array();// Needed to rig the menu list. ID has weird bug that keeps deleted items in an array
		for(place = 0;place < peuINFO.outDirNameArray.length;place++)
			menuArray[place] = peuINFO.outDirNameArray[place];
		var outDirNameArrayCnt = peuINFO.outDirNameArray.length;
		var tempDlog = app.dialogs.add({name:"Delete an Output Directory", canCancel:true} );
		with (tempDlog)
			with (dialogColumns.add() )
				with (dialogRows.add() )
					var killIt = dropdowns.add({stringList:menuArray, minWidth:204, selectedIndex: 0 } );
		if(tempDlog.show()){
			killIt = killIt.selectedIndex;
			if(killIt == peuINFO.defaultDir)
				peuINFO.defaultDir--;
			if(killIt == 0)
				alert("The \"Specific Directory\" entry can't be deleted.");
			else{
					peuINFO.outDirNameArray.splice(killIt,1);
					peuINFO.dirListArray.splice(killIt, 1);
			}
		}
		tempDlog.destroy();
		tempDlog = app.dialogs.add({name:"", canCancel:true} );
		with (tempDlog)
			with (dialogColumns.add() )
				with (dialogRows.add() )
					staticTexts.add({staticLabel:"Do you want to delete another default directory? (OK=yes, Cancel=no)"} );
		var prefGetOut = tempDlog.show();
	} while(prefGetOut);
}

