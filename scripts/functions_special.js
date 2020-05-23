﻿// Tuoreimman JSON2-paketin voi ladata osoitteesta https://github.com/douglascrockford/JSON-js/blob/master/json2.js
#include "json2.js"
#include "functions_basic.js"

function createTable(columns, options, partialText, wholeText){
	// Read JSON data
    var data = arrangeData(importData());
    var picName = importPicNames();
    var picSettings = importPicSettings();

	// Print missing pic names and settings in the console
    printMissingInformation();
	
	// set values for variables
    var columns = columns || 5;
    var topMargin      = options["topMargin"]      || 0;
    var globalScale    = options["scale"]          || 70;
    var globalY        = options["shiftY"]         || 0;
    var globalX        = options["shiftX"]         || 0;
	var insertPictures = options["insertPictures"] || true;
	var insertBacks    = options["insertBacks"]    || true;
	
    var partialText = partialText || {};
    var wholeText = wholeText || {};
	
	// Check that selection contains all the needed information
    var docRef = app.activeDocument;
    var selectedObjects = app.activeDocument.selection; // If this line gives an error, make sure the target is Illustrator (instead of ExtendScript Toolkit)
    if (selectedObjects.length != 4){throw new Error("Select exactly four elements in Illustrator.")};
	
	// Let's give nicer names to the selected objects
	var prototypeTextbox        = selectedObjects[0];
	var prototypeBack           = selectedObjects[1];
	var prototypeSectionTextbox = selectedObjects[2];
	var prototypePictureBox     = selectedObjects[3];
	// MUUTA JÄRKKÄ: back - picBox - name/title - sectionBox
	
	if (prototypeTextbox.textRange.lines.length != 2){ throw new Error("Make sure that the textbox for name and title has exactly two rows."); }
	
    var leftPadding = prototypePictureBox.position[0] - prototypeBack.position[0];
	
	// overlap between backs
    var cellOverlap = (prototypeBack.width - prototypePictureBox.width) / 2;
	var matrixCellDimensions = [prototypeBack.width  - 0.01 - cellOverlap,
								prototypeBack.height - 0.01        ];
	 // the little number lets the backs to overlap slightly at section breaks.
	 // it helps if you join backs with pathfinder
	 
	var matrixWidth = matrixCellDimensions[0] * columns - cellOverlap;
	
	// Vector from the up-left corner of the prototype back to the up-left corner of the whole matrix,
	// i.e., maps points from prototype cell to the first cell
    var prototypeToFirstCell = [-prototypeBack.position[0] + app.activeDocument.artboards[0].artboardRect[0] + (app.activeDocument.width - matrixWidth)/2,
					            -prototypeBack.position[1] + app.activeDocument.artboards[0].artboardRect[1] - topMargin];
	
	// create group for section names
    var grpSectionTextbox = activeDocument.groupItems.add();
    grpSectionTextbox.name = "Jaoksien nimet";
	
	// running index for cells
    var tileNo = 0;

    for (var section in data) {
    
        // Initialize new section
        var sectionTextbox = prototypeSectionTextbox.duplicate(grpSectionTextbox);
        sectionTextbox.contents = replaceText(section, wholeText, partialText);
        //var pos = sectionTextbox.position;
    
		// Offset of section back and text
        if (col(tileNo, columns) == 0){
			// Current cell is at the left edge
			var offsetText = (leftPadding - cellOverlap) / 2;
			var offsetBack = 0;
		}else{
			// Current cell is not at the left edge
            var offsetText = leftPadding / 2;
            var offsetBack = cellOverlap;
        }
        
        if (col(tileNo, columns) == columns-1){alert("A section title is placed on the right border. It won't look good, but I'll finish the job. Try changing the order of sections or vary the number of columns.")}
 
        //pathRef.position = 
		positionByTileNo(sectionTextbox, prototypeToFirstCell, matrixCellDimensions, tileNo, columns, offsetText);
        
        // Initialize new groups for this current section
		var grpBack    = activeDocument.groupItems.add();     grpBack.name = section.replaceAll("-\n","") + "-backs";
        var grpTextbox = activeDocument.groupItems.add();  grpTextbox.name = section.replaceAll("-\n","") + "-textboxes";
        var grpPics    = activeDocument.groupItems.add();     grpPics.name = section.replaceAll("-\n","") + "-pictures";
        //var grpWH      = activeDocument.groupItems.add();
        
        // Put a back for the section. Notice the little positive offset in the position
		if (insertBacks){
			positionByTileNo(prototypeBack.duplicate(grpBack), prototypeToFirstCell, matrixCellDimensions, tileNo, columns, offsetBack);
		}
   
        tileNo++;
        // Loop for the names in the section
        for (var nimi in data[section]){
            // Title
			var pathRef = positionByTileNo(prototypeTextbox.duplicate(grpTextbox), prototypeToFirstCell, matrixCellDimensions, tileNo, columns, 0);
            var nameLine  = pathRef.textRange.lines[0]
			var titleLine = pathRef.textRange.lines[1]
            // Remove hyphenation from names, just in case
            nameLine.paragraphAttributes.hyphenation = false;
            nameLine.contents  = replaceText(nimi,                wholeText, partialText);
            titleLine.contents = replaceText(data[section][nimi], wholeText, partialText);
            
			if (insertBacks){
				positionByTileNo(prototypeBack.duplicate(grpBack), prototypeToFirstCell, matrixCellDimensions, tileNo, columns, 0);
			}
			
            // Insert picture, if it's found and it has settings
            if (insertPictures && picName[nimi] && picSettings[picName[nimi]]){
                var picFile = picName[nimi];
				
                // Copy picture frame
				var picBox = prototypePictureBox.duplicate();
        
                // Load picture to a correct position
                var pic = insertPicture2(picDirectory(), picFile, '.jpg', picSettings[picFile])
        
                // Apply global scaling and translation
                objectScale(pic, globalScale, [100/2, 150/2]);
                pic.position = [pic.position[0] + globalX,
                                pic.position[1] + globalY]
        
                // Place the pic to the prototype position
                pic.position = [pic.position[0] - 100/2 + picBox.position[0] + picBox.width /2, 
                                pic.position[1] + 150/2 + picBox.position[1] - picBox.height/2];
        
                // Crop the picture with the frame
                var clipgroup = app.activeDocument.groupItems.add(grpPics); // meneekö grpPicsiin?

                pic.moveToBeginning(clipgroup);
				picBox.moveToBeginning(clipgroup);
                picBox.clipping = true;
                clipgroup.clipped = true;
				
                // Move the clipgroup to the final destination
				positionByTileNo(clipgroup, prototypeToFirstCell, matrixCellDimensions, tileNo, columns, 0)
            }
            // Moving to the next cell
            tileNo++;
        }
        // Remove empty groups
        removeIfEmptyGroup(grpBack); removeIfEmptyGroup(grpTextbox); removeIfEmptyGroup(grpPics);
    }
    // Bring Section name -group to the top
    grpSectionTextbox.zOrder(ZOrderMethod.BRINGTOFRONT)
    return "Success"
}

// moves the obj from prototype position to a correct matrix cell
function positionByTileNo(obj, prototypeToFirstCell, matrixCellDimensions, tileNo, columns, offset){ // 'offset' is for section cell offset
	obj.position = new Array(obj.position[0] + prototypeToFirstCell[0] + matrixCellDimensions[0] * col(tileNo, columns) + offset,
							 obj.position[1] + prototypeToFirstCell[1] - matrixCellDimensions[1] * row(tileNo, columns));
	return obj;
}

// These functions convert running tile number to row and column indices
function row(tileNo, columns){ return Math.floor((tileNo) / columns);};
function col(tileNo, columns){ return tileNo % columns; };

function saveWebJPGs(imageSize, options, quality){ // quality in [0,100]
	var globalScale = options["scale"] || 70;
    var globalY = options["shiftY"] || 0;
    var globalX = options["shiftX"] || 0;
	var quality = quality || 90;
	
    var doc = app.documents.add(); // create a doc with defaults. If error, target Illustrator
    var firstArtBoard = doc.artboards[0]; // get the default atboard
    var x1 = 0; var y1 = 0;
    var x2 = imageSize[0];
    var y2 = -imageSize[1];
    doc.artboards.add([x1, y1, x2, y2]); // adds new artboard
    firstArtBoard.remove(); // remove the default artboard
	
	// Export options
    var type = ExportType.JPEG;
    var options = new ExportOptionsJPEG();
    options.antiAliasing = true;
    options.artBoardClipping = true;
    options.optimization = true;
    options.qualitySetting = quality; // choose a value from 0 to 100
    //options.horizontalScale = 100; // percentage, initially 100. do we need to change these?
    //options.verticalScale = 100;
    
    var data = arrangeData(importData());
	
	// create a array containing already exported persons
    var nimet = [];
    
    for (var section in data){
        for (var nimi in data[section]){
            if (!isMember(nimi, nimet)){				
                nimet.push(nimi); // adds name into the array
                var pic = cropWithName(nimi); // set the pic onto the canvas
                objectScale(pic, globalScale, [100/2, 150/2]); // scaling with globalScale
                pic.position = [pic.position[0] - 100/2 - imageSize[0]/2 + globalX, // translation to center of canvas + global translation
                                pic.position[1] - 150/2 - imageSize[1]/2 + globalY];
				
				// retain the original file name in the exported file
                var pathSplit = pic.file.toString().split("/");
                var fileName = pathSplit[pathSplit.length -1].split(".")[0];
                fileName = decodeURIComponent(fileName); // Converts e.g. "%20" to " "
        
                var destinationFile = File(webPicDirectory() + fileName + '.jpg'); // initialize a new file
                activeDocument.exportFile(destinationFile, type, options); // export the jpg
                pic.remove(); // remove pic from canvas
            }
        }
    }
    return "The pictures were exported successfully!"
};

// FUNCTIONS WORKING WITH DATA

function arrangeData(dataJSON){
    var taulu = {};
    
    // Alphabetize the arrays containing lists of names
    for (var jaos in dataJSON){
        for (var virka in dataJSON[jaos]){
            dataJSON[jaos][virka] = alphabetizeByLastName(dataJSON[jaos][virka]);
        }
    }
    
    // Loop for sections
    for (var jaos in dataJSON){
        
        // Initialize an array for section's names and JSON
        var nimet = [];
        taulu[jaos] = {};
        for (var virka in dataJSON[jaos])
            for (var nimi = 0; nimi < dataJSON[jaos][virka].length; nimi++){
                // Here we have all the names in the order of the database
                
                // If the name is already processed, go on to the next name
                if (!isMember(dataJSON[jaos][virka][nimi], nimet)){
                    nimet.push(dataJSON[jaos][virka][nimi]);
                    var virat = [];
                    
                    // Gather all titles for the person, in the section.
                    for (var virka_2 in dataJSON[jaos]){
                        for (var nimi_2 = 0; nimi_2 < dataJSON[jaos][virka_2].length; nimi_2++){
                            if (dataJSON[jaos][virka_2][nimi_2] == dataJSON[jaos][virka][nimi]){ virat.push(virka_2); }
                    }
                }
                // Save person to the JSON along with its titles
                taulu[jaos][dataJSON[jaos][virka][nimi]] = virat.join(", ");
            }
        }
    }
    return taulu;
}

function printMissingInformation(){ // Works only with arranged data. Prints missing pics and pic settings in the console
    var data = arrangeData(importData());
    var picName = importPicNames();
    var picSettings = importPicSettings();
    
    var puuttuu_kuva = [];
    var puuttuu_rajaus = [];

    for (var section in data){
        for (var nimi in data[section]){
            if (!picName[nimi]){
                if (!isMember(nimi, puuttuu_kuva)){ puuttuu_kuva.push(nimi); }
            }else if (!picSettings[picName[nimi]]){
                if (!isMember(picName[nimi], puuttuu_rajaus)){ puuttuu_rajaus.push(picName[nimi]); }
            }
        }
    }
    
    if (puuttuu_kuva.length != 0)
        $.writeln("These persons are missing a picture file name:\n\t" + puuttuu_kuva.join("\n\t"))
    if (puuttuu_rajaus.length != 0)
        $.writeln("Following pictures don't have settings for cropping:\n\t" + puuttuu_rajaus.join("\n\t"))
        
    if (puuttuu_kuva.length == 0 && puuttuu_rajaus.length == 0){
        $.writeln("All picture file names and picture setting are found!")
    }
}

// FUNCTIONS RELATED TO INSERTING AND CROPPING OF PICTURES
////////////////////////////////////////////////////////////

function cropWithName(nimi){
    // Load JSON data
    var picName = importPicNames();
    var picSettings = importPicSettings();

    if (picName[nimi] && picSettings[picName[nimi]]){
        var pic = insertPicture2(currentDirectory() + '../pictures/', picName[nimi], '.jpg', picSettings[picName[nimi]]);
        return pic;
    }else if (picName[nimi]){
        var pic = insertPicture2(currentDirectory() + '../pictures/', picName[nimi], '.jpg', [1, 0, 0, 1, 0, 0]);
        return pic;
    }else{
        throw new Error("The person does not have a picture file name. Add it to data_picFileName.js.");
    }
}

function readCropSettings(){
    var artboardRect = app.activeDocument.artboards[0].artboardRect;
    if (artboardRect[0] != 0 || artboardRect[1] != 0 || artboardRect[2] != 100 || artboardRect[3] != -150){ throw new Error("The artboard in this document has wrong dimensions and/or is at wrong position. Create a working document by running script_crop_createDoc.jsx."); }

    pic1 = app.activeDocument.selection[0];
    if (!pic1){ throw new Error("Select the picture you want to crop.")}
    mat = pic1.matrix;
    array = [mat.mValueA,
             mat.mValueB,
             mat.mValueC,
             mat.mValueD,
             //mat.mValueTX,
             //mat.mValueTY,
             pic1.position[0],
             pic1.position[1]];
    var pathSplit = pic1.file.toString().split("/");
    var fileName = pathSplit[pathSplit.length -1].split(".")[0];
    fileName = decodeURIComponent(fileName); // Converts e.g. "%20" to " "
    var picSettings = JSON.parse(readFile(currentDirectory() + "data_picSettings.js"));
    picSettings[fileName] = array;
    writeFile(currentDirectory() + "data_picSettings.js", JSON.stringify(picSettings,null,1));
    return "Settings for '" + fileName + "' were saved successfully.";
}

function insertPicture2(path, name, suffix, matrixArray){
    var itemToPlace = app.activeDocument.placedItems.add(); // If error, chech Target illustrator
    filepath = path + name + suffix;
    $.writeln(filepath);
    itemToPlace.file = new File(filepath);
    itemToPlace.matrix.mValueA = matrixArray[0];
    itemToPlace.matrix.mValueB = matrixArray[1];
    itemToPlace.matrix.mValueC = matrixArray[2];
    itemToPlace.matrix.mValueD = matrixArray[3];
    //itemToPlace.matrix.mValueTX = matrixArray[4];
    //itemToPlace.matrix.mValueTY = matrixArray[5];
    itemToPlace.position = [matrixArray[4], matrixArray[5]];
    return itemToPlace;
}

function createPicFrameDocument() {
    var doc = app.documents.add(); // create a doc with defaults. If error, target Illustrator
    var firstArtBoard = doc.artboards[0]; // get the default atboard
    var x1 = 0; var y1 = 0; var x2 = 100; var y2 = -150;
    doc.artboards.add([x1, y1, x2, y2]);
    firstArtBoard.remove();
};

// FUNCTIONS FOR DATA IMPORT AND DIRECTORIES

function picDirectory() { return currentDirectory() + "../pictures/"; };

function webPicDirectory()  { // Returns the web picture directory. Creates the directory, if it doesn't exist.
    var f = new Folder(currentDirectory() + "../webPictures/");
    if (!f.exists)
        f.create();
    return currentDirectory() + "../webPictures/";
}

// JSON parsing functions
function importData()       { return JSON.parse(readFile(currentDirectory() + "data_structure.js")); };
function importPicNames()   { return JSON.parse(readFile(currentDirectory() + "data_picFileNames.js")); };
function importPicSettings(){ return JSON.parse(readFile(currentDirectory() + "data_picSettings.js")); };