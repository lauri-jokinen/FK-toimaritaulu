// Tuoreimman JSON2-paketin voi ladata osoitteesta https://github.com/douglascrockford/JSON-js/blob/master/json2.js
#include "json2.js"
#include "functions_basic.js"

function createTable(columns, composition, partialText, wholeText){
    var data = arrangeData(importData());
    var picName = importPicNames();
    var picSettings = importPicSettings();

    printMissingInformation();

    var columns = columns || 5;
    var topMargin = composition["topMargin"] || 0;
    var globalScale = composition["scale"] || 70;
    var globalY = composition["shiftY"] || 0;
    var globalX = composition["shiftX"] || 0;
    var partialText = partialText || {};
    var wholeText = wholeText || {};

    var docRef = app.activeDocument;
    var piRef = app.activeDocument.selection; // If this line gives an error, make sure the target is Illustrator (instead of ExtendScript Toolkit)
    if (piRef.length != 4){throw new Error("Select four elements in Illustrator.")};
    var objW = piRef[1].width; // dimensions of the back
    var objH = piRef[1].height;

    var leftPadding = piRef[3].position[0] - piRef[1].position[0];
    var dist = -.5*(piRef[1].width - piRef[3].width);
    
    var paperW = app.activeDocument.width;
    var paperH = app.activeDocument.height;

    var initPos = [(paperW - objW*columns - dist*(columns-1))/2, -topMargin] // Up-left corner of the first back
    var finalPos = [-piRef[1].position[0]+initPos[0], -piRef[1].position[1]+initPos[1]];
    finalPos[0] = finalPos[0] + app.activeDocument.artboards[0].artboardRect[0];
    finalPos[1] = finalPos[1] + app.activeDocument.artboards[0].artboardRect[1];
    var grpSect = activeDocument.groupItems.add();
    grpSect.name = "Jaoksien nimet";

    var tileNo = 0;

    // check box lines
    if (piRef[0].textRange.lines.length != 2){ throw new Error("Make sure that the textbox for name and title has exactly two rows."); }

    for (var section in data) {
    
        // Initialize new section
    
        var pathRef = piRef[2].duplicate(grpSect);
        pathRef.contents = replaceText(section, wholeText, partialText);
        var pos = pathRef.position;
    
        if (tileNo % columns == 0){var offsetText = dist * 0.5 + leftPadding * 0.5; var offsetBack = 0;}else{ // increment does not apply to the first section title
            var offsetText = leftPadding / 2;
            //var offsetBack = padding - leftPadding - dist;
            var offsetBack = -dist
        }
        
        if ((tileNo+1) % columns == 0){alert("A section title is placed on the right border. It won't look good, but I'll finish the job. Try changing the order of sections or vary the number of columns.")}
 
        pathRef.position = positionByTileNo(pathRef, finalPos, objW, objH, dist, tileNo, columns, offsetText)
        
        // Initialize new groups for this current section        
        var grpBack   = activeDocument.groupItems.add();     grpBack.name = section.replaceAll("-\n","") + "-tausta";
        var grpTitles = activeDocument.groupItems.add();   grpTitles.name = section.replaceAll("-\n","") + "-tittelit";
        var grpPics   = activeDocument.groupItems.add();     grpPics.name = section.replaceAll("-\n","") + "-kuvat";
        var grpWH     = activeDocument.groupItems.add();
        
        // Put a back for the section. Notice the little positive offset in the position
        var pathRef = piRef[1].duplicate(grpBack);
        pathRef.position = positionByTileNo(pathRef, finalPos, objW, objH, dist, tileNo, columns, offsetBack);
   
        tileNo++;
        // Loop for the names in the section
        for (var nimi in data[section]){
            // Title
            var pathRef = piRef[0].duplicate(grpTitles);
            pathRef.position = positionByTileNo(pathRef, finalPos, objW, objH, dist, tileNo, columns, 0);
            nameLine  = pathRef.textRange.lines[0]
            // Remove hyphenation from names
            nameLine.paragraphAttributes.hyphenation = false;
            titleLine = pathRef.textRange.lines[1]
            nameLine.contents = replaceText(nimi, wholeText, partialText);
            titleLine.contents = replaceText(data[section][nimi], wholeText, partialText);
            
            // Back
            var pathRef = piRef[1].duplicate(grpBack);
            pathRef.position = positionByTileNo(pathRef, finalPos, objW, objH, dist, tileNo, columns, 0);
        
            // Insert picture, if it's found and it has settings
            if (picName[nimi] && picSettings[picName[nimi]]){
                var picFile = picName[nimi];
            
                // Copy picture frame
                var pathRef = piRef[3].duplicate(grpWH);
                var pos = pathRef.position;
        
                // Load picture to a correct position
                var pic = insertPicture2(picDirectory(), picFile, '.jpg', picSettings[picFile])
        
                // Apply global scaling and translation
                objectScale(pic, globalScale, [100/2, 150/2]);
                pic.position = [pic.position[0] + globalX,
                                pic.position[1] + globalY]
        
                // Put the pic to the initial position
                pic.position = [pic.position[0] - 100/2 + pos[0] + pathRef.width/2, 
                                pic.position[1] + 150/2 + pos[1] - pathRef.height/2];
        
                // Crop the picture with the frame
                var clipgroup = app.activeDocument.groupItems.add();

                pic.moveToBeginning(clipgroup);
                pathRef.clipping = true;
                pathRef.moveToBeginning(clipgroup);
                clipgroup.clipped = true;
        
                // Move the clipgroup to the final destination
                clipgroup.position = positionByTileNo(clipgroup, finalPos, objW, objH, dist, tileNo, columns, 0)
                clipgroup.moveToBeginning(grpPics);
            }
            
            // The block is finished. Move to the next one
            tileNo++;
        }
        // Remove empty groups
        removeIfEmptyGroup(grpBack); removeIfEmptyGroup(grpTitles); removeIfEmptyGroup(grpPics); removeIfEmptyGroup(grpWH);
    }
    // Bring Section name -group to the top
    grpSect.zOrder(ZOrderMethod.BRINGTOFRONT)
    return "Success"
}

function positionByTileNo(obj, finalPos, objW, objH, dist, tileNo, columns, offset){
    var pos = obj.position;
    var epsilon = 0.01;
    return new Array(pos[0]+finalPos[0]+(objW+dist)*((tileNo) % columns) + offset, pos[1]+finalPos[1]-(objH-epsilon)*(Math.floor((tileNo) / columns)));
}

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

function printMissingInformation(){ // Works only with arranged data
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


function importData()       { return JSON.parse(readFile(currentDirectory() + "data_structure.js")); }
function importPicNames()   { return JSON.parse(readFile(currentDirectory() + "data_picFileNames.js")); }
function importPicSettings(){ return JSON.parse(readFile(currentDirectory() + "data_picSettings.js")); }
function picDirectory()     { return currentDirectory() + "../pictures/"; }

function webPicDirectory()  {
    var f = new Folder(currentDirectory() + "../webPictures/");
    if (!f.exists)
        f.create();
    return currentDirectory() + "../webPictures/";
}

function saveWebJPGs(imageSize, composition){
	var globalScale = composition["scale"] || 70;
    var globalY = composition["shiftY"] || 0;
    var globalX = composition["shiftX"] || 0;
	
    var doc = app.documents.add(); // create a doc with defaults. If error, target Illustrator
    //var doc = app.activeDocument;
    var firstArtBoard = doc.artboards[0]; // get the default atboard
    var x1 = 0;
    var y1 = 0;
    var x2 = imageSize[0];
    var y2 = -imageSize[1];
    doc.artboards.add([x1, y1, x2, y2]);
    firstArtBoard.remove();
    
    var data = arrangeData(importData());
    var nimet = [];
    
    for (var section in data){
        for (var nimi in data[section]){
            if (!isMember(nimi, nimet)){
                nimet.push(nimi);
                var pic = cropWithName(nimi);
                objectScale(pic, globalScale, [100/2, 150/2]);
                pic.position = [pic.position[0] + globalX,
                                pic.position[1] + globalY]
            
                // Put the pic to the initial position
                pic.position = [pic.position[0] - 100/2 + 0 + imageSize[0]/2, 
                                pic.position[1] + 150/2 + 0 - imageSize[1]/2];
    
                var type = ExportType.JPEG;
                var options = new ExportOptionsJPEG();
                options.antiAliasing = true;
                options.artBoardClipping = true;
                options.optimization = true;
                options.qualitySetting = 90; // choose a value from 0 to 100
                options.horizontalScale = scaling
                options.verticalScale = scaling
    
                var pathSplit = pic.file.toString().split("/");
                var fileName = pathSplit[pathSplit.length -1].split(".")[0];
                fileName = decodeURIComponent(fileName); // Converts e.g. "%20" to " "
        
                var destinationFile = File(webPicDirectory() + fileName + '.jpg');
                activeDocument.exportFile(destinationFile, type, options);
                pic.remove();
            }
        }
    }
    return "The pictures were exported successfully!"
};

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
             mat.mValueTX,
             mat.mValueTY,
             pic1.position[0],
             pic1.position[1]];
    var pathSplit = pic1.file.toString().split("/");
    var fileName = pathSplit[pathSplit.length -1].split(".")[0];
    fileName = decodeURIComponent(fileName); // Converts e.g. "%20" to " "
    var picSettings = JSON.parse(readFile(currentDirectory() + "data_picSettings.js"));
    picSettings[fileName] = array;
    writeFile(currentDirectory() + "data_picSettings.js", JSON.stringify(picSettings,null,1));
    return "Settings for '" + fileName + "' were saved successfully."
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
    itemToPlace.matrix.mValueTX = matrixArray[4];
    itemToPlace.matrix.mValueTY = matrixArray[5];
    itemToPlace.position = [matrixArray[6], matrixArray[7]];
    return itemToPlace;
}

function objectScale(obj, scale, point){
    // Scales an illustrator object with respect to a point on the canvas
    // obj: object, scale: double in percentages, and point = [x: double,y: double]
    var scaleMatrix = app.getScaleMatrix(scale, scale);
    scaleMatrix.mValueTX = (-obj.position[0] - obj.width/2 + point[0])*(1-scale/100);
    scaleMatrix.mValueTY = (-obj.position[1] + obj.height/2 + point[1])*(1-scale/100);
    obj.transform(scaleMatrix);
    return
}

function objectRotate(obj, angle, point){
    // Rotates an illustrator object with respect to a point on the canvas
    // obj: object, rotate: double in degrees, and point = [x: double,y: double]
    var rotationMatrix = app.getRotationMatrix(angle);
    var sine = rotationMatrix.mValueB;
    var cosine = rotationMatrix.mValueA;
    var tx = -obj.position[0] - obj.width/2 + point[0];
    var ty = -obj.position[1] + obj.height/2 + point[1];
    rotationMatrix.mValueTX = -tx*cosine + ty*sine + tx;
    rotationMatrix.mValueTY = -tx*sine - ty*cosine + ty;
    obj.transform(rotationMatrix);
    return
}

function createPicFrameDocument() {
    var doc = app.documents.add(); // create a doc with defaults. If error, target Illustrator
    var firstArtBoard = doc.artboards[0]; // get the default atboard
    var x1 = 0; var y1 = 0; var x2 = 100; var y2 = -150;
    doc.artboards.add([x1, y1, x2, y2]);
    firstArtBoard.remove();
}

// OLD STUFF

/*
function insertPicture2Old(path, name, suffix, Ex, Ey, Fx, Fy){
    var itemToPlace = app.activeDocument.placedItems.add(); // If error, chech Target illustrator
    filepath = path + name + suffix;
    $.writeln(filepath);
    itemToPlace.file = new File(filepath);
    lenEF = Math.sqrt(Math.pow(Ex-Fx,2) + Math.pow(Ey-Fy,2));
    itemToPlace.height = itemToPlace.height / itemToPlace.width * lenEF;
    itemToPlace.width = lenEF;
    rotation = Math.acos((Fx-Ex) / Math.sqrt(Math.pow(Ex-Fx,2) + Math.pow(Ey-Fy,2)))
    if (Fy > Ey){rotation = -rotation;}
    
    ex =-itemToPlace.width / 2;
    ey =-itemToPlace.height/ 2;

    s = Math.sin(-rotation);
    c = Math.cos(-rotation);

    ExNow = c*ex - s*ey;
    EyNow = s*ex + c*ey;

    ExNow = itemToPlace.position[0] + itemToPlace.width /2 + ExNow;
    EyNow = itemToPlace.position[1] - itemToPlace.height/2 + EyNow;

    itemToPlace.rotate(-rotation*180/Math.PI);
    itemToPlace.position = [itemToPlace.position[0] - ExNow + Ex, itemToPlace.position[1] - EyNow + Ey];
    return itemToPlace;
}
*/

/*
function readCropSettingsOld(){
    var artboardRect = app.activeDocument.artboards[0].artboardRect;
    if (artboardRect[0] != 0 || artboardRect[1] != 0 || artboardRect[2] != 100 || artboardRect[3] != -150){ throw new Error("The artboard in this document has wrong dimensions and/or is at wrong position. Create a working document by running script_crop_createDoc.jsx."); }

    pic1 = app.activeDocument.selection[0];
    if (!pic1){ throw new Error("Select the picture you want to crop.")}
    mat = pic1.matrix;

    var sine = mat.mValueB;
    var cosine = mat.mValueA;
    var alpha = Math.asin(sine / Math.sqrt(sine*sine + cosine*cosine))*180 / Math.PI;
    var rot = alpha
    pic1.rotate(rot);
    
    var ex =-pic1.width / 2;
    var ey =-pic1.height/ 2;
    var fx = pic1.width / 2;
    var fy =-pic1.height/ 2;

    alpha = alpha * Math.PI / 180;
    var s = Math.sin(-alpha);
    var c = Math.cos(-alpha);

    var Ex = c*ex - s*ey;
    var Ey = s*ex + c*ey;
    var Fx = c*fx - s*fy;
    var Fy = s*fx + c*fy;

    Ex = pic1.position[0] + pic1.width /2 + Ex;
    Ey = pic1.position[1] - pic1.height/2 + Ey;
    Fx = pic1.position[0] + pic1.width /2 + Fx;
    Fy = pic1.position[1] - pic1.height/2 + Fy;
    
    pic1.rotate(-rot);
    
    var pathSplit = pic1.file.toString().split("/");
    var fileName = pathSplit[pathSplit.length -1].split(".")[0];
    fileName = decodeURIComponent(fileName); // Converts e.g. "%20" to " "
    var picSettings = JSON.parse(readFile(currentDirectory() + "data_pictureSettings.js"));
    picSettings[fileName] = [Ex, Ey, Fx, Fy];
    writeFile(currentDirectory() + "data_pictureSettings.js", JSON.stringify(picSettings,null,1));
    return "Settings for '" + fileName + "' were saved successfully."
}
*/

/*
function cropWithNameOld(nimi){
    // Load JSON data
    var picName = importPicNames();
    var picSettings = importPicSettings();

    if (picName[nimi] && picSettings[picName[nimi]]){
        var pic = insertPicture2(currentDirectory() + '../pictures/', picName[nimi], '.jpg', picSettings[picName[nimi]][0], picSettings[picName[nimi]][1], picSettings[picName[nimi]][2], picSettings[picName[nimi]][3]);
        return pic;
    }else if (picName[nimi]){
        var pic = insertPicture2(currentDirectory() + '../pictures/', picName[nimi], '.jpg', 0, -150, 100, -150);
        return pic;
    }else{
        throw new Error("The person does not have a picture file name. Add it to data_picFileName.js.");
    }
}
*/