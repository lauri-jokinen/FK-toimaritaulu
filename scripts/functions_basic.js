// In this file we have simple/elementary functions that might be useful in many contexts

//// Basic array functions

function unique(array){ // Removes duplicates from array. Preserves first appearances.
    var res = [];
    for (var i = 0; i < array.length; i++){
        if (!isMember(res,array[i])){
            res.push(array[i]);
        };
    };
    return res;
};

function isMember(element, array){ // Checks if element is member of array
    for (var i = 0; i < array.length; i++){
        if (array[i] == element){ return true; }
    }
    return false;
};

function alphabetizeByLastName(array){ // sorts the array according to the last word in each string
    lastNamesFirst = [];
    for (i = 0; i < array.length; i++){
        splittedArray = array[i].split(" ");
        lastName = splittedArray[splittedArray.length-1]
        splittedArray.pop();
        restOfName = splittedArray;
        lastNamesFirst.push(lastName + " " + restOfName.join(" "));
    }
    lastNamesFirst.sort();
    sortedArray = lastNamesFirst;
    result = [];
    for (i = 0; i < array.length; i++){
        splittedArray = sortedArray[i].split(" ");
        lastName = splittedArray[0]
        splittedArray.splice(0,1);
        restOfName = splittedArray;
        result.push(restOfName.join(" ") + " " + lastName);
    }
    return result;
}

//// String modification

String.prototype.replaceAll = function(search, replacement) { // replaceAll
    var target = this;
    return target.replace(new RegExp(search, 'g'), replacement);
};

function replaceText(str, wholeText, partialText){
    // If whole replacement is found, use it. If not, try to replace pieces of text.
    return wholeText[str] || replacePiecesOfText(str, partialText);
};

function replacePiecesOfText(str, jsonData){
    var s = str;
    for (var key in jsonData){
        s = s.replaceAll(key, jsonData[key]);
    };
    return s;
};

//// File I/O

function currentDirectory(){
    var dir_arr = $.fileName.split("/");
    dir_arr.pop();
    return dir_arr.join("/") + "/";
};

function readFile(filepath){
    var file = new File(filepath);
    file.encoding = 'UTF8'; // set some encoding
    file.lineFeed = 'Macintosh'; // set the linefeeds
    file.open('r',undefined,undefined); // read the file
    var content = file.read(); // get the text in it
    file.close(); // close it again
    return content;
};

function writeFile(path, fileContent, encoding) {
    var fileObj = new File(path); // Path includes the filename
    encoding = encoding || "utf-8";
    fileObj = (fileObj instanceof File) ? fileObj : new File(fileObj);

    var parentFolder = fileObj.parent;

    if (!parentFolder.exists && !parentFolder.create()){
        throw new Error("Cannot create file in path " + fileObj.fsName)
    };

    fileObj.encoding = encoding;
    fileObj.open("w");
    fileObj.write(fileContent);
    fileObj.close();
    return fileObj;
};

//// Basic functions for Illustrator

function removeIfEmptyGroup(g){ // g is class groupItem
    for (var i = 0; i < g.groupItems.length; i++){
        removeIfEmptyGroup(g.groupItems[i]);
    }
    if (g.pathItems.length==0 &&
        g.textFrames.length==0 &&
        g.placedItems.length==0 &&
        g.compoundPathItems.length==0 &&
        g.graphItems.length==0 &&
        g.groupItems.length==0 &&
        g.legacyTextItems.length==0 &&
        g.meshItems.length==0 &&
        g.nonNativeItems.length==0 &&
        g.pageItems.length==0 &&
        g.pluginItems.length==0 &&
        g.rasterItems.length==0 &&
        g.symbolItems.length==0){
            g.remove();
        }
};

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
};