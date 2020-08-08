// This script copies art into a matrix form with given texts and creates a fitting artboard.
#target Illustrator
#include "functions_special.js";

/*
How to:
 Select objects in Illustrator
 Order of objects (from top): 
    Textbox:   Name and title(s) of a person, on two lines
    Rectangle: The back
    Textbox:   Section title
    Rectangle: From which the photo is masked.
*/

var columns = 7 // 12 in 2018 - 14 in 2019 - 13 in 2020

// Replaces left text with the one on the right. 
var partialText = {
    "Todennäköisyyslaskenta" : "Todennäköisyys-\nlaskenta",
    };

// Replaces text if the text is a perfect match
var wholeText = {
    };

var options = {
    "scale" : 57,
    "shiftX" : 0,
    "shiftY" : -62,
    "topMargin" : 180
    };

createTable(columns, options, partialText, wholeText);

// Ohjelman kaatuessa / kaikkea dataa ei ole saatavilla.
// 140 x 210