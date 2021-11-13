// This script copies art into a matrix form with given texts and creates a fitting artboard.
#target Illustrator
#include "functions_special.js";

/*
How to:
 Select following objects in Illustrator (order matters; first element of the list is on top): 
    Textbox:   Name and title(s) of a person, on two lines
    Rectangle: The back
    Textbox:   Section title
    Rectangle: From which the photo is masked.
 Run code!
*/

// Set the number of columns
var columns = 7 // 12 in 2018 - 14 in 2019 - 13 in 2020

// Replaces left text with the one on the right. It is applied everyhere on the matrix.
var partialText = {
    "Todennäköisyyslaskenta" : "Todennäköisyys-\nlaskenta",
    };

// Same as above, but replaces text if the text is a perfect match
var wholeText = {
    };

var options = {
    "scale" : 57,
    "shiftX" : 0,
    "shiftY" : -62,
    "topMargin" : 180
    };

createTable(columns, options, partialText, wholeText);
