// Function to read configuration from a text file
function readConfig(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
        // Alert if the file doesn't exist
        alert("File does not exist");
        return {};
    }
    file.open("r"); // Open the file for reading
    var content = file.read(); // Read the entire file content
    file.close(); // Close the file

    var lines = content.split("\n"); // Split content into lines
    var config = {};

    // Parse each line and store key-value pairs in the config object
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i]
        if (line && line.indexOf('=') !== -1) { // Check if line is not empty and contains '='
            var parts = line.split('=');
            var key = parts[0];
            var value = parts[1];
            config[key] = value;
        }
    }
    return config;
}

// Function to read and process CSV file, removing quotes from all fields
function readAndProcessCSV(filePath) {
    var file = new File(filePath);
    if (!file.exists) {
        // Alert if the file doesn't exist
        alert("File does not exist");
        return [];
    }
    file.open("r"); // Open the file for reading
    var content = file.read(); // Read the entire file content
    file.close(); // Close the file

    // Split the file content into rows
    var rows = content.split("\n");
    var processedData = [];

    // Function to remove quotes from a given string
    function removeQuotes(part) {
        // Use a regular expression to remove quotes from the string
        return part.replace(/"/g, "");
    }

    // Loop through rows starting from the second row to skip the header
    for (var i = 1; i < rows.length; i++) {
        var columns = rows[i].split(","); // Split each row into columns
        for (var j = 0; j < columns.length; j++) {
            columns[j] = removeQuotes(columns[j]); // Remove quotes from each column
        }
        processedData.push(columns); // Add processed columns to the data array
    }
    return processedData; // Return the processed data
}

// Function to create and style a text frame
function createTextFrame(bounds, content, fontSize, justification, verticalJustification) {
    // Add a new text frame to the first page of the document
    var textFrame = newDoc.pages[0].textFrames.add({
        geometricBounds: bounds, // Set the position and size of the text frame
    });

    var newText = textFrame.insertionPoints[0].parentStory; // Get the story object of the text frame
    newText.contents = content.toUpperCase(); // Convert content to uppercase
    textFrame.texts[0].pointSize = fontSize; // Set the font size of the text
    textFrame.texts[0].justification = justification; // Set the text justification

    // Set the applied font to "Syne", if available
    var font = app.fonts.item("Syne");
    if (font.isValid) {
        textFrame.texts[0].appliedFont = font; // Apply the font to the text

        // Create or retrieve a bold character style for the text
        var boldCharacterStyleName = "BoldStyle_" + content; // Unique name for the character style
        var boldCharacterStyle = newDoc.characterStyles.itemByName(boldCharacterStyleName).isValid ?
            newDoc.characterStyles.itemByName(boldCharacterStyleName) :
            newDoc.characterStyles.add({ name: boldCharacterStyleName, fontStyle: "Bold" });

        // Apply the bold character style to the text
        textFrame.texts[0].applyCharacterStyle(boldCharacterStyle);
    }

    // Set vertical justification if the text frame supports it
    if (textFrame.hasOwnProperty("textFramePreferences")) {
        textFrame.textFramePreferences.verticalJustification = verticalJustification;
    }

    return textFrame; // Return the created text frame
}

// Get Path
var scriptFile = new File($.fileName); // Get the current script file
var scriptDir = scriptFile.path; // Get the directory of the script file

// Read Config
var configFilePath = scriptDir + "/config.txt"; // Replace with your actual file path
var config = readConfig(configFilePath);

// Parse Config Data
var row1Bounds = [
    parseFloat(config.row1Top),
    parseFloat(config.row1Left),
    parseFloat(config.row1Bottom),
    parseFloat(config.row1Right)
];

// Parse the text frame bounds from the config for the second row
var row2Bounds = [
    parseFloat(config.row2Top),
    parseFloat(config.row2Left),
    parseFloat(config.row2Bottom),
    parseFloat(config.row2Right)
];

// Parse the text frame bounds from the config for the third row
var row3Bounds = [
    parseFloat(config.row3Top),
    parseFloat(config.row3Left),
    parseFloat(config.row3Bottom),
    parseFloat(config.row3Right)
];

// Parse the text frame bounds from the config for the bottom left frame
var bottomLeftBounds = [
    parseFloat(config.bottomleftTop),
    parseFloat(config.bottomleftLeft),
    parseFloat(config.bottomleftBottom),
    parseFloat(config.bottomleftRight)
];

// Parse the text frame bounds from the config for the bottom right frame
var bottomRightBounds = [
    parseFloat(config.bottomrightTop),
    parseFloat(config.bottomrightLeft),
    parseFloat(config.bottomrightBottom),
    parseFloat(config.bottomrightRight)
];

// Read CSV
var csvFilePath = scriptDir + "/CSV_Data.csv"; // Construct the path to the CSV file
var csvData = readAndProcessCSV(csvFilePath); // Read and process the CSV data

// Set document settings using inches
var newDoc = app.documents.add({
    viewPreferences: {
        // Set the measurement units to inches
        horizontalMeasurementUnits: MeasurementUnits.INCHES,
        verticalMeasurementUnits: MeasurementUnits.INCHES,
        cursorKeyIncrement: '1in', // Cursor movement increment set to 1 inch
    },
    documentPreferences: {
        // Use values from config for page size
        pageWidth: config.pageWidth || '5in', // Default to 5 inches if not specified in config
        pageHeight: config.pageHeight || '3in', // Default to 3 inches if not specified in config
        documentSlugUniformSize: true,
        slugTopOffset: '0.25in', // Slug area at the top set to 0.25 inches
        marginPreferences: {
            // Use values from config for margins
            top: parseFloat(config.marginTop) || 0.25,
            left: parseFloat(config.marginLeft) || 0.25,
            bottom: parseFloat(config.marginBottom) || 0.25,
            right: parseFloat(config.marginRight) || 0.25,
        },
    },
});

// Loop through CSV data
for (var row = 0; row < csvData.length; row++) {
    // Log SKU for debugging
    $.writeln("SKU for row " + (row + 1) + ": " + csvData[row][0]);

    // Create Text Box 1 with CSV data
    createTextFrame(row1Bounds, csvData[row][1], csvData[row][2], Justification.CENTER_ALIGN, VerticalJustification.CENTER_ALIGN);

    // Create Text Box 2 with CSV data
    createTextFrame(row2Bounds, csvData[row][3], csvData[row][4], Justification.CENTER_ALIGN, VerticalJustification.CENTER_ALIGN);

    // Create Text Box 3 with CSV data
    createTextFrame(row3Bounds, csvData[row][5], csvData[row][6], Justification.CENTER_ALIGN, VerticalJustification.CENTER_ALIGN);

    // Place the second image at the top and horizontally center it
    var imagePath2 = scriptDir + "/Logo.png"; // Path to the image file
    var imageFile2 = File(imagePath2); // Create a File object for the image
    if (imageFile2.exists) {
        // If the image file exists

        // Get the page width
        var pageWidth = newDoc.pages[0].bounds[3] - newDoc.pages[0].bounds[1];

        // Get the width of the image frame
        var imageWidth = parseFloat(config.imageWidth) || 4; // Desired width of the image frame

        // Calculate the left and right bounds to horizontally center the image
        var leftBound = (pageWidth - imageWidth) / 2;
        var rightBound = leftBound + imageWidth;

        // Add the image frame to the page and set its bounds
        var imageFrame2 = newDoc.pages[0].rectangles.add({
            geometricBounds: [0.16, leftBound, 0.5217, rightBound], // Position and size of the image frame
        });
        imageFrame2.place(imageFile2); // Place the image in the frame

        imageFrame2.fit(FitOptions.PROPORTIONALLY); // Fit the image proportionally in the frame

        imageFrame2.strokeWeight = 0; // Set the stroke weight of the image frame to 0
    }

    // Create additional text frames with CSV data
    createTextFrame(bottomLeftBounds, csvData[row][7], csvData[row][8], Justification.LEFT_ALIGN, VerticalJustification.BOTTOM_ALIGN);
    createTextFrame(bottomRightBounds, csvData[row][9], csvData[row][10], Justification.RIGHT_ALIGN, VerticalJustification.BOTTOM_ALIGN);

    // Export the document as a PDF using the SKU as the filename
    var pdfFilename = csvData[row][0]; // Use the SKU as the PDF filename
    var pdfFilePath = scriptDir + "/" + pdfFilename + ".pdf"; // Construct the path for the PDF file

    newDoc.exportFile(ExportFormat.PDF_TYPE, new File(pdfFilePath), false); // Export the PDF

    // Clear the page for the next iteration
    newDoc.pages[0].textFrames.everyItem().remove(); // Remove all text frames
    newDoc.pages[0].rectangles.everyItem().remove(); // Remove all rectangles
}

// Close the document without saving
newDoc.close(SaveOptions.NO);