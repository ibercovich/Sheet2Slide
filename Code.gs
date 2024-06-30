/**
 * The first slide on the deck will be used as a template
 * to generate the rest by replacing placeholders
 */

const PREFFIX = "__"
const MASTER_DECK_ID = "<INSERT_SLIDE_DECK_ID>";
const SLIDE_WIDTH = 720;
const SLIDE_HEIGHT = 405;
const IMAGE_WIDTH = 225;
const IMAGE_HEIGHT = 165;

/*
Function to create slides from a spreadsheet
It's invoked through a button or menu in the spreadsheet
*/
function createSlides() {
  // Open the Google Slides presentation by ID
  let deck = SlidesApp.openById(MASTER_DECK_ID);
  let slides = deck.getSlides();

  // Loop through all slides except the first one and delete them
  // The first slide is the templat
  for (let i = 1; i < slides.length; i++) {
    slides[i].remove();
  }

  // The first slide is used as a template
  let masterSlide = slides[0];

  // Get the active spreadsheet and the first sheet in it
  // This can be replaced with a specific sheet ID
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  let dataRange = sheet.getDataRange();

  let sheetContentsRaw = dataRange.getValues(); // Get raw values
  let sheetContents = dataRange.getDisplayValues(); // Get formatted values

  // Extract the header and remove it from the data
  let header = sheetContents.shift();
  sheetContentsRaw.shift();

  // Reverse the order of rows so slides are created in the same order as the sheet
  sheetContents.reverse();
  sheetContentsRaw.reverse();

  // Loop through each row and create a new slide
  for (let j = 0; j < sheetContents.length; j++) {
    let row = sheetContents[j];
    let rowRaw = sheetContentsRaw[j];

    // Only create a slide if there is data in the row
    if (row[0]) {
      let values = [];
      let labels = [];

      // Duplicate the master slide
      let slide = masterSlide.duplicate();

      // iterate over columns
      for (let i = 0; i < header.length; i++) {
        const placeholder = header[i];

        // Replace placeholders in the slide with values from the row
        // Not all columns on the table are intended to be replaced with placeholders
        // Therefore we add a preffix to the column name
        if (placeholder.startsWith(PREFFIX)) {

          // replace placeholder (e.g. column name) on slide with jth row data
          slide.replaceAllText(placeholder, row[i]);

          // we want to use some of the columns to create a chart
          if (placeholder.startsWith("__revenue_")) {
            labels.push(placeholder);
            values.push(rowRaw[i]);
          // the logo field contains a file also stored in google drive
          } else if (placeholder == "__logo_file_id__" && rowRaw[i]) {
            // Insert logo image if file ID is present
            let driveImage = DriveApp.getFileById(rowRaw[i]).getBlob();
            if (driveImage) {
              slide.insertImage(driveImage, 87, 60, 50, 50);
            }
          }
          // other special placeholder conditions here
        }
      }

      // Insert a bar chart into the slide
      insertBarChartToSlides(slide, labels, values, true);
    }
  }

  showSlidesDialog();
}

/*
Function to show a dialog with a link to the generated slides
*/
function showSlidesDialog() {
  let deck = SlidesApp.openById(MASTER_DECK_ID);
  let slideUrl = `https://docs.google.com/presentation/d/${deck.getId()}/edit`;
  let html = "<a href='" + slideUrl + "' target='_blank'>Generated Slides</a>";
  let htmlOutput = HtmlService.createHtmlOutput(html).setWidth(250).setHeight(50);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Slides Link');
}

/*
Function to insert a bar chart into a slide
*/
function insertBarChartToSlides(slide, labels, values, insertAsImage) {

  const leftPosition = SLIDE_WIDTH - IMAGE_WIDTH - 15;
  const topPosition = 75;

  // Create a data table for the chart
  let dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Quarter')
    .addColumn(Charts.ColumnType.NUMBER, 'Revenue Runrate');

  // Add data to the data table
  for (let i = 0; i < labels.length; i++) {
    let valueInMillions = values[i] / 1000000;
    dataTable.addRow([labels[i], valueInMillions]);
  }

  // Create the chart
  let chartBuilder = Charts.newColumnChart()
    .setDimensions(IMAGE_WIDTH, IMAGE_HEIGHT)
    .setLegendPosition(Charts.Position.NONE)
    .setColors(["grey"])
    .setDataTable(dataTable);

  let chart = chartBuilder.build();

  if (insertAsImage) {
    // Get a PNG image of the chart and insert it into the slide
    let image = chart.getAs('image/png');
    slide.insertImage(image, leftPosition, topPosition, IMAGE_WIDTH, IMAGE_HEIGHT);
  } else {
    // Insert the chart object into the slide
    slide.insertChart(chart);
  }
}
