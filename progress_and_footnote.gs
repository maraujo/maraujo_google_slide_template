/**
 * @OnlyCurrentDoc Adds progress bars to a presentation.
 */
var PROGRESS_BAR_ID = 'PROGRESS_BAR_ID';
var FOOTNOTE_ID = 'FOOTNOTE_BAR_ID';
var PRESENTER_NAME = "Matheus Araújo";
var TITLE = "The Title of Your Project";
var COLOR_1 = "#7a0019" // Darker
var COLOR_2 = "#FFFFFF" // Lighter
var COLOR_3 = "#EEEEEE" // Neutral
var PROGRESS_BAR_HEIGHT = 5; // px
var FONT_SIZE = 11;
var FOOTNOTE_HEIGHT = 20; // px
var NAME_WIDTH = 0.15;
var TITLE_WIDTH = 0.8;
var SLIDENO_WIDTH = 0.05;
var presentation = SlidesApp.getActivePresentation();

/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen();
}

/**
 * Trigger for opening a presentation.
 * @param {object} e The onOpen event.
 */
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Show progress bar', 'createBars')
      .addItem('Show footnote', 'createFootnote')
      .addItem('Hide progress bar', 'deleteBars')
      .addItem('Hide footnote', 'deleteFootnote')
      .addToUi();
}

/**
 * Create a rectangle on every slide with different bar widths.
 */
function createBars() {
  deleteBars(); // Delete any existing progress bars
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var ratioComplete = ((i +1) / (slides.length ));
    var x = 0;
    var y = 0;
    var barWidth = presentation.getPageWidth() * ratioComplete;
    if (barWidth > 0) {
      var bar = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,
                                      barWidth, PROGRESS_BAR_HEIGHT);
      bar.getBorder().setTransparent();
      bar.getFill().setSolidFill(COLOR_1)
      bar.setLinkUrl(PROGRESS_BAR_ID);
    }
  }
}

/**
 * Deletes all progress bar rectangles.
 */
function deleteBars() {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; ++j) {
      var el = elements[j];
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
          el.asShape().getLink() &&
          el.asShape().getLink().getUrl() === PROGRESS_BAR_ID) {
        el.remove();
      }
    }
  }
}

function createFootnote() {
  deleteBars(); // Delete any existing progress bars
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var x = 0;
    var y = presentation.getPageHeight() - FOOTNOTE_HEIGHT;
    var barWithNameWidth = presentation.getPageWidth() * NAME_WIDTH; //10% of the width
    var barWithTitleWidth = presentation.getPageWidth() * TITLE_WIDTH; //10% of the width
    var barWithSlidenoWidth = presentation.getPageWidth() * SLIDENO_WIDTH; //10% of the width
    
    // Setting bar with name
    var barWithName = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,barWithNameWidth, FOOTNOTE_HEIGHT);
    barWithName.getBorder().setTransparent();
    barWithName.getFill().setSolidFill(COLOR_1);
    barWithName.setLinkUrl(FOOTNOTE_ID);
    var nameText = barWithName.getText().setText(PRESENTER_NAME);
    var nameStyle = nameText.getTextStyle();
    nameStyle.setForegroundColor(COLOR_2);
    nameStyle.setFontSize(FONT_SIZE);
    nameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Setting bar with presentation title
    var barWithTitle = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, barWithName.getWidth(), y,barWithTitleWidth, FOOTNOTE_HEIGHT);
    barWithTitle.getBorder().setTransparent();
    barWithTitle.getFill().setSolidFill(COLOR_3);
    barWithTitle.setLinkUrl(FOOTNOTE_ID);
    var titleText = barWithTitle.getText().setText(TITLE);
    var titleStyle = titleText.getTextStyle();
    titleStyle.setForegroundColor(COLOR_1);
    titleStyle.setFontSize(FONT_SIZE);
    titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    
    // Setting bar with slide number
    var barWithSlideno = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, barWithName.getWidth() + barWithTitle.getWidth(), y, barWithSlidenoWidth, FOOTNOTE_HEIGHT);
    barWithSlideno.getBorder().setTransparent();
    barWithSlideno.getFill().setSolidFill(COLOR_1);
    barWithSlideno.setLinkUrl(FOOTNOTE_ID);
    var slidenoText = barWithSlideno.getText().setText(i);
    var slidenoStyle = slidenoText.getTextStyle();
    slidenoStyle.setForegroundColor(COLOR_2);
    slidenoStyle.setFontSize(FONT_SIZE);
    slidenoText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
}


function deleteFootnote() {
  var slides = presentation.getSlides();
  for (var i = 0; i < slides.length; ++i) {
    var elements = slides[i].getPageElements();
    for (var j = 0; j < elements.length; ++j) {
      var el = elements[j];
      if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE &&
          el.asShape().getLink() &&
          el.asShape().getLink().getUrl() === FOOTNOTE_ID) {
        el.remove();
      }
    }
  }
}
