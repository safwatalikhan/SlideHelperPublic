/**
 * @OnlyCurrentDoc Limits the script to only accessing the current presentation.
 */

/**
 * Create a open translate menu item.
 * @param {Event} event The open event.
 */
function onOpen(event) {
  SlidesApp.getUi()
    .createAddonMenu()
    .addItem("Slide Helper", "showSidebar")
    .addToUi();
}

/**
 * Open the Add-on upon install.
 * @param {Event} event The install event.
 */
function onInstall(event) {
  onOpen(event);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
    "Slide Helper"
  );
  SlidesApp.getUi().showSidebar(ui);
}

/**
 * Recursively gets child text elements a list of elements.
 * @param {PageElement[]} elements The elements to get text from.
 * @return {Text[]} An array of text elements.
 */
function getElementTexts(elements) {
  var texts = [];
  elements.forEach(function(element) {
    switch (element.getPageElementType()) {
      case SlidesApp.PageElementType.GROUP:
        element
          .asGroup()
          .getChildren()
          .forEach(function(child) {
            texts = texts.concat(getElementTexts(child));
          });
        break;
      case SlidesApp.PageElementType.TABLE:
        var table = element.asTable();
        for (var y = 0; y < table.getNumColumns(); ++y) {
          for (var x = 0; x < table.getNumRows(); ++x) {
            texts.push(table.getCell(x, y).getText());
          }
        }
        break;
      case SlidesApp.PageElementType.SHAPE:
        texts.push(element.asShape().getText());
        break;
    }
  });
  return texts;
}

/**
 * Translates selected slide elements to the target language using Apps Script's Language service.
 *
 * @param {string} targetLanguage The two-letter short form for the target language. (ISO 639-1)
 * @return {number} The number of elements translated.
 */

function selectedStuff() {
  var selection = SlidesApp.getActivePresentation().getSelection();
  var selectionType = selection.getSelectionType();
  var texts = [];
  switch (selectionType) {
    case SlidesApp.SelectionType.PAGE:
      var pages = selection
        .getPageRange()
        .getPages()
        .forEach(function(page) {
          texts = texts.concat(getElementTexts(page.getPageElements()));
        });
      break;
    case SlidesApp.SelectionType.PAGE_ELEMENT:
      var pageElements = selection.getPageElementRange().getPageElements();
      texts = texts.concat(getElementTexts(pageElements));
      break;
    case SlidesApp.SelectionType.TABLE_CELL:
      var cells = selection
        .getTableCellRange()
        .getTableCells()
        .forEach(function(cell) {
          texts.push(cell.getText());
        });
      break;
    case SlidesApp.SelectionType.TEXT:
      var elements = selection
        .getPageElementRange()
        .getPageElements()
        .forEach(function(element) {
          texts.push(element.asShape().getText());
        });
      break;
  }
  return texts;
}
function translateSelectedElements(targetLanguage) {
  // Get selected elements.
  var texts = selectedStuff();

  // Translate all elements in-place.
  texts.forEach(function(text) {
    text.setText(
      LanguageApp.translate(text.asRenderedString(), "", targetLanguage)
    );
  });

  return texts.length;
}
function italicizeSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setItalic(toggle);
  });
}

function boldSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setBold(toggle);
  });
}

function resizeAndPosition() {
  
  var highestWidth=0;
  var tempWidth;
  var left=9999999999999;
  var currPage=SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var pageElements=currPage.getPageElements();
  pageElements.forEach(function(pageElement) {
    tempWidth = pageElement.getWidth();
    if(highestWidth<tempWidth) {
      highestWidth = tempWidth;
      left = pageElement.getLeft();
    }
    if(pageElement.getLeft()<left) 
      left = pageElement.getLeft();
  });

  pageElements.forEach(function(pageElement) {
    pageElement.setWidth(highestWidth);
    pageElement.setLeft(left);
  });
}

function underlineSelectedElements(toggle) {
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setUnderline(toggle);
  });
}

function changeFontSize(size){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setFontSize(size);
  });
}
function changeFontType(font){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setFontFamily(font);
  });
}

function changeTextBackgroundColor(color){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setBackgroundColor(color);
  });
}
function changeTextForegroundColor(color){
  selectedStuff().forEach(function(text) {
    text.getTextStyle().setForegroundColor(color);
  });
}

function getTextProp() {
  var docProperties = [false,false,false,0,'','','',0,0,0,0,0];
  var textStyle= SlidesApp.getActivePresentation().getSelection().getTextRange().getTextStyle();
  var pageElements = SlidesApp.getActivePresentation().getSelection().getPageElementRange().getPageElements();
  docProperties[0] = textStyle.isBold()!=null?false:textStyle.isBold();
  docProperties[1] = textStyle.isItalic()!=null?false:textStyle.isItalic();
  docProperties[2] = textStyle.isUnderline()!=null?false:textStyle.isBold();
  docProperties[3] = textStyle.getFontSize();
  docProperties[4] = textStyle.getFontFamily();
  docProperties[5] = textStyle.getBackgroundColor();
  docProperties[6] = textStyle.getForegroundColor();
  docProperties[7] = textStyle.getBaselineOffset();
  pageElements.forEach(function(element) {
    docProperties[8] = element.getHeight();
    docProperties[9] = element.getWidth();
    docProperties[10] = element.getTop();
    docProperties[11] = element.getLeft();
    docProperties[12] = SlidesApp.getActivePresentation().getSelection().getTextRange().asString();
  })
  return docProperties;

}
function setTextProp() {
  var getProp = getTextProp();
  //Logger.log(getProp[0]+' '+getProp[1]+' '+getProp[2]);
  boldSelectedElements(getProp[0]);
  italicizeSelectedElements(getProp[1]);
  underlineSelectedElements(getProp[2]);
  changeFontSize(getProp[3]);
  changeFontType(getProp[4]);
  changeTextBackgroundColor(getProp[5]);
  changeTextForegroundColor(getProp[6]);

}

function alignParagraphText() {
  
  SlidesApp.getActivePresentation().getSelection().getCurrentPage().asSlide().getShapes().forEach(function(shape) {
    shape.getText().getParagraphs().forEach(function(paragraph) {
      paragraph.getRange().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    })
  });
}




/////////////////////////////////////////////Alignment like Master \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function getMasterAlignment() {
  var alignments = [];
  var i=0;
  var master = SlidesApp.getActivePresentation().getMasters()[0];
  master.getPageElements().forEach(function(pageElement) {
    alignments[i]=pageElement.asShape().getText().getParagraphStyle().getParagraphAlignment();
    i++;
  })
  return alignments;
}
function setAlignmentsAsMaster() {
  var alignments = getMasterAlignment();
  alignTextLikeMaster(alignments);
}
function alignTextLikeMaster(masterAlignments) {
  var i=0;
  SlidesApp.getActivePresentation().getSelection().getCurrentPage().getPageElements().forEach(function(pageElement) {
    pageElement.asShape().getText().getParagraphStyle().setParagraphAlignment(masterAlignments[i]);
      i++;
    
  });
}
/////////////////////////////////////////////////////Positions and Dimensions like Master\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

function getMasterPositionDimension() {

  var master = SlidesApp.getActivePresentation().getMasters()[0];
  var temp=[];
  var i=0;
  Logger.log('Number of placeholders in the master: ' + master.getPlaceholders().length);
  master.getPlaceholders().forEach(function(pageElement) {

    temp.push([pageElement.getLeft(),pageElement.getHeight(),pageElement.getTop(),pageElement.getWidth()]);
    i++;
  });
  return temp;
}
function setPositionsDimensionsAsMaster() {
  
  var positionsAndDimensions =  getMasterPositionDimension();
  resizeAndPositionLikeMaster(positionsAndDimensions);

}

function resizeAndPositionLikeMaster(masterIndents) {
  
  
  var currPage=SlidesApp.getActivePresentation().getSelection().getCurrentPage();
  var pageElements=currPage.getPageElements();
  
  var i=0;
  pageElements.forEach(function(pageElement) {
    
    if(masterIndents[i]) {
      pageElement.setLeft(masterIndents[i][0]);
      pageElement.setHeight(masterIndents[i][1]);
      pageElement.setTop(masterIndents[i][2]);
      pageElement.setWidth(masterIndents[i][3]);
    }
      
    i++;
  });
}

//////////////////////////////////////////////////Predictions\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function predictBold() {
  var categoryAttr = 'bold';
  var ignoredAttr = 'fgcolor';
  var selectedProp = getTextProp();
  var currElementInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  return prediction;
}
function predictItalic() {
  var categoryAttr = 'italic';
  var ignoredAttr = 'fgcolor';
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[0], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  return prediction;
}
function predictUnderline() {
  var categoryAttr = 'underline';
  var ignoredAttr = 'fgcolor';
  var selectedProp = getTextProp();
  var currElementInput = {bold: selectedProp[1], italic: selectedProp[1], fontsize: selectedProp[3], fgcolor: selectedProp[5], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  return prediction;
}
function predictFontSizeAfterRuns() {
  var start = new Date().getTime();
  var categoryAttr = 'fontsize';
  var ignoredAttr = 'fgcolor';
  var selectedText = SlidesApp.getActivePresentation().getSelection().getTextRange().asString();
  Logger.log('Selected Text: '+selectedText);
  var selectedProp = getTextProp();
  
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11], run: selectedProp[12]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  Logger.log('Predicted font size: '+prediction);
  var end = new Date().getTime();
  Logger.log("Time taken for predictFontSize(): "+(end - start) + ' miliseconds.');
  return prediction;
}
function predictFontSize() {
  var start = new Date().getTime();
  var categoryAttr = 'fontsize';
  var ignoredAttr = 'fgcolor';
  var selectedText = SlidesApp.getActivePresentation().getSelection().getTextRange().asString();
  Logger.log('Selected Text: '+selectedText);
  var selectedProp = getTextProp();
  
  var currElementInput = {bold: selectedProp[0], italic: selectedProp[1], underline: selectedProp[2], height: selectedProp[8], width: selectedProp[9], top: selectedProp[10], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  Logger.log('Predicted font size: '+prediction);
  var end = new Date().getTime();
  Logger.log("Time taken for predictFontSize(): "+(end - start) + ' miliseconds.');
  return prediction;
}
function predictTop() {
  var categoryAttr = 'top';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {height: selectedProp[8], width: selectedProp[9], left: selectedProp[11]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  
  
  return prediction;
}
function predictLeft() {
  var categoryAttr = 'left';
  var ignoredAttr = null;
  var selectedProp = getTextProp();
  var currElementInput = {height: selectedProp[8], width: selectedProp[9], top: selectedProp[10]};
  var prediction = extractTextProperties(categoryAttr,ignoredAttr, currElementInput);
  
  return prediction;
}
function fixPosition() {
  
  var changedPosition='';
  var top = predictTop();
  var left = predictLeft();
  var currProp = getTextProp();
  var currTop = currProp[10];
  var currLeft = currProp[11];
  Logger.log("Left is: "+left+", Top is: "+top);
  if(currTop===top && currLeft===left) {
    changedPosition ='Positioned perfectly.'
    return changedPosition;
  }
  else {
    var element = SlidesApp.getActivePresentation().getSelection().getPageElementRange().getPageElements()[0] ;
    element.setLeft(left);
    var leftTilt = (currProp[11]-left).toFixed(2);
    //element.setTop(top);
    changedPosition = 'Shifted element '+leftTilt+' points to the left.';
    return changedPosition;
    }
    
}
function extractTextPropertiesBeforeBigSlide(category, ignore) {

  var presentation = SlidesApp.getActivePresentation();//.getSlideById()[0];
  var slides = [];
  var elementId= [];
  var output= [[]];
  var p = [];
  output.pop();
  output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type"]);
  var slideNo="";
  var i=0,k=0;
  var data = [];
  slides=presentation.getSlides();
  var noOfSlidesBeforePresent = slides.length-1;
  Logger.log("Er aage jotogula slides: "+noOfSlidesBeforePresent);
  for(i=0;i<noOfSlidesBeforePresent;i++) {
    slideNo="Slide: "+(i+1);
    Logger.log(slideNo);
   
    
    slides[i].getPageElements().forEach(function(pageElement) {
      //pageElement.select();
      //Logger.log(pageElement.getObjectId()+" -> "+pageElement.getPageElementType()+"Bold: "+pageElement.asShape().getText().getTextStyle().isBold());
      k++;
      var obj = {};
      p[0]=pageElement.asShape().getText().getTextStyle().isBold();
      p[1]= pageElement.asShape().getText().getTextStyle().isItalic();
      p[2]=pageElement.asShape().getText().getTextStyle().isUnderline();
      p[3]=pageElement.asShape().getText().getTextStyle().getFontSize();
      p[4]=pageElement.asShape().getText().getTextStyle().getFontFamily();
      p[5]=pageElement.asShape().getText().getTextStyle().getForegroundColor();
      p[6]=pageElement.asShape().getText().getTextStyle().getBackgroundColor();
      p[7]=pageElement.getPageElementType();
      output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7]]);
      var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5]};
      data.push(newObj);
    });
  }
  Logger.log("Number of page elements: "+k);
  

  for(var j=0;j<=k;j++)
  {
    
     Logger.log(output[j]);   
  
  }
  var result = trainSet(data, category, ignore);
  Logger.log('extractProperties ferot dibe: '+result);
  return result;

}
function extractTextProperties(category, ignore, currElementInput) {
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();//.getSlideById()[0];
  var slides = [];
  var tempOutput= [[]];
  var output= [[]];
  var p = [];
  tempOutput.pop();
  output.pop();
  output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type", "Element Height", "Element Width", "Element position from top", "Element position from left"]);
  var slideNo="";
  var i=0,k=0;
  var data = [];
  slides=presentation.getSlides();
  
  //var noOfSlidesBeforePresent = slides.length-50;
  var currSlideNumber = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
  Logger.log("Currently at slide: "+currSlideNumber);
  for(i=0;i<currSlideNumber;i++) {
    var lStart = new Date().getTime();
    slideNo="Slide: "+(i+1);
    //Logger.log(slideNo+", Slide ID: "+slides[i].getObjectId().toString());
    
    
    slides[i].getPageElements().forEach(function(pageElement) {
      
      
      if(pageElement.getPageElementType().toString()==='SHAPE')
      {
        //pageElement.select();
        //Logger.log(pageElement.getObjectId()+" -> "+pageElement.getPageElementType());
        
        //var texts = pageElement.asShape().getText().asString().split(" ");
      
        var textStyle= pageElement.asShape().getText().getTextStyle();
        //Logger.log("lekha ja paisi "+texts);
        //var obj = {};
      
        p[0]=textStyle.isBold();
        p[1]= textStyle.isItalic();
        p[2]=textStyle.isUnderline();
        p[3]=textStyle.getFontSize();
        p[4]=textStyle.getFontFamily();
        p[5]=textStyle.getForegroundColor();
        p[6]=textStyle.getBackgroundColor();
        p[7]=pageElement.getPageElementType();
        p[8]=pageElement.getHeight();
        p[9]=pageElement.getWidth();
        p[10]=pageElement.getTop();
        p[11]=pageElement.getLeft();
        
        if(p[3]!==null )//&& p[10]<400 && p[11]<400
        {
          
          output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]);
          //tempOutput.push(([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]));
          //Logger.log(tempOutput.pop());
          var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11]};
          data.push(newObj);
          k++;
          
        }
        
      } 
      
    });
    var lEnd = new Date().getTime();
    Logger.log('Time taken for slide '+(i+1)+': '+(lEnd-lStart)+' miliseconds.');
  }
  Logger.log("Number of page elements: "+k);
  

   for(var j=0;j<=k;j++)
   {
    
      Logger.log(output[j]);   
  
   }
  var trainSetStart = new Date().getTime();
  var result = trainSet(data, category, ignore, currElementInput);
  var trainSetEnd = new Date().getTime();
  Logger.log('trainSet() takes '+(trainSetEnd-trainSetStart)+' miliseconds to compute.');
  Logger.log('extractProperties returns: '+result);
  var end = new Date().getTime();
  Logger.log('Time taken for extractProperties(): '+ (end-start)+' miliseconds.');
  return result;

}
function extractTextPropertiesAfterRuns(category, ignore, currElementInput) {
  var start = new Date().getTime();
  var presentation = SlidesApp.getActivePresentation();//.getSlideById()[0];
  var slides = [];
  var tempOutput= [[]];
  var output= [[]];
  var p = [];
  tempOutput.pop();
  output.pop();
  output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type", "Element Height", "Element Width", "Element position from top", "Element position from left", "Run"]);
  var slideNo="";
  var i=0,k=0;
  var data = [];
  slides=presentation.getSlides();
  
  //var noOfSlidesBeforePresent = slides.length-50;
  var currSlideNumber = SlidesApp.getActivePresentation().getSelection().getCurrentPage().getObjectId().toString().replace( /^\D+/g, '');
  Logger.log("Currently at slide: "+currSlideNumber);
  for(i=0;i<currSlideNumber;i++) {
    var lStart = new Date().getTime();
    slideNo="Slide: "+(i+1);
    //Logger.log(slideNo+", Slide ID: "+slides[i].getObjectId().toString());
    
    
    slides[i].getPageElements().forEach(function(pageElement) {
      
      
      if(pageElement.getPageElementType().toString()==='SHAPE')
      {
          var textRange = pageElement.asShape().getText(); // Get text belonging to current page element
          var elementType =pageElement.getPageElementType();
          var elementHeight=pageElement.getHeight();
          var elementWidth=pageElement.getWidth();
          var elementTop=pageElement.getTop();
          var elementLeft=pageElement.getLeft();
          textRange.getRuns().forEach(function(run) { // Loop through all runs in text
          var textStyle = run.getTextStyle(); // Get current row text style
          if(run.getLength()>1)
          {
            //Logger.log('Run: '+run.asString()+'--> Bold: '+textStyle.isBold()+', Italic: '+textStyle.isItalic()+', Underline: '+textStyle.isUnderline());
            p[0]=textStyle.isBold();
            p[1]= textStyle.isItalic();
            p[2]=textStyle.isUnderline();
            p[3]=textStyle.getFontSize();
            p[4]=textStyle.getFontFamily();
            p[5]=textStyle.getForegroundColor();
            p[6]=textStyle.getBackgroundColor();
            p[7]=elementType;
            p[8]=elementHeight;
            p[9]=elementWidth;
            p[10]=elementTop;
            p[11]=elementLeft;
            p[12]=run.asString();
            output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11],p[12]]);
          //tempOutput.push(([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]));
          //Logger.log(tempOutput.pop());
          var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11], run: p[12]};
          data.push(newObj);
          k++;
          }
        });
        
        
      } 
      
    });
    var lEnd = new Date().getTime();
    Logger.log('Time taken for slide '+(i+1)+': '+(lEnd-lStart)+' miliseconds.');
  }
  Logger.log("Number of page elements: "+k);
  

   for(var j=0;j<=k;j++)
   {
    
      Logger.log(output[j]);   
  
   }
  var result = trainSet(data, category, ignore, currElementInput);
  Logger.log('extractProperties returns: '+result);
  var end = new Date().getTime();
  Logger.log('Time taken for extractProperties(): '+ (end-start)+' miliseconds.');
  return result;

}
function extractTextPropertiesPreload(data, category, ignore, currElementInput) {
  var result = trainSet(data, category, ignore, currElementInput);
  Logger.log('extractProperties returns: '+result);
  return result;

}


function trainSet(data, category, ignore, currElementInput) {
  var start = new Date().getTime();
  var config = {
    trainingSet: data,
    categoryAttr: category,
    ignoredAttributes: [ignore]
  
  };
  var decisionTree = new dt.DecisionTree(config);
  var testPredict = {italic: true, underline: false, fontsize: 14, fgcolor: 'Color'};
  var selectedProp = getTextProp();
  //var currElementInput = {italic: selectedProp[1], underline: selectedProp[2], fontsize: selectedProp[3], fgcolor: selectedProp[5]};
  var decisionTreePrediction = decisionTree.predict(currElementInput); 

  Logger.log("Predicted value: "+decisionTreePrediction);
  Logger.log('trainSet returns: '+decisionTreePrediction);
  var end = new Date().getTime();
  Logger.log('Time taken for trainSet(): '+ (end-start)+' miliseconds.');
  return decisionTreePrediction;

  
}
// function buildData() {
//   const performance = require('perf_hooks').performance;
//   var t0 = performance.now();
//   var presentation = SlidesApp.getActivePresentation();//.getSlideById()[0];
//   var slides = [];
//   var elementId= [];
//   var output= [[]];
//   var p = [];
//   output.pop();
//   output.push(["Bold", "Italic", "Underline", "Font Size", "Font Family", "Foreground Color", "Background Color", "Element Type", "Element Height", "Element Width", "Element position from top", "Element position from left"]);
//   var slideNo="";
//   var i=0,k=0;
//   var data = [];
//   slides=presentation.getSlides();
  
//   var noOfSlidesBeforePresent = slides.length-1;
//   for(i=0;i<noOfSlidesBeforePresent;i++) {
//     slideNo="Slide: "+(i+1);
//     //Logger.log(slideNo+", Slide ID: "+slides[i].getObjectId().toString());
    
    
//     slides[i].getPageElements().forEach(function(pageElement) {
      
      
//       if(pageElement.getPageElementType().toString()==='SHAPE')
//       {
//         //pageElement.select();
//         //Logger.log(pageElement.getObjectId()+" -> "+pageElement.getPageElementType());
        
//         //var texts = pageElement.asShape().getText().asString().split(" ");
      
        
//         //Logger.log("lekha ja paisi "+texts);
//         var obj = {};
      
//         p[0]=pageElement.asShape().getText().getTextStyle().isBold();
//         p[1]= pageElement.asShape().getText().getTextStyle().isItalic();
//         p[2]=pageElement.asShape().getText().getTextStyle().isUnderline();
//         p[3]=pageElement.asShape().getText().getTextStyle().getFontSize();
//         p[4]=pageElement.asShape().getText().getTextStyle().getFontFamily();
//         p[5]=pageElement.asShape().getText().getTextStyle().getForegroundColor();
//         p[6]=pageElement.asShape().getText().getTextStyle().getBackgroundColor();
//         p[7]=pageElement.getPageElementType();
//         p[8]=pageElement.getHeight();
//         p[9]=pageElement.getWidth();
//         p[10]=pageElement.getTop();
//         p[11]=pageElement.getLeft();
//         if(p[3]!==null )//&& p[10]<400 && p[11]<400
//         {
          
//           output.push([p[0],p[1],p[2],p[3],p[4],p[5],p[6],p[7],p[8],p[9],p[10],p[11]]);
//           var newObj = {bold: p[0], italic: p[1], underline: p[2],fontsize: p[3],fgcolor: p[5], height: p[8], width: p[9], top: p[10], left: p[11]};
//           data.push(newObj);
//           k++;
//         }
        
//       } 
      
//     });
//   }
//   Logger.log("Number of page elements: "+k);
  

//   for(var j=0;j<=k;j++)
//   {
    
//      Logger.log(output[j]);   
  
//   }
//   var t1 = performance.now();
//   Logger.log('buildData() takes '+(t1 - t0).toFixed(4)+'miliseconds to run.');
//   return data;
// }