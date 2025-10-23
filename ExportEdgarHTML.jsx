/**
 * @@@BUILDINFO@@@ ExportEdgarHTML.jsx !Version! Thu Oct 23 2025 12:28:01 GMT+0530
 */
try {
  main();
} catch (e) {
  $.writeln(e.line + "\n" + e.message);
}

function main() {
  app.scriptPreferences.userInteractionLevel =
    UserInteractionLevels.NEVER_INTERACT;
  var myDoc = preCheck();
  exportFileToPDF(myDoc);
  convertTextVariablesToText(myDoc);
  overrideMasterItems(myDoc) ;
  findChangeByText(myDoc);
  splitStories(myDoc);
  exportFileToHTML(myDoc); 
  procssHTML(myDoc); // WIP
  // exportFileToIDML(myDoc);
  // processIDML(myDoc); // -- WIP
  app.scriptPreferences.userInteractionLevel =
    UserInteractionLevels.INTERACT_WITH_ALL;
}

function preCheck() {
  $.write(arguments.callee.name);
  var myDoc = app.documents[0];
  if (!myDoc.isValid) {
    alert("The document must be opened before proceeding.");
    exit(0);
  }

  if (!myDoc.saved) {
    alert("The document must be saved before proceeding.");
    exit(0);
  }

  if (myDoc.modified) {
    alert("The document must be saved before proceeding.");
    exit(0);
  }
  $.writeln(": completed.");
  return myDoc;
}

function exportFileToPDF(myDoc) {
  $.write(arguments.callee.name);

  var pdfFile = new File(myDoc.fullName.fsName.replace(".indd", ".pdf"));
  preset = getPreset();
  if (preset && preset.isValid) {
    myDoc.exportFile(ExportFormat.PDF_TYPE, pdfFile, false, preset);
  } else {
    myDoc.exportFile(ExportFormat.PDF_TYPE, pdfFile);
  }
  $.writeln(": completed.");
}

function getPreset() {
  var presets = app.pdfExportPresets;
  var listOfPreset = presets.everyItem().name;
  var window = new Window("dialog", "PDF Preset");
  var dd = window.add("dropdownlist", [0, 0, 200, 28], listOfPreset);
  var btn = window.add("button", [0, 0, 125, 28], "Go!", { name: "ok" });
  btn.enabled = false;
  dd.onChange = function () {
    btn.enabled = true;
  };
  if (window.show() == 1) {
    return presets[dd.selection.index];
  }
}

function convertTextVariablesToText(myDoc) {
  $.write(arguments.callee.name);
  var textVariables = myDoc.textVariables;
  var nTextVariable = textVariables.length;

  for (var i = nTextVariable - 1; i >= 0; i--) {
    textVariables[i].convertToText();
  }
  $.writeln(": completed.");
}

function overrideMasterItems(myDoc) {
	$.write(arguments.callee.name);
	var pages = myDoc.pages;
	var nPage = pages.length;
	for(var i=0; i<nPage; i++){
		var page = pages[i];
		var items = page.masterPageItems;
		var nItem = items.length;
		for(var j=0; j<nItem; j++){
			items[j].override(page);
		}
	}
	$.writeln(": completed.");
}

function findChangeByText(myDoc) {
  $.write(arguments.callee.name);
  var list = { "^m": "[em]", "^>": "[en]", "^/": "[fs]" };
  for (var key in list) {
    app.findTextPreferences = app.changeGrepPreferences = null;
    app.findTextPreferences.findWhat = key;
    app.changeTextPreferences.changeTo = list[key];
    myDoc.changeText();
    app.findTextPreferences = app.changeGrepPreferences = null;
  }
  $.writeln(": completed.");
}

function splitStories(myDoc) {
  $.write(arguments.callee.name);
  var stories = myDoc.stories;
  var nStory = stories.length;
  for (var i = 0; i < nStory; i++) {
    var story = stories[i];
    splitStory(story);
    removeStory(story);
  }
  $.writeln(": completed.");
}

function splitStory(story) {
  // $.write(arguments.callee.name);
  var textContainers = story.textContainers;
  var nTextContainer = textContainers.length;
  if (nTextContainer == 1) return;
  for (var i = nTextContainer - 1; i >= 0; i--) {
    textContainers[i].duplicate();
  }
  // $.writeln(": completed.");
}

function removeStory(story) {
  // $.write(arguments.callee.name);
  var textContainers = story.textContainers;
  var nTextContainer = textContainers.length;
  if (nTextContainer == 1) return;
  for (var i = nTextContainer - 1; i >= 0; i--) {
    textContainers[i].remove();
  }
  // $.writeln(": completed.");
}

function exportFileToHTML(myDoc) {
  $.write(arguments.callee.name);
  var htmlFile = new File(myDoc.fullName.fsName.replace(".indd", ".html"));
  myDoc.exportFile(ExportFormat.HTML, htmlFile);
  if (htmlFile.exists) {
    var edgarHtml = new File(htmlFile.fsName.replace(".html", "_EDGAR.htm"));
    htmlFile.copy(edgarHtml);
  }
  $.writeln(": completed.");
}

function procssHTML(myDoc) {
  $.write(arguments.callee.name);
  var idmlFile = new File(myDoc.fullName.fsName.replace(".indd", ".idml"));
  if(idmlFile.exists) {
	  // Work-in-progress
  }
  $.writeln(": completed.");
}

function exportFileToIDML(myDoc) {
  $.write(arguments.callee.name);
  var idmlFile = new File(myDoc.fullName.fsName.replace(".indd", ".idml"));
  myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile);
  $.writeln(": completed.");
}

function processIDML(myDoc) {
  $.write(arguments.callee.name);
  var docName = myDoc.name.replace(/\.indd$/i, "");
  var docPath = myDoc.filePath.fsName;
  var idmlFile = new File(docPath + "/" + docName + ".idml");
  if (idmlFile.exists) {
    var destinationFolder = new Folder(docPath + "/" + docName);
    if (!destinationFolder.exists) {
      destinationFolder.create();
    }
    app.unpackageUCF(idmlFile, destinationFolder);
    var designmapXmlFile = new File(destinationFolder + "/designmap.xml");
    if (designmapXmlFile.exists) {
      var designmapXml = getXml(designmapXmlFile);
      var spreads = designmapXml.xpath("/Document/idPkg:Spread");
      var nSpread = spreads.length();
      for (var i = 0; i < nSpread; i++) {
          var spreadXmlFile = new File(destinationFolder + "/" + spreads[i].@src);
        if (spreadXmlFile.exists) {
          var xmlOutputData = "<div>";
          var spreadXml = getXml(spreadXmlFile);
          var spread = spreadXml.xpath(
            "/idPkg:Spread/Spread"
          );
		  var bounds = spread.Page[0].@GeometricBounds.split(" ");
		  var pageWidth = bounds[3];
		  var pageHeight = bounds[2];
		  var textFrames = spread.TextFrame;
		  var nTextFrame = textFrames.length();
		  for(var j=0; j<nTextFrame; j++) {
			  var textFrame = textFrames[j];
			  // $.bp(textFrame.@ParentStory == "u107");
			  var ItemTransformArray = textFrame.@ItemTransform.split(" ");
			  var itemX = ItemTransformArray[4];
			  var itemY = ItemTransformArray[5];
			  if(itemX < 0) {
				  itemX = Number(pageWidth) + Number(itemX);
			  }
			itemY = Number(pageHeight) / 2 + Number(itemY);
			  var PathPointTypeArray = textFrame.Properties.PathGeometry.GeometryPathType.PathPointArray.PathPointType[0].@Anchor.split(" ");
			  var PathPointX = PathPointTypeArray[0];
			  var PathPointY = PathPointTypeArray[1];
			  $.writeln("Item X: " + (Number(itemX) + Number(PathPointX)) + "\rItem Y: " + (Number(itemY) + Number(PathPointY)));
		  }
//~           var nParagraph = paragraphs.length();
//~           for (var j = 0; j < nParagraph; j++) {
//~             var paragraph = paragraphs[j];
//~             xmlOutputData += "<p>";
//~             var characters = paragraph.CharacterStyleRange;
//~             var nCharacter = characters.length();
//~             for (var k = 0; k < nCharacter; k++) {
//~                 var character = characters[k];
//~                 xmlOutputData += "<span>";
//~                 xmlOutputData += character.Content.toString();
//~               xmlOutputData += "</span>";
//~             }
//~             xmlOutputData += "</p>";
//~           }
            xmlOutputData += "</div>";
            $.writeln(xmlOutputData);
        }
      }
    }
  }
  $.writeln(": completed.");
}

function getXml(file) {
  file.open("r");
  const data = file.read();
  file.close();
  return new XML(data);
}
