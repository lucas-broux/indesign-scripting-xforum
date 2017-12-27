////////////////////////////////////////////////////////////////////////////////
// Main content of the script.
////////////////////////////////////////////////////////////////////////////////

// Useful variable definition : link to working directory.
const link = "C:\\Path\\To\\Directory";


// User interface.
main();
function main(){
    setup();
    snippet();
}
function setup(){}
function snippet(){
  var dialog = app.dialogs.add({name:"Badges", canCancel:true});
	with(dialog){
		//Add a dialog column.
		with(dialogColumns.add()){
			//Create a border panel.
			with(borderPanels.add()){
				with(dialogColumns.add()){
					staticTexts.add({staticLabel:"Function:"});
				}
				with(dialogColumns.add()){
					//Create a pop-up menu ("dropdown") control.
					var function_choice = dropdowns.add({stringList:["Intervenants", "Organisateur", "Intervenant vide", "Eleve contact"], selectedIndex:0});
				}
			}
		}
	}
	//Display the dialog box.
	if(dialog.show() == true){
		// If the user didnt click the Cancel button,
		// then get the values back from the dialog box.
		// Get the choice of function.
		if(function_choice.selectedIndex == 0){
      // Badges intervenants.
      var documentName = link.concat("\\Intervenant\\badge_intervenant.indd");
      var pdfName = link.concat("\\Intervenant\\planche_intervenant");
      var csv_delimiter = ",";
      var link_to_csv = link.concat("\\Intervenant\\Intervenants.csv");
      var infos = infoFromCsv(link_to_csv, csv_delimiter);
      var background_link = link.concat("\\Intervenant\\Badge Intervenant.pdf");
		}
		else if(function_choice.selectedIndex == 1){
      // Badges Organisateurs
      var documentName = link.concat("\\Organisateur\\badge_organisateur.indd");
      var pdfName = link.concat("\\Organisateur\\planche_organisateur");
      var link_to_csv = link.concat("\\Organisateur\\Annuaire.csv");
      var csv_delimiter = ",";
      var infos = infoFromCsv(link_to_csv, csv_delimiter);
      var background_link = link.concat("\\Organisateur\\Badge Organisateur.pdf");
		}
    else if (function_choice.selectedIndex == 2){
      // Badges Intervenants vides.
      var documentName = link.concat("\\Intervenant_vide\\badge_intervenant.indd");
      var pdfName = link.concat("\\Intervenant_vide\\planche_intervenant");
      var background_link = link.concat("\\Intervenant_vide\\Badge Intervenant.pdf");
      var infos = [];
      for (i = 0; i < 10; i++) {
        infos.push({"Nom" : "........................................", "Fonction" : ".............................................", "Entreprise" : "................................." });
      }
    }
    else {
      // Badges élève contact.
      var documentName = link.concat("\\Eleve contact\\badge_contact.indd");
      var pdfName = link.concat("\\Eleve contact\\planche_contact");
      var link_to_csv = link.concat("\\Eleve contact\\eleve_contact.csv");
      var csv_delimiter = "#";
      var infos = [];
      for (i = 0; i < 10; i++) {
        infos.push({"Nom" : "........................................", "Fonction" : ".............................................", "Entreprise" : "................................." });
      }
      var background_link = link.concat("\\Eleve contact\\Badge contact.pdf");
    }

		publish(documentName, pdfName, infos, background_link, 10);
	}
	else{
		dialog.destroy()
	}
}


/**
  Generate and export the badges given all relevant parameters.
*/
function publish(documentName, pdfName, infos, background_link, nb_pages_per_batch) {
  // Define useful variables.
  var current_batch = 0;
  var nb_columns = 2;
  var nb_lines = 5;
  // Loop over badges.
  for (badge_counter = 0; badge_counter < infos.length; badge_counter++){
    // Get info.
    var person = infos[badge_counter];
    // Get grid position.
    var current_page_number =  Math.floor(badge_counter/(nb_columns * nb_lines));
    var position_in_page = badge_counter % (nb_columns * nb_lines);
    var current_line = Math.floor(position_in_page/nb_columns);
    var current_column = position_in_page % nb_columns;

    // If new batch.
    if (badge_counter % (nb_pages_per_batch * nb_columns * nb_lines) == 0){
      // Export file as pdf.
      if(app.documents.length != 0){
        export_name = pdfName + "" + current_batch + ".pdf";
        myDocument.exportFile(ExportFormat.pdfType, File(export_name), false);
      }
      // Increment number.
      current_batch++;

      // Close document (without saving it).
      app.documents.everyItem().close(SaveOptions.NO);

      // Open document.
      var myDocument = app.open(File(documentName));
      // Get layer "Layer 1" (original template layer).
      var myLayer = myDocument.layers.item("Layer 1");
      // Get active page (page 1).
      var myPage1 = myDocument.pages.item(0);
      // Lay out the badges.
      var myNewPage = myPage1;

    }

    // Create new page if needed.
    if ((position_in_page == 0) && (badge_counter % (nb_pages_per_batch * nb_columns * nb_lines) != 0)){
      // Copy page.
      var myNewPage = copyToNewPage(myLayer, myPage1, myDocument);
      // Actualize the first layer (top left).
      placeAttributes(person, myLayer, myNewPage, background_link);
    }
    else {
      // Place infos.
      placeAttributes(person, placeInGrid([current_column, current_line], myLayer, myNewPage), myNewPage, background_link);
    }
  }

  // Export file as pdf.
  export_name = pdfName + "" + current_batch + ".pdf";
  myDocument.exportFile(ExportFormat.pdfType, File(export_name), false);

  // Close document (without saving it).
  // app.activeDocument.close(SaveOptions.no);

}



////////////////////////////////////////////////////////////////////////////////
// Definition of useful functions.
////////////////////////////////////////////////////////////////////////////////

/**
  Read the given csv file and returns array of dictionaries
    info_entreprises = [{"Entreprise" : "Michelin", "Role" : "blablabla", "Nom" : ...}, ...]
*/
function infoFromCsv(link_to_csv, csv_delimiter){
  // Open csv file.
  var csvFile = File(link_to_csv);
  var Read1 = csvFile.open("r",undefined,undefined);

  // Get attributes and number of attributes.
  var firstLine = csvFile.readln().split(csv_delimiter);
  var numberOfAttributes = firstLine.length;

  // Extract info from file.
  var info_entreprises = [];
  while(csvFile.eof == false)
  {
      // Read line. line is the array of elements.
      var line = csvFile.readln().split(csv_delimiter);
      // Build dictionary.
      var dictionary = {};
      for (var counter = 0; counter < numberOfAttributes; counter++) {
        dictionary[firstLine[counter]] = line[counter];
      }
      info_entreprises.push(dictionary);
  }
  // Return array of dictionaries.
  return info_entreprises;

}


/**
  Get group that has the required name, in the given layer, and in the given page.
*/
function getGroupByName(name, layer, page){
  // List all groups in layer.
  allGroups = layer.allPageItems;
  for (var counter = 0; counter < allGroups.length; counter++) {
    // We return the group that has the correct name in the correct page.
    if (allGroups[counter].name == name && allGroups[counter].parentPage == page) {
      return allGroups[counter];
    }
  }
  // If the group is not found, return null.
  return null;
}

/**
  Set the content of the frame_nb-th text frame in the given group.
  Resizes in case of need.
*/
function setTextFrameContentInGroup(frame_nb, group, content){
  // Get considered text frame.
  var text_frame = group.textFrames[frame_nb];
  // Set font.
  // text_frame.parentStory.appliedFont = app.fonts.item("Lato");
  // Initialize contents.
  while (text_frame.contents.overflows) {
    text_frame.contents = "";
  }
  text_frame.contents = "";
  // Set contents.
  text_frame.contents = content;
  // Shrink size if overset. There are three possible sizes : 12 pts, 10 pts and 8 pts.
  // var font_size = [12, 10, 8];
  // for (var counter = 0; counter < font_size.length; counter++){
  //   // Set current size.
  //   var current_font_size = font_size[counter];
  //   text_frame.parentStory.pointSize = current_font_size;
  //   // If text not overset, break loop.
  //   if (! text_frame.overflows){
  //     // alert("Final font size for " + content + " : " + current_font_size);
  //     break;
  //   }
  // }
}

/**
  Copy original layer to new page. Copy the items that are in the given page.
  Returns the new page.
*/
function copyToNewPage(layer, page, document){
  if(layer.pageItems.length>0){
      // Get all items that are in the active page.
      var pageItemsOnLayer = layer.pageItems;
      var pageItemsOnLayerInActivePage = [];
      for (var counter = 0; counter < pageItemsOnLayer.length; counter++){
        var item = pageItemsOnLayer[counter];
        if (item.parentPage == page) {
          pageItemsOnLayerInActivePage.push(item);
        }
      }
      // Unlock every page item on the layer, and make all page items visible on the layer.
      for (var counter = 0; counter < pageItemsOnLayerInActivePage.length; counter++){
        var item = pageItemsOnLayerInActivePage[counter];
        item.locked = false;
        item.visible = true;
      }
      // Copy filtered items in new page.
      var myGroup = document.groups.add(pageItemsOnLayerInActivePage);
      var newPage = document.pages.add();
      var myDupOfGroup = myGroup.duplicate(newPage);
      myDupOfGroup.ungroup();
      myGroup.ungroup();
  };
  return newPage;
}

/**
  Get length and width of the whole selection by looping over all its elements.
  Returns array [length, height].
*/
function getSizeOfSelection(){
  var min_top_left_y = Infinity;
  var min_top_left_x = Infinity;
  var max_bottom_right_y = - Infinity;
  var max_bottom_right_x = - Infinity;
  for (var counter = 0; counter < app.selection.length; counter++) {
    current_bounds = app.selection[counter].geometricBounds;
    var min_top_left_y = Math.min(min_top_left_y, current_bounds[0]);
    var min_top_left_x = Math.min(min_top_left_x, current_bounds[1]);
    var max_bottom_right_y = Math.max(max_bottom_right_y, current_bounds[2]);
    var max_bottom_right_x = Math.max(max_bottom_right_x, current_bounds[3]);
  }
  var length_of_selection = max_bottom_right_x - min_top_left_x;
  var height_of_selection = max_bottom_right_y - min_top_left_y;
  return [length_of_selection, height_of_selection];
}

/**
  Copy the reference layer in a grid at the place given by the corresponding coordinates, in the given page.
  Returns the new layer.
*/
function placeInGrid(coordinates, reference_layer, page){
  // Get coordinates.
  m = coordinates[0] // Horizontal coordinate.
  n = coordinates[1] // Vertical coordinate.
  // Duplicate layer.
  var myNewLayer = reference_layer.duplicate();
  // Define new layer as active layer.
  app.activeDocument.activeLayer = myNewLayer;
  // Get page items of first order.
  var page_items = myNewLayer.pageItems;
  // Iterate through thisLayer.allPageItems and check if its parentPage property is equal to thisPage.
  // We keep the page items that are in the active page and hide those in the other pages.
  to_select = [];
  for (var counter = 0; counter < page_items.length; counter ++){
    var item = page_items[counter];
    if (item.parentPage == page) {
      to_select.push(item);
    }
    else{item.visible = false;}
  }
  // Select all page items.
  app.select(to_select);
  // Get size of selection.
  size_of_selection = getSizeOfSelection();
  length_of_selection = size_of_selection[0];
  height_of_selection = size_of_selection[1];
  // Change position of all page items.
  for (var counter = 0; counter < app.selection.length; counter++) {
    current_bounds = app.selection[counter].geometricBounds;
    var top_left_y = current_bounds[0];
    var top_left_x = current_bounds[1];
    var bottom_right_y = current_bounds[2];
    var bottom_right_x = current_bounds[3];
    app.selection[counter].geometricBounds = [top_left_y + n * height_of_selection, top_left_x + m * length_of_selection, bottom_right_y + n * height_of_selection, bottom_right_x + m * length_of_selection];
  }
  // Return layer.
  return myNewLayer;
}

/**
  Change the text frames in the layer according to the attributes of the person, in the given page.
*/
function placeAttributes(person, layer, page, background_link){
  // Get rectangle.
  for (var rectangle_counter = 0; rectangle_counter < layer.rectangles.length; rectangle_counter++){
    var rectangle = layer.rectangles.item(rectangle_counter)
    if (rectangle.parentPage.name == page.name){break}
  }
  // Place image.
  rectangle.place(background_link, false);
  // rectangle.fit(FitOptions.FILL_PROPORTIONALLY);
  rectangle.fit(FitOptions.centerContent);
  // Change text corresponding to attributes.
  for (var attribute in person){
    // Get group corresponding to attribute in the given page.
    myGroup = getGroupByName(attribute, layer, page);
    // If group is valid, change content.
    if (myGroup != null) {
      try{
        setTextFrameContentInGroup(0, myGroup, person[attribute]);
      }
      catch(e) {}
    }
  }
}
