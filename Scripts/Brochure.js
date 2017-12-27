////////////////////////////////////////////////////////////////////////////////
// Main content of the script.
// User interface proposes 3 functions :
//  - Layout pages from csv and export in folders.
//  - Layout whole document from folders.
//  - Update current document (index + logos).
// Il est nécessaire d'écrire 'é' dans ce header pour qu'indesign reconnaisse
// les caractères francais dans les layers.
////////////////////////////////////////////////////////////////////////////////

// Variable definitions.
// Link to the working directory.
const LINK = "C:\\Path\\To\\Directory";

// Link to the page model, index model, and intro.
const PAGE_MODEL = LINK.concat("\\ModèlePageEntreprise.indd");
const INDEX_MODEL = LINK.concat("\\Indexes.indd");
const INTRO = LINK.concat("\\Intro.indd");

// Define used text configurations. The script will loop over them in order to prevent overlapping text frames.
const TEXT_CONF = [{"font" : app.fonts.item("Lato"), "size" : 12 }, {"font" : app.fonts.item("Lato"), "size" : 10 }, {"font" : app.fonts.item("Lato"), "size" : 8 }];
const TITLE_CONF = [{"font" : app.fonts.item("Bebas Neue (OTF)"), "size" : 25 }, {"font" : app.fonts.item("Bebas Neue (OTF)"), "size" : 20 }, {"font" : app.fonts.item("Bebas Neue (OTF)"), "size" : 15 }];
const SUBTITLE_CONF = [{"font" : app.fonts.item("Lato"), "size" : 10 }, {"font" : app.fonts.item("Lato"), "size" : 8 }];

// Define array [[csv_LINK, corresponding_master_name, corresponding_folder_to_save_files]] for the different sectors.
const BA ={"csv" : [LINK.concat("\\Données CSV\\6-brochure.csv")], "master" : "BA-Gabarit", "folder" : LINK.concat("\\Pages Exportées\\BA"), "name" : "Banques et Assurances", "logos" : [LINK.concat("\\Logos\\Logos-6")]};
const AC ={"csv" : [LINK.concat("\\Données CSV\\7-brochure.csv")], "master" : "AC-Gabarit", "folder" : LINK.concat("\\Pages Exportées\\AC"), "name" : "Audit et conseil", "logos" : [LINK.concat("\\Logos\\Logos-7")]};
const IST ={"csv" : [LINK.concat("\\Données CSV\\8-brochure.csv")], "master" : "IST-Gabarit", "folder" : LINK.concat("\\Pages Exportées\\IST"), "name" : "Industries, Informatique et Télécom", "logos" : [LINK.concat("\\Logos\\Logos-8")]};
const PME ={"csv" : [LINK.concat("\\Données CSV\\10-brochure.csv")], "master" : "PME-Gabarit", "folder" : LINK.concat("\\Pages Exportées\\PME"), "name" : "Village PME", "logos" : [LINK.concat("\\Logos\\Logos-10")]};
const ECUI ={"csv" : [LINK.concat("\\Données CSV\\9-brochure.csv"), LINK.concat("\\Données CSV\\11-brochure.csv")] , "master" : "ECUI-Gabarit", "folder" : LINK.concat("\\Pages Exportées\\ECUI"), "name" : "Écoles, Corps et Universités", "logos" : [LINK.concat("\\Logos\\Logos-9"), LINK.concat("\\Logos\\Logos-11")]};

// Define arrays of all sectors.
const ARRAY_SECTEURS = [BA, AC, IST, PME, ECUI];


// User interface.
main();
function main(){
    setup();
    snippet();
}
function setup(){}
function snippet(){
  var dialog = app.dialogs.add({name:"Brochure generation", canCancel:true});
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
					var function_choice = dropdowns.add({stringList:["Layout pages from csv and export in folders", "Layout whole document from folders", "Update current document (index + logos)", "All"], selectedIndex:0});
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
      // Layout the pages from the csv document and export in folders.
			layoutPages();
		}
		else if(function_choice.selectedIndex == 1){
      // Layout the whole document from the folders.
			layoutDocument();
		}
    else if(function_choice.selectedIndex == 2){
      // Update the current document (indexes + logos).
      updateDocument();
    }
		else {
      // Do all of the above.
			layoutPages();
      layoutDocument();
      updateDocument();
		}
		dialog.destroy();
	}
	else{
		dialog.destroy()
	}
}


////////////////////////////////////////////////////////////////////////////////
// The main functions.
////////////////////////////////////////////////////////////////////////////////

/**
  Layout the different pages corresponding to the given secteur data.
*/
function layoutPages(){
  for (var secteur_counter = 0; secteur_counter < ARRAY_SECTEURS.length; secteur_counter++){
    var master_name =ARRAY_SECTEURS[secteur_counter]["master"];
    var folder_to_save = ARRAY_SECTEURS[secteur_counter]["folder"];
    for (csv_counter = 0; csv_counter < ARRAY_SECTEURS[secteur_counter]["csv"].length; csv_counter++){
      var csv_LINK = ARRAY_SECTEURS[secteur_counter]["csv"][csv_counter];
      layoutFromCsv(File(PAGE_MODEL), csv_LINK, master_name, folder_to_save);
    }
  }
}

/**
  Layout the whole document with all secteurs. (Does not close document).
*/
function layoutDocument(){
  // Open document.
  var brochure = app.open(File(INTRO));
  // Open file of indexes.
  var indexes = app.open(File(INDEX_MODEL));

  // Loop over secteurs.
  for (var secteur_counter = 0; secteur_counter < ARRAY_SECTEURS.length; secteur_counter++){
    // Initialize useful variables.
    var folder_to_save = ARRAY_SECTEURS[secteur_counter]["folder"];
    var secteur_name = ARRAY_SECTEURS[secteur_counter]["name"];

    // Get the list of generated pages in the folder.
    var list_pages = getFilesInFolder(folder_to_save);

    // Add index.
    indexes.spreads.item(secteur_counter + 1).duplicate(LocationOptions.AT_END, brochure.spreads.item(-1));
    // Get index text frame (TODO: do the work in 2 consecutive loops instead?).
    var index_text_frame = getIndexTextFrames(brochure)[secteur_counter];

    // Initialize content for index.
    var list_of_companies_to_index = [];

    // Loop over the files in the folder.
    for (var counter_file = 0; counter_file < list_pages.length; counter_file++){
      // Open generated file.
      var company_page = app.open(File(list_pages[counter_file]));
      // duplicate page in own document.
      company_page.pages.item(0).duplicate(LocationOptions.AFTER, company_page.pages.item(0));
      // Move duplicated page to new document.
      company_page.pages.item(-1).move(LocationOptions.AFTER, brochure.pages.item(-1), BindingOptions.DEFAULT_VALUE);
      // Get name of company and actualize content for index.
      var company_name = getGroupByName("Entreprise", company_page.layers.item("Upper layer"), company_page.pages.item(0)).textFrames[0].contents;
      list_of_companies_to_index.push({"company" : company_name, "page" : brochure.pages.item(-1).name});
      // Close generated file without saving.
      company_page.close(SaveOptions.no);
    }

    // Even out page numbers if needed.
    if (list_pages.length % 2 != 0) {
      var blank_page = brochure.pages.item(0);
      // duplicate page.
      blank_page = blank_page.duplicate(LocationOptions.AFTER, brochure.pages.item(-1));
      // Move page.
      blank_page.move(LocationOptions.AFTER, brochure.pages.item(-1), BindingOptions.DEFAULT_VALUE);
    }

    // Write index.
    writeIndexTextFrame(index_text_frame, secteur_name, list_of_companies_to_index);
  }

  // Export file as pdf.
  app.pdfExportPreferences.exportReaderSpreads = true;
  brochure.exportFile(ExportFormat.pdfType, File(LINK.concat("\\Brochure_Generated_Layout.pdf")), false);
  // Close index document.
  indexes.close(SaveOptions.no);
}

/**
  Update the whole document.
  Does 2 things:
    - Updates indexes.
    - Updates logos.
*/
function updateDocument(){
  // Only works if there is an active document.
  if (app.documents.length != 0) {
    // Get document.
    var brochure = app.activeDocument;

    // Get indexes pages i.e pages that have a "Index" group.
    var index_text_frames = getIndexTextFrames(brochure);

    // Process each index text frame :
    for (var index_counter = 0; index_counter < index_text_frames.length; index_counter++){
      // Get current secteur.
      var secteur_name = ARRAY_SECTEURS[index_counter]["name"];

      // Get current logo list.
      var logo_sources = ARRAY_SECTEURS[index_counter]["logos"];
      var logo_images = [];
      for (var sources_counter = 0; sources_counter < logo_sources.length; sources_counter++){
        logo_images = logo_images.concat(getFilesInFolder(logo_sources[sources_counter]));
      }

      // Initialize content for index.
      var list_of_companies_to_index = [];

      // Get current index text frame.
      var current_index_text_frame = index_text_frames[index_counter];

      // Define window of search. Last page to search is either the next index
      // or the end of the document.
      var first_page_to_process = parseFloat(current_index_text_frame.parentPage.name);
      if (index_counter == index_text_frames.length - 1){
        var last_page_to_process = brochure.pages.length - 1;
      }
      else {
        var last_page_to_process = index_text_frames[index_counter + 1].parentPage.name;
      }

      // Loop over the pages to search.
      for (var page_to_process = first_page_to_process; page_to_process < last_page_to_process; page_to_process++){
        // FIRST PART OF THE LOOP : Indexes.
        // Try to find group named "Entreprise".
        var wanted_group = getGroupByName("Entreprise", brochure.layers.item("Upper layer"), brochure.pages.item(page_to_process))
        if (wanted_group != null) {
          // If group found, actualize index list.
          var company_name = wanted_group.textFrames[0].contents;
          list_of_companies_to_index.push({"company" : company_name, "page" : page_to_process + 1});
        }
        // Center the "Message aux étudiants".
        var message_group = getGroupByName("Message aux étudiants", brochure.layers.item("Upper layer"), brochure.pages.item(page_to_process))
        if (message_group != null) {
          // Center text.
          // message_group.textFrames[0].parentStory.texts[0].select();
          // app.selection[0].parent.textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
          message_group.textFrames[0].textFramePreferences.verticalJustification = VerticalJustification.CENTER_ALIGN;
        }


        // SECOND PART OF THE LOOP : Logos.
        var logo_image = null;
        if (wanted_group != null){
          // Get company name.
          var company_name = wanted_group.textFrames[0].contents;
          // Get image to apply.
          var str = []
          // Loop over all file names.
          for (var logo_counter = 0; logo_counter < logo_images.length; logo_counter++){
            var file_name = String(logo_images[logo_counter]).split("/");
            file_name = file_name[file_name.length - 1].split(".");
            if (decodeURI(file_name[0]) == company_name) {
              var logo_image = new File(logo_images[logo_counter]);
            }
          }

          // Insert logo if found.
          if (logo_image != null){
            insertLogo(logo_image, page_to_process, brochure);
          }
        }
      }

      // Write index.
      writeIndexTextFrame(current_index_text_frame, secteur_name, list_of_companies_to_index);
    }
  // Finally, export pdf.
  brochure.exportFile(ExportFormat.pdfType, File(LINK.concat("\\Brochure_Generated_Current.pdf")), false);
  }
  else{
    alert("Please open a document to update.")
  }


}


////////////////////////////////////////////////////////////////////////////////
// The side functions.
////////////////////////////////////////////////////////////////////////////////
/**
  Get the list of all files in a the given folder link.
*/
function getFilesInFolder(folder_link){
  // Get to folder.
  var folder = new Folder(folder_link);

  // Optional code: get list of all file names.
  // var str = [];
  // for (var counter = 0; counter < list.length; counter++){
  //   var file_name = String(list[counter]).split("/");
  //   file_name = file_name[file_name.length - 1].split(".");
  //   if (file_name[1] == "indd") {str.push(file_name[0]);}
  // }

  return folder.getFiles();
}

/**
  Read the given csv file and returns array of dictionaries
    info_entreprises = [{"Entreprise" : "Michelin", "Texte de présentation" : "blablabla", ...}, ...]
*/
function infoFromCsv(link_to_csv){
  // Open csv file.
  var csvFile = File(link_to_csv);
  var Read1 = csvFile.open("r",undefined,undefined);

  // Get attributes and number of attributes.
  var firstLine = csvFile.readln().split("#");
  var numberOfAttributes = firstLine.length;

  // Extract info from file.
  var info_entreprises = [];
  while(csvFile.eof == false)
  {
      // Read line. line is the array of elements.
      var line = csvFile.readln().split("#");
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
function setTextFrameContentInGroup(frame_nb, group, content, text_configurations){
  // Get considered text frame.
  var text_frame = group.textFrames[frame_nb];
  // Shrink size if overset. There are three possible sizes : 12 pts, 10 pts and 8 pts.
  var font_size = [12, 10, 8];
  for (var counter = 0; counter < text_configurations.length; counter++){
    // Get current configuration.
    var current_configuration = text_configurations[counter];
    // Get font value and size.
    var current_font = current_configuration["font"];
    var current_font_size = current_configuration["size"];
    // Apply font.
    text_frame.parentStory.appliedFont = current_font;
    // Apply font size.
    text_frame.parentStory.pointSize = current_font_size;
    // Set contents.
    text_frame.contents = content;
    // If text not overset, break loop.
    if (! text_frame.overflows){
      // alert("Final font size for " + content + " : " + current_font_size);
      break;
    }
  }
}

/**
  Inserts text with title in the given text frame.
  The user inputs the font size for the title and for the body of the text.
*/
function insertTextWithTitle(text_frame, title, text, title_font_size, body_font_size){
  if (text != ""){
    // Insert title.
    var insertion_point = text_frame.insertionPoints.item(text_frame.insertionPoints.length - 1);
    insertion_point.fontStyle = "Bold";
    insertion_point.pointSize = title_font_size;
    insertion_point.contents = "\r\r" + title;

    // Insert text.
    var insertion_point = text_frame.insertionPoints.item(text_frame.insertionPoints.length - 1);
    insertion_point.fontStyle = "Regular";
    insertion_point.pointSize = body_font_size;
    insertion_point.contents = "\r" + text;
  }
}

/**
  Sets the presentation content in the corresponding text frame of the given group.
  The user inputs font size for the title and for the body of the text.
*/
function setPresentationContentWithSize(frame_nb, group, company_info, title_font_size, body_font_size){
  // Get considered text frame.
  var text_frame = group.textFrames[frame_nb];
  // Set font and font size.
  text_frame.parentStory.appliedFont = app.fonts.item("Lato");
  text_frame.parentStory.pointSize = body_font_size;
  // Initialize contents.
  while(text_frame.overflows){text_frame.contents = "";}
  text_frame.contents = "";
  // Insert texte de présentation.
  text_frame.insertionPoints.item(0).contents = company_info["Texte de présentation"];
  // Insert rest of content.
  insertTextWithTitle(text_frame, "Profil requis", company_info["Profil requis"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Type de poste", company_info["Type de poste"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Processus de recrutement", company_info["Processus de recrutement"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Perspectives de carrière", company_info["Perspectives de carrière"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Possibilités en interne", company_info["Possibilités en interne"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Salaire", company_info["Salaire"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Informations complémentaires 1", company_info["Informations complémentaires 1"], title_font_size, body_font_size);
  insertTextWithTitle(text_frame, "Informations complémentaires 2", company_info["Informations complémentaires 2"], title_font_size, body_font_size);
  // Return text frame.
  return text_frame;
}

/**
  Sets the presentation content in the corresponding text frame of the given group.
  We try two possibilities of font sizes to avoid text overflowing.
*/
function setPresentationContentInGroup(frame_nb, group, company_info){
  // We try first with title_font_size = 12, body_font_size = 10.
  var text_frame = setPresentationContentWithSize(frame_nb, group, company_info, 12, 10);
  // Else title_font_size = 10, body_font_size = 8.
  if (text_frame.overflows) {
    setPresentationContentWithSize(frame_nb, group, company_info, 10, 8);
  }
}

/**
  Layout all the given company information in the page.
*/
function layoutInformation(company_info, layer, page){
  /**
  Structure of the content:

  -Entreprise
  -Secteur d'activité
  -Chiffres:
  	-Effectif
  	-Implantation:
  		-Implantation en France
  		-Implantation à l'étranger
  	-Chiffre d'affaire
  -Texte de présentation:
  	-Profil requis
  	-Type de poste
  	-Processus de recrutement
  	-Perspectives de carrière
  	-Possibilités en interne
  	-Salaire
  	-Informations complémentaires 1
  	-Informations complémentaires 2
  -Message aux étudiants:
    -Email
  */

  // Info that does not require further consideration : entreprise, secteur d'activité, effectif, chiffre d'affaires.
  setTextFrameContentInGroup(0, getGroupByName("Entreprise", layer, page), company_info["Entreprise"], TITLE_CONF);
  setTextFrameContentInGroup(0, getGroupByName("Secteur d'activité", layer, page), company_info["Secteur d'activité"], SUBTITLE_CONF);
  setTextFrameContentInGroup(0, getGroupByName("Effectif", layer, page), company_info["Effectif"], TEXT_CONF);
  setTextFrameContentInGroup(0, getGroupByName("Chiffre d'affaire", layer, page), company_info["Chiffre d'affaire"], TEXT_CONF);

  // Implantation: combine France + abroad.
  if (company_info["Implantation en France"] == "") {var content_implantation = company_info["Implantation à l'étranger"];}
  else {var content_implantation = company_info["Implantation en France"] + "\n" + company_info["Implantation à l'étranger"];}
  setTextFrameContentInGroup(0, getGroupByName("Implantation", layer, page), content_implantation, TEXT_CONF);

  // Message aux étudiants: combine with email.
  if (company_info["Email"] == "") {var content_message = company_info["Message aux étudiants"]}
  else {var content_message = company_info["Message aux étudiants"] + "\n" + company_info["Email"];}
  setTextFrameContentInGroup(0, getGroupByName("Message aux étudiants", layer, page), content_message, TEXT_CONF);

  // Texte de présentation.
  setPresentationContentInGroup(0, getGroupByName("Texte de présentation", layer, page), company_info);
}

/**
  Layout the different pages according to the information given in the .csv file.
  Saves them in the given folder with filename  "company_name.indd"
  Returns the list of generated files.
*/
function layoutFromCsv(template_file, csv_link, master_name, folder_to_save){
  // Create new folder for saving files if it does not exist.
  var f = new Folder(folder_to_save);
  if (!f.exists)
      f.create();
  // Get array of informations.
  var info_entreprises = infoFromCsv(csv_link);
  // Initialize array of generated files.
  var generated_files = [];
  // Layout the pages in different files.
  for (var company_counter = 0; company_counter < info_entreprises.length; company_counter++){
    // Check if file already exists.
    var save_link = folder_to_save.concat("\\" + info_entreprises[company_counter]["Entreprise"] + ".indd");
    var myFile = new File(save_link);
    if (!myFile.exists) {
      // alert("new file:" + save_link);
      // Open document.
      var myDocument = app.open(template_file);

      // Apply corresponding master to document.
      var masters = myDocument.masterSpreads;
      var master = masters.item(master_name);
      myDocument.pages.item(0).appliedMaster = master;

      // Get layer "Upper layer" (original template layer).
      var myLayer = myDocument.layers.item("Upper layer");
      // Get active page (page 1).
      var myPage1 = myDocument.pages.item(0);

      // Layout the information for the given company.
      layoutInformation(info_entreprises[company_counter], myLayer, myPage1);

      // Close document (saving it into new file) and actualize array of generated files.
      app.activeDocument.save(new File(save_link));
      generated_files.push(save_link);
      app.activeDocument.close(SaveOptions.no);
      }
  }
  // Return array of generated files.
  return generated_files;
}

/**
  Returns the list of index pages in the brochure.
  It loops over the pages of the document and returns only pages that have a group named "Index" in the "Upper layer" layer.
*/
function getIndexTextFrames(brochure){
  // Initialize list of index pages.
  var index_text_frames = [];
  // Loop over the pages of the document and return only pages that have a group named "Index" in the "Upper layer" layer.
  for (var page_counter = 0; page_counter < brochure.pages.length; page_counter++){
    var current_page = brochure.pages.item(page_counter);
    var index_group = getGroupByName("Index", brochure.layers.item("Upper layer"), current_page);
    if ( index_group != null) {
      // Actualize current page.
      index_text_frames.push(index_group.textFrames[0]);
    }
  }
  // Return list.
  return index_text_frames;
}

/**
  This funcion writes the index text frame based on the array
    list_of_companies = [{"company" : "Michelin", "page" : 29}]
  It also writes the given title.
*/
function writeIndexTextFrame(index_text_frame, title, list_of_companies){
  // Set contents of index to "".
  while (index_text_frame.overflows){index_text_frame.contents = "";}
  index_text_frame.contents = "";
  // Write index title.
  var insertion_point = index_text_frame.insertionPoints.item(index_text_frame.insertionPoints.length - 1);
  insertion_point.fontStyle = "Bold";
  insertion_point.pointSize = 14;
  insertion_point.contents = title;
  var insertion_point = index_text_frame.insertionPoints.item(index_text_frame.insertionPoints.length - 1);

  // Loop over the list of companies.
  for (var company_counter = 0; company_counter < list_of_companies.length; company_counter++){
    // If overfowing, increase number of columns.
    if (index_text_frame.overflows){index_text_frame.textFramePreferences.textColumnCount++;}
    // Actualize content.
    var company_name = list_of_companies[company_counter]["company"];
    var page_number = list_of_companies[company_counter]["page"];
    var insertion_point = index_text_frame.insertionPoints.item(index_text_frame.insertionPoints.length - 1);
    insertion_point.fontStyle = "Regular";
    insertion_point.pointSize = 12;
    insertion_point.contents = "\n" + company_name + "\tp." + page_number;
  }
}

/**
  Attemps to insert the logo at the given page number in the document.
*/
function insertLogo(logo, page_number, myDocument){
  try{
    // Override master items in the current page.
    OverrideMasterItems(myDocument, page_number);
    // Get group where to insert logo. It depends on the parity of the page.
    if (page_number % 2 == 0){
      // Right pages.
      var logo_group = getGroupByName("Droite", myDocument.layers.item("Upper layer"), myDocument.pages.item(page_number));
      var anchor_point = AnchorPoint.LEFT_CENTER_ANCHOR;
      var moving_values =  [1, 0];
    }
    else{
      var logo_group = getGroupByName("Gauche", myDocument.layers.item("Upper layer"), myDocument.pages.item(page_number));
      var anchor_point = AnchorPoint.RIGHT_CENTER_ANCHOR;
      var moving_values =  [-1, 0];
    }
    // Get polygon where to place logo.
    var polygon = logo_group.polygons[0];
    // Place image.
    var image = polygon.place(logo, false)[0];
    // Fit image to frame.
    image.flip = Flip.none;
    polygon.fit(FitOptions.PROPORTIONALLY);
    polygon.fit(FitOptions.centerContent);
    // Slightly rezise image.
    image.resize(CoordinateSpaces.INNER_COORDINATES, AnchorPoint.CENTER_ANCHOR, ResizeMethods.MULTIPLYING_CURRENT_DIMENSIONS_BY, [0.8, 0.8]);
    // // Slightly move image.
    // image.move(undefined, moving_values)
  }
  catch(e){}
}

/**
  Overrides all master items of the given page in the given document.
*/
function OverrideMasterItems(myDocument, CurrentPage) {
  try{
    // Define constants.
    const CS_INNER = +CoordinateSpaces.innerCoordinates,
          ORIGIN = [[0,0],CS_INNER];
    // Remove all overrides of the current page.
    myDocument.pages[CurrentPage].removeOverride();
    // Get all items.
    var allItems= myDocument.pages[CurrentPage].appliedMaster.pageItems.everyItem().getElements();
    // Loop over all items
    for(var i=0;i< allItems.length;i++){
      try{
        // if (CurrentPage == 7) {alert(allItems[i].name)};
        allItems[i].override(myDocument.pages[CurrentPage]);
        (mx=myDocument.pages[CurrentPage].properties.masterPageTransform) &&  myDocument.pages[CurrentPage].groups.item(allItems[i].name).transform(CS_INNER, ORIGIN, mx)
      }
      catch(e){}
    }
  }
  catch(e){}
}
