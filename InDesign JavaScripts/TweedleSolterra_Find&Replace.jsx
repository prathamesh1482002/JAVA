var book = app.books[0];
var bookComps = book.bookContents;
var documentsToClose = [];

for (var j = 0; j < bookComps.length; j++) {
    myDoc = app.open(bookComps[j].fullName);
    $.writeln(myDoc.name);
    var k = 0;
    var ref = [];
    var pg = [];

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findChangeTextOptions.includeLockedLayersForFind = true;
    app.findChangeTextOptions.includeLockedStoriesForFind = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.wholeWord = true;

    findTitle(ref, pg, myDoc);
    replaceText(ref, pg);
    documentsToClose.push(myDoc);
}

// Close and save all opened documents
for (var i = 0; i < documentsToClose.length; i++) {
    var doc = documentsToClose[i];
    doc.save();
    doc.close(SaveOptions.yes);
}

book.save();

alert("Process Done");

function findTitle(ref, pg, myDoc) {
    var page = 0;
    while(myDoc.pages[page].isValid){
        app.select(myDoc.pages[page]);
        
    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findChangeTextOptions.includeLockedLayersForFind = true;
    app.findChangeTextOptions.includeLockedStoriesForFind = true;
    app.findChangeTextOptions.includeHiddenLayers = false;
    app.findChangeTextOptions.includeMasterPages = false;
    app.findChangeTextOptions.includeFootnotes = true;
    app.findChangeTextOptions.caseSensitive = false;
    app.findChangeTextOptions.wholeWord = false;
    var txtF = myDoc.pages[page].allPageItems;
    var t = 0;
    var textFrame = txtF[t];
    if(txtF.length>0){
        app.select(textFrame);
        var textColumn = textFrame.textColumns;
         for(var T = 0; T < textColumn.length; T++){
             var para = textColumn[T].paragraphs;
             for(var p = 0; p < para.length; p++){
                 var pwrds = para[p].words;
                 for(var w=0; w < pwrds.length; w++)
                 {
                     if(pwrds[w].appliedCharacterStyle.name == "Page Reference"){
                      var string_1 = pwrds[w].contents;
                      var page_reference = string_1.split('·')[1];
                      var pagenumber = myDoc.pages[page].name;
                      ref.push(page_reference);
                      pg.push(pagenumber);
                      app.select(pwrds[w]);
                      k++;
                      }
                     }
                 }
             }
        }
    page++;
    }
}

function replaceText(ref, pg) {
     if(ref.length > 0){
        var bookComp = book.bookContents;
        for (var l = 0; bookComp.length > l; l++) {
            myDocs = app.open (bookComp[l].fullName);
            app.findTextPreferences = NothingEnum.nothing;
            app.findChangeTextOptions.includeLockedLayersForFind = true;
            app.findChangeTextOptions.includeLockedStoriesForFind = true;
            app.findChangeTextOptions.includeHiddenLayers = false;
            app.findChangeTextOptions.includeMasterPages = false;
            app.findChangeTextOptions.includeFootnotes = true;
            app.findChangeTextOptions.caseSensitive = false;
            app.findChangeTextOptions.wholeWord = true;
            app.findTextPreferences.appliedCharacterStyle = "pagenumber";
            
            var refLength = ref.length;
            for(var v=0; v < refLength; v++){
                app.findTextPreferences.findWhat = ref[v];
                app.changeTextPreferences.changeTo = "P." + pg[v];
                myDocs.changeText();
            }
       //  myDocs.close(SaveOptions.yes);
            }
        }
}
