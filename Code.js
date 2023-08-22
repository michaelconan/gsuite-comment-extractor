/**
 * @module Gsuite Comment Extractor and Updater
 * @author Michael Conan
 * @version 1.0.3
 * @date 5/28/2020
 * Apps Script to leverage Google Drive Comments API to extract and summarize comments in a manageable format,
 * then provide user the ability to reply directly from the spreadsheet.
 * 
 */

/**
 * Simple trigger to create menu from which to execute script functions
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  .createMenu('Comments')
  .addItem('Get Comments', 'getComments')
  .addItem('Respond', 'respond')
  .addToUi()
}

// GLOBALS
var ss = SpreadsheetApp.getActive().getSheetByName('Comments');
var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
var heads = ['Document Name', 'Response', 'Action', 'Status', 'Comment', 'Author', 'Created', 'Modified', 'Context', 'Link'];

/**
 * Primary function to extract and summarize comment data based on user-provided
 * document URLs.
 */
function getComments() {
  
  // Get file IDs from spreadsheet
  var dcLnks = ss.getRange(1,2,1,50).getValues()[0].filter(Boolean);
  var dcIds = [];
  var typeErr = [];
  for (z=0;z< dcLnks.length;z++) {
    // Add document IDs to list, ignore values without IDs
    var idMatch = String(dcLnks[z]).match(/[-_\w]{25,}/)
    if (idMatch == null) {
      continue;
    } else {
      var file = DriveApp.getFileById(idMatch[0]);
      if (file.getMimeType().indexOf('google')+1 | file.getMimeType().indexOf('pdf')+1) {
        dcIds.push(file.getId());
      } else {
        typeErr.push(file.getName() + ': ' + file.getMimeType());
      }
    }
  }
  
  Logger.log(dcLnks)
  // Clear spreadsheet data
  ss.getRange(4,1,ss.getMaxRows(),ss.getMaxColumns()).clear();
  ss.getRange(4,1,1,heads.length).setValues([heads]);
  var flt = ss.getFilter();
  if (flt != null) {
    flt.remove();
  }
  if (ss.getRange(2,2).getValue() == 'Yes') {
    var del = 'true';
    }
  else {
    var del = 'false';
    }
  
  // Loop through documents to get comments
  var q = 0;
  var rpMx = 0;
  for (let x = 0; x< dcIds.length; x++) {
    
    // Get list of comments from file and set list of information to extract
    var fl = DriveApp.getFileById(dcIds[x]);
    var cmt = Drive.Comments.list(dcIds[x],{includeDeleted:del,pageSize:100});
    var cmts = cmt.items;
    while (cmt.nextPageToken != undefined) {
      cmt = Drive.Comments.list(dcIds[x],{includeDeleted:del,pageSize:100,pageToken:cmt.nextPageToken});
      for (i=0;i< cmt.items.length;i++) {
        cmts.push(cmt.items[i]);
      }
    }
    
    // Loop through comments and add information to list
    for (i = 0; i< cmts.length; i++) {
      var cmtAr = [];
      cmtAr.push(fl.getUrl(),
                 '',
                 '',
                 cmts[i].status,
                 cmts[i].content);
      try {
        cmtAr.push(cmts[i].author.displayName);
      }
      catch (e) {
        cmtAr.push("Non-PwC User");
      }
      cmtAr.push(Utilities.formatDate(new Date(cmts[i].createdDate),tz,"MM/dd/yyyy HH:mm:ss"),
                 Utilities.formatDate(new Date(cmts[i].modifiedDate),tz,"MM/dd/yyyy HH:mm:ss"));
      try {
        cmtAr.push(cmts[i].context.value);
      }
      catch (e) {
        cmtAr.push('n/a');
      }
      
      // Add comment link to list
      cmtAr.push(dcLnks[x].split(dcIds[x])[0]+dcIds[x]+'/edit?disco='+cmts[i].commentId);
      
      // Loop through each reply and add relevant fields
      var cmtReps = cmts[i].replies;
      for (r = 0; r< cmtReps.length; r++) {
        cmtAr.push(cmtReps[r].content);
        try {
          cmtAr.push(cmtReps[r].author.displayName);
        }
        catch (e) {
          cmtAr.push('n/a');
        }
        cmtAr.push(Utilities.formatDate(new Date(cmtReps[r].createdDate),tz,"MM/dd/yyyy HH:mm:ss"),
                   Utilities.formatDate(new Date(cmtReps[r].modifiedDate),tz,"MM/dd/yyyy HH:mm:ss"));
      }
      
      // Check number of replies to update headers for all replies
      var rep = [];
      if (cmtAr.length > 7) {
        for (j = 0;j< (cmtAr.length-7)/4;j++) {
          r = j+1;
          rep.push('Reply '+r,'Reply '+r+' Author','Reply '+r+' Created','Reply '+r+' Modified');
        }
        var rAr = new Array(1);
        rAr[0] = rep;
        ss.getRange(4,11,1,rep.length).setValues(rAr).setFontWeight('bold');
        if (rep.length > rpMx) {
          rpMx = rep.length;
        }
      }
      
      // Write comment details to spreadsheet
      var cmtDets = new Array(1);
      cmtDets[0] = cmtAr;
      ss.getRange(5+q,1,1,cmtAr.length).setValues(cmtDets).setBorder(true, true, true, true, true, true);
      var lk1 = ss.getRange(5+q,1).getValue()
      ss.getRange(5+q,1).setFormula('=HYPERLINK("'+lk1+'","'+fl.getName()+'")')
      var lnk = ss.getRange(5+q,10).getValue()
      ss.getRange(5+q,10).setFormula('=HYPERLINK("'+lnk+'","Comment'+cmts[i].commentId+'")')
      
      // Increment comment row
      q+= 1;
    }
  }
  
  // Update formatting
  if (q > 0) {
    var dv = SpreadsheetApp.newDataValidation().requireValueInList(['Yes','No']);
    ss.getRange(5, 1, q, rpMx+10).setBorder(true, true, true, true, true, true);
    ss.getRange(5, 3, q, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Resolve','Reopen']));
    //ss.getRange(5, 3, q, 1).setBackground('grey');
    ss.getRange(4, 1, q+1, rpMx+10).createFilter();
  }
  
  if (typeErr.length > 0) {
    SpreadsheetApp.getUi().alert('The following files were of unsupported file types: \n\n' + typeErr.join('\n'));
  }
}

/**
 * Subsequent functionality to respond to the summarized comments directly from the spreadsheet
 */
function respond() {
  
  var cmtDat = ss.getRange(5, 1, ss.getDataRange().getNumRows()-4, 12).getValues();
  
  // For each comment, apply response as specified
  for (i = 0;i< cmtDat.length;i++) {
    if (cmtDat[i][1] != '') {
      var fid = String(ss.getRange(i+5,10).getFormula()).split('"')[1].split('/')[5];
      var cid = String(ss.getRange(i+5,10).getFormula()).split('"')[1].split('=')[1];
      var action = cmtDat[i][2];
      
      // Create Javascript Object for reply
      var rp = {
        'resource':{
          'content':cmtDat[i][1]
        }
      }
      if (action != '') {
        rp.resource.verb = String(action).toLowerCase();
      }
      Drive.Replies.insert(Drive,fid,cid,rp);
    }
  }
}