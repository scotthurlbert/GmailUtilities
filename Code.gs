// Gave me the idea:
// http://lifehacker.com/5986204/automatically-clean-up-gmail-on-a-schedule-with-this-script

// Showed the details about setting up the timer:
// http://www.johneday.com/422/time-based-gmail-filters-with-google-apps-script

function check_auto_delete_mails() 
{
  Logger.log( "Running: check_auto_delete_mails()" );
  autoDeleteMails( "AutoClean/AutoClean7", 7 );
  autoDeleteMails( "AutoClean/AutoClean14", 14 );
  autoDeleteMails( "AutoClean/AutoClean30", 30 );
  autoDeleteMails( "AutoClean/AutoClean180", 180 );
  autoDeleteMails( "AutoClean/AutoClean365", 365 );
}

// Inspired by:
// https://ctrlq.org/code/19040-gmail-size-search

function tag_Large_Gmail_Messages() 
{
  var sizeLimit = "2000000";                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                ;
  var start = new Date();
  Logger.log("Now finding all the big emails in your Gmail mailbox.");
  var label = GmailApp.getUserLabelByName("LargeEmails");
  if(label == null)
  {
    GmailApp.createLabel('LargeEmails');
  }
  else
  {
    // Find all Gmail messages that have attachments
    var threads = GmailApp.search('has:attachment larger:' + sizeLimit + ' -LargeEmails -label:saved-to-pdf -label:Lists-idesign-alumni -label:okayToBeLarge');
    if (threads.length == 0) 
    {
      return;
    }
    Logger.log("Search returned " + threads.length + " threads larger than " + sizeLimit + ".");
    for (var i=0; i<threads.length; i++) 
    {
      if (isTimeUp(start)) 
      {
        Logger.log("Time's up.  Will have to finish processing later.");
        break;
      }
      label.addToThread(threads[i]);
      // It's not needed to calculate the size, just let the search return the results.
      // I'm leaving this code in place for reference.
      //var messages = threads[i].getMessages();
      //for (var m=0; m<messages.length; m++)
      //{
      //  var size = getMessageSize(messages[m].getAttachments());      
      //  // If the total size of attachments is > MB limit, log the messages
      //  // You can change this value as per requirement.
      //  if (size >= parseInt(sizeLimit))
      //  {
      //  }
      //  else
      //  {
      //    Logger.log("Message not being added because it's not >= " + sizeLimit + ".  Actual Size: " + size);
      //  }
      //}
    }
  }
  Logger.log( "Done tagging large emails.  Finished within time limit." );
}

// Adapted from:
// https://www.maketecheasier.com/google-scripts-to-automate-gmail/

// This gave me the idea
// http://www.labnol.org/internet/send-gmail-to-google-drive/21236/
// https://chrome.google.com/webstore/detail/save-emails-and-attachmen/nflmnfjphdbeagnilbihcodcophecebc

function save_Gmail_as_PDF()
{
  var start = new Date();
  var label = GmailApp.getUserLabelByName("Save as PDF");  
  var labelAfter = GmailApp.getUserLabelByName("Saved To PDF");  
  if(labelAfter == null){
    GmailApp.createLabel('Saved To PDF');
  }
  if(label == null){
    GmailApp.createLabel('Save as PDF');
  }
  else
  {
    var threads = label.getThreads();
    Logger.log( "Saving " + threads.length + " emails to google drive." );
    for (var i = 0; i < threads.length; i++) 
    {
      if (isTimeUp(start) ) 
      {
        Logger.log("Time's up.  Will have to finish processing later.");
        break;
      }
      var messages = threads[i].getMessages();  
      var message = messages[0];
      var body    = message.getBody();
      var dateTime = message.getDate();
      var subject = dateTime.toISOString().substring(0, 19).replace(/:/g, '.').replace(/-/g, '') + ", " + message.getFrom() + ", " + message.getSubject();
      var attachments  = message.getAttachments();
      
      Logger.log( "Creating message: " + subject );
      
      var jsonTextFormatted = getInfoText( message, threads[i] );
      
      for(var j = 1;j<messages.length;j++)
      {
        body += messages[j].getBody();
        var temp_attach = messages[j].getAttachments();
        if(temp_attach.length>0)
        {
          for(var k =0;k<temp_attach.length;k++)
          {
            attachments.push(temp_attach[k]);
          }
        }
      } 
      // Create an HTML File from the Message Body
      var bodydochtml = DriveApp.createFile(subject+'.html', body, "text/html")
      var bodyId=bodydochtml.getId();
 
      // Convert the HTML to PDF
      var bodydocpdf = bodydochtml.getAs('application/pdf');
      if(attachments.length > 0)
      {
        var gmailFolders = DriveApp.getFoldersByName( "Gmail PDFs" );
        var gmailFolder = gmailFolders.next();
        var folder = gmailFolder.createFolder(subject)
        gmailFolder.addFolder(folder);
        for (var j = 0; j < attachments.length; j++) 
        {
          folder.createFile(attachments[j]);
          Utilities.sleep(1000);
        }
        folder.createFile(bodydocpdf);
        folder.createFile("info.json", jsonTextFormatted, "application/json" );
      }
      else
      {
        DriveApp.createFile(bodydocpdf);
      }      
      DriveApp.getFileById(bodyId).setTrashed(true);
      label.removeFromThread(threads[i]);
      labelAfter.addToThread(threads[i]);
    }
  }  
  Logger.log( "Done saving emails as PDF to google drive.  Finished within time limit." );
}

// Adapted from:
// https://ctrlq.org/code/20016-maximum-execution-time-limit

function isTimeUp(pStart) 
{
  var now = new Date();
  return now.getTime() - pStart.getTime() > 300000 - 60000; // 5 minutes - 1 minute safety buffer.
}
 
function autoDeleteMails( labelToClean, numberOfDays ) 
{  
  var label = GmailApp.getUserLabelByName( labelToClean );  
  if(label == null)
  {
    Logger.log( "Could not find the label: " + labelToClean + ". Creating..." );
    GmailApp.createLabel(labelToClean);
  }
  else
  {
    var mySearch = "'label:" + labelToClean + " older_than:" + numberOfDays + "d'";
    var batchSize = 100; // Process up to 100 threads at once
    while (GmailApp.search(mySearch, 0, 1).length == 1)
    {
      GmailApp.moveThreadsToTrash(GmailApp.search(mySearch, 0, batchSize));
    }  
  }
  // Logger.log( "done: auto_delete_mails( " + labelToClean + ", " + numberOfDays + " )" );
}

function getInfo( pMessage, pThread )
{
  Logger.log("Processing message: " + pMessage.getId());
  if( pMessage )
  {
    var url = "https://mail.google.com/mail/u/0/#all/" + pMessage.getId();
    var labelsOnMessage = getLabelsOnMessage( pThread );

    // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3428
    var body = "";
    try
    {
      body = pMessage.getPlainBody();
    }
    catch(e)
    {
      body = "(unable to read the body of this message: https://code.google.com/p/google-apps-script-issues/issues/detail?id=3428)";
      Logger.log( "ERROR: " + e );
    }

    var info =
    {
      "Id" : pMessage.getId(),
      "From" : pMessage.getFrom(),
      "To" : pMessage.getTo(),
      "CC" : pMessage.getCc(),
      "BCC" : pMessage.getBcc(),
      "Date" : pMessage.getDate(),
      "Subject" : pMessage.getSubject(),
      "Body" : body,
      "Labels": labelsOnMessage,
      "Link" : url
    };
    return info;
  }
  return null;
}

function getInfoText( pMessage, pThread )
{
  var info = getInfo( pMessage, pThread );
  var jsonTextFormatted = JSON.stringify(info, null, 2);
  jsonTextFormatted = jsonTextFormatted.replace( /^(\s*)(.*):\{$/gim, "$1$2:\n$1{" )
  jsonTextFormatted = jsonTextFormatted.replace( /^(\s*)(.*):\[\{$/gim, "$1$2:\n$1[\n$1\t{" )
  return jsonTextFormatted;
}

function getLabelsOnMessage( pThread )
{
  var labelsOnMessage = "";
  if( pThread )
  {
    var labels = pThread.getLabels();
    for( var k = 0; k < labels.length; k++ )
    {
      if( labelsOnMessage === "" )
      {
        labelsOnMessage = labels[k].getName();
      }
      else
      {
        labelsOnMessage = labelsOnMessage + "; " + labels[k].getName();
      }
    }
  }
  return labelsOnMessage;
}
 
// Compute the size of email attachments in MB
 
function getMessageSize(att)
{
  var size = 0;
  for (var i=0; i<att.length; i++) 
  {
    //size += att[i].getBytes().length;
    size += att[i].getSize(); // Better and faster than getBytes()
  }
  // Wait for a second to avoid hitting the system limit
  // Utilities.sleep(1000);
  return Math.round(size*100/(1024*1024))/100;
}
