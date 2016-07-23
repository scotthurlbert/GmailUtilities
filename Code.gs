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
  var sizeLimit = "1000000";                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                ;
  var start = new Date();
  var timedOut = false;
  var processed = 0;
  var totalToProcess = 0;
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
    totalToProcess = threads.length;
    Logger.log("Search returned " + totalToProcess + " threads larger than " + sizeLimit + ".");
    for (var i=0; i<threads.length; i++) 
    {
      if (isTimeUp(start)) 
      {
      	timedOut = true;
        break;
      }
      label.addToThread(threads[i]);
      processed++;
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
  var msg = "Messages labeled: " + processed + " of " + totalToProcess;
  if( timedOut )
  {
    Logger.log("Time's up. " + msg );
  }
  else
  {
    Logger.log( "Done tagging large emails. " + msg );
  }
}

// Adapted from:
// https://www.maketecheasier.com/google-scripts-to-automate-gmail/

// This gave me the idea
// http://www.labnol.org/internet/send-gmail-to-google-drive/21236/
// https://chrome.google.com/webstore/detail/save-emails-and-attachmen/nflmnfjphdbeagnilbihcodcophecebc

// This had much of the code I ended up using to include the inline images, though it had a shit ton
// of bugs and issues - it was definetly a good start, but it was massively changed.
// http://pixelcog.com/blog/2015/gmail-to-pdf/

function save_Gmail_as_PDF()
{
  var start = new Date();
  var timedOut = false;
  var processed = 0;
  var totalToProcess = 0;
  //var label = GmailApp.getUserLabelByName("Save as PDF");  // inlineimage
  var label = GmailApp.getUserLabelByName("inlineimage");  // 
  var labelAfter = GmailApp.getUserLabelByName("Saved To PDF");  
  if(labelAfter == null)
  {
    GmailApp.createLabel('Saved To PDF');
  }
  if(label == null)
  {
    GmailApp.createLabel("Save as PDF"); // 'Save as PDF'
  }
  else
  {
    var threads = label.getThreads();
    totalToProcess = threads.length;
    Logger.log( "Saving " + threads.length + " emails to google drive." );
    for (var i = 0; i < threads.length; i++) 
    {
      if( isTimeUp(start) ) 
      {
        timedOut = true;
        break;
      }
      var messages = threads[i].getMessages();  
      var message = messages[0];
      var body    = message.getBody();
      var dateTime = message.getDate();
      var subject = dateTime.toISOString().substring(0, 19).replace(/:/g, '.').replace(/-/g, '') + ", " + message.getFrom() + ", " + message.getSubject();
      var attachments = message.getAttachments();
      var inlineimages = message.inlineImages;
      
      Logger.log( "Creating message: " + subject );
      
      var jsonTextFormatted = getInfoText( message, threads[i] );
      
      //Logger.log(message.inlineImages);
      //Logger.log(message.getRawContent());
      //body = embedInlineImages_(body, message.getRawContent());

      //if( message.inlineImages != null )
      //{
      //  Logger.log( "inlineImages" );
      //  body = embedInlineImages_(body, message.getRawContent());
      //}
      
      for(var j = 1;j<messages.length;j++)
      {
        //if( message.inlineImages != null )
        //{
        //  body += embedInlineImages_(messages[j].getBody(), messages[j].getRawContent());
        //}
        //else
        //{
        //  body += messages[j].getBody();
        //}
        //body += messages[j].getBody();
        
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
      //var bodydochtml = DriveApp.createFile(subject+'.html', body, "text/html")
      //var bodyId=bodydochtml.getId();
 
      // Convert the HTML to PDF
      //var bodydocpdf = bodydochtml.getAs('application/pdf');
      
      var gmailFolders = DriveApp.getFoldersByName( "Gmail PDFs" );
      var gmailFolder = gmailFolders.next();
      var folder = gmailFolder.createFolder(subject)
      gmailFolder.addFolder(folder);
      
      var opts = 
          {
            includeHeader: true,
            includeAttachments: false, // was true
            embedAttachments: false, // was true
            embedRemoteImages: true, // was true
            embedInlineImages: true,
            embedAvatar: true,
            width: 700,
            filename: null
          };
      
      // create a pdf of the message
      //var pdf = messageToPdf(threads[i], opts, folder);

      var html = messageToHtml(threads[i], opts, folder )
      var pdf = html.getAs('application/pdf');


      // prefix the pdf filename with a date string
      pdf.setName(formatDate(message, 'yyyyMMdd-hh.mm.ss ') + pdf.getName());
      
      // create the file
      folder.createFile(pdf);
      
      
      html.setName( subject + ".html" );
      // create the file
      folder.createFile(html);

      /*
      for( var k = 0; k < messages.length; k++ )
      {
        // create a pdf of the message
        var pdf = messageToPdf(messages[k]);

        // prefix the pdf filename with a date string
        pdf.setName(formatDate(messages[k], 'yyyyMMdd-hh.mm.ss ') + pdf.getName());

        folder.createFile(pdf);

      }
      */
      
      if(attachments.length > 0)
      {
        for (var j = 0; j < attachments.length; j++) 
        {
          folder.createFile(attachments[j]);
          // Utilities.sleep(1000);
        }
      }
      //else
      //{
      //  DriveApp.createFile(bodydocpdf);
      //}      
      //folder.createFile(bodydocpdf);
      //folder.createFile(pdf);
      folder.createFile("info.json", jsonTextFormatted, "application/json" );
      
      //DriveApp.getFileById(bodyId).setTrashed(true);
      //label.removeFromThread(threads[i]);
      labelAfter.addToThread(threads[i]);
      processed++;
    }
  }
  var msg = "Messages saved: " + processed + " of " + totalToProcess;
  if( timedOut )
  {
    Logger.log("Time's up. " + msg );
  }
  else
  {
    Logger.log( "Done saving to g-drive. " + msg );
  }
}

// Adapted from:
// https://ctrlq.org/code/20016-maximum-execution-time-limit

function isTimeUp(pStart) 
{
  var now = new Date();
  return now.getTime() - pStart.getTime() > 300000; // 5 minutes
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

    var body = "";
    try
    {
      // There is a bug in getPlainBody()
      // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3428
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
        labelsOnMessage = labels[k].getName() + ";";
      }
      else
      {
        labelsOnMessage = labelsOnMessage + " " + labels[k].getName() + ";";
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

// ***********************************************************************************************
// This had much of the code I ended up using to include the inline images, though it had a shit ton
// of bugs and issues - it was definetly a good start, but it was massively changed.
// http://pixelcog.com/blog/2015/gmail-to-pdf/
//
// LICENSE
//
// The MIT License (MIT)
//
// Copyright (c) 2015 PixelCog Inc.
//
// Permission is hereby granted, free of charge, to any person obtaining a copy of this software and 
// associated documentation files (the "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
// copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the 
// following conditions:
//
// The above copyright notice and this permission notice shall be included in all copies or substantial 
// portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT 
// LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN 
// NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
// WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE 
// SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// ***********************************************************************************************

/**
 * Wrapper for Utilities.formatDate() which provides sensible defaults
 *
 * @method formatDate
 * @param {string} message
 * @param {string} format
 * @param {string} timezone
 * @return {string} Formatted date
 */
function formatDate(message, format, timezone) {
  timezone = timezone || localTimezone_();
  format = format || "MMMMM dd, yyyy 'at' h:mm a '" + timezone + "'";
  return Utilities.formatDate(message.getDate(), timezone, format)
}

/**
 * Convert a Gmail message or thread to a PDF and return it as a blob
 *
 * @method messageToPdf
 * @param {GmailMessage|GmailThread} messages GmailMessage or GmailThread object (or an array of such objects)
 * @return {Blob}
 */
function messageToPdf(messages, opts, folder) 
{
  return messageToHtml(messages, opts, folder).getAs('application/pdf');
}


/**
 * Convert a Gmail message or thread to a HTML and return it as a blob
 *
 * @method messageToHtml
 * @param {GmailMessage|GmailThread} messages GmailMessage or GmailThread object (or an array of such objects)
 * @param {Object} options
 * @return {Blob}
 */
function messageToHtml(messages, opts, folder) 
{
  opts = opts || {};
  defaults_(opts, {
    includeHeader: true,
    includeAttachments: false, // was true
    embedAttachments: false, // was true
    embedRemoteImages: false, // was true
    embedInlineImages: true,
    embedAvatar: true,
    width: 700,
    filename: null
  });

  if (!(messages instanceof Array)) 
  {
    messages = isa_(messages, 'GmailThread') ? messages.getMessages() : [messages];
  }
  if (!messages.every(function(obj){ return isa_(obj, 'GmailMessage'); })) {
    throw "Argument must be of type GmailMessage or GmailThread.";
  }
  var name = opts.filename || sanitizeFilename_(messages[messages.length-1].getSubject()) + '.html';
  
  var images = [];
  var imageDict = {};
  var altDict = {};
  var imageStyles = "";

  for( var m=0; m < messages.length; m++ )
  {
    var message = messages[m],
        body = message.getBody();

    if (opts.embedInlineImages) 
    {
      //TODO: REMOVE
      folder.createFile( Utilities.newBlob(message.getRawContent(), "text/html", "raw" + m + ".html") );
      
      extractInlineImages( message.getRawContent(), imageDict, images, folder );
      // The idea of using styles to hold the images came from:
      // http://stackoverflow.com/questions/29187840/base64-inline-image-use-more-than-once
      for( var x=0; x < images.length; x++ )
      {
        
        Logger.log( "Adding to imageDict: CID: " + images[x] + ": " + imageDict[images[x]] + ", stylename: " + imageDict[images[x]].stylename || "no-style-found" );
        var imageVal = imageDict[images[x]];
        if( imageVal && imageVal.stylename && imageVal.status && imageVal.status != "published" )
        {
          imageStyles += imageVal.style + "\n\n";
          imageDict[images[x]].status = "published";
        }
        
        Logger.log( "Attempting to add to altDict: CID: " + images[x] + ": " + imageDict[images[x]].status + ", " + imageDict[images[x]].alt );
        if( imageVal && imageVal.alt && imageVal.stylename )
        {
          Logger.log( "ADDING ALTTEXT TO THE ALTDICT: " + imageVal.alt );
          altDict[imageVal.alt] = imageVal;
        }

      }
    }
  }
  
  // Logger.log( "ImageStyles finished: " + mystuff );
  
  var html = '<html>\n' +
      '<style type="text/css">\n' +
      'body{padding:0 10px;min-width:' + opts.width + 'px;-webkit-print-color-adjust: exact;}' +
      'body>dl.email-meta{font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size:14px;padding:0 0 10px;margin:0 0 5px;border-bottom:1px solid #ddd;page-break-before:always}' +
      'body>dl.email-meta:first-child{page-break-before:auto}' +
      'body>dl.email-meta dt{color:#808080;float:left;width:60px;clear:left;text-align:right;overflow:hidden;text-overf?low:ellipsis;white-space:nowrap;font-style:normal;font-weight:700;line-height:1.4}' +
      'body>dl.email-meta dd{margin-left:70px;line-height:1.4}' +
      'body>dl.email-meta dd a{color:#808080;font-size:0.85em;text-decoration:none;font-weight:normal}' +
      'body>dl.email-meta dd.avatar{float:right}' +
      'body>dl.email-meta dd.avatar img{max-height:72px;max-width:72px;border-radius:36px}' +
      'body>dl.email-meta dd.strong{font-weight:bold}' +
      'body>div.email-attachments{font-size:0.85em;color:#999}\n' +
       imageStyles +
      '</style>\n<body>\n';
  
  
  for( var m=0; m < messages.length; m++ )
  {
    var message = messages[m],
        subject = message.getSubject(),
        avatar = null,
        date = formatDate(message),
        from = formatEmails_(message.getFrom()),
        to   = formatEmails_(message.getTo()),
        body = message.getBody();

    if (opts.includeHeader) 
    {
      if (opts.embedAvatar && (avatar = emailGetAvatar(from))) 
      {
        avatar = '<dd class="avatar"><img src="' + renderDataUri_(avatar) + '" /></dd> ';
      } else 
      {
        avatar = '';
      }
      
      html += '<dl class="email-meta">\n' +
              '<dt>From:</dt>' + avatar + ' <dd class="strong">' + from + '</dd>\n' +
              '<dt>Subject:</dt> <dd>' + subject + '</dd>\n' +
              '<dt>Date:</dt> <dd>' + date + '</dd>\n' +
              '<dt>To:</dt> <dd>' + to + '</dd>\n' +
              '</dl>\n';      
    }
    if (opts.embedInlineImages) 
    {
      body = embedInlineImages_(body, message.getRawContent(), imageDict, altDict, images );
    }
    if (opts.embedRemoteImages) 
    {
      body = embedHtmlImages_(body);
    }
    if (opts.includeAttachments) 
    {
      var attachments = message.getAttachments();
      if (attachments.length > 0) 
      {
        body += '<br />\n<strong>Attachments:</strong>\n' +
                '<div class="email-attachments">\n';

        for (var a=0; a < attachments.length; a++) 
        {
          var filename = attachments[a].getName();
          var imageData;

          if (opts.embedAttachments && (imageData = renderDataUri_(attachments[a]))) 
          {
            body += '<img src="' + imageData + '" alt="&lt;' + filename + '&gt;" /><br />\n';
          } 
          else 
          {
            body += '&lt;' + filename + '&gt;<br />\n';
          }
        }
        body += '</div>\n';
      }
    }
    html += body;
    Logger.log( "The length of the body string is: " + body.length );
  }
  html += '</body>\n</html>';
  return Utilities.newBlob(html, 'text/html', name);
}

/**
 * Returns the name associated with an email string, or the domain name of the email.
 *
 * @method emailGetName
 * @param {string} email
 * @return {string} name or domain name
 */
function emailGetName(email) {
  return email.replace(/^<?(?:[^<\(]+@)?([^<\(,]+?|)(?:\s?[\(<>,].*|)$/i, '$1') || 'Unknown';
}

/**
 * Attempt to download an image representative of the email address provided. Using gravatar or
 * apple touch icons as appropriate.
 *
 * @method emailGetAvatar
 * @param {string} email
 * @return {Blob|boolean} Blob object or false
 */
function emailGetAvatar(email) {
  re = /[a-z0-9!#$%&'*+\/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+\/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?/gi
  if (!(email = email.match(re)) || !(email = email[0].toLowerCase())) {
    return false;
  }
  var domain = email.split('@')[1];
  var avatar = fetchRemoteFile_('http://www.gravatar.com/avatar/' + md5_(email) + '?s=128&d=404');
  if (!avatar && ['gmail','hotmail','yahoo.'].every(function(s){ return domain.indexOf(s) == -1 })) {
    avatar = fetchRemoteFile_('http://' + domain + '/apple-touch-icon.png') ||
             fetchRemoteFile_('http://' + domain + '/apple-touch-icon-precomposed.png');
  }
  return avatar;
}

/**
 * Download and embed all images referenced within an html document as data uris
 *
 * @param {string} html
 * @return {string} Html with embedded images
 */
function embedHtmlImages_(html) {
  // process all img tags
  html = html.replace(/(<img[^>]+src=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) {
    // Logger.log('Processing image src: ' + src);
    return tag + q + (renderDataUri_(src) || src) + q;
  });
  // process all style attributes
  html = html.replace(/(<[^>]+style=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, style) {
    style = style.replace(/url\((\\?["']?)([^\)]*)\1\)/gi, function(m, q, url) {
      return 'url(' + q + (renderDataUri_(url) || url) + q + ')';
    });
    return tag + q + style + q;
  });
  // process all style tags
  html = html.replace(/(<style[^>]*>)(.*?)(?:<\/style>)/gi, function(m, tag, style, end) {
    style = style.replace(/url\((["']?)([^\)]*)\1\)/gi, function(m, q, url) {
      return 'url(' + q + (renderDataUri_(url) || url) + q + ')';
    });
    return tag + style + end;
  });
  return html;
}

// The idea of using styles to hold the images came from:
// http://stackoverflow.com/questions/29187840/base64-inline-image-use-more-than-once

function extractInlineImages( raw, imageDict, images, folder )
{
  // <img src="#" class="myImage" />   

  // locate all inline content ids
  raw.replace(/<img[^>]+src=(?:3D)?(["'])cid:((?:(?!\1)[^\\]|\\.)*)\1.*?>/gi, function(m, q, cid) 
  {
    cid = cid.replace("\r\n", "").replace("=", "");
    Logger.log("Image with cid of: " + cid + ", and the match: " + m );
    var newImageVal = 
    { 
      "status" : "no file yet", 
      "alt" : "", 
      "cid" : cid, 
      "stylename" : null 
    };
    
    // We only want to use the new imageVal if the dictionary does not contain an existing one.
    var imageVal = imageDict[cid] || newImageVal;
    
    var imageTag = m.replace("=\r", "").replace("=\n", "").replace("=\r\n", "").replace("\r\n", "").replace("\r", "").replace("\n", "");
    Logger.log( "ImageTag after replacing linefeeds: " + imageTag );
    imageTag = imageTag.replace(/=3D/gi, "REALEQUALS");
    Logger.log( "ImageTag before replacing equals: " + imageTag );
    imageTag = imageTag.replace("=", "");
    imageTag = imageTag.replace(/REALEQUALS/gi, "=");
    Logger.log( "ImageTag after replacing equals: " + imageTag );
    
    Logger.log( "Looking for alt tag in: " + imageTag );
    
    if( imageTag.indexOf( "alt=" ) > 0 )
    {
      var altText = imageTag.replace( /(.*?<img[^>]+alt=)(["'])(.+)\2.*/gi, "$3" );
      Logger.log( "ALT TAG FOUND: " + altText );
      imageVal.alt = altText;
    }

    imageDict[cid] = imageVal;
    images.push(cid);
    
    return m;
  });

  // extract all inline images
  images = images.map(function(cid) 
  {
    var cidIndex = raw.search(new RegExp("Content-ID ?:.*?" + cid, 'i'));
    if (cidIndex === -1) return null;

    var prevBoundaryIndex = raw.lastIndexOf("\r\n--", cidIndex);
    var nextBoundaryIndex = raw.indexOf("\r\n--", prevBoundaryIndex+1);
    var part = raw.substring(prevBoundaryIndex, nextBoundaryIndex);

    var encodingLine = part.match(/Content-Transfer-Encoding:.*?\r\n/i)[0];
    var encoding = encodingLine.split(":")[1].trim();
    if (encoding != "base64") return null;

    var contentTypeLine = part.match(/Content-Type:.*?\r\n/i)[0];
    var contentType = contentTypeLine.split(":")[1].split(";")[0].trim();

    var startOfBlob = part.indexOf("\r\n\r\n");
    var blobText = part.substring(startOfBlob).replace("\r\n","");

    var files = folder.getFilesByName( cid );
    if( !files.hasNext() ) 
    {
      var imageBlob = Utilities.newBlob(Utilities.base64Decode(blobText), contentType, cid);
      file = folder.createFile(imageBlob);
      var renderedDataUri = renderDataUri_( imageBlob );
      // ToDo: create a clean cid for use as the stylename
      
      var stylename = cidToStylename(cid);
      
      // This will preserve the ALT name for this CID if there is one.
      var imageVal = imageDict[cid] || {};
      imageVal.status = "ready";
      imageVal.cid = cid;
      imageVal.style = "." + stylename + "{ content: url('" + renderedDataUri + "'); } ",
      imageVal.stylename = stylename

      imageDict[cid] = imageVal;
      Logger.log( "Added image for cid: " + cid + ", imageDict[cid].status: " + imageDict[cid].status + ", stylename: " + imageDict[cid].stylename );
    }
    return cid;
  }).filter(function(i){return i});
}
                      
function cidToStylename( pCid )
{
  var stylename = pCid || "default-style-from-cidToStylename";
  stylename = stylename.replace( /\./g, "" ).replace( /@/g, "" ).replace( /\&/g, "" );
  stylename = stylename.replace( /\[/g, "" ).replace( /\]/g, "" );
  stylename = stylename.replace( /\(/g, "" ).replace( /\)/g, "" );
  stylename = stylename.replace( /\{/g, "" ).replace( /\}/g, "" );
  stylename = stylename.replace( /\</g, "" ).replace( /\>/g, "" );
  return stylename;
}

/**
 * Extract and embed all inline images (experimental)
 *
 * @param {string} html Message body
 * @param {string} raw Unformatted message contents
 * @return {string} Html with embedded images
 */
function embedInlineImages_(html, raw, imageDict, altDict, images ) 
{
  /*
  var images = [];
  var imageDict = {};
  
   //var rObj = {};
   //rObj[obj.key] = obj.value;
   //return rObj;

  // locate all inline content ids
  raw.replace(/<img[^>]+src=(?:3D)?(["'])cid:((?:(?!\1)[^\\]|\\.)*)\1/gi, function(m, q, cid) 
  {
    cid = cid.replace("\r\n", "").replace("=", "");
    Logger.log("Image with cid of: " + cid );
    imageDict[cid] = {};
    images.push(cid);
    return m;
  });

  // extract all inline images
  images = images.map(function(cid) 
  {
    var cidIndex = raw.search(new RegExp("Content-ID ?:.*?" + cid, 'i'));
    if (cidIndex === -1) return null;

    var prevBoundaryIndex = raw.lastIndexOf("\r\n--", cidIndex);
    var nextBoundaryIndex = raw.indexOf("\r\n--", prevBoundaryIndex+1);
    var part = raw.substring(prevBoundaryIndex, nextBoundaryIndex);

    var encodingLine = part.match(/Content-Transfer-Encoding:.*?\r\n/i)[0];
    var encoding = encodingLine.split(":")[1].trim();
    if (encoding != "base64") return null;

    var contentTypeLine = part.match(/Content-Type:.*?\r\n/i)[0];
    var contentType = contentTypeLine.split(":")[1].split(";")[0].trim();

    var startOfBlob = part.indexOf("\r\n\r\n");
    var blobText = part.substring(startOfBlob).replace("\r\n","");

    imageDict[cid] = Utilities.newBlob(Utilities.base64Decode(blobText), contentType, cid);
    return Utilities.newBlob(Utilities.base64Decode(blobText), contentType, cid);
  }).filter(function(i){return i});

  Logger.log("sleeping for 30 seconds...");
  Utilities.sleep(30000);
  Logger.log("done sleeping." );
  */
  
  // process all img tags which reference "attachments"
  html = html.replace(/(<img[^>]+src=)(["'])(\?view=att(?:(?!\2)[^\\]|\\.)*)\2/gi, function(m, tag, q, src) 
  {
    Logger.log( "replacing tag: " + tag + ", with match: " + m + ", source: " + src + ", q: " + q );
	var start = src.indexOf( "realattid=" ) + 10;
    var key1 = src.substring( start );
    var key = key1.substring( 0, key1.indexOf( "&amp;" ) );
    var imageVal = imageDict[key];
    if( imageVal )
    {
      if( imageVal.stylename )
      {
        Logger.log( "Using imageDict for image! " + key + ", stylename: " + imageVal.stylename );
        //return tag + q + ( renderDataUri_( imageDict[key] ) || src ) + q;
        /*
        var files = DriveApp.getFilesByName( key );
        while( files.hasNext() ) 
        {
        var file = files.next();
        Logger.log( "renderDataUri for file " + file.getName() );
        var fileBlob = file.getBlob();
        Logger.log( "The size of our image blob is: " + fileBlob.getBytes().length + " bytes." );
        var renderedDataUri = renderDataUri_( fileBlob );
        Logger.log( "The length of the renderedDataUri is: " + renderedDataUri.length );
        return tag + q + ( renderedDataUri || src ) + q;
        // return tag + q + key + q;
        }
        */
        return tag + q + "#" + q + " class=" + q + imageVal.stylename + q;
        //<img src="#" class="myImage" />  
      }
      else
      {
        Logger.log( "No style for key: " + key );
      }
    }
    else
    {
      Logger.log( "Could not find an entry in imageDict for key: " + key + ", imageVal: " + imageVal );
    }
    return tag + q + src + q;
    // return tag + q + (renderDataUri_(images.shift() ) || src ) + q;
  });
  
  // process all image tags by the alt-text.  This is a fall back, and less reliable.
  // process all img tags which reference "attachments"
  html = html.replace(/(<img[^>]+src=)(["'])(\?view=att).+?>/gi, function(m, tag, q, src) 
  {
    Logger.log( "ALT tag: " + tag + ", match: " + m + ", source: " + src + ", q: " + q );
    
    if( m.indexOf( "alt=" ) > 0 )
    {
      var altText = m.replace( /(.*?<img[^>]+alt=)(["'])(.+)\2.*/gi, "$3" );
      Logger.log( "ALT TAG FOUND during replacement: " + altText );
      if( altDict[altText] && altDict[altText].stylename )
      {
        var imageVal = altDict[altText];
        Logger.log( "Using alt tag [" + altText + "] pointing to style: " + imageVal.stylename );
        return "<img src=" + q + "#" + q + " class=" + q + imageVal.stylename + q + ">";
      }
      else
      {
        Logger.log( "No entry found in altDict for: " + altText );
      }
    }
    
    /*
	var start = src.indexOf( "realattid=" ) + 10;
    var key1 = src.substring( start );
    var key = key1.substring( 0, key1.indexOf( "&amp;" ) );
    var imageVal = imageDict[key];
    if( imageVal && imageVal.status == "ready" )
    {
      Logger.log( "Using imageDict for image! " + key + ", stylename: " + imageVal.stylename );
      return tag + q + "#" + q + " class=" + q + imageVal.stylename + q;
    }
    else
    {
      Logger.log( "Could not find an entry in imageDict for key: " + key );
    }
    return tag + q + src + q;
    */
    return m;
  });

  return html;
 
}

/**
 * Convert an image into a base64-encoded data uri.
 *
 * @param {Blob|string} Blob object containing an image file or a remote url string
 * @return {string} Data uri
 */
function renderDataUri_(image) 
{
  if( typeof image == 'string' && !( isValidUrl_(image) && (image = fetchRemoteFile_(image) ) ) ) 
  {
    return null;
  }
  if (isa_(image, 'Blob') || isa_(image, 'GmailAttachment')) 
  {
    var type = image.getContentType().toLowerCase();
    var data = Utilities.base64Encode(image.getBytes());
    if (type.indexOf('image') == 0) 
    {
      return 'data:' + type + ';base64,' + data;
    }
  }
  return null;
}

/**
 * Fetch a remote file and return as a Blob object on success
 *
 * @param {string} url
 * @return {Blob}
 */
function fetchRemoteFile_(url) {
  try {
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    return response.getResponseCode() == 200 ? response.getBlob() : null;
  } catch (e) {
    return null;
  }
}

/**
 * Validate a url string (taken from jQuery)
 *
 * @param {string} url
 * @return {boolean}
 */
function isValidUrl_(url) {
  return /^(https?|ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(\#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i.test(url);
}

/**
 * Sanitize a filename by filtering out characters not allowed in most filesystems
 *
 * @param {string} filename
 * @return {string}
 */
function sanitizeFilename_(filename) {
  return filename.replace(/[\/\?<>\\:\*\|":\x00-\x1f\x80-\x9f]/g, '');
}

/**
 * Turn emails of the form "<handle@domain.tld>" into 'mailto:' links.
 *
 * @param {string} emails
 * @return {string}
 */
function formatEmails_(emails) {
  var pattern = new RegExp(/<(((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)>/i);
  return emails.replace(pattern, function(match, handle) {
    return '<a href="mailto:' + handle + '">' + handle + '</a>';
  });
}

/**
 * Test class name for Google Apps Script objects. They have no constructors so we must test them
 * with toString.
 *
 * @param {Object} obj
 * @param {string} class
 * @return {boolean}
 */
function isa_(obj, class) {
  return typeof obj == 'object' && typeof obj.constructor == 'undefined' && obj.toString() == class;
}

/**
 * Assign default attributes to an object.
 *
 * @param {Object} options
 * @param {Object} defaults
 */
function defaults_(options, defaults) {
  for (attr in defaults) {
    if (!options.hasOwnProperty(attr)) {
      options[attr] = defaults[attr];
    }
  }
}

/**
 * Get our current timezone string (or GMT if it cannot be determined)
 *
 * @return {string}
 */
function localTimezone_() {
  var timezone = new Date().toTimeString().match(/\(([a-z0-9]+)\)/i);
  return timezone.length ? timezone[1] : 'GMT';
}

/**
 * Create an MD5 hash of a string and return the reult as hexadecimal.
 *
 * @param {string} str
 * @return {string}
 */
function md5_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr) {
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1?'0':'') + chr;
  },'');
}