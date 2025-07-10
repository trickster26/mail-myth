function onOpen() {
  SpreadsheetApp.getUi().createMenu('Mail Myth')
    .addItem('Open Mail Myth Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Mail Myth')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getGmailDrafts() {
  var drafts = GmailApp.getDrafts();
  return drafts.map(function(draft) {
    var msg = draft.getMessage();
    return {
      id: draft.getId(),
      subject: msg.getSubject(),
      body: msg.getBody()
    };
  });
}

function getOrCreateColumn(sheet, header) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(header) + 1;
  if (col === 0) {
    // Add new column
    col = headers.length + 1;
    sheet.getRange(1, col).setValue(header);
  }
  return col;
}

function generateTrackingId() {
  return Utilities.getUuid();
}

function getDriveFiles() {
  var files = DriveApp.getFiles();
  var fileList = [];
  var count = 0;
  
  // Limit to first 50 files to avoid timeout
  while (files.hasNext() && count < 50) {
    var file = files.next();
    fileList.push({
      id: file.getId(),
      name: file.getName(),
      mimeType: file.getMimeType(),
      size: file.getSize()
    });
    count++;
  }
  
  return fileList;
}

function processLinksForTracking(htmlBody, trackingId) {
  var scriptUrl = ScriptApp.getService().getUrl();
  
  // Find all links in the HTML and replace with tracking URLs
  var linkRegex = /<a\s+[^>]*href\s*=\s*["']([^"']+)["'][^>]*>/gi;
  
  return htmlBody.replace(linkRegex, function(match, url) {
    // Skip if it's already a tracking URL or an email/tel link
    if (url.indexOf(scriptUrl) !== -1 || url.startsWith('mailto:') || url.startsWith('tel:')) {
      return match;
    }
    
    // Create tracking URL
    var trackingUrl = scriptUrl + '?action=click&tid=' + encodeURIComponent(trackingId) + '&url=' + encodeURIComponent(url);
    
    // Replace the href attribute
    return match.replace(url, trackingUrl);
  });
}

function sendMailMerge(draftId, sheetRange, attachmentIds) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange(sheetRange).getValues();
  var headers = data[0];
  var rows = data.slice(1);
  var draft = GmailApp.getDraft(draftId);
  var msg = draft.getMessage();
  var subject = msg.getSubject();
  var body = msg.getBody();
  var sentCount = 0;

  // Ensure TrackingID, Opened, and Clicked columns exist
  var trackingIdCol = getOrCreateColumn(sheet, 'TrackingID');
  var openedCol = getOrCreateColumn(sheet, 'Opened');
  var clickedCol = getOrCreateColumn(sheet, 'Clicked');

  // Get attachments from Drive if provided
  var attachments = [];
  if (attachmentIds && attachmentIds.length > 0) {
    attachmentIds.forEach(function(fileId) {
      try {
        var file = DriveApp.getFileById(fileId);
        attachments.push(file.getBlob());
      } catch (e) {
        console.error('Could not get attachment with ID:', fileId, e);
      }
    });
  }

  rows.forEach(function(row, idx) {
    var email = row[0]; // Assume first column is email
    if (!email) return; // Skip if no email
    var personalizedBody = body;
    headers.forEach(function(header, i) {
      var re = new RegExp('{{' + header + '}}', 'g');
      personalizedBody = personalizedBody.replace(re, row[i]);
    });
    // Generate or get tracking ID
    var trackingId = row[trackingIdCol - 1];
    if (!trackingId) {
      trackingId = generateTrackingId();
      sheet.getRange(idx + 2, trackingIdCol).setValue(trackingId);
    }
    
    // Process links for click tracking
    personalizedBody = processLinksForTracking(personalizedBody, trackingId);
    
    // Insert tracking pixel
    var scriptUrl = ScriptApp.getService().getUrl();
    var pixelUrl = scriptUrl + '?action=open&tid=' + encodeURIComponent(trackingId);
    personalizedBody += '<img src="' + pixelUrl + '" width="1" height="1" style="display:none">';
    
    // Send email with or without attachments
    var emailOptions = {htmlBody: personalizedBody};
    if (attachments.length > 0) {
      emailOptions.attachments = attachments;
    }
    
    GmailApp.sendEmail(email, subject, '', emailOptions);
    sentCount++;
  });
  return sentCount;
}

function scheduleMailMerge(draftId, sheetRange, scheduleDateTime) {
  var triggerTime = new Date(scheduleDateTime);
  var now = new Date();
  
  if (triggerTime <= now) {
    throw new Error('Schedule time must be in the future');
  }
  
  // Create the trigger
  var trigger = ScriptApp.newTrigger('executeScheduledMailMerge')
    .timeBased()
    .at(triggerTime)
    .create();
  
  // Store job details in script properties
  var properties = PropertiesService.getScriptProperties();
  var jobId = trigger.getUniqueId();
  var jobData = {
    id: jobId,
    draftId: draftId,
    sheetRange: sheetRange,
    scheduledTime: triggerTime.toISOString(),
    status: 'scheduled'
  };
  
  properties.setProperty('job_' + jobId, JSON.stringify(jobData));
  
  return {
    jobId: jobId,
    scheduledTime: triggerTime.toISOString()
  };
}

function executeScheduledMailMerge(e) {
  var triggerId = e.triggerUid;
  var properties = PropertiesService.getScriptProperties();
  var jobDataStr = properties.getProperty('job_' + triggerId);
  
  if (jobDataStr) {
    var jobData = JSON.parse(jobDataStr);
    try {
      // Execute the mail merge
      var result = sendMailMerge(jobData.draftId, jobData.sheetRange);
      
      // Update job status
      jobData.status = 'completed';
      jobData.sentCount = result;
      jobData.completedTime = new Date().toISOString();
      properties.setProperty('job_' + triggerId, JSON.stringify(jobData));
      
    } catch (error) {
      // Update job status with error
      jobData.status = 'failed';
      jobData.error = error.toString();
      jobData.completedTime = new Date().toISOString();
      properties.setProperty('job_' + triggerId, JSON.stringify(jobData));
    }
  }
}

function getScheduledJobs() {
  var properties = PropertiesService.getScriptProperties();
  var allProperties = properties.getProperties();
  var jobs = [];
  
  for (var key in allProperties) {
    if (key.indexOf('job_') === 0) { // Replace startsWith with indexOf
      try {
        var jobData = JSON.parse(allProperties[key]);
        jobs.push(jobData);
      } catch (e) {
        // Skip malformed job data
      }
    }
  }
  
  // Sort by scheduled time
  jobs.sort(function(a, b) {
    return new Date(a.scheduledTime) - new Date(b.scheduledTime);
  });
  
  return jobs;
}

function cancelScheduledJob(jobId) {
  var triggers = ScriptApp.getProjectTriggers();
  var properties = PropertiesService.getScriptProperties();
  
  // Find and delete the trigger
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getUniqueId() === jobId) {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }
  
  // Update job status
  var jobDataStr = properties.getProperty('job_' + jobId);
  if (jobDataStr) {
    var jobData = JSON.parse(jobDataStr);
    jobData.status = 'cancelled';
    jobData.cancelledTime = new Date().toISOString();
    properties.setProperty('job_' + jobId, JSON.stringify(jobData));
  }
  
  return true;
}

function doGet(e) {
  var action = e.parameter.action || 'open'; // Default to open tracking
  var tid = e.parameter.tid;
  
  if (tid) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var trackingIdCol = data[0].indexOf('TrackingID');
    
    if (trackingIdCol !== -1) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][trackingIdCol] === tid) {
          if (action === 'open') {
            // Handle open tracking
            var openedCol = data[0].indexOf('Opened');
            if (openedCol !== -1) {
              sheet.getRange(i + 1, openedCol + 1).setValue('Yes');
            }
          } else if (action === 'click') {
            // Handle click tracking
            var clickedCol = data[0].indexOf('Clicked');
            if (clickedCol !== -1) {
              var currentClicks = sheet.getRange(i + 1, clickedCol + 1).getValue();
              var clickCount = currentClicks ? parseInt(currentClicks) + 1 : 1;
              sheet.getRange(i + 1, clickedCol + 1).setValue(clickCount);
            }
            
            // Redirect to original URL
            var originalUrl = e.parameter.url;
            if (originalUrl) {
              return HtmlService.createHtmlOutput(
                '<script>window.location.href = "' + originalUrl + '";</script>'
              );
            }
          }
          break;
        }
      }
    }
  }
  
  if (action === 'click') {
    // If no redirect URL, show error
    return HtmlService.createHtmlOutput('<h3>Invalid link</h3>');
  }
  
  // Return tracking pixel for open tracking
  var img = Utilities.base64Decode('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==');
  return ContentService.createBinaryOutput(img).setMimeType(ContentService.MimeType.GIF);
} 

function getTrackingStats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return {
      totalEmails: 0,
      opens: 0,
      clicks: 0,
      openRate: 0,
      clickRate: 0
    };
  }
  
  var headers = data[0];
  var emailCol = headers.indexOf('Email');
  var trackingIdCol = headers.indexOf('TrackingID');
  var openedCol = headers.indexOf('Opened');
  var clickedCol = headers.indexOf('Clicked');
  
  var totalEmails = 0;
  var opens = 0;
  var totalClicks = 0;
  
  // Count emails that have tracking IDs (i.e., were sent through the add-on)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    // Only count rows that have a tracking ID (emails sent through add-on)
    if (trackingIdCol !== -1 && row[trackingIdCol]) {
      totalEmails++;
      
      // Count opens
      if (openedCol !== -1 && row[openedCol] === 'Yes') {
        opens++;
      }
      
      // Count clicks
      if (clickedCol !== -1 && row[clickedCol]) {
        var clickCount = parseInt(row[clickedCol]) || 0;
        if (clickCount > 0) {
          totalClicks += clickCount;
        }
      }
    }
  }
  
  var openRate = totalEmails > 0 ? (opens / totalEmails) * 100 : 0;
  var clickRate = totalEmails > 0 ? (totalClicks / totalEmails) * 100 : 0;
  
  return {
    totalEmails: totalEmails,
    opens: opens,
    clicks: totalClicks,
    openRate: Math.round(openRate * 10) / 10, // Round to 1 decimal
    clickRate: Math.round(clickRate * 10) / 10
  };
} 