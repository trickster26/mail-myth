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

function findColumnIndex(headers, target) {
  target = target.trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].trim().toLowerCase() === target) {
      return i;
    }
  }
  return -1;
}

function getOrCreateColumn(sheet, header) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var normalizedHeader = header.trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].trim().toLowerCase() === normalizedHeader) {
      return i + 1;
    }
  }
  // Add new column at the end
  var col = headers.length + 1;
  sheet.getRange(1, col).setValue(header);
  Logger.log('Created column: ' + header + ' at index ' + col);
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
  var scriptUrl = 'https://script.google.com/macros/s/AKfycbzkYDSgsBv2wfesBDgVs5NB7eTlUddZEynRRVQ0E4MMqCgEaldO1UoAzaOLEPMKZKM_/exec';
  
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

function addUnsubscribeLink(emailBody, trackingId) {
  var scriptUrl = 'https://script.google.com/macros/s/AKfycbzkYDSgsBv2wfesBDgVs5NB7eTlUddZEynRRVQ0E4MMqCgEaldO1UoAzaOLEPMKZKM_/exec';
  var unsubscribeUrl = scriptUrl + '?action=unsubscribe&tid=' + encodeURIComponent(trackingId);
  
  var unsubscribeHtml = 
    '<div style="margin-top: 20px; padding: 15px; background-color: #f8f9fa; border-top: 1px solid #dee2e6; text-align: center; font-size: 12px; color: #6c757d;">' +
    '<p style="margin: 0 0 10px 0;">Don\'t want to receive these emails?</p>' +
    '<a href="' + unsubscribeUrl + '" style="color: #dc3545; text-decoration: none;">Unsubscribe</a>' +
    '</div>';
  
  return emailBody + unsubscribeHtml;
}

function isEmailUnsubscribed(email) {
  var properties = PropertiesService.getScriptProperties();
  var unsubscribedEmails = properties.getProperty('unsubscribed_emails');
  
  if (!unsubscribedEmails) {
    return false;
  }
  
  try {
    var emailList = JSON.parse(unsubscribedEmails);
    return emailList.indexOf(email.toLowerCase()) !== -1;
  } catch (e) {
    return false;
  }
}

function addToUnsubscribeList(email) {
  var properties = PropertiesService.getScriptProperties();
  var unsubscribedEmails = properties.getProperty('unsubscribed_emails');
  
  var emailList = [];
  if (unsubscribedEmails) {
    try {
      emailList = JSON.parse(unsubscribedEmails);
    } catch (e) {
      emailList = [];
    }
  }
  
  var emailLower = email.toLowerCase();
  if (emailList.indexOf(emailLower) === -1) {
    emailList.push(emailLower);
    properties.setProperty('unsubscribed_emails', JSON.stringify(emailList));
  }
  
  // Update the 'Unsubscribed' column in the sheet if the email exists
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var emailCol = findColumnIndex(data[0], 'Email');
  var unsubscribedCol = findColumnIndex(data[0], 'Unsubscribed');
  if (emailCol !== -1 && unsubscribedCol !== -1) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailCol] && data[i][emailCol].toLowerCase() === emailLower) {
        sheet.getRange(i + 1, unsubscribedCol + 1).setValue('Yes');
      }
    }
  }
  
  return true;
}

function removeFromUnsubscribeList(email) {
  var properties = PropertiesService.getScriptProperties();
  var unsubscribedEmails = properties.getProperty('unsubscribed_emails');
  
  if (!unsubscribedEmails) {
    return false;
  }
  
  try {
    var emailList = JSON.parse(unsubscribedEmails);
    var emailLower = email.toLowerCase();
    var index = emailList.indexOf(emailLower);
    
    if (index !== -1) {
      emailList.splice(index, 1);
      properties.setProperty('unsubscribed_emails', JSON.stringify(emailList));
      return true;
    }
  } catch (e) {
    return false;
  }
  
  return false;
}

function getUnsubscribeList() {
  var properties = PropertiesService.getScriptProperties();
  var unsubscribedEmails = properties.getProperty('unsubscribed_emails');
  
  if (!unsubscribedEmails) {
    return [];
  }
  
  try {
    return JSON.parse(unsubscribedEmails);
  } catch (e) {
    return [];
  }
}

function getGmailQuota() {
  var remaining = MailApp.getRemainingDailyQuota();
  return {
    remaining: remaining,
    max: 2000 // Default for Google Workspace accounts; adjust if needed
  };
}

function clearTrackingColumns(sheet, sheetRange) {
  var range = sheet.getRange(sheetRange);
  var data = range.getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var openedCol = findColumnIndex(headers, 'Opened');
  var lastOpenedCol = findColumnIndex(headers, 'Last Opened');
  var clickedCol = findColumnIndex(headers, 'Clicked');
  var lastClickedCol = findColumnIndex(headers, 'Last Clicked');
  for (var i = 1; i < data.length; i++) {
    var rowIdx = range.getRow() + i - 1;
    if (openedCol !== -1) sheet.getRange(rowIdx, openedCol + 1).setValue('');
    if (lastOpenedCol !== -1) sheet.getRange(rowIdx, lastOpenedCol + 1).setValue('');
    if (clickedCol !== -1) sheet.getRange(rowIdx, clickedCol + 1).setValue('');
    if (lastClickedCol !== -1) sheet.getRange(rowIdx, lastClickedCol + 1).setValue('');
  }
}

function sendMailMerge(draftId, sheetRange, attachmentIds) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange(sheetRange).getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows = data.slice(1);
  var draft = GmailApp.getDraft(draftId);
  var msg = draft.getMessage();
  var subject = msg.getSubject();
  var body = msg.getBody();
  var sentCount = 0;
  var skippedCount = 0;

  // Ensure all tracking columns exist
  var trackingIdCol = getOrCreateColumn(sheet, 'TrackingID');
  var openedCol = getOrCreateColumn(sheet, 'Opened');
  var lastOpenedCol = getOrCreateColumn(sheet, 'Last Opened');
  var clickedCol = getOrCreateColumn(sheet, 'Clicked');
  var lastClickedCol = getOrCreateColumn(sheet, 'Last Clicked');
  var unsubscribedCol = getOrCreateColumn(sheet, 'Unsubscribed');

  // Clear tracking columns for the selected range
  clearTrackingColumns(sheet, sheetRange);

  // Quota check
  var quota = MailApp.getRemainingDailyQuota();
  var emailCol = findColumnIndex(headers, 'Email');
  var toSend = 0;
  for (var i = 1; i < data.length; i++) {
    var email = data[i][emailCol];
    if (!email) continue;
    if (!isEmailUnsubscribed(email)) toSend++;
  }
  if (toSend > quota) {
    throw new Error('Not enough Gmail quota to send ' + toSend + ' emails. Remaining quota: ' + quota);
  }

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

  for (var idx = 0; idx < rows.length; idx++) {
    var row = rows[idx];
    var email = row[emailCol];
    if (!email) continue;
    // Check if email is unsubscribed
    if (isEmailUnsubscribed(email)) {
      sheet.getRange(idx + 2, unsubscribedCol).setValue('Yes');
      skippedCount++;
      continue;
    }
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
    // Add unsubscribe link
    personalizedBody = addUnsubscribeLink(personalizedBody, trackingId);
    // Insert tracking pixel
    var scriptUrl = 'https://script.google.com/macros/s/AKfycbzkYDSgsBv2wfesBDgVs5NB7eTlUddZEynRRVQ0E4MMqCgEaldO1UoAzaOLEPMKZKM_/exec';
    var pixelUrl = scriptUrl + '?action=open&tid=' + encodeURIComponent(trackingId);
    personalizedBody += '<img src="' + pixelUrl + '" width="1" height="1" style="display:none">';
    // Send email with or without attachments
    var emailOptions = {htmlBody: personalizedBody};
    if (attachments.length > 0) {
      emailOptions.attachments = attachments;
    }
    GmailApp.sendEmail(email, subject, '', emailOptions);
    sentCount++;
  }
  return {
    sent: sentCount,
    skipped: skippedCount,
    total: sentCount + skippedCount
  };
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
      jobData.sentCount = result.sent;
      jobData.skippedCount = result.skipped;
      jobData.totalSent = result.total;
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
  
  if (action === 'unsubscribe') {
    // Handle unsubscribe request
    if (tid) {
      var sheet = SpreadsheetApp.getActiveSheet();
      var data = sheet.getDataRange().getValues();
      var trackingIdCol = findColumnIndex(data[0], 'TrackingID');
      var emailCol = findColumnIndex(data[0], 'Email');
      
      if (trackingIdCol !== -1 && emailCol !== -1) {
        for (var i = 1; i < data.length; i++) {
          if (data[i][trackingIdCol] === tid) {
            var email = data[i][emailCol];
            addToUnsubscribeList(email);
            
            // Return unsubscribe confirmation page
            return HtmlService.createHtmlOutput(
              '<!DOCTYPE html>' +
              '<html>' +
              '<head>' +
              '<title>Unsubscribed</title>' +
              '<style>' +
              'body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }' +
              '.container { max-width: 500px; margin: 0 auto; }' +
              '.success { color: #28a745; font-size: 24px; margin-bottom: 20px; }' +
              '.message { color: #6c757d; line-height: 1.6; }' +
              '</style>' +
              '</head>' +
              '<body>' +
              '<div class="container">' +
              '<div class="success">✓ Unsubscribed Successfully</div>' +
              '<div class="message">' +
              '<p>You have been unsubscribed from our mailing list.</p>' +
              '<p>You will no longer receive emails from us.</p>' +
              '<p><small>If you change your mind, please contact us to resubscribe.</small></p>' +
              '</div>' +
              '</div>' +
              '</body>' +
              '</html>'
            );
          }
        }
      }
    }
    
    // Fallback unsubscribe page
    return HtmlService.createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html>' +
      '<head>' +
      '<title>Unsubscribe</title>' +
      '<style>' +
      'body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }' +
      '.container { max-width: 500px; margin: 0 auto; }' +
      '.error { color: #dc3545; font-size: 24px; margin-bottom: 20px; }' +
      '.message { color: #6c757d; line-height: 1.6; }' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="container">' +
      '<div class="error">⚠️ Error</div>' +
      '<div class="message">' +
      '<p>Unable to process your unsubscribe request.</p>' +
      '<p>Please contact us directly to unsubscribe.</p>' +
      '</div>' +
      '</div>' +
      '</body>' +
      '</html>'
    );
  }
  
  if (tid) {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var trackingIdCol = getOrCreateColumn(sheet, 'TrackingID') - 1;
    var lastOpenedCol = getOrCreateColumn(sheet, 'Last Opened');
    var lastClickedCol = getOrCreateColumn(sheet, 'Last Clicked');
    var openedCol = getOrCreateColumn(sheet, 'Opened');
    var clickedCol = getOrCreateColumn(sheet, 'Clicked');
    
    var found = false;
    for (var i = 1; i < data.length; i++) {
      var rowTid = data[i][trackingIdCol];
      if (rowTid && String(rowTid).trim().toLowerCase() === String(tid).trim().toLowerCase()) {
        found = true;
        if (action === 'open') {
          sheet.getRange(i + 1, openedCol).setValue('Yes');
          sheet.getRange(i + 1, lastOpenedCol).setValue(new Date().toISOString());
          Logger.log('Updated Opened and Last Opened for row ' + (i + 1));
        } else if (action === 'click') {
          var currentClicks = sheet.getRange(i + 1, clickedCol).getValue();
          var clickCount = currentClicks ? parseInt(currentClicks) + 1 : 1;
          sheet.getRange(i + 1, clickedCol).setValue(clickCount);
          sheet.getRange(i + 1, lastClickedCol).setValue(new Date().toISOString());
          Logger.log('Updated Clicked and Last Clicked for row ' + (i + 1));
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
    if (!found) {
      Logger.log('TrackingID not found: ' + tid);
    }
  }
  
  if (action === 'click') {
    // If no redirect URL, show error
    return HtmlService.createHtmlOutput('<h3>Invalid link</h3>');
  }
  
  // Return tracking pixel for open tracking
  return HtmlService.createHtmlOutput('<img src="https://www.google.com/images/cleardot.gif" style="display:none" width="1" height="1">');
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
  var emailCol = findColumnIndex(headers, 'Email');
  var trackingIdCol = findColumnIndex(headers, 'TrackingID');
  var openedCol = findColumnIndex(headers, 'Opened');
  var clickedCol = findColumnIndex(headers, 'Clicked');
  
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

function getAnalyticsData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var emailCol = findColumnIndex(headers, 'Email');
  var openedCol = findColumnIndex(headers, 'Opened');
  var lastOpenedCol = findColumnIndex(headers, 'Last Opened');
  var clickedCol = findColumnIndex(headers, 'Clicked');
  var lastClickedCol = findColumnIndex(headers, 'Last Clicked');
  var unsubCol = findColumnIndex(headers, 'Unsubscribed');
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    rows.push({
      email: emailCol !== -1 ? row[emailCol] : '',
      opened: openedCol !== -1 ? row[openedCol] : '',
      lastOpened: lastOpenedCol !== -1 ? row[lastOpenedCol] : '',
      clicks: clickedCol !== -1 ? row[clickedCol] : '',
      lastClicked: lastClickedCol !== -1 ? row[lastClickedCol] : '',
      unsubscribed: unsubCol !== -1 ? row[unsubCol] : ''
    });
  }
  return rows;
} 

function saveTemplate(name, subject, body) {
  var properties = PropertiesService.getScriptProperties();
  var templates = properties.getProperty('mail_templates');
  var templateList = templates ? JSON.parse(templates) : {};
  templateList[name] = { subject: subject, body: body, updated: new Date().toISOString() };
  properties.setProperty('mail_templates', JSON.stringify(templateList));
  return true;
}

function loadTemplate(name) {
  var properties = PropertiesService.getScriptProperties();
  var templates = properties.getProperty('mail_templates');
  if (!templates) return null;
  var templateList = JSON.parse(templates);
  return templateList[name] || null;
}

function listTemplates() {
  var properties = PropertiesService.getScriptProperties();
  var templates = properties.getProperty('mail_templates');
  if (!templates) return [];
  var templateList = JSON.parse(templates);
  return Object.keys(templateList).map(function(name) {
    return {
      name: name,
      subject: templateList[name].subject,
      updated: templateList[name].updated
    };
  });
}

function deleteTemplate(name) {
  var properties = PropertiesService.getScriptProperties();
  var templates = properties.getProperty('mail_templates');
  if (!templates) return false;
  var templateList = JSON.parse(templates);
  if (templateList[name]) {
    delete templateList[name];
    properties.setProperty('mail_templates', JSON.stringify(templateList));
    return true;
  }
  return false;
} 