function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    let data;

    // Handle both JSON and form data
    if (e.postData.type === 'application/json') {
      data = JSON.parse(e.postData.contents);
    } else {
      // Handle form data
      data = e.parameter;
    }

    const spreadsheetId = '1ctH4WQQk5BwxBI1apVb0KxRuUj3bvHY6Q_suLo8WApo';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getActiveSheet();

    // Check if headers exist, if not create them
    if (sheet.getLastRow() === 0) {
      const headers = [
        'Timestamp',
        'Visit Date',
        'Visit Time',
        'Employee ID',
        'Employee Name',
        'Department',
        'Job Category',
        'Location',
        'Job Role',
        'Visit Type',
        'Primary Complaint',
        'Secondary Complaint',
        'Treatment Given',
        'Medication Dispensed',
        'Days Off',
        'Referral Required',
        'Referral Type',
        'Follow-up Required',
        'Follow-up Date',
        'Attending Staff',
        'Additional Notes',
        'Data Source'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#667eea');
      headerRange.setFontColor('white');
    }

    // Prepare the row data
    const rowData = [
      new Date().toLocaleString('en-US', {timeZone: 'Africa/Lagos'}),
      data.visitDate || '',
      data.visitTime || '',
      data.employeeId || '',
      data.employeeName || '',
      data.department || '',
      data.jobCategory || '',
      data.employeeLocation || '',
      data.employeeJobRole || '',
      data.visitType || '',
      data.primaryComplaint || '',
      data.secondaryComplaint || '',
      data.treatmentGiven || '',
      data.medicationDispensed || '',
      data.daysOff || '0',
      data.referralRequired === 'true' || data.referralRequired === true ? 'Yes' : 'No',
      data.referralType || '',
      data.followupRequired === 'true' || data.followupRequired === true ? 'Yes' : 'No',
      data.followupDate || '',
      data.attendingStaff || '',
      data.additionalNotes || '',
      data.dataSource || 'manual_entry'
    ];

    sheet.appendRow(rowData);
    sheet.autoResizeColumns(1, rowData.length);

    // Return a simple redirect response
    return HtmlService.createHtmlOutput(`
      <script>
        window.location.href = 'data:text/html,<html><head><title>Success</title></head><body><h2>✅ Data saved successfully!</h2><p>You can close this window.</p><script>setTimeout(function(){window.close();}, 2000);</script></body></html>';
      </script>
      <p>Data saved successfully!</p>
    `);

  } catch (error) {
    return HtmlService.createHtmlOutput(`
      <script>
        window.location.href = 'data:text/html,<html><head><title>Error</title></head><body><h2>❌ Error occurred!</h2><p>Error: ${error.toString()}</p><script>setTimeout(function(){window.close();}, 3000);</script></body></html>';
      </script>
      <p>Error: ${error.toString()}</p>
    `);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return HtmlService.createHtmlOutput('BPN Clinic Form Web App is running!');
}