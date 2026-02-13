// GOOGLE APPS SCRIPT CODE
// Copy this entire file and paste it into Google Apps Script

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);

    // Check if headers exist, if not create them
    if (sheet.getLastRow() === 0) {
      const headers = [
        'Timestamp',
        'Name',
        'Wine A Rating',
        'Wine A Comments',
        'Wine B Rating',
        'Wine B Comments',
        'Wine C Rating',
        'Wine C Comments',
        'Wine D Rating',
        'Wine D Comments',
        'Wine E Rating',
        'Wine E Comments',
        'Wine F Rating',
        'Wine F Comments',
        'Wine G Rating',
        'Wine G Comments',
        'Wine H Rating',
        'Wine H Comments',
        'Wine I Rating',
        'Wine I Comments'
      ];
      sheet.appendRow(headers);
    }

    // Append the data
    const row = [
      data.timestamp || new Date().toISOString(),
      data.name || '',
      data.wineA_rating || 'n/a',
      data.wineA_comments || '',
      data.wineB_rating || 'n/a',
      data.wineB_comments || '',
      data.wineC_rating || 'n/a',
      data.wineC_comments || '',
      data.wineD_rating || 'n/a',
      data.wineD_comments || '',
      data.wineE_rating || 'n/a',
      data.wineE_comments || '',
      data.wineF_rating || 'n/a',
      data.wineF_comments || '',
      data.wineG_rating || 'n/a',
      data.wineG_comments || '',
      data.wineH_rating || 'n/a',
      data.wineH_comments || '',
      data.wineI_rating || 'n/a',
      data.wineI_comments || ''
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      'status': 'success'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      'status': 'error',
      'message': error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function - you can run this to test if the script works
function test() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        name: 'Test User',
        wineA_rating: '8',
        wineA_comments: 'Great wine!',
        wineB_rating: '7',
        wineB_comments: 'Pretty good',
        wineC_rating: '9',
        wineC_comments: 'Excellent',
        wineD_rating: '6',
        wineD_comments: 'Decent',
        wineE_rating: '8',
        wineE_comments: 'Nice',
        wineF_rating: '7',
        wineF_comments: 'Good',
        wineG_rating: '9',
        wineG_comments: 'Amazing bubbles',
        wineH_rating: '8',
        wineH_comments: 'Very nice',
        wineI_rating: '10',
        wineI_comments: 'Perfect!'
      })
    }
  };

  Logger.log(doPost(testData));
}
