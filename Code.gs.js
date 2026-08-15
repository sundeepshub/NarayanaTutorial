function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var enquiryId = "ENQ-" + new Date().getTime();
    var timestamp = new Date();
    
    // Append row matching the form backend structure
    sheet.appendRow([
      enquiryId,
      timestamp,
      data.student_name || '',
      data.class_grade || '',
      data.board || '',
      data.school_name || '',
      data.subjects || '',
      data.preferred_batch || '',
      data.parent_name || '',
      data.mobile_number || '',
      data.student_location || '',
      data.colony_area || '',
      data.complete_address || '',
      data.rating_reading || '',
      data.rating_writing || '',
      data.rating_learning || '',
      data.rating_listening || '',
      data.rating_understanding || '',
      data.rating_participation || '',
      data.rating_punctuality || '',
      data.rating_discipline || '',
      data.additional_info || '',
      "New",
      data.source || 'Website',
      data.submissionType || 'Online'
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({"result": "success", "enquiryId": enquiryId}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"result": "error", "error": error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}