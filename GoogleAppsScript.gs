function onEdit(e) {
  var sheet = e.source.getActiveSheet(); 
  var range = e.range;  
  var editedColumn = range.getColumn();  
  var sheetName = sheet.getName();
  if (sheetName === "Преподаватели")
  {
  var statusColumn = 14;  
  var agentColumn = 15;   
  var contactDateColumn = 9;
  var contactTimeColumn = 10;
  var retriesColumn = 18;
  var resultColumn = 19;
  var postIdColumn = 20;
  var managersChat = false;
  }
  else if (sheetName === "Менеджеры")
  {
  var statusColumn = 12;  
  var agentColumn = 13;   
  var contactDateColumn = 9;
  var contactTimeColumn = 10;
  var retriesColumn = 16;
  var resultColumn = 17;
  var postIdColumn = 18;  
  var managersChat = true;
  }
 var columnsOfInterest = [
    contactDateColumn,
    contactTimeColumn,
    statusColumn,
    agentColumn,
    retriesColumn,
    resultColumn
  ];

 if (columnsOfInterest.indexOf(editedColumn) !== -1) {
    var row = range.getRow();  
   var newValue = e.value;
    
    var postId = sheet.getRange(row, postIdColumn).getValue();
    
    if (postId) {
       var columnName = '';
      switch (editedColumn) {
        case contactDateColumn:
          columnName = 'contactDate';
          break;
        case contactTimeColumn:
          columnName = 'contactTime';
          break;
        case statusColumn:
          columnName = 'Status';
          break;
        case agentColumn:
          columnName = 'Agent';
          break;
        case retriesColumn:
          columnName = 'Retries';
          break;
        case resultColumn:
          columnName = 'Result';
          break;
        default:
          columnName = 'dick_knows';
      }
      
      sendMattermostReply(postId, columnName, newValue, managersChat);
    }
  }
}

function sendMattermostReply(postIdS, column_name, new_value, isManagersChat) {
  var mattermostUrl = "http://YourServer:Port/BotPanel/SendUpdateOfPost";  
  var token = "p]QV3G$mn6T0";  

  // Create the payload to send to Mattermost
  var payload = JSON.stringify({
    secret: token,
    columnName : column_name,
    newValue : new_value,
    postId: postIdS,
    managersChat: isManagersChat
  });

  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json"
    },
    "payload": payload
  };

  UrlFetchApp.fetch(mattermostUrl, options);
}


