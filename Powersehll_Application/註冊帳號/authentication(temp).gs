/**
 * 1.Download the program to register first, link as below:
 *   https://drive.google.com/file/d/11ppvCaE1zSz-PLP9epgUNNJlbHGhx8nN/view?usp=sharing
 * 
 * 2.After registered, you will receive the email from administrator contains:
 *   - your username and password(protected)
 *   - Scripting ID for Library
 * 
 * 3.When you want to use Library at first time, you should switch your authentication to true,
 *   please select the class "executionPolicy", and send the parms as below:
 *   @{
 *   username = 'yourusername'
 *   password = 'yourpassword got from our email'
 *    }
 * 
 * 4.If you have any question, please contact me:
 *   hayashireii@yahoo.co.jp
 *   Thank you !
 */

 //passwordLibrary from mhjpj73 => 00.PasswordLibrary
var clientsPasswordLibrary = '1PVxiE5y6K5BmParWGj78jjyluK2RBbIR5WZ598U__Q8';

function authenticate(receivedata) {
  var username = receivedata.username;
  var password = receivedata.password;

  var ausheet = SpreadsheetApp.openById(clientsPasswordLibrary).getSheetByName('Clients');
  var audata = ausheet.getRange('A:B').getValues();
  var isAuthenticated = false;

  for (var i = 0; i < audata.length; i++) {
    if (audata[i][0] === username && audata[i][1] === password) {
      isAuthenticated = true;
      break;
    }
  }

  if (isAuthenticated) {
    updateUserState(username, true);
  } else {
    updateUserState(username, false);  // 當認證失敗時，更新狀態為 False
  }
  
  // 記錄驗證行為
  trackAuthentication(username);
  
  return isAuthenticated;
}

function updateUserState(username, isAuthenticated) {
  var ausheet2 = SpreadsheetApp.openById(clientsPasswordLibrary).getSheetByName('UserState');
  var audata2 = ausheet2.getRange('A:B').getValues();
  var found = false;

  for (var i = 0; i < audata2.length; i++) {
    if (audata2[i][0] === username) {
      ausheet2.getRange(i + 1, 2).setValue(isAuthenticated);  // 更新狀態
      found = true;
      break;
    }
  }

  if (!found) {
    ausheet2.appendRow([username, isAuthenticated]);
  }
}

function checkUserState(username) {
  var chsheet3 = SpreadsheetApp.openById(clientsPasswordLibrary).getSheetByName('UserState');
  var chdata3 = chsheet3.getRange('A:B').getValues();

  for (var i = 0; i < chdata3.length; i++) {
    if (chdata3[i][0] === username) {
      return chdata3[i][1];
    }
  }
  return false;
}

function trackAuthentication(username) {
  var trackSheet = SpreadsheetApp.openById(clientsPasswordLibrary).getSheetByName('Track');
  var timestamp = new Date();  // 獲取當前時間
  trackSheet.appendRow([username, timestamp]);
}

function executionPolicy(authDetails) {
  var currentState = checkUserState(authDetails.username);
  if (currentState === true) {
    return "You have got the authentication already!";
  }

  if (authDetails && authDetails.username && authDetails.password) {
    if (authenticate(authDetails.username, authDetails.password)) {
      return "Access granted. You can now use the protected Library";
    } else {
      return "Invalid credentials, please try again.";
    }
  } else {
    return "Authentication required, please provide your username and password.";
  }
}
