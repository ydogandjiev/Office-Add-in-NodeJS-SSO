// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/
Office.initialize = function (reason) {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    // Add any initialization logic to this function.
    $("#getGraphAccessTokenButton").click(function () {
      getOneDriveFiles();
    });
  });
}

var timesGetOneDriveFilesHasRun = 0;
var triedWithoutForceConsent = false;
var timesMSGraphErrorReceived = false;

function getOneDriveFiles() {
  timesGetOneDriveFilesHasRun++;
  triedWithoutForceConsent = true;
  getDataWithToken({ forceConsent: false });
}

function getDataWithToken(options) {
  Office.context.auth.getAccessTokenAsync(options,
    function (result) {
      if (result.status === "succeeded") {
        accessToken = result.value;
        getData("/api/values", accessToken);
      } else {
        handleClientSideErrors(result);
      }
    });
}

function getData(relativeUrl, accessToken) {
  $.ajax({
    url: relativeUrl,
    headers: { "Authorization": "Bearer " + accessToken },
    type: "GET"
  })
    .done(function (result) {
      showResult(result);
    })
    .fail(function (result) {
      handleServerSideErrors(result);
    });
}

function handleClientSideErrors(result) {
  switch (result.error.code) {
    case 13001:
      getDataWithToken({ forceAddAccount: true });
      break;

    case 13002:
      if (timesGetOneDriveFilesHasRun < 2) {
        showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
      } else {
        logError(result);
      }
      break;

    case 13003:
      showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
      break;

    case 13006:
      showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
      break;

    case 13007:
      showResult(['That operation cannot be done at this time. Please try again later.']);
      break;

    case 13008:
      showResult(['Please try that operation again after the current operation has finished.']);
      break;

    case 13009:
      if (triedWithoutForceConsent) {
        showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account.']);
      } else {
        getDataWithToken({ forceConsent: false });
      }
      break;

    default:
      logError(result);
      break;
  }
}

function handleServerSideErrors(result) {
  if (result.responseJSON.error.innerError
    && result.responseJSON.error.innerError.error_codes
    && result.responseJSON.error.innerError.error_codes[0] === 50076) {
    getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
  }
  else if (result.responseJSON.error.innerError
    && result.responseJSON.error.innerError.error_codes
    && result.responseJSON.error.innerError.error_codes[0] === 65001) {
    getDataWithToken({ forceConsent: true });
  }
  else if (result.responseJSON.error.innerError
    && result.responseJSON.error.innerError.error_codes
    && result.responseJSON.error.innerError.error_codes[0] === 70011) {
    showResult(['The add-in is asking for a type of permission that is not recognized.']);
  }
  else if (result.responseJSON.error.name
    && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1) {
    showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
  }
  else if (result.responseJSON.error.name
    && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
    if (!timesMSGraphErrorReceived) {
      timesMSGraphErrorReceived = true;
      timesGetOneDriveFilesHasRun = 0;
      triedWithoutForceConsent = false;
      getOneDriveFiles();
    } else {
      logError(result);
    }
  }
  else {
    logError(result);
  }
}

// Displays the data, assumed to be an array.
function showResult(data) {
  // Note that in this sample, the data parameter is an array of OneDrive file/folder
  // names. Encoding/sanitizing to protect against Cross-site scripting (XSS) attacks
  // is not needed because there are restrictions on what characters can be used in 
  // OneDrive file and folder names. These restrictions do not necessarily apply 
  // to other kinds of data including other kinds of Microsoft Graph data. So, to 
  // make this method safely reusable in other contexts, it uses the jQuery text() 
  // method which automatically encodes values that are passed to it.
  $.each(data, function (i) {
    var li = $('<li/>').addClass('ms-ListItem').appendTo($('#file-list'));
    var outerSpan = $('<span/>').addClass('ms-ListItem-secondaryText').appendTo(li);
    $('<span/>').addClass('ms-fontColor-themePrimary').appendTo(outerSpan).text(data[i]);
  });
}

function logError(result) {

  // Error messages can have a variety of structures depending on the ultimate
  // ultimate source and how intervening code restructures it before relaying it.
  console.log("Status: " + result.status);
  if (result.error.code) {
    console.log("Code: " + result.error.code);
  }
  if (result.error.name) {
    console.log("Code: " + result.error.name);
  }
  if (result.error.message) {
    console.log("Code: " + result.error.message);
  }
  if (result.responseJSON.error.name) {
    console.log("Code: " + result.responseJSON.error.name);
  }
  if (result.responseJSON.error.name) {
    console.log("Code: " + result.responseJSON.error.name);
  }
}
