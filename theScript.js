
//Node js script to convert a docx file to Google Doc and create a shortcut in its place

var ncp=require('copy-paste');
var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var ws=require('windows-shortcuts');
var child_process=require('child_process');
var args=process.argv.slice(2);
if(args.length==0){
	console.log('Enter a file path');
	process.exit();
}
var driveFolderName='GConvert';

var mtMap = {'xls':'application/vnd.google-apps.spreadsheet','xlsx':'application/vnd.google-apps.spreadsheet','doc':'application/vnd.google-apps.document','docx':'application/vnd.google-apps.document','pptx':'application/vnd.google-apps.presentation','ppt':'application/vnd.google-apps.presentation','txt':'text/plain'};

var filePath = args[0];
var fileName=filePath.substring(filePath.lastIndexOf('\\')+1);
var fileExtension=fileName.replace(/^.*\.(.*)$/,"$1");
var fileMainName=filePath.replace(/\.(.*)?$/,'');
/*
console.log('>>> filePath '+filePath);
console.log('>> fileName '+fileName);
console.log('>> fileExtnesion '+fileExtension);
console.log('fileMainName '+fileMainName);
*/

// If modifying these scopes, delete your previously saved credentials
// at ~/creds/gToken.json
var SCOPES = ['https://www.googleapis.com/auth/drive'];
var BASE_DIR=(process.env.HOME||'C:\\Windows\\temp')+'\\';
var TOKEN_DIR = BASE_DIR+'creds\\';
var TOKEN_PATH = TOKEN_DIR + 'gToken.json';
//console.log(TOKEN_PATH);
console.log('Converting....');
// Load client secrets from a local file.
fs.readFile(BASE_DIR+'client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log('Error loading client secret file: ' + err);
    console.log('Please download and copy client_secret.json to your HOME folder'); 
    setTimeout(function(){},5000);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Drive API.
 // authorize(JSON.parse(content), listFiles);
  authorize(JSON.parse(content), listFiles);
});
/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  ncp.copy(authUrl);
  console.log('The url has already been copied to your clipboard for your convenience');
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

/**
 * Check whether the folder we are looking for to store the files, exist. If not call createFolder. 
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function listFiles(auth) {
  var service = google.drive('v3');
  service.files.list({
    q:"name='"+driveFolderName+"'",
    auth: auth,
    pageSize: 10,
    fields: "nextPageToken, files(id, name)"
  }, function(err, response) {
    if (err) {
      console.log('The API returned an error: ' + err);
      return;
    }
    var files = response.files;
    if (files.length == 0) {
      console.log('No files found.');
      createFolder(auth);
    } else {
   /*   for (var i = 0; i < files.length; i++) {
        var file = files[i];
        console.log('%s (%s)', file.name, file.id);
      }*/
      insertFile(auth,files[0].id);
    }
  });
}
/**
 * Create the folder to store the converted files in your root folder (My Drive) on Google Drive 
 * 
 */ 

function createFolder(auth){
	var service=google.drive('v3');
	service.files.create({
		auth: auth,
		resource:{
			'name':driveFolderName,
			'mimeType':'application/vnd.google-apps.folder'
		},
		fields:'id'
	},function(err,file){
		if(err){
			console.log('Could not create folder '+err);
			insertFile(auth,'');
		}
		else
			insertFile(auth,file.id);
	});
}

/**
 * Upload the selected file to drive, convert it if applicable,create a shortcut to the file in the same folder and open the file
 *
 *
 */
function insertFile(auth,pFolderId) {
        var service=google.drive('v3');
        service.files.create({
            auth: auth,
	    fields:'id,webViewLink',
            resource:{
//		   mimeType:'application/vnd.google-apps.document',
		   mimeType:mtMap[fileExtension]||'application/octet-stream',
		   name:fileName,
		   parents:[pFolderId]
	    },
	    mediaUrl:'https://www.googleapis.com/upload/drive/v3/files',
            media:{
//                mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                body:fs.createReadStream(filePath)
            }
        },function (err,resp) {
           if(err){
               console.log('An error occured '+err);
           }else{
               console.log('upload succeeded');
               console.log(JSON.stringify(resp));
//		ncp.copy(resp.webViewLink);
		ws.create(fileMainName+'.g'+fileExtension+'.lnk',
			{target:'chrome.exe',args: '--app="'+resp.webViewLink+'"',runStyle:ws.MIN},
			function(err){
				if(err)
					console.log('An eror occured '+err);
				else
					child_process.execFile('chrome.exe',["--app="+resp.webViewLink]);
			}); 
           } 
        });

}

