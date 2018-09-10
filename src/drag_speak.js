'use strict';

if (WScript.Arguments(0) === "/ListVoices"){
    ShowVoices();
}else if(WScript.Arguments(0) === "/PlaySpeakTextFile"){
    PlaySpeakTextFile(WScript.Arguments(1), WScript.Arguments(2));
}else if(WScript.Arguments(0) === "/SaveSpeakTextFile"){
    SaveSpeakTextFile(WScript.Arguments(1), WScript.Arguments(2), WScript.Arguments(3));
}




function SaveSpeakTextFile(nVoice, sFromFilePath, sToFilePath){
  var oFileSystem = new ActiveXObject("Scripting.FileSystemObject");
  var oText = oFileSystem.OpenTextFile(sFromFilePath, 1, false);
  var sManuscript = oText.ReadAll();
  oText.Close();
  SaveSpeakString(nVoice, sManuscript, sToFilePath);
}

function SaveSpeakString (nVoice, sManuscript, sFilePath) {
    var oFile = new ActiveXObject("SAPI.SpFileStream");
    var oSpVoice = new ActiveXObject("SAPI.SpVoice");
    oSpVoice.Voice = oSpVoice.GetVoices().Item(nVoice);
    oSpVoice.Rate = 1;
    
    oFile.Open(sFilePath, 3);
    oSpVoice.AudioOutputStream = oFile;
    oSpVoice.Speak(sManuscript);
    
    oFile.Close();
}


function PlaySpeakTextFile(nVoice, sFileName) {
    var oFileSystem = new ActiveXObject("Scripting.FileSystemObject");
    var oText = oFileSystem.OpenTextFile(sFileName, 1, false);
    var sManuscript = oText.ReadAll();
    oText.Close();
    PlaySpeakString(nVoice, sManuscript); // change 1st argument to change voice.
}

function PlaySpeakString(nVoice, sManuscript){
    var oSpVoice = new ActiveXObject("SAPI.SpVoice");
    oSpVoice.Voice = oSpVoice.getVoices().Item(nVoice);
    oSpVoice.Rate = 1;
    oSpVoice.Speak(sManuscript);
}


function ShowVoices() {
    var oSpVoice = new ActiveXObject("SAPI.SpVoice");
    for (var i = 0; i<oSpVoice.GetVoices().Count; ++i){
	var oVoice = oSpVoice.GetVoices().Item(i);
	WScript.Echo (i + ": " + oVoice.GetDescription());
    }
}

