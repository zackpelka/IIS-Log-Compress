/*
====================================================================
  script: iislogcompress.js
  authored: matthew corpolongo
  email: matt.corpolongo@us-resources.com
  description: to be used in scheduled jobs to compress log files 
               that are older than a supplied number of days.
               these files are then zipped into a .zip file using
               zip.exe (available from http://www.info-zip.org/)
               with the .zip having the name:
               machinename-servicename-year.zip
  usage: 
    cscript iislogcompress.js <numberofdaystokeep> <fullpathtologs>
  example usage:
    cscript iislogcompress.js 14 c:\windows\system32\logfiles\w3svc1\
  history:
					07/07/2004
						- script created
					07/09/2004
						- modified script to work on servers with many services
						  running, in a much easier fashion.  now the script
						  should be targeted to the folder where all the sub
						  folders for each service reside.
						  ex: c:\winnt\system32\logfiles\
						    where the folders w3svc1,2,3.... reside
					07/12/2004
						- made changes to the script, because the testing of the
						  zip file, done by the -T argument, runs the unzip.exe
						  to test the file and make sure it is not corrupted. 
						  this would fail, because it appears zip.exe does not
						  call unzip.exe with quotes around the .zip file's name
						  meaning it fails if there are any spaces in the path
						  to the .zip.  so, instead the script now uses start.exe
						  to start in the directory that the .zip resides, and 
						  now as long as there is no space in the .zip file's
						  name, it should work fine.
  to do: nothing
====================================================================
*/


// change this line below to be the full path to zip.exe.  this folder
// should also have unzip.exe.  make sure you use two '\' for each
// folder dilimeter, so that '\' show up in the path.
// this path should NOT have any spaces in it.
//pathToZipExe = "c:\\unixutilities\\usr\\local\\wbin\\zip.exe";
pathToZipExe = "d:\\scripts\\iislogcompress\\zip.exe";

shell = WScript.createobject("WScript.shell");

// getting the machine name
network = WScript.createObject("WScript.network");
computerName = network.computername;

// command line argument stuff

if((WScript.Arguments.length != 2) || (isNaN(WScript.Arguments(0))) || (WScript.Arguments(0) == 0)  || (WScript.Arguments(0).indexOf(".") != -1) || (WScript.Arguments(0) <= -1)){
	WScript.echo("USAGE:");
	WScript.echo("You must enter only 2 command line arguments.");
	WScript.echo("The first argument must be a number, representing the number of days old a file must be to be zipped up.");
	WScript.echo("It must be greater than 0, and be an integer value (i.e. no '.' in it)");
	WScript.echo("The second argument must be the path to where the logs reside");
	WScript.echo("e.g. c:\\windows\\system32\\logfiles\\");
}
else{
	var daysToGoBack = WScript.Arguments(0);
	var logFileDirectory = WScript.Arguments(1);
	if(logFileDirectory.charAt(logFileDirectory.length-1) != "\\"){
		logFileDirectory = logFileDirectory + "\\";
	}
	
	// there are 86400000 miliseconds in one day
	var fullTargetDate = new Date(new Date() - (86400000 * daysToGoBack));
	fullTargetDate = new Date(fullTargetDate.getYear(), fullTargetDate.getMonth(), fullTargetDate.getDate());
	
	// have to add 1 to the month, because for some reason, jscript counts 0 to 11 for months
	targetMonth = (fullTargetDate.getMonth() + 1).toString();
	targetDate = fullTargetDate.getDate().toString();
	targetYear = fullTargetDate.getYear().toString();
	
	if(targetMonth.length != 2){
		// the month is less than 10, so it needs a leading 0
		targetMonth = "0" + targetMonth;
	}
	
	if(targetDate.length != 2){
		// the date is less than 10, so it needs a leading 0
		targetDate = "0" + targetDate;
	}
	
	// trim off the leading first two digits off of the year, so that 2004 becomes 04
	targetYear = targetYear.substr(2,2);
	
	// ok, find all sub folders inside of the parent folder, and loop over
	// them, going inside of each one to zip up the files.
	var fso, f, fc;
	fso = new ActiveXObject("Scripting.FileSystemObject");
	f = fso.GetFolder(logFileDirectory);
	fc = new Enumerator(f.SubFolders);
	for (;!fc.atEnd(); fc.moveNext())
	{
		// store the current directory into a variable. then check to see if
		// it has a trailing "/" on it. if not (which from my testing it looks
		// like it won't, add one on.
		// we must first add the empty string on to the end of it though
		// to change it into a string, instead of a folder object.
		currentWorkingDirectory = fc.item() + "";
		if(currentWorkingDirectory.charAt(currentWorkingDirectory.length-1) != "\\"){
			currentWorkingDirectory = currentWorkingDirectory + "\\";
		}
// ********************************************
		// get a list of the files in the directory	
		var objFileSystem, objFolder;
		objFileSystem = new ActiveXObject("Scripting.FileSystemObject");
		objFolder = objFileSystem.GetFolder(currentWorkingDirectory);
		objFileEnumerator = new Enumerator(objFolder.files);
		for (; !objFileEnumerator.atEnd(); objFileEnumerator.moveNext())
		{
			tempFileName = objFileEnumerator.item().Name;
			tempFilePath = objFileEnumerator.item().Path;
			tempvar = /^(in|ex)(\d\d)(\d\d)(\d\d)\.log$/i.exec(tempFileName);
			if( tempvar != null){
				// have to make sure that there is/are file(s) there to work with
				// else there will be some problems when using the regular expressions
				
				// ok, here is the regular expression
				// what it does is look for files that are named exXXXXXX.log
				// \d means look for a number.
				// the () around pairs of '\d' groups them together so we can
				// later get the year, month and date
				var match = /^(in|ex)(\d\d)(\d\d)(\d\d)\.log$/i.exec(tempFileName);
	
				// have to put the "20" back on the front of the year, and need to 
				// subtract 1 from the month, because for some reason it expects
				// the month to be between 0 and 11
				// then we make a date object out of the date we get from using
				// the regular expression on the file name.  then we are able
				// to compare the date of the file to the target date.
				tempFileDate = new Date("20" + match[2], match[3]-1, match[4]); 
				
				if(tempFileDate < fullTargetDate){
					// the file is older, so it should be added to the .zip
					// create command line to run.  store all files by year in a .zip
					
					tempPathForServiceName = tempFilePath.substr(0,tempFilePath.lastIndexOf("\\"));
					serviceName = tempPathForServiceName.substr(tempPathForServiceName.lastIndexOf("\\")+1);
					zipFileName = computerName + "-" + serviceName + "-20" + match[2] + ".zip";
					WScript.echo("adding " + tempFileName + " to the zip file " + zipFileName);
					
					// what the switches on zip.exe do:
					//    -9 means best compression
					//    -m means move the original files (delete it after putting it in the .zip
					//    -T means test the .zip before deleting the original file
					//    -D means do not add directories, although this isn't really needed for this script
					//        since we look for only names that end in .log
					//    -j means do not add the full path in to the zip.  if this wasn't there, then adding
					//        the file c:\windows\system32\logfiles\w3svc1\ex040707.log would mean the file
					//        would be stored in the zip with the directory structure of
					//          windows\
					//            system32\
					//              logfiles\
					//                w3svc1\
					//                  ex040707.log
					
					var cmdToRun = "cmd.exe /C start /B /WAIT /D \"" + currentWorkingDirectory + "\" " + pathToZipExe + " -9 -m -T -D -j " + zipFileName + " \"" + tempFilePath + "\"";
					//var cmdToRun = "cmd.exe /C zip.exe -9 -m -T -D -j \"" + currentWorkingDirectory + zipFileName + "\" \"" + tempFilePath + "\"";
					//var cmdToRun = "cmd.exe start dir";
					//WScript.echo(cmdToRun);
					shell.run(cmdToRun, 0, "True");
				}
				match[0] = "";
				match[1] = "";
				match[2] = "";
				match[3] = "";
				match[4] = "";
				match="";
			}
		}
// ********************************************
	}
}