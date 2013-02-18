// (c) Infocatcher 2010
// version 0.1.1 - 2010-11-01

// Usage:
//   [%SystemRoot%\system32\wscript.exe] backup.js fileOrFolderToBackup
// Force use ZIP:
//   backup.js :zip-on fileOrFolderToBackup
// Force don't use ZIP:
//   backup.js :zip-off fileOrFolderToBackup
// In Total Commander you can use %S argument.

//== Settings begin
var BACKUPS_DIR = "_backups\\"; // Or ""
var ZIP = true;
var ZIP_CMD = '"%COMMANDER_PATH%\\arch\\7-Zip\\7zG.exe" a -tzip -mx9 -ssw -- "<newFile>" "<curFile>"';
var checkEnvVars = true;
//== Settings end

var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.Shell");

var argsCount = WScript.Arguments.length;
if(!argsCount) {
	wsh.Popup(
		"Usage:\n\"" + WScript.ScriptName + "\" file1 file2", -1,
		"Wrong command line arguments! Ц " + WScript.ScriptName,
		48 /* icon exclamation */
	);
	WScript.Quit();
}

var zipCmd = wsh.ExpandEnvironmentStrings(ZIP_CMD);
if(checkEnvVars && zipCmd == ZIP_CMD) {
	wsh.Popup(
		"Ќе удалось развернуть переменные окружени€:\n" + ZIP_CMD, -1,
		"ќшибка! Ц " + WScript.ScriptName,
		16 /* icon stop */
	);
	WScript.Quit();
}

if(BACKUPS_DIR && !fso.FolderExists(BACKUPS_DIR))
	fso.CreateFolder(BACKUPS_DIR);

for(var i = 0; i < argsCount; i++) {
	var file = WScript.Arguments(i);
	if(file.charAt(0) == ":") {
		file = file.toLowerCase();
		if(file == ":zip-on")
			ZIP = true;
		else if(file == ":zip-off")
			ZIP = false;
		continue;
	}
	backup(file);
}

function backup(file) {
	var isFile = fso.FileExists(file);
	var backupName = BACKUPS_DIR + getBackupName(file, isFile && !ZIP);
	if(ZIP) {
		var cmd = zipCmd
			.replace(/<newFile>/g, backupName + ".zip")
			.replace(/<curFile>/g, file);
		wsh.Exec(cmd);
	}
	else {
		fso[isFile ? "CopyFile" : "CopyFolder"](file, backupName);
	}
}
function getBackupName(file, isFile) {
	if(isFile && /\.[^.]+$/.test(file)) {
		var fileName = RegExp.leftContext;
		var fileExt = RegExp.lastMatch;
	}
	else {
		var fileName = file;
		var fileExt = "";
	}

	var d = new Date();
	var timestamp = "_" + d.getFullYear()     + "-" + zeros(d.getMonth() + 1) + "-" + zeros(d.getDate())
	              + "_" + zeros(d.getHours()) + "-" + zeros(d.getMinutes())   + "-" + zeros(d.getSeconds());

	return fileName + timestamp + fileExt;
}
function zeros(n) {
	if(n > 9)
		return n;
	return "0" + n;
}