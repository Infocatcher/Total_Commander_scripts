// (c) Infocatcher 2010, 2012
// version 0.1.2 - 2012-12-15

// Usage in Total Commander:
// Command:     "%SystemRoot%\system32\wscript.exe" "%COMMANDER_PATH%\scripts\winMergeCompare.js"
// Parameters:  %S %T%M
// Start path:  <empty>

var wsh = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

var winMerge = "%ProgramFiles%\\WinMerge\\WinMergeU.exe";
if(!fso.FileExists(wsh.ExpandEnvironmentStrings(winMerge)))
	winMerge = "%COMMANDER_PATH%\\..\\WinMergePortable\\WinMergePortable.exe";

var compareCmd = '"' + winMerge + '" "<f1>" "<f2>"';
var checkEnvVars = true;

var argsCount = WScript.Arguments.length;
if(argsCount < 2) {
	wsh.Popup(
		"Usage:\n" + WScript.ScriptName + " file1 file2",
		-1,
		"Wrong command line arguments! – " + WScript.ScriptName,
		48 /*MB_ICONEXCLAMATION*/
	);
	WScript.Quit();
}

var cmd = wsh.ExpandEnvironmentStrings(compareCmd);
if(checkEnvVars && compareCmd.indexOf("%") != -1 && cmd == compareCmd) {
	wsh.Popup(
		"Can't expand environment variables:\n" + compareCmd,
		-1,
		"Error! – " + WScript.ScriptName,
		16 /*MB_ICONERROR*/
	);
	WScript.Quit();
}

var file1 = WScript.Arguments(0);
var file2 = WScript.Arguments(1);

try {
	// Always compare old file with new
	if(new Date(fso.GetFile(file1).DateLastModified) > new Date(fso.GetFile(file2).DateLastModified)) {
		var tmp = file1;
		file1 = file2;
		file2 = tmp;
	}
}
catch(e) {
}

cmd = cmd
	.replace(/<f1>/g, file1)
	.replace(/<f2>/g, file2);
wsh.Exec(cmd);