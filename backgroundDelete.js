// (c) Infocatcher 2010-2014
// version 0.1.6.3 - 2014-02-17

// Rename file(s) or folder(s) (fast!) and delete them.

// Usage:
//   [%SystemRoot%\system32\wscript.exe] backgroundDelete.js fileOrFolderToDelete
// In Total Commander:
//   Command:    wscript.exe "%COMMANDER_PATH%\scripts\backgroundDelete.js"
//   Parameters: %S
//   Start path: <empty>

var confirmDelete = true;
var confirmDeleteMaxItems = 10;
var waitBeforeDelete = 4000; // Rename -> wait -> delete, in milliseconds

function _localize(s) {
	var strings = {
		"Delete in background": {
			ru: "Удалить в фоне"
		},
		"Do you really want to delete %S selected files/directories?\n": {
			ru: "Вы действительно хотите удалить выбранные файлы/каталоги (%S шт.)?\n"
		},
		"All files and directories was successfully moved!\n[This message will be automatically closed in 1 second]": {
			ru: "Все файлы и каталоги успешно перемещены!\n[Данное сообщение будет закрыто через 1 секунду]"
		},
		"Wrong command line arguments!": {
			ru: "Неправильные параметры командной строки!"
		},
		"Usage:\n": {
			ru: "Использование:\n"
		},
		"Error!": {
			ru: "Ошибка!"
		},
		"Couldn't rename\n": {
			ru: "Не удалось переименовать\n"
		},
		"Couldn't delete\n": {
			ru: "Не удалось удалить\n"
		}
	};
	var lng = "en";
	try {
		var window = new ActiveXObject("htmlfile").parentWindow;
		lng = window.navigator.browserLanguage.match(/^\w*/)[0] || lng;
		window = null;
	}
	catch(e) {
	}
	_localize = function(s) {
		return strings[s] && strings[s][lng] || s;
	};
	return _localize(s);
}

var dialogTitle = _localize("Delete in background");
var fso = new ActiveXObject("Scripting.FileSystemObject");
var wsh = new ActiveXObject("WScript.shell");

var argsCount = WScript.Arguments.length;
if(!argsCount) {
	msg(_localize("Usage:\n") + dialogTitle + " file1 file2", _localize("Wrong command line arguments!"), true);
	WScript.Quit();
}

var files = [];
for(var i = 0; i < argsCount; ++i)
	files.push(WScript.Arguments(i));

if(confirmDelete) {
	var btn = wsh.Popup(
		_localize("Do you really want to delete %S selected files/directories?\n")
			.replace("%S", argsCount)
			+ printPaths(files),
		-1,
		dialogTitle,
		3|32 /*MB_YESNOCANCEL|MB_ICONQUESTION*/
	);
	if(btn != 6 /*IDYES*/)
		WScript.Quit();
}

var entries = [];
var moveErrors = [];
var deleteErrors = [];

for(var i = 0; i < argsCount; ++i)
	prepareDelete(files[i]);

if(moveErrors.length)
	msg(_localize("Couldn't rename\n") + cutArray(moveErrors).join("\n"), _localize("Error!"));
else {
	wsh.Popup(
		_localize("All files and directories was successfully moved!\n[This message will be automatically closed in 1 second]"),
		1,
		dialogTitle,
		64 /*MB_ICONINFORMATION*/
	);
}

if(waitBeforeDelete)
	WScript.Sleep(waitBeforeDelete);

for(var i = 0, len = entries.length; i < len; ++i)
	deleteEntry(entries[i]);

if(deleteErrors.length)
	msg(_localize("Couldn't delete\n") + cutArray(deleteErrors).join("\n"), _localize("Error!"));

function getRandomStr() {
	return getRandomStr._cached || (
		getRandomStr._cached = Math.random().toString(36).substr(2)
	);
}
function getRandomName(path) {
	for(;;) {
		var newPath = path.replace(/[^\\\/]+[\\\/]*$/, "~$&~" + getRandomStr() + ".deleted");
		if(!fso.FileExists(newPath) && !fso.FolderExists(newPath))
			return newPath;
		getRandomStr._cached = "";
	}
}
function prepareDelete(path) {
	var newPath = getRandomName(path);
	var type = fso.FileExists(path) ? "File" : "Folder";
	try {
		fso["Move" + type](path, newPath);
		entries.push({
			type: type,
			path: path,
			newPath: newPath
		});
	}
	catch(e) {
		moveErrors.push('"' + path + '": ' + e.message);
	}
}
function deleteEntry(entry) {
	try {
		fso["Delete" + entry.type](entry.newPath, true);
	}
	catch(e) {
		deleteErrors.push('"' + entry.path + '" => "' + entry.newPath + '": ' + e.message);
	}
}

function printPaths(paths, maxItems) {
	paths = cutArray(paths, maxItems);
	for(var i = 0, l = paths.length; i < l; ++i) {
		var path = paths[i];
		if(/\s/.test(path))
			paths[i] = '"' + path + '"';
	}
	return paths.join("\n");
}
function cutArray(arr, maxItems) {
	if(!maxItems)
		maxItems = confirmDeleteMaxItems;
	var len = arr.length;
	if(len <= maxItems)
		return arr.slice(0);
	var ret = [];
	for(var i = 0; i < maxItems; ++i)
		ret.push(arr[i]);
	ret.push("…");
	ret.push(arr[len - 1]);
	return ret;
}
function msg(text, title, warning) {
	wsh.Popup(
		text,
		-1,
		(title ? title + " – " : "") + dialogTitle,
		warning ? 48 /*MB_ICONEXCLAMATION*/ : 16 /*MB_ICONERROR*/
	);
}