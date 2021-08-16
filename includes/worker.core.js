(function(context){
	try {
		var wshShell = new ActiveXObject('WScript.Shell'),
			fso = new ActiveXObject('Scripting.FileSystemObject'),
			sid = wshShell.Environment('Process')('sid'),
			pid = 0,
			storageFileName = "streams.db",
			input = fso.OpenTextFile(storageFileName + ":" + sid + "@out",2,true),
			output = fso.OpenTextFile(storageFileName + ":" + sid + "@in",1,true),
			document = new ActiveXObject('htmlfile'),
			window = document.parentWindow,
			wmiService = GetObject("winmgmts:root/cimv2");
	} catch(e){}
	if(e) throw new Error('Failed to initialize worker ' + e.description);

	if(!output.AtEndOfStream){
		pid = output.ReadLine();
	} else {
		throw new Error('Failed to get parent process id');
	}

	hostCheckTimer = window.setInterval(function(){
		if(wmiService.ExecQuery('SELECT ProcessID FROM Win32_Process WHERE ProcessID=' + pid,'WQL',16).Count <=0){ 
			window.clearInterval(timer);
			window.clearInterval(hostCheckTimer);
			if(typeof onhostterminate === 'function') onhostterminate();
		}
	},1000);

	var timer = window.setInterval(function(){
		if(!output.AtEndOfStream){
			var data = output.ReadAll()
			try {
				data = JSON.parse(data);
			} catch(e){}
			if(e) throw new Error('Failed to parse "' + data + '"');
			if(typeof onmessage === 'function') onmessage(data);
		}
	},100);

	postMessage = function(data){
		input.WriteLine(JSON.stringify(data||''));
	}
	
})(this);