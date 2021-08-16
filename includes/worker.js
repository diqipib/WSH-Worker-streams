// Main worker object
var Worker = (function(fileName){
	var wshShell = new ActiveXObject('WScript.Shell'),
		fso = new ActiveXObject('Scripting.FileSystemObject'),
		document = new ActiveXObject('htmlfile'),
		window = document.parentWindow,
		wmiService = GetObject('winmgmts:root/cimv2'),
		instances = 0,
		storageFileName = "streams.db";
	
	// Removing storage if it exists
	if(fso.FileExists(storageFileName)) deleteStorage();
	
	// Function for removing main storage file
	function deleteStorage(){
		try {
			fso.DeleteFile(storageFileName,true);
		} catch(e) {
			WSH.Echo('Failed to delete storage')
		};
	}
	
	return function(fileName){
		var context = this,
			sid = new Date().getTime().toString(36) + Math.random().toString(36).substr(2),
			pid = 0;
		
		instances++;
			
		wshShell.Environment('Process')("sid") = sid;
		
		var	wshExec = wshShell.Exec('wscript.exe "' + fileName + '"'),
			input = fso.OpenTextFile(storageFileName + ":" + sid + "@in",2,true),
			output = fso.OpenTextFile(storageFileName + ":" + sid + "@out",1,true);
		
		// Setting storage file attributes - hidden and system
		with(fso.GetFile(storageFileName)){
			Attributes = Attributes | 2 | 4;
		}
		// Getting current process pid if it is not set
		if(!pid) pid = new Enumerator(wmiService.ExecQuery('SELECT parentProcessId FROM Win32_Process WHERE processId=' + wshExec.ProcessID)).item().ParentProcessId
		// Sending process pid to worker
		input.WriteLine(pid);
		// Method for process termination
		this.terminate = function(){
			wmiService.Get('Win32_Process.Handle=' + wshExec.ProcessID).Terminate();
		}
		// Method for sending data to worker
		this.postMessage = function(data){
			input.WriteLine(JSON.stringify(data||''));
		}
		
		var timer = window.setInterval(function(){
			if(wshExec.status == 1){
				window.clearInterval(timer);
				instances--;
				try {
					if(typeof context.onterminate === 'function') context.onterminate(wshExec.ExitCode);
				} finally {
					input.close();
					output.close();
					if(instances <= 0) deleteStorage();
					return
				}
			}
			if(!output.AtEndOfStream){
				var data = output.ReadAll();
				try {
					data = JSON.parse(data);
				} catch(e){}
				if(e) throw new Error('Failed to parse "' + data + '"');
				if(typeof context.onmessage === 'function') context.onmessage(data);
			}
		},100);
	}
})();