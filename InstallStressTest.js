var WshShell = new ActiveXObject("WScript.Shell");
var fso = new ActiveXObject("Scripting.FileSystemObject");

var testLoop = 1000; //times
var timeOut = 60; //seconds

//argument check
var args = WScript.Arguments;
if (args.length == 0) {
    WScript.Echo("Invalid input.");
    WScript.Quit(1);
}
else if (args.length == 1){
    WScript.Echo(["Usage:", args(0), "<msi file>"].join(" "));
    WScript.Quit(1);
}

//check msi file
var msiFile = args(1);
var msiPath = fso.GetAbsolutePathName([WshShell.CurrentDirectory, msiFile].join("\\"));
if (!fso.FileExists(msiPath)) {
    WScript.Echo("File Not found: " + msiPath);
    WScript.Quit(2);
}

//start test
var logTime = new Date().getTime() / 1000;

//test log
var logPath = WshShell.CurrentDirectory + "\\" + WScript.ScriptName.replace(".js", "_" + logTime + "_test.log");
//install log
var installLogPath = WshShell.CurrentDirectory + "\\" + WScript.ScriptName.replace(".js", "_" + logTime + "_install.log");
//uninstall log
var uninstallLogPath = WshShell.CurrentDirectory + "\\" + WScript.ScriptName.replace(".js", "_" + logTime + "_uninstall.log");

var logFile = null;
try {
    logFile = fso.CreateTextFile(logPath , true);
    testProcess();
}
catch(e) {
    logSpace("Error:", e);
}
finally{
    logFile.Close();
}

function testProcess() {
    logSpace("############Start time:", new Date(), "############");
    logSpace("Test Package:", msiPath);    
    
    var okTotalInstallDuration = 0;
    var okTotalUninstallDuration = 0;
    //loop start
    var successCnt = 0;
    for (var i = 1; i<=testLoop; i++){
        logSpace("Process index:", i, "/", testLoop);
        
        var installRlt = install(i);
        if (!installRlt.isOk) {
            log("Install failed. Stop Test.");
            break;
        }
        
        var uninstallRlt = uninstall(i);
        if (!uninstallRlt.isOk) {
            log("Uninstall failed. Stop Test.");
            break;
        }
        
        log("Install duration: ", installRlt.duration, ", ", "Uninstall duration: ", uninstallRlt.duration);
        
        successCnt++;
        okTotalInstallDuration += installRlt.duration;
        okTotalUninstallDuration += uninstallRlt.duration;
        
    }
    logSpace("----------------------------------------------------");
    logSpace("Summary>");
    
    logSpace("Install average duration(secs):", okTotalInstallDuration / successCnt);
    logSpace("Uninstall average duration(secs):", okTotalUninstallDuration / successCnt);
    logSpace("----------------------------------------------------");
    logSpace("############End time:", new Date(), "############");
}

function log() {
    var msg = "";
    if (arguments.length > 0) {
        var args = [];
        for (var k=0; k<arguments.length; k++) { args.push(arguments[k]); }
        msg = args.join("");
    }
    
    WScript.Echo(msg);
    logFile.WriteLine(msg)
}

function logSpace() {
    var msg = "";
    if (arguments.length > 0) {
        var args = [];
        for (var k=0; k<arguments.length; k++) { args.push(arguments[k]); }
        msg = args.join(" ");
    }
    
    WScript.Echo(msg);
    logFile.WriteLine(msg)
}

function install(index) {
    var cmd = "msiexec /i " + msiPath + " " + "/norestart /quiet /l*v " + installLogPath;
    return execProc(cmd);
}

function uninstall(index) {
    var cmd = "msiexec /x " + msiPath + " " + "/norestart /quiet /l*v " + uninstallLogPath;
    return execProc(cmd);
}

function execProc(cmd) {
    var oExec = WshShell.Exec(cmd);
    var isFinished = false;
    var startTime = new Date();
    var duration = timeOut;
    
    for (var t = timeOut * 10; t>0; t--)
    {
        if (oExec.Status != 0) {
            isFinished = true;
            duration = (new Date().getTime() - startTime.getTime()) / 1000;
            break;
        }
        WScript.Sleep(100);
    }
    if (!isFinished) {
        oExec.Terminate();    
        WScript.Sleep(10*1000);
    }    
    return {isOk: isFinished, duration: duration};    
}
