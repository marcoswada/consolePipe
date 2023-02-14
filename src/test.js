function main(){
    var wsh = new ActiveXObject("WScript.Shell");
    var wshOut = wsh.exec("..\\build\\readdyp.exe Q").StdOut;
    var s=wshOut.ReadAll();
    WScript.echo (s);
}
main();