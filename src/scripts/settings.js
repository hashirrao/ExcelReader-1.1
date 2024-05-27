const electron = require('electron')
const remote = electron.remote
const fs = require('fs') 
var dialog = remote.require('electron').dialog;

function closebutton() {
    var window = remote.getCurrentWindow();
    window.close()
}

function minimizeButton() {
    var window = remote.getCurrentWindow();
    window.minimize()
}

function browsefilebuttonclick() {
    var address = dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [
            { name: 'Excel', extensions: ['xlsx', 'xls'] }
        ]
    });
    if (address != null) {        
        // Requiring fs module in which 
        // writeFile function is defined. 
        
        // Data which will write in a file. 
        let data = address;
        
        // Write data in 'Output.txt' . 
        fs.writeFile('Output.txt', data, (err) => { 
            document.getElementById("addresslabel").innerHTML = data;
            // In case of a error throw err. 
            if (err) throw err; 
        }) 
    }
}

var address = fs.readFileSync('Output.txt').toString().split("\n");
document.getElementById("addresslabel").innerHTML = address[0];