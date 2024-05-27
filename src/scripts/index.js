const electron = require('electron')
const path = require('path')
const BrowserWindow = electron.remote.BrowserWindow
const remote = electron.remote

//const CryptoJS = require('crypto-js')
const HtmlTableToJson = require('html-table-to-json');

var xlsx = require("xlsx");
var dialog = remote.require('electron').dialog;


function closebutton() {
    var window = remote.getCurrentWindow();
    window.close()
}

function minimizeButton() {
    var window = remote.getCurrentWindow();
    window.minimize()
}

function loadprogram(){
    var fs = require("fs");
    var address = fs.readFileSync('Output.txt').toString().split("\n");
    if(address != ""){
        loadfiles(address);

        document.getElementById("deletebutton").style.visibility = "visible";
        document.getElementById("savebutton").style.visibility = "visible";
        document.getElementById("editbutton").style.visibility = "visible";
        document.getElementById("filterbutton").style.visibility = "visible";
        document.getElementById("exportbutton").style.visibility = "visible";
        document.getElementById("sheetnames").style.visibility = "visible";
    }
}

loadprogram();

function addfilebuttonclick() {
    var address = dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [
            { name: 'Excel', extensions: ['xlsx', 'xls'] }
        ]
    });
    if (address != null) {
        loadfiles(address);
        
        document.getElementById("deletebutton").style.visibility = "visible";
        document.getElementById("savebutton").style.visibility = "visible";
        document.getElementById("editbutton").style.visibility = "visible";
        document.getElementById("filterbutton").style.visibility = "visible";
        document.getElementById("exportbutton").style.visibility = "visible";
        document.getElementById("sheetnames").style.visibility = "visible";
    }
}

function loadfiles(add) {
    var file = xlsx.readFile(add[0], { cellDates: true });
    document.getElementById("sheetnames").innerHTML = "";
    for (var i = 0; i < file.SheetNames.length; i++) {
        document.getElementById("sheetnames").innerHTML += "<option>" + file.SheetNames[i] + "</option>";
    }
    sheetnameclick(add);
    document.getElementById("sheetnames").onchange = function(){
        sheetnameclick(add);
    }
}

function sheetnameclick(add){
    var data = new Array();
    var file = xlsx.readFile(add[0], { cellDates: true });
    var sheetnames = document.getElementById("sheetnames").value;
    var ws = file.Sheets[sheetnames];
    data[sheetnames] = xlsx.utils.sheet_to_json(ws);

    var headerarr = get_header_row(file.Sheets[sheetnames]);
    var newData = data[sheetnames].map(function (record) {
        return record;
    })
    //console.log(newData);
    //console.log(headerarr);
    str = "<tr><td style='padding:5px'><img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\" /></td>";
    for (var j = 0; j < headerarr.length; j++) {
            str += "<td style='padding:5px'>" + headerarr[j] + "</td>";
    }
    str += "</tr>";
    document.getElementById("sheetheadtable").innerHTML = str;
    
    document.getElementById("sheettable").innerHTML = "";
    str = "";
    for (var i = 0; i < newData.length; i++) {

        str += "<tr><td><img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\" /></td>";
        for (var j = 0; j < headerarr.length; j++) {
            if (headerarr[j] == "__EMPTY_1") {
                if ("<td>" + newData[i]["__EMPTY"] + "</td>" != "<td>undefined</td>") {
                    str += "<td>" + newData[i]["__EMPTY"] + "</td>";
                }
                else {
                    str += "<td></td>";
                }
            }

            if ("<td>" + newData[i][headerarr[j]] + "</td>" != "<td>undefined</td>") {
                str += "<td>" + newData[i][headerarr[j]] + "</td>";
            }
            else {
                str += "<td></td>";
            }
        }
        str += "</tr>";
    }
    document.getElementById("sheettable").innerHTML = str;
    //console.log(newData);
    settablecellclick();
}

function settablecellclick(){
    var t = document.getElementById("sheettable");
    var rIndex;
    var cIndex;
    var counti = 0;
    for(var i = 0; i < t.rows.length; i++){
        t.rows[i].onclick = function(){
            counti++;
                setTimeout(() => {
                    if(counti == 2){
                        //document.getElementById("updaterowpanel").style.visibility = "visible";
                        document.getElementById("updaterowinput").value = this.innerHTML;
                    }
                    counti = 0;
                }, 200);
            rIndex = this.rowIndex;
        }
        t.rows[i].cells[0].onclick = function(){
            if(this.innerHTML == "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">"){
                this.innerHTML = "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\">";
            }
            else{
                this.innerHTML = "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">";
            }
        }
        for(var j = 1; j < t.rows[i].cells.length; j++){
            var countj = 0;
            t.rows[i].cells[j].onclick = function(){
                countj++;
                setTimeout(() => {
                    if(countj == 1){
                        cIndex = this.cellIndex;
                        document.getElementById("updatecellpanel").style.visibility = "visible";
                        document.getElementById("updatecellinput").value = this.innerHTML;
                        document.getElementById("rowindexlabel").innerHTML = rIndex-1;
                        document.getElementById("cellindexlabel").innerHTML = cIndex;
                    }
                    countj = 0;
                }, 200);
            }
        }
    }
    setTimeout(rowcolumnscount, 2000);
}

function get_header_row(sheet) {
    var headers = [];
    var range = xlsx.utils.decode_range(sheet['!ref']);
    var C, R = range.s.r; /* start in the first row */
    /* walk every column in the range */
    for (C = range.s.c; C <= range.e.c; ++C) {
        var cell = sheet[xlsx.utils.encode_cell({ c: C, r: R })] /* find the cell in the first row */

        var hdr = "__EMPTY_" + C; // <-- replace with your desired default 
        if (cell && cell.t) hdr = xlsx.utils.format_cell(cell);

        headers.push(hdr);
    }
    return headers;
}

function deletebuttonclick(){
    var t = document.getElementById("sheettable");
    for(var i=t.rows.length-1; i>=0; i--){
        if(t.rows[i].cells[0].innerHTML == "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">"){
            document.getElementById("deleteconfirmationpanel").style.visibility = "visible";
            break;
        }
    }
}

function deleteconfirmyesbuttonclick(){
    var t = document.getElementById("sheettable");
    for(var i=t.rows.length-1; i>=0; i--){
        if(t.rows[i].cells[0].innerHTML == "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">"){
            t.deleteRow(i);
        }
    }
    sumcolumns();
    rowcolumnscount();
    document.getElementById("deleteconfirmationpanel").style.visibility = "hidden";
}

function deleteconfirmnobuttonclick(){
    document.getElementById("deleteconfirmationpanel").style.visibility = "hidden";
}

function savebuttonclick(){
    const {app} = require('electron');
    var fs = require("fs");
    var address = fs.readFileSync('Output.txt').toString().split("\n");
    var sheetnames = document.getElementById("sheetnames");
    var sheetval = document.getElementById("sheetnames").value;
    var currenttable = document.getElementById("sheettable").cloneNode(true);
    for (var j = 0; j < currenttable.rows.length; j++) {
        currenttable.rows[j].deleteCell(0);
    }
    if (address != "") {
        var newWB = xlsx.utils.book_new();
        for (var i = 0; i < sheetnames.length; i++) {
            sheetnames.value = sheetnames[i].value;
            sheetnames.onchange();
            var originalsheettable = document.getElementById("sheettable");
            var sheettable = originalsheettable.cloneNode(true);
            for (var j = 0; j < sheettable.rows.length; j++) {
                sheettable.rows[j].deleteCell(0);
            }
            if (sheettable.innerHTML != "") {
                if(sheetval == sheetnames[i].value){
                    var jsonTables = new HtmlTableToJson(
                        `<table>` +
                        currenttable.innerHTML
                        + `</table>`    
                    );        
                }
                else{
                    var jsonTables = new HtmlTableToJson(
                        `<table>` +
                        sheettable.innerHTML
                        + `</table>`    
                    );
                }
                var newData = jsonTables.results.map(function (record) {
                    return record;
                });

                var newWS = xlsx.utils.json_to_sheet(newData[0]);
                xlsx.utils.book_append_sheet(newWB, newWS, sheetnames[i].value);
                
                xlsx.writeFile(newWB, "temp.xlsx");
            }
        }
        const moveFile = require('move-file');
        (async () => {
            await moveFile("temp.xlsx", address[0]);
            //console.log('The file has been moved');
        })();
        document.getElementById("sheettable").innerHTML = "";
        setTimeout(function(){
            sheetnames.value = sheetval;
            sheetnames.onchange();
        }, 500)
    }   
}

function editbuttonclick(){
    if(document.getElementById("sheettable").rows.length > 0){
        document.getElementById("editcolumnpanel").style.visibility = "visible";
        loadcolumnstoselect();
    }
}

function loadcolumnstoselect(){
    var t = document.getElementById("sheettable");
    document.getElementById("editcolumnspanelcolumnsnumber").innerHTML = "<option>All</option>";
    document.getElementById("optionspanelfrontcolumnsnumber").innerHTML = "";
    document.getElementById("optionspanelfrontaddcolumnsnumber").innerHTML = "<option>At Last</option>";
    document.getElementById("optionspanelfrontdeletecolumnsnumber").innerHTML = "";
    for(var i = 1; i < t.rows[0].cells.length; i++){
        document.getElementById("editcolumnspanelcolumnsnumber").innerHTML += "<option>"+i+"</option>";
        document.getElementById("optionspanelfrontcolumnsnumber").innerHTML += "<option>"+i+"</option>";
        document.getElementById("optionspanelfrontaddcolumnsnumber").innerHTML += "<option>"+i+"</option>";
        document.getElementById("optionspanelfrontdeletecolumnsnumber").innerHTML += "<option>"+i+"</option>";
    }
}

function loadrowstoselect(){
    var t = document.getElementById("sheettable");
    document.getElementById("optionspanelfrontrowsnumber").innerHTML = "<option>At Last</option>";
    for(var i = 0; i < t.rows.length; i++){
        document.getElementById("optionspanelfrontrowsnumber").innerHTML += "<option>"+i+"</option>";
    }
}
/*
function editcolumnsreplacebuttonclick(){
    
}  */

function editcolumnsreplaceallbuttonclick(){
    var t = document.getElementById("sheettable");
    var x = document.getElementById("editcolumnspanelcolumnsnumber").value;
    var find = document.getElementById("editcolumnfindvalue").value;
    var replace = document.getElementById("editcolumnreplacevalue").value;
    var indexrangefrom = document.getElementById("editcolumnindexrangefrom").value;
    var indexrangeto = parseInt(document.getElementById("editcolumnindexrangeto").value)+1;
    if(x == "All"){
        for (var i = 0; i < t.rows.length; i++) {
            for (var j = 1; j < t.rows[i].cells.length; j++) {
                var string = t.rows[i].cells[j].innerHTML;
                if (indexrangefrom != "" && indexrangeto != "") {
                    var s1 = string.substring(0, indexrangefrom);
                    var s2 = string.substring(indexrangefrom, indexrangeto);
                    var s3 = string.substring(indexrangeto);
                    if (find == s2) {
                        t.rows[i].cells[j].innerHTML = s1 + replace + s3;
                    }
                }
                else {
                    var string = t.rows[i].cells[j].innerHTML;
                    var regex = new RegExp(find, "gi");
                    t.rows[i].cells[j].innerHTML = string.replace(regex, replace);
                }
            }
        }
    }
    else{
        for(var i = 0; i < t.rows.length; i++){     
            var string = t.rows[i].cells[x].innerHTML;
            if(indexrangefrom != "" && indexrangeto != ""){
                var s1 = string.substring(0, indexrangefrom); 
                var s2 = string.substring(indexrangefrom, indexrangeto);
                var s3 = string.substring(indexrangeto);
                
                if(find == s2){
                    t.rows[i].cells[x].innerHTML = s1+replace+s3; 
                }
            }
            else{
                for (var i = 0; i < t.rows.length; i++) {
                    var string = t.rows[i].cells[x].innerHTML;
                    var regex = new RegExp(find,"gi");
                    t.rows[i].cells[x].innerHTML = string.replace(regex, replace);
                }
            }
        }
    }
}

function editpanelclosebutton(){
    document.getElementById("editcolumnpanel").style.visibility = "hidden";
}

function filterbuttonclick() {
    if(document.getElementById("sheettable").rows.length > 0){
        document.getElementById("filterPanel").style.visibility = "visible";
    }
    else{
        alert("Table is empty..");
    }
    
}

function filterpanelclosebutton() {
    document.getElementById("filterPanel").style.visibility = "hidden";
}

function filterinputkeydown(event) {
    if(event.key == "Enter"){
        filtertable();
    }
}

function filtertable(){
    if(document.getElementById("methodselect").value == "method_1"){
        filtertableby1stmethod();
    }
    else if(document.getElementById("methodselect").value == "method_2"){
        filtertableby2ndmethod();
    }
}

function filtertableby1stmethod() {
    var check = false;
    var stb = document.getElementById("sheettable");
    var ftb = document.getElementById("filtertable");
    var val = document.getElementById("filterinput").value;
    if (stb.rows.length > 0) {
        var str = "";
        for (var i = 0; i < stb.rows.length; i++) {
            for (var j = 1; j < stb.rows[i].cells.length; j++) {
                var valarr = val.toLowerCase();
                var tbvalarr = stb.rows[i].cells[j].innerHTML.toLowerCase();
                for (var k = 0; k < valarr.length; k++) {
                    if (valarr[k] == tbvalarr[k]) {
                        check = true;
                    }
                    else {
                        check = false;
                        break;
                    }
                }
                if (check) {
                    str += "<tr>"+stb.rows[i].innerHTML+"</tr>";
                    break;
                }
            }
        }
        ftb.innerHTML = str;
        for (var j = 0; j < ftb.rows.length; j++) {
            ftb.rows[j].deleteCell(0);
        } 
    }
}

function filtertableby2ndmethod() {
    var stb = document.getElementById("sheettable");
    var ftb = document.getElementById("filtertable");
    var val = document.getElementById("filterinput").value;
    var count = 0;
    if (stb.rows.length > 0) {
        var str = "";
        count = 0;
        for (var i = 0; i < stb.rows.length; i++) {
            for (var j = 1; j < stb.rows[i].cells.length; j++) {
                var valarr = val.toLowerCase();
                var tbvalarr = stb.rows[i].cells[j].innerHTML.toLowerCase();
                var l = 0;
                count = 0;
                for (var k = 0; k <= tbvalarr.length; k++) {
                    if(count == valarr.length){
                        str += "<tr>"+stb.rows[i].innerHTML+"</tr>";
                        break;
                    }
                    if (valarr[l] == tbvalarr[k]) {
                        count++;
                        l++;
                    }
                    else {
                        count = 0;
                        l = 0;
                    }
                }
                if(count == valarr.length){
                    break;
                }
            }
        }
        ftb.innerHTML = str;
        for (var j = 0; j < ftb.rows.length; j++) {
            ftb.rows[j].deleteCell(0);
        } 
    }
}

function updatecellpanelclosebutton(){
    document.getElementById("updatecellpanel").style.visibility = "hidden";
}

function updaterowpanelclosebutton(){
    document.getElementById("updaterowpanel").style.visibility = "hidden";
}

function updatecell(){
    var ri = document.getElementById("rowindexlabel").innerHTML;
    var ci = document.getElementById("cellindexlabel").innerHTML;
    document.getElementById("sheettable").rows[ri].cells[ci].innerHTML = document.getElementById("updatecellinput").value;
    sumcolumns();
    document.getElementById("cellindexlabel").innerHTML = ++ci;
    var t = document.getElementById("sheettable");
    if(ci >= t.rows[0].cells.length){
        ci = 1;
        ri++;
        if(ri >= t.rows.length){
            document.getElementById("updatecellpanel").style.visibility = "hidden";
        }
    }
    document.getElementById("updatecellinput").value = t.rows[ri].cells[ci].innerHTML;
    document.getElementById("rowindexlabel").innerHTML = ri;
    document.getElementById("cellindexlabel").innerHTML = ci;
}

function exportbuttonclick(){
    document.getElementById("exportpanel").style.visibility = "visible";
}

function exportpanelclosebutton(){
    document.getElementById("exportpanel").style.visibility = "hidden";
}

function sheetexportbuttonclick(){
    var originalsheettable = document.getElementById("sheettable");
    var sheettable = originalsheettable.cloneNode(true);
    for (var i = 0; i < sheettable.rows.length; i++) {
        sheettable.rows[i].deleteCell(0);
    }
    if(sheettable.innerHTML != ""){
        const options = {
            //defaultPath: app.getPath('documents') + '/electron-tutorial-app.pdf',
            filters: [
                { name: 'Excel', extensions: ['xlsx', 'xls'] }
            ],
        }
        const savePath = dialog.showSaveDialog(null, options);
        if(savePath != null){
            var jsonTables = new HtmlTableToJson(
                `<table>`+
                    sheettable.innerHTML
                +`</table>`
            );
            var newData = jsonTables.results.map(function (record) {
                return record;
            });
            var newWB = xlsx.utils.book_new();
            var newWS = xlsx.utils.json_to_sheet(newData[0]);
            xlsx.utils.book_append_sheet(newWB, newWS, "New Data");
            
            xlsx.writeFile(newWB, savePath);
        }
    }
}

function wholefileexportbuttonclick(){
    var sheetnames = document.getElementById("sheetnames");
    var sheetval = document.getElementById("sheetnames").value;
    var currenttable = document.getElementById("sheettable").cloneNode(true);
    for (var j = 0; j < currenttable.rows.length; j++) {
        currenttable.rows[j].deleteCell(0);
    }
    const options = {
        //defaultPath: app.getPath('documents') + '/electron-tutorial-app.pdf',
        filters: [
            { name: 'Excel', extensions: ['xlsx', 'xls'] }
        ],
    }
    const savePath = dialog.showSaveDialog(null, options);
    if (savePath != null) {
        var newWB = xlsx.utils.book_new();
        for (var i = 0; i < sheetnames.length; i++) {
            sheetnames.value = sheetnames[i].value;
            sheetnames.onchange();
            var originalsheettable = document.getElementById("sheettable");
            var sheettable = originalsheettable.cloneNode(true);
            for (var j = 0; j < sheettable.rows.length; j++) {
                sheettable.rows[j].deleteCell(0);
            }
            if (sheettable.innerHTML != "") {
                if(sheetval == sheetnames[i].value){
                    var jsonTables = new HtmlTableToJson(
                        `<table>` +
                        currenttable.innerHTML
                        + `</table>`    
                    );        
                }
                else{
                    var jsonTables = new HtmlTableToJson(
                        `<table>` +
                        sheettable.innerHTML
                        + `</table>`    
                    );
                }
                var newData = jsonTables.results.map(function (record) {
                    return record;
                });

                var newWS = xlsx.utils.json_to_sheet(newData[0]);
                xlsx.utils.book_append_sheet(newWB, newWS, sheetnames[i].value);

                xlsx.writeFile(newWB, savePath);
            }
        }
        sheetnames.value = sheetval;
        sheetnames.onchange();
    }
}

function exportbuttonfilterclick(){
    var sheettable = document.getElementById("filtertable");
    if(sheettable.innerHTML != ""){
        const options = {
            //defaultPath: app.getPath('documents') + '/electron-tutorial-app.pdf',
            filters: [
                { name: 'Excel', extensions: ['xlsx', 'xls'] }
            ],
        }
        const savePath = dialog.showSaveDialog(null, options);
        if(savePath != null){
            var jsonTables = new HtmlTableToJson(
                `<table>`+
                    sheettable.innerHTML
                +`</table>`
            );
            var newData = jsonTables.results.map(function (record) {
                return record;
            });
            var newWB = xlsx.utils.book_new();
            var newWS = xlsx.utils.json_to_sheet(newData[0]);
            xlsx.utils.book_append_sheet(newWB, newWS, "New Data");
            
            xlsx.writeFile(newWB, savePath);
        }
    }
}

function rowcolumnscount(){
    var sheettable = document.getElementById("sheettable");
    document.getElementById("rowscountlabel").innerHTML = sheettable.rows.length;
    document.getElementById("columnscountlabel").innerHTML = sheettable.rows[0].cells.length-1;
}

function optionbuttonclick(){
    document.getElementById("optionspanelfront").style.visibility = "visible";
    loadcolumnstoselect();
    loadrowstoselect();
}

function optionspanelforntclosebutton(){
    document.getElementById("optionspanelfront").style.visibility = "hidden";
}

function addrowbuttonclick(){
    var t = document.getElementById("sheettable");
    var val = document.getElementById("optionspanelfrontrowsnumber").value;
    if(val == "At Last"){
        var cellslength = t.rows[0].cells.length;
        var row = t.insertRow(t.rows.length);
        row.innerHTML = "<td style='padding:5px'><img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\" /></td>";
        for(var i = 1; i < cellslength; i++){
            row.innerHTML += "<td style='padding:5px'></td>";
        }
    }
    else{
        var cellslength = t.rows[0].cells.length;
        var row = t.insertRow(val);
        row.innerHTML = "<td style='padding:5px'><img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\" /></td>";
        for(var i = 1; i < cellslength; i++){
            row.innerHTML += "<td style='padding:5px'></td>";
        }
    }

    var rIndex;
        var cIndex;
        var counti = 0;
        row.onclick = function(){
            counti++;
                setTimeout(() => {
                    if(counti == 2){
                        document.getElementById("updaterowpanel").style.visibility = "visible";
                        document.getElementById("updaterowinput").value = this.innerHTML;
                    }
                    counti = 0;
                }, 200);
            rIndex = this.rowIndex;
        }
        row.cells[0].onclick = function(){
            if(this.innerHTML == "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">"){
                this.innerHTML = "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-unchecked-checkbox-52.png\">";
            }
            else{
                this.innerHTML = "<img width=\"15px\" height=\"15px\" src=\"../assets/icons/images/icons8-tick-box-52.png\">";
            }
        }
        for(var j = 1; j < row.cells.length; j++){
            var countj = 0;
            row.cells[j].onclick = function(){
                countj++;
                setTimeout(() => {
                    if(countj == 1){
                        cIndex = this.cellIndex;
                        document.getElementById("updatecellpanel").style.visibility = "visible";
                        document.getElementById("updatecellinput").value = this.innerHTML;
                        document.getElementById("rowindexlabel").innerHTML = rIndex-1;
                        document.getElementById("cellindexlabel").innerHTML = cIndex;
                    }
                    countj = 0;
                }, 200);
            }
        }
        setTimeout(rowcolumnscount, 1000);
        loadrowstoselect();
}

function addcolumnbuttonclick(){
    var t = document.getElementById("sheettable");
    var val = document.getElementById("optionspanelfrontaddcolumnsnumber").value;
    if(val == "At Last"){
        var cellslength = t.rows[0].cells.length;
        for(var i = 0; i < t.rows.length; i++){
            var row = t.rows[i];
            var x = row.insertCell(cellslength);
        }
    }
    else{
        var cellslength = t.rows[0].cells.length;
        for(var i = 0; i < t.rows.length; i++){
            t.rows[i].insertCell(val);
        }
    }
    settablecellclick();
    sumcolumns();
    setTimeout(rowcolumnscount, 1000);
    loadcolumnstoselect();
}

function deletecolumnbuttonclick(){
    var t = document.getElementById("sheettable");
    var val = document.getElementById("optionspanelfrontdeletecolumnsnumber").value;
    
    var cellslength = t.rows[0].cells.length;
    for(var i = 0; i < t.rows.length; i++){
        t.rows[i].deleteCell(val);
    }
    settablecellclick();
    sumcolumns();
    setTimeout(rowcolumnscount, 2000);
    loadcolumnstoselect();
}

var sumsarray = new Array();
function addsumcolumnsbuttonclick(){
    var t = document.getElementById("sheettable");
    var optnum = document.getElementById("optionspanelfrontcolumnsnumber").value;
    var sum = 0;
    var check = false;
    var sumcheck = true;
    for(var i=0; i<sumsarray.length; i++){
        if(sumsarray[i] == optnum){
            check = true;
            break;
        }
    }
    if(!check){
        for(var i=0; i<t.rows.length; i++){
            if(t.rows[i].cells[optnum].innerHTML != 0 && t.rows[i].cells[optnum].innerHTML != ""){
                if(parseFloat(t.rows[i].cells[optnum].innerHTML)){
                    sum += parseFloat(t.rows[i].cells[optnum].innerHTML);
                    sumcheck = true;
                }
                else{
                    alert("This columns has other stuff than numbers..");
                    sumcheck = false;
                    break;
                }
            }
        }
        if(sumcheck){
            sumsarray[sumsarray.length] = optnum;
            document.getElementById("statuspaneltablerow").innerHTML += "<td><label class='commalabels'>, </label></td>"
                                                                    +"<td><label class='titlelabels'>Sum of Column "+optnum+": </label></td>"
                                                                    +"<td><label id='"+optnum+"sumvaluelabel' class='valuelabels'>"+sum+"</label></td>";
        }
    }
}

function sumcolumns(){
    var t = document.getElementById("sheettable");
    var sum = 0;
    for (var j = 0; j < sumsarray.length; j++) {    
        for (var i = 0; i < t.rows.length; i++) {
            if(t.rows[i].cells[sumsarray[j]].innerHTML != 0 && t.rows[i].cells[sumsarray[j]].innerHTML != "" && parseFloat(t.rows[i].cells[sumsarray[j]].innerHTML)){
                sum += parseFloat(t.rows[i].cells[sumsarray[j]].innerHTML);
            sumcheck = true;
            }
        }
        var x = sumsarray[j]+"sumvaluelabel";
        document.getElementById(x).innerHTML = sum;
    }
}

function settingsbuttonclick(){
    const modalPath = path.join('file://', __dirname, '../html/settings.html')
    let win = new BrowserWindow({ frame: false, transparent: true, width: 500, height: 600})
    win.on('close', function(){ win=null })
    win.loadURL(modalPath)
    win.show()
}




setInterval(() => {
    document.getElementById("processinglabel").innerHTML += "."
    if(document.getElementById("processinglabel").innerHTML == "Processing......"){
        document.getElementById("processinglabel").innerHTML = "Processing";
    }
}, 500);