function handler(e) {
    e = e || window.event;

    var pageX = e.pageX;
    var pageY = e.pageY;

    // IE 8
    if (pageX === undefined) {
        pageX = e.clientX + document.body.scrollLeft + document.documentElement.scrollLeft;
        pageY = e.clientY + document.body.scrollTop + document.documentElement.scrollTop;
    }

    if(pageY <= 30){
        //console.log(pageX, pageY);
        document.getElementById("optionspanel").style.visibility = "visible";
        document.getElementById("optionspanel").style.transition = "0.5s";
        document.getElementById("optionspanel").style.left = "0px";
        //document.getElementById("leftsidepanel").style.height = "calc(100vh - 50px)";
        //document.getElementById("leftsidepanel").style.marginTop = "0px";
    }
    else if(pageY >= 70){
        document.getElementById("optionspanel").style.transition = "0.5s";
        document.getElementById("optionspanel").style.left = "-2000px";
        document.getElementById("optionspanel").style.visibility = "hidden";
        //document.getElementById("leftsidepanel").style.height = "calc(100vh - 0px)";
        //document.getElementById("leftsidepanel").style.marginTop = "-50px";
    }
}

// attach handler to the click event of the document
if (document.attachEvent) document.attachEvent('onmousemove', handler);
else document.addEventListener('mousemove', handler);