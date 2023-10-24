const Toast = (nd) => {
    const x = document.getElementById("snackbar");
    x.textContent = nd;
    x.className = "show";

    setTimeout(() => {
        x.className = x.className.replace("show", "");
    }, 3000);
}


const delay = time => new Promise(res => {
    setTimeout(res, time)
});

function hienThi(eid) {
    document.getElementById(eid).style.display = "block";
}

function hienThiAn(eid,) {
    document.getElementById(eid).style.display = "none";
}

function log(s) {
    console.log(s);
    var textarea = document.getElementById('log');
    textarea.value += '> ' + s + '\n';
    textarea.scrollTop = textarea.scrollHeight;
}