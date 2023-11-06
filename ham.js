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
    var textarea = document.getElementById('log');
    textarea.value += '> ' + s + '\n';
    textarea.scrollTop = textarea.scrollHeight;
}

let objMXT = {
    ngay: "",
    mxt: [8648]
};

// localStorage.clear();



function khoiTaoMXT() {
    if (localStorage.getItem("MXT") == null || !ktNgay(objMXT.ngay)) {
        // Ghi
        objMXT.ngay = new Date().toString();
        objMXT.mxt = [8648];

        localStorage.setItem("MXT", JSON.stringify(objMXT));
        console.log("ghi mới");
    }
    objMXT = JSON.parse(localStorage.getItem("MXT"));
    // console.log("mxt");
    // console.log(objMXT);

}


function themMaXacThuc(m) {
    // nếu trình duyệt hỗ trợ localStorage
    if (typeof (Storage) === 'undefined') { return; }

    khoiTaoMXT();

    // đọc
    objMXT = JSON.parse(localStorage.getItem("MXT"));

    // sửa
    objMXT.mxt.push(m);

    // Ghi
    localStorage.setItem("MXT", JSON.stringify(objMXT));

}


function ktXacThuc(s) {

    // nếu trình duyệt hỗ trợ localStorage
    if (typeof (Storage) === 'undefined') { return; }

    khoiTaoMXT();

    for (i = 0; i < objMXT.mxt.length; i++) {
        if (objMXT.mxt[i] == s) {
            return true;
        }
    }
    return false;
}



function ktNgay(s) {

    var n1 = new Date(s);
    var dd1 = String(n1.getDate()).padStart(2, '0');
    var mm1 = String(n1.getMonth() + 1).padStart(2, '0');
    var yyyy1 = n1.getFullYear();

    var sn1 = mm1 + '/' + dd1 + '/' + yyyy1;

    var n2 = new Date(s);
    var dd2 = String(n2.getDate()).padStart(2, '0');
    var mm2 = String(n2.getMonth() + 1).padStart(2, '0');
    var yyyy2 = n2.getFullYear();

    var sn2 = mm2 + '/' + dd2 + '/' + yyyy2;

    if (sn1 === sn2)
        return true;
    else
        return false;
}






