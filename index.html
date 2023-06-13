
// lấy dữ liệu từ google app script
(function () {
    const urlApi = "https://script.google.com/macros/s/AKfycbyZPSsnfkay94_iSMJ0oMdtA3dm4JTici5_avaVf51oC_CvYgISdsE8PvvOb2dpMflp/exec";
    fetch(urlApi)
        .then(d => d.json())
        .then(d => {
            document.getElementById("log").value = 'Hà Nội, ngày ' + d[0].ngay + " " + d[0].gio + '\n';
        });
})();

// Lấy địa chỉ gmail người dùng
(function () {
    const urlApi = 'https://script.google.com/macros/s/AKfycbyLGsRs7t0SJKY74mVcTJZVOMsK6h33mmma-j1BxibVM6ks2wNMlYizjesYGr8fpUgD/exec';
    fetch(urlApi)
        .then(d => d.json())
        .then(d => {
            document.getElementById("myEmail").value = d[0].email;
        });
})();

// Hiển thị log lên textarea
function log(s) {
    var textarea = document.getElementById('log');
    textarea.value += s + '\n';
    textarea.scrollTop = textarea.scrollHeight;
}

// Hiển thị các file đã chọn 
function HienThiCacFile() {
    var trangThai = document.getElementById('trangThai');
    var input = document.getElementById('file');
    var output = document.getElementById('fileList');

    trangThai.textContent = `Đã chọn được ${input.files.length} file bao gồm:`;

    var element = "";

    for (var i = 0; i < input.files.length; ++i) {
        element += `<li> ${input.files.item(i).name}</li>`
    }
    output.innerHTML = '<ol>' + element + '</ol>';
}





//----------------------

function excelFileToJSON(file) {
    let output = {}
    try {
        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function (e) {
            let workbook = XLSX.read(e.target.result, {
                type: 'binary'
            });


            const tt = "thong tin";
            let roaThongTin = XLSX.utils.sheet_to_json(workbook.Sheets[tt], { header: "A" });
            if (roaThongTin.length <= 0) {
                alert(`không thấy sheet "${tt}"`);
                return;
            }
            let thongTin = {};
            thongTin[tt] = roaThongTin;



            const dm = 'danh muc';
            let roaDanhMuc = XLSX.utils.sheet_to_json(workbook.Sheets[dm]);
            if (roaDanhMuc.length <= 0) {
                alert(`không thấy sheet "${dm}"`);
                return;
            }
            let danhMuc = {};
            danhMuc[dm] = roaDanhMuc;

            output = {
                ...thongTin,
                ...danhMuc
            }

        }
    }
    catch (e) {
        console.error(e);
    }

    return output;
}

let dataJson = {}
function UploadCacFile() {
    var files = document.getElementById('file').files;
    if (files.length == 0) {
        alert("Chưa chọn được file");
        return;
    }

    var filename = files[0].name;
    var extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();

    if (extension != '.XLS' && extension != '.XLSX') {
        alert("Không tìm thấy file excel trong các file bạn đã chọn");
        return;
    }

    console.log(excelFileToJSON(files[0]));

    dataJson = excelFileToJSON(files[0]);

    //console.log(dataJson);

    // log("UploadCacFile");
    log(dataJson);

}



// http://www.ludovicperrichon.com/update-docx-with-js-and-optianally-upload-it-to-sp/
