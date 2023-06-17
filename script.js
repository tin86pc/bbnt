// // thêm cdn
var cdn = document.createElement('script');
cdn.setAttribute('src', 'https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js');
document.head.appendChild(cdn);

cdn.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.37.9/docxtemplater.min.js');
document.head.appendChild(cdn);






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
function excelFileToJsonPromise(file) {
    return new Promise((resolve, reject) => {
        try {
            let reader = new FileReader();
            reader.readAsBinaryString(file);
            reader.onload = function (e) {
                let workbook = XLSX.read(e.target.result, {
                    type: 'binary'
                });

                const tt = "thong tin";// Tên sheet
                let roaThongTin = XLSX.utils.sheet_to_json(workbook.Sheets[tt], { header: "A" });
                if (roaThongTin.length <= 0) {
                    alert(`không thấy sheet "${tt}" `);
                    return;
                }
                let thongTin = {};
                thongTin[tt] = roaThongTin;

                //xử lý thông tin mảng nhận được
                let arrThongTin = {};
                arrThongTin[tt] = thongTin[tt].map((e) => {
                    let obj = {};
                    obj[e.A] = e.B;
                    return obj;
                });


                const dm = 'danh muc';// Tên sheet
                let roaDanhMuc = XLSX.utils.sheet_to_json(workbook.Sheets[dm]);
                if (roaDanhMuc.length <= 0) {
                    alert(`không thấy sheet "${dm}"`);
                    return;
                }
                let danhMuc = {};
                danhMuc[dm] = roaDanhMuc;

                let output = {
                    ...arrThongTin,
                    ...danhMuc
                }

                //console.log(output);
                resolve(output);

            }
        }
        catch (e) {
            console.error(e);
            reject(e);
        }


    });
}


async function UploadCacFile() {
    var files = document.getElementById('file').files;
    if (files.length == 0) {
        alert("Chưa chọn được file");
        return;
    } else {
        log(`Đang xử lý ${files.length} file`)
    }

    // lấy ra file excell trong các file đã chọn
    let FileExcell;

    for (i = 0; i < files.length; i++) {
        const filename = files[i].name;
        const extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();

        // Kiểm tra có phải là file excell không
        if (extension === '.XLS' || extension === '.XLSX') {
            FileExcell = files[i];

        }
    }

    if (FileExcell === undefined) {
        alert("Không tìm thấy file excel trong các file bạn đã chọn");
        return;
    } else {
        log(`Đang đọc file excell "${FileExcell.name}" `);
    }

    // lấy dữ liệu từ file excell
    let json = await excelFileToJsonPromise(FileExcell);
    console.log(json);


    // Thay thế dữ liệu nhận được vào các file word



}


function test() {
    log('test');
}


// http://www.ludovicperrichon.com/update-docx-with-js-and-optianally-upload-it-to-sp/



