// // thêm cdn
// var cdn = document.createElement('script');
// cdn.setAttribute('src', 'https://cdn.sheetjs.com/xlsx-0.19.3/package/dist/xlsx.full.min.js');
// document.head.appendChild(cdn);

//const { docx } = require("./lib/docx_v7");

// cdn.setAttribute('src', 'https://cdnjs.cloudflare.com/ajax/libs/docxtemplater/3.37.9/docxtemplater.min.js');
// document.head.appendChild(cdn);


const dm = 'danh muc';// Tên sheet
const tt = "thong tin";// Tên sheet
const tenFile = 'Tên file';// Tên file đầu ra
const tenFileMau = 'Tên file mẫu'// Tên file mẫu
let Json = {};// obj chứa dữ liệu của file excell


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
    textarea.value += s + '\n>';
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

    XuLyToaBienBan()


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


async function XuLyToaBienBan() {
    let files = document.getElementById('file').files;
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
    Json = await excelFileToJsonPromise(FileExcell);
    //console.log(Json);

    // Render các file đã xử lý
    const output = document.getElementById('ketQua');
    let html = "";

    for (let i = 0; i < Json[dm].length; i++) {
        html += `<div>`
        html += `<li> <a onclick="Xulyfile(${i})"> ${Json[dm][i][tenFile]} </a> </li>`
        html += `</div>`
    }
    output.innerHTML = '<ol>' + html + '</ol>';
    output.scrollIntoView();

}

function Xulyfile(vitri) {
    const o = Json[dm][vitri];
    let tenFM = o[tenFileMau];
    log(`Đang tìm file ${tenFM}`)

    let files = document.getElementById('file').files;
    // lấy ra file word trong các file đã chọn
    let FileWord;

    for (let i = 0; i < files.length; i++) {
        // Tìm file mẫu
        if (files[i].name.toUpperCase() === tenFM.toUpperCase()) {
            FileWord = files[i];
        }
    }

    // Thông báo lỗi không tìm thấy file mẫu
    if (FileWord == undefined) {
        log(`Không tìm thấy file mẫu ${tenFM}`)
        return;
    } else {
        log(`Tìm thấy file ${tenFM}`);
    }

    //console.log(FileWord);

    onMyfileChange(FileWord, vitri);


    log("Đã tải file");

}



function onMyfileChange(fileInput, vitri) {
    const oDanhMuc = Json[dm][vitri];
    const tenF = oDanhMuc[tenFile];
    const arrThongTin = Json[tt];

    if (fileInput == undefined) {
        return;
    }

    var zip = new JSZip()
    zip.loadAsync(fileInput)
        .then(function (word) {
            let xml = word.file("word/document.xml").async("string");
            return xml
        })
        .then(function (xml) {

            for (const [key, value] of Object.entries(oDanhMuc)) {
                xml = replaceXml(xml, key, value);
            }

            arrThongTin.forEach(element => {
                xml = replaceXml(xml, Object.keys(element), Object.values(element));
            });

            zip.file("word/document.xml", xml);

            zip.generateAsync({ type: "blob" }).then(function (blob) {
                saveAs(blob, tenF);
            });

        })
}


function replaceXml(xml, cu, moi) {

    moi = moi.toString();
    moi = moi
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("'", "&apos;")
        .replaceAll("\"", "&quot;");

    cu = cu.toString();
    cu = cu
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("'", "&apos;")
        .replaceAll("\"", "&quot;");

    let aMoi = moi.split(/[\r\n]+/);

    if (aMoi.length > 1) {


        // let s=+aMoi[0]

        // for (let m = 1; m < aMoi.length; m++) {
        //    s+='</w:p></w:r></w:t>'+aMoi[m]+'</w:t></w:r></w:p>'

        // }

        // xml.replaceAll(cu, moi);









        //xml.replaceAll(cu, moi);

        // let out = '';
        // let aXml = xml.split(cu);

        // out += aXml[0].substring(0, aXml[0].indexOf('<w:p '));


        // for (let x = 0; x < aXml.length - 1; x++) {
        //     let dau = aXml[x].substring(aXml[x].indexOf('<w:p '), aXml[x].length);

        //     for (let m = 0; m < aMoi.length; m++) {

        //         out += dau + aMoi[m] + '</w:t></w:r></w:p>'
        //     }

        // }

        // let cuoi = aXml[aXml.length - 1];

        // out += cuoi.substring(cuoi.indexOf('<w:p '), cuoi.length);

        // console.log(out)

        // return out;

    }
    else {
        return xml.replaceAll(cu, moi);
    }

}
