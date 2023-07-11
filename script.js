
const dm = 'danh muc';// Tên sheet
const tt = "thong tin";// Tên sheet
const tenFile = 'Tên file';// Tên file đầu ra
const tenFileMau = 'Tên file mẫu'// Tên file mẫu
let Json = {};// obj chứa dữ liệu của file excell
const khoaMo = "%>";
const khoaDong = "<%";



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

            xml = xuLyBiTachXML(xml);

            for (const [key, value] of Object.entries(oDanhMuc)) {
                let cu = key.toString();
                let moi = value.toString();
                xml = replaceXml(xml, cu, moi);
            }

            arrThongTin.forEach(element => {
                let cu = Object.keys(element).toString();
                let moi = Object.values(element).toString();
                xml = replaceXml(xml, cu, moi);
            });



            // Thay thế file
            zip.file("word/document.xml", xml);


            // Tải xuống file
            zip.generateAsync({ type: "blob" }).then(function (blob) {
                saveAs(blob, tenF);
            });

        })
}

function ht(s) {
    console.log(s);
}





function test() {
    //xuLyTrungLap(x);

}

function xuLyBiTachXML(s) {
    const awp = s.split('<%');

    for (let i = 1; i < awp.length; i++) {
        const element = awp[i];
        const vtd = element.indexOf("</w:t>");
        if (vtd != -1) {

            const vtc = element.indexOf("<w:t", vtd);
            const vtcc = element.indexOf(">", vtc);

            let ss = "<%" + element.substring(0, vtd) + element.substring(vtcc + 1);
            awp[i] = ss;
        }

    }


    let t = ""
    for (let i = 0; i < awp.length; i++) {
        t += awp[i] + "<%";
    }

    t = t.substring(0, t.length - 2);
    return t;
}


function replaceXml(xml, cu, moi) {

    moi = moi
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("'", "&apos;")
        .replaceAll("\"", "&quot;");

    cu = cu
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll("'", "&apos;")
        .replaceAll("\"", "&quot;");



    let aMoi = moi.split(/\r\n|\n|\r/);
    let aXml = xml.split(cu);


    if (aMoi.length > 1 && aXml.length > 1) {

        for (let c = 0; c < aXml.length - 1; c++) {

            let s = '';

            for (let m = 0; m < aMoi.length; m++) {

                const sd = aXml[c];
                const dau = sd.substring(sd.lastIndexOf("w:p ") - 1, sd.length);

                const sc = aXml[c + 1];
                const cuoi = sc.substring(0, sc.indexOf("w:p") + 4);

                s += dau + aMoi[m] + cuoi;

            }

            if (s !== '') {

                // Cắt bỏ phần đầu thừa
                s = s.substring(s.indexOf("w:t"), s.length);
                s = s.substring(s.indexOf(">") + 1, s.length);


                // Cắt bỏ phần cuối thừa
                s = s.substring(0, s.lastIndexOf("w:t") - 2);

                aXml[c] = aXml[c] + s;
            }


        }

        let nd = ""
        aXml.forEach(element => {
            nd += element.trim();
        });

        //ht(nd);

        return nd;
    }

    return xml.replaceAll(cu, moi);

}
