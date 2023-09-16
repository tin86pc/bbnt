const phienban = "BBNT v2"

const dm = 'danh muc';// Tên sheet
const tt = "thong tin";// Tên sheet
const tenFile = 'Tên file';// Tên file đầu ra
const tenFileMau = 'Tên file mẫu'// Tên file mẫu
let Json = {};// obj chứa dữ liệu của file excell
const khoaMo = "%>";
const khoaDong = "<%";

const urlLayEmailThoiGian = "https://script.google.com/macros/s/AKfycbzmFLTgG7e-3co9QOVxr0SQgu0vn5ganUfJhuoeqCQTi03oKq_Qc6TPFKDX53luM_Eo_A/exec";

const urlSheet = 'https://script.google.com/macros/s/AKfycbyxhkj9iPqzwpBnX5ZEjJiDWfAw3XEcEWEPEO2ujK0TSSuSbC1eR7EQWtbA5q0PDNuWVQ/exec';

(function KhoiTaoBanDau() {
    document.title = phienban;
})();

// lấy dữ liệu từ google app script
(function LayEmailThoiGian() {
    fetch(urlLayEmailThoiGian)
        .then(d => d.json())
        .then(d => {
            const Ngay = d[0].ngay;
            const Gio = d[0].gio;
            const Email = d[0].email;
            document.getElementById("log").value = 'Hà Nội, ngày ' + Ngay + " " + Gio + '\n';
            document.getElementById("email").value = Email;

            objUser = d[0];
        })
}
)();

// Gửi dữ liệu về google sheet
let objUser = {};
setTimeout(() => {
    fetch(urlSheet, { method: "POST", body: objUser });
}, 5000);

// Gửi dữ liệu lên google sheet
function GuiPhanHoi() {

    let form = document.querySelector("form");
    let thongbao = document.getElementById("thongbao")
    thongbao.innerHTML = "Đang gửi phản hồi";
    let data = new FormData(form);

    fetch(urlSheet, { method: "POST", body: data })
        .then(res => res.text())
        .then(text => {
            document.getElementById("thongbao").innerHTML = text;
        });
}

// Hiển thị log lên textarea
function log(s) {
    var textarea = document.getElementById('log');
    textarea.value += s + '\n>';
    textarea.scrollTop = textarea.scrollHeight;
}

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

    document.getElementById("xuly").style.display = "block";
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


function LayDuLieuTuFileExcell() {
    ChayDelay()
        .then(KiemTraFileDaChon)
        .then(LayRaFileExcel)
        .then(DocFileExcell)
        .then(HienThiFileKetQua)

        .catch((nd) => {
            console.log(`lỗi: ${nd}`);
        })
}

function ChayDelay() {
    return new Promise((resolve, reject) => {
        let width = 1;
        //let id = setInterval(frame, 50);
        let id = setInterval(frame, 5);

        function frame() {
            if (width >= 100) {
                // kết thúc
                clearInterval(id);
                resolve();

            } else {
                // Hiển thị
                width++;
                const myBar = document.getElementById("myBar");
                myBar.style.width = width + "%";
                myBar.innerHTML = width + "%";
            }
        }
    })
}

function KiemTraFileDaChon() {
    return new Promise((resolve, reject) => {
        let files = document.getElementById('file').files;
        if (files.length == 0) {
            log("Chưa chọn được file");
            reject(`
            Kiểm tra file đã chọn
            Chưa chọn được file`);
        }
        else {
            log(`Đang xử lý ${files.length} file`)
            resolve({ 'files': files });
        }
    })
}

function LayRaFileExcel(data) {
    return new Promise((resolve, reject) => {

        let files = data.files;
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
            log("Không tìm thấy file excel trong các file bạn đã chọn");

            reject(`Lấy ra file Excell
            Không tìm thấy file excel trong các file bạn đã chọn`);
        } else {
            log(`Đang đọc file excell "${FileExcell.name}" `);
            data.FileExcell = FileExcell;
            resolve(data);
        }
    })
}

function DocFileExcell(data) {
    let file = data.FileExcell;
    return new Promise((resolve, reject) => {

        let reader = new FileReader();
        reader.readAsBinaryString(file);
        reader.onload = function (e) {
            let workbook = XLSX.read(e.target.result, {
                type: 'binary'
            });

            let roaThongTin = XLSX.utils.sheet_to_json(workbook.Sheets[tt], { header: "A" });
            if (roaThongTin.length <= 0) {
                alert(`không thấy sheet ${tt} `);
                reject(`không thấy sheet ${tt} `);

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
                alert(`không thấy sheet ${dm}`);
                reject(`không thấy sheet ${dm}`);
            }
            let danhMuc = {};
            danhMuc[dm] = roaDanhMuc;

            let output = {
                ...arrThongTin,
                ...danhMuc
            }

            data.json = output;

            resolve(data);
        }

    });
}

function HienThiFileKetQua(data) {
    return new Promise((resolve, reject) => {
        // lấy dữ liệu từ file excell
        Json = data.json

        // Render các file đã xử lý
        const output = document.getElementById('ketQua');
        let html = "";

        for (let i = 0; i < Json[dm].length; i++) {
            html += `<div>`
            html += `<li> <a onclick="XulyfileClichChon(${i})"> ${Json[dm][i][tenFile]} </a> </li>`
            html += `</div>`
        }
        output.innerHTML = '<ol>' + html + '</ol>';
        //output.scrollIntoView();

        document.getElementById("daura").style.display = "block";

    })
}

function LayFileMau(vitri) {
    const oDanhMuc = Json[dm][vitri];
    let tenFM = oDanhMuc[tenFileMau];
    log(`Đang tìm file ${tenFM}`)

    return new Promise((resolve, reject) => {
        let files = document.getElementById('file').files;

        // lấy ra file word mẫu trong các file đã chọn
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
            reject(`Không tìm thấy file mẫu ${tenFM}`);
        } else {
            log(`Tìm thấy file ${tenFM}`);

            let zip = new JSZip();
            let word = zip.loadAsync(FileWord);
            let tenF = Json[dm][vitri][tenFile];

            const data = {
                vitri: vitri,
                word: word,
                tenF: tenF
            }

            resolve(data);
        }
    })
}

function LayNoiDungXML(data) {
    return new Promise((resolve, reject) => {

        data.word.then(function (w) {
            w.file("word/document.xml")
                .async("string")
                .then(function (xml) {
                    data.xml = xml;
                    resolve(data);
                })
        })

    });

}

function newComic() {
    var elem = document.querySelector("#genreList").selectedOptions;
    console.log(elem);
    for (var i = 0; i < elem.length; i++) {
        console.log(elem[i].attributes[0].nodeValue); //second console output
    }
}


function SoSanh_rRr(a, b) {
    //console.log('SoSanh_rRr');

    let kq = true;

    for (i = 0; i < a[0].childNodes.length; i++) {

        const xa = new XMLSerializer().serializeToString(a[0].childNodes[i]);
        //console.log(xa);

        const xb = new XMLSerializer().serializeToString(b[0].childNodes[i]);
        //console.log(xb);

        if (xa !== xb) {
            kq = false
            return;
        }

    }
    //console.log("true");

    return kq;
}

function SuaLoiXml(data) {
    return new Promise((resolve, reject) => {
        const xmlDoc = new DOMParser().parseFromString(data.xml, "text/xml");
        const wp = xmlDoc.getElementsByTagName('w:p');

        if (wp.length > 0) {
            for (iwp = 0; iwp < wp.length; iwp++) {
                const wr = wp[iwp].getElementsByTagName('w:r');

                if (wr.length > 1) {

                    for (iwr = 1; iwr < wr.length; iwr++) {

                        const wrPr0 = wr[iwr - 1].getElementsByTagName("w:rPr")
                        const wt0 = wr[iwr - 1].getElementsByTagName("w:t")

                        const wrPr = wr[iwr].getElementsByTagName("w:rPr")
                        const wt = wr[iwr].getElementsByTagName("w:t")
                        //console.log(iwr);

                        if (wt[0] != undefined && wt0[0] != undefined) {

                            if (SoSanh_rRr(wrPr0, wrPr) == true) {

                                wt0[0].childNodes[0].nodeValue = wt0[0].childNodes[0].nodeValue + wt[0].childNodes[0].nodeValue;
                                wt[0].childNodes[0].nodeValue = "";

                            }

                        }

                    }

                }
            }
        }









        // chuyển XMLDocument thành string
        const sXml = new XMLSerializer().serializeToString(xmlDoc);
        data.xml = sXml;
        resolve(data);

    })
}

function XuLyXMLBiTach(data) {
    return new Promise((resolve, reject) => {
        //Chuyển string thành XMLDocument

        let xmlDoc = new DOMParser().parseFromString(data.xml, "text/xml");

        // Xử lý xml
        let awt = xmlDoc.getElementsByTagName('w:t');

        for (let i = 0; i < awt.length; i++) {

            // Sửa lỗi "TypeError: Cannot read properties of undefined (reading 'childNodes')"
            if (awt[i].childNodes.length > 1) {

                for (let j = 0; j < awt[i].childNodes.length; j++) {


                    // Nếu thấy khóa mở và không thấy khóa đóng
                    if (awt[i].childNodes[j].nodeValue.includes("<%") && !awt[i].childNodes[j].nodeValue.includes("%>")) {

                        let k = 0;
                        let ss = "";

                        do {
                            ss += awt[i + k].childNodes[j].nodeValue;
                            awt[i + k].childNodes[j].nodeValue = "";

                            k++;
                        } while (!awt[i + k].childNodes[j].nodeValue.includes("%>") || (awt[i + k].childNodes[j].nodeValue.includes("<%") && awt[i + k].childNodes[j].nodeValue.includes("%>")))// thoát khi false

                        ss += awt[i + k].childNodes[j].nodeValue;
                        awt[i + k].childNodes[j].nodeValue = "";

                        awt[i].childNodes[j].nodeValue = ss;

                    }

                }
            }

        }

        // chuyển XMLDocument thành string
        const sXml = new XMLSerializer().serializeToString(xmlDoc);

        data.xml = sXml;
        resolve(data);
    })
}

function ThayTheNoiDungDanhMuc(data) {
    return new Promise((resolve, reject) => {
        for (const [key, value] of Object.entries(Json[dm][data.vitri])) {
            let cu = key.toString();
            let moi = value.toString();
            data.xml = replaceXml(data.xml, cu, moi);
        }

        resolve(data);
    })
}

function ThayTheNoiDungThongTin(data) {
    return new Promise((resolve, reject) => {
        const arrThongTin = Json[tt];
        arrThongTin.forEach(element => {
            let cu = Object.keys(element).toString();
            let moi = Object.values(element).toString();
            data.xml = replaceXml(data.xml, cu, moi);
        });

        resolve(data);
    })
}

function ThayTheFileXML(data) {
    return new Promise((resolve, reject) => {

        data.word.then(function (word) {
            data.word = word.file("word/document.xml", data.xml);
            resolve(data);
        })

    })
}

function TaoFileWord(data) {

    let DinhDang = {
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    return new Promise((resolve, reject) => {
        data.word.generateAsync(DinhDang)
            .then(function (blob) {
                data.blob = blob;
                resolve(data);
            });
    })
}

function TaiFileWord(data) {
    return new Promise((resolve, reject) => {
        saveAs(data.blob, data.tenF);
        resolve(data);
    })
}

function tvw(vitri) {
    return new Promise((resolve, reject) => {
        LayFileMau(vitri) // trả về file mẫu
            .then(LayNoiDungXML)
            .then(SuaLoiXml)
            //.then(XuLyXMLBiTach)
            .then(ThayTheNoiDungDanhMuc)
            .then(ThayTheNoiDungThongTin)
            .then(ThayTheFileXML)
            .then(TaoFileWord)
            .then(function (data) {
                resolve(data)
            })

            .catch((nd) => {
                console.log(nd);
            })

    })

}

function TaiCacFileMau() {
    let files = document.getElementById('file').files;
    let zip = new JSZip();

    for (let i = 0; i < files.length; i++) {
        zip.file(files[i].name, files[i]);
    }

    zip.generateAsync({ type: "blob" })
        .then((content) => {
            saveAs(content, "File Mau.zip")
        })
}

function XulyfileClichChon(vitri) {
    tvw(vitri)
        .then(TaiFileWord);

}

async function TaiTatCaFile() {
    console.log("Tải toàn bộ file ");
    const DanhMuc = Json[dm];
    const length = DanhMuc.length;


    var zipCha = new JSZip();
    var ketxuat = zipCha.folder("ket xuat");

    for (vitri = 0; vitri < length; vitri++) {
        data = await tvw(vitri);
        ketxuat.file(data.tenF, data.blob);
    }


    var textarea = document.getElementById('log');
    zipCha.file("Ghi chu.txt", textarea.value);


    zipCha.generateAsync({ type: "blob" })
        .then(function (content) {
            saveAs(content, "Ket xuat.zip");
        });

}


