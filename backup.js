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

// Hàm console.log kiểu mới
Object.prototype.log = function (x) {
    if (x == undefined) {
        console.log(this.toString());
    }
    else {
        console.log(x + " " + this.toString());
    }
}

Object.prototype.show = function (o) {
    console.dir(o, { depth: null });
}



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

    document.getElementById("xuly").style.display = "block";


}


function XuLyToaBienBan() {


    var width = 1;
    var id = setInterval(frame, 50);
    function frame() {
        if (width >= 100) {
            // kết thúc
            clearInterval(id);
            TaoDanhSachBienBan();

        } else {
            // Hiển thị
            width++;
            const myBar = document.getElementById("myBar");
            myBar.style.width = width + "%";
            myBar.innerHTML = width + "%";
        }
    }

}

async function TaoDanhSachBienBan() {

    let files = document.getElementById('file').files;
    if (files.length == 0) {
        alert("Chưa chọn được file");
        log("Chưa chọn được file")
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
        log("Không tìm thấy file excel trong các file bạn đã chọn");
        return;
    } else {
        log(`Đang đọc file excell "${FileExcell.name}" `);
    }

    // lấy dữ liệu từ file excell
    Json = await excelFileToJsonPromise(FileExcell);

    // Render các file đã xử lý
    const output = document.getElementById('ketQua');
    let html = "";

    for (let i = 0; i < Json[dm].length; i++) {
        html += `<div>`
        html += `<li> <a onclick="Xulyfile(${i})"> ${Json[dm][i][tenFile]} </a> </li>`
        html += `</div>`
    }
    output.innerHTML = '<ol>' + html + '</ol>';
    //output.scrollIntoView();

    document.getElementById("daura").style.display = "block";

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

                resolve(output);
            }
        }
        catch (e) {
            console.error(e);
            reject(e);
        }

    });
}

function Xulyfile(vitri) {
    const oDanhMuc = Json[dm][vitri];
    let tenFM = oDanhMuc[tenFileMau];
    log(`Đang tìm file ${tenFM}`)

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
        alert(`Không tìm thấy file mẫu ${tenFM}`);
        log(`Không tìm thấy file mẫu ${tenFM}`)
        return;
    } else {
        log(`Tìm thấy file ${tenFM}`);
    }

    // Tạo file đầu ra
    const tenF = oDanhMuc[tenFile];
    const arrThongTin = Json[tt];

    var zip = new JSZip()
    zip.loadAsync(FileWord)
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
            zip.generateAsync({ type: "blob" })
                .then(function (blob) {
                    saveAs(blob, tenF);
                });



        })


    log("Đã tải file");

}


function TaiTatCaFile() {
    "Tải toàn bộ file ".log();

    let ghichu = "Tạo biên bản nghiệm thu\n"

    var zipCha = new JSZip();

    // Tạo folder kết xuất
    var ketxuat = zipCha.folder("ket xuat");


    const DanhMuc = Json[dm];
    const length = DanhMuc.length;

    for (vitri = 0; vitri < length; vitri++) {
        const oDanhMuc = Json[dm][vitri];
        let tenFM = oDanhMuc[tenFileMau];

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
            ghichu += `Không tìm thấy file mẫu ${tenFM}`;
        }
        else {

            const tenF = oDanhMuc[tenFile];
            const arrThongTin = Json[tt];

            var zip = new JSZip()
            zip.loadAsync(FileWord)
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

                    ketxuat.file(tenF, "zip");


                })

        }

    }

    zipCha.file("Ghi chu.txt", ghichu);

    zipCha.generateAsync({ type: "blob" })
        .then(function (content) {
            saveAs(content, "output.zip");
        });


}

function xuLyBiTachXML(xml) {
    //Chuyển string thành XMLDocument
    const xmlDoc = new DOMParser().parseFromString(xml, "text/xml");

    // Xử lý xml

    let awt = xmlDoc.getElementsByTagName('w:t');
    for (let i = 0; i < awt.length; i++) {

        for (let j = 0; j < awt[i].childNodes.length; j++) {

            if (awt[i].childNodes[j].nodeValue.includes("<%") && !awt[i].childNodes[j].nodeValue.includes("%>")) {

                let k = 0;
                let ss = "";

                do {
                    ss += awt[i + k].childNodes[j].nodeValue;
                    awt[i + k].childNodes[j].nodeValue = "";

                    k++;
                } while (
                    !awt[i + k].childNodes[j].nodeValue.includes("%>") ||
                    (awt[i + k].childNodes[j].nodeValue.includes("<%") && awt[i + k].childNodes[j].nodeValue.includes("%>"))

                )// thoát khi false

                ss += awt[i + k].childNodes[j].nodeValue;
                awt[i + k].childNodes[j].nodeValue = "";

                awt[i].childNodes[j].nodeValue = ss;
            }

        }
    }

    // chuyển XMLDocument thành string
    const sXml = new XMLSerializer().serializeToString(xmlDoc);
    //sXml.log();

    return sXml;
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


// create a new DOMParser object
var parser = new DOMParser();

// define the XML string to be parsed
var xmlString = "<books><book id='1'><title>JavaScript: The Definitive Guide</title><author>David Flanagan</author></book><book id='2'><title>JavaScript and JQuery: Interactive Front-End Web Development</title><author>Jon Duckett</author></book></books>";

// parse the XML string
var xmlDoc = parser.parseFromString(xmlString, "text/xml");

// create a new book element
var newBook = xmlDoc.createElement("book");

// set the ID attribute of the new book element
newBook.setAttribute("id", "3");

// create a new title element
var newTitle = xmlDoc.createElement("title");

// create a new text node for the title element
var titleText = xmlDoc.createTextNode("JavaScript Patterns");

// append the text node to the title element
newTitle.appendChild(titleText);

// create a new author element
var newAuthor = xmlDoc.createElement("author");

// create a new text node for the author element
var authorText = xmlDoc.createTextNode("Stoyan Stefanov");

// append the text node to the author element
newAuthor.appendChild(authorText);

// append the title and author elements to the new book element
newBook.appendChild(newTitle);
newBook.appendChild(newAuthor);

// append the new book element to the books element
xmlDoc.getElementsByTagName("books")[0].appendChild(newBook);

// serialize the modified XML document back into a string
var newXmlString = new XMLSerializer().serializeToString(xmlDoc);




////////////////////////////////////////////////////////////

// create a new XML document
var xmlDoc = document.implementation.createDocument(null, "books", null);

// create a new book element
var book1 = xmlDoc.createElement("book");

// set the ID attribute of the book element
book1.setAttribute("id", "1");

// create a new title element
var title1 = xmlDoc.createElement("title");

// create a new text node for the title element
var title1Text = xmlDoc.createTextNode("JavaScript: The Definitive Guide");

// append the text node to the title element
title1.appendChild(title1Text);

// create a new author element
var author1 = xmlDoc.createElement("author");

// create a new text node for the author element
var author1Text = xmlDoc.createTextNode("David Flanagan");

// append the text node to the author element
author1.appendChild(author1Text);

// append the title and author elements to the book element
book1.appendChild(title1);
book1.appendChild(author1);

// append the book element to the books element
xmlDoc.getElementsByTagName("books")[0].appendChild(book1);

// serialize the XML document into a string
var xmlString = new XMLSerializer().serializeToString(xmlDoc);


////////////////////////////