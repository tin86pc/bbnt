const obj = {}

const tenFile = 'Tên file';// Tên file đầu ra
const tenFileMau = 'Tên file mẫu'// Tên file mẫu

const NhanMaXacThuc = async () => {
    const email = document.getElementById('email').value

    if (email == "") {
        Toast(`Địa chỉ email đang để trống`);
        return;
    }

    var aa = email.indexOf("@");
    var dd = email.lastIndexOf(".");
    if (aa < 1 || dd < (aa + 2) || (dd + 2) >= email.length) {
        Toast(`Email không xác định`);
        return;
    }

    obj.email = email;
    const nd = `?email=${email}&mxt=0`
    let appScript = "https://script.google.com/macros/s/AKfycbyCrVKAdv3NgFkXg_dClsSebS36husGSs_HmOSURhK8A9ctfApmHBYq94OWmk6iC2WPNQ/exec"

    appScript = appScript + nd;

    fetch(appScript)
        .then(d => d.json())
        .then(d => {
            const Email = d[0].email;
            obj.mxt = d[0].mxt;
            Toast(`Mã xác thực đã gửi thành công đến email ${Email}`);
            document.getElementById('maxacthuc').focus();
        })
        .catch(() => {
            Toast("Lỗi kết nối");
        })

}

const XacThuc = async () => {

    const maxacthuc = document.getElementById('maxacthuc')
    if (maxacthuc.value == obj.mxt) {
        obj.on = true;
    }

    if (obj.on == true) {
        await delay(1000)
        maxacthuc.value = ""
        await delay(1000)
        Toast(`Đã xác thực thành công`)
        await delay(4000)
        hienThiAn('xacthuc');
        hienThi('taifilelen');

    }
    else {
        if (obj.email == undefined) {
            Toast(`Bạn chưa nhận mã xác thực`)
        } else {
            Toast(`Mã xác thực không đúng cần truy cập vào email ${obj.email} để lấy`)
        }
    }

}

const TaiFileLen = async () => {
    await delay(1000);
    Toast("Đang tải file lên");
    await delay(4000);
    hienThiAn('taifilelen');
    hienThi('taifilexuong');
}

let Json;
const tt = "thong tin";
const dm = 'danh muc';

const TaiFileXuong = async () => {
    await delay(1000);
    Json = await LayDuLieuTuFileExcell();

    const DanhMuc = Json[dm];
    const length = DanhMuc.length;

    var zipCha = new JSZip();
    var ketxuat = zipCha.folder("ket xuat");
    for (vitri = 0; vitri < length; vitri++) {
        try {
            data = await tvw(vitri);
            ketxuat.file(data.tenF, data.blob);
            log(data.tenF + "... ok")
        } catch (e) {
            log("TaiFileXuong " + e);
        }
    }

    var textarea = document.getElementById('log');
    zipCha.file("Ghi chu.txt", textarea.value);

    zipCha.generateAsync({ type: "blob" })
        .then(function (content) {
            saveAs(content, "Ket xuat.zip");
        });


    Toast("Đang chuẩn bị file tải xuống");

}

function LayDuLieuTuFileExcell() {

    return new Promise((res, rej) => {
        let files;
        let FileExcell;

        {// Kiểm tra các file tải lên
            files = document.getElementById('file').files;
            if (files.length == 0) {
                log("Chưa chọn được file");
                return;
            }
            else {
                log(`Đang xử lý ${files.length} file đầu vào`)
            }
        }

        {// Lấy ra file excell trong các file đã tải lên
            for (i = 0; i < files.length; i++) {
                const filename = files[i].name;
                const extension = filename.substring(filename.lastIndexOf(".")).toUpperCase();

                // Kiểm tra có phải là file excell không
                if (extension === '.XLS' || extension === '.XLSX') {
                    FileExcell = files[i];
                }
            }

            if (FileExcell === undefined) {
                log("Không tìm thấy file excel trong các file tải lên");
                return;
            } else {
                log(`Đang đọc file excell "${FileExcell.name}" `);
            }
        }

        {// Đọc file excell
            let reader = new FileReader();
            reader.readAsBinaryString(FileExcell);
            reader.onload = function (e) {
                let workbook = XLSX.read(e.target.result, {
                    type: 'binary'
                });


                // Đọc sheet thông tin chung
                let thongTin = {};

                let roaThongTin = XLSX.utils.sheet_to_json(workbook.Sheets[tt], { header: "A" });
                if (roaThongTin.length <= 0) {
                    log(`không thấy sheet ${tt} `);
                    return;
                }

                thongTin[tt] = roaThongTin;
                let arrThongTin = {};
                arrThongTin[tt] = thongTin[tt].map((e) => {
                    let obj = {};
                    obj[e.A] = e.B;
                    return obj;
                });

                // Đọc sheet danh mục
                let danhMuc = {};

                let roaDanhMuc = XLSX.utils.sheet_to_json(workbook.Sheets[dm]);
                if (roaDanhMuc.length <= 0) {
                    log(`không thấy sheet ${dm}`);
                    return;

                }
                danhMuc[dm] = roaDanhMuc;

                // Gom vào 1 đối tượng
                let o = {
                    ...arrThongTin,
                    ...danhMuc
                }
                res(o);

            }

        }

    })

}

function tvw(vitri) {
    return new Promise((resolve, reject) => {
        LayFileMau(vitri) // trả về file mẫu
            .then(LayNoiDungXML)
            .then(SuaLoiXml)
            .then(ThayTheNoiDungDanhMuc)
            .then(ThayTheNoiDungThongTin)
            .then(ThayTheFileXML)
            .then(TaoFileWord)
            .then(function (data) {
                resolve(data)
            })

            .catch((nd) => {
                log(nd)
                reject(nd)
            })

    })

}

function LayFileMau(vitri) {
    const oDanhMuc = Json[dm][vitri];
    let tenFM = oDanhMuc[tenFileMau];

    return new Promise((resolve, reject) => {
        let files = document.getElementById('file').files;

        // lấy ra file word mẫu trong các file đã chọn
        let FileWord;

        for (let i = 0; i < files.length; i++) {
            // Tìm file mẫu
            if (files[i].name === undefined || tenFM === undefined) {
                log("lỗi file excell ")
            } else {
                if (files[i].name.toUpperCase() == tenFM.toUpperCase()) {
                    FileWord = files[i];
                }
            }
        }

        // Thông báo lỗi không tìm thấy file mẫu
        if (FileWord == undefined) {
            reject(`Không tìm thấy file mẫu ${tenFM}`);
        } else {

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

function SuaLoiXml(data) {
    return new Promise((resolve, reject) => {
        const xmlDoc = new DOMParser().parseFromString(data.xml, "text/xml");
        const wt = xmlDoc.getElementsByTagName('w:t');

        // Gộp các nội dung bị tách
        if (wt.length > 0) {
            for (iwt = 0; iwt < wt.length; iwt++) {

                if (wt[iwt].textContent.includes('<%') && !wt[iwt].textContent.includes('%>')) {
                    for (i = 1; (iwt + i) < wt.length; i++) {
                        if (wt[iwt + i].textContent.includes('%>')) {
                            wt[iwt].textContent += wt[iwt + i].textContent;
                            wt[iwt + i].textContent = " ";
                            break;
                        }
                        if (!wt[iwt + i].textContent == " ") {
                            wt[iwt].textContent += wt[iwt + i].textContent;
                            wt[iwt + i].textContent = " ";
                        }
                    }
                }
            }
        }


        // xử lý < và % bị tách
        if (wt.length > 0) {
            for (iwt = 0; iwt < wt.length - 1; iwt++) {
                if (wt[iwt].textContent[wt[iwt].textContent.length - 1] == "<" && wt[iwt + 1].textContent[0] == "%") {
                    // console.log(wt[iwt]);
                    wt[iwt].textContent += wt[iwt + 1].textContent;
                    wt[iwt + 1].textContent = "";
                }
            }
        }

        // xử lý % và > bị tách
        if (wt.length > 0) {
            for (iwt = 0; iwt < wt.length - 1; iwt++) {
                if (wt[iwt].textContent[wt[iwt].textContent.length - 1] == "%" && wt[iwt + 1].textContent[0] == ">") {
                    // console.log(wt[iwt]);
                    wt[iwt].textContent += wt[iwt + 1].textContent;
                    wt[iwt + 1].textContent = "";
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


document.addEventListener('keydown', (event) => {
    var name = event.key;
    if (name == "q" || name == "Q") {
        hienThiAn('xacthuc');
        hienThi('taifilelen');
    }

}, false);


function GopY() {
    hienThi('gopy');
}

function DongGopY() {
    hienThiAn('gopy');
}
