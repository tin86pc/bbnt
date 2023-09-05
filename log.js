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


<textarea name="textarea" placeholder="Enter the text..."></textarea>

$(document).ready(function () {
    if ($("textarea").value !== "") {
        alert($("textarea").value);
    }

});


function () {
    var zip = new JSZip();
    for (var i = 0; i < 5; i++) {
        var txt = 'hello'; zip.file("file" + i + ".txt", txt);
    }
    zip.generateAsync({ type: "base64" })
        .then(function (content) { window.location.href = "data:application/zip;base64," + content; });
}


import JSZip from 'jszip';
import FileSaver from 'file-saver';

import validator from './validator';
import generatorRows from './formatters/rows/generatorRows';

import workbookXML from './statics/workbook.xml';
import workbookXMLRels from './statics/workbook.xml.rels';
import rels from './statics/rels';
import contentTypes from './statics/[Content_Types].xml';
import templateSheet from './templates/worksheet.xml';

export const generateXMLWorksheet = (rows) => {
    const XMLRows = generatorRows(rows);
    return templateSheet.replace('{placeholder}', XMLRows);
};

export default (config) => {
    if (!validator(config)) {
        throw new Error('Validation failed.');
    }

    const zip = new JSZip();
    const xl = zip.folder('xl');
    xl.file('workbook.xml', workbookXML);
    xl.file('_rels/workbook.xml.rels', workbookXMLRels);
    zip.file('_rels/.rels', rels);
    zip.file('[Content_Types].xml', contentTypes);

    const worksheet = generateXMLWorksheet(config.sheet.data);
    xl.file('worksheets/sheet1.xml', worksheet);

    return zip.generateAsync({
        type: 'blob',
        mimeType:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }).then((blob) => {
        FileSaver.saveAs(blob, `${config.filename}.xlsx`);
    });
};

function create_zip() {
    var zip = new JSZip();
    zip.add("hello1.txt", "Hello First World\n");
    zip.add("hello2.txt", "Hello Second World\n");
    content = zip.generate();
    location.href = "data:application/zip;base64," + content;
}