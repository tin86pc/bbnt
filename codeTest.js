const s = " 22<%1  2          3%> a  <1  2           3> b   <1  2  3> c <1  2  3> d"
xlc(s, "<%", "%>")

function ThayTheNoiDungDanhMuc2(data) {
    return new Promise((resolve, reject) => {
        const xmlDoc = new DOMParser().parseFromString(data.xml, "text/xml");
        const wp = xmlDoc.getElementsByTagName('w:p');

        for (const [key, value] of Object.entries(Json[dm][data.vitri])) {
            let cu = key.toString();
            let moi = value.toString();

            if (cu.includes('<%') && cu.includes('%>')) {
                if (wp.length > 0) {
                    for (iwp = 0; iwp < wp.length; iwp++) {

                        if (wp[iwp].textContent.includes('<%') && wp[iwp].textContent.includes('%>')) {
                            let aMoi = moi.split(/\r\n|\n|\r/);
                            console.log(aMoi.length);
                            if (aMoi.length > 1) {
                                for (iaMoi = 0; iaMoi < iaMoi.length; iaMoi++) {

                                    console.log(aMoi);


                                    let newWp = wp[iwp];



                                    //wp[iwp].insertAfterElement("afterend", newWp);
                                }
                            }
                            else {
                                wp[iwp].textContent = wp[iwp].textContent.replaceAll(cu, moi);
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

function ThayTheNoiDungThongTin2(data) {
    return new Promise((resolve, reject) => {
        const xmlDoc = new DOMParser().parseFromString(data.xml, "text/xml");
        const wp = xmlDoc.getElementsByTagName('w:p');

        const arrThongTin = Json[tt];
        arrThongTin.forEach(element => {
            let cu = Object.keys(element).toString();
            let moi = Object.values(element).toString();


            if (cu.includes('<%') && cu.includes('%>')) {
                if (wp.length > 0) {
                    for (iwp = 0; iwp < wp.length; iwp++) {

                        if (wp[iwp].textContent.includes('<%') && wp[iwp].textContent.includes('%>')) {
                            let aMoi = moi.split(/\r\n|\n|\r/);

                            if (aMoi.length > 1) {
                                for (iaMoi = 0; iaMoi < iaMoi.length; iaMoi++) {

                                    console.log(aMoi);


                                    let newWp = wp[iwp];
                                    console.log(newWp);



                                    wp[iwp].insertAfterElement("afterend", newWp);
                                }
                            }
                            else {
                                wp[iwp].textContent = wp[iwp].textContent.replaceAll(cu, moi);
                            }

                        }
                    }
                }
            }



        });



        // chuyển XMLDocument thành string
        const sXml = new XMLSerializer().serializeToString(xmlDoc);
        data.xml = sXml;
        resolve(data);
    })
}
