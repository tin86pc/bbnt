let aMoi = moi.split(/[\r\n]+/);
aMoi.forEach(element => {
    ht(element);
    ht("----");
});




if (aMoi.length > 1) {

    //----------------

    //----------------

    let out = '';
    let aXml = xml.split(cu);

    if (aXml.length > 1) {
        out += aXml[0].substring(0, aXml[0].indexOf('<w:p '));


        for (let x = 0; x < aXml.length - 1; x++) {
            let dau = aXml[x].substring(aXml[x].indexOf('<w:p '), aXml[x].length);

            for (let m = 0; m < aMoi.length; m++) {

                out += dau + aMoi[m] + '</w:t></w:r></w:p>'
            }

        }

        let cuoi = aXml[aXml.length - 1];

        out += cuoi.substring(cuoi.indexOf('<w:p '), cuoi.length);

        return out;
    }

}
