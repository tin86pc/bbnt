const awp = s.split('<%');

for (let i = 1; i < awp.length; i++) {
    const ewp = awp[i];
    if (ewp.indexOf("</w:t>") != -1) {

        const awt = ewp.split('</w:t>');

        for (let j = 1; j < awt.length - 1; j++) {
            const ewt = awt[j];

            const vtc = ewt.indexOf("<w:t", 0);
            const vtcc = ewt.indexOf(">", vtc);

            let swt = ewt.substring(vtcc + 1);
            //ht(swt);

            awt[j] = swt;
        }
        awt[awt.length - 2] += "</w:t>"

        let k = '';

        for (let n = 0; n < awt.length; n++) {
            k += awt[n];
        }

        awp[i] = k;
    }

}

let t = ""
for (let i = 0; i < awp.length; i++) {
    t += awp[i] + "<%";
}

t = t.substring(0, t.length - 2);
ht("đầu ra " + t);
return t;