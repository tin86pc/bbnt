"promise2".log();

function chay(a, b, c) {
    return new Promise((resolve, reject) => {
        `Tính diện tích tam giác`.log();
        `Cạnh a = ${a}`.log();
        `Cạnh b = ${b}`.log();
        `Cạnh c = ${c}`.log();
        resolve({ a: a, b: b, c: c });
    })
}

function kiemTraSo(o) {
    return new Promise((resolve, reject) => {

        if (!(typeof o.a === 'number')) {
            reject("chiều dài cạnh a không phải là số");
        }
        if (!(typeof o.b === 'number')) {
            reject("chiều dài cạnh b không phải là số");
        }
        if (!(typeof o.c === 'number')) {
            reject("chiều dài cạnh c không phải là số");
        }

        resolve(o);
    })
}

function kiemTraAm(o) {
    return new Promise((resolve, reject) => {
        if (o.a == 0) {
            reject("chiều dài cạnh a phải lớn hơn 0");
        }
        if (o.b == 0) {
            reject("chiều dài cạnh b phải lớn hơn 0");
        }
        if (o.c == 0) {
            reject("chiều dài cạnh c phải lớn hơn 0");
        }

        resolve(o);
    })
}

function kiemTraLogic(o) {
    return new Promise((resolve, reject) => {
        if (o.a + o.b < o.c) {
            reject("Chiều dài a + b < c")
        }
        if (o.b + o.c < o.a) {
            reject("Chiều dài b + c < a")
        }
        if (o.a + o.c < o.b) {
            reject("Chiều dài a + c < b")
        }

        resolve(o);
    })
}



function tinhDienTich(o) {
    return new Promise((resolve, reject) => {
        const chuvi = o.a + o.b + o.c;
        const dientich = Math.sqrt(chuvi * (chuvi - o.a) * (chuvi - o.b) * (chuvi - o.c));
        `Diện tích = ${dientich}`.log();
    })
}


chay(1, 1, 1)
    .then(kiemTraSo)
    .then(kiemTraAm)
    .then(kiemTraLogic)
    .then(tinhDienTich)
    .catch((nd) => {
        'lỗi: '.log();
        nd.log();
    })
