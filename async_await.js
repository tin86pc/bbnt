let congPr = (a, b) => {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject = 'Số hạng thứ nhất không phải là số';
            }
            if (typeof b != 'number') {
                return reject = 'Số hạng thứ hai không phải là số';
            }
            resolve(a + b)

        }, 2000);
    });
};

let nhanPr = (a, b) => {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject = 'Số hạng thứ nhất không phải là số';
            }
            if (typeof b != 'number') {
                return reject = 'Số hạng thứ hai không phải là số';
            }
            resolve(a * b)

        }, 2000);
    });
};

let chiaPr = (a, b) => {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject = 'Số hạng thứ nhất không phải là số';
            }
            if (typeof b != 'number') {
                return reject = 'Số hạng thứ hai không phải là số';
            }
            if (typeof b == 0) {
                return reject = 'Lối chia cho 0';
            }

            resolve(a / b)

        }, 2000);
    });
};


let tinhDienTichHinhThang = async (a, b, h) => {
    let res = await congPr(a, b);
    let nhan = await nhanPr(res, h);
    let kq = await chiaPr(nhan, 2);
    console.log('diện tích ' + kq);
}

tinhDienTichHinhThang(4, 5, 6);
