// promies return promies

const cong = function (a, b) {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject("Số hạng thứ nhất không phải là số");
            }
            if (typeof b != 'number') {
                return reject("Số hạng thứ hai không phải là số");
            }
            resolve(a + b)

        }, 2000)
    })
}

const nhan = function (a, b) {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject("Số hạng thứ nhất không phải là số");
            }
            if (typeof b != 'number') {
                return reject("Số hạng thứ hai không phải là số");
            }
            resolve(a * b)

        }, 2000)
    })
}

const chia = function (a, b) {
    return new Promise((resolve, reject) => {
        setTimeout(() => {
            if (typeof a != 'number') {
                return reject("Số hạng thứ nhất không phải là số");
            }
            if (typeof b != 'number') {
                return reject("Số hạng thứ hai không phải là số");
            }
            if (typeof b == 0) {
                return reject("Lỗi chia cho 0");
            }

            resolve(a / b)

        }, 2000)
    })
}


let tinhDienTichHinhThang = (a, b, h) => {
    cong(a, b)
        .then(res => nhan(res, h))
        .then(result => chia(result, 2))
        .then(dientich => console.log("diện tích: " + dientich))
        .catch(err => console.log(err));
}

tinhDienTichHinhThang(1, 2, 3);
