// Tính diện tích hình thang

let cong = (a, b, cb) => {
    setTimeout(() => {
        let loi, kq;
        if (typeof a != 'number') {
            loi = "Số hạng thứ nhất không phải là số";
            return cb(loi, kq);

            //return cb(new Error("Số hạng thứ nhất không phải là số"));
        }
        if (typeof b != 'number') {
            loi = "Số hạng thứ hai không phải là số";
            return cb(loi, kq);
        }
        kq = a + b;
        cb(loi, kq);
    }, 1000);
}

let nhan = (a, b, cb) => {
    setTimeout(() => {
        let loi, kq;
        if (typeof a != 'number') {
            loi = "Số hạng thứ nhất không phải là số";
            return cb(loi, kq);

            //return cb(new Error("Số hạng thứ nhất không phải là số"));
        }
        if (typeof b != 'number') {
            loi = "Số hạng thứ hai không phải là số";
            return cb(loi, kq);
        }
        kq = a * b;
        cb(loi, kq);
    }, 1000);
}

let chia = (a, b, cb) => {
    setTimeout(() => {
        let loi, kq;
        if (typeof a != 'number') {
            loi = "Số chia không phải là số";
            return cb(loi, kq);

            //return cb(new Error("Số hạng thứ nhất không phải là số"));
        }
        if (typeof b != 'number') {
            loi = "Số bị chia không phải là số";
            return cb(loi, kq);
        }
        if (b == 0) {
            loi = "Lỗi chia cho 0";
            return cb(loi, kq);
        }
        kq = a / b;
        cb(loi, kq);
    }, 1000);
}


let tinhDienTichHinhThang = (a, b, h, cb) => {
    cong(a, b, (loi, kq) => {
        if (loi) return console.log(loi);
        nhan(kq, h, (loi, kq) => {
            if (loi) return console.log(loi)
            chia(kq, 2, (loi, kq) => {
                if (loi) return console.log(loi)
                console.log("dien tich :" + kq)

            })
        })
    })
}


tinhDienTichHinhThang(1, 2, 3);
