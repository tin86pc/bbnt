const obj = {}

const NhanMaXacThuc = async () => {
    const email = document.getElementById('email').value

    if (email == "") {
        Toast(`Địa chỉ email đang để trống`);
        return;
    }

    obj.email = email;
    const nd = `?email=${email}&mxt=0`
    let appScript = "https://script.google.com/macros/s/AKfycbxVpUqWh4BYyN8f1cQTF8yZqtVXXknVD8kzlorPrPitGb2qwcnM4UYYc_FBwEuzD6ZHZQ/exec"

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
        document.getElementById('xacthuc').setAttribute("hidden", "hidden")
    }
    else {
        if (obj.email == undefined) {
            Toast(`Bạn chưa nhận mã xác thực`)
        } else {
            Toast(`Mã xác thực không đúng cần truy cập vào email ${obj.email} để lấy`)
        }
    }

}




