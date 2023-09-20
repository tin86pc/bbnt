const obj = {}

const NhanMaXacThuc = () => {
    const email = document.getElementById('email')
    obj.email = email.value;
    Toast(`Mã xác thực đã được gửi đến: ${email.value}`)
}

const XacThuc = async () => {
    obj.on = true;
    if (obj.on) {
        await delay(1000);
        Toast(`Đã xác thực thành công`)
        await delay(3000);

        const DivLogin = document.getElementById('login')
        DivLogin.setAttribute("hidden", "hidden");
    }

}