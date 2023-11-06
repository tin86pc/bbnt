let appScriptJson = {};
const LayJson = async () => {
    let appScript = "https://script.google.com/macros/s/AKfycbytA4rEc6taja3HzF1QUgUmrtnwkOAAu5mbOCKZmCPzyYtsTuIaeZ9nkOv8WA0TIuI2/exec"
    fetch(appScript)
        .then(d => d.json())
        .then(s => {
            s = s.trim();
            s = s.replace(/(\r\n|\n|\r)/gm, "");



            appScriptJson = JSON.parse(s);
            // console.log('bbnt xin chào !');
        })
        .catch(() => {
            console.log('lỗi file Json');
        })

}
LayJson();