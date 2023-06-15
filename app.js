function test2() {
    setTimeout(() => console.log('1'), 3000);
    console.log('2');
    console.log('3');
}
test2();

function httpGetAsync(url, callback) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.onreadystatechange = function () {
        if (xmlHttp.readyState == 4 && xmlHttp.status == 200)
            callback(xmlHttp);
    };

    xmlHttp.open('GET', url, true);// true for asynchrous
    xmlHttp.send(null);

}

httpGetAsync('https://picsum.photos/200/300', (data) => {
    document.getElementById('img_1').setAttribute('src', data.responseURL);
    httpGetAsync('https://picsum.photos/200/300', (data) => {
        document.getElementById('img_2').setAttribute('src', data.responseURL);
        httpGetAsync('https://picsum.photos/200/300', (data) => {
            document.getElementById('img_3').setAttribute('src', data.responseURL);

        });

    });

});



// Promise ---------------------------------------------------------

const currentPromise = new Promise((resolve, reject) => {
    let condition = true;
    if (condition) {
        setTimeout(() => {
            resolve('success')
        }, 3000);
    } else {
        reject('Error');
    }
});

currentPromise
    .then((data) => {
        console.log(data);
    })
    .catch((err) => {
        console.log(err);
    });



function httpGetAsync(url, resolve, reject) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.onreadystatechange = function () {
        if (xmlHttp.readyState == 4 && xmlHttp.status == 200)
            resolve(xmlHttp);
    };

    xmlHttp.open('GET', url, true);// true for asynchrous
    xmlHttp.send(null);

}

const httpProsemise = new Promise((resolve, reject) => {
    httpGetAsync('https://picsum.photos/200/300', resolve);
})

const httpProsemise2 = new Promise((resolve, reject) => {
    httpGetAsync('https://picsum.photos/200/300', resolve);
})

const httpProsemise3 = new Promise((resolve, reject) => {
    httpGetAsync('https://picsum.photos/200/300', resolve);
})

httpProsemise
    .then((data) => {
        document.getElementById('img_1').setAttribute('src', data.responseURL);
        return httpProsemise2;
    })
    .then((data) => {
        document.getElementById('img_2').setAttribute('src', data.responseURL);
        return httpProsemise3;
    })
    .then((data) => {
        document.getElementById('img_3').setAttribute('src', data.responseURL);
    })
    .catch((err) => {
        console.log(err);
    })



// Async await  ---------------------------------------------------------

const executeAsync = async () => {
    try {
        const response = await httpProsemise;
        document.getElementById('img_1').setAttribute('src', response.responseURL);

        const response2 = await httpProsemise2;
        document.getElementById('img_2').setAttribute('src', response2.responseURL);

        const response3 = await httpProsemise3;
        document.getElementById('img_3').setAttribute('src', response3.responseURL);

    } catch (error) {
        console.log(error)
    }

}

executeAsync();
