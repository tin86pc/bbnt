const Toast = (nd) => {
    const x = document.getElementById("snackbar");
    x.textContent = nd;
    x.className = "show";
    setTimeout(function () { x.className = x.className.replace("show", ""); }, 3000);
}

const delay = time => new Promise(res => setTimeout(res, time));