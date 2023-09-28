function doGet(e) {
    let Ngay = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    let Gio = Utilities.formatDate(new Date(), "GMT+7", "HH:mm")

    let o = {
        e: e,
        email: e.parameter.email,
        mxt: Math.floor((Math.random() * 10000) + 1),
        ngay: Ngay,
        gio: Gio
    }

    const response = [o];

    return ContentService
        .createTextOutput(JSON.stringify(response))
        .setMimeType(ContentService.MimeType.JSON);
}