const SOURCE_ID = "1i5KOk0KmAiYlrZCg3WxjNRvYSbFaj2tCvgeyZgGvd7E"; // spreadsheet file id
const SOURCE_SHEET_NAME = "testing";
const DOCUMENT_ID = "1WSBIyvXk-dyhTvcT7THtSkaisfvvtuyNExRBJ_ICqf0"; // document to write id
const DOCUMENT = loadDocument();
const TITLE = SpreadsheetApp.openById(SOURCE_ID).getName();
function loadData() {
    const SOURCE_FILE = SpreadsheetApp.openById(SOURCE_ID); // open by id
    const SOURCE_SHEET = SOURCE_FILE.getSheetByName(SOURCE_SHEET_NAME); // open current working sheet
    const DATA_RANGE = SOURCE_SHEET.getDataRange(); // get non empy cells
    const DATA = DATA_RANGE.getValues();

    return DATA;
}

function dummyData() {
    const dummyData = [
        [
            [
                "Timestamp",
                "Nama Lengkap",
                "NIP",
                "Bukti screenshot pengisian survei",
            ],
            [
                "Wed Jul 07 02:38:49 GMT-04:00 2021",
                "Faza Ahmad Sulaiman, S.Kom",
                "199410272017121001",
                "https://drive.google.com/open?id=1DXX8gfq0nDaoLBN8O-JHI0KLwtVf3T6p",
            ],
            [
                "Wed Jul 07 02:39:01 GMT-04:00 2021",
                "Hastomo Mawadya Sulistiyandi",
                "198805172017121002",
                "https://drive.google.com/open?id=1hvi6svXZlsDRU1AL0eeaK2KHn0Egxfgw",
            ],
            [
                "Wed Jul 07 02:39:07 GMT-04:00 2021",
                "I Gusti Gede Suputra",
                "198707222017121001",
                "https://drive.google.com/open?id=1oq5yN4LfKYxx7Enol9MybFSkICB-sVxP",
            ],
            [
                "Wed Jul 07 02:39:11 GMT-04:00 2021",
                "Brayan Anggita Linuwih",
                "199301122017121002",
                "https://drive.google.com/open?id=1T0wfIBiz-uEHyhIDYA8fEms7PV_6vqYw",
            ],
        ],
    ];
    return dummyData[0];
}
function loadDocument() {
    return DocumentApp.openById(DOCUMENT_ID);
}

function writeTable(body, data) {
    // data[0] = data[0].map((cell) => (cell === "Timestamp" ? "No." : cell));
    let tableHeader = data[0];
    tableHeader[0] = "No."; // change the table header
    data = data.map((row, index) => {
        if (index == 0) {
            return row;
        }
        row[0] = index.toFixed(0);
        return row;
    });
    // body = appendImage(body);
    // Logger.log(data);
    body.appendTable(data);
    appendImage();
}

function appendImage() {
    let body = DOCUMENT.getBody();
    let table = body.getTables()[0].setBold(false); // first table from collection of tables (if there is any)
    const imgCol = 3;
    const tableLength = table.getNumRows();

    for (let index = 0; index < tableLength; index++) {
        if (index === 0) {
            continue;
        }
        let cell = table.getRow(index).getCell(3);
        let imgText = cell.getText();
        Logger.log(imgText);
        // let fileID = imgText.match(/[\w\_\-]{25,}/).toString();
        let fileID = imgText.match(/[-\w]{25,}/).toString();

        let blob = DriveApp.getFileById(fileID).getBlob();
        // cell.clear().appendImage(blob);

        cell.clear().appendImage(blob).setWidth(135).setHeight(291);
        // cell.clear().appendImage(blob).scaleWidth(0.1).scaleHeight(0.1);
    }
    return body;
}

function writeTitle() {
    let body = DOCUMENT.getBody();
    body.appendParagraph(TITLE.toUpperCase())
        .setBold(true)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function main() {
    // const data = dummyData();
    const data = loadData();
    const docBody = DOCUMENT.getBody();
    writeTitle();
    writeTable(docBody, data);
    // docBody.appendTable([data]);
}
