import 'regenerator-runtime/runtime';
// @ts-ignore
import * as Excel from "exceljs";
import {saveAs} from "file-saver";
import moment from "moment";

declare let chrome:any;

export async function convertToExcel() {

    let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    chrome.runtime.onMessage.addListener(function(request:any, sender:any) {
        if (request.action == "getSource") {
            console.log(request.source);
            parseHTML(request.source);
        }
    });

    let tabId = [tab][0].id;

    chrome.scripting.executeScript({
        target: {tabId: tabId},
        files: ["./dist/getPagesSource.js"]
    });

}


async function parseHTML(sourceHTML:string) {

    let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    try {

        var xmlString = sourceHTML;
        var xmlDoc = new DOMParser().parseFromString(xmlString, "text/html");

        var elements = xmlDoc.getElementsByClassName("formitem form viewform")

        let propertyNodes = xmlDoc.querySelectorAll(".formitem.form.viewform");

        let wb = new Excel.Workbook();

        let col = 15;
        let row = 30;

        let wsData = [];
        for (let i = 0; i < row; i++) {
            let r = [];
            for (let j = 0; j < col; j++) {
                r[j] = "";
            }
            wsData[i] = r;
        }

        let ws = wb.addWorksheet("Sold Records");

        let attributes = ["MLS#", "Address", "Property Type", "Bedroom", "Washroom", "Kitchens", "Garage", "Lot", "Size", "Taxes", "Listing Price", "Offer Bidding", "Sold Price", "Difference", "Sold Date", "DOM"];

        let rowID = 1;

        ws.getColumn(1).width = 14;

        for (let i in attributes) {
            // let cellRef = XLSX.utils.encode_cell({c:1, r:Number(i)});
            ws.getCell("A" + rowID).value = attributes[i];
            rowID++;
        }

        if (!propertyNodes || propertyNodes.length <= 0) {
            throw "Error";
            return false;
        }



        let i = 1;
        for (let propertyNode of propertyNodes) {

            let colID = alphabet[i];

            let overviewNode = propertyNode.querySelector(".formitem.legacyBorder.formgroup.vertical > .formitem.formgroup.tabular:first-child > .formitem.formgroup.horizontal > .formitem.formgroup.vertical:nth-child(2)");
            let propertyTypeNodes = overviewNode?.querySelectorAll(".formitem.formgroup:nth-child(6) > .formitem.formgroup.horizontal:first-child > .formitem.formgroup.vertical:first-child > .formitem.formfield")
            let addrNode = overviewNode?.querySelector(".formitem.formgroup.horizontal > .formitem.formgroup.horizontal > .formitem.formgroup.vertical > .formitem.formgroup.horizontal > .formitem.formgroup.vertical > .formitem.formgroup.horizontal > .formitem.formfield:first-child > .value") as HTMLElement;
            let labels = propertyNode.querySelectorAll("label");

            let map: Record<string, any> = {};

            for (let label of labels) {
                let key = label.innerText.replace(":", "").trim();

                if (!map[key]) {
                    let lastChild = label?.parentNode?.lastChild as HTMLElement;
                    map[key] = lastChild?.innerText.trim();
                }

            }

            if (propertyTypeNodes && propertyTypeNodes.length > 0) {
                let node1FirstChild = propertyTypeNodes[0]?.firstElementChild as HTMLElement;
                let node2FirstChild = propertyTypeNodes[2].firstElementChild as HTMLElement;

                let driveParkSpcs = map["Drive Pk Spcs"] ? parseInt(map["Drive Pk Spcs"]) : (map["Drive Park Spcs"] ? parseInt(map["Drive Park Spcs"]) : 0);

                let mls = map["MLS#"];
                let addr = addrNode?.innerText;
                let propertyType = node1FirstChild?.innerText + " (" + node2FirstChild?.innerText + ")";
                let bedroom = map["Bedrooms"];
                let washroom = map["Washrms"];
                let kitchens = map["Kitchens"];
                let garage = parseInt(map["Tot Pk Spcs"]) - driveParkSpcs + " + " + driveParkSpcs;
                let log = map["Lot"];
                let size = map["Apx Sqft"];
                let taxes = map["Taxes"].replaceAll(",", "").replaceAll("$", "");
                let listingPrice = map["List"].replaceAll(",", "").replaceAll("$", "");
                let offerBidding = "";
                let soldPrice = map["Sold"].replaceAll(",", "").replaceAll("$", "");
                let difference: number = ((soldPrice - listingPrice) / listingPrice);
                let soldDate = map["Sold Date"];
                let DOM = map["DOM"];

                let rowID = 1;
                ws.getCell(colID + rowID++).value = mls;
                ws.getCell(colID + rowID++).value = addr;
                ws.getCell(colID + rowID++).value = propertyType;
                ws.getCell(colID + rowID++).value = bedroom;
                ws.getCell(colID + rowID++).value = washroom;
                ws.getCell(colID + rowID++).value = kitchens;
                ws.getCell(colID + rowID++).value = garage;
                ws.getCell(colID + rowID++).value = log;
                ws.getCell(colID + rowID++).value = size;
                ws.getCell(colID + rowID).value = parseFloat(taxes);
                ws.getCell(colID + rowID++).numFmt = "$#,##0.00";
                ws.getCell(colID + rowID).value = parseFloat(listingPrice);
                ws.getCell(colID + rowID++).numFmt = "$#,##0.00";
                ws.getCell(colID + rowID++).value = offerBidding != "" ? parseFloat(offerBidding) : "";
                ws.getCell(colID + rowID).value = parseFloat(soldPrice);
                ws.getCell(colID + rowID++).numFmt = "$#,##0.00";
                ws.getCell(colID + rowID).value = { formula: `(${colID}${rowID-1}-${colID}${rowID-3})/${colID}${rowID-3}`, result: difference };
                ws.getCell(colID + rowID++).numFmt = "0.00%";
                ws.getCell(colID + rowID++).value = moment(soldDate).format("DD-MMM-YYYY").toString();
                ws.getCell(colID + rowID++).value = parseInt(DOM);

                ws.getColumn(colID).eachCell(function(cell:any, rowNumber:any) {
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                });

                ws.getColumn(colID).width = 22;

                i++;
            }

        }
        /* Add the worksheet to the workbook */
        const buffer = await wb.xlsx.writeBuffer();
        const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const fileExtension = '.xlsx';

        const blob = new Blob([buffer], {type: fileType});

        saveAs(blob, 'sold-record-report.xlsx');
        return
    }
    catch (e) {
        console.log(e);
        alert("Cannot convert this page");
    }
}