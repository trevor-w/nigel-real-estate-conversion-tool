import {convertToExcel} from "./convertToExcel"

window.addEventListener("load", () => {
    let toExcelButton = document.getElementById("toExcelButton");
    toExcelButton?.addEventListener("click", () => {
        convertToExcel();
    });
});