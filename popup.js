// Initialize butotn with users's prefered color
let toExcelButton = document.getElementById("toExcelButton");

toExcelButton.addEventListener("click", async () => {

  let [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  chrome.runtime.onMessage.addListener(function(request, sender) {
    if (request.action == "getSource") {
      console.log(request.source);
      parseHTML(request.source);
    }
  });

  let tabId = [tab][0].id;

  chrome.scripting.executeScript({
    target: {tabId: tabId},
    files: ["getPagesSource.js"]
  });

});

function parseHTML(sourceHTML) {

  try {

    var xmlString = sourceHTML;
    var xmlDoc = new DOMParser().parseFromString(xmlString, "text/html");

    var elements = xmlDoc.getElementsByClassName("formitem form viewform")

    let propertyNodes = xmlDoc.querySelectorAll(".formitem.form.viewform");

    let excel = [];

    let attributes = ["MLS#", "Address", "Property Type", "Bedroom", "Washroom", "Kitchens", "Garage", "Lot", "Size", "Taxes", "Listing Price", "Offer Bidding", "Sold Price", "Difference", "Sold Date", "DOM"];

    for (let i in attributes) {
      excel[i] = [];
      excel[i].push(attributes[i]);
    }

    if (!propertyNodes || propertyNodes.length <= 0) {
      throw "Error";
      return false;
    }

    for (let propertyNode of propertyNodes) {
      let overviewNode = propertyNode.querySelector(".formitem.legacyBorder.formgroup.vertical > .formitem.formgroup.tabular:first-child > .formitem.formgroup.horizontal > .formitem.formgroup.vertical:nth-child(2)");
      let propertyTypeNodes = overviewNode.querySelectorAll(".formitem.formgroup:nth-child(6) > .formitem.formgroup.horizontal:first-child > .formitem.formgroup.vertical:first-child > .formitem.formfield")
      let addrNode = overviewNode.querySelector(".formitem.formgroup.horizontal > .formitem.formgroup.horizontal > .formitem.formgroup.vertical > .formitem.formgroup.horizontal > .formitem.formgroup.vertical > .formitem.formgroup.horizontal > .formitem.formfield:first-child > .value");
      let labels = propertyNode.querySelectorAll("label");

      let map = {};

      for (let label of labels) {
        let key = label.innerText.replace(":", "").trim();

        if (!map[key]) {
          map[key] = label.parentNode.lastChild.innerText.trim();
        }

      }

      let mls = map["MLS#"];
      let addr = addrNode.innerText;
      let propertyType = propertyTypeNodes[0].firstElementChild.innerText + " (" + propertyTypeNodes[2].firstElementChild.innerText + ")";
      let bedroom = map["Bedrooms"];
      let washroom = map["Washrms"];
      let kitchens = map["Kitchens"];
      let garage = parseInt(map["Tot Pk Spcs"]) - parseInt(map["Drive Pk Spcs"]) + " + " + parseInt(map["Drive Pk Spcs"]);
      let log = map["Lot"];
      let size = map["Apx Sqft"];
      let taxes = map["Taxes"].replaceAll(",", "").replaceAll("$", "");
      let listingPrice = map["List"].replaceAll(",", "").replaceAll("$", "");
      let offerBidding = "";
      let soldPrice = map["Sold"].replaceAll(",", "").replaceAll("$", "");
      let difference = (soldPrice - listingPrice) / listingPrice * 100;
      let soldDate = map["Sold Date"];
      let DOM = map["DOM"];

      let colIdx = 0;
      excel[colIdx++].push(mls);
      excel[colIdx++].push(addr);
      excel[colIdx++].push(propertyType);
      excel[colIdx++].push(bedroom);
      excel[colIdx++].push(washroom);
      excel[colIdx++].push(kitchens);
      excel[colIdx++].push(garage);
      excel[colIdx++].push(log);
      excel[colIdx++].push(size);
      excel[colIdx++].push("$" + parseFloat(taxes).toFixed(2));
      excel[colIdx++].push("$" + listingPrice);
      excel[colIdx++].push(offerBidding);
      excel[colIdx++].push("$" + soldPrice);
      excel[colIdx++].push(parseFloat(difference).toFixed(2) + "%");
      excel[colIdx++].push(soldDate);
      excel[colIdx++].push(DOM);

    }

    let csvContent = "";

    for (let rowID in excel) {
      csvContent += excel[rowID].join(",") + "\n"
    }

    var blob = new Blob([csvContent], {type: "text/csv"});
    var url = URL.createObjectURL(blob);
    chrome.downloads.download({
      url: url,
      filename: "torontomls.csv"
    });

  }
  catch (e) {
    alert("Cannot convert this page");
  }
}