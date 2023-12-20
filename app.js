const fs = require("fs");
const ExcelJS = require("exceljs");
const cheerio = require("cheerio");

const SPREADSHEET_PATH = "./addresses.xlsx";
const SPREADSHEET_BACKUP_PATH = "./addresses_old.xlsx";

const PAGE_NUMBER = 1;
const TARGET_URL = `https://www.baza-firm.com.pl/przemys%C5%82owe-maszyny-i-urz%C4%85dzenia/strona-${PAGE_NUMBER}`;
const ANCHOR_ELEMENT_SELECTOR =
  "a.pikto_txt.displayInlineBlock.piktoBt.allCornerRound3";

const CONTACT_PATH = "/kontakt";

async function getPageHTML(url) {
  try {
    const response = await fetch(url);
    const html = await response.text();

    return html;
  } catch {
    return null;
  }
}

function getPageURLs(html) {
  const $ = cheerio.load(html);
  const $a = $(ANCHOR_ELEMENT_SELECTOR);

  let links = Array.from($a).map((element) => {
    return element.attribs.href;
  });
  links = links.filter((url) => url.length > 0);

  return links;
}

function getEmailAddressesFromString(html) {
  try {
    const emailRegex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g;
    const email = html.match(emailRegex);

    return email;
  } catch {
    return null;
  }
}

async function writeAddressesToFile(addressArray) {
  const workbook = new ExcelJS.Workbook();

  // If file already exists, backup and read it
  // Else create new worksheet
  if (fs.existsSync(SPREADSHEET_PATH)) {
    fs.copyFile(SPREADSHEET_PATH, SPREADSHEET_BACKUP_PATH, (err) => {});
    await workbook.xlsx.readFile(SPREADSHEET_PATH);
  } else {
    workbook.addWorksheet("Addresses");
  }

  const worksheet = workbook.getWorksheet("Addresses");

  let currentRowIndex = 0;
  let currentAddressIndex = 0;

  while (true) {
    currentRowIndex += 1;

    const cell = worksheet.getRow(currentRowIndex).getCell(1);

    // Skip if cell already has value
    if (cell.value !== null) {
      continue;
    }

    cell.value = addressArray[currentAddressIndex];
    currentAddressIndex += 1;

    console.log(addressArray[currentAddressIndex - 1]);
    console.log(`${currentAddressIndex} / ${addressArray.length}`);

    // Break loop if all addresses had been written
    if (addressArray.length === currentAddressIndex) break;
  }

  workbook.xlsx.writeFile(SPREADSHEET_PATH);
}

async function main() {
  console.log("GETTING INITIAL TARGET HTML...");
  const html = await getPageHTML(TARGET_URL);

  console.log("GETTING URLS...");
  // Get URLs from target URL
  const pageURLsRaw = getPageURLs(html);

  console.log("ADDING ${CONTACT_PATH} TO URLS...");
  // Add URLs with contact path
  const pageURLs = pageURLsRaw
    .map((url) => {
      let cleanURL = url;
      if (url[url.length - 1] === "/") {
        cleanURL = url.slice(0, -1);
      }

      return [cleanURL, cleanURL + CONTACT_PATH];
    })
    .flat(1);

  console.log("GETTING HTMLS...");
  // Get HTML for each obtained URL
  const pageHTMLs = await Promise.all(pageURLs.map((url) => getPageHTML(url)));

  console.log("GETTING EMAIL ADDRESSES...");
  // Get email addresses from HTMLs
  const emailAddressesRaw = await Promise.all(
    pageHTMLs.map((html) => getEmailAddressesFromString(html))
  );

  console.log("FILTERING EMAIL ADDRESSES...");
  // Flatten, filter remove duplicates and lowercase
  const emailAddresses = [...new Set(emailAddressesRaw.flat(1))]
    .filter((address) => address !== null)
    .map((address) => address.toLowerCase());

  writeAddressesToFile(emailAddresses);
}

main();
