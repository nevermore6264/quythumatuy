const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");
const fs = require("fs");
const path = require("path");
const axios = require("axios");

const app = express();
const upload = multer({ dest: "uploads/" });
const cors = require("cors");
app.use(cors());
app.use(express.static(path.join(__dirname, "public")));

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    const filePath = req.file.path; // File path of uploaded Excel
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const results = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i]; // Access each row in the Excel sheet
      const searchQuery = row[1]; // Column E, index 4

      if (searchQuery && searchQuery.startsWith("Công an")) {
        console.log(`Row ${i}: Processing Công an - ${searchQuery}`);

        const link = await searchGoogle(searchQuery); // Search Facebook link
        const details =
          link && link !== "-"
            ? await getFanpageDetails(link)
            : { phone: "-", email: "-", address: "-" };

        row[2] = link || "-"; // Column F: Facebook link

        const phoneDetails = normalizePhoneNumber(details.phone);
        row[3] = phoneDetails.type === "Mobile" ? phoneDetails.phone : "-"; // Column D: DI ĐỘNG
        row[4] = phoneDetails.type === "Landline" ? phoneDetails.phone : "-"; // Column E: CỐ ĐỊNH

        row[5] =
          details.email && details.email !== "-"
            ? {
                t: "s",
                v: searchQuery,
                f: `HYPERLINK("${link}", "${searchQuery}")`,
              }
            : "-"; // Column F: Email

        row[6] = details.address || "-"; // Column G: Address

        row[1] =
          link && link !== "-"
            ? {
                t: "s",
                v: searchQuery,
                f: `HYPERLINK("${link}", "${searchQuery}")`,
              }
            : searchQuery;

        results.push({ searchQuery, link, details });
      } else if (searchQuery && searchQuery.includes("UBND")) {
        console.log(`Row ${i}: Processing UBND - ${searchQuery}`);

        const link = await searchGoogleWithGov(searchQuery); // Tìm GOV link

        row[2] = link || "-"; // Column C: GOV link
        row[3] = "-"; // Column D: DI ĐỘNG
        row[4] = "-"; // Column E: CỐ ĐỊNH
        row[5] = "-"; // Column F: EMAIL
        row[6] = "-"; // Column G: ĐỊA CHỈ

        // Update Column E with the link as a clickable hyperlink in Excel
        row[1] =
          link && link !== "-"
            ? {
                t: "s",
                v: searchQuery,
                f: `HYPERLINK("${link}", "${searchQuery}")`,
              }
            : searchQuery;

        results.push({ searchQuery, link });
      }
    }

    // Create updated worksheet with modified data
    const updatedWorksheet = xlsx.utils.aoa_to_sheet(data);
    const updatedWorkbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(
      updatedWorkbook,
      updatedWorksheet,
      "Updated Data"
    );

    // Save the updated Excel file
    const updatedFilePath = path.join(
      __dirname,
      "uploads",
      `updated_data_${Date.now()}.xlsm`
    );
    xlsx.writeFile(updatedWorkbook, updatedFilePath);

    fs.unlinkSync(filePath); // Delete original Excel file after processing
    res.json({ results, updatedFilePath }); // Return the updated file path
  } catch (error) {
    console.error(error);
    res.status(500).send("An error occurred");
  }
});

// Hàm tìm kiếm Facebook link
async function searchGoogle(query) {
  // const apiKey = "AIzaSyB-zjI4n-sXmad_ZQ76juPrzeX1WQq7xbg";
  // const cseId = "341005c8435be49e1";

  const apiKey = "AIzaSyCUwoIfESwtDcjb2kDDGASJNqJCLk-5LvM";
  const cseId = "67c1b1438f7244c19";

  async function performSearch(q) {
    const url = `https://www.googleapis.com/customsearch/v1/?q=${encodeURIComponent(
      q
    )}&cx=${cseId}&key=${apiKey}&excludeTerms=story.php&as_sitesearch=facebook.com`;
    try {
      const response = await axios.get(url);
      const searchResults = response.data.items;

      const facebookLinks = searchResults
        .map((item) => item.link)
        .filter(
          (link) =>
            link.startsWith("https://www.facebook.com") &&
            !link.includes("posts") &&
            !link.includes(".php") &&
            !link.includes("photos") &&
            !link.includes("photo") &&
            !link.includes("rell")
        );

      return facebookLinks[0] || null;
    } catch (error) {
      console.error("Error performing search:", error);
      return null;
    }
  }

  let facebookLink = await performSearch(query);
  if (!facebookLink) {
    facebookLink = await performSearch(`Tuổi Trẻ ${query}`);
  }

  return facebookLink || "-";
}

// Hàm tìm kiếm GOV link
async function searchGoogleWithGov(query) {
  // const apiKey = "AIzaSyB-zjI4n-sXmad_ZQ76juPrzeX1WQq7xbg";
  // const cseId = "341005c8435be49e1";

  const apiKey = "AIzaSyCUwoIfESwtDcjb2kDDGASJNqJCLk-5LvM";
  const cseId = "67c1b1438f7244c19";

  async function performSearch(q) {
    const url = `https://www.googleapis.com/customsearch/v1/?q=${encodeURIComponent(
      q
    )}&cx=${cseId}&key=${apiKey}&siteSearch=gov.vn`;
    try {
      const response = await axios.get(url);
      const searchResults = response.data.items;

      const govLinks = searchResults
        .map((item) => item.link)
        .filter((link) => link.includes(".gov.vn"));

      return govLinks[0] || null;
    } catch (error) {
      console.error("Error performing search:", error);
      return null;
    }
  }

  return await performSearch(query);
}

// Hàm thu thập thông tin fanpage
async function getFanpageDetails(url) {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto(url, { waitUntil: "networkidle2" });

  const details = await page.evaluate(() => {
    const data = {};
    const icons = {
      phone: "https://static.xx.fbcdn.net/rsrc.php/v4/yT/r/Dc7-7AgwkwS.png",
      email: "https://static.xx.fbcdn.net/rsrc.php/v4/yE/r/2PIcyqpptfD.png",
      address: "https://static.xx.fbcdn.net/rsrc.php/v4/yW/r/8k_Y-oVxbuU.png",
    };

    Object.keys(icons).forEach((key) => {
      const element = document.querySelector(`img[src*="${icons[key]}"]`);
      const parent = element?.closest("div + div");
      let value = parent?.textContent?.trim() || "-";

      data[key] = value;
    });

    return data;
  });

  await browser.close();
  return details;
}

function normalizePhoneNumber(rawPhone) {
  if (!rawPhone) return null; // Không có số trả về null

  // 1. Loại bỏ các ký tự không cần thiết
  let phone = rawPhone.replace(/[\s()-]/g, ""); // Xóa khoảng trắng, dấu ngoặc, gạch ngang

  // 2. Chuyển đổi đầu số quốc tế (+84 hoặc 84) thành 0
  phone = phone.replace(/^(\+84|84)/, "0");

  // 3. Xử lý số điện thoại của các nhà mạng với đầu số cũ
  const carrierMap = {
    // MOBIFONE
    "0120": "070",
    "0121": "079",
    "0122": "077",
    "0126": "076",
    "0128": "078",
    // GMOBILE
    "0199": "059",
    // VIETNAMOBILE
    "0186": "056",
    "0188": "058",
    // VINAPHONE
    "0123": "083",
    "0124": "084",
    "0125": "085",
    "0127": "081",
    "0129": "082",
    // VIETTEL
    "0169": "039",
    "0168": "038",
    "0167": "037",
    "0166": "036",
    "0165": "035",
    "0164": "034",
    "0163": "033",
    "0162": "032",

    // Mã vùng cũ và mã vùng mới Miền Bắc
    "0240": "0204",
    "0281": "0209",
    "0241": "0222",
    "026": "0206",
    "0230": "0215",
    "0219": "0219",
    "0351": "0226",
    "04": "024",
    "039": "0239",
    "0320": "0220",
    "031": "0225",
    "0218": "0218",
    "0321": "0221",
    "0231": "0213",
    "025": "0205",
    "020": "0214",
    "0350": "0228",
    "030": "0229",
    "038": "0238",
    "0210": "0210",
    "052": "0232",
    "033": "0203",
    "053": "0233",
    "022": "0212",
    "027": "0207",
    "036": "0227",
    "0280": "0208",
    "037": "0237",
    "054": "0234",
    "0211": "0211",
    "029": "0216",

    // Mã vùng mới Miền Nam
    "076": "0296",
    "064": "0254",
    "0781": "0291",
    "075": "0275",
    "0650": "0274",
    "056": "0256",
    "0651": "0271",
    "062": "0252",
    "0780": "0290",
    "0710": "0292",
    "0511": "0236",
    "0500": "0262",
    "0501": "0261",
    "061": "0251",
    "067": "0277",
    "059": "0269",
    "0711": "0293",
    "08": "028",
    "077": "0297",
    "060": "0260",
    "058": "0258",
    "063": "0263",
    "072": "0272",
    "068": "0259",
    "057": "0257",
    "0510": "0235",
    "055": "0255",
    "079": "0299",
    "066": "0276",
    "073": "0273",
    "074": "0294",
    "070": "0270",
  };

  Object.keys(carrierMap).forEach((oldPrefix) => {
    const regex = new RegExp(`^0${oldPrefix}`);
    if (regex.test(phone)) {
      phone = phone.replace(regex, `0${carrierMap[oldPrefix]}`);
    }
  });

  const mobilePrefixes = [
    // Đầu số 09
    "090",
    "091",
    "092",
    "093",
    "094",
    "095",
    "096",
    "097",
    "098",
    "099",

    // Đầu số 07
    "070",
    "079",
    "077",
    "076",
    "078",

    // Đầu số 08
    "083",
    "084",
    "085",
    "081",
    "082",

    // Đầu số 03
    "039",
    "038",
    "037",
    "036",
    "035",
    "034",
    "033",
    "032",

    // Đầu số 05
    "056",
    "058",
    "059",
  ];
  // Kiểm tra nếu là số di động
  const isMobile = mobilePrefixes.some((prefix) => phone.startsWith(prefix));

  // Nếu là số di động
  if (isMobile) {
    return { phone: phone, type: "Mobile" };
  } else {
    return { phone: phone, type: "Landline" };
  }
}

app.listen(3000, () =>
  console.log("Server is running on http://localhost:3000")
);
