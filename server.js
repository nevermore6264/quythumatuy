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
        console.log("Processing Công an:", searchQuery);

        const link = await searchGoogle(searchQuery); // Search Facebook link
        const details = link
          ? await getFanpageDetails(link)
          : { phone: "-", email: "-", address: "-" };

        row[2] = link || "-"; // Column F: Facebook link
        row[3] =
          phoneDetails.type === "Mobile"
            ? normalizePhoneNumber(phoneDetails.phone)
            : "-"; // Column D: DI ĐỘNG
        row[4] =
          phoneDetails.type === "Landline"
            ? normalizePhoneNumber(phoneDetails.phone)
            : "-"; // Column E: CỐ ĐỊNH
        row[5] = details.email || "-"; // Column F: Email
        row[6] = details.address || "-"; // Column G: Address

        // Update Column E with the link as a clickable hyperlink in Excel
        row[1] = link
          ? { f: `HYPERLINK("${link}", "${searchQuery}")` }
          : searchQuery;

        results.push({ searchQuery, link, details });
      } else if (searchQuery && searchQuery.includes("UBND")) {
        console.log("Processing UBND:", searchQuery);

        const link = await searchGoogleWithGov(searchQuery); // Tìm GOV link

        row[2] = link || "-"; // Column C: GOV link
        row[3] = "-"; // Column D: DI ĐỘNG
        row[4] = "-"; // Column E: CỐ ĐỊNH
        row[5] = "-"; // Column F: EMAIL
        row[6] = "-"; // Column G: ĐỊA CHỈ

        // Update Column E with the link as a clickable hyperlink in Excel
        row[1] = link
          ? { f: `HYPERLINK("${link}", "${searchQuery}")` }
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
      `updated_data_${Date.now()}.xlsx`
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
  const apiKey = "AIzaSyB-zjI4n-sXmad_ZQ76juPrzeX1WQq7xbg";
  const cseId = "341005c8435be49e1";

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
  const apiKey = "AIzaSyB-zjI4n-sXmad_ZQ76juPrzeX1WQq7xbg";
  const cseId = "341005c8435be49e1";

  async function performSearch(q) {
    const url = `https://www.googleapis.com/customsearch/v1/?q=${encodeURIComponent(
      q
    )}&cx=${cseId}&key=${apiKey}&siteSearch=gov.vn`;
    console.log(url);
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
  };

  Object.keys(carrierMap).forEach((oldPrefix) => {
    const regex = new RegExp(`^0${oldPrefix}`);
    if (regex.test(phone)) {
      phone = phone.replace(regex, `0${carrierMap[oldPrefix]}`);
    }
  });

  const mobilePrefixes = [
    "070",
    "079",
    "077",
    "076",
    "078",
    "083",
    "084",
    "085",
    "081",
    "082",
    "039",
    "038",
    "037",
    "036",
    "035",
    "034",
    "033",
    "032",
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
