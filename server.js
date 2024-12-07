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
    const logData = []; // To hold the data we want to log or save

    for (let i = 1; i < data.length; i++) {
      const row = data[i]; // Access each row in the Excel sheet
      const searchQuery = row[4]; // Column E, index 4

      // Capture the data from columns F (5), I (8), K (10), and L (11)
      const columnF = row[5] || "-";
      const columnG = row[6] || "-";
      const columnH = row[7] || "-";
      const columnI = row[11] || "-";
      const columnJ = row[11] || "-";

      // Save this data to logData
      logData.push([columnF, columnG, columnH, columnI, columnJ]);

      // Check if searchQuery starts with "Công an"
      if (searchQuery && searchQuery.startsWith("Công an")) {
        console.log(">>>>>>>>", searchQuery);

        const link = await searchGoogle(searchQuery); // Search Google for Facebook link
        const details = await getFanpageDetails(link);
        row[5] = link; // Column F for Facebook link
        row[6] = details.phone; // Column G for phone
        row[7] = details.phone; // Column H for phone
        row[8] = details.email; // Column I for email
        row[9] = details.address; // Column J for address

        results.push({ searchQuery, link, details });
      }
    }

    // Tạo một worksheet mới với dữ liệu đã cập nhật
    const updatedWorksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, updatedWorksheet, "Updated Data");

    // Lưu tệp Excel mới
    const updatedFilePath = path.join(
      __dirname,
      "uploads",
      `updated_data_${Date.now()}.xlsx`
    );
    xlsx.writeFile(workbook, updatedFilePath);

    fs.unlinkSync(filePath); // Xóa tệp Excel ban đầu sau khi xử lý
    res.json({ results, updatedFilePath }); // Trả về đường dẫn tệp đã cập nhật
  } catch (error) {
    console.error(error);
    res.status(500).send("An error occurred");
  }
});

async function searchGoogle(query) {
  const apiKey = "AIzaSyB-zjI4n-sXmad_ZQ76juPrzeX1WQq7xbg";
  const cseId = "341005c8435be49e1";

  async function performSearch(q) {
    const url = `https://www.googleapis.com/customsearch/v1/?q=${encodeURIComponent(
      q
    )}&cx=${cseId}&key=${apiKey}&excludeTerms=story.php&as_sitesearch=facebook.com`;
    console.log(url);
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

      if (key === "phone" && typeof value === "string") {
        value = value.replace(/\s+/g, "");
      }

      data[key] = value;
    });

    return data;
  });

  await browser.close();
  return details;
}

app.listen(3000, () =>
  console.log("Server is running on http://localhost:3000")
);
