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
      const searchQuery = row[4]; // Column E, index 4

      if (searchQuery && searchQuery.startsWith("Công an")) {
        console.log("Processing Công an:", searchQuery);

        const link = await searchGoogle(searchQuery); // Search Facebook link
        const details = link
          ? await getFanpageDetails(link)
          : { phone: "-", email: "-", address: "-" };

        // Update data vào các cột tương ứng
        row[5] = link || "-"; // Column F: Facebook link
        row[6] = details.phone || "-"; // Column G: Phone
        row[7] = details.phone || "-"; // Column H: Phone
        row[8] = details.email || "-"; // Column I: Email
        row[9] = details.address || "-"; // Column J: Address

        results.push({ searchQuery, link, details });
      } else if (searchQuery && searchQuery.includes("UBND")) {
        console.log("Processing UBND:", searchQuery);

        const link = await searchGoogleWithGov(searchQuery); // Tìm GOV link
        const details = link
          ? await getFooterDetails(link) // Trích xuất từ footer
          : { phone: "-", email: "-", address: "-" };

        // Update data vào các cột tương ứng
        row[5] = link || "-"; // Column F: GOV link
        row[6] = details.phone || "-"; // Column G: CỐ ĐỊNH
        row[7] = details.phone || "-"; // Column H: DI ĐỘNG
        row[8] = details.email || "-"; // Column I: EMAIL
        row[9] = details.address || "-"; // Column J: ĐỊA CHỈ

        results.push({ searchQuery, link, details });
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

async function getFooterDetails(url) {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();

  try {
    // Thiết lập User-Agent để tránh bị chặn
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    );

    // Tăng thời gian chờ và tải trang
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

    // Trích xuất thông tin từ thẻ footer
    const details = await page.evaluate(() => {
      const footer = document.getElementById("footer");
      if (!footer) {
        return { phone: "-", email: "-", address: "-" }; // Nếu không tìm thấy footer
      }

      const textContent = footer.textContent || ""; // Lấy toàn bộ nội dung văn bản
      const result = {
        address: (textContent.match(/Địa chỉ: (.+?)(\n|$)/) || [])[1] || "-", // Lọc địa chỉ
        phone: (textContent.match(/Điện thoại: (.+?)(\n|$)/) || [])[1] || "-", // Lọc điện thoại
        email:
          (textContent.match(
            /\b[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9.-]+\.)?hanoi\.gov\.vn\b/
          ) || [])[1] || "-", // Lọc email
      };

      return result;
    });

    return details;
  } catch (error) {
    console.error("Error scraping footer details:", error);
    return { phone: "-", email: "-", address: "-" };
  } finally {
    await browser.close();
  }
}

app.listen(3000, () =>
  console.log("Server is running on http://localhost:3000")
);
