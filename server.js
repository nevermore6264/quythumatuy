const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer");
const fs = require("fs");

const app = express();
const upload = multer({ dest: "uploads/" });
const cors = require("cors");
app.use(cors());
const path = require("path");
const axios = require("axios");

app.use(express.static(path.join(__dirname, "public")));

app.post("/process", upload.single("file"), async (req, res) => {
  try {
    const filePath = req.file.path; // Đường dẫn file Excel được tải lên
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    const results = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i]; // Lấy từng dòng
      const searchQuery = row[4]; // Cột E tương ứng với index 4 (bắt đầu từ 0)

      if (searchQuery && searchQuery.startsWith("Công an")) {
        console.log(">>>>>>>>", searchQuery);

        const link = await searchGoogle(searchQuery); // Tìm kiếm link fanpage

        // const details = link ? await getFanpageDetails(link) : {}; // Lấy thông tin từ fanpage nếu có
        results.push({ searchQuery, link });
      }
    }

    fs.unlinkSync(filePath); // Xóa file sau khi xử lý
    res.json(results); // Trả kết quả về client
  } catch (error) {
    console.error(error);
    res.status(500).send("Đã xảy ra lỗi");
  }
});

async function searchGoogle(query) {
  const apiKey = "AIzaSyBQzVfEynWjEksz59iqGkCNGxg03pmAJQ0";
  const cseId = "341005c8435be49e1";

  const url = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(
    query
  )}&cx=${cseId}&key=${apiKey}`;

  try {
    const response = await axios.get(url);
    const searchResults = response.data.items;

    // Lọc kết quả để lấy Facebook link
    const facebookLink = searchResults
      .map((item) => item.link)
      .find(
        (link) =>
          link.startsWith("https://www.facebook.com") && !link.includes("posts")
      );

    return facebookLink || null;
  } catch (error) {
    console.error("Error searching Google:", error);
    return null;
  }
}

// Lấy thông tin từ fanpage bằng Puppeteer
async function getFanpageDetails(url) {
  //   const browser = await puppeteer.launch();
  //   const page = await browser.newPage();
  //   await page.goto(url, { waitUntil: "load" });

  //   const details = await page.evaluate(() => {
  //     const phone = document.querySelector("[data-testid='phone']")?.textContent;
  //     const email = document.querySelector("[data-testid='email']")?.textContent;
  //     const address = document.querySelector(
  //       "[data-testid='address']"
  //     )?.textContent;
  //     return { phone, email, address };
  //   });

  //   await browser.close();
  //   return details;
  return null;
}

app.listen(3000, () =>
  console.log("Server is running on http://localhost:3000")
);
