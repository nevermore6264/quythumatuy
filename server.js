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

        const details = link ? await getFanpageDetails(link) : null; // Lấy thông tin từ fanpage nếu có
        results.push({ searchQuery, link, details });
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

  // Hàm thực hiện truy vấn Google Custom Search
  async function performSearch(q) {
    const url = `https://www.googleapis.com/customsearch/v1?q=${encodeURIComponent(
      q
    )}&cx=${cseId}&key=${apiKey}&excludeTerms=story.php&as_sitesearch=facebook.com`;

    try {
      const response = await axios.get(url);
      const searchResults = response.data.items;

      // Lọc kết quả để lấy link Facebook hợp lệ
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

      return facebookLinks[0] || null; // Trả về link đầu tiên hợp lệ hoặc null nếu không có
    } catch (error) {
      console.error("Error performing search:", error);
      return null;
    }
  }

  // Tìm kiếm lần đầu với truy vấn ban đầu
  let facebookLink = await performSearch(query);

  // Nếu không có kết quả, thêm "Tuổi Trẻ" vào trước truy vấn và tìm kiếm lại
  if (!facebookLink) {
    facebookLink = await performSearch(`Tuổi Trẻ ${query}`);
  }

  // Trả về kết quả cuối cùng (hoặc null nếu không tìm thấy sau cả hai lần tìm kiếm)
  return facebookLink || "-";
}

// Hàm lấy thông tin từ fanpage
async function getFanpageDetails(url) {
  const browser = await puppeteer.launch({ headless: true });
  const page = await browser.newPage();
  await page.goto(url, { waitUntil: "networkidle2" });

  // Lấy dữ liệu từ các biểu tượng
  const details = await page.evaluate(() => {
    const data = {};
    const icons = {
      // phone: "https://static.xx.fbcdn.net/rsrc.php/v4/yW/r/8k_Y-oVxbuU.png",
      // email: "https://static.xx.fbcdn.net/rsrc.php/v4/yE/r/2PIcyqpptfD.png",
      address: "https://static.xx.fbcdn.net/rsrc.php/v4/yW/r/8k_Y-oVxbuU.png",
    };

    Object.keys(icons).forEach((key) => {
      // const element = document.querySelector(`img[src*="${icons[key]}"]`);
      const element = document.querySelector(
        `img[src*="https://static.xx.fbcdn.net/rsrc.php/v4/yW/r/8k_Y-oVxbuU.png"]`
      );

      const parent = element?.closest("div")?.textContent.trim();
      if (parent) {
        data = parent.querySelector("div + div")?.textContent?.trim();
      }
    });

    return data;
  });

  await browser.close();
  return details;
}

app.listen(3000, () =>
  console.log("Server is running on http://localhost:3000")
);
