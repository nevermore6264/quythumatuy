const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const puppeteer = require("puppeteer-extra");
const StealthPlugin = require("puppeteer-extra-plugin-stealth");

puppeteer.use(StealthPlugin());
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
    const chunkSize = 1000; // Số lượng bản ghi mỗi file
    const filePaths = []; // Danh sách các file đã tạo
    console.log("Bắt đầu xử lý dữ liệu...");

    // Lặp qua dữ liệu, bắt đầu từ dòng 2 (bỏ qua tiêu đề)
    for (let i = 1; i < data.length; i++) {
      console.log(`Đang xử lý bản ghi thứ ${i}:`, data[i]);

      const row = data[i];
      const searchQuery = row[1];

      if (searchQuery && searchQuery.startsWith("Công an")) {
        const link = row[2];
        const details =
          link && link !== "-"
            ? await getFanpageDetails(link)
            : { phone: "-", email: "-", address: "-" };

        const phoneDetails = normalizePhoneNumber(details.phone);
        row[3] = phoneDetails.type === "Mobile" ? phoneDetails.phone : "-"; // Column D: DI ĐỘNG
        row[4] = phoneDetails.type === "Landline" ? phoneDetails.phone : "-"; // Column E: CỐ ĐỊNH
        row[5] =
          details.email && details.email !== "-"
            ? {
                t: "s",
                v: details.email,
                f: `HYPERLINK("mailto:${details.email}", "${details.email}")`,
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
        const link = row[2];
        row[2] = link || "-"; // Column C: GOV link
        row[3] = "-"; // Column D: DI ĐỘNG
        row[4] = "-"; // Column E: CỐ ĐỊNH
        row[5] = "-"; // Column F: EMAIL
        row[6] = "-"; // Column G: ĐỊA CHỈ
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

    // Chia dữ liệu thành các nhóm 1000 dòng
    const chunks = [];
    for (let i = 1; i < data.length; i += chunkSize) {
      chunks.push(data.slice(i, i + chunkSize));
    }

    // Tạo file Excel cho từng nhóm
    for (let i = 0; i < chunks.length; i++) {
      const chunk = chunks[i];
      const updatedWorksheet = xlsx.utils.aoa_to_sheet([data[0], ...chunk]); // Bao gồm tiêu đề
      const updatedWorkbook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, "Data");

      const updatedFilePath = path.join(
        __dirname,
        "uploads",
        `updated_data_${i + 1}_${Date.now()}.xlsx`
      );
      xlsx.writeFile(updatedWorkbook, updatedFilePath);
      filePaths.push(updatedFilePath); // Thêm file vào danh sách
    }

    fs.unlinkSync(filePath); // Xóa file gốc sau khi xử lý
    res.json({ filePaths }); // Trả về danh sách các file
  } catch (error) {
    console.error(error);
    res.status(500).send("An error occurred");
  }
});

async function getFanpageDetails(url) {
  const browser = await puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-blink-features=AutomationControlled",
    ],
  });
  const page = await browser.newPage();

  try {
    // Đặt User-Agent và các headers
    await page.setUserAgent(
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.90 Safari/537.36"
    );
    await page.setExtraHTTPHeaders({
      "Accept-Language": "en-US,en;q=0.9",
    });

    // Thiết lập viewport để giống người dùng thật
    await page.setViewport({
      width: 1366,
      height: 768,
    });

    // Điều hướng đến URL
    await page.goto(url, { waitUntil: "domcontentloaded" });

    // Cuộn trang để mô phỏng hành vi người dùng
    await page.evaluate(async () => {
      let totalHeight = 0;
      const distance = 100;
      while (totalHeight < document.body.scrollHeight) {
        window.scrollBy(0, distance);
        totalHeight += distance;
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    });

    // Đảm bảo phần tử chính xuất hiện trước khi thu thập dữ liệu
    await page.waitForSelector("body", { timeout: 10000 });

    // Thu thập thông tin
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

    return details;
  } catch (error) {
    console.error("Error fetching fanpage details:", error);
    return { phone: "-", email: "-", address: "-" }; // Trả về giá trị mặc định khi có lỗi
  } finally {
    await browser.close();
  }
}

function normalizePhoneNumber(rawPhone) {
  if (!rawPhone) return null; // Không có số trả về null

  // 1. Loại bỏ các ký tự không cần thiết
  let phone = rawPhone.replace(/[\s()-]/g, ""); // Xóa khoảng trắng, dấu ngoặc, gạch ngang

  // 2. Chuyển đổi đầu số quốc tế (+84 hoặc 84) thành 0
  phone = phone.replace(/^(\+84|84)/, "0");

  phone = phone.replace(/^(\+24|24)/, "024");

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
