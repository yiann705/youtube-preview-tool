// main.js
const apiKey = import.meta.env.VITE_GOOGLE_API_KEY;
const sheetId = import.meta.env.VITE_SHEET_ID;
const sheetName = "工作表1";

function formatMonth(raw) {
  const match = raw.match(/(\d{4})年(\d{1,2})月/);
  if (!match) return raw;
  const [, year, month] = match;
  return `${year}年${month.padStart(2, "0")}月`;
}

function parseNumber(val) {
  if (!val) return 0;
  const n = parseFloat(String(val).replace(/[,%]/g, ""));
  return isNaN(n) ? 0 : n;
}

function formatInteger(n) {
  return Math.round(parseNumber(n)).toLocaleString("zh-TW");
}

function formatPercentage(n) {
  const val = parseFloat(String(n).replace(/[,%]/g, ""));
  if (isNaN(val)) return "0.000%";
  const pct = val > 1 ? val / 100 : val;
  return (pct * 100).toFixed(3) + "%";
}

function extractVideoId(url) {
  try {
    const u = new URL(url);
    const id = u.searchParams.get("v");
    if (id) return id;
  } catch {}
  const match = url.match(/(?:\/|v=)([a-zA-Z0-9_-]{11})/);
  return match?.[1] || "";
}

function createTable(headers, rows) {
  const container = document.getElementById("dataContainer");
  const table = document.createElement("table");
  table.style.borderCollapse = "collapse";
  table.style.width = "100%";
  table.style.maxWidth = "1200px";
  table.style.margin = "0 auto";
  table.style.background = "#fff";

  const borderStyle = "1px solid #ccc";
  const thStyle = `border: ${borderStyle}; padding: 8px; font-weight: bold; text-align: left;`;
  const tdStyle = `border: ${borderStyle}; padding: 8px; vertical-align: top;`;

  const sheetKeyMap = {
    "聯播網": "聯播網 (及搜尋夥伴)",
    "素材": "斷字網址"
  };

  const displayOrder = [
    "年月",
    "素材",
    "聯播網",
    "鉤子",
    "風格",
    "秒數",
    "費用",
    "CPI",
    "IR%",
    "CPM"
  ];

  const head = document.createElement("thead");
  const tr = document.createElement("tr");
  for (const key of displayOrder) {
    const th = document.createElement("th");
    th.textContent = key;
    th.setAttribute("style", thStyle);
    tr.appendChild(th);
  }
  head.appendChild(tr);
  table.appendChild(head);

  const body = document.createElement("tbody");
  for (const row of rows) {
    const tr = document.createElement("tr");
    for (const key of displayOrder) {
      const td = document.createElement("td");
      td.setAttribute("style", tdStyle);
      if (key === "素材") td.classList.add("video-cell");

      if (key === "年月") {
        td.textContent = formatMonth(row["月"] || row["年月"]);
      } else if (key === "素材") {
        const url = row["斷字網址"] || "";
        const isImage = /\.(jpg|jpeg|png|gif|webp)$/.test(url) || url.includes("googlesyndication.com");
        const videoId = extractVideoId(url);

        if (isImage) {
          td.innerHTML = `<img src="${url}" style="max-width:200px; max-height:112px; object-fit: contain;" />`;
        } else if (videoId) {
          td.innerHTML = `<iframe src="https://www.youtube.com/embed/${videoId}" frameborder="0" allowfullscreen style="width:200px; height:112px;"></iframe>`;
        } else {
          td.textContent = "(不支援的素材連結)";
        }
      } else if (key === "費用") {
        td.textContent = formatInteger(row["費用"]);
      } else if (key === "CPI") {
        td.textContent = Math.round(parseNumber(row["CPI"]));
      } else if (key === "IR%") {
        td.textContent = formatPercentage(row["IR%"]);
      } else if (key === "CPM") {
        td.textContent = Math.round(parseNumber(row["CPM"]));
      } else {
        const sheetKey = sheetKeyMap[key] || key;
        td.textContent = row[sheetKey] || "";
      }
      tr.appendChild(td);
    }
    body.appendChild(tr);
  }

  table.appendChild(body);
  container.innerHTML = "";
  container.appendChild(table);
}

async function loadDescription() {
  try {
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/B1?key=${apiKey}`
    );
    const data = await res.json();
    document.getElementById("titleDescription").textContent =
      data.values?.[0]?.[0] || "(無說明文字)";
  } catch (err) {
    console.error("讀取說明文字失敗", err);
    document.getElementById("titleDescription").textContent = "(讀取失敗)";
  }
}

async function loadSubDescription() {
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/N2?key=${apiKey}`;
    console.log("🚀 N2 Request URL:", url);
    const res = await fetch(url);
    const data = await res.json();
    console.log("📄 副標題資料：", data);
    document.getElementById("subDescription").textContent =
      data.values?.[0]?.[0] || "(副標題讀取失敗)";
  } catch (err) {
    console.error("❌ 副標題讀取失敗：", err);
    document.getElementById("subDescription").textContent = "(副標題讀取失敗)";
  }
}

async function loadSheetData() {
  try {
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${sheetName}?key=${apiKey}`
    );
    const data = await res.json();
    const [headers, ...rows] = data.values || [];
    if (!headers || !rows.length) {
      document.getElementById("dataContainer").textContent = "沒有資料";
      return;
    }

    const normalized = rows.map((r) => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = r[i];
      });
      return obj;
    });

    createTable(headers, normalized);
  } catch (err) {
    console.error("讀取表格資料失敗", err);
    document.getElementById("dataContainer").textContent = "讀取資料失敗";
  }
}

loadDescription();
loadSubDescription();
loadSheetData();

const versionBox = document.createElement("div");
versionBox.textContent = __APP_VERSION__;
versionBox.setAttribute("style", `
  position: fixed;
  bottom: 10px;
  right: 10px;
  font-size: 12px;
  color: #666;
  background: rgba(255,255,255,0.8);
  padding: 4px 8px;
  border-radius: 6px;
  box-shadow: 0 0 4px rgba(0,0,0,0.1);
  z-index: 999;
`);
document.body.appendChild(versionBox);
console.log("🛠️ App Version:", __APP_VERSION__);
