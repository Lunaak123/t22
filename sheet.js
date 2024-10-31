const urlParams = new URLSearchParams(window.location.search);
const fileUrl = urlParams.get("fileUrl");

async function loadExcelFile(url) {
  const response = await fetch(url);
  const data = await response.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });

  const sheetNames = workbook.SheetNames;
  const sheetTabs = document.getElementById("sheet-tabs");

  sheetNames.forEach((sheetName, index) => {
    const button = document.createElement("button");
    button.textContent = sheetName;
    button.classList.add(index === 0 ? "active" : "");
    button.addEventListener("click", () => displaySheetData(workbook, sheetName));
    sheetTabs.appendChild(button);
  });

  // Display the first sheet by default
  displaySheetData(workbook, sheetNames[0]);
}

function displaySheetData(workbook, sheetName) {
  // Clear previous active tab style
  document.querySelectorAll("#sheet-tabs button").forEach(btn => btn.classList.remove("active"));
  // Set the clicked tab to active
  const activeTab = [...document.querySelectorAll("#sheet-tabs button")].find(
    btn => btn.textContent === sheetName
  );
  activeTab.classList.add("active");

  const sheet = workbook.Sheets[sheetName];
  const sheetData = XLSX.utils.sheet_to_html(sheet, { editable: false });
  document.getElementById("data-view").innerHTML = sheetData;
}

if (fileUrl) {
  loadExcelFile(fileUrl);
} else {
  alert("Invalid file URL. Please return to the previous page and enter a valid URL.");
}
