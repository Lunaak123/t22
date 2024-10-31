document.getElementById("fetch-btn").addEventListener("click", () => {
    const excelUrl = document.getElementById("excel-url").value.trim();
    if (excelUrl) {
      window.location.href = `sheet.html?fileUrl=${encodeURIComponent(excelUrl)}`;
    } else {
      alert("Please enter a valid Excel URL.");
    }
  });
  