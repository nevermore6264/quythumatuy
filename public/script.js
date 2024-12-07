document.getElementById("processBtn").addEventListener("click", async () => {
  const fileInput = document.getElementById("fileInput");
  if (!fileInput.files.length) {
    alert("Vui lòng tải lên tệp Excel!");
    return;
  }

  const formData = new FormData();
  formData.append("file", fileInput.files[0]);

  document.getElementById("output").textContent = "Đang xử lý...";

  try {
    const response = await fetch("/process", {
      method: "POST",
      body: formData,
    });
    const data = await response.json();
    document.getElementById("output").textContent = "Done";
    // document.getElementById("output").textContent = JSON.stringify(
    //   data,
    //   null,
    //   2
    // );
  } catch (error) {
    document.getElementById("output").textContent =
      "Đã xảy ra lỗi: " + error.message;
  }
});
