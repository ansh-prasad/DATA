// static/js/upload.js

document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("file");
  const fileText = document.getElementById("fileText");

  fileInput.addEventListener("change", function (e) {
    updateFileName(this.files);
  });

  // Add drag and drop functionality
  const fileUpload = document.querySelector(".file-upload");

  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    fileUpload.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    fileUpload.addEventListener(eventName, highlight, false);
  });

  ["dragleave", "drop"].forEach((eventName) => {
    fileUpload.addEventListener(eventName, unhighlight, false);
  });

  function highlight() {
    fileUpload.classList.add("highlight");
  }

  function unhighlight() {
    fileUpload.classList.remove("highlight");
  }

  fileUpload.addEventListener("drop", handleDrop, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;

    fileInput.files = files;
    updateFileName(files);
  }

  function updateFileName(files) {
    if (files && files.length > 0) {
      fileText.textContent = files[0].name;
    } else {
      fileText.textContent = "Drag files here";
    }
  }
});
