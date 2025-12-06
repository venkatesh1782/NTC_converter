const { ipcRenderer } = require("electron");

async function loadTemplates() {
  try {
    const templates = await ipcRenderer.invoke("get-image-templates");
    console.log("Image Templates Loaded:", templates);

    const area = document.getElementById("templateArea");
    if (!area) {
      console.error("No element with id 'templateArea' in HTML");
      return;
    }

    area.innerHTML = "";

    if (templates.length === 0) {
      area.innerHTML = "<p>No template images found.</p>";
      return;
    }

    templates.forEach((fileName) => {
      const img = document.createElement("img");
      img.src = fileName; // because image and HTML are in same folder
      img.style.width = "200px";
      img.style.margin = "10px";
      img.style.border = "1px solid #ccc";
      area.appendChild(img);
    });
  } catch (err) {
    console.error("Error loading templates:", err);
  }
}

window.onload = loadTemplates;
