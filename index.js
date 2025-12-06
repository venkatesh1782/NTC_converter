// ---------------------------
// Template selection logic
// ---------------------------
function selectTemplate(templateId) {
  const templates = document.querySelectorAll(".template-thumb");
  templates.forEach((t) => t.classList.remove("selected"));
  const selected = document.getElementById(templateId);
  selected.classList.add("selected");
  document.getElementById("template-input").value = templateId;

  const previewContainer = document.getElementById("template-preview");
  previewContainer.innerHTML = `
    <h3>Selected Template Preview:</h3>
    <img src="${selected.src}" class="preview-image" alt="Selected Template">
  `;
}

// ---------------------------
// Utility Functions
// ---------------------------
function updateImageCount() {
  const input = document.getElementById("images-input");
  document.getElementById("total-images").textContent =
    "Total Images: " + input.files.length;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(e);
    reader.readAsDataURL(file);
  });
}

async function imgToBase64(url) {
  const response = await fetch(url);
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// ---------------------------
// Folder Selection Logic (NEW)
// ---------------------------
let folderImages = []; // store images chosen via folder picker

function handleFolderSelection() {
  const input = document.getElementById("folder-input");
  const files = Array.from(input.files).filter(f => f.type.startsWith("image/"));
  folderImages = files;
  document.getElementById("total-folder-images").textContent =
    "Total Folder Images: " + folderImages.length;
}

// ---------------------------
// Generate PPT Logic
// ---------------------------
document.querySelector(".generate-btn").addEventListener("click", async function (e) {
  e.preventDefault();
  const loadingLine = document.getElementById("loading-line");
  if (loadingLine) {
    loadingLine.style.display = "block";
    loadingLine.textContent = "⏳ Generating PPT, please wait...";
    loadingLine.style.color = "#0078D4";
  }

  try {
    const fields = [
      "study_from", "display_title", "project_owner",
      "prepared_by", "submitted_on", "client",
      "doc_no", "project", "surveyed_on", "surveyed_by"
    ];
    const formData = {};
    fields.forEach(f => formData[f] = document.getElementById(f)?.value || "");

    let templatePath = document.getElementById("template-input").value.trim();
    let templateBase64;

    // ✅ Load selected template image (old or new)
    try {
      if (templatePath.endsWith(".jpg") || templatePath.endsWith(".png")) {
        templateBase64 = await imgToBase64(templatePath);
      } else {
        const templateImg = document.getElementById(templatePath);
        if (!templateImg) throw new Error("Template not found!");
        templateBase64 = await imgToBase64(templateImg.src);
      }
    } catch (err) {
      alert("⚠️ Could not load selected template image.\nPlease ensure the file exists in the same folder as index1.html");
      console.error(err);
      if (loadingLine) loadingLine.style.display = "none";
      return;
    }

    // ✅ Combine manual + folder images
    let imageFiles = Array.from(document.getElementById("images-input").files);
    if (folderImages.length > 0) {
      imageFiles = [...imageFiles, ...folderImages];
    }

    if (!imageFiles.length) {
      alert("Please select images or choose a folder to include in the PPT.");
      if (loadingLine) loadingLine.style.display = "none";
      return;
    }

    const pptx = new PptxGenJS();

    // ---------------------------
    // Template-based backgrounds
    // ---------------------------
   // ---------------------------
// Template background + cover / end page logic
// ---------------------------
// ---------------------------
// Template background + cover / end page logic
// ---------------------------
let frontImagePath = "";
let lastImagePath = "";

// ✅ Template 5 – use its own front & last pages
if (templatePath.includes("template5")) {
  frontImagePath = "template5 front page.jpg";
  lastImagePath = "template5 last page.jpg";
}

// ✅ Template 6 – use its own front & last pages
else if (templatePath.includes("template6")) {
  frontImagePath = "template6 front page.jpg";
  lastImagePath = "template6 last page.jpg";
}

// ✅ Template 8 – use its own front & last pages
else if (templatePath.includes("template8")) {
  frontImagePath = "template8 front page.jpg";
  lastImagePath = "template8 last page.jpg";
}

// ✅ Template 9 – use its own front & last pages
else if (templatePath.includes("template9")) {
  frontImagePath = "template9 front page.jpg";
  lastImagePath = "template9 last page.jpg";
}

// ✅ Template1 and Template3 (old NTC style)
else if (templatePath.includes("template1") || templatePath.includes("template3")) {
  frontImagePath = "ppt1 template1.jpg";
  lastImagePath = "ppt1 last page.jpg";
}

// ✅ Template2 (old GE style)
else if (templatePath.includes("template2")) {
  frontImagePath = "ppt2 template2.jpg";
  lastImagePath = "ppt2 lastpage.jpg";
}

// ✅ Any other template uses default NTC front/last
else {
  frontImagePath = "ppt1 template1.jpg";
  lastImagePath = "ppt1 last page.jpg";
}



    const frontBase64 = await imgToBase64(frontImagePath);
    const lastBase64 = await imgToBase64(lastImagePath);

    pptx.defineSlideMaster({
      title: "TemplateMaster",
      background: { data: templateBase64 },
    });

    // ---------------------------
    // Slide 1: Front Cover
    // ---------------------------
    const frontSlide = pptx.addSlide();
frontSlide.addImage({ data: frontBase64, x: 0, y: 0, w: 10, h: 7.5 });




    // ---------------------------
    // Slide 2: Table Details
    // ---------------------------
    const tableSlide = pptx.addSlide({ masterName: "TemplateMaster" });
    const tableData = [
      [
        {
          text: formData.display_title || "Project Information",
          options: { colspan: 4, bold: true, fontSize: 20, align: "center", fill: "F2F2F2" },
        },
      ],
      [
        { text: "Study From:", options: { bold: true } }, formData.study_from,
        { text: "Project:", options: { bold: true } }, formData.project,
      ],
      [
        { text: "Client:", options: { bold: true } }, formData.client,
        { text: "Project Owner:", options: { bold: true } }, formData.project_owner,
      ],
      [
        { text: "Doc No:", options: { bold: true } }, formData.doc_no,
        { text: "Surveyed On:", options: { bold: true } }, formData.surveyed_on,
      ],
      [
        { text: "Surveyed By:", options: { bold: true } }, formData.surveyed_by,
        { text: "Prepared By:", options: { bold: true } }, formData.prepared_by,
      ],
      [
        { text: "Submitted On:", options: { bold: true } }, formData.submitted_on, "", "",
      ],
    ];

    tableSlide.addTable(tableData, {
      x: 0.6, y: 1.0, w: 9,
      colW: [1.5, 2.3, 1.7, 3.0],
      border: { pt: 1, color: "000000" },
      fontSize: 14, color: "000000", align: "left",
    });

    tableSlide.addText("NTC Logistics India Pvt Limited", {
      x: 0, y: 5.8, w: 10, fontSize: 14, bold: true, color: "000000", align: "center",
    });

    // ---------------------------
    // Slides 3+: Image Slides
    // ---------------------------
    const base64Images = await Promise.all(
      Array.from(imageFiles).map(file => fileToBase64(file))
    );

    for (const imgBase64 of base64Images) {
      const slide = pptx.addSlide({ masterName: "TemplateMaster" });
      await new Promise(resolve => {
        const img = new Image();
        img.src = imgBase64;
        img.onload = () => {
          const slideW = 10.0, slideH = 7.5;
          const padX = 0.6, padY = 0.6;
          const maxW = slideW - padX * 2, maxH = (slideH - padY * 2) * 0.65;
          const imgRatio = img.width / img.height, areaRatio = maxW / maxH;
          let w, h;
          if (imgRatio > areaRatio) { w = maxW; h = w / imgRatio; }
          else { h = maxH; w = h * imgRatio; }
          const x = (slideW - w) / 2, y = padY + (maxH - h) / 2;
          slide.addImage({ data: imgBase64, x, y, w, h, sizing: { type: "contain", w, h } });
          resolve();
        };
      });
    }

    // ---------------------------
    // Final Slide: End Page
    // ---------------------------
    const lastSlide = pptx.addSlide();
   lastSlide.addImage({ data: lastBase64, x: 0, y: 0, w: 10, h: 7.5 });;


    // ---------------------------
    // Save PPT
    // ---------------------------
    if (loadingLine) loadingLine.textContent = "🧠 Finalizing slides...";
    await new Promise(res => setTimeout(res, 1000));

    await pptx.writeFile({
      fileName: `${formData.display_title || "NTC_Presentation"}.pptx`,
    });

    if (loadingLine) {
      loadingLine.textContent = "✅ PPT generated successfully!";
      loadingLine.style.color = "green";
      setTimeout(() => (loadingLine.style.display = "none"), 1000);
    }

  } catch (err) {
    console.error(err);
    if (loadingLine) {
      loadingLine.textContent = "❌ Error generating PPT!";
      loadingLine.style.color = "red";
      setTimeout(() => (loadingLine.style.display = "none"), 1500);
    }
  }
});

// ---------------------------
// Bind updateImageCount
// ---------------------------
let imageCountTimeout;
document.getElementById("images-input").addEventListener("change", () => {
  clearTimeout(imageCountTimeout);
  imageCountTimeout = setTimeout(updateImageCount, 300);
});

// ---------------------------
// Auto-load new templates dynamically
// ---------------------------
async function loadNewTemplates() {
  try {
    const response = await fetch("./");
    const text = await response.text();

    const matches = [...text.matchAll(/href="(.*?)"/g)]
      .map(m => decodeURIComponent(m[1]))
     .filter(f =>
  f.match(/\.(jpg|jpeg|png)$/i) &&
  (
    f.toLowerCase().includes("template5") ||
    f.toLowerCase().includes("template6") ||
    f.toLowerCase().includes("template7") ||
    f.toLowerCase().includes("template8") ||
    f.toLowerCase().includes("template9")
  ) &&
  !f.toLowerCase().includes("front") &&
  !f.toLowerCase().includes("last")
)



    const dropdown = document.getElementById("new-template-dropdown");
    if (!dropdown) return;

    dropdown.innerHTML = '<option value="">-- Select New Template --</option>';
    matches.forEach(file => {
      const cleanName = file.split("/").pop().replace(/\.[^/.]+$/, "").replace(/template/i, "Template ");
      const opt = document.createElement("option");
      opt.value = file;
      opt.textContent = cleanName;
      dropdown.appendChild(opt);
    });

    console.log("✅ Loaded new templates:", matches);
  } catch (err) {
    console.warn("⚠️ Auto-detection failed, using fallback list.");
    const fallbackFiles = ["template5.jpg", "template6.jpg", "template7.jpg"];
    const dropdown = document.getElementById("new-template-dropdown");
    dropdown.innerHTML = '<option value="">-- Select New Template --</option>';
    fallbackFiles.forEach(file => {
      const opt = document.createElement("option");
      opt.value = file;
      opt.textContent = file.replace(/template/i, "Template ").replace(/\.[^/.]+$/, "");
      dropdown.appendChild(opt);
    });
  }
}

function handleNewTemplateSelection() {
  const dropdown = document.getElementById("new-template-dropdown");
  const selectedTemplate = dropdown.value;
  if (!selectedTemplate) return;

  document.getElementById("template-input").value = selectedTemplate;

  const previewContainer = document.getElementById("template-preview");
  previewContainer.innerHTML = `
    <h3>Selected Template Preview:</h3>
    <img src="${selectedTemplate}" class="preview-image" alt="Selected Template">
  `;
}

window.addEventListener("DOMContentLoaded", loadNewTemplates);
