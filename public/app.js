const form = document.getElementById("planForm");
const steps = Array.from(document.querySelectorAll(".step"));
const backBtn = document.getElementById("backBtn");
const nextBtn = document.getElementById("nextBtn");
const downloadBtn = document.getElementById("downloadBtn");
const statusEl = document.getElementById("status");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");

let stepIndex = 0;

function refreshStep() {
  steps.forEach((step, index) => {
    step.classList.toggle("is-active", index === stepIndex);
  });

  backBtn.disabled = stepIndex === 0;

  const atLastStep = stepIndex === steps.length - 1;
  nextBtn.classList.toggle("hidden", atLastStep);
  downloadBtn.classList.toggle("hidden", !atLastStep);

  const percent = ((stepIndex + 1) / steps.length) * 100;
  progressBar.style.width = `${percent}%`;
  progressText.textContent = `Step ${stepIndex + 1} of ${steps.length}`;

  statusEl.textContent = "";
  statusEl.classList.remove("success");
}

function validateCurrentStep() {
  const fields = steps[stepIndex].querySelectorAll("input, textarea");

  for (const field of fields) {
    if (!field.checkValidity()) {
      field.reportValidity();
      return false;
    }
  }

  return true;
}

backBtn.addEventListener("click", () => {
  if (stepIndex > 0) {
    stepIndex -= 1;
    refreshStep();
  }
});

nextBtn.addEventListener("click", () => {
  if (!validateCurrentStep()) return;

  if (stepIndex < steps.length - 1) {
    stepIndex += 1;
    refreshStep();
  }
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!validateCurrentStep()) return;

  const formData = new FormData(form);
  const payload = Object.fromEntries(formData.entries());

  downloadBtn.disabled = true;
  statusEl.textContent = "Generating your Word project plan...";

  try {
    const response = await fetch("/api/project-plan", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error("Generation failed");
    }

    const blob = await response.blob();
    const projectName = (payload.projectName || "Project").trim() || "Project";
    const safeName = projectName.replace(/[^a-zA-Z0-9-_ ]/g, "").replace(/\s+/g, "-");

    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `${safeName}-Project-Plan.docx`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);

    statusEl.textContent = "Project plan generated and downloaded.";
    statusEl.classList.add("success");
  } catch (error) {
    statusEl.textContent = "Could not generate project plan. Please try again.";
    statusEl.classList.remove("success");
  } finally {
    downloadBtn.disabled = false;
  }
});

refreshStep();
