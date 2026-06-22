const form = document.querySelector("#field-book-form");
const fileInput = document.querySelector("#source-file");
const statusEl = document.querySelector("#status");

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusEl.textContent = "Generando archivo...";
  statusEl.className = "status";

  const data = new FormData(form);
  const response = await fetch("/api/libro-campo", {
    method: "POST",
    body: data,
  });

  if (!response.ok) {
    const payload = await response.json().catch(() => ({ error: "No se pudo generar el archivo." }));
    statusEl.textContent = payload.error;
    statusEl.className = "status error";
    return;
  }

  const blob = await response.blob();
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);
  const filename = match ? decodeURIComponent(match[1]) : "Libro de campo.xlsx";
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);

  statusEl.textContent = `Archivo generado desde ${fileInput.files[0]?.name || "el Excel fuente"}.`;
  statusEl.className = "status success";
});
