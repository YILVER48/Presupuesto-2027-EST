/* Navegación sin recargas y vistas del presupuesto, indicadores e ingresos */
(() => {
  "use strict";

  const WORKBOOK_URL = "PRESUPUESTO  PROYECTADO 2027.xlsx";
  const SHEET_NAME = "PPTO 2023 BASE 2022";
  const DISPLAY_COLUMNS = [0, 25, 26, 27];
  const VIEWS = {
    indicators: {
      eyebrow: "DESEMPEÑO FINANCIERO",
      title: "Indicadores",
      subtitle: "Indicadores clave para el seguimiento del cierre.",
      image: "imagnes/indicadores.png",
      alt: "Indicadores financieros"
    },
    income: {
      eyebrow: "RESULTADOS DEL PERIODO",
      title: "Ingresos",
      subtitle: "Resumen de ingresos de la empresa.",
      image: "imagnes/ingresos.png",
      alt: "Ingresos de la empresa"
    }
  };

  const homeView = document.getElementById("homeView");
  const contentView = document.getElementById("contentView");
  const homeLink = document.getElementById("homeLink");
  let loadedSheetRows = null;
  let activeImageZoom = 1;
  let isLoadingWorkbook = false;

  /* Plantilla común de las vistas internas */
  function showContent(title, eyebrow, subtitle, body) {
    homeView.hidden = true;
    contentView.hidden = false;
    contentView.innerHTML = `
      <article class="content-card">
        <header class="view-header">
          <div class="view-title">
            <span class="eyebrow">${eyebrow}</span>
            <h1>${title}</h1>
            <p class="view-subtitle">${subtitle}</p>
          </div>
          <div class="view-actions">
            ${body.actions || ""}
            <button class="back-button" type="button" data-action="home">← Inicio</button>
          </div>
        </header>
        ${body.content}
      </article>`;
  }

  function returnHome() {
    contentView.hidden = true;
    contentView.replaceChildren();
    homeView.hidden = false;
  }

  /* Elige una anchura proporcional al contenido de cada columna, con límites legibles. */
  function getColumnWidths(rows, columnCount) {
    const widths = Array(columnCount).fill(76);
    rows.forEach((row) => {
      for (let column = 0; column < columnCount; column += 1) {
        const value = row[column] == null ? "" : String(row[column]);
        const measured = Math.min(42, Math.max(9, value.length + 2)) * 7;
        widths[column] = Math.max(widths[column], measured);
      }
    });
    return widths.map((width) => Math.min(320, width));
  }

  function renderSheet(rows) {
    if (!rows.length || !rows.some((row) => row.some((value) => value !== ""))) {
      return '<div class="table-empty">La hoja seleccionada no contiene datos para mostrar.</div>';
    }

    const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
    const widths = getColumnWidths(rows, columnCount);
    const table = document.createElement("table");
    table.className = "budget-table";
    table.setAttribute("aria-label", `Contenido de la hoja ${SHEET_NAME}`);
    const colgroup = document.createElement("colgroup");
    widths.forEach((width) => {
      const col = document.createElement("col");
      col.style.width = `${width}px`;
      colgroup.append(col);
    });
    table.append(colgroup);

    const head = document.createElement("thead");
    const headerRow = document.createElement("tr");
    for (let column = 0; column < columnCount; column += 1) {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = rows[0][column] == null ? "" : String(rows[0][column]);
      headerRow.append(th);
    }
    head.append(headerRow);
    table.append(head);

    const body = document.createElement("tbody");
    rows.slice(1).forEach((row) => {
      const tr = document.createElement("tr");
      for (let column = 0; column < columnCount; column += 1) {
        const td = document.createElement("td");
        td.textContent = row[column] == null ? "" : String(row[column]);
        tr.append(td);
      }
      body.append(tr);
    });
    table.append(body);

    const viewport = document.createElement("div");
    viewport.className = "table-viewport";
    viewport.append(table);
    return viewport;
  }

  /* Algunas hojas incluyen títulos y filas vacías antes del encabezado tabular. */
  function findHeaderIndex(rows) {
    let headerIndex = 0;
    let highestCellCount = 0;
    rows.forEach((row, index) => {
      const populatedCells = row.filter((value) => value != null && String(value).trim() !== "").length;
      if (populatedCells > highestCellCount) {
        highestCellCount = populatedCells;
        headerIndex = index;
      }
    });
    return headerIndex;
  }

  function showWorkbookError(message) {
    const messageElement = contentView.querySelector("[data-role='sheet-message']");
    if (messageElement) {
      messageElement.hidden = false;
      messageElement.querySelector("[data-role='message-text']").textContent = message;
    }
    const status = contentView.querySelector("[data-role='sheet-status']");
    if (status) status.textContent = "No se pudo cargar el presupuesto.";
  }

  function drawBudget(rows, sourceLabel) {
    const tableRows = rows
      .slice(findHeaderIndex(rows))
      .map((row) => DISPLAY_COLUMNS.map((column) => row[column] == null ? "" : row[column]))
      .filter((row, index) => index === 0 || row.some((value) => String(value).trim() !== ""));
    contentView.querySelector("[data-role='table-container']").replaceChildren(renderSheet(tableRows));
    const status = contentView.querySelector("[data-role='sheet-status']");
    const sheetName = document.createElement("span");
    sheetName.append("Hoja: ");
    const sheetNameStrong = document.createElement("strong");
    sheetNameStrong.textContent = SHEET_NAME;
    sheetName.append(sheetNameStrong);
    const rowCount = document.createElement("span");
    rowCount.textContent = `${Math.max(0, tableRows.length - 1).toLocaleString("es-CO")} registros · ${sourceLabel}`;
    status.replaceChildren(sheetName, rowCount);
    const notice = contentView.querySelector("[data-role='sheet-message']");
    notice.hidden = true;
  }

  async function parseWorkbook(data, sourceLabel) {
    if (!window.XLSX) {
      throw new Error("No se pudo cargar SheetJS. Verifica tu conexión a internet e inténtalo de nuevo.");
    }
    const workbook = window.XLSX.read(data, { type: "array", cellDates: true });
    const worksheet = workbook.Sheets[SHEET_NAME];
    if (!worksheet) {
      throw new Error(`El archivo no contiene la hoja "${SHEET_NAME}".`);
    }
    const rows = window.XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: false
    });
    loadedSheetRows = rows;
    drawBudget(rows, sourceLabel);
  }

  async function loadDefaultWorkbook() {
    if (isLoadingWorkbook) return;
    isLoadingWorkbook = true;
    const status = contentView.querySelector("[data-role='sheet-status']");
    status.textContent = "Cargando archivo de presupuesto…";
    try {
      const response = await fetch(WORKBOOK_URL);
      if (!response.ok) throw new Error(`No fue posible abrir el archivo (${response.status}).`);
      await parseWorkbook(await response.arrayBuffer(), "archivo del portal");
    } catch (error) {
      showWorkbookError(`${error.message} Si abriste esta página directamente, inicia un servidor local o selecciona el Excel desde tu equipo.`);
    } finally {
      isLoadingWorkbook = false;
    }
  }

  function showBudget() {
    showContent(
      "Presupuesto 2027",
      "PLANIFICACIÓN ANUAL",
      `Libro proyectado · hoja ${SHEET_NAME}`,
      {
        actions: "",
        content: `
          <div class="sheet-status" data-role="sheet-status">Preparando el presupuesto…</div>
          <div class="sheet-message" data-role="sheet-message" hidden>
            <span data-role="message-text"></span>
            <label class="file-button" for="workbookFile">Seleccionar Excel</label>
            <input class="file-input" id="workbookFile" type="file" accept=".xlsx,.xls">
          </div>
          <div data-role="table-container"><div class="table-empty">Cargando la hoja de presupuesto…</div></div>`
      }
    );
    contentView.querySelector("#workbookFile").addEventListener("change", async (event) => {
      const [file] = event.target.files;
      if (!file) return;
      try {
        await parseWorkbook(await file.arrayBuffer(), file.name);
      } catch (error) {
        showWorkbookError(error.message);
      }
    });
    if (loadedSheetRows) drawBudget(loadedSheetRows, "archivo del portal");
    else loadDefaultWorkbook();
  }

  function showImageView(view) {
    activeImageZoom = 1;
    showContent(view.title, view.eyebrow, view.subtitle, {
      actions: `
        <div class="zoom-controls" aria-label="Controles de zoom">
          <button class="zoom-button" type="button" data-action="zoom-out" aria-label="Reducir imagen">−</button>
          <span class="zoom-value" data-role="zoom-value">100%</span>
          <button class="zoom-button" type="button" data-action="zoom-in" aria-label="Ampliar imagen">+</button>
        </div>`,
      content: `
        <div class="image-stage" data-role="image-stage">
          <img class="zoomable-image" src="${view.image}" alt="${view.alt}">
        </div>`
    });
    const image = contentView.querySelector(".zoomable-image");
    image.addEventListener("error", () => {
      const stage = contentView.querySelector("[data-role='image-stage']");
      stage.innerHTML = `
        <div class="missing-image" role="status">
          <strong>La imagen aún no está disponible</strong>
          Para mostrar esta sección, agrega el archivo <code>${view.image}</code> al proyecto.
        </div>`;
      contentView.querySelector("[data-action='zoom-in']").disabled = true;
      contentView.querySelector("[data-action='zoom-out']").disabled = true;
    }, { once: true });
  }

  /* Eventos delegados para botones de inicio, regreso y zoom */
  document.addEventListener("click", (event) => {
    const navigationButton = event.target.closest("[data-view]");
    if (navigationButton) {
      const { view } = navigationButton.dataset;
      if (view === "budget") showBudget();
      else if (VIEWS[view]) showImageView(VIEWS[view]);
      return;
    }

    const actionButton = event.target.closest("[data-action]");
    if (!actionButton) return;
    if (actionButton.dataset.action === "home") returnHome();
    if (actionButton.dataset.action === "zoom-in" || actionButton.dataset.action === "zoom-out") {
      activeImageZoom = Math.min(2.5, Math.max(.5, activeImageZoom + (actionButton.dataset.action === "zoom-in" ? .2 : -.2)));
      const image = contentView.querySelector(".zoomable-image");
      if (image) image.style.setProperty("--image-zoom", activeImageZoom);
      const zoomValue = contentView.querySelector("[data-role='zoom-value']");
      if (zoomValue) zoomValue.textContent = `${Math.round(activeImageZoom * 100)}%`;
    }
  });

  homeLink.addEventListener("click", (event) => {
    event.preventDefault();
    returnHome();
  });
})();
