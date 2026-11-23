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
    }
  };

  const VOLUMETRY_ROWS = [
    ["Certificado de Tradición Moto", "50", "135", "218", "291", "305", "378", "406", "1.783"],
    ["Certificado de Tradición Vehículo", "123", "356", "476", "532", "647", "835", "697", "3.666"],
    ["Duplicado Licencia Conducción Vehículo", "554", "668", "803", "829", "841", "1.043", "788", "5.526"],
    ["Duplicado de Licencia de Tránsito Motocicletas", "70", "104", "193", "321", "357", "396", "396", "1.837"],
    ["Duplicado de Placas Moto", "20", "33", "108", "145", "170", "246", "270", "992"],
    ["Expedición de fotocopias sin autenticar (ITC)", "19", "76", "122", "137", "149", "225", "228", "956"],
    ["Expedición Licencia de Conducción por Cambio de Documento Vehículo", "76", "167", "208", "299", "530", "399", "213", "1.892"],
    ["Expedición Licencia de Conducción por Cambio de Documento Moto", "1.149", "2.012", "2.828", "3.212", "2.316", "2.447", "2.514", "16.478"],
    ["Expedición Licencia Primera Vez Moto", "1.149", "2.012", "2.828", "3.212", "2.316", "2.447", "2.514", "16.478"],
    ["Expedición Licencia Primera Vez Vehículo", "2.144", "4.043", "4.915", "4.690", "2.876", "2.965", "2.977", "24.610"],
    ["Inscripción Persona Natural o Jurídica", "4.808", "8.456", "8.290", "6.723", "6.731", "7.076", "6.309", "48.393"],
    ["Inscripción de Prenda Vehículo", "459", "312", "221", "268", "250", "356", "363", "2.229"],
    ["Levantamiento de Prenda Vehículo", "358", "588", "804", "697", "804", "926", "855", "5.032"],
    ["Matrícula Inicial Motocicleta", "2.128", "3.825", "6.778", "5.828", "5.793", "6.397", "6.847", "37.596"],
    ["Matrícula Inicial Motocicleta con Prenda", "35", "113", "262", "244", "483", "650", "823", "2.610"],
    ["Matrícula Inicial Vehículo con Prenda", "168", "315", "567", "346", "456", "577", "494", "2.949"],
    ["Matrícula Inicial Vehículo Particular", "401", "307", "334", "237", "321", "363", "324", "2.287"],
    ["Matrícula Inicial Vehículo público", "173", "139", "247", "234", "253", "294", "327", "1.607"],
    ["Matrícula Inicial Vehículo público con Prenda", "79", "79", "109", "123", "228", "278", "250", "1.376"],
    ["Modificación de datos de inscripción de persona natural o jurídica, pública o privad.", "606", "855", "1.236", "1.761", "2.214", "2.008", "2.054", "10.734"],
    ["Radicación de Cuenta Vehículo", "104", "172", "336", "313", "338", "476", "331", "2.070"],
    ["Recategorización Licencia de Conducción hacia atrás Vehículo", "109", "240", "452", "698", "511", "590", "622", "3.222"],
    ["Refrendación Licencia Conducción Moto", "49", "360", "1.483", "5.785", "3.453", "2.967", "2.431", "16.528"],
    ["Refrendación Licencia Conducción Vehículo", "3.729", "4.701", "5.749", "12.091", "13.582", "10.932", "7.833", "58.617"],
    ["Tarjeta de Operación Individual (Renovación)", "72", "5", "233", "232", "227", "224", "173", "1.166"],
    ["Traspaso Moto", "762", "1.600", "2.536", "2.891", "3.182", "3.881", "3.947", "18.799"],
    ["Traspaso de Vehículo", "594", "1.345", "1.412", "1.676", "1.805", "2.302", "2.259", "11.393"]
  ];

  const DEMAND_RANKING = [
    ["1", "Refrendación Licencia Conducción Vehículo", "58.617"],
    ["2", "Inscripción Persona Natural o Jurídica", "48.393"],
    ["3", "Matrícula Inicial Motocicleta", "37.596"],
    ["4", "Expedición Licencia Primera Vez Vehículo", "24.610"],
    ["5", "Traspaso Moto", "18.799"],
    ["6", "Refrendación Licencia Conducción Moto", "16.528"],
    ["7", "Expedición Licencia de Conducción por Cambio de Documento Moto", "16.478"],
    ["8", "Expedición Licencia Primera Vez Moto", "16.478"],
    ["9", "Traspaso de Vehículo", "11.393"],
    ["10", "Modificación de datos de inscripción de persona natural o jurídica, pública o privad.", "10.734"],
    ["11", "Duplicado Licencia Conducción Vehículo", "5.526"],
    ["12", "Levantamiento de Prenda Vehículo", "5.032"],
    ["13", "Certificado de Tradición Vehículo", "3.666"],
    ["14", "Recategorización Licencia de Conducción hacia atrás Vehículo", "3.222"],
    ["15", "Matrícula Inicial Vehículo con Prenda", "2.949"],
    ["16", "Matrícula Inicial Motocicleta con Prenda", "2.610"],
    ["17", "Matrícula Inicial Vehículo Particular", "2.287"],
    ["18", "Inscripción de Prenda Vehículo", "2.229"],
    ["19", "Radicación de Cuenta Vehículo", "2.070"],
    ["20", "Expedición Licencia de Conducción por Cambio de Documento Vehículo", "1.892"],
    ["21", "Duplicado de Licencia de Tránsito Motocicletas", "1.837"],
    ["22", "Certificado de Tradición Moto", "1.783"],
    ["23", "Matrícula Inicial Vehículo público", "1.607"],
    ["24", "Matrícula Inicial Vehículo público con Prenda", "1.376"],
    ["25", "Tarjeta de Operación Individual (Renovación)", "1.166"],
    ["26", "Duplicado de Placas Moto", "992"],
    ["27", "Expedición de fotocopias sin autenticar (ITC)", "956"]
  ];

  function findVolumetryRow(name) {
    const normalize = (value) => value
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/\bde\b/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
    const normalizedName = normalize(name);
    return VOLUMETRY_ROWS.find((item) => normalize(item[0]) === normalizedName);
  }

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
      const rowLabel = row[0] == null ? "" : String(row[0]).trim();
      if (/^total\b/i.test(rowLabel)) {
        tr.className = "total-row";
      }
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

  function renderDataTable(headers, rows, label) {
    const table = document.createElement("table");
    table.className = "data-table";
    table.setAttribute("aria-label", label);

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    headers.forEach((header) => {
      const th = document.createElement("th");
      th.scope = "col";
      th.textContent = header;
      headerRow.append(th);
    });
    thead.append(headerRow);
    table.append(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((value, column) => {
        const td = document.createElement("td");
        td.textContent = value;
        if (column === row.length - 1) td.className = "total-value";
        tr.append(td);
      });
      tbody.append(tr);
    });
    table.append(tbody);
    return table;
  }

  function showVolumetry() {
    showContent(
      "Volumetría / Ingresos",
      "VOLUMEN DE TRÁMITES",
      "Selecciona un trámite para consultar su comportamiento por año.",
      {
        actions: "",
        content: `
          <div class="volumetry-content">
            <section class="demand-section" aria-labelledby="demandTitle">
              <div class="demand-heading">
                <div>
                  <span class="demand-eyebrow">RANKING ACUMULADO</span>
                  <h2 id="demandTitle">Trámites con mayor demanda</h2>
                </div>
                <span class="demand-period">2020 — 2026</span>
              </div>
              <div class="demand-grid" data-role="demand-grid"></div>
            </section>
            <section class="tramite-detail" data-role="tramite-detail" aria-live="polite"></section>
          </div>`
      }
    );

    const grid = contentView.querySelector("[data-role='demand-grid']");
    DEMAND_RANKING.forEach(([rank, name, total]) => {
      const row = findVolumetryRow(name);
      const card = document.createElement("button");
      card.type = "button";
      card.className = `demand-card${rank === "1" ? " is-selected" : ""}`;
      card.dataset.tramite = row ? row[0] : name;

      const cardRank = document.createElement("span");
      cardRank.className = "demand-rank";
      cardRank.textContent = rank.padStart(2, "0");

      const cardContent = document.createElement("span");
      cardContent.className = "demand-card-content";
      const cardTitle = document.createElement("strong");
      cardTitle.textContent = name;
      const cardTotal = document.createElement("b");
      cardTotal.textContent = total;
      const cardLabel = document.createElement("small");
      cardLabel.textContent = "Trámites acumulados";
      cardContent.append(cardTitle, cardTotal, cardLabel);
      card.append(cardRank, cardContent);
      grid.append(card);
    });

    showTramiteDetail(DEMAND_RANKING[0][1]);
  }

  function showTramiteDetail(name) {
    const row = findVolumetryRow(name);
    if (!row) return;

    const detail = contentView.querySelector("[data-role='tramite-detail']");
    const years = ["2020", "2021", "2022", "2023", "2024", "2025", "2026"];
    const values = row.slice(1, 8).map((value) => Number(String(value).replaceAll(".", "")));
    const maxValue = Math.max(...values);

    detail.innerHTML = "";
    const heading = document.createElement("div");
    heading.className = "detail-heading";
    const headingCopy = document.createElement("div");
    const eyebrow = document.createElement("span");
    eyebrow.textContent = "TRÁMITE SELECCIONADO";
    const title = document.createElement("h2");
    title.textContent = row[0];
    headingCopy.append(eyebrow, title);
    const total = document.createElement("div");
    total.className = "detail-total";
    total.innerHTML = `<strong>${row[8]}</strong><span>Total acumulado</span>`;
    heading.append(headingCopy, total);

    const chart = document.createElement("div");
    chart.className = "year-chart";
    years.forEach((year, index) => {
      const item = document.createElement("div");
      item.className = "year-bar";
      const value = document.createElement("strong");
      value.textContent = row[index + 1];
      const track = document.createElement("div");
      track.className = "year-bar-track";
      const fill = document.createElement("span");
      fill.style.height = `${Math.max(12, (values[index] / maxValue) * 100)}%`;
      track.append(fill);
      const label = document.createElement("small");
      label.textContent = year;
      item.append(value, track, label);
      chart.append(item);
    });

    detail.append(heading, chart);
  }

  /* Eventos delegados para botones de inicio, regreso y zoom */
  document.addEventListener("click", (event) => {
    const demandCard = event.target.closest("[data-tramite]");
    if (demandCard) {
      document.querySelectorAll("[data-tramite]").forEach((card) => card.classList.remove("is-selected"));
      demandCard.classList.add("is-selected");
      showTramiteDetail(demandCard.dataset.tramite);
      return;
    }

    const navigationButton = event.target.closest("[data-view]");
    if (navigationButton) {
      const { view } = navigationButton.dataset;
      if (view === "budget") showBudget();
      else if (view === "income") showVolumetry();
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
