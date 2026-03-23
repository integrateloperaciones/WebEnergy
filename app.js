const API_URL = "https://script.google.com/macros/s/AKfycbzSBGp3_lQ7EIHmLBzOGWN3hyY-ePiX53oDMQiT_-fIWZ2t-Hj1CPNVmfEcWP4PIe8I/exec";

function mostrarFormulario(tipo) {
  document.getElementById("impedancia").style.display = "none";
  document.getElementById("descarga").style.display = "none";
  document.getElementById(tipo).style.display = "block";
}

/* =========================
   UTILIDADES GENERALES
========================= */

function generarOpcionesBanco(valorSeleccionado = "") {
  let opciones = '<option value="">Seleccione</option>';
  for (let i = 1; i <= 10; i++) {
    const valor = `BB${i}`;
    opciones += `<option value="${valor}" ${valor === valorSeleccionado ? "selected" : ""}>${valor}</option>`;
  }
  return opciones;
}

function generarOpcionesCelda(valorSeleccionado = "") {
  let opciones = '<option value="">Seleccione</option>';
  for (let i = 1; i <= 24; i++) {
    const valor = `C${i}`;
    opciones += `<option value="${valor}" ${valor === valorSeleccionado ? "selected" : ""}>${valor}</option>`;
  }
  return opciones;
}

function normalizarResultado(valor) {
  const texto = String(valor || "").trim().toLowerCase();

  if (texto === "buena" || texto === "bueno") return "Buena";
  if (texto === "regular") return "Regular";
  if (texto === "mala" || texto === "malo") return "Mala";
  if (texto === "pesima" || texto === "pésima" || texto === "pesimo" || texto === "pésimo") return "Pésima";

  return texto ? String(valor).trim() : "Buena";
}

function formatearFechaExcel(valor) {
  if (!valor) return "";

  if (typeof valor === "number") {
    const fecha = XLSX.SSF.parse_date_code(valor);
    if (!fecha) return "";
    const mm = String(fecha.m).padStart(2, "0");
    const dd = String(fecha.d).padStart(2, "0");
    return `${fecha.y}-${mm}-${dd}`;
  }

  const texto = String(valor).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;

  const partes = texto.split(/[\/\-]/);
  if (partes.length === 3) {
    if (partes[0].length === 4) {
      return `${partes[0]}-${String(partes[1]).padStart(2, "0")}-${String(partes[2]).padStart(2, "0")}`;
    } else {
      return `${partes[2]}-${String(partes[1]).padStart(2, "0")}-${String(partes[0]).padStart(2, "0")}`;
    }
  }

  return "";
}

function exportarElementoPDF(idElemento, nombreArchivo) {
  const elemento = document.getElementById(idElemento);
  if (!elemento) return;

  elemento.classList.add("pdf-export");

  const opciones = {
    margin: [0.2, 0.2, 0.2, 0.2],
    filename: nombreArchivo,
    image: { type: "jpeg", quality: 1 },
    html2canvas: {
      scale: 3,
      useCORS: true,
      scrollX: 0,
      scrollY: 0
    },
    jsPDF: {
      unit: "in",
      format: "a4",
      orientation: "landscape"
    },
    pagebreak: {
      mode: ["css", "legacy"]
    }
  };

  html2pdf()
    .set(opciones)
    .from(elemento)
    .save()
    .then(() => {
      elemento.classList.remove("pdf-export");
    })
    .catch(() => {
      elemento.classList.remove("pdf-export");
    });
}

/* =========================
   IMPEDANCIA
========================= */

function agregarFila(
  banco = "",
  celda = "",
  vReposo = "",
  fechaFabricacion = "",
  riOrigen = "",
  riImpedancia = "",
  resultado = "Buena",
  observaciones = ""
) {
  const tabla = document.querySelector("#tabla tbody");

  const fila = `
    <tr>
      <td><select>${generarOpcionesBanco(banco)}</select></td>
      <td><select>${generarOpcionesCelda(celda)}</select></td>
      <td><input type="number" step="0.01" value="${vReposo}"></td>
      <td><input type="date" value="${fechaFabricacion}"></td>
      <td><input type="number" step="0.01" value="${riOrigen}"></td>
      <td><input type="number" step="0.01" value="${riImpedancia}"></td>
      <td>
        <select>
          <option value="Buena" ${resultado === "Buena" ? "selected" : ""}>Buena</option>
          <option value="Regular" ${resultado === "Regular" ? "selected" : ""}>Regular</option>
          <option value="Mala" ${resultado === "Mala" ? "selected" : ""}>Mala</option>
          <option value="Pésima" ${resultado === "Pésima" ? "selected" : ""}>Pésima</option>
        </select>
      </td>
      <td><textarea rows="1">${observaciones}</textarea></td>
    </tr>
  `;

  tabla.insertAdjacentHTML("beforeend", fila);
}

function cargarExcel(event) {
  const archivo = event.target.files[0];
  if (!archivo) return;

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });

      const tbody = document.querySelector("#tabla tbody");
      tbody.innerHTML = "";

      filas.forEach(fila => {
        const banco = fila["N° Banco"] ?? fila["No Banco"] ?? fila["Banco"] ?? fila["n°bancoCelda"] ?? "";
        const celda = fila["N° Celda"] ?? fila["No Celda"] ?? fila["Celda"] ?? fila["n°celda"] ?? "";
        const vReposo = fila["V reposo (V)"] ?? fila["V reposo"] ?? fila["reposo"] ?? "";
        const fechaFabricacion = formatearFechaExcel(
          fila["Fecha fabricación"] ??
          fila["Fecha de fabricación"] ??
          fila["fecha_de_fabric"] ??
          ""
        );
        const riOrigen = fila["RI Origen"] ?? fila["ri_origen"] ?? "";
        const riImpedancia = fila["RI Impedancia"] ?? fila["ri_impedancia"] ?? "";
        const resultado = normalizarResultado(
          fila["Resultado"] ?? fila["resultado"] ?? fila["Resultado por celda"] ?? "Buena"
        );
        const observaciones = fila["Observaciones"] ?? fila["observacion"] ?? "";

        const filaVacia =
          banco === "" &&
          celda === "" &&
          vReposo === "" &&
          fechaFabricacion === "" &&
          riOrigen === "" &&
          riImpedancia === "" &&
          observaciones === "";

        if (!filaVacia) {
          agregarFila(
            banco,
            celda,
            vReposo,
            fechaFabricacion,
            riOrigen,
            riImpedancia,
            resultado,
            observaciones
          );
        }
      });

      alert("✅ Excel cargado correctamente");
    } catch (error) {
      console.error(error);
      alert("Error al leer el archivo Excel");
    }
  };

  reader.readAsArrayBuffer(archivo);
}

function descargarPlantilla() {
  const datos = [
    ["N° Banco", "N° Celda", "V reposo (V)", "Fecha fabricación", "RI Origen", "RI Impedancia", "Resultado", "Observaciones"],
    ["BB1", "C1", "", "", "", "", "Buena", ""],
    ["BB1", "C2", "", "", "", "", "Regular", ""],
    ["BB1", "C3", "", "", "", "", "Mala", ""],
    ["BB1", "C4", "", "", "", "", "Pésima", ""]
  ];

  const hoja = XLSX.utils.aoa_to_sheet(datos);
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, "Plantilla");
  XLSX.writeFile(libro, "plantilla_detalle_celdas.xlsx");
}

async function guardarDatos() {
  const filas = document.querySelectorAll("#tabla tbody tr");
  let datos = [];

  filas.forEach(f => {
    const c = f.querySelectorAll("input, select, textarea");

    // Orden exacto según tu hoja IMPEDANCIA
    datos.push([
      document.getElementById("fecha").value,               // fecha
      document.getElementById("registro").value,            // n°registro
      document.getElementById("cu").value,                  // cu
      document.getElementById("local").value,               // local
      document.getElementById("ups").value,                 // rectificador|ups
      document.getElementById("num_bancos").value,          // n°banco
      document.getElementById("marca").value,               // marca
      document.getElementById("modelo_equipo").value,       // modelo
      document.getElementById("capacidad").value,           // capacidad
      document.getElementById("fecha_fabricacion").value,   // fecha_de_fabric

      c[0].value, // n°bancoCelda
      c[1].value, // n°celda
      c[2].value, // reposo
      c[3].value, // fecha_de_fabric
      c[4].value, // ri_origen
      c[5].value, // ri_impedancia
      c[6].value, // resultado
      c[7].value  // Observaciones
    ]);
  });

  try {
    const formData = new URLSearchParams();
    formData.append("sheet", "IMPEDANCIA");
    formData.append("data", JSON.stringify(datos));

    const response = await fetch(API_URL, {
      method: "POST",
      body: formData
    });

    const result = await response.json();
    if (result.status === "success") {
      alert("✅ Su registro fue exitoso");
    } else {
      alert("Error al guardar: " + (result.message || "Error desconocido"));
    }
  } catch (error) {
    console.error(error);
    alert("Error al guardar");
  }
}

function limpiar() {
  document.querySelector("#tabla tbody").innerHTML = "";

  document.querySelectorAll("#impedancia input").forEach(i => {
    if (i.type !== "file") i.value = "";
  });

  document.querySelectorAll("#impedancia textarea").forEach(t => {
    t.value = "";
  });

  document.querySelectorAll("#impedancia select").forEach(s => {
    s.selectedIndex = 0;
  });

  const archivoExcel = document.getElementById("archivoExcel");
  if (archivoExcel) archivoExcel.value = "";
}

function exportarPDF() {
  exportarElementoPDF("exportImpedancia", "reporte_impedancia.pdf");
}

/* =========================
   DESCARGA
========================= */

function agregarFilaDescarga(
  banco = "",
  disyuntor = "",
  iEstable = "",
  iMax = "",
  aporte = "",
  cablesPolos = "",
  calibreCable = "",
  longitudCable = "",
  vol1 = "",
  car1 = "",
  vol2 = "",
  car2 = "",
  vol3 = "",
  car3 = "",
  vol4 = "",
  car4 = ""
) {
  const tabla = document.querySelector("#tablaDescarga tbody");

  const fila = `
    <tr>
      <td><select>${generarOpcionesBanco(banco)}</select></td>
      <td><input type="number" step="0.01" value="${disyuntor}"></td>
      <td><input type="number" step="0.01" value="${iEstable}"></td>
      <td><input type="number" step="0.01" value="${iMax}"></td>
      <td><input type="number" step="0.01" value="${aporte}"></td>
      <td><input type="number" step="0.01" value="${cablesPolos}"></td>
      <td><input type="number" step="0.01" value="${calibreCable}"></td>
      <td><input type="number" step="0.01" value="${longitudCable}"></td>
      <td><input type="number" step="0.01" value="${vol1}"></td>
      <td><input type="number" step="0.01" value="${car1}"></td>
      <td><input type="number" step="0.01" value="${vol2}"></td>
      <td><input type="number" step="0.01" value="${car2}"></td>
      <td><input type="number" step="0.01" value="${vol3}"></td>
      <td><input type="number" step="0.01" value="${car3}"></td>
      <td><input type="number" step="0.01" value="${vol4}"></td>
      <td><input type="number" step="0.01" value="${car4}"></td>
    </tr>
  `;

  tabla.insertAdjacentHTML("beforeend", fila);
}

function cargarExcelDescarga(event) {
  const archivo = event.target.files[0];
  if (!archivo) return;

  const reader = new FileReader();

  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const hoja = workbook.Sheets[workbook.SheetNames[0]];
      const filas = XLSX.utils.sheet_to_json(hoja, { defval: "" });

      const tbody = document.querySelector("#tablaDescarga tbody");
      tbody.innerHTML = "";

      filas.forEach(fila => {
        const banco = fila["N° Banco"] ?? fila["Banco"] ?? fila["n°bancoBanco"] ?? "";
        const disyuntor = fila["Disyuntor (A)"] ?? fila["disyuntor"] ?? "";
        const iEstable = fila["I estable (A)"] ?? fila["estable"] ?? "";
        const iMax = fila["I máx (A)"] ?? fila["I max (A)"] ?? fila["max"] ?? "";
        const aporte = fila["% aporte"] ?? fila["aporte"] ?? "";
        const cablesPolos = fila["Cables x Polos"] ?? fila["cablePolos"] ?? "";
        const calibreCable = fila["Calibre cable"] ?? fila["CalibreCable"] ?? "";
        const longitudCable = fila["Longitud cable"] ?? fila["longitudCable"] ?? "";
        const vol1 = fila["Med1_Voltaje (1 min)"] ?? fila["vol1"] ?? "";
        const car1 = fila["Med1_Carga (A)"] ?? fila["car1"] ?? "";
        const vol2 = fila["Med2_Voltaje (10 min)"] ?? fila["vol2"] ?? "";
        const car2 = fila["Med2_Carga (A)"] ?? fila["car2"] ?? "";
        const vol3 = fila["Med3_Voltaje (20 min)"] ?? fila["vol3"] ?? "";
        const car3 = fila["Med3_Carga (A)"] ?? fila["car3"] ?? "";
        const vol4 = fila["Med4_Voltaje (30 min)"] ?? fila["vol4"] ?? "";
        const car4 = fila["Med4_Carga (A)"] ?? fila["car4"] ?? "";

        const filaVacia =
          banco === "" &&
          disyuntor === "" &&
          iEstable === "" &&
          iMax === "" &&
          aporte === "" &&
          cablesPolos === "" &&
          calibreCable === "" &&
          longitudCable === "" &&
          vol1 === "" &&
          car1 === "" &&
          vol2 === "" &&
          car2 === "" &&
          vol3 === "" &&
          car3 === "" &&
          vol4 === "" &&
          car4 === "";

        if (!filaVacia) {
          agregarFilaDescarga(
            banco,
            disyuntor,
            iEstable,
            iMax,
            aporte,
            cablesPolos,
            calibreCable,
            longitudCable,
            vol1,
            car1,
            vol2,
            car2,
            vol3,
            car3,
            vol4,
            car4
          );
        }
      });

      alert("✅ Excel de descarga cargado correctamente");
    } catch (error) {
      console.error(error);
      alert("Error al leer el archivo Excel de descarga");
    }
  };

  reader.readAsArrayBuffer(archivo);
}

function descargarPlantillaDescarga() {
  const datos = [
    [
      "N° Banco", "Disyuntor (A)", "I estable (A)", "I máx (A)", "% aporte",
      "Cables x Polos", "Calibre cable", "Longitud cable",
      "vol1", "car1", "vol2", "car2", "vol3", "car3", "vol4", "car4"
    ],
    ["BB1", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
    ["BB2", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
  ];

  const hoja = XLSX.utils.aoa_to_sheet(datos);
  const libro = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(libro, hoja, "PlantillaDescarga");
  XLSX.writeFile(libro, "plantilla_detalle_descarga.xlsx");
}

async function guardarDatosDescarga() {
  const filas = document.querySelectorAll("#tablaDescarga tbody tr");
  let datos = [];

  filas.forEach(f => {
    const c = f.querySelectorAll("input, select, textarea");

    // Orden exacto según tu hoja DESCARGA
    datos.push([
      document.getElementById("d_fecha").value,             // fecha
      document.getElementById("d_registro").value,          // n°registro
      document.getElementById("d_cu").value,                // cu
      document.getElementById("d_local").value,             // local
      document.getElementById("d_ups").value,               // rectificador|ups
      document.getElementById("d_num_bancos").value,        // n°banco
      document.getElementById("d_corriente_total").value,   // corriente_total
      document.getElementById("d_tiempo_descarga").value,   // tiempo_descarg
      document.getElementById("d_observacion").value,       // observacion
      document.getElementById("d_shunt").value,             // shunt
      document.getElementById("d_c").value,                 // c°

      c[0].value,  // n°bancoBanco
      c[1].value,  // disyuntor
      c[2].value,  // estable
      c[3].value,  // max
      c[4].value,  // aporte
      c[5].value,  // cablePolos
      c[6].value,  // CalibreCable
      c[7].value,  // longitudCable
      c[8].value,  // vol1
      c[9].value,  // car1
      c[10].value, // vol2
      c[11].value, // car2
      c[12].value, // vol3
      c[13].value, // car3
      c[14].value, // vol4
      c[15].value  // car4
    ]);
  });

  try {
    const formData = new URLSearchParams();
    formData.append("sheet", "DESCARGA");
    formData.append("data", JSON.stringify(datos));

    const response = await fetch(API_URL, {
      method: "POST",
      body: formData
    });

    const result = await response.json();
    if (result.status === "success") {
      alert("✅ Registro de descarga guardado correctamente");
    } else {
      alert("Error al guardar descarga: " + (result.message || "Error desconocido"));
    }
  } catch (error) {
    console.error(error);
    alert("Error al guardar descarga");
  }
}

function limpiarDescarga() {
  document.querySelector("#tablaDescarga tbody").innerHTML = "";

  document.querySelectorAll("#descarga input").forEach(i => {
    if (i.type !== "file") i.value = "";
  });

  document.querySelectorAll("#descarga textarea").forEach(t => {
    t.value = "";
  });

  document.querySelectorAll("#descarga select").forEach(s => {
    s.selectedIndex = 0;
  });

  const archivoExcel = document.getElementById("archivoExcelDescarga");
  if (archivoExcel) archivoExcel.value = "";
}

function exportarPDFDescarga() {
  exportarElementoPDF("exportDescarga", "reporte_descarga.pdf");
}
