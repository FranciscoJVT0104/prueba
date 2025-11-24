// =====================================================
// 1. MAPEO INTERNO CURSO_LARGO → CURSO_CORTO (TXT interno)
// =====================================================

const cursosTXT = `
COMPUTACIÓN E INFORMÁTICA;COMP-INF
CONTABILIDAD;CONT
ADMINISTRACIÓN DE EMPRESAS;ADM
ELECTRICIDAD INDUSTRIAL;ELEC
DISEÑO GRÁFICO;DG
MARKETING;MKT
CAJERO COMERCIAL;CAJ
PROGRAMACIÓN WEB;WEB
`;

// Convertir TXT a diccionario
const cursoMap = {};
cursosTXT.trim().split("\n").forEach(line => {
    const [largo, corto] = line.split(";");
    cursoMap[largo.trim().toUpperCase()] = corto.trim();
});

let originalData = []; // datos cargados del Excel

// =====================================================
// 2. CARGAR EXCEL
// =====================================================

document.getElementById("fileInput").addEventListener("change", function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: "binary" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        let json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

       originalData = json.map(row => {
    const cursoLargo = (row.CURSO || "").toUpperCase();
    const cursoCorto = cursoMap[cursoLargo] || "N/A";

    return {
        APELLIDOS: row.APELLIDOS || "",
        NOMBRES: row.NOMBRES || "",
        CURSO: row.CURSO || "",
        CURSO_CORTO: cursoCorto,
        DNI: row.DNI || "",
        CELULAR: row.CELULAR || "",
        FECHA: row.FECHA || "",
        INFORME: row["NUMERO DE INFORME"] || row.INFORME || "",
        OBSERVACION: row.OBSERVACION || row.OBS || ""
    };
});

        renderTable();
    };
    reader.readAsBinaryString(file);
});

// =====================================================
// 3. TABLA DE VISTA PREVIA
// =====================================================

function renderTable() {
    const table = document.getElementById("previewTable");
    table.innerHTML = "";

    if (originalData.length === 0) return;

    // Obtener TODOS los encabezados presentes en cualquier fila
    const headers = Array.from(
        originalData.reduce((set, row) => {
            Object.keys(row).forEach(k => set.add(k));
            return set;
        }, new Set())
    );

    // Crear encabezados
    let headerHTML = "<tr>" + headers.map(h => `<th>${h}</th>`).join("") + "</tr>";

    // Crear filas
    let rowsHTML = originalData
        .map(row =>
            "<tr>" +
            headers.map(h => `<td>${row[h] ?? ""}</td>`).join("") +
            "</tr>"
        )
        .join("");

    table.innerHTML = headerHTML + rowsHTML;
}

// =====================================================
// 4. EXPORTAR REV NOV (EXCEL)
// =====================================================

function exportRevNov() {
    if (originalData.length === 0) { alert("Sube un archivo primero"); return; }

    const output = originalData.map((row, index) => {
        const username = row.DNI;
        const password = row.DNI + row.APELLIDOS.charAt(0) + row.NOMBRES.charAt(0);
        const email = row.DNI + "s@actualizar.com";


        return {
            username,
            password,
            firstname: row.NOMBRES,
            lastname: row.APELLIDOS,
            email,
            city: "LIMA",
            course1: row.CURSO_CORTO,
            group1: row["NUMERO DE INFORME"] || row.INFORME || "",
            obs: row.OBSERVACION || row.OBS || ""
        };
    });

    const ws = XLSX.utils.json_to_sheet(output);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "REV NOV");

    XLSX.writeFile(wb, "REV_NOV.xlsx");
}

// =====================================================
// 5. EXPORTAR CONT NOV (CSV)
// =====================================================

function exportContNov() {
    if (originalData.length === 0) { alert("Sube un archivo primero"); return; }

    const output = originalData.map((row, index) => {

        return {
            Nombre: "",
            Apellido: `${row.APELLIDOS} ${row.NOMBRES} ${row["NUMERO DE INFORME"] || row.INFORME || ""}`,
            Telefono: row.CELULAR,
            "correo electronico": row.CELULAR + "s@actualizar.com",
            Direccion: "ESTANDAR",
            Cumpleaños: new Date().toLocaleDateString(),
            Observaciones: row["NUMERO DE INFORME"] || row.INFORME || ""
        };
    });

    const ws = XLSX.utils.json_to_sheet(output);
    const csv = XLSX.utils.sheet_to_csv(ws);

    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "CONT_NOV.csv";
    link.click();
}



