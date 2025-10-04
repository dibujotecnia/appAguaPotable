// ===============================
// CONFIGURACIÓN
// ===============================

// Enlace compartido de OneDrive con permisos de edición
const ENLACE_ONEDRIVE = "https://1drv.ms/x/c/3d7f55dc3c1f5283/EZ_UpfkgaHhFukgiZbatzlkBeKYVKpFAqXXHFmyRZqVmNA?e=Os17xb";

// Inputs del formulario
const inputCodigo = document.getElementById('codigo');
const inputNombre = document.getElementById('nombre');
const inputDireccion = document.getElementById('direccion');
const inputLecturaAnterior = document.getElementById('lectura-anterior');
const inputLecturaActual = document.getElementById('lectura-actual');

// Botones
const btnAbrirExcel = document.getElementById('btn-abrir-excel');
const btnImprimir = document.getElementById('btn-imprimir');
const btnQR = document.getElementById('btn-qr');

// Variables globales
let workbook;       // Archivo completo
let hojaLecturas;   // Hoja de lecturas visible
let hojaAviso;      // Hoja AVISO

// ===============================
// FUNCIONES AUXILIARES
// ===============================

// Convierte una hoja a JSON
function hojaAJSON(hoja) {
  return XLSX.utils.sheet_to_json(hoja, { defval: "" });
}

// Buscar socio por código en la hoja de lecturas
function buscarSocio(codigo) {
  if (!hojaLecturas) {
    alert('Primero abre el archivo Excel');
    return;
  }

  const datos = hojaAJSON(hojaLecturas);
  const socio = datos.find(row => row['CODIGO'] == codigo);

  if (!socio) {
    alert('Código no encontrado');
    inputNombre.value = '';
    inputDireccion.value = '';
    inputLecturaAnterior.value = '';
    return;
  }

  inputNombre.value = socio['NOMBRE Y APELLIDOS'] || '';
  inputDireccion.value = socio['UBICACION'] || '';
  inputLecturaAnterior.value = socio['LECT.ANT.'] || '';
}

// ===============================
// ABRIR EXCEL DESDE LINK COMPARTIDO
// ===============================

btnAbrirExcel.addEventListener('click', async () => {
  try {
    const response = await fetch(ENLACE_ONEDRIVE);
    const data = await response.arrayBuffer();
    workbook = XLSX.read(data, { type: "array" });

    // Detecta hojas automáticamente
    hojaAviso = workbook.Sheets['AVISO'];
    const hojasLecturas = workbook.SheetNames.filter(name => name.startsWith('LECTURAS'));
    if (hojasLecturas.length === 0) {
      alert('No se encontró hoja de lecturas');
      return;
    }
    hojaLecturas = workbook.Sheets[hojasLecturas[0]];

    alert('Archivo Excel cargado correctamente desde OneDrive');
  } catch (error) {
    console.error(error);
    alert('Error al abrir Excel desde OneDrive');
  }
});

// ===============================
// ESCANEAR QR CON CONFIRMACIÓN
// ===============================

btnQR.addEventListener('click', () => {
  const confirmar = confirm("¿Deseas iniciar el escaneo QR? Presiona Cancelar para abortar.");
  if (!confirmar) return;

  const html5QrCode = new Html5Qrcode("reader");
  html5QrCode.start(
    { facingMode: "environment" },
    { fps: 10, qrbox: 250 },
    qrCodeMessage => {
      inputCodigo.value = qrCodeMessage;
      buscarSocio(qrCodeMessage);
      html5QrCode.stop();
      document.getElementById('reader').innerHTML = "";
    },
    errorMessage => {
      console.log(errorMessage);
    }
  );
});

// ===============================
// GUARDAR LECTURA Y GENERAR AVISO
// ===============================

btnImprimir.addEventListener('click', async () => {
  const codigo = inputCodigo.value.trim();
  const lecturaActual = inputLecturaActual.value.trim();

  if (!codigo || !lecturaActual) {
    alert('Por favor ingresa el código y la lectura actual');
    return;
  }

  const datos = hojaAJSON(hojaLecturas);
  const filaIndex = datos.findIndex(row => row['CODIGO'] == codigo);
  if (filaIndex === -1) {
    alert('Código no encontrado');
    return;
  }

  // Actualiza lectura actual y fecha
  const fechaHoy = new Date().toLocaleDateString();
  const rangoLecturaActual = `F${filaIndex + 2}`; // LECT.ACT.
  const rangoFechaLectura = `G${filaIndex + 2}`;   // FECHA LECTURA
  hojaLecturas[rangoLecturaActual] = { t: 'n', v: Number(lecturaActual) };
  hojaLecturas[rangoFechaLectura] = { t: 's', v: fechaHoy };

  // Actualiza AVISO
  hojaAviso['D8'] = { t: 's', v: codigo };

  // Generar imagen del aviso
  const avisoDiv = document.getElementById('aviso-preview');
  if (avisoDiv) {
    avisoDiv.style.display = 'block';
    avisoDiv.innerHTML = `
      <p>Código: ${codigo}</p>
      <p>Nombre: ${inputNombre.value}</p>
      <p>Dirección: ${inputDireccion.value}</p>
      <p>Lectura Anterior: ${inputLecturaAnterior.value}</p>
      <p>Lectura Actual: ${lecturaActual}</p>
      <p>Fecha: ${fechaHoy}</p>
    `;

    html2canvas(avisoDiv).then(canvas => {
      const imgData = canvas.toDataURL('image/png');

      // Aquí se podría enviar a WhatsApp o impresora
      avisoDiv.style.display = 'none';
      alert('Lectura registrada y aviso generado');

      // Limpia campos
      inputCodigo.value = '';
      inputNombre.value = '';
      inputDireccion.value = '';
      inputLecturaAnterior.value = '';
      inputLecturaActual.value = '';
    });
  }

  // Guardar cambios al archivo XLSM en la nube
  const wbout = XLSX.write(workbook, { bookType: "xlsm", type: "array" });

  try {
    await fetch(ENLACE_ONEDRIVE, {
      method: "PUT",
      body: wbout,
      headers: {
        "Content-Type": "application/vnd.ms-excel"
      }
    });
    console.log("Archivo XLSM actualizado en OneDrive");
  } catch (err) {
    console.error("Error al actualizar XLSM en OneDrive", err);
  }
});

