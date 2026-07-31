// ============================================================
// CONTROLADOR.JS
// Coordina el Modelo con la Vista y gestiona los eventos.
// ============================================================
(function (global) {
  "use strict";

  let requerimientos = [];
  let eventosRegistrados = false;
  let inicializacionIniciada = false;
  let observadorChromeSharePoint = null;
  let documentoObservadoSharePoint = null;
  let requerimientoEnEdicion = null;
  let origenEdicion = "mis";
  // Usuarios seleccionados para el campo Solicitado por
  let usuariosSeleccionados = [];
  let requerimientosPropios = [];
  let indicadoresEstados = [];
  let indicadoresResponsables = [];
  let graficaEstados = null;
  let graficaResponsables = null;
  let paginaMisRequerimientos = 1;
  let paginaBacklog = 1;
  let requerimientosGestion = [];
  let paginaGestion = 1;
  let actividadesHistorial = [];
  let paginaHistorial = 1;
  let dialogoMensajeActual = null;
  let focoAntesDelDialogo = null;
  const REQUERIMIENTOS_POR_PAGINA = 5;
  const ACTIVIDADES_POR_PAGINA = 8;
  const TIPOS_DOCUMENTO_PREDETERMINADOS = [
    "01_Historia_Usuario",
    "02_Fuentes",
    "03_Pruebas",
    "04_Documentaci\u00f3n"
  ];
  let tiposDocumento = TIPOS_DOCUMENTO_PREDETERMINADOS.slice();
  let archivosSeleccionados = [];
  const tiposDocumentoSeleccionados = new Map();
  let requerimientoDetalleActual = null;
  let archivosDetalleSeleccionados = [];
  const tiposDocumentoDetalleSeleccionados = new Map();

  const SELECTORES_CHROME_SHAREPOINT = [
    "#suiteBar",
    "#suiteBarTop",
    "#s4-ribbonrow",
    "#s4-titlerow",
    "#sideNavBox",
    "#pageStatusBar",
    "#DeltaPlaceHolderPageTitleInTitleArea",
    ".ms-webpart-chrome-title",
    ".ms-webpart-titleText"
  ];

  const SELECTORES_ANCHO_SHAREPOINT = [
    "#s4-bodyContainer",
    "#contentRow",
    "#contentBox",
    "#DeltaPlaceHolderMain",
    ".ms-webpart-zone",
    ".ms-webpart-cell-vertical",
    ".ms-webpart-chrome-vertical",
    ".ms-WPBody"
  ];

  function obtenerDocumentoSharePoint() {
    try {
      if (
        global.parent &&
        global.parent !== global &&
        global.parent.document &&
        (
          global.parent.document.getElementById("s4-workspace") ||
          global.parent.document.getElementById("s4-ribbonrow") ||
          global.parent.document.getElementById("contentBox")
        )
      ) {
        return global.parent.document;
      }
    } catch (error) {
      console.warn(
        "La p\u00e1gina contenedora de SharePoint no permite acceso desde el iframe.",
        error
      );
    }

    return document;
  }

  function parametrosVentana(ventana) {
    try {
      return new URLSearchParams(ventana.location.search);
    } catch (error) {
      return new URLSearchParams("");
    }
  }

  function estaEnModoEdicionSharePoint(documentoSharePoint) {
    const ventanaSharePoint =
      documentoSharePoint.defaultView || global;
    const ventanas = [global];
    if (ventanaSharePoint !== global) {
      ventanas.push(ventanaSharePoint);
    }

    const modos = [];
    ventanas.forEach(function (ventana) {
      const parametros = parametrosVentana(ventana);
      ["Mode", "ControlMode", "DisplayMode"].forEach(function (nombre) {
        const valor = parametros.get(nombre);
        if (valor) {
          modos.push(valor.toLowerCase());
        }
      });
    });

    const campoDiseno =
      documentoSharePoint.getElementById("MSOLayout_InDesignMode");
    const valorDisenoGlobal = ventanaSharePoint.MSOLayout_InDesignMode;
    const disenoClasico =
      valorDisenoGlobal === true ||
      valorDisenoGlobal === 1 ||
      valorDisenoGlobal === "1" ||
      Boolean(campoDiseno && campoDiseno.value === "1");

    return (
      modos.indexOf("edit") !== -1 ||
      modos.indexOf("design") !== -1 ||
      disenoClasico
    );
  }

  function ocultarElemento(elemento) {
    elemento.style.setProperty("display", "none", "important");
  }

  function aplicarAjustesSharePoint() {
    const documentoSharePoint = obtenerDocumentoSharePoint();
    if (estaEnModoEdicionSharePoint(documentoSharePoint)) {
      return false;
    }

    documentoSharePoint.documentElement.classList.add(
      "bitacora-sharepoint-fullscreen"
    );

    SELECTORES_CHROME_SHAREPOINT.forEach(function (selector) {
      documentoSharePoint
        .querySelectorAll(selector)
        .forEach(ocultarElemento);
    });

    const contentBox = documentoSharePoint.getElementById("contentBox");
    if (contentBox) {
      contentBox.style.setProperty("min-width", "0", "important");
      contentBox.style.setProperty("margin-left", "0", "important");
    }

    const contentRow = documentoSharePoint.getElementById("contentRow");
    if (contentRow) {
      contentRow.style.setProperty("padding", "0", "important");
    }

    SELECTORES_ANCHO_SHAREPOINT.forEach(function (selector) {
      documentoSharePoint
        .querySelectorAll(selector)
        .forEach(function (elemento) {
          elemento.style.setProperty("width", "100%", "important");
          elemento.style.setProperty("max-width", "none", "important");
          elemento.style.setProperty("margin", "0", "important");
          elemento.style.setProperty("padding", "0", "important");
        });
    });

    const workspace =
      documentoSharePoint.getElementById("s4-workspace");
    if (workspace) {
      workspace.style.setProperty("inset", "0", "important");
      workspace.style.setProperty("width", "100%", "important");
      workspace.style.setProperty("height", "100%", "important");
      workspace.style.setProperty("margin", "0", "important");
      workspace.style.setProperty("padding", "0", "important");
      workspace.style.setProperty("overflow-x", "hidden", "important");
      workspace.style.setProperty("overflow-y", "auto", "important");
    }

    return true;
  }

  function observarCambiosSharePoint() {
    const documentoSharePoint = obtenerDocumentoSharePoint();
    const ventanaSharePoint =
      documentoSharePoint.defaultView || global;
    const Observador =
      ventanaSharePoint.MutationObserver || global.MutationObserver;

    if (
      observadorChromeSharePoint &&
      documentoObservadoSharePoint === documentoSharePoint
    ) {
      return;
    }

    if (observadorChromeSharePoint) {
      observadorChromeSharePoint.disconnect();
      observadorChromeSharePoint = null;
      documentoObservadoSharePoint = null;
    }

    if (
      !Observador ||
      estaEnModoEdicionSharePoint(documentoSharePoint)
    ) {
      return;
    }

    observadorChromeSharePoint = new Observador(function (cambios) {
      const agregoNodos = cambios.some(function (cambio) {
        return cambio.addedNodes && cambio.addedNodes.length > 0;
      });
      if (agregoNodos) {
        aplicarAjustesSharePoint();
      }
    });
    observadorChromeSharePoint.observe(
      documentoSharePoint.documentElement,
      {
        childList: true,
        subtree: true
      }
    );
    documentoObservadoSharePoint = documentoSharePoint;
  }

  function configurarVistaSharePoint() {
    const documentoSharePoint = obtenerDocumentoSharePoint();
    const debeOcultarChrome =
      !estaEnModoEdicionSharePoint(documentoSharePoint);

    documentoSharePoint.documentElement.classList.toggle(
      "bitacora-sharepoint-fullscreen",
      debeOcultarChrome
    );

    if (!debeOcultarChrome) {
      return;
    }

    aplicarAjustesSharePoint();
    observarCambiosSharePoint();
    global.setTimeout(aplicarAjustesSharePoint, 250);
    global.setTimeout(aplicarAjustesSharePoint, 1000);
  }

  async function iniciar() {
    if (inicializacionIniciada) {
      return;
    }
    inicializacionIniciada = true;
    configurarVistaSharePoint();
    registrarEventos();
    Vista.mostrarUsuario(Modelo.usuarioActual());
    const conectado = await comprobarConexion();
    if (conectado) {
      await cargarTiposDocumento();
      await cargarDashboard();
    } else {
      Vista.renderizarTarjetas([]);
      Vista.renderizarTabla([]);
    }
  }

  async function cargarTiposDocumento() {
    try {
      const opciones = await Modelo.obtenerTiposDocumento();
      if (opciones.length) {
        tiposDocumento = opciones;
      }
      global.BitacoraDiagnosticoTiposDocumento = {
        origen: opciones.length ? "SharePoint" : "Valores predeterminados",
        opciones: tiposDocumento.slice()
      };
    } catch (error) {
      tiposDocumento = TIPOS_DOCUMENTO_PREDETERMINADOS.slice();
      global.BitacoraDiagnosticoTiposDocumento = {
        origen: "Valores predeterminados",
        opciones: tiposDocumento.slice(),
        error: error.message
      };
      console.warn("Consulta de tipos de documento:", error);
    }
  }

  async function registrarActividadCompartida(accion, detalles) {
    try {
      await Modelo.agregarActividad(
        Modelo.usuarioActual().nombre,
        accion,
        detalles
      );
    } catch (error) {
      global.BitacoraDiagnosticoHistorial = {
        accion: accion,
        error: error.message,
        fecha: new Date().toISOString()
      };
      console.warn(
        "La operaci\u00f3n se complet\u00f3, pero no se pudo registrar en el historial:",
        error
      );
    }
  }

  async function comprobarConexion() {
    try {
      const diagnostico = await Modelo.verificarConexion();
      global.BitacoraDiagnosticoSharePoint = diagnostico;
      console.info("Diagnostico SharePoint Backlog:", diagnostico);
      console.table(diagnostico.campos);
      if (diagnostico.advertencias.length > 0) {
        console.warn(
          "Advertencias del esquema Backlog:",
          diagnostico.advertencias
        );
      }
      Vista.mostrarConexion(
        true,
        "Conectado a SharePoint: " +
          (diagnostico.sitio.Title || diagnostico.sitio.Url) +
          " / " +
          diagnostico.lista.Title +
          " (" +
          diagnostico.lista.ItemCount +
          " registros) / " +
          diagnostico.biblioteca.Title +
          " (" +
          diagnostico.biblioteca.ItemCount +
          " elementos)" +
          (diagnostico.advertencias.length > 0
            ? " - Esquema pendiente de revisar"
            : "")
      );
      return true;
    } catch (error) {
      global.BitacoraDiagnosticoSharePoint = {
        conectado: false,
        error: error.message,
        urlLista:
          "https://globalhitss.sharepoint.com/sites/AppsColombiaDesarrollo/" +
          "RetirosDeCesantiasDesarrollo/Lists/Backlog/AllItems.aspx",
        urlBiblioteca:
          "https://globalhitss.sharepoint.com/sites/AppsColombiaDesarrollo/" +
          "RetirosDeCesantiasDesarrollo/ArchivosRequerimientos/Forms/AllItems.aspx"
      };
      console.error("Validaci\u00f3n de SharePoint:", error);
      Vista.mostrarConexion(false, "Sin conexi\u00f3n: " + error.message);
      return false;
    }
  }

  async function cargarDashboard() {
    try {

      requerimientos = (await Modelo.obtenerTodos()).reverse();
      Vista.renderizarFiltros(requerimientos);
      paginaBacklog = 1;
      renderizarPaginaBacklog();

    } catch (error) {
      console.error("Carga del dashboard:", error);
      Vista.renderizarTarjetas([]);
      Vista.renderizarTabla([]);
      await mostrarAlerta({
        icon: "error",
        title: "No se pudieron cargar los requerimientos",
        text: "SharePoint devolvi\u00f3 el siguiente error: " + error.message
      });
    }
  }
  // ============================================================
// PROCESAR DATOS DE INDICADORES
// ============================================================

async function cargarIndicadores(){

    let datos = 
    await Modelo.obtenerDatosIndicadores();

    console.log("DATOS FILTRO FECHA:", datos);

const fechaInicio =
document.getElementById("fechaInicioIndicadores")?.value;


const fechaFin =
document.getElementById("fechaFinIndicadores")?.value;

console.log("FECHA INICIO INPUT:", fechaInicio);
console.log("FECHA FIN INPUT:", fechaFin);



if(fechaInicio && fechaFin){

    const inicio = new Date(fechaInicio);
    inicio.setHours(0,0,0,0);


    const fin = new Date(fechaFin);
    fin.setHours(23,59,59,999);



 datos = datos.filter(function(req){

    const fechaTexto = req.fechaSolicitud;

    console.log(
        "FECHA REQUERIMIENTO:",
        fechaTexto
    );


    if(!fechaTexto){
        return false;
    }


    const fecha = new Date(fechaTexto);


    if(isNaN(fecha)){
        console.log(
            "Fecha inválida:",
            fechaTexto
        );

        return false;
    }


    fecha.setHours(0,0,0,0);


    return (
        fecha >= inicio &&
        fecha <= fin
    );

});
console.log(
    "DATOS DESPUES DEL FILTRO:",
    datos
);
}
   // =======================================
   // ESTADOS FIJOS DEL FLUJO DE TRABAJO
    // =======================================

// =======================================
// ESTADOS FIJOS DEL FLUJO DE TRABAJO
// =======================================

const estadosFlujo = [
    "Pendiente",
    "Pruebas",
    "E.Cierre ",
    "E.Documentos",
    "Finalizado"
];


// Contar estados reales desde SharePoint
const conteoEstados = {};

datos.forEach(function(req){

    const estado = req.estado || "Sin estado";


    if(!conteoEstados[estado]){

        conteoEstados[estado] = 0;

    }


    conteoEstados[estado]++;

});


// Crear estructura fija para la gráfica

indicadoresEstados = estadosFlujo.map(function(estado){

    return {

        nombre: estado,

        cantidad: conteoEstados[estado] || 0

    };

});
    // ==============================
    // RESPONSABLES
    // ==============================

    const responsables = {};
    datos.forEach(function(req){

        const responsable =
            req.responsable || "No asignado";


        if(!responsables[responsable]){

            responsables[responsable] = 0;

        }
        responsables[responsable]++;

    });

    indicadoresResponsables =
        Object.keys(responsables).map(function(nombre){

            return {

                nombre:nombre,

                cantidad:responsables[nombre]

            };

        });
// =====================================
// ACTUALIZAR TARJETAS KPI
// =====================================

const total =
    datos.length;


let pendientes = 0;
let enPruebas = 0;

let finalizados = 0;


datos.forEach(function(req){

    if(req.estado === "Pendiente"){

        pendientes++;

    }


    if(req.estado === "Finalizado"){

    finalizados++;

}

});



const kpiTotal =
    document.getElementById("kpi-total");


const kpiPendientes =
    document.getElementById("kpi-pendientes");


const kpiFinalizados =
    document.getElementById("kpi-finalizados");



if(kpiTotal){

    kpiTotal.textContent = total;

}


if(kpiPendientes){

    kpiPendientes.textContent = pendientes;

}

if(kpiFinalizados){

    kpiFinalizados.textContent = finalizados;

}



    console.log(
        "Estados indicadores:",
        indicadoresEstados
    );


    console.log(
        "Responsables indicadores:",
        indicadoresResponsables
    );
renderizarGraficasIndicadores();

renderizarResumenIndicadores();
}

// ============================================================
// CREAR GRAFICAS DE INDICADORES
// ============================================================

function renderizarGraficasIndicadores(){



    if(graficaEstados){
        graficaEstados.destroy();
    }


    if(graficaResponsables){
        graficaResponsables.destroy();
    }


   
    const canvasEstados =
        document.getElementById("graficoEstados");


    const canvasResponsables =
        document.getElementById("graficoResponsables");

// destruir gráficas anteriores antes de crear nuevas

if(graficaEstados){

    graficaEstados.destroy();

}


if(graficaResponsables){

    graficaResponsables.destroy();

}


    if(!canvasEstados || !canvasResponsables){

        console.warn(
            "No existen los canvas de indicadores"
        );

        return;

    }



    // ===============================
    // GRAFICA DE ESTADOS
    // ===============================


    graficaEstados = new Chart(canvasEstados, {
        type: "bar",
        data: {
            labels: indicadoresEstados.map(
                function(item){
                    return item.nombre;
                }
            ),


            datasets:[{

                label:"Cantidad de requerimientos",


                data: indicadoresEstados.map(
                    function(item){
                        return item.cantidad;
                    }),

                borderWidth:0,
                borderRadius:8
            }]
        },
      options:{
    responsive:true,

    maintainAspectRatio:false,

    plugins:{
        legend:{
            display:false
        }
    },

    scales:{
        x:{
        ticks:{
            autoSkip:false,
            maxRotation:45,
            minRotation:45,
            font:{
                size:10
            }
        }
    },

        y:{
            beginAtZero:true,
            ticks:{
                precision:0
            }
        }
    }
}
    });


    // ===============================
    // GRAFICA RESPONSABLES
    // ===============================

    graficaResponsables = new Chart(canvasResponsables, {
        type:"doughnut",
        data:{
            labels:
            indicadoresResponsables.map(
                function(item){
                    return item.nombre;

                }
            ),

            datasets:[{
                data:
                indicadoresResponsables.map(
                    function(item){
                        return item.cantidad;

                    }
                )
            }]

        },
        options:{
            responsive:true,
            plugins:{
                legend:{
                    position:"bottom"
                }
            }
        }
    });



}
   function renderizarResumenIndicadores(){

    const tabla =
        document.getElementById(
            "tabla-resumen-indicadores"
        );


    if(!tabla){

        return;

    }


    const total =
        indicadoresEstados.reduce(
            function(total,item){

                return total + item.cantidad;

            },
            0
        );


    tabla.innerHTML =
        indicadoresEstados.map(function(item){


            const porcentaje =
                total > 0
                ? ((item.cantidad / total) * 100).toFixed(1)
                : 0;



            return `

            <tr>

                <td>${item.nombre}</td>

                <td>${item.cantidad}</td>

                <td>${porcentaje}%</td>


            </tr>

            `;


        }).join("");

}

  function coincideFecha(valorFecha, fechaFiltro) {
    if (!fechaFiltro) {
      return true;
    }
    if (!valorFecha) {
      return false;
    }

    const textoFecha = String(valorFecha).trim();
    const fechaISO = textoFecha.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (fechaISO) {
      return (
        fechaISO[1] + "-" + fechaISO[2] + "-" + fechaISO[3] ===
        fechaFiltro
      );
    }

    const fechaLatina = textoFecha.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (fechaLatina) {
      return (
        fechaLatina[3] + "-" + fechaLatina[2] + "-" + fechaLatina[1] ===
        fechaFiltro
      );
    }

    const fecha = new Date(textoFecha);
    if (isNaN(fecha.getTime())) {
      return false;
    }
    const fechaLocal =
      fecha.getFullYear() +
      "-" +
      String(fecha.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(fecha.getDate()).padStart(2, "0");
    return fechaLocal === fechaFiltro;
  }

  function filtrarRequerimientos() {
    const texto = document.getElementById("buscador").value.toLowerCase();
    const filtros = {
      app: document.getElementById("filtro-app").value,
      responsable: document.getElementById("filtro-responsable").value,
      estado: document.getElementById("filtro-estado").value,
      prioridad: document.getElementById("filtro-prioridad").value,
      fecha: document.getElementById("filtro-fecha").value
    };

    return requerimientos.filter(function (req) {
      const contenido = [
        req.id,
        req.app,
        req.tipoServicio,
        req.asunto,
        req.solicitadoPor,
        req.responsable
      ]
        .join(" ")
        .toLowerCase();
      return (
        contenido.indexOf(texto) !== -1 &&
        (!filtros.app || req.app === filtros.app) &&
        (!filtros.responsable ||
          req.responsable === filtros.responsable) &&
        (!filtros.estado || req.estado === filtros.estado) &&
        (!filtros.prioridad ||
          String(req.prioridad) === filtros.prioridad) &&
        coincideFecha(req.fechaSolicitud, filtros.fecha)
      );
    });
  }

  function renderizarPaginaBacklog() {
    const datosFiltrados = filtrarRequerimientos();
    Vista.renderizarTarjetas(datosFiltrados);
    const total = datosFiltrados.length;
    const totalPaginas = Math.max(
      1,
      Math.ceil(total / REQUERIMIENTOS_POR_PAGINA)
    );
    paginaBacklog = Math.min(Math.max(1, paginaBacklog), totalPaginas);
    const indiceInicial = (paginaBacklog - 1) * REQUERIMIENTOS_POR_PAGINA;
    const datosPagina = datosFiltrados.slice(
      indiceInicial,
      indiceInicial + REQUERIMIENTOS_POR_PAGINA
    );

    Vista.renderizarTabla(datosPagina, {
      pagina: paginaBacklog,
      totalPaginas: totalPaginas,
      total: total,
      inicio: total ? indiceInicial + 1 : 0,
      fin: Math.min(indiceInicial + REQUERIMIENTOS_POR_PAGINA, total)
    });
  }

  function cambiarPaginaBacklog(pagina) {
    const destino = Number(pagina);
    const totalPaginas = Math.max(
      1,
      Math.ceil(filtrarRequerimientos().length / REQUERIMIENTOS_POR_PAGINA)
    );
    if (
      !Number.isInteger(destino) ||
      destino < 1 ||
      destino > totalPaginas ||
      destino === paginaBacklog
    ) {
      return;
    }
    paginaBacklog = destino;
    renderizarPaginaBacklog();
  }

  function aplicarFiltros() {
    paginaBacklog = 1;
    renderizarPaginaBacklog();
  }

  function escaparCSV(valor) {
    const texto = String(valor == null ? "" : valor);
    if (/[";\r\n]/.test(texto)) {
      return '"' + texto.replace(/"/g, '""') + '"';
    }
    return texto;
  }

  async function exportarCSV() {
    const datos = filtrarRequerimientos();
    if (!datos.length) {
      await mostrarAlerta({
        icon: "info",
        title: "Sin datos para exportar",
        text: "No hay requerimientos que coincidan con los filtros actuales."
      });
      return;
    }

    const confirmacion = await mostrarDialogo({
      icon: "info",
      title: "Exportar CSV",
      text:
        "Se descargar\u00e1 un archivo CSV con " +
        datos.length +
        (datos.length === 1 ? " requerimiento" : " requerimientos") +
        " seg\u00fan los filtros actuales.",
      showCancelButton: true,
      confirmButtonText: "Descargar",
      cancelButtonText: "Cancelar"
    });
    if (!confirmacion.isConfirmed) {
      return;
    }

    const columnas = [
      ["id", "ID REQ."],
      ["app", "APP"],
      ["tipoServicio", "Tipo de servicio"],
      ["asunto", "Asunto"],
      ["solicitadoPor", "Solicitado por"],
      ["responsable", "Responsable"],
      ["prioridad", "Prioridad"],
      ["estado", "Estado"],
      ["fechaSolicitud", "F. solicitud"],
      ["fechaCierre", "F. cierre"]
    ];
    const filas = [columnas.map(function (columna) {
      return columna[1];
    })];

    datos.forEach(function (req) {
      filas.push(columnas.map(function (columna) {
        const valor = req[columna[0]];
        return columna[0] === "fechaSolicitud" || columna[0] === "fechaCierre"
          ? Vista.formatearFecha(valor)
          : valor;
      }));
    });

    const csv = filas
      .map(function (fila) {
        return fila.map(escaparCSV).join(";");
      })
      .join("\r\n");
    const blob = new Blob(["\ufeff" + csv], {
      type: "text/csv;charset=utf-8;"
    });
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download =
      "backlog_" + new Date().toISOString().slice(0, 10) + ".csv";
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
    URL.revokeObjectURL(url);
  }

  function limpiarFiltros() {
    document.getElementById("buscador").value = "";
    ["filtro-app", "filtro-responsable", "filtro-estado", "filtro-prioridad", "filtro-fecha"]
      .forEach(function (id) {
        document.getElementById(id).value = "";
      });
    paginaBacklog = 1;
    renderizarPaginaBacklog();
  }

  async function verRequerimiento(id) {
    try {
      const req = await Modelo.obtenerPorId(id);
      if (!req) {
        await mostrarAlerta({
          icon: "warning",
          title: "Requerimiento no encontrado",
          text: "No fue posible encontrar el requerimiento seleccionado."
        });
        return;
      }
      req.archivosAdjuntos =
        await Modelo.obtenerDocumentosRequerimiento(req);
      requerimientoDetalleActual = req;
      archivosDetalleSeleccionados = [];
      tiposDocumentoDetalleSeleccionados.clear();
      Vista.mostrarDetalle(req);
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudo consultar el detalle",
        text: error.message
      });
    }
  }

  function fechaActual() {
    return new Date().toLocaleDateString("es-CO", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric"
    });
  }

  function cerrarDialogoMensaje(confirmado) {
    if (!dialogoMensajeActual) {
      return;
    }

    const entrada = document.getElementById("entrada-modal-mensaje");
    const validacion = document.getElementById("validacion-modal-mensaje");
    const valor = entrada.value.trim();

    if (confirmado && dialogoMensajeActual.solicitaValor && !valor) {
      validacion.textContent = "Debe ingresar un nombre.";
      validacion.hidden = false;
      entrada.focus();
      return;
    }

    const resolver = dialogoMensajeActual.resolver;
    document.getElementById("modal-mensaje").hidden = true;
    dialogoMensajeActual = null;
    if (focoAntesDelDialogo && typeof focoAntesDelDialogo.focus === "function") {
      focoAntesDelDialogo.focus();
    }
    focoAntesDelDialogo = null;
    resolver({
      isConfirmed: confirmado,
      value: confirmado ? valor : ""
    });
  }

  function mostrarDialogo(opciones) {
    const overlay = document.getElementById("modal-mensaje");
    const icono = document.getElementById("icono-modal-mensaje");
    const entrada = document.getElementById("entrada-modal-mensaje");
    const campo = document.getElementById("campo-modal-mensaje");
    const validacion = document.getElementById("validacion-modal-mensaje");
    const cancelar = document.getElementById("cancelar-modal-mensaje");
    const confirmar = document.getElementById("confirmar-modal-mensaje");
    const tipo = ["success", "error", "warning", "info"]
      .indexOf(opciones.icon) !== -1
        ? opciones.icon
        : "info";
    const simbolos = {
      success: "\u2713",
      error: "\u00d7",
      warning: "!",
      info: "i"
    };

    focoAntesDelDialogo = document.activeElement;
    icono.className = "mensaje-icono " + tipo;
    icono.textContent = simbolos[tipo];
    document.getElementById("titulo-modal-mensaje").textContent =
      opciones.title || "Mensaje";
    document.getElementById("texto-modal-mensaje").textContent =
      opciones.text || "";
    campo.hidden = !opciones.input;
    entrada.value = opciones.inputValue || "";
    entrada.placeholder = opciones.inputPlaceholder || "";
    validacion.hidden = true;
    validacion.textContent = "";
    cancelar.hidden = !opciones.showCancelButton;
    cancelar.textContent = opciones.cancelButtonText || "Cancelar";
    confirmar.textContent = opciones.confirmButtonText || "Aceptar";
    overlay.hidden = false;

    return new Promise(function (resolver) {
      dialogoMensajeActual = {
        resolver: resolver,
        solicitaValor: Boolean(opciones.input)
      };
      global.setTimeout(function () {
        if (opciones.input) {
          entrada.focus();
        } else {
          confirmar.focus();
        }
      }, 0);
    });
  }

  function mostrarAlerta(opciones) {
    return mostrarDialogo(opciones).then(function () {
      return undefined;
    });
  }

  async function prepararFormulario() {
    requerimientoEnEdicion = null;
    document.getElementById("modo-formulario").textContent = "NUEVA SOLICITUD";
    document.getElementById("titulo-crear").textContent = "Registrar requerimiento";
    document.getElementById("subtitulo-formulario").textContent =
      "Los campos marcados con * son obligatorios.";
    document.getElementById("btn-guardar").textContent = "Guardar requerimiento";
    document.getElementById("btn-limpiar").hidden = false;

    [
      "id",
      "app",
      "tipoServicio",
      "asunto",
      "descripcion",
      "casoOrigen",
      "solicitadoPor",
      "estado",
      "fechaSolicitud"
    ].forEach(function (id) {
      const campo = document.getElementById(id);
      if (campo) {
        campo.value = "";
      }
    });
    document.getElementById("solicitadoPor").value =
      Modelo.usuarioActual().nombre;

    document.getElementById("estado").value = "Pendiente";
    document.getElementById("fechaSolicitud").value = fechaActual();
    limpiarClasificacionArchivos();
    usuariosSeleccionados = [];
    cargarSolicitanteActual();

    try {
      const datos = await Modelo.obtenerTodos();
      const hoy = new Date();
      const prefijo =
        String(hoy.getFullYear()).slice(-2) +
        String(hoy.getMonth() + 1).padStart(2, "0") +
        String(hoy.getDate()).padStart(2, "0");
      const cantidad = datos.filter(function (req) {
        return String(req.id).indexOf(prefijo) === 0;
      }).length;
      document.getElementById("id").value =
        prefijo + String(cantidad + 1).padStart(2, "0");
    } catch (error) {
      console.error("Generaci\u00f3n de ID:", error);
      document.getElementById("id").value = "";
    }
  }

  function esRequerimientoPropio(req) {
    const usuario = Modelo.usuarioActual();
    const solicitantes = String(req.solicitadoPor || "")
      .split(";")
      .map(function (valor) {
        return valor.trim().toLowerCase();
      })
      .filter(Boolean);
    return [usuario.nombre, usuario.correo]
      .filter(Boolean)
      .some(function (valor) {
        return solicitantes.indexOf(String(valor).trim().toLowerCase()) !== -1;
      });
  }

  async function editarRequerimiento(id, origen) {
    const origenFinal = origen === "dashboard" ? "dashboard" : "mis";
    try {
      const req = await Modelo.obtenerPorId(id);
      if (!req) {
        await mostrarAlerta({
          icon: "warning",
          title: "Requerimiento no encontrado",
          text: "No fue posible encontrar el requerimiento seleccionado."
        });
        return;
      }
      if (origenFinal !== "dashboard" && !esRequerimientoPropio(req)) {
        await mostrarAlerta({
          icon: "warning",
          title: "Edici\u00f3n no permitida",
          text: "Solo puede editar requerimientos solicitados por usted."
        });
        return;
      }

      origenEdicion = origenFinal;
      requerimientoEnEdicion = req;
      document.getElementById("modo-formulario").textContent =
        "EDICI\u00d3N DE SOLICITUD";
      document.getElementById("titulo-crear").textContent =
        "Editar requerimiento";
      document.getElementById("subtitulo-formulario").textContent =
        "Actualiza la informaci\u00f3n de tu solicitud.";
      document.getElementById("btn-guardar").textContent = "Guardar cambios";
      document.getElementById("btn-limpiar").hidden = true;
      limpiarClasificacionArchivos();
      document.getElementById("id").value = req.id || "";
      document.getElementById("app").value = req.app || "";
      document.getElementById("tipoServicio").value = req.tipoServicio || "";
      document.getElementById("casoOrigen").value = req.casoOrigen || "";
      document.getElementById("asunto").value = req.asunto || "";
      document.getElementById("descripcion").value = req.descripcion || "";
      document.getElementById("solicitadoPor").value = req.solicitadoPor || "";
      document.getElementById("estado").value = req.estado || "";
      document.getElementById("fechaSolicitud").value =
        Vista.formatearFecha(req.fechaSolicitud);

      Vista.mostrarFormularioEdicion();
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudo abrir el requerimiento",
        text: error.message
      });
    }
  }

  function datosFormulario() {
    return {
      id: document.getElementById("id").value,
      app: document.getElementById("app").value,
      tipoServicio: document.getElementById("tipoServicio").value,
      asunto: document.getElementById("asunto").value.trim(),
      descripcion: document.getElementById("descripcion").value.trim(),
      casoOrigen: document.getElementById("casoOrigen").value.trim(),
      solicitadoPor: document.getElementById("solicitadoPor").value.trim(),
      responsable: "No asignado",
      estado: document.getElementById("estado").value,
      fechaSolicitud: document.getElementById("fechaSolicitud").value
    };
  }

  function archivosFormulario() {
    return archivosSeleccionados.slice();
  }

  function claveArchivo(archivo) {
    return [
      archivo.name,
      archivo.size,
      archivo.lastModified || 0
    ].join("::");
  }

  function limpiarClasificacionArchivos() {
    archivosSeleccionados = [];
    tiposDocumentoSeleccionados.clear();
    const campo = document.getElementById("archivosAdjuntos");
    if (campo) {
      campo.value = "";
    }
    const contenedor = document.getElementById("clasificacionArchivos");
    contenedor.replaceChildren();
    contenedor.hidden = true;
    actualizarResumenArchivos();
  }

  function actualizarResumenArchivos() {
    const resumen = document.getElementById("resumenArchivos");
    if (!resumen) {
      return;
    }
    const cantidad = archivosSeleccionados.length;
    resumen.textContent = cantidad
      ? cantidad +
        (cantidad === 1
          ? " archivo seleccionado. Puedes elegir m\u00e1s archivos."
          : " archivos seleccionados. Puedes elegir m\u00e1s archivos.")
      : "No hay archivos agregados.";
  }

  function crearSelectorTipoDocumento(archivo, indice) {
    const selector = document.createElement("select");
    selector.id = "tipoDocumentoArchivo-" + indice;
    selector.dataset.indiceArchivo = String(indice);
    selector.setAttribute("aria-label", "Tipo de documento");
    selector.setAttribute("aria-required", "true");
    selector.required = true;

    const opcionInicial = document.createElement("option");
    opcionInicial.value = "";
    opcionInicial.textContent = "Seleccione el tipo";
    selector.appendChild(opcionInicial);

    tiposDocumento.forEach(function (tipo) {
      const opcion = document.createElement("option");
      opcion.value = tipo;
      opcion.textContent = tipo;
      selector.appendChild(opcion);
    });
    selector.value =
      tiposDocumentoSeleccionados.get(claveArchivo(archivo)) || "";
    selector.addEventListener("change", function () {
      tiposDocumentoSeleccionados.set(
        claveArchivo(archivo),
        selector.value.trim()
      );
      selector.removeAttribute("aria-invalid");
    });
    return selector;
  }

  function renderizarClasificacionArchivos() {
    const archivos = archivosFormulario();
    const contenedor = document.getElementById("clasificacionArchivos");
    contenedor.replaceChildren();
    contenedor.hidden = !archivos.length;
    actualizarResumenArchivos();

    archivos.forEach(function (archivo, indice) {
      const fila = document.createElement("div");
      fila.className = "file-category-row";

      const nombre = document.createElement("span");
      nombre.className = "file-category-name";
      nombre.textContent = archivo.name;
      nombre.title = archivo.name;

      const botonEliminar = document.createElement("button");
      botonEliminar.type = "button";
      botonEliminar.className = "file-category-remove";
      botonEliminar.textContent = "\u00d7";
      botonEliminar.title = "Quitar archivo";
      botonEliminar.setAttribute(
        "aria-label",
        "Quitar el archivo " + archivo.name
      );
      botonEliminar.addEventListener("click", function () {
        tiposDocumentoSeleccionados.delete(claveArchivo(archivo));
        archivosSeleccionados.splice(indice, 1);
        renderizarClasificacionArchivos();
      });

      fila.appendChild(nombre);
      fila.appendChild(crearSelectorTipoDocumento(archivo, indice));
      fila.appendChild(botonEliminar);
      contenedor.appendChild(fila);
    });
  }

  async function agregarArchivosSeleccionados(evento) {
    const nuevosArchivos = Array.from(evento.target.files || []);
    evento.target.value = "";
    if (!nuevosArchivos.length) {
      return;
    }

    const archivosPorClave = new Map(
      archivosSeleccionados.map(function (archivo) {
        return [claveArchivo(archivo), archivo];
      })
    );
    nuevosArchivos.forEach(function (archivo) {
      archivosPorClave.set(claveArchivo(archivo), archivo);
    });

    if (
      archivosPorClave.size >
      Modelo.configuracion.cantidadMaximaArchivos
    ) {
      await mostrarAlerta({
        icon: "warning",
        title: "Demasiados archivos",
        text:
          "Puede adjuntar m\u00e1ximo " +
          Modelo.configuracion.cantidadMaximaArchivos +
          " archivos por operaci\u00f3n."
      });
      return;
    }

    archivosSeleccionados = Array.from(archivosPorClave.values());
    renderizarClasificacionArchivos();
  }

  function archivosClasificadosFormulario() {
    const archivos = archivosFormulario();
    return archivos.map(function (archivo) {
      return {
        archivo: archivo,
        tipoDocumento:
          tiposDocumentoSeleccionados.get(claveArchivo(archivo)) || ""
      };
    });
  }

  function actualizarResumenArchivosDetalle() {
    const resumen = document.getElementById("detalle-resumen-archivos");
    const boton = document.getElementById("detalle-subir-archivos");
    const cantidad = archivosDetalleSeleccionados.length;
    if (resumen) {
      resumen.textContent = cantidad
        ? cantidad +
          (cantidad === 1
            ? " archivo listo para clasificar y subir."
            : " archivos listos para clasificar y subir.")
        : "No hay archivos agregados.";
    }
    if (boton) {
      boton.disabled = cantidad === 0;
    }
  }

  function crearSelectorTipoDocumentoDetalle(archivo, indice) {
    const selector = document.createElement("select");
    selector.id = "tipoDocumentoDetalle-" + indice;
    selector.setAttribute("aria-label", "Tipo de documento");
    selector.setAttribute("aria-required", "true");
    selector.required = true;

    const opcionInicial = document.createElement("option");
    opcionInicial.value = "";
    opcionInicial.textContent = "Seleccione el tipo";
    selector.appendChild(opcionInicial);

    tiposDocumento.forEach(function (tipo) {
      const opcion = document.createElement("option");
      opcion.value = tipo;
      opcion.textContent = tipo;
      selector.appendChild(opcion);
    });
    selector.value =
      tiposDocumentoDetalleSeleccionados.get(claveArchivo(archivo)) || "";
    selector.addEventListener("change", function () {
      tiposDocumentoDetalleSeleccionados.set(
        claveArchivo(archivo),
        selector.value.trim()
      );
      selector.removeAttribute("aria-invalid");
    });
    return selector;
  }

  function renderizarClasificacionArchivosDetalle() {
    const contenedor = document.getElementById(
      "detalle-clasificacion-archivos"
    );
    if (!contenedor) {
      return;
    }
    contenedor.replaceChildren();
    contenedor.hidden = !archivosDetalleSeleccionados.length;
    actualizarResumenArchivosDetalle();

    archivosDetalleSeleccionados.forEach(function (archivo, indice) {
      const fila = document.createElement("div");
      fila.className = "file-category-row detail-file-category-row";

      const nombre = document.createElement("span");
      nombre.className = "file-category-name";
      nombre.textContent = archivo.name;
      nombre.title = archivo.name;

      const botonEliminar = document.createElement("button");
      botonEliminar.type = "button";
      botonEliminar.className = "file-category-remove";
      botonEliminar.textContent = "\u00d7";
      botonEliminar.title = "Quitar archivo";
      botonEliminar.setAttribute(
        "aria-label",
        "Quitar el archivo " + archivo.name
      );
      botonEliminar.addEventListener("click", function () {
        tiposDocumentoDetalleSeleccionados.delete(claveArchivo(archivo));
        archivosDetalleSeleccionados.splice(indice, 1);
        renderizarClasificacionArchivosDetalle();
      });

      fila.appendChild(nombre);
      fila.appendChild(
        crearSelectorTipoDocumentoDetalle(archivo, indice)
      );
      fila.appendChild(botonEliminar);
      contenedor.appendChild(fila);
    });
  }

  async function agregarArchivosDetalle(evento) {
    const nuevosArchivos = Array.from(evento.target.files || []);
    evento.target.value = "";
    if (!nuevosArchivos.length) {
      return;
    }

    const archivosPorClave = new Map(
      archivosDetalleSeleccionados.map(function (archivo) {
        return [claveArchivo(archivo), archivo];
      })
    );
    nuevosArchivos.forEach(function (archivo) {
      archivosPorClave.set(claveArchivo(archivo), archivo);
    });

    if (
      archivosPorClave.size >
      Modelo.configuracion.cantidadMaximaArchivos
    ) {
      await mostrarAlerta({
        icon: "warning",
        title: "Demasiados archivos",
        text:
          "Puede adjuntar m\u00e1ximo " +
          Modelo.configuracion.cantidadMaximaArchivos +
          " archivos por operaci\u00f3n."
      });
      return;
    }

    archivosDetalleSeleccionados = Array.from(
      archivosPorClave.values()
    );
    renderizarClasificacionArchivosDetalle();
  }

  function archivosClasificadosDetalle() {
    return archivosDetalleSeleccionados.map(function (archivo) {
      return {
        archivo: archivo,
        tipoDocumento:
          tiposDocumentoDetalleSeleccionados.get(
            claveArchivo(archivo)
          ) || ""
      };
    });
  }

  async function subirArchivosDesdeDetalle() {
    if (
      !requerimientoDetalleActual ||
      !archivosDetalleSeleccionados.length
    ) {
      return;
    }
    const boton = document.getElementById("detalle-subir-archivos");
    const archivosClasificados = archivosClasificadosDetalle();
    if (boton) {
      boton.disabled = true;
    }

    try {
      validarArchivos(archivosClasificados, "tipoDocumentoDetalle-");
      const resultado = await Modelo.guardarArchivosRequerimiento(
        requerimientoDetalleActual,
        archivosClasificados
      );

      requerimientoDetalleActual.archivosAdjuntos =
        await Modelo.obtenerDocumentosRequerimiento(
          requerimientoDetalleActual
        );
      archivosDetalleSeleccionados = [];
      tiposDocumentoDetalleSeleccionados.clear();
      Vista.mostrarDetalle(requerimientoDetalleActual);

      if (resultado.errores.length) {
        await mostrarAlerta({
          icon: "warning",
          title: "Algunos archivos no se cargaron",
          text:
            "No fue posible cargar: " +
            resultado.errores
              .map(function (error) {
                return error.nombre;
              })
              .join(", ") +
            "."
        });
      } else {
        await registrarActividadCompartida(
          "Agreg\u00f3 " +
            resultado.cargados.length +
            (resultado.cargados.length === 1
              ? " archivo al requerimiento "
              : " archivos al requerimiento ") +
            requerimientoDetalleActual.id,
          {
            tipo: "Archivos",
            requerimiento: requerimientoDetalleActual.id
          }
        );
        await mostrarAlerta({
          icon: "success",
          title: "Archivos agregados",
          text:
            resultado.cargados.length +
            (resultado.cargados.length === 1
              ? " archivo fue cargado correctamente."
              : " archivos fueron cargados correctamente.")
        });
      }
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudieron agregar los archivos",
        text: error.message
      });
    } finally {
      const botonActual = document.getElementById(
        "detalle-subir-archivos"
      );
      if (botonActual) {
        botonActual.disabled =
          archivosDetalleSeleccionados.length === 0;
      }
    }
  }

  function validarArchivos(archivosClasificados, prefijoSelector) {
    const extensionesPermitidas = [
      "pdf",
      "doc",
      "docx",
      "xls",
      "xlsx",
      "png",
      "jpg",
      "jpeg",
      "zip"
    ];
    const configuracion = Modelo.configuracion;

    if (
      archivosClasificados.length > configuracion.cantidadMaximaArchivos
    ) {
      throw new Error(
        "Puede adjuntar m\u00e1ximo " +
          configuracion.cantidadMaximaArchivos +
          " archivos por operaci\u00f3n."
      );
    }

    archivosClasificados.forEach(function (clasificacion, indice) {
      const archivo = clasificacion.archivo;
      if (!clasificacion.tipoDocumento) {
        const selector = document.getElementById(
          (prefijoSelector || "tipoDocumentoArchivo-") + indice
        );
        if (selector) {
          selector.setAttribute("aria-invalid", "true");
          selector.focus();
        }
        throw new Error(
          "Seleccione el tipo de documento para " + archivo.name + "."
        );
      }
      const partesNombre = archivo.name.split(".");
      const extension =
        partesNombre.length > 1 ? partesNombre.pop().toLowerCase() : "";

      if (extensionesPermitidas.indexOf(extension) === -1) {
        throw new Error(
          "El archivo " +
            archivo.name +
            " tiene una extensi\u00f3n no permitida."
        );
      }
      if (archivo.size > configuracion.tamanoMaximoArchivo) {
        throw new Error(
          "El archivo " +
            archivo.name +
            " supera el tama\u00f1o m\u00e1ximo de 20 MB."
        );
      }
      if (/[#%]/.test(archivo.name)) {
        throw new Error(
          "El nombre " +
            archivo.name +
            " contiene # o %. Cambie el nombre del archivo antes de adjuntarlo."
        );
      }
    });
  }

  function textoResultadoArchivos(resultado) {
    if (!resultado || !resultado.cargados.length) {
      return "";
    }
    return (
      " Se cargaron " +
      resultado.cargados.length +
      (resultado.cargados.length === 1 ? " archivo." : " archivos.")
    );
  }

  async function mostrarErroresCarga(resultado, idRequerimiento) {
    if (!resultado || !resultado.errores.length) {
      return false;
    }
    const primerDetalle = resultado.errores[0].mensaje
      ? " Detalle: " + resultado.errores[0].mensaje
      : "";
    global.BitacoraDiagnosticoArchivos = {
      requerimiento: idRequerimiento,
      resultado: resultado,
      fecha: new Date().toISOString()
    };
    console.error(
      "Diagn\u00f3stico de carga de archivos:",
      global.BitacoraDiagnosticoArchivos
    );
    await mostrarAlerta({
      icon: "warning",
      title: "Requerimiento guardado con archivos pendientes",
      text:
        "El requerimiento " +
        idRequerimiento +
        " fue guardado, pero no se pudieron cargar: " +
        resultado.errores
          .map(function (error) {
            return error.nombre;
          })
          .join(", ") +
        "." +
        primerDetalle +
        " Puede editar el requerimiento e intentarlo nuevamente."
    });
    return true;
  }

  async function agregarPersona() {
    const resultado = await mostrarDialogo({
      icon: "info",
      title: "Agregar persona",
      text: "Ingrese el nombre de la persona que participara en el requerimiento.",
      input: true,
      inputPlaceholder: "Ejemplo: Cristian Pardo",
      showCancelButton: true,
      confirmButtonText: "Agregar",
      cancelButtonText: "Cancelar"
    });

    if (!resultado.isConfirmed) {
      return;
    }
    const nombre = resultado.value;

    const campo = document.getElementById("solicitadoPor");
    const personas = campo.value
      .split(";")
      .map(function (persona) {
        return persona.trim();
      })
      .filter(Boolean);
    const yaExiste = personas.some(function (persona) {
      return persona.toLowerCase() === nombre.toLowerCase();
    });

    if (yaExiste) {
      mostrarAlerta({
        icon: "info",
        title: "Persona ya agregada",
        text: "La persona seleccionada ya hace parte del requerimiento.",
        confirmButtonColor: "#2f6fed"
      });
      return;
    }

    personas.push(nombre);
    campo.value = personas.join("; ");
  }

  function formularioValido(datos) {
    if (!datos.id) {
      mostrarAlerta({
        icon: "error",
        title: "No se pudo generar el ID",
        text: "Verifique la conexi\u00f3n con SharePoint e intente nuevamente.",
        confirmButtonText: "Entendido",
        confirmButtonColor: "#2f6fed"
      });
      return false;
    }
   if (
 !datos.app ||
 !datos.tipoServicio ||
 !datos.asunto ||
 !datos.descripcion
)
    {
      mostrarAlerta({
        icon: "warning",
        title: "Informaci\u00f3n incompleta",
        text: "Complete todos los campos obligatorios antes de guardar el requerimiento.",
        confirmButtonText: "Entendido",
        confirmButtonColor: "#2f6fed"
      });
      return false;
    }
    return true;
  }
// ============================================================
// RENDERIZAR USUARIOS SELECCIONADOS
// ------------------------------------------------------------
// Muestra las personas elegidas en el formulario.
// ============================================================
function cargarSolicitanteActual(){

    const usuario = Modelo.usuarioActual();


    if(!usuario){
        return;
    }


    usuariosSeleccionados = [
        {
            nombre: usuario.nombre,
            correo: usuario.correo
        }
    ];


    renderizarUsuariosSeleccionados();

}

function renderizarUsuariosSeleccionados() {

    const contenedor =
        document.getElementById("usuariosSeleccionados");


    if (!contenedor) {
        return;
    }


    contenedor.innerHTML =
        usuariosSeleccionados.length
        ?
        usuariosSeleccionados.map(function(usuario, index){

            return (

                '<div class="usuario-chip">' +

                usuario.nombre +

                '<button type="button" data-index="' +
                index +
                '">&times;</button>' +

                '</div>'

            );

        }).join("")

        :

        "<span>No hay usuarios seleccionados</span>";


    // Actualiza el campo que se guarda en SharePoint

    const campo =
        document.getElementById("solicitadoPor");


    if(campo){

        campo.value =
            usuariosSeleccionados
            .map(function(usuario){
                return usuario.nombre;
            })
            .join("\n");

    }

}

// ============================================================
// BUSCAR USUARIOS MICROSOFT
// ============================================================

async function buscarUsuarios() {


    const entrada =
        document.getElementById("buscarUsuario");


    const resultados =
        document.getElementById("resultadoUsuarios");


    if(!entrada || !resultados){

        return;

    }


    const texto = entrada.value.trim();


    if(texto.length < 3){

        resultados.hidden = true;

        resultados.innerHTML = "";

        return;

    }


    try {


        const usuarios =
            await Modelo.buscarUsuariosMicrosoft(texto);



        resultados.innerHTML =
            usuarios.map(function(usuario){


                return (

                '<button type="button" class="usuario-opcion" ' +

                'data-nombre="' +
                usuario.nombre +
                '" ' +

                'data-correo="' +
                usuario.correo +
                '">' +

                usuario.nombre +

                '</button>'

                );


            }).join("");



        resultados.hidden = usuarios.length === 0;



    } catch(error){


        console.error(
            "Error buscando usuarios:",
            error
        );


    }

}

  async function guardarFormulario() {
    const datos = datosFormulario();
    const archivosClasificados = archivosClasificadosFormulario();
    const botonGuardar = document.getElementById("btn-guardar");
    if (!formularioValido(datos)) {
      return;
    }
    botonGuardar.disabled = true;
    try {
      validarArchivos(archivosClasificados);
      if (requerimientoEnEdicion) {
        const requerimientoActualizado = await Modelo.actualizar(
          requerimientoEnEdicion.id,
          {
            app: datos.app,
            tipoServicio: datos.tipoServicio,
            asunto: datos.asunto,
            descripcion: datos.descripcion,
            casoOrigen: datos.casoOrigen,
            prioridad: datos.prioridad,
            solicitadoPor: datos.solicitadoPor
          }
        );
        const resultadoArchivos = await Modelo.guardarArchivosRequerimiento(
          requerimientoActualizado,
          archivosClasificados
        );
        await registrarActividadCompartida(
          "Edit\u00f3 el requerimiento " + requerimientoEnEdicion.id,
          {
            tipo: "Edici\u00f3n",
            requerimiento: requerimientoEnEdicion.id
          }
        );
        requerimientoEnEdicion = null;
        if (
          await mostrarErroresCarga(
            resultadoArchivos,
            requerimientoActualizado.id
          )
        ) {
          if (origenEdicion === "dashboard") {
            Vista.mostrarDashboard();
          } else {
            Vista.mostrarMisRequerimientos();
          }
          return;
        }
        await mostrarAlerta({
          icon: "success",
          title: "Requerimiento actualizado",
          text:
            "Los cambios fueron guardados correctamente." +
            textoResultadoArchivos(resultadoArchivos)
        });
        if (origenEdicion === "dashboard") {
          Vista.mostrarDashboard();
        } else {
          Vista.mostrarMisRequerimientos();
        }
      } else {
        const requerimientoCreado = await Modelo.crear(datos);
        const resultadoArchivos = await Modelo.guardarArchivosRequerimiento(
          requerimientoCreado,
          archivosClasificados
        );
        await registrarActividadCompartida(
          "Cre\u00f3 el requerimiento " + datos.id + " - " + datos.asunto,
          {
            tipo: "Creaci\u00f3n",
            requerimiento: datos.id
          }
        );
        if (
          await mostrarErroresCarga(resultadoArchivos, requerimientoCreado.id)
        ) {
          Vista.mostrarDashboard();
          return;
        }
        await mostrarAlerta({
          icon: "success",
          title: "\u00a1Requerimiento registrado!",
          html:
            "<b>ID del requerimiento:</b> " +
            datos.id +
            "<br><br>El requerimiento fue registrado correctamente." +
            textoResultadoArchivos(resultadoArchivos),
          text:
            "Requerimiento " +
            datos.id +
            " registrado correctamente." +
            textoResultadoArchivos(resultadoArchivos),
          confirmButtonText: "Aceptar",
          confirmButtonColor: "#2f6fed"
        });
        Vista.mostrarDashboard();
      }
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudo guardar el requerimiento",
        text: error.message
      });
    } finally {
      botonGuardar.disabled = false;
    }
  }

  function renderizarPaginaMisRequerimientos() {
    const datosFiltrados = filtrarMisRequerimientos();
    const total = datosFiltrados.length;
    const totalPaginas = Math.max(
      1,
      Math.ceil(total / REQUERIMIENTOS_POR_PAGINA)
    );
    paginaMisRequerimientos = Math.min(
      Math.max(1, paginaMisRequerimientos),
      totalPaginas
    );
    const indiceInicial =
      (paginaMisRequerimientos - 1) * REQUERIMIENTOS_POR_PAGINA;
    const datosPagina = datosFiltrados.slice(
      indiceInicial,
      indiceInicial + REQUERIMIENTOS_POR_PAGINA
    );

    Vista.renderizarMisRequerimientos(datosPagina, {
      pagina: paginaMisRequerimientos,
      totalPaginas: totalPaginas,
      total: total,
      inicio: total ? indiceInicial + 1 : 0,
      fin: Math.min(indiceInicial + REQUERIMIENTOS_POR_PAGINA, total),
      filtrosActivos: Boolean(
        document.getElementById("buscador-mis-requerimientos").value.trim() ||
        document.getElementById("filtro-app-mis").value ||
        document.getElementById("filtro-estado-mis").value ||
        document.getElementById("filtro-fecha-mis").value
      )
    });
  }

  function filtrarMisRequerimientos() {
    const buscador = document.getElementById("buscador-mis-requerimientos");
    const texto = buscador ? buscador.value.trim().toLowerCase() : "";
    const app = document.getElementById("filtro-app-mis").value;
    const estado = document.getElementById("filtro-estado-mis").value;
    const fecha = document.getElementById("filtro-fecha-mis").value;

    return requerimientosPropios.filter(function (req) {
      const contenido = [
        req.id,
        req.app,
        req.tipoServicio,
        req.asunto,
        req.descripcion,
        req.solicitadoPor,
        req.responsable,
        req.estado
      ]
        .join(" ")
        .toLowerCase();
      return (
        contenido.indexOf(texto) !== -1 &&
        (!app || req.app === app) &&
        (!estado || req.estado === estado) &&
        coincideFecha(req.fechaSolicitud, fecha)
      );
    });
  }

  function buscarMisRequerimientos() {
    paginaMisRequerimientos = 1;
    renderizarPaginaMisRequerimientos();
  }

  function cambiarPaginaMisRequerimientos(pagina) {
    const destino = Number(pagina);
    const totalPaginas = Math.max(
      1,
      Math.ceil(filtrarMisRequerimientos().length / REQUERIMIENTOS_POR_PAGINA)
    );
    if (
      !Number.isInteger(destino) ||
      destino < 1 ||
      destino > totalPaginas ||
      destino === paginaMisRequerimientos
    ) {
      return;
    }
    paginaMisRequerimientos = destino;
    renderizarPaginaMisRequerimientos();
    document
      .getElementById("titulo-mis-requerimientos")
      .scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function cargarMisRequerimientos() {
    try {
      const datos = await Modelo.obtenerTodos();
      requerimientosPropios = datos.filter(function (req) {
        return esRequerimientoPropio(req);
      });
      Vista.renderizarFiltrosMisRequerimientos(requerimientosPropios);
      paginaMisRequerimientos = 1;
      renderizarPaginaMisRequerimientos();
    } catch (error) {
      console.error("Carga de requerimientos personales:", error);
      requerimientosPropios = [];
      Vista.renderizarFiltrosMisRequerimientos([]);
      paginaMisRequerimientos = 1;
      renderizarPaginaMisRequerimientos();
    }
  }

  function actualizarTiposHistorial() {
    const select = document.getElementById("filtro-tipo-historial");
    const valorActual = select.value;
    const tipos = actividadesHistorial
      .map(function (actividad) {
        return actividad.tipo;
      })
      .filter(Boolean)
      .filter(function (tipo, indice, arreglo) {
        return arreglo.indexOf(tipo) === indice;
      })
      .sort();
    select.innerHTML =
      '<option value="">Todos los tipos</option>' +
      tipos
        .map(function (tipo) {
          const seguro = String(tipo)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;");
          return '<option value="' + seguro + '">' + seguro + "</option>";
        })
        .join("");
    if (tipos.indexOf(valorActual) !== -1) {
      select.value = valorActual;
    }
  }

  function filtrarHistorial() {
    const texto = document
      .getElementById("buscador-historial")
      .value.trim()
      .toLowerCase();
    const tipo = document.getElementById("filtro-tipo-historial").value;
    const fecha = document.getElementById("filtro-fecha-historial").value;
    return actividadesHistorial.filter(function (actividad) {
      const contenido = [
        actividad.usuario,
        actividad.correo,
        actividad.accion,
        actividad.requerimiento,
        actividad.tipo
      ]
        .join(" ")
        .toLowerCase();
      return (
        contenido.indexOf(texto) !== -1 &&
        (!tipo || actividad.tipo === tipo) &&
        coincideFecha(actividad.fecha, fecha)
      );
    });
  }

  function renderizarPaginaHistorial() {
    const filtradas = filtrarHistorial();
    const total = filtradas.length;
    const totalPaginas = Math.max(
      1,
      Math.ceil(total / ACTIVIDADES_POR_PAGINA)
    );
    paginaHistorial = Math.min(Math.max(1, paginaHistorial), totalPaginas);
    const inicio = (paginaHistorial - 1) * ACTIVIDADES_POR_PAGINA;
    Vista.renderizarActividad(
      filtradas.slice(inicio, inicio + ACTIVIDADES_POR_PAGINA),
      {
        pagina: paginaHistorial,
        totalPaginas: totalPaginas,
        total: total,
        inicio: total ? inicio + 1 : 0,
        fin: Math.min(inicio + ACTIVIDADES_POR_PAGINA, total)
      }
    );
  }

  async function cargarHistorial() {
    const estado = document.getElementById("estado-historial");
    estado.textContent = "Consultando la actividad compartida en SharePoint\u2026";
    try {
      actividadesHistorial = await Modelo.obtenerBitacora();
      paginaHistorial = 1;
      actualizarTiposHistorial();
      renderizarPaginaHistorial();
      estado.textContent =
        actividadesHistorial.length +
        (actividadesHistorial.length === 1
          ? " actividad encontrada."
          : " actividades encontradas.");
    } catch (error) {
      actividadesHistorial = [];
      paginaHistorial = 1;
      actualizarTiposHistorial();
      renderizarPaginaHistorial();
      estado.textContent =
        error.message +
        " Revisa la configuraci\u00f3n indicada para la lista HistorialActividad.";
      console.error("Carga del historial:", error);
    }
  }

  function aplicarFiltrosHistorial() {
    paginaHistorial = 1;
    renderizarPaginaHistorial();
  }

  function restablecerFiltrosHistorial() {
    document.getElementById("buscador-historial").value = "";
    document.getElementById("filtro-tipo-historial").value = "";
    document.getElementById("filtro-fecha-historial").value = "";
    paginaHistorial = 1;
  }

  function limpiarFiltrosHistorial() {
    restablecerFiltrosHistorial();
    aplicarFiltrosHistorial();
  }

  async function actualizarHistorial() {
    const boton = document.getElementById("btn-actualizar-historial");
    restablecerFiltrosHistorial();
    boton.disabled = true;
    boton.setAttribute("aria-busy", "true");
    boton.textContent = "Actualizando\u2026";
    try {
      await cargarHistorial();
    } finally {
      boton.disabled = false;
      boton.removeAttribute("aria-busy");
      boton.textContent = "Actualizar";
    }
  }

  async function exportarHistorialCSV() {
    const datos = filtrarHistorial();
    if (!datos.length) {
      await mostrarAlerta({
        icon: "info",
        title: "Sin datos para exportar",
        text: "No hay actividades que coincidan con los filtros actuales."
      });
      return;
    }

    const confirmacion = await mostrarDialogo({
      icon: "info",
      title: "Exportar historial CSV",
      text:
        "Se descargar\u00e1 un archivo CSV con " +
        datos.length +
        (datos.length === 1 ? " actividad" : " actividades") +
        " seg\u00fan los filtros actuales.",
      showCancelButton: true,
      confirmButtonText: "Descargar",
      cancelButtonText: "Cancelar"
    });
    if (!confirmacion.isConfirmed) {
      return;
    }

    const columnas = [
      ["fecha", "Fecha"],
      ["usuario", "Usuario"],
      ["correo", "Correo"],
      ["tipo", "Tipo"],
      ["requerimiento", "Requerimiento"],
      ["accion", "Actividad"]
    ];
    const filas = [
      columnas.map(function (columna) {
        return columna[1];
      })
    ];

    datos.forEach(function (actividad) {
      filas.push(
        columnas.map(function (columna) {
          const valor = actividad[columna[0]];
          if (columna[0] !== "fecha" || !valor) {
            return valor;
          }
          const fecha = new Date(valor);
          return isNaN(fecha.getTime())
            ? valor
            : fecha.toLocaleString("es-CO");
        })
      );
    });

    const csv = filas
      .map(function (fila) {
        return fila.map(escaparCSV).join(";");
      })
      .join("\r\n");
    const blob = new Blob(["\ufeff" + csv], {
      type: "text/csv;charset=utf-8;"
    });
    const url = URL.createObjectURL(blob);
    const enlace = document.createElement("a");
    enlace.href = url;
    enlace.download =
      "historial_" + new Date().toISOString().slice(0, 10) + ".csv";
    document.body.appendChild(enlace);
    enlace.click();
    enlace.remove();
    URL.revokeObjectURL(url);
  }

  function cambiarPaginaHistorial(pagina) {
    const destino = Number(pagina);
    const totalPaginas = Math.max(
      1,
      Math.ceil(filtrarHistorial().length / ACTIVIDADES_POR_PAGINA)
    );
    if (
      !Number.isInteger(destino) ||
      destino < 1 ||
      destino > totalPaginas ||
      destino === paginaHistorial
    ) {
      return;
    }
    paginaHistorial = destino;
    renderizarPaginaHistorial();
  }

  async function cargarGestion() {
    try {
      Vista.renderizarGestion(await Modelo.obtenerTodos());
    } catch (error) {
      console.error("Carga de gesti\u00f3n:", error);
      Vista.renderizarGestion([]);
    }
  }

  async function cargarGestion() {
    try {
      const todos = await Modelo.obtenerTodos();
      // Un requerimiento con Responsable Y Observaciones ya diligenciados
      // se considera gestionado y deja de listarse aqui (sigue existiendo
      // y visible normalmente en Backlog / Mis requerimientos).
      requerimientosGestion = todos.filter(function (req) {
        const tieneResponsable = Boolean(String(req.responsable || "").trim());
        const tieneObservacion = Boolean(String(req.comentarios || "").trim());
        return !(tieneResponsable && tieneObservacion);
      });
      paginaGestion = 1;
      renderizarPaginaGestion();
    } catch (error) {
      console.error("Carga de gesti\u00f3n:", error);
      requerimientosGestion = [];
      paginaGestion = 1;
      renderizarPaginaGestion();
    }
  }

  function filtrarGestion() {
    const texto = document.getElementById("buscador-gestion").value.toLowerCase();
    const fecha = document.getElementById("filtro-fecha-gestion").value;
    return requerimientosGestion.filter(function (req) {
      const contenido = [req.id, req.asunto, req.responsable, req.mentor]
        .join(" ")
        .toLowerCase();
      return (
        contenido.indexOf(texto) !== -1 &&
        coincideFecha(req.fechaEntrega, fecha)
      );
    });
  }

  function renderizarPaginaGestion() {
    const datosFiltrados = filtrarGestion();
    const total = datosFiltrados.length;
    const totalPaginas = Math.max(
      1,
      Math.ceil(total / REQUERIMIENTOS_POR_PAGINA)
    );
    paginaGestion = Math.min(Math.max(1, paginaGestion), totalPaginas);
    const indiceInicial = (paginaGestion - 1) * REQUERIMIENTOS_POR_PAGINA;
    const datosPagina = datosFiltrados.slice(
      indiceInicial,
      indiceInicial + REQUERIMIENTOS_POR_PAGINA
    );

    Vista.renderizarGestion(datosPagina, {
      pagina: paginaGestion,
      totalPaginas: totalPaginas,
      total: total,
      inicio: total ? indiceInicial + 1 : 0,
      fin: Math.min(indiceInicial + REQUERIMIENTOS_POR_PAGINA, total)
    });
  }

  function cambiarPaginaGestion(pagina) {
    const destino = Number(pagina);
    const totalPaginas = Math.max(
      1,
      Math.ceil(filtrarGestion().length / REQUERIMIENTOS_POR_PAGINA)
    );
    if (
      !Number.isInteger(destino) ||
      destino < 1 ||
      destino > totalPaginas ||
      destino === paginaGestion
    ) {
      return;
    }
    paginaGestion = destino;
    renderizarPaginaGestion();
  }

  function aplicarFiltrosGestion() {
    paginaGestion = 1;
    renderizarPaginaGestion();
  }

  function limpiarFiltrosGestion() {
    document.getElementById("buscador-gestion").value = "";
    document.getElementById("filtro-fecha-gestion").value = "";
    paginaGestion = 1;
    renderizarPaginaGestion();
  }

  async function guardarGestion(id, fila) {
    const estado = fila.querySelector(".gestion-estado").value;
    const comentarios = fila.querySelector(".gestion-comentarios").value.trim();
    const campoResponsable = fila.querySelector(".gestion-responsable");
    const campoMentor = fila.querySelector(".gestion-mentor");
    const valorResponsable = campoResponsable.value.trim();
    const valorMentor = campoMentor.value.trim();
    const originalResponsable = campoResponsable.dataset.original || "";
    const originalMentor = campoMentor.dataset.original || "";

    const cambios = {
      estado: estado,
      comentarios: comentarios
    };

    try {
      if (valorResponsable !== originalResponsable) {
        if (!valorResponsable) {
          cambios.responsable = null;
        } else {
          const persona = await Modelo.resolverUsuarioPorCorreo(valorResponsable);
          cambios.responsable = persona.id;
        }
      }
      if (valorMentor !== originalMentor) {
        if (!valorMentor) {
          cambios.mentor = null;
        } else {
          const persona = await Modelo.resolverUsuarioPorCorreo(valorMentor);
          cambios.mentor = persona.id;
        }
      }
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudo asignar la persona",
        text: error.message
      });
      return;
    }

    try {
      await Modelo.actualizar(id, cambios);
      await registrarActividadCompartida(
        "Actualiz\u00f3 el requerimiento " + id,
        {
          tipo: "Gesti\u00f3n",
          requerimiento: id
        }
      );
      await mostrarAlerta({
        icon: "success",
        title: "Requerimiento actualizado",
        text: "Los cambios fueron guardados correctamente."
      });
      await cargarGestion();
    } catch (error) {
      await mostrarAlerta({
        icon: "error",
        title: "No se pudo actualizar el requerimiento",
        text: error.message
      });
    }
  }

  function registrarEventos() {
    if (eventosRegistrados) {
      return;
    }
    eventosRegistrados = true;

    const btnAgregarPersona =
   document.getElementById("btnAgregarPersona");

   if (btnAgregarPersona) {

   btnAgregarPersona.addEventListener(
    "click",
    agregarPersona
  );

// BOTON LIMPIAR FILTROS INDICADORES





const btnLimpiarIndicadores =
document.getElementById("btnLimpiarIndicadores");


if(btnLimpiarIndicadores){

    btnLimpiarIndicadores.onclick = function(e){

        e.preventDefault();


        document.getElementById(
            "fechaInicioIndicadores"
        ).value = "";


        document.getElementById(
            "fechaFinIndicadores"
        ).value = "";


        cargarIndicadores();


        return false;

    };

}

}
    document.getElementById("buscador").addEventListener("input", aplicarFiltros);
    document
      .getElementById("buscador-mis-requerimientos")
      .addEventListener("input", buscarMisRequerimientos);
    ["filtro-app-mis", "filtro-estado-mis", "filtro-fecha-mis"]
      .forEach(function (id) {
        document
          .getElementById(id)
          .addEventListener("change", buscarMisRequerimientos);
      });
    [
      "filtro-app",
      "filtro-responsable",
      "filtro-estado",
      "filtro-prioridad",
      "filtro-fecha"
    ]
      .forEach(function (id) {
        document.getElementById(id).addEventListener("change", aplicarFiltros);
      });
    document
      .getElementById("btn-limpiar-filtros")
      .addEventListener("click", limpiarFiltros);
    document
      .getElementById("btn-exportar")
      .addEventListener("click", exportarCSV);
   document

    document.querySelectorAll("[data-vista]").forEach(function (boton) {
      boton.addEventListener("click", function () {
        const vista = boton.dataset.vista;
        if (vista === "dashboard") {
          Vista.mostrarDashboard();
        } else if (vista === "mis-requerimientos") {
          Vista.mostrarMisRequerimientos();
        } else if (vista === "gestion") {
          Vista.mostrarGestion();
        } else if (vista === "historial") {
          Vista.mostrarHistorial();
        } else if (vista === "indicadores") {
        Vista.mostrarIndicadores();
        } else if (vista === "crear") {
        Vista.mostrarFormularioCrear();
        }
      });
    });
    document
      .getElementById("btn-crear-requerimiento")
      .addEventListener("click", Vista.mostrarFormularioCrear);
    document
      .getElementById("btn-guardar")
      .addEventListener("click", guardarFormulario);
    document
      .getElementById("archivosAdjuntos")
      .addEventListener("change", agregarArchivosSeleccionados);
    document
      .getElementById("btn-limpiar")
      .addEventListener("click", function (evento) {
        evento.preventDefault();
        prepararFormulario();
      });
    document
      .getElementById("btn-cancelar")
      .addEventListener("click", function () {
        if (requerimientoEnEdicion) {
          requerimientoEnEdicion = null;
          if (origenEdicion === "dashboard") {
            Vista.mostrarDashboard();
          } else {
            Vista.mostrarMisRequerimientos();
          }
        } else {
          Vista.mostrarDashboard();
        }
      });

    document
      .getElementById("tabla-requerimientos")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button");
        if (!boton || !boton.dataset.id) {
          return;
        }
        if (boton.classList.contains("view-btn")) {
          verRequerimiento(boton.dataset.id);
        } else if (boton.classList.contains("edit-btn")) {
          editarRequerimiento(boton.dataset.id, "dashboard");
        }
      });

    document
      .getElementById("tabla-mis-requerimientos")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button");
        if (!boton || !boton.dataset.id) {
          return;
        }
        if (boton.classList.contains("view-btn")) {
          verRequerimiento(boton.dataset.id);
        } else if (boton.classList.contains("edit-btn")) {
          editarRequerimiento(boton.dataset.id);
        }
      });

    document
      .getElementById("controles-paginacion-mis")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button[data-pagina]");
        if (boton && !boton.disabled) {
          cambiarPaginaMisRequerimientos(boton.dataset.pagina);
        }
      });
    document
      .getElementById("controles-paginacion")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button[data-pagina]");
        if (boton && !boton.disabled) {
          cambiarPaginaBacklog(boton.dataset.pagina);
        }
      });
document
      .getElementById("tabla-gestion")
      .addEventListener("click", function (evento) {
        const botonGuardar = evento.target.closest(".save-btn");
        if (botonGuardar) {
          guardarGestion(botonGuardar.dataset.id, botonGuardar.closest("tr"));
          return;
        }
        const botonVer = evento.target.closest(".view-btn");
        if (botonVer) {
          verRequerimiento(botonVer.dataset.id);
        }
      });

    document
      .getElementById("buscador-gestion")
      .addEventListener("input", aplicarFiltrosGestion);
    document
      .getElementById("filtro-fecha-gestion")
      .addEventListener("change", aplicarFiltrosGestion);
    document
      .getElementById("btn-limpiar-filtros-gestion")
      .addEventListener("click", limpiarFiltrosGestion);
    document
      .getElementById("controles-paginacion-gestion")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button[data-pagina]");
        if (boton && !boton.disabled) {
          cambiarPaginaGestion(boton.dataset.pagina);
        }
      });

    document
      .getElementById("buscador-historial")
      .addEventListener("input", aplicarFiltrosHistorial);
    document
      .getElementById("filtro-tipo-historial")
      .addEventListener("change", aplicarFiltrosHistorial);
    document
      .getElementById("filtro-fecha-historial")
      .addEventListener("change", aplicarFiltrosHistorial);
    document
      .getElementById("btn-limpiar-historial")
      .addEventListener("click", limpiarFiltrosHistorial);
    document
      .getElementById("btn-exportar-historial")
      .addEventListener("click", exportarHistorialCSV);
    document
      .getElementById("btn-actualizar-historial")
      .addEventListener("click", actualizarHistorial);
    document
      .getElementById("controles-paginacion-historial")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest("button[data-pagina]");
        if (boton && !boton.disabled) {
          cambiarPaginaHistorial(boton.dataset.pagina);
        }
      });


    const overlay = document.getElementById("modal-detalle");
    const overlayMensaje = document.getElementById("modal-mensaje");
    const contenidoDetalle = document.getElementById("modal-contenido");
    document
      .getElementById("modal-cerrar")
      .addEventListener("click", Vista.cerrarDetalle);
    overlay.addEventListener("click", function (evento) {
      if (evento.target === overlay) {
        Vista.cerrarDetalle();
      }
    });
    contenidoDetalle.addEventListener("change", function (evento) {
      if (evento.target.id === "detalle-archivos-nuevos") {
        agregarArchivosDetalle(evento);
      }
    });
    contenidoDetalle.addEventListener("click", function (evento) {
      const botonPagina = evento.target.closest(
        "[data-pagina-adjuntos]"
      );
      if (botonPagina && !botonPagina.disabled) {
        Vista.cambiarPaginaAdjuntos(
          Number(botonPagina.dataset.paginaAdjuntos)
        );
        return;
      }
      const botonSubir = evento.target.closest(
        "#detalle-subir-archivos"
      );
      if (botonSubir) {
        subirArchivosDesdeDetalle();
      }
    });
    document
      .getElementById("confirmar-modal-mensaje")
      .addEventListener("click", function () {
        cerrarDialogoMensaje(true);
      });
    document
      .getElementById("cancelar-modal-mensaje")
      .addEventListener("click", function () {
        cerrarDialogoMensaje(false);
      });
    document
      .getElementById("entrada-modal-mensaje")
      .addEventListener("keydown", function (evento) {
        if (evento.key === "Enter") {
          evento.preventDefault();
          cerrarDialogoMensaje(true);
        }
      });
    overlayMensaje.addEventListener("click", function (evento) {
      if (evento.target === overlayMensaje) {
        cerrarDialogoMensaje(false);
      }
    });
    document.addEventListener("keydown", function (evento) {
      if (evento.key === "Escape") {
        if (!overlayMensaje.hidden) {
          cerrarDialogoMensaje(false);
        } else if (!overlay.hidden) {
          Vista.cerrarDetalle();
        }
      }
    });
    // ============================================================
// EVENTO BUSCADOR DE USUARIOS
// ============================================================

const inputBuscarUsuario =
    document.getElementById("buscarUsuario");


if (inputBuscarUsuario) {

    inputBuscarUsuario.addEventListener(
        "input",
        buscarUsuarios
    );

}
// ============================================================
// SELECCIONAR USUARIO EN RESULTADOS
// ============================================================

const resultadosUsuarios =
    document.getElementById("resultadoUsuarios");


if (resultadosUsuarios) {

    resultadosUsuarios.addEventListener(
        "click",
        function(evento){


            const boton =
                evento.target.closest(
                    ".usuario-opcion"
                );


            if(!boton){
                return;
            }


            const usuario = {

                nombre:
                    boton.dataset.nombre,

                correo:
                    boton.dataset.correo

            };


            // Evita usuarios repetidos
            const existe =
                usuariosSeleccionados.some(
                    function(u){

                        return u.correo === usuario.correo;

                    }
                );


            if(!existe){

                usuariosSeleccionados.push(usuario);

            }


            renderizarUsuariosSeleccionados();


            // Limpia busqueda
            inputBuscarUsuario.value = "";


            resultadosUsuarios.hidden = true;


        }
    );

}
  }

  global.Controlador = Object.freeze({
    iniciar: iniciar,
    configurarVistaSharePoint: configurarVistaSharePoint,
    comprobarConexion: comprobarConexion,
    cargarDashboard: cargarDashboard,
    cargarIndicadores: cargarIndicadores,
    prepararFormulario: prepararFormulario,
    cargarMisRequerimientos: cargarMisRequerimientos,
    cambiarPaginaMisRequerimientos: cambiarPaginaMisRequerimientos,
    cambiarPaginaBacklog: cambiarPaginaBacklog,
    cargarGestion: cargarGestion,
    cambiarPaginaGestion: cambiarPaginaGestion,
    cargarHistorial: cargarHistorial,
    cambiarPaginaHistorial: cambiarPaginaHistorial
  });

  global.cargarDatosDashboard = cargarDashboard;
  configurarVistaSharePoint();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar, { once: true });
  } else {
    iniciar();
  }
})(window);
