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
  let dialogoMensajeActual = null;
  let focoAntesDelDialogo = null;

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
        });
    });

    const workspace =
      documentoSharePoint.getElementById("s4-workspace");
    if (workspace) {
      workspace.style.setProperty("overflow", "auto", "important");
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
      await cargarDashboard();
    } else {
      Vista.renderizarTarjetas([]);
      Vista.renderizarTabla([]);
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
          " registros)" +
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
          "RetirosDeCesantiasDesarrollo/Lists/Backlog/AllItems.aspx"
      };
      console.error("Validaci\u00f3n de SharePoint:", error);
      Vista.mostrarConexion(false, "Sin conexi\u00f3n: " + error.message);
      return false;
    }
  }

  async function cargarDashboard() {
    try {
      requerimientos = await Modelo.obtenerTodos();
      Vista.renderizarTarjetas(requerimientos);
      Vista.renderizarTabla(requerimientos);
      Vista.renderizarFiltros(requerimientos);
      Vista.renderizarActividad(Modelo.obtenerBitacora());
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

  function aplicarFiltros() {
    const texto = document.getElementById("buscador").value.toLowerCase();
    const filtros = {
      app: document.getElementById("filtro-app").value,
      responsable: document.getElementById("filtro-responsable").value,
      estado: document.getElementById("filtro-estado").value,
      prioridad: document.getElementById("filtro-prioridad").value
    };

    const datos = requerimientos.filter(function (req) {
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
          String(req.prioridad) === filtros.prioridad)
      );
    });
    Vista.renderizarTabla(datos);
  }

  function limpiarFiltros() {
    document.getElementById("buscador").value = "";
    ["filtro-app", "filtro-responsable", "filtro-estado", "filtro-prioridad"]
      .forEach(function (id) {
        document.getElementById(id).value = "";
      });
    Vista.renderizarTabla(requerimientos);
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
      "prioridad",
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

  async function editarRequerimiento(id) {
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
      if (!esRequerimientoPropio(req)) {
        await mostrarAlerta({
          icon: "warning",
          title: "Edici\u00f3n no permitida",
          text: "Solo puede editar requerimientos solicitados por usted."
        });
        return;
      }

      requerimientoEnEdicion = req;
      document.getElementById("modo-formulario").textContent =
        "EDICI\u00d3N DE SOLICITUD";
      document.getElementById("titulo-crear").textContent =
        "Editar requerimiento";
      document.getElementById("subtitulo-formulario").textContent =
        "Actualiza la informaci\u00f3n de tu solicitud.";
      document.getElementById("btn-guardar").textContent = "Guardar cambios";
      document.getElementById("btn-limpiar").hidden = true;

      document.getElementById("id").value = req.id || "";
      document.getElementById("app").value = req.app || "";
      document.getElementById("tipoServicio").value = req.tipoServicio || "";
      document.getElementById("casoOrigen").value = req.casoOrigen || "";
      document.getElementById("asunto").value = req.asunto || "";
      document.getElementById("descripcion").value = req.descripcion || "";
      document.getElementById("prioridad").value = req.prioridad || "";
      document.getElementById("solicitadoPor").value = req.solicitadoPor || "";
      document.getElementById("estado").value = req.estado || "";
      document.getElementById("fechaSolicitud").value =
        req.fechaSolicitud || "";

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
      prioridad: document.getElementById("prioridad").value,
      estado: document.getElementById("estado").value,
      fechaSolicitud: document.getElementById("fechaSolicitud").value
    };
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
      !datos.descripcion ||
      !datos.prioridad
    ) {
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

  async function guardarFormulario() {
    const datos = datosFormulario();
    if (!formularioValido(datos)) {
      return;
    }
    try {
      if (requerimientoEnEdicion) {
        await Modelo.actualizar(requerimientoEnEdicion.id, {
          app: datos.app,
          tipoServicio: datos.tipoServicio,
          asunto: datos.asunto,
          descripcion: datos.descripcion,
          casoOrigen: datos.casoOrigen,
          prioridad: datos.prioridad,
          solicitadoPor: datos.solicitadoPor
        });
        Modelo.agregarActividad(
          Modelo.usuarioActual().nombre,
          "Edit\u00f3 el requerimiento " + requerimientoEnEdicion.id
        );
        requerimientoEnEdicion = null;
        await mostrarAlerta({
          icon: "success",
          title: "Requerimiento actualizado",
          text: "Los cambios fueron guardados correctamente."
        });
        Vista.mostrarMisRequerimientos();
      } else {
        await Modelo.crear(datos);
        Modelo.agregarActividad(
          Modelo.usuarioActual().nombre,
          "Cre\u00f3 el requerimiento " + datos.id + " - " + datos.asunto
        );
        await mostrarAlerta({
          icon: "success",
          title: "\u00a1Requerimiento registrado!",
          html:
            "<b>ID del requerimiento:</b> " +
            datos.id +
            "<br><br>El requerimiento fue registrado correctamente.",
          text: "Requerimiento " + datos.id + " registrado correctamente.",
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
    }
  }

  async function cargarMisRequerimientos() {
    try {
      const datos = await Modelo.obtenerTodos();
      const propios = datos.filter(function (req) {
        return esRequerimientoPropio(req);
      });
      Vista.renderizarMisRequerimientos(propios);
    } catch (error) {
      console.error("Carga de requerimientos personales:", error);
      Vista.renderizarMisRequerimientos([]);
    }
  }

  async function cargarGestion() {
    try {
      Vista.renderizarGestion(await Modelo.obtenerTodos());
    } catch (error) {
      console.error("Carga de gesti\u00f3n:", error);
      Vista.renderizarGestion([]);
    }
  }

  async function guardarGestion(id, fila) {
    const estado = fila.querySelector(".gestion-estado").value;
    const comentarios = fila.querySelector(".gestion-comentarios").value.trim();
    try {
      await Modelo.actualizar(id, {
        estado: estado,
        comentarios: comentarios
      });
      Modelo.agregarActividad(
        Modelo.usuarioActual().nombre,
        "Actualiz\u00f3 el requerimiento " + id
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

    document
      .getElementById("btnAgregarPersona")
      .addEventListener("click", agregarPersona);
    document.getElementById("buscador").addEventListener("input", aplicarFiltros);
    ["filtro-app", "filtro-responsable", "filtro-estado", "filtro-prioridad"]
      .forEach(function (id) {
        document.getElementById(id).addEventListener("change", aplicarFiltros);
      });
    document
      .getElementById("btn-limpiar-filtros")
      .addEventListener("click", limpiarFiltros);
    document
      .getElementById("btn-actualizar")
      .addEventListener("click", cargarDashboard);
    document.querySelectorAll("[data-vista]").forEach(function (boton) {
      boton.addEventListener("click", function () {
        const vista = boton.dataset.vista;
        if (vista === "dashboard") {
          Vista.mostrarDashboard();
        } else if (vista === "mis-requerimientos") {
          Vista.mostrarMisRequerimientos();
        } else if (vista === "gestion") {
          Vista.mostrarGestion();
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
          Vista.mostrarMisRequerimientos();
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
      .getElementById("tabla-gestion")
      .addEventListener("click", function (evento) {
        const boton = evento.target.closest(".save-btn");
        if (boton) {
          guardarGestion(boton.dataset.id, boton.closest("tr"));
        }
      });

    const overlay = document.getElementById("modal-detalle");
    const overlayMensaje = document.getElementById("modal-mensaje");
    document
      .getElementById("modal-cerrar")
      .addEventListener("click", Vista.cerrarDetalle);
    document
      .getElementById("modal-cerrar-btn")
      .addEventListener("click", Vista.cerrarDetalle);
    overlay.addEventListener("click", function (evento) {
      if (evento.target === overlay) {
        Vista.cerrarDetalle();
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
  }

  global.Controlador = Object.freeze({
    iniciar: iniciar,
    configurarVistaSharePoint: configurarVistaSharePoint,
    comprobarConexion: comprobarConexion,
    cargarDashboard: cargarDashboard,
    prepararFormulario: prepararFormulario,
    cargarMisRequerimientos: cargarMisRequerimientos,
    cargarGestion: cargarGestion
  });

  global.cargarDatosDashboard = cargarDashboard;
  configurarVistaSharePoint();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciar, { once: true });
  } else {
    iniciar();
  }
})(window);
