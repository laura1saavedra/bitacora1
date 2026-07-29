// ============================================================
// MODELO.JS
// Datos, configuracion y acceso REST a SharePoint.
// ============================================================
(function (global) {
  "use strict";

  const CONFIG = Object.freeze({
    lista: "Backlog",
    listaHistorial: "HistorialActividad",
    bibliotecaArchivos: "ArchivosRequerimientos",
    tamanoMaximoArchivo: 20 * 1024 * 1024,
    cantidadMaximaArchivos: 10,
    sitioAlterno:
      "https://globalhitss.sharepoint.com/sites/AppsColombiaDesarrollo/RetirosDeCesantiasDesarrollo",
    tamanoPagina: 100,
    reintentos: 3
  });

  const CAMPOS = Object.freeze({
    id: {
      interno: "Title",
      tipo: "texto",
      titulos: ["ID REQ", "ID Requerimiento", "Title", "Titulo"]
    },
    app: { interno: "APP", tipo: "texto", titulos: ["APP", "Aplicacion"] },
    tipoServicio: {
      interno: "Tipo_x0020_de_x0020_servicio",
      tipo: "texto",
      titulos: ["Tipo de servicio", "Tipo de Servicio", "TipoServicio"]
    },
    asunto: { interno: "Asunto", tipo: "texto", titulos: ["Asunto"] },
    descripcion: {
      interno: "Descripci_x00f3_n_x0020_de_x0020_la_x0020_solicitud_x0020_",
      tipo: "texto",
      titulos: ["Descripcion de la solicitud", "Descripcion", "Detalle"]
    },
    solicitadoPor: {
      interno: "Solicitado_x0020_por",
      tipo: "texto",
      titulos: ["Solicitado por", "Solicitante"]
    },
    responsable: {
      interno: "Responsable",
      tipo: "persona",
      titulos: ["Responsable"]
    },
    prioridad: {
      interno: "Prioridad",
      tipo: "texto",
      titulos: ["Prioridad"]
    },
    estado: { interno: "Estado", tipo: "texto", titulos: ["Estado"] },
    comentarios: {
      interno: "Comentarios",
      tipo: "texto",
      titulos: ["Comentarios", "Observaciones"]
    },
    casoOrigen: {
      interno: "Caso_x0020_Origen",
      tipo: "texto",
      titulos: ["Caso Origen", "Caso de origen"]
    },
    mentor: { interno: "Mentor", tipo: "persona", titulos: ["Mentor"] },
    fechaEntrega: {
      interno: "Fecha_x0020_Entrega",
      tipo: "texto",
      titulos: ["F.E Entrega", "Fecha Entrega", "Fecha de entrega"]
    },
    complejidad: {
      interno: "Complejidad_x0020_",
      tipo: "texto",
      titulos: ["Complejidad"]
    },
    fechaSolicitud: {
      interno: "Fecha_x0020_Solicitud",
      tipo: "texto",
      titulos: ["Fecha Solicitud", "Fecha de solicitud"]
    },
    fechaPAP: {
      interno: "Fecha_x0020_PAP",
      tipo: "texto",
      titulos: ["F.E PAP", "Fecha PAP", "Fecha de PAP"]
    },
    fechaCierre: {
      interno: "Fecha_x0020_Cierre",
      tipo: "texto",
      titulos: ["F.E Cierre", "Fecha Cierre", "Fecha de cierre"]
    }
  });

  const ODATA = {
    Accept: "application/json;odata=verbose"
  };

  let tipoEntidadLista = null;
  let tipoEntidadHistorial = null;
  let esquemaLista = null;
  let ultimoDiagnosticoEscritura = null;

  function urlSitio() {
    const contexto = global._spPageContextInfo;
    return contexto && contexto.webAbsoluteUrl
      ? contexto.webAbsoluteUrl
      : CONFIG.sitioAlterno;
  }

  function rutaRelativaSitio() {
    const contexto = global._spPageContextInfo;
    if (contexto && contexto.webServerRelativeUrl) {
      return String(contexto.webServerRelativeUrl).replace(/\/$/, "");
    }
    return new URL(urlSitio()).pathname.replace(/\/$/, "");
  }

  function literalOData(valor) {
    return String(valor).replace(/'/g, "''");
  }

  function nombreSeguroCarpeta(valor) {
    return String(valor || "")
      .trim()
      .replace(/[^a-zA-Z0-9._-]/g, "_");
  }

  function nombreSeguroSegmento(valor) {
    return String(valor || "")
      .trim()
      .replace(/["*:<>?/\\|#%]/g, "_");
  }

  function rutaCarpetaRequerimiento(requerimiento) {
    const nombreCarpeta = nombreSeguroCarpeta(
      requerimiento.id + "_" + requerimiento.spItemId
    );
    return (
      rutaRelativaSitio() +
      "/" +
      CONFIG.bibliotecaArchivos +
      "/" +
      nombreCarpeta
    );
  }

  function estaEnSharePoint() {
    return Boolean(
      global._spPageContextInfo &&
      global._spPageContextInfo.webAbsoluteUrl
    );
  }

  function usuarioActual() {
    const contexto = global._spPageContextInfo || {};
    return {
      nombre:
        contexto.userDisplayName ||
        contexto.userLoginName ||
        contexto.userEmail ||
        "Usuario Microsoft 365",
      correo: contexto.userEmail || contexto.userLoginName || ""
    };
  }

  function endpointLista() {
    const tituloSeguro = CONFIG.lista.replace(/'/g, "''");
    return (
      urlSitio() +
      "/_api/web/lists/getbytitle('" +
      tituloSeguro +
      "')"
    );
  }

  function endpointHistorial() {
    const tituloSeguro = CONFIG.listaHistorial.replace(/'/g, "''");
    return (
      urlSitio() +
      "/_api/web/lists/getbytitle('" +
      tituloSeguro +
      "')"
    );
  }

  function endpointBibliotecaArchivos() {
    const tituloSeguro = CONFIG.bibliotecaArchivos.replace(/'/g, "''");
    return (
      urlSitio() +
      "/_api/web/lists/getbytitle('" +
      tituloSeguro +
      "')"
    );
  }

  async function obtenerTiposDocumento() {
    const respuesta = await solicitar(
      endpointBibliotecaArchivos() +
        "/fields/getbyinternalnameortitle('TipoDocumento')" +
        "?$select=Choices",
      {
        method: "GET",
        credentials: "same-origin",
        headers: ODATA
      }
    );
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudieron consultar las opciones de TipoDocumento (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }

    const datos = await respuesta.json();
    const opciones =
      datos.d &&
      datos.d.Choices &&
      Array.isArray(datos.d.Choices.results)
        ? datos.d.Choices.results
        : [];
    return opciones.filter(Boolean).map(String);
  }

  function seleccionarCampos() {
    const seleccion = ["Id"];
    Object.keys(CAMPOS).forEach(function (clave) {
      const campo = CAMPOS[clave];
      seleccion.push(
        campo.tipo === "persona"
          ? campo.interno + "/Title"
          : campo.interno
      );
    });
    return seleccion.join(",");
  }

  function expandirCampos() {
    return Object.keys(CAMPOS)
      .map(function (clave) {
        return CAMPOS[clave];
      })
      .filter(function (campo) {
        return campo.tipo === "persona";
      })
      .map(function (campo) {
        return campo.interno;
      })
      .filter(function (valor, indice, arreglo) {
        return arreglo.indexOf(valor) === indice;
      })
      .join(",");
  }

  function parametrosLectura() {
    const seleccion = [
      "*",
      "AttachmentFiles/FileName",
      "AttachmentFiles/ServerRelativeUrl"
    ];
    const expansion = ["AttachmentFiles"];

    Object.keys(CAMPOS).forEach(function (clave) {
      const campo = resolverCampo(clave);
      if (campo && campo.tipo.indexOf("User") === 0) {
        seleccion.push(campo.nombreInterno + "/Title");
        expansion.push(campo.nombreInterno);
      }
    });

    return (
      "$select=" +
      seleccion.filter(function (valor, indice, arreglo) {
        return arreglo.indexOf(valor) === indice;
      }).join(",") +
      "&$expand=" +
      expansion.filter(function (valor, indice, arreglo) {
        return arreglo.indexOf(valor) === indice;
      }).join(",")
    );
  }

  function desdeSharePoint(item) {
    const archivos =
      item.AttachmentFiles && Array.isArray(item.AttachmentFiles.results)
        ? item.AttachmentFiles.results
        : [];
    const resultado = {
      spItemId: item.Id,
      archivosAdjuntos: archivos.map(function (archivo) {
        return {
          nombre: archivo.FileName || "",
          url: archivo.ServerRelativeUrl || ""
        };
      })
    };
    Object.keys(CAMPOS).forEach(function (clave) {
      const configuracion = CAMPOS[clave];
      const campoReal = resolverCampo(clave);
      const nombreInterno = campoReal
        ? campoReal.nombreInterno
        : configuracion.interno;
      const valor = item[nombreInterno];
      const esPersona =
        campoReal
          ? campoReal.tipo.indexOf("User") === 0
          : configuracion.tipo === "persona";
      if (esPersona) {
        const idPersona = item[nombreInterno + "Id"];
        resultado[clave] =
          valor && valor.Title
            ? valor.Title
            : valor != null && typeof valor !== "object"
              ? valor
              : idPersona != null
                ? "Usuario #" + idPersona
                : "";
      } else {
        resultado[clave] = valor == null ? "" : valor;
      }
    });
    return resultado;
  }

  function normalizarNombre(valor) {
    return String(valor || "")
      .toLowerCase()
      .replace(/[\u00e1\u00e0\u00e4\u00e2]/g, "a")
      .replace(/[\u00e9\u00e8\u00eb\u00ea]/g, "e")
      .replace(/[\u00ed\u00ec\u00ef\u00ee]/g, "i")
      .replace(/[\u00f3\u00f2\u00f6\u00f4]/g, "o")
      .replace(/[\u00fa\u00f9\u00fc\u00fb]/g, "u")
      .replace(/\u00f1/g, "n")
      .replace(/[^a-z0-9]/g, "");
  }

  function resolverCampo(clave) {
    const configuracion = CAMPOS[clave];
    if (!configuracion || !esquemaLista) {
      return null;
    }

    const coincidenciaInterna = esquemaLista.find(function (campo) {
      return campo.nombreInterno === configuracion.interno;
    });
    if (coincidenciaInterna) {
      return coincidenciaInterna;
    }

    const alternativas = [configuracion.interno]
      .concat(configuracion.titulos || [])
      .map(normalizarNombre);
    return esquemaLista.find(function (campo) {
      return (
        alternativas.indexOf(normalizarNombre(campo.titulo)) !== -1 ||
        alternativas.indexOf(normalizarNombre(campo.nombreInterno)) !== -1
      );
    }) || null;
  }

  function fechaIso(valor) {
    if (valor == null || valor === "") {
      return null;
    }
    const texto = String(valor).trim();
    const fechaColombia = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(texto);
    if (fechaColombia) {
      return new Date(
        Date.UTC(
          Number(fechaColombia[3]),
          Number(fechaColombia[2]) - 1,
          Number(fechaColombia[1])
        )
      ).toISOString();
    }
    const fecha = new Date(texto);
    return isNaN(fecha.getTime()) ? null : fecha.toISOString();
  }

  function convertirValor(valor, campo) {
    const tipo = campo.tipo || "";
    if (valor == null || valor === "") {
      return null;
    }
    if (tipo === "DateTime") {
      return fechaIso(valor);
    }
    if (tipo === "Number" || tipo === "Currency") {
      const numero = Number(valor);
      return isNaN(numero) ? null : numero;
    }
    if (tipo === "Integer" || tipo === "Counter") {
      const entero = parseInt(valor, 10);
      return isNaN(entero) ? null : entero;
    }
    if (tipo === "Boolean") {
      return valor === true || valor === 1 || String(valor).toLowerCase() === "true";
    }
    if (tipo === "MultiChoice") {
      return {
        __metadata: { type: "Collection(Edm.String)" },
        results: (Array.isArray(valor) ? valor : [valor]).map(String)
      };
    }
    if (
      tipo === "Text" ||
      tipo === "Note" ||
      tipo === "Choice" ||
      tipo === "Guid"
    ) {
      return String(valor);
    }
    return typeof valor === "object" ? valor : String(valor);
  }

  function haciaSharePoint(item) {
    const resultado = {};
    const omitidos = [];
    Object.keys(CAMPOS).forEach(function (clave) {
      if (item[clave] === undefined) {
        return;
      }

      const campo = resolverCampo(clave);
      if (!campo || campo.soloLectura) {
        omitidos.push(clave);
        return;
      }

      if (campo.tipo.indexOf("User") === 0) {
        if (item[clave] === null) {
          resultado[campo.nombreInterno + "Id"] = null;
          return;
        }
        const idUsuario =
          clave === "solicitadoPor" &&
          global._spPageContextInfo &&
          global._spPageContextInfo.userId
            ? Number(global._spPageContextInfo.userId)
            : Number(item[clave]);
        if (!isNaN(idUsuario) && idUsuario > 0) {
          resultado[campo.nombreInterno + "Id"] = idUsuario;
        } else {
          omitidos.push(clave);
        }
        return;
      }


      if (campo.tipo.indexOf("Lookup") === 0) {
        const idBusqueda = Number(item[clave]);
        if (!isNaN(idBusqueda) && idBusqueda > 0) {
          resultado[campo.nombreInterno + "Id"] = idBusqueda;
        } else {
          omitidos.push(clave);
        }
        return;
      }

      const valor = convertirValor(item[clave], campo);
      if (valor === null && item[clave] !== null && item[clave] !== "") {
        omitidos.push(clave);
      } else {
        resultado[campo.nombreInterno] = valor;
      }
    });

    ultimoDiagnosticoEscritura = {
      enviados: Object.keys(resultado),
      tiposEnviados: Object.keys(resultado).map(function (nombreInterno) {
        const campo = esquemaLista.find(function (itemEsquema) {
          return itemEsquema.nombreInterno === nombreInterno.replace(/Id$/, "");
        });
        return {
          columna: nombreInterno,
          tipoSharePoint: campo ? campo.tipo : "Desconocido",
          tipoJavaScript: typeof resultado[nombreInterno]
        };
      }),
      omitidos: omitidos,
      esquemaConsultado: Boolean(esquemaLista)
    };
    global.BitacoraDiagnosticoEscritura = ultimoDiagnosticoEscritura;
    return resultado;
  }

  function esperar(ms) {
    return new Promise(function (resolver) {
      setTimeout(resolver, ms);
    });
  }

  async function solicitar(url, opciones, intento) {
    const numeroIntento = intento || 1;
    const respuesta = await fetch(url, opciones);
    if (
      (respuesta.status === 429 || respuesta.status === 503) &&
      numeroIntento <= CONFIG.reintentos
    ) {
      await esperar(500 * Math.pow(2, numeroIntento - 1));
      return solicitar(url, opciones, numeroIntento + 1);
    }
    return respuesta;
  }

  async function detalleErrorSharePoint(respuesta) {
    try {
      if (!respuesta.clone) {
        return "";
      }
      const datos = await respuesta.clone().json();
      const error = datos.error || datos["odata.error"];
      if (!error || !error.message) {
        return "";
      }
      return typeof error.message === "string"
        ? error.message
        : error.message.value || "";
    } catch (errorLectura) {
      return "";
    }
  }

  async function digest() {
    const campo = document.getElementById("__REQUESTDIGEST");
    if (campo && campo.value) {
      return campo.value;
    }

    const respuesta = await solicitar(
      urlSitio() + "/_api/contextinfo",
      {
        method: "POST",
        credentials: "same-origin",
        headers: ODATA
      }
    );
    if (!respuesta.ok) {
      throw new Error(
        "No fue posible obtener el X-RequestDigest (" +
        respuesta.status +
        ")."
      );
    }
    const datos = await respuesta.json();
    return datos.d.GetContextWebInformation.FormDigestValue;
  }
  async function resolverUsuarioPorCorreo(correo) {
    const correoNormalizado = String(correo || "").trim();
    if (!correoNormalizado) {
      return null;
    }
    const valorDigest = await digest();
    const respuesta = await solicitar(
      urlSitio() + "/_api/web/ensureuser",
      {
        method: "POST",
        credentials: "same-origin",
        headers: Object.assign({}, ODATA, {
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": valorDigest
        }),
        body: JSON.stringify({ logonName: correoNormalizado })
      }
    );
    if (!respuesta.ok) {
      throw new Error(
        "No se encontr\u00f3 ning\u00fan usuario de SharePoint con el correo \"" +
        correoNormalizado +
        "\"."
      );
    }
    const datos = await respuesta.json();
    const usuario = datos.d || datos;
    return {
      id: usuario.Id,
      nombre: usuario.Title,
      correo: usuario.Email || correoNormalizado
    };
  }
  async function obtenerTipoEntidad() {
    if (tipoEntidadLista) {
      return tipoEntidadLista;
    }
    const respuesta = await solicitar(
      endpointLista() + "/ListItemEntityTypeFullName",
      {
        method: "GET",
        credentials: "same-origin",
        headers: ODATA
      }
    );
    if (!respuesta.ok) {
      throw new Error(
        "No se pudo consultar el tipo de la lista Backlog (" +
        respuesta.status +
        ")."
      );
    }
    const datos = await respuesta.json();
    tipoEntidadLista = datos.d.ListItemEntityTypeFullName;
    return tipoEntidadLista;
  }

  async function verificarConexion() {
    if (!estaEnSharePoint()) {
      throw new Error(
        "La p\u00e1gina no est\u00e1 ejecut\u00e1ndose dentro del contexto de SharePoint."
      );
    }

    const opciones = {
      method: "GET",
      credentials: "same-origin",
      headers: ODATA
    };
    const resultados = await Promise.all([
      solicitar(urlSitio() + "/_api/web?$select=Title,Url", opciones),
      solicitar(
        endpointLista() +
          "?$select=Title,ItemCount,ListItemEntityTypeFullName",
        opciones
      ),
      solicitar(
        endpointLista() +
          "/fields?$select=Title,InternalName,TypeAsString,Hidden,ReadOnlyField,Required",
        opciones
      ),
      solicitar(
        endpointBibliotecaArchivos() +
          "?$select=Title,ItemCount,RootFolder/ServerRelativeUrl&$expand=RootFolder",
        opciones
      ),
      solicitar(
        endpointBibliotecaArchivos() +
          "/fields/getbyinternalnameortitle('TipoDocumento')" +
          "?$select=Title,InternalName,TypeAsString,Choices",
        opciones
      )
    ]);
    const respuestaSitio = resultados[0];
    const respuestaLista = resultados[1];
    const respuestaCampos = resultados[2];
    const respuestaBiblioteca = resultados[3];
    const respuestaTipoDocumento = resultados[4];

    if (!respuestaSitio.ok) {
      throw new Error(
        "No se pudo acceder al sitio de SharePoint (HTTP " +
          respuestaSitio.status +
          ")."
      );
    }
    if (!respuestaLista.ok) {
      const motivoLista =
        respuestaLista.status === 404
          ? "La lista Backlog no existe con ese nombre."
          : respuestaLista.status === 401 || respuestaLista.status === 403
            ? "El usuario no tiene permiso para leer la lista Backlog."
            : "SharePoint devolvi\u00f3 HTTP " + respuestaLista.status + ".";
      throw new Error(motivoLista);
    }
    if (!respuestaCampos.ok) {
      throw new Error(
        "No se pudo consultar el esquema de Backlog (HTTP " +
          respuestaCampos.status +
          ")."
      );
    }
    if (!respuestaBiblioteca.ok) {
      const motivoBiblioteca =
        respuestaBiblioteca.status === 404
          ? "La biblioteca " +
            CONFIG.bibliotecaArchivos +
            " no existe con ese nombre."
          : respuestaBiblioteca.status === 401 ||
              respuestaBiblioteca.status === 403
            ? "El usuario no tiene permiso para acceder a " +
              CONFIG.bibliotecaArchivos +
              "."
            : "No se pudo validar " +
              CONFIG.bibliotecaArchivos +
              " (HTTP " +
              respuestaBiblioteca.status +
              ").";
      throw new Error(motivoBiblioteca);
    }

    const datosSitio = await respuestaSitio.json();
    const datosLista = await respuestaLista.json();
    const datosCampos = await respuestaCampos.json();
    const datosBiblioteca = await respuestaBiblioteca.json();
    const datosTipoDocumento = respuestaTipoDocumento.ok
      ? await respuestaTipoDocumento.json()
      : null;
    const campos = datosCampos.d.results.map(function (campo) {
      return {
        titulo: campo.Title,
        nombreInterno: campo.InternalName,
        tipo: campo.TypeAsString,
        oculto: campo.Hidden,
        soloLectura: campo.ReadOnlyField,
        obligatorio: campo.Required
      };
    });
    esquemaLista = campos;
    const columnasFaltantes = Object.keys(CAMPOS).filter(function (clave) {
      return !resolverCampo(clave);
    }).map(function (clave) {
      return CAMPOS[clave].interno;
    });
    const columnasPersonaInvalidas = Object.keys(CAMPOS)
      .filter(function (clave) {
        return CAMPOS[clave].tipo === "persona";
      })
      .map(function (clave) {
        return CAMPOS[clave].interno;
      })
      .filter(function (nombre) {
        const claveCampo = Object.keys(CAMPOS).find(function (clave) {
          return CAMPOS[clave].interno === nombre;
        });
        const campo = claveCampo ? resolverCampo(claveCampo) : null;
        return campo && campo.tipo.indexOf("User") !== 0;
      });

    const advertencias = [];
    if (columnasFaltantes.length > 0) {
      advertencias.push(
        "Revisar nombres internos: " + columnasFaltantes.join(", ")
      );
    }
    if (columnasPersonaInvalidas.length > 0) {
      advertencias.push(
        "Revisar columnas Persona: " + columnasPersonaInvalidas.join(", ")
      );
    }
    if (!respuestaTipoDocumento.ok) {
      advertencias.push(
        respuestaTipoDocumento.status === 404
          ? "Crear en " +
            CONFIG.bibliotecaArchivos +
            " la columna de opciones TipoDocumento."
          : "No se pudo validar la columna TipoDocumento (HTTP " +
            respuestaTipoDocumento.status +
            ")."
      );
    } else if (
      datosTipoDocumento.d.TypeAsString !== "Choice" &&
      datosTipoDocumento.d.TypeAsString !== "MultiChoice"
    ) {
      advertencias.push(
        "La columna TipoDocumento debe ser de tipo Elecci\u00f3n."
      );
    }

    tipoEntidadLista = datosLista.d.ListItemEntityTypeFullName;
    return {
      sitio: datosSitio.d,
      lista: datosLista.d,
      biblioteca: datosBiblioteca.d,
      columnaTipoDocumento: datosTipoDocumento
        ? datosTipoDocumento.d
        : null,
      campos: campos,
      columnasFaltantes: columnasFaltantes,
      columnasPersonaInvalidas: columnasPersonaInvalidas,
      advertencias: advertencias,
      urlLista:
        urlSitio() + "/Lists/" + encodeURIComponent(CONFIG.lista) + "/AllItems.aspx",
      urlBiblioteca:
        urlSitio() +
        "/" +
        encodeURIComponent(CONFIG.bibliotecaArchivos) +
        "/Forms/AllItems.aspx"
    };
  }

  async function obtenerTodos() {
    let siguiente =
      endpointLista() +
      "/items?" +
      parametrosLectura() +
      "&$top=" +
      String(CONFIG.tamanoPagina);
    const resultados = [];

    while (siguiente) {
      const respuesta = await solicitar(siguiente, {
        method: "GET",
        credentials: "same-origin",
        headers: ODATA
      });
      if (!respuesta.ok) {
        const detalle = await detalleErrorSharePoint(respuesta);
        throw new Error(
          "No se pudo consultar la lista Backlog (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
        );
      }
      const datos = await respuesta.json();
      Array.prototype.push.apply(
        resultados,
        datos.d.results.map(desdeSharePoint)
      );
      siguiente = datos.d.__next || null;
    }
    return resultados;
  }

  async function obtenerPorId(id) {
    const idSeguro = String(id).replace(/'/g, "''");
    const consulta =
      "?" +
      parametrosLectura() +
      "&$filter=Title eq '" +
      encodeURIComponent(idSeguro) +
      "'&$top=1";

    const respuesta = await solicitar(
      endpointLista() + "/items" + consulta,
      {
        method: "GET",
        credentials: "same-origin",
        headers: ODATA
      }
    );
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo consultar el requerimiento (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    if (!datos.d.results.length) {
      return undefined;
    }

    const requerimiento = desdeSharePoint(datos.d.results[0]);
    try {
      const archivosBiblioteca = await obtenerArchivosRequerimiento(
        requerimiento
      );
      const archivosCombinados = requerimiento.archivosAdjuntos.concat(
        archivosBiblioteca
      );
      requerimiento.archivosAdjuntos = archivosCombinados.filter(function (
        archivo,
        indice
      ) {
        return (
          archivosCombinados.findIndex(function (candidato) {
            return candidato.url === archivo.url;
          }) === indice
        );
      });
    } catch (error) {
      console.warn("Consulta de archivos del requerimiento:", error);
    }
    return requerimiento;
  }

  async function crear(item) {
    if (!esquemaLista) {
      await verificarConexion();
    }
    const valores = await Promise.all([digest(), obtenerTipoEntidad()]);
    const carga = haciaSharePoint(item);
    const campoId = resolverCampo("id");
    if (!campoId || !carga[campoId.nombreInterno]) {
      throw new Error(
        "La columna de identificaci\u00f3n (Title/ID REQ) no se pudo resolver en Backlog."
      );
    }
    carga.__metadata = { type: valores[1] };

    const respuesta = await solicitar(endpointLista() + "/items", {
      method: "POST",
      credentials: "same-origin",
      headers: Object.assign({}, ODATA, {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": valores[0]
      }),
      body: JSON.stringify(carga)
    });
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo crear el requerimiento (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    return desdeSharePoint(datos.d);
  }

  async function asegurarCarpeta(rutaCarpeta) {
    const endpointCarpeta =
      urlSitio() +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      literalOData(rutaCarpeta) +
      "')";
    const existente = await solicitar(endpointCarpeta, {
      method: "GET",
      credentials: "same-origin",
      headers: ODATA
    });

    if (existente.ok) {
      return rutaCarpeta;
    }
    if (existente.status !== 404) {
      const detalleConsulta = await detalleErrorSharePoint(existente);
      throw new Error(
        "No se pudo comprobar la carpeta de archivos (HTTP " +
          existente.status +
          ")" +
          (detalleConsulta ? ": " + detalleConsulta : ".")
      );
    }

    const valorDigest = await digest();
    const respuesta = await solicitar(urlSitio() + "/_api/web/folders", {
      method: "POST",
      credentials: "same-origin",
      headers: Object.assign({}, ODATA, {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": valorDigest
      }),
      body: JSON.stringify({
        __metadata: { type: "SP.Folder" },
        ServerRelativeUrl: rutaCarpeta
      })
    });

    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo crear la carpeta del requerimiento (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }
    return rutaCarpeta;
  }

  async function subirArchivo(rutaCarpeta, archivo) {
    const valorDigest = await digest();
    const endpoint =
      urlSitio() +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      literalOData(rutaCarpeta) +
      "')/Files/Add(url='" +
      literalOData(archivo.name) +
      "',overwrite=false)";
    const respuesta = await solicitar(endpoint, {
      method: "POST",
      credentials: "same-origin",
      headers: Object.assign({}, ODATA, {
        "X-RequestDigest": valorDigest
      }),
      body: archivo
    });

    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo cargar " +
          archivo.name +
          " (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    return {
      nombre: archivo.name,
      url:
        datos.d && datos.d.ServerRelativeUrl
          ? datos.d.ServerRelativeUrl
          : rutaCarpeta + "/" + archivo.name
    };
  }

  async function guardarArchivosRequerimiento(
    requerimiento,
    archivosClasificados
  ) {
    const clasificaciones = Array.prototype.slice.call(
      archivosClasificados || []
    );
    if (!clasificaciones.length) {
      return { cargados: [], errores: [], rutaCarpeta: "" };
    }

    const cargados = [];
    const errores = [];
    const rutasCategorias = {};
    let rutaRequerimiento;

    try {
      rutaRequerimiento = await asegurarCarpeta(
        rutaCarpetaRequerimiento(requerimiento)
      );
    } catch (errorCarpeta) {
      clasificaciones.forEach(function (clasificacion) {
        errores.push({
          nombre: clasificacion.archivo.name,
          mensaje: errorCarpeta.message
        });
      });
      return {
        cargados: cargados,
        errores: errores,
        rutaCarpeta: ""
      };
    }

    for (const clasificacion of clasificaciones) {
      const archivo = clasificacion.archivo;
      const categoria = nombreSeguroSegmento(clasificacion.tipoDocumento);
      try {
        if (!categoria) {
          throw new Error(
            "El archivo " + archivo.name + " no tiene tipo de documento."
          );
        }
        if (!rutasCategorias[categoria]) {
          rutasCategorias[categoria] = await asegurarCarpeta(
            rutaRequerimiento + "/" + categoria
          );
        }
        const archivoCargado = await subirArchivo(
          rutasCategorias[categoria],
          archivo
        );
        archivoCargado.tipoDocumento = clasificacion.tipoDocumento;
        cargados.push(archivoCargado);
      } catch (error) {
        errores.push({
          nombre: archivo.name,
          mensaje: error.message
        });
      }
    }

    return {
      cargados: cargados,
      errores: errores,
      rutaCarpeta: rutaRequerimiento
    };
  }

  async function obtenerArchivosCarpeta(rutaCarpeta, tipoDocumento) {
    const endpoint =
      urlSitio() +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      literalOData(rutaCarpeta) +
      "')/Files?$select=Name,ServerRelativeUrl,TimeLastModified";
    const respuesta = await solicitar(endpoint, {
      method: "GET",
      credentials: "same-origin",
      headers: ODATA
    });

    if (respuesta.status === 404) {
      return [];
    }
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudieron consultar los archivos del requerimiento (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }

    const datos = await respuesta.json();
    return datos.d.results.map(function (archivo) {
      return {
        nombre: archivo.Name || "",
        url: archivo.ServerRelativeUrl || "",
        modificado: archivo.TimeLastModified || "",
        tipoDocumento: tipoDocumento || "Sin clasificaci\u00f3n"
      };
    });
  }

  async function obtenerArchivosRequerimiento(requerimiento) {
    const rutaCarpeta = rutaCarpetaRequerimiento(requerimiento);
    const archivos = await obtenerArchivosCarpeta(
      rutaCarpeta,
      "Sin clasificaci\u00f3n"
    );
    const endpointSubcarpetas =
      urlSitio() +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      literalOData(rutaCarpeta) +
      "')/Folders?$select=Name,ServerRelativeUrl";
    const respuesta = await solicitar(endpointSubcarpetas, {
      method: "GET",
      credentials: "same-origin",
      headers: ODATA
    });

    if (respuesta.status === 404) {
      return archivos;
    }
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudieron consultar las categor\u00edas de documentos (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }

    const datos = await respuesta.json();
    const archivosPorCategoria = await Promise.all(
      datos.d.results.map(function (carpeta) {
        return obtenerArchivosCarpeta(
          carpeta.ServerRelativeUrl,
          carpeta.Name
        );
      })
    );
    archivosPorCategoria.forEach(function (grupo) {
      Array.prototype.push.apply(archivos, grupo);
    });
    return archivos;
  }

  async function obtenerDocumentosRequerimiento(requerimientoOId) {
    const requerimiento =
      typeof requerimientoOId === "object" && requerimientoOId
        ? requerimientoOId
        : await obtenerPorId(requerimientoOId);
    if (!requerimiento) {
      return [];
    }
    return obtenerArchivosRequerimiento(requerimiento);
  }

  async function actualizar(id, cambios) {
    const existente = await obtenerPorId(id);
    if (!existente) {
      throw new Error("No se encontr\u00f3 el requerimiento " + id + ".");
    }

    if (!esquemaLista) {
      await verificarConexion();
    }
    const valores = await Promise.all([digest(), obtenerTipoEntidad()]);
    const carga = haciaSharePoint(cambios);
    carga.__metadata = { type: valores[1] };

    const respuesta = await solicitar(
      endpointLista() + "/items(" + existente.spItemId + ")",
      {
        method: "POST",
        credentials: "same-origin",
        headers: Object.assign({}, ODATA, {
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": valores[0],
          "X-HTTP-Method": "MERGE",
          "IF-MATCH": "*"
        }),
        body: JSON.stringify(carga)
      }
    );
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo actualizar el requerimiento (HTTP " +
        respuesta.status +
        ")" +
        (detalle ? ": " + detalle : ".")
      );
    }
    return obtenerPorId(id);
  }

  async function eliminar(id) {
    const existente = await obtenerPorId(id);
    if (!existente) {
      return;
    }
    const valorDigest = await digest();
    const respuesta = await solicitar(
      endpointLista() + "/items(" + existente.spItemId + ")",
      {
        method: "POST",
        credentials: "same-origin",
        headers: Object.assign({}, ODATA, {
          "Content-Type": "application/json;odata=verbose",
          "X-RequestDigest": valorDigest,
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*"
        })
      }
    );
    if (!respuesta.ok) {
      throw new Error(
        "No se pudo eliminar el requerimiento (" + respuesta.status + ")."
      );
    }
  }

// ============================================================
// BUSCAR USUARIOS MICROSOFT 365
// ------------------------------------------------------------
// Consulta los usuarios disponibles en SharePoint Online.
//
// Recibe:
// texto -> nombre o correo que escribe el usuario.
//
// Devuelve:
// lista de usuarios encontrados.
//
// No modifica listas ni informacion de SharePoint.
// Solo consulta usuarios.
// ============================================================

async function buscarUsuariosMicrosoft(texto) {

    // Si no hay texto suficiente no consulta
    if (!texto || texto.trim().length < 3) {

        return [];

    }


    const url =
        urlSitio() +
        "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";


    const datos = {

        queryParams: {

            __metadata: {
                type:
                "SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters"
            },

            AllowEmailAddresses: true,

            AllowMultipleEntities: true,

            MaximumEntitySuggestions: 10,

            PrincipalSource: 15,

            PrincipalType: 1,

            QueryString: texto

        }

    };


    const respuesta = await solicitar(url, {

        method: "POST",

        credentials: "same-origin",

        headers: Object.assign({}, ODATA, {

            "Content-Type":
            "application/json;odata=verbose"

        }),

        body: JSON.stringify(datos)

    });


    if (!respuesta.ok) {

        throw new Error(
            "No fue posible consultar usuarios de Microsoft 365."
        );

    }


    const resultado = await respuesta.json();


    const usuarios =
        JSON.parse(resultado.d.ClientPeoplePickerSearchUser);


    return usuarios.map(function(usuario){

        return {

            nombre: usuario.DisplayText,

            correo: usuario.EntityData?.Email || ""

        };

    });


}


  async function obtenerTipoEntidadHistorial() {
    if (tipoEntidadHistorial) {
      return tipoEntidadHistorial;
    }
    const respuesta = await solicitar(
      endpointHistorial() + "/ListItemEntityTypeFullName",
      {
        method: "GET",
        credentials: "same-origin",
        headers: ODATA
      }
    );
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        respuesta.status === 404
          ? "La lista " + CONFIG.listaHistorial + " no existe en SharePoint."
          : "No se pudo consultar la lista de historial (HTTP " +
            respuesta.status +
            ")" +
            (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    tipoEntidadHistorial = datos.d.ListItemEntityTypeFullName;
    return tipoEntidadHistorial;
  }

  function desdeActividadSharePoint(item) {
    return {
      id: item.Id,
      usuario: item.Usuario || "",
      correo: item.Correo || "",
      accion: item.Accion || item.Title || "",
      requerimiento: item.Requerimiento || "",
      tipo: item.TipoActividad || "General",
      fecha: item.Created || ""
    };
  }

  async function obtenerBitacora() {
    const consulta =
      "/items?$select=Id,Title,Usuario,Correo,Accion,Requerimiento," +
      "TipoActividad,Created&$orderby=Created desc&$top=500";
    const respuesta = await solicitar(endpointHistorial() + consulta, {
      method: "GET",
      credentials: "same-origin",
      headers: ODATA
    });
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        respuesta.status === 404
          ? "La lista " + CONFIG.listaHistorial + " no existe en SharePoint."
          : "No se pudo consultar el historial (HTTP " +
            respuesta.status +
            ")" +
            (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    return datos.d.results.map(desdeActividadSharePoint);
  }

  async function agregarActividad(usuario, accion, detalles) {
    const informacion = detalles || {};
    const usuarioSesion = usuarioActual();
    const valores = await Promise.all([
      digest(),
      obtenerTipoEntidadHistorial()
    ]);
    const textoAccion = String(accion || "").trim();
    const carga = {
      __metadata: { type: valores[1] },
      Title: textoAccion.substring(0, 250) || "Actividad de la plataforma",
      Usuario: String(usuario || usuarioSesion.nombre || "").substring(0, 250),
      Correo: String(usuarioSesion.correo || "").substring(0, 250),
      Accion: textoAccion,
      Requerimiento: String(informacion.requerimiento || "").substring(0, 250),
      TipoActividad: String(informacion.tipo || "General").substring(0, 250)
    };
    const respuesta = await solicitar(endpointHistorial() + "/items", {
      method: "POST",
      credentials: "same-origin",
      headers: Object.assign({}, ODATA, {
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": valores[0]
      }),
      body: JSON.stringify(carga)
    });
    if (!respuesta.ok) {
      const detalle = await detalleErrorSharePoint(respuesta);
      throw new Error(
        "No se pudo registrar la actividad (HTTP " +
          respuesta.status +
          ")" +
          (detalle ? ": " + detalle : ".")
      );
    }
    const datos = await respuesta.json();
    return desdeActividadSharePoint(datos.d);
  }

  global.Modelo = Object.freeze({
    configuracion: CONFIG,
    urlSitio: urlSitio,
    estaEnSharePoint: estaEnSharePoint,
    usuarioActual: usuarioActual,
    verificarConexion: verificarConexion,
    obtenerTiposDocumento: obtenerTiposDocumento,
    obtenerTodos: obtenerTodos,
    obtenerPorId: obtenerPorId,
    crear: crear,
    guardarArchivosRequerimiento: guardarArchivosRequerimiento,
    obtenerArchivosRequerimiento: obtenerArchivosRequerimiento,
    obtenerDocumentosRequerimiento: obtenerDocumentosRequerimiento,
    buscarUsuariosMicrosoft: buscarUsuariosMicrosoft,
    actualizar: actualizar,
    resolverUsuarioPorCorreo: resolverUsuarioPorCorreo,
    eliminar: eliminar,
    obtenerBitacora: obtenerBitacora,
    agregarActividad: agregarActividad,
    obtenerDiagnosticoEscritura: function () {
      return ultimoDiagnosticoEscritura;
    }
  });
})(window);
