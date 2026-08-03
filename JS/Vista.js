// ============================================================
// VISTA.JS
// Navegacion, renderizado y lectura de la interfaz.
// No consulta SharePoint directamente: comunica las acciones al Controlador.
// ============================================================
(function (global) {
  "use strict";

  // Relacion estable entre las rutas logicas y sus contenedores del DOM.
  const VISTAS = {
    dashboard: "view-dashboard",
    crear: "view-crear-requerimiento",
    mis: "view-mis-requerimientos",
    gestion: "view-gestion",
    historial: "view-historial",
     indicadores: "view-indicadores"
  };
  // SVG en linea para evitar dependencias externas y heredar el color del tema.
  const ICONO_OJO =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"></path>' +
    '<circle cx="12" cy="12" r="2.5"></circle></svg>';
  const ICONO_LAPIZ =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="m4 20 4.3-1.1L19 8.2 15.8 5 5.1 15.7 4 20Z"></path>' +
    '<path d="m14.5 6.3 3.2 3.2"></path></svg>';
  const ICONO_DESCARGA =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="M12 3v12"></path>' +
    '<path d="m7 10 5 5 5-5"></path>' +
    '<path d="M5 20h14"></path></svg>';
  const ICONO_PAPELERA =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="M4 7h16"></path>' +
    '<path d="M9 7V4h6v3"></path>' +
    '<path d="m6 7 1 13h10l1-13"></path>' +
    '<path d="M10 11v5M14 11v5"></path></svg>';
  const ICONO_CHECK =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="m4 12.5 5 5L20 6.5"></path></svg>';
  const ICONO_X =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="m5 5 14 14"></path>' +
    '<path d="m19 5-14 14"></path></svg>';

  const ADJUNTOS_POR_PAGINA = 4;
  let archivosAdjuntosDetalle = [];
  let paginaAdjuntosDetalle = 1;
  let esAdministradorSesion = false;

  // Convierte estados funcionales en variantes visuales conocidas por el CSS.
  function claseEstado(estado) {

    const mapa = {
      "pendiente": "orange",
      "en proceso": "orange",
      "pruebas": "teal",
      "en pruebas": "teal",
      "esperando documentaci\u00f3n usuario": "indigo",
      "esperando documentacion usuario": "indigo",
      "esperando cierre usuario": "purple",
      "finalizados": "green",
      "finalizado": "green",
      "cancelado": "danger"
    };
    const clave = String(estado || "").trim().toLowerCase();
    return mapa[clave] || "neutral";
  }

  // Todo valor dinamico insertado mediante innerHTML debe pasar por esta funcion.
  function textoSeguro(valor) {
    return String(valor == null ? "" : valor)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function formatearFecha(valor) {
    if (!valor) {
      return "";
    }

    const texto = String(valor).trim();
    const fechaISO = texto.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (fechaISO) {
      return fechaISO[3] + "/" + fechaISO[2] + "/" + fechaISO[1];
    }

    const fechaLatina = texto.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (fechaLatina) {
      return texto;
    }

    const fecha = new Date(texto);
    if (isNaN(fecha.getTime())) {
      return texto;
    }
    return String(fecha.getUTCDate()).padStart(2, "0") +
      "/" +
      String(fecha.getUTCMonth() + 1).padStart(2, "0") +
      "/" +
      fecha.getUTCFullYear();
  }

  // --------------------------------------------------------------------------
  // Navegacion entre vistas principales
  // --------------------------------------------------------------------------
  function mostrarVista(nombre) {
    Object.keys(VISTAS).forEach(function (clave) {
      const elemento = document.getElementById(VISTAS[clave]);
      if (elemento) {
        elemento.hidden = clave !== nombre;
      }
    });
  }

  function mostrarDashboard() {
    mostrarVista("dashboard");
    global.Controlador.cargarDashboard();
  }

  function mostrarFormularioCrear() {
    mostrarVista("crear");
    global.Controlador.prepararFormulario();
  }

  function mostrarFormularioEdicion() {
    mostrarVista("crear");
  }

  function mostrarMisRequerimientos() {
    mostrarVista("mis");
    global.Controlador.cargarMisRequerimientos();
  }

  function mostrarGestion() {
    mostrarVista("gestion");
    global.Controlador.cargarGestion();
  }

  function mostrarHistorial() {
    mostrarVista("historial");
    global.Controlador.cargarHistorial();
  }

 function mostrarIndicadores(){

    mostrarVista("indicadores");


    setTimeout(function(){

        global.Controlador.cargarIndicadores();

    },300);

}

  function mostrarConexion(estado, detalle) {
    const indicador = document.getElementById("estado-conexion");
    if (!indicador) {
      return;
    }
    indicador.classList.remove("conexion-ok", "conexion-error");
    indicador.classList.add(estado ? "conexion-ok" : "conexion-error");
    indicador.innerHTML =
      '<span class="status-dot" aria-hidden="true"></span>' +
      textoSeguro(detalle);
  }

  function mostrarUsuario(usuario) {
    const nombre = document.getElementById("usuario-sesion");
    const rol = document.getElementById("rol-sesion");
    if (nombre) {
      nombre.textContent = usuario.nombre;
    }
    if (rol) {
      rol.textContent = usuario.correo || "Sesi\u00f3n de SharePoint";
    }
  }

  function configurarAcceso(esAdministrador) {
    esAdministradorSesion = Boolean(esAdministrador);
    const rol = document.getElementById("rol-sesion");
    if (rol) {
      rol.textContent = esAdministradorSesion
        ? "Administrador"
        : "Cliente";
    }
  }

  function botonEditar(req) {
    if (!esAdministradorSesion) {
      return "";
    }
    return (
      '<button class="action-btn icon-action edit-btn" data-id="' +
      textoSeguro(req.id) +
      '" type="button" aria-label="Editar requerimiento ' +
      textoSeguro(req.id) +
      '" title="Editar">' + ICONO_LAPIZ + "</button>"
    );
  }

  // --------------------------------------------------------------------------
  // Tableros, tablas y paginacion
  // --------------------------------------------------------------------------
  function estadoCoincide(valorEstado, variantes) {
    // Compara ignorando mayusculas/minusculas y espacios de sobra, para no
    // depender de que el texto en CatalogoEstados sea IDENTICO caracter por
    // caracter al escrito aqui.
    const valor = String(valorEstado || "").trim().toLowerCase();
    return variantes.some(function (variante) {
      return valor === variante.toLowerCase();
    });
  }

  function renderizarTarjetas(datos) {
    const valores = {
      "total-requerimientos": datos.length,
      "total-pendientes": datos.filter(function (item) {
        return estadoCoincide(item.estado, ["Pendiente"]);
      }).length,
      "total-en-pruebas": datos.filter(function (item) {
        return estadoCoincide(item.estado, ["Pruebas", "En pruebas"]);
      }).length,
      "total-pruebas": datos.filter(function (item) {
        return estadoCoincide(item.estado, [
          "Esperando documentaci\u00f3n usuario",
          // Variante heredada sin tilde que puede existir en SharePoint.
          "Esperando Documentacion Usuario"
        ]);
      }).length,
      "total-cierre-usuario": datos.filter(function (item) {
        return estadoCoincide(item.estado, ["Esperando cierre usuario"]);
      }).length,
      "total-finalizados": datos.filter(function (item) {
        return estadoCoincide(item.estado, ["Finalizados", "Finalizado"]);
      }).length
    };
    Object.keys(valores).forEach(function (id) {
      const elemento = document.getElementById(id);
      if (elemento) {
        elemento.textContent = valores[id];
      }
    });
  }

function renderizarTabla(datos, paginacion) {
    const cuerpo = document.getElementById("tabla-requerimientos");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            return (
              "<tr>" +
              '<td class="req-id" data-label="ID req.">' + textoSeguro(req.id) + "</td>" +
              '<td data-label="APP">' + textoSeguro(req.app) + "</td>" +
              '<td data-label="Tipo de servicio">' + textoSeguro(req.tipoServicio) + "</td>" +
              '<td data-label="Asunto">' + textoSeguro(req.asunto) + "</td>" +
              '<td data-label="Solicitado por">' + textoSeguro(req.solicitadoPor) + "</td>" +
              '<td data-label="Responsable">' + textoSeguro(req.responsable || "No asignado") + "</td>" +
              '<td data-label="Prioridad">' + textoSeguro(req.prioridad) + "</td>" +
              '<td data-label="Estado"><span class="estado estado-' +
              claseEstado(req.estado) +
              '">' + textoSeguro(req.estado) + "</span></td>" +
              '<td data-label="F. solicitud">' +
              textoSeguro(formatearFecha(req.fechaSolicitud) || "Sin definir") +
              "</td>" +
              '<td data-label="F. cierre">' +
              textoSeguro(formatearFecha(req.fechaCierre) || "Sin definir") +
              "</td>" +
              '<td data-label="Acciones"><div class="request-action-buttons">' +
              '<button class="action-btn icon-action view-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Ver requerimiento ' +
              textoSeguro(req.id) +
              '" title="Ver detalle">' + ICONO_OJO + "</button>" +
              botonEditar(req) +
              "</div></td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="11">No hay requerimientos que coincidan con la b\u00fasqueda.</td></tr>';
    if (paginacion) {
      renderizarPaginacionBacklog(paginacion);
    }
  }

  function renderizarPaginacionBacklog(paginacion) {
    const pie = document.getElementById("pie-paginacion-backlog");
    const resumen = document.getElementById("resumen-paginacion");
    const controles = document.getElementById("controles-paginacion");
    const tieneDatos = paginacion && paginacion.total > 0;
    pie.hidden = !tieneDatos;

    if (!tieneDatos) {
      resumen.textContent = "";
      controles.innerHTML = "";
      return;
    }

    resumen.textContent =
      "Mostrando " +
      paginacion.fin +
      " Requerimientos " +
      " de " +
      paginacion.total;

    const anterior =
      '<button type="button" data-pagina="' +
      (paginacion.pagina - 1) +
      '" aria-label="P\u00e1gina anterior"' +
      (paginacion.pagina === 1 ? " disabled" : "") +
      ">&#8249;</button>";
    const paginas = paginasPaginacion(
      paginacion.pagina,
      paginacion.totalPaginas
    )
      .map(function (pagina) {
        if (pagina === "...") {
          return '<span class="pagination-ellipsis" aria-hidden="true">&hellip;</span>';
        }
        const activa = pagina === paginacion.pagina;
        return (
          '<button type="button" data-pagina="' +
          pagina +
          '" aria-label="Ir a la p\u00e1gina ' +
          pagina +
          '"' +
          (activa ? ' class="active" aria-current="page"' : "") +
          ">" +
          pagina +
          "</button>"
        );
      })
      .join("");
    const siguiente =
      '<button type="button" data-pagina="' +
      (paginacion.pagina + 1) +
      '" aria-label="P\u00e1gina siguiente"' +
      (paginacion.pagina === paginacion.totalPaginas ? " disabled" : "") +
      ">&#8250;</button>";
    controles.innerHTML = anterior + paginas + siguiente;
  }

  function paginasPaginacion(paginaActual, totalPaginas) {
    if (totalPaginas <= 7) {
      return Array.from({ length: totalPaginas }, function (_, indice) {
        return indice + 1;
      });
    }

    const paginas = [1];
    const inicio = Math.max(2, paginaActual - 1);
    const fin = Math.min(totalPaginas - 1, paginaActual + 1);
    if (inicio > 2) {
      paginas.push("...");
    }
    for (let pagina = inicio; pagina <= fin; pagina += 1) {
      paginas.push(pagina);
    }
    if (fin < totalPaginas - 1) {
      paginas.push("...");
    }
    paginas.push(totalPaginas);
    return paginas;
  }

  function renderizarPaginacionMisRequerimientos(paginacion) {
    const pie = document.getElementById("paginacion-mis-requerimientos");
    const resumen = document.getElementById("resumen-paginacion-mis");
    const controles = document.getElementById("controles-paginacion-mis");
    const tieneDatos = paginacion && paginacion.total > 0;
    pie.hidden = !tieneDatos;

    if (!tieneDatos) {
      resumen.textContent = "";
      controles.innerHTML = "";
      return;
    }

    resumen.textContent =
      "Mostrando " +
      paginacion.fin +
      " Requerimientos " +
      " de " +
      paginacion.total;

    const anterior =
      '<button type="button" data-pagina="' +
      (paginacion.pagina - 1) +
      '" aria-label="P\u00e1gina anterior"' +
      (paginacion.pagina === 1 ? " disabled" : "") +
      ">&#8249;</button>";
    const paginas = paginasPaginacion(
      paginacion.pagina,
      paginacion.totalPaginas
    )
      .map(function (pagina) {
        if (pagina === "...") {
          return '<span class="pagination-ellipsis" aria-hidden="true">\u2026</span>';
        }
        const activa = pagina === paginacion.pagina;
        return (
          '<button type="button" data-pagina="' +
          pagina +
          '" aria-label="Ir a la p\u00e1gina ' +
          pagina +
          '"' +
          (activa ? ' class="active" aria-current="page"' : "") +
          ">" +
          pagina +
          "</button>"
        );
      })
      .join("");
    const siguiente =
      '<button type="button" data-pagina="' +
      (paginacion.pagina + 1) +
      '" aria-label="P\u00e1gina siguiente"' +
      (paginacion.pagina === paginacion.totalPaginas ? " disabled" : "") +
      ">&#8250;</button>";
    controles.innerHTML = anterior + paginas + siguiente;
  }

  function renderizarMisRequerimientos(datos, paginacion) {
    const cuerpo = document.getElementById("tabla-mis-requerimientos");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            return (
              '<tr><td data-label="ID req.">' + textoSeguro(req.id) + "</td>" +
              '<td data-label="APP">' + textoSeguro(req.app) + "</td>" +
              '<td data-label="Tipo de servicio">' + textoSeguro(req.tipoServicio) + "</td>" +
              '<td data-label="Asunto">' + textoSeguro(req.asunto) + "</td>" +
              '<td class="request-description" data-label="Descripci\u00f3n" title="' +
              textoSeguro(req.descripcion) +
              '">' + textoSeguro(req.descripcion || "Sin descripci\u00f3n") + "</td>" +
              '<td data-label="Solicitado por">' + textoSeguro(req.solicitadoPor) + "</td>" +
              '<td data-label="Responsable">' + textoSeguro(req.responsable || "No asignado") + "</td>" +
              '<td data-label="Estado">' + textoSeguro(req.estado) + "</td>" +
              '<td class="request-actions" data-label="Acciones">' +
              '<div class="request-action-buttons">' +
              '<button class="action-btn icon-action view-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Ver requerimiento ' +
              textoSeguro(req.id) +
              '" title="Ver detalle">' + ICONO_OJO + "</button>" +
              botonEditar(req) +
              "</div></td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="9">' +
        (paginacion && paginacion.filtrosActivos
          ? "No hay requerimientos que coincidan con los filtros."
          : cuerpo.dataset.mensajeVacio || "No tienes requerimientos registrados.") +
        "</td></tr>";
    renderizarPaginacionMisRequerimientos(paginacion);
  }

  function activarPestanaMisRequerimientos(tipo) {
    const textos = {
      responsable: "Requerimientos asignados al usuario autenticado como responsable.",
      mentor: "Requerimientos asignados al usuario autenticado como mentor.",
      creados: "Requerimientos creados por el usuario autenticado."
    };
    const vacios = {
      responsable: "No tienes requerimientos asignados como responsable.",
      mentor: "No tienes requerimientos asignados como mentor.",
      creados: "No tienes requerimientos creados."
    };
    const panel = document.getElementById("panel-mis-requerimientos");
    document.querySelectorAll(".my-requests-tab").forEach(function (pestana) {
      const activa = pestana.dataset.tipoMis === tipo;
      pestana.classList.toggle("active", activa);
      pestana.setAttribute("aria-selected", activa ? "true" : "false");
      pestana.tabIndex = activa ? 0 : -1;
      if (activa && panel) {
        panel.setAttribute("aria-labelledby", pestana.id);
      }
    });
    document.getElementById("descripcion-mis-requerimientos").textContent =
      textos[tipo] || textos.responsable;
    const cuerpo = document.getElementById("tabla-mis-requerimientos");
    if (cuerpo) {
      cuerpo.dataset.mensajeVacio = vacios[tipo] || vacios.responsable;
    }
  }

function renderizarGestion(datos, paginacion) {
    const cuerpo = document.getElementById("tabla-gestion");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            const comentarioActual = String(req.comentarios || "");
            return (
              '<tr data-id="' + textoSeguro(req.id) + '">' +
              '<td data-label="ID req.">' + textoSeguro(req.id) + "</td>" +
              '<td data-label="Asunto">' + textoSeguro(req.asunto) + "</td>" +
              '<td data-label="Solicitado por">' + textoSeguro(req.solicitadoPor) + "</td>" +
              '<td data-label="Responsable"><div class="gestion-campo-autocompletar">' +
              '<input class="gestion-responsable" type="text" value="' +
              textoSeguro(req.responsable) +
              '" data-original="' +
              textoSeguro(req.responsable) +
              '" placeholder="Correo del responsable" aria-label="Responsable del requerimiento ' +
              textoSeguro(req.id) +
              '" readonly>' +
              '<div class="gestion-sugerencias" hidden></div></div></td>' +
              '<td data-label="Mentor"><div class="gestion-campo-autocompletar">' +
              '<input class="gestion-mentor" type="text" value="' +
              textoSeguro(req.mentor) +
              '" data-original="' +
              textoSeguro(req.mentor) +
              '" placeholder="Correo del mentor" aria-label="Mentor del requerimiento ' +
              textoSeguro(req.id) +
              '" readonly>' +
              '<div class="gestion-sugerencias" hidden></div></div></td>' +
              '<td data-label="F.E. Entrega">' + textoSeguro(req.fechaEntrega || "Sin definir") + "</td>" +
              '<td data-label="Estado"><span class="estado estado-' +
              claseEstado(req.estado) +
              '">' + textoSeguro(req.estado) + "</span></td>" +
              '<td data-label="Observaciones">' +
              '<input class="gestion-comentarios" type="text" value="' +
              textoSeguro(comentarioActual) +
              '" data-original="' +
              textoSeguro(comentarioActual) +
              '" maxlength="500" aria-label="Observaciones del requerimiento ' +
              textoSeguro(req.id) +
              '" readonly>' +
              '<span class="gestion-contador-caracteres">' +
              comentarioActual.length +
              "/500</span></td>" +
              '<td data-label="Acciones"><div class="request-action-buttons">' +
              '<button class="action-btn icon-action view-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Ver requerimiento ' +
              textoSeguro(req.id) +
              '" title="Ver detalle">' + ICONO_OJO + "</button>" +
              '<button class="action-btn icon-action edit-btn-gestion" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Editar requerimiento ' +
              textoSeguro(req.id) +
              '" title="Editar">' + ICONO_LAPIZ + "</button>" +
              '<button class="action-btn icon-action confirm-btn-gestion" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Guardar requerimiento ' +
              textoSeguro(req.id) +
              '" title="Guardar" hidden>' + ICONO_CHECK + "</button>" +
              '<button class="action-btn icon-action cancel-btn-gestion" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Cancelar edici\u00f3n de ' +
              textoSeguro(req.id) +
              '" title="Cancelar" hidden>' + ICONO_X + "</button>" +
              "</div></td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="9">No hay requerimientos para gestionar.</td></tr>';
    if (paginacion) {
      renderizarPaginacionGestion(paginacion);
    }
  }
  function renderizarPaginacionGestion(paginacion) {
    const pie = document.getElementById("pie-paginacion-gestion");
    const resumen = document.getElementById("resumen-paginacion-gestion");
    const controles = document.getElementById("controles-paginacion-gestion");
    const tieneDatos = paginacion && paginacion.total > 0;
    pie.hidden = !tieneDatos;

    if (!tieneDatos) {
      resumen.textContent = "";
      controles.innerHTML = "";
      return;
    }

    resumen.textContent =
      "Mostrando " +
      paginacion.fin +
      " Requerimientos " +
      " de " +
      paginacion.total;

    const anterior =
      '<button type="button" data-pagina="' +
      (paginacion.pagina - 1) +
      '" aria-label="P\u00e1gina anterior"' +
      (paginacion.pagina === 1 ? " disabled" : "") +
      ">&#8249;</button>";
    const paginas = paginasPaginacion(
      paginacion.pagina,
      paginacion.totalPaginas
    )
      .map(function (pagina) {
        if (pagina === "...") {
          return '<span class="pagination-ellipsis" aria-hidden="true">&hellip;</span>';
        }
        const activa = pagina === paginacion.pagina;
        return (
          '<button type="button" data-pagina="' +
          pagina +
          '" aria-label="Ir a la p\u00e1gina ' +
          pagina +
          '"' +
          (activa ? ' class="active" aria-current="page"' : "") +
          ">" +
          pagina +
          "</button>"
        );
      })
      .join("");
    const siguiente =
      '<button type="button" data-pagina="' +
      (paginacion.pagina + 1) +
      '" aria-label="P\u00e1gina siguiente"' +
      (paginacion.pagina === paginacion.totalPaginas ? " disabled" : "") +
      ">&#8250;</button>";
    controles.innerHTML = anterior + paginas + siguiente;
  }

  function mostrarSugerenciasGestion(campo, personas) {
    const contenedor = campo.closest(".gestion-campo-autocompletar");
    const lista = contenedor
      ? contenedor.querySelector(".gestion-sugerencias")
      : null;
    if (!lista) {
      return;
    }
    if (!personas || !personas.length) {
      ocultarSugerenciasGestion(campo);
      return;
    }
    lista.innerHTML = personas
      .map(function (persona) {
        return (
          '<button type="button" class="gestion-sugerencia-item" data-id="' +
          textoSeguro(persona.id) +
          '" data-nombre="' +
          textoSeguro(persona.nombre) +
          '">' +
          textoSeguro(persona.nombre) +
          (persona.correo
            ? '<span class="gestion-sugerencia-correo">' +
              textoSeguro(persona.correo) +
              "</span>"
            : "") +
          "</button>"
        );
      })
      .join("");
    lista.hidden = false;
  }

  function ocultarSugerenciasGestion(campo) {
    const contenedor = campo.closest(".gestion-campo-autocompletar");
    const lista = contenedor
      ? contenedor.querySelector(".gestion-sugerencias")
      : null;
    if (lista) {
      lista.hidden = true;
      lista.innerHTML = "";
    }
  }

  function renderizarFiltros(datos) {
    [
      ["filtro-app", "app", "APP"],
      ["filtro-responsable", "responsable", "Responsable"],
      ["filtro-estado", "estado", "Estado"],
      ["filtro-prioridad", "prioridad", "Prioridad"]
    ].forEach(function (configuracion) {
      const select = document.getElementById(configuracion[0]);
      const valores = datos
        .map(function (item) {
          return item[configuracion[1]];
        })
        .filter(Boolean)
        .filter(function (valor, indice, arreglo) {
          return arreglo.indexOf(valor) === indice;
        });
      select.innerHTML =
        '<option value="">' + configuracion[2] + "</option>" +
        valores
          .map(function (valor) {
            return (
              '<option value="' +
              textoSeguro(valor) +
              '">' +
              textoSeguro(valor) +
              "</option>"
            );
          })
          .join("");
    });
  }

  function renderizarFiltrosMisRequerimientos(datos) {
    [
      ["filtro-app-mis", "app", "Todas las APP"],
      ["filtro-estado-mis", "estado", "Todos los estados"]
    ].forEach(function (configuracion) {
      const select = document.getElementById(configuracion[0]);
      const valorActual = select.value;
      const valores = datos
        .map(function (item) {
          return item[configuracion[1]];
        })
        .filter(Boolean)
        .filter(function (valor, indice, arreglo) {
          return arreglo.indexOf(valor) === indice;
        });
      select.innerHTML =
        '<option value="">' + configuracion[2] + "</option>" +
        valores
          .map(function (valor) {
            return (
              '<option value="' +
              textoSeguro(valor) +
              '">' +
              textoSeguro(valor) +
              "</option>"
            );
          })
          .join("");
      if (valores.indexOf(valorActual) !== -1) {
        select.value = valorActual;
      }
    });
  }

  function formatearFechaHora(valor) {
    if (!valor) {
      return "Fecha no disponible";
    }
    const fecha = new Date(valor);
    if (isNaN(fecha.getTime())) {
      return String(valor);
    }
    return fecha.toLocaleString("es-CO", {
      dateStyle: "medium",
      timeStyle: "short"
    });
  }

  function renderizarActividad(actividades, paginacion) {
    const lista = document.getElementById("lista-actividad");
    lista.innerHTML = actividades.length
      ? actividades
          .map(function (item) {
            return (
              "<tr>" +
              '<td data-label="Fecha" class="history-date">' +
              textoSeguro(formatearFechaHora(item.fecha)) +
              "</td>" +
              '<td data-label="Usuario"><strong>' +
              textoSeguro(item.usuario || "Usuario de SharePoint") +
              "</strong></td>" +
              '<td data-label="Correo">' +
              textoSeguro(item.correo || "No disponible") +
              "</td>" +
              '<td data-label="Tipo"><span class="activity-type">' +
              textoSeguro(item.tipo || "General") +
              "</span></td>" +
              '<td data-label="Requerimiento">' +
              textoSeguro(item.requerimiento || "No aplica") +
              "</td>" +
              '<td data-label="Actividad" class="history-action">' +
              textoSeguro(item.accion) +
              "</td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="6">No hay actividad que coincida con los filtros.</td></tr>';
    renderizarPaginacionHistorial(paginacion);
  }

  function renderizarPaginacionHistorial(paginacion) {
    const pie = document.getElementById("pie-paginacion-historial");
    const resumen = document.getElementById("resumen-paginacion-historial");
    const controles = document.getElementById(
      "controles-paginacion-historial"
    );
    const tieneDatos = paginacion && paginacion.total > 0;
    pie.hidden = !tieneDatos;
    if (!tieneDatos) {
      resumen.textContent = "";
      controles.innerHTML = "";
      return;
    }
    resumen.textContent =
      "Mostrando " +
      paginacion.fin +
      " Registros " +
      " de " +
      paginacion.total;
    const botones = paginasPaginacion(
      paginacion.pagina,
      paginacion.totalPaginas
    )
      .map(function (pagina) {
        if (pagina === "...") {
          return '<span class="pagination-ellipsis" aria-hidden="true">\u2026</span>';
        }
        const activa = pagina === paginacion.pagina;
        return (
          '<button type="button" data-pagina="' +
          pagina +
          '"' +
          (activa ? ' class="active" aria-current="page"' : "") +
          ">" +
          pagina +
          "</button>"
        );
      })
      .join("");
    controles.innerHTML =
      '<button type="button" data-pagina="' +
      (paginacion.pagina - 1) +
      '" aria-label="P\u00e1gina anterior"' +
      (paginacion.pagina === 1 ? " disabled" : "") +
      ">&#8249;</button>" +
      botones +
      '<button type="button" data-pagina="' +
      (paginacion.pagina + 1) +
      '" aria-label="P\u00e1gina siguiente"' +
      (paginacion.pagina === paginacion.totalPaginas ? " disabled" : "") +
      ">&#8250;</button>";
  }

  // --------------------------------------------------------------------------
  // Biblioteca de adjuntos y acciones sobre archivos
  // --------------------------------------------------------------------------
  function fechaArchivo(valor) {
    if (!valor) {
      return "\u2014";
    }
    const fecha = new Date(valor);
    return isNaN(fecha.getTime())
      ? valor
      : fecha.toLocaleString("es-CO", {
          dateStyle: "short",
          timeStyle: "short"
        });
  }

  function urlDescargaArchivo(url) {
    const valor = String(url || "");
    return valor + (valor.indexOf("?") === -1 ? "?" : "&") + "download=1";
  }

  function urlVistaPreviaArchivo(url) {
    const valor = String(url || "");
    return valor + (valor.indexOf("?") === -1 ? "?" : "&") + "web=1";
  }

  function renderizarBibliotecaAdjuntos(archivos) {
    if (!archivos.length) {
      return (
        '<div id="biblioteca-adjuntos">' +
        '<span class="detail-empty">Sin archivos adjuntos</span>' +
        "</div>"
      );
    }

    const archivosOrdenados = archivos
      .slice()
      .sort(function (a, b) {
        const carpetaA = a.tipoDocumento || "Sin clasificaci\u00f3n";
        const carpetaB = b.tipoDocumento || "Sin clasificaci\u00f3n";
        return (
          carpetaA.localeCompare(carpetaB, "es") ||
          String(a.nombre || "").localeCompare(
            String(b.nombre || ""),
            "es"
          )
        );
      });
    const totalArchivos = archivosOrdenados.length;
    const totalPaginas = Math.max(
      1,
      Math.ceil(totalArchivos / ADJUNTOS_POR_PAGINA)
    );
    paginaAdjuntosDetalle = Math.min(
      Math.max(1, paginaAdjuntosDetalle),
      totalPaginas
    );
    const inicio =
      (paginaAdjuntosDetalle - 1) * ADJUNTOS_POR_PAGINA;
    const fin = Math.min(
      inicio + ADJUNTOS_POR_PAGINA,
      totalArchivos
    );

    const filas = archivosOrdenados
      .slice(inicio, fin)
      .map(function (archivo) {
        const carpeta =
          archivo.tipoDocumento || "Sin clasificaci\u00f3n";
        const nombre = archivo.nombre || "Abrir archivo";
        const extension = nombre.split(".").pop().toLowerCase();
        const admiteVistaPrevia = [
          "pdf",
          "doc",
          "docx",
          "xls",
          "xlsx",
          "png",
          "jpg",
          "jpeg"
        ].indexOf(extension) !== -1;
        const urlVistaPrevia = textoSeguro(
          urlVistaPreviaArchivo(archivo.url)
        );
        const urlDescarga = textoSeguro(
          urlDescargaArchivo(archivo.url)
        );
        const accionVistaPrevia = admiteVistaPrevia
          ? '<a class="sp-file-action" href="' +
            urlVistaPrevia +
            '" target="_blank" rel="noopener noreferrer" title="Visualizar" aria-label="Visualizar ' +
            textoSeguro(nombre) +
            ' en otra pesta\u00f1a">' +
            ICONO_OJO +
            "</a>"
          : '<span class="sp-file-action is-disabled" title="Vista previa no disponible para este formato" aria-label="Vista previa no disponible para ' +
            textoSeguro(nombre) +
            '">' +
            ICONO_OJO +
            "</span>";
        const enlaceNombre = admiteVistaPrevia
          ? '<a class="sp-file-name-link" href="' +
            urlVistaPrevia +
            '" target="_blank" rel="noopener noreferrer">' +
            textoSeguro(nombre) +
            "</a>"
          : '<span class="sp-file-name-no-preview">' +
            textoSeguro(nombre) +
            "</span>";
        return (
          '<div class="sp-library-row sp-file-row" role="row">' +
          '<span class="sp-file-icon" aria-hidden="true"></span>' +
          '<span class="sp-library-name">' +
          enlaceNombre +
          "</span>" +
          '<span class="sp-library-meta">' +
          textoSeguro(fechaArchivo(archivo.modificado)) +
          "</span>" +
          '<span class="sp-library-type"><span class="sp-folder-label">' +
          textoSeguro(carpeta) +
          "</span></span>" +
          '<span class="sp-file-actions">' +
          accionVistaPrevia +
          '<a class="sp-file-action" href="' +
          urlDescarga +
          '" download title="Descargar" aria-label="Descargar ' +
          textoSeguro(nombre) +
          '">' +
          ICONO_DESCARGA +
          "</a>" +
          "</span>" +
          "</div>"
        );
      })
      .join("");
    const botonesPagina = Array.from(
      { length: totalPaginas },
      function (_, indice) {
        const pagina = indice + 1;
        return (
          '<button type="button" class="sp-page-button' +
          (pagina === paginaAdjuntosDetalle ? " is-active" : "") +
          '" data-pagina-adjuntos="' +
          pagina +
          '" aria-label="P\u00e1gina ' +
          pagina +
          '"' +
          (pagina === paginaAdjuntosDetalle
            ? ' aria-current="page"'
            : "") +
          ">" +
          pagina +
          "</button>"
        );
      }
    ).join("");

    return (
      '<div id="biblioteca-adjuntos">' +
      '<div class="sp-library" role="table" aria-label="Archivos adjuntos">' +
      '<div class="sp-library-row sp-library-header" role="row">' +
      '<span aria-hidden="true"></span>' +
      '<strong class="sp-library-name">Nombre</strong>' +
      '<strong class="sp-library-meta">Modificado</strong>' +
      '<strong class="sp-library-type">Carpeta</strong>' +
      '<strong class="sp-file-actions-header">Acciones</strong>' +
      "</div>" +
      filas +
      "</div>" +
      '<footer class="sp-library-pagination">' +
      '<span>Mostrando ' +
      (inicio + 1) +
      "\u2013" +
      fin +
      " de " +
      totalArchivos +
      "</span>" +
      '<div class="sp-pagination-controls">' +
      '<button type="button" class="sp-page-button sp-page-nav" data-pagina-adjuntos="' +
      (paginaAdjuntosDetalle - 1) +
      '"' +
      (paginaAdjuntosDetalle === 1 ? " disabled" : "") +
      ">Anterior</button>" +
      botonesPagina +
      '<button type="button" class="sp-page-button sp-page-nav" data-pagina-adjuntos="' +
      (paginaAdjuntosDetalle + 1) +
      '"' +
      (paginaAdjuntosDetalle === totalPaginas ? " disabled" : "") +
      ">Siguiente</button>" +
      "</div>" +
      "</footer>" +
      "</div>"
    );
  }

  function cambiarPaginaAdjuntos(pagina) {
    const destino = Number(pagina);
    const totalPaginas = Math.max(
      1,
      Math.ceil(
        archivosAdjuntosDetalle.length / ADJUNTOS_POR_PAGINA
      )
    );
    if (
      !Number.isInteger(destino) ||
      destino < 1 ||
      destino > totalPaginas ||
      destino === paginaAdjuntosDetalle
    ) {
      return;
    }
    paginaAdjuntosDetalle = destino;
    const contenedor = document.getElementById(
      "biblioteca-adjuntos"
    );
    if (contenedor) {
      contenedor.outerHTML = renderizarBibliotecaAdjuntos(
        archivosAdjuntosDetalle
      );
    }
  }

  function renderizarArchivosFormulario(archivos) {
    const contenedor = document.getElementById(
      "archivos-existentes-edicion"
    );
    if (!contenedor) {
      return;
    }
    const lista = Array.isArray(archivos) ? archivos : [];
    contenedor.hidden = lista.length === 0;
    if (!lista.length) {
      contenedor.innerHTML = "";
      return;
    }

    const filas = lista
      .slice()
      .sort(function (a, b) {
        return String(a.nombre || "").localeCompare(
          String(b.nombre || ""),
          "es"
        );
      })
      .map(function (archivo) {
        const nombre = archivo.nombre || "Abrir archivo";
        const carpeta =
          archivo.tipoDocumento || "Sin clasificaci\u00f3n";
        const extension = nombre.split(".").pop().toLowerCase();
        const admiteVistaPrevia = [
          "pdf", "doc", "docx", "xls", "xlsx",
          "png", "jpg", "jpeg"
        ].indexOf(extension) !== -1;
        const urlVistaPrevia = textoSeguro(
          urlVistaPreviaArchivo(archivo.url)
        );
        const urlDescarga = textoSeguro(
          urlDescargaArchivo(archivo.url)
        );
        const enlaceNombre = admiteVistaPrevia
          ? '<a class="sp-file-name-link" href="' +
            urlVistaPrevia +
            '" target="_blank" rel="noopener noreferrer">' +
            textoSeguro(nombre) +
            "</a>"
          : '<span class="sp-file-name-no-preview">' +
            textoSeguro(nombre) +
            "</span>";
        const accionVistaPrevia = admiteVistaPrevia
          ? '<a class="sp-file-action" href="' +
            urlVistaPrevia +
            '" target="_blank" rel="noopener noreferrer" title="Visualizar" aria-label="Visualizar ' +
            textoSeguro(nombre) +
            '">' + ICONO_OJO + "</a>"
          : '<span class="sp-file-action is-disabled" title="Vista previa no disponible">' +
            ICONO_OJO +
            "</span>";

        return (
          '<div class="sp-library-row sp-file-row" role="row">' +
          '<span class="sp-file-icon" aria-hidden="true"></span>' +
          '<span class="sp-library-name">' + enlaceNombre + "</span>" +
          '<span class="sp-library-meta">' +
          textoSeguro(fechaArchivo(archivo.modificado)) +
          "</span>" +
          '<span class="sp-library-type"><span class="sp-folder-label">' +
          textoSeguro(carpeta) +
          "</span></span>" +
          '<span class="sp-file-actions">' +
          accionVistaPrevia +
          '<a class="sp-file-action" href="' +
          urlDescarga +
          '" download title="Descargar" aria-label="Descargar ' +
          textoSeguro(nombre) +
          '">' + ICONO_DESCARGA + "</a>" +
          '<button class="sp-file-action sp-file-delete" type="button" data-eliminar-archivo="' +
          textoSeguro(archivo.url) +
          '" data-nombre-archivo="' +
          textoSeguro(nombre) +
          '" title="Eliminar" aria-label="Eliminar ' +
          textoSeguro(nombre) +
          '">' + ICONO_PAPELERA + "</button>" +
          "</span></div>"
        );
      })
      .join("");

    contenedor.innerHTML =
      '<div class="sp-library sp-library-edit" role="table" aria-label="Archivos adjuntos existentes">' +
      '<div class="sp-library-row sp-library-header" role="row">' +
      '<span aria-hidden="true"></span>' +
      '<strong class="sp-library-name">Nombre</strong>' +
      '<strong class="sp-library-meta">Modificado</strong>' +
      '<strong class="sp-library-type">Carpeta</strong>' +
      '<strong class="sp-file-actions-header">Acciones</strong>' +
      "</div>" +
      filas +
      "</div>" +
      '<footer class="sp-library-pagination"><span>Mostrando 1\u2013' +
      lista.length +
      " de " +
      lista.length +
      "</span></footer>";
  }

  function renderizarCargaAdjuntos() {
    return (
      '<section class="detail-upload-panel" aria-labelledby="titulo-agregar-adjuntos">' +
      '<div class="detail-upload-heading">' +
      '<div><strong id="titulo-agregar-adjuntos">Agregar archivos</strong>' +
      "<p>Selecciona uno o varios archivos y asigna su tipo documental.</p></div>" +
      '<label class="detail-file-picker" for="detalle-archivos-nuevos">' +
      '<span>Elegir archivos</span>' +
      '<input type="file" id="detalle-archivos-nuevos" multiple ' +
      'accept=".pdf,.doc,.docx,.xls,.xlsx,.png,.jpg,.jpeg,.zip">' +
      "</label>" +
      "</div>" +
      '<p class="detail-upload-help">PDF, Word, Excel, im\u00e1genes o ZIP. M\u00e1ximo 10 archivos de 20 MB cada uno.</p>' +
      '<p id="detalle-resumen-archivos" class="detail-upload-summary" aria-live="polite">No hay archivos agregados.</p>' +
      '<div id="detalle-clasificacion-archivos" class="file-category-list" aria-live="polite" hidden></div>' +
      '<div class="detail-upload-actions">' +
      '<button class="btn-primary" id="detalle-subir-archivos" type="button" disabled>Subir archivos</button>' +
      "</div>" +
      "</section>"
    );
  }

  function mostrarDetalle(req) {
    const overlay = document.getElementById("modal-detalle");
    const contenido = document.getElementById("modal-contenido");
    const campos = [
      ["ID", req.id],
      ["APP", req.app],
      ["Tipo de servicio", req.tipoServicio],
      ["Asunto", req.asunto],
      ["Descripci\u00f3n", req.descripcion || "(sin descripci\u00f3n)"],
      ["Comentarios", req.comentarios || "(sin comentarios)"],
      ["Caso origen", req.casoOrigen || "No aplica"],
      ["Solicitado por", req.solicitadoPor],
      ["Responsable", req.responsable || "No asignado"],
      ["Mentor", req.mentor || "No asignado"],
      ["Estado", req.estado],
      ["Prioridad", req.prioridad],
      ["Fecha de solicitud", formatearFecha(req.fechaSolicitud) || "Sin definir"],
      ["F.E Entrega", formatearFecha(req.fechaEntrega) || "Sin definir"],
      ["Complejidad", req.complejidad || "Sin definir"],
      ["F.E PAP", formatearFecha(req.fechaPAP) || "Sin definir"],
      ["F.E Cierre", formatearFecha(req.fechaCierre) || "Sin definir"]
    ];
    const archivos = Array.isArray(req.archivosAdjuntos)
      ? req.archivosAdjuntos
      : [];
    archivosAdjuntosDetalle = archivos.slice();
    paginaAdjuntosDetalle = 1;
    const listaArchivos = renderizarBibliotecaAdjuntos(archivos);
    const cargaArchivos = renderizarCargaAdjuntos();

    contenido.innerHTML =
      campos
      .map(function (fila) {
        return (
          '<div class="detail-row"><strong>' +
          textoSeguro(fila[0]) +
          ":</strong><span>" +
          textoSeguro(fila[1]) +
          "</span></div>"
        );
      })
      .join("") +
      '<div class="detail-row detail-attachments"><strong>Archivos adjuntos:</strong><div>' +
      listaArchivos +
      cargaArchivos +
      "</div></div>";
    overlay.hidden = false;
  }

  function cerrarDetalle() {
    document.getElementById("modal-detalle").hidden = true;
  }

  global.Vista = Object.freeze({
    mostrarDashboard: mostrarDashboard,
    mostrarFormularioCrear: mostrarFormularioCrear,
    mostrarFormularioEdicion: mostrarFormularioEdicion,
    mostrarMisRequerimientos: mostrarMisRequerimientos,
    mostrarGestion: mostrarGestion,
    mostrarHistorial: mostrarHistorial,
    mostrarIndicadores: mostrarIndicadores,
    mostrarConexion: mostrarConexion,
    mostrarUsuario: mostrarUsuario,
    configurarAcceso: configurarAcceso,
    formatearFecha: formatearFecha,
    renderizarTarjetas: renderizarTarjetas,
    renderizarTabla: renderizarTabla,
    renderizarMisRequerimientos: renderizarMisRequerimientos,
    activarPestanaMisRequerimientos: activarPestanaMisRequerimientos,
    renderizarGestion: renderizarGestion,
    mostrarSugerenciasGestion: mostrarSugerenciasGestion,
    ocultarSugerenciasGestion: ocultarSugerenciasGestion,
    renderizarFiltros: renderizarFiltros,
    renderizarFiltrosMisRequerimientos: renderizarFiltrosMisRequerimientos,
    renderizarActividad: renderizarActividad,
    mostrarDetalle: mostrarDetalle,
    cambiarPaginaAdjuntos: cambiarPaginaAdjuntos,
    renderizarArchivosFormulario: renderizarArchivosFormulario,
    cerrarDetalle: cerrarDetalle
  });

  global.ViewManager = global.Vista;
})(window);
