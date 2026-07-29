// ============================================================
// VISTA.JS
// Navegacion, renderizado y lectura de la interfaz.
// ============================================================
(function (global) {
  "use strict";

  const VISTAS = {
    dashboard: "view-dashboard",
    crear: "view-crear-requerimiento",
    mis: "view-mis-requerimientos",
    gestion: "view-gestion"
  };
  const ICONO_OJO =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"></path>' +
    '<circle cx="12" cy="12" r="2.5"></circle></svg>';
  const ICONO_LAPIZ =
    '<svg viewBox="0 0 24 24" aria-hidden="true">' +
    '<path d="m4 20 4.3-1.1L19 8.2 15.8 5 5.1 15.7 4 20Z"></path>' +
    '<path d="m14.5 6.3 3.2 3.2"></path></svg>';

  function textoSeguro(valor) {
    return String(valor == null ? "" : valor)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

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

  function renderizarTarjetas(datos) {
    const valores = {
      "total-requerimientos": datos.length,
      "total-pendientes": datos.filter(function (item) {
        return item.estado === "Pendiente";
      }).length,
      "total-en-pruebas": datos.filter(function (item) {
        return item.estado === "En pruebas";
      }).length,
      "total-pruebas": datos.filter(function (item) {
        return item.estado === "Esperando Documentacion Usuario";
      }).length,
      "total-cierre-usuario": datos.filter(function (item) {
        return item.estado === "Esperando cierre usuario";
      }).length,
      "total-finalizados": datos.filter(function (item) {
        return item.estado === "Finalizado";
      }).length
    };
    Object.keys(valores).forEach(function (id) {
      const elemento = document.getElementById(id);
      if (elemento) {
        elemento.textContent = valores[id];
      }
    });
  }

  function renderizarTabla(datos) {
    const cuerpo = document.getElementById("tabla-requerimientos");
    cuerpo.innerHTML = datos
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
          '<td data-label="Estado">' + textoSeguro(req.estado) + "</td>" +
          '<td data-label="F. solicitud">' + textoSeguro(req.fechaSolicitud) + "</td>" +
          '<td data-label="F. cierre">' + textoSeguro(req.fechaCierre || "Sin definir") + "</td>" +
          '<td data-label="Acciones">' +
          '<button class="action-btn view-btn" data-id="' +
          textoSeguro(req.id) +
          '">Ver</button>' +
          "</td></tr>"
        );
      })
      .join("");
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
      paginacion.inicio +
      "\u2013" +
      paginacion.fin +
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
              '<button class="action-btn icon-action edit-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Editar requerimiento ' +
              textoSeguro(req.id) +
              '" title="Editar">' + ICONO_LAPIZ + "</button>" +
              "</div></td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="9">No tienes requerimientos registrados.</td></tr>';
    renderizarPaginacionMisRequerimientos(paginacion);
  }

  function renderizarGestion(datos) {
    const cuerpo = document.getElementById("tabla-gestion");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            return (
              '<tr data-id="' + textoSeguro(req.id) + '">' +
              '<td data-label="ID req.">' + textoSeguro(req.id) + "</td>" +
              '<td data-label="Asunto">' + textoSeguro(req.asunto) + "</td>" +
              '<td data-label="Responsable">' + textoSeguro(req.responsable || "No asignado") + "</td>" +
              '<td data-label="Mentor">' + textoSeguro(req.mentor || "No asignado") + "</td>" +
              '<td data-label="Estado"><select class="gestion-estado">' +
              opcionesEstado(req.estado) +
              "</select></td>" +
              '<td data-label="Observaciones"><input class="gestion-comentarios" type="text" value="' +
              textoSeguro(req.comentarios) +
              '" aria-label="Observaciones del requerimiento ' +
              textoSeguro(req.id) +
              '"></td>' +
              '<td data-label="Acciones"><button class="action-btn save-btn" data-id="' +
              textoSeguro(req.id) +
              '">Guardar</button></td></tr>'
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="7">No hay requerimientos para gestionar.</td></tr>';
  }

  function opcionesEstado(actual) {
    return [
      "Pendiente",
      "En proceso",
      "En pruebas",
      "Esperando Documentacion Usuario",
      "Esperando cierre usuario",
      "Finalizado",
      "Cancelado"
    ]
      .map(function (estado) {
        return (
          '<option value="' +
          textoSeguro(estado) +
          '"' +
          (estado === actual ? " selected" : "") +
          ">" +
          textoSeguro(estado) +
          "</option>"
        );
      })
      .join("");
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

  function renderizarActividad(actividades) {
    const lista = document.getElementById("lista-actividad");
    lista.innerHTML = actividades
      .map(function (item) {
        return (
          '<div class="activity-item"><div class="dot">\u2022</div><div class="txt">' +
          "<strong>" + textoSeguro(item.usuario) + "</strong><br>" +
          textoSeguro(item.accion) +
          '<div class="meta">' + textoSeguro(item.fecha) + "</div></div></div>"
        );
      })
      .join("");
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
      ["Fecha de solicitud", req.fechaSolicitud],
      ["F.E Entrega", req.fechaEntrega || "Sin definir"],
      ["Complejidad", req.complejidad || "Sin definir"],
      ["F.E PAP", req.fechaPAP || "Sin definir"],
      ["F.E Cierre", req.fechaCierre || "Sin definir"]
    ];
    const archivos = Array.isArray(req.archivosAdjuntos)
      ? req.archivosAdjuntos
      : [];
    const listaArchivos = archivos.length
      ? '<ul class="attachment-list">' +
        archivos
          .map(function (archivo) {
            return (
              '<li><a href="' +
              textoSeguro(archivo.url) +
              '" target="_blank" rel="noopener noreferrer">' +
              textoSeguro(archivo.nombre || "Abrir archivo") +
              "</a></li>"
            );
          })
          .join("") +
        "</ul>"
      : '<span class="detail-empty">Sin archivos adjuntos</span>';

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
    mostrarConexion: mostrarConexion,
    mostrarUsuario: mostrarUsuario,
    renderizarTarjetas: renderizarTarjetas,
    renderizarTabla: renderizarTabla,
    renderizarMisRequerimientos: renderizarMisRequerimientos,
    renderizarGestion: renderizarGestion,
    renderizarFiltros: renderizarFiltros,
    renderizarActividad: renderizarActividad,
    mostrarDetalle: mostrarDetalle,
    cerrarDetalle: cerrarDetalle
  });

  global.ViewManager = global.Vista;
})(window);
