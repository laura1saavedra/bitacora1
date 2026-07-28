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
          '<td class="req-id">' + textoSeguro(req.id) + "</td>" +
          "<td>" + textoSeguro(req.app) + "</td>" +
          "<td>" + textoSeguro(req.tipoServicio) + "</td>" +
          "<td>" + textoSeguro(req.asunto) + "</td>" +
          "<td>" + textoSeguro(req.solicitadoPor) + "</td>" +
          "<td>" + textoSeguro(req.responsable || "No asignado") + "</td>" +
          "<td>" + textoSeguro(req.prioridad) + "</td>" +
          "<td>" + textoSeguro(req.estado) + "</td>" +
          "<td>" + textoSeguro(req.fechaSolicitud) + "</td>" +
          "<td>" + textoSeguro(req.fechaCierre || "Sin definir") + "</td>" +
          "<td>" +
          '<button class="action-btn view-btn" data-id="' +
          textoSeguro(req.id) +
          '">Ver</button>' +
          "</td></tr>"
        );
      })
      .join("");
  }

  function renderizarMisRequerimientos(datos) {
    const cuerpo = document.getElementById("tabla-mis-requerimientos");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            return (
              "<tr><td>" + textoSeguro(req.id) + "</td>" +
              "<td>" + textoSeguro(req.app) + "</td>" +
              "<td>" + textoSeguro(req.tipoServicio) + "</td>" +
              "<td>" + textoSeguro(req.asunto) + "</td>" +
              '<td class="request-description" title="' +
              textoSeguro(req.descripcion) +
              '">' + textoSeguro(req.descripcion || "Sin descripci\u00f3n") + "</td>" +
              "<td>" + textoSeguro(req.solicitadoPor) + "</td>" +
              "<td>" + textoSeguro(req.responsable || "No asignado") + "</td>" +
              "<td>" + textoSeguro(req.estado) + "</td>" +
              '<td class="request-actions">' +
              '<button class="action-btn icon-action view-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Ver requerimiento ' +
              textoSeguro(req.id) +
              '" title="Ver detalle"><span aria-hidden="true">&#128065;</span></button>' +
              '<button class="action-btn icon-action edit-btn" data-id="' +
              textoSeguro(req.id) +
              '" type="button" aria-label="Editar requerimiento ' +
              textoSeguro(req.id) +
              '" title="Editar"><span aria-hidden="true">&#9998;</span></button>' +
              "</td></tr>"
            );
          })
          .join("")
      : '<tr class="empty-row"><td colspan="9">No tienes requerimientos registrados.</td></tr>';
  }

  function renderizarGestion(datos) {
    const cuerpo = document.getElementById("tabla-gestion");
    cuerpo.innerHTML = datos.length
      ? datos
          .map(function (req) {
            return (
              '<tr data-id="' + textoSeguro(req.id) + '">' +
              "<td>" + textoSeguro(req.id) + "</td>" +
              "<td>" + textoSeguro(req.asunto) + "</td>" +
              "<td>" + textoSeguro(req.responsable || "No asignado") + "</td>" +
              "<td>" + textoSeguro(req.mentor || "No asignado") + "</td>" +
              '<td><select class="gestion-estado">' +
              opcionesEstado(req.estado) +
              "</select></td>" +
              '<td><input class="gestion-comentarios" type="text" value="' +
              textoSeguro(req.comentarios) +
              '" aria-label="Observaciones del requerimiento ' +
              textoSeguro(req.id) +
              '"></td>' +
              '<td><button class="action-btn save-btn" data-id="' +
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
