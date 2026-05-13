let dbProductos = JSON.parse(localStorage.getItem('lico_db')) || [];
let historialVisual = JSON.parse(localStorage.getItem('lico_historial_visual')) || [];
let cierresDia = JSON.parse(localStorage.getItem('lico_cierres')) || [];

function init() {
    renderizarFilas();
    renderizarDB();
    document.getElementById('fecha-input').value = new Date().toLocaleDateString();
    document.getElementById('import-file').addEventListener('change', manejarImportacion);
    verificarRespaldoSemanal();
}

function verificarRespaldoSemanal() {
    const ultima = localStorage.getItem('lico_ultima_exportacion');
    if (ultima) {
        const dias = (new Date().getTime() - parseInt(ultima)) / (1000 * 60 * 60 * 24);
        if (dias >= 7) document.getElementById('bloqueo-seguridad').style.display = 'flex';
    } else {
        localStorage.setItem('lico_ultima_exportacion', new Date().getTime());
    }
}

function calcular(el) {
    const f = el.closest('tr');
    const caj = parseFloat(f.querySelector('.caj').value) || 0;
    const pre = parseFloat(f.querySelector('.pre').value) || 0;
    const uCaja = parseFloat(f.dataset.uCaja) || 1;
    const tUds = caj * uCaja;
    const tFila = caj * pre;
    f.querySelector('.uds').value = tUds;
    f.querySelector('.tot').innerText = tFila.toLocaleString('de-DE');
    f.querySelector('.pun').innerText = tUds > 0 ? (tFila/tUds).toLocaleString('de-DE', {maximumFractionDigits:0}) : 0;
    
    recalcularTotalesBase();
}

function recalcularTotalesBase() {
    let suma = 0;
    document.querySelectorAll('.tot').forEach(s => suma += parseFloat(s.innerText.replace(/\./g,'')) || 0);
    document.getElementById('subtotal-factura').innerText = '$ ' + suma.toLocaleString('de-DE');
    recalcularConNota();
}

function recalcularConNota() {
    const subtotalText = document.getElementById('subtotal-factura').innerText.replace('$ ','').replace(/\./g,'');
    const subtotal = parseFloat(subtotalText) || 0;
    const notaInput = document.getElementById('nota-credito').value;
    const notaValor = parseFloat(notaInput.replace(/[^\d]/g, '')) || 0;
    const totalFinal = subtotal - notaValor;
    document.getElementById('total-factura').innerText = '$ ' + totalFinal.toLocaleString('de-DE');
}

function guardarFacturaVisual() {
    const cli = document.getElementById('cliente-input').value;
    if(!cli) return alert("Ingrese el nombre del cliente");
    
    const area = document.getElementById('area-captura');
    const clon = area.cloneNode(true);
    clon.querySelectorAll('.fila-p').forEach(f => {
        if(!f.querySelector('.cod-in').value) f.remove();
    });

    clon.querySelectorAll('input').forEach(ins => {
        const span = document.createElement('span');
        span.innerText = ins.value;
        ins.parentNode.replaceChild(span, ins);
    });

    historialVisual.push({
        id: document.getElementById('n-factura').value,
        fecha: document.getElementById('fecha-input').value,
        cliente: cli,
        total: document.getElementById('total-factura').innerText,
        html: clon.innerHTML
    });

    localStorage.setItem('lico_historial_visual', JSON.stringify(historialVisual));
    alert("Factura guardada correctamente.");
}

// --- FUNCIÓN DE EXPORTACIÓN TOTAL MEJORADA ---
function exportarTodoExcel() {
    try {
        const wb = XLSX.utils.book_new();
        let todasLasVentas = [];

        // 1. Recolectar facturas del historial actual (las que no se han cerrado)
        historialVisual.forEach(f => extraerDatosFactura(f, todasLasVentas, "Pendiente de Cierre"));

        // 2. Recolectar facturas de todos los cierres guardados
        cierresDia.forEach(cierre => {
            cierre.facturas.forEach(f => {
                extraerDatosFactura(f, todasLasVentas, cierre.fechaCierre);
            });
        });

        // Crear las pestañas en el Excel
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(todasLasVentas), "Ventas Detalladas");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cierresDia.map(c=>({Fecha:c.fechaCierre, Cantidad_Facts:c.cantidadFacturas, Total_Cierre:c.totalAcumulado}))), "Resumen Cierres");
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dbProductos), "Inventario_DB");

        XLSX.writeFile(wb, "REPORTE_TOTAL_LICOEXPRESS.xlsx");
        
        localStorage.setItem('lico_ultima_exportacion', new Date().getTime());
        document.getElementById('bloqueo-seguridad').style.display = 'none';
    } catch (e) { 
        alert("Error al exportar: " + e.message); 
        console.error(e);
    }
}

// Función auxiliar para procesar el HTML de las facturas y volverlo filas de Excel
function extraerDatosFactura(f, arrayDestino, fechaCierreRef) {
    let tempDiv = document.createElement('div');
    tempDiv.innerHTML = f.html;
    tempDiv.querySelectorAll('.fila-p').forEach(fila => {
        let celdas = fila.querySelectorAll('span');
        if(celdas.length >= 6) {
            arrayDestino.push({
                "Estado/Cierre": fechaCierreRef,
                "Factura": f.id, 
                "Fecha": f.fecha, 
                "Cliente": f.cliente,
                "Código": celdas[0].innerText, 
                "Descripción": celdas[1].innerText,
                "Cajas": celdas[2].innerText, 
                "Cant_Unid": celdas[3].innerText,
                "Precio_Venta": celdas[4].innerText, 
                "Total_Producto": celdas[6].innerText,
                "Total_Factura": f.total
            });
        }
    });
}

function renderizarFilas() {
    const body = document.getElementById('body-factura');
    body.innerHTML = "";
    for(let i=0; i<15; i++) {
        body.innerHTML += `
        <tr class="fila-p">
            <td><input type="text" class="cod-in"></td>
            <td style="width:250px"><input type="text" class="nom-in" oninput="buscarSugerencias(this)"><div class="sugerencias"></div></td>
            <td><input type="number" class="caj" oninput="calcular(this)"></td>
            <td><input type="number" class="uds" readonly></td>
            <td><input type="number" class="pre" oninput="calcular(this)"></td>
            <td align="right">$ <span class="pun">0</span></td>
            <td align="right">$ <span class="tot">0</span></td>
        </tr>`;
    }
}

function buscarSugerencias(input) {
    const container = input.nextElementSibling;
    const texto = input.value.toLowerCase();
    container.innerHTML = "";
    if(texto.length < 2) { container.style.display = "none"; return; }
    dbProductos.filter(p => p.nombre.toLowerCase().includes(texto)).slice(0,8).forEach(p => {
        const div = document.createElement('div');
        div.className = "sugerencia-item";
        div.innerHTML = p.nombre;
        div.onclick = () => {
            const f = input.closest('tr');
            f.querySelector('.cod-in').value = p.codigo;
            f.querySelector('.nom-in').value = p.nombre;
            f.dataset.uCaja = p.unidad;
            f.querySelector('.caj').value = p.unidad;
            container.style.display = "none";
            calcular(f.querySelector('.caj'));
            f.querySelector('.pre').focus();
        };
        container.appendChild(div);
    });
    container.style.display = "block";
}

function ejecutarCierreDia() {
    if (historialVisual.length === 0) return alert("No hay facturas para cerrar.");
    if (confirm("¿Cerrar el día y limpiar historial?")) {
        let total = 0;
        historialVisual.forEach(f => total += parseFloat(f.total.replace('$ ','').replace(/\./g,'')) || 0);
        cierresDia.push({
            fechaCierre: new Date().toLocaleString(),
            cantidadFacturas: historialVisual.length,
            totalAcumulado: '$ ' + total.toLocaleString('de-DE'),
            facturas: [...historialVisual]
        });
        localStorage.setItem('lico_cierres', JSON.stringify(cierresDia));
        historialVisual = [];
        localStorage.setItem('lico_historial_visual', JSON.stringify(historialVisual));
        cambiarVista('vista-factura');
    }
}

function renderizarHistorial() {
    document.getElementById('lista-historial').innerHTML = historialVisual.map((f, i) => `<tr><td>${f.id}</td><td>${f.fecha}</td><td>${f.cliente}</td><td>${f.total}</td><td align="center"><button class="btn btn-db" style="padding:5px; width:auto;" onclick="verCaptura(${i})">👁️</button></td></tr>`).join('');
}

function verCaptura(i) {
    document.getElementById('contenido-modal').innerHTML = historialVisual[i].html;
    document.getElementById('modalFactura').style.display = "block";
}

function cambiarVista(id) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    if(id === 'vista-historial') renderizarHistorial();
    if(id === 'vista-lista-cierres') renderizarListaCierres();
}

function renderizarListaCierres() {
    document.getElementById('tabla-cierres-body').innerHTML = cierresDia.map((c, i) => `<tr><td>${c.fechaCierre}</td><td>${c.cantidadFacturas}</td><td>${c.totalAcumulado}</td><td align="center"><button class="btn btn-db" onclick="verDetalleCierre(${i})">👁️</button></td></tr>`).join('');
}

function verDetalleCierre(i) {
    const c = cierresDia[i];
    document.getElementById('contenido-modal').innerHTML = c.facturas.map(f => `<div style="border:1px solid #ccc; margin:10px; padding:10px;">${f.html}</div>`).join('');
    document.getElementById('modalFactura').style.display = "block";
}

function cerrarModal() { document.getElementById('modalFactura').style.display = "none"; }
function renderizarDB() { document.getElementById('lista-db').innerHTML = dbProductos.map(p => `<tr><td>${p.codigo}</td><td>${p.nombre}</td><td>${p.unidad}</td></tr>`).join(''); }

function manejarImportacion(e) {
    const r = new FileReader();
    r.onload = (x) => {
        const j = XLSX.utils.sheet_to_json(XLSX.read(new Uint8Array(x.target.result), {type:'array'}).Sheets[XLSX.read(new Uint8Array(x.target.result), {type:'array'}).SheetNames[0]]);
        dbProductos = j.map(i => ({codigo:String(i.CODIGO||i.codigo), nombre:i.DESCRIPCION||i.descripcion, unidad:i.UNIDAD||1}));
        localStorage.setItem('lico_db', JSON.stringify(dbProductos)); renderizarDB();
    };
    r.readAsArrayBuffer(e.target.files[0]);
}

function nuevaFactura() {
    if(confirm("¿Limpiar pantalla?")) {
        document.getElementById('n-factura').value = "CLAH-" + (historialVisual.length + 1).toString().padStart(3, '0');
        document.getElementById('cliente-input').value = "";
        renderizarFilas();
        document.getElementById('subtotal-factura').innerText = "$ 0";
        document.getElementById('total-factura').innerText = "$ 0";
        document.getElementById('nota-credito').value = "-$ 0";
    }
}
