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

function renderizarFilas() {
    const body = document.getElementById('body-factura');
    body.innerHTML = "";
    for(let i=0; i<15; i++) {
        body.innerHTML += `
        <tr class="fila-p">
            <td><input type="text" class="cod-in" onkeyup="buscarPorCodigo(this)" autocomplete="off"></td>
            <td style="width:250px; position:relative;">
                <input type="text" class="nom-in" oninput="buscarSugerencias(this)" autocomplete="off">
                <div class="sugerencias"></div>
            </td>
            <td><input type="number" class="caj" oninput="calcular(this)"></td>
            <td><input type="number" class="uds" readonly></td>
            <td><input type="number" class="pre" oninput="calcular(this)"></td>
            <td align="right">$ <span class="pun">0</span></td>
            <td align="right">$ <span class="tot">0</span></td>
        </tr>`;
    }
}

function buscarPorCodigo(input) {
    const cod = input.value.trim();
    if(cod === "") return;
    const p = dbProductos.find(x => String(x.codigo).trim() === cod);
    if(p) {
        const f = input.closest('tr');
        f.querySelector('.nom-in').value = p.nombre;
        f.dataset.uCaja = p.unidad;
        f.querySelector('.caj').value = 1;
        calcular(f.querySelector('.caj'));
    }
}

function buscarSugerencias(input) {
    const container = input.nextElementSibling;
    const texto = input.value.toLowerCase();
    container.innerHTML = "";
    if(texto.length < 2) { container.style.display = "none"; return; }
    
    const coincidencia = dbProductos.filter(p => p.nombre.toLowerCase().includes(texto)).slice(0,8);
    coincidencia.forEach(p => {
        const div = document.createElement('div');
        div.className = "sugerencia-item";
        div.innerHTML = `<strong>${p.codigo}</strong> - ${p.nombre}`;
        div.onclick = () => {
            const f = input.closest('tr');
            f.querySelector('.cod-in').value = p.codigo;
            f.querySelector('.nom-in').value = p.nombre;
            f.dataset.uCaja = p.unidad;
            f.querySelector('.caj').value = 1;
            container.style.display = "none";
            calcular(f.querySelector('.caj'));
            f.querySelector('.pre').focus();
        };
        container.appendChild(div);
    });
    container.style.display = coincidencia.length > 0 ? "block" : "none";
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
    const sub = parseFloat(document.getElementById('subtotal-factura').innerText.replace('$ ','').replace(/\./g,'')) || 0;
    const nota = parseFloat(document.getElementById('nota-credito').value.replace(/[^\d]/g,'')) || 0;
    document.getElementById('total-factura').innerText = '$ ' + (sub - nota).toLocaleString('de-DE');
}

function guardarFacturaVisual() {
    const cli = document.getElementById('cliente-input').value;
    if(!cli) return alert("Nombre de cliente requerido");
    const area = document.getElementById('area-captura');
    const clon = area.cloneNode(true);
    clon.querySelectorAll('.fila-p').forEach(f => { if(!f.querySelector('.cod-in').value) f.remove(); });
    clon.querySelectorAll('input').forEach(i => {
        const s = document.createElement('span'); s.innerText = i.value; i.parentNode.replaceChild(s, i);
    });
    historialVisual.push({
        id: document.getElementById('n-factura').value,
        fecha: document.getElementById('fecha-input').value,
        cliente: cli,
        total: document.getElementById('total-factura').innerText,
        html: clon.innerHTML
    });
    localStorage.setItem('lico_historial_visual', JSON.stringify(historialVisual));
    alert("Guardado.");
}

function exportarTodoExcel() {
    const wb = XLSX.utils.book_new();
    let ventasDetalladas = [];
    
    // Extraer de historial actual
    historialVisual.forEach(f => procesarHTMLFactura(f, ventasDetalladas, "Hoy"));
    // Extraer de cierres guardados
    cierresDia.forEach(c => c.facturas.forEach(f => procesarHTMLFactura(f, ventasDetalladas, c.fechaCierre)));

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ventasDetalladas), "VentasDetalle");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dbProductos), "DB_Productos");
    XLSX.writeFile(wb, "REPORTE_LICO_COMPLETO.xlsx");
    
    localStorage.setItem('lico_ultima_exportacion', new Date().getTime());
    document.getElementById('bloqueo-seguridad').style.display = 'none';
}

function procesarHTMLFactura(f, destino, ref) {
    const div = document.createElement('div'); div.innerHTML = f.html;
    div.querySelectorAll('.fila-p').forEach(fila => {
        const c = fila.querySelectorAll('span');
        if(c.length >= 6) {
            destino.push({
                Cierre: ref, Factura: f.id, Cliente: f.cliente,
                Codigo: c[0].innerText, Producto: c[1].innerText, Cajas: c[2].innerText, Total: c[6].innerText
            });
        }
    });
}

function ejecutarCierreDia() {
    if(historialVisual.length === 0 || !confirm("¿Cerrar día?")) return;
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
    location.reload();
}

function cambiarVista(id) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.getElementById(id).classList.add('active');
    if(id==='vista-historial') renderizarHistorial();
}

function renderizarHistorial() {
    document.getElementById('lista-historial').innerHTML = historialVisual.map((f,i)=>`<tr><td>${f.id}</td><td>${f.fecha}</td><td>${f.cliente}</td><td>${f.total}</td><td><button onclick="verCaptura(${i})">👁️</button></td></tr>`).join('');
}

function verCaptura(i) {
    document.getElementById('contenido-modal').innerHTML = historialVisual[i].html;
    document.getElementById('modalFactura').style.display = "block";
}

function cerrarModal() { document.getElementById('modalFactura').style.display = "none"; }

function manejarImportacion(e) {
    const r = new FileReader();
    r.onload = (x) => {
        const data = new Uint8Array(x.target.result);
        const workbook = XLSX.read(data, {type:'array'});
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        dbProductos = XLSX.utils.sheet_to_json(sheet).map(i => ({
            codigo: String(i.CODIGO || i.codigo),
            nombre: i.DESCRIPCION || i.descripcion,
            unidad: i.UNIDAD || 1
        }));
        localStorage.setItem('lico_db', JSON.stringify(dbProductos));
        renderizarDB();
        alert("Base de datos cargada");
    };
    r.readAsArrayBuffer(e.target.files[0]);
}

function renderizarDB() {
    document.getElementById('lista-db').innerHTML = dbProductos.map(p => `<tr><td>${p.codigo}</td><td>${p.nombre}</td><td>${p.unidad}</td></tr>`).join('');
}

function nuevaFactura() {
    if(confirm("¿Nueva factura?")) {
        renderizarFilas();
        document.getElementById('cliente-input').value = "";
        document.getElementById('total-factura').innerText = "$ 0";
    }
}
