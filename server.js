import express from 'express';
import { createServer } from 'http';
import { Server } from 'socket.io';
import basicAuth from 'express-basic-auth';
import ExcelJS from 'exceljs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);



const app = express();
const httpServer = createServer(app);
const io = new Server(httpServer,{
    pingTimeout: 300000,  // 5 minutos de inactividad sin respuesta (300,000 ms)
    pingInterval: 25000,  // Revisa el estado de la conexión cada 25 segundos
    connectTimeout: 30000 // Tiempo máximo para establecer la conexión inicial
});

// Configuración de clave para el Administrador / Encargado
const seguridadAdmin = basicAuth({
    users: { 'esub': '*guardia/9595' }, // Usuario: admin | Contraseña: tu_clave_aqui
    challenge: true, // Hace que el navegador muestre la ventana flotante de inicio de sesión
    unauthorizedResponse: 'Acceso no autorizado al Panel de Control de la PNA.'
});

// Aplicar la protección ÚNICAMENTE a las rutas que empiezan con /admin
app.use('/admin', seguridadAdmin);

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// -------------------------------------------------------------
// 1. BASE DE DATOS EN MEMORIA POR MES Y AÑO (Clave: "YYYY-MM")
// -------------------------------------------------------------
const baseDatosGuardias = {};

function obtenerOCrearMes(mes, año) {
    const clave = `${año}-${String(mes).padStart(2, '0')}`;

    if (!baseDatosGuardias[clave]) {
        const oficiales = [];
        const disponibles = [];

        // Cantidad de días reales del mes
        const totalDias = new Date(año, mes, 0).getDate();

        // Formato de nombre del mes (Ej: Octubre)
        const fechaObjMes = new Date(año, mes - 1, 1);
        let nombreMes = fechaObjMes.toLocaleString('es-AR', { month: 'long' });
        nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

        const strMes = String(mes).padStart(2, '0');

        for (let dia = 1; dia <= totalDias; dia++) {
            const strDia = String(dia).padStart(2, '0');
            const fechaObj = new Date(año, mes - 1, dia);

            // Obtener el día de la semana (Jue, Vie, Sáb...)
            let diaNombre = fechaObj.toLocaleString('es-AR', { weekday: 'short' });
            diaNombre = diaNombre.replace('.', '');
            diaNombre = diaNombre.charAt(0).toUpperCase() + diaNombre.slice(1);

            const fechaTexto = `${diaNombre} ${strDia}/${strMes}`;

            oficiales.push({
                id: dia,
                fecha: fechaTexto,
                estado: 'disponible',
                agente: null,
                reservadoEn: null
            });

            disponibles.push({
                id: dia,
                fecha: fechaTexto,
                estado: 'disponible',
                agente: null,
                reservadoEn: null
            });
        }

        baseDatosGuardias[clave] = {
            infoMes: { mesNombre: nombreMes, año: año, totalDias: totalDias, mesNumero: mes },
            oficiales: oficiales,
            disponibles: disponibles
        };
    }

    return baseDatosGuardias[clave];
}

// -------------------------------------------------------------
// 2. WEBSOCKETS EN TIEMPO REAL
// -------------------------------------------------------------
io.on('connection', (socket) => {
    // Al conectar, enviamos por defecto Octubre 2026 (o el mes actual)
    const datosIniciales = obtenerOCrearMes(10, 2026);
    socket.emit('cargarFechas', datosIniciales);

    // Consulta de un mes específico
    socket.on('obtenerFechas', (data) => {
        const mes = data && data.mes ? Number(data.mes) : 10;
        const año = data && data.año ? Number(data.año) : 2026;
        const datos = obtenerOCrearMes(mes, año);
        socket.emit('cargarFechas', datos);
    });

    // Reserva de guardia robusta por Mes y Año
    socket.on('solicitarReserva', (data) => {
        const { idFecha, tipoGuardia, jerarquia, apellido, nombre, mes, año } = data;
        
        const m = Number(mes) || 10;
        const a = Number(año) || 2026;

        const datosMes = obtenerOCrearMes(m, a);
        const lista = tipoGuardia === 'oficial' ? datosMes.oficiales : datosMes.disponibles;

        const fechaItem = lista.find(f => f.id === Number(idFecha));

        if (!fechaItem) {
            socket.emit('resultadoReserva', { exito: false, mensaje: 'Fecha no encontrada.' });
            return;
        }

        if (fechaItem.estado !== 'disponible') {
            socket.emit('resultadoReserva', { 
                exito: false, 
                mensaje: '¡Esta fecha ya está reservada por otro personal!' 
            });
            return;
        }

        // Asignar reserva exitosa
        fechaItem.estado = 'reservado';
        fechaItem.agente = { 
            jerarquia: jerarquia.trim(), 
            apellido: apellido.trim().toUpperCase(), 
            nombre: nombre.trim().toUpperCase() 
        };
        fechaItem.reservadoEn = new Date().toLocaleString('es-AR');

        const tipoTexto = tipoGuardia === 'oficial' ? 'Guardia Oficial' : 'Guardia Disponible';

        socket.emit('resultadoReserva', { 
            exito: true, 
            mensaje: `¡${tipoTexto} del ${fechaItem.fecha} reservada con éxito!` 
        });

        // Actualizar a todos los conectados
        io.emit('actualizarFechas', datosMes);
    });
});

// -------------------------------------------------------------
// 3. EXPORTAR A EXCEL Y REINICIAR
// -------------------------------------------------------------
app.get('/admin/exportar-excel', async (req, res) => {
    try {
        const mes = Number(req.query.mes) || 10;
        const año = Number(req.query.año) || 2026;
        const datos = obtenerOCrearMes(mes, año);

        const workbook = new ExcelJS.Workbook();

        // Estructura de columnas simplificada
        const columnasLimpia = [
            { header: 'Fecha', key: 'fecha', width: 15 },
            { header: 'Estado', key: 'estado', width: 14 },
            { header: 'Jerarquía', key: 'jerarquia', width: 12 },
            { header: 'Apellido', key: 'apellido', width: 22 },
            { header: 'Nombre', key: 'nombre', width: 22 }
        ];

        // Hoja 1: Guardias Oficiales
        const wsOficiales = workbook.addWorksheet('Guardias Oficiales');
        wsOficiales.columns = columnasLimpia;

        datos.oficiales.forEach(f => {
            wsOficiales.addRow({
                fecha: f.fecha,
                estado: f.estado.toUpperCase(),
                jerarquia: f.agente ? f.agente.jerarquia : '-',
                apellido: f.agente ? f.agente.apellido : 'SIN ASIGNAR',
                nombre: f.agente ? f.agente.nombre : '-'
            });
        });
        wsOficiales.getRow(1).font = { bold: true };

        // Hoja 2: Guardias Disponibles
        const wsDisponibles = workbook.addWorksheet('Guardias Disponibles');
        wsDisponibles.columns = columnasLimpia;

        datos.disponibles.forEach(f => {
            wsDisponibles.addRow({
                fecha: f.fecha,
                estado: f.estado.toUpperCase(),
                jerarquia: f.agente ? f.agente.jerarquia : '-',
                apellido: f.agente ? f.agente.apellido : 'SIN ASIGNAR',
                nombre: f.agente ? f.agente.nombre : '-'
            });
        });
        wsDisponibles.getRow(1).font = { bold: true };

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Guardias_${datos.infoMes.mesNombre}_${año}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        res.status(500).send('Error al generar Excel.');
    }
});
// -------------------------------------------------------------
// ELIMINAR REGISTROS DE UN MES ESPECÍFICO (ADMIN)
// -------------------------------------------------------------
app.get('/admin/eliminar-mes', (req, res) => {
    const mes = Number(req.query.mes) || 10;
    const año = Number(req.query.año) || 2026;
    const clave = `${año}-${String(mes).padStart(2, '0')}`;

    if (baseDatosGuardias[clave]) {
        delete baseDatosGuardias[clave]; // Borra el mes completamente de la memoria
    }

    // Volver a inicializar las fechas limpias para el mes consultado
    const datosNuevos = obtenerOCrearMes(mes, año);

    // Notificar a todos los usuarios conectados para actualizar la vista
    io.emit('actualizarFechas', datosNuevos);
    
    res.json({ exito: true, mensaje: `Se eliminaron todos los registros del mes ${mes}/${año}.` });
});
// -------------------------------------------------------------
// REINICIAR MES
// -------------------------------------------------------------
app.get('/admin/reiniciar-mes', (req, res) => {
    const mes = Number(req.query.mes) || 10;
    const año = Number(req.query.año) || 2026;
    const clave = `${año}-${String(mes).padStart(2, '0')}`;

    delete baseDatosGuardias[clave]; // Borra las reservas de ese mes específico
    const datosNuevos = obtenerOCrearMes(mes, año);

    io.emit('actualizarFechas', datosNuevos);
    res.send('Mes reiniciado.');
});

app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

const PORT = process.env.PORT || 3000;
httpServer.listen(PORT, () => {
    console.log(`>>> Servidor corriendo en http://localhost:${PORT}`);
});