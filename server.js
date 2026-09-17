import 'dotenv/config';
import express from 'express';
import { createServer } from 'http';
import { Server } from 'socket.io';
import basicAuth from 'express-basic-auth';
import path from 'path';
import { fileURLToPath } from 'url';
import pkg from 'pg';
import exceljs from 'exceljs';

console.log('>>> DATABASE_URL actual:', process.env.DATABASE_URL);

const { Pool } = pkg;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const httpServer = createServer(app);
const io = new Server(httpServer, {
    pingTimeout: 300000,
    pingInterval: 25000,
    connectTimeout: 30000
});

// -------------------------------------------------------------
// CONEXIÓN A POSTGRESQL
// -------------------------------------------------------------
const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// Inicializar tablas en PostgreSQL
async function inicializarBaseDatos() {
    try {
        // 1. Tabla de usuarios
        await pool.query(`
            CREATE TABLE IF NOT EXISTS usuarios (
                dni VARCHAR(20) PRIMARY KEY,
                pass VARCHAR(100) NOT NULL,
                jerarquia VARCHAR(50),
                apellido VARCHAR(100) NOT NULL,
                nombre VARCHAR(100) NOT NULL,
                creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        // 2. Tabla de reservas
        await pool.query(`
            CREATE TABLE IF NOT EXISTS reservas (
                id SERIAL PRIMARY KEY,
                clave_mes VARCHAR(10) NOT NULL,
                id_fecha INT NOT NULL,
                tipo_guardia VARCHAR(20) NOT NULL,
                dni_agente VARCHAR(20) REFERENCES usuarios(dni) ON DELETE CASCADE,
                reservado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (clave_mes, id_fecha, tipo_guardia)
            );
        `);

        console.log('>>> Base de datos PostgreSQL conectada y tablas ("usuarios" y "reservas") verificadas.');
    } catch (error) {
        console.error('>>> Error al inicializar la base de datos PostgreSQL:', error);
    }
}
inicializarBaseDatos();

// -------------------------------------------------------------
// MIDDLEWARES Y RUTAS HTTP
// -------------------------------------------------------------
app.use(express.json());

// Middleware de seguridad para la sección /admin
const seguridadAdmin = basicAuth({
    users: { 'esub': '*guardia/9595' },
    challenge: true,
    unauthorizedResponse: 'Acceso no autorizado al Panel de Control de la PNA.'
});

// Ruta del panel de administración (Protegida)
app.get('/admin', seguridadAdmin, (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// Ruta de exportación a Excel (Protegida)
app.get('/admin/exportar-excel', seguridadAdmin, async (req, res) => {
    const mes = req.query.mes || (new Date().getMonth() + 1);
    const año = req.query.año || new Date().getFullYear();
    const claveMes = `${año}-${String(mes).padStart(2, '0')}`;

    try {
        const query = `
            SELECT r.id_fecha, r.tipo_guardia, r.reservado_en,
                   u.dni, u.jerarquia, u.apellido, u.nombre
            FROM reservas r
            JOIN usuarios u ON TRIM(r.dni_agente) = TRIM(u.dni)
            WHERE r.clave_mes = $1
            ORDER BY r.id_fecha ASC, r.tipo_guardia ASC;
        `;
        const result = await pool.query(query, [claveMes]);

        const workbook = new exceljs.Workbook();
        const worksheet = workbook.addWorksheet(`Guardias ${claveMes}`);

        worksheet.columns = [
            { header: 'Fecha Guardia', key: 'dia', width: 15 },
            { header: 'Tipo de Guardia', key: 'tipo', width: 20 },
            { header: 'Jerarquía', key: 'jerarquia', width: 18 },
            { header: 'Apellido', key: 'apellido', width: 20 },
            { header: 'Nombre', key: 'nombre', width: 20 },
            { header: 'DNI', key: 'dni', width: 15 },
            { header: 'Fecha de Reserva', key: 'reservadoEn', width: 22 }
        ];

        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFF' } };
        worksheet.getRow(1).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '0D6EFD' }
        };

        const strMes = String(mes).padStart(2, '0');

        result.rows.forEach(row => {
            const strDia = String(row.id_fecha).padStart(2, '0');
            // Construimos la fecha en formato DD/MM/YYYY
            const fechaGuardia = `${strDia}/${strMes}/${año}`;

            worksheet.addRow({
                dia: fechaGuardia,
                tipo: row.tipo_guardia === 'oficial' ? 'Guardia Oficial' : 'Guardia Disponible',
                jerarquia: row.jerarquia || '',
                apellido: row.apellido || '',
                nombre: row.nombre || '',
                dni: row.dni || '',
                reservadoEn: row.reservado_en ? new Date(row.reservado_en).toLocaleString('es-AR') : ''
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Guardias_${claveMes}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();

    } catch (error) {
        console.error('>>> Error en /admin/exportar-excel:', error);
        res.status(500).send('Error al generar la planilla de Excel.');
    }
});

// Servir archivos estáticos de la carpeta public
app.use(express.static(path.join(__dirname, 'public')));

// -------------------------------------------------------------
// HELPER PARA ESTRUCTURAR EL MES Y CONSULTAR RESERVAS EN BD
// -------------------------------------------------------------
const contadoresDesreserva = {};

function generarPlantillaMes(mes, año) {
    const oficiales = [];
    const disponibles = [];
    const totalDias = new Date(año, mes, 0).getDate();

    const fechaObjMes = new Date(año, mes - 1, 1);
    let nombreMes = fechaObjMes.toLocaleString('es-AR', { month: 'long' });
    nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

    const strMes = String(mes).padStart(2, '0');

    for (let dia = 1; dia <= totalDias; dia++) {
        const strDia = String(dia).padStart(2, '0');
        const fechaObj = new Date(año, mes - 1, dia, 12, 0, 0);

        let diaNombre = fechaObj.toLocaleString('es-AR', { weekday: 'short', timeZone: 'America/Argentina/Buenos_Aires' });
        diaNombre = diaNombre.replace('.', '');
        diaNombre = diaNombre.charAt(0).toUpperCase() + diaNombre.slice(1);

        const fechaTexto = `${diaNombre} ${strDia}/${strMes}`;

        oficiales.push({ id: dia, fecha: fechaTexto, estado: 'disponible', agente: null, reservadoEn: null });
        disponibles.push({ id: dia, fecha: fechaTexto, estado: 'disponible', agente: null, reservadoEn: null });
    }

    return {
        infoMes: { mesNombre: nombreMes, año: Number(año), totalDias, mesNumero: Number(mes) },
        oficiales,
        disponibles
    };
}

// Función asíncrona que cruza la estructura del mes con las reservas de PostgreSQL
async function obtenerDatosMesPersistidos(mes, año) {
    const clave = `${año}-${String(mes).padStart(2, '0')}`;
    const datosMes = generarPlantillaMes(mes, año);

    try {
        const query = `
            SELECT r.id_fecha, r.tipo_guardia, r.reservado_en,
                   u.dni, u.jerarquia, u.apellido, u.nombre
            FROM reservas r
            JOIN usuarios u ON TRIM(r.dni_agente) = TRIM(u.dni)
            WHERE r.clave_mes = $1
        `;
        const result = await pool.query(query, [clave]);

        result.rows.forEach(row => {
            const lista = row.tipo_guardia === 'oficial' ? datosMes.oficiales : datosMes.disponibles;
            const fechaItem = lista.find(item => item.id === row.id_fecha);
            if (fechaItem) {
                fechaItem.estado = 'reservado';
                fechaItem.agente = {
                    dni: row.dni,
                    jerarquia: row.jerarquia,
                    apellido: row.apellido,
                    nombre: row.nombre
                };
                fechaItem.reservadoEn = row.reservado_en;
            }
        });
    } catch (error) {
        console.error('>>> Error al recuperar reservas de PostgreSQL:', error);
    }

    return datosMes;
}

// -------------------------------------------------------------
// WEBSOCKETS EN TIEMPO REAL CON PERSISTENCIA EN BD
// -------------------------------------------------------------
io.on('connection', (socket) => {

    // REGISTRO DE NUEVO USUARIO
    socket.on('solicitarRegistro', async (data) => {
        const { dni, pass, jerarquia, apellido, nombre } = data;
        const dniClean = String(dni).trim();

        try {
            const existeRes = await pool.query('SELECT dni FROM usuarios WHERE TRIM(dni) = $1', [dniClean]);
            if (existeRes.rows.length > 0) {
                socket.emit('resultadoRegistro', {
                    exito: false,
                    mensaje: 'El DNI ingresado ya se encuentra registrado en el sistema.'
                });
                return;
            }

            const queryInsert = `
                INSERT INTO usuarios (dni, pass, jerarquia, apellido, nombre)
                VALUES ($1, $2, $3, $4, $5)
                RETURNING dni, jerarquia, apellido, nombre;
            `;
            const values = [
                dniClean,
                String(pass).trim(),
                jerarquia ? jerarquia.trim() : '',
                apellido ? apellido.trim().toUpperCase() : '',
                nombre ? nombre.trim().toUpperCase() : ''
            ];

            const result = await pool.query(queryInsert, values);
            socket.emit('resultadoRegistro', {
                exito: true,
                mensaje: '¡Registro exitoso! Ya puedes iniciar sesión con tu DNI.',
                usuario: result.rows[0]
            });
        } catch (error) {
            console.error('Error en solicitarRegistro:', error);
            socket.emit('resultadoRegistro', {
                exito: false,
                mensaje: 'Error de servidor al procesar el registro.'
            });
        }
    });

    // INICIO DE SESIÓN
    socket.on('solicitarLogin', async (data) => {
        const { dni, pass } = data;
        const dniClean = String(dni).trim();
        const passClean = String(pass).trim();

        try {
            const userRes = await pool.query('SELECT * FROM usuarios WHERE TRIM(dni) = $1', [dniClean]);

            if (userRes.rows.length === 0) {
                socket.emit('resultadoLogin', {
                    exito: false,
                    mensaje: 'El usuario no está registrado. Debe registrarse primero.'
                });
                return;
            }

            const usuario = userRes.rows[0];

            if (usuario.pass.trim() !== passClean) {
                socket.emit('resultadoLogin', {
                    exito: false,
                    mensaje: 'Contraseña incorrecta. Intente nuevamente.'
                });
                return;
            }

            socket.emit('resultadoLogin', {
                exito: true,
                mensaje: `¡Bienvenido/a ${usuario.jerarquia} ${usuario.apellido}!`,
                usuario: {
                    dni: usuario.dni,
                    jerarquia: usuario.jerarquia,
                    apellido: usuario.apellido,
                    nombre: usuario.nombre
                }
            });
        } catch (error) {
            console.error('Error en solicitarLogin:', error);
            socket.emit('resultadoLogin', {
                exito: false,
                mensaje: 'Error en el servidor al intentar iniciar sesión.'
            });
        }
    });

    // RESTABLECER / CAMBIAR CONTRASEÑA
    socket.on('solicitarReseteoPass', async (data) => {
        const { dni, nuevaPass } = data;
        const dniClean = String(dni).trim();
        const nuevaPassClean = String(nuevaPass).trim();

        if (!dniClean || !nuevaPassClean) {
            socket.emit('resultadoReseteoPass', {
                exito: false,
                mensaje: 'Por favor, completa todos los campos.'
            });
            return;
        }

        try {
            const userRes = await pool.query('SELECT dni FROM usuarios WHERE TRIM(dni) = $1', [dniClean]);

            if (userRes.rows.length === 0) {
                socket.emit('resultadoReseteoPass', {
                    exito: false,
                    mensaje: 'El DNI ingresado no se encuentra registrado.'
                });
                return;
            }

            await pool.query('UPDATE usuarios SET pass = $1 WHERE TRIM(dni) = $2', [nuevaPassClean, dniClean]);

            socket.emit('resultadoReseteoPass', {
                exito: true,
                mensaje: 'La contraseña ha sido actualizada correctamente. Inicia sesión con tu nueva clave.'
            });
        } catch (error) {
            console.error('Error en solicitarReseteoPass:', error);
            socket.emit('resultadoReseteoPass', {
                exito: false,
                mensaje: 'Error interno en el servidor al intentar actualizar la contraseña.'
            });
        }
    });

    // OBTENER FECHAS (Lectura desde PostgreSQL)
    socket.on('obtenerFechas', async (data) => {
        const { mes, año, dni } = data;
        const datosMes = await obtenerDatosMesPersistidos(mes, año);
        const clave = `${año}-${String(mes).padStart(2, '0')}`;
        
        if (!contadoresDesreserva[clave]) {
            contadoresDesreserva[clave] = {};
        }
        const cancelaciones = contadoresDesreserva[clave][dni] || 0;

        socket.emit('cargarFechas', {
            ...datosMes,
            cancelacionesUsadas: cancelaciones
        });
    });

    // SOLICITAR RESERVA (Escritura en PostgreSQL)
    socket.on('solicitarReserva', async (data) => {
        const { idFecha, tipoGuardia, dni, mes, año } = data;
        const dniClean = String(dni).trim();
        const clave = `${año}-${String(mes).padStart(2, '0')}`;

        try {
            // 1. Verificar si el usuario ya alcanzó el límite permitido según el tipo de guardia
            const checkUser = `
                SELECT id FROM reservas 
                WHERE clave_mes = $1 AND tipo_guardia = $2 AND TRIM(dni_agente) = $3
            `;
            const checkUserRes = await pool.query(checkUser, [clave, tipoGuardia, dniClean]);

            // Definimos el límite máximo: 1 para Guardia Oficial, 2 para Guardia Disponible
            const limiteMaximo = tipoGuardia === 'disponible' ? 2 : 1;

            if (checkUserRes.rows.length >= limiteMaximo) {
                socket.emit('resultadoReserva', {
                    exito: false,
                    mensaje: tipoGuardia === 'disponible'
                        ? 'Ya alcanzaste el límite máximo de 2 Guardias Disponibles asignadas para este mes.'
                        : 'Ya posees una reserva de Guardia Oficial asignada en este mes.'
                });
                return;
            }

            // 2. Insertar en la base de datos
            const insertQuery = `
                INSERT INTO reservas (clave_mes, id_fecha, tipo_guardia, dni_agente)
                VALUES ($1, $2, $3, $4)
            `;
            await pool.query(insertQuery, [clave, idFecha, tipoGuardia, dniClean]);

            // 3. Recuperar el mes actualizado y notificar a todos los clientes
            const datosActualizados = await obtenerDatosMesPersistidos(mes, año);

            socket.emit('resultadoReserva', { exito: true, mensaje: 'Reserva realizada con éxito.' });
            io.emit('actualizarFechas', datosActualizados);

        } catch (error) {
            console.error('>>> Error al registrar reserva en PostgreSQL:', error);
            socket.emit('resultadoReserva', {
                exito: false,
                mensaje: 'La fecha seleccionada ya no se encuentra disponible o fue ocupada por otro agente.'
            });
        }
    });

    // SOLICITAR DESRESERVA (Eliminación en PostgreSQL)
    socket.on('solicitarDesreserva', async (data) => {
        const { idFecha, tipoGuardia, dni, mes, año } = data;
        const dniClean = String(dni).trim();
        const clave = `${año}-${String(mes).padStart(2, '0')}`;

        if (!contadoresDesreserva[clave]) {
            contadoresDesreserva[clave] = {};
        }
        if (!contadoresDesreserva[clave][dniClean]) {
            contadoresDesreserva[clave][dniClean] = 0;
        }

        if (contadoresDesreserva[clave][dniClean] >= 3) {
            socket.emit('resultadoDesreserva', {
                exito: false,
                mensaje: 'Has alcanzado el límite máximo de 3 cancelaciones permitidas para este mes.'
            });
            return;
        }

        try {
            const deleteQuery = `
                DELETE FROM reservas 
                WHERE clave_mes = $1 AND id_fecha = $2 AND tipo_guardia = $3 AND TRIM(dni_agente) = $4
            `;
            const deleteRes = await pool.query(deleteQuery, [clave, idFecha, tipoGuardia, dniClean]);

            if (deleteRes.rowCount > 0) {
                contadoresDesreserva[clave][dniClean] += 1;
                const datosActualizados = await obtenerDatosMesPersistidos(mes, año);

                socket.emit('resultadoDesreserva', { exito: true, mensaje: 'La reserva ha sido cancelada satisfactoriamente.' });
                io.emit('actualizarFechas', datosActualizados);
            } else {
                socket.emit('resultadoDesreserva', { exito: false, mensaje: 'No fue posible cancelar la reserva seleccionada.' });
            }

        } catch (error) {
            console.error('>>> Error al eliminar reserva en PostgreSQL:', error);
            socket.emit('resultadoDesreserva', { exito: false, mensaje: 'Error interno al procesar la cancelación.' });
        }
    });
});

const PORT = process.env.PORT || 3000;
httpServer.listen(PORT, () => {
    console.log(`>>> Servidor corriendo en el puerto ${PORT}`);
});