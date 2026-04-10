// ============================================================================
// CONCILIADOR CORPORATIVO V3 - MÓDULO PRINCIPAL
// ============================================================================

// ========== UTILIDAD: LIMPIEZA DE MONTOS ==========
/**
 * Limpia y convierte valores monetarios de string a número
 * Maneja formatos como "1,234.56" o 1234.56
 */
function limpiarMonto(valor) {
    if (valor === null || valor === undefined) return 0;
    if (typeof valor === 'number') return valor;
    if (typeof valor === 'string') {
        const limpio = valor.replace(/,/g, '').trim();
        const numero = parseFloat(limpio);
        return isNaN(numero) ? 0 : numero;
    }
    return 0;
}

/**
 * Normaliza un ID numérico a string limpio (elimina .0, espacios, etc.)
 * Maneja float64, int, string y valores nulos
 */
function normalizarId(valor) {
    if (valor === null || valor === undefined) return '';
    const str = valor.toString().trim();
    // Quitar .0 al final (IDs leídos como float64 por SheetJS)
    return str.replace(/\.0$/, '');
}

// ========== NORMALIZACIÓN DE DATOS ==========
const Normalizador = {
    /**
     * Detecta y salta filas de encabezado/resumen en Estado de Cuenta.
     * El archivo puede tener filas de resumen (INITIAL_BALANCE, CREDITS, etc.)
     * antes de los headers reales (RELEASE_DATE, TRANSACTION_TYPE, REFERENCE_ID...).
     */
    normalizarEstadoCuenta(sheet) {
        const COLUMNAS_ESPERADAS = ['REFERENCE_ID', 'TRANSACTION_NET_AMOUNT', 'TRANSACTION_TYPE'];

        // Intentar lectura directa primero
        let jsonData = XLSX.utils.sheet_to_json(sheet);
        if (jsonData.length > 0) {
            const cols = Object.keys(jsonData[0]);
            const tieneColumnas = COLUMNAS_ESPERADAS.every(c => cols.includes(c));
            if (tieneColumnas) {
                console.log('Estado de Cuenta: headers detectados en fila 0');
                return this._limpiarEstadoCuenta(jsonData);
            }
        }

        // Buscar la fila de headers reales (escanear primeras 10 filas)
        const range = XLSX.utils.decode_range(sheet['!ref']);
        for (let r = 1; r <= Math.min(10, range.e.r); r++) {
            const testData = XLSX.utils.sheet_to_json(sheet, { range: r });
            if (testData.length > 0) {
                const cols = Object.keys(testData[0]);
                const tieneColumnas = COLUMNAS_ESPERADAS.every(c => cols.includes(c));
                if (tieneColumnas) {
                    console.log(`Estado de Cuenta: headers detectados en fila ${r} (se saltaron ${r} filas de resumen)`);
                    return this._limpiarEstadoCuenta(testData);
                }
            }
        }

        // Fallback: usar lectura directa (dejar que falle más adelante si los datos son malos)
        console.warn('Estado de Cuenta: no se detectaron headers esperados, usando lectura directa');
        return this._limpiarEstadoCuenta(jsonData);
    },

    _limpiarEstadoCuenta(jsonData) {
        return jsonData.map(row => ({
            ...row,
            REFERENCE_ID: normalizarId(row.REFERENCE_ID),
            TRANSACTION_NET_AMOUNT: limpiarMonto(row.TRANSACTION_NET_AMOUNT),
            TRANSACTION_NET_AMOUNT_ORIGINAL: row.TRANSACTION_NET_AMOUNT,
            PARTIAL_BALANCE: limpiarMonto(row.PARTIAL_BALANCE)
        }));
    },

    /**
     * Normaliza datos de Liberaciones:
     * - Filtra filas sin SOURCE_ID (e.g. "Initial available balance")
     * - Convierte IDs float64 a string limpio
     * - Normaliza montos
     */
    normalizarLiberaciones(jsonData) {
        const filtrado = jsonData.filter(row => {
            const sourceId = row.SOURCE_ID;
            // Filtrar filas sin SOURCE_ID válido (balance inicial, totales, etc.)
            if (sourceId === null || sourceId === undefined || sourceId === '' || sourceId === 'undefined') return false;
            const str = sourceId.toString().trim();
            if (str === '' || str === 'NaN' || str === 'undefined') return false;
            return true;
        });

        const descartadas = jsonData.length - filtrado.length;
        if (descartadas > 0) {
            console.log(`Liberaciones: ${descartadas} filas descartadas (sin SOURCE_ID válido)`);
        }

        return filtrado.map(row => ({
            ...row,
            SOURCE_ID: normalizarId(row.SOURCE_ID),
            EXTERNAL_REFERENCE: normalizarId(row.EXTERNAL_REFERENCE),
            ORDER_ID: normalizarId(row.ORDER_ID),
            PACK_ID: normalizarId(row.PACK_ID),
            SHIPPING_ID: normalizarId(row.SHIPPING_ID),
            GROSS_AMOUNT: limpiarMonto(row.GROSS_AMOUNT),
            NET_CREDIT_AMOUNT: limpiarMonto(row.NET_CREDIT_AMOUNT),
            NET_DEBIT_AMOUNT: limpiarMonto(row.NET_DEBIT_AMOUNT),
            MP_FEE_AMOUNT: limpiarMonto(row.MP_FEE_AMOUNT),
            SHIPPING_FEE_AMOUNT: limpiarMonto(row.SHIPPING_FEE_AMOUNT),
            FINANCING_FEE_AMOUNT: limpiarMonto(row.FINANCING_FEE_AMOUNT),
            TAXES_AMOUNT: limpiarMonto(row.TAXES_AMOUNT),
            COUPON_AMOUNT: limpiarMonto(row.COUPON_AMOUNT),
            SELLER_AMOUNT: limpiarMonto(row.SELLER_AMOUNT)
        }));
    }
};

// ========== CATEGORÍAS GLOBALES ESTADO DE CUENTA ==========

/**
 * Normaliza texto eliminando acentos/diacríticos para comparación robusta.
 * "Liberación" → "Liberacion", "automático" → "automatico"
 */
function normalizarTexto(texto) {
    if (!texto) return '';
    return texto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

// Solo se usan para detectar REFERENCE_IDs que NO son pagos (no se enriquecen con API)
const CATEGORIAS_NO_PAGO = {
    FISCAL:        /Impuesto sobre la renta|Reintegro de impuestos|retenciones de ISR|retencion de IVA/i,
    TRANSFERENCIA: /^Transferencia (enviada|programada)/i,
};

// Patron que indica que un REFERENCE_ID es definitivamente un pago ML
const PATRON_TIPO_PAGO = /^(Liberacion de dinero|Dinero recibido|Entrada de dinero|Devolucion de dinero)/i;

/**
 * Determina si un REFERENCE_ID es un pago (debe enriquecerse con la API de ML)
 * o si es una transacción no-pago (TRANSFERENCIA, FISCAL) clasificable solo por TRANSACTION_TYPE.
 *
 * Reglas:
 * 1. TRANSFERENCIA siempre gana — es un retiro bancario, no un payment ID
 * 2. Si hay algún tipo de movimiento de pago (Liberación, Dinero recibido, Devolución...)
 *    → PAGO, aunque también haya tipos fiscales (son ajustes sobre el mismo payment)
 * 3. Solo si no hay tipos de pago → evaluar FISCAL
 *
 * @param {string[]} transactionTypes
 * @returns {'TRANSFERENCIA'|'FISCAL'|'PAGO'}
 */
function detectarTipoReferencia(transactionTypes) {
    const tipos = transactionTypes.map(t => normalizarTexto(t));

    // Transferencia siempre es no-pago
    if (tipos.some(t => CATEGORIAS_NO_PAGO.TRANSFERENCIA.test(t))) return 'TRANSFERENCIA';

    // Si hay algún tipo que indica movimiento de dinero de pago → es PAGO
    if (tipos.some(t => PATRON_TIPO_PAGO.test(t))) return 'PAGO';

    // Solo si no hay tipos de pago, clasificar como FISCAL
    if (tipos.some(t => CATEGORIAS_NO_PAGO.FISCAL.test(t))) return 'FISCAL';

    return 'PAGO';
}

/**
 * Clasifica una orden a partir del status y status_detail devueltos por la API de ML.
 * @param {string} status - Campo 'status' del payment
 * @param {string} statusDetail - Campo 'status_detail' del payment
 * @returns {string} Estatus del negocio
 */
function clasificarPorStatusML(status, statusDetail) {
    const s  = (status       || '').toLowerCase();
    const sd = (statusDetail || '').toLowerCase();

    if (s === 'approved'     && sd === 'accredited')    return 'VENTA';
    if (s === 'refunded'     && sd === 'bpp_refunded')  return 'DEVOLUCION_MP';
    if (s === 'refunded'     && sd === 'refunded')      return 'DEVOLUCION_VENDEDOR';
    if (s === 'refunded'     && sd === 'bpp_covered')   return 'DEVOLUCION_CUBIERTA';
    if (s === 'in_mediation' && sd === 'pending')       return 'MEDIACION_ABIERTA';
    if (s === 'charged_back' && sd === 'in_process')    return 'CONTRACARGO_EN_PROCESO';

    // Fallbacks por status principal
    if (s === 'approved')     return 'VENTA';
    if (s === 'refunded')     return 'DEVOLUCION_VENDEDOR';
    if (s === 'in_mediation') return 'MEDIACION_ABIERTA';
    if (s === 'charged_back') return 'CONTRACARGO_EN_PROCESO';

    return 'REVISAR';
}

// ========== CONFIGURACIÓN ==========
const CONFIG = {
    // Configuración de tolerancias y tasas de impuestos
    toleranciaNeto: 0.05,
    toleranciaSaldo: 1.00,
    tasaIVA: 0.08,
    
    STORAGE_KEYS: {
        CASOS: 'conciliador_v3_casos',
        HISTORIAL: 'conciliador_v3_historial'
    },
    ENDPOINTS: {
        development: 'https://api.atlasdeldescanso.com',
        production: 'https://api.atlasdeldescanso.com'
    },
    WORKFLOWS: {
        DEVOLUCION_PENDIENTE: {
            descripcion: 'Devolución con factura concluida - requiere cancelación contable',
            pasos: [
                { texto: 'Confirmar recepción física de mercancía en almacén', obligatorio: true },
                { texto: 'Verificar condición de la mercancía recibida', obligatorio: true },
                { texto: 'Cancelar factura en Intelisis', obligatorio: true },
                { texto: 'Registrar Nota de Crédito correspondiente', obligatorio: true },
                { texto: 'Actualizar inventario con mercancía devuelta', obligatorio: true },
                { texto: 'Notificar a finanzas sobre NC generada', obligatorio: false }
            ],
            documentosRequeridos: ['Factura', 'Guía de devolución ML'],
            sla: '24 horas'
        },
        DEVOLUCION_ORDEN_DEVUELTA: {
            descripcion: 'Orden devuelta sin facturar - solo cancelar pedido',
            pasos: [
                { texto: 'Verificar estado de cancelación en MercadoLibre', obligatorio: true },
                { texto: 'Confirmar que NO se generó factura', obligatorio: true },
                { texto: 'Cancelar pedido en Intelisis', obligatorio: true },
                { texto: 'Documentar razón de la devolución', obligatorio: false }
            ],
            documentosRequeridos: ['Comprobante de cancelación ML'],
            sla: '12 horas'
        },
        DEVOLUCION_EN_PROCESO: {
            descripcion: 'Cliente mencionó devolución pero orden aún activa',
            pasos: [
                { texto: 'Contactar al cliente para confirmar intención', obligatorio: true },
                { texto: 'Verificar estado actual en MercadoLibre', obligatorio: true },
                { texto: 'Esperar confirmación formal de devolución', obligatorio: true },
                { texto: 'Re-evaluar cuando cambie el estado en ML', obligatorio: true }
            ],
            documentosRequeridos: [],
            sla: '48 horas'
        },
        DISCREPANCIA: {
            descripcion: 'Diferencia de montos entre ML e Intelisis',
            pasos: [
                { texto: 'Comparar factura vs orden de MercadoLibre línea por línea', obligatorio: true },
                { texto: 'Verificar ajustes, descuentos o cupones en ML', obligatorio: true },
                { texto: 'Revisar comisiones de MercadoLibre aplicadas', obligatorio: true },
                { texto: 'Si persiste diferencia, contactar soporte ML', obligatorio: false },
                { texto: 'Ajustar factura en Intelisis o registrar NC/ND según corresponda', obligatorio: true }
            ],
            documentosRequeridos: ['Factura', 'Detalle de orden ML', 'Comprobante de comisiones'],
            sla: '48 horas'
        },
        ANTICIPO: {
            descripcion: 'Cliente pagó pero aún no se factura - registrar anticipo CxC',
            pasos: [
                { texto: 'Verificar forma de cobro en MercadoLibre', obligatorio: true },
                { texto: 'Crear anticipo en módulo CxC de Intelisis', obligatorio: true },
                { texto: 'Vincular anticipo con pedido pendiente', obligatorio: true },
                { texto: 'Programar aplicación automática al generar factura', obligatorio: true },
                { texto: 'Verificar que referencia ML quede registrada', obligatorio: true }
            ],
            documentosRequeridos: ['Comprobante de pago ML', 'Pedido pendiente'],
            sla: '24 horas'
        },
        COBRO_FACTURA: {
            descripcion: 'Factura concluida y pago recibido - registrar cobro',
            pasos: [
                { texto: 'Verificar factura concluida en Intelisis', obligatorio: true },
                { texto: 'Confirmar pago recibido en MercadoLibre', obligatorio: true },
                { texto: 'Registrar cobro en módulo CxC', obligatorio: true },
                { texto: 'Aplicar cobro a saldo pendiente de la factura', obligatorio: true },
                { texto: 'Verificar que saldo quede en cero', obligatorio: true }
            ],
            documentosRequeridos: ['Factura', 'Comprobante de pago ML'],
            sla: '24 horas'
        },
        PENDIENTE_ML: {
            descripcion: 'Orden pagada en ML pero sin pedido en Intelisis',
            pasos: [
                { texto: 'Verificar que orden esté realmente pagada en ML', obligatorio: true },
                { texto: 'Buscar si pedido existe con otro número/formato', obligatorio: true },
                { texto: 'Si no existe: crear pedido en Intelisis', obligatorio: true },
                { texto: 'Vincular orden ML con pedido creado', obligatorio: true },
                { texto: 'Proceder con facturación normal', obligatorio: true }
            ],
            documentosRequeridos: ['Orden completa de ML'],
            sla: '12 horas'
        },
        EN_DISPUTA: {
            descripcion: 'Orden en disputa activa con el comprador en MercadoLibre',
            pasos: [
                { texto: 'Revisar motivo de la disputa en MercadoLibre', obligatorio: true },
                { texto: 'Recopilar evidencia (fotos, guías, conversaciones)', obligatorio: true },
                { texto: 'Responder a la disputa dentro del plazo de ML', obligatorio: true },
                { texto: 'Documentar resolución esperada', obligatorio: true },
                { texto: 'Monitorear estado hasta resolución final', obligatorio: true }
            ],
            documentosRequeridos: ['Captura de disputa ML', 'Evidencia de envío', 'Conversaciones con cliente'],
            sla: '24 horas'
        },
        MEDIACION_GANADA: {
            descripcion: 'Mediación resuelta a favor del vendedor - documentar cierre',
            pasos: [
                { texto: 'Confirmar resolución favorable en MercadoLibre', obligatorio: true },
                { texto: 'Verificar que el pago se haya liberado correctamente', obligatorio: true },
                { texto: 'Documentar la resolución en el expediente', obligatorio: true },
                { texto: 'Archivar evidencia utilizada', obligatorio: false }
            ],
            documentosRequeridos: ['Resolución de ML', 'Comprobante de liberación de pago'],
            sla: '48 horas'
        },
        CONTRACARGO: {
            descripcion: 'Contracargo bancario - requiere atención urgente',
            pasos: [
                { texto: 'Identificar transacción afectada por el contracargo', obligatorio: true },
                { texto: 'Recopilar toda la evidencia de la venta (factura, guía, entrega)', obligatorio: true },
                { texto: 'Presentar documentación ante MercadoLibre', obligatorio: true },
                { texto: 'Dar seguimiento al proceso de disputa bancaria', obligatorio: true },
                { texto: 'Registrar provisión contable si es necesario', obligatorio: true },
                { texto: 'Actualizar estado según resolución final', obligatorio: true }
            ],
            documentosRequeridos: ['Factura', 'Guía de envío', 'Comprobante de entrega', 'Estado de cuenta'],
            sla: '12 horas'
        },
        SIN_COBRO: {
            descripcion: 'Orden sin cobro efectivo - neto en cero o negativo',
            pasos: [
                { texto: 'Verificar motivo del neto en cero (cancelación, devolución total)', obligatorio: true },
                { texto: 'Confirmar si hay pedido/factura en Intelisis', obligatorio: true },
                { texto: 'Si existe factura: proceder con cancelación/NC', obligatorio: true },
                { texto: 'Documentar razón del no cobro', obligatorio: true }
            ],
            documentosRequeridos: ['Detalle de movimientos ML'],
            sla: '24 horas'
        },
        REVISAR: {
            descripcion: 'Caso que requiere revisión manual - no clasificado automáticamente',
            pasos: [
                { texto: 'Analizar detalle de movimientos en MercadoLibre', obligatorio: true },
                { texto: 'Determinar tipo real de caso (devolución, disputa, etc.)', obligatorio: true },
                { texto: 'Reclasificar caso si es necesario', obligatorio: true },
                { texto: 'Aplicar workflow correspondiente', obligatorio: true }
            ],
            documentosRequeridos: ['Detalle completo de la orden ML'],
            sla: '48 horas'
        }
    },
    PRIORIDADES: {
        URGENTE: { color: 'var(--danger)', label: 'Urgentes', order: 0 },
        IMPORTANTE: { color: 'var(--warning)', label: 'Importantes', order: 1 },
        BAJA: { color: 'var(--success)', label: 'Revisar', order: 2 }
    },
    ROLES: {
        CONTADOR: { color: '#8b5cf6', label: 'Contador' },
        FINANZAS: { color: '#0891b2', label: 'Finanzas' },
        LOGISTICA: { color: '#ea580c', label: 'Logística' }
    }
};

// ========== ESTADO GLOBAL ==========
class EstadoGlobal {
    constructor() {
        this.casosPendientes = [];
        this.ordenesConsolidadas = [];
        this.archivo1Data = null;
        this.archivo2Data = null; // mantenido por compatibilidad, ya no se usa
        this.pagosCxCData = [];
        this.casoSeleccionado = null;
        this.filtroActual = { prioridad: null, responsable: 'todos', workflow: null, estado: null, busqueda: '', fecha: '' };
        this.endpointType = 'development';
        this.backendUrl = CONFIG.ENDPOINTS[this.endpointType];
        this.seguimientoApiDisponible = false;
        this.seguimientoApiUrl = 'https://api.atlasdeldescanso.com';
        this.estadisticasConciliacion = {
            encontradas: 0,
            noEncontradas: 0,
            conDiferencias: 0
        };
    }
}

const estado = new EstadoGlobal();

// ========== MÓDULO DE UTILIDADES ==========
const Utilidades = {
    formatMoney(value) {
        return '$' + parseFloat(value || 0).toLocaleString('es-MX', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        });
    },

    generarId() {
        return 'caso_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    },

    formatFecha(fecha) {
        if (!fecha) return '--';
        const d = new Date(fecha);
        return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
    },

    formatFechaHora(fecha) {
        if (!fecha) return '--';
        const d = new Date(fecha);
        return d.toLocaleString('es-MX', {
            day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
        });
    },

    roundMoney(value) {
        return Math.round((value || 0) * 100) / 100;
    },

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
};

// ========== MÓDULO DE PERSISTENCIA ==========
// ========== MÓDULO DE API PARA CASOS ==========
const ApiCasos = {
    async cargarCasosDesdeDB() {
        try {
            const url = `${estado.backendUrl}/seguimiento/casos`;
            console.log(`Consultando: ${url}`);

            const response = await apiFetch(url);

            if (!response.ok) {
                if (response.status === 404) {
                    console.warn('Endpoint /seguimiento/casos no encontrado - Agrega el endpoint GET al backend');
                    console.warn('📖 Ver archivo: endpoint_get_casos.py para el código');
                    Interfaz.mostrarToast('Endpoint de casos no encontrado. Usando modo local.', 'info');
                    return [];
                }
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }
            
            const data = await response.json();
            
            let casosArray = [];
            
            // Manejar tanto arrays como objetos agrupados por referencia
            if (Array.isArray(data)) {
                casosArray = data;
            } else if (typeof data === 'object' && data !== null) {
                // Convertir objeto a array
                casosArray = Object.values(data);
                console.log(`Convertido objeto agrupado a array: ${casosArray.length} casos`);
            } else {
                console.error('La respuesta del backend no es válida:', data);
                throw new Error('Respuesta inválida del backend');
            }
            
            console.log(`Cargados ${casosArray.length} casos desde la base de datos`);
            
            // Debug: mostrar estructura del primer caso
            if (casosArray.length > 0) {
                console.log('Estructura del primer caso del backend:', casosArray[0]);
            }
            
            // Mapear de la estructura del backend a la del frontend
            const casosMapeados = casosArray.map((caso, index) => {
                try {
                    return this.mapearCasoBackendAFrontend(caso);
                } catch (error) {
                    console.error(`Error mapeando caso ${index}:`, caso, error);
                    return null;
                }
            }).filter(caso => caso !== null);
            
            console.log(`${casosMapeados.length} casos mapeados correctamente`);
            return casosMapeados;
        } catch (error) {
            console.error('Error cargando casos desde DB:', error);
            
            // Si el endpoint no existe, usar localStorage como fallback
            if (error.message.includes('404')) {
                console.log('Intentando cargar desde localStorage como fallback...');
                try {
                    const localData = localStorage.getItem(CONFIG.STORAGE_KEYS.CASOS);
                    if (localData) {
                        const casos = JSON.parse(localData);
                        console.log(`Cargados ${casos.length} casos desde localStorage (fallback)`);
                        return casos;
                    }
                } catch (e) {
                    console.error('Error en fallback localStorage:', e);
                }
            }
            
            return [];
        }
    },

    async sincronizarCasosConDB() {
        try {
            // Convertir casos del frontend al formato del backend
            const casosParaBackend = estado.casosPendientes
                .map(caso => this.mapearCasoFrontendABackend(caso));

            const url = `${estado.backendUrl}/seguimiento/sincronizar`;
            console.log(`Sincronizando ${casosParaBackend.length} casos con el backend...`);

            const response = await apiFetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ casos: casosParaBackend })
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const resultado = await response.json();
            console.log(`Sincronización completada:`, resultado);
            
            Interfaz.mostrarToast(
                `Sincronizado: ${resultado.casosNuevos} nuevos, ${resultado.casosActualizados} actualizados`,
                'success'
            );

            // NO recargar casos automáticamente - esto puede causar bucles
            // Los casos ya están en memoria y sincronizados con la BD
            
            return resultado;
        } catch (error) {
            console.error('Error sincronizando casos:', error);
            console.warn('Guardando en localStorage como fallback');
            
            // Fallback a localStorage si el backend no está disponible
            try {
                localStorage.setItem(CONFIG.STORAGE_KEYS.CASOS, JSON.stringify(estado.casosPendientes));
                console.log('Casos guardados en localStorage (fallback)');
            } catch (e) {
                console.error('Error en fallback localStorage:', e);
            }
            
            throw error;
        }
    },

    async actualizarCasoIndividual(caso) {
        try {
            const casoBackend = this.mapearCasoFrontendABackend(caso);
            const url = `${estado.backendUrl}/seguimiento/casos/${encodeURIComponent(caso.orden.referencia)}`;

            const response = await apiFetch(url, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(casoBackend)
            });

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const casoActualizado = await response.json();
            console.log(`Caso actualizado en DB: ${caso.orden.referencia}`);
            
            return this.mapearCasoBackendAFrontend(casoActualizado);
        } catch (error) {
            console.error('Error actualizando caso en DB:', error);
            throw error;
        }
    },

    async resolverCaso(referencia, estadoFinal = 'resuelto') {
        try {
            const caso = estado.casosPendientes.find(c => c.orden.referencia === referencia);
            if (!caso) return;

            // Actualizar estado en el frontend
            caso.estado = 'RESUELTO';
            caso.fechaResolucion = new Date().toISOString();

            // Actualizar en la BD
            const url = `${estado.backendUrl}/seguimiento/casos/${encodeURIComponent(referencia)}`;
            await apiFetch(url, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    estadoSeguimiento: estadoFinal,
                    enUltimoProceso: false
                })
            });

            console.log(`Caso resuelto: ${referencia}`);
            Interfaz.mostrarToast(`Caso ${referencia} marcado como resuelto`, 'success');
        } catch (error) {
            console.error('Error resolviendo caso:', error);
            Interfaz.mostrarToast(`Error: ${error.message}`, 'error');
        }
    },

    mapearCasoFrontendABackend(caso) {
        return {
            referencia: caso.orden.referencia,
            tipoProblema: caso.tipo,
            estadoSeguimiento: caso.estado === 'RESUELTO' ? 'resuelto' : 
                              caso.estado === 'EN_REVISION' ? 'revision' : 'pendiente',
            montoNeto: caso.orden.monto,
            montoBruto: caso.orden.montoBruto || 0,
            costos: caso.orden.costos || 0,
            notas: caso.notas || '',
            enUltimoProceso: caso.estado !== 'RESUELTO',
            datosOrden: {
                sourceIds: caso.orden.sourceIds || [],
                metodoPago: caso.orden.metodoPago || '',
                fechaLiberacion: caso.orden.fechaLiberacion || '',
                estatus: caso.orden.estatus || '',
                statusML: caso.orden.statusML || '',
                statusDetailML: caso.orden.statusDetailML || '',
                pedido: caso.orden.pedido || '',
                factura: caso.orden.factura || '',
                cliente: caso.orden.cliente || '',
                responsable: caso.responsable || 'FINANZAS',
                // Información del ERP si existe
                erp: caso.erp || null
            },
            historial: caso.historial.map(h => ({
                fecha: h.fecha,
                tipo: h.usuario === 'SISTEMA' ? 'sistema' : 'manual',
                detalle: h.accion
            }))
        };
    },

    mapearCasoBackendAFrontend(caso) {
        // El backend usa camelCase: tipoProblema, estadoSeguimiento, montoNeto, etc.
        const datosOrden = caso.datosOrden || caso.datos_orden || {};
        
        // Validar y asignar valores por defecto
        const tipoProblema = caso.tipoProblema || caso.tipo_problema || 'REVISAR';
        const estadoSeguimiento = caso.estadoSeguimiento || caso.estado_seguimiento || 'pendiente';
        const montoNeto = caso.montoNeto ?? caso.monto_neto ?? 0;
        const montoBruto = caso.montoBruto ?? caso.monto_bruto ?? 0;
        const costos = caso.costos ?? 0;
        const referencia = caso.referencia || 'SIN-REF';
        
        const idStr = String(caso.id || '');
        const numeroCasoCalculado = /^\d+$/.test(idStr)
            ? '#' + idStr.padStart(4, '0')
            : '#TEMP';

        return {
            id: caso.id ? idStr : Utilidades.generarId(),
            numeroCaso: numeroCasoCalculado,
            tipo: tipoProblema,
            subtipo: null,
            estado: estadoSeguimiento === 'resuelto' ? 'RESUELTO' :
                    estadoSeguimiento === 'revision' ? 'EN_REVISION' : 'PENDIENTE',
            prioridad: Math.abs(montoNeto) > 5000 ? 'URGENTE' : 'IMPORTANTE',
            responsable: datosOrden.responsable || ((tipoProblema === 'RESERVA_EN_CONTRA' || tipoProblema === 'RESERVA_A_FAVOR') ? 'LOGISTICA' : 'FINANZAS'),
            titulo: `${tipoProblema} ${Utilidades.formatMoney(Math.abs(montoNeto))} - ${referencia}`,
            orden: {
                referencia: referencia,
                pedido: datosOrden.pedido || referencia,
                factura: datosOrden.factura || '',
                cliente: datosOrden.cliente || '',
                monto: montoNeto,
                montoBruto: montoBruto,
                costos: costos,
                fechaLiberacion: datosOrden.fechaLiberacion || '',
                estatus: datosOrden.estatus || '',
                sourceIds: datosOrden.sourceIds || [],
                metodoPago: datosOrden.metodoPago || ''
            },
            erp: datosOrden.erp || null,
            workflow: this.obtenerWorkflow(tipoProblema),
            accionesSugeridas: this.obtenerAcciones(tipoProblema),
            documentosRequeridos: this.obtenerDocumentos(tipoProblema),
            sla: '24 horas',
            notas: caso.notas || '',
            historial: (caso.historial || []).map(h => ({
                fecha: h.fecha || new Date().toISOString(),
                accion: h.detalle || h.accion || 'Sin detalle',
                usuario: h.tipo === 'sistema' ? 'SISTEMA' : 'Usuario'
            })),
            fechaCreacion: caso.fechaDeteccion || caso.fecha_deteccion || new Date().toISOString(),
            fechaUltimaActualizacion: caso.fechaUltimaActualizacion || caso.fecha_ultima_actualizacion || new Date().toISOString(),
            fechaResolucion: caso.fechaResolucion || caso.fecha_resolucion || null,
            creadoPor: 'SISTEMA'
        };
    },

    obtenerWorkflow(tipo) {
        // Mapeo especial para algunos tipos, el resto usa el mismo nombre
        const workflows = {
            'DEVOLUCION_MP': 'DEVOLUCION_PENDIENTE',
            'DEVOLUCION_VENDEDOR': 'DEVOLUCION_PENDIENTE',
            'DEVOLUCION_CUBIERTA': 'DEVOLUCION_PENDIENTE',
            'MEDIACION_ABIERTA': 'EN_DISPUTA',
            'CONTRACARGO_EN_PROCESO': 'CONTRACARGO',
            'DIFERENCIA_IMPORTE': 'DISCREPANCIA',
            'REVISAR': 'REVISAR'
        };
        return workflows[tipo] || tipo || 'REVISAR';
    },

    obtenerAcciones(tipo) {
        const workflow = CONFIG.WORKFLOWS[this.obtenerWorkflow(tipo)];
        return workflow ? workflow.pasos.map(p => ({
            texto: p.texto,
            obligatorio: p.obligatorio,
            completado: false
        })) : [];
    },

    obtenerDocumentos(tipo) {
        const workflow = CONFIG.WORKFLOWS[this.obtenerWorkflow(tipo)];
        return workflow ? workflow.documentosRequeridos : [];
    }
};

// ========== MÓDULO DE PERSISTENCIA (MIGRADO A API) ==========
const Persistencia = {
    async cargarCasos() {
        try {
            // Cargar desde la base de datos
            const casos = await ApiCasos.cargarCasosDesdeDB();
            estado.casosPendientes = casos;
            console.log(`${casos.length} casos cargados desde la base de datos`);
        } catch (error) {
            console.error('Error cargando casos:', error);
            estado.casosPendientes = [];
        }
    },

    async guardarCasos() {
        try {
            // Solo guardar en localStorage (rápido)
            // La sincronización con BD se hace manualmente cuando sea necesario
            localStorage.setItem(CONFIG.STORAGE_KEYS.CASOS, JSON.stringify(estado.casosPendientes));
        } catch (error) {
            console.error('Error guardando casos en localStorage:', error);
        }
    },

    async sincronizarCasosConBD() {
        // Función separada para sincronizar con la base de datos
        // Solo llamar cuando realmente sea necesario
        try {
            console.log('Iniciando sincronización con base de datos...');
            await ApiCasos.sincronizarCasosConDB();
            console.log('Sincronización con BD completada');
        } catch (error) {
            console.error('Error sincronizando con BD:', error);
            Interfaz.mostrarToast('Error sincronizando con base de datos', 'error');
        }
    },

    async guardarHistorialConciliacion(info) {
        // localStorage fallback
        try {
            let historial = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEYS.HISTORIAL) || '[]');
            historial.unshift(info);
            historial = historial.slice(0, 50);
            localStorage.setItem(CONFIG.STORAGE_KEYS.HISTORIAL, JSON.stringify(historial));
        } catch (e) {
            console.error('Error guardando historial en localStorage:', e);
        }

        // Guardar en BD vía API
        try {
            const ordenes = estado.ordenesConsolidadas || [];
            // Calcular resumen por estatus
            const resumenPorEstatus = {};
            let totalIngresosBruto = 0, totalCostos = 0, totalNetoReal = 0, ordenesValidadas = 0;
            ordenes.forEach(o => {
                const est = o.estatus || 'DESCONOCIDO';
                if (!resumenPorEstatus[est]) resumenPorEstatus[est] = { cantidad: 0, monto: 0 };
                resumenPorEstatus[est].cantidad++;
                resumenPorEstatus[est].monto += (o.neto?.real || 0);
                totalIngresosBruto += (o.ingresos?.bruto || 0);
                totalCostos += (o.costos?.total || 0);
                totalNetoReal += (o.neto?.real || 0);
                if (o.validacion?.saldoValidado) ordenesValidadas++;
            });
            const montoEstadoCuenta = (estado.archivo1Data || []).reduce((sum, row) => {
                const val = row.TRANSACTION_NET_AMOUNT;
                return sum + (typeof val === 'number' ? val : limpiarMonto(val));
            }, 0);

            const payload = {
                totalOrdenes: info.totalOrdenes || ordenes.length,
                coincidencias: info.coincidencias || 0,
                casosCreados: info.casosCreados || 0,
                ordenesValidadas,
                montoEstadoCuenta,
                resumenPorEstatus,
                totalIngresosBruto,
                totalCostos,
                totalNetoReal,
                ordenes: ordenes.map(o => ({
                    referenciaERP: o.referenciaERP || '',
                    estatus: o.estatus || '',
                    datosOrden: o
                })),
                notas: ''
            };

            const resp = await apiFetch(`${estado.backendUrl}/historico/conciliaciones`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            if (resp.ok) {
                const data = await resp.json();
                console.log('Conciliación guardada en BD:', data.id);
                // Sincronizar localStorage con la fecha real del backend
                try {
                    let histLocal = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEYS.HISTORIAL) || '[]');
                    histLocal[0] = { ...histLocal[0], fecha: data.fecha || new Date().toISOString() };
                    localStorage.setItem(CONFIG.STORAGE_KEYS.HISTORIAL, JSON.stringify(histLocal));
                } catch (_) {}
            } else {
                console.warn('Error guardando en BD:', resp.status);
            }
        } catch (e) {
            console.warn('No se pudo guardar en BD (API no disponible):', e.message);
        }
    }
};

// ========== AUTH: HELPER apiFetch ==========
async function apiFetch(url, options = {}) {
    const token = AuthManager.getToken();
    const headers = {
        ...(options.headers || {}),
        ...(token ? { 'Authorization': `Bearer ${token}` } : {})
    };
    const response = await fetch(url, { ...options, headers });
    if (response.status === 401) {
        AuthManager.logout();
        throw new Error('Sesión expirada. Por favor vuelve a iniciar sesión.');
    }
    return response;
}


// ========== AUTH MANAGER ==========
const AuthManager = {
    TOKEN_KEY: 'conciliador_token',
    USER_KEY:  'conciliador_user',

    getToken()        { return localStorage.getItem(this.TOKEN_KEY); },
    getUser()         { return JSON.parse(localStorage.getItem(this.USER_KEY) || 'null'); },
    isAuthenticated() { return !!this.getToken(); },

    async init() {
        if (!this.isAuthenticated()) { this.mostrarLogin(); return; }
        try {
            const resp = await apiFetch(`${estado.seguimientoApiUrl}/auth/me`);
            if (!resp.ok) throw new Error('token invalid');
            const user = await resp.json();
            this.onLoginSuccess(user, this.getToken(), true);
        } catch {
            this.logout(false);
        }
    },

    async login(event) {
        event.preventDefault();
        const btn = document.getElementById('login-btn');
        const errDiv = document.getElementById('login-error');
        btn.disabled = true;
        btn.textContent = 'Verificando...';
        errDiv.style.display = 'none';
        try {
            const resp = await fetch(`${estado.seguimientoApiUrl}/auth/login`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    username: document.getElementById('login-username').value.trim(),
                    password: document.getElementById('login-password').value
                })
            });
            if (!resp.ok) {
                const err = await resp.json().catch(() => ({}));
                throw new Error(err.detail || 'Credenciales incorrectas');
            }
            const data = await resp.json();
            this.onLoginSuccess(data.user, data.access_token, true);
        } catch (e) {
            errDiv.textContent = e.message;
            errDiv.style.display = 'block';
        } finally {
            btn.disabled = false;
            btn.textContent = 'Iniciar sesión';
        }
    },

    onLoginSuccess(user, token, cargarDatos = true) {
        localStorage.setItem(this.TOKEN_KEY, token);
        localStorage.setItem(this.USER_KEY, JSON.stringify(user));
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('app-layout').style.display  = 'flex';
        document.getElementById('header-user-nombre').textContent = user.nombreCompleto;
        const rolBadge = document.getElementById('header-user-rol');
        rolBadge.textContent = user.rol;
        rolBadge.className = `user-role-badge rol-${user.rol.toLowerCase()}`;
        document.getElementById('nav-btn-admin').style.display = user.rol === 'ADMIN' ? '' : 'none';
        this.aplicarRestriccionesPorRol(user.rol);
        if (cargarDatos) {
            Persistencia.cargarCasos().then(() => {
                CasosManager.renderizarSidebar();
                DashboardManager.actualizarDashboard();
            });
            DashboardManager.cargarUltimaConciliacion();
        }
    },

    logout(redirect = true) {
        localStorage.removeItem(this.TOKEN_KEY);
        localStorage.removeItem(this.USER_KEY);
        if (redirect) {
            document.getElementById('app-layout').style.display   = 'none';
            document.getElementById('login-screen').style.display = 'flex';
            const u = document.getElementById('login-username');
            const p = document.getElementById('login-password');
            if (u) u.value = '';
            if (p) p.value = '';
        }
    },

    mostrarLogin() {
        document.getElementById('login-screen').style.display = 'flex';
        document.getElementById('app-layout').style.display   = 'none';
    },

    togglePassword() {
        const inp = document.getElementById('login-password');
        inp.type = inp.type === 'password' ? 'text' : 'password';
    },

    aplicarRestriccionesPorRol(rol) {
        // VENTAS: ocultar tab Nueva Conciliación y Pagos CxC
        const btnConciliacion = document.querySelector('.nav-btn[data-view="conciliacion"]');
        if (btnConciliacion) {
            btnConciliacion.style.display = (rol === 'VENTAS') ? 'none' : '';
        }
    }
};


// ========== ADMIN MANAGER ==========
const AdminManager = {
    async cargarUsuarios() {
        try {
            const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios`);
            if (!resp.ok) throw new Error('Error cargando usuarios');
            const usuarios = await resp.json();
            this.renderizarTabla(usuarios);
        } catch (e) {
            document.getElementById('admin-tabla-body').innerHTML =
                `<tr><td colspan="6" style="text-align:center;color:var(--danger);padding:24px;">${e.message}</td></tr>`;
        }
    },

    renderizarTabla(usuarios) {
        const tbody = document.getElementById('admin-tabla-body');
        if (!usuarios.length) {
            tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;color:var(--gray-400);padding:24px;">No hay usuarios</td></tr>';
            return;
        }
        tbody.innerHTML = usuarios.map(u => {
            const lastLogin = u.lastLogin
                ? new Date(u.lastLogin).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' })
                : 'Nunca';
            const estadoBadge = u.activo
                ? '<span class="badge-activo">Activo</span>'
                : '<span class="badge-inactivo">Inactivo</span>';
            const rolBadge = `<span class="user-role-badge rol-${u.rol.toLowerCase()}">${u.rol}</span>`;
            return `<tr>
                <td><strong>${u.username}</strong></td>
                <td>${u.nombreCompleto}</td>
                <td>${rolBadge}</td>
                <td>${estadoBadge}</td>
                <td style="color:var(--gray-500);font-size:0.8rem;">${lastLogin}</td>
                <td>
                    <div class="admin-btn-group">
                        <button class="btn btn-secondary" style="font-size:0.75rem;padding:4px 10px"
                            onclick="AdminManager.abrirModalEditar(${JSON.stringify(u).replace(/"/g, '&quot;')})">Editar</button>
                        <button class="btn btn-secondary" style="font-size:0.75rem;padding:4px 10px"
                            onclick="AdminManager.resetPassword(${u.id}, '${u.username}')">Reset Pass</button>
                        ${u.activo
                            ? `<button class="btn" style="font-size:0.75rem;padding:4px 10px;background:#fef2f2;color:var(--danger);border:1px solid #fecaca"
                                onclick="AdminManager.toggleActivo(${u.id}, false, '${u.username}')">Desactivar</button>`
                            : `<button class="btn btn-secondary" style="font-size:0.75rem;padding:4px 10px"
                                onclick="AdminManager.toggleActivo(${u.id}, true, '${u.username}')">Activar</button>`
                        }
                    </div>
                </td>
            </tr>`;
        }).join('');
    },

    abrirModalCrear() {
        document.getElementById('admin-modal-titulo').textContent = 'Nuevo Usuario';
        document.getElementById('admin-modal-user-id').value = '';
        document.getElementById('admin-input-username').value = '';
        document.getElementById('admin-input-nombre').value = '';
        document.getElementById('admin-input-password').value = '';
        document.getElementById('admin-input-rol').value = 'FINANZAS';
        document.getElementById('admin-field-username').style.display = '';
        document.getElementById('admin-field-password').style.display = '';
        document.getElementById('admin-modal-overlay').style.display = 'flex';
        setTimeout(() => document.getElementById('admin-input-username').focus(), 50);
    },

    abrirModalEditar(usuario) {
        document.getElementById('admin-modal-titulo').textContent = 'Editar Usuario';
        document.getElementById('admin-modal-user-id').value = usuario.id;
        document.getElementById('admin-input-nombre').value = usuario.nombreCompleto;
        document.getElementById('admin-input-rol').value = usuario.rol;
        document.getElementById('admin-field-username').style.display = 'none';
        document.getElementById('admin-field-password').style.display = 'none';
        document.getElementById('admin-modal-overlay').style.display = 'flex';
    },

    cerrarModal(event) {
        if (!event || event.target === document.getElementById('admin-modal-overlay')) {
            document.getElementById('admin-modal-overlay').style.display = 'none';
        }
    },

    async guardarModal() {
        const id = document.getElementById('admin-modal-user-id').value;
        const nombre = document.getElementById('admin-input-nombre').value.trim();
        const rol = document.getElementById('admin-input-rol').value;
        if (!nombre) { Interfaz.mostrarToast('El nombre es requerido', 'error'); return; }

        try {
            if (id) {
                // Editar
                const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios/${id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ nombre_completo: nombre, rol })
                });
                if (!resp.ok) throw new Error((await resp.json()).detail || 'Error al guardar');
            } else {
                // Crear
                const username = document.getElementById('admin-input-username').value.trim();
                const password = document.getElementById('admin-input-password').value;
                if (!username || !password) { Interfaz.mostrarToast('Username y contraseña son requeridos', 'error'); return; }
                const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ username, nombre_completo: nombre, password, rol })
                });
                if (!resp.ok) throw new Error((await resp.json()).detail || 'Error al crear');
            }
            this.cerrarModal();
            Interfaz.mostrarToast('Usuario guardado', 'success');
            this.cargarUsuarios();
        } catch (e) {
            Interfaz.mostrarToast(e.message, 'error');
        }
    },

    async resetPassword(id, username) {
        if (!confirm(`¿Resetear contraseña de "${username}"?`)) return;
        try {
            const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios/${id}/reset-password`, {
                method: 'POST'
            });
            if (!resp.ok) throw new Error((await resp.json()).detail || 'Error');
            const data = await resp.json();
            alert(`Nueva contraseña temporal para ${data.username}:\n\n${data.tempPassword}\n\nCópiala ahora, no se mostrará de nuevo.`);
        } catch (e) {
            Interfaz.mostrarToast(e.message, 'error');
        }
    },

    async toggleActivo(id, activar, username) {
        const accion = activar ? 'activar' : 'desactivar';
        if (!confirm(`¿${accion.charAt(0).toUpperCase() + accion.slice(1)} al usuario "${username}"?`)) return;
        try {
            if (!activar) {
                const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios/${id}`, { method: 'DELETE' });
                if (!resp.ok) throw new Error((await resp.json()).detail || 'Error');
            } else {
                const resp = await apiFetch(`${estado.seguimientoApiUrl}/admin/usuarios/${id}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ activo: true })
                });
                if (!resp.ok) throw new Error((await resp.json()).detail || 'Error');
            }
            Interfaz.mostrarToast(`Usuario ${activar ? 'activado' : 'desactivado'}`, 'success');
            this.cargarUsuarios();
        } catch (e) {
            Interfaz.mostrarToast(e.message, 'error');
        }
    }
};


// ========== MÓDULO DE INTERFAZ ==========
const Interfaz = {
    mostrarToast(mensaje, tipo = 'info') {
        let toastContainer = document.getElementById('toast-container');
        if (!toastContainer) {
            toastContainer = document.createElement('div');
            toastContainer.id = 'toast-container';
            toastContainer.style.cssText = `
                position: fixed;
                top: 20px;
                right: 20px;
                z-index: 10000;
                display: flex;
                flex-direction: column;
                gap: 10px;
            `;
            document.body.appendChild(toastContainer);
        }

        const toast = document.createElement('div');
        toast.style.cssText = `
            background: ${tipo === 'success' ? 'var(--success)' : tipo === 'error' ? 'var(--danger)' : 'var(--info)'};
            color: white;
            padding: 12px 20px;
            border-radius: var(--radius);
            box-shadow: var(--shadow-lg);
            min-width: 250px;
            max-width: 400px;
            animation: slideIn 0.3s ease;
        `;
        toast.textContent = mensaje;

        toastContainer.appendChild(toast);

        setTimeout(() => {
            toast.style.animation = 'slideOut 0.3s ease';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    },

    toggleSidebar() {
        document.getElementById('sidebar').classList.toggle('collapsed');
    },

    cambiarVista(vista) {
        document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelector(`.nav-btn[data-view="${vista}"]`)?.classList.add('active');

        document.querySelectorAll('.main-body > section').forEach(section => {
            section.classList.remove('active');
        });
        document.getElementById(`view-${vista}`)?.classList.add('active');

        // Regenerar análisis cuando se navega a las vistas correspondientes
        if (vista === 'validacion' && estado.ordenesConsolidadas?.length > 0) {
            AnalisisValidacionManager.generarAnalisis();
        }
        if (vista === 'analisis-erp') {
            AnalisisERPManager.generarAnalisis();
        }
        if (vista === 'historico') {
            HistoricoManager.cargarHistorico();
        }
        if (vista === 'casos') {
            CasosManager.renderizarPortal();
            // Auto-actualizar ERP si hay casos con ERP encontrado pero embarqueInfo aún no fue consultado
            // Usamos 'in' para distinguir entre "nunca consultado" (clave ausente) vs "consultado y vacío" (null)
            const necesitaSync = estado.casosPendientes.some(
                c => c.estado !== 'RESUELTO' && c.erp?.encontrado && !('embarqueInfo' in (c.erp || {}))
            );
            if (necesitaSync) CasosManager.actualizarERPDeCasos(true);
        }
        if (vista === 'admin') {
            AdminManager.cargarUsuarios();
        }
    },

    cambiarEndpoint(tipo) {
        estado.endpointType = tipo;
        estado.backendUrl = CONFIG.ENDPOINTS[tipo];

        document.querySelectorAll('.endpoint-btn').forEach(btn => {
            if (btn.dataset.endpoint === tipo) {
                btn.classList.add('active');
                btn.style.background = 'var(--accent)';
                btn.style.color = 'white';
            } else {
                btn.classList.remove('active');
                btn.style.background = 'white';
                btn.style.color = 'var(--gray-700)';
            }
        });

        this.mostrarToast(`Endpoint cambiado a: ${tipo}`, 'info');
    },

    abrirModalImportar() {
        document.getElementById('inputImportarJSON').click();
    },

    cerrarModalImportar() {
        document.getElementById('inputImportarJSON').value = '';
    },

    setupUploadZones() {
        const zone  = document.getElementById('uploadZone1');
        const input = document.getElementById('fileInput1');
        const fileName = document.getElementById('fileName1');

        input.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.processFile(file);
                zone.classList.add('uploaded');
                fileName.textContent = file.name;
                fileName.style.display = 'block';
            }
        });
    },

    processFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];

                // Estado de Cuenta: detectar headers reales y normalizar
                estado.archivo1Data = Normalizador.normalizarEstadoCuenta(sheet);

                console.log('Estado de Cuenta cargado:', {
                    nombre: file.name,
                    registrosNormalizados: estado.archivo1Data.length,
                    columnas: Object.keys(estado.archivo1Data[0] || {})
                });

                const tieneReferenceId = estado.archivo1Data.some(row => row.REFERENCE_ID);
                console.log(`Tiene REFERENCE_ID: ${tieneReferenceId}`);

                if (estado.archivo1Data.length > 0) {
                    console.log('Primeros 3 registros:', estado.archivo1Data.slice(0, 3));
                }

                if (estado.archivo1Data && estado.archivo1Data.length > 0) {
                    document.getElementById('processBtn').disabled = false;
                    this.mostrarToast('Estado de Cuenta cargado correctamente', 'success');
                }
            } catch (error) {
                console.error('Error procesando archivo:', error);
                alert('Error al procesar archivo: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }
};

// ========== MÓDULO DE CASOS ==========
const CasosManager = {

    // Helper: devuelve tipo/workflowKey/prioridad/responsable según estatus de la orden
    _clasificarCaso(estatus, monto) {
        switch (estatus) {
            case 'DEVOLUCION_MP':
            case 'DEVOLUCION_VENDEDOR':
                return { tipo: estatus, workflowKey: 'DEVOLUCION_PENDIENTE', prioridad: Math.abs(monto) > 5000 ? 'URGENTE' : 'IMPORTANTE', responsable: 'LOGISTICA' };
            case 'DEVOLUCION_CUBIERTA':
                return { tipo: estatus, workflowKey: 'DEVOLUCION_PENDIENTE', prioridad: 'IMPORTANTE', responsable: 'LOGISTICA' };
            case 'MEDIACION_ABIERTA':
                return { tipo: estatus, workflowKey: 'EN_DISPUTA', prioridad: Math.abs(monto) > 10000 ? 'URGENTE' : 'IMPORTANTE', responsable: 'LOGISTICA' };
            case 'CONTRACARGO_EN_PROCESO':
                return { tipo: estatus, workflowKey: 'CONTRACARGO', prioridad: 'URGENTE', responsable: 'FINANZAS' };
            default:
                return { tipo: estatus, workflowKey: 'REVISAR', prioridad: 'BAJA', responsable: 'FINANZAS' };
        }
    },

    crearCasoAutomatico(orden) {
        // VENTA y TRANSFERENCIA son operaciones rutinarias, no generan caso
        if (orden.estatus === 'VENTA' || orden.estatus === 'TRANSFERENCIA' || orden.estatus === 'FISCAL') return null;

        const { tipo, workflowKey, prioridad, responsable } = this._clasificarCaso(orden.estatus, orden.ingresos.bruto);

        const workflow = CONFIG.WORKFLOWS[workflowKey] || CONFIG.WORKFLOWS.REVISAR;
        const ahora = new Date().toISOString();

        // Extraer información del ERP si está disponible
        const infoERP = orden.api ? {
            encontrado: orden.api.encontrado || false,
            estatusERP: orden.api.estatusERP || 'NO_ENCONTRADO',
            movimientos: orden.api.movimientos || 'N/A',
            importeERP: orden.api.importeERP || 0,
            diferencia: orden.api.diferencia || null,
            cliente: orden.api.cliente || '',
            movimientosDetalle: orden.api.movimientosDetalle || []
        } : null;

        // Extraer pedido y factura de los movimientos detallados
        let pedidoMovID = '';
        let facturaMovID = '';
        
        if (infoERP && infoERP.movimientosDetalle.length > 0) {
            const pedidos = infoERP.movimientosDetalle.filter(m =>
                m.mov && m.mov.toLowerCase().includes('pedido')
            );
            const facturas = infoERP.movimientosDetalle.filter(m =>
                m.mov && m.mov.toLowerCase().includes('factura')
            );
            
            if (pedidos.length > 0) pedidoMovID = pedidos[0].movID || '';
            if (facturas.length > 0) facturaMovID = facturas[0].movID || '';
        }

        return {
            id: Utilidades.generarId(),
            numeroCaso: '#TEMP',
            tipo,
            subtipo: null,
            estado: 'PENDIENTE',
            prioridad,
            responsable,
            titulo: `${tipo} ${Utilidades.formatMoney(Math.abs(orden.ingresos.bruto))} - ${orden.referenciaERP}`,
            orden: {
                referencia: orden.referenciaERP,
                pedido: pedidoMovID || orden.referenciaERP,
                factura: facturaMovID || '',
                cliente: infoERP?.cliente || '',
                monto: orden.ingresos.bruto,
                montoBruto: orden.ingresos.bruto,
                costos: orden.costos.total,
                fechaLiberacion: orden.fechaLiberacion,
                estatus: orden.estatus,
                statusML: orden.statusML || '',
                statusDetailML: orden.statusDetailML || '',
                sourceIds: orden.sourceIds,
                tieneMultiplesSourceIds: orden.tieneMultiplesSourceIds || false,
                validacionDetalle: {
                    diferencia: orden.validacion?.diferencia || 0,
                    saldoValidado: orden.validacion?.saldoValidado || false,
                    estadosCuentaDetalle: orden.validacion?.estadosCuentaDetalle || []
                }
            },
            // Información del ERP para seguimiento centralizado
            erp: infoERP,
            workflow: workflowKey,
            accionesSugeridas: workflow.pasos.map(p => ({
                texto: p.texto,
                obligatorio: p.obligatorio,
                completado: false
            })),
            documentosRequeridos: workflow.documentosRequeridos,
            sla: workflow.sla,
            notas: '',
            historial: [{
                fecha: ahora,
                accion: `Caso creado automáticamente - ${orden.estatus}`,
                usuario: 'SISTEMA'
            }],
            fechaCreacion: ahora,
            fechaUltimaActualizacion: ahora,
            creadoPor: 'SISTEMA'
        };
    },

    crearCasosAutomaticos() {
        const ahora = new Date().toISOString();
        const ESTATUS_RUTINARIOS = new Set(['VENTA', 'TRANSFERENCIA', 'FISCAL']);
        let casosCreados = 0;
        let casosResueltos = 0;
        let casosActualizados = 0;

        estado.ordenesConsolidadas.forEach(orden => {
            const esRutinario = ESTATUS_RUTINARIOS.has(orden.estatus);

            // Buscar caso abierto (no resuelto) para esta referencia
            const existente = estado.casosPendientes.find(
                c => c.orden.referencia === orden.referenciaERP && c.estado !== 'RESUELTO'
            );

            // ── 1. Orden rutinaria (VENTA / TRANSFERENCIA / FISCAL) ──────────────
            if (esRutinario) {
                if (existente) {
                    // La orden se liberó como rutinaria → resolver el caso automáticamente
                    existente.estado = 'RESUELTO';
                    existente.fechaUltimaActualizacion = ahora;
                    existente.historial.push({
                        fecha: ahora,
                        accion: `Caso resuelto automáticamente — orden liberada como ${orden.estatus}`,
                        usuario: 'SISTEMA'
                    });
                    casosResueltos++;
                }
                return; // nunca crear caso nuevo para estatus rutinarios
            }

            // ── 2. Caso existente ─────────────────────────────────────────────────
            if (existente) {
                // Siempre actualizar montos y validación con los datos frescos
                existente.orden.monto = orden.ingresos.bruto;
                existente.orden.montoBruto = orden.ingresos.bruto;
                existente.orden.costos = orden.costos.total;
                existente.orden.sourceIds = orden.sourceIds;
                existente.orden.tieneMultiplesSourceIds = orden.tieneMultiplesSourceIds || false;
                existente.orden.validacionDetalle = {
                    diferencia: orden.validacion?.diferencia || 0,
                    saldoValidado: orden.validacion?.saldoValidado || false,
                    estadosCuentaDetalle: orden.validacion?.estadosCuentaDetalle || []
                };
                existente.fechaUltimaActualizacion = ahora;

                // Detectar cambio de tipo (ej. MEDIACION_ABIERTA → DEVOLUCION_MP)
                if (existente.tipo !== orden.estatus) {
                    const tipoAnterior = existente.tipo;
                    const workflowAnterior = existente.workflow;
                    const attrs = this._clasificarCaso(orden.estatus, orden.ingresos.bruto);
                    const nuevoWorkflow = CONFIG.WORKFLOWS[attrs.workflowKey] || CONFIG.WORKFLOWS.REVISAR;

                    existente.tipo = attrs.tipo;
                    existente.prioridad = attrs.prioridad;
                    existente.responsable = attrs.responsable;
                    existente.workflow = attrs.workflowKey;

                    // Si el workflow cambió, resetear el checklist (el anterior ya no aplica)
                    if (workflowAnterior !== attrs.workflowKey) {
                        existente.accionesSugeridas = nuevoWorkflow.pasos.map(p => ({
                            texto: p.texto,
                            obligatorio: p.obligatorio,
                            completado: false
                        }));
                        existente.documentosRequeridos = nuevoWorkflow.documentosRequeridos;
                        existente.sla = nuevoWorkflow.sla;
                    }

                    existente.historial.push({
                        fecha: ahora,
                        accion: `Tipo actualizado: ${tipoAnterior} → ${orden.estatus}`,
                        usuario: 'SISTEMA'
                    });
                    casosActualizados++;
                }
                return;
            }

            // ── 3. Sin caso existente → crear nuevo ──────────────────────────────
            // Proteger casos ya resueltos: si ya existe uno resuelto para esta referencia, no reabrir
            const yaResuelto = estado.casosPendientes.some(
                c => c.orden.referencia === orden.referenciaERP && c.estado === 'RESUELTO'
            );
            if (yaResuelto) return;

            const nuevoCaso = this.crearCasoAutomatico(orden);
            if (nuevoCaso) {
                estado.casosPendientes.push(nuevoCaso);
                casosCreados++;
            }
        });

        Persistencia.guardarCasos();

        if (casosResueltos > 0)
            Interfaz.mostrarToast(`${casosResueltos} caso(s) resuelto(s) automáticamente por cambio de estatus`, 'success');
        if (casosActualizados > 0)
            Interfaz.mostrarToast(`${casosActualizados} caso(s) actualizado(s) por cambio de tipo`, 'info');

        return casosCreados; // mantiene compatibilidad con callers existentes
    },

    /**
     * Crea o actualiza casos de seguimiento para órdenes VENTA cuyo importe ML
     * difiere del importe registrado en el ERP (por encima de toleranciaSaldo).
     * Se llama después de conciliarConIntelisis(), cuando orden.api ya está poblado.
     * @returns {number} Número de casos nuevos creados
     */
    crearCasosVentaDiferencia() {
        const TOLERANCIA = CONFIG.toleranciaSaldo;
        const ahora = new Date().toISOString();
        let casosCreados = 0;

        estado.ordenesConsolidadas.forEach(orden => {
            if (orden.estatus !== 'VENTA') return;
            if (!orden.api || !orden.api.encontrado) return;
            if (orden.api.diferencia === null || orden.api.diferencia === undefined) return;
            if (Math.abs(orden.api.diferencia) <= TOLERANCIA) return;

            const diferencia = orden.api.diferencia;
            const existente = estado.casosPendientes.find(c => c.orden.referencia === orden.referenciaERP);

            if (existente) {
                // Actualizar datos del caso existente
                existente.erp = {
                    encontrado: true,
                    importeERP: orden.api.importeERP,
                    diferencia,
                    estatusERP: orden.api.estatusERP,
                    movimientos: orden.api.movimientos,
                    cliente: orden.api.cliente || '',
                    movimientosDetalle: orden.api.movimientosDetalle || [],
                    embarqueInfo: orden.api.embarqueInfo || {}
                };
                existente.orden.monto = orden.ingresos.bruto;
                existente.orden.montoBruto = orden.ingresos.bruto;
                existente.fechaUltimaActualizacion = ahora;
                return;
            }

            const workflow = CONFIG.WORKFLOWS.DISCREPANCIA;

            // Extraer pedido y factura de los movimientos del ERP
            const movsDet = orden.api.movimientosDetalle || [];
            const pedidoMovID  = (movsDet.find(m => m.mov?.toLowerCase().includes('pedido'))?.movID)  || '';
            const facturaMovID = (movsDet.find(m => m.mov?.toLowerCase().includes('factura'))?.movID) || '';

            const nuevoCaso = {
                id: Utilidades.generarId(),
                numeroCaso: '#TEMP',
                tipo: 'DIFERENCIA_IMPORTE',
                subtipo: diferencia < 0 ? 'IMPORTE_ERP_MAYOR' : 'IMPORTE_ML_MAYOR',
                estado: 'PENDIENTE',
                prioridad: Math.abs(diferencia) > 5000 ? 'URGENTE' : 'IMPORTANTE',
                responsable: 'FINANZAS',
                titulo: `DIFERENCIA ML vs ERP ${Utilidades.formatMoney(Math.abs(diferencia))} - ${orden.referenciaERP}`,
                orden: {
                    referencia: orden.referenciaERP,
                    pedido: pedidoMovID || orden.referenciaERP,
                    factura: facturaMovID || '',
                    cliente: orden.api.cliente || '',
                    monto: orden.ingresos.bruto,
                    montoBruto: orden.ingresos.bruto,
                    costos: orden.costos.total,
                    fechaLiberacion: orden.fechaLiberacion || '',
                    estatus: orden.estatus,
                    statusML: orden.statusML || '',
                    statusDetailML: orden.statusDetailML || '',
                    sourceIds: orden.sourceIds,
                    tieneMultiplesSourceIds: orden.tieneMultiplesSourceIds || false,
                    validacionDetalle: {
                        diferencia: orden.validacion?.diferencia || 0,
                        saldoValidado: orden.validacion?.saldoValidado || false
                    }
                },
                erp: {
                    encontrado: true,
                    importeERP: orden.api.importeERP,
                    diferencia,
                    estatusERP: orden.api.estatusERP,
                    movimientos: orden.api.movimientos,
                    cliente: orden.api.cliente || '',
                    movimientosDetalle: movsDet,
                    embarqueInfo: orden.api.embarqueInfo || {}
                },
                workflow: 'DISCREPANCIA',
                accionesSugeridas: workflow.pasos.map(p => ({ texto: p.texto, obligatorio: p.obligatorio, completado: false })),
                documentosRequeridos: workflow.documentosRequeridos,
                sla: workflow.sla,
                notas: `ML: ${Utilidades.formatMoney(orden.ingresos.bruto + (orden.costos?.shipping || 0))} | ERP: ${Utilidades.formatMoney(orden.api.importeERP)} | Diferencia: ${Utilidades.formatMoney(diferencia)}`,
                historial: [{
                    fecha: ahora,
                    accion: `Caso creado automáticamente — diferencia ML vs ERP: ${Utilidades.formatMoney(diferencia)}`,
                    usuario: 'SISTEMA'
                }],
                fechaCreacion: ahora,
                fechaUltimaActualizacion: ahora,
                creadoPor: 'SISTEMA'
            };

            estado.casosPendientes.push(nuevoCaso);
            casosCreados++;
            console.log(`Caso DIFERENCIA_IMPORTE creado: ${orden.referenciaERP} | dif=${diferencia}`);
        });

        if (casosCreados > 0) Persistencia.guardarCasos();
        return casosCreados;
    },

    actualizarCasosConERP() {
        let casosActualizados = 0;
        const ahora = new Date().toISOString();

        estado.casosPendientes.forEach(caso => {
            // Buscar la orden correspondiente en las órdenes consolidadas
            const orden = estado.ordenesConsolidadas.find(o => o.referenciaERP === caso.orden.referencia);
            
            if (orden && orden.api && orden.api.encontrado) {
                // Actualizar información del ERP
                caso.erp = {
                    encontrado: orden.api.encontrado,
                    estatusERP: orden.api.estatusERP || 'NO_DISPONIBLE',
                    movimientos: orden.api.movimientos || 'N/A',
                    importeERP: orden.api.importeERP || 0,
                    diferencia: orden.api.diferencia || null,
                    cliente: orden.api.cliente || '',
                    movimientosDetalle: orden.api.movimientosDetalle || [],
                    embarqueInfo: orden.api.embarqueInfo || {}
                };

                // Extraer pedido y factura de los movimientos detallados
                if (caso.erp.movimientosDetalle.length > 0) {
                    const pedidos = caso.erp.movimientosDetalle.filter(m =>
                        m.mov && m.mov.toLowerCase().includes('pedido')
                    );
                    const facturas = caso.erp.movimientosDetalle.filter(m =>
                        m.mov && m.mov.toLowerCase().includes('factura')
                    );
                    
                    if (pedidos.length > 0) {
                        caso.orden.pedido = pedidos[0].movID || caso.orden.pedido;
                    }
                    if (facturas.length > 0) {
                        caso.orden.factura = facturas[0].movID || '';
                    }
                    if (caso.erp.cliente) {
                        caso.orden.cliente = caso.erp.cliente;
                    }
                }

                // Agregar entrada al historial
                caso.historial.push({
                    fecha: ahora,
                    accion: `Caso actualizado con información de ERP - Estatus: ${caso.erp.estatusERP}`,
                    usuario: 'SISTEMA'
                });

                caso.fechaUltimaActualizacion = ahora;
                casosActualizados++;
            } else if (orden && (!orden.api || !orden.api.encontrado)) {
                // Marcar que no se encontró en ERP
                caso.erp = {
                    encontrado: false,
                    estatusERP: 'NO_ENCONTRADO',
                    movimientos: 'N/A',
                    importeERP: 0,
                    diferencia: null,
                    cliente: '',
                    movimientosDetalle: []
                };

                caso.historial.push({
                    fecha: ahora,
                    accion: 'Caso no encontrado en ERP - Requiere creación de pedido',
                    usuario: 'SISTEMA'
                });

                caso.fechaUltimaActualizacion = ahora;
                casosActualizados++;
            }
        });

        if (casosActualizados > 0) {
            Persistencia.guardarCasos();
            console.log(`${casosActualizados} casos actualizados con información del ERP`);
            Interfaz.mostrarToast(`${casosActualizados} casos actualizados con ERP`, 'success');
        }

        return casosActualizados;
    },

    filtrarCasosPorPrioridad(prioridad) {
        estado.filtroActual.prioridad = estado.filtroActual.prioridad === prioridad ? null : prioridad;

        // Compatibilidad con sidebar anterior (elementos pueden no existir)
        document.querySelectorAll('.sidebar-stat').forEach(el => el.classList.remove('active'));
        if (estado.filtroActual.prioridad) {
            document.querySelector(`.sidebar-stat[data-filter="${prioridad.toLowerCase()}"]`)?.classList.add('active');
        }

        this.renderizarSidebar();
    },

    filtrarCasosPorResponsable(responsable) {
        estado.filtroActual.responsable = responsable;

        // Compatibilidad con sidebar anterior
        document.querySelectorAll('.filter-chip:not([data-portal-filter])').forEach(el => el.classList.remove('active'));
        document.querySelector(`.filter-chip[data-filter="${responsable.toLowerCase()}"]`)?.classList.add('active');

        this.renderizarSidebar();
    },

    // ── Portal KPI filter ──────────────────────────────────────────────────────
    filtrarPortalPorPrioridad(prioridad) {
        estado.filtroActual.prioridad = estado.filtroActual.prioridad === prioridad ? null : prioridad;
        estado.filtroActual.estado = null; // limpiar filtro de estado al cambiar prioridad

        this.renderizarPortal();
    },

    filtrarPortalPorResponsable(responsable) {
        estado.filtroActual.responsable = responsable;

        document.querySelectorAll('[data-portal-filter]').forEach(el => el.classList.remove('active'));
        document.querySelector(`[data-portal-filter="${responsable}"]`)?.classList.add('active');

        this.renderizarPortal();
    },

    filtrarPortalPorWorkflow(workflow) {
        estado.filtroActual.workflow = workflow || null;
        this.renderizarPortal();
    },

    filtrarPortalPorEstado(estadoFiltro) {
        // null = Total Activos (sin filtro); toggle si ya está activo
        estado.filtroActual.estado = (estadoFiltro && estado.filtroActual.estado !== estadoFiltro)
            ? estadoFiltro
            : null;
        this.renderizarPortal();
    },

    async actualizarERPDeCasos(silencioso = false) {
        const casosActivos = estado.casosPendientes.filter(c => c.estado !== 'RESUELTO');
        if (casosActivos.length === 0) {
            if (!silencioso) Interfaz.mostrarToast('No hay casos activos para actualizar.', 'info');
            return;
        }

        const btn = document.getElementById('btn-actualizar-erp');
        if (!silencioso && btn) { btn.disabled = true; btn.textContent = 'Actualizando...'; }

        let actualizados = 0;
        let errores = 0;
        const ahora = new Date().toISOString();

        for (const caso of casosActivos) {
            const ref = caso.orden?.referencia;
            if (!ref || ref === 'undefined' || ref === 'SIN-REF') continue;
            try {
                const url = `${estado.backendUrl}/ventas/venta-id/${encodeURIComponent(ref.toString().trim())}`;
                const response = await apiFetch(url, { headers: { 'Accept': 'application/json' } });

                if (!response.ok) {
                    if (response.status === 404) {
                        caso.erp = { ...(caso.erp || {}), embarqueInfo: {} };
                        continue;
                    }
                    throw new Error(`HTTP ${response.status}`);
                }

                const apiData = await response.json();
                if (!apiData || !Array.isArray(apiData) || apiData.length === 0) {
                    caso.erp = { ...(caso.erp || {}), embarqueInfo: {} };
                    continue;
                }

                const primer = apiData[0];
                const estatuses = [...new Set(apiData.map(m => (m.Estatus || '').trim()))].join(', ');
                const movIDs    = apiData.map(m => m.MovID).join(', ');
                // Usar el primer record que tenga embarque; si ninguno, usar apiData[0]
                const conEmbarque = apiData.find(m => m.Embarque) || primer;
                const embarqueInfo = {
                    embarque:      conEmbarque.Embarque             || null,
                    observaciones: conEmbarque.EmbarqueObservaciones || null,
                    agente:        conEmbarque.EmbarqueAgente        || null,
                    estatus:       conEmbarque.EmbarqueEstatus       || null
                };

                caso.erp = {
                    encontrado: true,
                    estatusERP: estatuses || caso.erp?.estatusERP || '',
                    movimientos: movIDs   || caso.erp?.movimientos || 'N/A',
                    importeERP:  primer.PrecioTotal ?? (caso.erp?.importeERP || 0),
                    diferencia:  caso.erp?.diferencia ?? null,
                    cliente:     primer.Cliente?.trim() || caso.erp?.cliente || '',
                    movimientosDetalle: apiData.map(m => ({
                        mov: m.Mov?.trim() || '', movID: m.MovID || '',
                        estatus: m.Estatus?.trim() || '', importe: m.Importe || 0,
                        impuestos: m.Impuestos || 0, total: m.PrecioTotal || 0,
                        cliente: m.Cliente?.trim() || '', referencia: m.Referencia || '',
                        atencion: m.Atencion || '', fechaEmision: m.FechaEmision || ''
                    })),
                    embarqueInfo
                };
                caso.fechaUltimaActualizacion = ahora;
                actualizados++;
            } catch (e) {
                console.warn(`Error actualizando ERP para ${ref}:`, e.message);
                errores++;
            }
            // Delay entre peticiones para evitar rate limiting
            await new Promise(r => setTimeout(r, 80));
        }

        if (actualizados > 0) {
            Persistencia.guardarCasos();
            casosActivos.forEach(c => this._syncCasoEnBackground(c));
        }

        if (!silencioso) {
            if (btn) {
                btn.disabled = false;
                btn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width: 14px; height: 14px;"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg> Actualizar ERP`;
            }
            const msg = errores > 0
                ? `ERP actualizado: ${actualizados} casos (${errores} errores).`
                : `ERP actualizado correctamente en ${actualizados} caso(s).`;
            Interfaz.mostrarToast(msg, actualizados > 0 ? 'success' : 'warning');
        }

        // Re-renderizar para reflejar cambios
        this.renderizarPortal();
        if (estado.casoSeleccionado) this._renderizarDetalleEnPanel(estado.casoSeleccionado);
    },

    exportarPortalExcel() {
        // Aplicar los mismos filtros que renderizarPortal para exportar solo lo visible
        const activos = estado.casosPendientes.filter(c => c.estado !== 'RESUELTO');
        let casosFiltrados = [...activos];

        if (estado.filtroActual.prioridad) {
            casosFiltrados = casosFiltrados.filter(c => c.prioridad === estado.filtroActual.prioridad);
        }
        if (estado.filtroActual.responsable !== 'todos') {
            casosFiltrados = casosFiltrados.filter(c => c.responsable === estado.filtroActual.responsable);
        }
        if (estado.filtroActual.workflow) {
            casosFiltrados = casosFiltrados.filter(c => c.workflow === estado.filtroActual.workflow);
        }
        const busqueda = (document.getElementById('portalBuscarCasos')?.value || '').trim().toLowerCase();
        if (busqueda) {
            casosFiltrados = casosFiltrados.filter(c =>
                (c.orden?.referencia || '').toLowerCase().includes(busqueda) ||
                (c.tipo || '').toLowerCase().includes(busqueda)
            );
        }
        const fechaFiltro = document.getElementById('portalFiltroCasosFecha')?.value || '';
        if (fechaFiltro) {
            casosFiltrados = casosFiltrados.filter(c => c.fechaCreacion?.startsWith(fechaFiltro));
        }

        if (casosFiltrados.length === 0) {
            Interfaz.mostrarToast('No hay casos visibles para exportar', 'warning');
            return;
        }

        // Ordenar igual que el portal
        casosFiltrados.sort((a, b) => {
            const ord = { URGENTE: 0, IMPORTANTE: 1, BAJA: 2 };
            if (ord[a.prioridad] !== ord[b.prioridad]) return ord[a.prioridad] - ord[b.prioridad];
            return new Date(b.fechaCreacion) - new Date(a.fechaCreacion);
        });

        const rows = casosFiltrados.map(c => {
            const orden = (estado.ordenesConsolidadas || []).find(o => o.referenciaERP === c.orden?.referencia);
            const pasosTotal = c.accionesSugeridas?.length || 0;
            const pasosHechos = c.accionesSugeridas?.filter(a => a.completado).length || 0;
            const ultimaAccion = c.historial?.length
                ? c.historial[c.historial.length - 1].accion
                : '';

            return {
                'Caso':             c.numeroCaso || '#TEMP',
                'Referencia':       c.orden?.referencia || '',
                'Tipo':             c.tipo || '',
                'Workflow':         c.workflow || '',
                'Prioridad':        c.prioridad || '',
                'Responsable':      c.responsable || '',
                'Estado':           c.estado || '',
                'Monto':            c.orden?.monto || 0,
                'Ingreso Bruto':    +(orden?.ingresos?.bruto || 0),
                'Comisión MP':      -Math.abs(orden?.costos?.comision     || 0),
                'Costo Envío':      -Math.abs(orden?.costos?.envio        || 0),
                'Financiamiento':   -Math.abs(orden?.costos?.financiamiento || 0),
                'Total Gastos':     -Math.abs(orden?.costos?.total        || 0),
                'ERP Encontrado':   c.erp?.encontrado ? 'Sí' : 'No',
                'ERP Estatus':      c.erp?.estatusERP || '',
                'ERP Diferencia':   c.erp?.diferencia ?? '',
                'Embarque':         c.erp?.embarqueInfo?.embarque || '',
                'Estatus Embarque': c.erp?.embarqueInfo?.estatus || '',
                'Agente Embarque':  c.erp?.embarqueInfo?.agente || '',
                'Obs. Embarque':    c.erp?.embarqueInfo?.observaciones || '',
                'Pasos':            `${pasosHechos}/${pasosTotal}`,
                'SLA':              c.sla || '',
                'Fecha Creación':   c.fechaCreacion ? new Date(c.fechaCreacion).toLocaleDateString('es-MX') : '',
                'Última Actualización': c.fechaUltimaActualizacion ? new Date(c.fechaUltimaActualizacion).toLocaleDateString('es-MX') : '',
                'Última Acción':    ultimaAccion,
                'Notas':            c.notas || ''
            };
        });

        const ws = XLSX.utils.json_to_sheet(rows);

        // Anchos de columna
        ws['!cols'] = [
            { wch: 8  }, // Caso
            { wch: 20 }, // Referencia
            { wch: 22 }, // Tipo
            { wch: 22 }, // Workflow
            { wch: 12 }, // Prioridad
            { wch: 13 }, // Responsable
            { wch: 12 }, // Estado
            { wch: 12 }, // Monto
            { wch: 14 }, // Ingreso Bruto
            { wch: 13 }, // Comisión MP
            { wch: 13 }, // Costo Envío
            { wch: 14 }, // Financiamiento
            { wch: 13 }, // Total Gastos
            { wch: 14 }, // ERP Encontrado
            { wch: 14 }, // ERP Estatus
            { wch: 14 }, // ERP Diferencia
            { wch: 8  }, // Pasos
            { wch: 10 }, // SLA
            { wch: 16 }, // Fecha Creación
            { wch: 20 }, // Última Actualización
            { wch: 45 }, // Última Acción
            { wch: 40 }  // Notas
        ];

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Portal de Casos');

        // Nombre de archivo con filtros activos
        const partes = ['portal_casos'];
        if (estado.filtroActual.prioridad) partes.push(estado.filtroActual.prioridad.toLowerCase());
        if (estado.filtroActual.responsable !== 'todos') partes.push(estado.filtroActual.responsable.toLowerCase());
        if (estado.filtroActual.workflow) partes.push(estado.filtroActual.workflow.toLowerCase());
        partes.push(new Date().toISOString().split('T')[0]);

        XLSX.writeFile(wb, partes.join('_') + '.xlsx');
        Interfaz.mostrarToast(`${casosFiltrados.length} casos exportados a Excel`, 'success');
    },

    // ── Alias: mantiene compatibilidad con todos los call sites existentes ─────
    renderizarSidebar() {
        this.renderizarPortal();
    },

    // ── Renderizado principal del portal completo ──────────────────────────────
    renderizarPortal() {
        // Calcular estadísticas
        const activos = estado.casosPendientes.filter(c => c.estado !== 'RESUELTO');
        const pendientes = activos.filter(c => c.estado === 'PENDIENTE').length;
        const enRevision = activos.filter(c => c.estado === 'EN_REVISION').length;
        const resueltos = estado.casosPendientes.filter(c => c.estado === 'RESUELTO').length;

        // KPI bar del portal
        const elTotal      = document.getElementById('portal-count-total');
        const elPendientes = document.getElementById('portal-count-pendientes');
        const elEnRevision = document.getElementById('portal-count-en-revision');
        const elResueltos  = document.getElementById('portal-count-resueltos');
        if (elTotal)      elTotal.textContent      = activos.length;
        if (elPendientes) elPendientes.textContent = pendientes;
        if (elEnRevision) elEnRevision.textContent = enRevision;
        if (elResueltos)  elResueltos.textContent  = resueltos;

        // Marcar KPI activo
        document.querySelectorAll('.portal-kpi[data-filter-estado]').forEach(el => el.classList.remove('active'));
        if (estado.filtroActual.estado) {
            document.querySelector(`.portal-kpi[data-filter-estado="${estado.filtroActual.estado}"]`)?.classList.add('active');
        }

        // Badge en el botón de nav (muestra pendientes sin atender)
        const badge = document.getElementById('nav-casos-badge');
        if (badge) {
            badge.textContent = pendientes;
            badge.style.display = pendientes > 0 ? 'inline' : 'none';
        }

        // Filtrar casos
        let casosFiltrados = [...activos];

        if (estado.filtroActual.estado) {
            casosFiltrados = casosFiltrados.filter(c => c.estado === estado.filtroActual.estado);
        }

        if (estado.filtroActual.responsable !== 'todos') {
            casosFiltrados = casosFiltrados.filter(c => c.responsable === estado.filtroActual.responsable);
        }

        if (estado.filtroActual.workflow) {
            casosFiltrados = casosFiltrados.filter(c => c.workflow === estado.filtroActual.workflow);
        }

        // Sincronizar el select con el estado actual
        const wfSelect = document.getElementById('portalFiltroWorkflow');
        if (wfSelect) wfSelect.value = estado.filtroActual.workflow || '';

        const busqueda = (document.getElementById('portalBuscarCasos')?.value || '').trim().toLowerCase();
        if (busqueda) {
            casosFiltrados = casosFiltrados.filter(c =>
                (c.orden?.referencia || '').toLowerCase().includes(busqueda) ||
                (c.tipo || '').toLowerCase().includes(busqueda)
            );
        }

        const fechaFiltro = document.getElementById('portalFiltroCasosFecha')?.value || '';
        if (fechaFiltro) {
            casosFiltrados = casosFiltrados.filter(c => {
                if (!c.fechaCreacion) return false;
                return c.fechaCreacion.startsWith(fechaFiltro);
            });
        }

        // Ordenar por prioridad y fecha
        casosFiltrados.sort((a, b) => {
            const prioridadOrder = { URGENTE: 0, IMPORTANTE: 1, BAJA: 2 };
            if (prioridadOrder[a.prioridad] !== prioridadOrder[b.prioridad]) {
                return prioridadOrder[a.prioridad] - prioridadOrder[b.prioridad];
            }
            return new Date(b.fechaCreacion) - new Date(a.fechaCreacion);
        });

        // Actualizar count en header de la lista
        const countEl = document.getElementById('portalListaCount');
        if (countEl) countEl.textContent = casosFiltrados.length;

        // Renderizar lista de tarjetas
        const container = document.getElementById('portalListaCasos');
        if (!container) {
            DashboardManager.actualizarDashboard();
            return;
        }

        if (casosFiltrados.length === 0) {
            container.innerHTML = `
                <div class="sidebar-empty">
                    <div class="sidebar-empty-icon">📭</div>
                    <p>No hay casos pendientes</p>
                    <p class="sidebar-empty-sub">Los casos se crearán automáticamente al procesar una conciliación</p>
                </div>
            `;
        } else {
            container.innerHTML = casosFiltrados.map(caso => {
                if (!caso.tipo || !caso.prioridad || !caso.responsable) {
                    console.warn('Caso con campos faltantes:', caso);
                    return '';
                }
                const tipoClase = (caso.tipo || 'REVISAR').toLowerCase().replace('_', '-');
                const prioridadClase = (caso.prioridad || 'BAJA').toLowerCase();
                const responsableInicial = (caso.responsable || 'X').charAt(0);

                return `
                    <div class="caso-mini ${prioridadClase} ${estado.casoSeleccionado?.id === caso.id ? 'selected' : ''} ${caso.estado === 'EN_REVISION' ? 'en-revision' : ''}"
                         onclick="CasosManager.seleccionarCasoEnPortal('${caso.id}')">
                        <div class="caso-mini-header">
                            <span class="caso-mini-tipo ${tipoClase}">${caso.tipo}</span>
                            <span class="caso-mini-monto">${Utilidades.formatMoney(Math.abs(caso.orden?.monto || 0))}</span>
                        </div>
                        <div class="caso-mini-ref">${caso.orden?.referencia || 'Sin referencia'}</div>
                        <div class="caso-mini-footer">
                            <div class="caso-mini-responsable">
                                <span class="responsable-badge ${caso.responsable.toLowerCase()}">${responsableInicial}</span>
                                <span>${caso.responsable}</span>
                            </div>
                            <div style="display:flex;gap:4px;align-items:center;">
                                ${caso.estado === 'EN_REVISION' ? '<span class="badge-en-revision">En revisión</span>' : ''}
                                <span class="caso-num-badge ${caso.numeroCaso === '#TEMP' ? 'temp' : ''}">${caso.numeroCaso || '#TEMP'}</span>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }

        DashboardManager.actualizarDashboard();
    },

    // ── Seleccionar caso en el panel del portal ────────────────────────────────
    guardarNotasCaso(contexto) {
        if (!estado.casoSeleccionado) return;

        const notasId = contexto === 'portal' ? 'portal-detalle-notas' : 'detalle-notas';
        const btnId   = contexto === 'portal' ? 'btn-guardar-notas-portal' : 'btn-guardar-notas-modal';
        const notas   = document.getElementById(notasId)?.value || '';

        estado.casoSeleccionado.notas = notas;
        estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
        Persistencia.guardarCasos();

        const btn = document.getElementById(btnId);
        const textoOriginal = btn?.textContent;
        if (btn) { btn.textContent = 'Guardado ✓'; btn.disabled = true; }

        ApiCasos.actualizarCasoIndividual(estado.casoSeleccionado)
            .then(() => {
                if (btn) btn.textContent = 'Guardado ✓';
            })
            .catch(() => {
                if (btn) btn.textContent = 'Guardado (local)';
            })
            .finally(() => {
                setTimeout(() => {
                    if (btn) { btn.textContent = textoOriginal; btn.disabled = false; }
                }, 2000);
            });
    },

    // Sincroniza un caso con el backend en segundo plano (silencioso)
    _syncCasoEnBackground(caso) {
        ApiCasos.actualizarCasoIndividual(caso).catch(err => {
            console.warn('No se pudo sincronizar caso con backend:', err.message);
        });
    },

    toggleEnRevision() {
        if (!estado.casoSeleccionado) return;

        const caso = estado.casoSeleccionado;
        const enRevision = caso.estado !== 'EN_REVISION';
        caso.estado = enRevision ? 'EN_REVISION' : 'PENDIENTE';
        caso.fechaUltimaActualizacion = new Date().toISOString();
        caso.historial.push({
            fecha: new Date().toISOString(),
            accion: enRevision ? 'Marcado en revisión' : 'Revisión cancelada — vuelve a Pendiente',
            usuario: 'Usuario'
        });

        Persistencia.guardarCasos();
        this._syncCasoEnBackground(caso);

        const modalOpen = document.getElementById('casoDetalleOverlay').classList.contains('show');
        if (modalOpen) {
            this.abrirDetalleCaso(caso.id);
        } else {
            this._renderizarDetalleEnPanel(caso);
            this.renderizarPortal();
        }
    },

    seleccionarCasoEnPortal(casoId) {
        // Guardar notas del caso previo antes de cambiar
        if (estado.casoSeleccionado) {
            const notasEl = document.getElementById('portal-detalle-notas');
            if (notasEl) {
                const notas = notasEl.value;
                if (notas !== estado.casoSeleccionado.notas) {
                    estado.casoSeleccionado.notas = notas;
                    estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
                }
            }
            Persistencia.guardarCasos();
            this._syncCasoEnBackground(estado.casoSeleccionado);
        }

        const caso = estado.casosPendientes.find(c => c.id === casoId);
        if (!caso) return;

        estado.casoSeleccionado = caso;
        this._renderizarDetalleEnPanel(caso);
        this.renderizarPortal(); // marca tarjeta como selected
    },

    // ── Deseleccionar caso y volver al placeholder ─────────────────────────────
    deseleccionarCasoEnPortal() {
        if (estado.casoSeleccionado) {
            const notasEl = document.getElementById('portal-detalle-notas');
            if (notasEl) {
                const notas = notasEl.value;
                if (notas !== estado.casoSeleccionado.notas) {
                    estado.casoSeleccionado.notas = notas;
                    estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
                }
            }
            Persistencia.guardarCasos();
            this._syncCasoEnBackground(estado.casoSeleccionado);
        }
        estado.casoSeleccionado = null;
        this._mostrarPlaceholderPanel();
        this.renderizarPortal();
    },

    _mostrarPlaceholderPanel() {
        const container = document.getElementById('portalDetalleCaso');
        if (!container) return;
        container.innerHTML = `
            <div class="portal-detalle-placeholder">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1.5" style="width: 64px; height: 64px; margin-bottom: 16px; opacity: 0.3;">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                </svg>
                <p>Selecciona un caso de la lista para ver su detalle</p>
            </div>
        `;
    },

    // ── Renderizar detalle del caso en el panel (sin modal) ────────────────────
    _renderizarDetalleEnPanel(caso) {
        const container = document.getElementById('portalDetalleCaso');
        if (!container) return;

        const infoHTML = `
            <div class="caso-info-item">
                <label>Pedido</label>
                <span>${caso.orden.pedido || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Factura</label>
                <span>${caso.orden.factura || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Cliente</label>
                <span>${caso.orden.cliente || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Monto</label>
                <span style="color: ${caso.orden.monto >= 0 ? 'var(--success)' : 'var(--danger)'}">${Utilidades.formatMoney(caso.orden.monto)}</span>
            </div>
            <div class="caso-info-item">
                <label>Fecha Creación</label>
                <span>${Utilidades.formatFecha(caso.fechaCreacion)}</span>
            </div>
            ${caso.orden.tieneMultiplesSourceIds ? `
                <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 8px; padding-top: 12px; border-top: 1px solid var(--gray-200);">
                    <label style="font-weight: 600; color: #3730a3;">
                        <span style="background: #e0e7ff; color: #3730a3; padding: 2px 6px; border-radius: 3px; font-size: 0.75rem; margin-right: 6px;">MULTI</span>
                        Transacciones Múltiples (${caso.orden.sourceIds?.length || 0} SOURCE_IDs)
                    </label>
                </div>
                <div class="caso-info-item" style="grid-column: 1 / -1;">
                    <label>SOURCE_IDs</label>
                    <span style="font-family: monospace; font-size: 0.75rem; word-break: break-all;">
                        ${(caso.orden.sourceIds || []).join(', ')}
                    </span>
                </div>
                ${caso.orden.validacionDetalle?.estadosCuentaDetalle?.length > 0 ? `
                    <div class="caso-info-item" style="grid-column: 1 / -1;">
                        <label>Detalle por Transacción</label>
                        <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-top: 4px;">
                            ${caso.orden.validacionDetalle.estadosCuentaDetalle.map(det => `
                                <span style="background: var(--gray-100); padding: 4px 8px; border-radius: 4px; font-size: 0.75rem; font-family: monospace;">
                                    ${det.sourceId}: ${Utilidades.formatMoney(det.neto)}
                                </span>
                            `).join('')}
                        </div>
                    </div>
                    <div class="caso-info-item">
                        <label>Diferencia Consolidada</label>
                        <span style="color: ${caso.orden.validacionDetalle?.saldoValidado ? 'var(--success)' : 'var(--warning)'}; font-weight: 600;">
                            ${caso.orden.validacionDetalle?.saldoValidado ? 'Validado' : `Dif: ${Utilidades.formatMoney(caso.orden.validacionDetalle?.diferencia || 0)}`}
                        </span>
                    </div>
                ` : ''}
            ` : ''}
            ${caso.erp ? `
                <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 8px; padding-top: 12px; border-top: 1px solid var(--gray-200);">
                    <label style="font-weight: 600; color: var(--primary);">🔗 Estado en Intelisis ERP</label>
                </div>
                <div class="caso-info-item">
                    <label>Estado</label>
                    <span style="color: ${caso.erp.encontrado ? 'var(--success)' : 'var(--danger)'};">
                        ${caso.erp.encontrado ? 'Encontrado' : 'No Encontrado'}
                    </span>
                </div>
                ${caso.erp.encontrado ? `
                    <div class="caso-info-item">
                        <label>Estatus ERP</label>
                        <span style="font-family: monospace; font-size: 0.8125rem; background: var(--gray-100); padding: 2px 6px; border-radius: 3px;">${caso.erp.estatusERP}</span>
                    </div>
                    <div class="caso-info-item">
                        <label>Movimientos</label>
                        <span style="font-family: monospace; font-size: 0.8125rem;">${caso.erp.movimientos}</span>
                    </div>
                    <div class="caso-info-item">
                        <label>Importe ERP</label>
                        <span>${Utilidades.formatMoney(caso.erp.importeERP)}</span>
                    </div>
                    ${caso.erp.diferencia !== null ? `
                        <div class="caso-info-item">
                            <label>Diferencia</label>
                            <span style="color: ${Math.abs(caso.erp.diferencia) <= 0.01 ? 'var(--success)' : 'var(--warning)'}; font-weight: 600;">
                                ${Math.abs(caso.erp.diferencia) <= 0.01 ? 'Coincide' : `Dif: ${Utilidades.formatMoney(caso.erp.diferencia)}`}
                            </span>
                        </div>
                    ` : ''}
                    ${'embarqueInfo' in (caso.erp || {}) ? `
                        <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 6px; padding-top: 10px; border-top: 1px solid var(--gray-100);">
                            <label style="font-weight: 600; color: var(--gray-700);">📦 Embarque Intelisis</label>
                        </div>
                        ${caso.erp.embarqueInfo.embarque ? `<div class="caso-info-item"><label>Embarque</label><span style="font-family: monospace; font-size: 0.8125rem;">${caso.erp.embarqueInfo.embarque}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.estatus ? `<div class="caso-info-item"><label>Estatus Embarque</label><span style="background: var(--gray-100); padding: 2px 6px; border-radius: 3px; font-size: 0.8125rem;">${caso.erp.embarqueInfo.estatus}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.agente ? `<div class="caso-info-item"><label>Agente</label><span>${caso.erp.embarqueInfo.agente}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.observaciones ? `<div class="caso-info-item" style="grid-column: 1 / -1;"><label>Obs. Embarque</label><span style="font-size: 0.8125rem; color: var(--gray-600);">${caso.erp.embarqueInfo.observaciones}</span></div>` : ''}
                        ${!caso.erp.embarqueInfo.embarque && !caso.erp.embarqueInfo.estatus && !caso.erp.embarqueInfo.agente && !caso.erp.embarqueInfo.observaciones ? `
                            <div class="caso-info-item" style="grid-column: 1 / -1;"><span style="font-size: 0.8125rem; color: var(--gray-400); font-style: italic;">Sin información de embarque registrada</span></div>
                        ` : ''}
                    ` : ''}
                ` : ''}
            ` : ''}
        `;

        container.innerHTML = `
            <div class="portal-detalle-contenido">
                <div class="portal-detalle-header">
                    <div style="display: flex; align-items: center; gap: 12px; flex: 1; min-width: 0;">
                        <span class="prioridad-badge ${caso.prioridad.toLowerCase()}">${caso.prioridad}</span>
                        <div style="min-width: 0;">
                            <h2 style="font-size: 1rem; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                                <span class="caso-num-badge ${caso.numeroCaso === '#TEMP' ? 'temp' : ''}" style="margin-right: 6px;">${caso.numeroCaso || '#TEMP'}</span><a href="https://www.mercadolibre.com.mx/ventas/omni/listado?filters=&subFilters=&search=${encodeURIComponent(caso.orden.referencia)}&limit=50&offset=0" target="_blank" rel="noopener" title="Ver en MercadoLibre" style="color: inherit; text-decoration: none; border-bottom: 1px dashed var(--primary);">${caso.orden.referencia}&nbsp;<svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width:11px;height:11px;display:inline;vertical-align:middle;color:var(--primary);"><path stroke-linecap="round" stroke-linejoin="round" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" /></svg></a>
                            </h2>
                            <p style="font-size: 0.8125rem; color: var(--gray-500);">${caso.tipo} — Asignado a: ${caso.responsable}</p>
                        </div>
                    </div>
                    <button class="btn btn-sm btn-secondary" onclick="CasosManager.deseleccionarCasoEnPortal()" style="flex-shrink: 0;">&times; Cerrar</button>
                </div>
                <div class="portal-detalle-body">
                    <div class="caso-info-grid" id="portal-detalle-info">
                        ${infoHTML}
                    </div>

                    <div class="caso-section">
                        <h4>Historial</h4>
                        <div class="historial-list" id="portal-detalle-historial">
                            ${caso.historial.slice().reverse().map(h => `
                                <div class="historial-item">
                                    <span class="historial-fecha">${Utilidades.formatFechaHora(h.fecha)}</span>
                                    <span class="historial-detalle">${h.accion} <em>(${h.usuario})</em></span>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                    <div class="caso-section caso-notas">
                        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;">
                            <h4 style="margin: 0;">Notas</h4>
                            <button class="btn btn-sm btn-primary" id="btn-guardar-notas-portal" onclick="CasosManager.guardarNotasCaso('portal')">Guardar notas</button>
                        </div>
                        <textarea id="portal-detalle-notas" placeholder="Agregar notas sobre este caso...">${caso.notas || ''}</textarea>
                    </div>
                </div>
                <div class="portal-detalle-footer">
                    <div class="caso-sla">
                        <span>SLA:</span>
                        <strong>${caso.sla || '24 horas'}</strong>
                    </div>
                    <div class="caso-actions">
                        <select class="reasignar-select ${caso.responsable.toLowerCase()}" onchange="CasosManager.reasignarCaso(this.value)">
                            <option value="CONTADOR" ${caso.responsable === 'CONTADOR' ? 'selected' : ''}>Contador</option>
                            <option value="FINANZAS" ${caso.responsable === 'FINANZAS' ? 'selected' : ''}>Finanzas</option>
                            <option value="LOGISTICA" ${caso.responsable === 'LOGISTICA' ? 'selected' : ''}>Logística</option>
                        </select>
                        <button class="btn ${caso.estado === 'EN_REVISION' ? 'btn-warning' : 'btn-secondary'}" onclick="CasosManager.toggleEnRevision()">${caso.estado === 'EN_REVISION' ? '🔵 En revisión' : 'Marcar en revisión'}</button>
                        <button class="btn btn-success" onclick="CasosManager.resolverCaso()">Marcar Resuelto</button>
                    </div>
                </div>
            </div>
        `;
    },

    abrirDetalleCaso(casoId) {
        const caso = estado.casosPendientes.find(c => c.id === casoId);
        if (!caso) return;

        estado.casoSeleccionado = caso;
        const workflow = CONFIG.WORKFLOWS[caso.workflow] || CONFIG.WORKFLOWS.REVISAR;

        // Actualizar badge de prioridad
        const prioridadEl = document.getElementById('detalle-prioridad');
        prioridadEl.textContent = caso.prioridad;
        prioridadEl.className = 'prioridad-badge ' + caso.prioridad.toLowerCase();

        // Títulos
        const numEl = document.createElement('span');
        numEl.className = 'caso-num-badge' + (caso.numeroCaso === '#TEMP' ? ' temp' : '');
        numEl.textContent = caso.numeroCaso || '#TEMP';
        const tituloEl = document.getElementById('detalle-titulo');
        tituloEl.innerHTML = '';
        tituloEl.appendChild(numEl);
        const mlUrl = `https://www.mercadolibre.com.mx/ventas/omni/listado?filters=&subFilters=&search=${encodeURIComponent(caso.orden.referencia)}&limit=50&offset=0`;
        const refLink = document.createElement('a');
        refLink.href = mlUrl;
        refLink.target = '_blank';
        refLink.rel = 'noopener';
        refLink.title = 'Ver en MercadoLibre';
        refLink.style.cssText = 'color: inherit; text-decoration: none; border-bottom: 1px dashed var(--primary); margin-left: 4px;';
        refLink.innerHTML = caso.orden.referencia + ' <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width:11px;height:11px;display:inline;vertical-align:middle;color:var(--primary);"><path stroke-linecap="round" stroke-linejoin="round" d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" /></svg>';
        tituloEl.appendChild(refLink);
        document.getElementById('detalle-subtitulo').textContent = `${caso.tipo} - Asignado a: ${caso.responsable}`;
        const selectEl = document.getElementById('detalle-responsable-select');
        selectEl.value = caso.responsable;
        selectEl.className = 'reasignar-select ' + caso.responsable.toLowerCase();

        const btnRevision = document.getElementById('btn-en-revision-modal');
        if (btnRevision) {
            const enRevision = caso.estado === 'EN_REVISION';
            btnRevision.textContent = enRevision ? '🔵 En revisión' : 'Marcar en revisión';
            btnRevision.className = 'btn ' + (enRevision ? 'btn-warning' : 'btn-secondary');
        }

        // Info grid
        const infoHTML = `
            <div class="caso-info-item">
                <label>Pedido</label>
                <span>${caso.orden.pedido || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Factura</label>
                <span>${caso.orden.factura || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Cliente</label>
                <span>${caso.orden.cliente || '--'}</span>
            </div>
            <div class="caso-info-item">
                <label>Monto</label>
                <span style="color: ${caso.orden.monto >= 0 ? 'var(--success)' : 'var(--danger)'}">${Utilidades.formatMoney(caso.orden.monto)}</span>
            </div>
            <div class="caso-info-item">
                <label>Fecha Creación</label>
                <span>${Utilidades.formatFecha(caso.fechaCreacion)}</span>
            </div>
            ${caso.orden.tieneMultiplesSourceIds ? `
                <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 8px; padding-top: 12px; border-top: 1px solid var(--gray-200);">
                    <label style="font-weight: 600; color: #3730a3;">
                        <span style="background: #e0e7ff; color: #3730a3; padding: 2px 6px; border-radius: 3px; font-size: 0.75rem; margin-right: 6px;">MULTI</span>
                        Transacciones Múltiples (${caso.orden.sourceIds?.length || 0} SOURCE_IDs)
                    </label>
                </div>
                <div class="caso-info-item" style="grid-column: 1 / -1;">
                    <label>SOURCE_IDs</label>
                    <span style="font-family: monospace; font-size: 0.75rem; word-break: break-all;">
                        ${(caso.orden.sourceIds || []).join(', ')}
                    </span>
                </div>
                ${caso.orden.validacionDetalle?.estadosCuentaDetalle?.length > 0 ? `
                    <div class="caso-info-item" style="grid-column: 1 / -1;">
                        <label>Detalle por Transacción</label>
                        <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-top: 4px;">
                            ${caso.orden.validacionDetalle.estadosCuentaDetalle.map(det => `
                                <span style="background: var(--gray-100); padding: 4px 8px; border-radius: 4px; font-size: 0.75rem; font-family: monospace;">
                                    ${det.sourceId}: ${Utilidades.formatMoney(det.neto)}
                                </span>
                            `).join('')}
                        </div>
                    </div>
                    <div class="caso-info-item">
                        <label>Diferencia Consolidada</label>
                        <span style="color: ${caso.orden.validacionDetalle?.saldoValidado ? 'var(--success)' : 'var(--warning)'}; font-weight: 600;">
                            ${caso.orden.validacionDetalle?.saldoValidado ? 'Validado' : `Dif: ${Utilidades.formatMoney(caso.orden.validacionDetalle?.diferencia || 0)}`}
                        </span>
                    </div>
                ` : ''}
            ` : ''}
            ${caso.erp ? `
                <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 8px; padding-top: 12px; border-top: 1px solid var(--gray-200);">
                    <label style="font-weight: 600; color: var(--primary);">🔗 Estado en Intelisis ERP</label>
                </div>
                <div class="caso-info-item">
                    <label>Estado</label>
                    <span style="color: ${caso.erp.encontrado ? 'var(--success)' : 'var(--danger)'};">
                        ${caso.erp.encontrado ? 'Encontrado' : 'No Encontrado'}
                    </span>
                </div>
                ${caso.erp.encontrado ? `
                    <div class="caso-info-item">
                        <label>Estatus ERP</label>
                        <span style="font-family: monospace; font-size: 0.8125rem; background: var(--gray-100); padding: 2px 6px; border-radius: 3px;">${caso.erp.estatusERP}</span>
                    </div>
                    <div class="caso-info-item">
                        <label>Movimientos</label>
                        <span style="font-family: monospace; font-size: 0.8125rem;">${caso.erp.movimientos}</span>
                    </div>
                    <div class="caso-info-item">
                        <label>Importe ERP</label>
                        <span>${Utilidades.formatMoney(caso.erp.importeERP)}</span>
                    </div>
                    ${caso.erp.diferencia !== null ? `
                        <div class="caso-info-item">
                            <label>Diferencia</label>
                            <span style="color: ${Math.abs(caso.erp.diferencia) <= 0.01 ? 'var(--success)' : 'var(--warning)'}; font-weight: 600;">
                                ${Math.abs(caso.erp.diferencia) <= 0.01 ? 'Coincide' : `Dif: ${Utilidades.formatMoney(caso.erp.diferencia)}`}
                            </span>
                        </div>
                    ` : ''}
                    ${'embarqueInfo' in (caso.erp || {}) ? `
                        <div class="caso-info-item" style="grid-column: 1 / -1; margin-top: 6px; padding-top: 10px; border-top: 1px solid var(--gray-100);">
                            <label style="font-weight: 600; color: var(--gray-700);">📦 Embarque Intelisis</label>
                        </div>
                        ${caso.erp.embarqueInfo.embarque ? `<div class="caso-info-item"><label>Embarque</label><span style="font-family: monospace; font-size: 0.8125rem;">${caso.erp.embarqueInfo.embarque}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.estatus ? `<div class="caso-info-item"><label>Estatus Embarque</label><span style="background: var(--gray-100); padding: 2px 6px; border-radius: 3px; font-size: 0.8125rem;">${caso.erp.embarqueInfo.estatus}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.agente ? `<div class="caso-info-item"><label>Agente</label><span>${caso.erp.embarqueInfo.agente}</span></div>` : ''}
                        ${caso.erp.embarqueInfo.observaciones ? `<div class="caso-info-item" style="grid-column: 1 / -1;"><label>Obs. Embarque</label><span style="font-size: 0.8125rem; color: var(--gray-600);">${caso.erp.embarqueInfo.observaciones}</span></div>` : ''}
                        ${!caso.erp.embarqueInfo.embarque && !caso.erp.embarqueInfo.estatus && !caso.erp.embarqueInfo.agente && !caso.erp.embarqueInfo.observaciones ? `
                            <div class="caso-info-item" style="grid-column: 1 / -1;"><span style="font-size: 0.8125rem; color: var(--gray-400); font-style: italic;">Sin información de embarque registrada</span></div>
                        ` : ''}
                    ` : ''}
                ` : ''}
            ` : ''}
        `;

        document.getElementById('detalle-info').innerHTML = infoHTML;


        // Historial
        document.getElementById('detalle-historial').innerHTML = caso.historial
            .slice()
            .reverse()
            .map(h => `
                <div class="historial-item">
                    <span class="historial-fecha">${Utilidades.formatFechaHora(h.fecha)}</span>
                    <span class="historial-detalle">${h.accion} <em>(${h.usuario})</em></span>
                </div>
            `).join('');

        // SLA y notas
        document.getElementById('detalle-sla').textContent = caso.sla || '24 horas';
        document.getElementById('detalle-notas').value = caso.notas || '';

        // Mostrar modal
        document.getElementById('casoDetalleOverlay').classList.add('show');
        this.renderizarSidebar();
    },

    cerrarDetalleCaso(event) {
        if (event && event.target !== event.currentTarget) return;

        // Guardar notas antes de cerrar
        if (estado.casoSeleccionado) {
            const notas = document.getElementById('detalle-notas').value;
            if (notas !== estado.casoSeleccionado.notas) {
                estado.casoSeleccionado.notas = notas;
                estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
            }
            Persistencia.guardarCasos();
            this._syncCasoEnBackground(estado.casoSeleccionado);
        }

        document.getElementById('casoDetalleOverlay').classList.remove('show');
        estado.casoSeleccionado = null;
        this.renderizarSidebar();
    },

    toggleAccion(casoId, idx) {
        const caso = estado.casosPendientes.find(c => c.id === casoId);
        if (!caso) return;

        caso.accionesSugeridas[idx].completado = !caso.accionesSugeridas[idx].completado;
        caso.fechaUltimaActualizacion = new Date().toISOString();
        caso.historial.push({
            fecha: new Date().toISOString(),
            accion: caso.accionesSugeridas[idx].completado
                ? `Paso completado: "${caso.accionesSugeridas[idx].texto.substring(0, 50)}..."`
                : `Paso desmarcado: "${caso.accionesSugeridas[idx].texto.substring(0, 50)}..."`,
            usuario: 'Usuario'
        });

        Persistencia.guardarCasos();
        this._syncCasoEnBackground(caso);

        const modalOpen = document.getElementById('casoDetalleOverlay').classList.contains('show');
        if (modalOpen) {
            this.abrirDetalleCaso(casoId);
        } else if (estado.casoSeleccionado) {
            this._renderizarDetalleEnPanel(estado.casoSeleccionado);
        }
    },

    async resolverCaso() {
        if (!estado.casoSeleccionado) return;

        estado.casoSeleccionado.estado = 'RESUELTO';
        estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
        estado.casoSeleccionado.historial.push({
            fecha: new Date().toISOString(),
            accion: 'Caso marcado como RESUELTO',
            usuario: 'Usuario'
        });

        try {
            await ApiCasos.resolverCaso(estado.casoSeleccionado.orden.referencia, 'resuelto');
            Interfaz.mostrarToast('Caso resuelto correctamente', 'success');
        } catch (error) {
            console.error('Error resolviendo caso en BD:', error);
            await Persistencia.guardarCasos();
        }

        const modalOpen = document.getElementById('casoDetalleOverlay').classList.contains('show');
        if (modalOpen) {
            this.cerrarDetalleCaso();
        } else {
            estado.casoSeleccionado = null;
            this._mostrarPlaceholderPanel();
        }
        this.renderizarSidebar();
    },

    reasignarCaso(rol) {
        if (!estado.casoSeleccionado || !rol) return;
        if (rol === estado.casoSeleccionado.responsable) return;

        estado.casoSeleccionado.responsable = rol;
        const nuevo = rol;
        estado.casoSeleccionado.fechaUltimaActualizacion = new Date().toISOString();
        estado.casoSeleccionado.historial.push({
            fecha: new Date().toISOString(),
            accion: `Reasignado a ${nuevo}`,
            usuario: 'Usuario'
        });

        Persistencia.guardarCasos();
        this._syncCasoEnBackground(estado.casoSeleccionado);

        const modalOpen = document.getElementById('casoDetalleOverlay').classList.contains('show');
        if (modalOpen) {
            this.abrirDetalleCaso(estado.casoSeleccionado.id);
        } else {
            this.renderizarPortal();
            this._renderizarDetalleEnPanel(estado.casoSeleccionado);
        }
    }
};

// ========== MÓDULO DASHBOARD ==========
const DashboardManager = {
    async cargarUltimaConciliacion() {
        try {
            const resp = await apiFetch(`${estado.backendUrl}/historico/conciliaciones?limit=1`);
            if (resp.ok) {
                const data = await resp.json();
                const lista = Array.isArray(data) ? data : (data.items || data.conciliaciones || []);
                if (lista.length > 0) {
                    document.getElementById('dash-ultima-fecha').textContent = Utilidades.formatFecha(lista[0].fecha);
                    document.getElementById('dash-ultima-ordenes').textContent = `${lista[0].totalOrdenes} órdenes`;
                    return;
                }
            }
        } catch (e) {}
        // Fallback localStorage
        try {
            const historial = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEYS.HISTORIAL) || '[]');
            if (historial.length > 0) {
                document.getElementById('dash-ultima-fecha').textContent = Utilidades.formatFecha(historial[0].fecha);
                document.getElementById('dash-ultima-ordenes').textContent = `${historial[0].totalOrdenes} órdenes`;
            }
        } catch (e2) {}
    },

    actualizarDashboard() {
        const activos    = estado.casosPendientes.filter(c => c.estado !== 'RESUELTO');
        const pendientes = activos.filter(c => c.estado === 'PENDIENTE');
        const enRevision = activos.filter(c => c.estado === 'EN_REVISION');
        const resueltos  = estado.casosPendientes.filter(c => c.estado === 'RESUELTO');

        // Stat cards
        document.getElementById('dash-casos-activos').textContent = activos.length;
        document.getElementById('dash-pendientes').textContent    = pendientes.length;
        document.getElementById('dash-en-revision').textContent   = enRevision.length;
        document.getElementById('dash-resueltos').textContent     = resueltos.length;

        // Estadísticas por rol (solo activos)
        document.getElementById('role-contador').textContent  = activos.filter(c => c.responsable === 'CONTADOR').length;
        document.getElementById('role-finanzas').textContent  = activos.filter(c => c.responsable === 'FINANZAS').length;
        document.getElementById('role-logistica').textContent = activos.filter(c => c.responsable === 'LOGISTICA').length;

        // Lista de pendientes sin atender (máx 5, ordenados por fecha más antigua primero)
        const listaPendientes = document.getElementById('lista-pendientes');
        if (!listaPendientes) return;

        const sinAtender = pendientes
            .slice()
            .sort((a, b) => new Date(a.fechaCreacion) - new Date(b.fechaCreacion))
            .slice(0, 5);

        if (sinAtender.length === 0) {
            listaPendientes.innerHTML = '<p style="color: var(--gray-500); text-align: center; padding: 20px;">No hay casos pendientes</p>';
        } else {
            listaPendientes.innerHTML = sinAtender.map(caso => `
                <div class="urgente-item">
                    <div class="urgente-item-info">
                        <h4>${caso.tipo} — ${caso.orden.referencia}</h4>
                        <p>${Utilidades.formatMoney(Math.abs(caso.orden.monto))} • ${caso.responsable} • ${Utilidades.formatFecha(caso.fechaCreacion)}</p>
                    </div>
                    <div class="urgente-item-actions">
                        <button class="btn btn-sm btn-primary" onclick="CasosManager.abrirDetalleCaso('${caso.id}')">Ver</button>
                    </div>
                </div>
            `).join('');
        }
    }
};



// ========== FUNCIONES DE CONSOLIDACIÓN Y VALIDACIÓN ==========

/**
 * Extrae y clasifica impuestos de un movimiento
 */
function extraerImpuestos(movimiento) {
    const impuestos = [];

    if (movimiento.TAXES_AMOUNT && movimiento.TAXES_AMOUNT !== 0) {
        const montoImpuesto = movimiento.TAXES_AMOUNT;

        let tipo = 'Otro';
        let tasa = null;
        let baseGravable = null;

        if (movimiento.TAX_DETAIL) {
            const taxDetailLower = (movimiento.TAX_DETAIL || '').toLowerCase();

            if (taxDetailLower.includes('vat') || taxDetailLower.includes('iva')) {
                tipo = 'VAT (IVA)';
                tasa = '16%';
                baseGravable = Utilidades.roundMoney(montoImpuesto / CONFIG.tasaIVA);
            } else if (taxDetailLower.includes('isr') || taxDetailLower.includes('retenc')) {
                tipo = 'ISR (Retención)';
            } else if (taxDetailLower.includes('ieps')) {
                tipo = 'IEPS';
            }
        }

        impuestos.push({
            tipo: tipo,
            monto: montoImpuesto,
            tasa: tasa,
            baseGravable: baseGravable
        });
    }

    return impuestos;
}

/**
 * Calcula todos los costos operativos de una orden
 */
function calcularCostosOperativos(movimientos) {
    const refunds = movimientos.filter(m =>
        m.DESCRIPTION && m.DESCRIPTION.toLowerCase() === 'refund'
    );

    const tieneReembolsos = refunds.some(r =>
        (r.MP_FEE_AMOUNT && r.MP_FEE_AMOUNT > 0) ||
        (r.SHIPPING_FEE_AMOUNT && r.SHIPPING_FEE_AMOUNT > 0)
    );

    if (tieneReembolsos) {
        let totalComision = 0;
        let totalEnvio = 0;
        let totalFinanciamiento = 0;

        movimientos.forEach(m => {
            totalComision += (m.MP_FEE_AMOUNT || 0);
            totalEnvio += (m.SHIPPING_FEE_AMOUNT || 0);
            totalFinanciamiento += (m.FINANCING_FEE_AMOUNT || 0);
        });

        return {
            comision: Utilidades.roundMoney(Math.abs(totalComision)),
            envio: Utilidades.roundMoney(Math.abs(totalEnvio)),
            financiamiento: Utilidades.roundMoney(Math.abs(totalFinanciamiento)),
            total: Utilidades.roundMoney(Math.abs(totalComision) + Math.abs(totalEnvio) + Math.abs(totalFinanciamiento))
        };
    }

    let comision = 0;
    let financiamiento = 0;
    let envio = 0;

    movimientos.forEach(mov => {
        comision += (mov.MP_FEE_AMOUNT || 0);
        financiamiento += (mov.FINANCING_FEE_AMOUNT || 0);
        let envioMov = (mov.SHIPPING_FEE_AMOUNT || 0);
        if (mov.DESCRIPTION && mov.DESCRIPTION.toLowerCase() === 'shipping fee') {
            envioMov += (mov.NET_CREDIT_AMOUNT || 0);
        }
        envio += envioMov;
    });

    return {
        comision: Utilidades.roundMoney(comision),
        financiamiento: Utilidades.roundMoney(financiamiento),
        envio: Utilidades.roundMoney(envio),
        total: Utilidades.roundMoney(comision + financiamiento + envio)
    };
}

/**
 * Detecta alertas en los movimientos
 */
function detectarAlertas(movimientos, diferenciaNeto) {
    return [];
}

/**
 * FUNCIÓN PRINCIPAL: Consolida una orden con toda su lógica de validación
 */
/**
 * Construye una orden consolidada a partir de las filas del Estado de Cuenta
 * y los datos enriquecidos del API de MercadoPago.
 *
 * @param {string}   paymentId       - REFERENCE_ID del Estado de Cuenta
 * @param {object[]} edoCuentaRows   - Filas del Estado de Cuenta con ese REFERENCE_ID
 * @param {object|null} paymentData  - Datos del API (null = no es pago o no se encontró)
 * @param {string}   tipoReferencia  - 'PAGO' | 'TRANSFERENCIA' | 'FISCAL'
 */
function consolidarOrden(paymentId, edoCuentaRows, paymentData, tipoReferencia) {
    // Neto total registrado por ML: suma de TODOS los movimientos del estado de cuenta
    // para este REFERENCE_ID (liberaciones + retenciones + ajustes post-liberación)
    const netoEdoCuenta = Utilidades.roundMoney(
        edoCuentaRows.reduce((sum, r) => sum + (r.TRANSACTION_NET_AMOUNT || 0), 0)
    );
    const transactionTypes = edoCuentaRows.map(r => r.TRANSACTION_TYPE).filter(Boolean);

    // ── Clasificar movimientos del estado de cuenta ───────────────────────────
    // Detecta montos que afectan el neto pero no están en charges_details del payment.
    // El estado de cuenta es la fuente de verdad; estos valores explican las diferencias.
    let retencion_ml   = 0;  // Dinero bloqueado por reclamo/disputa (Dinero retenido...)
    let ajuste_isr_iva = 0;  // Cobros adicionales post-liberación (Recobro retención / Ajuste ISR)

    edoCuentaRows.forEach(row => {
        const tipo  = (row.TRANSACTION_TYPE || '').toLowerCase();
        const monto = row.TRANSACTION_NET_AMOUNT || 0;
        if (tipo.includes('retenido')) {
            retencion_ml += monto;    // negativo: ML congeló el dinero por reclamo/disputa
        } else if (tipo.includes('devoluci') && tipo.includes('reclam')) {
            retencion_ml += monto;    // positivo: ML liberó la retención (reclamo resuelto a favor vendedor)
        } else if (tipo.includes('ajuste en tus retenciones') || tipo.includes('recobro de retenci')) {
            ajuste_isr_iva += monto;  // negativo: cobro adicional post-liberación de ISR/IVA
        }
    });
    retencion_ml   = Utilidades.roundMoney(retencion_ml);
    ajuste_isr_iva = Utilidades.roundMoney(ajuste_isr_iva);

    const sourceIds = [paymentId];
    const tieneMultiplesSourceIds = false;

    // === TRANSFERENCIA / FISCAL: sin datos del API ===
    if (tipoReferencia === 'TRANSFERENCIA' || tipoReferencia === 'FISCAL') {
        return {
            paymentId,
            referenciaERP: paymentId,
            orderId: '',
            estatus: tipoReferencia,
            statusML: '',
            statusDetailML: '',
            fechaAprobacion: '',
            fechaLiberacion: '',
            ingresos: { bruto: 0 },
            costos:  { comision: 0, isr: 0, iva: 0, envio: 0, financiamiento: 0, shipping: 0, total: 0 },
            ajustes: { retencion_ml: 0, ajuste_isr_iva: 0, envio_comprador_sep: 0 },
            neto: { real: netoEdoCuenta, estadoCuenta: netoEdoCuenta },
            validacion: { saldoValidado: true, diferencia: 0 },
            transactionTypes, sourceIds, tieneMultiplesSourceIds,
            apiData: { encontrado: false },
            api: undefined
        };
    }

    // === PAGO sin datos del API (404 o error de red) ===
    if (!paymentData) {
        return {
            paymentId,
            referenciaERP: paymentId,
            orderId: '',
            estatus: 'REVISAR',
            statusML: '',
            statusDetailML: '',
            fechaAprobacion: '',
            fechaLiberacion: '',
            ingresos: { bruto: 0 },
            costos:  { comision: 0, isr: 0, iva: 0, envio: 0, financiamiento: 0, shipping: 0, total: 0 },
            ajustes: { retencion_ml, ajuste_isr_iva, envio_comprador_sep: 0 },
            neto: { real: netoEdoCuenta, estadoCuenta: netoEdoCuenta },
            validacion: { saldoValidado: false, diferencia: 0 },
            transactionTypes, sourceIds, tieneMultiplesSourceIds,
            apiData: { encontrado: false },
            api: undefined
        };
    }

    // === PAGO con datos del API ===
    const costos = parsearCostosAPI(paymentData);

    // Usamos el estado de cuenta como fuente de verdad del neto real.
    // net_received_amount de la API puede no reflejar retenciones ni ajustes post-liberación.
    const netoReal = netoEdoCuenta;

    // Calcular envio_comprador_sep: el shipping del comprador llega a veces en una liberación
    // separada (distinto REFERENCE_ID). Se detecta como residual positivo después de descontar
    // todos los costos conocidos y los ajustes del estado de cuenta.
    // NOTA: se suma costos.shipping porque ML incluye el pago del comprador por envío en el
    // net_received_amount (y por tanto en neto_ec), aunque no esté en transaction_amount.
    // Al incluirlo aquí el residual queda en 0 cuando el envío llega en la misma liberación,
    // reservando env_sep sólo para el caso real: shipping en un REFERENCE_ID distinto.
    const costosTotales = costos.comision + costos.isr + costos.iva + costos.envio + costos.financiamiento;
    const netoCalculado = Utilidades.roundMoney(
        (paymentData.transaction_amount || 0) + costos.shipping - costosTotales + retencion_ml + ajuste_isr_iva
    );
    const residual = Utilidades.roundMoney(netoReal - netoCalculado);
    const envio_comprador_sep = residual > 0.5 ? residual : 0;

    // Diferencia final: debería ser 0 si todos los ajustes fueron capturados
    const diferencia = Utilidades.roundMoney(
        netoReal - netoCalculado - envio_comprador_sep
    );

    return {
        paymentId,
        referenciaERP: paymentData.order_id || paymentData.external_reference || paymentId,
        orderId:        paymentData.order_id || '',
        estatus:        clasificarPorStatusML(paymentData.status, paymentData.status_detail),
        statusML:       paymentData.status || '',
        statusDetailML: paymentData.status_detail || '',
        fechaAprobacion: paymentData.date_approved || '',
        fechaLiberacion: paymentData.money_release_date || '',
        ingresos: {
            bruto: Utilidades.roundMoney(paymentData.transaction_amount || 0)
        },
        costos,
        ajustes: {
            retencion_ml,        // negativo: dinero congelado por reclamo
            ajuste_isr_iva,      // negativo: cobro adicional post-liberación
            envio_comprador_sep  // positivo: shipping comprador en liberación aparte
        },
        neto: {
            real:         netoReal,
            estadoCuenta: netoEdoCuenta
        },
        validacion: {
            saldoValidado: Math.abs(diferencia) <= CONFIG.toleranciaNeto,
            diferencia
        },
        transactionTypes, sourceIds, tieneMultiplesSourceIds,
        apiData: {
            encontrado:         true,
            moneyReleaseStatus: paymentData.money_release_status || '',
            metodoPago:         paymentData.payment_method_id || '',
            installments:       paymentData.installments || 1,
            descripcion:        paymentData.description || ''
        },
        api: undefined  // se llena después por ApiManager.conciliarConIntelisis
    };
}

// ========== HELPERS DE ENRIQUECIMIENTO API ML ==========

/**
 * Extrae el desglose de costos a partir de los datos parseados del API.
 * @param {object} paymentData - Objeto devuelto por parsear_payment del backend
 * @returns {object} { comision, isr, iva, envio, total }
 */
function parsearCostosAPI(paymentData) {
    const comision       = paymentData.comision       || 0;
    const isr            = paymentData.isr            || 0;
    const iva            = paymentData.iva            || 0;
    // envio          = costo cobrado al vendedor (shp_fulfillment) → EGRESO  → negativo en exports
    // shipping       = lo que pagó el comprador                    → INGRESO → positivo en exports
    // financiamiento = cuota MSI retenida por ML al vendedor       → EGRESO  → negativo en exports
    const envio          = paymentData.envio           || 0;
    const shipping       = paymentData.shipping_amount || 0;
    const financiamiento = paymentData.financiamiento  || 0;
    const total          = comision + isr + iva + envio + financiamiento;
    return { comision, isr, iva, envio, shipping, financiamiento, total: Math.round(total * 100) / 100 };
}

/**
 * Llama al backend para enriquecer una lista de payment IDs con datos de la API de ML.
 * Emite eventos de progreso a través del callback onProgress(procesados, total).
 * @param {string[]} paymentIds
 * @param {function} [onProgress]
 * @returns {Promise<object>} Mapa { paymentId: datosParsados | null }
 */
async function enriquecerConAPI(paymentIds, onProgress) {
    if (!paymentIds || paymentIds.length === 0) return {};

    const LOTE = 10;
    const resultado = {};
    let procesados = 0;
    const total = paymentIds.length;

    for (let i = 0; i < paymentIds.length; i += LOTE) {
        const lote = paymentIds.slice(i, i + LOTE);

        const resp = await apiFetch(`${estado.backendUrl}/payments/batch`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ payment_ids: lote })
        });

        if (!resp.ok) {
            console.error(`Error batch payments lote ${i}: ${resp.status}`);
            lote.forEach(id => { resultado[id] = null; });
        } else {
            const data = await resp.json();
            Object.assign(resultado, data.results || {});
        }

        procesados += lote.length;
        if (typeof onProgress === 'function') onProgress(procesados, total);
    }

    return resultado;
}

// ========== MÓDULO CONCILIACIÓN ==========
const ConciliacionManager = {
    async procesarArchivos() {
        document.getElementById('loading').classList.add('show');
        document.getElementById('loadingText').textContent = 'Leyendo Estado de Cuenta...';

        await Utilidades.sleep(300);

        try {
            // Solo se requiere el Estado de Cuenta (archivo1Data)
            if (!estado.archivo1Data || estado.archivo1Data.length === 0) {
                throw new Error('El Estado de Cuenta está vacío o no se cargó correctamente');
            }

            estado.ordenesConsolidadas = [];

            // 1. Agrupar filas del Estado de Cuenta por REFERENCE_ID
            const edoCuentaMap = {};
            estado.archivo1Data.forEach(row => {
                const refId = row.REFERENCE_ID;
                if (!refId || refId === 'undefined' || refId.toString().trim() === '') return;
                const id = normalizarId(refId);
                if (!id || id === 'NaN') return;
                if (!edoCuentaMap[id]) edoCuentaMap[id] = [];
                edoCuentaMap[id].push(row);
            });

            const todosIds = Object.keys(edoCuentaMap);
            console.log(`REFERENCE_IDs únicos en Estado de Cuenta: ${todosIds.length}`);

            // 2. Separar pagos de no-pagos (TRANSFERENCIA / FISCAL)
            const paymentIds = [];
            const noPagosMap = {};  // { id: tipoReferencia }
            todosIds.forEach(id => {
                const tipos = edoCuentaMap[id].map(r => r.TRANSACTION_TYPE).filter(Boolean);
                const tipo = detectarTipoReferencia(tipos);
                if (tipo === 'PAGO') {
                    paymentIds.push(id);
                } else {
                    noPagosMap[id] = tipo;
                }
            });

            console.log(`Pagos a consultar API: ${paymentIds.length} | No-pagos (TRANSFERENCIA/FISCAL): ${Object.keys(noPagosMap).length}`);

            // 3. Enriquecer pagos con la API de ML (con progreso en UI)
            let apiResultados = {};
            if (paymentIds.length > 0) {
                document.getElementById('loadingText').textContent = `Consultando API ML: 0/${paymentIds.length} pagos...`;

                apiResultados = await enriquecerConAPI(paymentIds, (procesados, total) => {
                    document.getElementById('loadingText').textContent =
                        `Consultando API ML: ${procesados}/${total} pagos...`;
                });
            }

            document.getElementById('loadingText').textContent = 'Consolidando órdenes...';
            await Utilidades.sleep(100);

            // 4. Construir órdenes consolidadas para pagos
            paymentIds.forEach(id => {
                const orden = consolidarOrden(id, edoCuentaMap[id], apiResultados[id] || null, 'PAGO');
                if (orden) estado.ordenesConsolidadas.push(orden);
            });

            // 5. Construir órdenes para no-pagos (sin API)
            Object.entries(noPagosMap).forEach(([id, tipo]) => {
                const orden = consolidarOrden(id, edoCuentaMap[id], null, tipo);
                if (orden) estado.ordenesConsolidadas.push(orden);
            });

            // 6. Post-consolidar por referenciaERP (múltiples REFERENCE_IDs → mismo order/external_reference)
            const ordenesAgrupadas = {};
            estado.ordenesConsolidadas.forEach(orden => {
                const key = orden.referenciaERP;
                if (!ordenesAgrupadas[key]) ordenesAgrupadas[key] = [];
                ordenesAgrupadas[key].push(orden);
            });

            estado.ordenesConsolidadas = Object.values(ordenesAgrupadas).map(grupo => {
                if (grupo.length === 1) return grupo[0];

                // Múltiples REFERENCE_IDs para la misma orden — tomar la VENTA como principal o la primera
                const principal = grupo.find(o => o.estatus === 'VENTA') || grupo[0];
                const allPaymentIds = grupo.map(o => o.paymentId);

                const bruto          = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.ingresos.bruto || 0), 0));
                const comision       = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.comision || 0), 0));
                const isr            = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.isr || 0), 0));
                const iva            = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.iva || 0), 0));
                const envio          = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.envio || 0), 0));
                const financiamiento = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.financiamiento || 0), 0));
                const shipping       = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.costos.shipping || 0), 0));
                const costoTotal     = Utilidades.roundMoney(comision + isr + iva + envio + financiamiento);

                const retencion_ml       = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.ajustes?.retencion_ml || 0), 0));
                const ajuste_isr_iva     = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.ajustes?.ajuste_isr_iva || 0), 0));
                const envio_comprador_sep= Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.ajustes?.envio_comprador_sep || 0), 0));

                const netoReal   = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.neto.real || 0), 0));
                const netoEdoCta = Utilidades.roundMoney(grupo.reduce((s, o) => s + (o.neto.estadoCuenta || 0), 0));
                const diferencia = Utilidades.roundMoney(netoReal -
                    (bruto - costoTotal + retencion_ml + ajuste_isr_iva + envio_comprador_sep)
                );

                return {
                    ...principal,
                    sourceIds: allPaymentIds,
                    tieneMultiplesSourceIds: true,
                    ingresos: { bruto },
                    costos:  { comision, isr, iva, envio, financiamiento, shipping, total: costoTotal },
                    ajustes: { retencion_ml, ajuste_isr_iva, envio_comprador_sep },
                    neto: { real: netoReal, estadoCuenta: netoEdoCta },
                    validacion: {
                        saldoValidado: Math.abs(diferencia) <= CONFIG.toleranciaNeto,
                        diferencia
                    },
                    transactionTypes: [...new Set(grupo.flatMap(o => o.transactionTypes || []))]
                };
            });

            console.log(`Órdenes consolidadas: ${estado.ordenesConsolidadas.length}`);
            estado.ordenesConsolidadas.forEach(o =>
                console.log(`  ${o.paymentId} → ${o.referenciaERP} | ${o.estatus} | $${o.neto.real}`)
            );

            // 6.5 Detectar payments pendientes de liberar (MSI / pago dividido parcial)
            // Para cada orden VENTA, consultamos la orden en ML y comparamos sus payment IDs
            // con los que ya están en el estado de cuenta. Los que faltan = aún no liberados.
            document.getElementById('loadingText').textContent = 'Detectando pagos pendientes de liberar...';
            await Utilidades.sleep(100);

            const ordenesVenta = estado.ordenesConsolidadas.filter(o => o.estatus === 'VENTA');
            const orderIdsUnicos = [...new Set(
                ordenesVenta
                    .map(o => o.orderId)
                    .filter(id => id && id !== '')
            )];

            if (orderIdsUnicos.length > 0) {
                try {
                    const respOrdenes = await apiFetch(`${estado.backendUrl}/orders/batch`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ order_ids: orderIdsUnicos })
                    });

                    if (respOrdenes.ok) {
                        const dataOrdenes = await respOrdenes.json();
                        const ordenesML = dataOrdenes.results || {};

                        // Construir set de payment IDs que SÍ están en el estado de cuenta
                        const paymentIdsEnEC = new Set(paymentIds);

                        estado.ordenesConsolidadas = estado.ordenesConsolidadas.map(orden => {
                            if (orden.estatus !== 'VENTA' || !orden.orderId) return orden;

                            const ordenML = ordenesML[orden.orderId];
                            if (!ordenML) return orden;

                            // Payments de esta orden que NO están en el estado de cuenta
                            const paymentsPendientes = (ordenML.payments || []).filter(p =>
                                !paymentIdsEnEC.has(p.id) &&
                                (p.status === 'approved' || p.status === 'in_mediation')
                            );

                            if (paymentsPendientes.length === 0) return orden;

                            // Sumar montos pendientes de liberar
                            const montoPendiente = Utilidades.roundMoney(
                                paymentsPendientes.reduce((s, p) => s + (p.amount || 0), 0)
                            );
                            const tieneMSI = paymentsPendientes.some(p => (p.installments || 1) > 1);
                            const maxCuotas = Math.max(...paymentsPendientes.map(p => p.installments || 1));

                            console.log(`Orden ${orden.referenciaERP}: ${paymentsPendientes.length} payment(s) pendiente(s) de liberar, monto: $${montoPendiente}${tieneMSI ? ` (MSI ${maxCuotas} meses)` : ''}`);

                            return {
                                ...orden,
                                paymentsPendientes,          // lista completa por si se necesita en detalle
                                montoPendienteLiberar: montoPendiente,
                                tienePagosMSI: tieneMSI,
                                msiMaxCuotas: tieneMSI ? maxCuotas : 1,
                            };
                        });

                        const conPendientes = estado.ordenesConsolidadas.filter(o => o.montoPendienteLiberar > 0).length;
                        if (conPendientes > 0) {
                            console.log(`${conPendientes} órdenes con payments pendientes de liberar`);
                        }
                    }
                } catch (err) {
                    // No es crítico — si falla, el cuadre sigue funcionando sin el desglose MSI
                    console.warn('No se pudo consultar /orders/batch:', err.message);
                }
            }

            // 6. Auto-crear casos
            const casosCreados = CasosManager.crearCasosAutomaticos();

            // 7. Guardar historial
            await Persistencia.guardarHistorialConciliacion({
                fecha: new Date().toISOString(),
                totalOrdenes: estado.ordenesConsolidadas.length,
                casosCreados,
                coincidencias: paymentIds.length
            });

            // 8. Mostrar resultados
            const encontrados = Object.values(apiResultados).filter(Boolean).length;
            this.mostrarResultados({
                totalReferenceIds: todosIds.length,
                totalPagos: paymentIds.length,
                totalEncontradosAPI: encontrados,
                totalNoPagos: Object.keys(noPagosMap).length,
                totalOrdenes: estado.ordenesConsolidadas.length,
                casosCreados
            });

            if (estado.ordenesConsolidadas.length > 0) {
                document.getElementById('processBtn').textContent = 'Procesar otra Conciliación';
                Interfaz.mostrarToast(`${estado.ordenesConsolidadas.length} órdenes procesadas (${encontrados} con datos API)`, 'success');
            }

            CasosManager.renderizarSidebar();

        } catch (error) {
            console.error('Error en procesarArchivos:', error);
            alert('Error: ' + error.message);
        } finally {
            document.getElementById('loading').classList.remove('show');
        }
    },

    mostrarResultados(stats) {
        document.getElementById('results').classList.add('show');

        document.getElementById('statsGrid').innerHTML = `
            <div class="stat-card">
                <h4>Edo. Cuenta (Total)</h4>
                <div class="number">${stats.totalReferenceIds}</div>
            </div>
            <div class="stat-card" style="border-top-color: var(--success)">
                <h4>Pagos consultados</h4>
                <div class="number">${stats.totalPagos}</div>
            </div>
            <div class="stat-card" style="border-top-color: #10b981">
                <h4>Encontrados en API</h4>
                <div class="number" style="color: #10b981">${stats.totalEncontradosAPI}</div>
            </div>
            <div class="stat-card" style="border-top-color: #0891b2">
                <h4>Transferencias/Fiscal</h4>
                <div class="number" style="color: #0891b2">${stats.totalNoPagos}</div>
            </div>
            <div class="stat-card" style="border-top-color: var(--success)">
                <h4>Órdenes Procesadas</h4>
                <div class="number">${stats.totalOrdenes}</div>
            </div>
            <div class="stat-card" style="border-top-color: var(--warning)">
                <h4>Casos Creados</h4>
                <div class="number">${stats.casosCreados}</div>
            </div>
        `;

        document.getElementById('dash-ultima-fecha').textContent = Utilidades.formatFecha(new Date());
        document.getElementById('dash-ultima-ordenes').textContent = `${stats.totalOrdenes} órdenes procesadas`;

        if (estado.ordenesConsolidadas && estado.ordenesConsolidadas.length > 0) {
            ResumenEstatusManager.generarResumen(estado.ordenesConsolidadas);
            AnalisisValidacionManager.generarAnalisis();
        }
    }
};

// ========== MÓDULO API ==========
const ApiManager = {
    async fetchVentaData(referenciaERP) {
        // Validar que la referencia no sea undefined, vacía o 'undefined'
        if (!referenciaERP || referenciaERP === 'undefined' || referenciaERP.toString().trim() === '') {
            console.warn('Referencia ERP inválida:', referenciaERP);
            return null;
        }

        try {
            // El backend espera el endpoint: /ventas/venta-id/{venta_id}
            const url = `${estado.backendUrl}/ventas/venta-id/${encodeURIComponent(referenciaERP.toString().trim())}`;
            console.log(`Consultando API: ${url}`);

            const response = await apiFetch(url, {
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    'Accept': 'application/json'
                }
            });

            if (!response.ok) {
                if (response.status === 404) {
                    console.log(`No encontrado en Intelisis: ${referenciaERP}`);
                    return null;
                }
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            console.log(`Respuesta API para ${referenciaERP}:`, data);
            return data;
        } catch (error) {
            console.error(`Error fetching venta ${referenciaERP}:`, error);
            return null;
        }
    },

    processApiData(apiData) {
        if (!apiData || !Array.isArray(apiData) || apiData.length === 0) {
            return {
                encontrado: false,
                movimientos: [],
                totalMovimientos: 0,
                importeTotal: 0,
                estatusERP: 'NO_ENCONTRADO',
                cliente: '',
                movIDs: '',
                detalle: []
            };
        }

        // El backend devuelve un array de objetos con la estructura del modelo PedidoMeli
        const movimientos = apiData.map(m => ({
            mov: m.Mov || '',
            movID: m.MovID || '',
            estatus: m.Estatus || '',
            importe: m.Importe || 0,
            impuestos: m.Impuestos || 0,
            saldo: m.Saldo || 0,
            total: m.PrecioTotal || 0,
            cliente: m.Cliente || '',
            referencia: m.Referencia || '',
            atencion: m.Atencion || '',
            fechaEmision: m.FechaEmision || ''
        }));

        // Extraer información relevante
        const importeTotal = apiData[0]?.PrecioTotal || 0;
        const estatuses = [...new Set(apiData.map(m => m.Estatus))];
        const cliente = apiData[0]?.Cliente || '';
        const movIDs = apiData.map(m => m.MovID).join(', ');
        const embarqueInfo = {
            embarque: apiData[0]?.Embarque || '',
            observaciones: apiData[0]?.EmbarqueObservaciones || '',
            agente: apiData[0]?.EmbarqueAgente || '',
            estatus: apiData[0]?.EmbarqueEstatus || ''
        };

        return {
            encontrado: true,
            movimientos,
            totalMovimientos: apiData.length,
            importeTotal,
            estatusERP: estatuses.join(', '),
            movIDs,
            cliente,
            embarqueInfo,
            detalle: apiData
        };
    },

    async conciliarConIntelisis() {
        // Filtrar solo órdenes con referenciaERP válida
        const ordenesValidas = estado.ordenesConsolidadas.filter(orden => 
            orden.referenciaERP && 
            orden.referenciaERP !== 'undefined' && 
            orden.referenciaERP.toString().trim() !== ''
        );

        if (ordenesValidas.length === 0) {
            alert('No hay órdenes válidas para conciliar. Asegúrate de procesar los archivos primero.');
            return;
        }

        const confirmacion = confirm(`¿Deseas conciliar ${ordenesValidas.length} órdenes válidas con Intelisis?\n\nEsto consultará la base de datos para verificar el estado de cada orden.`);
        if (!confirmacion) return;

        Interfaz.mostrarToast(`Iniciando conciliación con Intelisis para ${ordenesValidas.length} órdenes...`, 'info');

        let procesadas = 0;
        let encontradas = 0;
        let noEncontradas = 0;
        let conDiferencias = 0;

        // Limpiar estadísticas anteriores
        estado.estadisticasConciliacion = { encontradas: 0, noEncontradas: 0, conDiferencias: 0 };

        for (let orden of ordenesValidas) {
            console.log(`Procesando orden: ${orden.referenciaERP}`);
            let apiData = await this.fetchVentaData(orden.referenciaERP);
            const resultado = this.processApiData(apiData);

            orden.api = {
                encontrado: resultado.encontrado,
                estatusERP: resultado.estatusERP,
                movimientos: resultado.movIDs || 'N/A',
                importeERP: resultado.importeTotal,
                diferencia: resultado.encontrado ?
                    Utilidades.roundMoney((orden.ingresos.bruto + (orden.costos?.shipping || 0)) - resultado.importeTotal) : null,
                cliente: resultado.cliente || '',
                movimientosDetalle: resultado.movimientos || [],
                embarqueInfo: resultado.embarqueInfo || {},
                detalle: resultado.detalle
            };

            if (resultado.encontrado) {
                encontradas++;
                if (orden.api.diferencia && Math.abs(orden.api.diferencia) > 0.01) {
                    conDiferencias++;
                    console.log(`Diferencia encontrada en ${orden.referenciaERP}: ${orden.api.diferencia}`);
                }
            } else {
                noEncontradas++;
                console.log(`No encontrada en Intelisis: ${orden.referenciaERP}`);
            }

            procesadas++;

            // Actualizar progreso cada 5 órdenes
            if (procesadas % 5 === 0 || procesadas === ordenesValidas.length) {
                Interfaz.mostrarToast(`Procesadas: ${procesadas}/${ordenesValidas.length}`, 'info');
            }

            await Utilidades.sleep(200); // Pequeña pausa para no saturar el servidor
        }

        estado.estadisticasConciliacion = {
            encontradas,
            noEncontradas,
            conDiferencias
        };

        // Actualizar estadísticas en la UI
        document.getElementById('stat-encontradas').textContent = encontradas;
        document.getElementById('stat-no-encontradas').textContent = noEncontradas;
        document.getElementById('stat-diferencias').textContent = conDiferencias;
        document.getElementById('estadisticas-conciliacion').classList.add('show');

        // Identificar pagos CxC
        PagosCxCManager.identificarPagosCxC();

        // Actualizar casos existentes con información del ERP
        CasosManager.actualizarCasosConERP();

        // Crear casos para ventas VENTA con diferencia ML vs ERP
        const casosVentaDif = CasosManager.crearCasosVentaDiferencia();

        // Sincronizar casos con la base de datos (solo después de procesar conciliación)
        await Persistencia.sincronizarCasosConBD();

        const msgDif = casosVentaDif > 0 ? ` | ${casosVentaDif} caso(s) de diferencia ML/ERP` : '';
        Interfaz.mostrarToast(`Conciliación completada: ${encontradas} encontradas, ${noEncontradas} no encontradas, ${conDiferencias} con diferencias${msgDif}`, 'success');

        DashboardManager.actualizarDashboard();
        CasosManager.renderizarSidebar();

        // Generar análisis ERP
        AnalisisERPManager.generarAnalisis();
    }
};

// ========== MÓDULO DE PAGOS CXC ==========
const PagosCxCManager = {
    identificarPagosCxC() {
        estado.pagosCxCData = [];
        
        // Filtrar SOLO órdenes VENTA que coinciden con ERP (sin diferencia de monto)
        const ordenesPagadas = estado.ordenesConsolidadas.filter(orden =>
            orden.estatus === 'VENTA' &&
            orden.api &&
            orden.api.encontrado &&
            (!orden.api.diferencia || Math.abs(orden.api.diferencia) <= 0.01) &&
            orden.referenciaERP &&
            orden.referenciaERP !== 'undefined'
        );

        console.log(`Identificando pagos CxC para ${ordenesPagadas.length} órdenes PAGADAS con API`);

        ordenesPagadas.forEach(orden => {
            const movimientos = orden.api.movimientosDetalle || [];
            if (movimientos.length === 0) {
                console.warn(`Orden ${orden.referenciaERP} no tiene movimientos detallados`);
                return;
            }

            // Filtrar pedidos y facturas de los movimientos
            const pedidos = movimientos.filter(m =>
                m.mov && m.mov.trim().toLowerCase().includes('pedido')
            );
            const facturas = movimientos.filter(m =>
                m.mov && m.mov.trim().toLowerCase().includes('factura')
            );

            // REGLA 1: ANTICIPO - Pedido PENDIENTE sin factura
            const soloPedidoPendiente = pedidos.length > 0 &&
                facturas.length === 0 &&
                pedidos.some(p => p.estatus && p.estatus.trim().toLowerCase() === 'pendiente');

            // REGLA 2: COBRO_FACTURA - Factura CONCLUIDA
            const facturasConc = facturas.filter(f =>
                f.estatus && f.estatus.trim().toLowerCase() === 'concluido'
            );
            const tieneFacturaConcluida = facturasConc.length > 0;

            let tipoPago = null;
            let pedidoMovID = '';
            let facturaMovID = '';
            let cliente = (orden.api.cliente || '').trim();

            // Determinar tipo de pago según las reglas
            if (soloPedidoPendiente) {
                tipoPago = 'ANTICIPO';
                const pedidoPendiente = pedidos.find(p =>
                    p.estatus && p.estatus.trim().toLowerCase() === 'pendiente'
                );
                pedidoMovID = pedidoPendiente ? (pedidoPendiente.movID || '').trim() : '';
                if (!cliente && pedidoPendiente) cliente = (pedidoPendiente.cliente || '').trim();
                
                console.log(`ANTICIPO identificado: ${orden.referenciaERP} - Pedido: ${pedidoMovID}`);
                
            } else if (tieneFacturaConcluida) {
                tipoPago = 'COBRO_FACTURA';
                // Buscar la factura más reciente si hay varias
                const facturaReciente = facturasConc.reduce((latest, current) => {
                    if (!latest) return current;
                    const latestDate = latest.fechaEmision || '';
                    const currentDate = current.fechaEmision || '';
                    return currentDate > latestDate ? current : latest;
                }, null);
                
                facturaMovID = facturaReciente ? (facturaReciente.movID || '').trim() : '';
                if (!cliente && facturaReciente) cliente = (facturaReciente.cliente || '').trim();

                // Si hay pedido asociado, tomarlo también
                if (pedidos.length > 0) {
                    pedidoMovID = (pedidos[0].movID || '').trim();
                }
                
                console.log(`COBRO_FACTURA identificado: ${orden.referenciaERP} - Factura: ${facturaMovID} - Pedido: ${pedidoMovID}`);
            }

            // Solo agregar si se identificó un tipo de pago válido
            if (tipoPago) {
                estado.pagosCxCData.push({
                    referenciaML: orden.referenciaERP,
                    tipoPago: tipoPago,
                    pedido: pedidoMovID,
                    factura: facturaMovID,
                    cliente: cliente,
                    formaCobro: 'MercadoLibre',
                    monto: orden.neto.real,
                    fechaLiberacion: orden.fechaLiberacion || new Date().toISOString().split('T')[0]
                });

                console.log(`Pago CxC agregado: ${orden.referenciaERP} - ${tipoPago} - $${orden.neto.real}`);
            } else {
                console.warn(`Orden ${orden.referenciaERP} no cumple criterios: Pedidos=${pedidos.length}, Facturas=${facturas.length}, FacturasConc=${facturasConc.length}`);
            }
        });

        // Actualizar UI
        document.getElementById('total-pagos-cxc').textContent = estado.pagosCxCData.length;
        
        if (estado.pagosCxCData.length > 0) {
            const anticipos = estado.pagosCxCData.filter(p => p.tipoPago === 'ANTICIPO').length;
            const cobros = estado.pagosCxCData.filter(p => p.tipoPago === 'COBRO_FACTURA').length;
            Interfaz.mostrarToast(`Identificados ${estado.pagosCxCData.length} pagos: ${anticipos} anticipos, ${cobros} cobros`, 'success');
        } else {
            Interfaz.mostrarToast('No se identificaron pagos para CxC', 'info');
        }
    },

    downloadPagosCxCExcel() {
        if (estado.pagosCxCData.length === 0) {
            alert('No hay pagos CxC para exportar. Primero ejecuta la conciliación con Intelisis.');
            return;
        }

        try {
            const wb = XLSX.utils.book_new();
            
            // Preparar datos para Excel con los nombres de columnas que espera el backend
            const excelData = estado.pagosCxCData.map(pago => {
                // Buscar la orden consolidada para obtener gastos
                const orden = (estado.ordenesConsolidadas || []).find(o => o.referenciaERP === pago.referenciaML);
                return {
                    'REFERENCIA_ML': pago.referenciaML,
                    'TIPO': pago.tipoPago,
                    'PEDIDO': pago.pedido,
                    'FACTURA': pago.factura,
                    'CLIENTE': pago.cliente,
                    'FORMA_COBRO': pago.formaCobro,
                    'MONTO': pago.monto,
                    'INGRESO_BRUTO':   +(orden?.ingresos?.bruto              || 0),
                    'ENVIO_COMPRADOR': +(orden?.costos?.shipping             || 0),
                    'COMISION_MP':     -Math.abs(orden?.costos?.comision     || 0),
                    'COSTO_ENVIO':     -Math.abs(orden?.costos?.envio        || 0),
                    'FINANCIAMIENTO':  -Math.abs(orden?.costos?.financiamiento || 0),
                    'IMPUESTOS':       -Math.abs(orden?.costos?.impuestos    || 0),
                    'TOTAL_GASTOS':    -Math.abs(orden?.costos?.total        || 0),
                    'FECHA': pago.fechaLiberacion
                };
            });

            // Crear hoja de Excel
            const ws = XLSX.utils.json_to_sheet(excelData);

            // Ajustar anchos de columna
            const colWidths = [
                { wch: 20 }, // REFERENCIA_ML
                { wch: 18 }, // TIPO
                { wch: 15 }, // PEDIDO
                { wch: 15 }, // FACTURA
                { wch: 25 }, // CLIENTE
                { wch: 16 }, // FORMA_COBRO
                { wch: 14 }, // MONTO
                { wch: 14 }, // INGRESO_BRUTO
                { wch: 14 }, // COMISION_MP
                { wch: 14 }, // COSTO_ENVIO
                { wch: 14 }, // FINANCIAMIENTO
                { wch: 14 }, // IMPUESTOS
                { wch: 14 }, // TOTAL_GASTOS
                { wch: 12 }  // FECHA
            ];
            ws['!cols'] = colWidths;
            
            // Agregar la hoja al libro
            XLSX.utils.book_append_sheet(wb, ws, 'Pagos CxC');
            
            // Generar archivo
            const fecha = new Date().toISOString().split('T')[0];
            const nombreArchivo = `Pagos_CxC_${fecha}.xlsx`;
            XLSX.writeFile(wb, nombreArchivo);
            
            Interfaz.mostrarToast(`Excel exportado: ${estado.pagosCxCData.length} pagos`, 'success');
        } catch (error) {
            console.error('Error exportando Excel:', error);
            alert('Error al exportar Excel: ' + error.message);
        }
    },

    async enviarPagosCxCPorCorreo() {
        if (estado.pagosCxCData.length === 0) {
            alert('No hay pagos CxC para enviar. Primero ejecuta la conciliación con Intelisis.');
            return;
        }

        const confirmacion = confirm(`¿Deseas enviar ${estado.pagosCxCData.length} pagos CxC por correo?\n\nSe enviará un Excel con los pagos a los destinatarios configurados.`);
        if (!confirmacion) return;

        try {
            Interfaz.mostrarToast('Enviando correo...', 'info');

            // Preparar datos en el formato que espera el backend
            const pagosCxC = estado.pagosCxCData.map(pago => {
                const orden = (estado.ordenesConsolidadas || []).find(o => o.referenciaERP === pago.referenciaML);
                return {
                    PEDIDO: pago.pedido || '',
                    FACTURA: pago.factura || '',
                    CLIENTE: pago.cliente || 'N/A',
                    FORMA_COBRO: pago.formaCobro || 'MercadoLibre',
                    TIPO: pago.tipoPago || '',
                    MONTO: pago.monto || 0,
                    REFERENCIA_ML: pago.referenciaML || '',
                    INGRESO_BRUTO:   +(orden?.ingresos?.bruto              || 0),
                    ENVIO_COMPRADOR: +(orden?.costos?.shipping             || 0),
                    COMISION_MP:     -Math.abs(orden?.costos?.comision     || 0),
                    COSTO_ENVIO:     -Math.abs(orden?.costos?.envio        || 0),
                    FINANCIAMIENTO:  -Math.abs(orden?.costos?.financiamiento || 0),
                    IMPUESTOS:       -Math.abs(orden?.costos?.impuestos    || 0),
                    TOTAL_GASTOS:    -Math.abs(orden?.costos?.total        || 0)
                };
            });

            // Calcular resumen
            const anticipos = pagosCxC.filter(p => p.TIPO === 'ANTICIPO');
            const cobrosFactura = pagosCxC.filter(p => p.TIPO === 'COBRO_FACTURA');
            const totalMonto = pagosCxC.reduce((sum, p) => sum + p.MONTO, 0);

            const resumen = {
                total_pagos: pagosCxC.length,
                total_anticipos: anticipos.length,
                total_cobros_factura: cobrosFactura.length,
                total_monto: totalMonto,
                fecha: new Date().toLocaleDateString('es-MX')
            };

            // Enviar al backend
            const response = await apiFetch(`${estado.backendUrl}/enviar-correo-pagos-cxc`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    pagos: pagosCxC,
                    resumen: resumen
                })
            });

            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.detail || 'Error al enviar correo');
            }

            const result = await response.json();
            Interfaz.mostrarToast(`${result.message}`, 'success');

        } catch (error) {
            console.error('Error enviando correo CxC:', error);
            Interfaz.mostrarToast(`Error: ${error.message}`, 'error');
            alert('Error al enviar correo: ' + error.message);
        }
    }
};

// ========== FUNCIONES DE EXPORTACIÓN E IMPORTACIÓN ==========
function exportarSeguimiento() {
    if (estado.casosPendientes.length === 0) {
        alert('No hay casos para exportar');
        return;
    }

    const exportData = {
        version: '3.0',
        fecha_exportacion: new Date().toISOString(),
        total_casos: estado.casosPendientes.length,
        casos: estado.casosPendientes
    };

    const dataStr = JSON.stringify(exportData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(dataBlob);

    const link = document.createElement('a');
    link.href = url;
    link.download = `seguimiento_conciliacion_${new Date().toISOString().split('T')[0]}.json`;
    link.click();

    URL.revokeObjectURL(url);

    Interfaz.mostrarToast(`${estado.casosPendientes.length} casos exportados exitosamente`, 'success');
}

function exportarCasos() {
    const wb = XLSX.utils.book_new();
    const casosData = estado.casosPendientes.map(c => {
        // Buscar la orden consolidada para obtener gastos
        const orden = (estado.ordenesConsolidadas || []).find(o => o.referenciaERP === c.orden?.referencia);
        return {
            'ID': c.id,
            'TIPO': c.tipo,
            'PRIORIDAD': c.prioridad,
            'ESTADO': c.estado,
            'RESPONSABLE': c.responsable,
            'REFERENCIA': c.orden.referencia,
            'MONTO': c.orden.monto,
            'INGRESO_BRUTO':   +(orden?.ingresos?.bruto              || 0),
            'ENVIO_COMPRADOR': +(orden?.costos?.shipping             || 0),
            'COMISION_MP':     -Math.abs(orden?.costos?.comision     || 0),
            'COSTO_ENVIO':     -Math.abs(orden?.costos?.envio        || 0),
            'FINANCIAMIENTO':  -Math.abs(orden?.costos?.financiamiento || 0),
            'IMPUESTOS':       -Math.abs(orden?.costos?.impuestos    || 0),
            'TOTAL_GASTOS':    -Math.abs(orden?.costos?.total        || 0),
            'FECHA_CREACION': c.fechaCreacion,
            'NOTAS': c.notas
        };
    });

    const ws = XLSX.utils.json_to_sheet(casosData);
    XLSX.utils.book_append_sheet(wb, ws, 'Casos');
    XLSX.writeFile(wb, `Casos_${new Date().toISOString().split('T')[0]}.xlsx`);
    
    Interfaz.mostrarToast('Exportación de casos completada', 'success');
}

async function importarSeguimiento(file) {
    if (!file) return;

    const reader = new FileReader();

    reader.onload = async (e) => {
        try {
            const importData = JSON.parse(e.target.result);

            if (!importData.casos || !Array.isArray(importData.casos)) {
                throw new Error('Formato de archivo inválido');
            }

            const casosImportados = importData.casos;
            let nuevos = 0;
            let actualizados = 0;

            for (const casoImportado of casosImportados) {
                const existente = estado.casosPendientes.find(c => c.id === casoImportado.id);

                if (existente) {
                    Object.assign(existente, casoImportado);
                    actualizados++;
                } else {
                    estado.casosPendientes.push(casoImportado);
                    nuevos++;
                }
            }

            Persistencia.guardarCasos();
            CasosManager.renderizarSidebar();

            Interfaz.mostrarToast(`Importación exitosa: ${nuevos} casos nuevos, ${actualizados} actualizados`, 'success');
            Interfaz.cerrarModalImportar();
        } catch (error) {
            console.error('Error importando:', error);
            alert('Error al importar archivo: ' + error.message);
        }
    };

    reader.readAsText(file);
}

// ========== INICIALIZACIÓN ==========
document.addEventListener('DOMContentLoaded', async () => {
    // Configurar interfaz base (no requiere auth)
    Interfaz.setupUploadZones();

    // Configurar input de importación
    document.getElementById('inputImportarJSON').addEventListener('change', function(e) {
        if (e.target.files[0]) {
            importarSeguimiento(e.target.files[0]);
        }
    });

    // Autenticar — si hay token válido, carga datos; si no, muestra login
    await AuthManager.init();

    console.log('Centro de Control inicializado');
});

// ========== CONFIGURACIÓN DE RESUMEN POR ESTATUS ==========
const CONFIG_ESTATUS = {
    'VENTA': {
        badge: 'venta',
        color: '#10b981',
        icon: '✓',
        accion: 'Generar PAGO en ERP',
        prioridad: 'rutina',
        descripcion: 'Pago acreditado exitosamente. Proceder con cobro normal.'
    },
    'DEVOLUCION_MP': {
        badge: 'devolucion-mp',
        color: '#dc2626',
        icon: '↩',
        accion: 'Verificar devolución por Protección al Comprador',
        prioridad: 'urgente',
        descripcion: 'ML devolvió al comprador por Programa de Protección al Comprador. Cargo reversado por ML.'
    },
    'DEVOLUCION_VENDEDOR': {
        badge: 'devolucion-vendedor',
        color: '#b91c1c',
        icon: '↩',
        accion: 'Registrar devolución voluntaria en ERP',
        prioridad: 'normal',
        descripcion: 'Devolución iniciada manualmente por el vendedor. Cargo reversado voluntario.'
    },
    'DEVOLUCION_CUBIERTA': {
        badge: 'devolucion-cubierta',
        color: '#d97706',
        icon: '↩',
        accion: 'Verificar cobertura ML',
        prioridad: 'normal',
        descripcion: 'ML absorbió el costo de la devolución. Sin impacto financiero para el vendedor.'
    },
    'MEDIACION_ABIERTA': {
        badge: 'mediacion-abierta',
        color: '#f59e0b',
        icon: '⚖',
        accion: 'Dar seguimiento a mediación en ML',
        prioridad: 'urgente',
        descripcion: 'Disputa abierta en ML, sin resolución. Monto en suspenso.'
    },
    'CONTRACARGO_EN_PROCESO': {
        badge: 'contracargo',
        color: '#be185d',
        icon: '⚠',
        accion: 'Atender contracargo bancario con urgencia',
        prioridad: 'urgente',
        descripcion: 'Chargeback iniciado por el banco del comprador. Monto en riesgo.'
    },
    'FISCAL': {
        badge: 'fiscal',
        color: '#0891b2',
        icon: '%',
        accion: 'Verificar ajuste fiscal',
        prioridad: 'rutina',
        descripcion: 'Ajuste fiscal (ISR/IVA). Verificar con contabilidad.'
    },
    'TRANSFERENCIA': {
        badge: 'transferencia',
        color: '#7c3aed',
        icon: '→',
        accion: 'Verificar transferencia bancaria',
        prioridad: 'rutina',
        descripcion: 'Retiro a cuenta bancaria. Verificar recepción.'
    },
    'REVISAR': {
        badge: 'revisar',
        color: '#3b82f6',
        icon: '?',
        accion: 'Revisión manual',
        prioridad: 'normal',
        descripcion: 'No se pudo obtener datos del API o status desconocido. Revisión manual requerida.'
    }
};

// ========== MÓDULO DE RESUMEN POR ESTATUS ==========
const ResumenEstatusManager = {
    generarResumen(ordenes) {
        // 1. Agrupar órdenes por estatus
        const porEstatus = {};
        let totalOrdenes = 0;
        let montoTotalEdoCuenta = (estado.archivo1Data || []).reduce((sum, row) => sum + (typeof row.TRANSACTION_NET_AMOUNT === 'number' ? row.TRANSACTION_NET_AMOUNT : limpiarMonto(row.TRANSACTION_NET_AMOUNT)), 0);
        let ordenesValidadas = 0;

        ordenes.forEach(orden => {
            const estatus = orden.estatus || 'REVISAR';
            if (!porEstatus[estatus]) {
                porEstatus[estatus] = {
                    cantidad: 0,
                    montoTotal: 0,
                    ordenes: []
                };
            }
            porEstatus[estatus].cantidad++;
            porEstatus[estatus].montoTotal += orden.neto?.estadoCuenta || 0;
            porEstatus[estatus].ordenes.push(orden);

            totalOrdenes++;
            
            // Una orden está validada si saldoValidado es true
            if (orden.validacion?.saldoValidado) {
                ordenesValidadas++;
            }
        });
        
        // 2. Mostrar estadísticas rápidas
        this.mostrarEstadisticas(totalOrdenes, montoTotalEdoCuenta, ordenesValidadas, porEstatus);
        
        // 3. Generar paneles por estatus
        this.generarPaneles(porEstatus);
        
        // 4. Generar resumen de acciones
        this.generarAcciones(porEstatus);
        
        // 5. Mostrar el panel
        document.getElementById('resumenEstatusPanel').style.display = 'block';
    },

    mostrarEstadisticas(total, monto, validadas, porEstatus) {
        const porcentajeValidadas = total > 0 ? Math.round((validadas / total) * 100) : 0;
        
        document.getElementById('totalOrdenesResumen').textContent = total;
        
        document.getElementById('statsResumen').innerHTML = `
            <div class="stat-card">
                <h4>Total Órdenes</h4>
                <div class="number">${total}</div>
            </div>
            <div class="stat-card">
                <h4>Monto Edo. Cuenta</h4>
                <div class="number">${Utilidades.formatMoney(monto)}</div>
            </div>
            <div class="stat-card">
                <h4>Órdenes Validadas</h4>
                <div class="number" style="color: ${porcentajeValidadas >= 90 ? 'var(--success)' : porcentajeValidadas >= 70 ? 'var(--warning)' : 'var(--danger)'}">
                    ${porcentajeValidadas}%
                </div>
                <div class="subtitle">${validadas} de ${total} órdenes</div>
            </div>
            <div class="stat-card">
                <h4>Estatus Diferentes</h4>
                <div class="number">${Object.keys(porEstatus).length}</div>
            </div>
        `;
    },

    generarPaneles(porEstatus) {
        const filtro = document.getElementById('filtroResumenEstatus').value;
        const orden = document.getElementById('ordenarResumen').value;
        
        // Convertir a array y ordenar
        let estatusArray = Object.entries(porEstatus);
        
        // Aplicar filtro
        if (filtro !== 'todos') {
            estatusArray = estatusArray.filter(([estatus]) => estatus === filtro);
        }
        
        // Aplicar orden
        estatusArray.sort((a, b) => {
            const [estatusA, dataA] = a;
            const [estatusB, dataB] = b;
            
            switch (orden) {
                case 'cantidad':
                    return dataB.cantidad - dataA.cantidad;
                case 'monto':
                    return Math.abs(dataB.montoTotal) - Math.abs(dataA.montoTotal);
                case 'estatus':
                    return estatusA.localeCompare(estatusB);
                default:
                    return 0;
            }
        });
        
        const container = document.getElementById('resumenEstatusContent');
        
        if (estatusArray.length === 0) {
            container.innerHTML = `
                <div style="text-align: center; padding: 40px; color: var(--gray-500);">
                    <p>No hay órdenes que coincidan con el filtro seleccionado.</p>
                </div>
            `;
            return;
        }
        
        container.innerHTML = estatusArray.map(([estatus, data]) => {
            const config = CONFIG_ESTATUS[estatus] || CONFIG_ESTATUS['REVISAR'];
            
            return `
                <div class="panel" style="margin-bottom: 16px;">
                    <div class="panel-header" style="background: linear-gradient(135deg, ${config.color}20 0%, ${config.color}10 100%); border-bottom: 2px solid ${config.color};">
                        <h3 style="display: flex; align-items: center; gap: 10px;">
                            <span class="estatus-badge ${config.badge}">${config.icon} ${estatus}</span>
                            <span style="font-size: 0.75rem; color: var(--gray-600); background: rgba(255,255,255,0.9); padding: 2px 8px; border-radius: 4px;">
                                ${data.cantidad} órdenes
                            </span>
                        </h3>
                        <span style="font-size: 0.875rem; font-weight: 600; color: ${config.color};">${Utilidades.formatMoney(data.montoTotal)}</span>
                    </div>
                    <div class="panel-body">
                        <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 16px;">
                            <div style="padding: 12px; background: ${config.color}10; border-radius: 6px;">
                                <div style="font-size: 0.75rem; color: ${config.color}; text-transform: uppercase; margin-bottom: 4px;">Total Órdenes</div>
                                <div style="font-size: 1.5rem; font-weight: 700; color: ${config.color};">${data.cantidad}</div>
                            </div>
                            <div style="padding: 12px; background: ${config.color}10; border-radius: 6px;">
                                <div style="font-size: 0.75rem; color: ${config.color}; text-transform: uppercase; margin-bottom: 4px;">Monto Total</div>
                                <div style="font-size: 1.5rem; font-weight: 700; color: ${config.color};">${Utilidades.formatMoney(data.montoTotal)}</div>
                            </div>
                        </div>
                        
                        <div style="padding: 12px; background: ${config.color}08; border-left: 4px solid ${config.color}; border-radius: 6px; margin-bottom: 16px;">
                            <strong style="color: ${config.color};">Acción recomendada:</strong>
                            <span style="color: ${config.color}; font-weight: 600;">${config.accion}</span>
                            <div style="font-size: 0.8125rem; color: ${config.color}; margin-top: 4px; opacity: 0.9;">
                                ${config.descripcion}
                            </div>
                        </div>
                        
                        <details style="margin-top: 12px;">
                            <summary style="cursor: pointer; font-size: 0.875rem; font-weight: 600; color: var(--gray-700); padding: 8px 0; display: flex; align-items: center; gap: 8px;">
                                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width: 16px; height: 16px;">
                                    <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                                </svg>
                                Ver ${data.cantidad} órdenes detalladas
                            </summary>
                            <div style="margin-top: 8px; max-height: 300px; overflow-y: auto; border: 1px solid var(--gray-200); border-radius: 6px;">
                                ${data.ordenes.map(orden => `
                                    <div style="padding: 8px 12px; border-bottom: 1px solid var(--gray-200); font-size: 0.8125rem; display: flex; justify-content: space-between; align-items: center;">
                                        <div>
                                            <strong style="color: var(--gray-900);">${orden.referenciaERP || 'Sin referencia'}</strong>
                                            <div style="font-size: 0.75rem; color: var(--gray-600);">
                                                ${orden.fechaLiberacion ? orden.fechaLiberacion.substring(0, 10) : 'Sin fecha'} •
                                                Neto: ${Utilidades.formatMoney(orden.neto?.real || 0)}
                                                ${orden.tieneMultiplesSourceIds ? ` • <span style="color: var(--primary); font-weight: 500;">${orden.sourceIds?.length || 0} SOURCE_IDs</span>` : ''}
                                            </div>
                                            <div style="font-size: 0.7rem; color: var(--gray-500);">
                                                Bruto: ${Utilidades.formatMoney(orden.ingresos?.bruto || 0)} | Gastos: ${Utilidades.formatMoney(orden.costos?.total || 0)} (Com: ${Utilidades.formatMoney(orden.costos?.comision || 0)}, Env: ${Utilidades.formatMoney(orden.costos?.envio || 0)}, Fin: ${Utilidades.formatMoney(orden.costos?.financiamiento || 0)}, Imp: ${Utilidades.formatMoney(orden.costos?.impuestos || 0)})
                                            </div>
                                        </div>
                                        <div style="display: flex; gap: 8px; align-items: center;">
                                            ${orden.tieneMultiplesSourceIds ?
                                                '<span style="background: #e0e7ff; color: #3730a3; padding: 2px 6px; border-radius: 3px; font-size: 0.65rem;" title="Orden con múltiples transacciones">MULTI</span>' :
                                                ''
                                            }
                                            ${orden.validacion?.saldoValidado ?
                                                '<span style="color: var(--success); font-size: 0.7rem;">Validado</span>' :
                                                `<span style="color: var(--danger); font-size: 0.7rem;">Δ ${Utilidades.formatMoney(Math.abs(orden.validacion?.diferencia || 0))}</span>`
                                            }
                                            ${orden.api?.encontrado ?
                                                '<span style="background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 3px; font-size: 0.7rem;">ERP</span>' :
                                                '<span style="background: #fef2f2; color: #991b1b; padding: 2px 6px; border-radius: 3px; font-size: 0.7rem;">Sin ERP</span>'
                                            }
                                        </div>
                                    </div>
                                `).join('')}
                            </div>
                        </details>
                    </div>
                </div>
            `;
        }).join('');
    },

    generarAcciones(porEstatus) {
        const acciones = {};
        let accionesUrgentes = 0;
        let accionesNormales = 0;
        let accionesRutina = 0;
        
        Object.entries(porEstatus).forEach(([estatus, data]) => {
            const config = CONFIG_ESTATUS[estatus] || CONFIG_ESTATUS['REVISAR'];
            
            if (!acciones[config.accion]) {
                acciones[config.accion] = {
                    prioridad: config.prioridad,
                    estatus: [],
                    cantidad: 0,
                    monto: 0
                };
            }
            
            acciones[config.accion].estatus.push(estatus);
            acciones[config.accion].cantidad += data.cantidad;
            acciones[config.accion].monto += data.montoTotal;
            
            if (config.prioridad === 'urgente') accionesUrgentes++;
            else if (config.prioridad === 'normal') accionesNormales++;
            else accionesRutina++;
        });
        
        const container = document.getElementById('resumenAcciones');
        
        // Ordenar por prioridad (urgente primero)
        const accionesArray = Object.entries(acciones).sort((a, b) => {
            const prioridadOrder = { urgente: 0, normal: 1, rutina: 2 };
            return prioridadOrder[a[1].prioridad] - prioridadOrder[b[1].prioridad];
        });
        
        container.innerHTML = `
            <div style="grid-column: 1 / -1; margin-bottom: 12px;">
                <div style="display: flex; gap: 16px; font-size: 0.75rem;">
                    <div><span style="background: #dc2626; color: white; padding: 2px 8px; border-radius: 3px;">${accionesUrgentes}</span> Urgentes</div>
                    <div><span style="background: #f59e0b; color: white; padding: 2px 8px; border-radius: 3px;">${accionesNormales}</span> Normales</div>
                    <div><span style="background: #10b981; color: white; padding: 2px 8px; border-radius: 3px;">${accionesRutina}</span> Rutina</div>
                </div>
            </div>
            
            ${accionesArray.map(([accion, data]) => `
                <div class="accion-card ${data.prioridad}" style="${data.prioridad === 'urgente' ? 'grid-column: 1 / -1;' : ''}">
                    <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px;">
                        <div>
                            <div style="font-weight: 600; color: var(--gray-900); margin-bottom: 4px;">${accion}</div>
                            <div style="font-size: 0.75rem; color: var(--gray-600);">
                                ${data.estatus.map(e => `<span class="estatus-badge ${CONFIG_ESTATUS[e]?.badge || 'revisar'}">${e}</span>`).join(' ')}
                            </div>
                        </div>
                        <div style="text-align: right;">
                            <div style="font-weight: 700; font-size: 1.125rem; color: var(--gray-900);">${data.cantidad}</div>
                            <div style="font-size: 0.7rem; color: var(--gray-500);">órdenes</div>
                        </div>
                    </div>
                    <div style="font-size: 0.75rem; color: var(--gray-700); margin-top: 8px;">
                        Monto total: <strong>${Utilidades.formatMoney(data.monto)}</strong>
                    </div>
                </div>
            `).join('')}
        `;
    },

    exportarResumen() {
        if (!estado.ordenesConsolidadas || estado.ordenesConsolidadas.length === 0) {
            Interfaz.mostrarToast('No hay datos para exportar', 'warning');
            return;
        }

        const datos = estado.ordenesConsolidadas.map(orden => ({
            'Referencia': orden.referenciaERP,
            'Estatus': orden.estatus,
            'Fecha Liberacion': orden.fechaLiberacion || '',
            'SOURCE_IDs': (orden.sourceIds || []).join(', '),
            'Ingreso Bruto': orden.ingresos?.bruto || 0,
            'Envio Comprador':  +(orden.costos?.shipping || 0),
            'Comision MP':      -Math.abs(orden.costos?.comision || 0),
            '% Comision':       (orden.ingresos?.bruto) ? Utilidades.roundMoney((orden.costos?.comision || 0) / orden.ingresos.bruto * 100) : 0,
            'Envio ML':         -Math.abs(orden.costos?.envio || 0),
            'Financiamiento':   -Math.abs(orden.costos?.financiamiento || 0),
            'Retención ML':     +(orden.ajustes?.retencion_ml || 0),
            'Ajuste ISR/IVA':   +(orden.ajustes?.ajuste_isr_iva || 0),
            'Env.Comp.Sep.':    +(orden.ajustes?.envio_comprador_sep || 0),
            'Envío Cte.':       orden.costos?.shipping || 0,
            'Total Gastos':     -Math.abs(orden.costos?.total || 0),
            'Subtotal':         Utilidades.roundMoney(((orden.ingresos?.bruto || 0) - (orden.costos?.shipping || 0)) / 1.16),
            'IVA 8%':           orden.costos?.iva  || 0,
            'ISR 2.5%':         orden.costos?.isr  || 0,
            'Neto Edo. Cuenta': orden.neto?.estadoCuenta || 0,
            'Diferencia': orden.validacion?.diferencia || 0,
            'Validado': orden.validacion?.saldoValidado ? 'SI' : 'NO',
            'En ERP': orden.api?.encontrado ? 'SI' : orden.api !== undefined ? 'NO' : 'N/A',
            'Estatus ERP': orden.api?.estatusERP || 'N/A',
            'Importe ERP': orden.api?.importeERP || 0
        }));

        const ws = XLSX.utils.json_to_sheet(datos);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Resumen Estatus');

        const fecha = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `resumen_estatus_${fecha}.xlsx`);

        Interfaz.mostrarToast('Resumen por estatus exportado correctamente', 'success');
    },

    inicializar() {
        // Eventos de filtros
        const filtroEstatus = document.getElementById('filtroResumenEstatus');
        const ordenarResumen = document.getElementById('ordenarResumen');
        
        if (filtroEstatus) {
            filtroEstatus.addEventListener('change', () => {
                if (estado.ordenesConsolidadas && estado.ordenesConsolidadas.length > 0) {
                    this.generarResumen(estado.ordenesConsolidadas);
                }
            });
        }
        
        if (ordenarResumen) {
            ordenarResumen.addEventListener('change', () => {
                if (estado.ordenesConsolidadas && estado.ordenesConsolidadas.length > 0) {
                    this.generarResumen(estado.ordenesConsolidadas);
                }
            });
        }
        
        // Botón de exportación
        const exportarBtn = document.getElementById('exportarResumenBtn');
        if (exportarBtn) {
            exportarBtn.addEventListener('click', () => this.exportarResumen());
        }
    }
};

// Inicializar ResumenEstatusManager cuando el DOM esté listo
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => ResumenEstatusManager.inicializar());
} else {
    ResumenEstatusManager.inicializar();
}

// ========== MÓDULO HISTÓRICO DE CONCILIACIONES ==========
const HistoricoManager = {
    conciliaciones: [],
    seleccionadas: [],
    detalleActual: null,

    async cargarHistorico() {
        try {
            const resp = await apiFetch(`${estado.backendUrl}/historico/conciliaciones?limit=50`);
            if (!resp.ok) throw new Error('Error cargando histórico');
            this.conciliaciones = await resp.json();
            this.renderizarLista();
        } catch (e) {
            console.warn('No se pudo cargar histórico:', e.message);
            document.getElementById('historico-lista').innerHTML =
                '<p style="color: var(--gray-500); text-align: center; padding: 40px;">No se pudo conectar con el servidor para cargar el histórico.</p>';
        }
    },

    renderizarLista() {
        const container = document.getElementById('historico-lista');
        if (!this.conciliaciones.length) {
            container.innerHTML = '<p style="color: var(--gray-500); text-align: center; padding: 40px;">No hay conciliaciones registradas aún.</p>';
            return;
        }

        const fmt = (n) => typeof n === 'number' ? '$' + n.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '$0.00';
        const fmtFecha = (iso) => {
            if (!iso) return '--';
            const d = new Date(iso);
            return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
        };

        container.innerHTML = this.conciliaciones.map(c => {
            const checked = this.seleccionadas.includes(c.id) ? 'checked' : '';
            const estatusBadges = Object.entries(c.resumenPorEstatus || {}).map(([est, data]) =>
                `<span class="badge badge-sm" style="background: var(--gray-100); color: var(--gray-700); font-size: 0.7rem;">${est}: ${data.cantidad}</span>`
            ).join('');

            return `<div class="historico-item">
                <div class="historico-item-header">
                    <input type="checkbox" class="historico-check" ${checked} onchange="HistoricoManager.toggleSeleccion(${c.id})">
                    <div style="flex: 1; min-width: 0;">
                        <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                            <strong style="font-size: 0.9rem;">${fmtFecha(c.fecha)}</strong>
                            <span class="badge badge-sm" style="background: var(--primary); color: white;">${c.totalOrdenes} órdenes</span>
                            <span class="badge badge-sm" style="background: var(--success); color: white;">${c.coincidencias} coincidencias</span>
                            ${c.casosCreados > 0 ? `<span class="badge badge-sm" style="background: var(--warning); color: white;">${c.casosCreados} casos</span>` : ''}
                        </div>
                        <div style="display: flex; gap: 12px; margin-top: 4px; font-size: 0.8rem; color: var(--gray-600);">
                            <span>Edo. Cuenta: ${fmt(c.montoEstadoCuenta)}</span>
                            <span>Neto: ${fmt(c.totalNetoReal)}</span>
                            <span>Validadas: ${c.ordenesValidadas}</span>
                        </div>
                        <div class="historico-estatus-badges">${estatusBadges}</div>
                    </div>
                    <div style="display: flex; gap: 6px;">
                        <button class="btn btn-sm btn-secondary" onclick="HistoricoManager.verDetalle(${c.id})">Ver Detalle</button>
                        <button class="btn btn-sm" style="background: var(--danger); color: white;" onclick="HistoricoManager.eliminar(${c.id})">Eliminar</button>
                    </div>
                </div>
            </div>`;
        }).join('');
    },

    toggleSeleccion(id) {
        const idx = this.seleccionadas.indexOf(id);
        if (idx >= 0) {
            this.seleccionadas.splice(idx, 1);
        } else {
            if (this.seleccionadas.length >= 2) {
                Interfaz.mostrarToast('Solo puedes seleccionar 2 conciliaciones para comparar', 'warning');
                this.renderizarLista();
                return;
            }
            this.seleccionadas.push(id);
        }
        const btn = document.getElementById('btn-comparar-historico');
        if (btn) {
            btn.textContent = `Comparar Seleccionadas (${this.seleccionadas.length}/2)`;
            btn.disabled = this.seleccionadas.length !== 2;
        }
        this.renderizarLista();
    },

    async verDetalle(id) {
        try {
            const resp = await apiFetch(`${estado.backendUrl}/historico/conciliaciones/${id}`);
            if (!resp.ok) throw new Error('Error cargando detalle');
            this.detalleActual = await resp.json();
            this.renderizarDetalle();
        } catch (e) {
            Interfaz.mostrarToast('Error cargando detalle: ' + e.message, 'error');
        }
    },

    renderizarDetalle() {
        const c = this.detalleActual;
        if (!c) return;

        const fmt = (n) => '$' + (n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const fmtFecha = (iso) => {
            if (!iso) return '--';
            const d = new Date(iso);
            return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
        };

        // Build estatus badges
        const estatusBadges = Object.entries(c.resumenPorEstatus || {}).map(([est, data]) =>
            `<div style="display: flex; justify-content: space-between; padding: 4px 8px; background: var(--gray-50); border-radius: 4px;">
                <span style="font-weight: 500;">${est}</span>
                <span>${data.cantidad} (${fmt(data.monto)})</span>
            </div>`
        ).join('');

        // Build orders table
        const ordenes = c.ordenes || [];
        const ordenesRows = ordenes.map(o => {
            const d = o.datosOrden || {};
            return `<tr>
                <td style="padding: 6px 8px;">${o.referenciaERP}</td>
                <td style="padding: 6px 8px;">${o.estatus}</td>
                <td style="padding: 6px 8px; text-align: right;">${fmt(d.ingresos?.bruto)}</td>
                <td style="padding: 6px 8px; text-align: right;">${fmt(d.costos?.comision)}</td>
                <td style="padding: 6px 8px; text-align: right;">${fmt(d.neto?.real)}</td>
                <td style="padding: 6px 8px; text-align: center;">${d.validacion?.saldoValidado ? '<span style="color: var(--success);">Si</span>' : '<span style="color: var(--gray-400);">No</span>'}</td>
            </tr>`;
        }).join('');

        const modal = document.getElementById('modal-historico-detalle');
        modal.innerHTML = `
            <div class="caso-detalle" onclick="event.stopPropagation()" style="max-width: 900px;">
                <div class="caso-detalle-header" style="background: var(--gradient-primary);">
                    <div class="caso-detalle-title" style="color: white;">
                        <h2 style="color: white; margin: 0;">Conciliación del ${fmtFecha(c.fecha)}</h2>
                    </div>
                    <button class="caso-detalle-close" onclick="document.getElementById('modal-historico-detalle').style.display='none'" style="color: white;">&times;</button>
                </div>
                <div class="caso-detalle-body" style="max-height: 70vh; overflow-y: auto;">
                    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px;">
                        <div style="background: var(--gray-50); padding: 12px; border-radius: 8px; text-align: center;">
                            <div style="font-size: 0.75rem; color: var(--gray-500);">Total Órdenes</div>
                            <div style="font-size: 1.25rem; font-weight: 700;">${c.totalOrdenes}</div>
                        </div>
                        <div style="background: var(--gray-50); padding: 12px; border-radius: 8px; text-align: center;">
                            <div style="font-size: 0.75rem; color: var(--gray-500);">Coincidencias</div>
                            <div style="font-size: 1.25rem; font-weight: 700; color: var(--success);">${c.coincidencias}</div>
                        </div>
                        <div style="background: var(--gray-50); padding: 12px; border-radius: 8px; text-align: center;">
                            <div style="font-size: 0.75rem; color: var(--gray-500);">Validadas</div>
                            <div style="font-size: 1.25rem; font-weight: 700; color: var(--primary);">${c.ordenesValidadas}</div>
                        </div>
                        <div style="background: var(--gray-50); padding: 12px; border-radius: 8px; text-align: center;">
                            <div style="font-size: 0.75rem; color: var(--gray-500);">Neto Real</div>
                            <div style="font-size: 1.1rem; font-weight: 700; color: var(--primary);">${fmt(c.totalNetoReal)}</div>
                        </div>
                    </div>
                    <div style="margin-bottom: 20px;">
                        <h4 style="margin: 0 0 8px 0; font-size: 0.875rem;">Resumen por Estatus</h4>
                        <div style="display: grid; gap: 4px;">${estatusBadges}</div>
                    </div>
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                        <h4 style="margin: 0; font-size: 0.875rem;">Órdenes (${ordenes.length})</h4>
                        <button class="btn btn-sm btn-secondary" onclick="HistoricoManager.exportarDetalle()">Exportar Excel</button>
                    </div>
                    <div style="overflow-x: auto; max-height: 400px; overflow-y: auto;">
                        <table style="width: 100%; border-collapse: collapse; font-size: 0.8rem;">
                            <thead>
                                <tr style="background: var(--gray-100); position: sticky; top: 0;">
                                    <th style="padding: 8px; text-align: left;">Referencia</th>
                                    <th style="padding: 8px; text-align: left;">Estatus</th>
                                    <th style="padding: 8px; text-align: right;">Ingreso Bruto</th>
                                    <th style="padding: 8px; text-align: right;">Comisión</th>
                                    <th style="padding: 8px; text-align: right;">Neto</th>
                                    <th style="padding: 8px; text-align: center;">Validada</th>
                                </tr>
                            </thead>
                            <tbody>${ordenesRows || '<tr><td colspan="6" style="padding: 20px; text-align: center; color: var(--gray-400);">Sin órdenes</td></tr>'}</tbody>
                        </table>
                    </div>
                </div>
            </div>`;
        modal.style.display = 'flex';
    },

    async eliminar(id) {
        if (!confirm('¿Eliminar esta conciliación del histórico? Esta acción no se puede deshacer.')) return;
        try {
            const resp = await apiFetch(`${estado.backendUrl}/historico/conciliaciones/${id}`, { method: 'DELETE' });
            if (!resp.ok) throw new Error('Error eliminando');
            this.seleccionadas = this.seleccionadas.filter(s => s !== id);
            Interfaz.mostrarToast('Conciliación eliminada', 'success');
            await this.cargarHistorico();
        } catch (e) {
            Interfaz.mostrarToast('Error eliminando: ' + e.message, 'error');
        }
    },

    async compararSeleccionadas() {
        if (this.seleccionadas.length !== 2) return;
        try {
            const resp = await apiFetch(`${estado.backendUrl}/historico/comparar?ids=${this.seleccionadas.join(',')}`);
            if (!resp.ok) throw new Error('Error comparando');
            const data = await resp.json();
            this.renderizarComparacion(data[0], data[1]);
        } catch (e) {
            Interfaz.mostrarToast('Error al comparar: ' + e.message, 'error');
        }
    },

    renderizarComparacion(a, b) {
        const container = document.getElementById('historico-comparacion');
        const fmt = (n) => '$' + (n || 0).toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const fmtFecha = (iso) => {
            if (!iso) return '--';
            const d = new Date(iso);
            return d.toLocaleDateString('es-MX', { day: '2-digit', month: 'short', year: 'numeric' });
        };
        const deltaClass = (val, invertido = false) => {
            if (val === 0) return '';
            const positivo = invertido ? val < 0 : val > 0;
            return positivo ? 'delta-positive' : 'delta-negative';
        };
        const fmtDelta = (val, isMoney = false) => {
            if (val === 0) return '-';
            const prefix = val > 0 ? '+' : '';
            return isMoney ? prefix + fmt(val) : prefix + val;
        };

        const rows = [
            { label: 'Fecha', valA: fmtFecha(a.fecha), valB: fmtFecha(b.fecha), delta: '' },
            { label: 'Total Órdenes', valA: a.totalOrdenes, valB: b.totalOrdenes, delta: b.totalOrdenes - a.totalOrdenes, cls: deltaClass(b.totalOrdenes - a.totalOrdenes) },
            { label: 'Coincidencias', valA: a.coincidencias, valB: b.coincidencias, delta: b.coincidencias - a.coincidencias, cls: deltaClass(b.coincidencias - a.coincidencias) },
            { label: 'Validadas', valA: a.ordenesValidadas, valB: b.ordenesValidadas, delta: b.ordenesValidadas - a.ordenesValidadas, cls: deltaClass(b.ordenesValidadas - a.ordenesValidadas) },
            { label: 'Casos Creados', valA: a.casosCreados, valB: b.casosCreados, delta: b.casosCreados - a.casosCreados, cls: deltaClass(b.casosCreados - a.casosCreados, true) },
            { label: 'Monto Edo. Cuenta', valA: fmt(a.montoEstadoCuenta), valB: fmt(b.montoEstadoCuenta), delta: fmtDelta(b.montoEstadoCuenta - a.montoEstadoCuenta, true), cls: deltaClass(b.montoEstadoCuenta - a.montoEstadoCuenta), isMoney: true },
            { label: 'Ingresos Bruto', valA: fmt(a.totalIngresosBruto), valB: fmt(b.totalIngresosBruto), delta: fmtDelta(b.totalIngresosBruto - a.totalIngresosBruto, true), cls: deltaClass(b.totalIngresosBruto - a.totalIngresosBruto), isMoney: true },
            { label: 'Total Costos', valA: fmt(a.totalCostos), valB: fmt(b.totalCostos), delta: fmtDelta(b.totalCostos - a.totalCostos, true), cls: deltaClass(b.totalCostos - a.totalCostos, true), isMoney: true },
            { label: 'Neto Real', valA: fmt(a.totalNetoReal), valB: fmt(b.totalNetoReal), delta: fmtDelta(b.totalNetoReal - a.totalNetoReal, true), cls: deltaClass(b.totalNetoReal - a.totalNetoReal), isMoney: true },
        ];

        // Estatus comparison
        const allEstatus = new Set([...Object.keys(a.resumenPorEstatus || {}), ...Object.keys(b.resumenPorEstatus || {})]);
        allEstatus.forEach(est => {
            const cantA = (a.resumenPorEstatus?.[est]?.cantidad) || 0;
            const cantB = (b.resumenPorEstatus?.[est]?.cantidad) || 0;
            rows.push({
                label: `${est} (cantidad)`,
                valA: cantA, valB: cantB,
                delta: cantB - cantA,
                cls: est === 'VENTA' ? deltaClass(cantB - cantA) : deltaClass(cantB - cantA, true)
            });
        });

        const tableRows = rows.map(r => {
            const deltaVal = typeof r.delta === 'number' ? fmtDelta(r.delta, r.isMoney) : r.delta;
            return `<tr>
                <td style="padding: 8px 12px; font-weight: 500;">${r.label}</td>
                <td style="padding: 8px 12px; text-align: center;">${r.valA}</td>
                <td style="padding: 8px 12px; text-align: center;">${r.valB}</td>
                <td style="padding: 8px 12px; text-align: center;" class="${r.cls || ''}">${deltaVal}</td>
            </tr>`;
        }).join('');

        container.innerHTML = `
            <div class="dashboard-section" style="margin-top: 16px;">
                <div class="dashboard-section-header">
                    <h3>Comparación de Conciliaciones</h3>
                    <button class="btn btn-sm btn-secondary" onclick="document.getElementById('historico-comparacion').innerHTML=''">Cerrar</button>
                </div>
                <div class="dashboard-section-body" style="padding: 0;">
                    <table style="width: 100%; border-collapse: collapse; font-size: 0.85rem;">
                        <thead>
                            <tr style="background: var(--gray-100);">
                                <th style="padding: 10px 12px; text-align: left;">Métrica</th>
                                <th style="padding: 10px 12px; text-align: center;">Conciliación A</th>
                                <th style="padding: 10px 12px; text-align: center;">Conciliación B</th>
                                <th style="padding: 10px 12px; text-align: center;">Delta</th>
                            </tr>
                        </thead>
                        <tbody>${tableRows}</tbody>
                    </table>
                </div>
            </div>`;
        container.scrollIntoView({ behavior: 'smooth' });
    },

    exportarDetalle() {
        const c = this.detalleActual;
        if (!c || !c.ordenes?.length) {
            Interfaz.mostrarToast('No hay órdenes para exportar', 'warning');
            return;
        }

        const data = c.ordenes.map(o => {
            const d = o.datosOrden || {};
            return {
                'Referencia ERP': o.referenciaERP,
                'Estatus': o.estatus,
                'Ingreso Bruto': d.ingresos?.bruto || 0,
                'Comisión': d.costos?.comision || 0,
                'Costo Envío': d.costos?.envio || 0,
                'Total Costos': d.costos?.total || 0,
                'Neto Real': d.neto?.real || 0,
                'Neto Edo. Cuenta': d.neto?.estadoCuenta || 0,
                'Validada': d.validacion?.saldoValidado ? 'Sí' : 'No',
                'Diferencia': d.validacion?.diferencia || 0
            };
        });

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Detalle Conciliación');
        const fecha = c.fecha ? c.fecha.split('T')[0] : 'sin-fecha';
        XLSX.writeFile(wb, `historico_conciliacion_${fecha}.xlsx`);
        Interfaz.mostrarToast('Detalle exportado', 'success');
    }
};

// ========== MÓDULO ANÁLISIS DE VALIDACIÓN ==========
const AnalisisValidacionManager = {
    ordenSeleccionada: null,

    inicializar() {
        const filtro = document.getElementById('filtroValidacion');
        if (filtro) {
            filtro.addEventListener('change', () => this.renderizarTabla());
        }
    },

    generarAnalisis() {
        if (!estado.ordenesConsolidadas || estado.ordenesConsolidadas.length === 0) {
            return;
        }

        this.actualizarResumen();
        this.renderizarTabla();
    },

    actualizarResumen() {
        const ordenes = estado.ordenesConsolidadas;
        const total = ordenes.length;
        const validadas = ordenes.filter(o => o.validacion?.saldoValidado).length;
        const conDiferencias = total - validadas;
        const porcentaje = total > 0 ? Math.round((validadas / total) * 100) : 0;
        const montoEdoCuenta = (estado.archivo1Data || []).reduce((sum, row) => sum + (typeof row.TRANSACTION_NET_AMOUNT === 'number' ? row.TRANSACTION_NET_AMOUNT : limpiarMonto(row.TRANSACTION_NET_AMOUNT)), 0);

        document.getElementById('val-total').textContent = total;
        document.getElementById('val-validadas').textContent = validadas;
        document.getElementById('val-diferencias').textContent = conDiferencias;
        document.getElementById('val-monto-edocuenta').textContent = Utilidades.formatMoney(montoEdoCuenta);

        const porcentajeEl = document.getElementById('val-porcentaje');
        porcentajeEl.textContent = `${porcentaje}%`;
        porcentajeEl.style.color = porcentaje >= 90 ? 'var(--success)' :
                                    porcentaje >= 70 ? 'var(--warning)' : 'var(--danger)';
    },

    renderizarTabla() {
        const filtro = document.getElementById('filtroValidacion')?.value || 'todos';
        let ordenes = [...(estado.ordenesConsolidadas || [])];

        // Aplicar filtro
        switch (filtro) {
            case 'validadas':
                ordenes = ordenes.filter(o => o.validacion?.saldoValidado);
                break;
            case 'diferencias':
                ordenes = ordenes.filter(o => !o.validacion?.saldoValidado);
                break;
            case 'multiples':
                ordenes = ordenes.filter(o => o.tieneMultiplesSourceIds);
                break;
            case 'sin_api':
                ordenes = ordenes.filter(o => o.apiData && !o.apiData.encontrado);
                break;
        }

        // Ordenar: primero las que tienen diferencias, luego por monto de diferencia
        ordenes.sort((a, b) => {
            const aValidado = a.validacion?.saldoValidado ? 1 : 0;
            const bValidado = b.validacion?.saldoValidado ? 1 : 0;
            if (aValidado !== bValidado) return aValidado - bValidado;
            return Math.abs(b.validacion?.diferencia || 0) - Math.abs(a.validacion?.diferencia || 0);
        });

        const tbody = document.getElementById('validacionTablaBody');

        if (ordenes.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="16" style="padding: 40px; text-align: center; color: var(--gray-500);">
                        ${estado.ordenesConsolidadas?.length > 0
                            ? 'No hay órdenes que coincidan con el filtro seleccionado.'
                            : 'Procesa una conciliación para ver el análisis de validación.'}
                    </td>
                </tr>
            `;
            return;
        }

        tbody.innerHTML = ordenes.map(orden => {
            const validado = orden.validacion?.saldoValidado;
            const diferencia = orden.validacion?.diferencia || 0;
            const netoCalculado = orden.neto?.real || 0;
            const netoEsperado = orden.neto?.estadoCuenta || 0;
            const sourceIds = orden.sourceIds || [];
            const tieneMultiples = orden.tieneMultiplesSourceIds;

            let estadoHTML = '';
            let rowClass = '';

            if (orden.apiData && !orden.apiData.encontrado && orden.estatus !== 'TRANSFERENCIA' && orden.estatus !== 'FISCAL') {
                const config = CONFIG_ESTATUS[orden.estatus] || CONFIG_ESTATUS['REVISAR'];
                estadoHTML = `<span class="validacion-estado" style="background: ${config.color}15; color: ${config.color}; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem;">${orden.estatus} (Sin API)</span>`;
                rowClass = 'con-diferencia';
            } else if (netoEsperado === 0 && sourceIds.length > 0) {
                estadoHTML = `<span class="validacion-estado sin-estado-cuenta">Sin E. Cuenta</span>`;
                rowClass = 'con-diferencia';
            } else if (validado) {
                estadoHTML = `<span class="validacion-estado validado">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width: 12px; height: 12px;">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7" />
                    </svg>
                    Validado
                </span>`;
                rowClass = 'validada';
            } else {
                estadoHTML = `<span class="validacion-estado diferencia">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2" style="width: 12px; height: 12px;">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                    </svg>
                    Diferencia
                </span>`;
                rowClass = 'con-diferencia';
            }

            const sourceIdsHTML = sourceIds.length <= 3
                ? sourceIds.map(id => `<span class="validacion-sourceid ${tieneMultiples ? 'multiple' : ''}">${id}</span>`).join('')
                : `<span class="validacion-sourceid ${tieneMultiples ? 'multiple' : ''}">${sourceIds.length} IDs</span>`;

            return `
                <tr class="${rowClass}" data-ref="${orden.referenciaERP}">
                    <td>
                        <strong>${orden.referenciaERP}</strong>
                        ${tieneMultiples ? '<span style="background: #e0e7ff; color: #3730a3; padding: 1px 4px; border-radius: 2px; font-size: 0.65rem; margin-left: 4px;">MULTI</span>' : ''}
                        <div style="font-size: 0.7rem; color: var(--gray-500);">${orden.estatus}</div>
                        ${orden.statusML ? `<div style="font-size: 0.65rem; color: var(--gray-400); font-family: monospace;">${orden.statusML} / ${orden.statusDetailML}</div>` : ''}
                    </td>
                    <td>
                        <div class="validacion-sourceids">${sourceIdsHTML}</div>
                    </td>
                    <td style="text-align: right; font-family: monospace; color: #92400e;">
                        ${Utilidades.formatMoney(orden.ingresos?.bruto || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${Utilidades.formatMoney(orden.costos?.comision || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${(orden.ingresos?.bruto) ? ((orden.costos?.comision || 0) / orden.ingresos.bruto * 100).toFixed(2) + '%' : '--'}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${Utilidades.formatMoney(orden.costos?.envio || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: #92400e; font-size: 0.75rem;">
                        ${(orden.costos?.financiamiento || 0) > 0 ? Utilidades.formatMoney(orden.costos.financiamiento) : '--'}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--danger); font-size: 0.75rem;">
                        ${(orden.ajustes?.retencion_ml || 0) !== 0 ? Utilidades.formatMoney(orden.ajustes.retencion_ml) : '--'}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--danger); font-size: 0.75rem;">
                        ${(orden.ajustes?.ajuste_isr_iva || 0) !== 0 ? Utilidades.formatMoney(orden.ajustes.ajuste_isr_iva) : '--'}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--success); font-size: 0.75rem;">
                        ${(orden.ajustes?.envio_comprador_sep || 0) > 0 ? '+ ' + Utilidades.formatMoney(orden.ajustes.envio_comprador_sep) : '--'}
                    </td>
                    <td style="text-align: right; font-family: monospace; font-size: 0.75rem; color: ${(orden.costos?.shipping || 0) > 0 ? 'var(--success)' : 'var(--gray-400)'};">
                        ${(orden.costos?.shipping || 0) > 0 ? '+ ' : ''}${Utilidades.formatMoney(orden.costos?.shipping || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; font-weight: 600; color: var(--danger);">
                        ${Utilidades.formatMoney(orden.costos?.total || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${Utilidades.formatMoney(((orden.ingresos?.bruto || 0) - (orden.costos?.shipping || 0)) / 1.16)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${Utilidades.formatMoney(orden.costos?.iva || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">
                        ${Utilidades.formatMoney(orden.costos?.isr || 0)}
                    </td>
                    <td style="text-align: right; font-family: monospace;">
                        ${Utilidades.formatMoney(netoEsperado)}
                    </td>
                    <td style="text-align: right; font-family: monospace; color: ${validado ? 'var(--success)' : 'var(--danger)'}; font-weight: 600;">
                        ${diferencia >= 0 ? '' : '-'}${Utilidades.formatMoney(Math.abs(diferencia))}
                    </td>
                    <td style="text-align: center;">
                        ${estadoHTML}
                    </td>
                    <td style="text-align: center;">
                        <button class="validacion-btn-detalle" onclick="AnalisisValidacionManager.verDetalle('${orden.referenciaERP}')">
                            Ver detalle
                        </button>
                    </td>
                </tr>
            `;
        }).join('');
    },

    verDetalle(referenciaERP) {
        const orden = estado.ordenesConsolidadas.find(o => o.referenciaERP === referenciaERP);
        if (!orden) return;

        this.ordenSeleccionada = orden;

        document.getElementById('validacionDetalleTitle').textContent = `Detalle: ${referenciaERP}`;

        const netoCalculado = orden.neto?.real || 0;
        const netoEsperado = orden.neto?.estadoCuenta || 0;
        const diferencia = orden.validacion?.diferencia || 0;
        const validado = orden.validacion?.saldoValidado;
        const estadosCuenta = orden.validacion?.estadosCuentaDetalle || [];

        // Calcular totales de créditos y débitos desde transacciones
        let totalCreditos = 0;
        let totalDebitos = 0;
        const transacciones = orden.transacciones || [];

        transacciones.forEach(t => {
            totalCreditos += t.netCredit || 0;
            totalDebitos += t.netDebit || 0;
        });

        const detalleHTML = `
            <div class="validacion-detalle-grid">
                <!-- Columna izquierda: Movimientos -->
                <div class="validacion-detalle-section">
                    <h4>
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                        </svg>
                        Movimientos (${transacciones.length})
                    </h4>
                    <div class="validacion-movimientos-lista">
                        ${transacciones.length > 0 ? transacciones.map(t => `
                            <div class="validacion-movimiento-item">
                                <div>
                                    <div class="descripcion">${t.description || 'Sin descripción'}</div>
                                    <div style="font-size: 0.65rem; color: var(--gray-500);">${t.sourceId}</div>
                                </div>
                                <div>
                                    ${t.netCredit > 0 ? `<span class="credito">+${Utilidades.formatMoney(t.netCredit)}</span>` : ''}
                                    ${t.netDebit > 0 ? `<span class="debito">-${Utilidades.formatMoney(t.netDebit)}</span>` : ''}
                                </div>
                            </div>
                        `).join('') : '<p style="color: var(--gray-500); text-align: center; padding: 16px;">Sin movimientos</p>'}
                    </div>
                    <table class="validacion-detalle-table" style="margin-top: 12px; border-top: 2px solid var(--gray-300); padding-top: 8px;">
                        <tr>
                            <td>Total Créditos</td>
                            <td style="color: var(--success);">+${Utilidades.formatMoney(totalCreditos)}</td>
                        </tr>
                        <tr>
                            <td>Total Débitos</td>
                            <td style="color: var(--danger);">-${Utilidades.formatMoney(totalDebitos)}</td>
                        </tr>
                        <tr style="font-weight: 700; font-size: 0.9rem;">
                            <td>Neto API (ML)</td>
                            <td>${Utilidades.formatMoney(netoCalculado)}</td>
                        </tr>
                    </table>
                </div>

                <!-- Columna derecha: Estado de Cuenta -->
                <div class="validacion-detalle-section">
                    <h4>
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M9 14l6-6m-5.5.5h.01m4.99 5h.01M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16l3.5-2 3.5 2 3.5-2 3.5 2z" />
                        </svg>
                        Estado de Cuenta (${estadosCuenta.length} SOURCE_IDs)
                    </h4>
                    <div class="validacion-movimientos-lista">
                        ${estadosCuenta.length > 0 ? estadosCuenta.map(ec => `
                            <div class="validacion-movimiento-item">
                                <div>
                                    <div class="descripcion">SOURCE_ID: ${ec.sourceId}</div>
                                </div>
                                <div>
                                    <span style="font-weight: 600;">${Utilidades.formatMoney(ec.neto)}</span>
                                </div>
                            </div>
                        `).join('') : '<p style="color: var(--gray-500); text-align: center; padding: 16px;">Sin registros en estado de cuenta</p>'}
                    </div>
                    <table class="validacion-detalle-table" style="margin-top: 12px; border-top: 2px solid var(--gray-300); padding-top: 8px;">
                        <tr style="font-weight: 700; font-size: 0.9rem;">
                            <td>Neto Edo. Cuenta (Suma)</td>
                            <td>${Utilidades.formatMoney(netoEsperado)}</td>
                        </tr>
                    </table>
                </div>
            </div>

            <!-- Resultado de validación -->
            <div style="background: ${validado ? 'var(--success-light)' : 'var(--danger-light)'}; padding: 16px; border-radius: 8px; margin-top: 16px;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <h4 style="margin: 0 0 8px 0; color: ${validado ? 'var(--success-dark)' : 'var(--danger-dark)'};">
                            ${validado ? 'Orden Validada' : 'Orden con Diferencia'}
                        </h4>
                        <p style="margin: 0; font-size: 0.875rem; color: ${validado ? 'var(--success-dark)' : 'var(--danger-dark)'};">
                            ${validado
                                ? `La diferencia de ${Utilidades.formatMoney(Math.abs(diferencia))} está dentro de la tolerancia ($${CONFIG.toleranciaNeto}).`
                                : `Existe una diferencia de ${Utilidades.formatMoney(Math.abs(diferencia))} que excede la tolerancia de $${CONFIG.toleranciaNeto}.`
                            }
                        </p>
                    </div>
                    <div style="text-align: right;">
                        <div style="font-size: 0.75rem; color: ${validado ? 'var(--success-dark)' : 'var(--danger-dark)'};">Diferencia</div>
                        <div style="font-size: 1.5rem; font-weight: 700; color: ${validado ? 'var(--success-dark)' : 'var(--danger-dark)'};">
                            ${diferencia >= 0 ? '+' : ''}${Utilidades.formatMoney(diferencia)}
                        </div>
                    </div>
                </div>
            </div>

            <!-- Fórmula aplicada -->
            <div style="background: var(--gray-100); padding: 12px; border-radius: 6px; margin-top: 12px; font-family: monospace; font-size: 0.8125rem;">
                <strong>Fórmula:</strong> Diferencia = Neto API (ML) - Neto Edo. Cuenta<br>
                <strong>Cálculo:</strong> ${Utilidades.formatMoney(diferencia)} = ${Utilidades.formatMoney(netoCalculado)} - ${Utilidades.formatMoney(netoEsperado)}<br>
                <strong>Criterio:</strong> |${Utilidades.formatMoney(Math.abs(diferencia))}| ${validado ? '<' : '>='} $${CONFIG.toleranciaNeto} → ${validado ? 'VALIDADO' : 'CON DIFERENCIA'}
                ${orden.statusML ? `<br><strong>Status ML:</strong> ${orden.statusML} / ${orden.statusDetailML}` : ''}
            </div>
        `;

        document.getElementById('validacionDetalleBody').innerHTML = detalleHTML;
        document.getElementById('validacionDetalle').style.display = 'block';

        // Scroll al detalle
        document.getElementById('validacionDetalle').scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    cerrarDetalle() {
        document.getElementById('validacionDetalle').style.display = 'none';
        this.ordenSeleccionada = null;
    },

    exportar() {
        if (!estado.ordenesConsolidadas || estado.ordenesConsolidadas.length === 0) {
            Interfaz.mostrarToast('No hay datos para exportar', 'warning');
            return;
        }

        const datos = estado.ordenesConsolidadas.map(orden => ({
            'Referencia': orden.referenciaERP,
            'Estatus': orden.estatus,
            'Status ML': orden.statusML || '',
            'Status Detail ML': orden.statusDetailML || '',
            'SOURCE_IDs': (orden.sourceIds || []).join(', '),
            'Cantidad SOURCE_IDs': orden.sourceIds?.length || 0,
            'Ingreso Bruto': orden.ingresos?.bruto || 0,
            'Envio Comprador':  +(orden.costos?.shipping || 0),
            'Comision MP':      -Math.abs(orden.costos?.comision || 0),
            '% Comision':       (orden.ingresos?.bruto) ? Utilidades.roundMoney((orden.costos?.comision || 0) / orden.ingresos.bruto * 100) : 0,
            'Envio ML':         -Math.abs(orden.costos?.envio || 0),
            'Financiamiento':   -Math.abs(orden.costos?.financiamiento || 0),
            'Retención ML':     +(orden.ajustes?.retencion_ml || 0),
            'Ajuste ISR/IVA':   +(orden.ajustes?.ajuste_isr_iva || 0),
            'Env.Comp.Sep.':    +(orden.ajustes?.envio_comprador_sep || 0),
            'Envío Cte.':       orden.costos?.shipping || 0,
            'Total Gastos':     -Math.abs(orden.costos?.total || 0),
            'Subtotal':         Utilidades.roundMoney(((orden.ingresos?.bruto || 0) - (orden.costos?.shipping || 0)) / 1.16),
            'IVA 8%':           orden.costos?.iva  || 0,
            'ISR 2.5%':         orden.costos?.isr  || 0,
            'Neto Edo. Cuenta': orden.neto?.estadoCuenta || 0,
            'Diferencia': orden.validacion?.diferencia || 0,
            'Validado': orden.validacion?.saldoValidado ? 'SI' : 'NO',
            'Tiene Múltiples': orden.tieneMultiplesSourceIds ? 'SI' : 'NO',
            'Fecha Liberación': orden.fechaLiberacion || ''
        }));

        const ws = XLSX.utils.json_to_sheet(datos);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Análisis Validación');

        const fecha = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `analisis_validacion_${fecha}.xlsx`);

        Interfaz.mostrarToast('Análisis exportado correctamente', 'success');
    }
};

// Inicializar AnalisisValidacionManager
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => AnalisisValidacionManager.inicializar());
} else {
    AnalisisValidacionManager.inicializar();
}

// ========== MÓDULO ANÁLISIS ERP (INTELISIS) ==========
const AnalisisERPManager = {
    ordenSeleccionada: null,

    inicializar() {
        const filtro = document.getElementById('filtroERP');
        if (filtro) {
            filtro.addEventListener('change', () => this.renderizarTabla());
        }
    },

    tieneDataERP() {
        return estado.ordenesConsolidadas?.some(o => o.api !== undefined);
    },

    generarAnalisis() {
        if (!estado.ordenesConsolidadas || estado.ordenesConsolidadas.length === 0) {
            this.mostrarSinDatos();
            return;
        }

        if (!this.tieneDataERP()) {
            this.mostrarSinDatos();
            return;
        }

        document.getElementById('erpNoData').style.display = 'none';
        document.getElementById('erpAnalisisContent').style.display = 'block';

        this.actualizarResumen();
        this.renderizarTabla();
    },

    mostrarSinDatos() {
        document.getElementById('erpNoData').style.display = 'block';
        document.getElementById('erpAnalisisContent').style.display = 'none';
    },

    actualizarResumen() {
        const ordenes = estado.ordenesConsolidadas.filter(o => o.api !== undefined);
        const total = ordenes.length;
        const encontradas = ordenes.filter(o => o.api?.encontrado).length;
        const noEncontradas = total - encontradas;

        const ordenesEncontradas = ordenes.filter(o => o.api?.encontrado);
        const conDiferencias = ordenesEncontradas.filter(o =>
            o.api?.diferencia && Math.abs(o.api.diferencia) > 0.01
        ).length;

        const porcentaje = total > 0 ? Math.round((encontradas / total) * 100) : 0;

        let totalML = 0;
        let totalERP = 0;
        ordenesEncontradas.forEach(o => {
            totalML += o.ingresos?.bruto || 0;
            totalERP += o.api?.importeERP || 0;
        });
        const diferenciaMonto = totalML - totalERP;

        document.getElementById('erp-total').textContent = total;
        document.getElementById('erp-encontradas').textContent = encontradas;
        document.getElementById('erp-no-encontradas').textContent = noEncontradas;
        document.getElementById('erp-diferencias').textContent = conDiferencias;

        const porcentajeEl = document.getElementById('erp-porcentaje');
        porcentajeEl.textContent = `${porcentaje}%`;
        porcentajeEl.style.color = porcentaje >= 90 ? 'var(--success)' :
                                    porcentaje >= 70 ? 'var(--warning)' : 'var(--danger)';

        document.getElementById('erp-monto-ml').textContent = Utilidades.formatMoney(totalML);
        document.getElementById('erp-monto-erp').textContent = Utilidades.formatMoney(totalERP);
        document.getElementById('erp-monto-diferencia').textContent =
            (diferenciaMonto >= 0 ? '+' : '') + Utilidades.formatMoney(diferenciaMonto);

        // Resumen por Estatus con Validaciones ERP
        const porEstatus = {};
        ordenes.forEach(o => {
            const est = o.estatus || 'REVISAR';
            if (!porEstatus[est]) {
                porEstatus[est] = { total: 0, encontradas: 0, coinciden: 0, diferencias: 0, noEncontradas: 0, monto: 0 };
            }
            porEstatus[est].total++;
            porEstatus[est].monto += o.neto?.estadoCuenta || 0;
            if (o.api?.encontrado) {
                porEstatus[est].encontradas++;
                if (o.api.diferencia && Math.abs(o.api.diferencia) > 0.01) {
                    porEstatus[est].diferencias++;
                } else {
                    porEstatus[est].coinciden++;
                }
            } else {
                porEstatus[est].noEncontradas++;
            }
        });

        this.generarCuadre();

        const resumenContainer = document.getElementById('erpResumenEstatusBody');
        const resumenPanel = document.getElementById('erpResumenEstatus');
        if (resumenContainer && Object.keys(porEstatus).length > 0) {
            resumenPanel.style.display = 'block';
            resumenContainer.innerHTML = Object.entries(porEstatus).map(([estatus, data]) => {
                const config = CONFIG_ESTATUS[estatus] || CONFIG_ESTATUS['REVISAR'];
                const pctCoincide = data.total > 0 ? Math.round((data.coinciden / data.total) * 100) : 0;
                return `
                    <div style="background: ${config.color}10; border: 1px solid ${config.color}30; border-radius: 8px; padding: 12px;">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                            <span class="estatus-badge ${config.badge}">${config.icon} ${estatus}</span>
                            <span style="font-weight: 700; font-size: 1.1rem;">${data.total}</span>
                        </div>
                        <div style="font-size: 0.75rem; color: var(--gray-600); line-height: 1.6;">
                            <div style="display: flex; justify-content: space-between;">
                                <span>Coinciden:</span>
                                <span style="color: var(--success); font-weight: 600;">${data.coinciden}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between;">
                                <span>Diferencias:</span>
                                <span style="color: var(--warning); font-weight: 600;">${data.diferencias}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between;">
                                <span>No encontradas:</span>
                                <span style="color: var(--danger); font-weight: 600;">${data.noEncontradas}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between; margin-top: 4px; padding-top: 4px; border-top: 1px solid ${config.color}20;">
                                <span>Monto:</span>
                                <span style="font-weight: 600;">${Utilidades.formatMoney(data.monto)}</span>
                            </div>
                            <div style="display: flex; justify-content: space-between;">
                                <span>% Coincidencia:</span>
                                <span style="font-weight: 600; color: ${pctCoincide >= 90 ? 'var(--success)' : pctCoincide >= 70 ? 'var(--warning)' : 'var(--danger)'};">${pctCoincide}%</span>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }
    },

    generarCuadre() {
        const panel = document.getElementById('erpCuadreContable');
        const body  = document.getElementById('erpCuadreBody');
        if (!panel || !body) return;

        const todas = estado.ordenesConsolidadas || [];
        if (todas.length === 0) { panel.style.display = 'none'; return; }

        const R = v => Utilidades.roundMoney(v);
        const F = v => Utilidades.formatMoney(v);

        // --- VENTA: único ingreso real ---
        const ventas          = todas.filter(o => o.estatus === 'VENTA');
        const importeERPventa = R(ventas.filter(o => o.api?.encontrado).reduce((s, o) => s + (o.api?.importeERP || 0), 0));
        const comisionVenta   = R(ventas.reduce((s, o) => s + (o.costos?.comision || 0), 0));
        const envioVenta      = R(ventas.reduce((s, o) => s + (o.costos?.envio    || 0), 0));
        const ivaVenta        = R(ventas.reduce((s, o) => s + (o.costos?.iva      || 0), 0));
        const isrVenta        = R(ventas.reduce((s, o) => s + (o.costos?.isr      || 0), 0));
        const ajusteVenta     = R(ventas.reduce((s, o) => s + (o.ajustes?.ajuste_isr_iva || 0), 0));

        // --- Otros estatus: agrupados por estatus, usando neto.estadoCuenta ---
        const otrosGrupos = {};
        todas.filter(o => o.estatus !== 'VENTA').forEach(o => {
            const est = o.estatus || 'REVISAR';
            if (!otrosGrupos[est]) otrosGrupos[est] = { edoCuenta: 0 };
            const g = otrosGrupos[est];
            g.edoCuenta = R(g.edoCuenta + (o.neto?.estadoCuenta || 0));
        });

        // --- Total Edo. Cuenta (todos los estatus) ---
        const totalEdoCuenta = R(todas.reduce((s, o) => s + (o.neto?.estadoCuenta || 0), 0));

        // --- Resultado del cuadre ---
        // Importe ERP VENTA - costos VENTA ± cada otro estatus (edoCuenta - sus costos)
        let cuadreTotal = R(importeERPventa - comisionVenta - envioVenta - ivaVenta - isrVenta + ajusteVenta);
        Object.values(otrosGrupos).forEach(g => {
            if (g.edoCuenta !== 0) {
                cuadreTotal = R(cuadreTotal + g.edoCuenta);
            }
        });

        const diferencia = R(cuadreTotal - totalEdoCuenta);
        const cuadra     = Math.abs(diferencia) <= 0.5;

        // --- Desglose clasificado de la diferencia ---
        // La diferencia = (neto efectivo ERP de VENTAS) - (neto banco de VENTAS)
        // neto efectivo ERP = importeERP - comision - envio - iva - isr + ajuste
        // neto banco        = suma de neto.estadoCuenta de filas VENTA
        const ventasBancoNeto   = R(ventas.reduce((s, o) => s + (o.neto?.estadoCuenta || 0), 0));
        const ventasERPneto     = R(importeERPventa - comisionVenta - envioVenta - ivaVenta - isrVenta + ajusteVenta);

        // Agrupar diferencias VENTA por referencia ERP (una entrada por orden)
        const difPorReferencia = {};
        ventas.forEach(o => {
            const erpNet   = R((o.api?.importeERP || 0) - (o.costos?.comision || 0) - (o.costos?.envio || 0)
                               - (o.costos?.iva || 0) - (o.costos?.isr || 0) + (o.ajustes?.ajuste_isr_iva || 0));
            const bancoNet = o.neto?.estadoCuenta || 0;
            const difOrden = R(bancoNet - erpNet);
            if (Math.abs(difOrden) < 0.005) return;

            const pendiente = o.montoPendienteLiberar || 0;
            const ref = o.referenciaERP || o.paymentId;

            difPorReferencia[ref] = {
                monto:      difOrden,
                pendiente,
                esMSI:      pendiente > 0 && !!o.tienePagosMSI,
                maxCuotas:  o.msiMaxCuotas || 1,
                estatusERP: o.api?.encontrado
                    ? (o.api?.estatusERP || 'ERP s/estatus').toString().trim()
                    : 'NO_ENCONTRADO en ERP',
                encontrado: o.api?.encontrado || false,
                importeML:  R(o.ingresos?.bruto || 0),
                importeERP: o.api?.importeERP || 0,
            };
        });

        // Cashback / filas VENTA sin ERP que tienen importe en banco pero ERP = 0
        const difTotal = R(ventasBancoNeto - ventasERPneto); // debería coincidir con diferencia

        // --- Helpers HTML ---
        const row = (label, valor, cls = '', indent = false) => `
            <tr>
                <td style="padding: 7px 12px; color: var(--gray-600); ${indent ? 'padding-left:28px;' : 'font-weight:600;'}">${label}</td>
                <td style="padding: 7px 12px; text-align: right; font-family: monospace; ${cls}">${valor}</td>
            </tr>`;

        const rowSub = (label, valor, cls = '') => `
            <tr>
                <td style="padding: 4px 12px; padding-left:44px; color: var(--gray-500); font-size:0.75rem;">${label}</td>
                <td style="padding: 4px 12px; text-align: right; font-family: monospace; font-size:0.75rem; ${cls}">${valor}</td>
            </tr>`;

        const sep = label => `
            <tr style="background: var(--gray-100);">
                <td colspan="2" style="padding: 4px 12px; font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--gray-500); font-weight: 600;">${label}</td>
            </tr>`;

        const totalRow = (label, valor, cls = '') => `
            <tr style="border-top: 2px solid var(--gray-300);">
                <td style="padding: 9px 12px; font-weight: 700;">${label}</td>
                <td style="padding: 9px 12px; text-align: right; font-family: monospace; font-weight: 700; ${cls}">${valor}</td>
            </tr>`;

        // --- Filas de otros estatus ---
        const otrosRows = Object.entries(otrosGrupos)
            .filter(([, g]) => g.edoCuenta !== 0)
            .map(([est, g]) => {
                const cfg       = CONFIG_ESTATUS[est] || CONFIG_ESTATUS['REVISAR'];
                const esIngreso = g.edoCuenta > 0;
                return row(
                    `${cfg.icon} ${est}`,
                    (esIngreso ? '+ ' : '- ') + F(Math.abs(g.edoCuenta)),
                    `color:${esIngreso ? 'var(--success)' : 'var(--danger)'};`,
                    true
                );
            }).join('');

        body.innerHTML = `
            <table style="width: 100%; border-collapse: collapse; font-size: 0.8125rem; border: 1px solid var(--gray-200); border-radius: 8px; overflow: hidden;">
                <thead>
                    <tr style="background: var(--gray-700); color: white;">
                        <th style="padding: 10px 12px; text-align: left; font-weight: 600;">Concepto</th>
                        <th style="padding: 10px 12px; text-align: right; font-weight: 600;">Importe</th>
                    </tr>
                </thead>
                <tbody>
                    ${sep('Ingresos — Ventas ERP')}
                    ${row('(+) Importe ERP VENTA', '+ ' + F(importeERPventa), 'color:var(--success); font-weight:600;', true)}

                    ${sep('Egresos — Deducciones ML (VENTA)')}
                    ${row('(-) Comisión MP', '- ' + F(comisionVenta), 'color:var(--danger);', true)}
                    ${row('(-) Envío ML',    '- ' + F(envioVenta),    'color:var(--danger);', true)}
                    ${row('(-) IVA 8%',      '- ' + F(ivaVenta),      'color:var(--danger);', true)}
                    ${row('(-) ISR 2.5%',    '- ' + F(isrVenta),      'color:var(--danger);', true)}
                    ${ajusteVenta !== 0 ? row('(-) Ajuste ISR/IVA post-lib', (ajusteVenta >= 0 ? '+ ' : '- ') + F(Math.abs(ajusteVenta)), 'color:var(--danger);', true) : ''}

                    ${otrosRows ? sep('Otros estatus') + otrosRows : ''}

                    ${totalRow('= Resultado del Cuadre', F(cuadreTotal), 'color:var(--primary); font-size:1rem;')}
                    ${totalRow('(-) Total Importe Edo. Cuenta', '- ' + F(totalEdoCuenta), 'color:var(--gray-700);')}

                    <tr style="border-top: 2px solid ${cuadra ? '#16a34a' : '#dc2626'}; background: ${cuadra ? '#f0fdf4' : '#fef2f2'};">
                        <td style="padding: 10px 12px; font-weight: 700; color: ${cuadra ? '#15803d' : '#dc2626'};">
                            ${cuadra ? '✓ Cuadre correcto' : '⚠ Diferencia pendiente'}
                        </td>
                        <td style="padding: 10px 12px; text-align: right; font-family: monospace; font-weight: 700; font-size: 1rem; color: ${cuadra ? '#15803d' : '#dc2626'};">
                            ${diferencia >= 0 ? '+' : ''}${F(diferencia)}
                        </td>
                    </tr>
                    ${!cuadra ? `
                    ${(() => {
                        const rows = [];
                        // Encabezado del desglose
                        rows.push(`
                            <tr style="background:#fef9c3;">
                                <td colspan="2" style="padding:6px 12px; font-size:0.7rem; text-transform:uppercase; letter-spacing:0.05em; color:#92400e; font-weight:700;">
                                    📋 Desglose de la diferencia (VENTA: neto ERP vs. banco)
                                </td>
                            </tr>
                            <tr style="background:#fef9c3;">
                                <td style="padding:4px 12px; font-size:0.72rem; color:#78350f; padding-left:20px;">
                                    Neto efectivo ERP (Importe ERP − comisiones − impuestos)
                                </td>
                                <td style="padding:4px 12px; text-align:right; font-family:monospace; font-size:0.72rem; color:#78350f; font-weight:600;">
                                    ${F(ventasERPneto)}
                                </td>
                            </tr>
                            <tr style="background:#fef9c3;">
                                <td style="padding:4px 12px; font-size:0.72rem; color:#78350f; padding-left:20px;">
                                    Neto real banco (Importe Edo. Cuenta VENTA)
                                </td>
                                <td style="padding:4px 12px; text-align:right; font-family:monospace; font-size:0.72rem; color:#78350f; font-weight:600;">
                                    ${F(ventasBancoNeto)}
                                </td>
                            </tr>
                        `);

                        // Filas por referencia ERP (una fila por orden con discrepancia)
                        if (Object.keys(difPorReferencia).length > 0) {
                            rows.push(`
                                <tr style="background:#fef3c7;">
                                    <td colspan="2" style="padding:5px 12px; font-size:0.7rem; text-transform:uppercase; letter-spacing:0.04em; color:#92400e; font-weight:600; padding-left:20px;">
                                        Detalle por orden con discrepancia:
                                    </td>
                                </tr>
                            `);
                            Object.entries(difPorReferencia)
                                .sort((a, b) => Math.abs(b[1].monto) - Math.abs(a[1].monto))
                                .forEach(([ref, info]) => {
                                    const esNeg  = info.monto < 0;
                                    const badge  = info.encontrado
                                        ? `<span style="display:inline-block;margin-left:6px;padding:1px 6px;border-radius:9px;font-size:0.65rem;background:#e0e7ff;color:#3730a3;">${info.estatusERP}</span>`
                                        : `<span style="display:inline-block;margin-left:6px;padding:1px 6px;border-radius:9px;font-size:0.65rem;background:#fee2e2;color:#991b1b;">No encontrado</span>`;
                                    const msiTag = info.esMSI
                                        ? `<div style="font-size:0.67rem;color:#b45309;margin-top:1px;">📅 ${info.maxCuotas} MSI — $${F(info.pendiente)} pdte. de liberar</div>`
                                        : (info.pendiente > 0
                                            ? `<div style="font-size:0.67rem;color:#b45309;margin-top:1px;">⏳ $${F(info.pendiente)} pdte. de liberar</div>`
                                            : '');
                                    const detalle = `<div style="font-size:0.67rem;color:var(--gray-500);margin-top:1px;">ML ${F(info.importeML)} vs ERP ${F(info.importeERP)}</div>`;
                                    rows.push(`
                                        <tr style="background:#fffbeb;">
                                            <td style="padding:5px 12px 5px 32px; font-size:0.75rem; color:var(--gray-700);">
                                                <span style="font-weight:600;font-family:monospace;">${ref}</span>${badge}
                                                ${detalle}
                                                ${msiTag}
                                            </td>
                                            <td style="padding:5px 12px; text-align:right; font-family:monospace; font-size:0.75rem; font-weight:600; color:${esNeg ? '#dc2626' : '#16a34a'}; vertical-align:top;">
                                                ${info.monto >= 0 ? '+' : ''}${F(info.monto)}
                                            </td>
                                        </tr>
                                    `);
                                });
                        } else {
                            rows.push(`
                                <tr style="background:#fffbeb;">
                                    <td colspan="2" style="padding:5px 12px 5px 32px; font-size:0.72rem; color:var(--gray-500); font-style:italic;">
                                        Sin órdenes con diferencia individual identificada
                                    </td>
                                </tr>
                            `);
                        }

                        // Nota explicativa
                        rows.push(`
                            <tr style="background:#fef2f2;">
                                <td colspan="2" style="padding:6px 12px; font-size:0.68rem; color:#7f1d1d; font-style:italic; line-height:1.5;">
                                    💡 <strong>📅 Liberación pendiente:</strong> pagos a MSI o divididos donde ML aún no depositó el monto completo en este período.<br>
                                    Las diferencias negativas por estatus ERP indican que el ERP registra un neto mayor al depositado (ej. órdenes con ajustes de precio, CONCLUIDO+CANCELADO).<br>
                                    Las positivas indican depósitos en banco sin contrapartida en ERP (ej. cashback).
                                </td>
                            </tr>
                        `);

                        return rows.join('');
                    })()}
                    ` : ''}
                </tbody>
            </table>`;

        panel.style.display = 'block';
    },

    renderizarTabla() {
        const filtro = document.getElementById('filtroERP')?.value || 'todos';
        let ordenes = estado.ordenesConsolidadas.filter(o => o.api !== undefined);

        switch (filtro) {
            case 'encontradas':
                ordenes = ordenes.filter(o => o.api?.encontrado);
                break;
            case 'no-encontradas':
                ordenes = ordenes.filter(o => !o.api?.encontrado);
                break;
            case 'coinciden':
                ordenes = ordenes.filter(o =>
                    o.api?.encontrado && (!o.api?.diferencia || Math.abs(o.api.diferencia) <= 0.01)
                );
                break;
            case 'diferencias':
                ordenes = ordenes.filter(o =>
                    o.api?.encontrado && o.api?.diferencia && Math.abs(o.api.diferencia) > 0.01
                );
                break;
        }

        ordenes.sort((a, b) => {
            const aEncontrado = a.api?.encontrado ? 1 : 0;
            const bEncontrado = b.api?.encontrado ? 1 : 0;
            if (aEncontrado !== bEncontrado) return aEncontrado - bEncontrado;
            return Math.abs(b.api?.diferencia || 0) - Math.abs(a.api?.diferencia || 0);
        });

        const tbody = document.getElementById('erpTablaBody');

        if (ordenes.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="18" style="padding: 40px; text-align: center; color: var(--gray-500);">
                        No hay órdenes que coincidan con el filtro seleccionado.
                    </td>
                </tr>
            `;
            return;
        }

        tbody.innerHTML = ordenes.map(orden => {
            const encontrado = orden.api?.encontrado;
            const importeML = orden.ingresos?.bruto || 0;
            const importeERP = orden.api?.importeERP || 0;
            const diferencia = orden.api?.diferencia || 0;
            const estatusERP = orden.api?.estatusERP || 'N/A';
            const movimientos = orden.api?.movimientos || 'N/A';

            let estadoHTML = '';
            let rowClass = '';

            if (!encontrado) {
                estadoHTML = `<span class="erp-estado no-encontrado">No encontrado</span>`;
                rowClass = 'con-diferencia';
            } else if (Math.abs(diferencia) <= 0.01) {
                estadoHTML = `<span class="erp-estado coincide">Coincide</span>`;
                rowClass = 'validada';
            } else {
                estadoHTML = `<span class="erp-estado diferencia">Diferencia</span>`;
                rowClass = 'con-diferencia';
            }

            const erpSourceIds = orden.sourceIds || [];
            const erpTieneMultiples = orden.tieneMultiplesSourceIds;
            const erpSourceIdsHTML = erpSourceIds.length <= 3
                ? erpSourceIds.map(id => `<span class="validacion-sourceid ${erpTieneMultiples ? 'multiple' : ''}">${id}</span>`).join('')
                : `<span class="validacion-sourceid ${erpTieneMultiples ? 'multiple' : ''}">${erpSourceIds.length} IDs</span>`;

            return `
                <tr class="${rowClass}">
                    <td>
                        <strong>${orden.referenciaERP}</strong>
                        <div style="font-size: 0.7rem; color: var(--gray-500);">${orden.api?.cliente || ''}</div>
                    </td>
                    <td>
                        <div class="validacion-sourceids">${erpSourceIdsHTML}</div>
                    </td>
                    <td>
                        <span style="background: #fef3c7; color: #92400e; padding: 2px 6px; border-radius: 3px; font-size: 0.75rem;">${orden.estatus}</span>
                        ${orden.statusML ? `<div style="font-size: 0.65rem; color: var(--gray-400); font-family: monospace; margin-top: 2px;">${orden.statusML} / ${orden.statusDetailML}</div>` : ''}
                    </td>
                    <td>${encontrado ? `<span style="background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 3px; font-size: 0.75rem;">${estatusERP}</span><div style="font-size: 0.65rem; color: var(--gray-500);">Mov: ${movimientos}</div>` : '--'}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${Utilidades.formatMoney(orden.costos?.comision || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${importeML ? ((orden.costos?.comision || 0) / importeML * 100).toFixed(2) + '%' : '--'}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${Utilidades.formatMoney(orden.costos?.envio || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: #92400e; font-size: 0.75rem;">${(orden.costos?.financiamiento || 0) > 0 ? Utilidades.formatMoney(orden.costos.financiamiento) : '--'}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--danger); font-size: 0.75rem;">${(orden.ajustes?.retencion_ml || 0) !== 0 ? Utilidades.formatMoney(orden.ajustes.retencion_ml) : '--'}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--danger); font-size: 0.75rem;">${(orden.ajustes?.ajuste_isr_iva || 0) !== 0 ? Utilidades.formatMoney(orden.ajustes.ajuste_isr_iva) : '--'}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--success); font-size: 0.75rem;">${(orden.ajustes?.envio_comprador_sep || 0) > 0 ? '+ ' + Utilidades.formatMoney(orden.ajustes.envio_comprador_sep) : '--'}</td>
                    <td style="text-align: right; font-family: monospace; font-size: 0.75rem; color: ${(orden.costos?.shipping || 0) > 0 ? 'var(--success)' : 'var(--gray-400)'};">${(orden.costos?.shipping || 0) > 0 ? '+ ' : ''}${Utilidades.formatMoney(orden.costos?.shipping || 0)}</td>
                    <td style="text-align: right; font-family: monospace; font-weight: 600; color: var(--danger);">${Utilidades.formatMoney(orden.costos?.total || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${Utilidades.formatMoney(importeML / 1.16)}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${Utilidades.formatMoney(orden.costos?.iva || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--gray-600); font-size: 0.75rem;">${Utilidades.formatMoney(orden.costos?.isr || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: var(--primary);">${Utilidades.formatMoney(orden.neto?.estadoCuenta || 0)}</td>
                    <td style="text-align: right; font-family: monospace; color: #92400e;">${Utilidades.formatMoney(importeML + (orden.costos?.shipping || 0))}</td>
                    <td style="text-align: right; font-family: monospace; color: #1e40af;">${encontrado ? Utilidades.formatMoney(importeERP) : '--'}</td>
                    <td style="text-align: right; font-family: monospace; font-weight: 600; color: ${!encontrado ? 'var(--gray-400)' : Math.abs(diferencia) <= 0.01 ? 'var(--success)' : 'var(--danger)'};">${encontrado ? (diferencia >= 0 ? '+' : '') + Utilidades.formatMoney(diferencia) : '--'}</td>
                    <td style="text-align: center;">${estadoHTML}</td>
                    <td style="text-align: center;"><button class="validacion-btn-detalle" onclick="AnalisisERPManager.verDetalle('${orden.referenciaERP}')">Ver detalle</button></td>
                </tr>
            `;
        }).join('');
    },

    verDetalle(referenciaERP) {
        const orden = estado.ordenesConsolidadas.find(o => o.referenciaERP === referenciaERP);
        if (!orden) return;

        this.ordenSeleccionada = orden;
        document.getElementById('erpDetalleTitle').textContent = `Detalle ERP: ${referenciaERP}`;

        const encontrado = orden.api?.encontrado;
        const importeML = orden.ingresos?.bruto || 0;
        const importeERP = orden.api?.importeERP || 0;
        const diferencia = orden.api?.diferencia || 0;
        const movimientosDetalle = orden.api?.movimientosDetalle || [];
        const costos = orden.costos || {};

        const movimientosHTML = movimientosDetalle.length > 0
            ? movimientosDetalle.map(mov => `
                <div class="erp-movimiento-card">
                    <div class="mov-header">
                        <span class="mov-tipo">${mov.mov || 'Movimiento'}</span>
                        <span class="mov-id">${mov.movID || 'N/A'}</span>
                    </div>
                    <div class="mov-detalle">
                        <div class="mov-detalle-item"><span class="label">Estatus</span><span class="value">${mov.estatus || 'N/A'}</span></div>
                        <div class="mov-detalle-item"><span class="label">Importe</span><span class="value">${Utilidades.formatMoney(mov.importe || 0)}</span></div>
                        <div class="mov-detalle-item"><span class="label">Fecha</span><span class="value">${mov.fechaEmision ? mov.fechaEmision.substring(0, 10) : 'N/A'}</span></div>
                        <div class="mov-detalle-item"><span class="label">Situación</span><span class="value">${mov.situacion || 'N/A'}</span></div>
                    </div>
                </div>
            `).join('')
            : '<p style="color: var(--gray-500); text-align: center; padding: 20px;">Sin movimientos en ERP</p>';

        const detalleHTML = `
            <div class="erp-comparativa">
                <div class="erp-comparativa-col ml">
                    <h4 style="color: #92400e;">MercadoLibre</h4>
                    <table class="validacion-detalle-table">
                        <tr><td>Estatus</td><td><span style="background: #fef3c7; color: #92400e; padding: 2px 6px; border-radius: 3px;">${orden.estatus}</span></td></tr>
                        ${orden.statusML ? `<tr><td>Status ML</td><td style="font-family: monospace; font-size: 0.75rem;">${orden.statusML}</td></tr><tr><td>Status Detail</td><td style="font-family: monospace; font-size: 0.75rem;">${orden.statusDetailML}</td></tr>` : ''}
                        <tr><td>Ingreso Bruto</td><td style="color: #92400e; font-weight: 600;">${Utilidades.formatMoney(importeML)}</td></tr>
                        <tr><td>Comisión MP</td><td>-${Utilidades.formatMoney(costos.comision || 0)}</td></tr>
                        <tr><td>Envío</td><td>-${Utilidades.formatMoney(costos.envio || 0)}</td></tr>
                        <tr><td>IVA retenido</td><td>-${Utilidades.formatMoney(costos.iva || 0)}</td></tr>
                        <tr><td>ISR retenido</td><td>-${Utilidades.formatMoney(costos.isr || 0)}</td></tr>
                        <tr style="border-top: 2px solid var(--gray-300);"><td><strong>Neto ML</strong></td><td><strong>${Utilidades.formatMoney(orden.neto?.real || 0)}</strong></td></tr>
                    </table>
                </div>
                <div class="erp-comparativa-vs">VS</div>
                <div class="erp-comparativa-col erp">
                    <h4 style="color: #1e40af;">Intelisis ERP</h4>
                    ${encontrado ? `
                        <table class="validacion-detalle-table">
                            <tr><td>Estatus</td><td><span style="background: #dbeafe; color: #1e40af; padding: 2px 6px; border-radius: 3px;">${orden.api?.estatusERP || 'N/A'}</span></td></tr>
                            <tr><td>Movimientos</td><td style="font-family: monospace; font-size: 0.75rem;">${orden.api?.movimientos || 'N/A'}</td></tr>
                            <tr><td>Cliente</td><td>${orden.api?.cliente || 'N/A'}</td></tr>
                            <tr style="border-top: 2px solid var(--gray-300);"><td><strong>Importe ERP</strong></td><td style="color: #1e40af; font-weight: 600;">${Utilidades.formatMoney(importeERP)}</td></tr>
                        </table>
                    ` : `<div style="text-align: center; padding: 20px; color: var(--danger);"><p style="font-weight: 500;">No encontrado en ERP</p></div>`}
                </div>
            </div>
            <div style="background: ${!encontrado ? 'var(--danger-light)' : Math.abs(diferencia) <= 0.01 ? '#d1fae5' : 'var(--warning-light)'}; padding: 16px; border-radius: 8px; margin: 16px 0;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <div>
                        <h4 style="margin: 0 0 8px 0; color: ${!encontrado ? 'var(--danger-dark)' : Math.abs(diferencia) <= 0.01 ? '#065f46' : '#92400e'};">
                            ${!encontrado ? 'No Encontrada en ERP' : Math.abs(diferencia) <= 0.01 ? 'Montos Coinciden' : 'Diferencia Detectada'}
                        </h4>
                    </div>
                    ${encontrado ? `<div style="text-align: right;"><div style="font-size: 0.75rem;">Diferencia</div><div style="font-size: 1.5rem; font-weight: 700;">${diferencia >= 0 ? '+' : ''}${Utilidades.formatMoney(diferencia)}</div></div>` : ''}
                </div>
            </div>
            ${encontrado ? `<div style="background: var(--gray-50); padding: 16px; border-radius: 8px;"><h4 style="margin: 0 0 12px 0;">Movimientos en Intelisis (${movimientosDetalle.length})</h4><div style="max-height: 300px; overflow-y: auto;">${movimientosHTML}</div></div>` : ''}
        `;

        document.getElementById('erpDetalleBody').innerHTML = detalleHTML;
        document.getElementById('erpDetalle').style.display = 'block';
        document.getElementById('erpDetalle').scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    cerrarDetalle() {
        document.getElementById('erpDetalle').style.display = 'none';
        this.ordenSeleccionada = null;
    },

    exportar() {
        const ordenes = estado.ordenesConsolidadas.filter(o => o.api !== undefined);
        if (ordenes.length === 0) {
            Interfaz.mostrarToast('No hay datos para exportar', 'warning');
            return;
        }

        const datos = ordenes.map(orden => ({
            'Referencia': orden.referenciaERP,
            'SOURCE_IDs': (orden.sourceIds || []).join(', '),
            'Cantidad SOURCE_IDs': orden.sourceIds?.length || 0,
            'Estatus ML': orden.estatus,
            'Status ML': orden.statusML || '',
            'Status Detail ML': orden.statusDetailML || '',
            'Encontrado ERP': orden.api?.encontrado ? 'SI' : 'NO',
            'Estatus ERP': orden.api?.estatusERP || 'N/A',
            'Movimientos': orden.api?.movimientos || 'N/A',
            'Cliente': orden.api?.cliente || '',
            'Envio Comprador':  +(orden.costos?.shipping || 0),
            'Comision MP':      -Math.abs(orden.costos?.comision || 0),
            '% Comision':       (orden.ingresos?.bruto) ? Utilidades.roundMoney((orden.costos?.comision || 0) / orden.ingresos.bruto * 100) : 0,
            'Envio ML':         -Math.abs(orden.costos?.envio || 0),
            'Financiamiento':   -Math.abs(orden.costos?.financiamiento || 0),
            'Retención ML':     +(orden.ajustes?.retencion_ml || 0),
            'Ajuste ISR/IVA':   +(orden.ajustes?.ajuste_isr_iva || 0),
            'Env.Comp.Sep.':    +(orden.ajustes?.envio_comprador_sep || 0),
            'Envío Cte.':       orden.costos?.shipping || 0,
            'Total Gastos':     -Math.abs(orden.costos?.total || 0),
            'Subtotal':         Utilidades.roundMoney(((orden.ingresos?.bruto || 0) - (orden.costos?.shipping || 0)) / 1.16),
            'IVA 8%':           orden.costos?.iva  || 0,
            'ISR 2.5%':         orden.costos?.isr  || 0,
            'Importe Edo. Cuenta': +(orden.neto?.estadoCuenta || 0),
            'Importe ML':          (orden.ingresos?.bruto || 0),
            'Importe ERP':         +(orden.api?.importeERP || 0),
            'Diferencia':          orden.api?.diferencia || 0
        }));

        const ws = XLSX.utils.json_to_sheet(datos);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Análisis ERP');
        XLSX.writeFile(wb, `analisis_erp_${new Date().toISOString().split('T')[0]}.xlsx`);
        Interfaz.mostrarToast('Análisis ERP exportado', 'success');
    }
};

// Inicializar AnalisisERPManager
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => AnalisisERPManager.inicializar());
} else {
    AnalisisERPManager.inicializar();
}

// ========== MÓDULO WALMART MX ==========
const WalmartManager = {
    _tab:     'consolidado',   // tab activo
    _archivo: null,            // nombre del Excel generado (para descargar)
    _files:   {},              // { atlas: File, spring: File, individual: File }

    setTab(tab) {
        this._tab = tab;
        document.getElementById('wm-tab-consolidado').classList.toggle('active', tab === 'consolidado');
        document.getElementById('wm-tab-individual').classList.toggle('active', tab === 'individual');
        document.getElementById('wm-panel-consolidado').style.display = tab === 'consolidado' ? '' : 'none';
        document.getElementById('wm-panel-individual').style.display  = tab === 'individual'  ? '' : 'none';
        document.getElementById('wm-resultado').style.display = 'none';
    },

    onFile(key, input) {
        const file = input.files[0];
        if (!file) return;
        this._files[key] = file;

        const nameEl = document.getElementById(`wm-name-${key}`);
        if (nameEl) nameEl.textContent = file.name;

        const zoneEl = document.getElementById(`wm-zone-${key}`);
        if (zoneEl) zoneEl.classList.add('uploaded');

        this._updateButtons();
    },

    _updateButtons() {
        const btnC = document.getElementById('wm-btn-consolidado');
        const btnI = document.getElementById('wm-btn-individual');
        if (btnC) btnC.disabled = !(this._files.atlas && this._files.spring);
        if (btnI) btnI.disabled = !this._files.individual;
    },

    _setLoading(on, texto = 'Procesando conciliación Walmart...') {
        const el = document.getElementById('wm-loading');
        if (!el) return;
        if (on) {
            el.classList.add('show');
            document.getElementById('wm-loading-text').textContent = texto;
        } else {
            el.classList.remove('show');
        }
    },

    _mostrarResultado(resumen) {
        const stats = [
            { label: 'Registros Atlas',     value: resumen.atlas   ?? '—' },
            { label: 'Registros Spring',    value: resumen.spring  ?? '—' },
            { label: 'Casos creados',       value: resumen.casos_nuevos       ?? 0 },
            { label: 'Casos actualizados',  value: resumen.casos_actualizados ?? 0 },
        ];

        document.getElementById('wm-stats').innerHTML = stats.map(s => `
            <div class="dashboard-stat-card">
                <h4>${s.label}</h4>
                <div class="number">${s.value.toLocaleString?.() ?? s.value}</div>
            </div>
        `).join('');

        document.getElementById('wm-resultado').style.display = '';
    },

    async conciliarConsolidado() {
        if (!this._files.atlas || !this._files.spring) return;

        this._setLoading(true, 'Consultando API Walmart y ERP — puede tardar 1-2 minutos...');
        document.getElementById('wm-resultado').style.display = 'none';
        document.getElementById('wm-btn-consolidado').disabled = true;

        try {
            const fd = new FormData();
            fd.append('atlas_csv',  this._files.atlas);
            fd.append('spring_csv', this._files.spring);

            const resp = await apiFetch('/walmart/conciliar', { method: 'POST', body: fd });
            const data = await resp.json();

            if (!resp.ok) throw new Error(data.detail || 'Error en el servidor');

            this._archivo = data.archivo;
            this._mostrarResultado(data.resumen);

            const nc = (data.resumen.casos_nuevos ?? 0);
            Interfaz.mostrarToast(
                `Conciliación lista — ${nc} devolución${nc !== 1 ? 'es' : ''} registrada${nc !== 1 ? 's' : ''}`,
                'success'
            );
        } catch (err) {
            Interfaz.mostrarToast('Error: ' + err.message, 'error');
        } finally {
            this._setLoading(false);
            document.getElementById('wm-btn-consolidado').disabled = false;
        }
    },

    async conciliarIndividual() {
        if (!this._files.individual) return;

        const cuenta = document.getElementById('wm-cuenta-individual').value;
        this._setLoading(true, `Procesando cuenta ${cuenta}...`);
        document.getElementById('wm-resultado').style.display = 'none';
        document.getElementById('wm-btn-individual').disabled = true;

        try {
            const fd = new FormData();
            fd.append('csv_file', this._files.individual);
            fd.append('cuenta',   cuenta);

            const resp = await apiFetch(`/walmart/conciliar-individual?cuenta=${cuenta}`, { method: 'POST', body: fd });
            const data = await resp.json();

            if (!resp.ok) throw new Error(data.detail || 'Error en el servidor');

            this._archivo = data.archivo;
            this._mostrarResultado({ ...data.resumen, atlas: undefined, spring: undefined, [cuenta]: data.resumen[cuenta] });

            Interfaz.mostrarToast('Reporte individual listo', 'success');
        } catch (err) {
            Interfaz.mostrarToast('Error: ' + err.message, 'error');
        } finally {
            this._setLoading(false);
            document.getElementById('wm-btn-individual').disabled = false;
        }
    },

    descargar() {
        if (!this._archivo) return;
        const token = AuthManager.getToken();
        // Descarga directa via link con token en header no es posible en <a>,
        // usamos fetch + blob para respetar el auth
        apiFetch(`/walmart/descargar/${this._archivo}`)
            .then(r => {
                if (!r.ok) throw new Error('No se encontró el archivo');
                return r.blob();
            })
            .then(blob => {
                const url = URL.createObjectURL(blob);
                const a   = document.createElement('a');
                a.href     = url;
                a.download = this._archivo;
                a.click();
                URL.revokeObjectURL(url);
            })
            .catch(err => Interfaz.mostrarToast('Error al descargar: ' + err.message, 'error'));
    },
};

// ========== EXPOSICIÓN AL SCOPE GLOBAL ==========
window.CasosManager = CasosManager;
window.ConciliacionManager = ConciliacionManager;
window.ApiManager = ApiManager;
window.PagosCxCManager = PagosCxCManager;
window.Interfaz = Interfaz;
window.exportarSeguimiento = exportarSeguimiento;
window.exportarCasos = exportarCasos;
window.estado = estado;
window.ResumenEstatusManager = ResumenEstatusManager;
window.AnalisisValidacionManager = AnalisisValidacionManager;
window.AnalisisERPManager = AnalisisERPManager;
window.HistoricoManager = HistoricoManager;
window.WalmartManager = WalmartManager;