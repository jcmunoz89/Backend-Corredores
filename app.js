/* Kensa - Conecta v1.0 — app.js
   - Multi-tenant (localStorage)
   - CRUD de tareas + filtros + import/export
   - Dashboard por compañía/estados
   - Editor inline “Interior del caso” con datos de la tarea
   - Paginación robusta (8/15/30/60/100)
   - Selección por página + Eliminación masiva (FIX)
*/

/* ===========================
   Helpers y tema
=========================== */
(() => {
  const THEME_KEY = 'kensaTheme';

  const applyThemeAttrs = (theme) => {
    document.documentElement.setAttribute('data-theme', theme);
    document.body?.setAttribute('data-theme', theme);
    const appRoot = document.getElementById('app');
    if (appRoot) appRoot.setAttribute('data-theme', theme);
  };

  const setTheme = (theme) => {
    const next = theme || 'dark';
    applyThemeAttrs(next);
    localStorage.setItem(THEME_KEY, next);
  };

  const resolveInitialTheme = () => {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored) return stored;
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches) {
      return 'light';
    }
    return 'dark';
  };

  setTheme(resolveInitialTheme());

  const toggleTheme = () => {
    const current = localStorage.getItem(THEME_KEY) || 'dark';
    setTheme(current === 'dark' ? 'light' : 'dark');
  };

  document.addEventListener('click', (event) => {
    const toggleBtn = event.target.closest('[data-theme-toggle]');
    if (!toggleBtn) return;
    event.preventDefault();
    toggleTheme();
  });

  window.addEventListener('storage', (evt) => {
    if (evt.key === THEME_KEY && evt.newValue) applyThemeAttrs(evt.newValue);
  });

  if (window.matchMedia) {
    const media = window.matchMedia('(prefers-color-scheme: light)');
    media.addEventListener('change', (ev) => {
      const stored = localStorage.getItem(THEME_KEY);
      if (!stored) setTheme(ev.matches ? 'light' : 'dark');
    });
  }
})();

/* ===========================
   Estado + Storage keys
=========================== */
const kTasks   = (t) => `corredores_v1:${t}:tasks`;
const kUsers   = (t) => `corredores_v1:${t}:users`;
const kSession = (t) => `corredores_v1:${t}:session`;
const kClients = (t) => `corredores_v1:${t}:clients`;
const kTaskAuditLog = (t) => `corredores_v1:${t}:tasks_audit`;
const KENSA_GOALS_KEY = 'kensaGoalsByTenant';
const KENSA_METAS_HISTORIAL_KEY = 'kensa_metas_historial';
const REMEMBER_KEY = 'kensaRememberLogin';
const THREE_MONTHS_MS = 90 * 24 * 60 * 60 * 1000;
const DEFAULT_TENANT = 'demo-kensa';
const MASTER_SEED_EMAIL = 'jmunoz@kensa.cl';
const MASTER_SEED_PASSWORD = '874095Jc!';

const AUTO_RENEWAL_DAYS_BEFORE = 21;
const SAVE_OPTION_SOLO_GUARDAR = '__solo_guardar__';

const todayISO = () => new Date().toISOString().slice(0, 10);
const uuid = () =>
  Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
function toLocalDateString(dateInput) {
  if (!dateInput) return '';
  const date =
    dateInput instanceof Date ? new Date(dateInput.getTime()) : new Date(dateInput);
  if (Number.isNaN(date.getTime())) return '';
  const offset = date.getTimezoneOffset();
  date.setMinutes(date.getMinutes() - offset);
  return date.toISOString().slice(0, 10);
}
function normalizeDateOnly(value) {
  if (!value) return '';
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
  return toLocalDateString(value);
}
function getAutoRenewalLastRunKey(tenantId) {
  return `kensa_autoRenewalLastRun_${tenantId || 'default'}`;
}
const TYPE_CODES = {
  'Vehículo': '1',
  'Hogar': '2',
  'Vida': '3',
  'Salud': '4',
  'Viaje': '5',
  'SOAP': '6',
  'TRC': '6',
  'Otro': '8',
};
const DEFAULT_TASK_TYPE = 'Vehículo';
const DEFAULT_INSURER = 'Sin Asignar';
const normalizeInsurerKey = (value) => String(value || '').trim().toLowerCase();
const INSURER_DISPLAY_PAIRS = [
  ['Aseguradora Porvenir S.A. (76.598.625-7)', 'Porvenir'],
  ['Assurant Chile Compañía de Seguros Generales S.A. (76.212.519-6)', 'Assurant'],
  ['BCI Seguros Generales S.A. (99.147.000-K)', 'BCI'],
  ['BNP Paribas Cardif Seguros Generales S.A. (96.837.640-3)', 'Cardif'],
  ['Chubb Seguros Chile S.A. (99.225.000-3)', 'Chubb'],
  ['Compañía de Seguros Generales Consorcio Nacional de Seguros S.A. (96.654.180-6)', 'Consorcio'],
  ['Compañía de Seguros Generales Continental S.A. (76.039.758-K)', 'Continental'],
  ['Contempora Compañía de Seguros Generales S.A. (76.981.875-8)', 'Contempora'],
  ['Everest Compañía de Seguros Generales Chile S.A. (77.591.207-3)', 'Everest'],
  ['FID Chile Seguros Generales S.A. (77.096.952-2)', 'FID'],
  ['HDI Seguros S.A. (99.061.000-2)', 'HDI'],
  ['Liberty Mutual Surety Seguros Chile S.A. (78.027.718-1)', 'Liberty Mutual Surety'],
  ['MAPFRE Compañía de Seguros Generales de Chile S.A. (96.508.210-7)', 'MAPFRE'],
  ['MetLife Chile Seguros Generales S.A. (76.328.793-9)', 'MetLife'],
  ['Mutualidad de Carabineros (99.024.000-0)', 'Mutualidad de Carabineros'],
  ['Orion Seguros Generales S.A. (76.042.965-1)', 'Orion'],
  ['Reale Chile Seguros Generales S.A. (76.743.492-8)', 'Reale'],
  ['Renta Nacional Compañía de Seguros Generales S.A. (94.510.000-1)', 'Renta Nacional'],
  ['Seguros Generales Suramericana S.A. (99.017.000-2)', 'Suramericana (SURA)'],
  ['Southbridge Compañía de Seguros Generales S.A. (99.288.000-7)', 'Southbridge'],
  ['Starr International Seguros Generales S.A. (76.620.932-7)', 'Starr'],
  ['Unnio Seguros Generales S.A. (76.173.258-7)', 'Unnio'],
  ['Zenit Seguros Generales S.A. (76.061.223-5)', 'Zenit'],
  ['Zurich Chile Seguros Generales S.A. (99.037.000-1)', 'Zurich'],
  ['Zurich Santander Seguros Generales Chile S.A. (76.590.840-K)', 'Zurich Santander'],
  ['4 Life Seguros de Vida S.A. (76.418.751-2)', '4 Life'],
  ['Alemana Seguros S.A. (76.511.423-3)', 'Alemana'],
  ['Augustar Seguros de Vida S.A. (76.632.384-7)', 'Augustar'],
  ['BCI Seguros Vida S.A. (96.573.600-K)', 'BCI'],
  ['BICE Vida Compañía de Seguros S.A. (96.656.410-5)', 'BICE Vida'],
  ['BNP Paribas Cardif Seguros de Vida S.A. (96.837.630-6)', 'Cardif'],
  ['Bupa Compañía de Seguros de Vida S.A. (76.282.191-5)', 'Bupa'],
  ['CF Seguros de Vida S.A. (76.477.116-8)', 'CF'],
  ['Chubb Seguros de Vida Chile S.A. (99.588.060-1)', 'Chubb'],
  ['CN Life, Compañía de Seguros de Vida S.A. (96.579.280-5)', 'CN Life'],
  ['Colmena Compañía de Seguros de Vida S.A. (76.408.757-7)', 'Colmena'],
  ['Compañía de Seguros Confuturo S.A. (96.571.890-7)', 'Confuturo'],
  ['Compañía de Seguros de Vida Cámara S.A. (99.003.000-6)', 'Vida Cámara'],
  ['Compañía de Seguros de Vida Consorcio Nacional de Seguros S.A. (99.012.000-5)', 'Consorcio'],
  ['Divina Pastora Seguros de Vida S.A. (77.205.281-2)', 'Divina Pastora'],
  ['EuroAmerica Seguros de Vida S.A. (99.279.000-8)', 'EuroAmerica'],
  ['Help Seguros de Vida S.A. (76.213.329-6)', 'Help'],
  ['MAPFRE Compañía de Seguros de Vida de Chile S.A. (96.933.030-K)', 'MAPFRE'],
  ['MetLife Chile Seguros de Vida S.A. (99.289.000-2)', 'MetLife'],
  ['Mutual de Seguros de Chile (70.015.730-K)', 'Mutual'],
  ['Mutualidad del Ejército y Aviación (99.025.000-6)', 'Mutualidad del Ejército y Aviación'],
  ['Penta Vida Compañía de Seguros de Vida S.A. (96.812.960-0)', 'Penta Vida'],
  ['Principal Compañía de Seguros de Vida Chile S.A. (96.588.080-1)', 'Principal'],
  ['Renta Nacional Compañía de Seguros de Vida S.A. (94.716.000-1)', 'Renta Nacional'],
  ['Save Compañía de Seguros de Vida S.A. (76.034.737-K)', 'Save'],
  ['Seguros CLC S.A. (76.573.480-0)', 'CLC'],
  ['Seguros de Vida SURA S.A. (96.549.050-7)', 'SURA'],
  ['Seguros de Vida Suramericana S.A. (76.263.414-7)', 'Suramericana'],
  ['Seguros de Vida y Salud UC Christus S.A. (76.632.553-K)', 'UC Christus'],
  ['Seguros Vida Security Previsión S.A. (99.301.000-6)', 'Security Previsión'],
  ['Zurich Chile Seguros de Vida S.A. (99.185.000-7)', 'Zurich'],
  ['Zurich Santander Seguros de Vida Chile S.A. (96.819.630-8)', 'Zurich Santander'],
];
const INSURER_DISPLAY_MAP = new Map();
INSURER_DISPLAY_PAIRS.forEach(([legalName, displayName]) => {
  const key = normalizeInsurerKey(legalName);
  INSURER_DISPLAY_MAP.set(key, displayName);
  const baseName = legalName.replace(/\s*\([^)]*\)\s*$/, '').trim();
  if (baseName && baseName !== legalName) {
    const baseKey = normalizeInsurerKey(baseName);
    if (!INSURER_DISPLAY_MAP.has(baseKey)) INSURER_DISPLAY_MAP.set(baseKey, displayName);
  }
});
const TASK_STATUS_OPTIONS = [
  'Abierta',
  'En renovación',
  'Propuesta enviada',
  'Póliza emitida',
  'Desistida',
  'Pérdida',
  'Renovada',
  'Renovada con modificación',
];
const STATUS_ACTIVE = new Set(['Abierta', 'En renovación', 'Propuesta enviada']);
const STATUS_WON = new Set(['Póliza emitida', 'Renovada', 'Renovada con modificación']);
const STATUS_LOST = new Set(['Desistida', 'Pérdida']);
const FILTER_GROUPS = {
  nuevos: ['Abierta', 'En renovación'],
  gestion: ['Propuesta enviada'],
  finalizados: ['Póliza emitida', 'Renovada', 'Renovada con modificación'],
  desistidos: ['Desistida', 'Pérdida'],
};
// Normaliza estados históricos (v1.0) hacia el nuevo catálogo.
const LEGACY_STATUS_MAP = {
  cerrada: 'Póliza emitida',
  cerrado: 'Póliza emitida',
  'de alta': 'Póliza emitida',
  'en progreso': 'En renovación',
  'en renovacion': 'En renovación',
  perdida: 'Pérdida',
  'pérdida': 'Pérdida',
};

const REGIONES_CHILE = [
  'Arica y Parinacota',
  'Tarapacá',
  'Antofagasta',
  'Atacama',
  'Coquimbo',
  'Valparaíso',
  'Metropolitana de Santiago',
  'Libertador General Bernardo O\'Higgins',
  'Maule',
  'Ñuble',
  'Biobío',
  'La Araucanía',
  'Los Ríos',
  'Los Lagos',
  'Aysén del General Carlos Ibáñez del Campo',
  'Magallanes y de la Antártica Chilena',
];

const COMUNAS_POR_REGION = {
  'Arica y Parinacota': ['Arica', 'Camarones', 'Putre', 'General Lagos'],
  Tarapacá: ['Iquique', 'Alto Hospicio', 'Pozo Almonte', 'Camiña', 'Colchane', 'Huara', 'Pica'],
  Antofagasta: [
    'Antofagasta',
    'Mejillones',
    'Sierra Gorda',
    'Taltal',
    'Calama',
    'Ollagüe',
    'San Pedro de Atacama',
    'Tocopilla',
    'María Elena',
  ],
  Atacama: [
    'Copiapó',
    'Caldera',
    'Tierra Amarilla',
    'Chañaral',
    'Diego de Almagro',
    'Vallenar',
    'Alto del Carmen',
    'Freirina',
    'Huasco',
  ],
  Coquimbo: [
    'La Serena',
    'Coquimbo',
    'Andacollo',
    'La Higuera',
    'Paihuano',
    'Vicuña',
    'Illapel',
    'Canela',
    'Los Vilos',
    'Salamanca',
    'Ovalle',
    'Combarbalá',
    'Monte Patria',
    'Punitaqui',
    'Río Hurtado',
  ],
  Valparaíso: [
    'Valparaíso',
    'Casablanca',
    'Concón',
    'Juan Fernández',
    'Puchuncaví',
    'Quintero',
    'Viña del Mar',
    'Quilpué',
    'Villa Alemana',
    'Limache',
    'Olmué',
    'San Antonio',
    'Algarrobo',
    'Cartagena',
    'El Quisco',
    'El Tabo',
    'Santo Domingo',
    'San Felipe',
    'Catemu',
    'Llaillay',
    'Panquehue',
    'Putaendo',
    'Santa María',
    'Los Andes',
    'Calle Larga',
    'Rinconada',
    'San Esteban',
    'La Ligua',
    'Cabildo',
    'Papudo',
    'Zapallar',
    'Quillota',
    'La Calera',
    'Hijuelas',
    'La Cruz',
    'Nogales',
    'Isla de Pascua',
  ],
  'Metropolitana de Santiago': [
    'Cerrillos',
    'Cerro Navia',
    'Conchalí',
    'El Bosque',
    'Estación Central',
    'Huechuraba',
    'Independencia',
    'La Cisterna',
    'La Florida',
    'La Granja',
    'La Pintana',
    'La Reina',
    'Las Condes',
    'Lo Barnechea',
    'Lo Espejo',
    'Lo Prado',
    'Macul',
    'Maipú',
    'Ñuñoa',
    'Pedro Aguirre Cerda',
    'Peñalolén',
    'Providencia',
    'Pudahuel',
    'Quilicura',
    'Quinta Normal',
    'Recoleta',
    'Renca',
    'San Joaquín',
    'San Miguel',
    'San Ramón',
    'Santiago',
    'Vitacura',
    'Puente Alto',
    'Pirque',
    'San José de Maipo',
    'Colina',
    'Lampa',
    'Tiltil',
    'San Bernardo',
    'Buin',
    'Calera de Tango',
    'Paine',
    'Melipilla',
    'Alhué',
    'Curacaví',
    'María Pinto',
    'San Pedro',
    'Talagante',
    'El Monte',
    'Isla de Maipo',
    'Padre Hurtado',
    'Peñaflor',
  ],
  'Libertador General Bernardo O\'Higgins': [
    'Rancagua',
    'Codegua',
    'Coinco',
    'Coltauco',
    'Doñihue',
    'Graneros',
    'Las Cabras',
    'Machalí',
    'Malloa',
    'Mostazal',
    'Olivar',
    'Peumo',
    'Pichidegua',
    'Quinta de Tilcoco',
    'Rengo',
    'Requínoa',
    'San Vicente',
    'Pichilemu',
    'La Estrella',
    'Litueche',
    'Marchihue',
    'Navidad',
    'Paredones',
    'San Fernando',
    'Chépica',
    'Chimbarongo',
    'Lolol',
    'Nancagua',
    'Palmilla',
    'Peralillo',
    'Placilla',
    'Pumanque',
    'Santa Cruz',
  ],
  Maule: [
    'Talca',
    'Constitución',
    'Curepto',
    'Empedrado',
    'Maule',
    'Pelarco',
    'Pencahue',
    'Río Claro',
    'San Clemente',
    'San Rafael',
    'Linares',
    'Colbún',
    'Longaví',
    'Parral',
    'Retiro',
    'San Javier',
    'Villa Alegre',
    'Yerbas Buenas',
    'Cauquenes',
    'Chanco',
    'Pelluhue',
    'Curicó',
    'Hualañé',
    'Licantén',
    'Molina',
    'Rauco',
    'Romeral',
    'Sagrada Familia',
    'Teno',
    'Vichuquén',
  ],
  Ñuble: [
    'Chillán',
    'Chillán Viejo',
    'Cobquecura',
    'Coelemu',
    'Coihueco',
    'El Carmen',
    'Ninhue',
    'Ñiquén',
    'Pemuco',
    'Pinto',
    'Portezuelo',
    'Quillón',
    'Quirihue',
    'Ránquil',
    'San Carlos',
    'San Fabián',
    'San Ignacio',
    'San Nicolás',
    'Treguaco',
    'Yungay',
    'Bulnes',
  ],
  Biobío: [
    'Concepción',
    'Coronel',
    'Chiguayante',
    'Florida',
    'Hualqui',
    'Lota',
    'Penco',
    'San Pedro de la Paz',
    'Santa Juana',
    'Talcahuano',
    'Tomé',
    'Hualpén',
    'Los Ángeles',
    'Antuco',
    'Cabrero',
    'Laja',
    'Mulchén',
    'Nacimiento',
    'Negrete',
    'Quilaco',
    'Quilleco',
    'San Rosendo',
    'Santa Bárbara',
    'Tucapel',
    'Yumbel',
    'Alto Biobío',
    'Arauco',
    'Cañete',
    'Contulmo',
    'Curanilahue',
    'Lebu',
    'Los Álamos',
    'Tirúa',
  ],
  'La Araucanía': [
    'Temuco',
    'Carahue',
    'Cunco',
    'Curarrehue',
    'Freire',
    'Galvarino',
    'Gorbea',
    'Lautaro',
    'Loncoche',
    'Melipeuco',
    'Nueva Imperial',
    'Padre Las Casas',
    'Perquenco',
    'Pitrufquén',
    'Pucón',
    'Saavedra',
    'Teodoro Schmidt',
    'Toltén',
    'Vilcún',
    'Villarrica',
    'Angol',
    'Collipulli',
    'Curacautín',
    'Ercilla',
    'Lonquimay',
    'Los Sauces',
    'Lumaco',
    'Purén',
    'Renaico',
    'Traiguén',
    'Victoria',
  ],
  'Los Ríos': [
    'Valdivia',
    'Corral',
    'Lanco',
    'Los Lagos',
    'Máfil',
    'Mariquina',
    'Paillaco',
    'Panguipulli',
    'La Unión',
    'Futrono',
    'Lago Ranco',
    'Río Bueno',
  ],
  'Los Lagos': [
    'Puerto Montt',
    'Calbuco',
    'Cochamó',
    'Fresia',
    'Frutillar',
    'Los Muermos',
    'Llanquihue',
    'Maullín',
    'Puerto Varas',
    'Ancud',
    'Castro',
    'Chonchi',
    'Curaco de Vélez',
    'Dalcahue',
    'Puqueldón',
    'Queilén',
    'Quellón',
    'Quemchi',
    'Quinchao',
    'Osorno',
    'Puerto Octay',
    'Purranque',
    'Puyehue',
    'Río Negro',
    'San Juan de la Costa',
    'San Pablo',
    'Chaitén',
    'Futaleufú',
    'Hualaihué',
    'Palena',
  ],
  'Aysén del General Carlos Ibáñez del Campo': [
    'Coyhaique',
    'Lago Verde',
    'Aysén',
    'Cisnes',
    'Guaitecas',
    'Cochrane',
    'O\'Higgins',
    'Tortel',
    'Chile Chico',
    'Río Ibáñez',
  ],
  'Magallanes y de la Antártica Chilena': [
    'Punta Arenas',
    'Laguna Blanca',
    'Río Verde',
    'San Gregorio',
    'Cabo de Hornos',
    'Antártica',
    'Porvenir',
    'Primavera',
    'Timaukel',
    'Natales',
    'Torres del Paine',
  ],
};

function buscarRegiones(term = '') {
  const q = (term || '').trim().toLowerCase();
  if (!q) return REGIONES_CHILE.slice();
  return REGIONES_CHILE.filter((region) => region.toLowerCase().includes(q));
}

function buscarComunas(region = '', term = '') {
  const comunas = COMUNAS_POR_REGION[region] || [];
  const q = (term || '').trim().toLowerCase();
  if (!q) return comunas.slice();
  return comunas.filter((c) => c.toLowerCase().includes(q));
}

function normalizeTaskStatus(status) {
  const raw = (status || '').trim();
  if (!raw) return TASK_STATUS_OPTIONS[0];
  const lower = raw.toLowerCase();
  if (LEGACY_STATUS_MAP[lower]) return LEGACY_STATUS_MAP[lower];
  const found = TASK_STATUS_OPTIONS.find((opt) => opt.toLowerCase() === lower);
  return found || TASK_STATUS_OPTIONS[0];
}

const CHANNEL_TRADICIONAL = 'Tradicional';
const CHANNEL_CANAL_WEB = 'Canal web';
const CHANNEL_RENOVACION = 'Renovación';
const STATUS_TERMINAL = new Set([
  'Póliza emitida',
  'Renovada',
  'Renovada con modificación',
  'Desistida',
  'Pérdida',
]);
const FLOW_FROM_OPEN = ['Propuesta enviada', 'Desistida', 'Pérdida'];
const FLOW_FROM_PROPOSAL_TRAD = ['Póliza emitida', 'Desistida', 'Pérdida'];
const FLOW_FROM_PROPOSAL_REN = ['Renovada', 'Renovada con modificación', 'Desistida', 'Pérdida'];

function normalizeChannel(value = '') {
  const raw = (value || '').trim();
  if (!raw) return CHANNEL_TRADICIONAL;
  const lower = raw.toLowerCase();
  if (lower === 'canal web') return CHANNEL_CANAL_WEB;
  if (lower === 'renovacion' || lower === 'renovación') return CHANNEL_RENOVACION;
  if (lower === 'tradicional') return CHANNEL_TRADICIONAL;
  return raw;
}

function normalizeInitialStatus(value, canal) {
  const normalized = normalizeTaskStatus(value);
  if (normalized === 'En renovación') return 'En renovación';
  if (normalized === 'Abierta') return 'Abierta';
  const normalizedChannel = normalizeChannel(canal);
  if (normalizedChannel === CHANNEL_RENOVACION) return 'En renovación';
  return 'Abierta';
}

function inferInitialStatus(task = {}) {
  return normalizeInitialStatus(task.estadoInicial || task.status, task.canalIngreso);
}

function isStatusAllowedForChannel(status, canal) {
  const normalizedStatus = normalizeTaskStatus(status);
  const normalizedChannel = normalizeChannel(canal);
  if (normalizedStatus === 'Abierta') {
    return (
      normalizedChannel === CHANNEL_TRADICIONAL ||
      normalizedChannel === CHANNEL_CANAL_WEB
    );
  }
  if (normalizedStatus === 'En renovación') {
    return normalizedChannel === CHANNEL_RENOVACION;
  }
  if (normalizedStatus === 'Renovada' || normalizedStatus === 'Renovada con modificación') {
    return normalizedChannel === CHANNEL_RENOVACION;
  }
  if (normalizedStatus === 'Póliza emitida' && normalizedChannel === CHANNEL_RENOVACION) {
    return false;
  }
  return true;
}

function getAllowedNextStates(task = {}) {
  const canal = normalizeChannel(task.canalIngreso);
  const estadoActual = normalizeTaskStatus(task.status);
  const estadoInicial = normalizeInitialStatus(task.estadoInicial || estadoActual, canal);
  let allowed = [];
  if (estadoActual === 'Abierta' || estadoActual === 'En renovación') {
    allowed = FLOW_FROM_OPEN.slice();
  } else if (estadoActual === 'Propuesta enviada') {
    allowed =
      estadoInicial === 'En renovación'
        ? FLOW_FROM_PROPOSAL_REN.slice()
        : FLOW_FROM_PROPOSAL_TRAD.slice();
  } else if (!STATUS_TERMINAL.has(estadoActual)) {
    allowed = FLOW_FROM_OPEN.slice();
  }
  const normalizedOptions = allowed
    .map((state) => normalizeTaskStatus(state))
    .filter(
      (state) => isStatusAllowedForChannel(state, canal) && state !== estadoActual
    );
  const deduped = [];
  normalizedOptions.forEach((state) => {
    if (!deduped.includes(state)) deduped.push(state);
  });
  return deduped;
}

// ===========================
// Metas de venta (modelo + persistencia)
// ===========================
const GOAL_DEFAULT_LIST = { userGoals: [], globalGoals: [] };
const normalizeGoalPeriod = (period = '') => {
  const raw = String(period || '').trim();
  if (/^\d{4}-\d{2}$/.test(raw)) return raw;
  const date = new Date(raw || Date.now());
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
};
const getCurrentGoalPeriod = () => normalizeGoalPeriod(new Date());
const getGoalPeriodForFilters = () => {
  const { to } = getDashboardDateBounds();
  return normalizeGoalPeriod(to);
};
const normalizeGoalNumber = (val) => {
  const num = Number(val);
  return Number.isFinite(num) ? num : 0;
};
const formatDateInputValue = (date) => {
  if (!(date instanceof Date)) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};
const cloneTenantGoals = (data = {}) => ({
  userGoals: Array.isArray(data.userGoals) ? data.userGoals.slice() : [],
  globalGoals: Array.isArray(data.globalGoals) ? data.globalGoals.slice() : [],
});
function loadAllGoalsFromStorage() {
  try {
    const raw = localStorage.getItem(KENSA_GOALS_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    if (typeof parsed !== 'object' || parsed === null) return {};
    const result = {};
    Object.keys(parsed).forEach((rut) => {
      result[rut] = cloneTenantGoals(parsed[rut]);
    });
    return result;
  } catch {
    return {};
  }
}
function saveAllGoalsToStorage(goalsByTenant) {
  try {
    const payload = {};
    Object.keys(goalsByTenant || {}).forEach((rut) => {
      payload[rut] = cloneTenantGoals(goalsByTenant[rut]);
    });
    localStorage.setItem(KENSA_GOALS_KEY, JSON.stringify(payload));
  } catch {
    // Ignorar errores de escritura
  }
}
function ensureTenantGoals(goalsByTenant, rutCorredora) {
  if (!rutCorredora) return cloneTenantGoals(GOAL_DEFAULT_LIST);
  if (!goalsByTenant[rutCorredora]) goalsByTenant[rutCorredora] = cloneTenantGoals(GOAL_DEFAULT_LIST);
  return goalsByTenant[rutCorredora];
}
function getUserGoalsForTenant(rutCorredora, userEmail) {
  if (!rutCorredora || !userEmail) return [];
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const target = normalizeEmail(userEmail);
  return tenantGoals.userGoals.filter((goal) => normalizeEmail(goal.userEmail) === target);
}
function upsertUserGoal(rutCorredora, goal) {
  if (!rutCorredora || !goal?.userEmail || !goal?.period) return null;
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const period = normalizeGoalPeriod(goal.period);
  const targetEmail = normalizeEmail(goal.userEmail);
  const now = Date.now();
  const payload = {
    rutCorredora,
    userEmail: goal.userEmail,
    userName: goal.userName || '',
    period,
    goalPremiumUF: normalizeGoalNumber(goal.goalPremiumUF),
    goalCommissionUF: normalizeGoalNumber(goal.goalCommissionUF),
    goalDeals:
      goal.goalDeals !== undefined && goal.goalDeals !== null
        ? normalizeGoalNumber(goal.goalDeals)
        : undefined,
  };
  const idx = tenantGoals.userGoals.findIndex(
    (g) => normalizeEmail(g.userEmail) === targetEmail && normalizeGoalPeriod(g.period) === period
  );
  if (idx > -1) {
    tenantGoals.userGoals[idx] = {
      ...tenantGoals.userGoals[idx],
      ...payload,
      updatedAt: now,
    };
  } else {
    tenantGoals.userGoals.push({ ...payload, createdAt: now, updatedAt: now });
  }
  saveAllGoalsToStorage(goalsByTenant);
  return payload;
}
function deleteUserGoal(rutCorredora, userEmail, period) {
  if (!rutCorredora || !userEmail || !period) return;
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const targetEmail = normalizeEmail(userEmail);
  const comparePeriod = normalizeGoalPeriod(period);
  const next = tenantGoals.userGoals.filter(
    (g) =>
      !(
        normalizeEmail(g.userEmail) === targetEmail &&
        normalizeGoalPeriod(g.period) === comparePeriod
      )
  );
  if (next.length === tenantGoals.userGoals.length) return;
  tenantGoals.userGoals = next;
  saveAllGoalsToStorage(goalsByTenant);
}
function getGlobalGoalsForTenant(rutCorredora) {
  if (!rutCorredora) return [];
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  return tenantGoals.globalGoals;
}
function getGlobalGoalForPeriod(rutCorredora, period) {
  if (!rutCorredora || !period) return null;
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const comparePeriod = normalizeGoalPeriod(period);
  return (
    tenantGoals.globalGoals.find(
      (g) => normalizeGoalPeriod(g.period) === comparePeriod && g.rutCorredora === rutCorredora
    ) || null
  );
}
function upsertGlobalGoal(rutCorredora, goal) {
  if (!rutCorredora || !goal?.period) return null;
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const period = normalizeGoalPeriod(goal.period);
  const now = Date.now();
  const payload = {
    rutCorredora,
    period,
    goalPremiumUF: normalizeGoalNumber(goal.goalPremiumUF),
    goalCommissionUF: normalizeGoalNumber(goal.goalCommissionUF),
    goalDeals:
      goal.goalDeals !== undefined && goal.goalDeals !== null
        ? normalizeGoalNumber(goal.goalDeals)
        : undefined,
  };
  const idx = tenantGoals.globalGoals.findIndex(
    (g) => g.rutCorredora === rutCorredora && normalizeGoalPeriod(g.period) === period
  );
  if (idx > -1) {
    tenantGoals.globalGoals[idx] = {
      ...tenantGoals.globalGoals[idx],
      ...payload,
      updatedAt: now,
    };
  } else {
    tenantGoals.globalGoals.push({ ...payload, createdAt: now, updatedAt: now });
  }
  saveAllGoalsToStorage(goalsByTenant);
  return payload;
}
function deleteGlobalGoal(rutCorredora, period) {
  if (!rutCorredora || !period) return;
  const goalsByTenant = loadAllGoalsFromStorage();
  const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
  const comparePeriod = normalizeGoalPeriod(period);
  const next = tenantGoals.globalGoals.filter(
    (g) =>
      !(
        g.rutCorredora === rutCorredora &&
        normalizeGoalPeriod(g.period) === comparePeriod
      )
  );
  if (next.length === tenantGoals.globalGoals.length) return;
  tenantGoals.globalGoals = next;
  saveAllGoalsToStorage(goalsByTenant);
}

// Metas históricas con vigencia (localStorage)
function cargarHistorialMetas() {
  try {
    const raw = localStorage.getItem(KENSA_METAS_HISTORIAL_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function guardarHistorialMetas(metasArray) {
  try {
    const payload = Array.isArray(metasArray) ? metasArray.slice() : [];
    localStorage.setItem(KENSA_METAS_HISTORIAL_KEY, JSON.stringify(payload));
  } catch {
    // Ignorar errores de escritura
  }
}

function normalizeMetaScopeId(scopeId) {
  if (scopeId === undefined || scopeId === null) return null;
  const str = String(scopeId).trim();
  return str === '' ? null : str;
}

function buildMetaId() {
  return `meta_${Date.now()}_${Math.floor(Math.random() * 10000)}`;
}

function getPreviousGoalPeriod(period = '') {
  const norm = normalizeGoalPeriod(period);
  const [yearStr, monthStr] = norm.split('-');
  const year = parseInt(yearStr, 10);
  const month = parseInt(monthStr, 10);
  if (!year || !month) return norm;
  const d = new Date(year, month - 1, 1);
  d.setMonth(d.getMonth() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function obtenerScopeMetaParaUsuarioActual() {
  const role = roleKey(state.user?.role || '');
  const rutCorredora = getCurrentRutCorredora?.() || '';
  if (role === ROLE_KEY_USUARIO_MAESTRO) {
    return { tipo: 'global', scopeId: null };
  }
  if (role === ROLE_KEY_ADMINISTRADOR) {
    return { tipo: 'corredora', scopeId: rutCorredora || null };
  }
  if (state.user?.id || state.user?.email) {
    return { tipo: 'usuario', scopeId: state.user.id || normalizeEmail(state.user.email) };
  }
  return { tipo: 'corredora', scopeId: rutCorredora || null };
}

function seedMetaHistoryFromLegacy(scope) {
  const metas = cargarHistorialMetas();
  if (metas.length) return metas;
  const { tipo, scopeId } = scope || obtenerScopeMetaParaUsuarioActual();
  const rut = getCurrentRutCorredora?.() || '';
  const legacyGoals = rut ? getGlobalGoalsForTenant(rut) : [];
  const legacyGoal =
    (Array.isArray(legacyGoals) && legacyGoals.length
      ? legacyGoals[legacyGoals.length - 1]
      : null) || null;
  if (!legacyGoal) return metas;
  const identity = getCurrentUserIdentity ? getCurrentUserIdentity() : {};
  const period = normalizeGoalPeriod(legacyGoal.period || getCurrentGoalPeriod());
  const seed = {
    idMeta: buildMetaId(),
    tipo: tipo || 'global',
    scopeId: normalizeMetaScopeId(scopeId),
    periodoDesde: period,
    periodoHasta: null,
    metaPrimaUF: normalizeGoalNumber(legacyGoal.goalPremiumUF),
    metaComisionUF: normalizeGoalNumber(legacyGoal.goalCommissionUF),
    creadoPorUsuarioId: identity.email || identity.id || 'desconocido',
    creadoEn: new Date().toISOString(),
    comentario: 'Migrado desde metas previas',
  };
  const next = metas.concat(seed);
  guardarHistorialMetas(next);
  return next;
}

function obtenerMetaVigenteParaFecha(tipo, scopeId, fechaReferencia, metasCache) {
  const periodRef = normalizeGoalPeriod(fechaReferencia || new Date());
  const metas = Array.isArray(metasCache) ? metasCache : cargarHistorialMetas();
  const normalizedScopeId = normalizeMetaScopeId(scopeId);
  const sameScope = (meta) =>
    meta &&
    meta.tipo === tipo &&
    normalizeMetaScopeId(meta.scopeId) === normalizedScopeId;
  const ordered = metas
    .filter((meta) => sameScope(meta) && meta.periodoDesde)
    .map((meta) => ({
      ...meta,
      periodoDesde: normalizeGoalPeriod(meta.periodoDesde),
      periodoHasta: meta.periodoHasta ? normalizeGoalPeriod(meta.periodoHasta) : null,
    }))
    .sort((a, b) => a.periodoDesde.localeCompare(b.periodoDesde));
  for (let i = ordered.length - 1; i >= 0; i -= 1) {
    const meta = ordered[i];
    const desdeOk = meta.periodoDesde <= periodRef;
    const hastaOk = !meta.periodoHasta || periodRef <= meta.periodoHasta;
    if (desdeOk && hastaOk) return meta;
  }
  return null;
}

const ROLE_USUARIO_MAESTRO = 'Usuario Maestro';
const ROLE_ADMINISTRADOR = 'Administrador';
const ROLE_SUPERVISOR = 'Supervisor';
const ROLE_AGENTE = 'Agente';

function normalizeRoleName(role = '') {
  const raw = String(role || '').trim();
  if (!raw) return '';
  const lower = raw.toLowerCase();
  if (lower === 'maestro kensa' || lower === 'usuario maestro') return ROLE_USUARIO_MAESTRO;
  if (lower === 'admin' || lower === 'administrador') return ROLE_ADMINISTRADOR;
  if (lower === 'supervisor') return ROLE_SUPERVISOR;
  if (lower === 'agente') return ROLE_AGENTE;
  return raw;
}

function roleKey(role) {
  return normalizeRoleName(role).toLowerCase();
}
const ROLE_KEY_USUARIO_MAESTRO = roleKey(ROLE_USUARIO_MAESTRO);
const ROLE_KEY_ADMINISTRADOR = roleKey(ROLE_ADMINISTRADOR);
const ROLE_KEY_SUPERVISOR = roleKey(ROLE_SUPERVISOR);
const ROLE_KEY_AGENTE = roleKey(ROLE_AGENTE);
const pad2 = (n) => String(n).padStart(2, '0');
/**
 * Genera un ID de negocio correlativo por día con formato AAAAMMDD-####.
 * Persiste el último día y secuencia en localStorage para evitar duplicados.
 */
function generateBusinessId() {
  const now = new Date();
  const dateStr = `${now.getFullYear()}${pad2(now.getMonth() + 1)}${pad2(now.getDate())}`;
  let lastDate = null;
  let lastSeq = 0;
  try {
    lastDate = localStorage.getItem('kensa_lastBusinessDate');
    lastSeq = parseInt(localStorage.getItem('kensa_lastBusinessSeq') || '0', 10);
  } catch {
    lastDate = null;
    lastSeq = 0;
  }

  let seq = 1;
  if (lastDate === dateStr) {
    seq = Number.isFinite(lastSeq) && lastSeq > 0 ? lastSeq + 1 : 1;
  } else {
    seq = 1;
  }

  try {
    localStorage.setItem('kensa_lastBusinessDate', dateStr);
    localStorage.setItem('kensa_lastBusinessSeq', String(seq));
  } catch {
    // Si falla el storage, igual devolvemos el ID generado.
  }

  return `${dateStr}-${String(seq).padStart(4, '0')}`;
}

// Catálogos normalizados para motivos y estados (renovaciones/recaudación)
const MOTIVOS_PERDIDA = [
  'Precio',
  'Cobertura',
  'Servicio',
  'Competidor',
  'No asegurable',
  'Documentación insuficiente',
  'Cliente no responde',
  'Cancelación voluntaria',
  'Fraude sospechado',
  'Duplicado/Error de alta',
];

const MOTIVOS_REVERSA = [
  'Pago duplicado',
  'Monto incorrecto',
  'Chargeback/desconocido',
  'Fraude',
  'Imputación errónea',
  'Anulación de póliza',
  'Devolución normativa',
  'Nota de crédito comercial',
  'Cheque protestado/transferencia revertida',
  'Conciliación mal registrada',
];

const MOTIVOS_ANULACION = [
  'Duplicado',
  'Datos críticos erróneos',
  'Alta por error/prueba',
  'Solicitud del cliente antes de gestión',
  'Rechazo de suscripción',
];

const ESTADOS_RENOVACION = [
  'Preaviso iniciado',
  'En gestión',
  'Propuesta enviada',
  'Negociación',
  'Renovado',
  'Renovado con modificación',
  'Perdido',
  'Anulado',
];

const ESTADOS_RECAUDACION = [
  'Registrado',
  'Cobrado',
  'Conciliado',
  'Aplicado',
  'Reversado/Devolución',
];

// Estado compartido para el editor "Interior del caso"
let currentId = null;
let ce_docs = [];
let ce_comments = [];

// Validadores simples para prevenir estados/motivos fuera de catálogo
const esMotivoPerdidaValido = (motivo) => MOTIVOS_PERDIDA.includes(motivo);
const esMotivoReversaValido = (motivo) => MOTIVOS_REVERSA.includes(motivo);
const esMotivoAnulacionValido = (motivo) => MOTIVOS_ANULACION.includes(motivo);
const esEstadoRenovacionValido = (estado) => ESTADOS_RENOVACION.includes(estado);
const esEstadoRecaudacionValido = (estado) => ESTADOS_RECAUDACION.includes(estado);

/**
 * Modelo conceptual de una renovación de póliza.
 * No se integra aún en state.tasks; queda listo para usar luego.
 * @param {Object} partial
 * @returns {Object}
 */
function crearRenovacion(partial = {}) {
  const base = {
    estadoRenovacion: '',
    fechaPreaviso: '',
    fechaPropuesta: '',
    fechaDecision: '',
    motivoPerdidaOAnulacion: '',
    comentario: '',
  };
  const r = { ...base, ...partial };
  return r;
}

/**
 * Modelo conceptual de pago/cuota para recaudación.
 * No se integra aún en state.tasks; queda listo para usar luego.
 * @param {Object} partial
 * @returns {Object}
 */
function crearPagoCuota(partial = {}) {
  const base = {
    estadoRecaudacion: '',
    metodoPago: '',
    fechaCobro: '',
    referenciaPago: '',
    conciliado: false,
    polizaNumero: '',
    cuotaNumero: '',
    motivoReversa: '',
  };
  const p = { ...base, ...partial };
  return p;
}

// Flag de depuración para ver cómo se incorporan renovaciones/pagos en tareas
const DEBUG_RENOVACION_DATOS = false;

// Normaliza estructuras de caso y tareas para compatibilidad con versiones previas
function ensureRenovacion(val) {
  if (!val) return null;
  const r = crearRenovacion(val);
  if (!r.estadoRenovacion && ESTADOS_RENOVACION.length)
    r.estadoRenovacion = ESTADOS_RENOVACION[0];
  return r;
}
function ensurePagos(list) {
  if (!Array.isArray(list)) return [];
  return list.map((p) => crearPagoCuota(p));
}
const renNeedsReason = (state) => ['Perdido', 'Anulado'].includes(state);
const recNeedsReason = (state) => state === 'Reversado/Devolución';
function pickLatestPago(pagos = []) {
  if (!Array.isArray(pagos) || !pagos.length) return null;
  // Ordena por fechaCobro descendente si existe; si no, toma el primero
  const sorted = pagos
    .slice()
    .sort((a, b) => (b.fechaCobro || '').localeCompare(a.fechaCobro || ''));
  return crearPagoCuota(sorted[0]);
}
function normalizeCase(caseData = {}) {
  const merged = {
    personal: caseData.personal || {},
    pago: caseData.pago || {},
    poliza: caseData.poliza || {},
    vehiculo: caseData.vehiculo || {},
    riesgo: caseData.riesgo || {},
    docs: (caseData.docs || []).slice(),
    comments: (caseData.comments || []).slice(),
    renovacion: ensureRenovacion(caseData.renovacion),
    pagos: ensurePagos(caseData.pagos || caseData.cuotas || []),
  };
  return { ...caseData, ...merged };
}
function normalizeTask(task = {}) {
  const defaults = {
    creadoPorUserId: null,
    creadoPorEmail: '',
    creadoPorRut: '',
    creadoPorRutCorredora: '',
    rolCreador: '',
    creadoPorRol: '',
    asignadoAUserId: null,
    asignadoAEmail: '',
    asignadoARut: '',
    asignadoARutCorredora: '',
    fueEditado: false,
  };
  const t = { ...defaults, ...task };
  t.fueEditado = task.fueEditado === true;
  t.case = normalizeCase(task.case || {});
  const creatorRole = normalizeRoleName(task.creadoPorRol || task.rolCreador || t.creadoPorRol || t.rolCreador);
  if (creatorRole) {
    t.creadoPorRol = creatorRole;
    t.rolCreador = creatorRole;
  }
  t.canalIngreso = normalizeChannel(task.canalIngreso || t.canalIngreso || '');
  t.status = normalizeTaskStatus(task.status);
  const providedInitial =
    typeof task.estadoInicial === 'string'
      ? task.estadoInicial
      : typeof t.estadoInicial === 'string'
      ? t.estadoInicial
      : '';
  if (providedInitial) {
    t.estadoInicial = normalizeInitialStatus(providedInitial, t.canalIngreso);
  } else {
    t.estadoInicial = inferInitialStatus({ status: t.status, canalIngreso: t.canalIngreso });
  }
  const poliza = t.case?.poliza || {};
  const policy = t.case?.policy || {};
  poliza.issueDate = poliza.issueDate || poliza.fechaEmisionPoliza || task.fechaEmisionPoliza || policy.issueDate || '';
  t.fechaEmisionPoliza = poliza.issueDate || t.fechaEmisionPoliza || '';
  const normalizedInsurer = (t.insurer || poliza.insurer || '').trim();
  t.insurer = normalizedInsurer || DEFAULT_INSURER;
  const normalizedType = (t.type || poliza.type || poliza.ramo || '').trim();
  t.type = normalizedType || DEFAULT_TASK_TYPE;
  const policyNumber =
    task.policyNumber ??
    t.policyNumber ??
    poliza.number ??
    policy.number ??
    '';
  poliza.number = policyNumber || poliza.number || '';
  t.policyNumber = policyNumber || '';
  const premiumUF = normalizeGoalNumber(
    poliza.premiumUF ??
      policy.grossPremiumUF ??
      policy.premiumUF ??
      t.policyPremiumUF ??
      task.policyPremiumUF
  );
  const commissionUF = normalizeGoalNumber(
    poliza.commissionUF ??
      poliza.commissionTotal ??
      poliza.policyCommissionTotalUF ??
      policy.brokerCommissionUF ??
      policy.brokerCommission ??
      policy.commissionUF ??
      t.policyCommissionTotalUF ??
      task.policyCommissionTotalUF
  );
  t.case.poliza.premiumUF = premiumUF;
  t.case.poliza.commissionUF = commissionUF;
  t.policyPremiumUF = premiumUF;
  t.policyCommissionTotalUF = commissionUF;
  t.isRenewal = !!task.isRenewal;
  t.parentBusinessId = task.parentBusinessId || null;
  t.autoRenewalCloned = !!task.autoRenewalCloned;
  t.autoRenewalClonedAt = task.autoRenewalClonedAt || null;
  t.hasRenewalGenerated = !!task.hasRenewalGenerated;
  t.activeRenewalId = task.activeRenewalId || null;
  const policyEndDate = normalizeDateOnly(
    task.policyEndDate ||
      task.due ||
      poliza.endDate ||
      poliza.policyEndDate ||
      (t.case?.renovacion?.fechaDecision || '')
  );
  t.policyEndDate = policyEndDate || null;
  t.due = t.policyEndDate;
  if (!t.updatedAt) t.updatedAt = t.createdAt || new Date().toISOString();
  if (!t.fechaActualizacion) t.fechaActualizacion = t.updatedAt;
  return t;
}

const getCurrentRutCorredora = () =>
  state.user?.brokerRut || state.user?.rutCorredora || '';

// === Auto-renovación de negocios (fase 1: localStorage, job simulado una vez al día) ===
function runAutoRenewalOncePerDay() {
  if (!state.tenant) return;
  const todayStr = toLocalDateString(new Date());
  if (!todayStr) return;
  const key = getAutoRenewalLastRunKey(state.tenant);
  const lastRun = localStorage.getItem(key);
  if (lastRun === todayStr) {
    console.log('[AutoRenewal] Ya ejecutado hoy para', state.tenant);
    return;
  }
  runAutoRenewalForTenant(state.tenant);
  localStorage.setItem(key, todayStr);
}
function runAutoRenewalForTenant(tenantId) {
  if (!tenantId) return;
  if (!Array.isArray(state.tasks) || !state.tasks.length) {
    console.log('[AutoRenewal] No hay tareas para revisar');
    return;
  }
  const nowIso = new Date().toISOString();
  const todayStr = toLocalDateString(new Date());
  if (!todayStr) return;
  const todayDate = new Date(`${todayStr}T00:00:00`);
  const rutCorredora = getCurrentRutCorredora();
  const snapshot = state.tasks.slice();
  let reviewed = 0;
  let clones = 0;
  snapshot.forEach((task) => {
    reviewed++;
    if (!task || task.isRenewal) return;
    if (task.hasRenewalGenerated) return;
    if (!task.policyEndDate) return;
    if (rutCorredora && task.rutCorredora && !sameRutCorredora(task.rutCorredora, rutCorredora))
      return;
    const policyEnd = new Date(`${task.policyEndDate}T00:00:00`);
    if (Number.isNaN(policyEnd.getTime())) return;
    const diffDays = Math.floor((policyEnd - todayDate) / 86400000);
    if (diffDays < 0 || diffDays > AUTO_RENEWAL_DAYS_BEFORE) return;
    const cloned = buildAutoRenewalClone(task);
    state.tasks.push(cloned);
    task.hasRenewalGenerated = true;
    task.activeRenewalId = cloned.id;
    task.updatedAt = nowIso;
    task.fechaActualizacion = nowIso;
    clones++;
  });
  console.log(
    `[AutoRenewal] Revisadas ${reviewed} tareas para ${tenantId}, renovaciones generadas: ${clones}`
  );
  if (clones > 0) {
    saveTasks();
    if (typeof applyFilters === 'function') applyFilters();
  }
}
function buildAutoRenewalClone(originalTask) {
  const nowIso = new Date().toISOString();
  const baseCase = normalizeCase(originalTask.case || {});
  let clonedCase;
  try {
    clonedCase = JSON.parse(JSON.stringify(baseCase));
  } catch {
    clonedCase = { ...baseCase };
  }
  clonedCase = clonedCase || {};
  const poliza = {
    ...(clonedCase.poliza || {}),
    number: '',
    premiumUF: null,
    commissionPct: null,
    commissionUF: null,
    agentCommission: null,
    endDate: originalTask.policyEndDate || (clonedCase.poliza && clonedCase.poliza.endDate) || '',
  };
  clonedCase.poliza = poliza;
  const newTask = normalizeTask({
    ...originalTask,
    id: uuid(),
    title: generateBusinessId(originalTask.type || DEFAULT_TASK_TYPE),
    createdAt: nowIso,
    updatedAt: nowIso,
    fechaActualizacion: nowIso,
    status: 'En renovación',
    canalIngreso: 'Renovación',
    isRenewal: true,
    parentBusinessId: originalTask.id,
    autoRenewalCloned: true,
    autoRenewalClonedAt: nowIso,
    hasRenewalGenerated: false,
    activeRenewalId: null,
    policyPremiumUF: 0,
    policyCommissionTotalUF: 0,
    policyEndDate: originalTask.policyEndDate || null,
    due: originalTask.policyEndDate || null,
    policy: '',
    case: clonedCase,
  });
  const comments = ensureCaseComments(newTask);
  comments.unshift({
    by: `${getActorName()} (auto)`,
    text: 'Negocio de renovación generado automáticamente.',
    ts: Date.now(),
  });
  return newTask;
}

let state = {
  tenant: null,
  user: null,
  ui: {
    isDashboardVisible: false,
    isCaseDetailOpen: false,
    filtersMode: 'groups',
  },
  filters: {
    group: 'nuevos',
  },
  tasks: [],
  users: [],
  clients: [],
  tasksLoaded: false,
  usersLoaded: false,
  clientsLoaded: false,
  filtered: [],
  selection: new Set(), // ids seleccionados
  goals: { userGoals: [], globalGoals: [] },
};

let negocioAuditLog = [];
let negocioAuditPageIndex = 0;
let negocioAuditPageSize = 15;
const DASHBOARD_VIEW_KEY = 'corredores_v1:dashboardViewOnly';
let dashboardViewOnly = false;
function updateFiltersCardVisibility() {
  const filtersCard = document.getElementById('filtersCard');
  const layout = document.getElementById('appLayout');
  if (!filtersCard) return;
  state.ui = state.ui || {};
  const isDashboard = !!state.ui.isDashboardVisible;
  const isCaseOpen = !!state.ui.isCaseDetailOpen;
  const hide = isDashboard || isCaseOpen;
  filtersCard.classList.toggle('hidden', hide);
  if (layout) layout.classList.toggle('no-filters', hide);
}
const DASHBOARD_EXECUTIVE_ALL = 'all';
let dashboardFilters = {
  range: 'currentMonth',
  customFrom: '',
  customTo: '',
  rutCorredora: '',
  executive: DASHBOARD_EXECUTIVE_ALL,
  company: 'all',
  product: 'all',
};
let dashboardCharts = {
  company: null,
  meta: null,
};
let kpiGranularityMode = 'auto';
const hasChartLibrary = () =>
  typeof window !== 'undefined' && typeof window.Chart !== 'undefined';

function showChartFallback(canvas, message) {
  if (!canvas) return;
  const container = canvas.parentElement;
  if (!container) return;
  let fallback = container.querySelector('.chart-fallback');
  if (!fallback) {
    fallback = document.createElement('div');
    fallback.className = 'chart-fallback muted';
    container.appendChild(fallback);
  }
  fallback.textContent = message || 'No se puede renderizar el gráfico.';
  canvas.style.display = 'none';
}

function hideChartFallback(canvas) {
  if (!canvas) return;
  const container = canvas.parentElement;
  const fallback = container?.querySelector('.chart-fallback');
  if (fallback) fallback.remove();
  canvas.style.display = '';
}

function canManageGlobalGoals() {
  const role = roleKey(state.user?.role || '');
  return role === ROLE_KEY_USUARIO_MAESTRO || role === ROLE_KEY_ADMINISTRADOR;
}

function applyJsonDownloadVisibility(currentUser) {
  const btnJson = document.getElementById('btnBackup');
  if (!btnJson) return;
  const roleName = normalizeRoleName(currentUser?.role || '');
  const isMaestro = roleName === ROLE_USUARIO_MAESTRO;
  btnJson.classList.toggle('hidden', !isMaestro);
  btnJson.style.display = isMaestro ? '' : 'none';
}

function syncDashboardGoalUi(options = {}) {
  const { forceHidePanel = false } = options;
  const actionsRow = document.getElementById('dashboardActionsRow');
  const goalPanel = document.getElementById('dashboardGoalPanel');
  const btnToggleGoal = document.getElementById('btnDashboardToggleGoal');
  const canManage = canManageGlobalGoals();
  const inDashboard = dashboardViewOnly;
  const showActions = inDashboard;
  if (actionsRow) actionsRow.classList.toggle('hidden', !showActions);
  if (btnToggleGoal) btnToggleGoal.classList.toggle('is-hidden', !(showActions && canManage));
  if (goalPanel) {
    if (!showActions || forceHidePanel) goalPanel.classList.add('is-hidden');
    else goalPanel.classList.remove('is-hidden');
  }
}

function openMetaModal() {
  const modal = document.getElementById('metaModal');
  const panel = document.getElementById('dashboardGoalPanel');
  if (!modal) return;
  if (panel) panel.classList.remove('is-hidden');
  modal.classList.remove('hidden');
  modal.setAttribute('aria-hidden', 'false');
  const focusField =
    document.getElementById('metaPeriodoDesde') ||
    document.getElementById('globalGoalPremium');
  focusField?.focus();
}

function closeMetaModal() {
  const modal = document.getElementById('metaModal');
  if (!modal) return;
  modal.classList.add('hidden');
  modal.setAttribute('aria-hidden', 'true');
}

const normalizeEmail = (v = '') => String(v || '').trim().toLowerCase();
const normalizeRutValue = (v = '') =>
  String(v || '')
    .replace(/\./g, '')
    .replace(/-/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
function esRutValido(rutStr) {
  if (!rutStr) return false;
  const rut = rutStr.trim().toUpperCase();
  const match = rut.match(/^(\d{7,8})-([\dK])$/);
  if (!match) return false;
  const cuerpo = match[1];
  const dv = match[2];
  let suma = 0;
  let factor = 2;
  for (let i = cuerpo.length - 1; i >= 0; i--) {
    suma += parseInt(cuerpo[i], 10) * factor;
    factor = factor === 7 ? 2 : factor + 1;
  }
  const resto = suma % 11;
  const dvCalculado =
    11 - resto === 11 ? '0' : 11 - resto === 10 ? 'K' : String(11 - resto);
  return dv === dvCalculado;
}
function esTelefonoValido(value) {
  if (!value) return false;
  const soloDigitos = value.replace(/\D/g, '');
  return soloDigitos.length === 9;
}
function esAnioVehiculoValido(value) {
  if (!value) return false;
  const soloDigitos = value.replace(/\D/g, '');
  if (soloDigitos.length !== 4) return false;
  const anio = parseInt(soloDigitos, 10);
  return !Number.isNaN(anio) && anio >= 2000;
}
function esPatenteValida(value) {
  if (!value) return false;
  const normalizado = value.trim().toUpperCase();
  const regex = /^[A-Z0-9]{6}$/;
  return regex.test(normalizado);
}
function attachRutSanitizer(el) {
  if (!el) return;
  el.addEventListener('input', () => {
    const raw = (el.value || '').toString();
    const cleaned = raw.replace(/[^0-9kK-]/g, '').toUpperCase();
    const firstHyphen = cleaned.indexOf('-');
    let normalized = cleaned.replace(/-/g, '');
    if (firstHyphen !== -1) {
      const cuerpo = normalized.slice(0, Math.max(0, normalized.length - 1));
      const dv = normalized.slice(-1);
      normalized = cuerpo && dv ? `${cuerpo}-${dv}` : cleaned;
    }
    el.value = normalized;
    el.classList.remove('input-error');
  });
}
function attachPhoneSanitizer(el) {
  if (!el) return;
  el.addEventListener('input', () => {
    const digits = (el.value || '').replace(/\D/g, '').slice(0, 9);
    el.value = digits;
    el.classList.remove('input-error');
  });
}
function attachYearSanitizer(el) {
  if (!el) return;
  el.addEventListener('input', () => {
    const digits = (el.value || '').replace(/\D/g, '').slice(0, 4);
    el.value = digits;
    el.classList.remove('input-error');
  });
}
function attachPlateSanitizer(el) {
  if (!el) return;
  el.addEventListener('input', () => {
    const normalized = (el.value || '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase().slice(0, 6);
    el.value = normalized;
    el.classList.remove('input-error');
  });
}
let toastTimeoutId = null;
function showSuccessToast() {
  const toast = document.getElementById('successToast');
  if (!toast) return;
  const checkPath = toast.querySelector('.toast-check-check');
  if (checkPath) {
    checkPath.style.animation = 'none';
    // force reflow to restart animation
    // eslint-disable-next-line no-unused-expressions
    checkPath.offsetHeight;
    checkPath.style.animation = null;
  }
  toast.classList.remove('hidden');
  if (toastTimeoutId) {
    clearTimeout(toastTimeoutId);
    toastTimeoutId = null;
  }
  toastTimeoutId = setTimeout(() => {
    toast.classList.add('hidden');
    toastTimeoutId = null;
  }, 500);
}

function resolveUserByEmail(email) {
  const lower = normalizeEmail(email);
  if (!lower) return null;
  return (state.users || []).find((u) => normalizeEmail(u.email) === lower) || null;
}

function buildIdentityFromUser(user = {}) {
  const email = user.email || '';
  const id = user.id || (email ? `email:${normalizeEmail(email)}` : null);
  const brokerRut = user.brokerRut || user.rutCorredora || '';
  const roleName = normalizeRoleName(user.role);
  return {
    id,
    email,
    name: user.name || email || 'Usuario',
    role: roleName || '',
    rut: user.rut || '',
    brokerRut,
    rutCorredora: brokerRut,
  };
}

function getCurrentUserIdentity() {
  const sessionUser = state.user || {};
  const resolved = resolveUserByEmail(sessionUser.email) || {};
  return buildIdentityFromUser({
    ...resolved,
    ...sessionUser,
  });
}

function applyCommissionFieldsVisibilityForRole(currentUser) {
  const totalComisionEl = document.getElementById('field-total-comision-corredor');
  const comisionAgenteEl = document.getElementById('field-comision-agente');
  if (!currentUser || !totalComisionEl || !comisionAgenteEl) return;
  const role = normalizeRoleName(currentUser.role || '');
  const ocultarParaRoles = ['Agente', 'Supervisor'];
  const hide = ocultarParaRoles.includes(role);
  const apply = (el) => {
    el.classList.toggle('hidden', hide);
    if (!el.classList.contains('hidden')) el.style.display = '';
  };
  apply(totalComisionEl);
  apply(comisionAgenteEl);
}

function buildCreatorMetadata(userIdentity) {
  const identity = buildIdentityFromUser(userIdentity || {});
  const roleName = normalizeRoleName(identity.role);
  return {
    creadoPorUserId: identity.id || null,
    creadoPorEmail: identity.email || '',
    creadoPorRut: identity.rut || '',
    creadoPorRutCorredora: identity.brokerRut || '',
    rolCreador: roleName || '',
    creadoPorRol: roleName || '',
  };
}

function buildAssigneeMetadata(email) {
  if (!email) {
    return {
      asignadoAUserId: null,
      asignadoAEmail: '',
      asignadoARut: '',
      asignadoARutCorredora: '',
    };
  }
  const resolved = resolveUserByEmail(email) || {};
  const identity = buildIdentityFromUser({ ...resolved, email });
  return {
    asignadoAUserId: identity.id || null,
    asignadoAEmail: identity.email || '',
    asignadoARut: identity.rut || '',
    asignadoARutCorredora: identity.brokerRut || '',
  };
}

function applyAssigneeMetadata(target, email) {
  Object.assign(target, buildAssigneeMetadata(email));
}

function ensureCreatorMetadata(target, creatorIdentity) {
  const meta = buildCreatorMetadata(creatorIdentity);
  if (!target.creadoPorUserId && meta.creadoPorUserId) target.creadoPorUserId = meta.creadoPorUserId;
  if (!target.creadoPorEmail && meta.creadoPorEmail) target.creadoPorEmail = meta.creadoPorEmail;
  if (!target.creadoPorRut && meta.creadoPorRut) target.creadoPorRut = meta.creadoPorRut;
  if (!target.creadoPorRutCorredora && meta.creadoPorRutCorredora)
    target.creadoPorRutCorredora = meta.creadoPorRutCorredora;
  if (!target.rolCreador && meta.rolCreador) target.rolCreador = meta.rolCreador;
  if (!target.creadoPorRol && meta.creadoPorRol) target.creadoPorRol = meta.creadoPorRol;
}

function getTaskVisibilityMeta(task = {}) {
  const creatorEmail = task.creadoPorEmail || '';
  const creatorUser = creatorEmail ? resolveUserByEmail(creatorEmail) : null;
  const assigneeEmail = task.asignadoAEmail || task.assignee || '';
  const assigneeUser = assigneeEmail ? resolveUserByEmail(assigneeEmail) : null;
  const creatorId =
    task.creadoPorUserId || creatorUser?.id || (creatorEmail ? `email:${normalizeEmail(creatorEmail)}` : null);
  const assigneeId =
    task.asignadoAUserId || assigneeUser?.id || (assigneeEmail ? `email:${normalizeEmail(assigneeEmail)}` : null);
  return {
    creator: {
      id: creatorId,
      email: creatorEmail || creatorUser?.email || '',
      rut: task.creadoPorRut || creatorUser?.rut || '',
      brokerRut: task.creadoPorRutCorredora || creatorUser?.brokerRut || creatorUser?.rutCorredora || '',
      role: getTaskCreatorRole(task),
    },
    assignee: {
      id: assigneeId,
      email: assigneeEmail || assigneeUser?.email || '',
      rut: task.asignadoARut || assigneeUser?.rut || '',
      brokerRut: task.asignadoARutCorredora || assigneeUser?.brokerRut || assigneeUser?.rutCorredora || '',
    },
  };
}

function sameRutCorredora(a, b) {
  const na = normalizeRutValue(a);
  const nb = normalizeRutValue(b);
  if (!na || !nb) return false;
  return na === nb;
}

function getTaskAssigneeEmail(task = {}) {
  return (
    task.assignee ||
    task.assigneeEmail ||
    task.asignadoAEmail ||
    task.assignedUserEmail ||
    task.asignadoA?.email ||
    ''
  );
}

function matchesUser(metaUser, user) {
  if (!metaUser || !user) return false;
  const targetEmail = normalizeEmail(user.email);
  const metaEmail = normalizeEmail(metaUser.email);
  const targetId = user.id || (targetEmail ? `email:${targetEmail}` : null);
  const metaId = metaUser.id || (metaEmail ? `email:${metaEmail}` : null);
  if (metaId && targetId && metaId === targetId) return true;
  if (metaEmail && targetEmail && metaEmail === targetEmail) return true;
  return false;
}

function getTaskCreatorRole(task) {
  return normalizeRoleName(task?.creadoPorRol || task?.rolCreador || '');
}

function resolveUserById(id) {
  if (!id) return null;
  return (state.users || []).find((u) => u.id === id) || null;
}

function resolveUserDisplay(id, email) {
  const byId = id ? resolveUserById(id) : null;
  if (byId) {
    return {
      id: byId.id || id || null,
      email: byId.email || email || '',
      name: byId.name || '',
    };
  }
  const byEmail = email ? resolveUserByEmail(email) : null;
  if (byEmail) {
    return {
      id: byEmail.id || id || null,
      email: byEmail.email || email || '',
      name: byEmail.name || '',
    };
  }
  return { id: id || null, email: email || '', name: '' };
}

function getAssignmentSnapshot(task = {}) {
  const meta = getTaskVisibilityMeta(task);
  const email = meta.assignee.email || task.assignee || '';
  const display = resolveUserDisplay(meta.assignee.id, email);
  return {
    id: display.id || null,
    email: display.email || '',
    name: display.name || '',
  };
}

function logAssignmentChange(task, prevSnapshot = {}) {
  if (!task) return;
  const nextSnapshot = getAssignmentSnapshot(task);
  const prevId =
    prevSnapshot.id || (prevSnapshot.email ? `email:${normalizeEmail(prevSnapshot.email)}` : null);
  const nextId =
    nextSnapshot.id || (nextSnapshot.email ? `email:${normalizeEmail(nextSnapshot.email)}` : null);
  const prevEmail = normalizeEmail(prevSnapshot.email);
  const nextEmail = normalizeEmail(nextSnapshot.email);
  if (prevId === nextId && prevEmail === nextEmail) return;

  const fromDisplay = resolveUserDisplay(prevSnapshot.id, prevSnapshot.email);
  const toDisplay = resolveUserDisplay(nextSnapshot.id, nextSnapshot.email);
  const actor = getCurrentUserIdentity();
  const detail = `Reasignado de ${fromDisplay.name || fromDisplay.email || 'Sin asignar'} a ${
    toDisplay.name || toDisplay.email || 'Sin asignar'
  }`;
  const entry = {
    timestamp: new Date().toISOString(),
    action: 'reasignacion',
    negocioId: task.id,
    negocioTitulo: task.title || task.client || '',
    fromUserId: fromDisplay.id || null,
    fromUserEmail: fromDisplay.email || '',
    fromUserName: fromDisplay.name || '',
    toUserId: toDisplay.id || null,
    toUserEmail: toDisplay.email || '',
    toUserName: toDisplay.name || '',
    actorUserId: actor.id || null,
    actorUserEmail: actor.email || '',
    actorUserName: actor.name || '',
    detalle: detail,
  };
  negocioAuditLog.push(entry);
  if (negocioAuditLog.length > 1000) negocioAuditLog.shift();
  saveNegocioAuditLog();
  negocioAuditPageIndex = 0;
}

function getAssigneeDisplay(task) {
  const meta = getTaskVisibilityMeta(task || {});
  const assigneeId = meta.assignee.id;
  const assigneeEmail = meta.assignee.email || task.assignee || '';
  const userById = assigneeId ? resolveUserById(assigneeId) : null;
  const userByEmail = assigneeEmail ? resolveUserByEmail(assigneeEmail) : null;
  const resolved = userById || userByEmail;
  if (resolved?.name) return resolved.name;
  if (meta.assignee.email) return meta.assignee.email;
  if (task.assignee) return task.assignee;
  return '';
}

function getInsurerDisplayName(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const key = normalizeInsurerKey(raw);
  const mapped = INSURER_DISPLAY_MAP.get(key);
  if (mapped) return mapped;
  const base = raw.replace(/\s*\([^)]*\)\s*$/, '').trim();
  if (base && base !== raw) {
    const baseKey = normalizeInsurerKey(base);
    const baseMapped = INSURER_DISPLAY_MAP.get(baseKey);
    if (baseMapped) return baseMapped;
  }
  return raw;
}

function normalizeSearchText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function getCompanySearchOptions(select) {
  if (!select || !select.options) return [];
  return Array.from(select.options)
    .map((opt) => {
      const value = typeof opt.value === 'string' ? opt.value.trim() : '';
      const label = (opt.textContent || opt.value || '').trim();
      if (!value && !label) return null;
      return {
        value,
        label,
        labelKey: normalizeSearchText(label),
        valueKey: normalizeSearchText(value),
      };
    })
    .filter(Boolean);
}

function scoreCompanyOption(option, query) {
  if (!query) return 0;
  const labelKey = option.labelKey || '';
  const valueKey = option.valueKey || '';
  const labelIndex = labelKey.indexOf(query);
  const valueIndex = valueKey.indexOf(query);
  let score = Number.POSITIVE_INFINITY;
  if (labelIndex === 0 || valueIndex === 0) score = 0;
  const wordIndex = labelKey.split(' ').findIndex((word) => word.startsWith(query));
  if (wordIndex >= 0) score = Math.min(score, 1 + wordIndex);
  if (labelIndex >= 0) score = Math.min(score, 2 + labelIndex);
  if (valueIndex >= 0) score = Math.min(score, 3 + valueIndex);
  return score;
}

function applyCompanySelection(select, input, list, option) {
  if (!select || !input || !option) return;
  select.value = option.value;
  input.value = option.label;
  if (list) list.classList.add('hidden');
  select.dispatchEvent(new Event('change', { bubbles: true }));
}

function updateCompanySearchList(select, input, list) {
  if (!select || !input || !list) return;
  const query = normalizeSearchText(input.value);
  if (!query) {
    list.innerHTML = '';
    list.classList.add('hidden');
    return;
  }
  const options = getCompanySearchOptions(select)
    .map((opt) => ({ ...opt, score: scoreCompanyOption(opt, query) }))
    .filter((opt) => Number.isFinite(opt.score));
  options.sort(
    (a, b) =>
      a.score - b.score ||
      a.label.length - b.label.length ||
      a.label.localeCompare(b.label, 'es', { sensitivity: 'base' })
  );
  const top = options.slice(0, 5);
  list.innerHTML = '';
  top.forEach((opt) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'company-search-option';
    btn.textContent = opt.label;
    btn.dataset.value = opt.value;
    btn.dataset.label = opt.label;
    list.appendChild(btn);
  });
  list.classList.toggle('hidden', top.length === 0);
}

function commitCompanySearchInput(select) {
  const input = select?._companySearchInput;
  const list = select?._companySearchList;
  if (!select || !input) return;
  const query = normalizeSearchText(input.value);
  if (!query) {
    syncCompanySearchInput(select);
    if (list) list.classList.add('hidden');
    return;
  }
  const options = getCompanySearchOptions(select);
  const exact = options.find(
    (opt) => opt.labelKey === query || opt.valueKey === query
  );
  if (exact) {
    applyCompanySelection(select, input, list, exact);
    return;
  }
  syncCompanySearchInput(select);
  if (list) list.classList.add('hidden');
}

function syncCompanySearchInput(select) {
  const input = select?._companySearchInput;
  if (!select || !input) return;
  const value = typeof select.value === 'string' ? select.value.trim() : '';
  const selectedLabel = select.options?.[select.selectedIndex]?.textContent?.trim() || '';
  const hideDefaultLabel = select.dataset?.hideDefaultLabel === 'true';
  const isDefault =
    normalizeInsurerKey(value) === normalizeInsurerKey(DEFAULT_INSURER) ||
    normalizeInsurerKey(selectedLabel) === normalizeInsurerKey(DEFAULT_INSURER);
  if (hideDefaultLabel && isDefault) {
    input.value = '';
    return;
  }
  if (!value) {
    input.value = selectedLabel || '';
    return;
  }
  const options = getCompanySearchOptions(select);
  const match = options.find(
    (opt) => opt.value === value || opt.label === value
  );
  input.value = match ? match.label : selectedLabel || value;
}

function setupCompanySearch(select) {
  if (!select || select.dataset.companySearchReady) return;
  const parent = select.parentElement;
  if (!parent) return;
  select.dataset.companySearchReady = 'true';
  if (select.id === 'insurer' || select.id === 'ce_taskInsurer') {
    select.dataset.hideDefaultLabel = 'true';
  }
  const wrapper = document.createElement('div');
  wrapper.className = 'company-search';
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'company-search-input';
  input.placeholder = 'Buscar compañía...';
  input.autocomplete = 'off';
  input.spellcheck = false;
  const list = document.createElement('div');
  list.className = 'company-search-list hidden';

  parent.insertBefore(wrapper, select);
  wrapper.appendChild(input);
  wrapper.appendChild(list);
  wrapper.appendChild(select);

  select.classList.add('company-search-select');
  select.setAttribute('aria-hidden', 'true');
  select.tabIndex = -1;
  select._companySearchInput = input;
  select._companySearchList = list;
  syncCompanySearchInput(select);

  input.addEventListener('input', () => updateCompanySearchList(select, input, list));
  input.addEventListener('focus', () => updateCompanySearchList(select, input, list));
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      const first = list.querySelector('.company-search-option');
      if (first) {
        applyCompanySelection(select, input, list, {
          value: first.dataset.value,
          label: first.dataset.label || first.textContent || '',
        });
      } else {
        commitCompanySearchInput(select);
      }
    } else if (event.key === 'Escape') {
      list.classList.add('hidden');
    }
  });
  input.addEventListener('blur', () => {
    setTimeout(() => commitCompanySearchInput(select), 120);
  });

  list.addEventListener('pointerdown', (event) => {
    const option = event.target.closest('.company-search-option');
    if (option) event.preventDefault();
  });
  list.addEventListener('click', (event) => {
    const option = event.target.closest('.company-search-option');
    if (!option) return;
    applyCompanySelection(select, input, list, {
      value: option.dataset.value || option.textContent || '',
      label: option.dataset.label || option.textContent || '',
    });
  });
}

function getNameInitials(value = '') {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const base = raw.includes('@') ? raw.split('@')[0].replace(/[._-]+/g, ' ') : raw;
  const parts = base.split(/\s+/).filter(Boolean);
  if (!parts.length) return '';
  const firstChars = Array.from(parts[0]);
  const first = firstChars[0] || '';
  let second = '';
  if (parts.length > 1) {
    const lastChars = Array.from(parts[parts.length - 1]);
    second = lastChars[0] || '';
  } else {
    second = firstChars[1] || '';
  }
  return `${first}${second}`.toUpperCase();
}

function getUserCommission(email) {
  if (!email) return 0;
  const lower = String(email).toLowerCase();
  const user = (state.users || []).find(
    (u) => (u.email || '').toLowerCase() === lower
  );
  return typeof user?.commission === 'number' ? user.commission : 0;
}

function getActorName() {
  return (state.user && (state.user.name || state.user.email)) || 'Usuario';
}

function ensureCaseComments(task) {
  if (!task) return [];
  if (!task.case) task.case = {};
  if (!Array.isArray(task.case.comments)) task.case.comments = [];
  return task.case.comments;
}

function pushBitacoraEntry(task, text) {
  if (!task) return null;
  const entry = { by: getActorName(), text, ts: Date.now() };
  const comments = ensureCaseComments(task);
  comments.unshift(entry);
  try {
    if (
      typeof ce_comments !== 'undefined' &&
      Array.isArray(ce_comments) &&
      typeof currentId !== 'undefined' &&
      currentId &&
      task.id &&
      String(task.id) === String(currentId)
    ) {
      ce_comments.unshift(entry);
      if (typeof ce_renderComments === 'function') ce_renderComments();
    }
  } catch {
    // Ignorar si el editor no está activo
  }
  return entry;
}

/* ===========================
   Sesión, carga/persistencia
=========================== */
function seedInitialMasterUser() {
  // Usa un tenant preferido si ya existe sesión; si no, recurre al fallback.
  let tenantId = DEFAULT_TENANT;
  try {
    const raw = localStorage.getItem('kensaCurrentTenant');
    const parsed = raw ? JSON.parse(raw) : null;
    if (parsed?.tenantId) tenantId = parsed.tenantId;
  } catch {
    tenantId = DEFAULT_TENANT;
  }

  const key = kUsers(tenantId);
  let list = [];
  try {
    const raw = localStorage.getItem(key);
    list = raw ? JSON.parse(raw) : [];
  } catch {
    list = [];
  }
  if (!Array.isArray(list) || list.length > 0) return;

  const now = new Date().toISOString();
  const masterUser = {
    id: uuid(),
    name: 'Juan Carlos Muñoz',
    email: MASTER_SEED_EMAIL,
    rut: '',
    brokerRut: '',
    role: ROLE_USUARIO_MAESTRO,
    commission: 50,
    status: 'activo',
    limiteUsuarios: null,
    creadoPorId: null,
    permissions: [
      'ver_clientes',
      'crear_clientes',
      'ver_polizas',
      'crear_polizas',
      'modificar_polizas',
      'ver_ventas',
      'registrar_ventas',
      'ver_reportes',
      'gestionar_usuarios',
    ],
    password: MASTER_SEED_PASSWORD,
    createdAt: now,
    updatedAt: now,
  };
  list.push(masterUser);
  localStorage.setItem(key, JSON.stringify(list));
}

seedInitialMasterUser();

function sanitizeUserForSession(user = {}) {
  const identity = buildIdentityFromUser(user);
  const rawLimit = user?.limiteUsuarios;
  const parsedLimit =
    typeof rawLimit === 'number'
      ? rawLimit
      : rawLimit === undefined || rawLimit === null || rawLimit === ''
        ? null
        : parseInt(rawLimit, 10);
  return {
    id: identity.id,
    name: identity.name,
    email: identity.email,
    role: identity.role || ROLE_ADMINISTRADOR,
    rut: identity.rut || '',
    brokerRut: identity.brokerRut || '',
    rutCorredora: identity.rutCorredora || '',
    limiteUsuarios: Number.isNaN(parsedLimit) || parsedLimit < 0 ? null : parsedLimit,
    creadoPorId: user.creadoPorId ?? null,
  };
}
function isUserLoggedIn() {
  return !!(state.user && state.user.email);
}
function applyShortcutsHintVisibility(isLoggedIn) {
  const hint = document.getElementById('shortcutsHint');
  if (!hint) return;
  hint.style.display = isLoggedIn ? '' : 'none';
}
function applyLoginBody(isLogin) {
  const body = document.body;
  if (!body) return;
  body.classList.toggle('login-body', !!isLogin);
  if (isLogin) {
    const loginRoot = document.getElementById('login');
    if (loginRoot) {
      loginRoot.classList.remove('login-animate');
      // Force reflow to restart animation when shown again
      // eslint-disable-next-line no-unused-expressions
      loginRoot.offsetWidth;
      loginRoot.classList.add('login-animate');
    }
  }
}
applyShortcutsHintVisibility(false);

function setSession(tenant, user) {
  state.tenant = tenant;
  state.user = sanitizeUserForSession(user);
  state.clients = [];
  state.clientsLoaded = false;

  localStorage.setItem(
    'kensaCurrentTenant',
    JSON.stringify({ tenantId: tenant, user: state.user })
  );

  const isMaestro =
    roleKey(state.user.role) === ROLE_KEY_USUARIO_MAESTRO ||
    /@kensa\.cl$/i.test(state.user.email || '') ||
    (state.user.email || '') === 'master@kensa.cl';
  localStorage.setItem(
    kSession(tenant),
    JSON.stringify({
      tenant,
      email: state.user.email,
      name: state.user.name,
      role: state.user.role || ROLE_ADMINISTRADOR,
      isMaestro,
      brokerRut: state.user.brokerRut,
      rut: state.user.rut,
      id: state.user.id,
      limiteUsuarios: state.user.limiteUsuarios ?? null,
      creadoPorId: state.user.creadoPorId ?? null,
    })
  );

  const env = document.getElementById('env');
  if (env) env.textContent = `https://${tenant}.kensa.local (demo)`;

  applyHeaderUserInfo(state.user);
}

function loadUsers() {
  try {
    const raw = JSON.parse(localStorage.getItem(kUsers(state.tenant)) || '[]');
    state.users = Array.isArray(raw)
      ? raw.map((u) => ({ ...u, role: normalizeRoleName(u.role) }))
      : [];
  } catch {
    state.users = [];
  }
  state.usersLoaded = true;
}
function saveUsers() {
  localStorage.setItem(kUsers(state.tenant), JSON.stringify(state.users));
}

function loadClients(targetTenant = state.tenant) {
  const tenantId = targetTenant || state.tenant;
  let list = [];
  if (!tenantId) {
    state.clients = [];
    state.clientsLoaded = true;
    return [];
  }
  try {
    const raw = localStorage.getItem(kClients(tenantId));
    const parsed = raw ? JSON.parse(raw) : [];
    list = Array.isArray(parsed) ? parsed : [];
  } catch {
    list = [];
  }
  if (tenantId === state.tenant) {
    state.clients = list;
    state.clientsLoaded = true;
  }
  return list;
}

function saveClients(targetTenant = state.tenant) {
  const tenantId = targetTenant || state.tenant;
  if (!tenantId) return;
  localStorage.setItem(kClients(tenantId), JSON.stringify(state.clients || []));
}

function getClients(targetTenant = state.tenant) {
  const tenantId = targetTenant || state.tenant;
  if (!tenantId) return [];
  if (tenantId === state.tenant && state.clientsLoaded) {
    return (state.clients || []).slice();
  }
  return loadClients(tenantId);
}

function saveClientsList(clients, targetTenant = state.tenant) {
  const tenantId = targetTenant || state.tenant;
  const payload = Array.isArray(clients) ? clients.slice() : [];
  if (tenantId === state.tenant) {
    state.clients = payload;
    state.clientsLoaded = true;
  }
  if (tenantId) {
    localStorage.setItem(kClients(tenantId), JSON.stringify(payload));
  }
  return payload;
}

function createClient(clientData = {}, targetTenant = state.tenant) {
  const tenantId = targetTenant || state.tenant;
  const list = tenantId === state.tenant ? getClients(tenantId) : loadClients(tenantId);
  const now = new Date().toISOString();
  const payload = {
    id: clientData.id || uuid(),
    nombreCorredora: (clientData.nombreCorredora || '').trim(),
    rutCorredora: (clientData.rutCorredora || '').trim(),
    estado: clientData.estado === 'inactivo' ? 'inactivo' : 'activo',
    logoDataUrl: clientData.logoDataUrl || null,
    createdAt: now,
    updatedAt: now,
  };
  list.push(payload);
  saveClientsList(list, tenantId);
  return payload;
}

function updateClient(clientId, changes = {}, targetTenant = state.tenant) {
  if (!clientId) return null;
  const tenantId = targetTenant || state.tenant;
  const list = tenantId === state.tenant ? getClients(tenantId) : loadClients(tenantId);
  const idx = list.findIndex((c) => c.id === clientId);
  if (idx === -1) return null;
  const next = {
    ...list[idx],
    ...changes,
    updatedAt: new Date().toISOString(),
  };
  list[idx] = next;
  saveClientsList(list, tenantId);
  return next;
}

function deleteClient(clientId, targetTenant = state.tenant) {
  if (!clientId) return false;
  const tenantId = targetTenant || state.tenant;
  const list = tenantId === state.tenant ? getClients(tenantId) : loadClients(tenantId);
  const next = list.filter((c) => c.id !== clientId);
  if (next.length === list.length) return false;
  saveClientsList(next, tenantId);
  return true;
}

function findClientByRut(rutCorredora, targetTenant = state.tenant) {
  const targetRut = normalizeRutValue(rutCorredora || '');
  if (!targetRut) return null;
  const list = getClients(targetTenant);
  return (
    list.find((c) => normalizeRutValue(c.rutCorredora || '') === targetRut) || null
  );
}

function loadTasks() {
  try {
    const raw = JSON.parse(localStorage.getItem(kTasks(state.tenant)) || '[]');
    const nowIso = new Date().toISOString();
    state.tasks = Array.isArray(raw)
      ? raw.map((task) => {
          const normalized = normalizeTask(task);
          if (!normalized.fechaActualizacion) {
            const fallback = normalized.fechaCreacion || normalized.createdAt || nowIso;
            normalized.fechaActualizacion = fallback;
            normalized.updatedAt = normalized.updatedAt || fallback;
          }
          return normalized;
        })
      : [];
  } catch {
    state.tasks = [];
  }
  state.tasksLoaded = true;
}
function saveTasks() {
  localStorage.setItem(kTasks(state.tenant), JSON.stringify(state.tasks));
}

function loadNegocioAuditLog() {
  if (!state.tenant) {
    negocioAuditLog = [];
    return;
  }
  try {
    const raw = localStorage.getItem(kTaskAuditLog(state.tenant));
    const arr = raw ? JSON.parse(raw) : [];
    negocioAuditLog = Array.isArray(arr) ? arr : [];
    negocioAuditPageIndex = 0;
  } catch {
    negocioAuditLog = [];
    negocioAuditPageIndex = 0;
  }
}
function saveNegocioAuditLog() {
  if (!state.tenant) return;
  localStorage.setItem(kTaskAuditLog(state.tenant), JSON.stringify(negocioAuditLog));
}

/* ===========================
   UI auxiliares
=========================== */
function populateAssignees() {
  const dl = document.getElementById('assigneeList');
  if (!dl) return;
  const current = getCurrentUserIdentity();
  const role = roleKey(current.role);
  // Política de asignación:
  // - Usuario Maestro / roles sin restricción: ven a todos.
  // - Administrador y Supervisor: solo pueden asignar dentro de su misma corredora (mismo RUT).
  // - Agente: no asigna (lista vacía y UI oculta en otro punto).
  let list = (state.users || []).filter((u) => u.active !== false);
  if (role === ROLE_KEY_SUPERVISOR || role === ROLE_KEY_ADMINISTRADOR) {
    const rut = current.brokerRut || current.rutCorredora || '';
    list = list.filter((u) =>
      sameRutCorredora(u.brokerRut || u.rutCorredora || '', rut)
    );
  } else if (role === ROLE_KEY_AGENTE) {
    dl.innerHTML = '';
    return;
  }
  dl.innerHTML = list
    .map((u) => `<option value="${u.email}">${u.name || u.email} (${u.role || '–'})</option>`)
    .join('');
}

function applyClientesButtonVisibility(currentUser) {
  const btnClientes = document.getElementById('btnClientes');
  if (!btnClientes || !currentUser) return;
  if (normalizeRoleName(currentUser.role || '') === ROLE_USUARIO_MAESTRO) {
    btnClientes.style.display = '';
    btnClientes.classList.remove('hidden');
  } else {
    btnClientes.style.display = 'none';
    btnClientes.classList.add('hidden');
  }
}

function applyRoleUiVisibility() {
  const role = roleKey(state.user?.role || '');
  applyClientesButtonVisibility(state.user || {});
  const canSeeUsers = role === ROLE_KEY_USUARIO_MAESTRO || role === ROLE_KEY_ADMINISTRADOR;
  const btnUsers = document.getElementById('btnUsers');
  if (btnUsers) btnUsers.classList.toggle('hidden', !canSeeUsers);

  const isAgent = role === ROLE_KEY_AGENTE;
  const bulkAssignBtn = document.getElementById('btnBulkAssign');
  if (bulkAssignBtn) {
    bulkAssignBtn.classList.toggle('hidden', isAgent);
    bulkAssignBtn.disabled = isAgent;
  }

  const ceAssignee = document.getElementById('ce_taskAssignee');
  const ceAssigneeLabel = ceAssignee?.closest('label');
  if (ceAssigneeLabel) ceAssigneeLabel.style.display = isAgent ? 'none' : '';

  const assignField = document.getElementById('assignee');
  const assignWrap = assignField?.closest('label');
  if (assignWrap) assignWrap.style.display = isAgent ? 'none' : '';

  applyJsonDownloadVisibility(state.user || {});

  syncDashboardGoalUi();

  const importBtn = document.getElementById('importButton');
  if (importBtn) {
    const hideImport = role === ROLE_KEY_ADMINISTRADOR || role === ROLE_KEY_SUPERVISOR || role === ROLE_KEY_AGENTE;
    importBtn.classList.toggle('hidden', hideImport);
    importBtn.disabled = hideImport;
  }
  const importInput = document.getElementById('importFileInput');
  if (importInput) importInput.disabled = role === ROLE_KEY_ADMINISTRADOR || role === ROLE_KEY_SUPERVISOR || role === ROLE_KEY_AGENTE;
}

function applyCorredoraLogo(currentUser) {
  const logoEl = document.getElementById('logoCorredora');
  const divider = document.querySelector('.header-divider');
  if (!logoEl) return;
  const rutCorredora = currentUser?.rutCorredora || currentUser?.brokerRut || '';
  const client = rutCorredora ? findClientByRut(rutCorredora) : null;
  const hasLogo = !!(client && client.logoDataUrl);
  if (hasLogo) {
    logoEl.src = client.logoDataUrl;
    logoEl.classList.remove('hidden');
  } else {
    logoEl.removeAttribute('src');
    logoEl.classList.add('hidden');
  }
  if (divider) {
    divider.classList.toggle('hidden', !hasLogo);
  }
}

function applyHeaderUserInfo(currentUser) {
  const roleTitleEl = document.getElementById('headerUserRole');
  const nameEl = document.getElementById('headerUserName');
  const chip = document.getElementById('userChip');
  if (roleTitleEl) roleTitleEl.textContent = 'USUARIO';
  const nameValue = currentUser?.name || currentUser?.email || 'Usuario';
  if (nameEl) nameEl.textContent = nameValue;
  if (chip) chip.textContent = nameValue;
}

/* ===========================
   Render de tabla de tareas
=========================== */
function renderTable() {
  const tbody = document.getElementById('tbody');
  if (!tbody) return;

  const formatDateTimeCell = (value) => {
    const d = value ? new Date(value) : null;
    if (!(d instanceof Date) || Number.isNaN(d.getTime()))
      return '<span style="color:var(--muted)">–</span>';
    const dateStr = d.toLocaleDateString();
    const timeStr = d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: false });
    return `<span>${dateStr}</span><div class="negocio-prioridad">${timeStr}</div>`;
  };

  const rows = state.filtered.map((t) => {
    const tagsMarkup = (t.tags || [])
      .map((x) => `<span class="chip">${x}</span>`)
      .join(' ');
    const tagItems = (t.tags || [])
      .filter((tag) => tag != null && String(tag).trim())
      .flatMap((tag) => {
        const value = String(tag).trim();
        if (!value) return [];
        const compact = value.replace(/\s+/g, '');
        const starMatches = compact.match(/\u2B50\uFE0F?/g);
        const starOnly = starMatches && starMatches.join('') === compact;
        if (starOnly) return starMatches;
        return [value];
      });
    const mobileTagsStack = tagItems.length
      ? `<div class="mobile-tags-stack">${tagItems
          .map((tag) => `<span class="mobile-tag-item">${tag}</span>`)
          .join('')}</div>`
      : '';

    const checked = state.selection.has(String(t.id)) ? 'checked' : '';
    const createdCell = formatDateTimeCell(t.createdAt);
    const updatedCell = formatDateTimeCell(t.fechaActualizacion || t.updatedAt || t.createdAt);
    const policyNumber = t.policyNumber || (t.case?.poliza?.number ?? '') || '';
    const negocioId = t.title || t.id || '—';
    const prioridadLabel = t.priority || '—';
    const personal = t.case?.personal || {};
    const firstName = personal.firstName || (t.client?.split(' ')[0] ?? t.client ?? '');
    const lastName =
      personal.lastName ||
      (t.client && t.client.includes(' ') ? t.client.split(' ').slice(1).join(' ') : '');
    const clientFirstLine = firstName || t.client || '';
    const clientSecondLine = lastName || personal.rut || t.rut || '';
    const policyMeta = t.type || (!policyNumber ? t.title || '' : '');
    const mobileStack = `
      <div class="mobile-case-stack">
        <div class="mobile-case-line mobile-case-priority">${prioridadLabel}</div>
        <div class="mobile-case-line mobile-case-first">${clientFirstLine}</div>
        <div class="mobile-case-line mobile-case-last">${clientSecondLine}</div>
      </div>
    `;
    const assigneeDisplay = getAssigneeDisplay(t);
    const assigneeInitials = getNameInitials(assigneeDisplay);
    const assigneeMarkup = assigneeDisplay ? `<span class="assignee-full">${assigneeDisplay}</span>` : '';
    const assigneeActionMarkup = assigneeInitials
      ? `<span class="assignee-initials" aria-hidden="true">${assigneeInitials}</span>`
      : '';
    const statusText = t.status || '';
    const mobileStatusMarkup = statusText ? `<span class="mobile-status">${statusText}</span>` : '';

    return `<tr data-id="${t.id}">
      <td><input type="checkbox" class="rowCheck" ${checked} />${mobileTagsStack}</td>
      <td>
        <div class="negocio-id"><span>${negocioId}</span></div>
        <div class="negocio-prioridad"><span class="negocio-prioridad-label">Prioridad:</span> ${prioridadLabel}</div>
        ${mobileStack}
      </td>
      <td><span>${policyNumber}</span><div class="negocio-prioridad">${policyMeta}</div></td>
      <td>${getInsurerDisplayName(t.insurer || '')}</td>
      <td>${clientFirstLine}<div class="negocio-prioridad">${clientSecondLine}</div></td>
      <td>${assigneeMarkup}</td>
      <td>${createdCell}</td>
      <td>${t.due ? new Date(t.due + 'T00:00:00').toLocaleDateString() : ''}</td>
      <td>${t.status || ''}</td>
      <td>${tagsMarkup}</td>
      <td>${updatedCell}</td>
      <td>
        <div class="mobile-detail-actions">
          ${assigneeActionMarkup}
          <button class="btn secondary btnEdit" aria-label="Gestionar negocio">→</button>
        </div>
        ${mobileStatusMarkup}
      </td>
    </tr>`;
  }).join('');

  tbody.innerHTML =
    rows ||
    '<tr><td colspan="12" style="text-align:center;color:var(--muted);padding:1rem">Sin negocios</td></tr>';

  const countLabel = document.getElementById('countLabel');
  if (countLabel) {
    countLabel.textContent = `${state.filtered.length} ${
      state.filtered.length === 1 ? 'negocio' : 'negocios'
    }`;
  }

  // Reaplicar paginación y sincronizar checkbox/acciones
  if (typeof window.refreshTasksPagination === 'function') window.refreshTasksPagination();
  if (typeof window.__syncHeaderCheckbox === 'function') window.__syncHeaderCheckbox();
}

function getTasksVisibleForCurrentUser(source = state.tasks) {
  const currentUser = getCurrentUserIdentity();
  if (!state.user) return (source || []).slice();
  const role = roleKey(currentUser.role);
  const rutCorredora = currentUser.brokerRut || currentUser.rutCorredora || '';
  const currentId = currentUser.id || (currentUser.email ? `email:${normalizeEmail(currentUser.email)}` : null);
  const matchIdentity = { ...currentUser, id: currentId };
  const isUsuarioMaestro = role === ROLE_KEY_USUARIO_MAESTRO;
  const isAdministrador = role === ROLE_KEY_ADMINISTRADOR;
  const isSupervisor = role === ROLE_KEY_SUPERVISOR;
  const isAgente = role === ROLE_KEY_AGENTE;

  // Visibilidad por rol:
  // - Usuario Maestro ve todo, pero el resto oculta negocios creados por Usuario Maestro salvo que estén asignados a ellos.
  // - Administrador: propios + Supervisores/Agentes de su corredora + cualquier negocio asignado.
  // - Supervisor: creados por Supervisor/Agente de su corredora + cualquier negocio asignado al supervisor.
  // - Agente: solo creados por él o asignados a él.
  // La metadata creadoPor*/asignadoA* es la fuente de verdad para estas reglas.
  let arr = (source || []).slice();

  if (!isUsuarioMaestro) {
    arr = arr.filter((t) => {
      const creatorRole = roleKey(getTaskCreatorRole(t));
      if (creatorRole !== ROLE_KEY_USUARIO_MAESTRO) return true;
      const meta = getTaskVisibilityMeta(t);
      return matchesUser(meta.assignee, matchIdentity);
    });
  }

  if (isAdministrador) {
    arr = arr.filter((t) => {
      const meta = getTaskVisibilityMeta(t);
      const creatorRole = roleKey(getTaskCreatorRole(t));
      const creatorMatches = matchesUser(meta.creator, matchIdentity);
      const assignedMatches = matchesUser(meta.assignee, matchIdentity);
      const sameRutCreator = sameRutCorredora(meta.creator.brokerRut, rutCorredora);
      const creatorSupOrAgent = creatorRole === ROLE_KEY_SUPERVISOR || creatorRole === ROLE_KEY_AGENTE;
      return creatorMatches || assignedMatches || (sameRutCreator && creatorSupOrAgent);
    });
  } else if (isSupervisor) {
    arr = arr.filter((t) => {
      const meta = getTaskVisibilityMeta(t);
      if (matchesUser(meta.assignee, matchIdentity)) return true;
      const creatorRole = roleKey(getTaskCreatorRole(t));
      return (
        sameRutCorredora(meta.creator.brokerRut, rutCorredora) &&
        (creatorRole === ROLE_KEY_SUPERVISOR || creatorRole === ROLE_KEY_AGENTE)
      );
    });
  } else if (isAgente) {
    arr = arr.filter((t) => {
      const meta = getTaskVisibilityMeta(t);
      return matchesUser(meta.assignee, matchIdentity) || matchesUser(meta.creator, matchIdentity);
    });
  }

  return arr;
}

/* ===========================
   Filtros
=========================== */
function applyFilters() {
  const $ = (s) => document.querySelector(s);

  const qval = ($('#q')?.value || '').trim().toLowerCase();
  const tv = $('#fTipo')?.value || '';
  const ev = $('#fEstado')?.value || '';
  const pv = $('#fPrioridad')?.value || '';
  const vv = $('#fVenc')?.value || '';
  const iv = $('#fCompania')?.value || '';
  const tg = ($('#fTags')?.value || '').trim();
  const onlyMine = $('#fSoloMias')?.checked || false;

  const today = todayISO();
  state.filters = state.filters || { group: 'nuevos' };
  const visibleTasks = getTasksVisibleForCurrentUser();
  let normalizedTimestamps = false;
  visibleTasks.forEach((t) => {
    if (!t.updatedAt) {
      t.updatedAt = t.createdAt || new Date().toISOString();
      normalizedTimestamps = true;
    }
    if (!t.fechaActualizacion) {
      t.fechaActualizacion = t.updatedAt;
      normalizedTimestamps = true;
    }
  });
  if (normalizedTimestamps) saveTasks();
  let arr = visibleTasks.slice();

  const visibleIds = new Set(visibleTasks.map((t) => String(t.id)));
  state.selection.forEach((id) => {
    if (!visibleIds.has(String(id))) state.selection.delete(id);
  });

  if (qval) {
    arr = arr.filter((t) => {
      const poliza = t.case?.poliza || {};
      const personal = t.case?.personal || {};
      const vehiculo = t.case?.vehiculo || {};
      const riesgo = t.case?.riesgo || {};
      const haystack = [
        t.title,
        t.client,
        t.policy,
        t.rut,
        poliza.number,
        personal.phone,
        personal.email,
        vehiculo.plate,
        riesgo.address,
      ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase();
      return haystack.includes(qval);
    });
  }
  if (tv) arr = arr.filter((t) => t.type === tv);
  if (ev)
    arr = arr.filter((t) => normalizeTaskStatus(t.status) === normalizeTaskStatus(ev));
  if (pv) arr = arr.filter((t) => t.priority === pv);
  if (iv) arr = arr.filter((t) => t.insurer === iv);
  if (tg)
    arr = arr.filter((t) => (t.tags || []).some((x) => (x || '').toLowerCase() === tg.toLowerCase()));
  if (onlyMine)
    arr = arr.filter(
      (t) => (t.assignee || '').toLowerCase() === (state.user?.email || '').toLowerCase()
    );
  const tasksForGroupCounts = arr.slice();
  if (state.filters.group) {
    const allowed = FILTER_GROUPS[state.filters.group] || [];
    if (allowed.length) {
      const normalizedAllowed = allowed.map((status) => normalizeTaskStatus(status));
      arr = arr.filter((t) => normalizedAllowed.includes(normalizeTaskStatus(t.status)));
    }
  }

  if (vv) {
    if (vv === 'vencidas')
      arr = arr.filter((t) => {
        const status = normalizeTaskStatus(t.status);
        return !STATUS_WON.has(status) && !STATUS_LOST.has(status) && t.due && t.due < today;
      });
    else if (vv === 'hoy') arr = arr.filter((t) => t.due === today);
    else {
      const end = new Date(new Date().getTime() + Number(vv) * 86400000);
      arr = arr.filter(
        (t) => t.due && t.due >= today && new Date(t.due) <= end
      );
    }
  }

  const getUpdateStamp = (t) =>
    new Date(t.fechaActualizacion || t.updatedAt || t.createdAt || 0).getTime();
  arr.sort((a, b) => getUpdateStamp(b) - getUpdateStamp(a));
  state.filtered = arr;
  renderTable();
  populateCompanyProductOptions(
    document.getElementById('dashFilterCompany'),
    document.getElementById('dashFilterProduct'),
    { preserveSelection: true }
  );

  if (dashboardViewOnly) {
    refreshDashboard();
  }
  computeAndLogKpis(state.filtered);
  updateFilterGroupCounts(tasksForGroupCounts);
}

function updateFilterGroupUI() {
  const pills = document.querySelectorAll('.filter-group-pill');
  if (!pills || !pills.length) return;
  const active = (state.filters && state.filters.group) || null;
  pills.forEach((pill) => {
    const key = pill.getAttribute('data-group');
    if (key && active === key) pill.classList.add('active');
    else pill.classList.remove('active');
  });
}

function initFilterGroupsBehavior() {
  const container = document.querySelector('.filter-groups-grid');
  if (!container) return;
  const pills = container.querySelectorAll('.filter-group-pill');
  if (!pills.length) return;
  state.filters = state.filters || { group: 'nuevos' };
  pills.forEach((pill) => {
    pill.addEventListener('click', () => {
      const key = pill.getAttribute('data-group');
      if (!key) return;
      state.filters.group = state.filters.group === key ? null : key;
      updateFilterGroupUI();
      applyFilters();
    });
  });
  updateFilterGroupUI();
  updateFilterGroupCounts();
}

function initFiltersModeUI() {
  const btnGroups = document.getElementById('filtersModeGroups');
  const btnFilters = document.getElementById('filtersModeFilters');
  const groupsSection = document.getElementById('filtersGroupsSection');
  const formSection = document.getElementById('filtersFormSection');
  if (!btnGroups || !btnFilters || !groupsSection || !formSection) return;
  state.ui = state.ui || {};
  const applyMode = () => {
    const mode = state.ui.filtersMode === 'filters' ? 'filters' : 'groups';
    if (mode === 'groups') {
      btnGroups.classList.add('active');
      btnFilters.classList.remove('active');
      groupsSection.classList.remove('hidden');
      formSection.classList.add('hidden');
    } else {
      btnGroups.classList.remove('active');
      btnFilters.classList.add('active');
      groupsSection.classList.add('hidden');
      formSection.classList.remove('hidden');
    }
  };
  btnGroups.addEventListener('click', () => {
    state.ui.filtersMode = 'groups';
    state.filters = state.filters || { group: 'nuevos' };
    if (!state.filters.group) {
      state.filters.group = 'nuevos';
      updateFilterGroupUI();
      applyFilters();
    }
    applyMode();
  });
  btnFilters.addEventListener('click', () => {
    state.ui.filtersMode = 'filters';
    state.filters = state.filters || { group: null };
    if (state.filters.group) {
      state.filters.group = null;
      updateFilterGroupUI();
      applyFilters();
    }
    applyMode();
  });
  if (!state.ui.filtersMode) state.ui.filtersMode = 'groups';
  applyMode();
}

function computeFilterGroupCounts(tasks = []) {
  const counts = {};
  const normalizedGroups = {};
  Object.entries(FILTER_GROUPS).forEach(([key, statuses]) => {
    normalizedGroups[key] = statuses.map((status) => normalizeTaskStatus(status));
    counts[key] = 0;
  });
  tasks.forEach((task) => {
    const normalizedStatus = normalizeTaskStatus(task.status);
    Object.entries(normalizedGroups).forEach(([groupKey, statusList]) => {
      if (statusList.includes(normalizedStatus)) counts[groupKey]++;
    });
  });
  return counts;
}

function updateFilterGroupCounts(sourceTasks) {
  const tasks = Array.isArray(sourceTasks) ? sourceTasks : getTasksVisibleForCurrentUser();
  const counts = computeFilterGroupCounts(tasks);
  document.querySelectorAll('[data-group-count]').forEach((el) => {
    const key = el.getAttribute('data-group-count');
    if (!key) return;
    el.textContent = counts[key] ?? 0;
  });
}

function computeKpis(tasks = []) {
  let renovadas = 0;
  let perdidas = 0;
  let anuladas = 0;
  let pagosReversados = 0;

  tasks.forEach((t) => {
    const c = normalizeCase(t.case || {});
    const ren = c.renovacion;
    const estRen = ren?.estadoRenovacion || '';
    if (estRen === 'Renovado' || estRen === 'Renovado con modificación') renovadas++;
    else if (estRen === 'Perdido') perdidas++;
    else if (estRen === 'Anulado') anuladas++;

    const pagos = Array.isArray(c.pagos) ? c.pagos : [];
    pagos.forEach((p) => {
      if ((p.estadoRecaudacion || '') === 'Reversado/Devolución') pagosReversados++;
    });
  });

  const totalRenovaciones = renovadas + perdidas + anuladas;
  const retencion =
    totalRenovaciones > 0 ? renovadas / totalRenovaciones : 0;

  return { renovadas, perdidas, anuladas, totalRenovaciones, retencion, pagosReversados };
}

function computeAndLogKpis(tasks) {
  const k = computeKpis(tasks);
  if (typeof console !== 'undefined') {
    console.log(
      '[KPIs] Renovadas:',
      k.renovadas,
      'Pérdidas:',
      k.perdidas,
      'Anuladas:',
      k.anuladas,
      'Retención:',
      (k.retencion * 100).toFixed(1) + '%',
      'Pagos reversados:',
      k.pagosReversados
    );
  }
}

// ===========================
// Dashboard KPIs – filtros y cálculos
// ===========================
const DASHBOARD_RANGE_PRESETS = [
  { value: 'last30', label: 'Últimos 30 días' },
  { value: 'currentMonth', label: 'Mes actual' },
  { value: 'last3', label: 'Últimos 3 meses' },
  { value: 'last6', label: 'Últimos 6 meses' },
  { value: 'last12', label: 'Últimos 12 meses' },
  { value: 'year', label: 'Año calendario' },
  { value: 'custom', label: 'Rango personalizado' },
];

function initDashboardFiltersUI() {
  const rangeSelect = document.getElementById('dashFilterRange');
  if (!rangeSelect) return;
  rangeSelect.innerHTML = '';
  DASHBOARD_RANGE_PRESETS.forEach((preset) => {
    const opt = document.createElement('option');
    opt.value = preset.value;
    opt.textContent = preset.label;
    rangeSelect.appendChild(opt);
  });
  if (!dashboardFilters.range) dashboardFilters.range = 'currentMonth';
  rangeSelect.value = dashboardFilters.range;
  const execSelect = document.getElementById('dashFilterExecutive');
  const companySelect = document.getElementById('dashFilterCompany');
  const productSelect = document.getElementById('dashFilterProduct');
  const customGroup = document.getElementById('dashCustomRangeGroup');
  const customFrom = document.getElementById('dashCustomFrom');
  const customTo = document.getElementById('dashCustomTo');

  const currentRut = getCurrentRutCorredora() || '';
  dashboardFilters.rutCorredora = currentRut;
  populateExecutiveOptions(execSelect, currentRut);
  populateCompanyProductOptions(companySelect, productSelect);
  const syncCustomInputs = () => {
    if (!customGroup || !customFrom || !customTo) return;
    const isCustom = dashboardFilters.range === 'custom';
    customGroup.classList.toggle('visible', isCustom);
    if (isCustom) {
      if (!dashboardFilters.customFrom || !dashboardFilters.customTo) {
        const today = new Date();
        const defaultFrom = new Date(today);
        defaultFrom.setDate(defaultFrom.getDate() - 30);
        if (!dashboardFilters.customFrom) {
          dashboardFilters.customFrom = formatDateInputValue(defaultFrom);
        }
        if (!dashboardFilters.customTo) {
          dashboardFilters.customTo = formatDateInputValue(today);
        }
      }
      customFrom.value = dashboardFilters.customFrom || '';
      customTo.value = dashboardFilters.customTo || '';
    } else {
      customFrom.value = '';
      customTo.value = '';
    }
  };
  execSelect?.addEventListener('change', () => {
    dashboardFilters.executive = execSelect.value;
    refreshDashboard();
  });
  companySelect?.addEventListener('change', () => {
    dashboardFilters.company = companySelect.value;
    refreshDashboard();
  });
  productSelect?.addEventListener('change', () => {
    dashboardFilters.product = productSelect.value;
    refreshDashboard();
  });
  rangeSelect.addEventListener('change', () => {
    dashboardFilters.range = rangeSelect.value;
    syncCustomInputs();
    refreshDashboard();
  });
  customFrom?.addEventListener('change', () => {
    dashboardFilters.customFrom = customFrom.value;
    refreshDashboard();
  });
  customTo?.addEventListener('change', () => {
    dashboardFilters.customTo = customTo.value;
    refreshDashboard();
  });
  syncCustomInputs();
  initGranularityControls();
  refreshDashboard();
}

function populateSelect(select, options, selected) {
  if (!select) return null;
  select.innerHTML = '';
  let selectionApplied = false;
  options.forEach((opt) => {
    const option = document.createElement('option');
    option.value = opt.value;
    option.textContent = opt.label;
    if (
      !selectionApplied &&
      selected !== undefined &&
      selected !== null &&
      opt.value === selected
    ) {
      option.selected = true;
      selectionApplied = true;
    }
    select.appendChild(option);
  });
  if (!selectionApplied && select.options.length > 0) {
    select.options[0].selected = true;
  }
  return select.value || null;
}

function populateExecutiveOptions(select, rut, opts = {}) {
  const { preserveSelection = false } = opts;
  if (!select) return;
  const role = roleKey(state.user?.role || '');
  const normalizedRut = normalizeRutValue(rut);
  const userList = Array.isArray(state.users) ? state.users : [];
  const users = normalizedRut
    ? userList.filter((user) =>
        sameRutCorredora(user.brokerRut || user.rutCorredora || '', normalizedRut)
      )
    : userList.slice();
  let options = [{ value: DASHBOARD_EXECUTIVE_ALL, label: 'Todos' }];
  if (role === ROLE_KEY_AGENTE && state.user?.email) {
    options = [{ value: state.user.email, label: state.user.name || state.user.email }];
    dashboardFilters.executive = state.user.email;
    select.disabled = true;
  } else {
    const execOptions = users.map((u) => ({
      value: u.email,
      label: u.name || u.email,
    }));
    options = options.concat(execOptions);
    const hasPrev =
      preserveSelection &&
      options.some((opt) => opt.value === dashboardFilters.executive);
    dashboardFilters.executive = hasPrev ? dashboardFilters.executive : DASHBOARD_EXECUTIVE_ALL;
    select.disabled = false;
  }
  const applied = populateSelect(select, options, dashboardFilters.executive);
  if (applied) dashboardFilters.executive = applied;
}

function getSelectOptionValues(id) {
  const select = document.getElementById(id);
  if (!select || !select.options) return [];
  return Array.from(select.options)
    .map((opt) => (opt.value || opt.textContent || '').trim())
    .filter(Boolean);
}

function getKnownInsurerValues() {
  const base = new Set();
  getSelectOptionValues('ce_taskInsurer').forEach((val) => base.add(val));
  getSelectOptionValues('insurer').forEach((val) => base.add(val));
  if (!Array.from(base).some((val) => val.toLowerCase() === DEFAULT_INSURER.toLowerCase())) {
    base.add(DEFAULT_INSURER);
  }
  return Array.from(base);
}

function getKnownTypeValues() {
  const base = new Set(Object.keys(TYPE_CODES || {}));
  getSelectOptionValues('ce_taskType').forEach((val) => base.add(val));
  getSelectOptionValues('type').forEach((val) => base.add(val));
  if (!base.size) base.add(DEFAULT_TASK_TYPE);
  return Array.from(base);
}

// Fix: filtros de compañía y tipo unificados con catálogos del interior del caso
function populateCompanyProductOptions(companySelect, productSelect, opts = {}) {
  const { preserveSelection = false } = opts;
  const rutRaw = getCurrentRutCorredora() || '';
  const currentRut = normalizeRutValue(rutRaw);
  const rutFilter = currentRut || null;
  dashboardFilters.rutCorredora = rutRaw;
  const sourceTasks = Array.isArray(state.tasks) ? state.tasks : [];
  const tasks = rutFilter ? sourceTasks.filter((t) => taskMatchesRut(t, rutFilter)) : sourceTasks.slice();
  const baseInsurers = getKnownInsurerValues();
  const baseInsurerKeys = new Set(baseInsurers.map((val) => val.toLowerCase()));
  const dynamicInsurers = [];
  tasks
    .map((t) => (t.insurer || '').trim())
    .filter(Boolean)
    .forEach((val) => {
      const key = val.toLowerCase();
      if (!baseInsurerKeys.has(key)) {
        baseInsurerKeys.add(key);
        dynamicInsurers.push(val);
      }
    });
  dynamicInsurers.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
  const companyValues = baseInsurers.slice();
  dynamicInsurers.forEach((val) => companyValues.push(val));

  const baseTypes = getKnownTypeValues();
  const baseTypeKeys = new Set(baseTypes.map((val) => val.toLowerCase()));
  const dynamicTypes = [];
  tasks
    .map((t) => (t.type || '').trim())
    .filter(Boolean)
    .forEach((val) => {
      const key = val.toLowerCase();
      if (!baseTypeKeys.has(key)) {
        baseTypeKeys.add(key);
        dynamicTypes.push(val);
      }
    });
  dynamicTypes.sort((a, b) => a.localeCompare(b, 'es', { sensitivity: 'base' }));
  const productValues = baseTypes.slice();
  dynamicTypes.forEach((val) => productValues.push(val));
  const companyOptions = [
    { value: 'all', label: 'Todas' },
    ...companyValues.map((c) => ({ value: c, label: c })),
  ];
  const productOptions = [
    { value: 'all', label: 'Todos' },
    ...productValues.map((p) => ({ value: p, label: p })),
  ];
  const currentCompanyValid =
    preserveSelection &&
    companyOptions.some((opt) => opt.value === dashboardFilters.company);
  const currentProductValid =
    preserveSelection &&
    productOptions.some((opt) => opt.value === dashboardFilters.product);
  const companyValue = currentCompanyValid ? dashboardFilters.company : 'all';
  const productValue = currentProductValid ? dashboardFilters.product : 'all';
  dashboardFilters.company = companyValue;
  dashboardFilters.product = productValue;
  const appliedCompany = populateSelect(companySelect, companyOptions, companyValue);
  const appliedProduct = populateSelect(productSelect, productOptions, productValue);
  if (appliedCompany) dashboardFilters.company = appliedCompany;
  if (appliedProduct) dashboardFilters.product = appliedProduct;
  syncCompanySearchInput(companySelect);
}

function taskMatchesRut(task, rut) {
  const target = normalizeRutValue(rut);
  if (!target) return true;
  const visibility = task?.visibility || {};
  const candidates = [
    visibility.rutCorredora,
    visibility.brokerRut,
    task.rutCorredora,
    task.brokerRut,
  ];
  const matchedCandidate = candidates
    .map((val) => normalizeRutValue(val))
    .find((val) => val && val === target);
  if (matchedCandidate) return true;

  const meta = getTaskVisibilityMeta(task);
  const creatorRut = normalizeRutValue(meta.creator?.brokerRut || meta.creator?.rutCorredora);
  const assigneeRut = normalizeRutValue(meta.assignee?.brokerRut || meta.assignee?.rutCorredora);
  if (creatorRut && creatorRut === target) return true;
  if (assigneeRut && assigneeRut === target) return true;

  const hasAnyRut =
    matchedCandidate ||
    creatorRut ||
    assigneeRut ||
    candidates.some((val) => normalizeRutValue(val));
  if (!hasAnyRut) {
    // Tareas legacy sin RUT explícito se incluyen solo cuando no se filtra por RUT.
    return false;
  }
  return false;
}

function getDashboardDateBounds() {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  let from = new Date(today);
  switch (dashboardFilters.range) {
    case 'last30':
      from.setDate(from.getDate() - 30);
      break;
    case 'currentMonth':
      from = new Date(today.getFullYear(), today.getMonth(), 1);
      break;
    case 'last3':
      from = new Date(today.getFullYear(), today.getMonth() - 3, today.getDate());
      break;
    case 'last6':
      from = new Date(today.getFullYear(), today.getMonth() - 6, today.getDate());
      break;
    case 'last12':
      from = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      break;
    case 'year':
      from = new Date(today.getFullYear(), 0, 1);
      break;
    case 'custom': {
      const customFrom = dashboardFilters.customFrom
        ? new Date(dashboardFilters.customFrom)
        : new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30);
      const customTo = dashboardFilters.customTo
        ? new Date(dashboardFilters.customTo)
        : today;
      return {
        from: new Date(customFrom.getFullYear(), customFrom.getMonth(), customFrom.getDate()),
        to: new Date(customTo.getFullYear(), customTo.getMonth(), customTo.getDate(), 23, 59, 59),
      };
    }
    default:
      from.setDate(from.getDate() - 30);
      break;
  }
  return { from, to: today };
}

function ensureDashboardSourcesLoaded() {
  if (!state.tenant) return;
  if (!state.usersLoaded) loadUsers();
  if (!state.tasksLoaded) loadTasks();
}

function getFilteredDealsForDashboard() {
  const currentRut = normalizeRutValue(getCurrentRutCorredora());
  const rutFilter = currentRut || null;
  ensureDashboardSourcesLoaded();
  const { from, to } = getDashboardDateBounds();
  const executiveFilter =
    dashboardFilters.executive !== DASHBOARD_EXECUTIVE_ALL &&
    dashboardFilters.executive
      ? normalizeEmail(dashboardFilters.executive)
      : null;
  const data = (state.tasks || []).filter((task) => {
    if (rutFilter && !taskMatchesRut(task, rutFilter)) return false;
    const stamp = new Date(task.updatedAt || task.createdAt || Date.now());
    if (stamp < from || stamp > to) return false;
    if (executiveFilter) {
      const assigneeEmail = normalizeEmail(getTaskAssigneeEmail(task));
      if (assigneeEmail !== executiveFilter) return false;
    }
    if (
      dashboardFilters.company !== 'all' &&
      (task.insurer || '').trim() !== dashboardFilters.company
    )
      return false;
    if (
      dashboardFilters.product !== 'all' &&
      (task.type || '').trim() !== dashboardFilters.product
    )
      return false;
    return true;
  });
  const won = data.filter((task) => STATUS_WON.has(normalizeTaskStatus(task.status)));
  console.log('[DASHBOARD] filtros', {
    rutFilter,
    dateBounds: { from: from.toISOString(), to: to.toISOString() },
    executive: dashboardFilters.executive,
    company: dashboardFilters.company,
    product: dashboardFilters.product,
    totalTasks: state.tasks?.length,
    filteredCount: data.length,
  });
  return { all: data, won };
}


function buildDashboardGoalContext() {
  const scope = obtenerScopeMetaParaUsuarioActual();
  const metas = seedMetaHistoryFromLegacy(scope);
  return {
    ...scope,
    metas,
    resolve(period) {
      const compare = normalizeGoalPeriod(period);
      const meta = obtenerMetaVigenteParaFecha(scope.tipo, scope.scopeId, compare, metas);
      if (meta) return meta;
      const rut = getCurrentRutCorredora() || '';
      const legacy = rut ? getGlobalGoalForPeriod(rut, compare) : null;
      if (legacy) {
        return {
          ...legacy,
          tipo: scope.tipo || 'corredora',
          scopeId: scope.scopeId ?? rut ?? null,
          periodoDesde: compare,
          periodoHasta: null,
          metaPrimaUF: legacy.goalPremiumUF,
          metaComisionUF: legacy.goalCommissionUF,
        };
      }
      return null;
    },
  };
}

// Fix Dashboard KPIs: Prima emitida y Comisión corredor usan datos de Póliza
function getPolicyPremiumValue(task = {}) {
  if (!task) return 0;
  const premium =
    task.case?.poliza?.premiumUF ??
    task.case?.policy?.grossPremiumUF ??
    task.case?.policy?.premiumUF ??
    task.policyPremiumUF;
  return normalizeGoalNumber(premium);
}
function getPolicyCommissionValue(task = {}) {
  if (!task) return 0;
  const commission =
    task.case?.poliza?.commissionUF ??
    task.case?.policy?.brokerCommissionUF ??
    task.case?.policy?.brokerCommission ??
    task.case?.policy?.commissionUF ??
    task.case?.policy?.commissionTotalUF ??
    task.policyCommissionTotalUF;
  return normalizeGoalNumber(commission);
}
function getUserCommissionMultiplierForTask(task = {}) {
  const assigneeEmail = normalizeEmail(getTaskAssigneeEmail(task));
  if (!assigneeEmail) return 1;
  const user = resolveUserByEmail(assigneeEmail);
  if (!user) return 1;
  const role = roleKey(user.role);
  const pct = Number.isFinite(user.commission) ? Number(user.commission) : null;
  if (role === ROLE_KEY_AGENTE && pct !== null) {
    return pct / 100;
  }
  return 1;
}
function getDashboardCommissionValue(task = {}, { debug = false } = {}) {
  const baseCommissionUF = getPolicyCommissionValue(task);
  const multiplier = getUserCommissionMultiplierForTask(task);
  const commissionUF = baseCommissionUF * multiplier || 0;
  if (debug && typeof console !== 'undefined') {
    console.log('[DASHBOARD] comisión negocio', {
      taskId: task.id,
      assignee: getTaskAssigneeEmail(task),
      baseCommissionUF,
      commissionPercent: multiplier * 100,
      commissionUF,
    });
  }
  return commissionUF;
}
function calcPrimaEmitidaUF(deals = []) {
  return deals.reduce((sum, task) => sum + getPolicyPremiumValue(task), 0);
}
function calcComisionCorredorUF(deals = []) {
  return deals.reduce((sum, task) => sum + getDashboardCommissionValue(task, { debug: true }), 0);
}

function computeDashboardKpis(dealsData, goalContext) {
  const wonDeals = Array.isArray(dealsData?.won) ? dealsData.won : [];
  const allDeals = Array.isArray(dealsData?.all) ? dealsData.all : [];
  const prima = calcPrimaEmitidaUF(wonDeals);
  const commission = calcComisionCorredorUF(wonDeals);
  const dealsWon = wonDeals.length;
  const totalDeals = allDeals.length;
  const totalCreated = allDeals.length;
  const metaPeriod = getGoalPeriodForFilters();
  const goal = goalContext?.resolve(metaPeriod);
  const metaPrima = goal?.metaPrimaUF ?? goal?.goalPremiumUF ?? 0;
  const metaCommission = goal?.metaComisionUF ?? goal?.goalCommissionUF ?? 0;
  const metaDeals = goal?.goalDeals;
  console.log(
    '[DASHBOARD] deals usados:',
    wonDeals.length,
    wonDeals.slice(0, 3)
  );
  console.log('[DASHBOARD] Prima emitida UF:', prima);
  console.log('[DASHBOARD] Comision corredor UF:', commission);
  const primaPctRaw = metaPrima > 0 ? (prima / metaPrima) * 100 : null;
  const commissionPctRaw = metaCommission > 0 ? (commission / metaCommission) * 100 : null;
  const primaPct = primaPctRaw !== null ? Number(primaPctRaw.toFixed(1)) : null;
  const commissionPct = commissionPctRaw !== null ? Number(commissionPctRaw.toFixed(1)) : null;
  return {
    prima,
    commission,
    metaPrima,
    metaCommission,
    metaDeals,
    dealsWon,
    totalDeals,
    totalCreated,
    primaPct,
    commissionPct,
  };
}

function renderDashboardKpis(kpis) {
  const fmt = (val) =>
    typeof val === 'number'
      ? val.toLocaleString('es-CL', { maximumFractionDigits: 1 })
      : '—';
  const primaValue = document.getElementById('kpiPrimaValue');
  const primaSubtitle = document.getElementById('kpiPrimaSubtitle');
  const commissionValue = document.getElementById('kpiCommissionValue');
  const commissionSubtitle = document.getElementById('kpiCommissionSubtitle');
  const commissionPct = document.getElementById('kpiCommissionPct');
  const commissionPctSubtitle = document.getElementById('kpiCommissionPctSubtitle');
  const dealsValue = document.getElementById('kpiDealsValue');
  const dealsSubtitle = document.getElementById('kpiDealsSubtitle');

  if (primaValue) primaValue.textContent = fmt(kpis.prima);
  if (primaSubtitle) {
    if (kpis.metaPrima > 0) {
      const diff = ((kpis.prima / kpis.metaPrima) * 100 - 100).toFixed(1);
      primaSubtitle.textContent = `${diff >= 0 ? '+' : ''}${diff}% vs meta`;
    } else {
      primaSubtitle.textContent = 'Sin meta definida';
    }
  }
  if (commissionValue) commissionValue.textContent = fmt(kpis.commission);
  if (commissionSubtitle) {
    if (kpis.metaCommission > 0) {
      const diff = ((kpis.commission / kpis.metaCommission) * 100 - 100).toFixed(1);
      commissionSubtitle.textContent = `${diff >= 0 ? '+' : ''}${diff}% vs meta`;
    } else {
      commissionSubtitle.textContent = 'Sin meta definida';
    }
  }
  if (commissionPct) commissionPct.textContent = kpis.commissionPct ? `${kpis.commissionPct}%` : '—';
  if (commissionPctSubtitle) {
    if (kpis.metaCommission <= 0) {
      commissionPctSubtitle.textContent = 'Sin meta registrada';
    } else {
      const pct = typeof kpis.commissionPct === 'number' ? kpis.commissionPct : parseFloat(kpis.commissionPct) || 0;
      commissionPctSubtitle.textContent = pct >= 100 ? 'Sobre meta de comisión' : 'Bajo meta de comisión';
    }
  }
  if (dealsValue) dealsValue.textContent = kpis.dealsWon.toLocaleString('es-CL');
  if (dealsSubtitle) {
    const baseTotal = kpis.totalCreated ?? kpis.totalDeals ?? 0;
    const pctTotal =
      baseTotal > 0 ? Number(((kpis.dealsWon / baseTotal) * 100).toFixed(1)) : 0;
    dealsSubtitle.textContent = `${pctTotal}% sobre negocios creados`;
  }
}

function getTaskLossReason(task = {}) {
  const direct = (task.motivoPerdida || '').trim();
  if (direct) return direct;
  return (task.case?.renovacion?.motivoPerdidaOAnulacion || '').trim();
}

function buildSaveOptionsForTask(task, allowedStates = []) {
  const options = [];
  const motivo = getTaskLossReason(task);
  const hasMotivo = !!motivo;

  if (hasMotivo) {
    const cierreStates = allowedStates.filter((s) => s === 'Desistida' || s === 'Pérdida');
    cierreStates.forEach((estado) => {
      options.push({ type: 'estado', value: estado, label: estado });
    });
    if (!options.length) {
      options.push({ type: 'solo_guardar', value: SAVE_OPTION_SOLO_GUARDAR, label: 'Solo guardar' });
    }
  } else {
    const nonLossStates = allowedStates.filter((s) => s !== 'Desistida' && s !== 'Pérdida');
    nonLossStates.forEach((estado) => {
      options.push({ type: 'estado', value: estado, label: estado });
    });
    options.push({ type: 'solo_guardar', value: SAVE_OPTION_SOLO_GUARDAR, label: 'Solo guardar' });
  }

  if (!options.length) {
    options.push({ type: 'solo_guardar', value: SAVE_OPTION_SOLO_GUARDAR, label: 'Solo guardar' });
  }
  return options;
}

function openSaveOptionsModal(task, allowedStates = [], onDecision) {
  const options = buildSaveOptionsForTask(task, allowedStates);
  const picker = window.__openPicker || window.openPicker;
  if (!options || !options.length || typeof picker !== 'function') {
    // Fallback simple si no existe picker
    const labels = options.map((o) => o.label).join(', ');
    const choice = window.prompt(`Guardar como: ${labels}`);
    if (choice && typeof onDecision === 'function') onDecision(choice);
    return;
  }
  picker(
    'Guardar como',
    options.map((opt) => ({ label: opt.label, value: opt.value })),
    (selected) => {
      if (selected && typeof onDecision === 'function') {
        onDecision(selected.value ?? selected);
      }
    },
    { confirm: true }
  );
}

function prepareChartCanvas(canvas, fallbackHeight = 320) {
  if (!canvas) return null;
  const parent = canvas.parentElement;
  const width = parent?.clientWidth || canvas.clientWidth || 400;
  const height = parent?.clientHeight || fallbackHeight;
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

function renderCompanyChart(deals) {
  const ctx = document.getElementById('chartPrimaCompany');
  if (!ctx) return;
  const existingTooltip = document.getElementById('chart-company-tooltip');
  if (existingTooltip) existingTooltip.classList.add('hidden');
  if (!hasChartLibrary()) {
    showChartFallback(ctx, 'Chart.js no disponible');
    return;
  }
  hideChartFallback(ctx);
  prepareChartCanvas(ctx);
  const getOrCreateCompanyTooltip = () => {
    let tooltipEl = document.getElementById('chart-company-tooltip');
    if (!tooltipEl) {
      tooltipEl = document.createElement('div');
      tooltipEl.id = 'chart-company-tooltip';
      tooltipEl.className = 'chart-tooltip hidden';
      document.body.appendChild(tooltipEl);
    }
    return tooltipEl;
  };
  const companyTooltip = (context) => {
    const tooltip = context.tooltip;
    const tooltipEl = getOrCreateCompanyTooltip();
    if (!tooltip || tooltip.opacity === 0) {
      tooltipEl.classList.add('hidden');
      return;
    }
    tooltipEl.classList.remove('hidden');
    const title = (tooltip.title && tooltip.title[0]) || '';
    const rows = (tooltip.dataPoints || []).map((dp) => {
      const val = Number(dp.raw || 0);
      const formatted = Number.isFinite(val)
        ? val.toLocaleString('es-CL', { maximumFractionDigits: 1 })
        : '0';
      const color =
        dp.dataset?.borderColor ||
        dp.dataset?.backgroundColor ||
        'rgba(59,130,246,1)';
      return `<div class="chart-tooltip-row">
        <span class="chart-tooltip-dot" style="background:${color}"></span>
        <span class="chart-tooltip-label">${dp.dataset?.label || ''}</span>
        <strong>${formatted}</strong>
      </div>`;
    });
    tooltipEl.innerHTML = `
      <div class="chart-tooltip-title">${title}</div>
      ${rows.join('')}
    `;
    const { left, top } = context.chart.canvas.getBoundingClientRect();
    tooltipEl.style.left = `${left + window.scrollX + tooltip.caretX}px`;
    tooltipEl.style.top = `${top + window.scrollY + tooltip.caretY - 12}px`;
  };
  const aggregate = {};
  deals.forEach((deal) => {
    const raw = (deal.insurer || '').trim();
    const display = getInsurerDisplayName(raw);
    const key = display || 'Sin compañía';
    if (!aggregate[key]) aggregate[key] = { prima: 0, commission: 0 };
    aggregate[key].prima += getPolicyPremiumValue(deal);
    aggregate[key].commission += getDashboardCommissionValue(deal);
  });
  const entries = Object.entries(aggregate).sort((a, b) => b[1].prima - a[1].prima);
  const labels = entries.map(([k]) => k);
  const primaData = entries.map(([, val]) => val.prima);
  const commissionData = entries.map(([, val]) => val.commission);
  if (dashboardCharts.company) dashboardCharts.company.destroy();
  dashboardCharts.company = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Prima Emitida', data: primaData, backgroundColor: 'rgba(59,130,246,.7)' },
        { label: 'Comisión', data: commissionData, backgroundColor: 'rgba(16,185,129,.7)' },
      ],
    },
    options: {
      responsive: false,
      animation: false,
      maintainAspectRatio: false,
      resizeDelay: 0,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        tooltip: { enabled: false, external: companyTooltip },
        legend: { display: true },
      },
      scales: { y: { beginAtZero: true } },
    },
  });
}

function renderFunnel(dealsData) {
  const container = document.getElementById('funnelContainer');
  if (!container) return;
  container.innerHTML = '';
  const tooltip = document.createElement('div');
  tooltip.className = 'funnel-tooltip hidden';
  container.appendChild(tooltip);

  const positionTooltip = (layer) => {
    const layerRect = layer.getBoundingClientRect();
    const containerRect = container.getBoundingClientRect();
    const left = layerRect.left - containerRect.left + layerRect.width / 2;
    const top = layerRect.top - containerRect.top;
    tooltip.style.left = `${left}px`;
    tooltip.style.top = `${top}px`;
  };
  const showTooltip = (layer, text) => {
    tooltip.textContent = text;
    positionTooltip(layer);
    tooltip.classList.remove('hidden');
  };
  const hideTooltip = () => {
    tooltip.classList.add('hidden');
  };
  const abiertas = dealsData.all.filter(
    (task) => normalizeTaskStatus(task.status) === 'Abierta'
  ).length;
  const enRenovacion = dealsData.all.filter(
    (task) => normalizeTaskStatus(task.status) === 'En renovación'
  ).length;
  const enGestion = dealsData.all.filter(
    (task) => normalizeTaskStatus(task.status) === 'Propuesta enviada'
  ).length;
  const perdidos = dealsData.all.filter((task) =>
    STATUS_LOST.has(normalizeTaskStatus(task.status))
  ).length;
  const steps = [
    { key: 'created', label: 'Negocios creados', value: dealsData.all.length },
    { key: 'renewal', label: 'En renovación', value: enRenovacion },
    { key: 'inProgress', label: 'En gestión', value: enGestion },
    { key: 'won', label: 'Ganados', value: dealsData.won.length },
    { key: 'lost', label: 'Perdidos', value: perdidos },
  ];
  const totalCreated = steps[0].value || 0;
  steps.forEach((step, index) => {
    const base = totalCreated || 1;
    const pct = base > 0 ? (step.value / base) * 100 : 0;
    const layer = document.createElement('div');
    layer.className = `funnel-layer funnel-layer-${index}`;
    layer.innerHTML = `
      <div class="funnel-layer-content">
        <span class="funnel-label">${step.label}</span>
        <span class="funnel-value">${step.value}</span>
        <span class="funnel-percent">${pct.toFixed(1)}%</span>
      </div>
    `;
    if (step.key === 'created') {
      const detail = `Abierta: ${abiertas} · En renovación: ${enRenovacion}`;
      layer.addEventListener('mouseenter', () => showTooltip(layer, detail));
      layer.addEventListener('mousemove', () => positionTooltip(layer));
      layer.addEventListener('mouseleave', hideTooltip);
    }
    container.appendChild(layer);
  });
}

const formatShortDateLabel = (date) =>
  date.toLocaleDateString('es-CL', { day: '2-digit', month: '2-digit' });

function startOfWeek(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0 domingo, 1 lunes
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
function endOfWeek(date) {
  const start = startOfWeek(date);
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  return end;
}
function startOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}
function endOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59);
}

function getEffectiveGranularityMode(fromDate, toDate) {
  if (!(fromDate instanceof Date) || !(toDate instanceof Date)) return 'day';
  const diffMs = toDate.getTime() - fromDate.getTime();
  const days = Math.max(1, Math.floor(diffMs / (1000 * 60 * 60 * 24)) + 1);
  if (kpiGranularityMode === 'day' || kpiGranularityMode === 'week' || kpiGranularityMode === 'month') {
    return kpiGranularityMode;
  }
  if (days <= 13) return 'day';
  if (days >= 14 && days <= 49) return 'week';
  return 'month';
}
function updateKpiGranularityLabel(mode) {
  const labelEl = document.getElementById('kpiGranularityLabel');
  if (!labelEl) return;
  const map = {
    day: 'Detalle por día',
    week: 'Detalle por semana',
    month: 'Detalle por mes',
    auto: 'Detalle (auto)',
  };
  labelEl.textContent = map[mode] || map.auto;
}
function updateGranularityButtonsActive(mode) {
  const buttons = document.querySelectorAll('.granularity-btn');
  if (!buttons || !buttons.length) return;
  buttons.forEach((btn) => {
    const btnMode = btn.getAttribute('data-mode');
    btn.classList.toggle('granularity-active', btnMode === mode);
  });
}
function initGranularityControls() {
  const container = document.querySelector('.kpi-granularity-toggle');
  if (!container) return;
  container.addEventListener('click', (event) => {
    const btn = event.target.closest('.granularity-btn');
    if (!btn) return;
    const mode = btn.getAttribute('data-mode');
    if (!mode || mode === kpiGranularityMode) return;
    kpiGranularityMode = mode;
    updateGranularityButtonsActive(kpiGranularityMode);
    refreshDashboard();
  });
  updateGranularityButtonsActive(kpiGranularityMode);
}

function getBucketIdentity(date, mode) {
  const base = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  if (mode === 'week') {
    const start = startOfWeek(base);
    const end = endOfWeek(start);
    return {
      key: `${start.toISOString().slice(0, 10)}_${end.toISOString().slice(0, 10)}`,
      start,
      end,
      label: `${formatShortDateLabel(start)} - ${formatShortDateLabel(end)}`,
    };
  }
  if (mode === 'month') {
    const start = startOfMonth(base);
    const end = endOfMonth(base);
    const label = start.toLocaleDateString('es-CL', { month: 'short', year: 'numeric' });
    return { key: `${start.getFullYear()}-${String(start.getMonth() + 1).padStart(2, '0')}`, start, end, label };
  }
  return {
    key: base.toISOString().slice(0, 10),
    start: base,
    end: new Date(base.getFullYear(), base.getMonth(), base.getDate(), 23, 59, 59),
    label: formatShortDateLabel(base),
  };
}

function resolvePrimaMetaDate(task) {
  if (!task) return null;
  const issueDate =
    task.fechaEmisionPoliza ||
    task.case?.poliza?.issueDate ||
    task.case?.poliza?.fechaEmisionPoliza;
  if (issueDate) {
    const parsed =
      typeof issueDate === 'string' && issueDate.indexOf('T') === -1
        ? new Date(`${issueDate}T00:00:00`)
        : new Date(issueDate);
    const d = parsed;
    if (!Number.isNaN(d.getTime())) return d;
  }
  const fallback = task.updatedAt || task.createdAt || task.fechaCreacion;
  if (!fallback) return null;
  const d = new Date(fallback);
  return Number.isNaN(d.getTime()) ? null : d;
}

function groupPrimaMetaByPeriod(deals = [], fromDate, toDate, mode) {
  const buckets = new Map();
  const safeMode = mode || 'day';
  deals.forEach((task) => {
    const statusNorm = normalizeTaskStatus(task.status);
    if (!STATUS_WON.has(statusNorm)) return;
    const stamp = resolvePrimaMetaDate(task);
    if (!stamp || Number.isNaN(stamp.getTime())) return;
    if (stamp < fromDate || stamp > toDate) return;
    const { key, label, start, end } = getBucketIdentity(stamp, safeMode);
    if (!buckets.has(key)) {
      buckets.set(key, { key, label, start, end, prima: 0, commission: 0, count: 0 });
    }
    const bucket = buckets.get(key);
    bucket.prima += getPolicyPremiumValue(task);
    bucket.commission += getDashboardCommissionValue(task);
    bucket.count += 1;
  });
  return Array.from(buckets.values()).sort((a, b) => a.start - b.start);
}

function limitBucketsByGranularity(buckets, mode) {
  if (!Array.isArray(buckets) || !buckets.length) return buckets;
  let maxCount;
  if (mode === 'day') {
    maxCount = 13;
  } else if (mode === 'week') {
    maxCount = 7;
  } else if (mode === 'month') {
    maxCount = 12;
  } else {
    return buckets;
  }
  if (buckets.length > maxCount) {
    return buckets.slice(buckets.length - maxCount);
  }
  return buckets;
}

function getBucketRange(bucket) {
  if (!bucket) return null;
  const from = new Date(bucket.start.getFullYear(), bucket.start.getMonth(), bucket.start.getDate());
  const to = new Date(bucket.end.getFullYear(), bucket.end.getMonth(), bucket.end.getDate(), 23, 59, 59);
  return { from, to };
}

function applyDashboardRangeFromBucket(bucket) {
  const range = getBucketRange(bucket);
  if (!range) return;
  dashboardFilters.range = 'custom';
  dashboardFilters.customFrom = formatDateInputValue(range.from);
  dashboardFilters.customTo = formatDateInputValue(range.to);
  const rangeSelect = document.getElementById('dashFilterRange');
  const customGroup = document.getElementById('dashCustomRangeGroup');
  const customFrom = document.getElementById('dashCustomFrom');
  const customTo = document.getElementById('dashCustomTo');
  if (rangeSelect) rangeSelect.value = 'custom';
  if (customGroup) customGroup.classList.add('visible');
  if (customFrom) customFrom.value = dashboardFilters.customFrom;
  if (customTo) customTo.value = dashboardFilters.customTo;
  refreshDashboard();
}

function renderLineChart(deals, goalContext) {
  const ctx = document.getElementById('chartPrimaMeta');
  if (!ctx) return;
  const existingTooltip = document.getElementById('chart-prima-tooltip');
  if (existingTooltip) existingTooltip.classList.add('hidden');
  if (!hasChartLibrary()) {
    showChartFallback(ctx, 'Chart.js no disponible');
    return;
  }
  const { from, to } = getDashboardDateBounds();
  const effectiveMode = getEffectiveGranularityMode(from, to);
  updateKpiGranularityLabel(effectiveMode);
  updateGranularityButtonsActive(kpiGranularityMode);

  let buckets = groupPrimaMetaByPeriod(Array.isArray(deals) ? deals : [], from, to, effectiveMode);
  buckets = limitBucketsByGranularity(buckets, effectiveMode);
  if (dashboardCharts.meta) {
    dashboardCharts.meta.destroy();
    dashboardCharts.meta = null;
  }
  if (!buckets.length) {
    showChartFallback(ctx, 'No hay datos para el rango seleccionado.');
    return;
  }
  hideChartFallback(ctx);
  prepareChartCanvas(ctx);
  const labels = buckets.map((b) => b.label);
  const primaData = buckets.map((b) => b.prima);
  const commissionData = buckets.map((b) => b.commission);
  const metaPeriod = normalizeGoalPeriod(to);
  const goal = goalContext?.resolve(metaPeriod);
  const metaData = buckets.map((bucket) => {
    const meta =
      goalContext && obtenerMetaVigenteParaFecha(goalContext.tipo, goalContext.scopeId, bucket.start, goalContext.metas);
    const resolved = meta || goal;
    return resolved?.metaPrimaUF ?? resolved?.goalPremiumUF ?? 0;
  });
  if (dashboardCharts.meta) dashboardCharts.meta.destroy();
  const getOrCreateMetaTooltip = () => {
    let tooltipEl = document.getElementById('chart-prima-tooltip');
    if (!tooltipEl) {
      tooltipEl = document.createElement('div');
      tooltipEl.id = 'chart-prima-tooltip';
      tooltipEl.className = 'chart-tooltip hidden';
      document.body.appendChild(tooltipEl);
    }
    return tooltipEl;
  };
  const metaTooltip = (context) => {
    const tooltip = context.tooltip;
    const tooltipEl = getOrCreateMetaTooltip();
    if (!tooltip || tooltip.opacity === 0) {
      tooltipEl.classList.add('hidden');
      return;
    }
    tooltipEl.classList.remove('hidden');
    const title = (tooltip.title && tooltip.title[0]) || '';
    const rows = (tooltip.dataPoints || []).map((dp) => {
      const val = Number(dp.raw || 0);
      const formatted = Number.isFinite(val)
        ? `${val.toLocaleString('es-CL', { maximumFractionDigits: 1 })} UF`
        : '—';
      const color =
        dp.dataset?.borderColor ||
        dp.dataset?.backgroundColor ||
        'rgba(59,130,246,1)';
      return `<div class="chart-tooltip-row">
        <span class="chart-tooltip-dot" style="background:${color}"></span>
        <span class="chart-tooltip-label">${dp.dataset?.label || ''}</span>
        <strong>${formatted}</strong>
      </div>`;
    });
    tooltipEl.innerHTML = `
      <div class="chart-tooltip-title">${title}</div>
      ${rows.join('')}
    `;
    const { left, top } = context.chart.canvas.getBoundingClientRect();
    tooltipEl.style.left = `${left + window.scrollX + tooltip.caretX}px`;
    tooltipEl.style.top = `${top + window.scrollY + tooltip.caretY - 12}px`;
  };
  dashboardCharts.meta = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'Prima real',
          data: primaData,
          backgroundColor: 'rgba(59,130,246,.15)',
          borderColor: 'rgba(59,130,246,1)',
          fill: false,
          tension: 0.25,
          borderWidth: 2,
          pointRadius: 3,
          pointHitRadius: 10,
          pointHoverRadius: 6,
        },
        {
          label: 'Comisión real',
          data: commissionData,
          backgroundColor: 'rgba(16,185,129,.15)',
          borderColor: 'rgba(16,185,129,1)',
          fill: false,
          tension: 0.25,
          borderWidth: 2,
          pointRadius: 3,
          pointHitRadius: 10,
          pointHoverRadius: 6,
        },
        {
          label: 'Meta Prima',
          data: metaData,
          borderDash: [5, 5],
          borderColor: 'rgba(148,163,184,1)',
          backgroundColor: 'rgba(148,163,184,.35)',
          tension: 0,
          fill: false,
          pointRadius: 0,
        },
      ],
    },
    options: {
      responsive: false,
      animation: false,
      maintainAspectRatio: false,
      resizeDelay: 0,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        tooltip: { enabled: false, external: metaTooltip },
        legend: { display: true },
      },
      scales: {
        y: { beginAtZero: true },
      },
    },
  });
}

const DASHBOARD_DETAIL_COLUMNS = [
  { key: 'id', header: 'ID de negocio' },
  { key: 'client', header: 'Cliente' },
  { key: 'executive', header: 'Ejecutivo' },
  { key: 'company', header: 'Compañía' },
  { key: 'product', header: 'Ramo' },
  { key: 'prima', header: 'Prima (UF)' },
  { key: 'commission', header: 'Comisión (UF)' },
  { key: 'status', header: 'Estado' },
  { key: 'date', header: 'Fecha' },
  { key: 'issueDate', header: 'Fecha de emisión' },
];

function getDashboardDetailDataset(dealsData, opts = {}) {
  const { displayCompany = false } = opts;
  const fmtDate = (dateInput) => {
    const d = dateInput instanceof Date ? dateInput : new Date(dateInput);
    if (!d || Number.isNaN(d.getTime())) return '';
    return d.toLocaleDateString();
  };
  const formatNumber = (value) => {
    const num = Number(value);
    const safe = Number.isFinite(num) ? num : 0;
    return safe.toLocaleString('es-CL', { maximumFractionDigits: 1 });
  };
  const rows = (dealsData?.all || [])
    .map((task) => {
      const prima = getPolicyPremiumValue(task);
      const commission = getDashboardCommissionValue(task);
      const issueDate =
        task.fechaEmisionPoliza ||
        task.case?.poliza?.issueDate ||
        task.case?.poliza?.fechaEmisionPoliza ||
        '';
      return {
        id: task.title || task.id || '',
        client: task.client || task.case?.personal?.firstName || '',
        executive: task.assigneeName || task.assignee || task.asignadoAEmail || '',
        company: task.insurer || '',
        product: task.type || '',
        primaValue: prima,
        commissionValue: commission,
        status: normalizeTaskStatus(task.status),
        dateValue: task.updatedAt || task.createdAt || Date.now(),
        issueDateValue: issueDate,
      };
    })
    .sort((a, b) => Number(b.primaValue || 0) - Number(a.primaValue || 0));

  return rows.map((row) => ({
    id: row.id,
    client: row.client,
    executive: row.executive,
    company: displayCompany ? getInsurerDisplayName(row.company) : row.company,
    product: row.product,
    prima: formatNumber(row.primaValue),
    commission: formatNumber(row.commissionValue),
    status: row.status,
    date: fmtDate(row.dateValue),
    issueDate: row.issueDateValue ? fmtDate(row.issueDateValue) : '',
  }));
}

function renderDashboardTable(dealsData) {
  const tbody = document.getElementById('dashDealsTableBody');
  if (!tbody) return;
  const rows = getDashboardDetailDataset(dealsData, { displayCompany: true });
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="10" class="muted">No hay datos para el filtro seleccionado.</td></tr>';
    return;
  }
  const html = rows
    .map(
      (row) =>
        `<tr>${DASHBOARD_DETAIL_COLUMNS.map((col) => `<td>${row[col.key] ?? ''}</td>`).join('')}</tr>`
    )
    .join('');
  tbody.innerHTML = html;
}

function renderHistorialMetas(scope) {
  const tbody = document.getElementById('historialMetasBody');
  if (!tbody) return;
  const targetScope = scope || obtenerScopeMetaParaUsuarioActual();
  const metas = seedMetaHistoryFromLegacy(targetScope);
  const normalizedScope = normalizeMetaScopeId(targetScope?.scopeId);
  const fmtPeriod = (period) => {
    if (!period) return '';
    const norm = normalizeGoalPeriod(period);
    const [year, month] = norm.split('-');
    return `${month}/${year}`;
  };
  const fmtNumber = (value) => {
    const num = Number(value);
    if (!Number.isFinite(num)) return '—';
    return num.toLocaleString('es-CL', { maximumFractionDigits: 1 });
  };
  const fmtDateTime = (value) => {
    if (!value) return '-';
    const d = new Date(value);
    if (!d || Number.isNaN(d.getTime())) return '-';
    return d.toLocaleString('es-CL');
  };
  const escapeCell = (value) =>
    String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  const filtered = metas
    .filter(
      (meta) =>
        meta &&
        meta.tipo === targetScope.tipo &&
        normalizeMetaScopeId(meta.scopeId) === normalizedScope &&
        meta.periodoDesde
    )
    .sort((a, b) =>
      normalizeGoalPeriod(b.periodoDesde || '').localeCompare(normalizeGoalPeriod(a.periodoDesde || ''))
    );
  tbody.innerHTML =
    filtered
      .map((meta) => {
        const usuario = escapeCell(meta.creadoPorUsuarioId || meta.creadoPorEmail || '-');
        const comentario = meta.comentario ? escapeCell(meta.comentario) : '-';
        return `<tr>
          <td>${fmtPeriod(meta.periodoDesde)}</td>
          <td>${meta.periodoHasta ? fmtPeriod(meta.periodoHasta) : 'Vigente'}</td>
          <td>${fmtNumber(meta.metaPrimaUF)}</td>
          <td>${fmtNumber(meta.metaComisionUF)}</td>
          <td>${usuario}</td>
          <td>${fmtDateTime(meta.creadoEn)}</td>
          <td>${comentario}</td>
        </tr>`;
      })
      .join('') ||
    '<tr><td colspan="7" class="muted">Sin metas registradas.</td></tr>';
}

function refreshDashboard() {
  if (!dashboardViewOnly) return;
  const data = getFilteredDealsForDashboard();
  const goalContext = buildDashboardGoalContext();
  const kpis = computeDashboardKpis(data, goalContext);
  renderDashboardKpis(kpis);
  renderCompanyChart(data.won);
  renderFunnel(data);
  renderLineChart(data.won, goalContext);
  renderDashboardTable(data);
  renderHistorialMetas(goalContext);
}

function initGlobalGoalsCard() {
  const card = document.getElementById('dashboardGoalPanel');
  if (!card) return;
  const role = roleKey(state.user?.role || '');
  const allowed =
    role === ROLE_KEY_USUARIO_MAESTRO || role === ROLE_KEY_ADMINISTRADOR;
  card.classList.toggle('is-hidden', !allowed);
  if (!allowed) return;
  const rut = state.user?.brokerRut || state.user?.rutCorredora || '';
  const periodInput = document.getElementById('metaPeriodoDesde') || document.getElementById('globalGoalPeriod');
  const premiumInput = document.getElementById('globalGoalPremium');
  const commissionInput = document.getElementById('globalGoalCommission');
  const dealsInput = document.getElementById('globalGoalDeals');
  const commentInput = document.getElementById('metaComentario');
  const saveBtn = document.getElementById('btnSaveGlobalGoal');
  const statusLabel = document.getElementById('globalGoalStatus');
  const scope = obtenerScopeMetaParaUsuarioActual();
  if (periodInput && !periodInput.value) periodInput.value = getCurrentGoalPeriod();

  const fillForm = () => {
    const metas = seedMetaHistoryFromLegacy(scope);
    if (!periodInput) return;
    const period = normalizeGoalPeriod(periodInput.value || getCurrentGoalPeriod());
    periodInput.value = period;
    const meta = obtenerMetaVigenteParaFecha(scope.tipo, scope.scopeId, period, metas);
    if (premiumInput) premiumInput.value = meta?.metaPrimaUF ?? meta?.goalPremiumUF ?? '';
    if (commissionInput) commissionInput.value = meta?.metaComisionUF ?? meta?.goalCommissionUF ?? '';
    if (dealsInput) dealsInput.value = meta?.goalDeals ?? '';
    if (commentInput) commentInput.value = '';
    if (statusLabel) statusLabel.textContent = meta ? 'Meta vigente cargada' : '';
    renderHistorialMetas(scope);
  };

  periodInput?.addEventListener('change', fillForm);
  saveBtn?.addEventListener('click', () => {
    const period = normalizeGoalPeriod(periodInput?.value || getCurrentGoalPeriod());
    const premiumVal = premiumInput?.value ?? '';
    const commissionVal = commissionInput?.value ?? '';
    const premium =
      premiumVal === '' || premiumVal === null ? null : Number(premiumVal);
    const commission =
      commissionVal === '' || commissionVal === null ? null : Number(commissionVal);
    if ((premium !== null && premium < 0) || (commission !== null && commission < 0)) {
      alert('Las metas deben ser mayores o iguales a cero.');
      return;
    }
    if (scope.tipo === 'corredora' && !rut) {
      alert('No se detecta el RUT de la corredora actual.');
      return;
    }
    let metas = seedMetaHistoryFromLegacy(scope).slice();
    const normalizedScope = normalizeMetaScopeId(scope.scopeId);
    const metasScope = metas
      .map((meta, idx) => ({ meta, idx }))
      .filter(
        ({ meta }) =>
          meta &&
          meta.tipo === scope.tipo &&
          normalizeMetaScopeId(meta.scopeId) === normalizedScope
      )
      .sort((a, b) =>
        normalizeGoalPeriod(a.meta.periodoDesde || '').localeCompare(
          normalizeGoalPeriod(b.meta.periodoDesde || '')
        )
      );
    const vigente = metasScope
      .slice()
      .reverse()
      .find(({ meta }) => {
        const desde = normalizeGoalPeriod(meta.periodoDesde);
        const hasta = meta.periodoHasta ? normalizeGoalPeriod(meta.periodoHasta) : null;
        return desde <= period && (!hasta || hasta >= period);
      });
    if (vigente) {
      const prevPeriod = getPreviousGoalPeriod(period);
      const minDesde = normalizeGoalPeriod(vigente.meta.periodoDesde);
      if (!vigente.meta.periodoHasta || prevPeriod >= minDesde) {
        metas[vigente.idx] = { ...vigente.meta, periodoHasta: prevPeriod };
      }
    }
    const identity = getCurrentUserIdentity();
    const nuevaMeta = {
      idMeta: buildMetaId(),
      tipo: scope.tipo,
      scopeId: normalizedScope,
      periodoDesde: period,
      periodoHasta: null,
      metaPrimaUF: premium !== null && Number.isFinite(premium) ? premium : null,
      metaComisionUF:
        commission !== null && Number.isFinite(commission) ? commission : null,
      creadoPorUsuarioId:
        identity.email || identity.name || identity.id || identity.rut || 'usuario',
      creadoEn: new Date().toISOString(),
      comentario: commentInput?.value?.trim() ? commentInput.value.trim() : null,
    };
    metas.push(nuevaMeta);
    guardarHistorialMetas(metas);
    if (statusLabel) {
      statusLabel.textContent = 'Meta guardada';
      setTimeout(() => (statusLabel.textContent = ''), 2500);
    }
    closeMetaModal();
    renderHistorialMetas(scope);
    refreshDashboard();
  });

  fillForm();
}

// Fix: toggle Dashboard vs listado de negocios para evitar pantalla congelada
function applyDashboardView() {
  const dashEl = document.getElementById('dashboardSummary');
  const tableCard = document.getElementById('tableCard');
  const btn = document.getElementById('btnDashboardToggle');
  const btnWrap = document.getElementById('dashboardToggleWrap');
  const caseEditor = document.getElementById('caseEditor');
  const goalPanel = document.getElementById('dashboardGoalPanel');
  const wasEditing = document.body.classList.contains('editing');
  state.ui = state.ui || {};
  state.ui.isDashboardVisible = !!dashboardViewOnly;
  if (dashboardViewOnly) {
    document.body.classList.add('dashboard-open');
    if (dashEl) {
      dashEl.classList.remove('hidden');
      dashEl.style.display = '';
    }
    if (tableCard) {
      tableCard.classList.add('hidden');
      tableCard.style.display = 'none';
    }
    if (btn) {
      btn.classList.add('active');
      btn.textContent = 'Dashboard KPIs';
    }
    if (btnWrap) btnWrap.classList.add('hidden');
    if (wasEditing) {
      document.getElementById('btnEditorExit')?.click();
    }
    setFormVisibility(false);
  } else {
    document.body.classList.remove('dashboard-open');
    if (tableCard) {
      tableCard.classList.remove('hidden');
      tableCard.style.display = '';
    }
    if (dashEl) {
      dashEl.classList.add('hidden');
      dashEl.style.display = 'none';
    }
    if (btn) {
      btn.classList.remove('active');
      btn.textContent = 'Dashboard KPIs';
    }
    if (btnWrap) btnWrap.classList.remove('hidden');
  }
  if (!dashboardViewOnly) {
    // ensure editor stays hidden unless user opens it again
    if (!document.body.classList.contains('editing') && caseEditor) {
      caseEditor.classList.add('hidden');
    }
    if (goalPanel) goalPanel.classList.add('is-hidden');
  }
  syncDashboardGoalUi({ forceHidePanel: !dashboardViewOnly });
  updateFiltersCardVisibility();
}

function loadDashboardViewState(options = {}) {
  const { forceDefault = false } = options;
  if (forceDefault) {
    dashboardViewOnly = false;
  } else {
    try {
      const raw = localStorage.getItem(DASHBOARD_VIEW_KEY);
      dashboardViewOnly = raw === '1';
    } catch {
      dashboardViewOnly = false;
    }
  }
  applyDashboardView();
  if (dashboardViewOnly) {
    refreshDashboard();
  }
}

function saveDashboardViewState() {
  localStorage.setItem(DASHBOARD_VIEW_KEY, dashboardViewOnly ? '1' : '0');
}

/* ===========================
   CRUD de tareas (panel derecho)
=========================== */
function resetForm() {
  const $ = (s) => document.querySelector(s);
  $('#taskId').value = '';
  $('#taskForm').reset();
  $('#title').value = '';
  $('#priority').value = 'Media';
  $('#type').value = 'Vehículo';
  const channelField = document.getElementById('taskChannel');
  if (channelField) channelField.value = '';
  const insurerField = document.getElementById('insurer');
  if (insurerField) {
    insurerField.value = DEFAULT_INSURER;
    syncCompanySearchInput(insurerField);
  }
  const assignField = document.getElementById('assignee');
  if (assignField) assignField.value = state.user?.email || '';
  const customerTypeField = document.getElementById('customerType');
  if (customerTypeField) customerTypeField.value = 'natural';
  const firstField = document.getElementById('firstName');
  if (firstField) firstField.value = '';
  const lastField = document.getElementById('lastName');
  if (lastField) lastField.value = '';
  const emailField = document.getElementById('email');
  if (emailField) emailField.value = '';
  const phoneField = document.getElementById('phone');
  if (phoneField) phoneField.value = '';
  const rutField = document.getElementById('nuevoNegocioRut');
  if (rutField) rutField.classList.remove('input-error');
  const ft = document.getElementById('formTitle');
  if (ft) ft.textContent = 'Nuevo negocio';
}
function setFormVisibility(show) {
  const body = document.body;
  if (body) body.classList.toggle('show-form', !!show);
  const modal = document.getElementById('formModal');
  if (modal) {
    modal.classList.toggle('hidden', !show);
    modal.setAttribute('aria-hidden', show ? 'false' : 'true');
  }
  const formPanel = modal?.querySelector('.panel-right') || document.querySelector('.panel-right');
  if (formPanel) formPanel.setAttribute('aria-hidden', show ? 'false' : 'true');
}
function upsertTask(p, options = {}) {
  const { silent = false, deferSave = false } = options;
  const existing = p.id ? state.tasks.find((x) => x.id === p.id) : null;
  const nowIso = new Date().toISOString();
  if (Object.prototype.hasOwnProperty.call(p, 'status')) {
    p.status = normalizeTaskStatus(p.status);
  }
  applyAssigneeMetadata(p, p.assignee || existing?.assignee || '');
  if (!existing) {
    ensureCreatorMetadata(p, getCurrentUserIdentity());
  } else {
    // Mantener metadata de creador incluso si no viene en p.
    const creatorFallback =
      existing && (existing.creadoPorUserId || existing.creadoPorEmail || existing.creadoPorRutCorredora)
        ? {
            id: existing.creadoPorUserId,
            email: existing.creadoPorEmail,
            rut: existing.creadoPorRut,
            brokerRut: existing.creadoPorRutCorredora,
            role: existing.rolCreador,
          }
        : getCurrentUserIdentity();
    ensureCreatorMetadata(p, creatorFallback);
  }
  if (p.id) {
    const i = state.tasks.findIndex((x) => x.id === p.id);
    if (i > -1) {
      const prev = state.tasks[i];
      const oldStatus = prev.status || 'Sin estado';
      const newStatus = p.status || oldStatus;
      const oldSubStatus = prev.subStatus || '';
      const newSubStatus =
        Object.prototype.hasOwnProperty.call(p, 'subStatus') && p.subStatus !== undefined
          ? p.subStatus
          : oldSubStatus;
      if (newStatus !== oldStatus) {
        pushBitacoraEntry(prev, `Estado cambiado de "${oldStatus}" a "${newStatus}"`);
      }
      if (newSubStatus !== oldSubStatus) {
        const fromLabel = oldSubStatus || 'Sin subestado';
        const toLabel = newSubStatus || 'Sin subestado';
        pushBitacoraEntry(prev, `Subestado cambiado de "${fromLabel}" a "${toLabel}"`);
      }
      const mergedCase = {
        ...prev.case,
        ...p.case,
      };
      mergedCase.comments = ensureCaseComments(prev).slice(0, 500);
      state.tasks[i] = normalizeTask({
        ...prev,
        ...p,
        case: mergedCase,
        status: newStatus,
        subStatus: newSubStatus,
        title: prev.title,
        type: prev.type,
        updatedAt: nowIso,
        fechaActualizacion: nowIso,
      });
    }
  } else {
    const newTask = normalizeTask({
      ...p,
      id: uuid(),
      title: p.title || generateBusinessId(p.type),
      createdAt: nowIso,
      updatedAt: nowIso,
      fechaActualizacion: nowIso,
    });
    state.tasks.push(newTask);
    if (DEBUG_RENOVACION_DATOS)
      console.log('Nueva tarea creada con renovacion/pagos', newTask);
  }
  if (!deferSave) {
    saveTasks();
    applyFilters();
  }
  if (!silent) {
    resetForm();
    setFormVisibility(false);
  }
}

/* ===========================
   Login / Eventos globales
=========================== */
(() => {
  const $ = (s) => document.querySelector(s);
  const normalizeEmail = (v = '') => v.trim().toLowerCase();

  const ensureUserInTenant = (tenantId, attrs = {}) => {
    const key = kUsers(tenantId);
    let list = [];
    try {
      const raw = localStorage.getItem(key);
      list = raw ? JSON.parse(raw) : [];
    } catch {
      list = [];
    }
    if (!Array.isArray(list)) list = [];
    const emailLower = normalizeEmail(attrs.email || '');
    const idx = list.findIndex((u) => normalizeEmail(u.email || '') === emailLower);
    const now = new Date().toISOString();
    if (idx === -1) {
      const newUser = {
        id: attrs.id || uuid(),
        name: attrs.name || attrs.email || 'Usuario',
        email: attrs.email,
        rut: attrs.rut || '',
        brokerRut: attrs.brokerRut || '',
        role: attrs.role || ROLE_ADMINISTRADOR,
        commission: attrs.commission ?? 0,
        status: 'activo',
        permissions: attrs.permissions || [],
        password: attrs.password || '',
        limiteUsuarios: attrs.limiteUsuarios ?? null,
        creadoPorId: attrs.creadoPorId ?? null,
        createdAt: now,
        updatedAt: now,
      };
      list.push(newUser);
    } else {
      list[idx] = { ...list[idx], ...attrs, updatedAt: now };
    }
    localStorage.setItem(key, JSON.stringify(list));
    return idx === -1 ? list[list.length - 1] : list[idx];
  };
  function findUserMatchesByEmail(email) {
    const matches = [];
    const lower = normalizeEmail(email);
    if (!lower) return matches;
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      const kMatch = key && key.match(/^corredores_v1:([^:]+):users$/);
      if (!kMatch) continue;
      const tenantId = kMatch[1];
      try {
        const raw = localStorage.getItem(key);
        const list = raw ? JSON.parse(raw) : [];
        if (!Array.isArray(list)) continue;
        const found = list.find(
          (u) => normalizeEmail(u.email || '') === lower
        );
        if (found) matches.push({ tenantId, user: found });
      } catch (e) {
        // ignorar entradas corruptas
      }
    }
    return matches;
  }

  function tryAutoLoginFromRemember() {
    let data = null;
    try {
      const raw = localStorage.getItem(REMEMBER_KEY);
      data = raw ? JSON.parse(raw) : null;
    } catch {
      data = null;
    }
    if (!data) return false;
    const now = Date.now();
    if (!data.expiresAt || now > data.expiresAt) {
      localStorage.removeItem(REMEMBER_KEY);
      return false;
    }
    const rememberedEmail = normalizeEmail(data.email || '');
    if (!rememberedEmail || !data.password) {
      localStorage.removeItem(REMEMBER_KEY);
      return false;
    }
    const matches = findUserMatchesByEmail(rememberedEmail);
    if (!matches || matches.length !== 1) {
      localStorage.removeItem(REMEMBER_KEY);
      return false;
    }
    const { tenantId, user } = matches[0];
    const storedPwd = user?.password;
    if (storedPwd === undefined || storedPwd === null || String(storedPwd) !== data.password) {
      localStorage.removeItem(REMEMBER_KEY);
      return false;
    }
    enterApp(tenantId, {
      id: user.id,
      name: user.name || user.email,
      email: user.email,
      role: user.role,
      rut: user.rut,
      brokerRut: user.brokerRut || user.rutCorredora,
      rutCorredora: user.rutCorredora,
    });
    return true;
  }

  function updateUserPasswordInTenant(tenantId, emailLower, newPassword) {
    let list = [];
    try {
      const raw = localStorage.getItem(kUsers(tenantId));
      list = raw ? JSON.parse(raw) : [];
    } catch {
      list = [];
    }
    if (!Array.isArray(list)) list = [];
    const idx = list.findIndex((u) => normalizeEmail(u.email || '') === emailLower);
    if (idx === -1) return false;
    list[idx].password = newPassword;
    list[idx].updatedAt = new Date().toISOString();
    localStorage.setItem(kUsers(tenantId), JSON.stringify(list));
    if (state.tenant === tenantId) {
      loadUsers();
    }
    return true;
  }
function enterApp(tenantId, user){
  sessionStorage.removeItem('kensaLogout');
  setSession(tenantId, user);
  applyLoginBody(false);
  applyShortcutsHintVisibility(true);
  const rutCorredora = state.user?.brokerRut || state.user?.rutCorredora || '';
  if (rutCorredora) {
    const goalsByTenant = loadAllGoalsFromStorage();
    const tenantGoals = ensureTenantGoals(goalsByTenant, rutCorredora);
    state.goals = cloneTenantGoals(tenantGoals);
    console.log('[Metas] Metas cargadas para', rutCorredora, state.goals);
  } else {
    state.goals = cloneTenantGoals(GOAL_DEFAULT_LIST);
    console.log('[Metas] usuario sin RUT de corredora, metas vacías');
  }
  loadClients();
  loadUsers();
  loadTasks();
  runAutoRenewalOncePerDay();
  loadNegocioAuditLog();
  state.ui = state.ui || {};
  state.ui.isCaseDetailOpen = false;
  state.ui.isDashboardVisible = false;
  state.ui.filtersMode = 'groups';
  state.filters = state.filters || { group: 'nuevos' };
  state.filters.group = 'nuevos';
  updateFiltersCardVisibility();
    populateAssignees();
    applyRoleUiVisibility();
    applyCorredoraLogo(state.user);
    initGlobalGoalsCard();
    initDashboardFiltersUI();
    initFiltersModeUI();
    initFilterGroupsBehavior();
    const login = document.getElementById('login');
    const app = document.getElementById('app');
    if (login) login.classList.add('hidden');
    if (app) app.classList.remove('hidden');
    syncAppHeaderOffset();
  loadDashboardViewState({ forceDefault: true });
  applyFilters();
}

  function bootstrapFromStoredSession(){
    if (state.user) return;
    let stored = null;
    try{
      const raw = localStorage.getItem('kensaCurrentTenant');
      stored = raw ? JSON.parse(raw) : null;
    }catch{ stored = null; }
    if(!stored || !stored.tenantId || !stored.user) return;
    enterApp(stored.tenantId, stored.user);
  }

  const ensureUserPasswordInTenant = (tenantId, emailLower, password) => {
    const updated = ensureUserInTenant(tenantId, { email: emailLower, password });
    if (state.tenant === tenantId) loadUsers();
    return !!updated;
  };

  const getPreferredTenant = () => {
    const params = new URLSearchParams(location.search);
    const queryTenant = params.get('tenant');
    if (queryTenant) return queryTenant;
    try {
      const raw = localStorage.getItem('kensaCurrentTenant');
      const parsed = raw ? JSON.parse(raw) : null;
      if (parsed?.tenantId) return parsed.tenantId;
    } catch {
      // ignore
    }
    return DEFAULT_TENANT;
  };

  const btnLogin = $('#btnLogin');
  if (btnLogin) {
    btnLogin.addEventListener('click', () => {
      const email = normalizeEmail($('#loginEmail')?.value || '');
      const pwd = $('#loginPassword')?.value || '';
      if (!email) {
        alert('Debe ingresar un correo');
        return;
      }
      if (!pwd) {
        alert('Debe ingresar una clave');
        return;
      }

      const preferTenantFromSession = () => {
        try {
          const raw = localStorage.getItem('kensaCurrentTenant');
          const parsed = raw ? JSON.parse(raw) : null;
          return parsed?.tenantId || null;
        } catch {
          return null;
        }
      };

      let matches = findUserMatchesByEmail(email);
      // Si no hay usuarios y se intenta con el master, resembrar y reintentar.
      if (matches.length === 0 && email === normalizeEmail(MASTER_SEED_EMAIL)) {
        const tenantId = getPreferredTenant();
        const seeded = ensureUserInTenant(tenantId, {
          email: MASTER_SEED_EMAIL,
          name: 'Usuario Maestro',
          role: ROLE_USUARIO_MAESTRO,
          password: MASTER_SEED_PASSWORD,
        });
        matches = seeded ? [{ tenantId, user: seeded }] : findUserMatchesByEmail(email);
      }
      if (matches.length === 0) {
        alert(
          'No existe un usuario registrado con ese correo. Solicite su creación al administrador.'
        );
        return;
      }
      if (matches.length > 1) {
        const preferred = preferTenantFromSession();
        const picked =
          (preferred && matches.find((m) => m.tenantId === preferred)) ||
          matches[0];
        matches = [picked];
        console.warn('Login: correo encontrado en múltiples corredoras, usando', picked.tenantId);
      }

      const { tenantId, user } = matches[0];
      const masterPwd = MASTER_SEED_PASSWORD;
      const isMasterUser = email === normalizeEmail(MASTER_SEED_EMAIL);

      // Si es master y la clave coincide, normalizamos y entramos directo
      if (isMasterUser && pwd === masterPwd) {
        ensureUserInTenant(tenantId, {
          email: MASTER_SEED_EMAIL,
          name: user.name || 'Usuario Maestro',
          role: user.role || ROLE_USUARIO_MAESTRO,
          password: masterPwd,
        });
        user.password = masterPwd;
      } else {
        const storedPwd = user?.password;
        if (storedPwd === undefined || storedPwd === null || String(storedPwd).trim() === '') {
          if (pwd !== masterPwd) {
            alert('Este usuario no tiene una clave definida. Usa la clave maestra o pide al admin que asigne una.');
            return;
          }
          updateUserPasswordInTenant(tenantId, email, masterPwd);
          user.password = masterPwd;
        } else if (String(storedPwd) !== pwd) {
          alert('Clave incorrecta');
          return;
        }
      }

      enterApp(tenantId, {
        id: user.id,
        name: user.name || user.email,
        email: user.email,
        role: user.role,
        rut: user.rut,
        brokerRut: user.brokerRut || user.rutCorredora,
        rutCorredora: user.rutCorredora,
      });
      const rememberToggle = document.getElementById('rememberMeToggle');
      const rememberOn = rememberToggle?.checked || false;
      if (rememberOn) {
        const now = Date.now();
        const record = {
          email,
          password: pwd,
          createdAt: now,
          expiresAt: now + THREE_MONTHS_MS,
        };
        localStorage.setItem(REMEMBER_KEY, JSON.stringify(record));
      } else {
        localStorage.removeItem(REMEMBER_KEY);
      }
    });
  }

  const btnSeed = $('#btnSeed');
  if (btnSeed) {
    btnSeed.addEventListener('click', () => {
      const email = $('#loginEmail');
      const pwd = $('#loginPassword');
      if (email) email.value = MASTER_SEED_EMAIL;
      if (pwd) pwd.value = MASTER_SEED_PASSWORD;
    });
  }

  let headerOffsetObserver;
  const syncAppHeaderOffset = () => {
    const appRoot = document.getElementById('app');
    const header = document.querySelector('#app > header.app');
    if (!appRoot || !header) return;
    const height = header.getBoundingClientRect().height;
    appRoot.style.setProperty('--app-header-offset', `${height}px`);
  };
  const setupAppHeaderOffsetSync = () => {
    const header = document.querySelector('#app > header.app');
    if (!header || headerOffsetObserver) return;
    syncAppHeaderOffset();
    if (window.ResizeObserver) {
      headerOffsetObserver = new ResizeObserver(syncAppHeaderOffset);
      headerOffsetObserver.observe(header);
    } else {
      window.addEventListener('resize', syncAppHeaderOffset);
    }
  };

  function initApp() {
    const finishLoading = () => {
      if (document?.body) document.body.classList.remove('app-loading');
    };
    try {
      setupAppHeaderOffsetSync();
      if (!state.user) bootstrapFromStoredSession();
      if (!state.user && !sessionStorage.getItem('kensaLogout')) {
        tryAutoLoginFromRemember();
      }
      if (!state.user) {
        const login = document.getElementById('login');
        const app = document.getElementById('app');
        if (login) login.classList.remove('hidden');
        if (app) app.classList.add('hidden');
        applyLoginBody(true);
        applyShortcutsHintVisibility(false);
      }
      [
        document.getElementById('insurer'),
        document.getElementById('ce_taskInsurer'),
        document.getElementById('fCompania'),
        document.getElementById('dashFilterCompany'),
      ].forEach((select) => setupCompanySearch(select));
    } finally {
      finishLoading();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }


  const btnMobileNav = document.getElementById('btnMobileNav');
  if (btnMobileNav) {
    btnMobileNav.addEventListener('click', () => {
      document.body.classList.toggle('is-mobile-nav-open');
    });
  }
  document.addEventListener('click', (event) => {
    if (!document.body.classList.contains('is-mobile-nav-open')) return;
    const filtersCard = document.getElementById('filtersCard');
    const sidebarActionsEl = document.getElementById('sidebarActions');
    if (filtersCard && filtersCard.contains(event.target)) return;
    if (sidebarActionsEl && sidebarActionsEl.contains(event.target)) return;
    if (btnMobileNav && btnMobileNav.contains(event.target)) return;
    document.body.classList.remove('is-mobile-nav-open');
    event.preventDefault();
    event.stopPropagation();
  });

  const sidebarActions = document.getElementById('sidebarActions');
  const headerActions = document.querySelector('.app-header-actions');
  const sidebarBrand = document.getElementById('sidebarBrand');
  const headerBrand = document.querySelector('.brand.header-left');
  const kensaLogo = document.getElementById('logoKonectaCuadrado');
  const headerUserInfo = document.querySelector('.header-user-info');
  const headerDivider = document.querySelector('.header-divider');
  const logoCorredora = document.getElementById('logoCorredora');
  const sidebarButtons = [
    document.getElementById('btnUsers'),
    document.getElementById('btnClientes'),
    document.getElementById('btnTheme'),
    document.getElementById('btnLogout'),
  ];
  const placeSidebarButtons = () => {
    if (!sidebarActions || !headerActions) return;
    const useSidebar =
      window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    const target = useSidebar ? sidebarActions : headerActions;
    sidebarButtons.forEach((btn) => {
      if (!btn || btn.parentElement === target) return;
      target.appendChild(btn);
    });
  };
  const placeSidebarBrand = () => {
    if (!sidebarBrand || !headerBrand || !kensaLogo || !headerUserInfo) return;
    const useSidebar =
      window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    if (useSidebar) {
      if (kensaLogo.parentElement !== sidebarBrand) sidebarBrand.appendChild(kensaLogo);
      if (headerUserInfo.parentElement !== sidebarBrand) sidebarBrand.appendChild(headerUserInfo);
      return;
    }
    if (kensaLogo.parentElement !== headerBrand) {
      if (headerDivider && headerDivider.parentElement === headerBrand) {
        headerDivider.insertAdjacentElement('afterend', kensaLogo);
      } else if (logoCorredora && logoCorredora.parentElement === headerBrand) {
        logoCorredora.insertAdjacentElement('afterend', kensaLogo);
      } else {
        headerBrand.appendChild(kensaLogo);
      }
    }
    if (headerUserInfo.parentElement !== headerBrand) {
      headerBrand.appendChild(headerUserInfo);
    }
  };
  placeSidebarButtons();
  placeSidebarBrand();
  if (window.matchMedia) {
    const sidebarMedia = window.matchMedia('(max-width: 768px)');
    if (sidebarMedia.addEventListener) {
      sidebarMedia.addEventListener('change', () => {
        placeSidebarButtons();
        placeSidebarBrand();
      });
    } else {
      window.addEventListener('resize', () => {
        placeSidebarButtons();
        placeSidebarBrand();
      });
    }
  } else {
    window.addEventListener('resize', () => {
      placeSidebarButtons();
      placeSidebarBrand();
    });
  }

  const bulkActionsModal = document.getElementById('bulkActionsModal');
  const bulkActionsOpen = document.getElementById('btnBulkActions');
  const bulkActionsClose = document.getElementById('bulkActionsClose');
  const bulkActionsBody = document.getElementById('bulkActionsBody');
  const bulkActionsInline = document.getElementById('bulkActionsInline');
  const dashboardFiltersToggle = document.getElementById('btnDashboardFiltersToggle');
  const bulkActionButtons = [
    document.getElementById('btnDeleteSelected'),
    document.getElementById('btnBulkStatus'),
    document.getElementById('btnBulkAssign'),
  ];
  const openBulkActionsModal = () => {
    if (!bulkActionsModal) return;
    bulkActionsModal.classList.remove('hidden');
    bulkActionsModal.setAttribute('aria-hidden', 'false');
  };
  const closeBulkActionsModal = () => {
    if (!bulkActionsModal) return;
    bulkActionsModal.classList.add('hidden');
    bulkActionsModal.setAttribute('aria-hidden', 'true');
  };
  if (bulkActionsOpen) {
    bulkActionsOpen.addEventListener('click', () => {
      if (!bulkActionsModal) return;
      openBulkActionsModal();
    });
  }
  if (bulkActionsClose) {
    bulkActionsClose.addEventListener('click', () => closeBulkActionsModal());
  }
  if (bulkActionsModal) {
    bulkActionsModal.addEventListener('click', (event) => {
      if (event.target === bulkActionsModal) closeBulkActionsModal();
    });
  }
  if (dashboardFiltersToggle) {
    dashboardFiltersToggle.addEventListener('click', () => {
      document.body.classList.toggle('is-dashboard-filters-open');
    });
  }
  if (bulkActionsBody) {
    bulkActionsBody.addEventListener('click', (event) => {
      const targetBtn = event.target.closest('button');
      if (!targetBtn) return;
      if (bulkActionsModal && !bulkActionsModal.classList.contains('hidden')) {
        closeBulkActionsModal();
      }
    });
  }
  const placeBulkActionButtons = () => {
    if (!bulkActionsBody || !bulkActionsInline) return;
    const useModal =
      window.matchMedia && window.matchMedia('(max-width: 1024px)').matches;
    const target = useModal ? bulkActionsBody : bulkActionsInline;
    const beforeNode = !useModal && bulkActionsOpen ? bulkActionsOpen : null;
    bulkActionButtons.forEach((btn) => {
      if (!btn || btn.parentElement === target) return;
      target.insertBefore(btn, beforeNode);
    });
    if (!useModal) closeBulkActionsModal();
  };
  placeBulkActionButtons();
  if (window.matchMedia) {
    const bulkMedia = window.matchMedia('(max-width: 1024px)');
    if (bulkMedia.addEventListener) {
      bulkMedia.addEventListener('change', placeBulkActionButtons);
    } else {
      window.addEventListener('resize', placeBulkActionButtons);
    }
  } else {
    window.addEventListener('resize', placeBulkActionButtons);
  }

  const btnLogout = $('#btnLogout');
  const logoutModal = document.getElementById('logoutModal');
  const logoutConfirm = document.getElementById('logoutConfirm');
  const logoutCancel = document.getElementById('logoutCancel');
  const showLogoutModal = () => {
    if (!logoutModal) return;
    logoutModal.classList.remove('hidden');
    logoutModal.setAttribute('aria-hidden', 'false');
  };
  const hideLogoutModal = () => {
    if (!logoutModal) return;
    logoutModal.classList.add('hidden');
    logoutModal.setAttribute('aria-hidden', 'true');
  };
  if (logoutModal) {
    logoutModal.addEventListener('click', (e) => {
      if (e.target === logoutModal) hideLogoutModal();
    });
  }
  if (btnLogout) {
    btnLogout.addEventListener('click', () => {
      showLogoutModal();
    });
  }
  const loginRootEl = document.getElementById('login');
  if (loginRootEl) {
    loginRootEl.addEventListener('animationend', () => {
      loginRootEl.classList.remove('login-animate');
    });
  }
  if (logoutConfirm) {
    logoutConfirm.addEventListener('click', () => {
      hideLogoutModal();
      localStorage.removeItem('kensaCurrentTenant');
      sessionStorage.setItem('kensaLogout', '1');
      location.reload();
    });
  }
  if (logoutCancel) {
    logoutCancel.addEventListener('click', () => hideLogoutModal());
  }

  const linkForgot = document.getElementById('linkForgotPassword');
  const forgotDialog = document.getElementById('forgotDialog');
  const btnForgotCancel = document.getElementById('btnForgotCancel');
  const btnForgotSubmit = document.getElementById('btnForgotSubmit');
  const passwordRegex = /^(?=.*[A-Z])(?=.*[!"#$)=?]).{8,}$/;
  const forgotEmailInput = document.getElementById('forgotEmail');
  const forgotPassInput = document.getElementById('forgotPassword');
  const forgotPassConfirmInput = document.getElementById('forgotPasswordConfirm');

  function closeForgotDialog() {
    if (forgotDialog) forgotDialog.classList.add('hidden');
    if (forgotEmailInput) forgotEmailInput.value = '';
    if (forgotPassInput) forgotPassInput.value = '';
    if (forgotPassConfirmInput) forgotPassConfirmInput.value = '';
  }

  if (linkForgot && forgotDialog) {
    linkForgot.addEventListener('click', (e) => {
      e.preventDefault();
      forgotDialog.classList.remove('hidden');
    });
  }
  if (btnForgotCancel && forgotDialog) {
    btnForgotCancel.addEventListener('click', () => closeForgotDialog());
  }
  if (forgotDialog) {
    forgotDialog.addEventListener('click', (e) => {
      if (e.target === forgotDialog) closeForgotDialog();
    });
  }
  if (btnForgotSubmit && forgotDialog) {
    btnForgotSubmit.addEventListener('click', () => {
      const email = normalizeEmail(forgotEmailInput?.value || '');
      const pass = forgotPassInput?.value || '';
      const pass2 = forgotPassConfirmInput?.value || '';
      if (!email) {
        alert('Debes ingresar tu correo.');
        return;
      }
      if (!pass || !pass2) {
        alert('Debes ingresar y confirmar la nueva contraseña.');
        return;
      }
      if (pass !== pass2) {
        alert('La contraseña y la confirmación no coinciden.');
        return;
      }
      if (!passwordRegex.test(pass)) {
        alert('La contraseña debe tener al menos 8 caracteres, 1 mayúscula y 1 carácter especial de: !"#$)=?');
        return;
      }
      const matches = findUserMatchesByEmail(email);
      if (!matches || matches.length === 0) {
        alert('No existe un usuario con ese correo. Contacta a tu administrador.');
        return;
      }
      if (matches.length > 1) {
        alert('Este correo existe en más de una corredora. Contacta a soporte para resolverlo.');
        return;
      }
      const { tenantId } = matches[0];
      const updated = updateUserPasswordInTenant(tenantId, email, pass);
      if (!updated) {
        alert('No fue posible actualizar la contraseña. Intenta nuevamente.');
        return;
      }
      localStorage.removeItem(REMEMBER_KEY);
      alert('Tu contraseña ha sido actualizada. Ahora puedes iniciar sesión con la nueva clave.');
      closeForgotDialog();
    });
  }

  const btnNew = $('#btnNew');
  if (btnNew) {
    btnNew.addEventListener('click', () => {
      setFormVisibility(true);
      resetForm();
      $('#type')?.focus();
    });
  }

  const formModal = document.getElementById('formModal');
  if (formModal) {
    formModal.addEventListener('click', (e) => {
      if (e.target === formModal) {
        resetForm();
        setFormVisibility(false);
      }
    });
  }

  const taskForm = document.getElementById('taskForm');
  const newBusinessRutInput = document.getElementById('nuevoNegocioRut');
  attachRutSanitizer(newBusinessRutInput);
  attachPhoneSanitizer(document.getElementById('phone'));
  if (taskForm) {
    taskForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const titleInput = document.getElementById('title');
      const typeInput = document.getElementById('type');
      const assigneeInput = document.getElementById('assignee');
      const customerTypeInput = document.getElementById('customerType');
      const firstNameInput = document.getElementById('firstName');
      const lastNameInput = document.getElementById('lastName');
      const emailInput = document.getElementById('email');
      const phoneInput = document.getElementById('phone');
      const rutInput = newBusinessRutInput || document.getElementById('nuevoNegocioRut');
      const channelInput = document.getElementById('taskChannel');
      const firstName = firstNameInput?.value.trim() || '';
      const lastName = lastNameInput?.value.trim() || '';
      const clientName = [firstName, lastName].filter(Boolean).join(' ').trim();
      const canalIngreso = normalizeChannel(channelInput?.value || '');
      const tagValue = (document.getElementById('tags').value || '').trim();
      const policyInput = document.getElementById('policy');
      const policyValue =
        policyInput && typeof policyInput.value === 'string' ? policyInput.value.trim() : '';
      const insurerInput = document.getElementById('insurer');
      const insurerValue = (insurerInput?.value || DEFAULT_INSURER).trim() || DEFAULT_INSURER;
      const dueInput = document.getElementById('due');
      const dueValue = dueInput?.value || null;
      const rutValue = rutInput?.value.trim() || '';
      const normalizedRutValue = rutValue.trim().toUpperCase();
      const p = {
        id: document.getElementById('taskId').value || undefined,
        title: titleInput ? titleInput.value.trim() : '',
        type: typeInput?.value || '',
        priority: document.getElementById('priority').value,
        due: dueValue,
        policyEndDate: dueValue,
        assignee: (assigneeInput?.value?.trim() || state.user?.email || '') ?? '',
        client: clientName,
        email: emailInput?.value.trim() || '',
        phone: phoneInput?.value.trim() || '',
        customerType: customerTypeInput?.value || 'natural',
        rut: normalizedRutValue,
        policy: policyValue,
        insurer: insurerValue,
        tags: tagValue ? [tagValue] : [],
        canalIngreso,
        fueEditado: false,
      };
      const currentIdentity = getCurrentUserIdentity();
      const currentRole = roleKey(currentIdentity.role);
      if (currentRole === ROLE_KEY_AGENTE) {
        p.assignee = currentIdentity.email || '';
      } else if (currentRole === ROLE_KEY_SUPERVISOR) {
        p.assignee = p.assignee || currentIdentity.email || '';
      }
      const initialStatus =
        canalIngreso === CHANNEL_RENOVACION ? 'En renovación' : 'Abierta';
      p.status = normalizeTaskStatus(initialStatus);
      p.estadoInicial = p.status;
      if (!p.type) {
        alert('Tipo requerido');
        return;
      }
      if (!rutValue || !esRutValido(rutValue)) {
        if (rutInput) {
          rutInput.classList.add('input-error');
          rutInput.focus();
        }
        alert('RUT inválido. Formato esperado: 12345678-9');
        return;
      }
      if (rutInput) {
        rutInput.classList.remove('input-error');
        rutInput.value = normalizedRutValue;
      }
      p.case = {
        personal: {
          firstName,
          lastName,
          rut: p.rut,
          email: p.email,
          phone: p.phone,
        },
        pago: {
          firstName: '',
          lastName: '',
          tipo: 'natural',
          usarPersonales: false,
          rut: '',
          email: '',
          phone: '',
        },
        poliza: {},
        vehiculo: {},
        riesgo: {},
        docs: [],
        comments: [],
        renovacion: crearRenovacion({
          estadoRenovacion: ESTADOS_RENOVACION[0] || '',
        }),
        pagos: [],
      };
      if (!Array.isArray(p.case.comments)) p.case.comments = [];
      if (!p.id) {
        const actorName = getActorName();
        const initialStatus = p.status || TASK_STATUS_OPTIONS[0];
        p.case.comments.unshift({
          by: actorName,
          ts: Date.now(),
          text: `Negocio creado en estado "${initialStatus}" por ${actorName}`,
        });
      }
      p.insurer = p.insurer || DEFAULT_INSURER;
      if (!p.id) {
        p.title = generateBusinessId(p.type);
        if (titleInput) titleInput.value = p.title;
      } else if (!p.title) {
        p.title = generateBusinessId(p.type);
      }
      upsertTask(p);
    });
  }

  const btnReset = document.getElementById('btnReset');
  if (btnReset)
    btnReset.addEventListener('click', () => {
      resetForm();
      setFormVisibility(false);
    });
  setFormVisibility(false);

  // Tabla: editar
  const tbody = document.getElementById('tbody');
  if (tbody) {
    tbody.addEventListener('click', (e) => {
      const btn = e.target.closest('.btnEdit');
      if (!btn) return;
      const tr = btn.closest('tr');
      if (!tr) return;
      const id = tr.getAttribute('data-id');
      if (!id) return;
      const t = state.tasks.find((x) => String(x.id) === String(id));
      if (!t) return;
      openEditorFor(t);
    });
  }

  // Atajos
  document.addEventListener('keydown', (e) => {
    if (!isUserLoggedIn()) return;
    if (e.key === 'Escape' && document.body.classList.contains('show-form')) {
      e.preventDefault();
      resetForm();
      setFormVisibility(false);
      return;
    }
    if (e.key === '/' && document.activeElement !== document.getElementById('q')) {
      e.preventDefault();
      document.getElementById('q')?.focus();
    }
    if (
      e.key.toLowerCase() === 'n' &&
      !['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName)
    ) {
      e.preventDefault();
      const newBtn = document.getElementById('btnNew');
      if (!newBtn || newBtn.disabled) return;
      newBtn.click();
    }
  });

  // Filtros (input en vivo)
  [
    '#q',
    '#fTipo',
    '#fEstado',
    '#fPrioridad',
    '#fVenc',
    '#fCompania',
    '#fTags',
    '#fSoloMias',
  ].forEach((sel) => {
    const el = document.querySelector(sel);
    if (el) el.addEventListener('input', applyFilters);
  });

  // Export/Import
function download(name, content, type = 'application/json') {
  const blob =
    content instanceof Blob ? content : new Blob([content], { type });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1500);
}
function getScopedTasksForExport() {
  const role = roleKey(state.user?.role || '');
  const tasks = Array.isArray(state.tasks) ? state.tasks : [];
  if (role === ROLE_KEY_USUARIO_MAESTRO) return tasks.slice();
  return getTasksVisibleForCurrentUser(tasks);
}
const btnExport = document.getElementById('btnExport');
const exportRangeModal = document.getElementById('exportRangeModal');
const exportRangeWarning = document.getElementById('exportRangeWarning');
const exportDesde = document.getElementById('exportDesde');
const exportHasta = document.getElementById('exportHasta');
const btnExportCerrar = document.getElementById('btnExportCerrar');
const btnExportDescargar = document.getElementById('btnExportDescargar');
let currentExportType = 'excel';

function isDashboardActive() {
  return !!dashboardViewOnly;
}
function formatDateLabel(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return date.toISOString().slice(0, 10);
}
function formatHumanDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  return date.toLocaleDateString('es-CL');
}
function isRangeWithinOneYear(from, to) {
  if (!(from instanceof Date) || !(to instanceof Date)) return false;
  const diffMs = to.getTime() - from.getTime();
  const dias = diffMs / (1000 * 60 * 60 * 24);
  return dias <= 365;
}

function closeExportRangeModal() {
  if (exportRangeModal) {
    exportRangeModal.classList.add('hidden');
    exportRangeModal.setAttribute('aria-hidden', 'true');
  }
}

function openExportRangeModal(tipo = 'excel') {
  currentExportType = tipo || 'excel';
  if (!exportRangeModal || !exportRangeWarning || !exportDesde || !exportHasta) return;
  exportRangeWarning.classList.add('hidden');
  exportRangeWarning.textContent = 'Excede el rango máximo permitido de un año';
  exportDesde.value = '';
  exportHasta.value = '';
  if (btnExportDescargar) btnExportDescargar.disabled = false;
  exportRangeModal.classList.remove('hidden');
  exportRangeModal.setAttribute('aria-hidden', 'false');
}

function validarRangoExportacion() {
  if (!exportRangeWarning || !exportDesde || !exportHasta) return { valido: false };
  exportRangeWarning.classList.add('hidden');
  let mensaje = '';
  const desdeStr = exportDesde.value;
  const hastaStr = exportHasta.value;
  if (!desdeStr || !hastaStr) {
    mensaje = 'Debe seleccionar ambas fechas (Desde y Hasta).';
  }
  const desde = desdeStr ? new Date(desdeStr) : null;
  const hasta = hastaStr ? new Date(hastaStr) : null;
  if (!mensaje) {
    if (!desde || !hasta || Number.isNaN(desde.getTime()) || Number.isNaN(hasta.getTime()) || desde > hasta) {
      mensaje = 'El rango de fechas es inválido (Desde debe ser menor o igual que Hasta).';
    }
  }
  if (!mensaje && desde && hasta) {
    const diffMs = hasta.getTime() - desde.getTime();
    const dias = diffMs / (1000 * 60 * 60 * 24);
    if (dias > 365) {
      mensaje = 'Excede el rango máximo permitido de un año';
    }
  }
  if (mensaje) {
    exportRangeWarning.textContent = mensaje;
    exportRangeWarning.classList.remove('hidden');
    return { valido: false };
  }
  return { valido: true, desdeStr, hastaStr };
}

function filtrarTasksPorRango(desdeStr, hastaStr) {
  const desde = new Date(desdeStr);
  const hasta = new Date(hastaStr);
  // asegurar inclusión del día final
  hasta.setHours(23, 59, 59, 999);
  const scoped = getScopedTasksForExport();
  const filtered = scoped.filter((task) => {
    const dateCandidate = task.createdAt || task.updatedAt || task.due || task.fechaCreacion;
    if (!dateCandidate) return false;
    const d = new Date(dateCandidate);
    if (Number.isNaN(d.getTime())) return false;
    return d >= desde && d <= hasta;
  });
  return filtered;
}

function exportarExcelConRango(desdeStr, hastaStr) {
  const filtered = filtrarTasksPorRango(desdeStr, hastaStr);
  const rows = buildExportRows(filtered);
  const blob = createXLSXBlob(rows);
  download(
    `tasks-${state.tenant}.xlsx`,
    blob,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
}

function exportarCsvConRango(desdeStr, hastaStr) {
  const filtered = filtrarTasksPorRango(desdeStr, hastaStr);
  const rows = buildExportRows(filtered);
  const csv = rows.map((row) => row.map(csvEscape).join(',')).join('\n');
  download(`tasks-${state.tenant}.csv`, csv, 'text/csv');
}

function getDashboardRange() {
  const { from, to } = getDashboardDateBounds();
  return { from, to };
}
function getDashboardFiltersSnapshot() {
  const rangePreset = DASHBOARD_RANGE_PRESETS.find((p) => p.value === dashboardFilters.range);
  return {
    ...dashboardFilters,
    rangeLabel: rangePreset?.label || dashboardFilters.range,
  };
}
function buildDashboardTasksRows(tasks = []) {
  return tasks.map((t) => {
    const prima = getPolicyPremiumValue(t);
    const commission = getDashboardCommissionValue(t);
    const date = new Date(t.updatedAt || t.createdAt || Date.now());
    return {
      ID: t.id || t.title || '',
      Cliente: t.client || t.case?.personal?.firstName || '',
      Ejecutivo: t.assigneeName || t.assignee || t.asignadoAEmail || '',
      Compañía: t.insurer || '',
      Ramo: t.type || '',
      Estado: normalizeTaskStatus(t.status || ''),
      Canal: t.canalIngreso || '',
      'Prima (UF)': Number(prima) || 0,
      'Comisión (UF)': Number(commission) || 0,
      Fecha: formatDateLabel(date),
    };
  });
}
function buildDashboardPeriodRows(buckets = [], metaPrima = 0) {
  return buckets.map((b) => ({
    Periodo: b.label,
    Desde: formatDateLabel(b.start),
    Hasta: formatDateLabel(b.end),
    'Prima total (UF)': Number(b.prima) || 0,
    'Comisión corredor (UF)': Number(b.commission) || 0,
    'Meta Prima (UF)': Number(metaPrima) || 0,
    'N° negocios': Number(b.count) || 0,
  }));
}
function buildDashboardSummaryRows(kpis, filtros, from, to) {
  return [
    { Concepto: 'Rango', Valor: `${formatHumanDate(from)} - ${formatHumanDate(to)}` },
    { Concepto: 'Preset de rango', Valor: filtros.rangeLabel || filtros.range || '' },
    { Concepto: 'Ejecutivo', Valor: filtros.executive === DASHBOARD_EXECUTIVE_ALL ? 'Todos' : filtros.executive || 'Todos' },
    { Concepto: 'Compañía', Valor: filtros.company === 'all' ? 'Todas' : filtros.company || 'Todas' },
    { Concepto: 'Producto', Valor: filtros.product === 'all' ? 'Todos' : filtros.product || 'Todos' },
    { Concepto: 'Prima emitida (UF)', Valor: Number(kpis.prima) || 0 },
    { Concepto: 'Comisión corredor (UF)', Valor: Number(kpis.commission) || 0 },
    { Concepto: 'Negocios ganados', Valor: Number(kpis.dealsWon) || 0 },
    { Concepto: 'Meta Prima (UF)', Valor: Number(kpis.metaPrima) || 0 },
    { Concepto: '% Cumplimiento Prima', Valor: kpis.primaPct ?? '' },
    { Concepto: '% Recaudación', Valor: kpis.commissionPct ?? '' },
  ];
}
function exportarDetalleNegociosDashboard() {
  const { from, to } = getDashboardRange();
  if (!from || !to) {
    alert('Debe seleccionar un rango de fechas válido en el dashboard antes de exportar.');
    return;
  }
  if (!isRangeWithinOneYear(from, to)) {
    alert('Excede el rango máximo permitido de un año');
    return;
  }
  const data = getFilteredDealsForDashboard();
  const detailRows = getDashboardDetailDataset(data);
  if (!Array.isArray(detailRows) || !detailRows.length) {
    alert('No hay datos en Detalle de negocios para exportar con los filtros actuales.');
    return;
  }
  if (typeof XLSX === 'undefined' || !XLSX?.utils) {
    alert('No se encontró la librería de exportación XLSX.');
    return;
  }
  const headers = DASHBOARD_DETAIL_COLUMNS.map((col) => col.header);
  const values = detailRows.map((row) => DASHBOARD_DETAIL_COLUMNS.map((col) => row[col.key] ?? ''));
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...values]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Detalle de negocios');

  const fileName = `dashboard-detalle-${formatDateLabel(from)}_${formatDateLabel(to)}.xlsx`;
  XLSX.writeFile(workbook, fileName);
}
function exportarDashboardExcel() {
  exportarDetalleNegociosDashboard();
}

if (btnExport)
  btnExport.addEventListener('click', (e) => {
    e.preventDefault();
    if (isDashboardActive()) {
      exportarDetalleNegociosDashboard();
    } else {
      openExportRangeModal('excel');
    }
  });

if (btnExportCerrar) {
  btnExportCerrar.addEventListener('click', (e) => {
    e.preventDefault();
    closeExportRangeModal();
  });
}
if (btnExportDescargar) {
  btnExportDescargar.addEventListener('click', (e) => {
    e.preventDefault();
    const result = validarRangoExportacion();
    if (!result.valido) return;
    if (currentExportType === 'csv') exportarCsvConRango(result.desdeStr, result.hastaStr);
    else exportarExcelConRango(result.desdeStr, result.hastaStr);
    closeExportRangeModal();
  });
}

if (exportRangeModal) {
  exportRangeModal.addEventListener('click', (ev) => {
    if (ev.target === exportRangeModal) closeExportRangeModal();
  });
}
const btnBackup = document.getElementById('btnBackup');
if (btnBackup)
  btnBackup.addEventListener('click', () =>
    download(
      `backup-${state.tenant}-${Date.now()}.json`,
      JSON.stringify({ tenant: state.tenant, tasks: getScopedTasksForExport() }, null, 2)
    )
  );

  const exportFields = [
    { label: 'ID del negocio', get: (t) => t.title || '' },
    { label: 'Ramo', get: (t) => t.type || '' },
    { label: 'Prioridad', get: (t) => t.priority || '' },
    { label: 'Estado', get: (t) => t.status || '' },
    { label: 'Vencimiento', get: (t) => t.due || '' },
    { label: 'Asignado a', get: (t) => t.assignee || '' },
    { label: 'Tags', get: (t) => t.tags?.[0] || '' },
    { label: 'Cliente - Nombres', get: (_, ctx) => ctx.personal.firstName || '' },
    { label: 'Cliente - Apellidos', get: (_, ctx) => ctx.personal.lastName || '' },
    { label: 'Cliente - Correo', get: (_, ctx) => ctx.personal.email || '' },
    { label: 'Cliente - RUT', get: (_, ctx) => ctx.personal.rut || '' },
    { label: 'Cliente - Teléfono', get: (_, ctx) => ctx.personal.phone || '' },
    { label: 'Cliente - Tipo', get: (t) => t.customerType || '' },
    { label: 'Asegurado - Nombres', get: (_, ctx) => ctx.pago.firstName || '' },
    { label: 'Asegurado - Apellidos', get: (_, ctx) => ctx.pago.lastName || '' },
    { label: 'Asegurado - Correo', get: (_, ctx) => ctx.pago.email || '' },
    { label: 'Asegurado - RUT', get: (_, ctx) => ctx.pago.rut || '' },
    { label: 'Asegurado - Teléfono', get: (_, ctx) => ctx.pago.phone || '' },
    { label: 'Asegurado - Tipo', get: (_, ctx) => ctx.pago.tipo || '' },
    { label: 'Póliza - Compañía', get: (_, ctx) => ctx.poliza.insurer || '' },
    { label: 'Póliza - N° de póliza', get: (_, ctx) => ctx.poliza.number || '' },
    { label: 'Póliza - Prima total (UF)', get: (_, ctx) => ctx.poliza.premiumUF ?? '' },
    { label: 'Póliza - % Comisión corredor', get: (_, ctx) => ctx.poliza.commissionPct ?? '' },
    { label: 'Póliza - Total comisión corredor (UF)', get: (_, ctx) => ctx.poliza.commissionUF ?? '' },
    { label: 'Póliza - Comisión de Agente', get: (_, ctx) => ctx.poliza.agentCommission || '' },
    { label: 'Vehículo - Marca', get: (_, ctx) => ctx.vehiculo.make || '' },
    { label: 'Vehículo - Modelo', get: (_, ctx) => ctx.vehiculo.model || '' },
    { label: 'Vehículo - Año', get: (_, ctx) => ctx.vehiculo.year || '' },
    { label: 'Vehículo - Patente', get: (_, ctx) => ctx.vehiculo.plate || '' },
    { label: 'Vehículo - Color', get: (_, ctx) => ctx.vehiculo.color || '' },
    { label: 'Vehículo - Chasis', get: (_, ctx) => ctx.vehiculo.chassis || '' },
    { label: 'Vehículo - Tipo de combustible', get: (_, ctx) => ctx.vehiculo.fuel || '' },
    { label: 'Vehículo - Tipo de vehículo', get: (_, ctx) => ctx.vehiculo.type || '' },
    {
      label: 'Vehículo - Flota',
      get: (_, ctx) => (ctx.vehiculo.fleet ? 'Sí' : 'No'),
    },
    { label: 'Dirección del riesgo', get: (_, ctx) => ctx.riesgo.address || '' },
    { label: 'Dirección del riesgo - Comuna', get: (_, ctx) => ctx.riesgo.commune || '' },
    { label: 'Dirección del riesgo - Región', get: (_, ctx) => ctx.riesgo.region || '' },
    {
      label: 'Dirección del riesgo - Tipo de inmueble',
      get: (_, ctx) => ctx.riesgo.propertyType || '',
    },
    // Renovación
    { label: 'Renovación - Estado', get: (_, ctx) => ctx.renovacion?.estadoRenovacion || '' },
    { label: 'Renovación - Fecha preaviso', get: (_, ctx) => ctx.renovacion?.fechaPreaviso || '' },
    { label: 'Renovación - Fecha decisión', get: (_, ctx) => ctx.renovacion?.fechaDecision || '' },
    {
      label: 'Renovación - Motivo pérdida/anulación',
      get: (_, ctx) =>
        renNeedsReason(ctx.renovacion?.estadoRenovacion || '')
          ? ctx.renovacion?.motivoPerdidaOAnulacion || ''
          : '',
    },
    { label: 'Renovación - Comentario', get: (_, ctx) => ctx.renovacion?.comentario || '' },

    // Recaudación (último pago)
    { label: 'Recaudación - Estado último pago', get: (_, ctx) => ctx.lastPago?.estadoRecaudacion || '' },
    { label: 'Recaudación - Método último pago', get: (_, ctx) => ctx.lastPago?.metodoPago || '' },
    { label: 'Recaudación - Fecha último cobro', get: (_, ctx) => ctx.lastPago?.fechaCobro || '' },
    {
      label: 'Recaudación - Motivo reversa',
      get: (_, ctx) =>
        recNeedsReason(ctx.lastPago?.estadoRecaudacion || '')
          ? ctx.lastPago?.motivoReversa || ''
          : '',
    },
  ];

  function buildCaseContext(task) {
    const c = normalizeCase(task.case || {});
    const pagos = Array.isArray(c.pagos) ? c.pagos : [];
    const lastPago = pickLatestPago(pagos);
    return {
      personal: c.personal || {},
      pago: c.pago || {},
      poliza: c.poliza || {},
      vehiculo: c.vehiculo || {},
      riesgo: c.riesgo || {},
      renovacion: c.renovacion || null,
      pagos,
      lastPago,
    };
  }

  function buildExportRows(tasks) {
    const headers = exportFields.map((f) => f.label);
    const rows = tasks.map((task) => {
      const ctx = buildCaseContext(task);
      return exportFields.map((field) => {
        const value = field.get(task, ctx);
        return value == null ? '' : value;
      });
    });
    return [headers, ...rows];
  }

  const csvEscape = (val) => `"${String(val ?? '').replace(/"/g, '""')}"`;

const btnCSV = document.getElementById('btnCSV');
if (btnCSV)
  btnCSV.addEventListener('click', (e) => {
    e.preventDefault();
    openExportRangeModal('csv');
  });

  function createXLSXBlob(rows) {
    const sheetXml = buildSheetXML(rows);
    const files = [
      {
        name: '[Content_Types].xml',
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
      },
      {
        name: '_rels/.rels',
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
      },
      {
        name: 'xl/workbook.xml',
        data: `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Negocios" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
      },
      {
        name: 'xl/_rels/workbook.xml.rels',
        data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
      },
      {
        name: 'xl/styles.xml',
        data: `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="1">
    <fill>
      <patternFill patternType="none"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>`,
      },
      {
        name: 'xl/worksheets/sheet1.xml',
        data: sheetXml,
      },
    ];

    return createZipBlob(
      files,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
  }

  function buildSheetXML(rows) {
    const columnCount = rows[0]?.length || 1;
    const lastColumn = columnName(columnCount - 1);
    const lastRow = Math.max(rows.length, 1);
    const dimension = `A1:${lastColumn}${lastRow}`;
    const rowsXml = rows
      .map((row, i) => {
        const cells = row
          .map((value, j) => {
            const cellRef = `${columnName(j)}${i + 1}`;
            const safe = escapeXml(String(value ?? ''));
            return `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${safe}</t></is></c>`;
          })
          .join('');
        return `<row r="${i + 1}">${cells}</row>`;
      })
      .join('');

    return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="${dimension}"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData>${rowsXml}</sheetData>
</worksheet>`;
  }

  function columnName(idx) {
    let name = '';
    let n = idx;
    while (n >= 0) {
      name = String.fromCharCode((n % 26) + 65) + name;
      n = Math.floor(n / 26) - 1;
    }
    return name;
  }

  function escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function createZipBlob(files, mime) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    const now = new Date();
    const { dosTime, dosDate } = toDosDateTime(now);
    let offset = 0;

    files.forEach((file) => {
      const fileNameBytes = encoder.encode(file.name);
      let dataBytes;
      if (file.data instanceof Uint8Array) dataBytes = file.data;
      else if (file.data instanceof ArrayBuffer)
        dataBytes = new Uint8Array(file.data);
      else dataBytes = encoder.encode(String(file.data));

      const crc = crc32(dataBytes);
      const localHeader = new Uint8Array(30 + fileNameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(6, 0, true);
      localView.setUint16(8, 0, true);
      localView.setUint16(10, dosTime, true);
      localView.setUint16(12, dosDate, true);
      localView.setUint32(14, crc, true);
      localView.setUint32(18, dataBytes.length, true);
      localView.setUint32(22, dataBytes.length, true);
      localView.setUint16(26, fileNameBytes.length, true);
      localView.setUint16(28, 0, true);
      localHeader.set(fileNameBytes, 30);

      localParts.push(localHeader, dataBytes);

      const centralHeader = new Uint8Array(46 + fileNameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(8, 0, true);
      centralView.setUint16(10, 0, true);
      centralView.setUint16(12, dosTime, true);
      centralView.setUint16(14, dosDate, true);
      centralView.setUint32(16, crc, true);
      centralView.setUint32(20, dataBytes.length, true);
      centralView.setUint32(24, dataBytes.length, true);
      centralView.setUint16(28, fileNameBytes.length, true);
      centralView.setUint16(30, 0, true);
      centralView.setUint16(32, 0, true);
      centralView.setUint16(34, 0, true);
      centralView.setUint16(36, 0, true);
      centralView.setUint32(38, 0, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(fileNameBytes, 46);

      centralParts.push(centralHeader);

      offset += localHeader.length + dataBytes.length;
    });

    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const centralOffset = offset;
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, files.length, true);
    eocdView.setUint16(10, files.length, true);
    eocdView.setUint32(12, centralSize, true);
    eocdView.setUint32(16, centralOffset, true);
    eocdView.setUint16(20, 0, true);

    const parts = [...localParts, ...centralParts, eocd];
    return new Blob(parts, {
      type: mime || 'application/zip',
    });
  }

  function toDosDateTime(date) {
    const year = date.getFullYear();
    const dosYear = Math.max(year - 1980, 0);
    const dosDate =
      (dosYear << 9) | ((date.getMonth() + 1) << 5) | date.getDate();
    const dosTime =
      (date.getHours() << 11) |
      (date.getMinutes() << 5) |
      Math.floor(date.getSeconds() / 2);
    return { dosDate, dosTime };
  }

  const crcTable = (() => {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
      let c = i;
      for (let k = 0; k < 8; k++) {
        if (c & 1) c = 0xedb88320 ^ (c >>> 1);
        else c >>>= 1;
      }
      table[i] = c >>> 0;
    }
    return table;
  })();

  function crc32(input) {
    let crc = 0xffffffff;
    for (let i = 0; i < input.length; i++) {
      const byte = input[i];
      crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  const importButton = document.getElementById('importButton');
  const importFileInput = document.getElementById('importFileInput');
  const btnDashboardToggleGoal = document.getElementById('btnDashboardToggleGoal');
  const dashboardGoalPanel = document.getElementById('dashboardGoalPanel');
  const btnDashboardViewDeals = document.getElementById('btnDashboardViewDeals');
  const dashboardActionsRow = document.getElementById('dashboardActionsRow');
  const IMPORT_COLUMNS = {
    RAMO: 'Ramo',
    PRIORIDAD: 'Prioridad',
    ESTADO: 'Estado',
    TAGS: 'Tags',
    CLI_NOMBRES: 'Cliente - Nombres',
    CLI_APELLIDOS: 'Cliente - Apellidos',
    CLI_CORREO: 'Cliente - Correo',
    CLI_RUT: 'Cliente - RUT',
    CLI_TELEFONO: 'Cliente - Teléfono',
    CLI_TIPO: 'Cliente - Tipo',
    ASEG_TIPO: 'Asegurado - Tipo',
    VEH_MARCA: 'Vehículo - Marca',
    VEH_MODELO: 'Vehículo - Modelo',
    VEH_ANIO: 'Vehículo - Año',
    VEH_PATENTE: 'Vehículo - Patente',
    VEH_TIPO: 'Vehículo - Tipo de vehículo',
    VEH_FLOTA: 'Vehículo - Flota',
    RIESGO_DIRECCION: 'Dirección del riesgo',
    RIESGO_COMUNA: 'Dirección del riesgo - Comuna',
    RIESGO_REGION: 'Dirección del riesgo - Región',
    RIESGO_TIPO: 'Dirección del riesgo - Tipo de inmueble',
  };
  const REQUIRED_IMPORT_FIELDS = [
    IMPORT_COLUMNS.CLI_NOMBRES,
    IMPORT_COLUMNS.CLI_APELLIDOS,
    IMPORT_COLUMNS.CLI_CORREO,
    IMPORT_COLUMNS.CLI_RUT,
    IMPORT_COLUMNS.CLI_TELEFONO,
    IMPORT_COLUMNS.CLI_TIPO,
    IMPORT_COLUMNS.ASEG_TIPO,
  ];
  if (importButton && importFileInput) {
    importButton.addEventListener('click', () => {
      importFileInput.value = '';
      importFileInput.click();
    });
    importFileInput.addEventListener('change', handleImportFileSelected);
  }

  function handleImportFileSelected(event) {
    const input = event.target;
    const file = input.files?.[0];
    if (!file) return;
    if (!file.name?.toLowerCase().endsWith('.xlsx')) {
      alert('Por favor selecciona un archivo con extensión .xlsx');
      input.value = '';
      return;
    }
    if (typeof XLSX === 'undefined') {
      alert('No se pudo cargar el lector de Excel (XLSX). Intenta nuevamente.');
      input.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onerror = () => {
      alert('Ocurrió un error al leer el archivo seleccionado.');
      input.value = '';
    };
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames?.[0];
        if (!firstSheetName) {
          alert('El archivo no contiene hojas con datos.');
          return;
        }
        const sheet = workbook.Sheets[firstSheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        processImportedRows(rows);
      } catch (err) {
        console.error('Error procesando archivo XLSX', err);
        alert('No se pudo procesar el archivo. Verifica que sea un .xlsx válido.');
      } finally {
        input.value = '';
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function processImportedRows(rawRows = []) {
    const entries = (rawRows || [])
      .map((row, index) => ({ row, index }))
      .filter((entry) => !isImportRowEmpty(entry.row));
    if (!entries.length) {
      alert('El archivo no contiene filas con datos para importar.');
      return;
    }
    const missingHeaders = findMissingImportHeaders(entries.map((entry) => entry.row));
    if (missingHeaders.length) {
      alert(
        `Faltan columnas obligatorias en el archivo: ${missingHeaders.join(
          ', '
        )}`
      );
      return;
    }
    const created = [];
    const errors = [];
    const seenVehicleKeys = new Set();
    entries.forEach(({ row, index }) => {
      const rowNumber = index + 2; // encabezados en fila 1
      const result = validateAndMapRow(row, rowNumber, seenVehicleKeys);
      if (result.ok) created.push(result);
      else errors.push(result);
    });
    if (created.length) {
      createBusinessesFromImport(created);
    }
    showImportSummary(created, errors);
  }

  function validateAndMapRow(row, rowNumber, seenVehicleKeys) {
    const errors = [];
    const value = (key) => normalizeImportString(row[key]);
    const ramo = value(IMPORT_COLUMNS.RAMO) || DEFAULT_TASK_TYPE;
    const prioridad = value(IMPORT_COLUMNS.PRIORIDAD) || 'Media';
    const tags = parseTags(value(IMPORT_COLUMNS.TAGS));

    const nombres = value(IMPORT_COLUMNS.CLI_NOMBRES);
    const apellidos = value(IMPORT_COLUMNS.CLI_APELLIDOS);
    const correo = value(IMPORT_COLUMNS.CLI_CORREO);
    const rut = value(IMPORT_COLUMNS.CLI_RUT);
    const telefono = value(IMPORT_COLUMNS.CLI_TELEFONO);
    const cliTipo = normalizeCustomerType(value(IMPORT_COLUMNS.CLI_TIPO));
    const asegTipo = normalizeCustomerType(value(IMPORT_COLUMNS.ASEG_TIPO));

    if (!nombres) errors.push('Falta Cliente - Nombres');
    if (!apellidos) errors.push('Falta Cliente - Apellidos');
    if (!correo) errors.push('Falta Cliente - Correo');
    if (!rut) errors.push('Falta Cliente - RUT');
    if (!telefono) errors.push('Falta Cliente - Teléfono');
    if (!cliTipo) errors.push('Falta Cliente - Tipo');
    if (!asegTipo) errors.push('Falta Asegurado - Tipo');
    if (correo && !isValidEmailBasic(correo)) {
      errors.push('Correo inválido');
    }

    const patenteRaw = value(IMPORT_COLUMNS.VEH_PATENTE).toUpperCase();
    const patenteNormalized = normalizePlate(patenteRaw);
    const anio = value(IMPORT_COLUMNS.VEH_ANIO);
    let pendingVehicleKey = null;
    if (patenteNormalized && anio) {
      const key = `${patenteNormalized}|${anio}`;
      if (seenVehicleKeys.has(key)) {
        errors.push('Duplicado por Vehículo - Patente + Año');
      } else {
        pendingVehicleKey = key;
      }
    }

    if (errors.length) {
      return {
        ok: false,
        rowNumber,
        row,
        errorMessage: errors.join('; '),
      };
    }
    if (pendingVehicleKey) seenVehicleKeys.add(pendingVehicleKey);

    const identity = getCurrentUserIdentity();
    const assigneeEmail =
      identity.email || state.user?.email || state.user?.correo || '';
    const clientFullName = [nombres, apellidos].filter(Boolean).join(' ').trim();
    const vehiculoFlota = parseFleetValue(value(IMPORT_COLUMNS.VEH_FLOTA));
    const vehiculoTipo = value(IMPORT_COLUMNS.VEH_TIPO);
    const vehiculoMarca = value(IMPORT_COLUMNS.VEH_MARCA);
    const vehiculoModelo = value(IMPORT_COLUMNS.VEH_MODELO);
    const riesgoDireccion = value(IMPORT_COLUMNS.RIESGO_DIRECCION);
    const riesgoComuna = value(IMPORT_COLUMNS.RIESGO_COMUNA);
    const riesgoRegion = value(IMPORT_COLUMNS.RIESGO_REGION);
    const riesgoTipo = value(IMPORT_COLUMNS.RIESGO_TIPO);

    const caseData = {
      personal: {
        firstName: nombres,
        lastName: apellidos,
        rut,
        email: correo,
        phone: telefono,
        tipo: cliTipo || 'natural',
      },
      pago: {
        firstName: '',
        lastName: '',
        rut: '',
        email: '',
        phone: '',
        tipo: asegTipo || 'natural',
        usarPersonales: false,
      },
      poliza: {
        insurer: DEFAULT_INSURER,
        ramo,
      },
      vehiculo: {
        make: vehiculoMarca,
        model: vehiculoModelo,
        year: anio,
        plate: patenteNormalized,
        type: vehiculoTipo,
        fleet: vehiculoFlota,
      },
      riesgo: {
        address: riesgoDireccion,
        commune: riesgoComuna,
        region: riesgoRegion,
        propertyType: riesgoTipo,
      },
      docs: [],
      comments: [],
      renovacion: crearRenovacion({
        estadoRenovacion: ESTADOS_RENOVACION[0] || '',
      }),
      pagos: [],
    };

    const canalIngreso = CHANNEL_TRADICIONAL;
    const initialStatus =
      canalIngreso === CHANNEL_RENOVACION ? 'En renovación' : 'Abierta';
    const businessData = {
      type: ramo || DEFAULT_TASK_TYPE,
      priority: prioridad || 'Media',
      status: initialStatus,
      estadoInicial: initialStatus,
      tags,
      client: clientFullName,
      email: correo,
      phone: telefono,
      rut,
      customerType: cliTipo || 'natural',
      canalIngreso,
      assignee: assigneeEmail,
      case: caseData,
    };

    return {
      ok: true,
      rowNumber,
      businessData,
    };
  }

  function createBusinessesFromImport(items) {
    if (!Array.isArray(items) || !items.length) return;
    const actorName = getActorName();
    items.forEach(({ businessData, rowNumber }) => {
      if (!businessData) return;
      businessData.case = businessData.case || {};
      if (!Array.isArray(businessData.case.comments)) businessData.case.comments = [];
      businessData.case.comments.unshift({
        by: actorName,
        ts: Date.now(),
        text: `Negocio importado desde Excel (fila ${rowNumber})`,
      });
      upsertTask(businessData, { silent: true, deferSave: true });
    });
    saveTasks();
    applyFilters();
  }

  function showImportSummary(created, errors) {
    const createdCount = Array.isArray(created) ? created.length : 0;
    const errorCount = Array.isArray(errors) ? errors.length : 0;
    alert(
      `Importación finalizada.\nNegocios creados: ${createdCount}\nFilas con errores: ${errorCount}`
    );
    if (createdCount) {
      console.table(
        created.map((item) => ({
          fila: item.rowNumber,
          cliente: item.businessData?.client || '',
          estado: item.businessData?.status || '',
        }))
      );
    }
    if (errorCount) {
      console.table(
        errors.map((err) => ({
          fila: err.rowNumber,
          error: err.errorMessage,
        }))
      );
    }
  }

  function isImportRowEmpty(row) {
    if (!row) return true;
    return Object.values(row).every((val) => !normalizeImportString(val));
  }

  function findMissingImportHeaders(rows) {
    const seen = new Set();
    rows.forEach((row) => {
      Object.keys(row || {}).forEach((key) => {
        if (key) seen.add(key);
      });
    });
    return REQUIRED_IMPORT_FIELDS.filter((header) => !seen.has(header));
  }

  function normalizeImportString(value) {
    if (value === undefined || value === null) return '';
    if (typeof value === 'string') return value.trim();
    return String(value).trim();
  }

  function normalizeCustomerType(value) {
    const normalized = normalizeImportString(value).toLowerCase();
    if (!normalized) return '';
    if (normalized.includes('juri') || normalized.includes('empresa')) return 'juridica';
    return 'natural';
  }

  function parseTags(value) {
    if (!value) return [];
    return value
      .split(/[,;]+/)
      .map((tag) => tag.trim())
      .filter(Boolean);
  }

  function parseFleetValue(value) {
    if (!value) return false;
    const v = value.toLowerCase();
    return ['si', 'sí', 'true', '1', 'yes', 'y', 'x'].includes(v);
  }

  function normalizePlate(value) {
    if (!value) return '';
    return value.replace(/[^0-9A-Z]/gi, '').toUpperCase();
  }

  function isValidEmailBasic(email) {
    if (!email) return false;
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  const btnDashboardToggle = document.getElementById('btnDashboardToggle');
  if (btnDashboardToggle) {
    btnDashboardToggle.addEventListener('click', () => {
      dashboardViewOnly = !dashboardViewOnly;
      state.ui.isDashboardVisible = dashboardViewOnly;
      saveDashboardViewState();
      applyDashboardView();
      updateFiltersCardVisibility();
      if (dashboardViewOnly) {
        refreshDashboard();
      }
    });
  }
  if (btnDashboardViewDeals && btnDashboardToggle) {
    btnDashboardViewDeals.addEventListener('click', () => {
      if (dashboardViewOnly) btnDashboardToggle.click();
    });
  }
  if (btnDashboardToggleGoal) {
    btnDashboardToggleGoal.addEventListener('click', () => {
      if (!canManageGlobalGoals()) return;
      openMetaModal();
    });
  }
  const closeMetaModalBtn = document.getElementById('closeMetaModal');
  const cancelMetaBtn = document.getElementById('btnMetaCancel');
  if (closeMetaModalBtn) closeMetaModalBtn.addEventListener('click', closeMetaModal);
  if (cancelMetaBtn) cancelMetaBtn.addEventListener('click', closeMetaModal);
  if (dashboardActionsRow) {
    dashboardActionsRow.classList.add('hidden');
  }

  // Ventana usuarios
  const btnUsers = document.getElementById('btnUsers');
  if (btnUsers) {
    btnUsers.addEventListener('click', () => {
      const url = `admin-users.html?tenant=${encodeURIComponent(state.tenant || '')}`;
      window.location.href = url;
    });
  }
  const btnClientes = document.getElementById('btnClientes');
  if (btnClientes) {
    btnClientes.addEventListener('click', () => {
      const url = `admin-clients.html?tenant=${encodeURIComponent(state.tenant || '')}`;
      window.location.href = url;
    });
  }
})();

/* ===========================
   Interior del caso (Editor)
=========================== */
(function () {
  const tbody = document.getElementById('tbody');
  const caseTitleEl = document.getElementById('caseTitle');
  const progress = document.getElementById('globalProgress');
  const infoToast = document.getElementById('infoToast');

  const ce = {
    root: document.getElementById('caseEditor'),

    // Datos personales
    firstName: document.getElementById('ce_firstName'),
    lastName: document.getElementById('ce_lastName'),
    email: document.getElementById('ce_email'),
    rut: document.getElementById('ce_rut'),
    phone: document.getElementById('ce_phone'),

    // Pago / Asegurado
    payType: document.getElementById('ce_payType'),
    usePersonal: document.getElementById('ce_usePersonal'),
    payFirstName: document.getElementById('ce_payFirstName'),
    payLastName: document.getElementById('ce_payLastName'),
    payRut: document.getElementById('ce_payRut'),
    payEmail: document.getElementById('ce_payEmail'),
    payPhone: document.getElementById('ce_payPhone'),
    policyNumber: document.getElementById('ce_policyNumber'),
    policyIssueDate: document.getElementById('ce_policyIssueDate'),
    policyPremium: document.getElementById('ce_policyPremium'),
    policyCommissionPct: document.getElementById('ce_policyCommissionPct'),
    policyCommissionTotal: document.getElementById('ce_policyCommissionTotal'),
    policyAgentCommission: document.getElementById('ce_policyAgentCommission'),
    vehicleMake: document.getElementById('ce_vehicleMake'),
    vehicleModel: document.getElementById('ce_vehicleModel'),
    vehicleYear: document.getElementById('ce_vehicleYear'),
    vehiclePlate: document.getElementById('ce_vehiclePlate'),
    vehicleColor: document.getElementById('ce_vehicleColor'),
    vehicleChassis: document.getElementById('ce_vehicleChassis'),
    vehicleFuel: document.getElementById('ce_vehicleFuel'),
    vehicleType: document.getElementById('ce_vehicleType'),
    vehicleFleet: document.getElementById('ce_vehicleFleet'),
    riskAddress: document.getElementById('ce_riskAddress'),
    riskCommune: document.getElementById('ce_riskCommune'),
    riskRegion: document.getElementById('ce_riskRegion'),
    riskPropertyType: document.getElementById('ce_riskPropertyType'),
    useRiskClient: document.getElementById('ce_useRiskClient'),

    // Documentos
    files: document.getElementById('ce_files'),
    filesList: document.getElementById('ce_filesList'),

    // Comentarios
    comment: document.getElementById('ce_comment'),
    comments: document.getElementById('ce_comments'),
    btnAddComment: document.getElementById('ce_addComment'),

    // Botones
    btnSave: document.getElementById('ce_save'),
    btnExit: document.getElementById('ce_exit'),

    // Datos de la tarea (editables desde interior del caso)
    taskTitle: document.getElementById('ce_taskTitle'),
    taskType: document.getElementById('ce_taskType'),
    taskPriority: document.getElementById('ce_taskPriority'),
    taskStatus: document.getElementById('ce_taskStatus'),
    taskDue: document.getElementById('ce_taskDue'),
    taskAssignee: document.getElementById('ce_taskAssignee'),
    taskPolicy: document.getElementById('ce_taskPolicy'),
    taskInsurer: document.getElementById('ce_taskInsurer'),
    taskTags: document.getElementById('ce_taskTags'),
    taskChannel: document.getElementById('ce_taskChannel'),
    customerType: document.getElementById('ce_customerType'),
    clientAddress: document.getElementById('ce_clientAddress'),
    clientCommune: document.getElementById('ce_clientCommune'),
    clientRegion: document.getElementById('ce_clientRegion'),

    // Renovación
    renPreaviso: document.getElementById('ce_renPreaviso'),
    renDecision: document.getElementById('ce_renDecision'),
    renMotivo: document.getElementById('ce_renMotivo'),
    renComentario: document.getElementById('ce_renComentario'),

  };

  attachRutSanitizer(ce.rut);
  attachRutSanitizer(ce.payRut);
  attachPhoneSanitizer(ce.phone);
  attachPhoneSanitizer(ce.payPhone);
  attachYearSanitizer(ce.vehicleYear);
  attachPlateSanitizer(ce.vehiclePlate);

  let ceBaseline = null;
  let ceDirty = false;
  let ceCurrentTask = null;
  let caseCardObserver = null;
  let caseCardResizeAttached = false;

  const syncCaseCarouselHeight = () => {
    if (!ce.root) return;
    const carousel = ce.root.querySelector('.case-carousel');
    const baseCard = ce.root.querySelector('.case-card-base');
    if (!carousel || !baseCard) return;
    const isMobile =
      window.matchMedia && window.matchMedia('(max-width: 768px)').matches;
    if (!isMobile) {
      carousel.style.removeProperty('--case-card-height');
      return;
    }
    const height = baseCard.getBoundingClientRect().height;
    if (height > 0) {
      carousel.style.setProperty('--case-card-height', `${Math.round(height)}px`);
    }
  };

  const setupCaseCarouselHeightSync = () => {
    if (!ce.root) return;
    const carousel = ce.root.querySelector('.case-carousel');
    const baseCard = ce.root.querySelector('.case-card-base');
    if (!carousel || !baseCard) return;
    syncCaseCarouselHeight();
    if (window.ResizeObserver) {
      if (!caseCardObserver) {
        caseCardObserver = new ResizeObserver(() => {
          syncCaseCarouselHeight();
        });
      }
      caseCardObserver.disconnect();
      caseCardObserver.observe(baseCard);
    } else if (!caseCardResizeAttached) {
      window.addEventListener('resize', syncCaseCarouselHeight);
      caseCardResizeAttached = true;
    }
  };

  function renderStatusOptions(task) {
    if (!ce.taskStatus || !task) return;
    const channelValue = normalizeChannel(ce.taskChannel?.value || task.canalIngreso || '');
    const context = { ...task, canalIngreso: channelValue };
    const allowed = getAllowedNextStates(context);
    const currentStatus = normalizeTaskStatus(task.status || 'Abierta');
    const optionOrder = [currentStatus];
    allowed.forEach((state) => {
      if (!optionOrder.includes(state)) optionOrder.push(state);
    });
    ce.taskStatus.innerHTML = optionOrder
      .map((state) => `<option value="${state}">${state}</option>`)
      .join('');
    ce.taskStatus.value = currentStatus;
    const disableForTransitions = allowed.length === 0;
    ce.taskStatus.disabled = disableForTransitions;
    ce.taskStatus.dataset.baseDisabled = disableForTransitions ? '1' : '0';
    enforcePolicyIssuedStatusIfNeeded();
  }

  // Helpers UI
  const fillSelect = (el, opts) => {
    if (!el || !Array.isArray(opts)) return;
    el.innerHTML = opts.map((o) => `<option>${o}</option>`).join('');
  };

  const MOTIVOS_PERDIDA_ANULACION = Array.from(
    new Set([...MOTIVOS_PERDIDA, ...MOTIVOS_ANULACION])
  );
  fillSelect(ce.renMotivo, MOTIVOS_PERDIDA_ANULACION);

  function captureCaseState() {
    if (!ce.root) return '';
    const data = {};
    ce.root.querySelectorAll('input, select, textarea').forEach((el) => {
      if (el.type === 'file') return;
      const key = el.id || el.name;
      if (!key) return;
      if (el.type === 'checkbox' || el.type === 'radio') data[key] = !!el.checked;
      else data[key] = el.value;
    });
    return JSON.stringify(data);
  }
  function refreshDirtyFlag() {
    if (ceBaseline === null) return;
    ceDirty = captureCaseState() !== ceBaseline;
  }
  function attachDirtyListeners() {
    if (!ce.root) return;
    ce.root.querySelectorAll('input, select, textarea').forEach((el) => {
      if (el.type === 'file') return;
      el.addEventListener('input', refreshDirtyFlag);
      el.addEventListener('change', refreshDirtyFlag);
    });
  }
  attachDirtyListeners();
  function hasCompletePolicyData() {
    const policyNumberFilled = !!(ce.policyNumber?.value || '').trim();
    const startFilled = !!(ce.renPreaviso?.value || '').trim();
    const endFilled = !!(ce.renDecision?.value || '').trim();
    const premiumFilled = parseFloat(ce.policyPremium?.value || '') > 0;
    const commissionPctFilled = parseFloat(ce.policyCommissionPct?.value || '') > 0;
    const commissionTotalFilled = parseFloat(ce.policyCommissionTotal?.value || '') > 0;
    const commissionFilled = commissionPctFilled || commissionTotalFilled;
    return policyNumberFilled && startFilled && endFilled && premiumFilled && commissionFilled;
  }
  function validarCamposPolizaCompletos(task) {
    const errores = [];
    const poliza = task?.case?.poliza || {};
    const renovacion = task?.case?.renovacion || {};
    const numeroPoliza = String(poliza.number || task.policyNumber || '').trim();
    const inicioVigencia = renovacion.fechaPreaviso || poliza.startDate || '';
    const terminoVigencia =
      poliza.endDate || task.policyEndDate || renovacion.fechaDecision || '';
    const primaTotal = poliza.premiumUF ?? task.policyPremiumUF;
    const comisionCorredor = poliza.commissionPct ?? task.policyCommissionPct;
    const primaNum = Number(primaTotal);
    const comisionNum = Number(comisionCorredor);

    if (!numeroPoliza) errores.push('N° de póliza');
    if (!inicioVigencia) errores.push('Inicio de vigencia');
    if (!terminoVigencia) errores.push('Término de vigencia');
    if (!Number.isFinite(primaNum) || primaNum <= 0) errores.push('Prima total (UF)');
    if (!Number.isFinite(comisionNum) || comisionNum <= 0) errores.push('% Comisión corredor');

    if (errores.length) {
      const lista = errores.join(', ');
      alert(
        `Para guardar en este estado debe completar los siguientes campos de la tarjeta Póliza: ${lista}.`
      );
      return false;
    }
    return true;
  }
  function enforcePolicyIssuedStatusIfNeeded() {
    if (!ce.taskStatus) return;
    const channelValue = normalizeChannel(ce.taskChannel?.value || ceCurrentTask?.canalIngreso || '');
    const baseStatus = ceCurrentTask ? normalizeTaskStatus(ceCurrentTask.status || '') : '';
    const isEligibleChannel =
      channelValue === CHANNEL_TRADICIONAL || channelValue === CHANNEL_CANAL_WEB;
    const isPropuesta = baseStatus === 'Propuesta enviada';
    const shouldForce =
      isEligibleChannel && isPropuesta && hasCompletePolicyData();
    const baseDisabled = ce.taskStatus.dataset?.baseDisabled === '1';
    const roleLocked = ce.taskStatus.dataset?.roleLock === '1';
    if (shouldForce) {
      if (ce.taskStatus.value !== 'Póliza emitida') ce.taskStatus.value = 'Póliza emitida';
      ce.taskStatus.disabled = true;
      ce.taskStatus.dataset.policyLock = '1';
    } else {
      ce.taskStatus.disabled = baseDisabled || roleLocked;
      ce.taskStatus.dataset.policyLock = '0';
    }
  }
  function applyEstadoFieldLockForRole(currentUser) {
    if (!ce.taskStatus) return;
    const roleName = normalizeRoleName(currentUser?.role || currentUser?.rol || '');
    const blockedRoles = ['Administrador', 'Supervisor', 'Agente'];
    const locked = blockedRoles.includes(roleName);
    ce.taskStatus.dataset.roleLock = locked ? '1' : '0';
    const baseDisabled = ce.taskStatus.dataset?.baseDisabled === '1';
    const policyLock = ce.taskStatus.dataset?.policyLock === '1';
    ce.taskStatus.disabled = locked || baseDisabled || policyLock;
  }
  function syncPolicyEndInputs() {
    if (!ce.renDecision || !ce.taskDue) return;
    ce.taskDue.value = ce.renDecision.value || '';
  }
  if (ce.renDecision) {
    ['input', 'change'].forEach((evt) => {
      ce.renDecision.addEventListener(evt, syncPolicyEndInputs);
    });
  }

  function setSearchSelectValue(hiddenId, searchInputId, value) {
    const hidden = document.getElementById(hiddenId);
    const search = document.getElementById(searchInputId);
    const val = value || '';
    if (hidden) hidden.value = val;
    if (search) search.value = val;
  }

  function setupRegionSearch(inputId, listId, hiddenId, onSelect) {
    const input = document.getElementById(inputId);
    const list = document.getElementById(listId);
    const hidden = document.getElementById(hiddenId);
    if (!input || !list || !hidden) return null;

    const closeList = () => {
      list.innerHTML = '';
      list.classList.remove('is-open');
    };

    const handleSelect = (region) => {
      hidden.value = region;
      hidden.dispatchEvent(new Event('input', { bubbles: true }));
      input.value = region;
      closeList();
      if (typeof onSelect === 'function') onSelect(region);
    };

    const renderOptions = (term = '') => {
      const regiones = buscarRegiones(term);
      list.innerHTML = '';
      if (!regiones.length) {
        list.innerHTML =
          '<div class="search-select-option is-empty">Sin resultados</div>';
        list.classList.add('is-open');
        return;
      }
      regiones.forEach((region) => {
        const option = document.createElement('div');
        option.className = 'search-select-option';
        option.textContent = region;
        option.addEventListener('mousedown', (e) => e.preventDefault());
        option.addEventListener('click', () => handleSelect(region));
        list.appendChild(option);
      });
      list.classList.add('is-open');
    };

    input.addEventListener('input', () => {
      hidden.value = '';
      hidden.dispatchEvent(new Event('input', { bubbles: true }));
      renderOptions(input.value);
    });
    input.addEventListener('focus', () => renderOptions(input.value));
    input.addEventListener('blur', () => setTimeout(closeList, 120));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeList();
    });

    if (hidden.value) input.value = hidden.value;

    return { input, hidden, renderOptions, closeList };
  }

  function setupComunaSearch(inputId, listId, hiddenId, getRegionFn, onSelect) {
    const input = document.getElementById(inputId);
    const list = document.getElementById(listId);
    const hidden = document.getElementById(hiddenId);
    if (!input || !list || !hidden) return null;

    const closeList = () => {
      list.innerHTML = '';
      list.classList.remove('is-open');
    };

    const handleSelect = (comuna) => {
      hidden.value = comuna;
      hidden.dispatchEvent(new Event('input', { bubbles: true }));
      input.value = comuna;
      closeList();
      if (typeof onSelect === 'function') onSelect(comuna);
    };

    const renderOptions = (term = '') => {
      const region =
        (typeof getRegionFn === 'function' && getRegionFn()) || '';
      list.innerHTML = '';
      if (!region) {
        list.innerHTML =
          '<div class="search-select-option is-empty">Selecciona una región</div>';
        list.classList.add('is-open');
        return;
      }
      const comunas = buscarComunas(region, term);
      if (!comunas.length) {
        list.innerHTML =
          '<div class="search-select-option is-empty">Sin resultados</div>';
        list.classList.add('is-open');
        return;
      }
      comunas.forEach((comuna) => {
        const option = document.createElement('div');
        option.className = 'search-select-option';
        option.textContent = comuna;
        option.addEventListener('mousedown', (e) => e.preventDefault());
        option.addEventListener('click', () => handleSelect(comuna));
        list.appendChild(option);
      });
      list.classList.add('is-open');
    };

    input.addEventListener('input', () => {
      hidden.value = '';
      hidden.dispatchEvent(new Event('input', { bubbles: true }));
      renderOptions(input.value);
    });
    input.addEventListener('focus', () => renderOptions(input.value));
    input.addEventListener('blur', () => setTimeout(closeList, 120));
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') closeList();
    });

    if (hidden.value) input.value = hidden.value;

    return { input, hidden, renderOptions, closeList };
  }

  function initRegionComunaSearchSelects() {
    const clienteComuna = setupComunaSearch(
      'clienteComunaSearch',
      'clienteComunaList',
      'ce_clientCommune',
      () => document.getElementById('ce_clientRegion')?.value || '',
      () => {
        if (ce.useRiskClient?.checked) syncRiskFromClient();
      }
    );
    setupRegionSearch(
      'clienteRegionSearch',
      'clienteRegionList',
      'ce_clientRegion',
      () => {
        setSearchSelectValue('ce_clientCommune', 'clienteComunaSearch', '');
        document
          .getElementById('ce_clientCommune')
          ?.dispatchEvent(new Event('input', { bubbles: true }));
        clienteComuna?.renderOptions('');
        if (ce.useRiskClient?.checked) syncRiskFromClient();
      }
    );

    const riesgoComuna = setupComunaSearch(
      'riesgoComunaSearch',
      'riesgoComunaList',
      'ce_riskCommune',
      () => document.getElementById('ce_riskRegion')?.value || '',
      null
    );
    setupRegionSearch(
      'riesgoRegionSearch',
      'riesgoRegionList',
      'ce_riskRegion',
      () => {
        setSearchSelectValue('ce_riskCommune', 'riesgoComunaSearch', '');
        document
          .getElementById('ce_riskCommune')
          ?.dispatchEvent(new Event('input', { bubbles: true }));
        riesgoComuna?.renderOptions('');
      }
    );
  }
  initRegionComunaSearchSelects();

  function syncRiskFromClient() {
    if (!ce.useRiskClient?.checked) return;
    if (ce.riskAddress && ce.clientAddress) ce.riskAddress.value = ce.clientAddress.value || '';
    setSearchSelectValue('ce_riskRegion', 'riesgoRegionSearch', ce.clientRegion?.value || '');
    setSearchSelectValue('ce_riskCommune', 'riesgoComunaSearch', ce.clientCommune?.value || '');
  }

  function applyRiskClientToggle() {
    const locked = !!ce.useRiskClient?.checked;
    [ce.riskAddress].forEach((input) => {
      if (!input) return;
      input.readOnly = locked;
      input.classList.toggle('readonly', locked);
    });
    const regionSearch = document.getElementById('riesgoRegionSearch');
    const comunaSearch = document.getElementById('riesgoComunaSearch');
    [regionSearch, comunaSearch].forEach((input) => {
      if (!input) return;
      input.readOnly = locked;
      input.classList.toggle('readonly', locked);
    });
    if (locked) syncRiskFromClient();
  }

function getAssigneeInputEmail() {
    if (!ce.taskAssignee) return '';
    const stored = (ce.taskAssignee.dataset?.email || '').trim();
    if (stored) return stored;
    return (ce.taskAssignee.value || '').trim();
  }

  function setAssigneeInputFromEmail(email) {
    if (!ce.taskAssignee) return;
    const trimmed = (email || '').trim();
    const user = resolveUserByEmail(trimmed);
    const display = user?.name || trimmed || '';
    ce.taskAssignee.value = display;
    ce.taskAssignee.dataset.email = trimmed;
    if (trimmed) ce.taskAssignee.setAttribute('title', trimmed);
    else ce.taskAssignee.removeAttribute('title');
  }

  function syncAssigneeInputMetadata({ convertDisplay = false } = {}) {
    if (!ce.taskAssignee) return;
    const raw = (ce.taskAssignee.value || '').trim();
    if (!raw) {
      ce.taskAssignee.dataset.email = '';
      ce.taskAssignee.removeAttribute('title');
      return;
    }
    const byEmail = resolveUserByEmail(raw);
    if (byEmail) {
      ce.taskAssignee.dataset.email = byEmail.email || '';
      ce.taskAssignee.setAttribute('title', byEmail.email || '');
      if (convertDisplay) ce.taskAssignee.value = byEmail.name || byEmail.email || '';
      return;
    }
    const byName = (state.users || []).find(
      (u) => (u.name || '').trim().toLowerCase() === raw.toLowerCase()
    );
    if (byName) {
      ce.taskAssignee.dataset.email = byName.email || '';
      ce.taskAssignee.setAttribute('title', byName.email || '');
      if (convertDisplay) ce.taskAssignee.value = byName.name || '';
      return;
    }
    ce.taskAssignee.dataset.email = raw;
    ce.taskAssignee.setAttribute('title', raw);
  }

  function calcAgentCommission(totalUF, assignee) {
    const pct = getUserCommission(assignee);
    if (!totalUF || !pct) return 0;
    return totalUF * (pct / 100);
  }

  function updateAgentCommission(totalUF) {
    if (!ce.policyAgentCommission) return;
    const total =
      typeof totalUF === 'number'
        ? totalUF
        : parseFloat(ce.policyCommissionTotal?.value) || 0;
    const assignee = getAssigneeInputEmail() || state.user?.email || '';
    const value = calcAgentCommission(total, assignee);
    ce.policyAgentCommission.value = value ? value.toFixed(2) : '';
  }

  function updateCommissionTotal() {
    if (!ce.policyPremium || !ce.policyCommissionPct || !ce.policyCommissionTotal) return 0;
    const premium = parseFloat(ce.policyPremium.value) || 0;
    const pct = parseFloat(ce.policyCommissionPct.value) || 0;
    const total = premium * (pct / 100);
    ce.policyCommissionTotal.value = total ? total.toFixed(2) : '';
    updateAgentCommission(total);
    enforcePolicyIssuedStatusIfNeeded();
    return total;
  }

  let progressTimer = null;
  function showProgress() {
    if (!progress) return;
    progress.classList.remove('hidden');
    progress.classList.add('is-visible');
    if (progressTimer) {
      clearTimeout(progressTimer);
      progressTimer = null;
    }
  }
  function hideProgress() {
    if (!progress) return;
    if (progressTimer) return;
    progressTimer = setTimeout(() => {
      progress.classList.remove('is-visible');
      setTimeout(() => progress.classList.add('hidden'), 200);
      progressTimer = null;
    }, 300);
  }

  if (ce.policyPremium) {
    ['input', 'change'].forEach((evt) =>
      ce.policyPremium.addEventListener(evt, updateCommissionTotal)
    );
  }
  if (ce.policyCommissionPct) {
    ['input', 'change'].forEach((evt) =>
      ce.policyCommissionPct.addEventListener(evt, updateCommissionTotal)
    );
  }
  const policyAutoFields = [
    ce.policyNumber,
    ce.renPreaviso,
    ce.renDecision,
    ce.policyPremium,
    ce.policyCommissionPct,
    ce.policyCommissionTotal,
  ];
  policyAutoFields.forEach((field) => {
    if (!field) return;
    ['input', 'change'].forEach((evt) =>
      field.addEventListener(evt, enforcePolicyIssuedStatusIfNeeded)
    );
  });
  const requiredParamModal = document.getElementById('requiredParamModal');
  const requiredParamMessage = document.getElementById('requiredParamMessage');
  const requiredParamClose = document.getElementById('requiredParamClose');
  const PERDIDA_MOTIVES = new Set([
    'precio',
    'cobertura',
    'servicio',
    'competidor',
    'documentación insuficiente',
    'documentacion insuficiente',
    'rechazo de suscripción',
    'rechazo de suscripcion',
    'y rechazo de suscripción',
    'y rechazo de suscripcion',
  ]);
  function setFieldError(el, show) {
    if (!el) return;
    el.classList.toggle('input-error', !!show);
  }
  function setInsurerErrorState(show) {
    setFieldError(ce.taskInsurer, show);
  }
  function showRequiredParamModal(message) {
    if (!requiredParamModal || !requiredParamMessage || !requiredParamClose) {
      window.alert(message);
      return;
    }
    requiredParamMessage.textContent = message;
    requiredParamModal.classList.remove('hidden');
    requiredParamModal.setAttribute('aria-hidden', 'false');
    requiredParamClose.focus();
  }
  if (requiredParamClose) {
    requiredParamClose.addEventListener('click', () => {
      requiredParamModal?.classList.add('hidden');
      requiredParamModal?.setAttribute('aria-hidden', 'true');
    });
  }
  if (ce.taskInsurer) {
    ce.taskInsurer.addEventListener('change', () => setInsurerErrorState(false));
  }
  [
    ce.firstName,
    ce.lastName,
    ce.rut,
    ce.email,
    ce.phone,
    ce.payFirstName,
    ce.payLastName,
    ce.payRut,
    ce.payEmail,
    ce.payPhone,
    ce.taskType,
    ce.policyPremium,
    ce.policyCommissionPct,
    ce.policyCommissionTotal,
    ce.policyIssueDate,
    ce.renDecision,
    ce.renPreaviso,
    ce.vehicleMake,
    ce.vehicleModel,
    ce.vehicleYear,
    ce.vehiclePlate,
    ce.vehicleType,
  ].forEach((field) => {
    if (!field) return;
    field.addEventListener('input', () => setFieldError(field, false));
    field.addEventListener('change', () => setFieldError(field, false));
  });
  if (ce.taskAssignee) {
    const handleAssigneeChange = (convertDisplay) => {
      syncAssigneeInputMetadata({ convertDisplay });
      const total = parseFloat(ce.policyCommissionTotal?.value) || 0;
      updateAgentCommission(total);
    };
    ce.taskAssignee.addEventListener('input', () => handleAssigneeChange(false));
    ['change', 'blur'].forEach((evt) =>
      ce.taskAssignee.addEventListener(evt, () => handleAssigneeChange(true))
    );
  }
  if (ce.useRiskClient) {
    ce.useRiskClient.addEventListener('change', applyRiskClientToggle);
  }
  ['clientAddress', 'clientCommune', 'clientRegion'].forEach((key) => {
    const el = ce[key];
    if (!el) return;
    el.addEventListener('input', () => {
      if (ce.useRiskClient?.checked) syncRiskFromClient();
    });
  });



  let toastTimer = null;
  function showInfoToast(message = 'Información actualizada', duration = 1000) {
    if (!infoToast) return;
    infoToast.textContent = message;
    infoToast.classList.remove('hidden');
    infoToast.classList.add('is-visible');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(() => {
      infoToast.classList.remove('is-visible');
      setTimeout(() => infoToast.classList.add('hidden'), 200);
      toastTimer = null;
    }, duration);
  }

  function ce_renderDocs() {
    if (!ce.filesList) return;
    ce.filesList.innerHTML =
      ce_docs
        .map(
          (d, i) =>
            `<div class="pill" data-i="${i}">${d.name} · ${Math.round(
              (d.size || 0) / 1024
            )} KB</div>`
        )
        .join('') || '<div class="muted">Sin documentos</div>';
  }
  function ce_renderComments() {
    if (!ce.comments) return;
    ce.comments.innerHTML =
      ce_comments
        .map((c) => {
          const when = new Date(c.ts).toLocaleString();
          return `<div class="card"><div class="muted" style="font-size:.85rem">${when} — ${
            c.by || 'Usuario'
          }</div><div>${c.text || ''}</div></div>`;
        })
        .join('') || '<div class="muted">Sin comentarios</div>';
  }
  function setTitleFromTask(t) {
    const ttl =
      (t && (t.title || t.titulo || t.Título || t.name || t.client)) ||
      ('Caso ' + (t?.id || ''));
    if (caseTitleEl && ttl) caseTitleEl.textContent = ttl;
  }

  function toggleTaskCreationLock(lock) {
    if (lock) setFormVisibility(false);
    const newBtn = document.getElementById('btnNew');
    if (newBtn) {
      newBtn.disabled = lock;
      if (!lock) newBtn.blur();
    }
  }

  window.openEditorFor = function openEditorFor(t) {
    if (!ce.root || !t) return;
    currentId = t.id;
    ceCurrentTask = t;
    state.ui = state.ui || {};
    state.ui.isCaseDetailOpen = true;
    updateFiltersCardVisibility();
    const chipFirst = document.getElementById('chipFirstName');
    const chipLast = document.getElementById('chipLastName');
    const chipStatus = document.getElementById('chipStatus');
    const chipTags = document.getElementById('chipTags');

    // Cargar sub-secciones
    // Normaliza para preservar renovacion/pagos aunque no haya UI
    const c = normalizeCase(t.case || {});
    t.case = c;
    const p = c.personal || {};
    const g = c.pago || {};
    const poliza = c.poliza || {};
    const veh = c.vehiculo || {};
    const riesgo = c.riesgo || {};
    const ren = ensureRenovacion(c.renovacion) ||
      crearRenovacion({ estadoRenovacion: ESTADOS_RENOVACION[0] || '' });
    const policyEndValue = t.policyEndDate || ren.fechaDecision || '';
    // Personales
    ce.firstName.value = p.firstName || '';
    ce.lastName.value = p.lastName || '';
    ce.email.value = p.email || '';
    ce.rut.value = p.rut || '';
    ce.phone.value = p.phone || '';
    if (ce.clientAddress) ce.clientAddress.value = p.address || '';
    if (ce.clientCommune) ce.clientCommune.value = p.commune || '';
    if (ce.clientRegion) ce.clientRegion.value = p.region || '';
    setSearchSelectValue('ce_clientRegion', 'clienteRegionSearch', ce.clientRegion?.value || '');
    setSearchSelectValue('ce_clientCommune', 'clienteComunaSearch', ce.clientCommune?.value || '');
    if (chipFirst) chipFirst.textContent = p.firstName || '-';
    if (chipLast) chipLast.textContent = p.lastName || '-';
    if (chipStatus) chipStatus.textContent = normalizeTaskStatus(t.status || '-');
    if (chipTags) {
      const tagValue = (t.tags && t.tags[0]) || '';
      const showTag = tagValue && tagValue.toLowerCase() !== 'sin clasificación' && tagValue.toLowerCase() !== 'sin clasificar';
      chipTags.textContent = showTag ? tagValue : '';
      chipTags.style.display = showTag ? 'inline-flex' : 'none';
    }
    if (ce.customerType) ce.customerType.value = t.customerType || 'natural';

    // Pago / Asegurado
    const insuredFirst = g.firstName || g.nombre || '';
    const insuredLast = g.lastName || '';
    const payTypeValue =
      g.tipo === 'empresa' ? 'juridica' : g.tipo || 'natural';
    ce.payType.value = payTypeValue;
    ce.usePersonal.checked = !!g.usarPersonales;
    if (ce.payFirstName) ce.payFirstName.value = insuredFirst || '';
    if (ce.payLastName) ce.payLastName.value = insuredLast || '';
    ce.payRut.value = g.rut || '';
    ce.payEmail.value = g.email || '';
    ce.payPhone.value = g.phone || '';

    // Póliza
    if (ce.policyNumber) ce.policyNumber.value = poliza.number || '';
    if (ce.policyIssueDate) ce.policyIssueDate.value = poliza.issueDate || poliza.fechaEmisionPoliza || t.fechaEmisionPoliza || '';
    if (ce.policyPremium) ce.policyPremium.value = poliza.premiumUF ?? '';
    if (ce.policyCommissionPct) ce.policyCommissionPct.value = poliza.commissionPct ?? '';
    if (ce.policyCommissionTotal)
      ce.policyCommissionTotal.value =
        poliza.commissionUF !== undefined && poliza.commissionUF !== null
          ? String(poliza.commissionUF)
          : '';
    if (ce.policyAgentCommission) ce.policyAgentCommission.value = poliza.agentCommission || '';
    const totalInitial = updateCommissionTotal();
    if (!totalInitial && ce.policyAgentCommission)
      ce.policyAgentCommission.value = poliza.agentCommission || '';

    // Vehículo
    if (ce.vehicleMake) ce.vehicleMake.value = veh.make || '';
    if (ce.vehicleModel) ce.vehicleModel.value = veh.model || '';
    if (ce.vehicleYear) ce.vehicleYear.value = veh.year || '';
    if (ce.vehiclePlate) ce.vehiclePlate.value = veh.plate || '';
    if (ce.vehicleColor) ce.vehicleColor.value = veh.color || '';
    if (ce.vehicleChassis) ce.vehicleChassis.value = veh.chassis || '';
    if (ce.vehicleFuel) ce.vehicleFuel.value = veh.fuel || '';
    if (ce.vehicleType) ce.vehicleType.value = veh.type || '';
    if (ce.vehicleFleet) ce.vehicleFleet.checked = !!veh.fleet;

    // Dirección del riesgo
    if (ce.riskAddress) ce.riskAddress.value = riesgo.address || '';
    if (ce.riskCommune) ce.riskCommune.value = riesgo.commune || '';
    if (ce.riskRegion) ce.riskRegion.value = riesgo.region || '';
    setSearchSelectValue('ce_riskRegion', 'riesgoRegionSearch', ce.riskRegion?.value || '');
    setSearchSelectValue('ce_riskCommune', 'riesgoComunaSearch', ce.riskCommune?.value || '');
    if (ce.riskPropertyType) ce.riskPropertyType.value = riesgo.propertyType || '';
    if (ce.useRiskClient) {
      ce.useRiskClient.checked = false;
      applyRiskClientToggle();
    }

    // Archivos y comentarios
    ce_docs = (c.docs || []).slice();
    ce_comments = (c.comments || []).slice();
    ce_renderDocs();
    ce_renderComments();

    // Datos de la tarea
    if (ce.taskTitle) ce.taskTitle.value = t.title || '';
    if (ce.taskType) ce.taskType.value = t.type || 'Vehículo';
    if (ce.taskPriority) ce.taskPriority.value = t.priority || 'Media';
    if (ce.taskChannel) ce.taskChannel.value = t.canalIngreso || '';
    renderStatusOptions(t);
    if (ce.taskStatus && !ce.taskStatus.innerHTML) ce.taskStatus.value = t.status || 'Abierta';
    if (ce.taskDue) ce.taskDue.value = policyEndValue || '';
    syncPolicyEndInputs();
    if (ce.taskAssignee) setAssigneeInputFromEmail(t.assignee || state.user?.email || '');
    if (ce.taskPolicy) ce.taskPolicy.value = t.policy || '';
    if (ce.taskInsurer) {
      ce.taskInsurer.value = t.insurer || 'Sin Asignar';
      syncCompanySearchInput(ce.taskInsurer);
    }
    if (ce.taskTags) ce.taskTags.value = t.tags?.[0] || '';

    // Renovación
    if (ce.renPreaviso) ce.renPreaviso.value = ren.fechaPreaviso || '';
    if (ce.renDecision) ce.renDecision.value = policyEndValue;
    if (ce.renMotivo) ce.renMotivo.value = ren.motivoPerdidaOAnulacion || '';
    if (ce.renComentario) ce.renComentario.value = ren.comentario || '';

    setTitleFromTask(t);
    applyCommissionFieldsVisibilityForRole(getCurrentUserIdentity());
    applyEstadoFieldLockForRole(getCurrentUserIdentity());

    // Entrar a modo edición (oculta tabla/paginador via CSS)
    document.body.classList.add('editing');
    if (typeof window.__syncHeaderCheckbox === 'function') window.__syncHeaderCheckbox();
    toggleTaskCreationLock(true);
    ce.root.classList.remove('hidden');
    ce.root.scrollIntoView({ behavior: 'smooth', block: 'start' });
    requestAnimationFrame(() => {
      setupCaseCarouselHeightSync();
    });
    enforcePolicyIssuedStatusIfNeeded();
    ceBaseline = captureCaseState();
    ceDirty = false;
  };

  function closeEditor() {
    document.body.classList.remove('editing');
    if (typeof window.__syncHeaderCheckbox === 'function') window.__syncHeaderCheckbox();
    toggleTaskCreationLock(false);
    if (ce.root) ce.root.classList.add('hidden');
    currentId = null;
    ceCurrentTask = null;
    ceBaseline = null;
    ceDirty = false;
    state.ui = state.ui || {};
    state.ui.isCaseDetailOpen = false;
    updateFiltersCardVisibility();
    const chipFirst = document.getElementById('chipFirstName');
    const chipLast = document.getElementById('chipLastName');
    const chipStatus = document.getElementById('chipStatus');
    const chipTags = document.getElementById('chipTags');
    if (chipFirst) chipFirst.textContent = '-';
    if (chipLast) chipLast.textContent = '-';
    if (chipStatus) chipStatus.textContent = '-';
    if (chipTags) {
      chipTags.textContent = '';
      chipTags.style.display = 'none';
    }
  }

  // Usar datos personales -> pago
  if (ce.usePersonal) {
    ce.usePersonal.addEventListener('change', () => {
      if (ce.usePersonal.checked) {
        if (ce.payFirstName) ce.payFirstName.value = ce.firstName.value || '';
        if (ce.payLastName) ce.payLastName.value = ce.lastName.value || '';
        ce.payRut.value = ce.rut.value || '';
        ce.payEmail.value = ce.email.value || '';
        ce.payPhone.value = ce.phone.value || '';
        if (ce.payType && ce.customerType)
          ce.payType.value = ce.customerType.value || 'natural';
      }
    });
  }

  if (ce.taskChannel) {
    ce.taskChannel.addEventListener('change', () => {
      if (ceCurrentTask) {
        renderStatusOptions(ceCurrentTask);
        enforcePolicyIssuedStatusIfNeeded();
      }
    });
  }

  // Carga de archivos (solo metadatos en demo)
  if (ce.files) {
    ce.files.addEventListener('change', () => {
      const allowed = [
        'application/pdf',
        'application/msword',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'image/png',
        'image/jpeg',
      ];
      Array.from(ce.files.files || []).forEach((f) => {
        if (
          allowed.includes(f.type) ||
          /\.pdf$|\.docx?$|\.png$|\.jpe?g$/i.test(f.name)
        ) {
          ce_docs.push({ name: f.name, size: f.size, type: f.type, ts: Date.now() });
        }
      });
      ce.files.value = '';
      ce_renderDocs();
    });
  }

  // Comentarios
  if (ce.btnAddComment) {
    ce.btnAddComment.addEventListener('click', () => {
      const text = (ce.comment.value || '').trim();
      if (!text) return;
      const who = (state.user && (state.user.name || state.user.email)) || 'Usuario';
      ce_comments.unshift({ by: who, text, ts: Date.now() });
      ce.comment.value = '';
      ce_renderComments();
    });
    if (ce.comment) {
      ce.comment.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          ce.btnAddComment.click();
        }
      });
    }
  }

  // Guardar (NO cierra el editor)
  if (ce.btnSave) {
    ce.btnSave.addEventListener('click', () => {
      if (!currentId) return;
      const t = (state.tasks || []).find((x) => String(x.id) === String(currentId));
      if (!t) return;
      const nowIso = new Date().toISOString();
      const prevAssignmentSnapshot = getAssignmentSnapshot(t);
      const requiresInsurer = hasCompletePolicyData();
      const insurerValue = (ce.taskInsurer?.value || '').trim();
      const clientRutValue = (ce.rut?.value || '').trim();
      const clientPhoneValue = (ce.phone?.value || '').trim();
      const clientPhoneDigits = clientPhoneValue.replace(/\D/g, '');
      const insuredRutValue = (ce.payRut?.value || '').trim();
      const insuredPhoneValue = (ce.payPhone?.value || '').trim();
      const insuredPhoneDigits = insuredPhoneValue.replace(/\D/g, '');
      const vehicleYearValue = ce.vehicleYear?.value || '';
      const vehiclePlateValue = (ce.vehiclePlate?.value || '').trim();
      const requiredClientFields = [
        { el: ce.firstName, value: (ce.firstName?.value || '').trim(), message: 'Debes ingresar los nombres del cliente.' },
        { el: ce.lastName, value: (ce.lastName?.value || '').trim(), message: 'Debes ingresar los apellidos del cliente.' },
        { el: ce.rut, value: clientRutValue, message: 'Debes ingresar el RUT del cliente.' },
        { el: ce.email, value: (ce.email?.value || '').trim(), message: 'Debes ingresar el correo del cliente.' },
        { el: ce.phone, value: clientPhoneValue, message: 'Debes ingresar el teléfono del cliente.' },
      ];
      for (const field of requiredClientFields) {
        if (!field.value) {
          setFieldError(field.el, true);
          showRequiredParamModal(field.message);
          return;
        }
        setFieldError(field.el, false);
      }
      if (!esRutValido(clientRutValue)) {
        setFieldError(ce.rut, true);
        showRequiredParamModal('RUT de cliente inválido. Formato esperado: 12345678-9');
        return;
      }
      setFieldError(ce.rut, false);
      if (!esTelefonoValido(clientPhoneValue)) {
        setFieldError(ce.phone, true);
        showRequiredParamModal('Teléfono de cliente inválido. Debe contener exactamente 9 dígitos numéricos.');
        return;
      }
      setFieldError(ce.phone, false);
      if (requiresInsurer && (!insurerValue || insurerValue.toLowerCase() === 'sin asignar')) {
        setInsurerErrorState(true);
        showRequiredParamModal('Debe asignar una compañía antes de guardar esta póliza.');
        return;
      }
      setInsurerErrorState(false);
      const policyNumberValue = (ce.policyNumber?.value || '').trim();
      const hasPolicyNumber = !!policyNumberValue;
      const ramoValue = (ce.taskType?.value || '').trim();
      const policyPremiumValue = (ce.policyPremium?.value || '').trim();
      const policyCommissionValue = (ce.policyCommissionPct?.value || '').trim();
      const policyCommissionTotalValue = (ce.policyCommissionTotal?.value || '').trim();
      const policyStart = ce.renPreaviso?.value || '';
      const policyEnd = ce.renDecision?.value || '';
      const ramoLower = ramoValue.toLowerCase();
      const requiresVehicleData = ramoLower === 'vehículo' || ramoLower === 'vehiculo' || ramoLower === 'soap';
      if (requiresVehicleData) {
        const vehicleFields = [
          { el: ce.vehicleMake, message: 'Debes ingresar la marca del vehículo para continuar.' },
          { el: ce.vehicleModel, message: 'Debes ingresar el modelo del vehículo para continuar.' },
          { el: ce.vehicleYear, message: 'Debes ingresar el año del vehículo para continuar.' },
          { el: ce.vehiclePlate, message: 'Debes ingresar la patente del vehículo para continuar.' },
          { el: ce.vehicleType, message: 'Debes seleccionar el tipo de vehículo para continuar.' },
        ];
        for (const { el, message } of vehicleFields) {
          const value = (el?.value || '').trim();
          if (!value) {
            setFieldError(el, true);
            showRequiredParamModal(message);
            return;
          }
          setFieldError(el, false);
        }
      } else {
        [ce.vehicleMake, ce.vehicleModel, ce.vehicleYear, ce.vehiclePlate, ce.vehicleType].forEach((field) =>
          setFieldError(field, false)
        );
      }
      if ((vehicleYearValue || requiresVehicleData) && !esAnioVehiculoValido(vehicleYearValue)) {
        setFieldError(ce.vehicleYear, true);
        showRequiredParamModal('Año del vehículo inválido. Debe ser un número de 4 dígitos mayor o igual a 2000.');
        return;
      }
      setFieldError(ce.vehicleYear, false);
      if ((vehiclePlateValue || requiresVehicleData) && !esPatenteValida(vehiclePlateValue)) {
        setFieldError(ce.vehiclePlate, true);
        showRequiredParamModal('Patente inválida. Debe contener exactamente 6 caracteres alfanuméricos (letras y números, sin espacios ni símbolos).');
        return;
      }
      setFieldError(ce.vehiclePlate, false);
      if (hasPolicyNumber) {
        setFieldError(ce.taskType, false);
        setFieldError(ce.policyPremium, false);
        setFieldError(ce.policyCommissionPct, false);
        setFieldError(ce.policyCommissionTotal, false);
        if (!ramoValue) {
          setFieldError(ce.taskType, true);
          showRequiredParamModal('Debes seleccionar un ramo para continuar.');
          return;
        }
        if (!policyPremiumValue) {
          setFieldError(ce.policyPremium, true);
          showRequiredParamModal('Debes ingresar la prima total en UF para continuar.');
          return;
        }
        if (!policyCommissionValue) {
          setFieldError(ce.policyCommissionPct, true);
          showRequiredParamModal('Debes ingresar la comisión del corredor para continuar.');
          return;
        }
        setFieldError(ce.policyCommissionTotal, false);
      }
      if (!hasPolicyNumber) {
        setFieldError(ce.taskType, false);
        setFieldError(ce.policyPremium, false);
        setFieldError(ce.policyCommissionPct, false);
        setFieldError(ce.policyCommissionTotal, false);
      }
      if (policyStart) {
        if (!policyEnd) {
          setFieldError(ce.renDecision, true);
          showRequiredParamModal('Debes ingresar el término de vigencia para continuar.');
          return;
        }
        if (new Date(policyEnd) < new Date(policyStart)) {
          setFieldError(ce.renDecision, true);
          showRequiredParamModal('El término de vigencia no puede ser anterior al inicio.');
          return;
        }
      }
      setFieldError(ce.renDecision, false);
      if (hasPolicyNumber && policyStart) {
        const issueDateValue = ce.policyIssueDate?.value || '';
        if (!issueDateValue) {
          setFieldError(ce.policyIssueDate, true);
          showRequiredParamModal('La Fecha de emisión es obligatoria cuando existe N° de póliza e inicio de vigencia.');
          return;
        }
        setFieldError(ce.policyIssueDate, false);
      }
      const requiresInsuredData =
        hasPolicyNumber &&
        !!policyStart &&
        !!policyEnd &&
        !!policyPremiumValue &&
        !!policyCommissionValue;
      if (requiresInsuredData) {
        const insuredFields = [
          { el: ce.payFirstName, message: 'Debes ingresar los nombres del asegurado.' },
          { el: ce.payLastName, message: 'Debes ingresar los apellidos del asegurado.' },
          { el: ce.payRut, message: 'Debes ingresar el RUT del asegurado.' },
          { el: ce.payEmail, message: 'Debes ingresar el correo del asegurado.' },
          { el: ce.payPhone, message: 'Debes ingresar el teléfono del asegurado.' },
        ];
        for (const { el, message } of insuredFields) {
          const value = (el?.value || '').trim();
          if (!value) {
            setFieldError(el, true);
            showRequiredParamModal(message);
            return;
          }
          setFieldError(el, false);
        }
      } else {
        [ce.payFirstName, ce.payLastName, ce.payRut, ce.payEmail, ce.payPhone].forEach((field) =>
          setFieldError(field, false)
        );
      }
      if ((insuredRutValue || requiresInsuredData) && !esRutValido(insuredRutValue)) {
        setFieldError(ce.payRut, true);
        showRequiredParamModal('RUT del asegurado inválido. Formato esperado: 12345678-9');
        return;
      }
      setFieldError(ce.payRut, false);
      if ((insuredPhoneValue || requiresInsuredData) && !esTelefonoValido(insuredPhoneValue)) {
        setFieldError(ce.payPhone, true);
        showRequiredParamModal('Teléfono del asegurado inválido. Debe contener exactamente 9 dígitos numéricos.');
        return;
      }
      setFieldError(ce.payPhone, false);

      // Guardar secciones del caso
      const premiumUF = parseFloat(ce.policyPremium?.value) || 0;
      const commissionPct = parseFloat(ce.policyCommissionPct?.value) || 0;
      const commissionUF = parseFloat(ce.policyCommissionTotal?.value) || 0;
      const normalizedClientRut = clientRutValue.trim().toUpperCase();
      const normalizedInsuredRut = insuredRutValue ? insuredRutValue.trim().toUpperCase() : '';
      const normalizedClientPhone = clientPhoneDigits.slice(0, 9);
      const normalizedInsuredPhone = insuredPhoneDigits.slice(0, 9);
      const normalizedVehicleYear = vehicleYearValue
        ? vehicleYearValue.replace(/\D/g, '').slice(0, 4)
        : '';
      const normalizedVehiclePlate = vehiclePlateValue ? vehiclePlateValue.trim().toUpperCase() : '';
      if (ce.rut) ce.rut.value = normalizedClientRut;
      if (ce.payRut) ce.payRut.value = normalizedInsuredRut;
      if (ce.phone) ce.phone.value = normalizedClientPhone;
      if (ce.payPhone) ce.payPhone.value = normalizedInsuredPhone;
      if (ce.vehicleYear) ce.vehicleYear.value = normalizedVehicleYear;
      if (ce.vehiclePlate) ce.vehiclePlate.value = normalizedVehiclePlate;

      const assigneeEmail = getAssigneeInputEmail() || state.user?.email || '';
      const agentCommissionValue = calcAgentCommission(commissionUF, assigneeEmail);

      const prevCase = normalizeCase(t.case || {});
      const prevRenovacion =
        ensureRenovacion(prevCase.renovacion) ||
        crearRenovacion({ estadoRenovacion: ESTADOS_RENOVACION[0] || '' });
      const prevPagos = Array.isArray(prevCase.pagos) ? prevCase.pagos : [];
      const policyEndInputValue = ce.renDecision?.value || '';
      const policyEndValue =
        policyEndInputValue ||
        prevCase.poliza?.endDate ||
        prevRenovacion.fechaDecision ||
        '';

      const estadoRen = prevRenovacion.estadoRenovacion || '';
      const motivoRen = (ce.renMotivo?.value || '').trim();
      const comentarioRen = (ce.renComentario?.value || '').trim();
      const needsComment = !!motivoRen;
      const commentTooShort = needsComment && comentarioRen.length < 10;
      if (commentTooShort) {
        showRequiredParamModal('Debes ingresar un comentario para continuar.');
        setFieldError(ce.renComentario, true);
        return;
      }
      setFieldError(ce.renComentario, false);
      const renovacion = crearRenovacion({
        ...prevRenovacion,
        estadoRenovacion: estadoRen,
        fechaPreaviso: ce.renPreaviso?.value || prevRenovacion.fechaPreaviso || '',
        fechaDecision: policyEndValue,
        motivoPerdidaOAnulacion: motivoRen,
        comentario:
          comentarioRen !== ''
            ? comentarioRen
            : prevRenovacion.comentario || '',
      });

      const pagos = prevPagos.slice(0, 200);

      const oldStatus = t.status || 'Sin estado';
      let newStatus = ce.taskStatus ? normalizeTaskStatus(ce.taskStatus.value) : oldStatus;
      const channelValue = normalizeChannel(ce.taskChannel?.value || t.canalIngreso || '');
      const allowedTransitions = getAllowedNextStates({ ...t, canalIngreso: channelValue });
      const oldSubStatus = t.subStatus || '';
      const newSubStatus = ce.taskSubStatus ? ce.taskSubStatus.value : oldSubStatus;

      t.case = {
        ...prevCase,
        personal: {
          firstName: ce.firstName.value.trim(),
          lastName: ce.lastName.value.trim(),
          email: ce.email.value.trim(),
          rut: normalizedClientRut,
          phone: normalizedClientPhone,
          address: (ce.clientAddress?.value || '').trim(),
          commune: (ce.clientCommune?.value || '').trim(),
          region: (ce.clientRegion?.value || '').trim(),
        },
        pago: {
          tipo: ce.payType.value,
          usarPersonales: ce.usePersonal.checked,
          firstName: (ce.payFirstName?.value || '').trim(),
          lastName: (ce.payLastName?.value || '').trim(),
          nombre: [ce.payFirstName?.value || '', ce.payLastName?.value || '']
            .join(' ')
            .trim(),
          rut: normalizedInsuredRut,
          email: ce.payEmail.value.trim(),
          phone: normalizedInsuredPhone,
        },
        poliza: {
          insurer: ce.taskInsurer?.value || 'Sin Asignar',
          number: (ce.policyNumber?.value || '').trim(),
          premiumUF,
          commissionPct,
          commissionUF,
          agentCommission: agentCommissionValue,
          endDate: policyEndValue,
          issueDate: ce.policyIssueDate?.value || '',
        },
        vehiculo: {
          make: (ce.vehicleMake?.value || '').trim(),
          model: (ce.vehicleModel?.value || '').trim(),
          year: normalizedVehicleYear,
          plate: normalizedVehiclePlate,
          color: (ce.vehicleColor?.value || '').trim(),
          chassis: (ce.vehicleChassis?.value || '').trim(),
          fuel: ce.vehicleFuel?.value || '',
          type: ce.vehicleType?.value || '',
          fleet: !!ce.vehicleFleet?.checked,
        },
        riesgo: {
          address: (ce.riskAddress?.value || '').trim(),
          commune: (ce.riskCommune?.value || '').trim(),
          region: (ce.riskRegion?.value || '').trim(),
          propertyType: ce.riskPropertyType?.value || '',
        },
        docs: ce_docs.slice(0, 200),
        comments: ce_comments.slice(0, 500),
        renovacion,
        pagos,
      };
      t.policyNumber = t.case.poliza?.number || '';
      t.fechaEmisionPoliza = t.case.poliza?.issueDate || '';
      t.policyPremiumUF = premiumUF;
      t.policyCommissionTotalUF = commissionUF;
      console.log(
        '[SAVE DEAL] policyPremiumUF:',
        t.policyPremiumUF,
        'policyCommissionTotalUF:',
        t.policyCommissionTotalUF
      );

      // Guardar datos de la tarea (desde interior del caso)
      if (ce.taskPriority) t.priority = ce.taskPriority.value;
      t.motivoPerdida = motivoRen;
      t.canalIngreso = channelValue;

      const persistTask = () => {
        showProgress();
        Promise.resolve()
          .then(() => {
            saveTasks();
            applyFilters();
            ceBaseline = captureCaseState();
            ceDirty = false;
            if (DEBUG_RENOVACION_DATOS)
              console.log('Caso guardado (renovacion/pagos preservados):', t.case);
            showSuccessToast();
          })
          .finally(() => hideProgress());
      };

      const applyCommonFields = (targetStatus) => {
        t.fueEditado = true;
        const normalizedTarget = normalizeTaskStatus(targetStatus || oldStatus);
        if (normalizedTarget !== oldStatus) {
          pushBitacoraEntry(t, `Estado cambiado de "${oldStatus}" a "${normalizedTarget}"`);
        }
        if (newSubStatus !== oldSubStatus) {
          const fromLabel = oldSubStatus || 'Sin subestado';
          const toLabel = newSubStatus || 'Sin subestado';
          pushBitacoraEntry(t, `Subestado cambiado de "${fromLabel}" a "${toLabel}"`);
        }
        t.status = normalizedTarget;
        t.subStatus = newSubStatus;
        t.updatedAt = nowIso;
        t.fechaActualizacion = nowIso;
        t.policyEndDate = policyEndValue || null;
        t.due = t.policyEndDate;
        if (ce.taskStatus) ce.taskStatus.value = normalizedTarget;
        renderStatusOptions(t);
        const currentIdentity = getCurrentUserIdentity();
        const currentRole = roleKey(currentIdentity.role);
        if (currentRole === ROLE_KEY_AGENTE) {
          t.assignee = currentIdentity.email || '';
        } else if (ce.taskAssignee) {
          t.assignee = assigneeEmail.trim();
        }
        applyAssigneeMetadata(t, t.assignee);
        logAssignmentChange(t, prevAssignmentSnapshot);
        if (ce.taskPolicy) t.policy = (ce.taskPolicy.value || '').trim();
        if (ce.taskInsurer) t.insurer = ce.taskInsurer.value;
        if (ce.taskTags) {
          const val = (ce.taskTags.value || '').trim();
          t.tags = val ? [val] : [];
        }
        if (ce.customerType) t.customerType = ce.customerType.value || 'natural';
        t.email = t.case.personal.email;
        t.phone = t.case.personal.phone;

        // Mantener 'client' si hay nombres
        const fullName = [t.case.personal.firstName, t.case.personal.lastName]
          .filter(Boolean)
          .join(' ')
          .trim();
        if (fullName) t.client = fullName;
      };

      const allowedStatesSnapshot = Array.isArray(allowedTransitions) ? allowedTransitions.slice() : [];
      openSaveOptionsModal(
        { ...t, motivoPerdida: motivoRen },
        allowedStatesSnapshot,
        (choice) => {
          if (!choice) return;
          if (choice === SAVE_OPTION_SOLO_GUARDAR) {
            applyCommonFields(oldStatus);
            persistTask();
            return;
          }
          const targetStatus = normalizeTaskStatus(choice);
          if (!allowedTransitions.includes(targetStatus)) {
            alert('El estado seleccionado no es válido para este negocio.');
            return;
          }
          if (!isStatusAllowedForChannel(targetStatus, channelValue)) {
            alert('El estado seleccionado no es válido para el canal de ingreso actual.');
            return;
          }
          const estadosQueRequierenPolizaCompleta = [
            'Póliza emitida',
            'Renovada',
            'Renovada con modificación',
          ];
          if (estadosQueRequierenPolizaCompleta.includes(targetStatus)) {
            const ok = validarCamposPolizaCompletos(t);
            if (!ok) return;
          }
          if ((targetStatus === 'Pérdida' || targetStatus === 'Desistida') && !motivoRen) {
            const estadoLabel = targetStatus === 'Desistida' ? '"Desistida"' : '"Pérdida"';
            alert(`Debes seleccionar un Motivo de pérdida cuando el estado es ${estadoLabel}.`);
            return;
          }
          applyCommonFields(targetStatus);
          persistTask();
          closeEditor();
        }
      );
      // NO cerrar editor
    });
  }
  const headerSaveBtn = document.getElementById('btnEditorSave');
  if (headerSaveBtn)
    headerSaveBtn.addEventListener('click', () => {
      if (!document.body.classList.contains('editing')) return;
      ce.btnSave?.click();
    });

  if (ce.btnExit)
    ce.btnExit.addEventListener('click', () => {
      if (ceDirty && !confirm('Seguro que desea salir sin guardar?')) return;
      closeEditor();
    });
  const headerExitBtn = document.getElementById('btnEditorExit');
  if (headerExitBtn)
    headerExitBtn.addEventListener('click', () => {
      if (!document.body.classList.contains('editing')) return;
      ce.btnExit?.click();
    });

  // Delegación: abrir editor desde botón Editar
  if (tbody) {
    tbody.addEventListener('click', (ev) => {
      const btn = ev.target.closest('.btnEdit');
      if (!btn) return;
      const tr = btn.closest('tr');
      if (!tr) return;
      const id = tr.getAttribute('data-id');
      if (!id) return;
      const t = (state.tasks || []).find((x) => String(x.id) === String(id));
      if (!t) return;
      openEditorFor(t);
    });
  }
})();

/* ===========================
   Paginador robusto
   Requiere: #perPage, #btnPrev, #btnNext, #pageInfo
   CSS: .is-hidden-pager { display:none!important }
=========================== */
(function () {
  const STORAGE_KEY = 'tasks_page_size';

  const PAGER = (window.__TASKS_PAGER__ = {
    page: 1,
    pageSize: parseInt(localStorage.getItem(STORAGE_KEY) || '8', 10),

    get tbody() {
      return document.getElementById('tbody');
    },
    get rowsAll() {
      return this.tbody ? Array.from(this.tbody.querySelectorAll('tr')) : [];
    },
    get rows() {
      // Solo filas de tarea reales
      let core = this.rowsAll.filter(
        (tr) =>
          tr.hasAttribute('data-id') ||
          tr.hasAttribute('data-task') ||
          tr.querySelector('.btnEdit')
      );
      return core.length ? core : this.rowsAll;
    },
    get total() {
      return this.rows.length;
    },
    get totalPages() {
      return Math.max(1, Math.ceil(this.total / this.pageSize));
    },
  });

  const $ = (s) => document.querySelector(s);
  const els = {
    perPage: $('#perPage'),
    btnPrev: $('#btnPrev'),
    btnNext: $('#btnNext'),
    pageInfo: $('#pageInfo'),
  };

  function clampPage() {
    if (PAGER.page < 1) PAGER.page = 1;
    if (PAGER.page > PAGER.totalPages) PAGER.page = PAGER.totalPages;
  }

  function applyVisibility() {
    const start = (PAGER.page - 1) * PAGER.pageSize;
    const end = start + PAGER.pageSize;

    if (PAGER.rows.length === 0) {
      PAGER.rowsAll.forEach((tr) => tr.classList.remove('is-hidden-pager'));
      return;
    }
    PAGER.rowsAll.forEach((tr) => tr.classList.add('is-hidden-pager'));
    PAGER.rows.forEach((tr, idx) => {
      if (idx >= start && idx < end) tr.classList.remove('is-hidden-pager');
    });
  }

  function render() {
    if (!PAGER.tbody) return;
    clampPage();
    applyVisibility();
    if (els.pageInfo) els.pageInfo.textContent = `${PAGER.page}/${PAGER.totalPages}`;
    if (els.btnPrev) els.btnPrev.disabled = PAGER.page <= 1;
    if (els.btnNext) els.btnNext.disabled = PAGER.page >= PAGER.totalPages;
    if (els.perPage) els.perPage.value = String(PAGER.pageSize);

    // Sincroniza checkbox/acciones tras paginar
    if (typeof window.__syncHeaderCheckbox === 'function') window.__syncHeaderCheckbox();
  }

  if (els.perPage) {
    els.perPage.addEventListener('change', (e) => {
      PAGER.pageSize = parseInt(e.target.value, 10) || 8;
      localStorage.setItem(STORAGE_KEY, String(PAGER.pageSize));
      PAGER.page = 1;
      render();
    });
  }
  if (els.btnPrev) {
    els.btnPrev.addEventListener('click', () => {
      if (PAGER.page > 1) {
        PAGER.page--;
        render();
      }
    });
  }
  if (els.btnNext) {
    els.btnNext.addEventListener('click', () => {
      if (PAGER.page < PAGER.totalPages) {
        PAGER.page++;
        render();
      }
    });
  }

  // Re-render cuando cambie el contenido del tbody (filtros, import, etc.)
  let mo = null;
  function attachObserver() {
    if (mo) {
      mo.disconnect();
      mo = null;
    }
    if (!PAGER.tbody) return;
    mo = new MutationObserver(() => render());
    mo.observe(PAGER.tbody, { childList: true });
    window.refreshTasksPagination = render;
  }

  attachObserver();
  render();
  window.addEventListener('load', render);
})();

/* ===========================
   Selección por página + Eliminar seleccionados (FIX)
=========================== */
(function () {
  // Normaliza a string para evitar desajustes tipo/valor
  function toId(x) { return String(x ?? ''); }

  // Filas visibles (página actual)
  function visibleTaskRows() {
    const tbody = document.getElementById('tbody');
    if (!tbody) return [];
    return Array.from(tbody.querySelectorAll('tr')).filter(
      (tr) =>
        !tr.classList.contains('is-hidden-pager') &&
        tr.hasAttribute('data-id')
    );
  }

  // Marca/desmarca selección para una fila concreta
  function setRowSelectedByTr(tr, checked) {
    const cb = tr.querySelector('.rowCheck');
    if (cb) cb.checked = !!checked;
    const id = toId(tr.getAttribute('data-id'));
    if (state?.selection instanceof Set) {
      if (checked) state.selection.add(id);
      else state.selection.delete(id);
    }
    tr.classList.toggle('row-selected', !!checked); // opcional visual
  }

  function closePicker() {
    const overlay = document.getElementById('bulkPicker');
    if (!overlay) return;
    overlay.classList.add('hidden');
    overlay.dataset.selected = '';
    const accept = overlay.querySelector('.picker-accept');
    if (accept) {
      accept.disabled = true;
      accept.style.display = '';
    }
    overlay.querySelectorAll('.picker-option').forEach((btn) => btn.classList.remove('is-selected'));
    const searchWrap = overlay.querySelector('.picker-search');
    const searchInput = searchWrap?.querySelector('input');
    if (searchInput) searchInput.value = '';
    if (searchWrap) searchWrap.style.display = 'none';
  }
  function ensurePicker() {
    let overlay = document.getElementById('bulkPicker');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'bulkPicker';
      overlay.className = 'picker-overlay hidden';
      overlay.innerHTML = `
        <div class="picker-card">
          <div class="picker-title"></div>
          <div class="picker-search"><input type="search" placeholder="Buscar..." /></div>
          <div class="picker-options"></div>
          <div class="picker-footer">
            <button type="button" class="btn secondary picker-cancel">Cerrar</button>
            <button type="button" class="btn picker-accept" disabled>Aceptar</button>
          </div>
        </div>`;
      document.body.appendChild(overlay);
      overlay.addEventListener('click', (e) => {
        if (e.target === overlay) closePicker();
      });
      overlay.querySelector('.picker-cancel').addEventListener('click', closePicker);
    }
    return overlay;
  }
  function openPicker(title, options, onSelect, opts = {}) {
    if (!options || !options.length) return;
    const overlay = ensurePicker();
    overlay.classList.remove('hidden');
    const titleEl = overlay.querySelector('.picker-title');
    const listEl = overlay.querySelector('.picker-options');
    const acceptBtn = overlay.querySelector('.picker-accept');
    const searchWrap = overlay.querySelector('.picker-search');
    const searchInput = searchWrap?.querySelector('input');
    const requireConfirm = !!opts.confirm;
    if (titleEl) titleEl.textContent = title;
    let selected = null;
    const originalOpts = options.slice();
    const assignSelect = (opt, btnEl) => {
      selected = opt;
      overlay.dataset.selected = opt.value ?? opt.label;
      overlay.querySelectorAll('.picker-option').forEach((btn) => btn.classList.remove('is-selected'));
      btnEl?.classList.add('is-selected');
      if (acceptBtn) acceptBtn.disabled = !selected;
    };
    const renderList = (source) => {
      if (!listEl) return;
      listEl.innerHTML = '';
      source.forEach((opt) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'picker-option';
        btn.textContent = opt.label;
        btn.addEventListener('click', () => {
          if (requireConfirm) assignSelect(opt, btn);
          else {
            closePicker();
            if (typeof onSelect === 'function') onSelect(opt);
          }
        });
        listEl.appendChild(btn);
      });
      if (acceptBtn) acceptBtn.disabled = !selected;
    };
    renderList(originalOpts);
    if (acceptBtn) {
      acceptBtn.style.display = requireConfirm ? '' : 'none';
      acceptBtn.onclick = () => {
        if (!selected) return;
        closePicker();
        if (typeof onSelect === 'function') onSelect(selected);
      };
    }
    if (opts.searchable && searchWrap && searchInput) {
      searchWrap.style.display = 'block';
      searchInput.value = '';
      searchInput.oninput = () => {
        const val = searchInput.value.trim().toLowerCase();
        const filtered = originalOpts.filter((opt) =>
          opt.label.toLowerCase().includes(val)
        );
        renderList(filtered);
      };
    } else if (searchWrap) {
      searchWrap.style.display = 'none';
    }
  }
  window.__openPicker = openPicker;

  function bulkUpdateSelected(mutator) {
    const ids = new Set(Array.from(state?.selection || [], toId));
    if (!ids.size) return 0;
    let touched = 0;
    const nowIso = new Date().toISOString();
    state.tasks.forEach((task) => {
      if (!ids.has(toId(task.id))) return;
      mutator(task);
      task.updatedAt = nowIso;
      task.fechaActualizacion = nowIso;
      touched++;
    });
    if (touched) {
      saveTasks();
      applyFilters();
      setTimeout(syncHeaderAndBulk, 0);
    }
    return touched;
  }

  // Sincroniza checkbox del encabezado y botones masivos
  function syncHeaderAndBulk() {
    const hdr = document.getElementById('checkAll');
    if (hdr) {
      const rows = visibleTaskRows();
      if (rows.length === 0) {
        hdr.checked = false;
        hdr.indeterminate = false;
      } else {
        let checked = 0;
        rows.forEach((tr) => {
          const cb = tr.querySelector('.rowCheck');
          if (cb && cb.checked) checked++;
        });
        hdr.checked = checked === rows.length;
        hdr.indeterminate = checked > 0 && checked < rows.length;
      }
    }
    const n = (state?.selection && state.selection.size) || 0;
    const isEditing = document.body.classList.contains('editing');
    const isAgent = roleKey(state.user?.role || '') === ROLE_KEY_AGENTE;
    const bulkButtons = [
      { id: 'btnDeleteSelected', label: 'Eliminar' },
      { id: 'btnBulkStatus', label: 'Cambiar estado' },
      { id: 'btnBulkAssign', label: 'Asignar A:' },
      { id: 'btnBulkActions', label: '+' },
    ];
    bulkButtons.forEach(({ id, label }) => {
      const el = document.getElementById(id);
      if (!el) return;
      if ((id === 'btnDeleteSelected' || id === 'btnBulkAssign') && isAgent) {
        el.classList.add('hidden');
        el.disabled = true;
        el.textContent = label;
        return;
      }
      el.disabled = false;
      if (n > 0 && !isEditing) {
        el.classList.remove('hidden');
        el.textContent = id === 'btnBulkActions' ? label : `${label} (${n})`;
      } else {
        el.classList.add('hidden');
        el.textContent = label;
      }
    });
  }
  window.__syncHeaderCheckbox = syncHeaderAndBulk;

  // Encabezado: seleccionar SOLO visibles
  const hdr = document.getElementById('checkAll');
  if (hdr) {
    hdr.addEventListener('change', (e) => {
      const rows = visibleTaskRows();
      rows.forEach((tr) => setRowSelectedByTr(tr, e.target.checked));
      syncHeaderAndBulk();
    });
  }

  // Checkbox de fila
  const tbody = document.getElementById('tbody');
  if (tbody) {
    tbody.addEventListener('change', (e) => {
      if (!e.target.classList.contains('rowCheck')) return;
      const tr = e.target.closest('tr');
      if (!tr) return;
      setRowSelectedByTr(tr, e.target.checked);
      syncHeaderAndBulk();
    });

    // Alternar al hacer click en la fila (evita botones/inputs)
    tbody.addEventListener('click', (e) => {
      const tr = e.target.closest('tr');
      if (!tr) return;
      if (
        e.target.closest('button') ||
        e.target.closest('a') ||
        e.target.closest('input') ||
        e.target.closest('select') ||
        e.target.closest('textarea')
      ) return;
      const cb = tr.querySelector('.rowCheck');
      if (cb) {
        cb.checked = !cb.checked;
        setRowSelectedByTr(tr, cb.checked);
        syncHeaderAndBulk();
      }
    });
  }

  // Botón Eliminar (masivo)
  const btnDel = document.getElementById('btnDeleteSelected');
  if (btnDel) {
    btnDel.addEventListener('click', () => {
      // Para Agente, ocultamos y bloqueamos eliminación por ahora (posible apertura futura por estado/autor).
      if (roleKey(state.user?.role || '') === ROLE_KEY_AGENTE) {
        alert('Los agentes no pueden eliminar negocios en esta versión.');
        return;
      }
      const selectedIds = Array.from(state?.selection || [], toId);
      if (!selectedIds.length) return;

      const selectedTasks = state.tasks.filter((t) => selectedIds.includes(toId(t.id)));
      const deletable = selectedTasks.filter((t) => t?.fueEditado !== true);
      const blocked = selectedTasks.filter((t) => t?.fueEditado === true);

      if (!deletable.length) {
        alert(
          'Los negocios seleccionados ya fueron gestionados y no pueden ser eliminados. Cámbialos de estado en lugar de eliminarlos.'
        );
        return;
      }
      if (blocked.length) {
        alert(
          'Algunos negocios seleccionados ya fueron gestionados y no pueden ser eliminados. Solo se eliminarán los que no han sido modificados.'
        );
      }

      const deleteCount = deletable.length;
      if (
        !confirm(
          `¿Eliminar ${deleteCount} negocio${deleteCount === 1 ? '' : 's'} seleccionad${
            deleteCount === 1 ? 'o' : 'os'
          } no editad${deleteCount === 1 ? 'o' : 'os'}?`
        )
      ) {
        return;
      }

      const ids = new Set(deletable.map((t) => toId(t.id)));
      state.tasks = state.tasks.filter((t) => !ids.has(toId(t.id)));

      // Limpiar selección y refrescar (pueden quedar seleccionados los bloqueados, se limpian para evitar confusión)
      state.selection.clear();
      saveTasks();
      applyFilters();

      // Reset header checkbox
      const hdr = document.getElementById('checkAll');
      if (hdr) {
        hdr.checked = false;
        hdr.indeterminate = false;
      }
      setTimeout(syncHeaderAndBulk, 0);
    });
  }

  const btnBulkStatus = document.getElementById('btnBulkStatus');
  if (btnBulkStatus) {
    btnBulkStatus.addEventListener('click', () => {
      const n = state?.selection?.size || 0;
      if (!n) return;
      const idSet = new Set(Array.from(state.selection || [], toId));
      const selectedTasks = state.tasks.filter((task) => idSet.has(toId(task.id)));
      const commonOptions = TASK_STATUS_OPTIONS.filter((status) => {
        const normalized = normalizeTaskStatus(status);
        return selectedTasks.every((task) => getAllowedNextStates(task).includes(normalized));
      });
      if (!commonOptions.length) {
        alert('Los negocios seleccionados no tienen cambios de estado disponibles en común.');
        return;
      }
      const statusOptions = commonOptions.map((s) => ({ label: s, value: s }));
      openPicker('Cambiar de estado a:', statusOptions, (opt) => {
        if (!opt) return;
        const scope = n === 1 ? 'el negocio seleccionado' : `${n} negocios seleccionados`;
        if (!confirm(`¿Cambiar ${scope} al estado "${opt.label}"?`)) return;
        const targetStatus = normalizeTaskStatus(opt.value);
        let updated = 0;
        let blocked = 0;
        const nowIso = new Date().toISOString();
        state.tasks.forEach((task) => {
          if (!idSet.has(toId(task.id))) return;
          const allowed = getAllowedNextStates(task);
          if (!allowed.includes(targetStatus) || !isStatusAllowedForChannel(targetStatus, task.canalIngreso)) {
            blocked++;
            return;
          }
          const oldStatus = task.status || 'Sin estado';
          if (oldStatus === targetStatus) return;
          pushBitacoraEntry(task, `Estado cambiado de "${oldStatus}" a "${targetStatus}"`);
          task.status = targetStatus;
          task.updatedAt = nowIso;
          task.fechaActualizacion = nowIso;
          updated++;
        });
        if (!updated) {
          alert('Ningún negocio pudo cambiar de estado con la opción seleccionada.');
          return;
        }
        saveTasks();
        applyFilters();
        setTimeout(syncHeaderAndBulk, 0);
        let message = `Estado actualizado a "${opt.label}" para ${updated} negocio${updated === 1 ? '' : 's'}.`;
        if (blocked)
          message += ` ${blocked} negocio${blocked === 1 ? '' : 's'} no permitieron este cambio.`;
        alert(message);
      });
    });
  }

  const btnBulkAssign = document.getElementById('btnBulkAssign');
  if (btnBulkAssign) {
    btnBulkAssign.addEventListener('click', () => {
      if (roleKey(state.user?.role || '') === ROLE_KEY_AGENTE) return;
      const n = state?.selection?.size || 0;
      if (!n) return;
      const options = [];
      const current = getCurrentUserIdentity();
      const role = roleKey(current.role);
      const rut = current.brokerRut || current.rutCorredora || '';
      (state.users || [])
        .filter((u) => u.active !== false)
        .filter((u) => {
          if (role === ROLE_KEY_SUPERVISOR || role === ROLE_KEY_ADMINISTRADOR) {
            return sameRutCorredora(u.brokerRut || u.rutCorredora || '', rut);
          }
          return true;
        })
        .forEach((u) => {
          options.push({
            label: `${u.name || u.email} (${u.email})`,
            value: u.email,
          });
        });
      if (!options.length) {
        alert('No hay usuarios disponibles para asignar.');
        return;
      }
      openPicker('Asignar a:', options, (opt) => {
        if (!opt) return;
        const targetLabel = opt.value ? opt.label : 'Sin asignar';
        const scope = n === 1 ? 'el negocio seleccionado' : `${n} negocios seleccionados`;
        if (!confirm(`¿Asignar ${scope} a ${targetLabel}?`)) return;
        const updated = bulkUpdateSelected((task) => {
          const prevSnapshot = getAssignmentSnapshot(task);
          task.assignee = opt.value || '';
          applyAssigneeMetadata(task, opt.value || '');
          logAssignmentChange(task, prevSnapshot);
        });
        if (updated) {
          alert(`Asignación actualizada para ${updated} negocio${updated === 1 ? '' : 's'}.`);
        }
      }, { searchable: true });
    });
  }

  document.addEventListener('keydown', (e) => {
    if (!isUserLoggedIn()) return;
    if (e.key === 'Escape') {
      const overlay = document.getElementById('bulkPicker');
      if (overlay && !overlay.classList.contains('hidden')) {
        closePicker();
      }
    }
  });

  // Re-sincroniza al cambiar de página o tamaño
  ['btnPrev', 'btnNext', 'perPage'].forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.addEventListener('click', () => setTimeout(syncHeaderAndBulk, 0));
    el.addEventListener('change', () => setTimeout(syncHeaderAndBulk, 0));
  });

  // Hook: tras repintar/paginar
  const oldRefresh = window.refreshTasksPagination;
  window.refreshTasksPagination = function () {
    if (typeof oldRefresh === 'function') oldRefresh();
    syncHeaderAndBulk();
  };

  // Primer sync
  setTimeout(syncHeaderAndBulk, 0);
})();

/* ===========================
   Debug / ejemplos de catálogos
   (Preparación de datos, no conectado a la UI)
=========================== */
// Mostrar catálogos cargados
console.log('Catálogo motivos pérdida:', MOTIVOS_PERDIDA);
console.log('Catálogo motivos reversa:', MOTIVOS_REVERSA);
console.log('Catálogo motivos anulación:', MOTIVOS_ANULACION);
console.log('Estados de renovación:', ESTADOS_RENOVACION);
console.log('Estados de recaudación:', ESTADOS_RECAUDACION);

// Ejemplo de uso de los modelos conceptuales y validadores
const demoRenovacion = crearRenovacion({
  estadoRenovacion: 'Perdido',
  fechaPreaviso: '2024-05-01',
  fechaPropuesta: '2024-05-10',
  fechaDecision: '2024-05-20',
  motivoPerdidaOAnulacion: 'Precio',
  comentario: 'Cliente tomó oferta más barata.',
});
console.log(
  'Renovación demo válida:',
  esEstadoRenovacionValido(demoRenovacion.estadoRenovacion) &&
    (['Perdido', 'Anulado'].includes(demoRenovacion.estadoRenovacion)
      ? esMotivoPerdidaValido(demoRenovacion.motivoPerdidaOAnulacion) ||
        esMotivoAnulacionValido(demoRenovacion.motivoPerdidaOAnulacion)
      : true),
  demoRenovacion
);

const demoPago = crearPagoCuota({
  estadoRecaudacion: 'Reversado/Devolución',
  metodoPago: 'Transferencia',
  fechaCobro: '2024-06-15',
  referenciaPago: 'REF-12345',
  conciliado: false,
  polizaNumero: 'POL-2024-001',
  cuotaNumero: '1',
  motivoReversa: 'Pago duplicado',
});
console.log(
  'Pago demo válido:',
  esEstadoRecaudacionValido(demoPago.estadoRecaudacion) &&
    (demoPago.estadoRecaudacion === 'Reversado/Devolución'
      ? esMotivoReversaValido(demoPago.motivoReversa)
      : true),
  demoPago
);

/* ===========================
   Autologin demo (opcional)
=========================== */
(function () {
  const params = new URLSearchParams(location.search);
  const tenant = params.get('tenant') || 'andes-corredores';
  const sessionKey = 'corredores_v1:' + tenant + ':session';

  function seedSession() {
    const sess = {
      tenant,
      email: 'admin@andescorredores.cl',
      name: 'Admin Andescorredores',
      role: ROLE_ADMINISTRADOR,
      isMaestro: false,
    };
    try {
      localStorage.setItem(sessionKey, JSON.stringify(sess));
    } catch {}
  }

  function goIn() {
    try {
      const btn =
        document.getElementById('btnLogin') ||
        document.querySelector('[data-action="login"]') ||
        document.querySelector('button.btnLogin');
      if (btn) {
        btn.click();
        return;
      }
    } catch {}
    location.href = location.pathname + '?tenant=' + encodeURIComponent(tenant);
  }

  function seedAndEnter() {
    seedSession();
    goIn();
  }

  window.__kensaDemoLogin = seedAndEnter;

  if (
    params.get('demo') === '1' ||
    params.get('autologin') === '1' ||
    location.hash === '#demo'
  ) {
    if (document.readyState === 'loading')
      document.addEventListener('DOMContentLoaded', seedAndEnter);
    else seedAndEnter();
  }
})();
