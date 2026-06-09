import { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

const ADMIN_USER = {
  id: "ADMIN",
  name: "Administrador UAH",
  password: "uah2024admin",
  role: "admin",
};

const SUPERVISOR_USER = {
  id: "SUPER",
  name: "Supervisor UAH",
  password: "uah2024super",
  role: "supervisor",
};

const INITIAL_USERS = [
  { id: "UAH001", name: "RAUL",                  password: "RAUL", role: "user" },
  { id: "UAH002", name: "Carlos Ramírez",         password: "1234", role: "user" },
  { id: "UAH003", name: "Sofía Hernández",        password: "1234", role: "user" },
  { id: "UAH004", name: "Luis Torres",            password: "1234", role: "user" },
  { id: "UAH005", name: "Ana Martínez",           password: "1234", role: "user" },
  { id: "UAH006", name: "Pedro López",            password: "1234", role: "user" },
  { id: "UAH007", name: "Laura Sánchez",          password: "1234", role: "user" },
  { id: "UAH008", name: "Miguel Flores",          password: "1234", role: "user" },
  { id: "UAH009", name: "Valeria Reyes",          password: "1234", role: "user" },
  { id: "UAH010", name: "Jorge Castro",           password: "1234", role: "user" },
  { id: "UAH011", name: "Diana Morales",          password: "1234", role: "user" },
  { id: "UAH012", name: "Andrés Jiménez",         password: "1234", role: "user" },
  { id: "UAH013", name: "Paola Vargas",           password: "1234", role: "user" },
  { id: "UAH014", name: "Roberto Díaz",           password: "1234", role: "user" },
  { id: "UAH015", name: "Gabriela Núñez",         password: "1234", role: "user" },
  { id: "UAH016", name: "Fernando Castillo",      password: "1234", role: "user" },
  { id: "UAH017", name: "Mariana Gutiérrez",      password: "1234", role: "user" },
  { id: "UAH018", name: "Alejandro Mendoza",      password: "1234", role: "user" },
  { id: "UAH019", name: "Patricia Ruiz",          password: "1234", role: "user" },
  { id: "UAH020", name: "Héctor Vega",            password: "1234", role: "user" },
  { id: "UAH021", name: "Carmen Ríos",            password: "1234", role: "user" },
  { id: "UAH022", name: "Eduardo Peña",           password: "1234", role: "user" },
  { id: "UAH023", name: "Lucía Campos",           password: "1234", role: "user" },
  { id: "UAH024", name: "Gustavo Medina",         password: "1234", role: "user" },
  { id: "UAH025", name: "Beatriz Ortega",         password: "1234", role: "user" },
  { id: "UAH026", name: "Óscar Romero",           password: "1234", role: "user" },
  { id: "UAH027", name: "Verónica Aguilar",       password: "1234", role: "user" },
  { id: "UAH028", name: "Ramón Silva",             password: "1234", role: "user" },
  { id: "UAH029", name: "Claudia Espinoza",       password: "1234", role: "user" },
  { id: "UAH030", name: "Ernesto Guerrero",       password: "1234", role: "user" },
  { id: "UAH031", name: "Irma Delgado",           password: "1234", role: "user" },
  { id: "UAH032", name: "Sergio Navarro",         password: "1234", role: "user" },
  { id: "UAH033", name: "Norma Rojas",            password: "1234", role: "user" },
  { id: "UAH034", name: "Arturo Fuentes",         password: "1234", role: "user" },
  { id: "UAH035", name: "Gloria Herrera",         password: "1234", role: "user" },
  { id: "UAH036", name: "Javier Cervantes",       password: "1234", role: "user" },
  { id: "UAH037", name: "Rosa Sandoval",          password: "1234", role: "user" },
  { id: "UAH038", name: "Felipe Contreras",       password: "1234", role: "user" },
  { id: "UAH039", name: "Adriana Ramos",          password: "1234", role: "user" },
  { id: "UAH040", name: "Manuel Acosta",          password: "1234", role: "user" },
  { id: "UAH041", name: "Silvia Mendez",          password: "1234", role: "user" },
  { id: "UAH042", name: "Rubén Ibarra",           password: "1234", role: "user" },
  { id: "UAH043", name: "Alicia Montes",          password: "1234", role: "user" },
  { id: "UAH044", name: "Daniel Palacios",        password: "1234", role: "user" },
  { id: "UAH045", name: "Elena Cabrera",          password: "1234", role: "user" },
  { id: "UAH046", name: "Rodrigo Estrada",        password: "1234", role: "user" },
  { id: "UAH047", name: "Mónica Villanueva",      password: "1234", role: "user" },
  { id: "UAH048", name: "Hugo Salinas",           password: "1234", role: "user" },
  { id: "UAH049", name: "Teresa Ávila",           password: "1234", role: "user" },
  { id: "UAH050", name: "Enrique Bravo",          password: "1234", role: "user" },
  { id: "UAH051", name: "Lorena Cárdenas",        password: "1234", role: "user" },
  { id: "UAH052", name: "Ignacio Paredes",        password: "1234", role: "user" },
  { id: "UAH053", name: "Yolanda Esquivel",       password: "1234", role: "user" },
  { id: "UAH054", name: "Ernesto Villalba",       password: "1234", role: "user" },
  { id: "UAH055", name: "Concepción Trujillo",    password: "1234", role: "user" },
  { id: "UAH056", name: "Braulio Soria",          password: "1234", role: "user" },
  { id: "UAH057", name: "Graciela Zamora",        password: "1234", role: "user" },
  { id: "UAH058", name: "Alfredo Leal",           password: "1234", role: "user" },
  { id: "UAH059", name: "Esperanza Quintero",     password: "1234", role: "user" },
  { id: "UAH060", name: "Nicolás Carrillo",       password: "1234", role: "user" },
  { id: "UAH061", name: "Fabiola Domínguez",      password: "1234", role: "user" },
  { id: "UAH062", name: "César Tapia",            password: "1234", role: "user" },
  { id: "UAH063", name: "Maribel Fuentes",        password: "1234", role: "user" },
  { id: "UAH064", name: "Aurelio Ponce",          password: "1234", role: "user" },
  { id: "UAH065", name: "Dolores Serrano",        password: "1234", role: "user" },
  { id: "UAH066", name: "Marcos Galván",          password: "1234", role: "user" },
  { id: "UAH067", name: "Pilar Elizalde",         password: "1234", role: "user" },
  { id: "UAH068", name: "Gerardo Lozano",         password: "1234", role: "user" },
  { id: "UAH069", name: "Minerva Cortés",         password: "1234", role: "user" },
  { id: "UAH070", name: "Adolfo Meza",            password: "1234", role: "user" },
  { id: "UAH071", name: "Rebeca Solis",           password: "1234", role: "user" },
  { id: "UAH072", name: "Genaro Pedraza",         password: "1234", role: "user" },
  { id: "UAH073", name: "Leticia Vargas",         password: "1234", role: "user" },
  { id: "UAH074", name: "Mauricio Ibáñez",        password: "1234", role: "user" },
  { id: "UAH075", name: "Amparo Cisneros",        password: "1234", role: "user" },
  { id: "UAH076", name: "Edmundo Olvera",         password: "1234", role: "user" },
];

const DEFAULT_MUNICIPIOS = {
  "Pachuca de Soto":     ["Centro","Morales","Cubitos","Venta Prieta","Caminos"],
  "Tulancingo":          ["Centro","San Juan","Tepepa","Jaltepec","El Chico"],
  "Tizayuca":            ["Centro","Fraccionamiento Valle","San Pedro","La Laguna","Colinas"],
  "Actopan":             ["Centro","El Mondongo","Santuario","Tepetitlán","Mixquiahuala"],
  "Ixmiquilpan":         ["Centro","El Alberto","Dios Padre","Tasquillo","Defay"],
  "Otro":                ["Sin sección"],
};

const QUESTIONS = [
  {
    id: 1,
    text: "¿Cuál considera usted que es el principal obstáculo para que los jóvenes y adultos de su municipio continúen con su formación profesional licenciatura o técnica?",
    options: [
      "Falta de escuelas cercanas",
      "Costos elevados de colegiaturas y/o materiales",
      "Falta de horarios flexibles para trabajar y estudiar",
      "Problemas de transporte o acceso a internet",
      "Falta de interés o motivación",
    ],
  },
  {
    id: 2,
    text: "¿Qué nivel educativo considera más necesario fortalecer en su municipio?",
    options: [
      "Bachillerato",
      "Licenciaturas",
      "Cursos técnicos y de oficio",
      "Capacitación para emprender",
      "Educación digital y tecnológica",
    ],
  },
  {
    id: 3,
    text: "¿Qué sector económico necesita más apoyo técnico y capacitación en su municipio?",
    options: [
      "Comercio local",
      "Agricultura y campo",
      "Emprendimiento y negocios",
      "Tecnología y sistemas",
      "Servicios y turismo",
    ],
  },
  {
    id: 4,
    text: "¿Cuál de los siguientes criterios considera más justo para otorgar becas educativas en su comunidad?",
    options: [
      "Excelencia académica",
      "Madres solteras y jefes de familia",
      "Jóvenes de escasos recursos",
      "Vocación de servicio comunitario",
      "Beca deportiva",
    ],
  },
  {
    id: 5,
    text: "Si el Maestro Luis Ángel Carrasco Gasca acercara a su municipio brigadas especializadas de apoyo por parte de la Universidad Alianza Hispana, ¿cuál aprovecharía usted o su familia?",
    options: [
      "Consultoría legal y administrativa",
      "Brigadas de atención psicológica",
      "Asesoría en proyectos productivos y para el campo",
    ],
  },
  {
    id: 6,
    text: "¿Qué tan importante considera que el Mtro. Luis Ángel Carrasco Gasca recorra personalmente los municipios para gestionar soluciones a las necesidades de la región?",
    options: ["Poco importante", "Importante", "Muy importante"],
  },
];

const SURVEY_DONE = QUESTIONS.length + 1;
// 3 minutos exactos
const RUSH_THRESHOLD_SEC = 180;

const CHART_PALETTE = [
  { bar: "#4F8EF7", light: "#E8F0FE" },
  { bar: "#22C987", light: "#D6F5EB" },
  { bar: "#F7844F", light: "#FEE8D6" },
  { bar: "#C94FBF", light: "#F5D6F3" },
  { bar: "#F7C94F", light: "#FEF6D6" },
  { bar: "#4FC9F7", light: "#D6F3FE" },
];

const STORAGE_KEY   = "uah_encuesta_db";
const IDB_NAME      = "uah_offline_db";
const IDB_STORE     = "pending_responses";
const IDB_VERSION   = 1;

// ─── IndexedDB helpers ────────────────────────────────────────────────
function openIDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE, { keyPath: "id" });
      }
    };
    req.onsuccess  = (e) => resolve(e.target.result);
    req.onerror    = ()  => reject(new Error("IDB error"));
  });
}

async function idbSavePending(response) {
  try {
    const db = await openIDB();
    return new Promise((res, rej) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).put({ ...response, _pending: true });
      tx.oncomplete = () => res(true);
      tx.onerror    = () => rej();
    });
  } catch { return false; }
}

async function idbGetAllPending() {
  try {
    const db = await openIDB();
    return new Promise((res) => {
      const tx  = db.transaction(IDB_STORE, "readonly");
      const req = tx.objectStore(IDB_STORE).getAll();
      req.onsuccess = () => res(req.result || []);
      req.onerror   = () => res([]);
    });
  } catch { return []; }
}

async function idbDeletePending(id) {
  try {
    const db = await openIDB();
    return new Promise((res) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).delete(id);
      tx.oncomplete = () => res(true);
      tx.onerror    = () => res(false);
    });
  } catch { return false; }
}

async function idbClearAllPending() {
  try {
    const db = await openIDB();
    return new Promise((res) => {
      const tx = db.transaction(IDB_STORE, "readwrite");
      tx.objectStore(IDB_STORE).clear();
      tx.oncomplete = () => res(true);
      tx.onerror    = () => res(false);
    });
  } catch { return false; }
}

function loadDb() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      // ensure municipios key exists
      if (!parsed.municipios) parsed.municipios = DEFAULT_MUNICIPIOS;
      return parsed;
    }
  } catch {}
  return { users: INITIAL_USERS, responses: [], municipios: DEFAULT_MUNICIPIOS };
}

function persistDb(db) {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); } catch {}
}

// ─── EXCEL EXPORT COMPLETO ─────────────────────────────────────────────
function buildXlsx(db) {
  const wb = XLSX.utils.book_new();

  // ── Helper: ajusta ancho de columnas al contenido más largo ──────────
  function autoWidth(ws) {
    if (!ws["!ref"]) return ws;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const colWidths = [];
    for (let C = range.s.c; C <= range.e.c; C++) {
      let maxLen = 8; // mínimo
      for (let R = range.s.r; R <= range.e.r; R++) {
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.v != null) {
          const len = String(cell.v).length;
          if (len > maxLen) maxLen = len;
        }
      }
      colWidths.push({ wch: Math.min(maxLen + 2, 80) }); // +2 padding, max 80
    }
    ws["!cols"] = colWidths;
    // Congelar la primera fila (encabezados siempre visibles al scrollear)
    if (!ws["!freeze"]) ws["!freeze"] = { xSplit: 0, ySplit: 1 };
    return ws;
  }

  // ── Helper: aplica hoja con autoWidth y la agrega al libro ───────────
  function addSheet(wb, data, name) {
    const ws = XLSX.utils.aoa_to_sheet(data);
    autoWidth(ws);
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  const now = new Date();
  const todayStr = now.toISOString().slice(0,10);
  const weekAgo  = new Date(now - 7  * 86400000);
  const monthAgo = new Date(now - 30 * 86400000);
  const total    = db.responses.length;
  const rushAll  = db.responses.filter(r => r.isRush).length;

  // ── calcular stats de entrevistadores (mismo algoritmo que la UI) ─────
  const byUser = {};
  db.users.forEach(u => {
    byUser[u.id] = { user: u, all: [], today: 0, week: 0, month: 0, rush: 0, totalSec: 0, countSec: 0 };
  });
  db.responses.forEach(r => {
    if (!byUser[r.userId]) byUser[r.userId] = { user: { id: r.userId, name: r.userName }, all: [], today: 0, week: 0, month: 0, rush: 0, totalSec: 0, countSec: 0 };
    const s = byUser[r.userId];
    s.all.push(r);
    const d = new Date(r.timestamp);
    if (r.timestamp?.slice(0,10) === todayStr) s.today++;
    if (d >= weekAgo)  s.week++;
    if (d >= monthAgo) s.month++;
    if (r.isRush) s.rush++;
    if (r.elapsedSec > 0) { s.totalSec += r.elapsedSec; s.countSec++; }
  });
  const rankRows = Object.values(byUser).filter(s => s.all.length > 0).sort((a,b) => b.all.length - a.all.length);
  const avgTotal = rankRows.length > 0 ? rankRows.reduce((a,b) => a + b.all.length, 0) / rankRows.length : 0;

  // ── stats por municipio ───────────────────────────────────────────────
  const munMap = {};
  db.responses.forEach(r => {
    const k = r.municipio || "Sin municipio";
    if (!munMap[k]) munMap[k] = { total: 0, rush: 0 };
    munMap[k].total++;
    if (r.isRush) munMap[k].rush++;
  });

  // ── horas ─────────────────────────────────────────────────────────────
  const hourBuckets = Array(24).fill(0);
  db.responses.forEach(r => { if (r.hour !== undefined) hourBuckets[r.hour]++; });

  // ────────────────────────────────────────────────────────────────────
  // HOJA 1 — ANÁLISIS GRÁFICO POR PREGUNTA (primera hoja, la más importante)
  // ────────────────────────────────────────────────────────────────────
  const h1 = [
    ["UNIVERSIDAD ALIANZA HISPANA — DIAGNÓSTICO EDUCATIVO"],
    [`Reporte generado: ${now.toLocaleString("es-MX")}   |   Total encuestas: ${total}   |   ⚡ Apresuradas: ${rushAll}`],
    [],
  ];

  QUESTIONS.forEach((q) => {
    const counts  = q.options.map(opt => db.responses.filter(r => r.answers[q.id] === opt).length);
    const rushCts = q.options.map(opt => db.responses.filter(r => r.isRush && r.answers[q.id] === opt).length);
    const totQ    = counts.reduce((a, b) => a + b, 0);
    const maxVal  = Math.max(...counts, 1);
    const maxIdx  = counts.indexOf(Math.max(...counts));
    const rushTotal = rushCts.reduce((a, b) => a + b, 0);
    const pctLider  = totQ > 0 ? Math.round(((counts[maxIdx] || 0) / totQ) * 100) : 0;

    // Separador y título
    h1.push(["════════════════════════════════════════════════════════════"]);
    h1.push([`PREGUNTA ${q.id}`, rushTotal > 0 ? `⚡ apr.` : ""]);
    h1.push([q.text]);
    h1.push([]);

    // Strip de stats (igual que en pantalla)
    h1.push([
      `${totQ} respuestas`,
      rushTotal > 0 ? `${rushTotal} apresuradas` : "",
      `${pctLider}% opción líder`,
    ]);
    h1.push([]);

    // Cabecera de tabla
    h1.push(["Opción de Respuesta", "Votos", "%", "⚡ Apresuradas", "Barra visual"]);

    // Una fila por opción
    q.options.forEach((opt, i) => {
      const pct   = totQ > 0 ? Math.round((counts[i] / totQ) * 100) : 0;
      const rCt   = rushCts[i];
      const bar   = counts[i] > 0 ? "█".repeat(Math.round((counts[i] / maxVal) * 20)) : "·";
      const lider = counts[i] === counts[maxIdx] && totQ > 0 ? " ★" : "";
      const rushLabel = rCt > 0 ? `⚡ ${rCt} apr.` : "—";
      h1.push([`${opt}${lider}`, counts[i], `${pct}%`, rushLabel, bar]);
    });

    h1.push(["TOTAL", totQ, "100%", rushTotal > 0 ? `⚡ ${rushTotal}` : "—", ""]);
    h1.push([]);
    h1.push([]);
  });

  // ── Mini resumen ejecutivo al final ──────────────────────────────────
  h1.push(["════════════════════════════════════════════════════════════"]);
  h1.push(["RESUMEN EJECUTIVO"]);
  h1.push([]);
  h1.push(["Total encuestas respondidas",     total]);
  h1.push(["Total encuestadores registrados", db.users.length]);
  h1.push(["Tasa de participación (%)",       db.users.length > 0 ? Math.round((total/db.users.length)*100) : 0]);
  h1.push(["⚡ Encuestas apresuradas (<3 min)", rushAll]);
  h1.push(["% encuestas apresuradas",         total > 0 ? Math.round((rushAll/total)*100) : 0]);
  h1.push([]);
  h1.push(["Municipio","Total","Apresuradas","% del total","% apr. del municipio"]);
  Object.entries(munMap).sort((a,b)=>b[1].total-a[1].total).forEach(([m,st]) => {
    h1.push([m, st.total, st.rush,
      total > 0 ? Math.round((st.total/total)*100) : 0,
      st.total > 0 ? Math.round((st.rush/st.total)*100) : 0]);
  });
  h1.push([]);
  h1.push(["Posición","ID","Nombre","Total","Hoy","Esta semana","Este mes","⚡ Apresuradas","% apresuradas","Tiempo prom.","Rendimiento"]);
  rankRows.forEach((s, i) => {
    const avgSec = s.countSec > 0 ? Math.round(s.totalSec/s.countSec) : 0;
    const avgStr = avgSec < 60 ? `${avgSec}s` : `${Math.floor(avgSec/60)}m ${avgSec%60}s`;
    const rushPct = s.all.length > 0 ? Math.round((s.rush/s.all.length)*100) : 0;
    const isLow = s.all.length < avgTotal*0.5 && avgTotal > 0;
    h1.push([i+1, s.user.id, s.user.name, s.all.length, s.today, s.week, s.month,
      s.rush, rushPct, avgStr,
      isLow ? "⚠ Rendimiento bajo" : rushPct > 30 ? "⚠ Muchas apresuradas" : "✓ Normal"]);
  });

  addSheet(wb, h1, "Analisis-Preguntas");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 2 — TODAS LAS ENCUESTAS (con apresuradas destacadas)
  // ────────────────────────────────────────────────────────────────────
  const h2header = ["#","ID Respuesta","Encuestador ID","Nombre Encuestador","👤 Nombre Encuestado","📞 Teléfono Encuestado","Municipio","Sección","Fecha","Hora","Duración","⚡ Apresurada (<3min)"];
  QUESTIONS.forEach(q => h2header.push(`P${q.id}: ${q.text.slice(0,45)}`));
  const h2 = [h2header];
  db.responses.forEach((r, i) => {
    const row = [
      i+1, r.id, r.userId, r.userName,
      r.encuestadoNombre || "Sin nombre",
      r.encuestadoTel    || "Sin teléfono",
      r.municipio || "Sin municipio",
      r.seccion   || "Sin sección",
      r.dateStr, r.timeStr,
      r.duration || "—",
      r.isRush ? "SÍ ⚡" : "NO"
    ];
    QUESTIONS.forEach(q => row.push(r.answers[q.id] || "Sin respuesta"));
    h2.push(row);
  });
  addSheet(wb, h2, "Encuestas");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 3 — ENCUESTAS APRESURADAS (solo las <3 min)
  // ────────────────────────────────────────────────────────────────────
  const rushList = db.responses.filter(r => r.isRush);
  const h3header = ["#","ID Respuesta","Usuario ID","Nombre","Municipio","Sección","Fecha","Hora","Duración (seg aprox)"];
  QUESTIONS.forEach(q => h3header.push(`P${q.id}`));
  const h3 = [
    ["⚡ ENCUESTAS APRESURADAS — respondidas en menos de 3 minutos"],
    [`Total apresuradas: ${rushList.length} de ${total} (${total>0?Math.round((rushList.length/total)*100):0}%)`],
    [],
    h3header,
    ...rushList.map((r, i) => {
      const row = [i+1, r.id, r.userId, r.userName, r.municipio||"Sin municipio", r.seccion||"Sin sección", r.dateStr, r.timeStr, r.duration||"—"];
      QUESTIONS.forEach(q => row.push(r.answers[q.id] || "Sin respuesta"));
      return row;
    })
  ];
  addSheet(wb, h3, "Apresuradas");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 4 — GRÁFICAS / RESULTADOS POR PREGUNTA
  // ────────────────────────────────────────────────────────────────────
  const h4 = [
    ["RESULTADOS GENERALES POR PREGUNTA — mismo dato que las gráficas en pantalla"],
    [`Total encuestas: ${total}`],
    [],
  ];
  QUESTIONS.forEach((q, qi) => {
    const counts  = q.options.map(opt => db.responses.filter(r => r.answers[q.id] === opt).length);
    const rushCts = q.options.map(opt => db.responses.filter(r => r.isRush && r.answers[q.id] === opt).length);
    const maxVal  = Math.max(...counts, 1);
    const maxIdx  = counts.indexOf(Math.max(...counts));
    h4.push([`PREGUNTA ${q.id}`, q.text]);
    h4.push(["Opción","Votos","Porcentaje (%)","Apresuradas en esta opción","% apresuradas / opción","Barra visual"]);
    q.options.forEach((opt, i) => {
      const pct  = total > 0 ? Math.round((counts[i]/total)*100) : 0;
      const rPct = counts[i] > 0 ? Math.round((rushCts[i]/counts[i])*100) : 0;
      const bar  = "█".repeat(Math.round((counts[i]/maxVal)*20));
      h4.push([
        opt, counts[i], pct, rushCts[i], rPct, bar
      ]);
    });
    h4.push([
      "→ OPCIÓN LÍDER",
      q.options[maxIdx] || "—",
      `${total > 0 ? Math.round((counts[maxIdx]||0)/total*100) : 0}%`,
      "", "", ""
    ]);
    h4.push([]);
  });
  addSheet(wb, h4, "Graficas-Resultados");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 5 — RANKING DE PRODUCTIVIDAD
  // ────────────────────────────────────────────────────────────────────
  const h5 = [
    ["🏆 RANKING DE PRODUCTIVIDAD — Entrevistadores"],
    [`Generado: ${now.toLocaleString("es-MX")}`],
    [],
    ["Posición","ID","Nombre","Total encuestas","Hoy","Esta semana","Este mes","⚡ Apresuradas","% Apresuradas","Tiempo promedio","Eficiencia (%)","Alerta"],
    ...rankRows.map((s, i) => {
      const avgSec  = s.countSec > 0 ? Math.round(s.totalSec/s.countSec) : 0;
      const avgStr  = avgSec < 60 ? `${avgSec}s` : `${Math.floor(avgSec/60)}m ${avgSec%60}s`;
      const rushPct = s.all.length > 0 ? Math.round((s.rush/s.all.length)*100) : 0;
      const eff     = 100 - rushPct;
      const medal   = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : `${i+1}°`;
      const isLow   = s.all.length < avgTotal*0.5 && avgTotal > 0;
      const alerta  = isLow ? "⚠ Rendimiento bajo" : rushPct > 30 ? "⚠ Muchas apresuradas" : "✓ OK";
      return [medal, s.user.id, s.user.name, s.all.length, s.today, s.week, s.month, s.rush, rushPct, avgStr, eff, alerta];
    }),
    [],
    ["─── ENTREVISTAS POR HORA DEL DÍA ───"],
    ["Hora","Encuestas"],
    ...hourBuckets.map((v, h) => [`${h}:00 - ${h}:59`, v]),
  ];
  addSheet(wb, h5, "Ranking-Entrevistadores");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 6 — RESULTADOS POR MUNICIPIO
  // ────────────────────────────────────────────────────────────────────
  const h6 = [
    ["RESULTADOS POR MUNICIPIO"],
    [],
  ];
  const municipios = [...new Set(db.responses.map(r => r.municipio || "Sin municipio"))].sort();
  municipios.forEach(mun => {
    const rsMun = db.responses.filter(r => (r.municipio || "Sin municipio") === mun);
    h6.push([`══ MUNICIPIO: ${mun} ══`, `Total: ${rsMun.length}`, `Apresuradas: ${rsMun.filter(r=>r.isRush).length}`, `% del total: ${total>0?Math.round((rsMun.length/total)*100):0}%`]);
    h6.push([]);
    QUESTIONS.forEach(q => {
      const counts = q.options.map(opt => rsMun.filter(r => r.answers[q.id] === opt).length);
      const maxIdx = counts.indexOf(Math.max(...counts));
      h6.push([`P${q.id}: ${q.text.slice(0,60)}`, "Votos", "Porcentaje (%)", "Apresuradas"]);
      q.options.forEach((opt, i) => {
        const pct = rsMun.length > 0 ? Math.round((counts[i]/rsMun.length)*100) : 0;
        const rc  = rsMun.filter(r => r.isRush && r.answers[q.id] === opt).length;
        h6.push([opt, counts[i], pct, rc]);
      });
      h6.push(["→ Opción líder", q.options[maxIdx]||"—", `${rsMun.length>0?Math.round((counts[maxIdx]||0)/rsMun.length*100):0}%`, ""]);
      h6.push([]);
    });
    h6.push(["─".repeat(60)]);
    h6.push([]);
  });
  addSheet(wb, h6, "Por-Municipio");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 7 — USUARIOS Y CONTRASEÑAS
  // ────────────────────────────────────────────────────────────────────
  const h7 = [
    ["CREDENCIALES Y ACCESOS — UAH"],
    [`Generado: ${now.toLocaleString("es-MX")}`],
    [],
    ["═══ ROLES ESPECIALES ═══"],
    ["Rol","ID de acceso","Nombre","Contraseña","Permisos"],
    ["Administrador", ADMIN_USER.id, ADMIN_USER.name, ADMIN_USER.password, "Lectura · Escritura · Eliminación · Exportación · Gestión completa"],
    ["Supervisor",    SUPERVISOR_USER.id, SUPERVISOR_USER.name, SUPERVISOR_USER.password, "Solo lectura de gráficas con porcentajes"],
    [],
    ["═══ ENCUESTADORES (76) ═══"],
    ["#","ID","Nombre","Contraseña","Respondió","Fecha respuesta","Municipio","Sección","Duración","⚡ Apresurada"],
    ...db.users.map((u, i) => {
      const res = db.responses.find(r => r.userId === u.id);
      return [
        i+1, u.id, u.name, u.password,
        res ? "✓ SÍ" : "Pendiente",
        res ? `${res.dateStr} ${res.timeStr}` : "—",
        res ? (res.municipio||"Sin municipio") : "—",
        res ? (res.seccion||"Sin sección")    : "—",
        res ? (res.duration||"—")             : "—",
        res ? (res.isRush ? "SÍ ⚡" : "NO")   : "—",
      ];
    }),
  ];
  addSheet(wb, h7, "Usuarios-Contrasenas");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 8 — UBICACIONES GPS Y DISPOSITIVOS
  // ────────────────────────────────────────────────────────────────────
  const h8 = [
    ["UBICACIONES GPS Y DATOS DE ACCESO"],
    [`Generado: ${now.toLocaleString("es-MX")}`],
    [],
    ["#","ID Respuesta","Usuario ID","Nombre","Fecha","Hora","Municipio","Latitud","Longitud","Precisión GPS (m)","Google Maps Link","Dispositivo","Navegador","Sistema Operativo","Idioma","Resolución pantalla","Zona horaria"],
    ...db.responses.map((r, i) => [
      i+1, r.id, r.userId, r.userName,
      r.dateStr, r.timeStr, r.municipio||"Sin municipio",
      r.gps ? r.gps.lat : "Sin GPS",
      r.gps ? r.gps.lng : "Sin GPS",
      r.gps ? `±${r.gps.accuracy}m` : "—",
      r.gps ? `https://maps.google.com/?q=${r.gps.lat},${r.gps.lng}` : "—",
      r.device ? (r.device.isMobile ? "Móvil" : "Desktop") : "Desconocido",
      r.device?.browser   || "—",
      r.device?.platform  || "—",
      r.device?.lang      || "—",
      r.device?.screen    || "—",
      r.device?.tz        || "—",
    ]),
  ];
  addSheet(wb, h8, "GPS-Dispositivos");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 9 — DIRECTORIO DE CONTACTOS (nombre + teléfono + datos clave)
  // ────────────────────────────────────────────────────────────────────
  const contactos = db.responses.filter(r => r.encuestadoNombre && r.encuestadoNombre !== "Sin nombre");
  const h9 = [
    ["📞 DIRECTORIO DE CONTACTOS — UAH"],
    [`Generado: ${now.toLocaleString("es-MX")}`],
    [`Total contactos con nombre: ${contactos.length} de ${total}`],
    [],
    ["#", "👤 Nombre Completo", "📞 Teléfono / WhatsApp", "🗺️ Municipio", "📍 Sección", "📅 Fecha", "⏰ Hora", "⏱ Duración", "⚡ Apresurada", "🆔 Encuestador ID", "👤 Nombre Encuestador", "WhatsApp Link"],
  ];
  contactos.forEach((r, i) => {
    const tel = r.encuestadoTel && r.encuestadoTel !== "Sin teléfono" ? r.encuestadoTel : "";
    const waLink = tel ? `https://wa.me/52${tel.replace(/\D/g, "")}` : "—";
    h9.push([
      i + 1,
      r.encuestadoNombre || "—",
      tel || "Sin teléfono",
      r.municipio || "Sin municipio",
      r.seccion || "Sin sección",
      r.dateStr || "—",
      r.timeStr || "—",
      r.duration || "—",
      r.isRush ? "SÍ ⚡" : "NO",
      r.userId || "—",
      r.userName || "—",
      waLink,
    ]);
  });
  h9.push([]);
  h9.push(["─── RESUMEN POR MUNICIPIO ───"]);
  h9.push(["Municipio", "Total Contactos", "Con Teléfono", "Sin Teléfono", "% Con Teléfono"]);
  const munContactMap = {};
  contactos.forEach(r => {
    const m = r.municipio || "Sin municipio";
    if (!munContactMap[m]) munContactMap[m] = { total: 0, conTel: 0 };
    munContactMap[m].total++;
    if (r.encuestadoTel && r.encuestadoTel !== "Sin teléfono") munContactMap[m].conTel++;
  });
  Object.entries(munContactMap).sort((a, b) => b[1].total - a[1].total).forEach(([m, st]) => {
    const sinTel = st.total - st.conTel;
    const pctTel = st.total > 0 ? Math.round((st.conTel / st.total) * 100) : 0;
    h9.push([m, st.total, st.conTel, sinTel, pctTel]);
  });
  addSheet(wb, h9, "Directorio-Contactos");

  // ────────────────────────────────────────────────────────────────────
  // HOJA 10 — ANÁLISIS GRÁFICO (formato idéntico a la pantalla)
  // ────────────────────────────────────────────────────────────────────
  const h10 = [
    ["📊 ANÁLISIS GRÁFICO DE RESULTADOS — Universidad Alianza Hispana"],
    [`Reporte generado: ${now.toLocaleString("es-MX")}   |   Total encuestas: ${total}`],
    [],
  ];

  QUESTIONS.forEach((q) => {
    const counts  = q.options.map(opt => db.responses.filter(r => r.answers[q.id] === opt).length);
    const rushCts = q.options.map(opt => db.responses.filter(r => r.isRush && r.answers[q.id] === opt).length);
    const totQ    = counts.reduce((a, b) => a + b, 0);
    const maxVal  = Math.max(...counts, 1);
    const maxIdx  = counts.indexOf(Math.max(...counts));
    const rushTotal = rushCts.reduce((a, b) => a + b, 0);
    const pctLider  = totQ > 0 ? Math.round(((counts[maxIdx] || 0) / totQ) * 100) : 0;
    const hasRush   = rushTotal > 0;

    // ── Bloque encabezado de pregunta (igual que la tarjeta en pantalla) ──
    h10.push([`PREGUNTA ${q.id}`, hasRush ? "⚡ con apresuradas" : ""]);
    h10.push([q.text]);
    h10.push([]);

    // ── Fila de resumen (igual al strip de stats de la UI) ──
    h10.push(["RESUMEN", "Respuestas", "Apresuradas", "% opción líder", "Opción líder"]);
    h10.push(["", totQ, rushTotal, `${pctLider}%`, q.options[maxIdx] || "—"]);
    h10.push([]);

    // ── Una fila por opción (igual a cada barra en pantalla) ──
    h10.push(["Opción de Respuesta", "Votos", "%", "⚡ Apresuradas en esta opción", "Barra visual (20 blks)"]);
    q.options.forEach((opt, i) => {
      const pct  = totQ > 0 ? Math.round((counts[i] / totQ) * 100) : 0;
      const rCt  = rushCts[i];
      const bar  = counts[i] > 0 ? "█".repeat(Math.round((counts[i] / maxVal) * 20)) : "·";
      const lider = counts[i] === counts[maxIdx] && totQ > 0 ? " ★" : "";
      const rushLabel = rCt > 0 ? `⚡ ${rCt} apr.` : "";
      h10.push([`${opt}${lider}`, counts[i], pct, rushLabel, bar]);
    });

    h10.push(["TOTAL", totQ, 100, "", ""]);
    h10.push([]);
    h10.push([]);
  });

  // ── TABLA RESUMEN PLANA (para gráficas de Excel con un clic) ─────────
  h10.push(["══════ TABLA RESUMEN PARA INSERTAR GRÁFICA EN EXCEL ══════"]);
  h10.push(["Instrucción: selecciona las columnas Opción + % y luego Insertar > Gráfico de barras"]);
  h10.push([]);
  QUESTIONS.forEach(q => {
    const counts  = q.options.map(opt => db.responses.filter(r => r.answers[q.id] === opt).length);
    const rushCts = q.options.map(opt => db.responses.filter(r => r.isRush && r.answers[q.id] === opt).length);
    const totQ    = counts.reduce((a, b) => a + b, 0);
    const maxIdx  = counts.indexOf(Math.max(...counts));

    h10.push([`P${q.id}: ${q.text.slice(0, 70)}`]);
    h10.push(["Opción", "Votos", "% (graficar)", "⚡ Apresuradas", "% Apresuradas", "★ Líder"]);
    q.options.forEach((opt, i) => {
      const pct  = totQ > 0 ? parseFloat(((counts[i] / totQ) * 100).toFixed(1)) : 0;
      const rPct = counts[i] > 0 ? parseFloat(((rushCts[i] / counts[i]) * 100).toFixed(1)) : 0;
      h10.push([opt, counts[i], pct, rushCts[i], rPct, i === maxIdx && totQ > 0 ? "★" : ""]);
    });
    h10.push(["TOTAL", totQ, 100, rushCts.reduce((a,b)=>a+b,0), "", ""]);
    h10.push([]);
  });

  // ── CRUCE POR MUNICIPIO ───────────────────────────────────────────────
  const munList10 = [...new Set(db.responses.map(r => r.municipio || "Sin municipio"))].sort();
  if (munList10.length > 0) {
    h10.push(["══════ RESULTADOS POR MUNICIPIO ══════"]);
    h10.push([]);
    QUESTIONS.forEach(q => {
      h10.push([`P${q.id}: ${q.text.slice(0, 70)}`]);
      const hdr = ["Opción", "TOTAL", "% global", ...munList10.flatMap(m => [`${m} votos`, `${m} %`])];
      h10.push(hdr);
      const totQ = db.responses.filter(r => r.answers[q.id]).length;
      q.options.forEach(opt => {
        const totalOpt = db.responses.filter(r => r.answers[q.id] === opt).length;
        const pctG = totQ > 0 ? parseFloat(((totalOpt / totQ) * 100).toFixed(1)) : 0;
        const munCols = munList10.flatMap(m => {
          const mTotal = db.responses.filter(r => (r.municipio||"Sin municipio") === m && r.answers[q.id]).length;
          const mVotos = db.responses.filter(r => (r.municipio||"Sin municipio") === m && r.answers[q.id] === opt).length;
          return [mVotos, mTotal > 0 ? parseFloat(((mVotos/mTotal)*100).toFixed(1)) : 0];
        });
        h10.push([opt, totalOpt, pctG, ...munCols]);
      });
      h10.push([]);
    });
  }

  addSheet(wb, h10, "Analisis-Grafico");

  return wb;
}

function exportToExcel(db) {
  try {
    // Incluir encuestas offline pendientes en el Excel
    let mergedDb = db;
    try {
      const pendingLocal = JSON.parse(localStorage.getItem("uah_pending_local") || "[]");
      if (pendingLocal.length > 0) {
        const existingIds = new Set(db.responses.map(r => r.id));
        const newOnes = pendingLocal.filter(r => !existingIds.has(r.id));
        if (newOnes.length > 0) {
          mergedDb = { ...db, responses: [...db.responses, ...newOnes] };
        }
      }
    } catch {}
    const wb = buildXlsx(mergedDb);
    // Configurar propiedades del libro para apertura sin restricciones
    wb.Props = {
      Title: "UAH Encuestas — Diagnóstico Educativo",
      Subject: "Resultados de encuesta UAH",
      Author: "Universidad Alianza Hispana",
      CreatedDate: new Date(),
    };
    // Primera hoja activa al abrir
    wb.Workbook = {
      Views: [{ RTL: false }],
      Sheets: wb.SheetNames.map((name, i) => ({
        Hidden: 0,           // 0 = visible, nunca ocultar
        state: "visible",
      })),
    };
    const fileName = `UAH_Encuestas_Completo_${new Date().toISOString().slice(0,10)}.xlsx`;
    XLSX.writeFile(wb, fileName, {
      bookType: "xlsx",
      bookSST: false,
      type: "binary",
      compression: true,
    });
  } catch(e) {
    exportAsCsv(db);
  }
}

function exportAsCsv(db) {
  // Fallback CSV si SheetJS no carga
  const rows = [["ID Respuesta","Usuario ID","Nombre","Municipio","Sección","Fecha","Hora","Duración","Apresurada"]];
  QUESTIONS.forEach(q => rows[0].push(`P${q.id}`));
  db.responses.forEach(r => {
    const row = [r.id, r.userId, r.userName, r.municipio, r.seccion, r.dateStr, r.timeStr, r.duration||"", r.isRush?"SÍ":"NO"];
    QUESTIONS.forEach(q => row.push(r.answers[q.id]||""));
    rows.push(row);
  });
  const csv = "\uFEFF" + rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  try {
    const a = document.createElement("a");
    a.href = "data:text/csv;charset=utf-8," + encodeURIComponent(csv);
    a.download = `UAH_Encuestas_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
  } catch(e) {
    navigator.clipboard?.writeText(csv).then(() => alert("Datos copiados. Pega en Excel con Ctrl+V.")).catch(()=>{});
  }
}

// ─── CHART COMPONENTS ────────────────────────────────────────────────

function ModernBarChart({ data, labels, palette, total, rushData }) {
  const max = Math.max(...data, 1);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
      {labels.map((label, i) => {
        const pct = total > 0 ? Math.round((data[i] / total) * 100) : 0;
        const barW = max > 0 ? (data[i] / max) * 100 : 0;
        const { bar, light } = palette[i % palette.length];
        const rushCount = rushData ? rushData[i] : 0;
        return (
          <div key={i}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "4px" }}>
              <span style={{ fontSize: "12px", color: "#444", fontWeight: 500, flex: 1, lineHeight: 1.3 }}>{label}</span>
              <div style={{ display: "flex", gap: "6px", alignItems: "center", flexShrink: 0, marginLeft: "8px" }}>
                {rushCount > 0 && (
                  <span style={{
                    background: "#FFF0F0", border: "1px solid #FFCDD2",
                    color: "#C62828", fontSize: "10px", fontWeight: "bold",
                    padding: "1px 6px", borderRadius: "20px", letterSpacing: "0.02em"
                  }}>⚡ {rushCount} apr.</span>
                )}
                <span style={{ fontSize: "13px", fontWeight: "bold", color: bar }}>{pct}%</span>
                <span style={{ fontSize: "11px", color: "#aaa" }}>({data[i]})</span>
              </div>
            </div>
            <div style={{ height: "10px", background: "#F0F0F0", borderRadius: "99px", overflow: "hidden" }}>
              <div style={{
                width: `${barW}%`, height: "100%",
                background: `linear-gradient(90deg, ${bar}, ${bar}cc)`,
                borderRadius: "99px",
                transition: "width 0.8s cubic-bezier(.4,0,.2,1)",
              }} />
            </div>
          </div>
        );
      })}
    </div>
  );
}

function DonutChart({ data, labels, palette }) {
  const total = data.reduce((a, b) => a + b, 0);
  if (total === 0) return (
    <div style={{ textAlign: "center", color: "#aaa", padding: "1.5rem", fontSize: "13px" }}>Sin respuestas aún</div>
  );

  const size = 140, cx = 70, cy = 70, r = 52, strokeW = 20;
  const circumference = 2 * Math.PI * r;
  let offset = 0;

  const slices = data.map((val, i) => {
    const pct = val / total;
    const dashArray = `${pct * circumference} ${circumference}`;
    const rotation = offset * 360 - 90;
    offset += pct;
    return { dashArray, rotation, color: palette[i % palette.length].bar, val, pct };
  });

  const maxIdx = data.indexOf(Math.max(...data));

  return (
    <div style={{ display: "flex", gap: "16px", alignItems: "center", flexWrap: "wrap" }}>
      <div style={{ position: "relative", width: size, height: size, flexShrink: 0 }}>
        <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
          <circle cx={cx} cy={cy} r={r} fill="none" stroke="#F0F0F0" strokeWidth={strokeW} />
          {slices.map((s, i) => (
            <circle key={i} cx={cx} cy={cy} r={r} fill="none"
              stroke={s.color} strokeWidth={strokeW}
              strokeDasharray={s.dashArray}
              strokeDashoffset={0}
              transform={`rotate(${s.rotation} ${cx} ${cy})`}
              style={{ transition: "all 0.6s ease" }}
            />
          ))}
        </svg>
        <div style={{
          position: "absolute", top: "50%", left: "50%",
          transform: "translate(-50%,-50%)", textAlign: "center"
        }}>
          <div style={{ fontSize: "22px", fontWeight: "bold", color: palette[maxIdx % palette.length].bar, lineHeight: 1 }}>
            {Math.round((data[maxIdx] / total) * 100)}%
          </div>
          <div style={{ fontSize: "9px", color: "#999", lineHeight: 1.2, maxWidth: "50px" }}>más votado</div>
        </div>
      </div>
      <div style={{ flex: 1, minWidth: "140px" }}>
        {labels.map((l, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "7px" }}>
            <div style={{ width: "10px", height: "10px", borderRadius: "50%", background: palette[i % palette.length].bar, flexShrink: 0 }} />
            <span style={{ fontSize: "11px", color: "#555", flex: 1, lineHeight: 1.3 }}>{l}</span>
            <span style={{ fontSize: "12px", fontWeight: "bold", color: palette[i % palette.length].bar }}>
              {total > 0 ? Math.round((data[i] / total) * 100) : 0}%
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────

export default function App() {
  const [screen, setScreen]               = useState("login");
  const [currentUser, setCurrentUser]     = useState(null);
  const [loginId, setLoginId]             = useState("");
  const [loginPass, setLoginPass]         = useState("");
  const [loginError, setLoginError]       = useState("");
  const [answers, setAnswers]             = useState({});
  const [step, setStep]                   = useState(0);
  const [adminTab, setAdminTab]           = useState("responses");
  const [selectedResponse, setSelectedResponse] = useState(null);
  const [editUser, setEditUser]           = useState(null);
  const [surveyStartTime, setSurveyStartTime] = useState(null);
  const [showLoginPass, setShowLoginPass] = useState(false);
  const [showNewUserForm, setShowNewUserForm] = useState(false);
  const [newUser, setNewUser]             = useState({ id: "", name: "", password: "" });
  const [surveyMunicipio, setSurveyMunicipio]   = useState("");
  const [surveySeccion, setSurveySeccion]       = useState("");
  const [encuestadoNombre, setEncuestadoNombre] = useState("");
  const [encuestadoTel, setEncuestadoTel]       = useState("");
  const [filterMunicipio, setFilterMunicipio] = useState("Todos");
  const [filterSeccion, setFilterSeccion]     = useState("Todas");
  const [supTab, setSupTab]                   = useState("charts");
  const [responsePage, setResponsePage]       = useState(0);
  const [responseSearch, setResponseSearch]   = useState("");
  const [showSuperPass, setShowSuperPass]     = useState(false);
  const [showAdminPass, setShowAdminPass]     = useState(false);
  const RESPONSES_PER_PAGE = 50;

  // Modal de confirmación propio (reemplaza window.confirm bloqueado en iframe)
  const [confirmModal, setConfirmModal] = useState(null); // { msg, onOk }
  const [toastMsg, setToastMsg]         = useState("");

  const showConfirm = (msg, onOk) => setConfirmModal({ msg, onOk });
  const showToast   = (msg) => { setToastMsg(msg); setTimeout(() => setToastMsg(""), 2800); };
  const safeAlert   = (msg) => showToast(msg);

  // Municipios editor state
  const [editingMun, setEditingMun]       = useState(null); // { key, newName, secciones: [] }
  const [showAddMun, setShowAddMun]       = useState(false);
  const [newMunName, setNewMunName]       = useState("");
  const [newMunSecs, setNewMunSecs]       = useState("");
  const [newSecByMun, setNewSecByMun]     = useState({}); // { [munKey]: "nueva seccion" }

  const [db, setDb] = useState(loadDb);
  const [isOnline, setIsOnline]         = useState(navigator.onLine);
  const [pendingCount, setPendingCount] = useState(0);
  const [syncing, setSyncing]           = useState(false);
  const [lastSync, setLastSync]         = useState(null);

  // ── Permisos y geolocalización ──
  const [geoLocation, setGeoLocation]   = useState(null); // { lat, lng, accuracy, city? }
  const [permsDone, setPermsDone]       = useState(false);
  const [permsScreen, setPermsScreen]   = useState(false);
  const [deviceInfo, setDeviceInfo]     = useState(null);

  const collectDeviceInfo = () => {
    const ua = navigator.userAgent;
    const isMobile = /Android|iPhone|iPad|iPod/i.test(ua);
    const platform = isMobile
      ? /iPhone|iPad/.test(ua) ? "iOS" : "Android"
      : /Windows/.test(ua) ? "Windows" : /Mac/.test(ua) ? "macOS" : /Linux/.test(ua) ? "Linux" : "Desconocido";
    const browser = /Chrome\//.test(ua) && !/Edg/.test(ua) ? "Chrome"
      : /Firefox\//.test(ua) ? "Firefox"
      : /Safari\//.test(ua) && !/Chrome/.test(ua) ? "Safari"
      : /Edg\//.test(ua) ? "Edge" : "Otro";
    const lang     = navigator.language || "es";
    const screen   = `${window.screen.width}×${window.screen.height}`;
    const tz       = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return { platform, browser, isMobile, lang, screen, tz, ua: ua.slice(0,120) };
  };

  const requestPermissions = async () => {
    const info = collectDeviceInfo();
    setDeviceInfo(info);
    // GPS — el más importante
    const geoPromise = new Promise((resolve) => {
      if (!navigator.geolocation) { resolve(null); return; }
      navigator.geolocation.getCurrentPosition(
        (pos) => resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude, accuracy: Math.round(pos.coords.accuracy) }),
        ()    => resolve(null),
        { enableHighAccuracy: true, timeout: 10000 }
      );
    });
    // Micrófono
    try { const s = await navigator.mediaDevices?.getUserMedia({ audio: true }); s?.getTracks().forEach(t => t.stop()); } catch {}
    // Notificaciones
    try { await Notification?.requestPermission(); } catch {}
    // Esperar GPS
    const geo = await geoPromise;
    if (geo) setGeoLocation(geo);
    setPermsDone(true);
    setPermsScreen(false);
  };

  // ── Cargar cuántas pendientes hay al inicio ──
  useEffect(() => {
    idbGetAllPending().then(p => {
      const localPending = JSON.parse(localStorage.getItem("uah_pending_local") || "[]");
      const allIds = new Set([...p.map(r => r.id), ...localPending.map(r => r.id)]);
      setPendingCount(allIds.size);
    });
  }, []);

  // ── Detectar online/offline ──
  useEffect(() => {
    const goOnline  = () => { setIsOnline(true); };
    const goOffline = () => setIsOnline(false);
    window.addEventListener("online",  goOnline);
    window.addEventListener("offline", goOffline);
    return () => {
      window.removeEventListener("online",  goOnline);
      window.removeEventListener("offline", goOffline);
    };
  }, []);

  // ── Auto-sync cuando recupera conexión ──
  useEffect(() => {
    if (isOnline && pendingCount > 0) {
      syncPending();
    }
  // eslint-disable-next-line
  }, [isOnline]);

  const syncPending = async () => {
    setSyncing(true);
    const pending = await idbGetAllPending();
    // También revisar los pendientes en localStorage
    const pendingLocal = JSON.parse(localStorage.getItem("uah_pending_local") || "[]");
    const allPending = [...pending, ...pendingLocal];
    if (allPending.length === 0) { setSyncing(false); return; }
    // Fusionar pendientes con db (evitar duplicados por id)
    const existingIds = new Set(db.responses.map(r => r.id));
    const newOnes = allPending.filter(r => !existingIds.has(r.id)).map(({ _pending, ...r }) => r);
    if (newOnes.length > 0) {
      const merged = { ...db, responses: [...db.responses, ...newOnes] };
      saveDb(merged);
      await idbClearAllPending();
      localStorage.removeItem("uah_pending_local");
      setPendingCount(0);
      setLastSync(new Date().toLocaleTimeString("es-MX"));
      showToast(`✅ ${newOnes.length} encuesta${newOnes.length > 1 ? "s" : ""} sincronizada${newOnes.length > 1 ? "s" : ""} correctamente`);
    } else {
      await idbClearAllPending();
      localStorage.removeItem("uah_pending_local");
      setPendingCount(0);
    }
    setSyncing(false);
  };

  const dbGeneratedAt = useMemo(() => new Date().toISOString(), []);

  const responsesByUser = useMemo(
    () => new Map(db.responses.map((r) => [r.userId, r])),
    [db.responses]
  );

  // ── Incluir encuestas offline pendientes en todas las estadísticas ──
  const allResponsesMerged = useMemo(() => {
    try {
      const pendingLocal = JSON.parse(localStorage.getItem("uah_pending_local") || "[]");
      if (pendingLocal.length === 0) return db.responses;
      const existingIds = new Set(db.responses.map(r => r.id));
      const newOnes = pendingLocal.filter(r => !existingIds.has(r.id));
      return [...db.responses, ...newOnes];
    } catch { return db.responses; }
  }, [db.responses]);

  const allStats = useMemo(
    () => QUESTIONS.reduce((acc, q) => {
      acc[q.id] = q.options.map(
        (opt) => allResponsesMerged.filter((r) => r.answers[q.id] === opt).length
      );
      return acc;
    }, {}),
    [allResponsesMerged]
  );

  // Rush stats per question per option
  const rushStats = useMemo(
    () => QUESTIONS.reduce((acc, q) => {
      acc[q.id] = q.options.map(
        (opt) => allResponsesMerged.filter((r) => r.isRush && r.answers[q.id] === opt).length
      );
      return acc;
    }, {}),
    [allResponsesMerged]
  );

  // Supervisor stats
  const supStats = useMemo(
    () => QUESTIONS.reduce((acc, q) => {
      acc[q.id] = q.options.map(opt => allResponsesMerged.filter(r => r.answers[q.id] === opt).length);
      return acc;
    }, {}),
    [allResponsesMerged]
  );

  const supMunStats = useMemo(() => {
    const m = {};
    allResponsesMerged.forEach(r => {
      const key = r.municipio || "Sin municipio";
      if (!m[key]) m[key] = { total: 0, rush: 0 };
      m[key].total++;
      if (r.isRush) m[key].rush++;
    });
    return m;
  }, [allResponsesMerged]);

  useEffect(() => { persistDb(db); }, [db]);

  useEffect(() => {
    const handler = (e) => {
      if (e.key === "Escape") { setSelectedResponse(null); setEditUser(null); setEditingMun(null); }
    };
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, []);

  const saveDb = (newDb) => setDb(newDb);

  const logout = () => {
    setCurrentUser(null); setScreen("login");
    setLoginId(""); setLoginPass("");
    setSurveyMunicipio(""); setSurveySeccion("");
    setEncuestadoNombre(""); setEncuestadoTel("");
    setAnswers({}); setStep(0);
    setPermsDone(false); setPermsScreen(false);
  };

  const handleLogin = () => {
    setLoginError("");
    const trimId   = loginId.trim();
    const trimPass = loginPass;
    if (trimId === ADMIN_USER.id && trimPass === ADMIN_USER.password) {
      setCurrentUser(ADMIN_USER); setScreen("admin"); return;
    }
    if (trimId === SUPERVISOR_USER.id && trimPass === SUPERVISOR_USER.password) {
      setCurrentUser(SUPERVISOR_USER); setScreen("supervisor"); return;
    }
    const user = db.users.find((u) => u.id === trimId && u.password === trimPass);
    if (user) {
      setCurrentUser(user); setStep(0); setAnswers({});
      setScreen("survey");
    } else {
      setLoginError("ID o contraseña incorrectos.");
    }
  };

  const handleAnswer = (qId, option) => setAnswers((prev) => ({ ...prev, [qId]: option }));

  const handleSubmit = async () => {
    if (Object.keys(answers).length !== QUESTIONS.length) return;
    const userExists = db.users.find((u) => u.id === currentUser.id);
    if (!userExists) { logout(); return; }
    const now = new Date();
    const elapsedMs  = surveyStartTime ? Date.now() - surveyStartTime : 0;
    const elapsedSec = Math.round(elapsedMs / 1000);
    const durationStr = elapsedSec < 60
      ? `${elapsedSec}s`
      : `${Math.floor(elapsedSec / 60)}m ${elapsedSec % 60}s`;
    const isRush = elapsedSec > 0 && elapsedSec < RUSH_THRESHOLD_SEC;
    const response = {
      id: `R${Date.now()}`,
      userId: currentUser.id,
      userName: currentUser.name,
      timestamp: now.toISOString(),
      dateStr: now.toLocaleDateString("es-MX", { year: "numeric", month: "long", day: "numeric" }),
      timeStr: now.toLocaleTimeString("es-MX"),
      hour: now.getHours(),
      duration: durationStr,
      elapsedSec,
      isRush,
      municipio: surveyMunicipio || "Sin municipio",
      seccion: surveySeccion || "Sin sección",
      encuestadoNombre: encuestadoNombre.trim() || "Sin nombre",
      encuestadoTel:    encuestadoTel.trim()    || "Sin teléfono",
      answers: { ...answers },
      // ── Geolocalización y dispositivo ──
      gps: geoLocation ? { lat: geoLocation.lat, lng: geoLocation.lng, accuracy: geoLocation.accuracy } : null,
      device: deviceInfo || collectDeviceInfo(),
    };

    if (!navigator.onLine) {
      // ── MODO OFFLINE: guardar en IndexedDB Y en localStorage pendientes ──
      await idbSavePending(response);
      // Guardar también en localStorage para que el admin vea las pendientes
      const pendingLocal = JSON.parse(localStorage.getItem("uah_pending_local") || "[]");
      pendingLocal.push(response);
      localStorage.setItem("uah_pending_local", JSON.stringify(pendingLocal));
      const newCount = (await idbGetAllPending()).length;
      setPendingCount(newCount);
    } else {
      // ── MODO ONLINE: guardar normal ──
      saveDb({ ...db, responses: [...db.responses, response] });
    }
    setScreen("thanks");
  };

  const getStats = (qId) => allStats[qId] ?? [];

  const handleDeleteResponse = (id) => {
    showConfirm("¿Eliminar esta respuesta? Esta acción no se puede deshacer.", () => {
      saveDb({ ...db, responses: db.responses.filter((r) => r.id !== id) });
      setSelectedResponse(null);
      showToast("✅ Respuesta eliminada.");
    });
  };

  const handleDeleteUser = (userId) => {
    const u = db.users.find(u => u.id === userId);
    showConfirm(`¿Eliminar a "${u?.name || userId}" y todas sus respuestas? Esta acción no se puede deshacer.`, () => {
      saveDb({ ...db, users: db.users.filter((u) => u.id !== userId), responses: db.responses.filter((r) => r.userId !== userId) });
      showToast("✅ Usuario eliminado.");
    });
  };

  const handleEditUser = (user) => setEditUser({ ...user, _originalId: user.id });

  const handleSaveUser = () => {
    const originalId = editUser._originalId;
    const { _originalId, ...userToSave } = editUser;
    if (!userToSave.id.trim()) { safeAlert("El ID no puede estar vacío."); return; }
    const duplicate = db.users.find((u) => u.id === userToSave.id && u.id !== originalId);
    if (duplicate) { safeAlert(`El ID "${userToSave.id}" ya está en uso.`); return; }
    const idChanged = originalId !== userToSave.id;
    const updatedUsers = db.users.map((u) => u.id === originalId ? userToSave : u);
    const updatedResponses = idChanged
      ? db.responses.map((r) => r.userId === originalId ? { ...r, userId: userToSave.id } : r)
      : db.responses;
    saveDb({ ...db, users: updatedUsers, responses: updatedResponses });
    setEditUser(null);
  };

  const handleAddUser = () => {
    const trimId   = newUser.id.trim().toUpperCase();
    const trimName = newUser.name.trim();
    const trimPass = newUser.password.trim();
    if (!trimId)   { safeAlert("El ID no puede estar vacío."); return; }
    if (!trimName) { safeAlert("El nombre no puede estar vacío."); return; }
    if (!trimPass) { safeAlert("La contraseña no puede estar vacía."); return; }
    if (db.users.find((u) => u.id === trimId)) { safeAlert(`El ID "${trimId}" ya existe.`); return; }
    saveDb({ ...db, users: [...db.users, { id: trimId, name: trimName, password: trimPass, role: "user" }] });
    setNewUser({ id: "", name: "", password: "" });
    setShowNewUserForm(false);
  };

  // ── MUNICIPIOS ADMIN HANDLERS ──
  const handleSaveMun = () => {
    if (!editingMun) return;
    const { key, newName, secciones } = editingMun;
    const trimName = newName.trim();
    if (!trimName) { safeAlert("El nombre no puede estar vacío."); return; }
    const munObj = { ...db.municipios };
    if (trimName !== key) {
      // rename: copy secciones to new key, delete old
      if (munObj[trimName] && trimName !== key) { safeAlert("Ese nombre ya existe."); return; }
      munObj[trimName] = secciones.filter(s => s.trim());
      delete munObj[key];
      // update responses that had the old municipio name
      const updatedResponses = db.responses.map(r =>
        r.municipio === key ? { ...r, municipio: trimName } : r
      );
      saveDb({ ...db, municipios: munObj, responses: updatedResponses });
    } else {
      munObj[key] = secciones.filter(s => s.trim());
      saveDb({ ...db, municipios: munObj });
    }
    setEditingMun(null);
  };

  const handleDeleteMun = (key) => {
    showConfirm(`¿Eliminar el municipio "${key}"? Esta acción no se puede deshacer.`, () => {
      const munObj = { ...db.municipios };
      delete munObj[key];
      saveDb({ ...db, municipios: munObj });
      showToast("✅ Municipio eliminado.");
    });
  };

  const handleAddMun = () => {
    const trimName = newMunName.trim();
    if (!trimName) { safeAlert("El nombre no puede estar vacío."); return; }
    if (db.municipios[trimName]) { safeAlert("Ya existe un municipio con ese nombre."); return; }
    const secs = newMunSecs.split(",").map(s => s.trim()).filter(Boolean);
    saveDb({ ...db, municipios: { ...db.municipios, [trimName]: secs.length ? secs : ["Sin sección"] } });
    setNewMunName(""); setNewMunSecs(""); setShowAddMun(false);
  };

  const totalResponses = allResponsesMerged.length;
  const totalUsers     = db.users.length;
  const responseRate   = totalUsers > 0 ? Math.round((totalResponses / totalUsers) * 100) : 0;
  const rushCount      = allResponsesMerged.filter((r) => r.isRush).length;

  // ─── STYLES ──────────────────────────────────────────────────────────

  const inputSt = {
    width: "100%", padding: "10px 14px",
    background: "rgba(255,255,255,0.07)",
    border: "1px solid rgba(255,255,255,0.15)",
    borderRadius: "8px", color: "#fff", fontSize: "14px",
    outline: "none", boxSizing: "border-box",
  };

  const adminInputSt = {
    width: "100%", padding: "8px 12px", border: "1px solid #E0E0E0",
    borderRadius: "8px", fontSize: "13px", boxSizing: "border-box",
    outline: "none", fontFamily: "inherit",
  };

  // ─────────────────────────────────────────────────────────────────────
  // LOGIN
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "login") return (
    <>
      {!isOnline && (
        <div style={{ background: "#E53935", color: "#fff", textAlign: "center", padding: "8px", fontSize: "13px", fontWeight: "bold", fontFamily: "'Georgia',serif" }}>
          📴 SIN CONEXIÓN — Las encuestas se guardarán localmente y se enviarán al reconectarse
        </div>
      )}
      {isOnline && pendingCount > 0 && (
        <div style={{ background: "#FF9800", color: "#fff", textAlign: "center", padding: "8px", fontSize: "13px", fontWeight: "bold", fontFamily: "'Georgia',serif" }}>
          🔄 Conexión restaurada — {syncing ? "Sincronizando..." : `${pendingCount} encuesta${pendingCount>1?"s":""} pendiente${pendingCount>1?"s":""} por sincronizar`}
          {!syncing && <button onClick={syncPending} style={{ marginLeft: "12px", background: "#fff", color: "#E65100", border: "none", borderRadius: "4px", padding: "2px 10px", cursor: "pointer", fontSize: "12px", fontWeight: "bold" }}>Sincronizar ahora</button>}
        </div>
      )}
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0A1628 0%,#162B50 60%,#0A1628 100%)", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem", fontFamily: "'Georgia', serif" }}>
      <div style={{ background: "rgba(255,255,255,0.04)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: "20px", padding: "2.5rem", width: "100%", maxWidth: "380px", backdropFilter: "blur(10px)" }}>
        <div style={{ textAlign: "center", marginBottom: "2rem" }}>
          <div style={{ width: "64px", height: "64px", background: "linear-gradient(135deg,#C8A951,#E8C96B)", borderRadius: "14px", margin: "0 auto 1rem", display: "flex", alignItems: "center", justifyContent: "center", fontSize: "18px", fontWeight: "bold", color: "#0F1A2E", boxShadow: "0 4px 20px rgba(200,169,81,0.4)" }}>UAH</div>
          <h1 style={{ color: "#C8A951", fontSize: "20px", margin: "0 0 4px", letterSpacing: "0.05em" }}>Universidad Alianza Hispana</h1>
          <p style={{ color: "rgba(255,255,255,0.45)", fontSize: "13px", margin: 0 }}>Diagnóstico Educativo y Desarrollo Regional</p>
        </div>
        <div style={{ marginBottom: "1rem" }}>
          <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "6px", letterSpacing: "0.1em", textTransform: "uppercase" }}>ID de Acceso</label>
          <input value={loginId} onChange={(e) => setLoginId(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="Ej: UAH001" style={inputSt} />
        </div>
        <div style={{ marginBottom: "1.5rem" }}>
          <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "6px", letterSpacing: "0.1em", textTransform: "uppercase" }}>Contraseña</label>
          <div style={{ position: "relative" }}>
            <input type={showLoginPass ? "text" : "password"} value={loginPass} onChange={(e) => setLoginPass(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleLogin()} placeholder="••••" style={{ ...inputSt, paddingRight: "42px" }} />
            <button onClick={() => setShowLoginPass(!showLoginPass)} style={{ position: "absolute", right: "12px", top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: "rgba(255,255,255,0.4)", fontSize: "16px", padding: 0 }}>{showLoginPass ? "🙈" : "👁️"}</button>
          </div>
        </div>
        {loginError && <div style={{ color: "#F09595", fontSize: "13px", marginBottom: "1rem", textAlign: "center" }}>{loginError}</div>}
        <button onClick={handleLogin} style={{ width: "100%", padding: "13px", background: "linear-gradient(135deg,#C8A951,#E8C96B)", border: "none", borderRadius: "10px", color: "#0F1A2E", fontSize: "15px", fontWeight: "bold", cursor: "pointer", letterSpacing: "0.04em", boxShadow: "0 4px 14px rgba(200,169,81,0.35)" }}>Ingresar</button>
      </div>
    </div>
    </>
  );

  // ─────────────────────────────────────────────────────────────────────
  // YA RESPONDIÓ
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "already") return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0A1628,#162B50)", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem", fontFamily: "Georgia, serif" }}>
      <div style={{ textAlign: "center", color: "#fff" }}>
        <div style={{ fontSize: "56px", marginBottom: "1rem" }}>✓</div>
        <h2 style={{ color: "#C8A951", marginBottom: "1rem" }}>Ya respondiste la encuesta</h2>
        <p style={{ color: "rgba(255,255,255,0.6)", marginBottom: "2rem" }}>Tu participación ha sido registrada. ¡Gracias, {currentUser?.name}!</p>
        <button onClick={logout} style={{ padding: "10px 24px", background: "#C8A951", border: "none", borderRadius: "8px", color: "#0F1A2E", fontWeight: "bold", cursor: "pointer" }}>Salir</button>
      </div>
    </div>
  );

  // ─────────────────────────────────────────────────────────────────────
  // GRACIAS (post-submit)
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "thanks") return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(135deg,#0A1628,#162B50)", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Georgia, serif", padding: "1rem" }}>
      <div style={{ textAlign: "center", color: "#fff", maxWidth: "480px" }}>
        <div style={{ fontSize: "56px", marginBottom: "1rem" }}>🎓</div>
        <h2 style={{ color: "#C8A951", marginBottom: "1rem", fontSize: "22px" }}>¡Gracias por participar!</h2>
        <p style={{ color: "rgba(255,255,255,0.7)", lineHeight: "1.7", marginBottom: "2rem" }}>A nombre del Mtro. Luis Ángel Carrasco Gasca, le damos las gracias por responder la encuesta. Su participación ayudará a impulsar mejores oportunidades y programas para la comunidad.</p>
        <button onClick={logout} style={{ padding: "12px 32px", background: "linear-gradient(135deg,#C8A951,#E8C96B)", border: "none", borderRadius: "8px", color: "#0F1A2E", fontWeight: "bold", cursor: "pointer", fontSize: "15px" }}>Finalizar y Salir</button>
      </div>
    </div>
  );

  // ─────────────────────────────────────────────────────────────────────
  // ENCUESTA
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "survey") {
    const bg = "linear-gradient(135deg,#0A1628,#162B50)";
    const munKeys = Object.keys(db.municipios);


    // ── PANTALLA DE CIERRE: recopila nombre y teléfono antes del submit ──
    if (step === SURVEY_DONE) return (
      <div style={{ minHeight: "100vh", background: bg, fontFamily: "Georgia, serif", padding: "1rem 0" }}>
        {!isOnline && (
          <div style={{ background: "#E53935", color: "#fff", textAlign: "center", padding: "7px", fontSize: "12px", fontWeight: "bold", marginBottom: "4px" }}>
            📴 SIN CONEXIÓN — Tu encuesta se guardará aquí y se enviará automáticamente al reconectarte
          </div>
        )}
        <div style={{ maxWidth: "640px", margin: "0 auto", padding: "0 1rem" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
            <div style={{ color: "#C8A951", fontSize: "14px", fontWeight: "bold" }}>UAH — Encuesta</div>
            <div style={{ color: "rgba(255,255,255,0.4)", fontSize: "12px" }}>{currentUser?.name}</div>
          </div>
          <div style={{ background: "rgba(255,255,255,0.05)", border: "1px solid rgba(200,169,81,0.25)", borderRadius: "16px", padding: "2rem" }}>
            {/* Encabezado cierre */}
            <div style={{ textAlign: "center", marginBottom: "1.5rem" }}>
              <div style={{ fontSize: "42px", marginBottom: "0.5rem" }}>🎓</div>
              <h2 style={{ color: "#C8A951", fontSize: "20px", margin: "0 0 6px" }}>¡Encuesta completada!</h2>
              <p style={{ color: "rgba(255,255,255,0.45)", fontSize: "12px", margin: 0 }}>Cierre de la encuesta</p>
            </div>
            {/* Discurso de cierre */}
            <div style={{ background: "rgba(200,169,81,0.08)", border: "1px solid rgba(200,169,81,0.2)", borderRadius: "10px", padding: "1rem 1.2rem", marginBottom: "1.5rem" }}>
              <p style={{ color: "rgba(255,255,255,0.85)", fontSize: "13px", lineHeight: "1.8", margin: 0 }}>
                A nombre del <strong style={{ color: "#C8A951" }}>Mtro. Luis Ángel Carrasco Gasca</strong>, le damos las gracias por responder la encuesta.
              </p>
              <p style={{ color: "rgba(255,255,255,0.75)", fontSize: "13px", lineHeight: "1.8", margin: "0.75rem 0 0" }}>
                Con estos datos, la universidad va a planear brigadas y proyectos para lo que se necesita aquí en <strong style={{ color: "#C8A951" }}>{surveyMunicipio || "su municipio"}</strong>. Además, nos gustaría avisarle directamente cuando vengan estos servicios gratuitos y enviarle descuentos y beneficios especiales que la universidad tiene para usted y su familia.
              </p>
              <p style={{ color: "#C8A951", fontSize: "13px", fontWeight: "bold", margin: "0.75rem 0 0", lineHeight: "1.6" }}>
                ¿Me regala su nombre completo y su WhatsApp para mandarle la invitación directo a su teléfono?
              </p>
            </div>
            {/* Campos nombre y teléfono */}
            <div style={{ display: "flex", flexDirection: "column", gap: "12px", marginBottom: "1.5rem" }}>
              <div>
                <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "5px", textTransform: "uppercase", letterSpacing: "0.07em" }}>Nombre completo del encuestado *</label>
                <input
                  type="text"
                  value={encuestadoNombre}
                  onChange={e => setEncuestadoNombre(e.target.value)}
                  placeholder="Nombre completo de quien respondió…"
                  style={{ width: "100%", padding: "10px 12px", background: "rgba(255,255,255,0.07)", border: `1px solid ${encuestadoNombre.trim() ? "rgba(200,169,81,0.5)" : "rgba(255,255,255,0.2)"}`, borderRadius: "8px", color: "#fff", fontSize: "14px", outline: "none", boxSizing: "border-box" }}
                />
              </div>
              <div>
                <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "5px", textTransform: "uppercase", letterSpacing: "0.07em" }}>WhatsApp / Teléfono <span style={{ color: "rgba(255,255,255,0.3)", textTransform: "none", fontSize: "10px" }}>(opcional)</span></label>
                <input
                  type="tel"
                  value={encuestadoTel}
                  onChange={e => setEncuestadoTel(e.target.value.replace(/[^0-9+\-()\s]/g, ""))}
                  placeholder="10 dígitos…"
                  maxLength={15}
                  style={{ width: "100%", padding: "10px 12px", background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.2)", borderRadius: "8px", color: "#fff", fontSize: "14px", outline: "none", boxSizing: "border-box" }}
                />
              </div>
            </div>
            {/* Botones */}
            <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
              <button
                onClick={handleSubmit}
                disabled={!encuestadoNombre.trim()}
                style={{ width: "100%", padding: "14px", background: encuestadoNombre.trim() ? "linear-gradient(135deg,#C8A951,#E8C96B)" : "rgba(200,169,81,0.2)", border: "none", borderRadius: "10px", color: encuestadoNombre.trim() ? "#0F1A2E" : "rgba(200,169,81,0.4)", fontSize: "15px", fontWeight: "bold", cursor: encuestadoNombre.trim() ? "pointer" : "not-allowed" }}>
                Enviar Encuesta ✓
              </button>
              <button
                onClick={() => setStep(QUESTIONS.length)}
                style={{ width: "100%", padding: "11px", background: "transparent", border: "1px solid rgba(255,255,255,0.12)", borderRadius: "10px", color: "rgba(255,255,255,0.4)", fontSize: "13px", cursor: "pointer" }}>
                ← Regresar a la última pregunta
              </button>
            </div>
          </div>
        </div>
      </div>
    );



    return (
      <div style={{ minHeight: "100vh", background: bg, fontFamily: "Georgia, serif", padding: "1rem 0" }}>
        {/* ── BARRA OFFLINE ── */}
        {!isOnline && (
          <div style={{ background: "#E53935", color: "#fff", textAlign: "center", padding: "7px", fontSize: "12px", fontWeight: "bold", marginBottom: "4px" }}>
            📴 SIN CONEXIÓN — Tu encuesta se guardará aquí y se enviará automáticamente al reconectarte
          </div>
        )}
        {isOnline && pendingCount > 0 && (
          <div style={{ background: "#FF9800", color: "#fff", textAlign: "center", padding: "7px", fontSize: "12px", fontWeight: "bold", marginBottom: "4px" }}>
            {syncing ? "🔄 Sincronizando encuestas guardadas..." : `✅ Conexión restaurada — ${pendingCount} encuesta${pendingCount>1?"s":""} sincronizándose`}
          </div>
        )}
        <div style={{ maxWidth: "640px", margin: "0 auto", padding: "0 1rem" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
            <div style={{ color: "#C8A951", fontSize: "14px", fontWeight: "bold" }}>UAH — Encuesta</div>
            <div style={{ color: "rgba(255,255,255,0.4)", fontSize: "12px" }}>{currentUser?.name}</div>
          </div>

          {step === 0 ? (
            <div style={{ background: "rgba(255,255,255,0.05)", border: "1px solid rgba(200,169,81,0.25)", borderRadius: "16px", padding: "2rem" }}>
              <div style={{ textAlign: "center", marginBottom: "1.5rem" }}>
                <div style={{ display: "inline-block", background: "linear-gradient(135deg,#C8A951,#E8C96B)", borderRadius: "12px", padding: "12px 24px", marginBottom: "1rem" }}>
                  <div style={{ fontSize: "11px", letterSpacing: "0.15em", color: "#0F1A2E", fontWeight: "bold" }}>GAFETE OFICIAL</div>
                  <div style={{ fontSize: "18px", fontWeight: "bold", color: "#0F1A2E" }}>Universidad Alianza Hispana</div>
                  <div style={{ fontSize: "12px", color: "#1B2F4A" }}>Encuestador Autorizado</div>
                </div>
              </div>
              <div style={{ color: "rgba(255,255,255,0.85)", lineHeight: "1.8", fontSize: "14px" }}>
                <p><strong style={{ color: "#C8A951" }}>Buen día,</strong> nos identificamos como parte del equipo <strong style={{ color: "#C8A951" }}>UAH</strong>. Estamos realizando una encuesta de <strong style={{ color: "#C8A951" }}>Diagnóstico Educativo y Desarrollo Regional</strong> para la Universidad Alianza Hispana, por instrucción de nuestro Rector, el <strong style={{ color: "#C8A951" }}>Mtro. Luis Ángel Carrasco Gasca</strong>.</p>
                <div style={{ background: "rgba(200,169,81,0.1)", border: "1px solid rgba(200,169,81,0.2)", borderRadius: "8px", padding: "1rem", margin: "1rem 0" }}>
                  <p style={{ margin: "0 0 6px", fontWeight: "bold", color: "#C8A951", fontSize: "13px" }}>🎯 Objetivo de la encuesta:</p>
                  <p style={{ margin: 0, color: "rgba(255,255,255,0.7)", fontSize: "13px" }}>Conocer las necesidades educativas, sociales y de desarrollo del municipio para impulsar apoyos, becas, capacitaciones y programas que beneficien a la comunidad.</p>
                </div>
                <p style={{ color: "rgba(255,255,255,0.7)", fontSize: "13px" }}>La encuesta consta de <strong style={{ color: "#C8A951" }}>{QUESTIONS.length} preguntas</strong> de opción múltiple. Su participación es voluntaria y anónima. <strong style={{ color: "#C8A951" }}>Tiempo estimado: 2–5 minutos.</strong></p>
              </div>
              <div style={{ marginTop: "1.25rem", display: "flex", flexDirection: "column", gap: "10px" }}>

                <div>
                  <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "5px", textTransform: "uppercase", letterSpacing: "0.07em" }}>Municipio *</label>
                  <select value={surveyMunicipio} onChange={(e) => { setSurveyMunicipio(e.target.value); setSurveySeccion(""); }}
                    style={{ width: "100%", padding: "10px 12px", background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.2)", borderRadius: "8px", color: surveyMunicipio ? "#fff" : "rgba(255,255,255,0.4)", fontSize: "14px" }}>
                    <option value="" style={{ background: "#0F1A2E" }}>Selecciona un municipio…</option>
                    {munKeys.map((m) => <option key={m} value={m} style={{ background: "#0F1A2E" }}>{m}</option>)}
                  </select>
                </div>
                {surveyMunicipio && (
                  <div>
                    <label style={{ display: "block", color: "rgba(255,255,255,0.6)", fontSize: "11px", marginBottom: "5px", textTransform: "uppercase", letterSpacing: "0.07em" }}>Sección / Área</label>
                    <select value={surveySeccion} onChange={(e) => setSurveySeccion(e.target.value)}
                      style={{ width: "100%", padding: "10px 12px", background: "rgba(255,255,255,0.07)", border: "1px solid rgba(255,255,255,0.2)", borderRadius: "8px", color: surveySeccion ? "#fff" : "rgba(255,255,255,0.4)", fontSize: "14px" }}>
                      <option value="" style={{ background: "#0F1A2E" }}>Selecciona una sección…</option>
                      {(db.municipios[surveyMunicipio] || []).map((s) => <option key={s} value={s} style={{ background: "#0F1A2E" }}>{s}</option>)}
                    </select>
                  </div>
                )}

              </div>
              <button onClick={() => { setSurveyStartTime(Date.now()); setStep(1); }}
                disabled={!surveyMunicipio}
                style={{ width: "100%", padding: "14px", background: surveyMunicipio ? "linear-gradient(135deg,#C8A951,#E8C96B)" : "rgba(200,169,81,0.2)", border: "none", borderRadius: "10px", color: surveyMunicipio ? "#0F1A2E" : "rgba(200,169,81,0.4)", fontSize: "15px", fontWeight: "bold", cursor: surveyMunicipio ? "pointer" : "not-allowed", marginTop: "1rem" }}>
                Comenzar Encuesta →
              </button>
            </div>
          ) : (
            <div>
              <div style={{ marginBottom: "1.5rem" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "6px" }}>
                  <span style={{ color: "rgba(255,255,255,0.5)", fontSize: "12px" }}>Pregunta {step} de {QUESTIONS.length}</span>
                  <span style={{ color: "#C8A951", fontSize: "12px" }}>{Math.round(((step - 1) / QUESTIONS.length) * 100)}%</span>
                </div>
                <div style={{ height: "4px", background: "rgba(255,255,255,0.08)", borderRadius: "2px" }}>
                  <div style={{ height: "100%", width: `${((step - 1) / QUESTIONS.length) * 100}%`, background: "linear-gradient(90deg,#C8A951,#E8C96B)", borderRadius: "2px", transition: "width 0.4s" }} />
                </div>
              </div>
              <div style={{ background: "rgba(255,255,255,0.05)", border: "1px solid rgba(200,169,81,0.2)", borderRadius: "16px", padding: "2rem" }}>
                <div style={{ color: "rgba(200,169,81,0.5)", fontSize: "11px", letterSpacing: "0.1em", marginBottom: "12px", textTransform: "uppercase" }}>Pregunta {step}</div>
                <h3 style={{ color: "#fff", fontSize: "16px", lineHeight: "1.6", marginBottom: "1.5rem", fontWeight: "normal" }}>{QUESTIONS[step - 1].text}</h3>
                <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                  {QUESTIONS[step - 1].options.map((opt, i) => {
                    const selected = answers[step] === opt;
                    return (
                      <button key={i} onClick={() => handleAnswer(step, opt)}
                        style={{ padding: "12px 16px", background: selected ? "rgba(200,169,81,0.18)" : "rgba(255,255,255,0.04)", border: `1.5px solid ${selected ? "#C8A951" : "rgba(255,255,255,0.1)"}`, borderRadius: "8px", color: selected ? "#C8A951" : "rgba(255,255,255,0.75)", fontSize: "14px", cursor: "pointer", textAlign: "left", transition: "all 0.15s", display: "flex", alignItems: "center", gap: "12px" }}>
                        <div style={{ width: "18px", height: "18px", borderRadius: "50%", border: `2px solid ${selected ? "#C8A951" : "rgba(255,255,255,0.3)"}`, flexShrink: 0, background: selected ? "#C8A951" : "transparent", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          {selected && <div style={{ width: "7px", height: "7px", borderRadius: "50%", background: "#0F1A2E" }} />}
                        </div>
                        {opt}
                      </button>
                    );
                  })}
                </div>
                <div style={{ display: "flex", gap: "10px", marginTop: "1.5rem" }}>
                  {step > 1 && (
                    <button onClick={() => setStep(step - 1)} style={{ flex: 1, padding: "12px", background: "transparent", border: "1px solid rgba(255,255,255,0.15)", borderRadius: "8px", color: "rgba(255,255,255,0.6)", cursor: "pointer", fontSize: "14px" }}>← Anterior</button>
                  )}
                  {step < QUESTIONS.length ? (
                    <button onClick={() => answers[step] && setStep(step + 1)} disabled={!answers[step]}
                      style={{ flex: 2, padding: "12px", background: answers[step] ? "linear-gradient(135deg,#C8A951,#E8C96B)" : "rgba(200,169,81,0.15)", border: "none", borderRadius: "8px", color: answers[step] ? "#0F1A2E" : "rgba(200,169,81,0.35)", fontWeight: "bold", cursor: answers[step] ? "pointer" : "not-allowed", fontSize: "14px" }}>Siguiente →</button>
                  ) : (
                    <button onClick={() => answers[step] && setStep(SURVEY_DONE)} disabled={!answers[step]}
                      style={{ flex: 2, padding: "12px", background: answers[step] ? "linear-gradient(135deg,#22C987,#3AE09B)" : "rgba(34,201,135,0.15)", border: "none", borderRadius: "8px", color: answers[step] ? "#0F1A2E" : "rgba(34,201,135,0.35)", fontWeight: "bold", cursor: answers[step] ? "pointer" : "not-allowed", fontSize: "14px" }}>Continuar →</button>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>
      </div>
    );
  }

  // ─────────────────────────────────────────────────────────────────────
  // ADMIN
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "admin") {
    const tabs = [
      { id: "responses",    label: "Respuestas",      icon: "📋" },
      { id: "contactos",    label: "Contactos",       icon: "📞" },
      { id: "charts",       label: "Gráficas",        icon: "📊" },
      { id: "interviewers", label: "Entrevistadores",  icon: "📈" },
      { id: "municipios",   label: "Municipios",      icon: "🗺️" },
      { id: "ubicaciones",  label: "Ubicaciones GPS", icon: "📍" },
      { id: "users",        label: "Usuarios",        icon: "👥" },
      { id: "accesos",      label: "Accesos",         icon: "🔑" },
      { id: "database",     label: "Base de Datos",   icon: "🗄️" },
    ];

    return (
      <>
      {confirmModal && (
        <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 9999 }}>
          <div style={{ background: "#fff", borderRadius: "16px", padding: "2rem", maxWidth: "420px", width: "90%", boxShadow: "0 20px 60px rgba(0,0,0,0.3)", fontFamily: "'Georgia', serif" }}>
            <div style={{ fontSize: "32px", textAlign: "center", marginBottom: "1rem" }}>⚠️</div>
            <p style={{ color: "#0A1628", fontSize: "15px", textAlign: "center", lineHeight: "1.6", margin: "0 0 1.5rem" }}>{confirmModal.msg}</p>
            <div style={{ display: "flex", gap: "12px" }}>
              <button onClick={() => setConfirmModal(null)} style={{ flex: 1, padding: "11px", background: "transparent", border: "1px solid #ddd", borderRadius: "8px", cursor: "pointer", fontSize: "14px", color: "#555" }}>Cancelar</button>
              <button onClick={() => { confirmModal.onOk(); setConfirmModal(null); }} style={{ flex: 1, padding: "11px", background: "linear-gradient(135deg,#E53935,#C62828)", border: "none", borderRadius: "8px", cursor: "pointer", fontSize: "14px", color: "#fff", fontWeight: "bold" }}>Sí, eliminar</button>
            </div>
          </div>
        </div>
      )}
      {toastMsg && (
        <div style={{ position: "fixed", bottom: "28px", left: "50%", transform: "translateX(-50%)", background: "#0A1628", color: "#fff", padding: "12px 24px", borderRadius: "30px", fontSize: "14px", fontWeight: "bold", zIndex: 9999, boxShadow: "0 4px 20px rgba(0,0,0,0.3)", whiteSpace: "nowrap" }}>
          {toastMsg}
        </div>
      )}
      <div style={{ minHeight: "100vh", background: "#F4F6FA", fontFamily: "'Georgia', serif" }}>
        {/* Top bar */}
        <div style={{ background: "#0A1628", padding: "0.75rem 1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ background: "linear-gradient(135deg,#C8A951,#E8C96B)", borderRadius: "8px", padding: "4px 10px", fontWeight: "bold", color: "#0A1628", fontSize: "13px" }}>UAH</div>
            <span style={{ color: "rgba(255,255,255,0.85)", fontSize: "14px" }}>Panel de Administrador</span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <span style={{ color: "rgba(255,255,255,0.4)", fontSize: "12px" }}>🔒 {ADMIN_USER.name}</span>
            <button onClick={() => exportToExcel(db)} style={{ padding: "6px 16px", background: "linear-gradient(135deg,#22C987,#1AAF75)", border: "none", borderRadius: "6px", color: "#fff", cursor: "pointer", fontSize: "12px", fontWeight: "bold", boxShadow: "0 2px 8px rgba(34,201,135,0.4)" }}>📥 Exportar Excel</button>
            {/* ── indicador online/offline ── */}
            <div style={{ display: "flex", alignItems: "center", gap: "6px", padding: "4px 10px", background: isOnline ? "rgba(34,201,135,0.15)" : "rgba(229,57,53,0.2)", borderRadius: "20px", border: `1px solid ${isOnline ? "rgba(34,201,135,0.4)" : "rgba(229,57,53,0.4)"}` }}>
              <div style={{ width: "8px", height: "8px", borderRadius: "50%", background: isOnline ? "#22C987" : "#E53935" }} />
              <span style={{ color: isOnline ? "#22C987" : "#FF6B6B", fontSize: "11px", fontWeight: "bold" }}>{isOnline ? "En línea" : "Sin conexión"}</span>
            </div>
            {pendingCount > 0 && (
              <button onClick={syncPending} disabled={!isOnline || syncing}
                style={{ padding: "4px 10px", background: syncing ? "rgba(255,152,0,0.15)" : "rgba(255,152,0,0.2)", border: "1px solid rgba(255,152,0,0.4)", borderRadius: "20px", color: "#FF9800", cursor: isOnline && !syncing ? "pointer" : "default", fontSize: "11px", fontWeight: "bold" }}>
                {syncing ? "🔄 Sincronizando..." : `⚡ ${pendingCount} pendiente${pendingCount>1?"s":""}`}
              </button>
            )}
            {lastSync && !pendingCount && (
              <span style={{ color: "rgba(255,255,255,0.3)", fontSize: "10px" }}>✓ Sync {lastSync}</span>
            )}
            <button onClick={logout} style={{ padding: "6px 14px", background: "transparent", border: "1px solid rgba(255,255,255,0.15)", borderRadius: "6px", color: "rgba(255,255,255,0.6)", cursor: "pointer", fontSize: "12px" }}>Salir</button>
          </div>
        </div>

        {/* Stats bar */}
        <div style={{ background: "#162B50", padding: "1rem 1.5rem", display: "flex", gap: "2rem", flexWrap: "wrap" }}>
          {[
            { label: "Total Respuestas",     value: totalResponses, color: "#C8A951" },
            { label: "Usuarios Registrados", value: totalUsers,     color: "#C8A951" },
            { label: "Tasa de Respuesta",     value: `${responseRate}%`, color: "#22C987" },
            { label: "⚡ Apresuradas (<3min)", value: rushCount,    color: rushCount > 0 ? "#FF6B6B" : "#aaa", warn: rushCount > 0 },
            { label: "Última Respuesta",      value: db.responses.length > 0 ? db.responses[db.responses.length - 1].timeStr : "—", color: "#C8A951" },
          ].map((s, i) => (
            <div key={i} style={{ textAlign: "center" }}>
              <div style={{ color: s.color, fontSize: "22px", fontWeight: "bold" }}>{s.value}</div>
              <div style={{ color: "rgba(255,255,255,0.4)", fontSize: "11px" }}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{ borderBottom: "2px solid #E0E4EE", background: "#fff", padding: "0 1.5rem", display: "flex", overflowX: "auto" }}>
          {tabs.map((t) => (
            <button key={t.id} onClick={() => setAdminTab(t.id)}
              style={{ padding: "14px 18px", background: "transparent", border: "none", borderBottom: `2px solid ${adminTab === t.id ? "#C8A951" : "transparent"}`, color: adminTab === t.id ? "#C8A951" : "#888", cursor: "pointer", fontSize: "13px", fontWeight: adminTab === t.id ? "bold" : "normal", marginBottom: "-2px", whiteSpace: "nowrap" }}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>

        <div style={{ padding: "1.5rem", maxWidth: "1100px", margin: "0 auto" }}>

          {/* ── RESPONSES TAB ── */}
          {adminTab === "responses" && (() => {
            const filtered = responseSearch.trim()
              ? db.responses.filter(r =>
                  r.userName.toLowerCase().includes(responseSearch.toLowerCase()) ||
                  r.userId.toLowerCase().includes(responseSearch.toLowerCase()) ||
                  (r.municipio||"").toLowerCase().includes(responseSearch.toLowerCase())
                )
              : db.responses;
            const totalPages = Math.ceil(filtered.length / RESPONSES_PER_PAGE);
            const paginated  = filtered.slice(responsePage * RESPONSES_PER_PAGE, (responsePage + 1) * RESPONSES_PER_PAGE);
            return (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", flexWrap: "wrap", gap: "10px" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>Respuestas Registradas <span style={{ color: "#aaa", fontWeight: "normal", fontSize: "14px" }}>({filtered.length.toLocaleString()} de {db.responses.length.toLocaleString()})</span></h2>
                <input
                  value={responseSearch}
                  onChange={e => { setResponseSearch(e.target.value); setResponsePage(0); }}
                  placeholder="Buscar por nombre, ID o municipio…"
                  style={{ padding: "8px 14px", border: "1px solid #E0E4EE", borderRadius: "8px", fontSize: "13px", width: "260px", outline: "none" }}
                />
              </div>
              {filtered.length === 0 ? (
                <div style={{ textAlign: "center", padding: "3rem", color: "#aaa", background: "#fff", borderRadius: "12px", border: "1px solid #E0E4EE" }}>No hay respuestas registradas aún.</div>
              ) : (
                <>
                  <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                    {paginated.map((r) => (
                      <div key={r.id} style={{ background: "#fff", border: `1px solid ${r.isRush ? "#FFCDD2" : "#E0E4EE"}`, borderLeft: `4px solid ${r.isRush ? "#E53935" : "#C8A951"}`, borderRadius: "10px", padding: "1rem 1.25rem", display: "flex", justifyContent: "space-between", alignItems: "center", gap: "12px" }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px", display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                            {r.userName}
                            {r.isRush && (
                              <span style={{ background: "linear-gradient(135deg,#FF5252,#E53935)", color: "#fff", fontSize: "10px", fontWeight: "bold", padding: "2px 10px", borderRadius: "20px", letterSpacing: "0.05em" }}>⚡ APRESURADA</span>
                            )}
                          </div>
                          <div style={{ color: "#888", fontSize: "12px", marginTop: "2px" }}>
                            {r.userId} · {r.dateStr} · {r.timeStr}
                            {r.duration && <span style={{ marginLeft: "8px", color: r.isRush ? "#E53935" : "#22C987", fontWeight: "bold" }}>⏱ {r.duration}</span>}
                            <span style={{ marginLeft: "8px" }}>📍 {r.municipio} › {r.seccion}</span>
                            {r.encuestadoNombre && r.encuestadoNombre !== "Sin nombre" && (
                              <span style={{ marginLeft: "8px", color: "#C8A951" }}>👤 {r.encuestadoNombre}</span>
                            )}
                            {r.encuestadoTel && r.encuestadoTel !== "Sin teléfono" && (
                              <span style={{ marginLeft: "8px", color: "#4F8EF7" }}>📞 {r.encuestadoTel}</span>
                            )}
                          </div>
                        </div>
                        <div style={{ display: "flex", gap: "8px", flexShrink: 0 }}>
                          <button onClick={() => setSelectedResponse(r)} style={{ padding: "6px 14px", background: "#E8F0FE", border: "1px solid #C5D8FC", borderRadius: "6px", color: "#1A73E8", cursor: "pointer", fontSize: "12px" }}>Ver</button>
                          <button onClick={() => handleDeleteResponse(r.id)} style={{ padding: "6px 14px", background: "#FEE8E8", border: "1px solid #FFCDD2", borderRadius: "6px", color: "#C62828", cursor: "pointer", fontSize: "12px" }}>Eliminar</button>
                        </div>
                      </div>
                    ))}
                  </div>
                  {/* Paginación */}
                  {totalPages > 1 && (
                    <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "8px", marginTop: "1.5rem", flexWrap: "wrap" }}>
                      <button onClick={() => setResponsePage(0)} disabled={responsePage === 0} style={{ padding: "6px 12px", background: responsePage === 0 ? "#F0F0F0" : "#fff", border: "1px solid #E0E4EE", borderRadius: "6px", cursor: responsePage === 0 ? "default" : "pointer", color: "#555", fontSize: "12px" }}>« Primera</button>
                      <button onClick={() => setResponsePage(p => Math.max(0, p-1))} disabled={responsePage === 0} style={{ padding: "6px 12px", background: responsePage === 0 ? "#F0F0F0" : "#fff", border: "1px solid #E0E4EE", borderRadius: "6px", cursor: responsePage === 0 ? "default" : "pointer", color: "#555", fontSize: "12px" }}>‹ Ant</button>
                      <span style={{ fontSize: "13px", color: "#555", padding: "6px 12px", background: "#F8F9FC", borderRadius: "6px", border: "1px solid #E0E4EE" }}>Página {responsePage + 1} de {totalPages} · {(responsePage * RESPONSES_PER_PAGE + 1).toLocaleString()}–{Math.min((responsePage+1)*RESPONSES_PER_PAGE, filtered.length).toLocaleString()}</span>
                      <button onClick={() => setResponsePage(p => Math.min(totalPages-1, p+1))} disabled={responsePage >= totalPages-1} style={{ padding: "6px 12px", background: responsePage >= totalPages-1 ? "#F0F0F0" : "#fff", border: "1px solid #E0E4EE", borderRadius: "6px", cursor: responsePage >= totalPages-1 ? "default" : "pointer", color: "#555", fontSize: "12px" }}>Sig ›</button>
                      <button onClick={() => setResponsePage(totalPages-1)} disabled={responsePage >= totalPages-1} style={{ padding: "6px 12px", background: responsePage >= totalPages-1 ? "#F0F0F0" : "#fff", border: "1px solid #E0E4EE", borderRadius: "6px", cursor: responsePage >= totalPages-1 ? "default" : "pointer", color: "#555", fontSize: "12px" }}>Última »</button>
                    </div>
                  )}
                </>
              )}

              {selectedResponse && (
                <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
                  <div style={{ background: "#fff", borderRadius: "16px", padding: "2rem", maxWidth: "560px", width: "90%", maxHeight: "80vh", overflow: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
                      <div>
                        <h3 style={{ margin: "0 0 4px", color: "#0A1628" }}>Respuesta de {selectedResponse.userName}</h3>
                        {selectedResponse.isRush && <span style={{ background: "linear-gradient(135deg,#FF5252,#E53935)", color: "#fff", fontSize: "11px", padding: "2px 10px", borderRadius: "20px", fontWeight: "bold" }}>⚡ ENCUESTA APRESURADA (&lt;3 min)</span>}
                      </div>
                      <button onClick={() => setSelectedResponse(null)} style={{ background: "none", border: "none", fontSize: "22px", cursor: "pointer", color: "#aaa" }}>×</button>
                    </div>
                    <div style={{ marginBottom: "1rem", padding: "12px", background: "#F4F6FA", borderRadius: "8px", fontSize: "13px", color: "#555" }}>
                      <strong>ID Encuestador:</strong> {selectedResponse.userId} · <strong>Fecha:</strong> {selectedResponse.dateStr} · <strong>Hora:</strong> {selectedResponse.timeStr}
                      {selectedResponse.duration && <span> · <strong>⏱ Tiempo:</strong> <span style={{ color: selectedResponse.isRush ? "#C62828" : "#1B8040", fontWeight: "bold" }}>{selectedResponse.duration}</span></span>}
                      {selectedResponse.encuestadoNombre && selectedResponse.encuestadoNombre !== "Sin nombre" && (
                        <div style={{ marginTop: "8px" }}><strong>👤 Encuestado:</strong> {selectedResponse.encuestadoNombre}</div>
                      )}
                      {selectedResponse.encuestadoTel && selectedResponse.encuestadoTel !== "Sin teléfono" && (
                        <div style={{ marginTop: "4px" }}><strong>📞 Teléfono:</strong> <a href={`tel:${selectedResponse.encuestadoTel}`} style={{ color: "#1A73E8" }}>{selectedResponse.encuestadoTel}</a></div>
                      )}
                    </div>
                    {QUESTIONS.map((q) => (
                      <div key={q.id} style={{ marginBottom: "1rem", paddingBottom: "1rem", borderBottom: "1px solid #E0E4EE" }}>
                        <p style={{ fontWeight: "bold", color: "#0A1628", fontSize: "13px", margin: "0 0 6px" }}>{q.id}. {q.text}</p>
                        <p style={{ color: "#1A73E8", fontSize: "13px", margin: 0, padding: "8px 12px", background: "#E8F0FE", borderRadius: "6px" }}>✓ {selectedResponse.answers[q.id] || "Sin respuesta"}</p>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
            );
          })()}

          {/* ── CONTACTOS TAB ── */}
          {adminTab === "contactos" && (() => {
            const contactos = db.responses.filter(r => r.encuestadoNombre && r.encuestadoNombre !== "Sin nombre");
            return (
              <div>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.25rem", flexWrap: "wrap", gap: "10px" }}>
                  <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>
                    📞 Directorio de Contactos
                    <span style={{ color: "#aaa", fontWeight: "normal", fontSize: "14px", marginLeft: "8px" }}>({contactos.length} registros)</span>
                  </h2>
                  <button onClick={() => exportToExcel(db)} style={{ padding: "8px 18px", background: "linear-gradient(135deg,#22C987,#1AAF75)", border: "none", borderRadius: "8px", color: "#fff", cursor: "pointer", fontSize: "13px", fontWeight: "bold", boxShadow: "0 2px 8px rgba(34,201,135,0.35)" }}>
                    📥 Exportar Excel (Directorio + Gráficas)
                  </button>
                </div>
                {contactos.length === 0 ? (
                  <div style={{ textAlign: "center", padding: "3rem", color: "#aaa", background: "#fff", borderRadius: "12px", border: "1px solid #E0E4EE" }}>
                    Aún no hay contactos registrados.
                  </div>
                ) : (
                  <div style={{ background: "#fff", borderRadius: "12px", border: "1px solid #E0E4EE", overflow: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                      <thead>
                        <tr style={{ background: "#0A1628" }}>
                          {["#", "👤 Nombre", "📞 Teléfono / WhatsApp", "🗺️ Municipio", "📅 Fecha", "⏱ Tiempo"].map(h => (
                            <th key={h} style={{ padding: "12px 16px", color: "#C8A951", textAlign: "left", fontWeight: "bold", fontSize: "12px", letterSpacing: "0.05em", whiteSpace: "nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {contactos.map((r, i) => (
                          <tr key={r.id} style={{ borderBottom: "1px solid #F0F0F0", background: i % 2 === 0 ? "#fff" : "#FAFBFC" }}>
                            <td style={{ padding: "11px 16px", color: "#aaa", fontSize: "12px" }}>{i + 1}</td>
                            <td style={{ padding: "11px 16px", color: "#0A1628", fontWeight: "600" }}>{r.encuestadoNombre}</td>
                            <td style={{ padding: "11px 16px" }}>
                              {r.encuestadoTel && r.encuestadoTel !== "Sin teléfono"
                                ? <a href={`https://wa.me/52${r.encuestadoTel.replace(/\D/g,"")}`} target="_blank" rel="noopener noreferrer"
                                    style={{ color: "#25D366", fontWeight: "bold", textDecoration: "none" }}>
                                    📱 {r.encuestadoTel}
                                  </a>
                                : <span style={{ color: "#ccc", fontSize: "12px" }}>— sin teléfono</span>
                              }
                            </td>
                            <td style={{ padding: "11px 16px", color: "#555" }}>{r.municipio || "—"}</td>
                            <td style={{ padding: "11px 16px", color: "#555", fontSize: "12px", whiteSpace: "nowrap" }}>{r.dateStr}</td>
                            <td style={{ padding: "11px 16px" }}>
                              <span style={{ background: r.isRush ? "#FFF0F0" : "#F0FFF8", border: `1px solid ${r.isRush ? "#FFCDD2" : "#B2DFDB"}`, color: r.isRush ? "#C62828" : "#1B8040", fontSize: "11px", fontWeight: "bold", padding: "2px 8px", borderRadius: "20px", whiteSpace: "nowrap" }}>
                                {r.isRush ? "⚡ Apresurada" : "✓ Normal"} · {r.duration || "—"}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })()}

          {/* ── CHARTS TAB ── */}
          {adminTab === "charts" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>Análisis Gráfico de Resultados</h2>
                <div style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap" }}>
                  {rushCount > 0 && (
                    <div style={{ background: "linear-gradient(135deg,#FF5252,#E53935)", color: "#fff", padding: "6px 14px", borderRadius: "20px", fontSize: "12px", fontWeight: "bold" }}>
                      ⚡ {rushCount} encuesta{rushCount !== 1 ? "s" : ""} apresurada{rushCount !== 1 ? "s" : ""}
                    </div>
                  )}
                  <button onClick={() => exportToExcel(db)} style={{ padding: "8px 18px", background: "linear-gradient(135deg,#22C987,#1AAF75)", border: "none", borderRadius: "8px", color: "#fff", cursor: "pointer", fontSize: "13px", fontWeight: "bold", boxShadow: "0 2px 8px rgba(34,201,135,0.35)" }}>
                    📥 Exportar Excel (incluye hoja Analisis-Grafico)
                  </button>
                </div>
              </div>

              {db.responses.length === 0 ? (
                <div style={{ textAlign: "center", padding: "3rem", color: "#aaa", background: "#fff", borderRadius: "12px", border: "1px solid #E0E4EE" }}>Sin datos. Espera a que los usuarios respondan la encuesta.</div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(440px,1fr))", gap: "1.25rem" }}>
                  {QUESTIONS.map((q, qi) => {
                    const data  = getStats(q.id);
                    const total = data.reduce((a, b) => a + b, 0);
                    const rush  = rushStats[q.id] ?? [];
                    const hasRush = rush.some(v => v > 0);
                    return (
                      <div key={q.id} style={{ background: "#fff", border: `1px solid ${hasRush ? "#FFCDD2" : "#E0E4EE"}`, borderRadius: "16px", padding: "1.5rem", boxShadow: "0 2px 12px rgba(0,0,0,0.04)" }}>
                        {/* Header */}
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "12px" }}>
                          <div>
                            <div style={{ color: CHART_PALETTE[qi % CHART_PALETTE.length].bar, fontSize: "10px", letterSpacing: "0.12em", fontWeight: "bold", textTransform: "uppercase", marginBottom: "4px" }}>PREGUNTA {q.id}</div>
                            <p style={{ fontSize: "13px", color: "#0A1628", fontWeight: "bold", margin: 0, lineHeight: "1.5", maxWidth: "360px" }}>{q.text}</p>
                          </div>
                          {hasRush && (
                            <span style={{ background: "#FFF0F0", border: "1px solid #FFCDD2", color: "#C62828", fontSize: "10px", fontWeight: "bold", padding: "2px 8px", borderRadius: "20px", flexShrink: 0, marginLeft: "8px" }}>⚡ apr.</span>
                          )}
                        </div>

                        {/* Stats summary */}
                        <div style={{ display: "flex", gap: "12px", marginBottom: "16px", padding: "10px 14px", background: "#F8F9FC", borderRadius: "10px" }}>
                          <div style={{ textAlign: "center" }}>
                            <div style={{ fontSize: "18px", fontWeight: "bold", color: CHART_PALETTE[qi % CHART_PALETTE.length].bar }}>{total}</div>
                            <div style={{ fontSize: "10px", color: "#999" }}>respuestas</div>
                          </div>
                          {hasRush && (
                            <div style={{ textAlign: "center" }}>
                              <div style={{ fontSize: "18px", fontWeight: "bold", color: "#E53935" }}>{rush.reduce((a,b)=>a+b,0)}</div>
                              <div style={{ fontSize: "10px", color: "#999" }}>apresuradas</div>
                            </div>
                          )}
                          <div style={{ textAlign: "center", marginLeft: "auto" }}>
                            <div style={{ fontSize: "18px", fontWeight: "bold", color: "#0A1628" }}>{data.indexOf(Math.max(...data)) >= 0 ? Math.round((Math.max(...data)/Math.max(total,1))*100) : 0}%</div>
                            <div style={{ fontSize: "10px", color: "#999" }}>opción líder</div>
                          </div>
                        </div>

                        {/* Bar chart */}
                        <ModernBarChart data={data} labels={q.options} palette={Array(q.options.length).fill(CHART_PALETTE[qi % CHART_PALETTE.length])} total={total} rushData={rush} />

                        {/* Divider + Donut */}
                        {total > 0 && (
                          <>
                            <div style={{ height: "1px", background: "#F0F0F0", margin: "16px 0" }} />
                            <DonutChart data={data} labels={q.options} palette={Array(q.options.length).fill(0).map((_, i) => CHART_PALETTE[i % CHART_PALETTE.length])} />
                          </>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

          {/* ── INTERVIEWERS TAB ── */}
          {adminTab === "interviewers" && (() => {
            const now = new Date();
            const todayStr = now.toISOString().slice(0,10);
            const weekAgo  = new Date(now - 7 * 86400000);
            const monthAgo = new Date(now - 30 * 86400000);

            // Agrupar por fecha
            const byDate = {};
            db.responses.forEach(r => {
              const d = r.timestamp?.slice(0,10) || "Sin fecha";
              if (!byDate[d]) byDate[d] = { date: d, total: 0, rush: 0 };
              byDate[d].total++;
              if (r.isRush) byDate[d].rush++;
            });
            const dateRows = Object.values(byDate).sort((a,b) => b.date.localeCompare(a.date));

            // Agrupar por usuario
            const byUser = {};
            db.users.forEach((u) => { byUser[u.id] = { user: u, all: [], today: 0, week: 0, month: 0, rush: 0, totalSec: 0, countSec: 0 }; });
            db.responses.forEach((r) => {
              if (!byUser[r.userId]) byUser[r.userId] = { user: { id: r.userId, name: r.userName }, all: [], today: 0, week: 0, month: 0, rush: 0, totalSec: 0, countSec: 0 };
              const slot = byUser[r.userId];
              slot.all.push(r);
              const d = new Date(r.timestamp);
              if (r.timestamp?.slice(0,10) === todayStr) slot.today++;
              if (d >= weekAgo)  slot.week++;
              if (d >= monthAgo) slot.month++;
              if (r.isRush) slot.rush++;
              if (r.elapsedSec > 0) { slot.totalSec += r.elapsedSec; slot.countSec++; }
            });

            // Ordenar por EFICIENCIA (% encuestas en tiempo adecuado), no por cantidad
            const rows = Object.values(byUser).filter(s => s.all.length > 0);
            rows.sort((a, b) => {
              const effA = a.all.length > 0 ? (a.all.length - a.rush) / a.all.length : 0;
              const effB = b.all.length > 0 ? (b.all.length - b.rush) / b.all.length : 0;
              return effB - effA;
            });

            const hourBuckets = Array(24).fill(0);
            db.responses.forEach((r) => { if (r.hour !== undefined) hourBuckets[r.hour]++; });
            const maxHour = Math.max(...hourBuckets, 1);

            return (
              <div>
                <h2 style={{ color: "#0A1628", marginBottom: "1.5rem", fontSize: "18px" }}>Análisis de Entrevistadores</h2>
                {rows.length === 0 && dateRows.length === 0 ? (
                  <div style={{ textAlign: "center", padding: "3rem", color: "#aaa", background: "#fff", borderRadius: "12px" }}>Sin datos aún.</div>
                ) : (
                  <>
                    {/* ── ENTREVISTAS POR FECHA ── */}
                    <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", overflow: "hidden", marginBottom: "1.5rem" }}>
                      <div style={{ padding: "1rem 1.25rem", fontWeight: "bold", color: "#0A1628", fontSize: "14px", borderBottom: "1px solid #E0E4EE", background: "#FAFBFC" }}>
                        📅 Entrevistas por Fecha
                      </div>
                      {dateRows.length === 0 ? (
                        <div style={{ padding: "2rem", textAlign: "center", color: "#aaa", fontSize: "13px" }}>Sin datos.</div>
                      ) : (
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                          <thead>
                            <tr style={{ background: "#F4F6FA" }}>
                              {["Fecha", "Total encuestas", "✓ En tiempo", "⚡ Apresuradas", "% Eficiencia del día"].map(h => (
                                <th key={h} style={{ padding: "9px 14px", textAlign: h === "Fecha" ? "left" : "center", color: "#666", borderBottom: "1px solid #E0E4EE", fontSize: "12px", fontWeight: "bold" }}>{h}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            {dateRows.map((d, i) => {
                              const ok = d.total - d.rush;
                              const eff = d.total > 0 ? Math.round((ok / d.total) * 100) : 0;
                              const effColor = eff >= 80 ? "#1B8040" : eff >= 50 ? "#E65100" : "#C62828";
                              // Format date nicely
                              let dateLabel = d.date;
                              try { dateLabel = new Date(d.date + "T12:00:00").toLocaleDateString("es-MX", { weekday: "short", year: "numeric", month: "short", day: "numeric" }); } catch {}
                              return (
                                <tr key={d.date} style={{ borderBottom: "1px solid #F4F6FA", background: d.date === todayStr ? "#FFFBEF" : i % 2 === 0 ? "#fff" : "#FAFBFC" }}>
                                  <td style={{ padding: "9px 14px", color: "#0A1628", fontWeight: d.date === todayStr ? "bold" : "normal" }}>
                                    {dateLabel}{d.date === todayStr && <span style={{ marginLeft: "6px", background: "#C8A951", color: "#0A1628", fontSize: "10px", borderRadius: "8px", padding: "1px 6px", fontWeight: "bold" }}>HOY</span>}
                                  </td>
                                  <td style={{ padding: "9px 14px", textAlign: "center", fontWeight: "bold", color: "#0A1628" }}>{d.total}</td>
                                  <td style={{ padding: "9px 14px", textAlign: "center", color: "#1B8040", fontWeight: "bold" }}>{ok}</td>
                                  <td style={{ padding: "9px 14px", textAlign: "center", color: d.rush > 0 ? "#C62828" : "#aaa", fontWeight: d.rush > 0 ? "bold" : "normal" }}>{d.rush > 0 ? `⚡ ${d.rush}` : "—"}</td>
                                  <td style={{ padding: "9px 14px", textAlign: "center" }}>
                                    <span style={{ background: eff >= 80 ? "#F0FFF8" : eff >= 50 ? "#FFF3E0" : "#FFF0F0", border: `1px solid ${eff >= 80 ? "#B2DFDB" : eff >= 50 ? "#FFCC80" : "#FFCDD2"}`, color: effColor, fontWeight: "bold", fontSize: "12px", padding: "2px 10px", borderRadius: "20px" }}>{eff}%</span>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      )}
                    </div>

                    {/* ── RANKING POR EFICIENCIA ── */}
                    <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", padding: "1.25rem", marginBottom: "1.5rem" }}>
                      <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px", marginBottom: "4px" }}>🏆 Ranking de Eficiencia</div>
                      <div style={{ color: "#aaa", fontSize: "11px", marginBottom: "1rem" }}>Ordenado por % de encuestas realizadas en tiempo adecuado (≥ 3 min)</div>
                      {rows.map((s, i) => {
                        const ok     = s.all.length - s.rush;
                        const eff    = s.all.length > 0 ? Math.round((ok / s.all.length) * 100) : 0;
                        const avgSec = s.countSec > 0 ? Math.round(s.totalSec / s.countSec) : 0;
                        const avgStr = avgSec < 60 ? `${avgSec}s` : `${Math.floor(avgSec/60)}m ${avgSec%60}s`;
                        const effColor = eff >= 80 ? "#1B8040" : eff >= 50 ? "#E65100" : "#C62828";
                        const medals = ["🥇","🥈","🥉"];
                        return (
                          <div key={s.user.id} style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "10px" }}>
                            <div style={{ width: "24px", textAlign: "center", fontSize: i < 3 ? "16px" : "12px", color: "#888", flexShrink: 0 }}>{i < 3 ? medals[i] : i + 1}</div>
                            <div style={{ width: "140px", fontSize: "12px", color: "#333", flexShrink: 0 }}>{s.user.name}</div>
                            <div style={{ flex: 1, background: "#F4F6FA", borderRadius: "4px", height: "22px", overflow: "hidden" }}>
                              <div style={{ width: `${eff}%`, height: "100%", background: eff >= 80 ? "#22C987" : eff >= 50 ? "#F7844F" : "#FF6B6B", borderRadius: "4px", transition: "width 0.6s", display: "flex", alignItems: "center", paddingLeft: "6px" }}>
                                {eff > 15 && <span style={{ fontSize: "10px", color: "#fff", fontWeight: "bold" }}>{eff}%</span>}
                              </div>
                            </div>
                            <div style={{ fontSize: "12px", fontWeight: "bold", color: effColor, width: "36px", flexShrink: 0 }}>{eff}%</div>
                            <div style={{ fontSize: "11px", color: "#888", width: "55px", flexShrink: 0 }}>⏱ {avgStr}</div>
                            <div style={{ fontSize: "11px", color: "#0A1628", width: "40px", flexShrink: 0 }}>📋 {s.all.length}</div>
                          </div>
                        );
                      })}
                    </div>

                    {/* ── TABLA DETALLE POR PERIODO ── */}
                    <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", overflow: "hidden", marginBottom: "1.5rem" }}>
                      <div style={{ padding: "1rem 1.25rem", fontWeight: "bold", color: "#0A1628", fontSize: "14px", borderBottom: "1px solid #E0E4EE", background: "#FAFBFC" }}>📊 Detalle por Entrevistador y Período</div>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                        <thead>
                          <tr style={{ background: "#F4F6FA" }}>
                            {["Entrevistador","Total","Hoy","Esta Semana","Este Mes","✓ En tiempo","⚡ Apresuradas","% Eficiencia","Tiempo prom."].map(h => (
                              <th key={h} style={{ padding: "8px 12px", textAlign: h === "Entrevistador" ? "left" : "center", color: "#666", borderBottom: "1px solid #E0E4EE", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {rows.map((s) => {
                            const ok      = s.all.length - s.rush;
                            const eff     = s.all.length > 0 ? Math.round((ok / s.all.length) * 100) : 0;
                            const avgSec  = s.countSec > 0 ? Math.round(s.totalSec / s.countSec) : 0;
                            const avgStr  = avgSec < 60 ? `${avgSec}s` : `${Math.floor(avgSec/60)}m ${avgSec%60}s`;
                            const effColor = eff >= 80 ? "#1B8040" : eff >= 50 ? "#E65100" : "#C62828";
                            return (
                              <tr key={s.user.id} style={{ borderBottom: "1px solid #F4F6FA" }}>
                                <td style={{ padding: "8px 12px", color: "#0A1628" }}>{s.user.name}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center" }}>{s.all.length}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center" }}>{s.today}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center" }}>{s.week}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center" }}>{s.month}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center", color: "#1B8040", fontWeight: "bold" }}>{ok}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center", color: s.rush > 0 ? "#C62828" : "#aaa" }}>{s.rush > 0 ? `⚡ ${s.rush}` : "—"}</td>
                                <td style={{ padding: "8px 12px", textAlign: "center" }}>
                                  <span style={{ color: effColor, fontWeight: "bold" }}>{eff}%</span>
                                </td>
                                <td style={{ padding: "8px 12px", textAlign: "center", color: "#666" }}>{avgStr}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    {/* ── HORAS DEL DÍA ── */}
                    <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", padding: "1.25rem" }}>
                      <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px", marginBottom: "1rem" }}>🕐 Entrevistas por Hora del Día</div>
                      <div style={{ display: "flex", alignItems: "flex-end", gap: "3px", height: "110px" }}>
                        {hourBuckets.map((v, h) => (
                          <div key={h} style={{ flex: 1, display: "flex", flexDirection: "column", alignItems: "center", gap: "2px" }}>
                            <span style={{ fontSize: "9px", fontWeight: "bold", color: v > 0 ? CHART_PALETTE[h % CHART_PALETTE.length].bar : "transparent" }}>{v > 0 ? v : "·"}</span>
                            <div style={{ width: "100%", background: v > 0 ? CHART_PALETTE[h % CHART_PALETTE.length].bar : "#E0E4EE", height: `${(v / maxHour) * 70}px`, borderRadius: "3px 3px 0 0", minHeight: v > 0 ? "4px" : "2px", transition: "height 0.4s" }} />
                            <span style={{ fontSize: "8px", color: h % 6 === 0 ? "#888" : "transparent" }}>{h}h</span>
                          </div>
                        ))}
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", marginTop: "2px" }}>
                        {["Medianoche","6 AM","Mediodía","6 PM"].map(label => (
                          <span key={label} style={{ fontSize: "10px", color: "#bbb" }}>{label}</span>
                        ))}
                      </div>
                    </div>
                  </>
                )}
              </div>
            );
          })()}

          {/* ── MUNICIPIOS TAB ── */}
          {adminTab === "municipios" && (() => {
            const munKeys = Object.keys(db.municipios);
            const munList = ["Todos", ...munKeys];
            const secList = filterMunicipio === "Todos" ? ["Todas"] : ["Todas", ...(db.municipios[filterMunicipio] || [])];

            const filtered = db.responses.filter((r) => {
              if (filterMunicipio !== "Todos" && r.municipio !== filterMunicipio) return false;
              if (filterSeccion !== "Todas" && r.seccion !== filterSeccion) return false;
              return true;
            });

            const munStats = {};
            db.responses.forEach((r) => {
              const m = r.municipio || "Sin municipio";
              if (!munStats[m]) munStats[m] = { total: 0, rush: 0 };
              munStats[m].total++;
              if (r.isRush) munStats[m].rush++;
            });

            const secStats = {};
            filtered.forEach((r) => {
              const sec = r.seccion || "Sin sección";
              if (!secStats[sec]) secStats[sec] = { total: 0, rush: 0 };
              secStats[sec].total++;
              if (r.isRush) secStats[sec].rush++;
            });
            const secEntries = Object.entries(secStats).sort((a,b) => b[1].total - a[1].total);
            const maxSecTotal = secEntries.length > 0 ? secEntries[0][1].total : 1;

            const filtStats = QUESTIONS.reduce((acc, q) => {
              acc[q.id] = q.options.map(opt => filtered.filter(r => r.answers[q.id] === opt).length);
              return acc;
            }, {});

            return (
              <div>
                {/* ── CATALOG EDITOR ── */}
                <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "16px", padding: "1.5rem", marginBottom: "1.5rem" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
                    <h3 style={{ margin: 0, color: "#0A1628", fontSize: "15px" }}>✏️ Catálogo de Municipios y Secciones</h3>
                    <button onClick={() => setShowAddMun(!showAddMun)}
                      style={{ padding: "7px 16px", background: showAddMun ? "#F4F6FA" : "linear-gradient(135deg,#C8A951,#E8C96B)", border: showAddMun ? "1px solid #ddd" : "none", borderRadius: "8px", color: showAddMun ? "#555" : "#0A1628", fontWeight: "bold", cursor: "pointer", fontSize: "12px" }}>
                      {showAddMun ? "✕ Cancelar" : "+ Agregar Municipio"}
                    </button>
                  </div>

                  {/* Add new municipio form */}
                  {showAddMun && (
                    <div style={{ background: "#F8F9FC", border: "1px solid #C8A951", borderRadius: "10px", padding: "1rem", marginBottom: "1rem", display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "flex-end" }}>
                      <div style={{ flex: "1 1 160px" }}>
                        <label style={{ display: "block", fontSize: "11px", color: "#666", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Nombre del Municipio</label>
                        <input value={newMunName} onChange={(e) => setNewMunName(e.target.value)} placeholder="Ej: Mineral del Monte" style={adminInputSt} />
                      </div>
                      <div style={{ flex: "2 1 220px" }}>
                        <label style={{ display: "block", fontSize: "11px", color: "#666", marginBottom: "4px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Secciones iniciales (separadas por coma)</label>
                        <input value={newMunSecs} onChange={(e) => setNewMunSecs(e.target.value)} placeholder="Ej: Centro, Norte, Sur" style={adminInputSt} />
                      </div>
                      <button onClick={handleAddMun} style={{ padding: "9px 20px", background: "#22C987", border: "none", borderRadius: "8px", color: "#fff", fontWeight: "bold", cursor: "pointer", fontSize: "13px", flexShrink: 0 }}>Guardar</button>
                    </div>
                  )}

                  {/* Municipios list — each with inline secciones table */}
                  <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                    {munKeys.map((key) => {
                      const secciones = db.municipios[key] || [];
                      const secInput  = newSecByMun[key] || "";
                      return (
                        <div key={key} style={{ border: "1px solid #E0E4EE", borderRadius: "12px", overflow: "hidden" }}>
                          {/* Municipio header row */}
                          <div style={{ display: "flex", alignItems: "center", gap: "12px", background: "#F8F9FC", padding: "10px 14px", borderBottom: "1px solid #E0E4EE" }}>
                            <div style={{ flex: 1 }}>
                              <span style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px" }}>{key}</span>
                              <span style={{ marginLeft: "8px", background: "#fff", border: "1px solid #E0E4EE", borderRadius: "20px", padding: "1px 8px", fontSize: "11px", color: "#888" }}>{secciones.length} secciones · {munStats[key]?.total || 0} enc.</span>
                            </div>
                            <button onClick={() => setEditingMun({ key, newName: key, secciones: [...secciones] })}
                              style={{ padding: "5px 12px", background: "#E8F0FE", border: "1px solid #C5D8FC", borderRadius: "6px", color: "#1A73E8", cursor: "pointer", fontSize: "12px" }}>✏️ Renombrar</button>
                            <button onClick={() => handleDeleteMun(key)}
                              style={{ padding: "5px 12px", background: "#FEE8E8", border: "1px solid #FFCDD2", borderRadius: "6px", color: "#C62828", cursor: "pointer", fontSize: "12px" }}>🗑 Eliminar</button>
                          </div>

                          {/* Rename form if active */}
                          {editingMun?.key === key && (
                            <div style={{ background: "#FFF8E8", padding: "12px 14px", borderBottom: "1px solid #E0E4EE", display: "flex", gap: "8px", alignItems: "center" }}>
                              <input value={editingMun.newName} onChange={(e) => setEditingMun({ ...editingMun, newName: e.target.value })}
                                placeholder="Nuevo nombre…" style={{ ...adminInputSt, maxWidth: "260px" }} />
                              <button onClick={handleSaveMun} style={{ padding: "7px 16px", background: "linear-gradient(135deg,#C8A951,#E8C96B)", border: "none", borderRadius: "7px", color: "#0A1628", fontWeight: "bold", cursor: "pointer", fontSize: "12px" }}>Guardar</button>
                              <button onClick={() => setEditingMun(null)} style={{ padding: "7px 12px", background: "transparent", border: "1px solid #ddd", borderRadius: "7px", cursor: "pointer", fontSize: "12px", color: "#555" }}>Cancelar</button>
                            </div>
                          )}

                          {/* Secciones table */}
                          <div style={{ padding: "10px 14px" }}>
                            {secciones.length === 0 ? (
                              <div style={{ color: "#bbb", fontSize: "12px", fontStyle: "italic", padding: "4px 0" }}>Sin secciones registradas.</div>
                            ) : (
                              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", marginBottom: "8px" }}>
                                <thead>
                                  <tr style={{ background: "#F4F6FA" }}>
                                    <th style={{ padding: "6px 10px", textAlign: "left", color: "#666", fontWeight: "bold", borderBottom: "1px solid #E0E4EE" }}>Sección</th>
                                    <th style={{ padding: "6px 10px", textAlign: "center", color: "#666", fontWeight: "bold", borderBottom: "1px solid #E0E4EE" }}>Encuestas</th>
                                    <th style={{ padding: "6px 10px", textAlign: "center", color: "#666", fontWeight: "bold", borderBottom: "1px solid #E0E4EE" }}>Acción</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {secciones.map((sec, si) => {
                                    const secCount = db.responses.filter(r => r.municipio === key && r.seccion === sec).length;
                                    return (
                                      <tr key={si} style={{ borderBottom: "1px solid #F8F9FC" }}>
                                        <td style={{ padding: "6px 10px", color: "#333" }}>{sec}</td>
                                        <td style={{ padding: "6px 10px", textAlign: "center", color: "#888" }}>{secCount}</td>
                                        <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                          <button
                                            onClick={() => {
                                              const newSecs = secciones.filter((_, idx) => idx !== si);
                                              const munObj = { ...db.municipios, [key]: newSecs };
                                              saveDb({ ...db, municipios: munObj });
                                              showToast(`✅ Sección "${sec}" eliminada.`);
                                            }}
                                            style={{ padding: "3px 10px", background: "#FEE8E8", border: "1px solid #FFCDD2", borderRadius: "5px", color: "#C62828", cursor: "pointer", fontSize: "11px" }}>
                                            🗑 Borrar
                                          </button>
                                        </td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                            )}
                            {/* Add new section inline */}
                            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                              <input
                                value={secInput}
                                onChange={e => setNewSecByMun(prev => ({ ...prev, [key]: e.target.value }))}
                                placeholder="Nueva sección…"
                                style={{ ...adminInputSt, flex: 1, maxWidth: "220px", padding: "6px 10px" }}
                              />
                              <button
                                onClick={() => {
                                  const trimSec = secInput.trim();
                                  if (!trimSec) return;
                                  if (secciones.includes(trimSec)) { showToast("Esa sección ya existe."); return; }
                                  const munObj = { ...db.municipios, [key]: [...secciones, trimSec] };
                                  saveDb({ ...db, municipios: munObj });
                                  setNewSecByMun(prev => ({ ...prev, [key]: "" }));
                                  showToast(`✅ Sección "${trimSec}" agregada.`);
                                }}
                                style={{ padding: "6px 14px", background: "#22C987", border: "none", borderRadius: "7px", color: "#fff", fontWeight: "bold", cursor: "pointer", fontSize: "12px", flexShrink: 0 }}>
                                + Agregar
                              </button>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* ── STATS SECTION ── */}
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem", flexWrap: "wrap", gap: "10px" }}>
                  <h3 style={{ color: "#0A1628", margin: 0, fontSize: "15px" }}>📊 Estadísticas por Municipio</h3>
                  <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
                    <select value={filterMunicipio} onChange={(e) => { setFilterMunicipio(e.target.value); setFilterSeccion("Todas"); }}
                      style={{ padding: "7px 12px", border: "1px solid #E0E4EE", borderRadius: "8px", fontSize: "13px", background: "#fff" }}>
                      {munList.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                    <select value={filterSeccion} onChange={(e) => setFilterSeccion(e.target.value)}
                      style={{ padding: "7px 12px", border: "1px solid #E0E4EE", borderRadius: "8px", fontSize: "13px", background: "#fff" }}>
                      {secList.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: "10px", marginBottom: "1.5rem" }}>
                  {Object.entries(munStats).sort((a,b) => b[1].total - a[1].total).map(([m, st]) => (
                    <div key={m} onClick={() => { setFilterMunicipio(m); setFilterSeccion("Todas"); }}
                      style={{ background: filterMunicipio === m ? "#FFF8E8" : "#fff", border: `1.5px solid ${filterMunicipio === m ? "#C8A951" : "#E0E4EE"}`, borderRadius: "12px", padding: "1rem", cursor: "pointer", transition: "all 0.15s" }}>
                      <div style={{ fontWeight: "bold", fontSize: "13px", color: "#0A1628", marginBottom: "4px" }}>{m}</div>
                      <div style={{ fontSize: "24px", fontWeight: "bold", color: "#C8A951" }}>{st.total}</div>
                      <div style={{ fontSize: "11px", color: "#aaa" }}>encuestas{st.rush > 0 ? ` · ⚡${st.rush} apr.` : ""}</div>
                    </div>
                  ))}
                  {Object.keys(munStats).length === 0 && <div style={{ color: "#aaa", fontSize: "13px", gridColumn: "1/-1" }}>Sin datos aún.</div>}
                </div>

                {secEntries.length > 0 && (
                  <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", padding: "1.25rem", marginBottom: "1.5rem" }}>
                    <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px", marginBottom: "1rem" }}>
                      Secciones {filterMunicipio !== "Todos" ? `— ${filterMunicipio}` : "(todos los municipios)"}
                    </div>
                    {secEntries.map(([sec, st]) => (
                      <div key={sec} style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "8px" }}>
                        <div style={{ width: "120px", fontSize: "12px", color: "#333", flexShrink: 0, textAlign: "right" }}>{sec}</div>
                        <div style={{ flex: 1, background: "#F4F6FA", borderRadius: "4px", height: "22px", overflow: "hidden" }}>
                          <div style={{ width: `${(st.total / maxSecTotal) * 100}%`, height: "100%", background: "#4F8EF7", borderRadius: "4px", display: "flex", alignItems: "center", paddingLeft: "6px", transition: "width 0.6s" }}>
                            {st.total > 0 && <span style={{ fontSize: "10px", color: "#fff", fontWeight: "bold" }}>{st.total}</span>}
                          </div>
                        </div>
                        {st.rush > 0 && <span style={{ fontSize: "10px", color: "#C62828", fontWeight: "bold" }}>⚡{st.rush}</span>}
                      </div>
                    ))}
                  </div>
                )}

                {filtered.length > 0 && (
                  <div>
                    <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px", marginBottom: "1rem" }}>
                      Resultados — {filterMunicipio}{filterSeccion !== "Todas" ? ` › ${filterSeccion}` : ""}
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(380px,1fr))", gap: "1.25rem" }}>
                      {QUESTIONS.map((q, qi) => {
                        const data  = filtStats[q.id];
                        const total = data.reduce((a,b)=>a+b,0);
                        return (
                          <div key={q.id} style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", padding: "1.25rem" }}>
                            <div style={{ color: CHART_PALETTE[qi % CHART_PALETTE.length].bar, fontSize: "10px", letterSpacing: "0.1em", marginBottom: "6px", fontWeight: "bold" }}>PREGUNTA {q.id}</div>
                            <p style={{ fontSize: "12px", color: "#0A1628", fontWeight: "bold", marginBottom: "1rem", lineHeight: "1.5" }}>{q.text}</p>
                            <ModernBarChart data={data} labels={q.options} palette={Array(q.options.length).fill(CHART_PALETTE[qi % CHART_PALETTE.length])} total={total} />
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
            );
          })()}

          {/* ── USERS TAB ── */}
          {adminTab === "users" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1rem" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>Gestión de Usuarios ({db.users.length})</h2>
                <button onClick={() => setShowNewUserForm(!showNewUserForm)}
                  style={{ padding: "8px 18px", background: showNewUserForm ? "#F4F6FA" : "linear-gradient(135deg,#C8A951,#E8C96B)", border: showNewUserForm ? "1px solid #ddd" : "none", borderRadius: "8px", color: showNewUserForm ? "#555" : "#0A1628", fontWeight: "bold", cursor: "pointer", fontSize: "13px" }}>
                  {showNewUserForm ? "✕ Cancelar" : "+ Agregar Usuario"}
                </button>
              </div>

              {showNewUserForm && (
                <div style={{ background: "#fff", border: "1px solid #C8A951", borderRadius: "12px", padding: "1.25rem 1.5rem", marginBottom: "1rem", display: "flex", gap: "12px", flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div style={{ flex: "1 1 120px" }}>
                    <label style={{ display: "block", fontSize: "11px", color: "#666", marginBottom: "4px", textTransform: "uppercase" }}>ID de Acceso</label>
                    <input value={newUser.id} onChange={(e) => setNewUser({ ...newUser, id: e.target.value.toUpperCase() })} placeholder="Ej: UAH016" style={adminInputSt} />
                  </div>
                  <div style={{ flex: "2 1 180px" }}>
                    <label style={{ display: "block", fontSize: "11px", color: "#666", marginBottom: "4px", textTransform: "uppercase" }}>Nombre Completo</label>
                    <input value={newUser.name} onChange={(e) => setNewUser({ ...newUser, name: e.target.value })} placeholder="Nombre" style={adminInputSt} />
                  </div>
                  <div style={{ flex: "1 1 140px" }}>
                    <label style={{ display: "block", fontSize: "11px", color: "#666", marginBottom: "4px", textTransform: "uppercase" }}>Contraseña</label>
                    <input type="password" value={newUser.password} onChange={(e) => setNewUser({ ...newUser, password: e.target.value })} placeholder="Contraseña" style={adminInputSt} />
                  </div>
                  <button onClick={handleAddUser} style={{ padding: "9px 20px", background: "#22C987", border: "none", borderRadius: "8px", color: "#fff", fontWeight: "bold", cursor: "pointer", fontSize: "13px", flexShrink: 0 }}>Guardar</button>
                </div>
              )}

              <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", overflow: "hidden" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                  <thead>
                    <tr style={{ background: "#F4F6FA" }}>
                      {["ID","Nombre","Respondió","Fecha","Acciones"].map((h) => (
                        <th key={h} style={{ padding: "12px 16px", textAlign: "left", color: "#555", fontWeight: "bold", borderBottom: "1px solid #E0E4EE" }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {db.users.map((u) => {
                      const res = responsesByUser.get(u.id);
                      return (
                        <tr key={u.id} style={{ borderBottom: "1px solid #F4F6FA" }}>
                          <td style={{ padding: "10px 16px", fontFamily: "monospace", color: "#0A1628", fontWeight: "bold" }}>{u.id}</td>
                          <td style={{ padding: "10px 16px", color: "#333" }}>{u.name}</td>
                          <td style={{ padding: "10px 16px" }}>
                            <span style={{ padding: "3px 10px", borderRadius: "20px", fontSize: "11px", background: res ? "#E6F9F0" : "#F4F6FA", color: res ? "#1B8040" : "#888" }}>
                              {res ? "✓ Sí" : "Pendiente"}
                            </span>
                          </td>
                          <td style={{ padding: "10px 16px", color: "#888", fontSize: "12px" }}>
                            {res ? `${res.dateStr} ${res.timeStr}` : "—"}
                            {res?.duration && <span style={{ marginLeft: "6px", color: res.isRush ? "#C62828" : "#1B8040", fontWeight: "bold" }}>⏱ {res.duration}</span>}
                          </td>
                          <td style={{ padding: "10px 16px" }}>
                            <div style={{ display: "flex", gap: "6px" }}>
                              <button onClick={() => handleEditUser(u)} style={{ padding: "4px 10px", background: "#E8F0FE", border: "1px solid #C5D8FC", borderRadius: "5px", color: "#1A73E8", cursor: "pointer", fontSize: "11px" }}>Editar</button>
                              <button onClick={() => handleDeleteUser(u.id)} style={{ padding: "4px 10px", background: "#FEE8E8", border: "1px solid #FFCDD2", borderRadius: "5px", color: "#C62828", cursor: "pointer", fontSize: "11px" }}>Eliminar</button>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {editUser && (
                <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000 }}>
                  <div style={{ background: "#fff", borderRadius: "16px", padding: "2rem", width: "360px", boxShadow: "0 20px 60px rgba(0,0,0,0.3)" }}>
                    <h3 style={{ margin: "0 0 1.5rem", color: "#0A1628" }}>Editar Usuario</h3>
                    <div style={{ marginBottom: "1rem" }}>
                      <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "4px" }}>ID de Acceso</label>
                      <input value={editUser.id} onChange={(e) => setEditUser({ ...editUser, id: e.target.value.toUpperCase() })} style={adminInputSt} />
                      <span style={{ fontSize: "11px", color: "#aaa", marginTop: "3px", display: "block" }}>Las respuestas se actualizarán automáticamente.</span>
                    </div>
                    <div style={{ marginBottom: "1rem" }}>
                      <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "4px" }}>Nombre</label>
                      <input value={editUser.name} onChange={(e) => setEditUser({ ...editUser, name: e.target.value })} style={adminInputSt} />
                    </div>
                    <div style={{ marginBottom: "1rem" }}>
                      <label style={{ display: "block", fontSize: "12px", color: "#666", marginBottom: "4px" }}>Nueva contraseña</label>
                      <input type="password" value={editUser.password} onChange={(e) => setEditUser({ ...editUser, password: e.target.value })} style={adminInputSt} />
                    </div>
                    <div style={{ display: "flex", gap: "10px", marginTop: "1.5rem" }}>
                      <button onClick={() => setEditUser(null)} style={{ flex: 1, padding: "10px", background: "transparent", border: "1px solid #ddd", borderRadius: "8px", cursor: "pointer" }}>Cancelar</button>
                      <button onClick={handleSaveUser} style={{ flex: 1, padding: "10px", background: "linear-gradient(135deg,#C8A951,#E8C96B)", border: "none", borderRadius: "8px", color: "#0A1628", fontWeight: "bold", cursor: "pointer" }}>Guardar</button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ── UBICACIONES GPS TAB ── */}
          {adminTab === "ubicaciones" && (() => {
            const withGps    = db.responses.filter(r => r.gps);
            const withoutGps = db.responses.filter(r => !r.gps);
            return (
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "1.5rem", flexWrap: "wrap" }}>
                  <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>📍 Ubicaciones GPS de Encuestas</h2>
                  <span style={{ background: "#E8F0FE", border: "1px solid #C5D8FC", borderRadius: "20px", padding: "3px 12px", fontSize: "11px", color: "#1A73E8" }}>{withGps.length} con GPS · {withoutGps.length} sin GPS</span>
                </div>

                {/* Estadísticas rápidas */}
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: "12px", marginBottom: "1.5rem" }}>
                  {[
                    { icon: "📍", label: "Con GPS registrado",  value: withGps.length,     color: "#22C987" },
                    { icon: "📵", label: "Sin GPS / denegado",  value: withoutGps.length,  color: "#E53935" },
                    { icon: "📱", label: "Desde móvil",         value: db.responses.filter(r => r.device?.isMobile).length, color: "#4F8EF7" },
                    { icon: "💻", label: "Desde desktop",       value: db.responses.filter(r => r.device && !r.device.isMobile).length, color: "#C8A951" },
                  ].map(s => (
                    <div key={s.label} style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "12px", padding: "1rem", textAlign: "center" }}>
                      <div style={{ fontSize: "22px" }}>{s.icon}</div>
                      <div style={{ fontSize: "24px", fontWeight: "bold", color: s.color, margin: "4px 0" }}>{s.value}</div>
                      <div style={{ fontSize: "11px", color: "#888" }}>{s.label}</div>
                    </div>
                  ))}
                </div>

                {/* Link a Google Maps con todas las coordenadas */}
                {withGps.length > 0 && (
                  <div style={{ background: "#E8F5E9", border: "1px solid #A5D6A7", borderRadius: "12px", padding: "1rem 1.25rem", marginBottom: "1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px" }}>
                    <div>
                      <div style={{ fontWeight: "bold", color: "#1B5E20", fontSize: "14px" }}>🗺️ Ver en Google Maps</div>
                      <div style={{ color: "#388E3C", fontSize: "12px", marginTop: "2px" }}>Abre un mapa con todas las ubicaciones registradas</div>
                    </div>
                    <a href={`https://www.google.com/maps/search/?api=1&query=${withGps[0].gps.lat},${withGps[0].gps.lng}`} target="_blank" rel="noreferrer"
                      style={{ padding: "8px 16px", background: "#4CAF50", border: "none", borderRadius: "8px", color: "#fff", fontSize: "13px", fontWeight: "bold", cursor: "pointer", textDecoration: "none" }}>
                      Abrir mapa →
                    </a>
                  </div>
                )}

                {/* Tabla de registros con GPS */}
                <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "16px", overflow: "hidden", marginBottom: "1.5rem" }}>
                  <div style={{ background: "#F4F6FA", padding: "12px 18px", borderBottom: "1px solid #E0E4EE", fontWeight: "bold", color: "#0A1628", fontSize: "14px" }}>
                    Registro de accesos con ubicación
                  </div>
                  {db.responses.length === 0 ? (
                    <div style={{ textAlign: "center", padding: "2rem", color: "#aaa", fontSize: "13px" }}>Sin encuestas registradas aún.</div>
                  ) : (
                    <div style={{ overflowX: "auto", maxHeight: "520px", overflowY: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                        <thead style={{ position: "sticky", top: 0, background: "#F8F9FC", zIndex: 1 }}>
                          <tr>
                            {["#","Encuestador","Fecha","Hora","Municipio","GPS","Precisión","Dispositivo","Navegador","SO","Ver mapa"].map(h => (
                              <th key={h} style={{ padding: "9px 12px", textAlign: "left", color: "#555", fontWeight: "bold", borderBottom: "1px solid #E0E4EE", whiteSpace: "nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {db.responses.map((r, i) => (
                            <tr key={r.id} style={{ borderBottom: "1px solid #F4F6FA", background: i % 2 === 0 ? "#fff" : "#FAFBFD" }}>
                              <td style={{ padding: "8px 12px", color: "#bbb", fontSize: "11px" }}>{i+1}</td>
                              <td style={{ padding: "8px 12px", color: "#0A1628", fontWeight: "500" }}>{r.userName}</td>
                              <td style={{ padding: "8px 12px", color: "#555", whiteSpace: "nowrap" }}>{r.dateStr}</td>
                              <td style={{ padding: "8px 12px", color: "#555", whiteSpace: "nowrap" }}>{r.timeStr}</td>
                              <td style={{ padding: "8px 12px", color: "#555" }}>{r.municipio}</td>
                              <td style={{ padding: "8px 12px" }}>
                                {r.gps
                                  ? <span style={{ color: "#1B8040", fontFamily: "monospace", fontSize: "11px" }}>{r.gps.lat.toFixed(5)}, {r.gps.lng.toFixed(5)}</span>
                                  : <span style={{ color: "#bbb", fontSize: "11px" }}>Sin GPS</span>}
                              </td>
                              <td style={{ padding: "8px 12px", color: "#888", fontSize: "11px" }}>
                                {r.gps ? `±${r.gps.accuracy}m` : "—"}
                              </td>
                              <td style={{ padding: "8px 12px", color: "#555", fontSize: "11px" }}>
                                {r.device ? (r.device.isMobile ? "📱 Móvil" : "💻 Desktop") : "—"}
                              </td>
                              <td style={{ padding: "8px 12px", color: "#555", fontSize: "11px" }}>{r.device?.browser || "—"}</td>
                              <td style={{ padding: "8px 12px", color: "#555", fontSize: "11px" }}>{r.device?.platform || "—"}</td>
                              <td style={{ padding: "8px 12px" }}>
                                {r.gps
                                  ? <a href={`https://www.google.com/maps?q=${r.gps.lat},${r.gps.lng}`} target="_blank" rel="noreferrer"
                                      style={{ color: "#1A73E8", fontSize: "11px", textDecoration: "none" }}>🗺️ Ver</a>
                                  : <span style={{ color: "#ccc", fontSize: "11px" }}>—</span>}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                {/* Dispositivos usados */}
                {db.responses.some(r => r.device) && (
                  <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "16px", overflow: "hidden" }}>
                    <div style={{ background: "#F4F6FA", padding: "12px 18px", borderBottom: "1px solid #E0E4EE", fontWeight: "bold", color: "#0A1628", fontSize: "14px" }}>
                      📱 Dispositivos y Navegadores Utilizados
                    </div>
                    <div style={{ padding: "1rem 1.25rem", display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: "1rem" }}>
                      {/* SO */}
                      <div>
                        <div style={{ fontSize: "12px", color: "#888", fontWeight: "bold", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Sistema Operativo</div>
                        {Object.entries(db.responses.filter(r=>r.device).reduce((acc,r)=>{ acc[r.device.platform]=(acc[r.device.platform]||0)+1; return acc; },{}))
                          .sort((a,b)=>b[1]-a[1]).map(([k,v]) => (
                          <div key={k} style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px", fontSize:"13px" }}>
                            <span style={{ color:"#333" }}>{k}</span>
                            <span style={{ fontWeight:"bold", color:"#0A1628" }}>{v}</span>
                          </div>
                        ))}
                      </div>
                      {/* Navegador */}
                      <div>
                        <div style={{ fontSize: "12px", color: "#888", fontWeight: "bold", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Navegador</div>
                        {Object.entries(db.responses.filter(r=>r.device).reduce((acc,r)=>{ acc[r.device.browser]=(acc[r.device.browser]||0)+1; return acc; },{}))
                          .sort((a,b)=>b[1]-a[1]).map(([k,v]) => (
                          <div key={k} style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px", fontSize:"13px" }}>
                            <span style={{ color:"#333" }}>{k}</span>
                            <span style={{ fontWeight:"bold", color:"#0A1628" }}>{v}</span>
                          </div>
                        ))}
                      </div>
                      {/* Idioma */}
                      <div>
                        <div style={{ fontSize: "12px", color: "#888", fontWeight: "bold", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" }}>Idioma del dispositivo</div>
                        {Object.entries(db.responses.filter(r=>r.device).reduce((acc,r)=>{ acc[r.device.lang]=(acc[r.device.lang]||0)+1; return acc; },{}))
                          .sort((a,b)=>b[1]-a[1]).map(([k,v]) => (
                          <div key={k} style={{ display:"flex", justifyContent:"space-between", marginBottom:"6px", fontSize:"13px" }}>
                            <span style={{ color:"#333" }}>{k}</span>
                            <span style={{ fontWeight:"bold", color:"#0A1628" }}>{v}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })()}

          {/* ── ACCESOS TAB ── */}
          {adminTab === "accesos" && (
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "1.5rem" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>🔑 Panel de Accesos y Contraseñas</h2>
                <span style={{ background: "#FEF9EC", border: "1px solid #C8A951", borderRadius: "20px", padding: "3px 12px", fontSize: "11px", color: "#8B6914" }}>🔒 Solo visible para Admin</span>
              </div>

              {/* ── CREDENCIALES DE ROLES ESPECIALES ── */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(320px,1fr))", gap: "1rem", marginBottom: "1.5rem" }}>

                {/* ADMIN */}
                <div style={{ background: "linear-gradient(135deg,#0A1628,#162B50)", borderRadius: "16px", padding: "1.5rem", border: "2px solid #C8A951" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "1rem" }}>
                    <div style={{ background: "linear-gradient(135deg,#C8A951,#E8C96B)", borderRadius: "8px", padding: "6px 12px", fontWeight: "bold", color: "#0A1628", fontSize: "12px" }}>ADMIN</div>
                    <span style={{ color: "#C8A951", fontWeight: "bold", fontSize: "15px" }}>Administrador UAH</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                    <div style={{ background: "rgba(255,255,255,0.06)", borderRadius: "8px", padding: "10px 14px" }}>
                      <div style={{ color: "rgba(255,255,255,0.45)", fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "4px" }}>ID de Acceso</div>
                      <div style={{ color: "#fff", fontFamily: "monospace", fontSize: "15px", fontWeight: "bold", letterSpacing: "0.05em" }}>{ADMIN_USER.id}</div>
                    </div>
                    <div style={{ background: "rgba(255,255,255,0.06)", borderRadius: "8px", padding: "10px 14px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                        <div style={{ color: "rgba(255,255,255,0.45)", fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em" }}>Contraseña</div>
                        <button onClick={() => setShowAdminPass(v => !v)} style={{ background: "rgba(200,169,81,0.2)", border: "1px solid rgba(200,169,81,0.4)", borderRadius: "4px", color: "#C8A951", cursor: "pointer", fontSize: "11px", padding: "2px 8px" }}>{showAdminPass ? "🙈 Ocultar" : "👁 Ver"}</button>
                      </div>
                      <div style={{ color: "#C8A951", fontFamily: "monospace", fontSize: "15px", fontWeight: "bold", letterSpacing: "0.08em" }}>
                        {showAdminPass ? ADMIN_USER.password : "•".repeat(ADMIN_USER.password.length)}
                      </div>
                    </div>
                  </div>
                  <div style={{ marginTop: "10px", padding: "6px 10px", background: "rgba(200,169,81,0.1)", borderRadius: "6px", fontSize: "11px", color: "rgba(200,169,81,0.7)" }}>
                    ✅ Permisos: Lectura · Escritura · Eliminación · Exportación · Gestión completa
                  </div>
                </div>

                {/* SUPERVISOR */}
                <div style={{ background: "linear-gradient(135deg,#162B50,#1E3A6E)", borderRadius: "16px", padding: "1.5rem", border: "2px solid #4F8EF7" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "1rem" }}>
                    <div style={{ background: "linear-gradient(135deg,#4F8EF7,#6FA8FF)", borderRadius: "8px", padding: "6px 12px", fontWeight: "bold", color: "#fff", fontSize: "12px" }}>SUPER</div>
                    <span style={{ color: "#6FA8FF", fontWeight: "bold", fontSize: "15px" }}>Supervisor UAH</span>
                  </div>
                  <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                    <div style={{ background: "rgba(255,255,255,0.06)", borderRadius: "8px", padding: "10px 14px" }}>
                      <div style={{ color: "rgba(255,255,255,0.45)", fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: "4px" }}>ID de Acceso</div>
                      <div style={{ color: "#fff", fontFamily: "monospace", fontSize: "15px", fontWeight: "bold", letterSpacing: "0.05em" }}>{SUPERVISOR_USER.id}</div>
                    </div>
                    <div style={{ background: "rgba(255,255,255,0.06)", borderRadius: "8px", padding: "10px 14px" }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                        <div style={{ color: "rgba(255,255,255,0.45)", fontSize: "10px", textTransform: "uppercase", letterSpacing: "0.1em" }}>Contraseña</div>
                        <button onClick={() => setShowSuperPass(v => !v)} style={{ background: "rgba(79,142,247,0.2)", border: "1px solid rgba(79,142,247,0.4)", borderRadius: "4px", color: "#6FA8FF", cursor: "pointer", fontSize: "11px", padding: "2px 8px" }}>{showSuperPass ? "🙈 Ocultar" : "👁 Ver"}</button>
                      </div>
                      <div style={{ color: "#6FA8FF", fontFamily: "monospace", fontSize: "15px", fontWeight: "bold", letterSpacing: "0.08em" }}>
                        {showSuperPass ? SUPERVISOR_USER.password : "•".repeat(SUPERVISOR_USER.password.length)}
                      </div>
                    </div>
                  </div>
                  <div style={{ marginTop: "10px", padding: "6px 10px", background: "rgba(79,142,247,0.1)", borderRadius: "6px", fontSize: "11px", color: "rgba(111,168,255,0.7)" }}>
                    👁 Permisos: Solo lectura de gráficas con porcentajes
                  </div>
                </div>
              </div>

              {/* ── TABLA DE ENCUESTADORES ── */}
              <div style={{ background: "#fff", border: "1px solid #E0E4EE", borderRadius: "16px", overflow: "hidden" }}>
                <div style={{ background: "#F4F6FA", padding: "14px 20px", borderBottom: "1px solid #E0E4EE", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px" }}>👥 Contraseñas de Encuestadores ({db.users.length})</div>
                  <span style={{ fontSize: "11px", color: "#aaa" }}>Contraseñas visibles solo para el administrador</span>
                </div>
                <div style={{ overflowX: "auto", maxHeight: "500px", overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
                    <thead style={{ position: "sticky", top: 0, background: "#F8F9FC", zIndex: 1 }}>
                      <tr>
                        {["#","ID","Nombre","Contraseña","Estado"].map(h => (
                          <th key={h} style={{ padding: "10px 16px", textAlign: "left", color: "#555", fontWeight: "bold", borderBottom: "1px solid #E0E4EE", whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {db.users.map((u, i) => {
                        const respondio = db.responses.some(r => r.userId === u.id);
                        return (
                          <tr key={u.id} style={{ borderBottom: "1px solid #F4F6FA", background: i % 2 === 0 ? "#fff" : "#FAFBFD" }}>
                            <td style={{ padding: "9px 16px", color: "#bbb", fontSize: "11px" }}>{i + 1}</td>
                            <td style={{ padding: "9px 16px", fontFamily: "monospace", color: "#0A1628", fontWeight: "bold", fontSize: "12px" }}>{u.id}</td>
                            <td style={{ padding: "9px 16px", color: "#333" }}>{u.name}</td>
                            <td style={{ padding: "9px 16px", fontFamily: "monospace", color: "#C8A951", fontWeight: "bold", letterSpacing: "0.05em", fontSize: "13px" }}>{u.password}</td>
                            <td style={{ padding: "9px 16px" }}>
                              <span style={{ padding: "3px 10px", borderRadius: "20px", fontSize: "11px", background: respondio ? "#E6F9F0" : "#F4F6FA", color: respondio ? "#1B8040" : "#888" }}>
                                {respondio ? "✓ Respondió" : "Pendiente"}
                              </span>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── DATABASE TAB ── */}
          {adminTab === "database" && (
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: "12px", marginBottom: "1.5rem" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>Base de Datos</h2>
                <span style={{ background: "#FEE8E8", border: "1px solid #FFCDD2", borderRadius: "20px", padding: "3px 12px", fontSize: "11px", color: "#C62828" }}>🔒 Acceso Restringido</span>
              </div>

              {/* ── PANEL SINCRONIZACIÓN OFFLINE ── */}
              <div style={{ background: isOnline ? "linear-gradient(135deg,#E6FAF2,#F0FDF8)" : "linear-gradient(135deg,#FFF3F3,#FFF8F8)", border: `1px solid ${isOnline ? "#B2DFDB" : "#FFCDD2"}`, borderRadius: "14px", padding: "1.25rem 1.5rem", marginBottom: "1.5rem" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "12px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <div style={{ width: "12px", height: "12px", borderRadius: "50%", background: isOnline ? "#22C987" : "#E53935", boxShadow: `0 0 0 3px ${isOnline ? "rgba(34,201,135,0.2)" : "rgba(229,57,53,0.2)"}` }} />
                    <div>
                      <div style={{ fontWeight: "bold", color: "#0A1628", fontSize: "14px" }}>
                        {isOnline ? "🌐 Conectado a internet" : "📴 Sin conexión a internet"}
                      </div>
                      <div style={{ fontSize: "12px", color: "#888", marginTop: "2px" }}>
                        {isOnline
                          ? pendingCount > 0
                            ? `Hay ${pendingCount} encuesta${pendingCount>1?"s":""} guardada${pendingCount>1?"s":""} localmente sin sincronizar`
                            : lastSync ? `Última sincronización: ${lastSync}` : "Todas las encuestas están sincronizadas"
                          : "Las encuestas se guardan en este dispositivo hasta recuperar conexión"}
                      </div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: "10px", flexWrap: "wrap" }}>
                    {pendingCount > 0 && isOnline && (
                      <button onClick={syncPending} disabled={syncing}
                        style={{ padding: "8px 18px", background: syncing ? "#aaa" : "linear-gradient(135deg,#22C987,#1AAF75)", border: "none", borderRadius: "8px", color: "#fff", cursor: syncing ? "default" : "pointer", fontWeight: "bold", fontSize: "13px" }}>
                        {syncing ? "🔄 Sincronizando..." : `🔄 Sincronizar ${pendingCount} encuesta${pendingCount>1?"s":""}`}
                      </button>
                    )}
                    <button onClick={async () => { const p = await idbGetAllPending(); setPendingCount(p.length); showToast(`📊 ${p.length} encuesta${p.length!==1?"s":""} pendiente${p.length!==1?"s":""}`); }}
                      style={{ padding: "8px 14px", background: "#E8F0FE", border: "1px solid #C5D8FC", borderRadius: "8px", color: "#1A73E8", cursor: "pointer", fontSize: "12px" }}>
                      🔍 Revisar cola offline
                    </button>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(130px,1fr))", gap: "10px", marginTop: "1rem" }}>
                  {[
                    { label: "En base de datos", value: db.responses.length, color: "#22C987" },
                    { label: "Pendientes sync", value: pendingCount, color: pendingCount > 0 ? "#FF9800" : "#22C987" },
                    { label: "Total encuestadores", value: db.users.length, color: "#C8A951" },
                    { label: "Municipios", value: Object.keys(db.municipios).length, color: "#4F8EF7" },
                  ].map(stat => (
                    <div key={stat.label} style={{ background: "#fff", borderRadius: "10px", padding: "10px 14px", border: "1px solid #E0E4EE", textAlign: "center" }}>
                      <div style={{ fontSize: "22px", fontWeight: "bold", color: stat.color }}>{stat.value}</div>
                      <div style={{ fontSize: "11px", color: "#888", marginTop: "2px" }}>{stat.label}</div>
                    </div>
                  ))}
                </div>
              </div>
              <div style={{ background: "#0A1628", borderRadius: "12px", padding: "1.5rem", fontFamily: "monospace", fontSize: "12px", color: "#C8A951", overflowX: "auto", maxHeight: "600px", overflowY: "auto" }}>
                <div style={{ color: "rgba(200,169,81,0.4)", marginBottom: "1rem" }}>// Base de Datos — UAH Encuesta Diagnóstico Educativo</div>
                <div style={{ color: "rgba(200,169,81,0.4)", marginBottom: "0.5rem" }}>// Generado: {dbGeneratedAt}</div>
                <br />
                <div style={{ color: "#22C987" }}>{"{"}</div>
                <div style={{ paddingLeft: "1.5rem" }}>
                  <div style={{ color: "#fff", marginBottom: "0.5rem" }}>"metadata": {"{"}<div style={{ paddingLeft: "1.5rem", color: "#aaa" }}><div>"totalUsuarios": {db.users.length},</div><div>"totalRespuestas": {db.responses.length},</div><div>"tasaRespuesta": "{responseRate}%"</div></div>{"},"}</div>
                  <div style={{ color: "#fff", marginBottom: "0.5rem" }}>"usuarios": [</div>
                  <div style={{ paddingLeft: "1.5rem" }}>
                    {db.users.map((u, i) => (
                      <div key={u.id} style={{ marginBottom: "4px", color: "#aaa" }}>
                        {"{"}
                        "id":"{u.id}", "nombre":"{u.name}", "respondio":{db.responses.some((r) => r.userId === u.id) ? "true" : "false"}
                        {"}"}{i < db.users.length - 1 ? "," : ""}
                      </div>
                    ))}
                  </div>
                  <div style={{ color: "#fff", marginBottom: "0.5rem" }}>],</div>
                  <div style={{ color: "#fff", marginBottom: "0.5rem" }}>"respuestas": [</div>
                  <div style={{ paddingLeft: "1.5rem" }}>
                    {db.responses.length === 0 && <div style={{ color: "#555" }}>// Sin respuestas aún</div>}
                    {db.responses.map((r, i) => (
                      <div key={r.id} style={{ marginBottom: "8px" }}>
                        <div style={{ color: "#aaa" }}>{"{"}</div>
                        <div style={{ paddingLeft: "1.5rem", color: "#888" }}>
                          <div>"id": "{r.id}",</div>
                          <div>"usuario": "{r.userId} — {r.userName}",</div>
                          <div>"timestamp": "{r.timestamp}",</div>
                          <div>"duracion": "{r.duration || "—"}",</div>
                          <div style={{ color: r.isRush ? "#FF6B6B" : "#888" }}>"apresurada": {r.isRush ? "true" : "false"},</div>
                          <div>"municipio": "{r.municipio}",</div>
                          <div>"respuestas": {"{"}</div>
                          <div style={{ paddingLeft: "1.5rem" }}>
                            {QUESTIONS.map((q) => (
                              <div key={q.id}>"P{q.id}": "{r.answers[q.id] || "—"}"{q.id < QUESTIONS.length ? "," : ""}</div>
                            ))}
                          </div>
                          <div>{"}"}</div>
                        </div>
                        <div style={{ color: "#aaa" }}>{"}"}{ i < db.responses.length - 1 ? "," : ""}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ color: "#fff" }}>]</div>
                </div>
                <div style={{ color: "#22C987" }}>{"}"}</div>
              </div>
            </div>
          )}
        </div>
      </div>
      </>
    );
  }

  // ─────────────────────────────────────────────────────────────────────
  // SUPERVISOR — solo lectura de gráficas
  // ─────────────────────────────────────────────────────────────────────
  if (screen === "supervisor") {
    const total = db.responses.length;
    const supTabs = [
      { id: "charts", label: "Gráficas", icon: "📊" },
    ];

    return (
      <div style={{ minHeight: "100vh", background: "#F4F6FA", fontFamily: "'Georgia', serif" }}>
        {/* Top bar */}
        <div style={{ background: "#162B50", padding: "0.75rem 1.5rem", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <div style={{ background: "linear-gradient(135deg,#C8A951,#E8C96B)", borderRadius: "8px", padding: "4px 10px", fontWeight: "bold", color: "#0A1628", fontSize: "13px" }}>UAH</div>
            <span style={{ color: "rgba(255,255,255,0.85)", fontSize: "14px" }}>Panel Supervisor — Solo Lectura</span>
            <span style={{ background: "rgba(34,201,135,0.2)", border: "1px solid rgba(34,201,135,0.4)", borderRadius: "20px", padding: "2px 10px", fontSize: "11px", color: "#22C987" }}>👁 Vista supervisión</span>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
            <span style={{ color: "rgba(255,255,255,0.4)", fontSize: "12px" }}>{SUPERVISOR_USER.name}</span>
            <button onClick={logout} style={{ padding: "6px 14px", background: "transparent", border: "1px solid rgba(255,255,255,0.15)", borderRadius: "6px", color: "rgba(255,255,255,0.6)", cursor: "pointer", fontSize: "12px" }}>Salir</button>
          </div>
        </div>

        {/* Stats bar */}
        <div style={{ background: "#0A1628", padding: "1rem 1.5rem", display: "flex", gap: "2rem", flexWrap: "wrap" }}>
          {[
            { label: "Total Encuestas",      value: total,                    color: "#C8A951" },
            { label: "Municipios activos",    value: Object.keys(supMunStats).length, color: "#4F8EF7" },
            { label: "⚡ Apresuradas (<3min)", value: db.responses.filter(r=>r.isRush).length, color: db.responses.filter(r=>r.isRush).length > 0 ? "#FF6B6B" : "#aaa" },
            { label: "Tasa de participación", value: db.users.length > 0 ? `${Math.round((total/db.users.length)*100)}%` : "—", color: "#22C987" },
          ].map((s, i) => (
            <div key={i} style={{ textAlign: "center" }}>
              <div style={{ color: s.color, fontSize: "22px", fontWeight: "bold" }}>{s.value}</div>
              <div style={{ color: "rgba(255,255,255,0.4)", fontSize: "11px" }}>{s.label}</div>
            </div>
          ))}
        </div>

        {/* Tabs */}
        <div style={{ borderBottom: "2px solid #E0E4EE", background: "#fff", padding: "0 1.5rem", display: "flex" }}>
          {supTabs.map(t => (
            <button key={t.id} onClick={() => setSupTab(t.id)}
              style={{ padding: "14px 18px", background: "transparent", border: "none", borderBottom: `2px solid ${supTab === t.id ? "#C8A951" : "transparent"}`, color: supTab === t.id ? "#C8A951" : "#888", cursor: "pointer", fontSize: "13px", fontWeight: supTab === t.id ? "bold" : "normal", marginBottom: "-2px" }}>
              {t.icon} {t.label}
            </button>
          ))}
        </div>

        <div style={{ padding: "1.5rem", maxWidth: "1100px", margin: "0 auto" }}>

          {/* ── GRÁFICAS ── */}
          {supTab === "charts" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "1.5rem" }}>
                <h2 style={{ color: "#0A1628", margin: 0, fontSize: "18px" }}>Resultados de la Encuesta</h2>
                {db.responses.filter(r=>r.isRush).length > 0 && (
                  <div style={{ background: "linear-gradient(135deg,#FF5252,#E53935)", color: "#fff", padding: "6px 14px", borderRadius: "20px", fontSize: "12px", fontWeight: "bold" }}>
                    ⚡ {db.responses.filter(r=>r.isRush).length} apresurada{db.responses.filter(r=>r.isRush).length !== 1 ? "s" : ""}
                  </div>
                )}
              </div>
              {total === 0 ? (
                <div style={{ textAlign: "center", padding: "3rem", color: "#aaa", background: "#fff", borderRadius: "12px", border: "1px solid #E0E4EE" }}>Sin datos disponibles aún.</div>
              ) : (
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(440px,1fr))", gap: "1.25rem" }}>
                  {QUESTIONS.map((q, qi) => {
                    const data = supStats[q.id] ?? [];
                    const tot  = data.reduce((a,b)=>a+b,0);
                    const maxIdx = data.indexOf(Math.max(...data));
                    const rushD  = q.options.map(opt => db.responses.filter(r=>r.isRush && r.answers[q.id]===opt).length);
                    const hasRush = rushD.some(v=>v>0);
                    return (
                      <div key={q.id} style={{ background: "#fff", border: `1px solid ${hasRush ? "#FFCDD2" : "#E0E4EE"}`, borderRadius: "16px", padding: "1.5rem", boxShadow: "0 2px 12px rgba(0,0,0,0.04)" }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "12px" }}>
                          <div>
                            <div style={{ color: CHART_PALETTE[qi % CHART_PALETTE.length].bar, fontSize: "10px", letterSpacing: "0.12em", fontWeight: "bold", textTransform: "uppercase", marginBottom: "4px" }}>PREGUNTA {q.id}</div>
                            <p style={{ fontSize: "13px", color: "#0A1628", fontWeight: "bold", margin: 0, lineHeight: "1.5", maxWidth: "360px" }}>{q.text}</p>
                          </div>
                          {hasRush && <span style={{ background: "#FFF0F0", border: "1px solid #FFCDD2", color: "#C62828", fontSize: "10px", fontWeight: "bold", padding: "2px 8px", borderRadius: "20px", flexShrink: 0, marginLeft: "8px" }}>⚡ apr.</span>}
                        </div>
                        {/* Summary strip */}
                        <div style={{ display: "flex", gap: "12px", marginBottom: "16px", padding: "10px 14px", background: "#F8F9FC", borderRadius: "10px" }}>
                          <div style={{ textAlign: "center" }}>
                            <div style={{ fontSize: "18px", fontWeight: "bold", color: CHART_PALETTE[qi % CHART_PALETTE.length].bar }}>{tot}</div>
                            <div style={{ fontSize: "10px", color: "#999" }}>respuestas</div>
                          </div>
                          <div style={{ textAlign: "center", marginLeft: "auto" }}>
                            <div style={{ fontSize: "18px", fontWeight: "bold", color: "#0A1628" }}>{tot > 0 ? Math.round((Math.max(...data)/tot)*100) : 0}%</div>
                            <div style={{ fontSize: "10px", color: "#999" }}>opción líder</div>
                          </div>
                        </div>
                        <ModernBarChart data={data} labels={q.options} palette={Array(q.options.length).fill(CHART_PALETTE[qi % CHART_PALETTE.length])} total={tot} rushData={rushD} />
                        {tot > 0 && <>
                          <div style={{ height: "1px", background: "#F0F0F0", margin: "16px 0" }} />
                          <DonutChart data={data} labels={q.options} palette={Array(q.options.length).fill(0).map((_,i) => CHART_PALETTE[i % CHART_PALETTE.length])} />
                        </>}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          )}

        </div>
      </div>
    );
  }

  return null;
}
