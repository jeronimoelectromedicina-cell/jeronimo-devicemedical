// Genera la Memoria justificativa de monitorización de ciberseguridad
// con el mismo formato que la Memoria justificativa de equipamiento (modelo HRUM).
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  AlignmentType, WidthType, ShadingType, BorderStyle, VerticalAlign, PageBreak,
  ExternalHyperlink,
} = require('docx');

const ASSETS = process.argv[2];
const OUT = process.argv[3];

const VERDE = '007A3D';
const FUENTE = 'Source Sans Pro';
const GRIS = 'F5F5F5';

const logoJunta = fs.readFileSync(path.join(ASSETS, 'logo_junta.png'));
const logoHospital = fs.readFileSync(path.join(ASSETS, 'logo_hospital.png'));

const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};
const thin = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const allThin = { top: thin, bottom: thin, left: thin, right: thin };

// ---------- helpers de formato (idénticos al modelo) ----------
const run = (text, opts = {}) => new TextRun({ text, font: FUENTE, ...opts });

function header() {
  return [
    new Table({
      width: { size: 8504, type: WidthType.DXA },
      columnWidths: [4252, 4252],
      rows: [new TableRow({ children: [
        new TableCell({
          borders: noBorders, verticalAlign: VerticalAlign.TOP,
          width: { size: 4252, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 200 },
          children: [
            new Paragraph({ spacing: { before: 0, after: 60 }, children: [
              new ImageRun({ data: logoJunta, transformation: { width: 102, height: 59 }, type: 'png' }),
            ]}),
            new Paragraph({ spacing: { before: 0, after: 0 }, children: [
              run('Consejería de Sanidad, Presidencia y Emergencias', { bold: true, size: 18, color: VERDE }),
            ]}),
            new Paragraph({ spacing: { before: 0, after: 0 }, children: [
              run('Servicio Andaluz de Salud', { size: 18 }),
            ]}),
          ],
        }),
        new TableCell({
          borders: noBorders, verticalAlign: VerticalAlign.TOP,
          width: { size: 4252, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 200, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 0 }, children: [
              new ImageRun({ data: logoHospital, transformation: { width: 179, height: 89 }, type: 'png' }),
            ]}),
          ],
        }),
      ]})],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 0 }, children: [
      run('Subdirección de Innovación Tecnológica', { bold: true, size: 18, color: VERDE }),
    ]}),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 0, after: 100 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: VERDE, space: 1 } },
      children: [run('Servicio de Electromedicina', { size: 18 })],
    }),
  ];
}

const titleLine = (t) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
  children: [run(t, { bold: true, size: 24, color: VERDE })],
});

const H = (t) => new Paragraph({
  spacing: { before: 240, after: 80 }, keepNext: true,
  children: [run(t, { bold: true, size: 22, color: VERDE })],
});

const H3 = (t) => new Paragraph({
  spacing: { before: 160, after: 60 }, keepNext: true,
  children: [run(t, { bold: true, size: 20, color: VERDE })],
});

const P = (t) => new Paragraph({
  spacing: { after: 120 }, alignment: AlignmentType.JUSTIFIED,
  children: [run(t, { size: 20 })],
});

const B = (t) => new Paragraph({
  spacing: { after: 80 }, indent: { left: 280 }, alignment: AlignmentType.JUSTIFIED,
  children: [run('• ' + t, { size: 20 })],
});

const empty = () => new Paragraph({ spacing: { after: 120 }, children: [] });

function cell(text, { fill, header = false, width } = {}) {
  return new TableCell({
    borders: allThin,
    width: { size: width, type: WidthType.DXA },
    shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    children: [new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [run(text, { size: 18, bold: header, color: header ? 'FFFFFF' : undefined })],
    })],
  });
}

// Tabla corporativa: cabecera verde con texto blanco, filas alternas gris/blanco.
function T(widths, head, rows) {
  const total = widths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, { fill: VERDE, header: true, width: widths[i] })) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, i) => cell(c, { fill: ri % 2 === 0 ? GRIS : undefined, width: widths[i] })) })),
    ],
  });
}

// Tabla de datos de la solicitud (como el modelo: fila superior verde vacía, filas alternas).
function metaTable(rows) {
  const w = [3104, 5616];
  return new Table({
    width: { size: 8720, type: WidthType.DXA },
    columnWidths: w,
    alignment: AlignmentType.CENTER,
    rows: [
      new TableRow({ children: w.map((x) => new TableCell({ borders: allThin, width: { size: x, type: WidthType.DXA }, shading: { fill: VERDE, type: ShadingType.CLEAR, color: 'auto' }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [] })] })) }),
      ...rows.map(([k, v], ri) => new TableRow({ children: [
        cell(k, { fill: ri % 2 === 0 ? GRIS : undefined, width: w[0] }),
        cell(v, { fill: ri % 2 === 0 ? GRIS : undefined, width: w[1] }),
      ]})),
    ],
  });
}

function source(n, text, url) {
  const children = [run(`[${n}] ${text} `, { size: 18 })];
  if (url) {
    children.push(new ExternalHyperlink({
      link: url,
      children: [run('Consultar fuente', { size: 18, color: '0563C1', underline: {} })],
    }));
  }
  return new Paragraph({ spacing: { after: 40 }, alignment: AlignmentType.JUSTIFIED, children });
}

// ---------- contenido ----------
const body = [
  ...header(),
  titleLine('MEMORIA JUSTIFICATIVA DE LA NECESIDAD'),
  titleLine('DE MONITORIZACIÓN DE CIBERSEGURIDAD'),
  titleLine('DEL EQUIPAMIENTO ELECTROMÉDICO'),
  new Paragraph({ spacing: { after: 120 }, children: [] }),
  metaTable([
    ['NOMBRE DEL CENTRO', 'HOSPITAL UNIVERSITARIO REGIONAL DE MÁLAGA'],
    ['NOMBRE DEL SERVICIO/UNIDAD/UGC SOLICITANTE', 'Servicio de Electromedicina en coordinación con TIC.'],
    ['NOMBRE DE LA NECESIDAD', 'Monitorización proactiva de ciberseguridad del equipamiento electromédico conectado y dotación técnica conjunta de Electromedicina y TIC'],
    ['DESTINO', 'Hospital General, Materno Infantil, Hospital Civil y CARE José Estrada'],
    ['FECHA', '22 de octubre de 2026.'],
  ]),

  H('NECESIDAD DEL CENTRO SANITARIO.'),

  H('1. Introducción'),
  P('La presente memoria justificativa tiene como objetivo solicitar la consolidación de un programa permanente de monitorización de ciberseguridad del equipamiento electromédico conectado del HOSPITAL UNIVERSITARIO REGIONAL DE MÁLAGA, integrado con los servicios TIC corporativos, y dotarlo de un técnico especialista de Electromedicina y un técnico informático con dedicación específica conjunta.'),
  P('El parque gestionado comprende aproximadamente 11.000 equipos, de los cuales más de 1.000 están conectados. Su protección requiere conocer qué dispositivos, servidores y aplicaciones están en uso, evaluar las alertas que les afectan y ejecutar medidas compatibles con la continuidad asistencial. La revisión realizada el 22 de octubre de 2026 con Claroty xDome ante la alerta PHP2Shell acredita la utilidad de disponer de información técnica centralizada y de una organización capaz de interpretarla y coordinar la respuesta. La necesidad abarca plataforma, cobertura de red, procedimientos y personal; no se limita a la alerta actual.'),
  P('La solicitud se fundamenta en datos objetivos procedentes de:'),
  B('Inventario corporativo de equipamiento electromédico (MANSIS) e información operativa del parque conectado.'),
  B('Consulta de la plataforma Claroty xDome y avisos de seguridad de organismos oficiales y fabricantes.'),

  H('1.1. Cartera de Servicios a la que responde'),
  P('El programa dará soporte a las áreas asistenciales que dependen de equipamiento conectado en el Hospital General, Materno Infantil, Hospital Civil y CARE José Estrada, incluyendo los equipos propios y los cedidos, arrendados o aportados por proveedores cuando se conecten a las redes o participen en los procesos asistenciales.'),
  T([2400, 6100], ['Ámbito', 'Sistemas incluidos en el alcance'], [
    ['Cuidados críticos y hospitalización', 'Centrales de monitorización, telemetría, monitores y pasarelas de integración.'],
    ['Quirófanos y procedimientos', 'Integración quirúrgica, archivo de imágenes, anestesia y servidores asociados.'],
    ['Diagnóstico y tratamiento', 'Radiología, hemodinámica, radioterapia y estaciones de adquisición y tratamiento.'],
    ['Laboratorios y diagnóstico especializado', 'Analizadores, estaciones de control, anatomía patológica digital y sistemas de consulta.'],
    ['Infraestructura de soporte', 'Servidores físicos y virtuales, interfaces web, accesos remotos y conexiones de fabricantes.'],
  ]),
  empty(),

  H('1.2. Necesidad e indicación clínica de la necesidad'),
  P('Un incidente de ciberseguridad puede interrumpir la transmisión de alarmas, impedir la consulta de resultados o retrasar procedimientos aunque el equipo siga encendido: la disponibilidad clínica depende también de las redes, del software y de los servidores. El mantenimiento preventivo y correctivo debe complementarse con vigilancia continua de versiones, comunicaciones y exposiciones, ya que altas, cambios de configuración y nuevas vulnerabilidades modifican el riesgo durante toda la vida útil del sistema.'),
  P('Los dispositivos con soporte limitado requieren medidas compensatorias (restricción de comunicaciones, segmentación) con riesgo residual documentado y aceptado. La vigilancia se extenderá a la red interna, priorizando servicios críticos y dependencias compartidas. No se realizarán desconexiones, exploraciones activas ni cambios automáticos sobre equipos asistenciales sin evaluación y autorización acordada con TIC, Electromedicina, el fabricante y el servicio usuario.'),
  P('Alertas y casos recientes que fundamentan la necesidad:'),
  T([2300, 6200], ['Caso', 'Lección para el HRUM'], [
    ['PHP2Shell (CERT-EU 2026-018, 21/10/2026) [1]', 'Ejecución remota de código sin autenticación en instalaciones PHP (CVE-2026-109618 y CVE-2026-109619). Responder exige inventario técnico previo, localizar los equipos afectados y confirmar la exposición real.'],
    ['Monitores Contec CMS8000 y Epsimed MN-120 (FDA, 2025) [2]', 'El parche elimina la funcionalidad de red: una corrección de seguridad puede modificar la prestación clínica y requiere evaluación conjunta antes de implantarse.'],
    ['jQuery CVE-2020-11023 en productos Philips (2025) [3]', 'El análisis debe alcanzar librerías y aplicaciones asociadas; la respuesta depende de la versión y de las instrucciones del fabricante.'],
    ['Synnovis (NHS, 2024-2025) [4]', 'Dependencias entre hospitales y proveedores con consecuencias prolongadas tras recuperar los sistemas.'],
  ]),
  empty(),
  P('Estos antecedentes justifican mantener capacidad de vigilancia y respuesta; no se presentan como incidentes ocurridos en el HRUM.'),
  P('En la consulta de Claroty xDome del 22 de octubre de 2026 se accedió a inventario de dispositivos, aplicaciones, comunicaciones de red y fichas de vulnerabilidades. La ficha de PHP2Shell mostraba "Sin impacto actual" [5]. Este resultado orienta la revisión, pero no equivale a una certificación de ausencia de PHP vulnerable ni garantiza cobertura total: los totales de la consola mezclan categorías IT, IoT, OT y médicas, y una coincidencia de "PHP" en un identificador o la observación de HTTP no demuestran la instalación del lenguaje. La eficacia local de la plataforma debe medirse por cobertura verificable, calidad de identificación y cierre de actuaciones [6].'),

  H('1.3. Datos de la actividad del Servicio de destino'),
  P('El volumen de 11.000 equipos y más de 1.000 conectados exige una actividad sostenida de inventario, revisión de alertas, coordinación con proveedores y seguimiento de medidas. La herramienta no elimina ese trabajo: permite concentrarlo en los equipos y exposiciones relevantes. Se propone un equipo funcional conjunto de dos profesionales, revisable a los seis meses según carga registrada:'),
  T([2400, 6100], ['Perfil', 'Responsabilidad principal'], [
    ['Técnico especialista de Electromedicina', 'Identificación física y clínica; conciliación con MANSIS; criticidad; contacto con fabricantes; documentación MDS²; comprobación funcional tras cambios.'],
    ['Técnico informático de TIC', 'Cobertura de sensores y redes; análisis de servicios y versiones; verificación de exposición y accesos remotos; ejecución autorizada de cambios de red y registros.'],
    ['Trabajo conjunto', 'Validación de alertas, priorización, plan de tratamiento, pruebas, documentación y verificación del cierre. Las alertas fuera de jornada se canalizan por los mecanismos corporativos.'],
  ]),
  empty(),
  P('Escenario de dimensionamiento inicial sobre 1.000 equipos conectados (hipótesis de planificación en horas persona, a contrastar en los primeros seis meses):'),
  T([3300, 3700, 1500], ['Actividad estimada', 'Hipótesis anual', 'Horas persona'], [
    ['Revisión técnica de la ficha por equipo', '1.000 equipos × 1,5 h/año', '1.500'],
    ['Altas y cambios de sistemas', '20 actuaciones/mes × 2 h × 12 meses', '480'],
    ['Triaje de alertas y comunicaciones', '2 h/día × 220 días', '440'],
    ['Análisis de avisos y consultas a fabricantes', '6 h/semana × 46 semanas', '276'],
    ['Pruebas y cambios coordinados', '8 h/semana × 46 semanas', '368'],
    ['Seguimiento e informes', '8 h/mes × 12 meses', '96'],
    ['Total de referencia', 'Suma del escenario inicial', '3.160'],
  ]),
  empty(),
  P('El escenario equivale a dos dedicaciones anuales con una capacidad efectiva de 1.600 horas por profesional, a validar por Recursos Humanos. Implantación: primer mes, designar responsables, comprobar cobertura y delimitar el parque conectado; meses segundo y tercero, conciliar inventarios, validar alertas prioritarias y probar el procedimiento de cambios; meses cuarto a sexto, ampliar cobertura y presentar la evaluación de carga y resultados. Se informará mensualmente de cobertura, equipos con responsable y versión verificados, sistemas críticos pendientes, tiempo entre alerta y valoración, actuaciones cerradas y horas por actividad.'),

  H('1.4. Estimación de impacto económico'),
  H('1.4.1. Estimación de impacto compra/arrendamiento'),
  P('La necesidad consiste en consolidar o ampliar la capacidad de monitorización de los más de 1.000 equipos conectados. Antes de contratar se verificará qué licencias, sensores, módulos e integraciones ya están disponibles. La valoración desglosará implantación, cobertura, licencia o suscripción, soporte, conservación de evidencias, formación e integración, con coste total a tres años diferenciando gasto inicial y recurrente. No se fija importe porque no se dispone de oferta económica para este alcance.'),
  H('1.4.2. Estimación de impacto capítulo I y II'),
  P('La propuesta incorpora dos dedicaciones específicas (técnico especialista de Electromedicina y técnico informático), cuyo coste anual calcularán Recursos Humanos y Gestión Económica según categoría y modalidad de provisión. La reasignación de personal existente tiene coste de oportunidad y requiere cobertura de las funciones que se dejan de realizar. No se atribuyen ahorros ni reducciones de incidentes sin medición local.'),

  H('1.5. Requisitos de funcionamiento y coordinación'),
  P('El programa priorizará descubrimiento pasivo, control de acceso por funciones y registro de actuaciones, integración con los servicios TIC y procedimientos de cambio con validación del fabricante, pruebas y reversión. El Esquema Nacional de Seguridad contempla vigilancia continua, inventario de activos y gestión de actualizaciones; no prescribe una marca ni una dotación fija [7].'),

  H('2. Conclusión'),
  P('Según los criterios de este informe se propone consolidar un programa permanente de monitorización proactiva del equipamiento electromédico conectado del HRUM, orientado inicialmente a los más de 1.000 equipos conectados, y se solicita a Dirección autorizar la valoración del alcance técnico, el uso coordinado de la plataforma corporativa y la asignación de un técnico especialista de Electromedicina y un técnico informático a jornada completa, con revisión a los seis meses.'),
  new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.JUSTIFIED, keepNext: true, children: [run('Se presentará un informe de seguimiento a los tres meses y una evaluación de resultados, cobertura y dimensionamiento a los seis meses.', { size: 20 })] }),
  new Paragraph({ spacing: { after: 120 }, keepNext: true, children: [] }),
  new Paragraph({ spacing: { after: 0 }, keepNext: true, children: [run('Jerónimo López Sanz', { size: 20 })] }),
  new Paragraph({ spacing: { after: 0 }, keepNext: true, children: [run('Jefe del Servicio de Electromedicina', { size: 20 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [run('Hospital Regional Universitario de Málaga', { size: 20 })] }),

  H('Fuentes'),
  source(1, 'CERT-EU. Security Advisory 2026-018. Critical Vulnerabilities in PHP. 21/10/2026.', 'https://cert.europa.eu/publications/security-advisories/2026-018'),
  source(2, 'FDA. Cybersecurity Vulnerabilities with Certain Patient Monitors from Contec and Epsimed. 30/01/2025; act. 02/07/2025.', 'https://www.fda.gov/medical-devices/safety-communications/cybersecurity-vulnerabilities-certain-patient-monitors-contec-and-epsimed-fda-safety-communication'),
  source(3, 'Philips. Security Advisory Archives 2025. JQuery Advisory CVE-2020-11023. Febrero 2025.', 'https://www.philips.com/a-w/security/security-advisories/product-security-2025.html'),
  source(4, 'Portsmouth Hospitals University NHS Trust. Synnovis Cyber Incident.', 'https://www.porthosp.nhs.uk/departments-and-services/pathology/synnovis-cyber-incident'),
  source(5, 'Evidencia interna HRUM. Capturas de la sesión de Claroty xDome del 22/10/2026.', null),
  source(6, 'Claroty. xDome for Healthcare Environments.', 'https://es.claroty.com/healthcare-cybersecurity/xdome-for-healthcare'),
  source(7, 'BOE. Real Decreto 311/2022, de 3 de mayo, Esquema Nacional de Seguridad. Art. 10 y anexo II (op.exp.1, op.exp.4).', 'https://www.boe.es/eli/es/rd/2022/05/03/311/con'),
];

const doc = new Document({
  creator: 'Servicio de Electromedicina · HRUM',
  title: 'Memoria justificativa de la necesidad de monitorización de ciberseguridad del equipamiento electromédico',
  styles: {
    default: {
      document: {
        run: { font: FUENTE, size: 20, language: { value: 'es-ES' } },
        paragraph: { spacing: { after: 200, line: 276, lineRule: 'auto' } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 720, right: 1701, bottom: 850, left: 1701, header: 720, footer: 720 },
      },
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, buf);
  console.log('OK', OUT, buf.length, 'bytes');
});
