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
  const children = [run(`[${n}] ${text} `, { size: 20 })];
  if (url) {
    children.push(new ExternalHyperlink({
      link: url,
      children: [run('Consultar fuente', { size: 20, color: '0563C1', underline: {} })],
    }));
  }
  return new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.JUSTIFIED, children });
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
  P('La presente memoria justificativa tiene como objetivo solicitar la consolidación de un programa permanente de monitorización de ciberseguridad del equipamiento electromédico conectado del HOSPITAL UNIVERSITARIO REGIONAL DE MÁLAGA, integrado con los servicios TIC corporativos, y dotarlo de un técnico especialista de Electromedicina con dedicación específica y un técnico informático que trabaje conjuntamente en esta función.'),
  P('El parque gestionado comprende aproximadamente 11.000 equipos, de los cuales más de 1.000 están conectados, según la información operativa del Servicio. Su protección requiere conocer qué dispositivos, servidores y aplicaciones están en uso, evaluar las alertas que les afectan y ejecutar medidas compatibles con la continuidad asistencial.'),
  P('La revisión realizada el 22 de octubre de 2026 con Claroty xDome ante la alerta PHP2Shell acredita la utilidad de disponer de información técnica centralizada. Justifica además una organización capaz de interpretar sus resultados, completar los datos pendientes y coordinar la respuesta. La necesidad abarca plataforma, cobertura de red, procedimientos y personal; su alcance no se limita a la alerta actual.'),
  P('La presente solicitud se enmarca en una gestión planificada y responsable del parque tecnológico y se fundamenta en datos objetivos y verificables procedentes de:'),
  B('Inventario corporativo de equipamiento electromédico (MANSIS).'),
  B('Información operativa del Servicio sobre el parque conectado y sus dependencias técnicas.'),
  B('Consulta de la plataforma de monitorización Claroty xDome y avisos de seguridad publicados por organismos oficiales y fabricantes.'),

  H('1.1. Cartera de Servicios a la que responde'),
  P('El programa dará soporte a las áreas asistenciales que dependen de equipamiento conectado en el Hospital General, Materno Infantil, Hospital Civil y CARE José Estrada. Incluirá los equipos propiedad del centro y los cedidos, arrendados o aportados por proveedores cuando se conecten a las redes o participen en los procesos asistenciales.'),
  T([2400, 6100], ['Ámbito', 'Sistemas que deben incorporarse al alcance'], [
    ['Cuidados críticos y hospitalización', 'Centrales de monitorización, telemetría, monitores y pasarelas de integración.'],
    ['Quirófanos y procedimientos', 'Sistemas de integración quirúrgica, archivo de imágenes, anestesia y servidores asociados.'],
    ['Diagnóstico y tratamiento', 'Radiología, hemodinámica, radioterapia y estaciones de adquisición y tratamiento.'],
    ['Laboratorios y diagnóstico especializado', 'Analizadores, estaciones de control, anatomía patológica digital y sistemas de consulta.'],
    ['Infraestructura de soporte', 'Servidores físicos y virtuales, interfaces web, accesos remotos y conexiones de fabricantes.'],
  ]),
  empty(),

  H('1.2. Necesidad e indicación clínica de la necesidad'),
  P('Un incidente de ciberseguridad puede interrumpir la transmisión de alarmas, impedir la consulta de resultados o retrasar procedimientos aunque el equipo siga encendido. Por ello, la disponibilidad clínica depende también de las redes, del software y de los servidores que conectan los dispositivos con sus usuarios.'),
  P('El mantenimiento preventivo y correctivo debe complementarse con vigilancia de versiones, comunicaciones y exposiciones. Una revisión puntual aporta una fotografía; las altas, cambios de configuración, actualizaciones y nuevas vulnerabilidades modifican el riesgo durante toda la vida útil del sistema.'),
  P('Los dispositivos con soporte limitado requieren atención particular. Cuando un fabricante no permite actualizar de inmediato, deben valorarse restricciones de comunicaciones, segmentación u otras medidas compensatorias, documentando el riesgo residual y su aceptación por el responsable competente.'),
  P('La vigilancia se extenderá a la red interna. Un portal no publicado en Internet puede seguir siendo relevante ante accesos remotos, conexiones desde segmentos comprometidos o movimientos laterales. El programa priorizará servicios críticos, dependencias compartidas y equipos cuya interrupción tenga consecuencias asistenciales.'),
  P('Se evitarán desconexiones, exploraciones activas o cambios automáticos sobre equipos asistenciales sin evaluación y autorización. Las medidas se acordarán con TIC, Electromedicina, el fabricante y el servicio usuario, con alternativas de continuidad y comprobación posterior.'),

  H('1.2.1. Alertas y casos recientes que fundamentan la necesidad'),
  H3('PHP2Shell y respuesta urgente en octubre de 2026'),
  P('CERT-EU publicó el 21 de octubre de 2026 el aviso 2026-018 sobre CVE-2026-109618 y CVE-2026-109619. Ambas vulnerabilidades pueden permitir ejecución remota de código sin autenticación en instalaciones PHP que atienden solicitudes web. Las correcciones se publicaron para las ramas soportadas y el aviso recomienda actualizar o aislar temporalmente las aplicaciones expuestas cuando no sea viable corregirlas de inmediato. [1]'),
  P('Para responder a una alerta de este tipo hay que identificar componentes de software, localizar sus equipos y confirmar la exposición real. La marca del dispositivo o la ausencia de ".php" en una dirección web no resuelven esa comprobación. La urgencia exige disponer de inventario técnico previo y responsables de verificación.'),
  H3('Monitores Contec y Epsimed en 2025'),
  P('La FDA comunicó el 30 de enero de 2025 vulnerabilidades en Contec CMS8000 y Epsimed MN-120 con posibles efectos sobre el funcionamiento y los datos. La actualización del 2 de julio informó de un parche que elimina la funcionalidad de red. Esto ilustra que una corrección de seguridad puede modificar la prestación clínica y requerir evaluación conjunta antes de su implantación. [2]'),
  H3('Software de terceros en productos Philips en 2025'),
  P('El aviso de Philips de febrero de 2025 sobre jQuery CVE-2020-11023 identifica versiones de productos que podrían verse afectadas, entre ellas IntelliBridge Enterprise e IntelliVue Guardian Software. Muestra que el análisis debe alcanzar las librerías y aplicaciones asociadas, y que la respuesta depende de la versión y de las instrucciones verificadas del fabricante. No acredita afectación de productos Philips instalados en el HRUM. [3]'),
  H3('Interrupción de servicios de laboratorio en Synnovis'),
  P('El incidente de junio de 2024 en Synnovis alteró servicios de laboratorio y comprometió datos. Portsmouth Hospitals comunicó que en diciembre de 2025 recibió confirmación de que la información sustraída incluía datos de su organización. El caso evidencia dependencias entre hospitales y proveedores y consecuencias que pueden prolongarse tras recuperar los sistemas. [4]'),
  P('Estos antecedentes justifican mantener capacidad de vigilancia y respuesta. No se presentan como incidentes ocurridos en el HRUM ni como prueba de que una plataforma, por sí sola, hubiera evitado sus consecuencias.'),

  H('1.2.2. Experiencia de revisión con Claroty xDome en el HRUM'),
  P('En las capturas de la consulta del 22 de octubre de 2026 se observa acceso a inventario de dispositivos, aplicaciones y versiones, comunicaciones de red, documentación del fabricante y fichas de vulnerabilidades. La ficha específica de PHP2Shell mostraba "Sin impacto actual" e indicaba que no se habían identificado dispositivos afectados en la red en ese momento. [5]'),
  P('Este resultado aporta información para orientar la revisión y comunicar la situación identificada. Debe acompañarse del alcance monitorizado, de la fecha de consulta y de las comprobaciones pendientes. No equivale a una certificación de ausencia de PHP vulnerable ni garantiza que todos los equipos del centro estén cubiertos.'),
  P('La revisión puso de manifiesto la necesidad de interpretar correctamente las búsquedas: una coincidencia de las letras PHP en el identificador de un dispositivo no demuestra la instalación del lenguaje. Del mismo modo, observar HTTP o HTTPS no confirma PHP, y una conexión saliente a Internet no prueba que un portal sea accesible desde el exterior.'),
  T([2700, 5800], ['Aportación observada', 'Comprobación profesional necesaria'], [
    ['Inventario y aplicaciones detectadas', 'Conciliar identidad, ubicación, versión y servidor asociado con MANSIS y fabricante.'],
    ['Ficha de alerta y vulnerabilidades', 'Validar aplicabilidad, vigencia del dato y posibles equipos sin cobertura.'],
    ['Comunicaciones de red', 'Distinguir cliente y servidor; contrastar exposición con cortafuegos, NAT, proxy, VPN y túneles.'],
    ['Documentación MDS²', 'Verificar correspondencia con el modelo y versión instalados; completar componentes no descritos.'],
  ]),
  empty(),
  P('Los totales de la consola incluyen categorías IT, IoT, OT y médicas. No deben equipararse automáticamente al parque de Electromedicina ni utilizarse como número de equipos afectados. Es necesario delimitar el centro, los segmentos y los servicios, y reconciliar la clasificación con el inventario corporativo.'),
  P('La plataforma facilita la vigilancia, pero su eficacia local debe medirse mediante cobertura verificable, calidad de identificación, tiempos de análisis y cierre de actuaciones. La ausencia de resultados puede obedecer a ausencia de afectación, falta de datos o cobertura incompleta. Claroty describe capacidades de inventario y gestión de exposición; ello no sustituye la validación de cada despliegue. [6]'),

  H('1.3. Datos de la actividad del Servicio de destino y necesidad de dedicación profesional'),
  P('El volumen de aproximadamente 11.000 equipos y más de 1.000 conectados exige una actividad sostenida de inventario, revisión de alertas, coordinación con proveedores y seguimiento de medidas. La incorporación de una herramienta no elimina ese trabajo: permite concentrarlo en los equipos y exposiciones relevantes.'),

  H('1.3.1. Dotación conjunta de Electromedicina y TIC'),
  P('Se propone asignar un técnico especialista de Electromedicina a jornada completa a este programa y un técnico informático con dedicación específica equivalente, constituyendo un equipo funcional conjunto de dos profesionales. Esta dotación inicial es una propuesta de organización, revisable a los seis meses según cobertura y carga registrada; no deriva de una ratio normativa por número de dispositivos.'),
  T([2400, 6100], ['Perfil', 'Responsabilidad principal'], [
    ['Técnico especialista de Electromedicina', 'Identificación física y clínica; conciliación con MANSIS; criticidad; contacto técnico con fabricantes; documentación MDS² y componentes; coordinación asistencial y comprobación funcional tras cambios.'],
    ['Técnico informático de TIC', 'Cobertura de sensores y redes; análisis de servicios y versiones; verificación de exposición y accesos remotos; integración con herramientas corporativas; ejecución autorizada de cambios de red y conservación de registros.'],
    ['Trabajo conjunto', 'Validación de alertas, priorización, plan de tratamiento, pruebas, documentación, seguimiento y verificación del cierre.'],
    ['Responsables y servicios de apoyo', 'Jefaturas de Electromedicina y TIC coordinan el programa; Seguridad corporativa dirige la respuesta a incidentes; los servicios clínicos validan continuidad y ventanas de intervención.'],
  ]),
  empty(),
  P('El conocimiento clínico y técnico del dispositivo permite valorar qué función se perdería al aislarlo o actualizarlo. La competencia de TIC permite determinar las rutas de acceso, desplegar controles y asegurar la integración corporativa. Ambas funciones son necesarias para adoptar una medida eficaz sin producir una interrupción asistencial evitable.'),
  P('La dedicación debe figurar en la planificación de trabajo, con responsables y sustituciones. Si se utiliza personal existente, se deberá liberar y reconocer esa capacidad, evitando acumular la función a las tareas ordinarias sin tiempo asignado. La modalidad de provisión se determinará por Dirección y Recursos Humanos.'),
  P('La captura y generación de alertas podrá funcionar de forma continua. Dos profesionales no constituyen una cobertura humana presencial de 24 horas: las alertas urgentes fuera de jornada se canalizarán a través de los mecanismos corporativos de seguridad y continuidad que se acuerden.'),

  H('1.3.2. Dimensionamiento inicial y plan de implantación'),
  P('Para orientar la dotación se plantea un escenario de trabajo sobre 1.000 equipos conectados. Los tiempos siguientes son hipótesis de planificación expresadas en horas persona, no mediciones de actividad ni compromisos de productividad. Se contrastarán durante los primeros seis meses y se ajustarán al número real de equipos, automatización y alertas.'),
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
  P('El escenario aproxima el trabajo a dos dedicaciones anuales si se adopta, solo para dimensionamiento, una capacidad efectiva de 1.600 horas por profesional. La jornada aplicable deberá validarse por Recursos Humanos. Se registrará el trabajo por tarea para evitar duplicidades y revisar la asignación entre perfiles.'),
  H3('Implantación propuesta'),
  B('Primer mes: designar responsables, comprobar la cobertura existente, delimitar el parque conectado y registrar accesos externos, servidores y sistemas críticos.'),
  B('Meses segundo y tercero: conciliar inventarios, validar alertas prioritarias, establecer el circuito con fabricantes y probar el procedimiento de cambios.'),
  B('Meses cuarto a sexto: ampliar cobertura, revisar medidas compensatorias y presentar la evaluación de carga y resultados.'),
  H3('Indicadores de seguimiento'),
  P('Se informará mensualmente de los siguientes indicadores. Cada resultado indicará denominador, fecha y datos pendientes. Las metas se aprobarán tras establecer la línea de base.'),
  B('Porcentaje de equipos conectados cubiertos sobre los identificados.'),
  B('Porcentaje de equipos con responsable y versión verificados.'),
  B('Sistemas críticos pendientes de revisión.'),
  B('Tiempo entre alerta y valoración documentada.'),
  B('Actuaciones cerradas y verificadas.'),
  B('Accesos externos pendientes de regularizar.'),
  B('Horas persona por actividad.'),

  H('1.4. Estimación de impacto económico'),
  H('1.4.1. Estimación de impacto compra/arrendamiento (plataforma y medios técnicos)'),
  P('La necesidad consiste en consolidar o ampliar, según la cobertura y las condiciones corporativas existentes, la capacidad de monitorización de los más de 1.000 equipos conectados y sus sistemas asociados. Antes de contratar se verificará qué licencias, sensores, módulos e integraciones están ya disponibles para evitar duplicidades.'),
  P('La valoración económica deberá desglosar implantación, extensión de cobertura, licencia o suscripción, soporte, almacenamiento y conservación de evidencias, formación e integración con el inventario y los circuitos TIC. Se solicitará un coste total a tres años, diferenciando gasto inicial y recurrente. No se fija importe porque no se dispone de oferta económica para este alcance.'),
  H('1.4.2. Estimación de impacto capítulo I y II (personal y gasto recurrente)'),
  P('La propuesta incorpora dos dedicaciones específicas: un técnico especialista de Electromedicina y un técnico informático. Recursos Humanos y Gestión Económica deberán calcular el coste anual completo según categoría, modalidad de provisión y situación de plantilla. La reasignación de personal existente también tiene un coste de oportunidad y requiere cobertura de las funciones que se dejan de realizar.'),
  P('El coste del primer año será la suma de implantación y medios iniciales, gastos recurrentes, dedicación de ambos perfiles y formación. Los años siguientes incluirán personal, servicio y soporte, actualización de competencias y reposición o ampliación necesaria. La imputación a capítulos presupuestarios se determinará conforme a la modalidad finalmente elegida.'),
  P('Se espera reducir la exposición no identificada y mejorar la rapidez de respuesta. No se atribuyen porcentajes de reducción de incidentes ni ahorros económicos sin medición local. La evaluación económica comparará el coste del programa con escenarios razonados de indisponibilidad, recuperación y actividad diferida, sin presentar como ahorro cierto el coste de un incidente hipotético.'),

  H('1.5. Requisitos de funcionamiento y coordinación'),
  P('El programa priorizará descubrimiento pasivo y consultas compatibles con los dispositivos; control de acceso por funciones y registro de actuaciones; integración con los servicios TIC; y procedimientos de cambio con validación del fabricante, pruebas, alternativa de continuidad y reversión. Cualquier actuación activa se someterá a autorización técnica y clínica según su impacto.'),
  P('El Esquema Nacional de Seguridad contempla vigilancia continua, reevaluación e inventario de activos, así como gestión de actualizaciones. Las medidas concretas se determinarán por el alcance y la categorización del sistema y la política corporativa. El ENS no prescribe una marca ni una dotación fija de dos técnicos. [7]'),

  H('2. Conclusión'),
  P('Según los criterios de este informe se considera necesaria la consolidación de un programa permanente de monitorización proactiva del equipamiento electromédico conectado del HRUM, vinculado al inventario de aproximadamente 11.000 equipos y orientado inicialmente a los más de 1.000 equipos conectados y sus dependencias técnicas.'),
  P('Se solicita a Dirección autorizar la definición y valoración del alcance técnico, la utilización coordinada de la plataforma corporativa disponible y la asignación de un técnico especialista de Electromedicina y un técnico informático con dedicación específica conjunta. La dotación inicial propuesta es de una jornada completa por perfil, con revisión a los seis meses sobre actividad registrada.'),
  P('El episodio PHP2Shell muestra la necesidad de responder con información verificable y responsables definidos. La consulta en Claroty aporta una evidencia útil sobre la afectación identificada en ese momento, pero su interpretación y las actuaciones posteriores requieren continuidad profesional, contraste con fabricantes y capacidad de intervención de TIC.'),
  new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.JUSTIFIED, keepNext: true, children: [run('La aprobación solicitada permitirá concretar presupuesto, responsabilidades y cobertura antes del despliegue o ampliación. Se presentará un informe de seguimiento a los tres meses y una evaluación de resultados, cobertura y dimensionamiento a los seis meses.', { size: 20 })] }),
  new Paragraph({ spacing: { after: 120 }, keepNext: true, children: [] }),
  new Paragraph({ spacing: { after: 0 }, keepNext: true, children: [run('Jerónimo López Sanz', { size: 20 })] }),
  new Paragraph({ spacing: { after: 0 }, keepNext: true, children: [run('Jefe del Servicio de Electromedicina', { size: 20 })] }),
  new Paragraph({ spacing: { after: 120 }, children: [run('Hospital Regional Universitario de Málaga', { size: 20 })] }),

  new Paragraph({ children: [new PageBreak()] }),

  H('ANEXO. Fuentes y documentación de soporte'),
  source(1, 'CERT-EU. Security Advisory 2026-018. Critical Vulnerabilities in PHP. 21 de octubre de 2026.', 'https://cert.europa.eu/publications/security-advisories/2026-018'),
  source(2, 'FDA. Cybersecurity Vulnerabilities with Certain Patient Monitors from Contec and Epsimed. 30 de enero de 2025; actualización de 2 de julio de 2025.', 'https://www.fda.gov/medical-devices/safety-communications/cybersecurity-vulnerabilities-certain-patient-monitors-contec-and-epsimed-fda-safety-communication'),
  source(3, 'Philips. Security Advisory Archives 2025. JQuery Advisory CVE-2020-11023. Publicación 21 de febrero; actualización 25 de febrero de 2025.', 'https://www.philips.com/a-w/security/security-advisories/product-security-2025.html'),
  source(4, 'Portsmouth Hospitals University NHS Trust. Synnovis Cyber Incident. Comunicación sobre el incidente de 2024 y la notificación recibida en diciembre de 2025.', 'https://www.porthosp.nhs.uk/departments-and-services/pathology/synnovis-cyber-incident'),
  source(5, 'Evidencia interna HRUM. Capturas de la sesión de Claroty xDome aportadas el 22 de octubre de 2026: inventario, aplicaciones, comunicaciones y alerta PHP2Shell. Magnitud del parque: información operativa facilitada por el Servicio.', null),
  source(6, 'Claroty. xDome for Healthcare Environments. Capacidades generales de inventario y gestión de exposición.', 'https://es.claroty.com/healthcare-cybersecurity/xdome-for-healthcare'),
  source(7, 'BOE. Real Decreto 311/2022, de 3 de mayo. Esquema Nacional de Seguridad. Artículo 10 y anexo II, medidas op.exp.1 y op.exp.4.', 'https://www.boe.es/eli/es/rd/2022/05/03/311/con'),
  P('Fuentes públicas consultadas el 22 de octubre de 2026. Los ejemplos externos fundamentan la necesidad de vigilancia y no acreditan por sí mismos afectación del parque del HRUM.'),
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
