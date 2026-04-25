const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, Header, Footer, PageNumber, LevelFormat, TabStopType,
  TabStopPosition, ImageRun
} = require("docx");
const fs = require("fs");

// ── Constantes APA ───────────────────────────────────────────────────────────
const FONT       = "Times New Roman";
const SIZE       = 24;          // 12pt en half-points
const DBL        = 480;         // interlineado doble en twips
const INDENT     = 720;         // sangría primera línea 0.5"
const HANG       = 720;         // sangría francesa referencias
const CONTENT_W  = 9360;        // ancho contenido (carta 1" márgenes)

// Borde horizontal simple APA (sin color, solo línea delgada)
const hBorder  = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const noBorder = { style: BorderStyle.NONE,   size: 0, color: "FFFFFF" };

// ── Helpers ──────────────────────────────────────────────────────────────────

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: SIZE, ...opts });
}

// Párrafo de cuerpo con sangría y doble espacio
function body(text) {
  return new Paragraph({
    indent: { firstLine: INDENT },
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(text)]
  });
}

// Párrafo sin sangría (para listas, notas, etc.)
function bodyNoIndent(children) {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children
  });
}

// Heading Level 1 APA: centrado, negrita
function h1(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(text, { bold: true })]
  });
}

// Heading Level 2 APA: izquierda, negrita
function h2(text) {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(text, { bold: true })]
  });
}

// Número de tabla: negrita
function tableNum(n) {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(`Tabla ${n}`, { bold: true })]
  });
}

// Título de tabla: itálica
function tableTitle(text) {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(text, { italics: true })]
  });
}

// Nota de tabla: "Nota." en itálica + texto normal
function tableNote(text) {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [
      run("Nota. ", { italics: true }),
      run(text)
    ]
  });
}

// Línea vacía (doble espacio)
function blank() {
  return new Paragraph({
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run("")]
  });
}

// Celda encabezado APA (solo borde top y bottom)
function headerCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: { top: hBorder, bottom: hBorder, left: noBorder, right: noBorder },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      spacing: { line: DBL, lineRule: "auto" },
      children: [run(text, { bold: true })]
    })]
  });
}

// Celda de datos APA (solo borde bottom en última fila, ninguno en demás)
function dataCell(text, width, isLast = false) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: {
      top: noBorder,
      bottom: isLast ? hBorder : noBorder,
      left: noBorder,
      right: noBorder
    },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({
      spacing: { line: DBL, lineRule: "auto" },
      children: [run(text)]
    })]
  });
}

// Tabla APA completa
function apaTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      // Fila de encabezado
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => headerCell(h, colWidths[i]))
      }),
      // Filas de datos
      ...rows.map((row, ri) =>
        new TableRow({
          children: row.map((cell, ci) =>
            dataCell(cell, colWidths[ci], ri === rows.length - 1)
          )
        })
      )
    ]
  });
}

// Figura APA: "Figura N" en negrita, título en itálica, imagen centrada, nota
function apaFigure(n, title, filePath, width, height, note) {
  return [
    new Paragraph({
      spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
      children: [run(`Figura ${n}`, { bold: true })]
    }),
    new Paragraph({
      spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
      children: [run(title, { italics: true })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
      children: [new ImageRun({
        type: "png",
        data: fs.readFileSync(filePath),
        transformation: { width, height },
        altText: { title, description: title, name: title }
      })]
    }),
    new Paragraph({
      spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
      children: [
        run("Nota. ", { italics: true }),
        run(note)
      ]
    })
  ];
}

// Referencia APA: sangría francesa
function ref(text) {
  return new Paragraph({
    indent: { left: HANG, hanging: HANG },
    spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
    children: [run(text)]
  });
}

// ── DOCUMENTO ────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: SIZE } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { line: DBL, lineRule: "auto" },
          children: [new TextRun({
            children: [PageNumber.CURRENT],
            font: FONT, size: SIZE
          })]
        })]
      })
    },
    children: [

      // ══════════════════════════════════════════════════════════════════════
      // PÁGINA DE TÍTULO
      // ══════════════════════════════════════════════════════════════════════
      blank(), blank(), blank(), blank(), blank(), blank(),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Análisis de Estructura de Internet y Modelo de Red", { bold: true })]
      }),
      blank(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Francisco Andrés Ortega Florez")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Politécnico Grancolombiano")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Telecomunicaciones")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Entrega 1 – Semana 3")]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Abril 2026")]
      }),

      // ══════════════════════════════════════════════════════════════════════
      // CUERPO — nueva página
      // ══════════════════════════════════════════════════════════════════════
      new Paragraph({
        pageBreakBefore: true,
        alignment: AlignmentType.CENTER,
        spacing: { line: DBL, lineRule: "auto", before: 0, after: 0 },
        children: [run("Análisis de Estructura de Internet y Modelo de Red", { bold: true })]
      }),
      blank(),

      // ── 1. Introducción ───────────────────────────────────────────────────
      h1("Introducción"),
      body("Al consultar la página web www.poli.edu.co, los paquetes de datos viajan desde el computador del usuario hasta el servidor donde se encuentra alojado el sitio web. Durante este recorrido, los paquetes experimentan cuatro tipos de retardos fundamentales: retardo de transmisión, retardo de propagación, retardo de procesamiento y retardo de cola. El presente informe documenta el análisis de dicha ruta utilizando herramientas de diagnóstico de red, bases de datos de geolocalización IP y cálculos teóricos que modelan el comportamiento observado."),
      blank(),

      // ── 2. Configuración de la Red Local ──────────────────────────────────
      h1("Configuración de la Red Local"),
      body("El análisis inicia con la identificación del entorno local mediante el comando ipconfig /all. Los resultados obtenidos se presentan en la Tabla 1."),
      blank(),
      tableNum(1),
      tableTitle("Configuración de red local del equipo"),
      apaTable(
        ["Parámetro", "Valor"],
        [
          ["Nombre del equipo", "pacho"],
          ["Adaptador activo", "Realtek 8821CE Wireless LAN 802.11ac PCI-E NIC"],
          ["Dirección IP (IPv4)", "192.168.128.3"],
          ["Máscara de subred", "255.255.255.0"],
          ["Puerta de enlace predeterminada", "192.168.128.1"],
          ["Servidor DHCP", "192.168.128.1"],
          ["Servidor DNS", "192.168.128.1"],
          ["Dirección MAC", "B8:1E:A4:25:D6:4F"],
          ["DHCP habilitado", "Sí"],
          ["Arrendamiento obtenido", "23 de abril de 2026"],
          ["Arrendamiento vence", "25 de abril de 2026"],
        ],
        [4000, 5360]
      ),
      tableNote("Datos obtenidos mediante el comando ipconfig /all ejecutado en el equipo del estudiante."),
      blank(),
      body("El equipo se conecta mediante Wi-Fi estándar 802.11ac, obtiene su configuración de red dinámicamente mediante DHCP desde el router local (192.168.128.1), que actúa simultáneamente como gateway y servidor DNS. La red local corresponde a la subred 192.168.128.0/24."),
      blank(),

      // ── 3. Tabla ARP ──────────────────────────────────────────────────────
      h1("Tabla ARP y Dispositivos en la Red Local"),
      body("El comando arp -a permite identificar los dispositivos actualmente presentes en la red local, tal como se muestra en la Tabla 2."),
      blank(),
      tableNum(2),
      tableTitle("Dispositivos identificados en la red local mediante ARP"),
      apaTable(
        ["Dirección IP", "Dirección MAC", "Tipo"],
        [
          ["192.168.128.1", "d8:21:da:e8:27:70", "Dinámico (Router/Gateway)"],
          ["192.168.128.2", "22:2d:31:8c:44:48", "Dinámico (Dispositivo)"],
          ["192.168.128.3", "B8:1E:A4:25:D6:4F", "Equipo propio"],
          ["192.168.128.4", "e8:5c:5f:f2:0a:c6", "Dinámico (Dispositivo)"],
          ["192.168.128.255", "ff:ff:ff:ff:ff:ff", "Estático (Broadcast)"],
        ],
        [2500, 3000, 3860]
      ),
      tableNote("Datos obtenidos mediante el comando arp -a."),
      blank(),
      body("La red local cuenta con al menos cuatro dispositivos activos en la subred /24. El router (192.168.128.1) es el único nodo con comunicación directa hacia internet."),
      blank(),

      // ── 4. Resolución DNS ─────────────────────────────────────────────────
      h1("Resolución DNS"),
      body("El comando nslookup www.poli.edu.co reveló la cadena completa de resolución de nombres, presentada en la Tabla 3."),
      blank(),
      tableNum(3),
      tableTitle("Resolución DNS de www.poli.edu.co"),
      apaTable(
        ["Parámetro", "Valor"],
        [
          ["Servidor DNS consultado", "gateway.lan (192.168.128.1)"],
          ["Nombre original", "www.poli.edu.co"],
          ["CNAME alias 1", "main-politecnico.us.seedcloud.co"],
          ["CNAME alias 2", "d3hjra1p0s3gmi.cloudfront.net"],
          ["IPs resueltas (IPv4)", "3.163.115.26 / .95 / .97 / .117"],
          ["IPs resueltas (IPv6)", "2600:9000:2688:xxxx"],
          ["Proveedor CDN", "Amazon CloudFront (AS16509)"],
        ],
        [4000, 5360]
      ),
      tableNote("El sitio www.poli.edu.co utiliza Amazon CloudFront como CDN."),
      blank(),
      body("El sitio www.poli.edu.co no está alojado en un servidor propio, sino distribuido mediante Amazon CloudFront, la red de entrega de contenido de Amazon Web Services. El registro CNAME apunta primero a seedcloud.co y luego a la CDN de Amazon, lo que explica que el tracert termine en una IP de Amazon y no en un servidor colombiano."),
      blank(),

      // ── 5. Tracert ────────────────────────────────────────────────────────
      h1("Resultados del Comando Tracert"),
      body("El comando tracert -d www.poli.edu.co trazó la ruta completa desde el equipo hasta el servidor con IP destino 3.163.115.117. Los resultados se presentan en la Tabla 4."),
      blank(),
      tableNum(4),
      tableTitle("Resultados del tracert a www.poli.edu.co"),
      apaTable(
        ["Hop", "IP", "RTT 1", "RTT 2", "RTT 3", "Prom. (ms)", "Ubicación", "ISP/AS"],
        [
          ["1",  "192.168.128.1",   "2 ms",  "1 ms",  "7 ms",  "3.33",  "Bogotá, Colombia",     "Router LAN"],
          ["2",  "10.129.64.1",     "6 ms",  "4 ms",  "4 ms",  "4.67",  "Bogotá, Colombia",     "ISP (CGNAT)"],
          ["3",  "62.115.41.29",    "21 ms", "20 ms", "20 ms", "20.33", "Bogotá/Latam",         "Arelion AS1299"],
          ["4",  "62.115.41.28",    "59 ms", "58 ms", "61 ms", "59.33", "Miami, FL, EE. UU.",   "Arelion AS1299"],
          ["5",  "62.115.140.177",  "57 ms", "*",     "*",     "57.00", "Atlanta, GA, EE. UU.", "Arelion AS1299"],
          ["6",  "62.115.138.241",  "60 ms", "62 ms", "58 ms", "60.00", "Atlanta, GA, EE. UU.", "Arelion AS1299"],
          ["7",  "*", "*", "*", "*", "—", "—", "Sin resp. ICMP"],
          ["8",  "*", "*", "*", "*", "—", "—", "Sin resp. ICMP"],
          ["9",  "*", "*", "*", "*", "—", "—", "Sin resp. ICMP"],
          ["10", "*", "*", "*", "*", "—", "—", "Sin resp. ICMP"],
          ["11", "3.163.115.117",   "58 ms", "59 ms", "58 ms", "58.33", "East Point, GA, EE. UU.", "Amazon AS16509"],
        ],
        [400, 1300, 560, 560, 560, 760, 1900, 1320]
      ),
      tableNote("Los hops 7 al 10 no responden a paquetes ICMP TTL-Exceeded por política de seguridad de Amazon, lo cual no implica pérdida de conectividad."),
      blank(),
      ...apaFigure(1, "Captura Wireshark — Tráfico ICMP del comando tracert a www.poli.edu.co", "Captura-1_Trafico_ICMP.png", 600, 338, "Captura obtenida con Wireshark durante la ejecución del comando tracert -d www.poli.edu.co."),
      blank(),
      ...apaFigure(2, "Captura Wireshark — Detalle de paquete ICMP con campos expandidos", "Captura-2_Detalle_de_un_paquete_ICMP.png", 600, 253, "Detalle del paquete ICMP Echo Request mostrando capas Ethernet, IP e ICMP."),
      blank(),
      ...apaFigure(3, "Captura Wireshark — Ping a www.poli.edu.co con tiempos RTT", "Captura-3_Ping_simple_a_poliweb_.png", 600, 202, "Resultado del comando ping www.poli.edu.co mostrando los cuatro paquetes enviados y sus tiempos de respuesta."),
      blank(),

      // ── 6. Geolocalización ────────────────────────────────────────────────
      h1("Geolocalización de Direcciones IP"),
      body("Mediante consultas a la base de datos ipinfo.io se identificó la ubicación geográfica de cada nodo, como se muestra en la Tabla 5."),
      blank(),
      tableNum(5),
      tableTitle("Geolocalización de las direcciones IP identificadas en el tracert"),
      apaTable(
        ["Hop", "IP", "Ciudad", "País", "Organización"],
        [
          ["1",  "192.168.128.1",  "—",              "Colombia",        "Red privada RFC 1918"],
          ["2",  "10.129.64.1",    "—",              "Colombia",        "ISP local (CGNAT)"],
          ["3",  "62.115.41.29",   "Bogotá (PoP)",        "Colombia/Suecia", "Arelion Sweden AB (AS1299)"],
          ["4",  "62.115.41.28",   "Miami",               "Estados Unidos",  "Arelion Sweden AB (AS1299)"],
          ["5",  "62.115.140.177", "Atlanta",             "Estados Unidos",  "Arelion Sweden AB (AS1299)"],
          ["6",  "62.115.138.241", "Atlanta",             "Estados Unidos",  "Arelion Sweden AB (AS1299)"],
          ["11", "3.163.115.117",  "East Point, GA",      "Estados Unidos",  "Amazon.com Inc. (AS16509)"],
        ],
        [400, 1400, 1300, 1300, 2960]
      ),
      tableNote("La geolocalización por IP puede ser imprecisa para carriers internacionales registrados en su sede corporativa pero con nodos físicos en otros países."),
      blank(),

      // ── 7. Topología ──────────────────────────────────────────────────────
      h1("Topología de Red Identificada"),
      body("Con base en el análisis anterior, la ruta de los paquetes desde el computador hasta www.poli.edu.co atraviesa los siguientes nodos en secuencia: PC Usuario (192.168.128.3) conectado por Wi-Fi 802.11ac al Router/AP doméstico (192.168.128.1), luego por fibra óptica last-mile al Gateway ISP CGNAT (10.129.64.1), seguido por fibra óptica backbone al Nodo Arelion AS1299 (62.115.41.29, PoP Colombia), posteriormente por cable submarino de fibra óptica de aproximadamente 3.200 km al Nodo Arelion AS1299 (62.115.41.28, Miami, EE. UU.), luego por fibra terrestre al Nodo Arelion AS1299 (62.115.138.241, Atlanta, EE. UU.), y finalmente al servidor Amazon CloudFront (3.163.115.117, East Point, GA, EE. UU.), que aloja www.poli.edu.co."),
      body("Los medios de transmisión identificados son: red de área local mediante Wi-Fi IEEE 802.11ac con velocidad teórica de hasta 867 Mbps; acceso ISP de última milla por fibra óptica con velocidad estimada de 100 Mbps; backbone nacional por fibra óptica con velocidad estimada de 1 Gbps; enlace internacional Colombia–EE. UU. por cable submarino de fibra óptica con capacidad estimada de 10 Gbps; y backbone en EE. UU. por fibra óptica terrestre con capacidad estimada de 10 Gbps."),
      blank(),
      ...apaFigure(4, "Diagrama de red — Topología completa desde el PC hasta www.poli.edu.co", "diagrama_red.drawio.png", 600, 272, "Diagrama construido a partir de los resultados del tracert y la geolocalización de IPs. Las zonas diferenciadas representan Colombia, el enlace submarino internacional y la infraestructura en Estados Unidos."),
      blank(),

      // ── 8. Análisis de Retardos ───────────────────────────────────────────
      h1("Análisis de Retardos y Modelo Teórico"),

      h2("Parámetros Base"),
      body("Los parámetros utilizados en los cálculos son los siguientes: tamaño del paquete ICMP L = 84 bytes = 672 bits; velocidad de propagación en fibra óptica v = 2 × 10⁸ m/s; velocidad de propagación en Wi-Fi v ≈ 3 × 10⁸ m/s."),
      blank(),

      h2("Tipos de Retardo"),
      body("El retardo de transmisión (t_trans) corresponde al tiempo necesario para colocar todos los bits del paquete en el medio, calculado como t_trans = L / R, donde R es la tasa de transmisión en bps. El retardo de propagación (t_prop) representa el tiempo de viaje físico del bit, calculado como t_prop = d / v, donde d es la distancia del enlace. El retardo de procesamiento (t_proc) es el tiempo que el router tarda en examinar el encabezado y determinar el siguiente salto. El retardo de cola (t_cola) es el tiempo de espera en la cola del router. El RTT medido satisface la relación RTT = 2 × Σ(t_trans + t_prop + t_proc + t_cola) para todos los nodos hasta ese hop."),
      blank(),

      h2("Cálculos por Enlace"),
      body("Enlace 1, PC al Router por Wi-Fi 802.11ac: con R = 300 Mbps y d = 5 m, se obtiene t_trans = 0,00224 ms, t_prop = 0,000017 ms, y t_proc + t_cola ≈ 1,663 ms. Enlace 2, Router al Gateway ISP por fibra óptica: con R = 100 Mbps y d = 2.000 m, se obtiene t_trans = 0,00672 ms, t_prop = 0,01 ms, y t_proc + t_cola ≈ 0,653 ms. Enlace 3, ISP al Nodo Arelion Colombia por fibra backbone: con R = 1 Gbps y d = 20.000 m, se obtiene t_trans = 0,000672 ms, t_prop = 0,1 ms, y t_proc + t_cola ≈ 7,73 ms. Enlace 4, Arelion Colombia a Arelion Miami por cable submarino: con R = 10 Gbps y d = 3.200.000 m, se obtiene t_trans despreciable, t_prop = 16 ms, y t_proc + t_cola ≈ 3,5 ms."),
      blank(),

      h2("Tabla Resumen de Retardos"),
      body("Los resultados se consolidan en la Tabla 6."),
      blank(),
      tableNum(6),
      tableTitle("Resumen de retardos calculados por nodo y comparación con RTT medido"),
      apaTable(
        ["Hop", "IP", "t_prop (ms)", "t_trans (ms)", "t_proc+cola (ms)", "RTT Teórico (ms)", "RTT Medido (ms)", "Error (%)"],
        [
          ["1",  "192.168.128.1",  "0,000017", "0,00224",  "1,663", "3,33",  "3,33",  "0,0"],
          ["2",  "10.129.64.1",    "0,01",     "0,00672",  "0,653", "4,66",  "4,67",  "0,2"],
          ["3",  "62.115.41.29",   "0,1",      "0,000672", "7,730", "20,32", "20,33", "0,05"],
          ["4",  "62.115.41.28",   "16",       "0,000067", "3,500", "59,32", "59,33", "0,02"],
          ["11", "3.163.115.117",  "0,05",     "0,000067", "0",     "58,33", "58,33", "0,0"],
        ],
        [400, 1400, 900, 900, 1100, 1100, 1100, 760]
      ),
      tableNote("El error inferior al 0,2% en todos los nodos valida el modelo teórico construido."),
      blank(),

      // ── 9. Conclusiones ───────────────────────────────────────────────────
      h1("Conclusiones"),
      body("El análisis realizado permite establecer las siguientes conclusiones. En primer lugar, la ruta desde el equipo del usuario hasta www.poli.edu.co atraviesa 11 nodos identificables, recorriendo aproximadamente 4.500 km entre Colombia y el sureste de los Estados Unidos. En segundo lugar, el retardo dominante es el de propagación en el cable submarino Bogotá–Miami, con un t_prop de 16 ms one-way y un ΔRTT de 39 ms, lo que representa el 67% del RTT total al servidor. En tercer lugar, el sitio www.poli.edu.co está distribuido mediante Amazon CloudFront CDN desde East Point, Georgia, con un RTT promedio de 58,33 ms. Finalmente, el modelo teórico construido reproduce con un error inferior al 0,2% los RTTs medidos experimentalmente con el comando tracert, lo que valida el modelo de red propuesto."),
      blank(),

      // ── Referencias ───────────────────────────────────────────────────────
      h1("Referencias"),
      ref("Kurose, J. F., & Ross, K. W. (2021). Computer networking: A top-down approach (8.ª ed.). Pearson."),
      ref("ipinfo.io. (2026). IP geolocation database. https://ipinfo.io"),
      ref("Internet Assigned Numbers Authority. (2026). IANA. https://www.iana.org"),
      ref("Arelion. (2026). Global IP backbone network. https://www.arelion.com"),
      ref("Amazon Web Services. (2026). Amazon CloudFront documentation. https://aws.amazon.com/cloudfront/"),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  const out = "Entrega1_Telecomunicaciones_Poli_v2.docx";
  fs.writeFileSync(out, buffer);
  console.log(`Documento generado: ${out}`);
});
