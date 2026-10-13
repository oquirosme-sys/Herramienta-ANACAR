/* Biblioteca de símbolos del diagrama unifilar (bloques de AutoCAD), con la geometría de los bloques de "DU prueba.dwg" y "DU ELÉCTRICO 3.dwg":
   interruptor (dos contactos con arco y trazos por polo), transformador seco, ATS y MTS (con puertos normal / alterno / carga), generador y tierra.
   Coordenadas en unidades del bloque, y hacia arriba; el punto de inserción es la conexión superior (el generador, la inferior).
   Primitivas: ['L',x1,y1,x2,y2] línea · ['C',cx,cy,r] círculo · ['F',cx,cy,r] punto relleno · ['A',cx,cy,r,a1,a2] arco antihorario en grados · ['T',x,y,h,texto] texto centrado.
   attrs: atributos del bloque (AutoCAD los muestra al hacer doble clic y permite extraerlos): [tag, prompt, x, y, alto, visible]. */
(function (g) {
  'use strict';
  const trazos = n => ({ 1: [-6.59], 2: [-6.19, -6.99], 3: [-5.79, -6.59, -7.39] }[n] || [-6.59]).map(y => ['L', 2.29, y, 4.56, y]);
  const interruptor = n => ({
    capa: 'PROTECCION', k: 2.6, alto: 13.28,
    prims: [['L', 0, 0, 0, -2.57], ['F', 0, -3.23, 1], ['A', 0, -6.59, 3.46, 289.9, 70.1]].concat(trazos(n), [['F', 0, -9.95, 1], ['L', 0, -10.61, 0, -13.28]]),
    attrs: [['POLOS', 'Polos', -4, -4.2, 1.6, 0], ['AMPERIOS', 'Amperios', -4, -6.8, 1.6, 0], ['UNIDAD', 'Unidad de disparo', -4, -9.4, 1.6, 0], ['MODELO', 'Modelo de referencia', -4, -12, 1.6, 0], ['SCCR', 'SCCR (kA)', -4, -14.6, 1.6, 0]],
  });
  const SIMB = {
    INTERRUPTOR_1P: interruptor(1), INTERRUPTOR_2P: interruptor(2), INTERRUPTOR_3P: interruptor(3),
    // Transformador seco (bloque "Trafo Seco"): caja con dos bobinas de cinco arcos; origen = centro del borde superior
    TRAFO_SECO: {
      capa: 'TRAFO', k: 1.8, alto: 16.8,
      prims: [['L', -8.82, 0, 8.86, 0], ['L', 8.86, 0, 8.86, -16.8], ['L', 8.86, -16.8, -8.82, -16.8], ['L', -8.82, -16.8, -8.82, 0], ['L', 0, 3, 0, 0], ['L', 0, -16.8, 0, -19.8]]
        .concat([-5.41, -2.7, 0, 2.7, 5.41].map(x => ['A', x, -3.35, 1.35, 180, 0]), [-5.41, -2.7, 0, 2.7, 5.41].map(x => ['A', x, -13.45, 1.35, 0, 180]),
          [['L', -6.76, -5.84, 6.76, -5.84], ['L', -6.76, -7.86, 6.76, -7.86], ['L', -6.76, -10.96, 6.76, -10.96], ['L', -6.76, -8.94, 6.76, -8.94]]),
      attrs: [['KVA', 'Capacidad (kVA)', 0, -22, 1.6, 0], ['IMPEDANCIA', 'Impedancia (%)', 0, -24.6, 1.6, 0], ['VOLTAJES', 'Voltajes', 0, -27.2, 1.6, 0]],
    },
    // Transferencia automática (bloque "ATS"): normal arriba, alterna a la derecha, carga abajo
    ATS: {
      capa: 'EQUIPOS', k: 1.8, alto: 20,
      prims: [['L', -10, 0, 10, 0], ['L', 10, 0, 10, -20], ['L', 10, -20, -10, -20], ['L', -10, -20, -10, 0],
        ['C', 3.03, -10, 0.72], ['C', -3.03, -10, 0.72], ['C', 0, -13.03, 0.72], ['A', 0, -10, 2.31, 49.7, 130.3], ['L', 0, -12.31, 0, -7.69],
        ['L', -3.03, -9.28, -3.03, -4], ['L', -3.03, -4, 0, -4], ['L', 0, -4, 0, 0], ['L', 3.75, -10, 10, -10], ['L', 0, -13.75, 0, -20]],
      attrs: [['MODELO', 'Modelo del ATS', 12, -4, 1.6, 0], ['AMPERIOS', 'Amperios', 12, -6.6, 1.6, 0], ['WCR', 'WCR (kA)', 12, -9.2, 1.6, 0], ['TRANSICION', 'Transición', 12, -11.8, 1.6, 0]],
      puertos: { normal: [0, 0], alterna: [10, -10], carga: [0, -20] },
    },
    // Transferencia manual (bloque "MTS")
    MTS: {
      capa: 'EQUIPOS', k: 2.2, alto: 16.99,
      prims: [['L', -6.56, 0, 6.56, 0], ['L', 6.56, 0, 6.56, -16.99], ['L', 6.56, -16.99, -6.56, -16.99], ['L', -6.56, -16.99, -6.56, 0],
        ['C', -3.11, -4.45, 0.79], ['C', 3.11, -4.45, 0.79], ['C', 0, -12.55, 0.79], ['A', 0, -6.09, 1.23, 42.3, 137.7],
        ['L', -0.91, -5.26, 0, -11.76], ['L', 0, -11.76, 0.91, -5.26], ['L', -3.11, -3.66, -3.11, -1.54], ['L', -3.11, -1.54, 0, -1.54], ['L', 0, -1.54, 0, 0], ['L', 3.9, -4.45, 6.56, -4.45], ['L', 0, -13.34, 0, -16.99]],
      attrs: [['MODELO', 'Modelo del MTS', 8, -4, 1.6, 0], ['AMPERIOS', 'Amperios', 8, -6.6, 1.6, 0], ['WCR', 'WCR (kA)', 8, -9.2, 1.6, 0]],
      puertos: { normal: [0, 0], alterna: [6.56, -4.45], carga: [0, -16.99] },
    },
    // Generador: círculo con G; el origen es su terminal inferior
    GENERADOR: { capa: 'EQUIPOS', k: 3, alto: 10, prims: [['C', 0, 5, 5], ['T', 0, 5, 4.6, 'G']], attrs: [['NOMBRE', 'Nombre', 7, 6, 1.2, 0], ['KVA', 'Capacidad (kVA)', 7, 3.6, 1.2, 0], ['XD', "Reactancia X''d (%)", 7, 1.2, 1.2, 0]] },
    TIERRA: { capa: 'EQUIPOS', k: 2, alto: 5.9, prims: [['L', 0, 0, 0, -2.8], ['L', -2.2, -2.8, 2.2, -2.8], ['L', -1.6, -4.4, 1.6, -4.4], ['L', -0.95, -5.9, 0.95, -5.9]] },
  };
  g.SIMB = SIMB;
})(window);
