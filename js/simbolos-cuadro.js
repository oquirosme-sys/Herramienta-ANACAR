/* Símbolos del bloque "Cantidad de dispositivos" del cuadro de cargas de Panamá (hoja TL4).
   Se dibujan como SVG en pantalla (cuadro y memoria de cálculo) y se rasterizan a PNG al exportar a Excel.
   `disp` define, de izquierda a derecha, el campo del circuito donde se escribe cada cantidad. */
(function (g) {
  'use strict';
  const L = (a, b, c, d) => '<line x1="' + a + '" y1="' + b + '" x2="' + c + '" y2="' + d + '"/>';
  const CLIP = (id, cx, cy, r) => '<clipPath id="' + id + '"><circle cx="' + cx + '" cy="' + cy + '" r="' + r + '"/></clipPath>';
  const DIB = {
    s: '<text x="20" y="40" font-family="Century Gothic, Arial, sans-serif" font-size="34" text-anchor="middle" fill="#000" stroke="none">S</text>',
    spot: '<circle cx="20" cy="30" r="9"/>' + L(3, 30, 37, 30) + L(20, 13, 20, 47),
    flechas: '<polygon points="3,22 10,10 17,22"/><polygon points="23,22 30,10 37,22"/>' + L(10, 22, 10, 48) + L(30, 22, 30, 48) + L(3, 48, 37, 48) + '<circle cx="20" cy="38" r="8"/>' + L(17, 22, 17, 38) + L(23, 22, 23, 38),
    toma: '<circle cx="20" cy="38" r="10"/>' + L(16.5, 14, 16.5, 48) + L(23.5, 14, 23.5, 48),
    gfci: CLIP('cg', 20, 38, 10) + '<circle cx="20" cy="38" r="10"/><rect x="16.5" y="28" width="7" height="20" fill="#000" clip-path="url(#cg)"/>' + L(16.5, 14, 16.5, 28) + L(23.5, 14, 23.5, 28),
    esp: CLIP('ce', 20, 38, 10) + '<circle cx="20" cy="38" r="10"/><g fill="#000" clip-path="url(#ce)"><rect x="11" y="28" width="3" height="20"/><rect x="16" y="28" width="3" height="20"/><rect x="21" y="28" width="3" height="20"/><rect x="26" y="28" width="3" height="20"/></g>' + L(16.5, 14, 16.5, 28) + L(23.5, 14, 23.5, 28),
    triangulo: '<circle cx="20" cy="30" r="12"/><polygon points="20,20 29,37 11,37" fill="#000"/>',
    caja: '<rect x="4" y="24" width="24" height="15"/><polyline points="28,34 35,34 35,40"/>',
    rayo: '<circle cx="20" cy="30" r="10"/>' + L(4, 47, 36, 13),
  };
  /** Dispositivos por circuito, de izquierda a derecha: campo, símbolo y nombre. */
  const disp = [
    { campo: 'dApag', clave: 's', nombre: 'Apagadores' }, { campo: 'dSpot', clave: 'spot', nombre: 'Dispositivo 2' }, { campo: 'dFlechas', clave: 'flechas', nombre: 'Dispositivo 3' },
    { campo: 'dTomas', clave: 'toma', nombre: 'Tomas' }, { campo: 'dGfci', clave: 'gfci', nombre: 'Tomas GFCI' }, { campo: 'dEsp', clave: 'esp', nombre: 'Tomas especiales' },
    { campo: 'dTri', clave: 'triangulo', nombre: 'Dispositivo 7' }, { campo: 'dCaja', clave: 'caja', nombre: 'Dispositivo 8' }, { campo: 'dRayo', clave: 'rayo', nombre: 'Dispositivo 9' },
  ];
  const svg = (k, w, h) => '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 40 60"' + (w ? ' width="' + w + '" height="' + h + '"' : '') + ' fill="none" stroke="#000" stroke-width="1.5" stroke-linejoin="round" stroke-linecap="round">' + (DIB[k] || '') + '</svg>';

  /** PNG de cada símbolo para el libro de Excel: { clave: {bytes, w, h} } (vacío si el navegador no puede rasterizar). */
  async function png(claves) {
    const out = {};
    for (const k of claves) {
      try {
        const W = 120, H = 180, img = new Image();
        img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg(k, W, H));
        await img.decode();
        const cv = document.createElement('canvas'); cv.width = W; cv.height = H; cv.getContext('2d').drawImage(img, 0, 0, W, H);
        const blob = await new Promise(res => cv.toBlob(res, 'image/png'));
        out[k] = { bytes: new Uint8Array(await blob.arrayBuffer()), w: W, h: H };
      } catch (e) { /* sin imagen: la celda queda vacía */ }
    }
    return out;
  }
  g.Simbolos = { claves: Object.keys(DIB), disp, svg, png };
})(window);
