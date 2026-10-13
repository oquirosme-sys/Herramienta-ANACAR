"""Agrega al catálogo (js/data/seed.js) las palabras clave de Revit de cada detalle de carga y el detalle de deshumidificador."""
import json, sys
p = sys.argv[1]
s = open(p, encoding='utf8').read()
i, j = s.index('{'), s.rindex('}') + 1
D = json.loads(s[i:j])
CLAVES = {
    1: 'iluminacion, iluminación, luminaria, luces, lamparas, luz', 2: 'ducto ascensor, ducto de ascensor', 3: 'escalera', 4: 'iluminacion exterior, iluminación exterior, exterior',
    5: 'rotulo, rótulo, letrero', 6: 'iluminacion, iluminación, luminaria, luces', 7: 'iluminacion exterior 208',
    8: 'toma, tomas, tomacorriente, tomacorrientes, receptaculo, cortinero, cortineros, prevista',
    9: 'tomas ups, ups', 10: 'calentador de camisas, camisas', 11: 'cargador de baterias, cargador de baterías, baterias', 12: 'cargador de vehiculo, cargador de vehículo, ev, evse, vehiculo electrico',
    13: 'secamanos, secador de manos', 14: 'tomas 208, toma 208', 15: 'porton, portón, motor de porton', 16: 'equipo 208, equipos 208',
    23: 'elevador, ascensor', 24: 'ups 208', 26: 'equipo de cocina 120', 27: 'equipo de cocina 208, horno, plantilla, cocina, estufa', 28: 'equipo de cocina 480',
    29: 'campana, campana de extraccion, campana de extracción', 30: 'microondas, horno microondas, horno mw', 31: 'lavaplatos, lavavajillas',
    32: 'cafetera, maquina de cafe, máquina de café', 33: 'refrigerador, refri, nevera, wine storage, ice maker, maquina de hielo, gaveta caliente, triturador',
    34: 'control de puertas, puerta', 35: 'control de acceso, acceso', 36: 'facp, incendio, panel de incendio', 37: 'seguridad, cctv, alarma', 38: 'rack, racks, ti, datos, comunicaciones',
    39: 'pc-a', 40: 'uc-, cu-, condensadora, unidad condensadora, unidad cu', 41: 'ah-, ue, ue-, evaporadora, evaporadoras, manejadora, unidad ah', 42: 'uc-, cu-, condensadora 3f, unidad cu 3f',
    43: 'iny-, inyector', 44: 'ext-, extractor', 45: 'ms-, minisplit, mini split', 46: 'cassette, cassettes, fan coil, fancoil, fc-', 47: 'tanque de agua caliente, tanque, termo',
    48: 'calentador de paso, calentador', 49: 'bomba de agua caliente, recirculacion', 50: 'bomba diesel, trasiego', 51: 'bomba jockey, jockey', 52: 'bomba pluvial, pluvial',
    53: 'ptar, planta de tratamiento', 54: 'bomba de agua, bombas de agua, bomba de potable, hidroneumatico', 55: 'equipo mecanico, equipos mecanicos, equipo mecánico',
    56: 'esterilizadora', 57: 'lavadora de instrumentos', 58: 'equipo medico, equipos médicos', 59: 'bomba de vacio, vacío', 60: 'gases medicos, gases médicos', 61: 'compresor',
}
for d in D['detallesCarga']:
    d['claves'] = CLAVES.get(d['id'], d.get('claves', ''))
if not any(d['id'] == 67 for d in D['detallesCarga']):
    base = next(d for d in D['detallesCarga'] if d['id'] == 44)
    nuevo = dict(base); nuevo.update({'id': 67, 'tipo': 5, 'descripcion': 'DESHUMIDIFICADOR DH-XX', 'v': 120, 'fases': 1, 'hilos': 3, 'fd': 0.8, 'fdiv': 1, 'fp': 0.9, 'continua': True, 'mult': 1.25,
                  'fundamento': 'Motor/compresor de uso continuo (430.22). Demanda: equipos mecánicos.', 'claves': 'dh-, deshumidificador, deshumidificadores'})
    D['detallesCarga'].append(nuevo)
s2 = s[:i] + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + s[j:]
open(p, 'w', encoding='utf8').write(s2)
print('detalles', len(D['detallesCarga']))
