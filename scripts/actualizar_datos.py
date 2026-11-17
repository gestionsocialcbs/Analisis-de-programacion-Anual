"""
Actualiza los datos del Tablero de Control del PAA.

Lee los archivos de Excel del PAA (programado) y del registro de inscripciones
(realizado), los consolida por año, OOAD, UOPSI, área normativa, tipo y clave de
actividad, y escribe el archivo datos/paa_datos.js que usa index.html.

Uso:
    python scripts/actualizar_datos.py "C:/ruta/a/Bases claude"

La carpeta debe contener un archivo por año con estos nombres (o parecidos):
    PAA_2019.xlsx ... PAA_2026.xlsx
    Inscritos2019.xlsx / Inscritos_2020.xlsx ... Inscritos_2026.xlsx

Requisitos: Python 3.9+, pandas y python-calamine (ver scripts/requirements.txt).
"""
import sys, re, json, datetime
from pathlib import Path
import numpy as np
import pandas as pd
from python_calamine import CalamineWorkbook

# ---------------------------------------------------------------- reglas de negocio
AREAS = ['PROMOCIÓN DE LA SALUD', 'DESARROLLO CULTURAL', 'CULTURA FÍSICA Y DEPORTE',
         'CAPACITACIÓN Y ADIESTRAMIENTO TÉCNICO', 'ATENCIÓN A PERSONAS CON DISCAPACIDAD', 'SERVICIOS DE ENLACE']
AREA_EQUIV = {  # áreas que se integran a otra
    'CASSAAM': 'PROMOCIÓN DE LA SALUD',
    'ATENCIÓN A PERSONAS CON DISCAPACIDAD': 'PROMOCIÓN DE LA SALUD',
    'SERVICIOS DE ENLACE': 'PROMOCIÓN DE LA SALUD',
    'CECART': 'CAPACITACIÓN Y ADIESTRAMIENTO TÉCNICO',
    'CAPACITACION Y ADIESTRAMIENTO TECNICO': 'CAPACITACIÓN Y ADIESTRAMIENTO TÉCNICO',
}
TIPOS = ['Curso', 'Taller']                       # tercer dígito de CVE_ACTIVIDAD: 1 curso; 7 y 8 taller
TIPO_POR_DIGITO = {'1': 'Curso', '7': 'Taller', '8': 'Taller'}
K = ['year', 'ooad', 'uopsi', 'area', 'tipo', 'act']

# ---------------------------------------------------------------- lectura
def leer_excel(path):
    """Devuelve la hoja con más renglones que tenga las columnas CVE_ACTIVIDAD y CVE_GRUPO."""
    wb = CalamineWorkbook.from_path(str(path)); mejor = None
    for s in wb.sheet_names:
        sh = wb.get_sheet_by_name(s)
        for i, r in enumerate(sh.to_python(nrows=40)):
            v = [str(x).strip() for x in r]
            if 'CVE_ACTIVIDAD' in v and 'CVE_GRUPO' in v:
                if mejor is None or sh.height > mejor[2]: mejor = (s, i, sh.height)
                break
    if mejor is None: raise ValueError(f'No encontré el encabezado en {path.name}')
    data = wb.get_sheet_by_name(mejor[0]).to_python()
    hdr = [str(x).strip() for x in data[mejor[1]]]
    df = pd.DataFrame(data[mejor[1] + 1:], columns=hdr)
    return df.loc[:, [c for c in df.columns if c]].astype(str)

def code(s):
    s = s.astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return s.where(~s.isin(['', 'nan', 'None']), '')
def num(s):
    return pd.to_numeric(s.astype(str).str.strip().replace({'': '0', 'nan': '0'}), errors='coerce').fillna(0)

def normalizar(d, kind, y):
    d = d.rename(columns={'CVE_DELEGACION': 'CVE_OOAD', 'DES_DELEGACION': 'DES_OOAD'})
    o = pd.DataFrame({
        'ooad': code(d.CVE_OOAD).str.zfill(2), 'ooad_n': d.DES_OOAD.str.strip().str.upper(),
        'uopsi': code(d.CVE_UOPSI), 'uopsi_n': d.DES_UOPSI.str.strip().str.upper(),
        'area': d.DES_AREA.astype(str).str.strip().str.upper().replace(AREA_EQUIV),
        'act': code(d.CVE_ACTIVIDAD), 'act_n': d.DES_ACTIVIDAD.str.strip().str.upper(),
        'mes': num(d.CVE_MES).astype(int)})
    if kind == 'P':
        o['ins'] = num(d.CAN_CUPO); o['asi'] = num(d.CAN_ASISTENCIAS)
        o['_ini'] = d.FEC_FECHA_INICIO; o['_fin'] = d.FEC_FECHA_TERMINO
    else:
        o['ins'] = num(d.CAN_INSCRITOS); o['asi'] = num(d.CAN_ASISTENCIAS)
    o['year'] = y
    o = o[o.act.ne('') & o.ooad.ne('00') & o.uopsi.ne('')]          # excluye renglones de totales
    o['tipo'] = o.act.str[2].map(TIPO_POR_DIGITO).fillna('Otro')
    return o[o.tipo.ne('Otro') & o.area.isin(AREAS)].reset_index(drop=True)

def prorratear(P, y):
    """Reparte las asistencias programadas entre los meses que dura cada apertura, por días."""
    s = pd.to_datetime(P['_ini'], errors='coerce'); e = pd.to_datetime(P['_fin'], errors='coerce')
    y0, y1 = pd.Timestamp(y, 1, 1), pd.Timestamp(y, 12, 31)
    s = s.fillna(y0).clip(y0, y1); e = e.fillna(s).clip(y0, y1); e = e.where(e >= s, s)
    tot = (e - s).dt.days + 1; out = np.zeros((len(P), 12))
    for m in range(1, 13):
        ms = pd.Timestamp(y, m, 1); me = ms + pd.offsets.MonthEnd(0)
        out[:, m - 1] = np.clip((np.minimum(e, me) - np.maximum(s, ms)).dt.days + 1, 0, None) / tot
    return out * P['asi'].values[:, None], s.dt.month.values

def archivos(carpeta):
    paa, ins = {}, {}
    for f in sorted(Path(carpeta).glob('*.xls*')):
        anios = re.findall(r'(20\d\d)', f.stem)
        if len(anios) != 1: continue                                   # ignora acumulados (p. ej. 2019_2026)
        y = int(anios[0]); n = f.stem.upper()
        if n.startswith('PAA'): paa[y] = f
        elif n.startswith('INSCRI'): ins[y] = f
    years = sorted(set(paa) & set(ins))
    if not years: raise SystemExit('No encontré pares PAA / Inscritos por año en la carpeta.')
    return years, paa, ins

# ---------------------------------------------------------------- proceso
def main(carpeta, salida):
    years, paa, ins = archivos(carpeta); LAST = years[-1]
    print('Años encontrados:', years)
    datos = {}
    for y in years:
        print(f'  Leyendo {y}…', flush=True)
        datos[y] = (normalizar(leer_excel(paa[y]), 'P', y), normalizar(leer_excel(ins[y]), 'I', y))
    # mes de corte: último mes del último año con inscripciones realizadas
    I_last = datos[LAST][1]; por_mes = I_last[I_last.mes.between(1, 12)].groupby('mes').ins.sum()
    CUT = int(por_mes[por_mes > 0].index.max()) if (por_mes > 0).any() else 12
    corte = datetime.date.fromtimestamp(ins[LAST].stat().st_mtime).strftime('%d/%m/%Y')
    print(f'Último año: {LAST}; datos realizados hasta el mes {CUT}; archivo del {corte}')
    F, M, N = [], [], []
    for y in years:
        P, I = datos[y]
        am, sm = prorratear(P, y); P['sm'] = sm
        P['pic'] = P.ins.where(P.sm <= CUT, 0); P['pac'] = am[:, :CUT].sum(1)
        I['ric'] = I.ins.where(I.mes.between(1, CUT), 0); I['rac'] = I.asi.where(I.mes.between(1, CUT), 0)
        p = P.groupby(K).agg(pi=('ins', 'sum'), pa=('asi', 'sum'), pic=('pic', 'sum'), pac=('pac', 'sum'))
        r = I.groupby(K).agg(ri=('ins', 'sum'), ra=('asi', 'sum'), ric=('ric', 'sum'), rac=('rac', 'sum'))
        F.append(p.join(r, how='outer').fillna(0).reset_index())
        G = ['year', 'ooad', 'uopsi', 'area', 'tipo']
        pm = pd.DataFrame(am, columns=range(1, 13)); pm[G] = P[G].values
        pm = pm.melt(id_vars=G, var_name='mes', value_name='pa').groupby(G + ['mes']).pa.sum()
        pim = P.rename(columns={'sm': 'mes2'}).groupby(G + ['mes2']).ins.sum().rename('pi'); pim.index = pim.index.set_names(G + ['mes'])
        rm = I[I.mes.between(1, 12)].groupby(G + ['mes']).agg(ri=('ins', 'sum'), ra=('asi', 'sum'))
        M.append(pd.concat([pm, pim, rm], axis=1).fillna(0).reset_index())
        N.append(pd.concat([P, I])[['year', 'ooad', 'ooad_n', 'uopsi', 'uopsi_n', 'act', 'act_n', 'area']])
        print(f'  {y}: programado {int(F[-1].pi.sum()):,} · realizado {int(F[-1].ri.sum()):,} inscripciones')
    F = pd.concat(F); M = pd.concat(M); N = pd.concat(N).sort_values('year')
    M['mes'] = M.mes.astype(int)
    # catálogos (el nombre más reciente gana)
    oo = N.groupby('ooad').ooad_n.last(); uo = N.groupby(['ooad', 'uopsi']).uopsi_n.last(); an = N.groupby('act').act_n.last()
    ooads = sorted(F.ooad.unique(), key=lambda c: oo[c]); oi = {c: i for i, c in enumerate(ooads)}
    ukeys = sorted(F[['ooad', 'uopsi']].drop_duplicates().itertuples(index=False), key=lambda r: (oo[r[0]], uo[(r[0], r[1])]))
    ui = {(r[0], r[1]): i for i, r in enumerate(ukeys)}
    ai = {a: i for i, a in enumerate(AREAS)}; ti = {t: i for i, t in enumerate(TIPOS)}
    acts = sorted(F.act.unique(), key=lambda c: (an[c], c)); ci = {c: i for i, c in enumerate(acts)}
    actarea = F.groupby('act').area.agg(lambda s: s.value_counts().index[0])
    def cols(df, keys):
        o = {'y': (df.year - years[0]).tolist(), 'u': [ui[(a, b)] for a, b in zip(df.ooad, df.uopsi)],
             'a': df.area.map(ai).tolist(), 't': df.tipo.map(ti).tolist()}
        for k in keys: o[k] = [int(round(v)) for v in df[k]]
        return o
    fact = cols(F, ['pi', 'pa', 'ri', 'ra', 'pic', 'pac', 'ric', 'rac']); fact['c'] = F.act.map(ci).tolist()
    mon = cols(M, ['pi', 'pa', 'ri', 'ra']); mon['m'] = M.mes.tolist()
    D = {'years': years, 'cut': CUT, 'corte': corte, 'generado': datetime.date.today().isoformat(),
         'ooads': [[c, oo[c]] for c in ooads], 'uopsis': [[oi[o], u, uo[(o, u)]] for o, u in ukeys],
         'areas': AREAS, 'tipos': TIPOS, 'acts': [[c, an[c], ai[actarea[c]]] for c in acts], 'F': fact, 'M': mon}
    Path(salida).parent.mkdir(parents=True, exist_ok=True)
    Path(salida).write_text('window.PAA_DATA=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + ';\n', encoding='utf-8')
    print(f'Listo: {salida} ({Path(salida).stat().st_size/1e6:.1f} MB)')

if __name__ == '__main__':
    if len(sys.argv) < 2: raise SystemExit(__doc__)
    raiz = Path(__file__).resolve().parent.parent
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else raiz / 'datos' / 'paa_datos.js')
