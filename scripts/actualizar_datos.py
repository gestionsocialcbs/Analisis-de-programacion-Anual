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
    Catálogo agrupado.xlsx (hoja «Agrupaciones»: grupo de cada clave de actividad por año)
    Nominales/Nominal_2021.xlsx ... (bases nominales por año; columna REF_CURP para contar personas únicas)

Requisitos: Python 3.9+, pandas y python-calamine (ver scripts/requirements.txt).
"""
import sys, re, json, datetime, gzip, base64, gc, pickle, subprocess, tempfile
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
K = ['year', 'ooad', 'uopsi', 'cec', 'area', 'tipo', 'act']
CECS = ['UOPSI', 'CEC']                          # CVE_CEC = 0: UOPSI; CVE_CEC > 0: CEC
# Perfil de la población inscrita (solo existe en el registro de inscripciones)
GENERO = ['CAN_MASCULINOS', 'CAN_FEMENINOS']
EDAD = {'jov': ['CAN_00_A_02', 'CAN_03_A_05', 'CAN_06_A_07', 'CAN_08_A_09', 'CAN_10_A_12', 'CAN_13_A_15', 'CAN_16_A_19'],
        'adu': ['CAN_20_A_59'], 'may': ['CAN_MAS_DE_60']}
CALIDAD = ['CAN_DERECHOHABIENTE', 'CAN_NODERECHOHABIENTE', 'CAN_JUBILADOIMSS', 'CAN_JUBILADONOIMSS',
           'CAN_PENSIONADOIMSS', 'CAN_PENSIONADONOIMSS', 'CAN_GRUPOREPRESENTATIVOIMSS', 'CAN_GRUPOREPRESENTATIVONOIMSS',
           'CAN_PERSONALVOLUNTARIODERECHO', 'CAN_PERSONALVOLUNTARIONODERECHO', 'CAN_TRABAJADORIMSS',
           'CAN_PERSONAL_DISC_DERER', 'CAN_PERSONAL_DISC_NODERER', 'CAN_REFERIDOS_AREA_MEDICA',
           'CAN_ADULTO_MAYOR_DERER', 'CAN_ADULTO_MAYOR_NODERER']
PERFIL = ['g0', 'g1', 'e0', 'e1', 'e2'] + [f'q{i}' for i in range(len(CALIDAD))]

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
        z = lambda c: num(d[c]) if c in d else 0
        o['g0'], o['g1'] = z(GENERO[0]), z(GENERO[1])
        for j, k in enumerate(['jov', 'adu', 'may']): o[f'e{j}'] = sum(z(c) for c in EDAD[k])
        for j, c in enumerate(CALIDAD): o[f'q{j}'] = z(c)
    o['cec'] = (num(d.CVE_CEC) > 0).astype(int) if 'CVE_CEC' in d else 0
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


# ---------------------------------------------------------------- grupos de actividad
SIN_GRUPO = 'Sin agrupación'
# claves que no están en el catálogo agrupado (el catálogo tiene prioridad si después se agregan)
GRUPO_MANUAL = {'31806': 'Curso Vacacional',               # Taller de actividad teatral Verano entre telones
                '43701': 'Personas mayores y cuidadores'}  # Curso de habilidad física (club)
def agrupaciones(carpeta, D):
    """Agrega D['grupos'] y D['F']['g'] a partir de la hoja «Agrupaciones» del Catálogo agrupado
    (columnas Agrupación, Clave, Nombre de la actividad, Año). La agrupación se asigna por año y clave;
    si una clave no aparece en ese año, se usa su agrupación más frecuente."""
    fs = [f for f in Path(carpeta).glob('*.xls*') if 'AGRUPAD' in f.stem.upper() and not f.name.startswith('~$')]
    if not fs:
        print('  Aviso: no encontré el «Catálogo agrupado»; el filtro de grupo quedará vacío.'); return
    wb = CalamineWorkbook.from_path(str(fs[0])); data = wb.get_sheet_by_name('Agrupaciones').to_python()
    hr = hc = None
    for i, r in enumerate(data[:40]):
        v = [str(x).strip() for x in r]
        for j in range(len(v) - 1):
            if v[j] == 'Agrupación' and v[j + 1] == 'Clave': hr, hc = i, j
    if hr is None: raise ValueError('No encontré las columnas «Agrupación» y «Clave» en la hoja Agrupaciones')
    por_anio, por_clave = {}, {}
    for r in data[hr + 1:]:
        g, y = str(r[hc]).strip(), r[hc + 3]; cv = re.sub(r'\.0$', '', str(r[hc + 1]).strip())
        if not g or not cv: continue
        try: y = int(float(y))
        except (TypeError, ValueError): y = None
        if y: por_anio[(y, cv)] = g
        por_clave.setdefault(cv, []).append(g)
    por_clave = {c: max(set(v), key=v.count) for c, v in por_clave.items()}
    claves = [a[0] for a in D['acts']]; F = D['F']
    asig = [por_anio.get((D['years'][y], claves[c])) or por_clave.get(claves[c]) or GRUPO_MANUAL.get(claves[c]) or SIN_GRUPO for y, c in zip(F['y'], F['c'])]
    grupos = sorted(set(asig) - {SIN_GRUPO}, key=lambda s: s.lower()) + ([SIN_GRUPO] if SIN_GRUPO in asig else [])
    gi = {g: i for i, g in enumerate(grupos)}
    D['grupos'] = grupos; F['g'] = [gi[g] for g in asig]
    sin = sorted({claves[c] for c, g in zip(F['c'], asig) if g == SIN_GRUPO})
    print(f'  Grupos de actividad: {len(grupos)}' + (f' (claves sin agrupación: {", ".join(sin)})' if sin else ''))


# ---------------------------------------------------------------- usuarios nominales
NOM_COLS = ['CVE_OOAD', 'DES_OOAD', 'CVE_UOPSI', 'DES_UOPSI', 'CVE_CEC', 'DES_AREA', 'CVE_ACTIVIDAD', 'DES_ACTIVIDAD',
            'CVE_ALUMNO', 'NOM_ALUMNO', 'NOM_APELLIDO_PATERNO', 'NOM_APELLIDO_MATERNO', 'FEC_NACIMIENTO', 'REF_CURP',
            'DES_SEXO', 'EDAD_CALCULADA', 'EDAD', 'CAN_EDAD', 'NOM_CONDICION', 'NOM_TIPO_USUARIO']
# calidad jurídica nominal -> mismas 16 categorías del registro de inscripciones (q0..q15); (condición, ¿derechohabiente?)
CJ_NOM = {'DERECHOHABIENTE': (0, 0), 'NO DERECHOHABIENTE': (1, 1), 'JUBILADO IMSS': (2, 2), 'JUBILADO NO IMSS': (3, 3),
          'PENSIONADO IMSS': (4, 4), 'PENSIONADO NO IMSS': (5, 5), 'GRUPO REPRESENTATIVO IMSS': (6, 6), 'GRUPO REPRESENTATIVO NO IMSS': (7, 7),
          'PERSONAL VOLUNTARIO': (8, 9), 'TRABAJADOR IMSS (PREST. SOC.)': (10, 10), 'PERSONA CON DISCAPACIDAD': (11, 12),
          'PACIENTE REFERIDO DEL AREA MEDICA': (13, 13), 'PERSONA ADULTA MAYOR': (14, 15)}
NPF = 2 + 3 + 16                                            # columnas del perfil nominal: sexo, grupo etario, calidad jurídica

def leer_nominal(path):
    """Une todas las hojas de la base nominal que tengan REF_CURP y CVE_ACTIVIDAD (una fila por inscripción)."""
    wb = CalamineWorkbook.from_path(str(path)); partes = []
    for s in wb.sheet_names:
        hdr = idx = None; rows = []
        for n, r in enumerate(wb.get_sheet_by_name(s).iter_rows()):
            if hdr is None:
                v = ['CVE_OOAD' if str(x).strip().upper() == 'CVE_DELEGACION' else 'DES_OOAD' if str(x).strip().upper() == 'DES_DELEGACION'
                     else str(x).strip().upper() for x in r]
                if 'REF_CURP' in v and 'CVE_ACTIVIDAD' in v: hdr = v; idx = [v.index(c) if c in v else None for c in NOM_COLS]
                elif n > 30: break
                continue
            rows.append([r[i] if i is not None and i < len(r) else '' for i in idx])
        if hdr: partes.append(pd.DataFrame(rows, columns=NOM_COLS, dtype=object)); del rows
    if not partes: raise ValueError(f'No encontré la hoja nominal en {path.name}')
    return pd.concat(partes, ignore_index=True)

def _limpia(s):
    s = s.astype(str).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)
    return s.where(~s.isin(['', 'NAN', 'NONE', 'NULL', 'NAT']), '')

def personas(d):
    """Identificador de persona por inscripción. Se usa REF_CURP; si está vacío, se busca el CURP de la misma
    persona por clave de alumno y, después, por nombre completo y fecha de nacimiento. Si no existe, se le asigna
    un identificador propio (clave de alumno, nombre y fecha, o la inscripción misma) para que cuente."""
    curp = _limpia(d.REF_CURP); alu = _limpia(d.CVE_ALUMNO)
    nom = (_limpia(d.NOM_ALUMNO) + '|' + _limpia(d.NOM_APELLIDO_PATERNO) + '|' + _limpia(d.NOM_APELLIDO_MATERNO) + '|' + _limpia(d.FEC_NACIMIENTO).str[:10])
    nom = nom.where(~nom.str.startswith('|'), '')
    vacios = int((curp == '').sum()); unicos = int(curp[curp != ''].nunique())
    for llave in (alu, nom):
        con = (curp != '') & (llave != '')
        m = pd.Series(curp[con].values, index=llave[con].values); m = m[~m.index.duplicated()]
        f = (curp == '') & (llave != ''); curp[f] = llave[f].map(m).fillna('')
    sin = curp == ''
    propio = ('ALU:' + alu).where(alu != '', ('NOM:' + nom).where(nom != '', 'INS:' + pd.Series(range(len(d)), index=d.index).astype(str)))
    curp[sin] = propio[sin]
    st = {'inscripciones': len(d), 'curp_unicos': unicos, 'curp_vacios': vacios, 'curp_recuperados': vacios - int(sin.sum()),
          'id_asignados': int(sin.sum()), 'usuarios': int(curp.nunique())}
    return curp, st

def atributos(d):
    """Sexo (0 hombre, 1 mujer), grupo etario (0: 0-19, 1: 20-59, 2: 60+) y calidad jurídica (0-15); -1 = sin dato.
    La edad se toma de EDAD_CALCULADA (si no existe, de EDAD o CAN_EDAD)."""
    sx = _limpia(d.DES_SEXO).str[:1].map({'M': 0, 'H': 0, 'F': 1}).fillna(-1).astype(int)
    ed = num(d.EDAD_CALCULADA) if (_limpia(d.EDAD_CALCULADA) != '').any() else num(d.EDAD) if (_limpia(d.EDAD) != '').any() else num(d.CAN_EDAD)
    vac = _limpia(d.EDAD_CALCULADA if (_limpia(d.EDAD_CALCULADA) != '').any() else d.EDAD if (_limpia(d.EDAD) != '').any() else d.CAN_EDAD) == ''
    ag = pd.Series(np.select([ed < 0, ed < 20, ed < 60, ed <= 110], [-1, 0, 1, 2], -1), index=d.index).where(~vac, -1)
    cond = _limpia(d.NOM_CONDICION); dh = _limpia(d.NOM_TIPO_USUARIO).eq('DERECHOHABIENTE')
    par = cond.map(CJ_NOM)
    cj = pd.Series([(-1 if not isinstance(v, tuple) else (v[0] if h else v[1])) for v, h in zip(par, dh)], index=d.index)
    return sx, ag, cj

def moda(pid, v):
    """Valor más frecuente por persona (sin contar -1)."""
    t = pd.DataFrame({'p': pid, 'v': v}); t = t[t.v >= 0]
    t = t.groupby(['p', 'v']).size().reset_index(name='n').sort_values(['p', 'n'], ascending=[True, False]).drop_duplicates('p')
    return pd.Series(t.v.values, index=t.p.values)

def procesar_nominal(f, y):
    """Lee un archivo nominal y lo reduce a los pares únicos persona-renglón del tablero, más el perfil de cada persona."""
    d = leer_nominal(f); d = d[code(d.CVE_ACTIVIDAD) != ''].reset_index(drop=True)      # quita notas o renglones vacíos
    pid, st = personas(d); sx, ag, cj = atributos(d)
    o = pd.DataFrame({'ooad': code(d.CVE_OOAD).str.zfill(2), 'ooad_n': d.DES_OOAD.astype(str).str.strip().str.upper(),
                      'uopsi': code(d.CVE_UOPSI), 'uopsi_n': d.DES_UOPSI.astype(str).str.strip().str.upper(),
                      'cec': (num(d.CVE_CEC) > 0).astype(int),
                      'area': d.DES_AREA.astype(str).str.strip().str.upper().replace(AREA_EQUIV),
                      'act': code(d.CVE_ACTIVIDAD), 'act_n': d.DES_ACTIVIDAD.astype(str).str.strip().str.upper(), 'pid': pid.values})
    del d; gc.collect()
    o['year'] = y; o['tipo'] = o.act.str[2].map(TIPO_POR_DIGITO).fillna('Otro')
    ok = o.act.ne('') & o.ooad.ne('00') & o.uopsi.ne('') & o.tipo.ne('Otro') & o.area.isin(AREAS)
    st['fuera_de_catalogo'] = int((~ok).sum())
    pf = pd.DataFrame({'sx': moda(o.pid.values, sx.values), 'ag': moda(o.pid.values, ag.values), 'cj': moda(o.pid.values, cj.values)})
    o = o[ok]; st['usuarios_tablero'] = int(o.pid.nunique())
    o = o.drop_duplicates(['pid'] + K).reset_index(drop=True)
    cat = o.pid.astype('category'); o['pid'] = cat.cat.codes.astype(np.int32)
    pf = pf.reindex(cat.cat.categories).fillna(-1).astype(np.int8).reset_index(drop=True)   # una fila por persona (código pid)
    st.update({'sin_sexo': int((pf.sx < 0).sum()), 'sin_edad': int((pf.ag < 0).sum()), 'sin_calidad': int((pf.cj < 0).sum())})
    for c in ['ooad', 'ooad_n', 'uopsi', 'uopsi_n', 'area', 'act', 'act_n', 'tipo']: o[c] = o[c].astype('category')
    return o, st, pf

def leer_nominales(carpeta):
    """Lee la carpeta «Nominales» (un archivo por año). Cada archivo se procesa en un proceso aparte para
    liberar la memoria de la lectura del Excel."""
    dirn = next((p for p in Path(carpeta).iterdir() if p.is_dir() and p.name.upper().startswith('NOMINAL')), None)
    if dirn is None: print('  Aviso: no encontré la carpeta «Nominales».'); return {}
    out = {}; tmp = Path(tempfile.mkdtemp())
    for f in sorted(dirn.glob('*.xls*')):
        if f.name.startswith('~$'): continue
        a = re.findall(r'(20\d\d)', f.stem)
        if len(a) != 1: continue
        y = int(a[0]); print(f'  Leyendo nominal {y} ({f.name})…', flush=True); pk = tmp / f'nom_{y}.pkl'
        subprocess.run([sys.executable, str(Path(__file__).resolve()), '--nominal', str(f), str(y), str(pk)], check=True)
        out[y] = pickle.load(open(pk, 'rb')); pk.unlink(); print(f'    {out[y][1]}', flush=True)
        for c in ['ooad', 'ooad_n', 'uopsi', 'uopsi_n', 'area', 'act', 'act_n', 'tipo']: out[y][0][c] = out[y][0][c].astype(str)
    return out

def codificar_nominales(Fy, nom, pf):
    """Fy: renglones del año (en orden). Devuelve n1 (personas con un solo renglón), su perfil por renglón,
    la lista de personas con varios renglones y el perfil de cada una de ellas."""
    llave = pd.MultiIndex.from_frame(Fy[K]); pos = pd.Series(np.arange(len(Fy)), index=llave)
    loc = pos.reindex(pd.MultiIndex.from_frame(nom[K])).values
    if np.isnan(loc).any(): raise ValueError('Hay renglones nominales sin renglón en el tablero')
    t = pd.DataFrame({'p': nom.pid.values, 'r': loc.astype(np.int64)}).drop_duplicates().sort_values(['p', 'r'])
    c = t.groupby('p').r.size(); uno = t[t.p.map(c) == 1]
    n1 = np.bincount(uno.r.values, minlength=len(Fy))
    perf = np.zeros((len(Fy), NPF), dtype=np.int64); a = pf.iloc[uno.p.values]
    for col, base in (('sx', 0), ('ag', 2), ('cj', 5)):
        v = a[col].values; m = v >= 0; np.add.at(perf, (uno.r.values[m], base + v[m]), 1)
    varios = t[t.p.map(c) > 1]; grupos = varios.groupby('p').r.apply(list)
    if len(Fy) >= 65536: raise ValueError('Demasiados renglones en el año para codificar en 16 bits')
    flujo = [len(grupos) >> 16, len(grupos) & 0xFFFF]
    for g in grupos: flujo.append(len(g)); flujo.extend(g)
    attrs = pf.iloc[grupos.index.values][['sx', 'ag', 'cj']].values.astype(np.int8).ravel()
    return n1, perf, flujo, attrs

# ---------------------------------------------------------------- proceso
def main(carpeta, salida):
    years, paa, ins = archivos(carpeta); LAST = years[-1]
    print('Años encontrados:', years)
    NOM = leer_nominales(carpeta); gc.collect()
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
        # perfil: año completo; en el último año, solo hasta el mes de corte (igual que el resto del tablero)
        Ip = I[I.mes.between(1, CUT)] if y == LAST else I
        pf = Ip.groupby(K)[PERFIL].sum()
        Fy = p.join(r, how='outer').join(pf, how='left')
        if y in NOM:                                                    # renglones que solo existen en la base nominal
            nk = pd.MultiIndex.from_frame(NOM[y][0][K].drop_duplicates()); Fy = Fy.reindex(Fy.index.union(nk))
        F.append(Fy.fillna(0).reset_index())
        G = ['year', 'ooad', 'uopsi', 'cec', 'area', 'tipo']
        pm = pd.DataFrame(am, columns=range(1, 13)); pm[G] = P[G].values
        pm = pm.melt(id_vars=G, var_name='mes', value_name='pa').groupby(G + ['mes']).pa.sum()
        pim = P.rename(columns={'sm': 'mes2'}).groupby(G + ['mes2']).ins.sum().rename('pi'); pim.index = pim.index.set_names(G + ['mes'])
        rm = I[I.mes.between(1, 12)].groupby(G + ['mes']).agg(ri=('ins', 'sum'), ra=('asi', 'sum'))
        M.append(pd.concat([pm, pim, rm], axis=1).fillna(0).reset_index())
        N.append(pd.concat([P, I] + ([NOM[y][0]] if y in NOM else []))[['year', 'ooad', 'ooad_n', 'uopsi', 'uopsi_n', 'act', 'act_n', 'area']])
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
             'a': df.area.map(ai).tolist(), 't': df.tipo.map(ti).tolist(), 'k': df.cec.astype(int).tolist()}
        for k in keys: o[k] = [int(round(v)) for v in df[k]]
        return o
    fact = cols(F, ['pi', 'pa', 'ri', 'ra', 'pic', 'pac', 'ric', 'rac'] + PERFIL); fact['c'] = F.act.map(ci).tolist()
    # usuarios nominales: n1 por renglón + personas con varios renglones (índices locales al año), comprimido
    n1 = np.zeros(len(F), dtype=np.int64); perf = np.zeros((len(F), NPF), dtype=np.int64); flujo = []; attrs = []
    nom_years = []; ini = 0; F = F.reset_index(drop=True)
    for y in years:
        m = (F.year == y).values; idx = np.flatnonzero(m)
        assert len(idx) == 0 or (idx[0] == ini and idx[-1] == ini + len(idx) - 1), 'F debe estar ordenado por año'
        if y in NOM:
            a, pr, b, at = codificar_nominales(F.iloc[idx], NOM[y][0], NOM[y][2]); n1[idx] = a; perf[idx] = pr
            flujo += b; attrs.append(at); nom_years.append(y)
        ini += len(idx)
    fact['n1'] = n1.tolist()
    nom_info = {'years': nom_years, 'stats': {str(y): NOM[y][1] for y in nom_years},
                'multi': base64.b64encode(gzip.compress(np.array(flujo, dtype='<u2').tobytes(), 9)).decode(),
                # perfil nominal: por renglón, personas con un solo renglón (21 columnas: 2 sexo, 3 edad, 16 calidad), por columnas
                'perfil': base64.b64encode(gzip.compress(np.ascontiguousarray(perf.T).astype('<i4').tobytes(), 9)).decode(),
                # perfil de cada persona con varios renglones (sexo, edad, calidad; -1 = sin dato), en el orden de «multi»
                'attrs': base64.b64encode(gzip.compress((np.concatenate(attrs) if attrs else np.zeros(0)).astype(np.int8).tobytes(), 9)).decode()}
    mon = cols(M, ['pi', 'pa', 'ri', 'ra']); mon['m'] = M.mes.tolist()
    D = {'years': years, 'cut': CUT, 'corte': corte, 'generado': datetime.date.today().isoformat(),
         'ooads': [[c, oo[c]] for c in ooads], 'uopsis': [[oi[o], u, uo[(o, u)]] for o, u in ukeys],
         'areas': AREAS, 'tipos': TIPOS, 'cecs': CECS, 'acts': [[c, an[c], ai[actarea[c]]] for c in acts], 'F': fact, 'M': mon, 'nom': nom_info}
    agrupaciones(carpeta, D)
    Path(salida).parent.mkdir(parents=True, exist_ok=True)
    Path(salida).write_text('window.PAA_DATA=' + json.dumps(D, ensure_ascii=False, separators=(',', ':')) + ';\n', encoding='utf-8')
    print(f'Listo: {salida} ({Path(salida).stat().st_size/1e6:.1f} MB)')

if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == '--nominal':           # uso interno: un archivo nominal
        pickle.dump(procesar_nominal(Path(sys.argv[2]), int(sys.argv[3])), open(sys.argv[4], 'wb')); sys.exit()
    if len(sys.argv) < 2: raise SystemExit(__doc__)
    raiz = Path(__file__).resolve().parent.parent
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else raiz / 'datos' / 'paa_datos.js')
