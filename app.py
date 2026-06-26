"""
Eco SR Reader — Versão Streamlit
Lê DICOM SR Philips EPIQ, preenche formulário editável por seções e exporta CSV/Excel.

Dependências:
    pip install streamlit pydicom openpyxl
"""

import io, os, re
import streamlit as st

try:
    import pydicom
except ImportError:
    st.error("Execute: pip install pydicom"); st.stop()

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("Execute: pip install openpyxl"); st.stop()


# ═══════════════════════════════════════════════════════════════════════
# DEFINIÇÃO DO FORMULÁRIO
# ═══════════════════════════════════════════════════════════════════════

FORMULARIO = {
    "DADOS ANTROPOMÉTRICOS": [
        {"name": "Peso",             "unit": "kg",    "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "Altura",           "unit": "cm",    "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "Superfície corp.", "unit": "m²",    "ref_mas": "", "ref_fem": "", "calc": True},
        {"name": "IMC",              "unit": "kg/m²", "ref_mas": "", "ref_fem": "", "calc": True},
    ],
    "CÂMARAS ESQUERDAS": [
        {"name": "Aorta ascend.",         "unit": "mm",    "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Aorta ascend. index",   "unit": "mm/m²", "ref_mas": "",         "ref_fem": "",         "calc": True},
        {"name": "Diâm VSVE",             "unit": "mm",    "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Seio aórtico",          "unit": "mm",    "ref_mas": "31 - 37",  "ref_fem": "27 - 33",  "calc": False},
        {"name": "AE - Diâm.",            "unit": "mm",    "ref_mas": "30 - 40",  "ref_fem": "27 - 38",  "calc": False},
        {"name": "AE - Vol. bipl.",       "unit": "mL",    "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "AE - Vol. bipl. index", "unit": "mL/m²", "ref_mas": "16 - 34",  "ref_fem": "16 - 34",  "calc": True},
        {"name": "Septo",                 "unit": "mm",    "ref_mas": "6 - 10",   "ref_fem": "6 - 9",    "calc": False},
        {"name": "Parede post.",          "unit": "mm",    "ref_mas": "6 - 10",   "ref_fem": "6 - 9",    "calc": False},
        {"name": "DdVE",                  "unit": "mm",    "ref_mas": "42 - 58",  "ref_fem": "38 - 52",  "calc": False},
        {"name": "DsVE",                  "unit": "mm",    "ref_mas": "25 - 40",  "ref_fem": "22 - 35",  "calc": False},
        {"name": "VDVE",                  "unit": "mL",    "ref_mas": "62 - 150", "ref_fem": "46 - 106", "calc": True},
        {"name": "VSVE",                  "unit": "mL",    "ref_mas": "21 - 61",  "ref_fem": "14 - 42",  "calc": True},
        {"name": "VDVE index",            "unit": "mL/m²", "ref_mas": "24 - 74",  "ref_fem": "29 - 61",  "calc": True},
        {"name": "FEVE (Teichholz)",      "unit": "%",     "ref_mas": "52 - 72",  "ref_fem": "54 - 74",  "calc": True},
        {"name": "FEVE (Simpson)",        "unit": "%",     "ref_mas": "52 - 72",  "ref_fem": "54 - 74",  "calc": False},
        {"name": "ERP",                   "unit": "",      "ref_mas": "< 0,42",   "ref_fem": "< 0,42",   "calc": True},
        {"name": "Massa VE",              "unit": "g",     "ref_mas": "88 - 224", "ref_fem": "67 - 162", "calc": True},
        {"name": "Massa index",           "unit": "g/m²",  "ref_mas": "49 - 115", "ref_fem": "43 - 95",  "calc": True},
    ],
    "CÂMARAS DIREITAS": [
        {"name": "AD - Área",        "unit": "cm²",   "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "AD - Vol.",        "unit": "mL",    "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "AD - Vol. index",  "unit": "mL/m²", "ref_mas": "18 - 32", "ref_fem": "15 - 27", "calc": True},
        {"name": "AD - PSAP",        "unit": "mmHg",  "ref_mas": "<= 35",   "ref_fem": "<= 35",   "calc": False},
        {"name": "VD - Diâm. basal", "unit": "mm",    "ref_mas": "25 - 41", "ref_fem": "25 - 41", "calc": False},
        {"name": "VD - TAPSE",       "unit": "mm",    "ref_mas": ">= 17",   "ref_fem": ">= 17",   "calc": False},
        {"name": "VD - FAC",         "unit": "%",     "ref_mas": ">= 35",   "ref_fem": ">= 35",   "calc": False},
        {"name": "VD - Onda S",      "unit": "cm/s",  "ref_mas": ">= 9,5",  "ref_fem": ">= 9,5",  "calc": False},
        {"name": "VD - PSAP",        "unit": "mmHg",  "ref_mas": "<= 35",   "ref_fem": "<= 35",   "calc": False},
        {"name": "AP - Diâm.",       "unit": "mm",    "ref_mas": "",        "ref_fem": "",        "calc": False},
    ],
    "VALVA MITRAL": [
        {"name": "Vel. onda E",    "unit": "m/s",  "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Vel. onda A",    "unit": "m/s",  "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Relação E/A",    "unit": "",     "ref_mas": "",         "ref_fem": "",         "calc": True},
        {"name": "Decel. Time",    "unit": "ms",   "ref_mas": "150 - 250","ref_fem": "150 - 250","calc": False},
        {"name": "PHT",            "unit": "ms",   "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Área (PHT)",     "unit": "cm²",  "ref_mas": "",         "ref_fem": "",         "calc": True},
        {"name": "Área (PISA)",    "unit": "cm²",  "ref_mas": "",         "ref_fem": "",         "calc": True},
        {"name": "VTI mit.",       "unit": "cm",   "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Grad. máx.",     "unit": "mmHg", "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Grad. méd.",     "unit": "mmHg", "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "MAPSE",          "unit": "mm",   "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Dur. onda A",    "unit": "ms",   "ref_mas": "",         "ref_fem": "",         "calc": False},
        {"name": "Fluxo (PISA)",   "unit": "mL/s", "ref_mas": "",         "ref_fem": "",         "calc": True},
        {"name": "Volume Regurg.", "unit": "mL",   "ref_mas": "< 30",     "ref_fem": "< 30",     "calc": True},
    ],
    "AORTA / VSVE": [
        {"name": "Diâm VSVE",      "unit": "mm",    "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "VTI VSVE",       "unit": "cm",    "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "VTI Ao",         "unit": "cm",    "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "AVAo (EC-VTI)",  "unit": "cm²",   "ref_mas": "",        "ref_fem": "",        "calc": True},
        {"name": "AVAo (EC-Vmax)", "unit": "cm²",   "ref_mas": "",        "ref_fem": "",        "calc": True},
        {"name": "AVAo index",     "unit": "cm²/m²","ref_mas": "",        "ref_fem": "",        "calc": True},
        {"name": "Vel. Ratio",     "unit": "",      "ref_mas": "> 0,25",  "ref_fem": "> 0,25",  "calc": True},
        {"name": "Grad. máx.",     "unit": "mmHg",  "ref_mas": "< 20",    "ref_fem": "< 20",    "calc": False},
        {"name": "Grad. méd.",     "unit": "mmHg",  "ref_mas": "< 10",    "ref_fem": "< 10",    "calc": False},
        {"name": "Decel. Slope",   "unit": "mm/s²", "ref_mas": "",        "ref_fem": "",        "calc": False},
        {"name": "PHT Ao",         "unit": "ms",    "ref_mas": "",        "ref_fem": "",        "calc": False},
    ],
    "TRICÚSPIDE / PULMONAR": [
        {"name": "Vel. RT",    "unit": "m/s",  "ref_mas": "",      "ref_fem": "",      "calc": False},
        {"name": "PSAP",       "unit": "mmHg", "ref_mas": "<= 35", "ref_fem": "<= 35", "calc": False},
        {"name": "TAPSE",      "unit": "mm",   "ref_mas": ">= 17", "ref_fem": ">= 17", "calc": False},
        {"name": "AP - Diâm.", "unit": "mm",   "ref_mas": "",      "ref_fem": "",      "calc": False},
        {"name": "VTI AP",     "unit": "cm",   "ref_mas": "",      "ref_fem": "",      "calc": False},
    ],
    "TDI": [
        {"name": "Vel. e' septal",    "unit": "cm/s", "ref_mas": ">= 7",  "ref_fem": ">= 7",  "calc": False},
        {"name": "Rel. E/E' septal",  "unit": "",     "ref_mas": "<= 14", "ref_fem": "<= 14", "calc": True},
        {"name": "Vel. e' lateral",   "unit": "cm/s", "ref_mas": ">= 10", "ref_fem": ">= 10", "calc": False},
        {"name": "Rel. E/E' lateral", "unit": "",     "ref_mas": "<= 14", "ref_fem": "<= 14", "calc": True},
        {"name": "E/e' MÉDIO",        "unit": "",     "ref_mas": "<= 14", "ref_fem": "<= 14", "calc": True},
        {"name": "Vel. a' septal",    "unit": "cm/s", "ref_mas": "",      "ref_fem": "",      "calc": False},
        {"name": "Vel. a' lateral",   "unit": "cm/s", "ref_mas": "",      "ref_fem": "",      "calc": False},
        {"name": "E/A tecidual",      "unit": "",     "ref_mas": "",      "ref_fem": "",      "calc": True},
    ],
    "STRAIN": [
        {"name": "VE - SLG",  "unit": "%", "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "VD - SLPL", "unit": "%", "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "AE - R",    "unit": "%", "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "AE - CD",   "unit": "%", "ref_mas": "", "ref_fem": "", "calc": False},
        {"name": "AE - B",    "unit": "%", "ref_mas": "", "ref_fem": "", "calc": False},
    ],
}


# ═══════════════════════════════════════════════════════════════════════
# MAPEAMENTO DICOM SR → CAMPOS DO FORMULÁRIO
# ═══════════════════════════════════════════════════════════════════════

def _mm(v):    return round(v, 1)
def _plain(v): return round(v, 2)
def _pct(v):   return round(v, 1)

MAPA_DICOM = {
    "Patient Weight":       [("DADOS ANTROPOMÉTRICOS", "Peso", _plain)],
    "Patient Height":       [("DADOS ANTROPOMÉTRICOS", "Altura",
                              lambda v: round(v*100,1) if v < 10 else round(v,1))],
    "Body Surface Area":    [("DADOS ANTROPOMÉTRICOS", "Superfície corp.", _plain)],

    "Ascending Aortic Diameter":   [("CÂMARAS ESQUERDAS", "Aorta ascend.", _mm)],
    "Aortic Root Diameter":        [("CÂMARAS ESQUERDAS", "Seio aórtico", _mm)],
    "Cardiovascular Orifice Diameter": [
        ("CÂMARAS ESQUERDAS", "Diâm VSVE", _mm),
        ("AORTA / VSVE",      "Diâm VSVE", _mm)],
    "Left Atrium Antero-posterior Systolic Dimension": [("CÂMARAS ESQUERDAS", "AE - Diâm.", _mm)],
    "Left Atrium Systolic Volume": [("CÂMARAS ESQUERDAS", "AE - Vol. bipl.",
        lambda v: round(v/1000,1) if v > 100 else round(v,1))],
    "Left Atrium Systolic Volume Index": [("CÂMARAS ESQUERDAS", "AE - Vol. bipl. index", _plain)],
    "Interventricular Septum Diastolic Thickness": [("CÂMARAS ESQUERDAS", "Septo", _mm)],
    "Left Ventricle Posterior Wall Diastolic Thickness": [("CÂMARAS ESQUERDAS", "Parede post.", _mm)],
    "Left Ventricle Internal End Diastolic Dimension": [("CÂMARAS ESQUERDAS", "DdVE", _mm)],
    "LVIDd": [("CÂMARAS ESQUERDAS", "DdVE", _mm)],
    "Left Ventricle Internal Systolic Dimension": [("CÂMARAS ESQUERDAS", "DsVE", _mm)],
    "LVIDs": [("CÂMARAS ESQUERDAS", "DsVE", _mm)],
    "Left Ventricular Ejection Fraction": [("CÂMARAS ESQUERDAS", "FEVE (Simpson)", _pct)],
    "Relative Wall Thickness": [("CÂMARAS ESQUERDAS", "ERP", _plain)],
    "Left Ventricle Mass":         [("CÂMARAS ESQUERDAS", "Massa VE", _plain)],
    "Left Ventricle Mass by M-mode":[("CÂMARAS ESQUERDAS", "Massa VE", _plain)],

    "Right Atrium Systolic Area": [("CÂMARAS DIREITAS", "AD - Área",
        lambda v: round(v/100,1) if v > 100 else round(v,1))],
    "Right Atrium Systolic Volume": [("CÂMARAS DIREITAS", "AD - Vol.",
        lambda v: round(v/1000,1) if v > 100 else round(v,1))],
    "Right Atrium Systolic Volume Index": [("CÂMARAS DIREITAS", "AD - Vol. index", _plain)],
    "Right Atrium Systolic Pressure": [("CÂMARAS DIREITAS", "AD - PSAP", _plain)],
    "Right Ventricle Basal Diameter": [("CÂMARAS DIREITAS", "VD - Diâm. basal", _mm)],
    "Tricuspid Annular Plane Systolic Excursion": [
        ("CÂMARAS DIREITAS",       "VD - TAPSE", _mm),
        ("TRICÚSPIDE / PULMONAR",  "TAPSE",      _mm)],
    "Right Ventricle S Velocity": [("CÂMARAS DIREITAS", "VD - Onda S",
        lambda v: round(v/10,1) if v > 10 else round(v,1))],
    "Right Ventricular Peak Systolic Pressure": [("CÂMARAS DIREITAS", "VD - PSAP", _plain)],
    "Right Ventricle Outflow Tract Distal Diameter": [
        ("CÂMARAS DIREITAS",      "AP - Diâm.", _mm),
        ("TRICÚSPIDE / PULMONAR", "AP - Diâm.", _mm)],

    "Mitral Valve E-Wave Peak Velocity": [("VALVA MITRAL", "Vel. onda E",
        lambda v: round(v/1000,2) if v > 10 else round(v,2))],
    "Mitral Valve A-Wave Peak Velocity": [("VALVA MITRAL", "Vel. onda A",
        lambda v: round(v/1000,2) if v > 10 else round(v,2))],
    "Mitral Valve E to A Ratio": [("VALVA MITRAL", "Relação E/A", _plain)],
    "Deceleration Time": [("VALVA MITRAL", "Decel. Time", _plain)],
    "Pressure Half-Time": [("VALVA MITRAL", "PHT", _plain)],
    "Area by Pressure Half-Time": [("VALVA MITRAL", "Área (PHT)",
        lambda v: round(v/100,2) if v > 10 else round(v,2))],
    "Mitral Valve Flow Area": [("VALVA MITRAL", "Área (PISA)",
        lambda v: round(v/100,2) if v > 10 else round(v,2))],
    "Mitral Annular Plane Systolic Excursion": [("VALVA MITRAL", "MAPSE", _mm)],
    "Mitral Valve A-Wave Duration": [("VALVA MITRAL", "Dur. onda A", _plain)],
    "Peak Instantaneous Flow Rate": [("VALVA MITRAL", "Fluxo (PISA)",
        lambda v: round(v/1000,1) if v > 100 else round(v,1))],
    "Volume Flow": [("VALVA MITRAL", "Volume Regurg.",
        lambda v: round(v/1000,1) if v > 100 else round(v,1))],

    "Velocity Time Integral":   [("AORTA / VSVE", "_vti_ctx", _plain)],
    "Mean Velocity":            [("AORTA / VSVE", "_meanvel_ctx", _plain)],
    "Mean Gradient":            [("AORTA / VSVE", "_meangrad_ctx", _plain)],
    "Peak Velocity":            [("AORTA / VSVE", "_peakvel_ctx", _plain)],
    "Peak Gradient":            [("AORTA / VSVE", "_peakgrad_ctx", _plain)],
    "Continuity Equation by Velocity Time Integral": [("AORTA / VSVE", "AVAo (EC-VTI)",
        lambda v: round(v/100,2) if v > 10 else round(v,2))],
    "Continuity Equation by Peak Velocity": [("AORTA / VSVE", "AVAo (EC-Vmax)",
        lambda v: round(v/100,2) if v > 10 else round(v,2))],
    "Aortic Valve Area Indexed To BSA": [("AORTA / VSVE", "AVAo index", _plain)],
    "Aortic Valve Velocity Ratio": [("AORTA / VSVE", "Vel. Ratio", _plain)],
    "Cardiovascular Orifice Area": [("AORTA / VSVE", "AVAo (EC-VTI)",
        lambda v: round(v/100,2) if v > 10 else round(v,2))],
    "Deceleration Slope": [("AORTA / VSVE", "Decel. Slope", _plain)],
    "Pressure Half-Time Peak velocity": [("AORTA / VSVE", "PHT Ao", _plain)],

    "Left Ventricular Peak Early Diastolic Tissue Velocity": [("TDI", "_e_prime_raw", None)],
    "LV Peak Diastolic Tissue Velocity During Atrial Systole": [("TDI", "_a_prime_raw", None)],
    "Ratio of MV Peak Velocity to LV Peak Tissue Velocity E-Wave": [("TDI", "_e_e_prime_raw", None)],
    "Ratio of MV Peak Velocity to avg LV Peak Tissue Velocity E-Wave": [("TDI", "E/e' MÉDIO", _plain)],
    "Mean Myocardial Velocity of E' sep and E' lat": [("TDI", "E/e' MÉDIO",
        lambda v: round(v/10,1) if v > 10 else round(v,1))],
    "Left Ventricle E to A Tissue Velocity Ratio": [("TDI", "E/A tecidual", _plain)],
}

MEDIDAS_OCULTAS = {
    "Left Ventricle MOD Diam", "Left Atrium MOD Diam",
    "Right Atrium MOD Diam", "Value",
    "Left Ventricle diastolic major axis", "Left Ventricle systolic major axis",
    "Left Ventricular Diastolic Area", "Left Ventricular Systolic Area",
    "Left Atrium systolic major axis", "Left Atrium Systolic Area",
    "Right Atrium Systolic Major Axis",
    "Simpson's Disk Number", "End Diastole", "End Systole",
    "Left Atrium to Aortic Root Ratio",
    "Interventricular Septum to Posterior Wall Thickness Ratio",
}


# ═══════════════════════════════════════════════════════════════════════
# DICOM — leitura
# ═══════════════════════════════════════════════════════════════════════

SR_SOP = {
    "1.2.840.10008.5.1.4.1.1.88.11","1.2.840.10008.5.1.4.1.1.88.22",
    "1.2.840.10008.5.1.4.1.1.88.33","1.2.840.10008.5.1.4.1.1.88.34",
    "1.2.840.10008.5.1.4.1.1.88.35","1.2.840.10008.5.1.4.1.1.88.67",
    "1.2.840.10008.5.1.4.1.1.88.68","1.2.840.10008.5.1.4.1.1.88.72",
}

def eh_sr(path):
    try:
        ds = pydicom.dcmread(path, stop_before_pixels=True, force=True)
        if str(getattr(ds,"Modality","")).strip().upper() == "SR": return True
        if str(getattr(ds,"SOPClassUID","")) in SR_SOP: return True
    except: pass
    return False

def _conceito(item):
    try:
        c = item.ConceptNameCodeSequence
        return c[0].CodeMeaning if c else None
    except: return None

def _eh_finding_site(item):
    try:
        if str(getattr(item, "RelationshipType", "")).upper() == "HAS CONCEPT MOD":
            nc = _conceito(item)
            if nc == "Finding Site":
                ccs = getattr(item, "ConceptCodeSequence", None)
                if ccs:
                    return ccs[0].CodeMeaning
    except: pass
    return None

def _capturar(item, nome, ctx, site, out):
    try:
        tem_selecao = False
        if hasattr(item, "ContentSequence"):
            for sub in item.ContentSequence:
                if _conceito(sub) == "Selection Status":
                    tem_selecao = True; break
        for mv in item.MeasuredValueSequence:
            v = getattr(mv,"NumericValue",None)
            u = ""
            try: u = mv.MeasurementUnitsCodeSequence[0].CodeMeaning
            except: pass
            if v is not None:
                out.append({"nome": nome or "Value", "valor": float(v),
                            "unidade": u, "contexto": ctx, "site": site or "", "selecao": tem_selecao})
    except: pass

def extrair_raw(ds, out=None, ctx="", site_herdado=""):
    if out is None: out = []
    for elem in ds:
        if elem.VR != "SQ": continue
        site_deste_nivel = site_herdado
        label_deste_nivel = None
        for item in elem.value:
            site_encontrado = _eh_finding_site(item)
            if site_encontrado: site_deste_nivel = site_encontrado
            nc = _conceito(item)
            if getattr(item, "ValueType", "") == "TEXT" and nc == "Label":
                try: label_deste_nivel = str(item.TextValue)
                except: pass
        for item in elem.value:
            if _eh_finding_site(item): continue
            nc = _conceito(item)
            if getattr(item, "ValueType", "") == "NUM" and nc == "Value" and label_deste_nivel:
                nc = label_deste_nivel
            nctx = f"{ctx} > {nc}" if nc else ctx
            if hasattr(item, "MeasuredValueSequence"):
                _capturar(item, nc, nctx, site_deste_nivel, out)
            else:
                extrair_raw(item, out, nctx, site_deste_nivel)
    return out

def info_paciente(ds):
    campos = {"PatientName":"Nome","PatientID":"ID","PatientSex":"Sexo",
              "PatientBirthDate":"Nascimento","StudyDate":"Data do Exame",
              "InstitutionName":"Instituição","StudyDescription":"Descrição"}
    r = {}
    for a,l in campos.items():
        try:
            v = getattr(ds,a,None)
            if v: r[l] = str(v)
        except: pass
    return r


# ═══════════════════════════════════════════════════════════════════════
# MAPEAMENTO SR → FORMULÁRIO
# ═══════════════════════════════════════════════════════════════════════

def _media_vals(vals):
    return round(sum(vals)/len(vals), 2) if vals else None

def _ctx_tem(site, *palavras):
    site_low = (site or "").lower()
    return all(p.lower() in site_low for p in palavras)

def mapear_para_form(medidas_raw):
    grupos = {}
    for m in medidas_raw:
        grupos.setdefault(m["nome"], []).append(m)

    resultado = {}

    for nome_dicom, itens in grupos.items():
        if nome_dicom in MEDIDAS_OCULTAS or nome_dicom not in MAPA_DICOM:
            continue
        destinos = MAPA_DICOM[nome_dicom]
        itens_sel = [i for i in itens if i.get("selecao")]
        vals_float = [i["valor"] for i in (itens_sel if itens_sel else itens)]

        for (secao, campo, conv_fn) in destinos:
            if campo == "_e_prime_raw":
                if len(vals_float) >= 2:
                    resultado[("TDI","Vel. e' septal")]  = round(min(vals_float)/10, 1)
                    resultado[("TDI","Vel. e' lateral")] = round(max(vals_float)/10, 1)
                elif len(vals_float)==1:
                    resultado[("TDI","Vel. e' septal")] = round(vals_float[0]/10,1)
                continue
            if campo == "_a_prime_raw":
                if len(vals_float) >= 2:
                    resultado[("TDI","Vel. a' septal")]  = round(min(vals_float)/10,1)
                    resultado[("TDI","Vel. a' lateral")] = round(max(vals_float)/10,1)
                continue
            if campo == "_e_e_prime_raw":
                if len(vals_float) >= 2:
                    resultado[("TDI","Rel. E/E' septal")]  = round(max(vals_float),1)
                    resultado[("TDI","Rel. E/E' lateral")] = round(min(vals_float),1)
                elif len(vals_float)==1:
                    resultado[("TDI","Rel. E/E' septal")] = round(vals_float[0],1)
                continue
            if campo == "_vti_ctx":
                for it in itens:
                    site = it["site"]; v = it["valor"]
                    val_cm = round(v/10,1) if v > 10 else round(v,1)
                    if _ctx_tem(site,"outflow") and _ctx_tem(site,"left ventricle"):
                        resultado[("AORTA / VSVE","VTI VSVE")] = val_cm
                    elif _ctx_tem(site,"aortic"):
                        resultado[("AORTA / VSVE","VTI Ao")] = val_cm
                    elif _ctx_tem(site,"mitral"):
                        resultado[("VALVA MITRAL","VTI mit.")] = val_cm
                    elif _ctx_tem(site,"outflow") and _ctx_tem(site,"right ventricle"):
                        resultado[("TRICÚSPIDE / PULMONAR","VTI AP")] = val_cm
                continue
            if campo in ("_meanvel_ctx","_meangrad_ctx","_peakvel_ctx","_peakgrad_ctx"):
                for it in itens:
                    site = it["site"]; v = it["valor"]
                    is_aortic    = _ctx_tem(site,"aortic")
                    is_mitral    = _ctx_tem(site,"mitral")
                    is_tricuspid = _ctx_tem(site,"tricuspid")
                    is_vsve      = _ctx_tem(site,"outflow") and _ctx_tem(site,"left ventricle")
                    if campo == "_meangrad_ctx":
                        if is_aortic:   resultado[("AORTA / VSVE","Grad. méd.")] = round(v,1)
                        elif is_mitral: resultado[("VALVA MITRAL","Grad. méd.")] = round(v,1)
                    elif campo == "_peakgrad_ctx":
                        if is_aortic:      resultado[("AORTA / VSVE","Grad. máx.")] = round(v,1)
                        elif is_mitral:    resultado[("VALVA MITRAL","Grad. máx.")] = round(v,1)
                        elif is_tricuspid: resultado[("TRICÚSPIDE / PULMONAR","PSAP")] = round(v,1)
                    elif campo == "_peakvel_ctx":
                        val_ms = round(v/1000,2) if v > 10 else round(v,2)
                        if is_aortic:
                            resultado[("AORTA / VSVE","_peak_ao")] = val_ms
                        elif is_tricuspid:
                            resultado[("TRICÚSPIDE / PULMONAR","Vel. RT")] = val_ms
                    elif campo == "_meanvel_ctx":
                        if is_vsve:
                            resultado[("AORTA / VSVE","_mean_vsve")] = round(v/10,1) if v > 10 else round(v,1)
                continue

            val = _media_vals(vals_float)
            if val is None: continue
            if conv_fn: val = conv_fn(val)
            resultado[(secao, campo)] = val

    _calcular_derivados(resultado)
    return resultado


def _calcular_derivados(resultado):
    def g(sec, campo): return resultado.get((sec, campo))

    peso   = g("DADOS ANTROPOMÉTRICOS","Peso")
    altura = g("DADOS ANTROPOMÉTRICOS","Altura")
    bsa    = g("DADOS ANTROPOMÉTRICOS","Superfície corp.")

    if peso and altura and ("DADOS ANTROPOMÉTRICOS","IMC") not in resultado:
        altura_m = altura/100 if altura > 3 else altura
        if altura_m > 0:
            resultado[("DADOS ANTROPOMÉTRICOS","IMC")] = round(peso/(altura_m**2),1)

    ddve = g("CÂMARAS ESQUERDAS","DdVE")
    dsve = g("CÂMARAS ESQUERDAS","DsVE")
    sep  = g("CÂMARAS ESQUERDAS","Septo")
    pp   = g("CÂMARAS ESQUERDAS","Parede post.")

    if ddve and ("CÂMARAS ESQUERDAS","VDVE") not in resultado:
        d = ddve/10; resultado[("CÂMARAS ESQUERDAS","VDVE")] = round((7*d**3)/(2.4+d),1)
    if dsve and ("CÂMARAS ESQUERDAS","VSVE") not in resultado:
        d = dsve/10; resultado[("CÂMARAS ESQUERDAS","VSVE")] = round((7*d**3)/(2.4+d),1)

    vdve = g("CÂMARAS ESQUERDAS","VDVE")
    vsve = g("CÂMARAS ESQUERDAS","VSVE")
    if vdve and vsve and vdve > 0 and ("CÂMARAS ESQUERDAS","FEVE (Teichholz)") not in resultado:
        resultado[("CÂMARAS ESQUERDAS","FEVE (Teichholz)")] = round(((vdve-vsve)/vdve)*100,1)

    if ddve and sep and pp and ("CÂMARAS ESQUERDAS","ERP") not in resultado:
        resultado[("CÂMARAS ESQUERDAS","ERP")] = round((sep+pp)/ddve,2)

    if ddve and sep and pp and ("CÂMARAS ESQUERDAS","Massa VE") not in resultado:
        resultado[("CÂMARAS ESQUERDAS","Massa VE")] = round(
            0.8*(1.04*(((ddve+sep+pp)/10)**3-(ddve/10)**3))+0.6, 1)

    if bsa and bsa > 0:
        for (s,c,src_s,src_c) in [
            ("CÂMARAS ESQUERDAS","Aorta ascend. index","CÂMARAS ESQUERDAS","Aorta ascend."),
            ("CÂMARAS ESQUERDAS","VDVE index","CÂMARAS ESQUERDAS","VDVE"),
            ("CÂMARAS ESQUERDAS","AE - Vol. bipl. index","CÂMARAS ESQUERDAS","AE - Vol. bipl."),
            ("CÂMARAS DIREITAS","AD - Vol. index","CÂMARAS DIREITAS","AD - Vol."),
        ]:
            v = g(src_s,src_c)
            if v and (s,c) not in resultado:
                resultado[(s,c)] = round(v/bsa,1)

        massa = g("CÂMARAS ESQUERDAS","Massa VE")
        if massa and ("CÂMARAS ESQUERDAS","Massa index") not in resultado:
            resultado[("CÂMARAS ESQUERDAS","Massa index")] = round(massa/bsa,1)

        d  = g("AORTA / VSVE","Diâm VSVE")
        vv = g("AORTA / VSVE","VTI VSVE")
        va = g("AORTA / VSVE","VTI Ao")
        avao = g("AORTA / VSVE","AVAo (EC-VTI)")
        if not avao and d and vv and va:
            avao = round((3.1416*((d/10)/2)**2*vv)/va,2)
            resultado[("AORTA / VSVE","AVAo (EC-VTI)")] = avao
        if avao and ("AORTA / VSVE","AVAo index") not in resultado:
            resultado[("AORTA / VSVE","AVAo index")] = round(avao/bsa,2)

    s_tdi = g("TDI","Rel. E/E' septal")
    l_tdi = g("TDI","Rel. E/E' lateral")
    if s_tdi and l_tdi and ("TDI","E/e' MÉDIO") not in resultado:
        resultado[("TDI","E/e' MÉDIO")] = round((s_tdi+l_tdi)/2,1)

    peak_ao = g("AORTA / VSVE","_peak_ao")
    if peak_ao and ("AORTA / VSVE","Grad. máx.") not in resultado:
        resultado[("AORTA / VSVE","Grad. máx.")] = round(4*(peak_ao**2),1)


# ═══════════════════════════════════════════════════════════════════════
# FÓRMULAS REATIVAS (recalculadas a cada interação no Streamlit)
# ═══════════════════════════════════════════════════════════════════════

FORMULAS_CALCULADAS = [
    (("DADOS ANTROPOMÉTRICOS","Superfície corp."),
        [("DADOS ANTROPOMÉTRICOS","Peso"),("DADOS ANTROPOMÉTRICOS","Altura")],
        lambda p,a: round(0.007184*(p**0.425)*((a if a>3 else a*100)**0.725),2) if p and a else None),

    (("DADOS ANTROPOMÉTRICOS","IMC"),
        [("DADOS ANTROPOMÉTRICOS","Peso"),("DADOS ANTROPOMÉTRICOS","Altura")],
        lambda p,a: round(p/((a/100 if a>3 else a)**2),1) if a else None),

    (("CÂMARAS ESQUERDAS","VDVE"),
        [("CÂMARAS ESQUERDAS","DdVE")],
        lambda dd: round((7*((dd/10)**3))/(2.4+(dd/10)),1) if dd else None),

    (("CÂMARAS ESQUERDAS","VSVE"),
        [("CÂMARAS ESQUERDAS","DsVE")],
        lambda ds: round((7*((ds/10)**3))/(2.4+(ds/10)),1) if ds else None),

    (("CÂMARAS ESQUERDAS","FEVE (Teichholz)"),
        [("CÂMARAS ESQUERDAS","VDVE"),("CÂMARAS ESQUERDAS","VSVE")],
        lambda vd,vs: round(((vd-vs)/vd)*100,1) if vd and vd>0 else None),

    (("CÂMARAS ESQUERDAS","ERP"),
        [("CÂMARAS ESQUERDAS","Septo"),("CÂMARAS ESQUERDAS","Parede post."),("CÂMARAS ESQUERDAS","DdVE")],
        lambda s,pp,dd: round((s+pp)/dd,2) if dd else None),

    (("CÂMARAS ESQUERDAS","Massa VE"),
        [("CÂMARAS ESQUERDAS","DdVE"),("CÂMARAS ESQUERDAS","Septo"),("CÂMARAS ESQUERDAS","Parede post.")],
        lambda dd,s,pp: round(0.8*(1.04*(((dd+s+pp)/10)**3-(dd/10)**3))+0.6,1) if dd and s and pp else None),

    (("CÂMARAS ESQUERDAS","Massa index"),
        [("CÂMARAS ESQUERDAS","Massa VE"),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda m,bsa: round(m/bsa,1) if bsa else None),

    (("CÂMARAS ESQUERDAS","Aorta ascend. index"),
        [("CÂMARAS ESQUERDAS","Aorta ascend."),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda ao,bsa: round(ao/bsa,1) if bsa else None),

    (("CÂMARAS ESQUERDAS","VDVE index"),
        [("CÂMARAS ESQUERDAS","VDVE"),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda v,bsa: round(v/bsa,1) if bsa else None),

    (("CÂMARAS ESQUERDAS","AE - Vol. bipl. index"),
        [("CÂMARAS ESQUERDAS","AE - Vol. bipl."),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda v,bsa: round(v/bsa,1) if bsa else None),

    (("CÂMARAS DIREITAS","AD - Vol. index"),
        [("CÂMARAS DIREITAS","AD - Vol."),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda v,bsa: round(v/bsa,1) if bsa else None),

    (("AORTA / VSVE","AVAo (EC-VTI)"),
        [("AORTA / VSVE","Diâm VSVE"),("AORTA / VSVE","VTI VSVE"),("AORTA / VSVE","VTI Ao")],
        lambda d,vv,va: round((3.1416*((d/10)/2)**2*vv)/va,2) if va else None),

    (("AORTA / VSVE","AVAo index"),
        [("AORTA / VSVE","AVAo (EC-VTI)"),("DADOS ANTROPOMÉTRICOS","Superfície corp.")],
        lambda a,bsa: round(a/bsa,2) if bsa else None),

    (("VALVA MITRAL","Relação E/A"),
        [("VALVA MITRAL","Vel. onda E"),("VALVA MITRAL","Vel. onda A")],
        lambda e,a: round(e/a,2) if a else None),

    (("TDI","Rel. E/E' septal"),
        [("VALVA MITRAL","Vel. onda E"),("TDI","Vel. e' septal")],
        lambda e,ep: round((e*100)/ep,1) if ep else None),

    (("TDI","Rel. E/E' lateral"),
        [("VALVA MITRAL","Vel. onda E"),("TDI","Vel. e' lateral")],
        lambda e,ep: round((e*100)/ep,1) if ep else None),

    (("TDI","E/e' MÉDIO"),
        [("TDI","Rel. E/E' septal"),("TDI","Rel. E/E' lateral")],
        lambda s,l: round((s+l)/2,1)),

    (("TDI","E/A tecidual"),
        [("TDI","Vel. e' septal"),("TDI","Vel. a' septal")],
        lambda e,a: round(e/a,2) if a else None),
]

CAMPOS_2_DECIMAIS = {("CÂMARAS ESQUERDAS","ERP")}


def _fmt(dest, valor_float) -> str:
    """Formata um float para exibição usando vírgula como separador decimal."""
    if dest in CAMPOS_2_DECIMAIS:
        return f"{valor_float:.2f}".replace(".", ",")
    s = f"{valor_float:.10g}"
    return s.replace(".", ",")


def _to_float(s):
    """Converte string com vírgula ou ponto para float."""
    try:
        return float(str(s).strip().replace(",", ".")) if str(s).strip() != "" else None
    except (ValueError, TypeError):
        return None


# ═══════════════════════════════════════════════════════════════════════
# RECÁLCULO — aplica todas as fórmulas em cascata sobre o dict de valores
# ═══════════════════════════════════════════════════════════════════════

def recalcular(valores: dict) -> dict:
    """Aplica FORMULAS_CALCULADAS em cascata até estabilizar."""
    v = dict(valores)
    for _ in range(len(FORMULAS_CALCULADAS) + 1):
        mudou = False
        for (dest, fontes, fn) in FORMULAS_CALCULADAS:
            srcs = [_to_float(v.get(k)) for k in fontes]
            if any(s is None for s in srcs):
                continue
            try:
                novo = fn(*srcs)
            except (ZeroDivisionError, TypeError, ValueError):
                continue
            if novo is None:
                continue
            novo_fmt = _fmt(dest, novo)
            if v.get(dest) != novo_fmt:
                v[dest] = novo_fmt
                mudou = True
        if not mudou:
            break
    return v


# ═══════════════════════════════════════════════════════════════════════
# REFERÊNCIAS
# ═══════════════════════════════════════════════════════════════════════

def _parse_ref(ref_str):
    if not ref_str or ref_str.strip() in ("","-"): return None
    s = ref_str.strip().replace(",",".")
    m = re.match(r'^([\d.]+)\s*-\s*([\d.]+)$', s)
    if m: return ('range', float(m.group(1)), float(m.group(2)))
    m = re.match(r'^<=\s*([\d.]+)$', s)
    if m: return ('max_eq', float(m.group(1)))
    m = re.match(r'^<\s*([\d.]+)$', s)
    if m: return ('max', float(m.group(1)))
    m = re.match(r'^>=\s*([\d.]+)$', s)
    if m: return ('min_eq', float(m.group(1)))
    m = re.match(r'^>\s*([\d.]+)$', s)
    if m: return ('min', float(m.group(1)))
    return None

def _dentro_ref(valor_str, ref_str):
    try: v = float(str(valor_str).strip().replace(",","."))
    except: return None
    p = _parse_ref(ref_str)
    if p is None: return None
    t = p[0]
    if t=='range':  return p[1] <= v <= p[2]
    if t=='max_eq': return v <= p[1]
    if t=='max':    return v < p[1]
    if t=='min_eq': return v >= p[1]
    if t=='min':    return v > p[1]
    return None


# ═══════════════════════════════════════════════════════════════════════
# LAUDO DESCRITIVO
# ═══════════════════════════════════════════════════════════════════════

def _gv(valores, secao, campo):
    """Pega valor float de valores dict."""
    s = str(valores.get((secao, campo), "")).strip()
    try: return float(s.replace(",","."))
    except: return None

def _ref(sexo, secao, campo_nome):
    for sec, campos in FORMULARIO.items():
        if sec == secao:
            for c in campos:
                if c["name"] == campo_nome:
                    return c["ref_mas"] if sexo=="M" else c["ref_fem"]
    return ""

def gerar_laudo(valores, sexo):
    def av(sec, campo):
        v = _gv(valores, sec, campo)
        r = _ref(sexo, sec, campo)
        ok = _dentro_ref(str(v), r) if v is not None else None
        return v, ok

    ddve, ddve_ok = av("CÂMARAS ESQUERDAS","DdVE")
    erp,  erp_ok  = av("CÂMARAS ESQUERDAS","ERP")
    massa,massa_ok = av("CÂMARAS ESQUERDAS","Massa index")
    feve_s,_ = av("CÂMARAS ESQUERDAS","FEVE (Simpson)")
    feve_t,_ = av("CÂMARAS ESQUERDAS","FEVE (Teichholz)")
    feve_val = feve_s if feve_s is not None else feve_t
    ae_vol, ae_ok   = av("CÂMARAS ESQUERDAS","AE - Vol. bipl. index")
    ae_diam,ae_d_ok = av("CÂMARAS ESQUERDAS","AE - Diâm.")
    _,      vd_ok   = av("CÂMARAS DIREITAS","VD - Diâm. basal")
    ad_vol, ad_ok   = av("CÂMARAS DIREITAS","AD - Vol. index")

    ve_dim_txt = "Ventrículo esquerdo (VE) com dimensões aumentadas (dilatado)" if ddve_ok is False \
                 else "Ventrículo esquerdo (VE) com dimensões normais"

    if erp is not None and massa is not None:
        if erp >= 0.42 and massa_ok is False: geom = "Geometria ventricular: hipertrofia concêntrica"
        elif erp >= 0.42:                     geom = "Geometria ventricular: remodelamento concêntrico"
        elif massa_ok is False:               geom = "Geometria ventricular: hipertrofia excêntrica"
        else:                                 geom = "Geometria ventricular: normal"
    elif erp is not None:
        geom = "Geometria ventricular: remodelamento concêntrico" if erp >= 0.42 else "Geometria ventricular: normal"
    else:
        geom = "Geometria ventricular: normal"

    if feve_val is not None:
        lim = 52 if sexo=="M" else 54
        if feve_val >= lim:    fsist = f"Função sistólica normal do VE (FEVE {feve_val:.0f}%)"
        elif feve_val >= 40:   fsist = f"Função sistólica do VE com redução de grau leve (FEVE {feve_val:.0f}%)"
        elif feve_val >= 30:   fsist = f"Função sistólica do VE com redução de grau moderado (FEVE {feve_val:.0f}%)"
        else:                  fsist = f"Função sistólica do VE com redução de grau importante (FEVE {feve_val:.0f}%)"
    else:
        fsist = "Função sistólica do VE"

    if ae_ok is False and ae_vol:
        ae_txt = f"Átrio esquerdo (AE) com volume aumentado (índice {ae_vol:.1f} mL/m²)"
    elif ae_d_ok is False and ae_diam:
        ae_txt = f"Átrio esquerdo (AE) com dimensão aumentada ({ae_diam:.1f} mm)"
    else:
        ae_txt = "Átrio esquerdo (AE) com volume normal"

    vd_txt = "Ventrículo direito (VD) com dimensões aumentadas" if vd_ok is False \
             else "Ventrículo direito (VD) com dimensões normais"
    ad_txt = f"Átrio direito (AD) com área e volume aumentados (índice {ad_vol:.1f} mL/m²)" \
             if ad_ok is False and ad_vol else "Átrio direito (AD) com área e volume normais"

    return [
        "**CÂMARAS ESQUERDAS**",
        ve_dim_txt, geom,
        "Espessamento sistólico normal em todos os segmentos do VE", fsist, "",
        ae_txt, "",
        "**CÂMARAS DIREITAS**", vd_txt, "", ad_txt, "",
        "**VALVAS CARDÍACAS**",
        "Valva aórtica com textura, mobilidade e abertura normais dos folhetos | Ausência de sinais de refluxo",
        "Valva mitral com textura, mobilidade e abertura normais dos folhetos | Ausência de sinais de refluxo",
        "Valva tricúspide com textura, mobilidade e abertura normais dos folhetos | Ausência de sinais de refluxo",
        "Valva pulmonar com textura, mobilidade e abertura normais dos folhetos | Ausência de sinais de refluxo", "",
        "**VASOS DA BASE**",
        "Aorta ascendente com calibre normal | Paredes com textura normal | Fluxo normal",
        "Artéria Pulmonar com calibre normal | Fluxo normal", "",
        "**PERICÁRDICO**", "Textura e deslizamento normais do pericárdico", "",
        "**CONGÊNITAS**",
        "Situs solitus, levocardia | Concordâncias veno-atrial, átrio-ventricular e ventrículo-arterial | Septos íntegros | Canal arterial não visualizado", "",
        "**CONCLUSÃO**",
        "- Câmaras cardíacas com dimensões normais",
        "- Funções sistólica e diastólica biventricular normais",
        "- Valvas cardíacas com aspectos morfofuncionais normais",
        "- Ecodopplercardiograma transtorácico normal",
    ]


# ═══════════════════════════════════════════════════════════════════════
# EXPORTAÇÃO
# ═══════════════════════════════════════════════════════════════════════

def exportar_csv_bytes(paciente, valores, sexo):
    W_MEDIDA=34; W_VALOR=8; W_UNIDADE=8; W_REF=20
    TOTAL=W_MEDIDA+W_VALOR+W_UNIDADE+W_REF+7
    SEP_H="="*TOTAL; SEP_L="-"*TOTAL

    def fmt(medida,valor,unidade,ref):
        return f"  {str(medida)[:W_MEDIDA].ljust(W_MEDIDA)} | {str(valor)[:W_VALOR].rjust(W_VALOR)} | {str(unidade)[:W_UNIDADE].ljust(W_UNIDADE)} | {str(ref)[:W_REF].ljust(W_REF)}"

    buf = io.StringIO()
    def w(line=""): buf.write(line+"\r\n")

    w(SEP_H); w("  ECOCARDIOGRAMA - Banco de Dados de Pesquisa"); w(SEP_H)
    for k,v in paciente.items(): w(fmt(k,v,"-","-"))
    w(SEP_L); w()
    w(SEP_H); w(fmt("MEDIDA","VALOR","UNIDADE","REFERÊNCIA")); w(SEP_H)

    secoes = {}
    for secao,campos in FORMULARIO.items():
        for campo in campos:
            val = str(valores.get((secao,campo["name"]),"")).strip()
            if val and val!="-":
                ref = campo["ref_mas"] if sexo=="M" else campo["ref_fem"]
                secoes.setdefault(secao,[]).append((campo["name"],val,campo["unit"],ref))

    for secao,itens in secoes.items():
        w(); w(SEP_L); w(f"  {secao}"); w(SEP_L)
        for nome,val,unit,ref in itens:
            w(fmt(nome,val,unit,ref or "-"))

    w(); w(SEP_H); w(); w(SEP_H); w("  LAUDO DESCRITIVO"); w(SEP_H); w()
    for linha in gerar_laudo(valores, sexo):
        linha_limpa = linha.replace("**","")
        w(f"  {linha_limpa}")
    w(); w(SEP_H)

    return buf.getvalue().encode("utf-8-sig")


def exportar_excel_bytes(paciente, valores, sexo):
    thin  = Side(style="thin",color="CCCCCC")
    BRD   = Border(left=thin,right=thin,top=thin,bottom=thin)
    CTR   = Alignment(horizontal="center",vertical="center",wrap_text=True)
    LEFT  = Alignment(horizontal="left",  vertical="center",wrap_text=True)
    H_FILL= PatternFill("solid",fgColor="1F3864")
    H_FONT= Font(name="Calibri",bold=True,color="FFFFFF",size=10)
    S_FILL= PatternFill("solid",fgColor="2E4D8A")
    S_FONT= Font(name="Calibri",bold=True,color="FFFFFF",size=10)
    Z_FILL= PatternFill("solid",fgColor="EBF3FB")
    C_FILL= PatternFill("solid",fgColor="FFF2CC")
    D_FONT= Font(name="Calibri",size=10)

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title="Ecocardiograma"
    row=1
    for k,v in paciente.items():
        ws.cell(row=row,column=1,value=k).font=Font(name="Calibri",bold=True,size=9,color="1F3864")
        ws.cell(row=row,column=2,value=v).font=Font(name="Calibri",size=9)
        row+=1
    row+=1

    for c,(h,w2) in enumerate(zip(
        ["Seção","Medida","Valor","Unidade","Ref. Masc.","Ref. Fem.","Calc."],
        [24,      24,      10,     10,        18,          18,         8]),1):
        cell=ws.cell(row=row,column=c,value=h)
        cell.font=H_FONT; cell.fill=H_FILL; cell.alignment=CTR; cell.border=BRD
        ws.column_dimensions[get_column_letter(c)].width=w2
    row+=1

    sec_ant=""; zebra=False
    for secao,campos in FORMULARIO.items():
        if secao!=sec_ant:
            for c in range(1,8):
                cell=ws.cell(row=row,column=c,value=secao if c==1 else "")
                cell.font=S_FONT; cell.fill=S_FILL; cell.border=BRD
                cell.alignment=LEFT if c==1 else CTR
            row+=1; sec_ant=secao; zebra=False
        for campo in campos:
            val=str(valores.get((secao,campo["name"]),"")).strip()
            fill=C_FILL if campo["calc"] else (Z_FILL if zebra else None)
            for c,v in enumerate(["",campo["name"],val,campo["unit"],
                                   campo["ref_mas"],campo["ref_fem"],
                                   "⚙" if campo["calc"] else ""],1):
                cell=ws.cell(row=row,column=c,value=v)
                cell.font=D_FONT; cell.border=BRD
                cell.alignment=CTR if c in (3,4,5,6,7) else LEFT
                if fill: cell.fill=fill
            row+=1; zebra=not zebra
    ws.freeze_panes=f"A{len(paciente)+3}"

    # Aba banco de dados
    ws2=wb.create_sheet("Banco de Dados")
    all_campos=[(sec,c["name"]) for sec,campos in FORMULARIO.items() for c in campos]
    info_keys=list(paciente.keys())
    for c,k in enumerate(info_keys,1):
        cell=ws2.cell(row=1,column=c,value=k)
        cell.font=H_FONT; cell.fill=H_FILL; cell.alignment=CTR; cell.border=BRD
        ws2.column_dimensions[get_column_letter(c)].width=16
    offset=len(info_keys)+1
    for c,(sec,nome) in enumerate(all_campos,offset):
        cell=ws2.cell(row=1,column=c,value=f"{nome}\n({sec})")
        cell.font=H_FONT; cell.fill=H_FILL; cell.alignment=CTR; cell.border=BRD
        ws2.column_dimensions[get_column_letter(c)].width=14
    for c,k in enumerate(info_keys,1):
        ws2.cell(row=2,column=c,value=paciente.get(k,"")).font=D_FONT
    for c,(sec,nome) in enumerate(all_campos,offset):
        ws2.cell(row=2,column=c,value=str(valores.get((sec,nome),""))).font=D_FONT
    ws2.row_dimensions[1].height=36; ws2.freeze_panes="A2"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════════════════════════════
# INICIALIZAÇÃO DO SESSION STATE
# ═══════════════════════════════════════════════════════════════════════

def _init_state():
    if "valores" not in st.session_state:
        st.session_state.valores = {}   # {(secao, campo): str}
    if "paciente" not in st.session_state:
        st.session_state.paciente = {}
    if "sexo" not in st.session_state:
        st.session_state.sexo = "F"


# ═══════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ═══════════════════════════════════════════════════════════════════════

def main():
    st.set_page_config(
        page_title="Eco SR Reader",
        page_icon="🫀",
        layout="wide",
    )
    _init_state()

    # ── CSS customizado ───────────────────────────────────────────────
    st.markdown("""
    <style>
    .sec-header {
        background: #2E4D8A; color: white;
        font-weight: bold; font-size: 13px;
        padding: 6px 12px; border-radius: 4px;
        margin: 12px 0 4px 0;
    }
    .ref-ok  { color: #a6e3a1; font-size: 12px; }
    .ref-bad { color: #f38ba8; font-size: 12px; font-weight: bold; }
    .ref-neu { color: #89dceb; font-size: 12px; }
    .calc-label { color: #f9e2af; font-size: 11px; }
    div[data-testid="stNumberInput"] input { text-align: center; }
    </style>
    """, unsafe_allow_html=True)

    # ── Cabeçalho ────────────────────────────────────────────────────
    col_title, col_sexo = st.columns([6,1])
    with col_title:
        st.title("🫀 Eco SR Reader")
        st.caption("Formulário Estruturado de Ecocardiograma")
    with col_sexo:
        st.session_state.sexo = st.radio(
            "Sexo (referência)", ["M","F"],
            index=0 if st.session_state.sexo=="M" else 1,
            horizontal=True)

    st.divider()

    # ── Upload DICOM SR pelo navegador ───────────────────────────────
    with st.expander("📂 Carregar DICOM SR", expanded=True):
        uploaded = st.file_uploader(
            "Envie um ou mais arquivos DICOM SR",
            type=None,
            accept_multiple_files=True,
            help="Selecione os arquivos DICOM SR do ecocardiograma diretamente pelo navegador.",
        )

        if uploaded:
            # Filtra apenas SRs válidos lendo os bytes em memória
            srs_validos = []
            for f in uploaded:
                raw_bytes = f.read()
                try:
                    ds_test = pydicom.dcmread(io.BytesIO(raw_bytes), force=True,
                                              stop_before_pixels=True)
                    modality = str(getattr(ds_test, "Modality", "")).strip().upper()
                    sop      = str(getattr(ds_test, "SOPClassUID", ""))
                    if modality == "SR" or sop in SR_SOP:
                        srs_validos.append((f.name, raw_bytes))
                except Exception:
                    pass

            if not srs_validos:
                st.warning("Nenhum arquivo SR válido encontrado nos arquivos enviados.")
            else:
                nomes = [n for n, _ in srs_validos]
                idx = 0
                if len(srs_validos) > 1:
                    sel = st.selectbox("Selecione o SR para carregar:", nomes)
                    idx = nomes.index(sel)
                else:
                    st.info(f"Arquivo SR detectado: **{nomes[0]}**")

                nome_sel, bytes_sel = srs_validos[idx]

                if st.button("✅ Carregar SR selecionado", type="primary"):
                    with st.spinner("Lendo DICOM SR..."):
                        ds = pydicom.dcmread(io.BytesIO(bytes_sel), force=True)
                        st.session_state.paciente = info_paciente(ds)
                        raw = extrair_raw(ds)
                        mapeado = mapear_para_form(raw)

                        novos = {}
                        for (sec, campo), val in mapeado.items():
                            if campo.startswith("_"): continue
                            dest = (sec, campo)
                            novos[dest] = _fmt(dest, float(val))

                        novos_calc = recalcular(novos)
                        st.session_state.valores = novos_calc

                        # Sincroniza os inputs com os novos valores
                        for (sec, campo), val in novos_calc.items():
                            st.session_state[f"inp_{sec}_{campo}"] = val

                        sx = st.session_state.paciente.get("Sexo", "")
                        if sx.upper() in ("M", "MALE", "MASCULINO"):
                            st.session_state.sexo = "M"
                        elif sx.upper() in ("F", "FEMALE", "FEMININO"):
                            st.session_state.sexo = "F"

                    preenchidos = sum(1 for v in st.session_state.valores.values() if v)
                    st.success(f"✅ {preenchidos} campos preenchidos a partir de {nome_sel}")
                    st.rerun()

    # ── Dados do Paciente ────────────────────────────────────────────
    if st.session_state.paciente:
        pac = st.session_state.paciente
        cols = st.columns(len(pac))
        for col, (k,v) in zip(cols, pac.items()):
            col.metric(k, v)
        st.divider()

    # ── Recalcula fórmulas a partir dos valores atuais ───────────────
    # Converte strings → floats para as fontes, aplica fórmulas
    valores_str = st.session_state.valores   # {(sec,campo): str}

    def get_float(sec, campo):
        return _to_float(valores_str.get((sec, campo), ""))

    # Aplica fórmulas em cascata e atualiza os valores calculados
    valores_calc = {}
    for (dest, fontes, fn) in FORMULAS_CALCULADAS:
        srcs = [get_float(s,c) for s,c in fontes]
        if any(x is None for x in srcs): continue
        try:
            novo = fn(*srcs)
        except: continue
        if novo is None: continue
        valores_calc[dest] = _fmt(dest, novo)

    # Valores finais: manuais têm prioridade, calculados preenchem o resto
    valores_exibir = {**valores_calc, **{k:v for k,v in valores_str.items() if v}}

    # ── Pré-popula session_state dos widgets antes de renderizar ─────
    # Isso garante que os valores apareçam nos inputs após carregar SR
    for key, val in valores_exibir.items():
        wkey = f"inp_{key[0]}_{key[1]}"
        if wkey not in st.session_state:
            st.session_state[wkey] = val

    # ── Formulário por seções ────────────────────────────────────────
    st.subheader("📋 Formulário de Medidas")

    sexo = st.session_state.sexo
    valores_editados = {}

    for secao, campos in FORMULARIO.items():
        st.markdown(f'<div class="sec-header">{secao}</div>', unsafe_allow_html=True)

        # Cabeçalho das colunas
        h1,h2,h3,h4,h5 = st.columns([3,1.2,0.7,1.8,0.3])
        h1.markdown("**Medida**"); h2.markdown("**Valor**")
        h3.markdown("**Un.**");   h4.markdown("**Referência**")

        for campo in campos:
            key = (secao, campo["name"])
            wkey = f"inp_{secao}_{campo['name']}"
            ref = campo["ref_mas"] if sexo=="M" else campo["ref_fem"]
            is_calc = campo["calc"]

            # Atualiza o widget com valor calculado se não há entrada manual
            val_calc = valores_calc.get(key, "")
            val_manual = valores_str.get(key, "")
            val_widget = st.session_state.get(wkey, "")

            # Se calculado mudou e o usuário não digitou nada diferente, atualiza
            if val_calc and val_widget == val_manual and val_widget != val_calc:
                st.session_state[wkey] = val_calc

            val_exibir = valores_exibir.get(key, "")

            c1,c2,c3,c4,c5 = st.columns([3,1.2,0.7,1.8,0.3])

            with c1:
                label = campo["name"]
                if is_calc:
                    st.markdown(f'<span class="calc-label">⚙ {label}</span>', unsafe_allow_html=True)
                else:
                    st.markdown(f"<span style='font-size:13px'>{label}</span>", unsafe_allow_html=True)

            with c2:
                novo_val = st.text_input(
                    label=campo["name"],
                    key=wkey,
                    label_visibility="collapsed",
                )
                valores_editados[key] = novo_val

            with c3:
                st.markdown(f"<span style='color:#6c7086;font-size:12px'>{campo['unit']}</span>",
                            unsafe_allow_html=True)

            with c4:
                dentro = _dentro_ref(val_exibir, ref) if val_exibir else None
                if dentro is True:
                    st.markdown(f'<span class="ref-ok">✓ {ref}</span>', unsafe_allow_html=True)
                elif dentro is False:
                    st.markdown(f'<span class="ref-bad">⚠ {ref}</span>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<span class="ref-neu">{ref}</span>', unsafe_allow_html=True)

            with c5:
                if is_calc:
                    st.markdown('<span class="calc-label">⚙</span>', unsafe_allow_html=True)

    # Persiste edições manuais e recalcula
    st.session_state.valores = recalcular(valores_editados)

    st.divider()

    # ── Exportação ───────────────────────────────────────────────────
    st.subheader("💾 Exportar")
    col_csv, col_xls, col_limpar = st.columns([1,1,1])

    pac = st.session_state.paciente
    nome_pac = pac.get("Nome","paciente").replace(" ","_").replace("/","-")

    tem_dados = any(v for v in st.session_state.valores.values())

    with col_csv:
        if tem_dados:
            csv_bytes = exportar_csv_bytes(pac, valores_exibir, sexo)
            st.download_button(
                "📄 Baixar CSV",
                data=csv_bytes,
                file_name=f"eco_{nome_pac}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.button("📄 Baixar CSV", disabled=True, use_container_width=True)

    with col_xls:
        if tem_dados:
            xls_bytes = exportar_excel_bytes(pac, valores_exibir, sexo)
            st.download_button(
                "📊 Baixar Excel",
                data=xls_bytes,
                file_name=f"eco_{nome_pac}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("📊 Baixar Excel", disabled=True, use_container_width=True)

    with col_limpar:
        if st.button("🗑 Limpar Formulário", use_container_width=True):
            st.session_state.valores = {}
            st.session_state.paciente = {}
            st.rerun()

    # ── Laudo Descritivo ─────────────────────────────────────────────
    if tem_dados:
        st.divider()
        with st.expander("📝 Laudo Descritivo", expanded=False):
            linhas = gerar_laudo(valores_exibir, sexo)
            texto_laudo = "\n".join(l.replace("**","") for l in linhas)
            st.text_area("Laudo (editável)", value=texto_laudo, height=420,
                         key="laudo_texto")
            laudo_bytes = texto_laudo.encode("utf-8")
            st.download_button("📄 Baixar Laudo (.txt)",
                               data=laudo_bytes,
                               file_name=f"laudo_{nome_pac}.txt",
                               mime="text/plain")


if __name__ == "__main__":
    main()