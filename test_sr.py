import pydicom, io

caminho = r'c:\Users\thiago.freire\OneDrive\Python\SR_extractor\eco_huol\53749__MARIA NAZARE BATISTA\0F9B27B4.dcm'

with open(caminho, 'rb') as f:
    raw_bytes = f.read()

print(f'Tamanho do arquivo: {len(raw_bytes)} bytes')

SR_SOP = {
    '1.2.840.10008.5.1.4.1.1.88.11','1.2.840.10008.5.1.4.1.1.88.22',
    '1.2.840.10008.5.1.4.1.1.88.33','1.2.840.10008.5.1.4.1.1.88.34',
    '1.2.840.10008.5.1.4.1.1.88.35','1.2.840.10008.5.1.4.1.1.88.67',
    '1.2.840.10008.5.1.4.1.1.88.68','1.2.840.10008.5.1.4.1.1.88.72',
}

try:
    ds_test = pydicom.dcmread(io.BytesIO(raw_bytes), force=True, stop_before_pixels=True)
    modality = str(getattr(ds_test, 'Modality', '')).strip().upper()
    sop      = str(getattr(ds_test, 'SOPClassUID', '')).strip()
    print(f'Modality: [{modality}]')
    print(f'SOPClassUID: [{sop}]')
    print(f'SOPClassUID in SR_SOP: {sop in SR_SOP}')
    print(f'modality == SR: {modality == "SR"}')
    print(f'ACEITO: {modality == "SR" or sop in SR_SOP}')
except Exception as e:
    print(f'ERRO ao ler: {e}')
