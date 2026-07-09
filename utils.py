def limpiar_monto(texto):
    if not texto: return 0.0
    t = texto.strip().replace('$', '').replace('Bs', '').replace('"', '')
    pos_coma = t.rfind(',')
    pos_punto = t.rfind('.')
    if pos_coma > pos_punto:
        t = t.replace('.', '').replace(',', '.')
    elif pos_punto > pos_coma:
        t = t.replace(',', '')
    try:
        return float(t)
    except ValueError:
        return 0.0
