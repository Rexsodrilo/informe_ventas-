#!/usr/bin/env python3
"""Genera dataCOM.json.gz a partir de la hoja 'Corte' del Excel.

Uso (desde la terminal):
    python actualizar_datos.py                       # pide la clave y usa la ruta por defecto
    python actualizar_datos.py "C:\\ruta\\archivo.xlsm"
    python actualizar_datos.py --push                # ademas hace git add/commit/push

Requiere: pip install openpyxl
"""
import argparse, collections, datetime as dt, getpass, gzip, hashlib, json, os, subprocess, sys

RUTA_DEFECTO = r"T:\Corte_OC_FIFO\Base\TEst_4_Corte_v4.xlsm"
SALIDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dataCOM.json.gz")
# Se guarda solo el hash (SHA-256) de la clave, no la clave en texto plano.
CLAVE_SHA256 = "79ab7e4a022c5796db53ffff61c508401ae1523feed9e6662b452414d3728151"


def pedir_clave():
    for _ in range(3):
        c = os.environ.get("BODEGA_CLAVE") or getpass.getpass("Clave de actualizacion: ")
        if hashlib.sha256(c.encode()).hexdigest() == CLAVE_SHA256:
            return
        print("Clave incorrecta.")
    sys.exit("Acceso denegado.")


def leer_corte(ruta):
    try:
        import openpyxl
    except ImportError:
        sys.exit("Falta openpyxl. Instale con: python -m pip install openpyxl")
    print("Leyendo Excel (puede tardar 1-2 minutos)...")
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    it = wb["Corte"].iter_rows(values_only=True)
    ix = {str(h).strip(): i for i, h in enumerate(next(it)) if h is not None}
    need = ["NVNumero", "NomAux", "VenDes", "Canal_EN", "Fec_Cre", "Dias NV", "Dias_em_NV", "Valor SdespNV",
            "CodProd", "DetProd", "Sal_X_ Des", "ES_INV", "Bloqueado", "Cum_Val"]
    falta = [n for n in need if n not in ix]
    if falta:
        sys.exit("Faltan columnas en la hoja Corte: " + ", ".join(falta))
    g = lambda r, n: r[ix[n]]
    lineas, ref = [], collections.Counter()
    for r in it:
        if g(r, "NVNumero") is None:
            continue
        fec, dn = g(r, "Fec_Cre"), g(r, "Dias NV")
        if isinstance(fec, dt.datetime) and isinstance(dn, (int, float)):
            ref[(fec + dt.timedelta(days=int(dn))).date()] += 1
        if (g(r, "Sal_X_ Des") or 0) <= 0:
            continue
        lineas.append([int(g(r, "NVNumero")), (g(r, "NomAux") or "").strip(), g(r, "VenDes") or "Sin vendedor",
                       g(r, "Canal_EN") or "", fec.strftime("%Y-%m-%d"), round(g(r, "Dias_em_NV") or 0, 1),
                       round(g(r, "Valor SdespNV") or 0), str(g(r, "CodProd")), (g(r, "DetProd") or "").strip(),
                       g(r, "Sal_X_ Des"), 1 if g(r, "ES_INV") == "QUIEBRE" else 0,
                       1 if g(r, "Bloqueado") == "S" else 0, 1 if g(r, "Cum_Val") == "VALOR_BAJO" else 0])
    corte = ref.most_common(1)[0][0] if ref else dt.date.today()
    return lineas, corte


def bucket(a):
    return 0 if a <= 15 else 1 if a <= 30 else 2 if a <= 90 else 3


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel", nargs="?", default=RUTA_DEFECTO)
    ap.add_argument("--push", action="store_true", help="git add/commit/push del archivo generado")
    a = ap.parse_args()
    pedir_clave()
    if not os.path.exists(a.excel):
        sys.exit("No existe el archivo: " + a.excel)
    lineas, corte = leer_corte(a.excel)

    edad = {}
    monto = collections.Counter()
    cli = {}
    for l in lineas:
        edad[l[0]] = max(edad.get(l[0], 0), l[5]); monto[l[0]] += l[6]; cli[l[0]] = l[1]
    b, mb = [0] * 4, [0] * 4
    for n, e in edad.items():
        b[bucket(e)] += 1; mb[bucket(e)] += monto[n]
    snap = {"fecha": corte.isoformat(), "b": b, "monto_b": mb, "total": len(edad),
            "monto": sum(monto.values()), "clientes": len(set(cli.values()))}

    hist = []
    if os.path.exists(SALIDA):
        try:
            hist = json.loads(gzip.decompress(open(SALIDA, "rb").read())).get("historial", [])
        except Exception:
            hist = []
    hist = [h for h in hist if h["fecha"] != snap["fecha"]] + [snap]
    hist.sort(key=lambda h: h["fecha"])

    out = {"meta": {"generado": dt.datetime.now().strftime("%Y-%m-%d %H:%M"), "corte": corte.isoformat(),
                    "fuente": os.path.basename(a.excel), "lineas": len(lineas)},
           "historial": hist[-104:], "lines": lineas}
    raw = json.dumps(out, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    with open(SALIDA, "wb") as f:
        f.write(gzip.compress(raw, 9))
    print(f"OK: {len(edad)} NV pendientes ({len(lineas)} lineas) -> {SALIDA} ({os.path.getsize(SALIDA)//1024} KB)")

    if a.push:
        d = os.path.dirname(SALIDA)
        try:
            subprocess.run(["git", "add", "dataCOM.json.gz"], cwd=d, check=True)
            subprocess.run(["git", "commit", "-m", "Actualizacion de datos " + out["meta"]["generado"]], cwd=d, check=True)
            subprocess.run(["git", "push"], cwd=d, check=True)
            print("Publicado en GitHub.")
        except subprocess.CalledProcessError as e:
            sys.exit("Fallo git (¿hay cambios? ¿repositorio configurado?): " + str(e))


if __name__ == "__main__":
    main()
