# -*- coding: utf-8 -*-
"""
Genera los entregables del Maestro PK-16 a partir de las fuentes del repositorio:
  salidas/PK16_MAESTRO_CRITICIDAD_DIMENSIONAMIENTO.xlsx  (Excel maestro con fórmulas, recalculado)
  salidas/PK16_MODELO_INTERACTIVO.html                    (HTML interactivo autocontenido)
Uso:  python modelo/build.py            (descarga de la release los ficheros pesados si faltan)
"""
import json
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import pk16_datos as D  # noqa: E402
import pk16_excel as X  # noqa: E402

OUT = HERE.parent / "salidas"
XLSX = OUT / "PK16_MAESTRO_CRITICIDAD_DIMENSIONAMIENTO.xlsx"
HTML = OUT / "PK16_MODELO_INTERACTIVO.html"


def recalcular(path: Path, timeout: int = 900) -> bool:
    """Recalcula con LibreOffice Calc para guardar valores en caché (visores, pandas, revisión).
    Excel recalcula igualmente al abrir (fullCalcOnLoad)."""
    import os
    import shutil
    import tempfile
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        print("    LibreOffice no disponible: el libro se recalculará al abrirlo en Excel.")
        return False
    env = dict(os.environ, SAL_USE_VCLPLUGIN="svp")
    with tempfile.TemporaryDirectory(prefix="lo_pk16_") as prof:
        url = Path(prof).as_uri()
        subprocess.run([soffice, "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"],
                       env=env, capture_output=True, timeout=120)
        mdir = Path(prof) / "user" / "basic" / "Standard"
        if not mdir.exists():
            print("    No se pudo preparar el perfil de LibreOffice.")
            return False
        (mdir / "Module1.xba").write_text(
            '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">\n'
            "Sub RecalculateAndSave()\n  ThisComponent.calculateAll()\n  ThisComponent.store()\n  ThisComponent.close(True)\nEnd Sub\n"
            "</script:module>", encoding="utf-8")
        before = path.stat().st_mtime_ns
        subprocess.run([soffice, "--headless", "--norestore", f"-env:UserInstallation={url}",
                        "vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application", str(path)],
                       env=env, capture_output=True, timeout=timeout)
        ok = path.stat().st_mtime_ns != before
        print("    recalculado" if ok else "    LibreOffice no reescribió el fichero (¿falta el módulo Calc?)")
        return ok


def main():
    OUT.mkdir(exist_ok=True)
    print("1/4 Cargando y cruzando fuentes…")
    M = D.construir_modelo()
    S = D.puntuar(M["inv"])
    print(f"    registros: {len(M['inv'])} · parque fila a fila: {len(S)}")
    print("2/4 Generando Excel maestro…")
    X.construir_excel(M, S, XLSX)
    print("3/4 Recalculando con LibreOffice…")
    recalcular(XLSX)
    print("4/4 Generando HTML interactivo…")
    try:
        import pk16_html as H
        H.construir_html(M, S, XLSX, HTML)
    except ImportError:
        print("    (pk16_html no disponible todavía)")
    print("Listo:", XLSX, HTML)


if __name__ == "__main__":
    main()
