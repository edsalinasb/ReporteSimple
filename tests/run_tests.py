#!/usr/bin/env python3
"""Ejecuta tests/D52_Pruebas.bas contra RSimple_Diario52.bas en LibreOffice Basic.

Requisitos: soffice (LibreOffice) en el PATH. Uso: python3 tests/run_tests.py
Los modulos .bas se leen tal cual estan en el repositorio (ANSI/cp1252, CRLF);
solo se quitan las lineas "Attribute" y se activa "Option VBASupport 1".
Esto valida la logica pura; NO valida DAO, Crystal ni Excel (ver README).
"""
import os
import re
import shutil
import subprocess
import sys
import tempfile
from xml.sax.saxutils import escape

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODULOS = {
    "Diario52": os.path.join(RAIZ, "RSimple_Diario52.bas"),
    "Pruebas": os.path.join(RAIZ, "tests", "D52_Pruebas.bas"),
    # Solo se compila junto con los demas (verificacion de sintaxis/nombres);
    # la automatizacion de Excel no se ejecuta aqui.
    "ExcelF52": os.path.join(RAIZ, "RSimple_Excel.bas"),
}

RUNNER = """Option VBASupport 1
Option Explicit
Sub Main
    Dim s As String, n As Integer
    On Error GoTo Fallo
    s = D52_EjecutarPruebas()
    GoTo Escribir
Fallo:
    s = "ERROR BASIC " & Err & ": " & Error$ & " (linea " & Erl & ")" & Chr$(10)
Escribir:
    n = FreeFile
    Open "{salida}" For Output As #n
    Print #n, s
    Close #n
    StarDesktop.terminate()
End Sub
"""


def leer_bas(ruta):
    with open(ruta, "rb") as f:
        texto = f.read().decode("cp1252")
    lineas = [l for l in texto.replace("\r\n", "\n").split("\n")
              if not l.startswith("Attribute ")]
    codigo = "\n".join(lineas)
    # LibreOffice Basic no admite vbObjectError en expresiones Const (VB6 si):
    # se sustituye por su valor literal (&H80040000 = -2147221504).
    codigo = re.sub(r"vbObjectError \+ (\d+)",
                    lambda m: str(-2147221504 + int(m.group(1))), codigo)
    return "Option VBASupport 1\n" + codigo


def xba(nombre, codigo):
    return ('<?xml version="1.0" encoding="UTF-8"?>\n'
            '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
            '<script:module xmlns:script="http://openoffice.org/2000/script" '
            'script:name="%s" script:language="StarBasic" script:moduleType="normal">%s</script:module>\n'
            % (nombre, escape(codigo)))


def main():
    if not shutil.which("soffice"):
        print("soffice no disponible: instale LibreOffice para ejecutar las pruebas")
        return 2
    tmp = tempfile.mkdtemp(prefix="d52_")
    perfil = os.path.join(tmp, "perfil")
    salida = os.path.join(tmp, "resultado.txt")
    env = dict(os.environ, HOME=tmp)
    url_perfil = "file://" + perfil
    # 1) crear perfil
    subprocess.run(["soffice", "--headless", "--norestore",
                    "-env:UserInstallation=" + url_perfil, "--terminate_after_init"],
                   env=env, timeout=180, check=False,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    lib = os.path.join(perfil, "user", "basic", "Standard")
    os.makedirs(lib, exist_ok=True)
    modulos = dict((n, leer_bas(r)) for n, r in MODULOS.items())
    modulos["Runner"] = RUNNER.replace("{salida}", salida)
    for nombre, codigo in modulos.items():
        with open(os.path.join(lib, nombre + ".xba"), "w", encoding="utf-8") as f:
            f.write(xba(nombre, codigo))
    with open(os.path.join(lib, "script.xlb"), "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
                '<library:library xmlns:library="http://openoffice.org/2000/library" '
                'library:name="Standard" library:readonly="false" library:passwordprotected="false">\n'
                + "".join(' <library:element library:name="%s"/>\n' % n for n in modulos)
                + '</library:library>\n')
    # 2) ejecutar
    # Un error de compilacion Basic abre un dialogo modal (cuelga en headless):
    # se limita el tiempo y se termina el proceso.
    proc = subprocess.Popen(["soffice", "--headless", "--norestore", "--invisible",
                             "-env:UserInstallation=" + url_perfil,
                             "macro:///Standard.Runner.Main"],
                            env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        proc.wait(timeout=120)
    except subprocess.TimeoutExpired:
        proc.kill()
        subprocess.run(["pkill", "-9", "-f", "UserInstallation=" + url_perfil], check=False)
    if not os.path.exists(salida):
        print("No se obtuvo resultado (error de compilacion Basic o macro no ejecutada)")
        return 1
    with open(salida, encoding="utf-8", errors="replace") as f:
        texto = f.read()
    print(texto)
    shutil.rmtree(tmp, ignore_errors=True)
    ok = "RESUMEN:" in texto and " 0 fallas" in texto
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
