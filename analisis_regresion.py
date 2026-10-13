#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
analisis_regresion.py — Regresion lineal sobre el control mensual del Centro.

Replica en Python (solo biblioteca estandar, sin pip install) el analisis que
hace el aplicativo web de `index.html` (pestana Regresion): lee los mismos
archivos, arma las mismas variables del dia y ajusta el mismo modelo OLS.

    Uso rapido
    ----------
    python analisis_regresion.py                       # modelo por defecto
    python analisis_regresion.py --listar              # ver variables disponibles
    python analisis_regresion.py --y monto_total --x cantidad_grandes
    python analisis_regresion.py --x cantidad_grandes,cantidad_medianos \\
                                 --predecir cantidad_grandes=5,cantidad_medianos=8
    python analisis_regresion.py --archivo "260901 CONTROL MENSUAL CENTRO.xlsm" \\
                                 --hoja caja
    python analisis_regresion.py --unidad item-dia --y monto --x cantidad,seccion \\
                                 --cat seccion

    Formatos de entrada
    -------------------
    1. CSV transaccional ......... fecha, dia_semana, seccion, item, cantidad, precio, monto
                                   (uno por linea). Ej.: control_mensual_sept.csv
    2. CSV / Excel matriz diaria . un concepto por fila y un dia por columna
                                   (como salen las hojas ventas, caja, gastos, insumos).
    3. .xlsx / .xlsm directo ..... se elige la hoja con --hoja (las macros se ignoran,
                                   se leen los valores ya calculados). El .xls viejo
                                   no se puede leer sin dependencias: guardalo como .xlsx.

    De un libro se leen los consolidados (ventas, caja, gastos, insumos): las hojas
    diarias `dd.mm.aa` no tienen formato de tabla y se rechazan a proposito, con un
    aviso, en vez de armar un modelo con datos mal leidos.

    Del formato matriz se saltea solo la fila/columna de totales; metricas como
    "Total recaudado" se conservan. Si hay mas de una columna de texto (la hoja
    ventas tiene SECCION e ITEM), se elige con --concepto ITEM.

    Unidad de analisis
    ------------------
    --unidad dia      (por defecto) una fila por dia  -> n = 28, Y tipica monto_total
    --unidad item-dia  una fila por linea de venta     -> n = 330, Y tipica monto

    Variables que construye sola (nivel dia)
    -----------------------------------------
    monto_total (Y sugerida), unidades_total, cantidad_grandes, cantidad_medianos,
    monto_grandes, monto_medianos, monto_toppings, monto_pasteleria, monto_bubblebath,
    monto_merchandise, monto_extras (incluye delivery), pedidos_ya, unidades_promo,
    ticket_promedio, precio_promedio_vaso, items_distintos, dia_mes, fin_de_semana y
    dia_semana (categorica). Con --listar se ven todas, con su descripcion.

    Las categoricas se codifican con k-1 dummies (se omite el nivel base) cuando el
    modelo lleva intercepto, para no caer en la colinealidad perfecta entre las
    dummies y el intercepto.

    Requisitos
    ----------
    Python 3.7 o superior. **No usa dependencias**: solo la biblioteca estandar
    (csv, json, math, zipfile, xml.etree, argparse, datetime, re). Se puede correr
    en cualquier maquina con Python, sin pip install.

    Que informa
    -----------
    Resumen de datos, avisos, ecuacion del modelo, coeficientes (beta, error estandar,
    t, p, IC 95%, beta estandarizado), bondad de ajuste (R2, R2 ajustado, sigma, F),
    error de prediccion (MAE, RMSE, MAPE, sesgo, validacion cruzada k-fold),
    diagnostico (VIF, Durbin-Watson, apalancamiento, residuo estudentizado, distancia
    de Cook, asimetria y curtosis) y prediccion con IC 95% para la media y para una
    observacion nueva, con aviso de extrapolacion.

    No escribe ningun archivo salvo que se pida --json o --escenarios-csv.

    Equivalencia con el aplicativo web
    ----------------------------------
    El motor es el mismo que el de la pestana "Regresion" de index.html (OLS por
    descomposicion QR, mismas metricas y misma mezcla de la validacion cruzada),
    verificado contra statsmodels y contra el propio motor JS con diferencias
    menores a 1e-12 (ver _val/validar_regresion_py.py).
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

VERSION = "1.0.0"

ARCHIVO_POR_DEFECTO = "control_mensual_sept.csv"
Y_POR_DEFECTO = "monto_total"
X_POR_DEFECTO = ["cantidad_grandes", "cantidad_medianos", "ticket_promedio"]

DIAS_SEMANA = ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"]


# ============================================================================
#  1. Utilidades de texto y numeros
# ============================================================================

def texto(v) -> str:
    """Cadena recortada; None y vacios -> \"\"."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:  # NaN
            return ""
        if v == int(v):
            return str(int(v))
    return str(v).strip()


def clave(v) -> str:
    """Nombre de variable normalizado: minusculas, sin espacios, parentesis ni comas.

    Las comas se vuelven puntos para que el nombre se pueda pasar por linea de
    comandos (--x, --predecir) sin chocar con el separador, por ej.
    "QR neto (-1,2%)" -> qr_neto_-1.2%.
    """
    s = texto(v).lower().replace("(", "").replace(")", "").replace(",", ".")
    return "_".join(s.split())


def es_fila_total(etiqueta) -> bool:
    """¿Es una fila/columna de totales? No confundir con metricas como "Total caja"."""
    k = clave(etiqueta)
    if not k:
        return False
    if k in ("total", "totales", "subtotal", "subtotales"):
        return True
    return bool(re.match(r"^(sub)?total(es)?[\s_\-]*([$#%]|cant|cantidad|unidades|monto|importe|general)$", k))


def num(v) -> float:
    """Convierte a numero tolerando formato espanol (1.234,56) e ingles (1,234.56)."""
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(float(v)) else 0.0
    s = texto(v)
    if not s:
        return 0.0
    s = "".join(ch for ch in s if ch.isdigit() or ch in ".,-")
    if not s:
        return 0.0
    ultima_coma = s.rfind(",")
    ultimo_punto = s.rfind(".")
    if ultima_coma > ultimo_punto:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def es_numero(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(float(v))


def fmt(v, dec: int = 4) -> str:
    """Numero con separador de miles '.' y decimal ',' (formato espanol)."""
    if v is None:
        return "-"
    if isinstance(v, float) and not math.isfinite(v):
        return "-"
    s = "{:,.{d}f}".format(float(v), d=dec)
    return s.replace(",", "\u00a0").replace(".", ",").replace("\u00a0", ".")


def fmtp(p) -> str:
    """p-valor en notacion legible."""
    if p is None or not math.isfinite(p):
        return "-"
    if p < 1e-4:
        return "< 0,0001"
    return fmt(p, 5)


# ============================================================================
#  2. Distribuciones (t de Student y F) — equivalen a js/estadistica.js
# ============================================================================

def log_gamma(x: float) -> float:
    """log(Gamma(x)); usa math.lgamma (biblioteca estandar, alta precision)."""
    if x <= 0:
        raise ValueError("log_gamma requiere x > 0")
    return math.lgamma(x)


def beta_inc_cf(a: float, b: float, x: float) -> float:
    """Fraccion continua para la beta incompleta (Numerical Recipes)."""
    fpmim = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - (qab * x) / qap
    if abs(d) < fpmim:
        d = fpmim
    d = 1.0 / d
    h = d
    for m in range(1, 301):
        m2 = 2 * m
        aa = (m * (b - m) * x) / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < fpmim:
            d = fpmim
        c = 1.0 + aa / c
        if abs(c) < fpmim:
            c = fpmim
        d = 1.0 / d
        h *= d * c
        aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < fpmim:
            d = fpmim
        c = 1.0 + aa / c
        if abs(c) < fpmim:
            c = fpmim
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < 3e-14:
            break
    return h


def betai(a: float, b: float, x: float) -> float:
    """Beta incompleta regularizada I_x(a, b)."""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    bt = math.exp(
        log_gamma(a + b) - log_gamma(a) - log_gamma(b) + a * math.log(x) + b * math.log(1.0 - x)
    )
    if x < (a + 1.0) / (a + b + 2.0):
        return (bt * beta_inc_cf(a, b, x)) / a
    return 1.0 - (bt * beta_inc_cf(b, a, 1.0 - x)) / b


def tcdf(t: float, df: float) -> float:
    """Funcion de distribucion de la t de Student."""
    if df <= 0:
        return float("nan")
    x = df / (df + t * t)
    ib = betai(df / 2.0, 0.5, x)
    return 1.0 - ib / 2.0 if t > 0 else ib / 2.0


def tinv(p: float, df: float) -> float:
    """Inversa de la t por biseccion (200 iteraciones, sobra para IC)."""
    if p <= 0:
        return -math.inf
    if p >= 1:
        return math.inf
    lo, hi = -1e3, 1e3
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if tcdf(mid, df) < p:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2.0


def fcdf(f: float, d1: float, d2: float) -> float:
    """Funcion de distribucion F."""
    if f <= 0:
        return 0.0
    return betai(d1 / 2.0, d2 / 2.0, (d1 * f) / (d1 * f + d2))


# ============================================================================
#  3. Algebra lineal (mismas rutinas que el motor JS)
# ============================================================================

def transpose(a):
    if not a:
        return []
    m, n = len(a), len(a[0])
    return [[a[i][j] for i in range(m)] for j in range(n)]


def matmul(a, b):
    m, p, n = len(a), len(b), len(b[0])
    out = [[0.0] * n for _ in range(m)]
    for i in range(m):
        oi = out[i]
        ai = a[i]
        for k in range(p):
            aik = ai[k]
            if aik == 0:
                continue
            bk = b[k]
            for j in range(n):
                oi[j] += aik * bk[j]
    return out


def identidad(n):
    return [[1.0 if i == j else 0.0 for j in range(n)] for i in range(n)]


def qr_householder(a_in):
    """Descomposicion QR por reflexiones de Householder. Devuelve (Q, R)."""
    m, n = len(a_in), len(a_in[0])
    r = [row[:] for row in a_in]
    q = identidad(m)
    for k in range(n):
        norm = math.sqrt(sum(r[i][k] * r[i][k] for i in range(k, m)))
        if norm < 1e-300:
            continue
        alpha = -norm if r[k][k] > 0 else norm
        v = [0.0] * m
        v[k] = r[k][k] - alpha
        for i in range(k + 1, m):
            v[i] = r[i][k]
        vnorm2 = sum(v[i] * v[i] for i in range(k, m))
        if vnorm2 < 1e-300:
            continue
        # R <- H . R
        for j in range(k, n):
            s = 0.0
            for i in range(k, m):
                s += v[i] * r[i][j]
            s = (2.0 * s) / vnorm2
            for i in range(k, m):
                r[i][j] -= s * v[i]
        # Q <- Q . H
        for i in range(m):
            s = 0.0
            for j in range(k, m):
                s += q[i][j] * v[j]
            s = (2.0 * s) / vnorm2
            for j in range(k, m):
                q[i][j] -= s * v[j]
    return q, r


def solve_upper(r, b):
    """Resuelve R . x = b con R triangular superior cuadrada."""
    n = len(r)
    x = [0.0] * n
    for i in range(n - 1, -1, -1):
        s = b[i]
        for j in range(i + 1, n):
            s -= r[i][j] * x[j]
        if abs(r[i][i]) < 1e-300:
            raise ValueError("matriz singular (columna constante o colineal)")
        x[i] = s / r[i][i]
    return x


def invertir_upper(r):
    """Inversa de una triangular superior cuadrada."""
    n = len(r)
    inv = identidad(n)
    for c in range(n):
        col = [0.0] * n
        col[c] = 1.0
        x = solve_upper(r, col)
        for i in range(n):
            inv[i][c] = x[i]
    return inv


def xtx_inversa(xa, p):
    """(X'X)^-1 a partir del QR de la matriz de diseno."""
    _q, r = qr_householder(xa)
    r11 = [r[i][:p] for i in range(p)]
    rinv = invertir_upper(r11)
    return matmul(rinv, transpose(rinv))


# ============================================================================
#  4. Regresion por minimos cuadrados (OLS)
# ============================================================================

def ols(y, x, intercepto=True, nombres_x=None):
    """Ajusta y = b0 + b1·X1 + ... por QR. x es una lista de filas (sin intercepto)."""
    n = len(y)
    k = len(x[0]) if x else 0
    if nombres_x is None:
        nombres_x = ["X" + str(j + 1) for j in range(k)]

    xa = [[1.0] + row[:] for row in x] if intercepto else [row[:] for row in x]
    p = k + 1 if intercepto else k
    if n < p + 1:
        raise ValueError("n insuficiente: %d filas para %d parametros" % (n, p))

    q, r = qr_householder(xa)
    qt = transpose(q)
    qty = [fila[0] for fila in matmul(qt, [[v] for v in y])]
    r11 = [r[i][:p] for i in range(p)]
    beta = solve_upper(r11, qty[:p])

    ajustado = [sum(v * beta[j] for j, v in enumerate(row)) for row in xa]
    resid = [y[i] - ajustado[i] for i in range(n)]

    ybar = sum(y) / n
    sst = sum((v - ybar) ** 2 for v in y)
    sse = sum(e * e for e in resid)
    df = n - p
    sigma2 = sse / df
    sigma = math.sqrt(sigma2)

    rinv = invertir_upper(r11)
    vcov = [[v * sigma2 for v in row] for row in matmul(rinv, transpose(rinv))]
    se = [math.sqrt(max(0.0, vcov[j][j])) for j in range(p)]

    r2 = 1.0 - sse / sst if sst > 0 else 0.0
    r2adj = 1.0 - (1.0 - r2) * (n - (1 if intercepto else 0)) / df if df > 0 else float("nan")

    f_glob = fp_glob = float("nan")
    if intercepto and k > 0 and df > 0:
        den = (1.0 - r2) / df
        f_glob = (r2 / k) / den if den > 0 else math.inf
        fp_glob = 1.0 - fcdf(f_glob, k, df) if math.isfinite(f_glob) else 0.0

    tcrit = tinv(0.975, df)
    sdy = math.sqrt(sst / (n - 1)) if n > 1 else float("nan")
    coef = []
    for j in range(p):
        b = beta[j]
        es_int = (j == 0 and intercepto)
        t = b / se[j] if se[j] > 0 else float("nan")
        pv = float("nan") if not math.isfinite(t) else 2.0 * (1.0 - tcdf(abs(t), df))
        std_b = float("nan")
        if not es_int:
            col = [row[j - 1] for row in x]
            m = sum(col) / n
            sd = math.sqrt(sum((v - m) ** 2 for v in col) / (n - 1)) if n > 1 else float("nan")
            std_b = (b * sd) / sdy if sdy and math.isfinite(sdy) and sdy > 0 else float("nan")
        coef.append({
            "nombre": "(Intercepto)" if es_int else nombres_x[j - (1 if intercepto else 0)],
            "beta": b,
            "se": se[j],
            "t": t,
            "p": pv,
            "ci_lo": b - tcrit * se[j],
            "ci_hi": b + tcrit * se[j],
            "beta_std": std_b,
        })

    return {
        "n": n, "k": k, "p": p, "df": df, "intercepto": intercepto,
        "beta": beta, "coef": coef, "ajustado": ajustado, "resid": resid,
        "sigma": sigma, "sigma2": sigma2, "r2": r2, "r2adj": r2adj,
        "sst": sst, "sse": sse, "F": f_glob, "Fp": fp_glob, "dfModelo": k,
        "vcov": vcov, "Xa": xa, "nombresX": nombres_x, "ybar": ybar, "tcrit": tcrit,
    }


def diagnostico(modelo):
    """Apalancamiento, residuo estudentizado, distancia de Cook y outliers (Cook > 4/n)."""
    xa, resid, sigma2, p = modelo["Xa"], modelo["resid"], modelo["sigma2"], modelo["p"]
    n = len(xa)
    xtx_inv = xtx_inversa(xa, p)
    hat = []
    for row in xa:
        s = 0.0
        for i in range(p):
            s2 = sum(xtx_inv[i][j] * row[j] for j in range(p))
            s += row[i] * s2
        hat.append(s)
    estudent = [
        resid[i] / math.sqrt(sigma2 * max(1e-12, 1.0 - hat[i])) for i in range(n)
    ]
    cook = []
    for i in range(n):
        h = min(hat[i], 0.999999)
        cook.append((resid[i] ** 2 / (p * sigma2)) * (h / ((1.0 - h) ** 2)))
    return {
        "hat": hat,
        "estudent": estudent,
        "cook": cook,
        "outliers": [i for i in range(n) if cook[i] > 4.0 / n],
    }


def durbin_watson(resid):
    num_, den = 0.0, 0.0
    for i, e in enumerate(resid):
        den += e * e
        if i > 0:
            num_ += (resid[i] - resid[i - 1]) ** 2
    return num_ / den if den > 0 else float("nan")


def asimetria_curtosis(resid):
    n = len(resid)
    if not n:
        return float("nan"), float("nan")
    m = sum(resid) / n
    m2 = m3 = m4 = 0.0
    for r0 in resid:
        d = r0 - m
        m2 += d * d
        m3 += d ** 3
        m4 += d ** 4
    m2 /= n
    m3 /= n
    m4 /= n
    if m2 <= 0:
        return float("nan"), float("nan")
    return m3 / m2 ** 1.5, m4 / (m2 * m2) - 3.0


def vif(x):
    """VIF y tolerancia de cada columna de X (sin intercepto)."""
    k = len(x[0]) if x else 0
    salida = []
    for j in range(k):
        yj = [row[j] for row in x]
        otras = [[v for c, v in enumerate(row) if c != j] for row in x]
        r2 = 0.0
        if otras and otras[0]:
            try:
                r2 = ols(yj, otras, intercepto=True)["r2"]
            except ValueError:
                r2 = 0.0
        salida.append({"vif": 1.0 / (1.0 - r2) if r2 < 1 else math.inf, "tol": 1.0 - r2, "r2": r2})
    return salida


SEMILLA_MEZCLA = 20260901


def mezclar_indices(n, semilla=SEMILLA_MEZCLA):
    """Indices mezclados con la misma secuencia pseudoaleatoria que el motor JS.

    La mezcla replica el calculo de JavaScript, que hace la multiplicacion en
    punto flotante de 64 bits antes de recortar a entero; asi la particion de la
    validacion cruzada coincide exactamente con la del aplicativo web.
    """
    idx = list(range(n))
    for i in range(n - 1, 0, -1):
        semilla = (int(float(semilla) * 1103515245.0 + 12345.0) % (1 << 32)) & 0x7FFFFFFF
        j = semilla % (i + 1)
        idx[i], idx[j] = idx[j], idx[i]
    return idx


def kfold(y, x, k=5, intercepto=True):
    """Validacion cruzada k-fold con mezcla determinista (misma semilla que la web)."""
    k = max(2, int(k))
    n = len(y)
    idx = mezclar_indices(n)

    pred = [float("nan")] * n
    tam = math.ceil(n / k)
    for f in range(k):
        prueba = idx[f * tam:(f + 1) * tam]
        if not prueba:
            continue
        restantes = set(prueba)
        train = [i for i in idx if i not in restantes]
        if len(train) < (2 if intercepto else 1):
            continue
        try:
            m = ols([y[i] for i in train], [x[i] for i in train], intercepto=intercepto)
        except ValueError as e:
            return {"ok": False, "error": str(e)}
        base = 1 if intercepto else 0
        for i in prueba:
            pred[i] = sum(v * m["beta"][j + base] for j, v in enumerate(x[i])) + (
                m["beta"][0] if intercepto else 0.0
            )

    ybar = sum(y) / n
    ss_res = ss_tot = ae = 0.0
    m_ = 0
    for i in range(n):
        if not math.isfinite(pred[i]):
            continue
        e = y[i] - pred[i]
        ss_res += e * e
        ss_tot += (y[i] - ybar) ** 2
        ae += abs(e)
        m_ += 1
    if not m_:
        return {"ok": False, "error": "sin predicciones fuera de muestra"}
    return {
        "ok": True, "k": k, "n": m_, "pred": pred,
        "r2": 1.0 - ss_res / ss_tot if ss_tot > 0 else float("nan"),
        "rmse": math.sqrt(ss_res / m_),
        "mae": ae / m_,
    }


def predecir(modelo, x0):
    """Y estimada con IC 95% para la media y para una observacion nueva."""
    beta, xa, sigma, tcrit, intercepto = (
        modelo["beta"], modelo["Xa"], modelo["sigma"], modelo["tcrit"], modelo["intercepto"]
    )
    fila = ([1.0] + list(x0)) if intercepto else list(x0)
    yhat = sum(v * beta[j] for j, v in enumerate(fila))
    xtx_inv = xtx_inversa(xa, modelo["p"])
    cuad = 0.0
    for i in range(modelo["p"]):
        cuad += fila[i] * sum(xtx_inv[i][j] * fila[j] for j in range(modelo["p"]))
    cuad = max(0.0, cuad)
    se_media = sigma * math.sqrt(cuad)
    se_nueva = sigma * math.sqrt(1.0 + cuad)
    return {
        "yhat": yhat,
        "se_media": se_media,
        "se_nueva": se_nueva,
        "ic_media": [yhat - tcrit * se_media, yhat + tcrit * se_media],
        "ic_nueva": [yhat - tcrit * se_nueva, yhat + tcrit * se_nueva],
    }


# ============================================================================
#  5. Lectura de archivos (CSV y .xlsx/.xlsm, sin dependencias)
# ============================================================================

ALIAS = {
    "fecha": ["fecha", "date", "dia", "día", "fecha_dia"],
    "dia_semana": ["dia_semana", "diasemana", "weekday", "dayofweek", "dia_de_la_semana", "ds"],
    "seccion": ["seccion", "sección", "section", "categoria", "categoría", "rubro"],
    "item": ["item", "producto", "nombre", "articulo", "artículo", "concepto", "detalle"],
    "cantidad": ["cantidad", "qty", "quantity", "unidades", "cant", "vasos"],
    "precio": ["precio", "price", "precio_u", "punit", "unitario"],
    "monto": ["monto", "total", "importe", "amount", "venta", "ventas", "valor", "subtotal"],
}


def canon(k) -> str:
    """Devuelve el nombre canonico de una columna (o "" si no la reconoce)."""
    n = clave(k)
    for c, alias in ALIAS.items():
        if n in alias:
            return c
    return ""


def partes_fecha(f):
    """(anio, mes, dia) de una fecha en texto, o None."""
    s = texto(f)
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})", s)
    if m:
        y = int(m.group(3))
        if y < 100:
            y += 2000
        return y, int(m.group(2)), int(m.group(1))
    return None


def fecha_iso(f) -> str:
    """Normaliza a AAAA-MM-DD cuando se puede; si no, devuelve el texto original."""
    p = partes_fecha(f)
    if not p:
        return texto(f)
    return "%04d-%02d-%02d" % p


def dia_semana_de_fecha(f) -> str:
    p = partes_fecha(f)
    if not p:
        return ""
    return dia_semana_de(p[0], p[1], p[2])


def dia_semana_de(y, mes, d) -> str:
    """Nombre del dia de la semana (Domingo..Sabado)."""
    from datetime import date
    return DIAS_SEMANA[(date(y, mes, d).weekday() + 1) % 7]


def fecha_de_serial(n: float) -> str:
    """Numero de serie de Excel -> AAAA-MM-DD (base 1899-12-30)."""
    from datetime import date, timedelta
    return (date(1899, 12, 30) + timedelta(days=int(round(n)))).isoformat()


def es_etiqueta_dia(v) -> bool:
    """¿La celda parece un dia: 1..31, una fecha corta o un serial de Excel?"""
    if es_numero(v):
        n = float(v)
        if n == int(n) and 1 <= n <= 31:
            return True
        return 20000 <= n <= 80000  # serial de Excel (1954-2119)
    s = texto(v)
    if not s:
        return False
    if s.isdigit() and 1 <= int(s) <= 31:
        return True
    return partes_fecha(s) is not None


def etiqueta_a_fecha(et) -> str:
    """Rotulo de columna de dia -> clave de fecha."""
    if es_numero(et) and float(et) >= 20000:
        return fecha_de_serial(float(et))
    s = texto(et)
    if s.isdigit():
        return s.zfill(2)
    return fecha_iso(s)


def detectar_delimitador(muestra: str) -> str:
    linea = muestra.splitlines()[0] if muestra.splitlines() else muestra
    candidatos = [",", ";", "\t", "|"]
    conteos = [(linea.count(c), c) for c in candidatos]
    mejor = max(conteos)
    return mejor[1] if mejor[0] > 0 else ","


def leer_csv(path):
    """CSV -> matriz de celdas (texto), con BOM y delimitador detectados."""
    with open(path, "r", encoding="utf-8-sig", newline="") as fh:
        muestra = fh.read(65536)
        fh.seek(0)
        delim = detectar_delimitador(muestra)
        return [fila for fila in csv.reader(fh, delimiter=delim)]


# --- .xlsx / .xlsm: se leen con zipfile + ElementTree (la libreria estandar) ---

XML_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XML_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
XML_PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _col_idx(ref: str) -> int:
    letras = ""
    for ch in ref:
        if ch.isalpha():
            letras += ch
        else:
            break
    n = 0
    for ch in letras.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _texto_elementos(nodo) -> str:
    return "".join((t.text or "") for t in nodo.iter(XML_MAIN + "t"))


def iter_hojas_xlsx(path, hoja=None):
    """Genera (nombre_hoja, matriz) de un libro .xlsx/.xlsm, hoja por hoja."""
    if not zipfile.is_zipfile(path):
        raise ValueError("El archivo no es un libro de Excel (.xlsx/.xlsm).")
    with zipfile.ZipFile(path) as zf:
        nombres_zip = set(zf.namelist())

        # 1) cadenas compartidas
        compartidas = []
        if "xl/sharedStrings.xml" in nombres_zip:
            raiz = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in raiz.iter(XML_MAIN + "si"):
                compartidas.append(_texto_elementos(si))

        # 2) mapa id de relacion -> ruta de la hoja
        relaciones = {}
        if "xl/_rels/workbook.xml.rels" in nombres_zip:
            raiz = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in raiz.iter(XML_PKG + "Relationship"):
                destino = rel.get("Target") or ""
                if destino.startswith("/"):
                    destino = destino[1:]
                elif not destino.startswith("xl/"):
                    destino = "xl/" + destino.lstrip("./")
                relaciones[rel.get("Id")] = destino

        # 3) hojas, en el orden del libro
        raiz = ET.fromstring(zf.read("xl/workbook.xml"))
        orden = []
        for sh in raiz.iter(XML_MAIN + "sheet"):
            rid = sh.get(XML_REL + "id")
            orden.append((sh.get("name") or "Hoja", relaciones.get(rid, "")))

        pedidas = orden
        if hoja:
            pedidas = [s for s in orden if clave(s[0]) == clave(hoja)]
            if not pedidas:
                disponibles = ", ".join(s[0] for s in orden)
                raise ValueError("No existe la hoja '%s'. Disponibles: %s" % (hoja, disponibles))

        # 4) celdas de cada hoja
        for nombre, ruta in pedidas:
            if ruta not in nombres_zip:
                continue
            raiz = ET.fromstring(zf.read(ruta))
            filas = []
            for fila in raiz.iter(XML_MAIN + "row"):
                celdas = []
                sig = 0
                for c in fila.iter(XML_MAIN + "c"):
                    ref = c.get("r") or ""
                    idx = _col_idx(ref) if ref else sig
                    sig = idx + 1
                    while len(celdas) < idx:
                        celdas.append(None)
                    tipo = c.get("t") or "n"
                    v = c.find(XML_MAIN + "v")
                    if tipo == "s":
                        valor = compartidas[int(v.text)] if v is not None and v.text else ""
                    elif tipo == "inlineStr":
                        valor = _texto_elementos(c)
                    elif tipo == "str":
                        valor = (v.text or "") if v is not None else ""
                    elif tipo == "b":
                        valor = 1 if (v is not None and v.text == "1") else 0
                    else:
                        valor = num(v.text) if v is not None and v.text not in (None, "") else None
                    celdas.append(valor)
                filas.append(celdas)
            ancho = max((len(f) for f in filas), default=0)
            for f in filas:
                while len(f) < ancho:
                    f.append(None)
            yield nombre, filas


def leer_xlsx(path, hoja=None):
    """Libro .xlsx/.xlsm -> lista de (nombre_hoja, matriz). Sin macros, solo valores."""
    return list(iter_hojas_xlsx(path, hoja))


# ============================================================================
#  6. Reconocimiento del formato y construccion de las variables
# ============================================================================

DESC = {
    "monto_total": "Venta del dia (monto total)",
    "unidades_total": "Unidades vendidas en el dia",
    "cantidad_grandes": "Vasos grandes vendidos",
    "cantidad_medianos": "Vasos medianos vendidos",
    "monto_grandes": "Monto en seccion GRANDES",
    "monto_medianos": "Monto en seccion MEDIANOS",
    "monto_toppings": "Monto en TOPPINGS",
    "monto_pasteleria": "Monto en PASTELERIA",
    "monto_bubblebath": "Monto en BUBBLEBATH",
    "monto_merchandise": "Monto en MERCHANDISE",
    "monto_extras": "Monto en EXTRAS (incluye delivery)",
    "monto_delivery": "Pedidos Ya (monto = cantidad de pedidos)",
    "pedidos_ya": "Pedidos por delivery (cantidad)",
    "unidades_promo": "Unidades en promocion",
    "ticket_promedio": "Ticket promedio (monto por unidad)",
    "precio_promedio_vaso": "Precio promedio por vaso (grande/mediano)",
    "items_distintos": "Cantidad de items distintos del dia",
    "dia_mes": "Dia del mes (tendencia)",
    "fin_de_semana": "Fin de semana (0/1)",
    "dia_semana": "Dia de la semana (categorica)",
    "fecha": "Fecha",
}

EXTERNAS = {
    "efectivo", "qr_neto", "qr", "tarjeta_neto", "tarjeta", "transferencia",
    "total_recaudado", "total_gastos", "gastos", "insumos_total", "insumos",
    "diferencia", "resultado", "diferencia_vasos", "domos", "selladora",
    "caja_inicial", "caja_real",
}


def es_fila_vacia(fila) -> bool:
    return not fila or all(texto(v) == "" for v in fila)


def es_transaccional(encabezado) -> bool:
    c = [canon(h) for h in encabezado]
    return "fecha" in c and ("monto" in c or "cantidad" in c)


def detectar_columnas_dias(encabezado, col_etq=0):
    """Columnas de un encabezado que son dias, con las verificaciones de una matriz.

    Ademas de parecer dias, se exige que: sean al menos 3, no se repitan, avancen
    en orden creciente y que despues de la ultima solo haya celdas vacias o de
    totales. Asi una hoja diaria cualquiera no se confunde con un consolidado.
    """
    cols = []
    for j, celda in enumerate(encabezado):
        if j == col_etq or not es_etiqueta_dia(celda) or es_fila_total(celda):
            continue
        cols.append((j, celda))
    if len(cols) < 3:
        return []
    claves = [etiqueta_a_fecha(et) for _j, et in cols]
    if len(set(claves)) != len(claves):
        return []
    if any(claves[i] >= claves[i + 1] for i in range(len(claves) - 1)):
        return []
    for j in range(cols[-1][0] + 1, len(encabezado)):
        celda = encabezado[j]
        if texto(celda) and not es_fila_total(celda) and not es_etiqueta_dia(celda):
            return []
    return cols


def es_matriz(encabezado) -> bool:
    return len(detectar_columnas_dias(encabezado)) >= 3


def buscar_encabezado(filas, max_filas=15):
    """Ubica la fila de encabezado (los consolidados traen titulos arriba).

    Devuelve (indice, formato), con formato "transaccional", "matriz" o None.
    """
    if not filas:
        return 0, None
    ini = 0
    while ini < len(filas) and es_fila_vacia(filas[ini]):
        ini += 1
    for k in range(ini, min(len(filas), ini + max_filas)):
        head = [texto(v) for v in filas[k]]
        if es_transaccional(head):
            return k, "transaccional"
        if es_matriz(head):
            return k, "matriz"
    return ini, None


def construir_dataset(filas, nombre, concepto=None):
    """Matriz de celdas -> dataset transaccional o matriz diaria (o None)."""
    if not filas:
        return None
    ini, formato = buscar_encabezado(filas)
    if formato is None:
        return None
    if formato == "transaccional":
        encabezado = [texto(v) for v in filas[ini]]
        cuerpo = [f for f in filas[ini + 1:] if not es_fila_vacia(f)]
        return dataset_transaccional(encabezado, cuerpo, nombre)
    return dataset_matriz(filas[ini:], nombre, concepto)


def dataset_transaccional(encabezado, cuerpo, nombre):
    mapa = {}
    for j, h in enumerate(encabezado):
        c = canon(h)
        if c and c not in mapa:
            mapa[c] = j
    if "fecha" not in mapa:
        return None

    filas = []
    for r in cuerpo:
        f = fecha_iso(r[mapa["fecha"]] if mapa["fecha"] < len(r) else "")
        if not f:
            continue
        def celda(campo):
            j = mapa.get(campo)
            return r[j] if j is not None and j < len(r) else ""
        item = texto(celda("item"))
        monto = num(celda("monto"))
        cant = num(celda("cantidad"))
        if not item and not monto and not cant:
            continue
        if es_fila_total(item):
            continue
        filas.append({
            "fecha": f,
            "dia_semana": texto(celda("dia_semana")) or dia_semana_de_fecha(f),
            "seccion": texto(celda("seccion")).upper(),
            "item": item,
            "cantidad": cant,
            "precio": num(celda("precio")),
            "monto": monto,
        })
    if not filas:
        return None

    dias = agregar_por_dia(filas)
    return {
        "nombre": nombre, "formato": "transaccional", "filas": filas,
        "dias": dias, "variables": catalogo(dias, filas), "meta": meta(dias, filas),
    }


def agregar_por_dia(filas):
    """Lineas de venta -> una fila por dia con las variables derivadas."""
    por_fecha = {}
    for r in filas:
        d = por_fecha.get(r["fecha"])
        if d is None:
            d = {
                "fecha": r["fecha"], "dia_semana": r["dia_semana"],
                "monto_total": 0.0, "unidades_total": 0.0,
                "cantidad_grandes": 0.0, "cantidad_medianos": 0.0,
                "monto_grandes": 0.0, "monto_medianos": 0.0,
                "monto_toppings": 0.0, "monto_pasteleria": 0.0, "monto_bubblebath": 0.0,
                "monto_merchandise": 0.0, "monto_extras": 0.0,
                "pedidos_ya": 0.0, "unidades_promo": 0.0,
                "_items": set(),
            }
            por_fecha[r["fecha"]] = d
        d["monto_total"] += r["monto"]
        d["unidades_total"] += r["cantidad"]
        if r["item"]:
            d["_items"].add(r["item"])
        sec = r["seccion"]
        if sec == "GRANDES":
            d["monto_grandes"] += r["monto"]
            d["cantidad_grandes"] += r["cantidad"]
        elif sec == "MEDIANOS":
            d["monto_medianos"] += r["monto"]
            d["cantidad_medianos"] += r["cantidad"]
        elif sec == "TOPPINGS":
            d["monto_toppings"] += r["monto"]
        elif sec == "PASTELERIA":
            d["monto_pasteleria"] += r["monto"]
        elif sec == "BUBBLEBATH":
            d["monto_bubblebath"] += r["monto"]
        elif sec == "MERCHANDISE":
            d["monto_merchandise"] += r["monto"]
        elif sec in ("EXTRAS", "INSUMOS"):
            d["monto_extras"] += r["monto"]
        if "pedidos ya" in r["item"].lower():
            d["pedidos_ya"] += r["cantidad"]
        if "promo" in r["item"].lower():
            d["unidades_promo"] += r["cantidad"]

    dias = sorted(por_fecha.values(), key=lambda d: d["fecha"])
    for d in dias:
        p = partes_fecha(d["fecha"]) or (0, 0, 0)
        d["dia_mes"] = float(p[2]) if p[2] else float("nan")
        d["fin_de_semana"] = 1.0 if d["dia_semana"] in ("Sabado", "Domingo") else 0.0
        d["ticket_promedio"] = d["monto_total"] / d["unidades_total"] if d["unidades_total"] else 0.0
        vasos = d["cantidad_grandes"] + d["cantidad_medianos"]
        d["precio_promedio_vaso"] = (d["monto_grandes"] + d["monto_medianos"]) / vasos if vasos else 0.0
        d["items_distintos"] = float(len(d["_items"]))
        d["monto_delivery"] = d["pedidos_ya"]
        del d["_items"]
    return dias


def resolver_columna_concepto(encabezado, spec):
    """--concepto acepta un numero de columna (1 = primera) o el nombre del encabezado."""
    if spec is None or texto(spec) == "":
        return 0
    s = texto(spec)
    if s.isdigit():
        j = int(s) - 1
        if 0 <= j < len(encabezado):
            return j
        raise ValueError("La columna de concepto %s esta fuera de rango (hay %d columnas)."
                         % (s, len(encabezado)))
    for j, h in enumerate(encabezado):
        if clave(h) == clave(s):
            return j
    raise ValueError("No encontre la columna de concepto '%s'. Encabezados: %s" % (
        s, ", ".join(h for h in encabezado if texto(h))))


def dataset_matriz(filas, nombre, concepto=None):
    """Matriz diaria (concepto por fila, un dia por columna) -> una fila por dia."""
    encabezado = [texto(v) for v in filas[0]]
    cuerpo = [f for f in filas[1:] if not es_fila_vacia(f)]
    col_etq = resolver_columna_concepto(encabezado, concepto)
    col_dias = detectar_columnas_dias(encabezado, col_etq)
    if len(col_dias) < 3:
        return None

    claves_fecha = [etiqueta_a_fecha(et) for _j, et in col_dias]
    por_fecha = {}
    for k in claves_fecha:
        por_fecha.setdefault(k, {"fecha": k, "dia_semana": "", "_vars": {}})

    for r in cuerpo:
        etiqueta = texto(r[col_etq]) if (r and col_etq < len(r)) else ""
        if not etiqueta or es_fila_total(etiqueta):
            continue
        concepto = etiqueta
        k = clave(concepto) or concepto
        for (j, _et), fk in zip(col_dias, claves_fecha):
            v = r[j] if j < len(r) else None
            if v is None or texto(v) == "":
                continue
            rec = por_fecha[fk]
            rec["_vars"][k] = num(v)
            rec["_vars"]["_" + k] = concepto

    dias = sorted(por_fecha.values(), key=lambda d: d["fecha"])
    if not dias:
        return None
    for i, d in enumerate(dias):
        p = partes_fecha(d["fecha"])
        d["dia_semana"] = dia_semana_de_fecha(d["fecha"]) if p else ""
        d["dia_mes"] = float(p[2]) if p else float(i + 1)
        d["fin_de_semana"] = 1.0 if d["dia_semana"] in ("Sabado", "Domingo") else 0.0

    claves = set()
    for d in dias:
        claves.update(k for k in d["_vars"] if not k.startswith("_"))
    for d in dias:
        for k in claves:
            d["_vars"].setdefault(k, 0.0)

    dias_planos = []
    for d in dias:
        o = {
            "fecha": d["fecha"], "dia_semana": d["dia_semana"],
            "dia_mes": d["dia_mes"], "fin_de_semana": d["fin_de_semana"],
        }
        for k in claves:
            o[k] = d["_vars"][k]
        dias_planos.append(o)

    filas_item = []
    for d in dias:
        for k in claves:
            v = d["_vars"][k]
            if v == 0:
                continue
            filas_item.append({
                "fecha": d["fecha"], "dia_semana": d["dia_semana"], "seccion": "",
                "item": d["_vars"].get("_" + k) or k,
                "cantidad": v, "precio": 0.0, "monto": v,
            })

    return {
        "nombre": nombre, "formato": "matriz", "filas": filas_item,
        "dias": dias_planos, "variables": catalogo(dias_planos, filas_item),
        "meta": meta(dias_planos, filas_item),
    }


def catalogo(dias, filas):
    """Variables disponibles por unidad de analisis (dia / item-dia)."""
    muestra = dias[0] if dias else {}
    vars_dia = []
    for k, v in muestra.items():
        if k.startswith("_") or k == "fecha":
            continue
        tipo = "categorica" if k == "dia_semana" else ("numerica" if es_numero(v) else "texto")
        vars_dia.append({
            "clave": k, "nombre": k, "tipo": tipo,
            "rol": "y" if k == "monto_total" else "x",
            "descripcion": DESC.get(k, ""),
            "externa": k in EXTERNAS,
        })
    vars_dia.sort(key=lambda v: (0 if v["clave"] == "monto_total" else 1, v["clave"]))

    mf = filas[0] if filas else {}
    candidatas = [
        {"clave": "monto", "tipo": "numerica", "rol": "y", "descripcion": "Monto de la linea item-dia"},
        {"clave": "cantidad", "tipo": "numerica", "rol": "x", "descripcion": "Unidades de la linea"},
        {"clave": "precio", "tipo": "numerica", "rol": "x", "descripcion": "Precio unitario"},
        {"clave": "seccion", "tipo": "categorica", "rol": "x", "descripcion": "Seccion del item"},
        {"clave": "item", "tipo": "categorica", "rol": "x", "descripcion": "Item vendido"},
        {"clave": "dia_semana", "tipo": "categorica", "rol": "x", "descripcion": "Dia de la semana"},
        {"clave": "dia_mes", "tipo": "numerica", "rol": "x", "descripcion": "Dia del mes (si existe)"},
    ]
    vars_item = []
    for v in candidatas:
        if v["clave"] in ("monto", "dia_mes"):
            if v["clave"] in mf:
                vars_item.append(dict(v, nombre=v["clave"], externa=False))
        elif v["clave"] in mf and mf[v["clave"]] not in ("", None):
            vars_item.append(dict(v, nombre=v["clave"], externa=False))
    return {"dia": vars_dia, "itemDia": vars_item}


def meta(dias, filas):
    total = sum(d.get("monto_total", 0.0) for d in dias)
    montos = [d["monto_total"] for d in dias if es_numero(d.get("monto_total"))]
    return {
        "nFilas": len(filas),
        "nDias": len(dias),
        "totalMonto": total,
        "promedio": total / len(montos) if montos else 0.0,
        "min": min(montos) if montos else 0.0,
        "max": max(montos) if montos else 0.0,
        "items": len({f["item"] for f in filas}) if filas else 0,
        "secciones": len({f["seccion"] for f in filas if f["seccion"]}) if filas else 0,
    }


def construir_matriz(ds, unidad, y_clave, x_vars, intercepto=True):
    """Arma Y y X (con dummies si hace falta), filtra faltantes y avisa.

    Las categoricas se codifican con k-1 dummies (se omite el nivel base) cuando
    el modelo lleva intercepto; asi se evita la colinealidad perfecta entre las
    dummies y el intercepto (trampa de las variables ficticias).
    """
    filas = ds["filas"] if unidad == "itemDia" else ds["dias"]
    if not filas:
        raise ValueError("No hay filas para la unidad elegida.")
    avisos = []

    usadas = []
    for fila in filas:
        yr = fila.get(y_clave)
        if yr is None or texto(yr) == "":
            continue
        usadas.append(fila)
    faltantes_y = len(filas) - len(usadas)
    if faltantes_y:
        avisos.append("%d fila(s) sin Y se excluyeron." % faltantes_y)
    if not usadas:
        raise ValueError("La variable Y '%s' no tiene valores en las filas." % y_clave)

    y = [num(f.get(y_clave)) for f in usadas]
    columnas = []
    niveles = {}
    for v in x_vars:
        k = v["clave"]
        if v["tipo"] == "categorica":
            valores = [texto(f.get(k)) for f in usadas]
            uniq = sorted({s for s in valores if s != ""})
            niveles[k] = uniq
            usados = uniq[1:] if (intercepto and len(uniq) > 1) else uniq
            if intercepto and len(uniq) > 1:
                avisos.append("Categorica '%s': se usa '%s' como nivel base (omitido)." % (k, uniq[0]))
            for li, lvl in enumerate(usados):
                columnas.append({
                    "nombre": k + " = " + lvl,
                    "valores": [1.0 if s == lvl else 0.0 for s in valores],
                    "categoria": k, "nivel": lvl, "base": li == 0,
                })
        else:
            vals = []
            nulos = False
            for f in usadas:
                raw = f.get(k)
                if raw is None or texto(raw) == "":
                    nulos = True
                    vals.append(float("nan"))
                else:
                    vals.append(num(raw))
            if nulos:
                avisos.append("La X '%s' tiene valores faltantes; esas filas se excluyen." % k)
            columnas.append({"nombre": k, "valores": vals, "categoria": None})

    keep = []
    for i in range(len(y)):
        if any(not math.isfinite(c["valores"][i]) for c in columnas):
            continue
        keep.append(i)
    y2 = [y[i] for i in keep]

    utiles = []
    for c in columnas:
        col = [c["valores"][i] for i in keep]
        if min(col) == max(col):
            avisos.append("La columna '%s' es constante y se omitio." % c["nombre"])
            continue
        utiles.append(c)
    if not utiles:
        raise ValueError("No quedó ninguna predictora válida.")

    x2 = [[c["valores"][i] for c in utiles] for i in keep]
    nombres = [c["nombre"] for c in utiles]
    columnas_info = [{"nombre": c["nombre"], "categoria": c["categoria"],
                      "nivel": c.get("nivel")} for c in utiles]

    if len(y2) < len(nombres) * 10:
        avisos.append(
            "n = %d para %d predictoras: pocos datos por parametro. "
            "Mira el R2 de validacion cruzada, no el del ajuste." % (len(y2), len(nombres))
        )
    return {
        "y": y2, "X": x2, "nombres_x": nombres, "niveles": niveles,
        "columnas": columnas_info, "n": len(y2),
        "filas_usadas": [usadas[i] for i in keep], "avisos": avisos,
        "unidad": unidad, "excluidas_x": len(y) - len(keep),
    }


# ============================================================================
#  7. Informe por consola
# ============================================================================

LINEA = "=" * 74
SUB = "-" * 74


def identificador_fila(fila, unidad) -> str:
    if unidad == "itemDia":
        return "%s %s" % (fila.get("fecha", ""), fila.get("item", ""))
    return texto(fila.get("fecha", ""))


def descripcion_x(matriz) -> str:
    return ", ".join(matriz["nombres_x"])


def errores_muestra(modelo) -> dict:
    n = modelo["n"]
    resid, ajustado = modelo["resid"], modelo["ajustado"]
    mae = sum(abs(e) for e in resid) / n
    rmse = math.sqrt(modelo["sse"] / n)
    sesgo = sum(resid) / n
    mape = sum(
        abs(resid[i] / ajustado[i]) if ajustado[i] != 0 else 0.0 for i in range(n)
    ) / n * 100.0
    return {"mae": mae, "rmse": rmse, "mape": mape, "sesgo": sesgo}


def ecuacion(modelo) -> str:
    partes = []
    base = 1 if modelo["intercepto"] else 0
    for j, nombre in enumerate(modelo["nombresX"]):
        b = modelo["beta"][j + base]
        signo = "-" if b < 0 else "+"
        partes.append(" %s %s*%s" % (signo, fmt(abs(b), 4), nombre))
    encabezado = "%s =" % modelo.get("nombre_y", "Y")
    if base:
        encabezado += " " + fmt(modelo["beta"][0], 4)
    return encabezado + "".join(partes)


def informe(ds, matriz, modelo, cv, preds, k_vif, diag, args):
    """Arma el texto completo del informe."""
    out = []
    m = ds["meta"]
    unidad = "item-dia" if matriz["unidad"] == "itemDia" else "dia"

    out.append(LINEA)
    out.append(" ANALISIS DE REGRESION LINEAL — control mensual del Centro")
    out.append(LINEA)
    out.append("Archivo     : %s  (%s)" % (ds["nombre"], ds["formato"]))
    if ds.get("hoja") or getattr(args, "hoja", None):
        out.append("Hoja        : %s" % (ds.get("hoja") or args.hoja))
    out.append("Unidad      : %s   (n = %d filas usadas)" % (unidad, modelo["n"]))
    detalle = "%d filas, %d dias" % (m["nFilas"], m["nDias"])
    if m["items"]:
        detalle += ", %d items" % m["items"]
    if m["secciones"]:
        detalle += ", %d secciones" % m["secciones"]
    out.append("Datos       : " + detalle)

    ys = matriz["y"]
    total_y = sum(ys)
    out.append("%-12s: total %s | promedio %s | rango %s ... %s" % (
        modelo.get("nombre_y", "Y")[:12], fmt(total_y, 1),
        fmt(total_y / len(ys), 2), fmt(min(ys), 2), fmt(max(ys), 2)))
    out.append("")

    out.append("Variables   : Y = %s  |  X = %s" % (
        modelo.get("nombre_y", "Y"), descripcion_x(matriz)))
    out.append("Codificacion: %d predictoras tras dummies/omisiones (k = %d)" % (
        len(matriz["nombres_x"]), modelo["k"]))
    if matriz["excluidas_x"]:
        out.append("Filas con X faltante excluidas: %d" % matriz["excluidas_x"])
    for a in matriz["avisos"]:
        out.append("AVISO: " + a)
    out.append("")

    out.append(SUB)
    out.append(" MODELO AJUSTADO  (OLS, %s)" % ("con intercepto" if modelo["intercepto"] else "sin intercepto"))
    out.append(SUB)
    out.append("  " + ecuacion(modelo))
    out.append("")
    out.append("  %-26s %11s %11s %9s %10s %12s %12s %10s" % (
        "Variable", "Beta", "Error est.", "t", "p", "IC95% inf", "IC95% sup", "Beta est."))
    for c in modelo["coef"]:
        out.append("  %-26s %11s %11s %9s %10s %12s %12s %10s" % (
            c["nombre"][:26], fmt(c["beta"], 4), fmt(c["se"], 4), fmt(c["t"], 4),
            fmtp(c["p"]), fmt(c["ci_lo"], 4), fmt(c["ci_hi"], 4),
            "-" if not math.isfinite(c["beta_std"]) else fmt(c["beta_std"], 4)))
    out.append("")

    out.append(bloque_bondad(modelo))
    out.append(bloque_error(modelo, cv))
    out.append(bloque_diagnostico(modelo, k_vif, diag, matriz))
    out.append(bloque_prediccion(matriz, preds))
    return "\n".join(out)


def bloque_bondad(modelo):
    out = [SUB, " BONDAD DE AJUSTE", SUB]
    out.append("  n = %d    k = %d    grados de libertad = %d" % (modelo["n"], modelo["k"], modelo["df"]))
    out.append("  R2            = %s" % fmt(modelo["r2"], 4))
    out.append("  R2 ajustado   = %s" % fmt(modelo["r2adj"], 4))
    out.append("  Error est. (sigma) = %s" % fmt(modelo["sigma"], 4))
    if modelo["k"] > 0 and math.isfinite(modelo["F"]):
        out.append("  F(%d, %d)      = %s    p = %s" % (
            modelo["dfModelo"], modelo["df"], fmt(modelo["F"], 4), fmtp(modelo["Fp"])))
    out.append("")
    return "\n".join(out)


def bloque_error(modelo, cv):
    e = errores_muestra(modelo)
    out = [SUB, " ERROR DE PREDICCION", SUB]
    out.append("  En la muestra     : MAE = %s   RMSE = %s   MAPE = %s%%   sesgo = %s" % (
        fmt(e["mae"], 3), fmt(e["rmse"], 3), fmt(e["mape"], 2), fmt(e["sesgo"], 3)))
    if cv and cv.get("ok"):
        out.append("  Validacion cruzada (%d-fold, fuera de muestra):" % cv["k"])
        out.append("      R2 = %s   RMSE = %s   MAE = %s   (n = %d)" % (
            fmt(cv["r2"], 4), fmt(cv["rmse"], 3), fmt(cv["mae"], 3), cv["n"]))
        if cv["r2"] < modelo["r2"] - 0.15:
            out.append("      AVISO: el R2 cae fuerte fuera de muestra: el ajuste esta sobreestimado.")
    else:
        out.append("  Validacion cruzada: no se pudo calcular (%s)" % (
            cv.get("error") if cv else "deshabilitada"))
    out.append("")
    return "\n".join(out)


def bloque_diagnostico(modelo, k_vif, diag, matriz):
    dw = durbin_watson(modelo["resid"])
    asim, curt = asimetria_curtosis(modelo["resid"])
    out = [SUB, " DIAGNOSTICO", SUB]
    out.append("  Durbin-Watson = %s   (2 = sin autocorrelacion)" % fmt(dw, 4))
    out.append("  Residuos: asimetria = %s   curtosis (exceso) = %s" % (fmt(asim, 4), fmt(curt, 4)))
    tolerancias = []
    for nombre, v in zip(modelo["nombresX"], k_vif):
        alerta = "  <-- VIF > 10: colinealidad alta" if v["vif"] > 10 else ""
        tolerancias.append("      %-26s VIF = %8s   tol = %8s%s" % (
            nombre[:26], fmt(v["vif"], 3), fmt(v["tol"], 3), alerta))
    if tolerancias:
        out.append("  Multicolinealidad:")
        out.extend(tolerancias)
    n = modelo["n"]
    out.append("  Outliers por distancia de Cook (> 4/n = %s):" % fmt(4.0 / n, 4))
    if diag["outliers"]:
        filas = matriz["filas_usadas"]
        orden = sorted(diag["outliers"], key=lambda i: -diag["cook"][i])
        for i in orden[:10]:
            out.append("      %-28s Cook = %s   res. estudentizado = %s" % (
                identificador_fila(filas[i], matriz["unidad"])[:28],
                fmt(diag["cook"][i], 4), fmt(diag["estudent"][i], 2)))
        if len(orden) > 10:
            out.append("      ... y %d mas" % (len(orden) - 10))
    else:
        out.append("      ninguno")
    out.append("")
    return "\n".join(out)


def bloque_prediccion(matriz, preds):
    out = [SUB, " PREDICCION", SUB]
    if not preds:
        out.append("  (usá --predecir variable=valor[,variable=valor] para calcular escenarios)")
        out.append("")
        return "\n".join(out)
    rangos = rangos_observados(matriz)
    for caso in preds:
        valores = caso["valores"]
        out.append("  Escenario: %s" % ", ".join("%s = %s" % (k, fmt(v, 4)) for k, v in valores.items()))
        out.append("      Y estimada             = %s" % fmt(caso["yhat"], 4))
        out.append("      IC 95%% para la media   = %s ... %s" % (
            fmt(caso["ic_media"][0], 4), fmt(caso["ic_media"][1], 4)))
        out.append("      IC 95%% obs. nueva      = %s ... %s" % (
            fmt(caso["ic_nueva"][0], 4), fmt(caso["ic_nueva"][1], 4)))
        for aviso in caso["avisos"]:
            out.append("      AVISO: " + aviso)
        fuera = []
        for k, v in valores.items():
            r = rangos.get(k)
            if r and (v < r[0] or v > r[1]):
                fuera.append("%s = %s fuera del rango observado %s ... %s" % (
                    k, fmt(v, 4), fmt(r[0], 4), fmt(r[1], 4)))
        for f in fuera:
            out.append("      EXTRAPOLACION: " + f)
        out.append("")
    return "\n".join(out)


def rangos_observados(matriz):
    """min/max de cada predictora numerica en las filas usadas (para avisar extrapolacion)."""
    rangos = {}
    filas = matriz["filas_usadas"]
    for col in matriz["columnas"]:
        if col["categoria"]:
            continue
        vals = [num(f.get(col["nombre"])) for f in filas]
        if vals:
            rangos[col["nombre"]] = (min(vals), max(vals))
    return rangos


# ============================================================================
#  8. Interfaz de linea de comandos
# ============================================================================

def vector_x(matriz, valores):
    """Valores dados por el usuario -> fila de X en el orden codificado."""
    vec = []
    faltan = []
    for col in matriz["columnas"]:
        if col["categoria"]:
            dado = valores.get(col["categoria"])
            vec.append(1.0 if dado is not None and texto(dado) == col["nivel"] else 0.0)
        else:
            if col["nombre"] not in valores:
                faltan.append(col["nombre"])
                vec.append(0.0)
            else:
                vec.append(num(valores[col["nombre"]]))
    if faltan:
        raise ValueError("Faltan valores para: %s" % ", ".join(faltan))
    return vec


def parsear_predecir(argumentos):
    """--predecir "a=1,b=2" (repetible) -> lista de escenarios."""
    escenarios = []
    for arg in argumentos or []:
        valores = {}
        for trozo in arg.split(","):
            trozo = trozo.strip()
            if not trozo:
                continue
            if "=" not in trozo:
                raise ValueError("Escenario invalido '%s': se espera variable=valor" % trozo)
            k, v = trozo.split("=", 1)
            valores[clave(k)] = v.strip()
        if valores:
            escenarios.append(valores)
    return escenarios


def variables_disponibles(ds, unidad):
    return ds["variables"]["itemDia" if unidad == "itemDia" else "dia"]


def lista_de_argumento(arg):
    """"a,b c" -> ["a", "b", "c"] (acepta comas y espacios)."""
    if not arg:
        return []
    return [s for s in re.split(r"[,\s]+", arg.strip()) if s]


def resolver_variables(ds, unidad, y_pedida, x_pedidas, cat_pedidas):
    disponibles = variables_disponibles(ds, unidad)
    mapa = {v["clave"]: dict(v) for v in disponibles}
    claves_ok = list(mapa.keys())

    if unidad == "itemDia":
        por_defecto_y = "monto"
        por_defecto_x = ["cantidad", "seccion"]
    else:
        por_defecto_y = Y_POR_DEFECTO
        por_defecto_x = X_POR_DEFECTO

    y_clave = clave(y_pedida) if y_pedida else por_defecto_y
    x_pedidas = lista_de_argumento(x_pedidas) if isinstance(x_pedidas, str) else (x_pedidas or [])
    cat_pedidas = lista_de_argumento(cat_pedidas) if isinstance(cat_pedidas, str) else (cat_pedidas or [])
    if y_clave not in mapa:
        raise ValueError("La variable Y '%s' no existe en la unidad '%s'. Disponibles: %s" % (
            y_clave, unidad, ", ".join(claves_ok)))

    x_claves = [clave(c) for c in x_pedidas] if x_pedidas else list(por_defecto_x)
    x_claves = [c for c in x_claves if c]
    faltan = [c for c in x_claves if c not in mapa]
    if faltan:
        raise ValueError("No existe(n) la(s) variable(s) X: %s\nDisponibles: %s" % (
            ", ".join(faltan), ", ".join(claves_ok)))
    if y_clave in x_claves:
        raise ValueError("La variable Y '%s' no puede usarse tambien como X." % y_clave)

    categoricas = {clave(c) for c in (cat_pedidas or []) if clave(c)}
    x_vars = []
    for c in x_claves:
        v = mapa[c]
        if c in categoricas:
            v["tipo"] = "categorica"
        x_vars.append(v)
    return y_clave, x_vars, categoricas


def cargar_dataset(args):
    """Lee el archivo pedido (CSV transaccional, CSV matriz o Excel) y arma el dataset."""
    # --archivo es un alias del argumento posicional; cualquiera de los dos sirve.
    path = getattr(args, "archivo_opt", None) or args.archivo or ARCHIVO_POR_DEFECTO
    if not os.path.exists(path):
        if path != ARCHIVO_POR_DEFECTO:
            raise ValueError("No existe el archivo: %s" % path)
        raise ValueError(
            "No encontre '%s' en el directorio actual. Indicalo con --archivo." % path)

    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        if ext == ".xls":
            raise ValueError("El formato .xls viejo no se puede leer sin dependencias. "
                             "Guardalo como .xlsx o .xlsm.")
        ds = None
        for nombre, filas in iter_hojas_xlsx(path, args.hoja):
            ds = construir_dataset(filas, nombre, args.concepto)
            if ds:
                ds["hoja"] = nombre
                break
        if not ds:
            raise ValueError(
                "No encontre una tabla reconocible en la(s) hoja(s) leida(s). "
                "Pasá --hoja con el nombre de la hoja (por ej. ventas, caja, gastos, insumos).")
        return ds

    filas = leer_csv(path)
    ds = construir_dataset(filas, os.path.basename(path), args.concepto)
    if not ds:
        raise ValueError(
            "El archivo no tiene el formato esperado: CSV transaccional "
            "(fecha, item, cantidad, precio, monto) o matriz diaria (concepto por fila, "
            "un dia por columna).")
    return ds


def analizar(args):
    """Corre todo el analisis y devuelve (texto, resultado, dataset)."""
    ds = cargar_dataset(args)
    unidad = "itemDia" if (args.unidad or "dia").lower() in ("item-dia", "itemdia", "item_dia") else "dia"
    y_clave, x_vars, _cats = resolver_variables(ds, unidad, args.y, args.x, args.cat)

    matriz = construir_matriz(ds, unidad, y_clave, x_vars, intercepto=not args.sin_intercepto)
    modelo = ols(matriz["y"], matriz["X"],
                 intercepto=not args.sin_intercepto, nombres_x=matriz["nombres_x"])
    modelo["nombre_y"] = y_clave
    diag = diagnostico(modelo)
    k_vif = vif(matriz["X"])
    cv = None
    if args.kfold and args.kfold > 1 and modelo["n"] > args.kfold:
        cv = kfold(matriz["y"], matriz["X"], k=args.kfold, intercepto=not args.sin_intercepto)

    preds = []
    categoricas_cod = {c["categoria"] for c in matriz["columnas"] if c["categoria"]}
    numericas_cod = {c["nombre"] for c in matriz["columnas"] if not c["categoria"]}
    for bruto in parsear_predecir(args.predecir):
        desconocidas = [k for k in bruto if k not in categoricas_cod and k not in numericas_cod]
        if desconocidas:
            raise ValueError("El escenario usa variable(s) que no estan en el modelo: %s\n"
                             "Predictoras del modelo: %s" % (
                                 ", ".join(desconocidas),
                                 ", ".join(sorted(categoricas_cod | numericas_cod))))
        valores = {k: (texto(v) if k in categoricas_cod else num(v)) for k, v in bruto.items()}
        x0 = vector_x(matriz, valores)
        r0 = predecir(modelo, x0)
        r0["valores"] = valores
        r0["avisos"] = []
        preds.append(r0)

    resultado = {
        "version": VERSION,
        "archivo": ds["nombre"],
        "formato": ds["formato"],
        "unidad": unidad,
        "y": y_clave,
        "x": [v["clave"] for v in x_vars],
        "meta": ds["meta"],
        "avisos": matriz["avisos"],
        "n": modelo["n"],
        "k": modelo["k"],
        "df": modelo["df"],
        "intercepto": modelo["intercepto"],
        "coeficientes": modelo["coef"],
        "r2": modelo["r2"],
        "r2_ajustado": modelo["r2adj"],
        "sigma": modelo["sigma"],
        "F": modelo["F"],
        "F_p": modelo["Fp"],
        "error_muestra": errores_muestra(modelo),
        "validacion_cruzada": {kk: vv for kk, vv in cv.items() if kk != "pred"} if cv else None,
        "vif": [{"variable": n_, **v_} for n_, v_ in zip(modelo["nombresX"], k_vif)],
        "durbin_watson": durbin_watson(modelo["resid"]),
        "residuos": dict(zip(("asimetria", "curtosis"), asimetria_curtosis(modelo["resid"]))),
        "outliers_cook": [
            {"fila": identificador_fila(matriz["filas_usadas"][i], unidad),
             "cook": diag["cook"][i], "residuo_estudentizado": diag["estudent"][i]}
            for i in diag["outliers"]
        ],
        "predicciones": [
            {"valores": p["valores"], "yhat": p["yhat"], "ic_media": p["ic_media"],
             "ic_nueva": p["ic_nueva"], "se_media": p["se_media"], "se_nueva": p["se_nueva"]}
            for p in preds
        ],
    }
    texto_informe = informe(ds, matriz, modelo, cv, preds, k_vif, diag, args)
    return texto_informe, resultado, ds, matriz


def listar_variables(ds):
    out = [SUB, " VARIABLES DISPONIBLES", SUB]
    for unidad, titulo in (("dia", "UNIDAD DIA (una fila por dia)"),
                           ("itemDia", "UNIDAD ITEM-DIA (una linea de venta)")):
        out.append("  %s" % titulo)
        for v in ds["variables"][unidad]:
            marca = " [Y sugerida]" if v["rol"] == "y" else ""
            if v.get("externa"):
                marca += " [externa: es consecuencia de las ventas]"
            desc = (" — " + v["descripcion"]) if v.get("descripcion") else ""
            out.append("      %-24s %-11s%s%s" % (v["clave"], v["tipo"], desc, marca))
        out.append("")
    return "\n".join(out)


def escribir_escenarios_csv(path, preds):
    with open(path, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["escenario", "valores", "y_estimada", "ic95_media_inf", "ic95_media_sup",
                    "ic95_nueva_inf", "ic95_nueva_sup"])
        for i, p in enumerate(preds, 1):
            w.writerow([
                i,
                "; ".join("%s=%s" % (k, fmt(v, 4)) for k, v in p["valores"].items()),
                fmt(p["yhat"], 4), fmt(p["ic_media"][0], 4), fmt(p["ic_media"][1], 4),
                fmt(p["ic_nueva"][0], 4), fmt(p["ic_nueva"][1], 4),
            ])


def construir_parser():
    p = argparse.ArgumentParser(
        prog="analisis_regresion.py",
        description="Regresion lineal simple o multiple sobre el control mensual del Centro "
                    "(solo biblioteca estandar de Python).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Ejemplos:\n"
               "  python analisis_regresion.py\n"
               "  python analisis_regresion.py --listar\n"
               "  python analisis_regresion.py --x cantidad_grandes --predecir cantidad_grandes=5\n"
               "  python analisis_regresion.py --archivo \"260901 CONTROL MENSUAL CENTRO.xlsm\" --hoja caja\n"
               "  python analisis_regresion.py --unidad item-dia --y monto --x cantidad,seccion --cat seccion\n",
    )
    p.add_argument("archivo", nargs="?", default=None,
                   help="CSV transaccional, CSV matriz diaria o libro .xlsx/.xlsm "
                        "(por defecto %s)" % ARCHIVO_POR_DEFECTO)
    p.add_argument("--archivo", dest="archivo_opt", default=None,
                   help="Alias del argumento posicional (util cuando el nombre tiene espacios).")
    p.add_argument("--hoja", default=None, help="Hoja del libro de Excel (por ej. ventas, caja, gastos).")
    p.add_argument("--concepto", default=None, metavar="COLUMNA",
                   help="Columna de concepto en el formato matriz: numero (1 = primera) o nombre "
                        "del encabezado (por ej. ITEM en la hoja ventas).")
    p.add_argument("--unidad", default="dia", help="Unidad de analisis: dia (defecto) o item-dia.")
    p.add_argument("--y", default=None, help="Variable dependiente (por defecto %s)." % Y_POR_DEFECTO)
    p.add_argument("--x", default=None,
                   help="Predictoras separadas por coma (por defecto %s)." % ",".join(X_POR_DEFECTO))
    p.add_argument("--cat", default=None,
                   help="Predictoras que deben tratarse como categoricas (dummies), separadas por coma.")
    p.add_argument("--sin-intercepto", action="store_true", help="Ajusta el modelo sin intercepto.")
    p.add_argument("--kfold", type=int, default=5, help="Folds de la validacion cruzada (0 = deshabilitada).")
    p.add_argument("--predecir", action="append", default=None, metavar="VAR=VALOR[,VAR=VALOR]",
                   help="Escenario de prediccion; se puede repetir.")
    p.add_argument("--escenarios-csv", default=None, metavar="ARCHIVO",
                   help="Guarda la tabla de escenarios en un CSV (separador ;).")
    p.add_argument("--json", dest="json_salida", default=None, metavar="ARCHIVO",
                   help="Guarda todos los resultados en un JSON.")
    p.add_argument("--listar", action="store_true", help="Lista las variables disponibles y sale.")
    p.add_argument("--version", action="version", version="analisis_regresion.py " + VERSION)
    return p


def main(argv=None) -> int:
    try:  # consola de Windows en cp1252: evitar UnicodeEncodeError con acentos
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

    args = construir_parser().parse_args(argv)

    try:
        if args.listar:
            ds = cargar_dataset(args)
            m = ds["meta"]
            resumen = "%s — %d filas, %d dias, %d items" % (
                ds["formato"], m["nFilas"], m["nDias"], m["items"])
            if any(v["clave"] == "monto_total" for v in ds["variables"]["dia"]):
                resumen += ", venta total $%s" % fmt(m["totalMonto"], 1)
            print("%s\n%s" % (ds["nombre"], resumen))
            print(listar_variables(ds))
            return 0

        texto_informe, resultado, _ds, _matriz = analizar(args)
        print(texto_informe)

        if args.escenarios_csv and resultado["predicciones"]:
            casos = []
            for p in resultado["predicciones"]:
                casos.append({"valores": p["valores"], "yhat": p["yhat"],
                              "ic_media": p["ic_media"], "ic_nueva": p["ic_nueva"]})
            escribir_escenarios_csv(args.escenarios_csv, casos)
            print("Escenarios guardados en %s" % args.escenarios_csv)
        if args.json_salida:
            with open(args.json_salida, "w", encoding="utf-8") as fh:
                json.dump(resultado, fh, ensure_ascii=False, indent=2, default=str)
            print("Resultados guardados en %s" % args.json_salida)
        return 0
    except (ValueError, OSError) as e:
        print("ERROR: %s" % e, file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
