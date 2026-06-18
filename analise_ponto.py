import csv
import json
from datetime import datetime, timedelta
from urllib.request import Request, urlopen
from urllib.error import URLError

OLLAMA_URL = "http://localhost:11434/api/generate"
MODEL = "llama3.1:8b"
ARQUIVO_CSV = "folha_ponto.csv"


def ler_csv(caminho):
    with open(caminho, "r", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f, delimiter=";"))


def parse_hora(valor):
    if not valor or not valor.strip():
        return None
    return datetime.strptime(valor.strip(), "%H:%M")


def calcular_horas_trabalhadas(entrada, saida_almoco, retorno_almoco, saida):
    if not all([entrada, saida]):
        return 0.0
    e = parse_hora(entrada)
    sm = parse_hora(saida_almoco)
    rm = parse_hora(retorno_almoco)
    s = parse_hora(saida)

    total = (s - e).total_seconds() / 3600
    if sm and rm:
        almoco = (rm - sm).total_seconds() / 3600
        total -= almoco
    return round(total, 2)


def processar_registros(rows):
    colaboradores = {}
    for row in rows:
        cid = row["colaborador_id"]
        if cid not in colaboradores:
            colaboradores[cid] = {
                "nome": row["nome"],
                "cargo": row["cargo"],
                "carga_diaria": float(row["carga_horaria_diaria"]),
                "dias": [],
            }
        data = row["data"]
        problemas = []
        for campo, rotulo in [
            ("entrada", "entrada"),
            ("saida_almoco", "saida almoco"),
            ("retorno_almoco", "retorno almoco"),
            ("saida", "saida"),
        ]:
            if not row.get(campo, "").strip():
                problemas.append(f"{rotulo} nao registrada")

        horas_trab = calcular_horas_trabalhadas(
            row["entrada"], row["saida_almoco"], row["retorno_almoco"], row["saida"]
        )
        carga = colaboradores[cid]["carga_diaria"]
        saldo = round(horas_trab - carga, 2)

        registro = {
            "data": data,
            "entrada": row["entrada"],
            "saida_almoco": row["saida_almoco"],
            "retorno_almoco": row["retorno_almoco"],
            "saida": row["saida"],
            "horas_trabalhadas": horas_trab,
            "saldo": saldo,
            "problemas": problemas,
        }
        colaboradores[cid]["dias"].append(registro)

    return colaboradores


def gerar_resumo(colaboradores):
    linhas = []
    for cid, info in colaboradores.items():
        linhas.append(
            f"Colaborador: {info['nome']} (ID {cid}) - {info['cargo']} - Carga diaria: {info['carga_diaria']}h"
        )
        saldo_total = 0.0
        dias_faltantes = []
        for dia in info["dias"]:
            saldo_total += dia["saldo"]
            if dia["problemas"]:
                dias_faltantes.append(f"  - {dia['data']}: {', '.join(dia['problemas'])}")
        for dia in info["dias"]:
            if not dia["problemas"]:
                continue
        if dias_faltantes:
            linhas.append("  Dias com registro incompleto:")
            linhas.extend(dias_faltantes)
        else:
            linhas.append("  Todos os dias com registro completo.")
        if saldo_total > 0:
            linhas.append(f"  Saldo: +{saldo_total}h (horas extras)")
        elif saldo_total < 0:
            linhas.append(f"  Saldo: {saldo_total}h (horas devendo)")
        else:
            linhas.append("  Saldo: 0h (dentro do esperado)")
        linhas.append("")
    return "\n".join(linhas)


def consultar_ollama(prompt):
    payload = json.dumps({
        "model": MODEL,
        "prompt": prompt,
        "stream": False,
    }).encode("utf-8")
    req = Request(OLLAMA_URL, data=payload, headers={"Content-Type": "application/json"})
    try:
        with urlopen(req, timeout=300) as resp:
            return json.loads(resp.read())["response"]
    except URLError as e:
        return f"[Erro ao conectar no Ollama: {e.reason}]"
    except Exception as e:
        return f"[Erro: {e}]"


def main():
    print("=== ANALISE DE FOLHA DE PONTO ===\n")

    rows = ler_csv(ARQUIVO_CSV)
    colaboradores = processar_registros(rows)
    resumo = gerar_resumo(colaboradores)

    print("--- Dados processados ---")
    print(resumo)

    print("\n--- Analise com IA (Ollama) ---")
    prompt = f"""Voce e um analista de RH. Analise os seguintes dados de folha de ponto e identifique:

1. Quais colaboradores tem dias com registros de ponto faltando (nao bateram o ponto em determinados horarios)
2. Quem esta com saldo de horas negativo (devendo horas) e quem esta com saldo positivo (horas extras)
3. Recomendacoes para cada caso

Dados:

{resumo}

Responda de forma clara e concisa em portugues."""
    resposta = consultar_ollama(prompt)
    print(resposta)

    print("\n--- Detalhamento por colaborador ---")
    for cid, info in colaboradores.items():
        print(f"\n{info['nome']} | Saldo total: {sum(d['saldo'] for d in info['dias']):+.2f}h")
        for dia in info["dias"]:
            prob = f" [PROBLEMA: {', '.join(dia['problemas'])}]" if dia["problemas"] else ""
            print(
                f"  {dia['data']}: {dia['entrada'] or '--'} | {dia['saida_almoco'] or '--'} | {dia['retorno_almoco'] or '--'} | {dia['saida'] or '--'} "
                f"=> {dia['horas_trabalhadas']}h (saldo: {dia['saldo']:+.2f}h){prob}"
            )


if __name__ == "__main__":
    main()
