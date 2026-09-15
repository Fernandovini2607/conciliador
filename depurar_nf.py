"""Depurador: confirma que a Fase 4 esta no codigo carregado e lista
as parcelas do Dominio pra uma NF especifica.

Uso:
  python depurar_nf.py <numero_nf> [codi_emp]

Exemplos:
  python depurar_nf.py 171255              (usa a empresa do config)
  python depurar_nf.py 171255 82           (busca so em codi_emp=82)
  python depurar_nf.py 171255 81,82,133    (busca em varios codi_emp)
"""

from __future__ import annotations

import sys

import config
import main
import parser_dominio


def _fase4_no_codigo() -> bool:
    doc = (main.App._filtrar_conciliados_por_dominio.__doc__ or "").upper()
    return "FASE 4" in doc


def main_cli(nf_alvo: str, codis_arg: str | None) -> None:
    print(f"NF alvo: {nf_alvo!r}\n")

    if _fase4_no_codigo():
        print("[OK] Fase 4 esta presente no main.py carregado.")
    else:
        print(
            "[ERRO] main.py carregado NAO tem Fase 4.\n"
            "  Feche o app, rode 'git pull' na pasta do projeto e\n"
            "  abra o Conciliador de novo pelo iniciar.bat."
        )
        return

    cfg = config.carregar()
    fonte = cfg.get("dominio_fonte_pagamentos", {})
    if not fonte.get("mapeamento"):
        print("[ERRO] Fonte de pagamentos nao configurada.")
        return

    if codis_arg:
        codis = [int(x) for x in codis_arg.split(",") if x.strip()]
    else:
        emp = cfg.get("dominio_empresa") or {}
        codi = emp.get("codi_emp")
        cnpj = emp.get("cnpj", "")
        print(f"Empresa do config: [{codi}] {emp.get('razao','')} CNPJ={cnpj}")
        if not codi:
            print(
                "[ERRO] Nenhuma empresa selecionada no config. Passe o "
                "codi_emp na linha de comando: python depurar_nf.py "
                f"{nf_alvo} 82"
            )
            return
        codis = [codi]
        # Expande pro grupo (matriz + filiais)
        if cnpj:
            with parser_dominio.connect_dominio() as conn:
                filiais = parser_dominio.listar_filiais(conn, cnpj)
            if filiais:
                codis = [f["codi_emp"] for f in filiais]

    print(f"codi_emp a consultar: {codis}\n")

    achou = 0
    with parser_dominio.connect_dominio() as conn:
        for codi in codis:
            try:
                txs = parser_dominio.extrair_pagamentos(
                    conn, fonte, codi_emp=codi,
                )
            except Exception as e:
                print(f"  codi={codi}: ERRO {e}")
                continue
            for t in txs:
                nf = str(t.extras.get("numero_nf", "")).strip()
                if nf == nf_alvo:
                    achou += 1
                    forn = t.extras.get("fornecedor", "")
                    print(
                        f"  codi={codi:>4}  data={t.data}  "
                        f"valor=R$ {t.valor}  "
                        f"status={t.extras.get('status','')}  "
                        f"pago={t.extras.get('valor_pago','')}"
                    )
                    print(
                        f"    CNPJ={t.extras.get('cnpj','')!r}  "
                        f"forn={forn!r}"
                    )
                    print(
                        f"    forn normalizado="
                        f"{main.App._normaliza_nome_fornecedor(forn)!r}"
                    )
    if not achou:
        print("  (nenhuma parcela com essa NF encontrada nas empresas listadas)")
    else:
        print(f"\nTotal: {achou} parcela(s) encontrada(s) no Dominio.")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    codis = sys.argv[2] if len(sys.argv) > 2 else None
    main_cli(sys.argv[1], codis)
