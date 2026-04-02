#!/usr/bin/env python3
"""
test_audit_demo.py
==================
Demonstra o sistema de auditoria de linhas rejeitadas.

Simula a leitura de um arquivo Demanda.csv com dados problemáticos:
- Datas inválidas
- Quantidades não numéricas
- Linhas em branco

Resultado: relatório de auditoria em audit/rejected_demanda.csv
"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import mrp

mrp.DIR_DADOS = "data_sim"
mrp.DIR_SAIDA = os.path.join("data_sim", "output_with_errors")
os.makedirs(mrp.DIR_SAIDA, exist_ok=True)

print("\n" + "="*72)
print("  TESTE DE AUDITORIA — Leitura de arquivo com erros")
print("="*72)

dem_com_erros = mrp.transformar_demanda_dtm(os.path.join("data_sim", "demanda_com_erros.csv"))

print("\n  Demanda processada com sucesso (linhas válidas):")
print(dem_com_erros[["material", "mes", "quantidade"]].to_string(index=False))

# Salva auditoria
print("\n" + "-"*72)
resumo = mrp.audit.save_all()
print("\n  Arquivos de auditoria salvos:")
for source, count in resumo.items():
    print(f"    ✓ {source}: {count} linha(s) rejeitada(s)")

# Mostra conteúdo dos arquivos de rejeição
if resumo:
    import pandas as pd
    audit_dir = os.path.join(mrp.DIR_SAIDA, "audit")
    for source in resumo:
        path = os.path.join(audit_dir, f"rejected_{source}.csv")
        df = pd.read_csv(path)
        print(f"\n  {source.upper()} — Linhas rejeitadas ({len(df)}):")
        print(df.to_string(index=False))
        print(f"    └─ Motivos: {', '.join(df['reason'].unique())}")

print("\n" + "="*72 + "\n")
