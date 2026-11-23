import sys
import os
sys.path.insert(0, os.path.abspath("."))
import openpyxl
from core.agente_organizador import AgenteOrganizador
from core.gerador_excel import GeradorExcel

dados_teste = [
    {
        'Código': 'B181711',
        '_cod_base': 'B1817',
        '_fmi': '11',
        'Localização da Avaria': 'Circuito de acionamento do farol baixo',
        'Causa': 'Curto-circuito ao aterramento (GND)',
        'Descrição': '-'
    },
    {
        'Código': 'B100115',
        '_cod_base': 'B1001',
        '_fmi': '15',
        'Localização da Avaria': 'Airbag frontal do passageiro',
        'Causa': 'Curto-circuito ao positivo (Vbatt)',
        'Descrição': 'Possíveis causas incluem falhas elétricas ou desgaste mecânico. Antes de substituir peças, realize testes detalhados e utilize diagramas elétricos e manuais técnicos SIMPLO para um diagnóstico preciso e confiável.'
    },
    {
        'Código': 'P0100',
        '_cod_base': 'P0100',
        '_fmi': '-',
        'Localização da Avaria': 'Sensor de fluxo de massa de ar (MAF)',
        'Causa': 'Circuito aberto',
        'Descrição': ''
    },
    {
        'Código': 'U0100',
        '_cod_base': 'U0100',
        '_fmi': '-',
        'Localização da Avaria': 'Módulo de controle do motor (ECM)',
        'Causa': 'Perda de comunicação na rede CAN',
        'Descrição': '-'
    },
    {
        'Código': 'P0401',
        '_cod_base': 'P0401',
        '_fmi': '-',
        'Localização da Avaria': 'Válvula EGR',
        'Causa': 'Fluxo insuficiente detectado / Obstrução mecânica',
        'Descrição': '-'
    }
]

print("=== TESTE DE SÍNTESE DIRETA ===")
for d in dados_teste:
    sint = AgenteOrganizador.sintetizar_descricao_modo_completo(d['Causa'], d['Localização da Avaria'], d['Descrição'])
    print(f"[{d['Código']}] {d['Causa']} + {d['Localização da Avaria']}")
    print(f"-> {sint}\n")

print("=== TESTE DE GERAÇÃO EXCEL (SIMPLO 3 Colunas + Descrição) ===")
gerador = GeradorExcel()
caminho_teste = "testes_e_exemplos/teste_modo_completo_3col.xlsx"
gerador.exportar_para_excel(
    dados=dados_teste,
    colunas_ordenadas=["Código", "Causa", "Localização da Avaria", "Descrição"],
    caminho_saida=caminho_teste,
    nome_modulo="MÓDULO DE TESTE MODO COMPLETO",
    tema_visual="simplo_3_colunas_desc"
)

wb = openpyxl.load_workbook(caminho_teste)
ws = wb.active
print(f"Planilha gerada com sucesso! Linhas totais: {ws.max_row}")
for r in range(4, ws.max_row + 1):
    c1 = ws.cell(r, 1).value
    c2 = ws.cell(r, 2).value
    c3 = ws.cell(r, 3).value
    print(f"Linha {r:2d} | Col 1: {str(c1):12s} | Col 2: {str(c2)[:40]:40s} | Col 3: {str(c3)[:40]:40s}")
