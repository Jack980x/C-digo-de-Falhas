import os
import sys
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.abspath('.'))

from app import executar_lote_multiplos_pdfs, executar_extracao_padrao

class TestOp4ArquivosDuplos(unittest.TestCase):

    def test_op4_lote_gera_dois_arquivos_distintos(self):
        caminhos_pdf = ["doc1.pdf", "doc2.pdf"]
        leitor = MagicMock()
        leitor.obter_total_paginas.return_value = 1
        leitor.ler_documento.return_value = [{
            "texto": "P0100 Falha MAF Circuito aberto",
            "metodo": "vetorial",
            "codigos_ground_truth": ["P0100"]
        }]
        
        tradutor = MagicMock()
        tradutor.estruturar_buscar_codigo.return_value = [{
            "Código": "P0100",
            "Localização da Avaria": "Sensor MAF",
            "Causa": "Circuito aberto",
            "Descrição": "Circuito aberto no Sensor MAF"
        }]

        agente_organizador = MagicMock()
        gerador_excel = MagicMock()
        gerador_excel.exportar_para_excel.return_value = os.path.abspath("saida/Unificado_2_PDFs_traduzido.xlsx")
        gerador_excel.exportar_buscar_codigo.return_value = os.path.abspath("saida/Unificado_2_PDFs_buscar_codigo.xlsx")

        config = {}
        padrao_op4 = {
            "colunas": ["Código", "Causa", "Localização da Avaria", "Descrição"],
            "larguras": {"Código": 16.57, "Causa": 52.43, "Localização da Avaria": 52.57},
            "tema_visual": "simplo_3_colunas_desc"
        }

        with patch('app.Prompt.ask', side_effect=["3", "1", "1"]), \
             patch('app.Confirm.ask', return_value=False), \
             patch('app.exibir_diagnostico_leitura_automatica', return_value=False):
            
            executar_lote_multiplos_pdfs(
                caminhos_pdf=caminhos_pdf,
                operacao="4",
                leitor=leitor,
                tradutor=tradutor,
                agente_organizador=agente_organizador,
                gerador_excel=gerador_excel,
                config=config,
                padrao_predefinido=padrao_op4
            )

            # Verifica se ambos os geradores foram chamados
            gerador_excel.exportar_para_excel.assert_called_once()
            gerador_excel.exportar_buscar_codigo.assert_called_once()

            # Pega os caminhos de saída passados para cada função
            caminho_padrao = gerador_excel.exportar_para_excel.call_args[0][2]
            caminho_bc = gerador_excel.exportar_buscar_codigo.call_args[0][1]

            # Garante que os caminhos são DIFERENTES e terminam corretamente
            self.assertTrue(caminho_padrao.endswith("_traduzido.xlsx"), f"Esperado _traduzido.xlsx, obtido: {caminho_padrao}")
            self.assertTrue(caminho_bc.endswith("_buscar_codigo.xlsx"), f"Esperado _buscar_codigo.xlsx, obtido: {caminho_bc}")
            self.assertNotEqual(caminho_padrao, caminho_bc, "Os dois arquivos NÃO podem ter o mesmo nome!")

    def test_op4_individual_gera_dois_arquivos_distintos(self):
        caminho_pdf = "doc_unico.pdf"
        leitor = MagicMock()
        leitor.obter_total_paginas.return_value = 1
        leitor.ler_documento.return_value = [{
            "texto": "P0100 Falha MAF Circuito aberto",
            "metodo": "vetorial",
            "codigos_ground_truth": ["P0100"]
        }]

        tradutor = MagicMock()
        tradutor.estruturar_buscar_codigo.return_value = [{
            "Código": "P0100",
            "Localização da Avaria": "Sensor MAF",
            "Causa": "Circuito aberto",
            "Descrição": "Circuito aberto no Sensor MAF"
        }]

        agente_organizador = MagicMock()
        gerador_excel = MagicMock()
        gerador_excel.exportar_para_excel.return_value = os.path.abspath("saida/doc_unico_traduzido.xlsx")
        gerador_excel.exportar_buscar_codigo.return_value = os.path.abspath("saida/doc_unico_buscar_codigo.xlsx")

        config = {}
        padrao_op4 = {
            "colunas": ["Código", "Causa", "Localização da Avaria", "Descrição"],
            "larguras": {"Código": 16.57, "Causa": 52.43, "Localização da Avaria": 52.57},
            "tema_visual": "simplo_3_colunas_desc"
        }

        with patch('app.selecionar_modulos', return_value=(["Módulo Único"], "mesmo_arquivo", {})), \
             patch('app.selecionar_paginas', return_value=[1]), \
             patch('app.exibir_diagnostico_leitura_automatica', return_value=False), \
             patch('app.Confirm.ask', return_value=False):

            executar_extracao_padrao(
                caminho_pdf=caminho_pdf,
                leitor=leitor,
                tradutor=tradutor,
                agente_organizador=agente_organizador,
                gerador_excel=gerador_excel,
                total_paginas=1,
                config=config,
                exportar_duplo=True,
                padrao_predefinido=padrao_op4
            )

            gerador_excel.exportar_para_excel.assert_called_once()
            gerador_excel.exportar_buscar_codigo.assert_called_once()

            caminho_padrao = gerador_excel.exportar_para_excel.call_args[0][2]
            caminho_bc = gerador_excel.exportar_buscar_codigo.call_args[0][1]

            self.assertTrue(caminho_padrao.endswith("_traduzido.xlsx"))
            self.assertTrue(caminho_bc.endswith("_buscar_codigo.xlsx"))
            self.assertNotEqual(caminho_padrao, caminho_bc)

if __name__ == '__main__':
    unittest.main()

