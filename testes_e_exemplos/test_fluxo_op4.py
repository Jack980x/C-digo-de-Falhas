import os
import sys
import unittest
from unittest.mock import patch, MagicMock

sys.path.insert(0, os.path.abspath('.'))

import app

class TestFluxoOperacao4(unittest.TestCase):
    @patch('app.Prompt.ask')
    @patch('app.selecionar_arquivos_pdf')
    @patch('app.selecionar_padrao_planilha')
    @patch('app.executar_extracao_padrao')
    @patch('app.Confirm.ask', return_value=False)
    def test_op4_chama_pergunta_padrao_imediatamente_arquivo_unico(
        self, mock_confirm, mock_exec_extracao, mock_selecionar_padrao, mock_selecionar_arquivos, mock_prompt
    ):
        mock_selecionar_arquivos.side_effect = [["arquivo_teste.pdf"], []]
        mock_prompt.return_value = "4"
        mock_padrao = {
            "colunas": ["Código", "Causa", "Localização da Avaria", "Descrição"],
            "larguras": {"Código": 16.57, "Causa": 52.43, "Localização da Avaria": 52.57},
            "tema_visual": "simplo_3_colunas_desc"
        }
        mock_selecionar_padrao.return_value = mock_padrao

        with patch('app.LeitorPDF') as mock_leitor_cls, \
             patch('app.TradutorIA'), \
             patch('app.AgenteOrganizador'), \
             patch('app.GeradorExcel'), \
             patch('app.OCREngine'), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestKeyValidLength12345"}):
            
            mock_leitor_inst = MagicMock()
            mock_leitor_inst.obter_total_paginas.return_value = 5
            mock_leitor_cls.return_value = mock_leitor_inst

            app.main()

            mock_selecionar_padrao.assert_called_once()
            call_kwargs = mock_selecionar_padrao.call_args[1]
            self.assertIn("Qual o padrão do Excel a ser gerado?", call_kwargs.get("titulo", ""))

            mock_exec_extracao.assert_called_once()
            exec_kwargs = mock_exec_extracao.call_args[1]
            self.assertTrue(exec_kwargs.get("exportar_duplo"))
            self.assertEqual(exec_kwargs.get("padrao_predefinido"), mock_padrao)

    @patch('app.Prompt.ask')
    @patch('app.selecionar_arquivos_pdf')
    @patch('app.selecionar_padrao_planilha')
    @patch('app.executar_lote_multiplos_pdfs')
    @patch('app.Confirm.ask', return_value=False)
    def test_op4_chama_pergunta_padrao_imediatamente_lote_pdfs(
        self, mock_confirm, mock_exec_lote, mock_selecionar_padrao, mock_selecionar_arquivos, mock_prompt
    ):
        mock_selecionar_arquivos.side_effect = [["arquivo1.pdf", "arquivo2.pdf"], []]
        mock_prompt.return_value = "4"
        mock_padrao = {
            "colunas": ["Códigos", "Componente", "FMI", "Descrição"],
            "larguras": {"Códigos": 15.86, "Componente": 40.29, "FMI": 14.00, "Descrição": 48.57},
            "tema_visual": "simplo_azul"
        }
        mock_selecionar_padrao.return_value = mock_padrao

        with patch('app.LeitorPDF') as mock_leitor_cls, \
             patch('app.TradutorIA'), \
             patch('app.AgenteOrganizador'), \
             patch('app.GeradorExcel'), \
             patch('app.OCREngine'), \
             patch.dict(os.environ, {"GEMINI_API_KEY": "AIzaSyTestKeyValidLength12345"}):

            app.main()

            mock_selecionar_padrao.assert_called_once()
            mock_exec_lote.assert_called_once()
            exec_kwargs = mock_exec_lote.call_args[1]
            self.assertEqual(exec_kwargs.get("padrao_predefinido"), mock_padrao)


if __name__ == '__main__':
    unittest.main()

