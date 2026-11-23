import os
import sys
import tempfile
import unittest

# Adiciona o diretório base ao sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import (
    listar_arquivos_entrada,
    obter_pasta_saida_para_arquivo,
    obter_pasta_saida_para_lote,
    parsear_selecao_arquivos_pdf,
    agrupar_arquivos_por_pasta,
    PASTA_ENTRADA,
    PASTA_SAIDA
)


class TestSubpastasEntradaSaida(unittest.TestCase):
    def setUp(self):
        # Cria estrutura temporária para teste dentro de entrada/
        self.pasta_teste = os.path.join(PASTA_ENTRADA, "_teste_unitario_subpasta")
        self.sub_a = os.path.join(self.pasta_teste, "MarcaA")
        self.sub_b = os.path.join(self.pasta_teste, "MarcaB")
        os.makedirs(self.sub_a, exist_ok=True)
        os.makedirs(self.sub_b, exist_ok=True)

        self.arq_a1 = os.path.join(self.sub_a, "doc1.pdf")
        self.arq_a2 = os.path.join(self.sub_a, "doc2.pdf")
        self.arq_b1 = os.path.join(self.sub_b, "doc3.pdf")

        for arq in [self.arq_a1, self.arq_a2, self.arq_b1]:
            with open(arq, "w", encoding="utf-8") as f:
                f.write("%PDF-1.4 dummy")

    def tearDown(self):
        # Remove arquivos e pastas de teste
        for arq in [self.arq_a1, self.arq_a2, self.arq_b1]:
            if os.path.exists(arq):
                os.remove(arq)
        if os.path.exists(self.sub_a):
            os.rmdir(self.sub_a)
        if os.path.exists(self.sub_b):
            os.rmdir(self.sub_b)
        if os.path.exists(self.pasta_teste):
            os.rmdir(self.pasta_teste)

        # Remove subpastas geradas em saida se houver
        saida_teste = os.path.join(PASTA_SAIDA, "_teste_unitario_subpasta")
        if os.path.exists(os.path.join(saida_teste, "MarcaA")):
            os.rmdir(os.path.join(saida_teste, "MarcaA"))
        if os.path.exists(os.path.join(saida_teste, "MarcaB")):
            os.rmdir(os.path.join(saida_teste, "MarcaB"))
        if os.path.exists(saida_teste):
            os.rmdir(saida_teste)

    def test_listar_arquivos_recursivo(self):
        arquivos = listar_arquivos_entrada()
        caminhos_norm = [os.path.abspath(a) for a in arquivos]
        self.assertIn(os.path.abspath(self.arq_a1), caminhos_norm)
        self.assertIn(os.path.abspath(self.arq_a2), caminhos_norm)
        self.assertIn(os.path.abspath(self.arq_b1), caminhos_norm)

    def test_obter_pasta_saida_para_arquivo(self):
        # Arquivo na raiz de entrada -> saida/
        arq_raiz = os.path.join(PASTA_ENTRADA, "exemplo_raiz.pdf")
        pasta_raiz = obter_pasta_saida_para_arquivo(arq_raiz)
        self.assertEqual(os.path.abspath(pasta_raiz), os.path.abspath(PASTA_SAIDA))

        # Arquivo em subpasta -> saida/Subpasta/
        pasta_sub = obter_pasta_saida_para_arquivo(self.arq_a1)
        esperado = os.path.join(PASTA_SAIDA, "_teste_unitario_subpasta", "MarcaA")
        self.assertEqual(os.path.abspath(pasta_sub), os.path.abspath(esperado))
        self.assertTrue(os.path.exists(pasta_sub))

    def test_obter_pasta_saida_para_lote(self):
        # Todos os arquivos na mesma subpasta -> espelha a subpasta
        pasta_lote_a = obter_pasta_saida_para_lote([self.arq_a1, self.arq_a2])
        esperado_a = os.path.join(PASTA_SAIDA, "_teste_unitario_subpasta", "MarcaA")
        self.assertEqual(os.path.abspath(pasta_lote_a), os.path.abspath(esperado_a))

        # Arquivos de subpastas diferentes -> raiz saida/
        pasta_lote_misto = obter_pasta_saida_para_lote([self.arq_a1, self.arq_b1])
        self.assertEqual(os.path.abspath(pasta_lote_misto), os.path.abspath(PASTA_SAIDA))

    def test_parsear_selecao_por_nome_subpasta(self):
        arquivos = [self.arq_a1, self.arq_a2, self.arq_b1]

        # Digitando "MarcaA" deve selecionar os 2 arquivos de MarcaA
        selecionados = parsear_selecao_arquivos_pdf("MarcaA", arquivos)
        self.assertEqual(len(selecionados), 2)
        self.assertIn(self.arq_a1, selecionados)
        self.assertIn(self.arq_a2, selecionados)

        # Digitando "MarcaB" deve selecionar o arquivo de MarcaB
        selecionados_b = parsear_selecao_arquivos_pdf("MarcaB", arquivos)
        self.assertEqual(len(selecionados_b), 1)
        self.assertIn(self.arq_b1, selecionados_b)

    def test_agrupar_arquivos_por_pasta(self):
        arquivos = [self.arq_a1, self.arq_a2, self.arq_b1]
        grupos = agrupar_arquivos_por_pasta(arquivos)
        
        # Chaves relativas esperadas
        chave_a = os.path.relpath(self.sub_a, PASTA_ENTRADA).replace("\\", "/")
        chave_b = os.path.relpath(self.sub_b, PASTA_ENTRADA).replace("\\", "/")

        self.assertIn(chave_a, grupos)
        self.assertIn(chave_b, grupos)
        self.assertEqual(len(grupos[chave_a]), 2)
        self.assertEqual(len(grupos[chave_b]), 1)


if __name__ == "__main__":
    unittest.main()

