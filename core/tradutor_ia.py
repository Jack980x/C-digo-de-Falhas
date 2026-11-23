import os
import json
from typing import List, Dict, Any, Optional
from dotenv import load_dotenv

from core.agente_tradutor import AgenteExtrator, AgenteTradutor, AgenteAuditor, reparar_e_carregar_json
from core.agente_organizador import AgenteOrganizador

load_dotenv()

class TradutorIA:
    def __init__(self, api_key: str = None, model_name: str = "gemini-3.8-flash"):
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        self.model_name = model_name
        self.modelos_candidatos = [model_name, "gemini-3.8-flash", "gemini-flash-latest"]
        # Remove duplicados mantendo ordem
        self.modelos_candidatos = list(dict.fromkeys(self.modelos_candidatos))
        self.client = None
        
        if self.api_key and self.api_key != "sua_chave_aqui":
            try:
                from google import genai
                self.client = genai.Client(api_key=self.api_key)
            except Exception as e:
                print(f"[AVISO IA] Não foi possível inicializar o cliente Gemini: {e}")

        # Instancia os agentes especialistas
        self.agente_extrator = AgenteExtrator(self.client, self.modelos_candidatos)
        self.agente_tradutor = AgenteTradutor(self.client, self.modelos_candidatos)
        self.agente_auditor = AgenteAuditor()
        self.pasta_cache = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", ".cache_extracoes")
        os.makedirs(self.pasta_cache, exist_ok=True)

    def _gerar_chave_cache(self, operacao: str, texto_bruto: str, colunas: List[str] = None, modulo_alvo: str = None, codigos_gt: List[str] = None) -> str:
        import hashlib
        payload = json.dumps({
            "v": "2.1",
            "op": operacao,
            "txt": (texto_bruto or "").strip(),
            "cols": [str(c).strip() for c in (colunas or [])],
            "mod": str(modulo_alvo or "").strip(),
            "gt": [str(c).strip().upper() for c in (codigos_gt or [])]
        }, ensure_ascii=False, sort_keys=True)
        return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:28]

    def _ler_cache_disco(self, chave_hash: str) -> Optional[List[Dict[str, Any]]]:
        caminho = os.path.join(self.pasta_cache, f"{chave_hash}.json")
        if os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    dados = json.load(f)
                if isinstance(dados, list) and dados:
                    return dados
            except Exception:
                pass
        return None

    def _salvar_cache_disco(self, chave_hash: str, registros: List[Dict[str, Any]]) -> None:
        if not registros:
            return
        caminho = os.path.join(self.pasta_cache, f"{chave_hash}.json")
        try:
            with open(caminho, "w", encoding="utf-8") as f:
                json.dump(registros, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    @staticmethod
    def _aplicar_colunas_codigo_secundario(
        registros: List[Dict[str, Any]],
        colunas_desejadas: List[str],
        texto_bruto: str,
        codigos_ground_truth: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """
        Quando a tabela possui duas colunas de código lado a lado (ex: 'DTC Designação do Veículo' U240088
        e 'DTCs' E40088), preenche automaticamente a 2ª coluna de código se ambas foram solicitadas,
        ou converte para o código da 2ª coluna caso o usuário tenha solicitado especificamente 'DTCs'.
        """
        if not registros or not colunas_desejadas or not texto_bruto:
            return registros
        from core.leitor_pdf import LeitorPDF
        from core.gerador_excel import GeradorExcel
        mapa_sec = LeitorPDF.mapear_codigos_secundarios_pagina(texto_bruto, codigos_ground_truth or [])
        if not mapa_sec:
            return registros

        cols_cod = [c for c in colunas_desejadas if GeradorExcel.classificar_tipo_coluna(c) == "codigo"]
        if len(cols_cod) >= 2:
            col_pri, col_sec = cols_cod[0], cols_cod[1]
            for r in registros:
                val_pri = str(r.get(col_pri, "")).strip().upper()
                if (not val_pri or val_pri == "-") and str(r.get(col_sec, "")).strip().upper() in mapa_sec:
                    val_pri = str(r.get(col_sec, "")).strip().upper()
                    r[col_pri] = val_pri
                if val_pri in mapa_sec:
                    r[col_sec] = mapa_sec[val_pri]
        elif len(cols_cod) == 1:
            col_unica = cols_cod[0]
            if col_unica.strip().lower() in {"dtcs", "dtc hex", "código hex", "codigo hex", "dtcs (hex)"}:
                for r in registros:
                    val_pri = str(r.get(col_unica, "")).strip().upper()
                    if val_pri in mapa_sec:
                        r[col_unica] = mapa_sec[val_pri]
        return registros

    def carregar_dados_glossario(self) -> Dict[str, Any]:
        """Carrega os dados completos do glossario_tecnico.json."""
        caminho_glossario = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "glossario_tecnico.json")
        if os.path.exists(caminho_glossario):
            try:
                with open(caminho_glossario, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"termos_e_traducoes": {}, "siglas_preservadas": []}

    def salvar_dados_glossario(self, dados: Dict[str, Any]) -> str:
        """Salva ou atualiza os dados no glossario_tecnico.json."""
        caminho_glossario = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "glossario_tecnico.json")
        os.makedirs(os.path.dirname(caminho_glossario), exist_ok=True)
        with open(caminho_glossario, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=2, ensure_ascii=False)
        return caminho_glossario

    def _obter_texto_glossario(self) -> str:
        """Carrega os termos do glossario_tecnico.json dinamicamente para o prompt."""
        dados = self.carregar_dados_glossario()
        termos = dados.get("termos_e_traducoes", {})
        if termos:
            return "\n".join([f'   - "{k}" -> "{v}"' for k, v in termos.items()])
        return '   - "Short to ground" -> "Curto-circuito ao aterramento"\n   - "Short to battery" -> "Curto-circuito ao positivo"'

    def _obter_texto_siglas(self) -> str:
        """Carrega as siglas preservadas do glossario_tecnico.json dinamicamente para o prompt."""
        dados = self.carregar_dados_glossario()
        siglas = dados.get("siglas_preservadas", [])
        if siglas:
            return ", ".join(siglas)
        return "ECU, TCU, ABS, EBS, ESP, EGR, DPF, SCR, CAN, PWM, OBD, DTC, SPN, FMI, MAF, MAP, VGT"

    def extrair_novos_termos(self, texto_bruto: str, contexto_segmento: str = "Veículos Pesados e Máquinas Agrícolas") -> Dict[str, Any]:
        """
        Analisa o texto de um manual e extrai vocabulário técnico relevante:
        - Termos em inglês com sugestão de tradução técnica em PT-BR
        - Siglas técnicas a preservar sem tradução
        """
        if not self.client:
            raise ValueError("Cliente Gemini não inicializado.")

        glossario_atual = self.carregar_dados_glossario()
        termos_existentes = list(glossario_atual.get("termos_e_traducoes", {}).keys())

        prompt_sistema = f"""Você é um lexicógrafo e engenheiro especialista em terminologia técnica automotiva, diesel pesado e máquinas agrícolas ({contexto_segmento}).
Sua missão é analisar o texto original de um manual de códigos de falhas e extrair termos técnicos importantes em inglês para enriquecer o glossário do sistema.

DIRETRIZES:
1. Extraia termos técnicos de hardware, testes elétricos, mecânica, eletrônica embarcada, tipos de erro e condições de funcionamento.
   Exemplos de termos relevantes: "Open circuit", "Short to ground", "Pull-up resistor", "Gate Interlock", "Solenoid valve", "Speed sensor", "Derated engine".
2. Para cada termo identificado em inglês, forneça a tradução técnica ideal para o Português do Brasil (PT-BR) conforme o padrão das montadoras (Scania, Iveco, Volvo, John Deere).
   REGRA FUNDAMENTAL: Para termos relacionados a GND/Ground, traduza SEMPRE como "aterramento" (evite "terra").
3. Identifique também siglas técnicas consagradas (ex: CAN, ECU, VGT, PWM) que devem ser preservadas sem tradução.
4. Desconsidere palavras genéricas do inglês comum (ex: "the", "and", "is", "page").
5. Retorne EXCLUSIVAMENTE um objeto JSON no formato:
{{
  "termos_sugeridos": {{
    "Termo em Inglês 1": "Tradução Técnica em PT-BR 1",
    "Termo em Inglês 2": "Tradução Técnica em PT-BR 2"
  }},
  "siglas_detectadas": ["SIGLA1", "SIGLA2"]
}}"""

        prompt_usuario = f"""Analise o texto do manual abaixo e identifique todos os termos técnicos relevantes e siglas:
---
{texto_bruto[:25000]}
---
Termos já cadastrados no glossário (para sua referência, evite duplicar se já estiver perfeito):
{', '.join(termos_existentes[:50])}"""

        for modelo in self.modelos_candidatos:
            try:
                resposta = self.client.models.generate_content(
                    model=modelo,
                    contents=[prompt_sistema, prompt_usuario],
                    config={
                        "temperature": 0.2,
                        "response_mime_type": "application/json"
                    }
                )
                texto = resposta.text.strip()
                if texto.startswith("```"):
                    linhas = texto.split("\n")
                    if linhas[0].startswith("```"):
                        linhas = linhas[1:]
                    if linhas and linhas[-1].startswith("```"):
                        linhas = linhas[:-1]
                    texto = "\n".join(linhas).strip()
                return json.loads(texto)
            except Exception as e:
                if "404" in str(e) or "NOT_FOUND" in str(e) or "not found" in str(e).lower():
                    continue
                else:
                    raise e
        return {"termos_sugeridos": {}, "siglas_detectadas": []}

    def identificar_modulos_documento(
        self,
        resumo_texto: str,
        nome_arquivo: str = "",
        modulos_regex: List[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """
        Identifica os módulos / centrais eletrônicas (ECUs / Sistemas) presentes nas tabelas ou seções do PDF.
        Combina detecção estrutural via Regex, análise inteligente via IA e sugestão do nome do arquivo.
        """
        import re
        modulos_finais = []
        vistos = set()

        def adicionar_modulo(nome: str, paginas: str = "", descricao: str = ""):
            if not nome:
                return
            n_limpo = str(nome).strip()
            if not n_limpo or len(n_limpo) < 2:
                return
            chave = re.sub(r'[\s_-]+', '', n_limpo).upper()
            if chave in vistos:
                return
            vistos.add(chave)
            info_partes = [p for p in [paginas.strip(), descricao.strip()] if p]
            modulos_finais.append({
                "nome": n_limpo,
                "paginas": paginas.strip(),
                "descricao": descricao.strip(),
                "info": " — ".join(info_partes) if info_partes else "Detectado no documento"
            })

        # 1. Consulta IA se disponível e se houver texto
        if self.client and resumo_texto and len(resumo_texto.strip()) > 30:
            prompt_sistema = """Você é um engenheiro especialista em manuais de diagnóstico automotivo, caminhões, ônibus e máquinas agrícolas.
Sua missão é analisar os cabeçalhos e trechos das páginas de um manual PDF e listar APENAS os MÓDULOS ELETRÔNICOS (ECUs / Centrais / Sistemas principais, ex: HCML, HCMR, EM_SH, IC M06, EDC7UC31, BCM2, VCM, EBS, TCU, ADM) que possuem tabelas ou seções de códigos de falhas no documento.

REGRAS ESTRITAS:
1. Liste APENAS módulos/centrais que sejam donos de seções/tabelas do documento ou que apareçam na coluna de Módulo/ECU da tabela.
2. PROIBIDO listar componentes internos, sensores, válvulas, pinos, relés ou sub-circuitos citados apenas dentro das descrições das falhas (ex: NÃO liste BIC1..BIC20, Pin A01, CAN 1, SBC, IG3, EEPROM, Sensor).
3. Se o documento inteiro for uma lista genérica sem dividir por módulos específicos no texto, retorne "modulos": [].
4. Retorne EXCLUSIVAMENTE um JSON válido no formato:
{
  "modulos": [
    {
      "nome": "SIGLA_OU_NOME_DO_MODULO",
      "paginas": "Pág. 1-4",
      "descricao": "Identificação breve do módulo/seção"
    }
  ]
}"""
            prompt_usuario = f"""Nome do arquivo PDF: {nome_arquivo}
Resumo estruturado por página:
---
{resumo_texto[:16000]}
---
Identifique os módulos/centrais que possuem tabelas ou seções de códigos neste documento:"""

            for modelo in self.modelos_candidatos:
                try:
                    resposta = self.client.models.generate_content(
                        model=modelo,
                        contents=[prompt_sistema, prompt_usuario],
                        config={
                            "temperature": 0.0,
                            "response_mime_type": "application/json"
                        }
                    )
                    dados = reparar_e_carregar_json(resposta.text)
                    lista_ia = []
                    if isinstance(dados, dict):
                        lista_ia = dados.get("modulos", [])
                    elif isinstance(dados, list):
                        lista_ia = dados
                    if isinstance(lista_ia, list):
                        for item in lista_ia:
                            if isinstance(item, dict) and item.get("nome"):
                                adicionar_modulo(
                                    nome=item.get("nome"),
                                    paginas=str(item.get("paginas", "Detectado no PDF")),
                                    descricao=str(item.get("descricao", "Módulo identificado no documento"))
                                )
                    break
                except Exception:
                    continue

        # 2. Adiciona módulos detectados via Regex estrutural (se ainda não incluídos)
        if modulos_regex:
            for m_reg in modulos_regex:
                adicionar_modulo(
                    nome=m_reg.get("nome", ""),
                    paginas=m_reg.get("paginas", ""),
                    descricao=m_reg.get("descricao", "Detectado na tabela do PDF")
                )

        # 3. Se o nome do arquivo contiver múltiplos módulos separados por ' e ' ou ',' (ex: "HCML e HCMR")
        if nome_arquivo:
            nome_limpo = nome_arquivo.replace("_traduzido", "").strip()
            dividido_arquivo = False
            if re.search(r'\s+e\s+|,', nome_limpo, flags=re.I):
                partes_arq = [p.strip() for p in re.split(r'\s+e\s+|,', nome_limpo, flags=re.I) if p.strip()]
                if len(partes_arq) >= 2 and all(len(p) <= 15 for p in partes_arq):
                    for p in partes_arq:
                        adicionar_modulo(p, paginas="Arquivo", descricao="Sugestão do nome do arquivo")
                    dividido_arquivo = True

            # Se não foi dividido e nenhum módulo específico foi detectado no PDF, usa o nome do arquivo
            if not dividido_arquivo and not modulos_finais:
                adicionar_modulo(nome_limpo, paginas="Documento", descricao="Sugestão baseada no nome do arquivo")

        return modulos_finais

    def estruturar_buscar_codigo(
        self,
        texto_bruto: str,
        contexto_segmento: str = "Linha Pesada e Agrícola",
        modulo_alvo: str = None,
        texto_nativo: str = None,
        codigos_ground_truth: List[str] = None
    ) -> List[Dict[str, Any]]:
        """
        Operação 3: 'Excel do Buscar Código'
        Extrai e estrutura tabelas de códigos de falhas com as seguintes regras fundamentais:
        1. 'Código': Junção direta do código base da falha + FMI em uma única string alfanumérica contínua (ex: 40021 + 01 -> 4002101, 40B16 + 08 -> 40B1608).
           - Preserva zeros à esquerda (ex: 047228 + 03 -> 04722803).
           - Formata FMI sempre com 2 dígitos se tiver 1 (ex: 1 -> 01, 8 -> 08).
           - Cada FMI gera uma linha independente (sem mesclagem ou agrupamento).
        2. 'Localização da Avaria': Componente físico / sensor / subsistema traduzido estritamente para PT-BR técnico.
        3. 'Causa': Causa física ou elétrica traduzida de forma puramente técnica e literal do FMI Description do manual.
        4. 'Descrição': Tradução puramente literal da coluna Description do manual (estritamente '-' se vazia).
        """
        if not self.client:
            raise ValueError(
                "Chave da API do Gemini (GEMINI_API_KEY) não configurada.\n"
                "Edite o arquivo .env e insira sua chave da API do Google AI Studio."
            )

        if not texto_bruto.strip():
            return []

        chave_cache = self._gerar_chave_cache("buscar_codigo", texto_bruto, ["Código", "Localização da Avaria", "Causa", "Descrição"], modulo_alvo, codigos_ground_truth)
        em_cache = self._ler_cache_disco(chave_cache)
        if em_cache:
            return em_cache

        # Se a tabela extraída já estiver em Português no formato estruturado 'CODIGO | DESCRICAO'
        registros_ptbr_bc = self._tentar_buscar_codigo_ptbr_direto(texto_bruto, codigos_ground_truth)
        if registros_ptbr_bc:
            res_bc = self._higienizar_registros_buscar_codigo(
                registros_ptbr_bc,
                codigos_ground_truth=codigos_ground_truth,
                texto_fonte=texto_bruto
            )
            self._salvar_cache_disco(chave_cache, res_bc)
            return res_bc

        # Se a página/imagem for muito longa (> 45 códigos, ex: print rolante de tabela inteira),
        # divide em lotes de até 40 códigos em paralelo para evitar truncamento de JSON ou timeout 504
        if (codigos_ground_truth and len(codigos_ground_truth) > 45) or len(texto_bruto.splitlines()) > 100:
            import concurrent.futures
            blocos = self.agente_extrator._segmentar_texto_em_lotes(texto_bruto, codigos_ground_truth, tamanho_lote=40)
            if len(blocos) > 1:
                resultados_blocos = [None] * len(blocos)
                with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, len(blocos))) as executor:
                    futures = {
                        executor.submit(
                            self._estruturar_bloco_buscar_codigo,
                            b_txt,
                            contexto_segmento,
                            modulo_alvo,
                            texto_nativo,
                            b_gt
                        ): idx
                        for idx, (b_txt, b_gt) in enumerate(blocos)
                    }
                    for future in concurrent.futures.as_completed(futures):
                        idx = futures[future]
                        resultados_blocos[idx] = future.result() or []
                todos_brutos = []
                for rb in resultados_blocos:
                    if rb:
                        todos_brutos.extend(rb)
                res_lotes = self._higienizar_registros_buscar_codigo(
                    todos_brutos,
                    codigos_ground_truth=codigos_ground_truth,
                    texto_fonte=texto_bruto
                )
                self._salvar_cache_disco(chave_cache, res_lotes)
                return res_lotes

        lista_registros = self._estruturar_bloco_buscar_codigo(
            texto_bruto=texto_bruto,
            contexto_segmento=contexto_segmento,
            modulo_alvo=modulo_alvo,
            texto_nativo=texto_nativo,
            codigos_ground_truth=codigos_ground_truth
        )
        res_final_bc = self._higienizar_registros_buscar_codigo(
            lista_registros,
            codigos_ground_truth=codigos_ground_truth,
            texto_fonte=texto_bruto
        )
        self._salvar_cache_disco(chave_cache, res_final_bc)
        return res_final_bc

    def _estruturar_bloco_buscar_codigo(
        self,
        texto_bruto: str,
        contexto_segmento: str = "Linha Pesada e Agrícola",
        modulo_alvo: str = None,
        texto_nativo: str = None,
        codigos_ground_truth: List[str] = None
    ) -> List[Dict[str, Any]]:
        colunas_esperadas = ["Código", "Localização da Avaria", "Causa", "Descrição"]
        colunas_str = ', '.join([f'"{c}"' for c in colunas_esperadas])

        import re
        tem_coluna_modulo = bool(re.search(
            r'\b(?:ECU\s*Name|Module\s*Name|Control\s*Unit|Central\s*Eletr[oô]nica|Nome\s*do\s*M[oó]dulo)\b',
            texto_bruto,
            re.I
        ))
        alvo_presente = False
        if modulo_alvo:
            partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', modulo_alvo.strip()) if x]
            if partes_m:
                pat_m = re.compile(r'\b' + r'[\s_-]*'.join(partes_m) + r'\b', re.I)
                alvo_presente = bool(pat_m.search(texto_bruto))

        aplicar_filtro_estrito = bool(modulo_alvo and (tem_coluna_modulo or alvo_presente))

        instrucao_modulo = ""
        if aplicar_filtro_estrito:
            instrucao_modulo = f"""
FILTRO DE MÓDULO SOLICITADO: '{modulo_alvo}'.
- Extraia TODAS as falhas e códigos pertencentes ao módulo '{modulo_alvo}' (ou que estejam sob a seção/continuação de '{modulo_alvo}').
- Se a página contiver uma coluna explícita de módulo (ex: 'ECU Name') identificando códigos de OUTROS módulos diferentes de '{modulo_alvo}', ignore os de outros módulos.
- Se a tabela NÃO possuir coluna explícita de módulo por linha, extraia TODOS os códigos da página! NUNCA descarte códigos apenas porque a descrição menciona siglas como ECU, BCM, ESP, EPS, ABS, CAN, etc.
"""
        elif modulo_alvo:
            instrucao_modulo = f"""
MÓDULO DE DESTINO: '{modulo_alvo}'.
- Esta página é uma tabela geral de códigos de falha (sem coluna restritiva de módulo). Extraia ABSOLUTAMENTE TODOS os códigos de falha presentes nesta página!
- NUNCA descarte ou ignore códigos porque a descrição menciona siglas como ECU, BCM, ESP, EPS, ABS, LKA, CAN, etc.
"""

        instrucao_gt = ""
        if codigos_ground_truth:
            instrucao_gt = f"""
CÓDIGOS DETECTADOS NO DOCUMENTO (TOTAL: {len(codigos_ground_truth)} CÓDIGOS OBRIGATÓRIOS — REFERÊNCIA DE GRAFIA EXATA):
{', '.join(codigos_ground_truth)}
- É TERMINANTEMENTE OBRIGATÓRIO extrair TODOS os {len(codigos_ground_truth)} códigos acima sem omitir nenhum!
- Use esta lista como referência para manter a grafia exata de cada código.
- Quando uma linha da tabela possuir duas colunas de código lado a lado (ex: 'DTC Designação do Veículo' U240088 / C110116 na 1ª coluna e 'DTCs' hexadecimal E40088 / 510116 na 2ª coluna), extraia APENAS o código principal listado acima (1 registro por linha). NUNCA duplique a linha extraindo a 2ª coluna como outro código!
"""

        prompt_sistema = f"""Você é um tradutor técnico e engenheiro especialista em diagnósticos automotivos e máquinas agrícolas ({contexto_segmento}).
Sua missão única é ler a tabela do manual e extrair EXATAMENTE as falhas pertencentes ao módulo solicitado, realizando ESTRITAMENTE A TRADUÇÃO TÉCNICA LITERAL DO TEXTO ORIGINAL, SEM NENHUMA INTERPRETAÇÃO E SEM INVENTAR NADA.{instrucao_modulo}{instrucao_gt}

REGRAS DE OURO OBRIGATÓRIAS (ZERO INTERPRETAÇÃO, TRADUÇÃO PURAMENTE LITERAL):

1. TRADUÇÃO PURAMENTE LITERAL E TÉCNICA (PROIBIDO INTERPRETAR OU ADICIONAR):
   - Traduza de forma técnica, direta e enxuta o que está escrito no manual.
   - NUNCA adicione explicações, comentários, suposições, hipóteses ou sinônimos explicativos com barras (ex: NUNCA acrescente '/ modo de falha não identificável', '/ ECU', 'via CAN...', etc.).
   - Se o texto em inglês diz 'Root cause not known', traduza estritamente como 'Causa raiz desconhecida'.
   - Se diz 'bad intelligent device or component', traduza estritamente como 'Dispositivo ou componente inteligente com defeito'.
   - Se diz 'Abnormal update rate', traduza estritamente como 'Taxa de atualização anormal'.
   - Use 'aterramento' rigorosamente para GND/ground (nunca 'terra').
   - Use 'positivo' ou 'positivo (Vbatt)' para Vbatt/battery.

2. COLUNA "DESCRIÇÃO" - SÍNTESE INTELIGENTE EM "MODO COMPLETO" (FALHA TÉCNICA + RECOMENDAÇÃO DE TESTE):
   - Se a coluna Description no manual estiver vazia, em branco, ausente, contiver '-' ou se for apenas um link de diagnóstico repetindo a falha:
     É OBRIGATÓRIO gerar a "Descrição" no MODO COMPLETO combinando a Localização da Avaria e a Causa em duas partes fluentes:
     * PARTE 1 (Identificação Técnica da Falha): Frase técnica em português automotivo associando a Causa ao Componente/Localização (ex: "Circuito aberto detectado no resistor interno do comutador do espelho.", "Curto-circuito ao positivo (Vbatt) no circuito do airbag frontal do passageiro.", "Falha de comunicação no barramento CAN da ECU do motor.").
     * PARTE 2 (Pista de Teste Recomendada): Orientação prática de oficina baseada na natureza da falha (ex: teste de continuidade e terminais para circuito aberto; isolamento contra massa para curto ao aterramento; inspeção de chicote contra esmagamento com linha de alimentação para curto ao positivo; medição de terminação de 60 Ω e tensões CAN-H/CAN-L para barramento CAN; verificação de 5V de referência e aterramento de sinal para plausibilidade/faixa).
   - NUNCA repita o texto padrão genérico fixo se puder gerar a síntese baseada na Causa e Localização da Avaria.
   - Se houver texto explicativo adicional REAL impresso na coluna Description do manual (com passos ou sintomas detalhados além do nome da falha), traduza fielmente esse texto original.

3. DIVISÃO OBRIGATÓRIA DE DESCRIÇÕES MULTILINHA / TÓPICOS POR FMI (1-PARA-1):
   - Quando um componente/código possui múltiplos FMIs (ex: 00, 01, 02, 0B, 12) e a coluna Description no manual contiver múltiplos tópicos ou linhas (separadas por hífens '-', quebras de linha ou frases com ponto-e-vírgula ';'):
     * CADA PARTE/TÓPICO DA DESCRIÇÃO CORRESPONDE RIGOROSAMENTE A UMA CAUSA / FMI NA MESMA ORDEM!
     * JAMAIS copie o bloco inteiro com todas as frases reunidas em todas as linhas daquele código!
     * O 1º tópico da descrição vai para o 1º FMI;
     * O 2º tópico da descrição vai para o 2º FMI;
     * O 3º tópico da descrição vai para o 3º FMI;
     * E assim por diante.
     * Exemplo para o código 42774 (VIM):
       - "4277400" (FMI 00): "Descrição": "Velocidade filtrada do veículo MTCO maior que a velocidade de frenagem de um determinado limite EOL km/h"
       - "4277401" (FMI 01): "Descrição": "Velocidade filtrada do veículo MTCO menor que a velocidade de frenagem de um determinado limite EOL km/h"
       - "4277402" (FMI 02): "Descrição": "Velocidade do veículo totalmente ausente"
       - "427740B" (FMI 0B): "Descrição": "Informação oficial TCO1 ausente, mas uma velocidade de recuperação está disponível"
       - "4277412" (FMI 12): "Descrição": "Velocidade do veículo MTCO maior que o limite EOL Km/h"

4. FIDELIDADE ABSOLUTA E TRATAMENTO OBRIGATÓRIO DE FMIs COM 2 NÚMEROS (DEC E HEX OU SEPARADOS POR BARRA):
   - Quando a tabela do manual contiver colunas de FMI Decimal [DEC] e Hexadecimal [HEX] (ou contiver dois números para a mesma falha, ex: 12 e 0C, ou 04 e 06):
     * Se os dois números forem idênticos (ex: 02 e 02, 03 e 03, 04 e 04):
       Gere exatamente UMA linha (ex: "Código": "1003102", "_cod_base": "10031", "_fmi": "02").
     * Se os dois números forem DIFERENTES (ex: 12 em Decimal e 0C em Hexadecimal, 10 e 0A, 11 e 0B, 13 e 0D, 14 e 0E, 15 e 0F):
       É TERMINANTEMENTE OBRIGATÓRIO GERAR DUAS LINHAS SEPARADAS NO ARRAY DE SAÍDA:
       - Linha 1 com o FMI Decimal: {{"Código": "1003112", "_cod_base": "10031", "_fmi": "12"}}
       - Linha 2 com o FMI Hexadecimal: {{"Código": "100310C", "_cod_base": "10031", "_fmi": "0C"}}
       Ambas as linhas com a mesma Localização da Avaria, Causa e Descrição!
       Isso é indispensável para que o profissional encontre a falha independentemente de seu scanner exibir o código em decimal (1003112) ou em hexadecimal (100310C).
   - Se o manual trouxer múltiplos FMIs separados por barra ou na mesma célula (ex: "04 / 06"):
     Gere duas linhas: uma para o primeiro FMI ("1003104") e outra para o segundo FMI ("1003106").
   - Se um código no manual só aparece com um FMI único (ex: apenas 02), gere apenas uma única linha ("4270102"). NUNCA adicione FMIs inventados que não estejam no manual.

5. REGRAS DE CÓDIGOS E DIFERENCIAÇÃO ENTRE FMI (J1939) E STATUS DE DTC (UDS / ISO 14229):
   - CASO 1: PROTOCOLO SAE J1939 (CÓDIGO SPN DE 2 A 5 DÍGITOS + FMI):
     * O código base (SPN) possui de 2 a 5 dígitos (ex: 10031, 42701, 42741, 40064, 652).
     * O FMI possui 2 dígitos (ex: 01, 02, 0A, 0C, 12).
     * NESTE CASO, o "Código" da tabela final DEVE ser a unificação: SPN + FMI (ex: "1003102", "100310C", "1003112", "4270102").
     * Preencha "_cod_base": "10031", "_fmi": "02".
   
   - CASO 2: PROTOCOLO UDS / ISO 14229 / ISO 13400 / OBD-UDS (DTC COMPLETO DE 5 A 7 CARACTERES):
     * O código de falha (DTC identifier) já é completo e possui 6 caracteres hexadecimais (ex: "094249", "400100", "056000", "101087", "05FF04", "062C96") OU começa com P, C, B ou U seguido de 4 a 6 caracteres hexadecimais (ex: "C130242", "U014087", "C135004", "P0087").
     * A coluna "DTC status", "Status" ou valores como "02", "2F", "60" representam apenas flags operacionais de ciclo de teste da ECU (ex: Bit 1: Test failed this operation cycle = 02) e NÃO FAZEM PARTE DO CÓDIGO!
     * NESTE CASO, É TERMINANTEMENTE PROIBIDO UNIFICAR O STATUS AO CÓDIGO!
     * O "Código" DEVE SER ESTRITAMENTE O IDENTIFICADOR ORIGINAL COMPLETO (ex: "094249", "400100", "C130242", "U014087").
     * NUNCA crie códigos de 8 dígitos como "09424902" ou "40010002".
     * Preencha "_cod_base": com o código completo (ex: "094249" ou "C130242"), "_fmi": "-".

   - REGRAS DAS OUTRAS COLUNAS DO 'BUSCAR CÓDIGO':
     * "Localização da Avaria": Tradução técnica literal direta da coluna de componente/sensor/sistema do manual. NUNCA misture causas ou condições elétricas aqui.
     * "Causa": Tradução técnica da coluna de defeito/causa/condição/fmi_description do manual.
       - Quando contiver fórmulas elétricas ou inequações (ex: "V>8V", "V<2V", "6V<V<8V", "Vout/Vc>1V"): interprete o significado técnico da falha em português (ex: "Tensão acima do normal / Curto ao positivo (V > 8V)", "Tensão abaixo do normal / Curto ao aterramento (V < 2V)") mantendo a fórmula original entre parênteses!
       - Se não houver causa no manual, coloque "-".
     * "Descrição": Tradução literal da parte correspondente da descrição para aquele FMI específico. Se vazia ou repetitiva no manual, gere em "Modo Completo" (Parte 1: Frase técnica da falha + Parte 2: Pista prática de teste).

6. GLOSSÁRIO TÉCNICO APLICÁVEL:
{self._obter_texto_glossario()}

7. FORMATO DE SAÍDA EXCLUSIVO (ARRAY JSON):
[
  {{
    "Código": "4270102",
    "_cod_base": "42701",
    "_fmi": "02",
    "Localização da Avaria": "Controle de cruzeiro Set + (Entrada digital 5)",
    "Causa": "Dados erráticos, intermitentes ou incorretos (aplicação (Eltrac))",
    "Descrição": "Dados erráticos, intermitentes ou incorretos detectados no circuito de Controle de cruzeiro Set + (Entrada digital 5). Verifique a alimentação do sensor (tipicamente 5,0 V ou 12 V) e a referência de aterramento, monitorando o sinal via scanner durante a operação."
  }}
]"""

        prompt_usuario = f"""Texto extraído da página do manual:
---
{texto_bruto}
---

Extraia com fidelidade literal e estrita apenas as falhas existentes no texto acima para o modo 'Buscar Código'."""

        ultimo_erro = None
        for modelo in self.modelos_candidatos:
            try:
                resposta = self.client.models.generate_content(
                    model=modelo,
                    contents=[prompt_sistema, prompt_usuario],
                    config={
                        "temperature": 0.0,
                        "response_mime_type": "application/json"
                    }
                )

                dados = reparar_e_carregar_json(resposta.text)
                self.model_name = modelo

                lista_registros = []
                if isinstance(dados, dict):
                    for v in dados.values():
                        if isinstance(v, list):
                            lista_registros = v
                            break
                    if not lista_registros:
                        lista_registros = [dados]
                elif isinstance(dados, list):
                    lista_registros = dados

                return lista_registros

            except Exception as e:
                ultimo_erro = e
                err_low = str(e).lower()
                if any(x in err_low for x in ["404", "not_found", "not found", "504", "deadline_exceeded", "503", "unavailable", "timeout"]):
                    print(f"[AVISO IA BUSCAR CÓDIGO] Indisponibilidade/Timeout no modelo {modelo} ({e}). Tentando próximo modelo...")
                    continue
                if isinstance(e, json.JSONDecodeError) or any(x in err_low for x in ["delimiter", "json", "expecting"]):
                    print(f"[AVISO IA BUSCAR CÓDIGO] Formato JSON malformado no modelo {modelo} ({e}). Tentando próximo modelo...")
                    continue
                else:
                    print(f"[ERRO IA BUSCAR CÓDIGO] Falha no modelo {modelo}: {e}")
                    raise e

        print(f"[ERRO IA BUSCAR CÓDIGO] Nenhum modelo disponível: {ultimo_erro}")
        raise ultimo_erro

    def _higienizar_registros_buscar_codigo(
        self,
        registros: List[Dict[str, Any]],
        codigos_ground_truth: List[str] = None,
        texto_fonte: str = None
    ) -> List[Dict[str, Any]]:
        """
        Higieniza e valida os registros para o modo Buscar Código:
        - Diferencia com precisão SAE J1939 (unifica SPN + FMI, ex: 1003102, 1003112, 100310C) vs UDS/ISO 14229 (DTC de 6 dígitos)
        - Desmembra se múltiplos códigos ou FMIs foram combinados com '/' ou na mesma linha (ex: 12 / 0C gera 1003112 e 100310C)
        - Garante chaves corretas: Código, Localização da Avaria, Causa, Descrição
        - Corrige 'terra' para 'aterramento'
        - Organiza mantendo cada código base e todos os seus FMIs agrupados juntos um abaixo do outro
        """
        import re
        from core.agente_organizador import AgenteOrganizador

        resultado = []
        gt_set = (
            {c.upper() for c in (codigos_ground_truth or [])}
            | {re.sub(r'[\s:_-]+', '', c).upper() for c in (codigos_ground_truth or [])}
        )

        for reg in registros:
            if not isinstance(reg, dict):
                continue

            novo_reg = {}
            cod_val = None
            fmi_val = None
            loc_val = None
            causa_val = None
            desc_val = None

            for k, v in reg.items():
                k_low = k.lower()
                # Não confundir coluna de status com código ou FMI
                if any(x in k_low for x in ["código", "codigo", "code", "dtc", "spn"]) and "status" not in k_low:
                    cod_val = str(v).strip()
                elif "fmi" in k_low and "status" not in k_low:
                    fmi_val = str(v).strip()
                elif any(x in k_low for x in ["localização", "localizacao", "componente", "component", "sensore"]):
                    loc_val = str(v).strip()
                elif any(x in k_low for x in ["causa", "cause", "defeito"]):
                    causa_val = str(v).strip()
                elif any(x in k_low for x in ["descrição", "descricao", "description", "sinal", "segnale"]):
                    desc_val = str(v).strip()

            if cod_val:
                from core.ocr_engine import OCREngine
                cod_val = re.sub(r'[\s:]+', '', cod_val).upper()
                cod_val = OCREngine.normalizar_codigo_dtc_token(cod_val).upper()
                cod_val, _ = AgenteAuditor.reconciliar_codigo(
                    cod=cod_val,
                    texto_fonte=texto_fonte,
                    codigos_ground_truth=codigos_ground_truth
                )
                if (
                    codigos_ground_truth
                    and len(codigos_ground_truth) >= 2
                    and all(re.match(r'^[PCBU]', str(c).strip().upper()) for c in codigos_ground_truth if c)
                    and not re.match(r'^[PCBU]', cod_val)
                ):
                    continue
            else:
                cod_val = "-"

            # Se o LLM já colocou _cod_base no registro
            base_sugerida = str(reg.get("_cod_base", "")).strip().upper()
            is_obd_completo = bool(re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', cod_val))
            if not base_sugerida or base_sugerida == "-":
                if (len(cod_val) == 6 and all(c in '0123456789ABCDEF' for c in cod_val)) or is_obd_completo:
                    base_sugerida = cod_val
                elif len(cod_val) >= 4:
                    base_sugerida = cod_val[:-2] if len(cod_val) > 5 else cod_val
                else:
                    base_sugerida = cod_val

            # Extração de tokens de FMI (suporta múltiplos como "12 / 0C", "04 / 06", "12 0C")
            fmi_raw = fmi_val or reg.get("_fmi") or ""
            fmi_tokens = []
            if fmi_raw and str(fmi_raw).strip() not in ["-", "", "None", "null"]:
                partes_fmi = [p.strip() for p in re.split(r'[/,]+|\s+(?=[0-9A-Fa-f]{2}\b)', str(fmi_raw)) if p.strip()]
                for p in partes_fmi:
                    p_clean = re.sub(r'^(?:FMI|fmi|HEX|hex|DEC|dec)[:\s#]*', '', p).strip()
                    if p_clean and p_clean not in ["-", ""]:
                        fmi_tokens.append(p_clean.zfill(2).upper())
            fmi_tokens = list(dict.fromkeys(fmi_tokens))

            # Decisão inteligente para concatenar ou não FMI / Status
            is_hex_6 = len(cod_val) == 6 and all(c in '0123456789ABCDEF' for c in cod_val)

            desc_val = AgenteOrganizador.sintetizar_descricao_modo_completo(causa_val, loc_val, desc_val)

            novo_reg["Localização da Avaria"] = loc_val or "-"
            novo_reg["Causa"] = causa_val or "-"
            novo_reg["Descrição"] = desc_val

            # Caso UDS/ISO 14229 ou OBD-UDS (DTC completo de 6 dígitos ou código OBD P/C/B/U de 5 a 7 caracteres)
            if (cod_val in gt_set and len(cod_val) >= 5) or is_hex_6 or is_obd_completo:
                novo_reg["Código"] = cod_val
                novo_reg["_cod_base"] = cod_val
                novo_reg["_fmi"] = "-"
                resultado.append(novo_reg)
                continue

            # Caso DTC de 8 dígitos onde o status operacional de 2 dígitos foi concatenado indevidamente
            if len(cod_val) == 8 and all(c in '0123456789ABCDEF' for c in cod_val):
                base_6 = cod_val[:-2]
                sufixo_2 = cod_val[-2:]
                if base_6 in gt_set or sufixo_2 in ["02", "00", "2F", "60"]:
                    novo_reg["Código"] = base_6
                    novo_reg["_cod_base"] = base_6
                    novo_reg["_fmi"] = "-"
                    resultado.append(novo_reg)
                    continue

            # Caso SAE J1939 (SPN de 2 a 5 dígitos + FMI)
            base_cod = base_sugerida
            if len(cod_val) > 2 and len(cod_val) != len(base_cod) and cod_val.startswith(base_cod):
                # O cod_val já veio concatenado com o FMI
                sufixo_fmi = cod_val[len(base_cod):]
                if sufixo_fmi and not fmi_tokens:
                    fmi_tokens = [sufixo_fmi.zfill(2).upper()]

            # Se temos múltiplos FMIs (ex: "12 / 0C" ou "04 / 06"): desmembra em linhas separadas
            if len(fmi_tokens) > 1 and len(base_cod) <= 5:
                for ft in fmi_tokens:
                    sub_item = novo_reg.copy()
                    sub_item["_cod_base"] = base_cod
                    sub_item["_fmi"] = ft
                    sub_item["Código"] = f"{base_cod}{ft}"
                    resultado.append(sub_item)
                continue

            # Se o código contém barra (ex: "1003112 / 100310C" ou "1003112 / 0C")
            if "/" in cod_val or "\n" in cod_val:
                cods_split = [c.strip() for c in re.split(r'[/,\n]+', cod_val) if c.strip()]
                for c_item in cods_split:
                    sub_item = novo_reg.copy()
                    c_clean = re.sub(r'[\s:_-]+', '', c_item).upper()
                    if len(c_clean) <= 2 and base_cod:
                        ft = c_clean.zfill(2).upper()
                        sub_item["_fmi"] = ft
                        sub_item["Código"] = f"{base_cod}{ft}"
                    else:
                        sub_item["Código"] = c_clean
                        if len(c_clean) > 2 and len(c_clean) <= 7:
                            sub_item["_cod_base"] = c_clean[:-2]
                            sub_item["_fmi"] = c_clean[-2:]
                        else:
                            sub_item["_cod_base"] = c_clean
                            sub_item["_fmi"] = "-"
                    resultado.append(sub_item)
                continue

            # Caso FMI único normal
            if fmi_tokens:
                ft = fmi_tokens[0]
                cod_final = f"{base_cod}{ft}" if not cod_val.endswith(ft) else cod_val
                novo_reg["Código"] = cod_final
                novo_reg["_cod_base"] = base_cod
                novo_reg["_fmi"] = ft
            else:
                novo_reg["Código"] = cod_val
                novo_reg["_cod_base"] = base_cod
                novo_reg["_fmi"] = "-"

            resultado.append(novo_reg)

        # Expansão automática de equivalência DEC / HEX quando o manual possui colunas [DEC] e [HEX]
        # Aplica-se exclusivamente aos FMIs com letras hexadecimais (0A a 0F <-> 10 a 15)
        MAPA_HEX_PARA_DEC = {
            "0A": "10",
            "0B": "11",
            "0C": "12",
            "0D": "13",
            "0E": "14",
            "0F": "15"
        }
        MAPA_DEC_PARA_HEX = {v: k for k, v in MAPA_HEX_PARA_DEC.items()}

        tem_dec_hex = False
        if texto_fonte:
            tf_up = texto_fonte.upper()
            if ("[DEC]" in tf_up and "[HEX]" in tf_up) or ("FMI [DEC]" in tf_up) or ("FMI DEC" in tf_up) or ("FMI[DEC]" in tf_up):
                tem_dec_hex = True

        if tem_dec_hex:
            codigos_presentes = {r.get("Código", "") for r in resultado}
            novos_equivalentes = []
            for r in resultado:
                fmi = str(r.get("_fmi", "")).strip().upper()
                base = str(r.get("_cod_base", "")).strip()
                if len(base) <= 5 and fmi:
                    if fmi in MAPA_HEX_PARA_DEC:
                        dec_equiv = MAPA_HEX_PARA_DEC[fmi]
                        cod_dec = f"{base}{dec_equiv}"
                        if cod_dec not in codigos_presentes:
                            sub = r.copy()
                            sub["Código"] = cod_dec
                            sub["_fmi"] = dec_equiv
                            novos_equivalentes.append(sub)
                            codigos_presentes.add(cod_dec)
                    elif fmi in MAPA_DEC_PARA_HEX:
                        hex_equiv = MAPA_DEC_PARA_HEX[fmi]
                        cod_hex = f"{base}{hex_equiv}"
                        if cod_hex not in codigos_presentes:
                            sub = r.copy()
                            sub["Código"] = cod_hex
                            sub["_fmi"] = hex_equiv
                            novos_equivalentes.append(sub)
                            codigos_presentes.add(cod_hex)
            resultado.extend(novos_equivalentes)

        for r in resultado:
            for k in ["Localização da Avaria", "Causa", "Descrição"]:
                v = r.get(k, "")
                if isinstance(v, str):
                    v_mod = v
                    v_mod = re.sub(r'\b(?:fio terra|cabo terra)\b', 'fio de aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ao terra / massa|ao terra|ao massa|à terra|à massa|com terra|com a terra)\b', 'ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:curto-circuito ao terra|curto ao terra|curto à terra|curto com terra|curto para terra)\b', 'curto-circuito ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ligação ao terra|ligação terra|conexão terra|ponto de terra)\b', 'conexão de aterramento', v_mod, flags=re.I)
                    r[k] = v_mod

            # Interpretação técnica de fórmulas na coluna Causa
            causa_val = r.get("Causa", "")
            fmi_val = r.get("_fmi", "")
            if causa_val and causa_val != "-":
                r["Causa"] = AgenteAuditor.interpretar_fmi_formula(fmi_val, causa_val)

        # Divisão inteligente de descrições múltiplas por FMI (correspondência 1-para-1)
        from collections import defaultdict
        grupos_base = defaultdict(list)
        for r in resultado:
            cod_str = str(r.get("Código", "")).strip()
            base = r.get("_cod_base") or (cod_str[:-2] if len(cod_str) >= 3 else cod_str)
            grupos_base[base].append(r)

        for base, itens in grupos_base.items():
            if len(itens) <= 1:
                continue

            descs = [it.get("Descrição", "").strip() for it in itens]
            primeira_desc = descs[0]

            # Não divide se for a descrição padrão SIMPLO
            if primeira_desc.startswith("Possíveis causas incluem"):
                continue

            # Se todos os itens do grupo têm a mesma descrição múltipla repetida
            if all(d == primeira_desc for d in descs) or len(set(descs)) == 1:
                partes = []
                # 1. Se tem quebras de linha
                if "\n" in primeira_desc:
                    linhas = [re.sub(r'^[•\-\*\d\.\)\s]+', '', l).strip(' ;-\n') for l in primeira_desc.split('\n')]
                    linhas = [l for l in linhas if l]
                    if len(linhas) > 1:
                        partes = linhas
                # 2. Se tem ponto e vírgula
                elif ";" in primeira_desc:
                    sub_p = [re.sub(r'^[•\-\*\d\.\)\s]+', '', p).strip(' ;-\n') for p in primeira_desc.split(';')]
                    sub_p = [p for p in sub_p if p]
                    if len(sub_p) > 1:
                        partes = sub_p
                # 3. Se tem hífens separando tópicos " - item1 - item2"
                elif " - " in primeira_desc or primeira_desc.startswith("-"):
                    sub_p = [re.sub(r'^[•\-\*\d\.\)\s]+', '', p).strip(' ;-\n') for p in re.split(r'(?:^|\s+)[\-•\*]\s+', primeira_desc)]
                    sub_p = [p for p in sub_p if p]
                    if len(sub_p) > 1:
                        partes = sub_p

                if len(partes) > 1:
                    for idx, it in enumerate(itens):
                        if idx < len(partes):
                            it["Descrição"] = partes[idx]

        # Remove duplicatas exatas de Código (ex: quando OCR corrigiu dois tokens para o mesmo código)
        dedup = []
        codigos_vistos = set()
        for r in resultado:
            chave_dedup = (
                str(r.get("Código", "")).strip().upper(),
                str(r.get("Localização da Avaria", "")).strip().lower(),
                str(r.get("Causa", "")).strip().lower()
            )
            cod_unico = str(r.get("Código", "")).strip().upper()
            # Se for código OBD/UDS de 6-7 caracteres que faz parte do Ground Truth, garante unicidade por código
            if cod_unico in gt_set and len(cod_unico) >= 6:
                if cod_unico in codigos_vistos:
                    continue
                codigos_vistos.add(cod_unico)
            elif chave_dedup in codigos_vistos:
                continue
            else:
                codigos_vistos.add(chave_dedup)
            dedup.append(r)

        # Ordenação exclusiva do Buscar Código: cada código e seus FMIs um abaixo do outro
        resultado = AgenteOrganizador.organizar_buscar_codigo(dedup)

        return resultado

    def estruturar_e_traduzir(self, texto_bruto: str, colunas_desejadas: List[str], contexto_segmento: str = "Linha Pesada e Agrícola", modulo_alvo: str = None, texto_nativo: str = None, codigos_ground_truth: List[str] = None) -> List[Dict[str, Any]]:
        """
        Pipeline Multi-Agente Especialista:
        - Agente 1 (Extrator): Extrai códigos originais em inglês, preservando integridade, zeros à esquerda e isolando componentes.
        - Agente 2 (Tradutor): Aplica o glossário técnico e traduz para PT-BR no estilo direto de concessionária.
        - Agente 3 (Auditor): Audita e assegura fidelidade absoluta dos FMIs ao documento original, eliminando alucinações e conversões indevidas.
        """
        if not self.client:
            raise ValueError(
                "Chave da API do Gemini (GEMINI_API_KEY) não configurada.\n"
                "Edite o arquivo .env e insira sua chave da API do Google AI Studio."
            )

        if not texto_bruto.strip():
            return []

        chave_cache = self._gerar_chave_cache("padrao", texto_bruto, colunas_desejadas, modulo_alvo, codigos_ground_truth)
        em_cache = self._ler_cache_disco(chave_cache)
        if em_cache:
            return self._aplicar_colunas_codigo_secundario(em_cache, colunas_desejadas, texto_bruto, codigos_ground_truth)

        # Se a tabela extraída já estiver em Português e estruturada como 'CODIGO | DESCRICAO',
        # estrutura diretamente preservando 100% dos códigos e textos originais sem risco de alteração por LLM
        registros_ptbr = self._tentar_extracao_ptbr_direta(
            texto_bruto=texto_bruto,
            colunas_desejadas=colunas_desejadas,
            modulo_alvo=modulo_alvo,
            texto_nativo=texto_nativo,
            codigos_ground_truth=codigos_ground_truth
        )
        if registros_ptbr:
            registros_ptbr = self._aplicar_colunas_codigo_secundario(registros_ptbr, colunas_desejadas, texto_bruto, codigos_ground_truth)
            self._salvar_cache_disco(chave_cache, registros_ptbr)
            return registros_ptbr

        try:
            # 1. Agente Extrator Estrutural
            registros_originais = self.agente_extrator.extrair(
                texto_bruto=texto_bruto,
                modulo_alvo=modulo_alvo,
                contexto=contexto_segmento,
                codigos_ground_truth=codigos_ground_truth
            )

            if not registros_originais:
                import re
                tem_coluna_modulo = bool(re.search(
                    r'\b(?:ECU\s*Name|Module\s*Name|Control\s*Unit|Central\s*Eletr[oô]nica|Nome\s*do\s*M[oó]dulo)\b',
                    texto_bruto,
                    re.I
                ))
                if modulo_alvo and tem_coluna_modulo:
                    return []
                res_dir = self._estruturar_e_traduzir_direto(texto_bruto, colunas_desejadas, contexto_segmento, modulo_alvo)
                res_dir = self._aplicar_colunas_codigo_secundario(res_dir, colunas_desejadas, texto_bruto, codigos_ground_truth)
                self._salvar_cache_disco(chave_cache, res_dir)
                return res_dir

            # 2. Agente Lexicógrafo & Tradutor Técnico
            glossario_txt = self._obter_texto_glossario()
            siglas_txt = self._obter_texto_siglas()
            registros_traduzidos = self.agente_tradutor.traduzir(
                registros_originais=registros_originais,
                colunas_desejadas=colunas_desejadas,
                glossario_texto=glossario_txt,
                siglas_texto=siglas_txt,
                contexto=contexto_segmento
            )

            if not registros_traduzidos:
                registros_traduzidos = registros_originais

            # 3. Agente Auditor de Qualidade e Regras (Auditoria Minuciosa de Grounding e Reconciliação Alfanumérica)
            registros_auditados, relatorio = self.agente_auditor.auditar(
                registros=registros_traduzidos,
                colunas_desejadas=colunas_desejadas,
                registros_originais=registros_originais,
                texto_bruto=texto_bruto,
                texto_nativo=texto_nativo,
                codigos_ground_truth=codigos_ground_truth
            )

            # Higienização complementar do módulo e zeros
            res_hig = self._higienizar_registros(registros_auditados, colunas_desejadas, modulo_alvo=modulo_alvo)
            res_hig = self._aplicar_colunas_codigo_secundario(res_hig, colunas_desejadas, texto_bruto, codigos_ground_truth)
            self._salvar_cache_disco(chave_cache, res_hig)
            return res_hig

        except Exception as e:
            print(f"[PIPELINE MULTI-AGENTE] Recorrendo ao modo direto de segurança: {e}")
            res_fb = self._estruturar_e_traduzir_direto(texto_bruto, colunas_desejadas, contexto_segmento, modulo_alvo)
            res_fb = self._aplicar_colunas_codigo_secundario(res_fb, colunas_desejadas, texto_bruto, codigos_ground_truth)
            self._salvar_cache_disco(chave_cache, res_fb)
            return res_fb

    def _estruturar_e_traduzir_direto(self, texto_bruto: str, colunas_desejadas: List[str], contexto_segmento: str = "Linha Pesada e Agrícola", modulo_alvo: str = None) -> List[Dict[str, Any]]:
        """
        Método de extração e tradução direta (fallback de segurança).
        """
        if not self.client:
            raise ValueError(
                "Chave da API do Gemini (GEMINI_API_KEY) não configurada.\n"
                "Edite o arquivo .env e insira sua chave da API do Google AI Studio."
            )

        if not texto_bruto.strip():
            return []

        colunas_str = ", ".join([f'"{col}"' for col in colunas_desejadas])

        import re
        tem_coluna_modulo = bool(re.search(
            r'\b(?:ECU\s*Name|Module\s*Name|Control\s*Unit|Central\s*Eletr[oô]nica|Nome\s*do\s*M[oó]dulo)\b',
            texto_bruto,
            re.I
        ))
        alvo_presente = False
        if modulo_alvo:
            partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', modulo_alvo.strip()) if x]
            if partes_m:
                pat_m = re.compile(r'\b' + r'[\s_-]*'.join(partes_m) + r'\b', re.I)
                alvo_presente = bool(pat_m.search(texto_bruto))

        aplicar_filtro_estrito = bool(modulo_alvo and (tem_coluna_modulo or alvo_presente))

        if aplicar_filtro_estrito:
            instrucao_modulo = f"""
FILTRO DE MÓDULO SOLICITADO: '{modulo_alvo}'.
- Extraia códigos de falha pertencentes ao módulo '{modulo_alvo}' (ou que estejam sob a seção/continuação de '{modulo_alvo}').
- Se o documento possuir uma coluna explícita de módulo (ex: 'ECU Name') identificando códigos de outros módulos diferentes de '{modulo_alvo}', ignore os de outros módulos.
- Se a tabela NÃO possuir coluna explícita de módulo por linha, extraia TODOS os códigos da página! NUNCA descarte códigos apenas porque a descrição menciona siglas como ECU, BCM, ESP, EPS, ABS, CAN, etc.
- Preserve todos os códigos originais mantendo estritamente os zeros à esquerda quando existirem."""
        elif modulo_alvo:
            instrucao_modulo = f"""
MÓDULO DE DESTINO: '{modulo_alvo}'.
- Esta página é uma tabela geral de códigos de falha (sem coluna restritiva de módulo). Extraia ABSOLUTAMENTE TODOS os códigos de falha presentes nesta página!
- NUNCA descarte ou ignore códigos porque a descrição menciona siglas como ECU, BCM, ESP, EPS, ABS, LKA, CAN, etc."""
        else:
            instrucao_modulo = ""

        prompt_sistema = f"""Você é um engenheiro e especialista sênior em diagnóstico automotivo, veículos pesados (caminhões/ônibus) e maquinário agrícola.
Sua tarefa é analisar o texto bruto extraído de um manual de códigos de falhas ({contexto_segmento}), extrair cada código e gerar uma tabela técnica estruturada com dados traduzidos para Português do Brasil (PT-BR).{instrucao_modulo}

REGRAS OBRIGATÓRIAS DE EXTRAÇÃO E TRADUÇÃO TÉCNICA:
1. SEPARAÇÃO ESTRITA ENTRE MÓDULO E CÓDIGO:
   - O nome do módulo/central eletrônica (ex: "EM_SH", "VCM", "ECU", "EDC", "ADM", "BCM") NUNCA deve ser misturado ou juntado com o código da falha!

   - Na coluna de código ("SPN", "PSN", "Código", "DTC", "Código/SPN", "SPN [Hex]"), coloque ESTRITAMENTE o código puro (ex: "047228", "427C9", "47221", "047221").
   - NUNCA retorne "EM_SH 047228" ou "VCM 427C9" na coluna de código. O nome do módulo ("EM_SH", "VCM") pertence exclusivamente à coluna de Módulo/ECU. Se não houver coluna de módulo solicitada, apenas descarte o nome do módulo e mantenha o código puro.
   - PRESERVE ZEROS À ESQUERDA no código (ex: mantenha "047228", nunca remova o zero).

2. FIDELIDADE ABSOLUTA AOS NÚMEROS DE FMI (PROIBIÇÃO DE INVENTAR OU CONVERTER):
   - O campo "fmi" deve ser 100% FIEL ao que está impresso no manual original para aquele código:
     * Extraia EXATAMENTE o número de FMI que consta no documento (ex: se o manual tem "02", extraia "02"; se tem "0B", extraia "0B"; se tem "0C", extraia "0C"; se tem "12", extraia "12").
     * JAMAIS invente ou converta números de FMI (NUNCA adicione "11 / 0B" se o manual só tem "0B"; NUNCA crie "12 / 0C" se o manual só tem "0C"; NUNCA crie "18 / 12" se o manual só tem "12").
     * Use barra "/" APENAS se o manual trouxer expressamente múltiplos FMIs (ex: "04 / 06") ou duas colunas distintas [DEC] e [HEX] preenchidas no documento original (ex: "12 / 0C").
     * PROIBIÇÃO TOTAL DE ADICIONAR FMIs QUE NÃO EXISTEM: Extraia estritamente os FMIs do manual. NUNCA adicione FMIs extras da sua memória ou norma SAE J1939.
   - Para permitir a mesclagem vertical correta no Excel: se um código possuir múltiplos FMIs com descrições diferentes no manual, gere um objeto JSON separado para cada FMI que realmente constar no documento.

3. EXTRAÇÃO E TRADUÇÃO DO COMPONENTE RESPONSÁVEL (NUNCA DEIXE COMPONENTE VAZIO SE HOUVER INFORMAÇÃO):
   - A coluna de Componente ("Componente Culpa", "Fault Component", "Componente") DEVE CONTER O NOME DO COMPONENTE, VÁLVULA, SENSOR, PINO, SAÍDA OU BARRAMENTO TRADUZIDO!
   - Exemplo: "Válvula de engate da embreagem menor (Small eng. clutch valve)", "Válvula de desengate da embreagem maior (Large diseng. clutch valve)", "Saída ADVP", "Controladores 1 e 2 (Falha na ECU)", "Barramento CAN (EBC1)", "Sinal de ABS ativo (EBC1)", "Sensor de pressão", "Sensor de posição da embreagem".
   - Quando o texto original do manual trouxer a descrição de falha junto com o componente (ex: "Short circuit to positive at output stage to small eng. clutch valve" ou "Mechanical failure of large engagement clutch valve" ou "Error on 'ABS active' signal (EBC1)"):
     * EXTRAIA O COMPONENTE E COLOQUE-O NA COLUNA DE COMPONENTE!
     * JAMAIS deixe a coluna de componente vazia ou com "-" se o texto do manual mencionar qualquer válvula, sensor, saída, sinal, barramento ou componente!
   - Não inclua descrições longas de chaveamento ("Low Side switch: ...") no nome do componente.

4. PADRÃO DE ENGENHARIA E OFICINAS AUTORIZADAS (PORTUGUÊS DO BRASIL):
   - Redija as descrições técnicas de forma limpa, direta, fluida e concisa (como nos manuais e scanners oficiais de montadoras como Iveco, Scania, Volvo e John Deere).
   - Use ESTILO TELEGRÁFICO E OBJETIVO DE SCANNER: prefira ordem direta (ex: "Tensão de sinal acima do limite máximo", "Sensor de pressão com falha de plausibilidade"), evitando orações passivas longas ou prolixas.
   - Use terminologia elétrica padronizada: utilize SEMPRE "aterramento" em vez de "terra".
   - Evite repetições desnecessárias ou prolixas.

5. GLOSSÁRIO TÉCNICO RIGOROSO (PRIORIZE ESTES TERMOS PADRONIZADOS):
{self._obter_texto_glossario()}

6. Mantenha siglas internacionais consagradas sem traduzir: {self._obter_texto_siglas()}, etc.
7. Para cada código de falha encontrado no texto bruto, gere um objeto JSON contendo EXATAMENTE estas chaves/colunas:
   [{colunas_str}]
8. NUNCA deixe colunas de descrição ou modo de falha (ex: 'FMI Description', 'Descrição') vazias ou com '-' se o manual fornecer qualquer informação descritiva da falha ou se o FMI for conhecido! Se a coluna 'FMI Description' não tiver texto próprio no manual, aproveite a descrição do defeito na coluna 'Description'. Se ambas estiverem vazias para aquele FMI, utilize a definição técnica padrão da norma SAE J1939 correspondente ao FMI.
9. RETORNE EXCLUSIVAMENTE UM ARRAY JSON VÁLIDO. Não inclua comentários, textos adicionais ou marcações markdown que não sejam o bloco json."""

        filtro_msg = f"do módulo '{modulo_alvo}'" if aplicar_filtro_estrito else "de toda a página"
        prompt_usuario = f"""Texto extraído da página do manual:
---
{texto_bruto}
---

Extraia rigorosamente TODOS os códigos de falha {filtro_msg} presentes no texto, gerando um registro estruturado para cada FMI com as colunas:
[{colunas_str}]"""

        ultimo_erro = None
        for modelo in self.modelos_candidatos:
            try:
                resposta = self.client.models.generate_content(
                    model=modelo,
                    contents=[prompt_sistema, prompt_usuario],
                    config={
                        "temperature": 0.1,  # Baixa temperatura para precisão cirúrgica
                        "response_mime_type": "application/json"
                    }
                )

                dados = reparar_e_carregar_json(resposta.text)
                
                # Atualiza modelo bem-sucedido para os próximos chamados
                self.model_name = modelo
                
                # Garante que seja uma lista de dicionários
                lista_registros = []
                if isinstance(dados, dict):
                    for v in dados.values():
                        if isinstance(v, list):
                            lista_registros = v
                            break
                    if not lista_registros:
                        lista_registros = [dados]
                elif isinstance(dados, list):
                    lista_registros = dados
                    
                return self._higienizar_registros(lista_registros, colunas_desejadas)

            except Exception as e:
                ultimo_erro = e
                err_low = str(e).lower()
                # Se for erro de modelo não encontrado ou timeout/indisponibilidade, tenta o próximo modelo da lista
                if any(x in err_low for x in ["404", "not_found", "not found", "504", "deadline_exceeded", "503", "unavailable", "timeout"]):
                    print(f"[AVISO IA] Indisponibilidade/Timeout no modelo {modelo} ({e}). Tentando próximo modelo...")
                    continue
                if isinstance(e, json.JSONDecodeError) or any(x in err_low for x in ["delimiter", "json", "expecting"]):
                    print(f"[AVISO IA] Formato JSON malformado no modelo {modelo} ({e}). Tentando próximo modelo...")
                    continue
                else:
                    print(f"[ERRO IA] Falha no modelo {modelo}: {e}")
                    raise e

        print(f"[ERRO IA] Nenhum modelo disponível respondeu com sucesso: {ultimo_erro}")
        raise ultimo_erro

    def _higienizar_registros(self, registros: List[Dict[str, Any]], colunas_desejadas: List[str], modulo_alvo: str = None) -> List[Dict[str, Any]]:
        """
        Higieniza os registros garantindo:
        1. Que nomes de módulos (como EM_SH, VCM, ECU) não fiquem colados no código (SPN/DTC).
        2. Que códigos com zeros à esquerda (ex: 047228) permaneçam como strings.
        3. Se o usuário pediu coluna de Módulo/ECU, preenche com o módulo isolado.
        4. Se modulo_alvo foi especificado, descarta registros pertencentes a outros módulos.
        """
        import re
        
        # Identifica colunas de código e módulo solicitadas
        col_codigo = None
        col_modulo = None
        
        for col in colunas_desejadas:
            c_low = col.lower()
            if any(k in c_low for k in ["spn", "psn", "código", "codigo", "dtc"]):
                col_codigo = col
            if any(k in c_low for k in ["módulo", "modulo", "ecu"]):
                col_modulo = col
                
        modulos_conhecidos = ["EM_SH", "VCM", "ECU", "EDC", "ADM", "BCM", "EBS", "ABS", "TCU", "IC", "DD", "PDM"]
        regex_modulo = re.compile(r"^([A-Za-z0-9_]{2,8})\s+([0-9A-Fa-f]{3,8})$")

        resultado = []
        for reg in registros:
            if not isinstance(reg, dict):
                continue
                
            novo_reg = reg.copy()
            
            if col_codigo and col_codigo in novo_reg:
                val_cod = str(novo_reg[col_codigo]).strip()
                
                # Checa se o código veio no formato "MÓDULO CÓDIGO" (ex: "EM_SH 047228" ou "VCM 427C9")
                match = regex_modulo.match(val_cod)
                if match:
                    modulo_extraido = match.group(1).strip()
                    codigo_puro = match.group(2).strip()
                    
                    novo_reg[col_codigo] = codigo_puro
                    if col_modulo and (not novo_reg.get(col_modulo) or novo_reg[col_modulo] == "-"):
                        novo_reg[col_modulo] = modulo_extraido
                else:
                    # Remove prefixos conhecidos se houver
                    for mod in modulos_conhecidos:
                        if val_cod.upper().startswith(mod + " "):
                            codigo_puro = val_cod[len(mod):].strip()
                            novo_reg[col_codigo] = codigo_puro
                            if col_modulo and (not novo_reg.get(col_modulo) or novo_reg[col_modulo] == "-"):
                                novo_reg[col_modulo] = mod
                            break

            # Filtro estrito de módulo alvo
            if modulo_alvo:
                mod_alvo_c = re.sub(r'[\s_-]+', '', modulo_alvo).upper()
                # Verifica no campo do módulo (se existir na coluna ou nos dados)
                val_m = None
                if col_modulo and col_modulo in novo_reg:
                    val_m = str(novo_reg[col_modulo]).strip()
                elif "module" in novo_reg:
                    val_m = str(novo_reg["module"]).strip()
                elif "modulo" in novo_reg:
                    val_m = str(novo_reg["modulo"]).strip()

                if val_m and val_m not in ["-", "None", ""]:
                    val_m_c = re.sub(r'[\s_-]+', '', val_m).upper()
                    mod_raiz = re.split(r'[\s_-]+', modulo_alvo.strip())[0].upper()
                    val_m_raiz = re.split(r'[\s_-]+', val_m.strip())[0].upper()
                    if mod_raiz and val_m_raiz and mod_raiz == val_m_raiz and len(mod_raiz) >= 2:
                        pass
                    elif val_m_c != mod_alvo_c and mod_alvo_c not in val_m_c and val_m_c not in mod_alvo_c:
                        continue
                
                # Se a coluna de módulo estiver vazia ou com '-', preenche com modulo_alvo
                if col_modulo and (not novo_reg.get(col_modulo) or novo_reg[col_modulo] == "-"):
                    novo_reg[col_modulo] = modulo_alvo

            # 2. Limpeza estrita e resgate da coluna de Componente
            col_comp = next((c for c in colunas_desejadas if any(k in c.lower() for k in ["component", "componente", "culpa"])), None)
            if col_comp and col_comp in novo_reg:
                val_comp = str(novo_reg[col_comp]).strip()

                # Resgate se estiver vazio ou '-'
                if val_comp in ["-", "", "None", "null", "nan"]:
                    orig_d = str(novo_reg.get("description", novo_reg.get("Descrição", ""))).strip()
                    if orig_d and orig_d not in ["-", "None"]:
                        comp_resg = AgenteAuditor.extrair_e_traduzir_componente(orig_d)
                        if comp_resg and comp_resg != "-":
                            val_comp = comp_resg

                # Corta textos de condições de teste elétrico se tiverem sido mesclados no componente
                padroes_corte = [
                    "Low Side switch", "High Side switch", "low side switch", "high side switch",
                    "S.C. to Vbat", "S.C. to Gnd", "Open Load detected", "Open Circuit and Short",
                    "comutador Low Side", "comutador High Side", "(comutador Low Side)", "(comutador High Side)",
                    "- S.C.", "- Curto-circuito"
                ]
                
                # Se tiver quebra de linha no componente, pega apenas a primeira linha (nome do componente)
                if "\n" in val_comp:
                    linhas_c = val_comp.split("\n")
                    val_comp = linhas_c[0].strip()

                for padrao in padroes_corte:
                    if padrao in val_comp:
                        idx_corte = val_comp.find(padrao)
                        if idx_corte > 0:
                            val_comp = val_comp[:idx_corte].strip(" -:;,()")
                            
                novo_reg[col_comp] = val_comp

            # 3. Desmembra caso FMI e Descrição tenham vindo com quebra de linha (\n)
            col_fmi_desc = next((c for c in colunas_desejadas if "fmi" in c.lower() and any(k in c.lower() for k in ["desc", "descrição", "descricao"])), None)
            col_fmi = next((c for c in colunas_desejadas if "fmi" in c.lower() and c != col_fmi_desc), None)
            col_desc = next((c for c in colunas_desejadas if any(k in c.lower() for k in ["desc", "descrição", "descricao", "localização", "localizacao", "condição", "condicao", "defeito", "causa"]) and c != col_fmi_desc), None)
            col_alvo_desc = col_fmi_desc or col_desc

            def formatar_fmi_campo(val_f):
                return AgenteAuditor.formatar_fmi_hex_2digitos(val_f)

            if col_fmi and col_fmi in novo_reg:
                val_fmi = str(novo_reg[col_fmi]).strip()
                
                # Se veio com quebra de linha ou delimitador múltiplo
                if col_alvo_desc and col_alvo_desc in novo_reg:
                    val_desc = str(novo_reg[col_alvo_desc]).strip()
                    if "\n" in val_desc:
                        descs = [d.strip() for d in val_desc.split("\n") if d.strip()]
                        if "\n" in val_fmi:
                            fmis = [f.strip() for f in val_fmi.split("\n") if f.strip()]
                        else:
                            # Se FMI veio separado por barra mas NÃO é par Dec/Hex único (ex: '03 / 04')
                            atomos_fmi = AgenteAuditor._decompor_fmi_atomos(val_fmi)
                            if len(atomos_fmi) > 1:
                                fmis = atomos_fmi
                            else:
                                fmis = []

                        if len(fmis) == len(descs) and len(fmis) > 1:
                            for f, d in zip(fmis, descs):
                                sub_reg = novo_reg.copy()
                                sub_reg[col_fmi] = formatar_fmi_campo(f)
                                sub_reg[col_alvo_desc] = d
                                resultado.append(sub_reg)
                            continue

                novo_reg[col_fmi] = formatar_fmi_campo(val_fmi)

            resultado.append(novo_reg)
            
        # 4. Padronização terminológica rigorosa (GND / Terra -> Aterramento)
        for r in resultado:
            for k, v in r.items():
                if isinstance(v, str):
                    v_mod = v
                    v_mod = re.sub(r'\b(?:fio terra|cabo terra)\b', 'fio de aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ao terra / massa|ao terra|ao massa|à terra|à massa|com terra|com a terra)\b', 'ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:curto-circuito ao terra|curto ao terra|curto à terra|curto com terra|curto para terra)\b', 'curto-circuito ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ligação ao terra|ligação terra|conexão terra|ponto de terra)\b', 'conexão de aterramento', v_mod, flags=re.I)
                    r[k] = v_mod

        # 5. Consolidação de FMIs diferentes com a mesma descrição (ex: 04 e 06 -> 04 / 06)
        col_codigo = next((c for c in colunas_desejadas if any(k in c.lower() for k in ["spn", "psn", "código", "codigo", "dtc"])), None)
        col_comp = next((c for c in colunas_desejadas if any(k in c.lower() for k in ["component", "componente", "culpa", "fault"])), None)
        resultado = self.agente_auditor._consolidar_fmis_mesma_descricao(
            registros=resultado,
            col_codigo=col_codigo,
            col_fmi=col_fmi,
            col_desc=col_desc,
            col_componente=col_comp
        )

        # 6. Ordenação crescente dos códigos, FMIs e descrições
        from core.gerador_excel import GeradorExcel
        def chave_ordenacao_ia(item):
            val_cod = item.get(col_codigo) if col_codigo else ""
            val_fmi = item.get(col_fmi) if col_fmi else ""
            val_desc = str(item.get(col_desc, "")).strip().lower() if col_desc else ""
            val_comp = str(item.get(col_comp, "")).strip().lower() if col_comp else ""
            return (
                GeradorExcel.extrair_chave_codigo(val_cod),
                GeradorExcel.extrair_valores_fmi(val_fmi),
                val_desc,
                val_comp
            )

        resultado = sorted(resultado, key=chave_ordenacao_ia)

        return resultado

    @staticmethod
    def _texto_ja_em_portugues(texto: str) -> bool:
        import re
        if not texto:
            return False
        marcadores_pt = re.findall(
            r'\b(?:tens[aã]o|tra[cç][aã]o|inv[aá]lido|diagn[oó]stico|el[eé]trica|velocidade|'
            r'press[aã]o|dire[cç][aã]o|comunica[cç][aã]o|falha|circuito|m[oó]dulo|válvula|valvula|'
            r'dianteira|traseira|esquerda|direita|incoerente|desativada|bateria|freio)\b',
            texto,
            flags=re.I
        )
        return len(marcadores_pt) >= 5

    def _tentar_buscar_codigo_ptbr_direto(
        self,
        texto_bruto: str,
        codigos_ground_truth: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """
        Quando a tabela extraída já está em Português no formato 'CODIGO | COMPONENTE - CAUSA',
        estrutura deterministicamente para o modo Buscar Código com 100% de fidelidade aos códigos originais.
        """
        import re
        if not self._texto_ja_em_portugues(texto_bruto):
            return []

        linhas_validas = []
        gt_set = {c.upper() for c in (codigos_ground_truth or []) if c}
        for l in texto_bruto.splitlines():
            if "|" not in l:
                continue
            partes = [p.strip() for p in l.split("|") if p.strip()]
            if len(partes) < 2:
                continue
            cod = partes[0].strip().rstrip(",.;:").upper()
            if (gt_set and cod in gt_set) or re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', cod):
                linhas_validas.append((cod, " - ".join(partes[1:])))

        if not linhas_validas or (gt_set and len(linhas_validas) < len(gt_set) * 0.85):
            return []

        registros = []
        for cod, desc_full in linhas_validas:
            if " - " in desc_full:
                loc, causa = desc_full.split(" - ", 1)
            else:
                loc, causa = desc_full, desc_full
            loc_str = loc.strip().rstrip(".") or "-"
            causa_str = causa.strip().rstrip(".") or "-"
            desc_sintetizada = AgenteOrganizador.sintetizar_descricao_modo_completo(causa_str, loc_str)
            registros.append({
                "Código": cod,
                "_cod_base": cod,
                "_fmi": "-",
                "Localização da Avaria": loc_str,
                "Causa": causa_str,
                "Descrição": desc_sintetizada
            })
        return registros

    def _tentar_extracao_ptbr_direta(
        self,
        texto_bruto: str,
        colunas_desejadas: List[str],
        modulo_alvo: Optional[str] = None,
        texto_nativo: Optional[str] = None,
        codigos_ground_truth: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """
        Quando a tabela extraída já está em Português no formato 'CODIGO | DESCRICAO',
        constrói os registros diretamente sem risco de omissão ou alteração de códigos pelo LLM.
        """
        import re
        if not self._texto_ja_em_portugues(texto_bruto):
            return []

        linhas_validas = []
        gt_set = {c.upper() for c in (codigos_ground_truth or []) if c}
        for l in texto_bruto.splitlines():
            if "|" not in l:
                continue
            partes = [p.strip() for p in l.split("|") if p.strip()]
            if len(partes) < 2:
                continue
            cod = partes[0].strip().rstrip(",.;:").upper()
            if (gt_set and cod in gt_set) or re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', cod):
                linhas_validas.append((cod, " | ".join(partes[1:])))

        if not linhas_validas or (gt_set and len(linhas_validas) < len(gt_set) * 0.85):
            return []

        registros_base = []
        for cod, desc_full in linhas_validas:
            comp = desc_full.split(" - ", 1)[0].strip() if " - " in desc_full else desc_full
            reg = {}
            for col in colunas_desejadas:
                c_low = col.lower()
                if any(k in c_low for k in ["spn", "psn", "código", "codigo", "dtc", "code"]) and "desc" not in c_low:
                    reg[col] = cod
                elif "fmi" in c_low and "desc" not in c_low:
                    reg[col] = "-"
                elif any(k in c_low for k in ["módulo", "modulo", "ecu"]):
                    reg[col] = modulo_alvo or "-"
                elif any(k in c_low for k in ["component", "componente", "culpa", "localização", "localizacao"]):
                    reg[col] = comp
                else:
                    reg[col] = desc_full
            registros_base.append(reg)

        registros_auditados, _ = self.agente_auditor.auditar(
            registros=registros_base,
            colunas_desejadas=colunas_desejadas,
            registros_originais=[
                {"code": c, "component": d.split(" - ", 1)[0] if " - " in d else d, "description": d, "fmi": "-", "fmi_description": d}
                for c, d in linhas_validas
            ],
            texto_bruto=texto_bruto,
            texto_nativo=texto_nativo,
            codigos_ground_truth=codigos_ground_truth
        )
        return self._higienizar_registros(registros_auditados, colunas_desejadas, modulo_alvo=modulo_alvo)



