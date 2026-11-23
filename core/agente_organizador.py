import re
from typing import List, Dict, Any, Tuple, Optional

class AgenteOrganizador:
    """
    Agente 4: Especialista em Ordenação e Hierarquia de Dados Técnicos (Pure Python).
    
    Missão:
    Organizar os códigos de falhas e FMIs em ordem numérica crescente através
    de Chaves Numéricas Virtuais (Sort Keys), garantindo 100% de imutabilidade
    dos dados originais (nenhum zero à esquerda, caractere ou texto é alterado).
    """

    @staticmethod
    def extrair_chave_codigo(cod_str: Any) -> Tuple:
        """
        Gera uma chave virtual matemática para o código de falha.
        O valor original NÃO é modificado na tabela, a chave serve apenas para ordenação.
        
        Suporta com máxima precisão:
        1. Códigos Hexadecimais / UDS / CAN (ex: 003D13, 003D1A, 003D20, 0472E1)
        2. Códigos Numéricos com ou sem zeros à esquerda (ex: 047221 e 47222 ficam juntos)
        3. Códigos OBD-II / DTC com letras (ex: P0020 antes de P0100)
        4. Códigos SAE J1939 SPN puros (ex: SPN 91 antes de SPN 108)
        """
        if not cod_str or str(cod_str).strip() in ['-', '', 'None', 'null', 'nan']:
            return (9, 0, '')

        s = str(cod_str).strip().upper()
        # Remove prefixos comuns de manuais para avaliar o valor real
        limpo = re.sub(r'^(?:0X|SPN|DTC|PSN)[\s:#_-]*', '', s, flags=re.I).strip()

        if not limpo:
            return (9, 0, s)

        # 1. Padrão Numérico Decimal Puro (Ex: 42701, 42702 ... 42799, 108)
        # Todos os códigos compostos apenas por números vêm primeiro em ordem decimal crescente
        if limpo.isdigit():
            return (0, int(limpo, 10), len(limpo), s)

        # 2. Padrão Hexadecimal com Letras (Ex: 4270A, 4270B, 4270F, 4271A, 427FF)
        # Códigos com caracteres hexadecimais (A-F) aparecem logo após o término dos numéricos (ex: após 42799)
        if all(c in '0123456789ABCDEF' for c in limpo):
            try:
                valor_hex = int(limpo, 16)
                return (1, valor_hex, len(limpo), s)
            except ValueError:
                pass

        # 3. Padrão OBD-II / DTC tradicional ou hifenizado (Ex: P0020, C0100, B235A-68, C003A-4A, U19B9-83)
        m_obd = re.match(r'^([PCBU])\s*([0-9A-F]+)(?:[-.\s]+(?:FMI\s*)?([0-9A-F]+))?$', limpo)
        if m_obd:
            letra = m_obd.group(1)
            try:
                num = int(m_obd.group(2), 16)
                sub = int(m_obd.group(3), 16) if m_obd.group(3) else -1
                return (2, letra, num, sub, s)
            except ValueError:
                pass

        # 4. Padrão Alfanumérico Natural Genérico (Divide em blocos de texto e números)
        partes = []
        for chunk in re.split(r'(\d+)', limpo):
            if chunk.isdigit():
                partes.append((0, int(chunk)))
            elif chunk:
                partes.append((1, chunk.lower()))
        return (3, partes, len(s), s)

    @staticmethod
    def extrair_chave_fmi(fmi_str: Any) -> List[Tuple]:
        """
        Gera uma chave virtual para o número do FMI (Failure Mode Identifier).
        
        Suporta:
        1. Formato Decimal / Hexadecimal equivalente (ex: "12 / 0C" -> 12)
        2. Múltiplos FMIs unificados (ex: "04 / 06" -> [4, 6])
        3. FMIs simples (ex: "00", "01", "09", "14")
        4. FMIs vazios ou não aplicáveis ao final
        """
        if not fmi_str or str(fmi_str).strip() in ['-', '', 'None', 'null', 'nan']:
            return [(999, 999)]

        s = str(fmi_str).strip()

        # Verifica se é formato de equivalência Decimal / Hexadecimal (ex: "12 / 0C" ou "10 / 0A")
        m_eq = re.match(r'^(\d+)\s*/\s*([0-9A-Fa-f]{2})$', s)
        if m_eq:
            dec_val = int(m_eq.group(1))
            try:
                hex_val = int(m_eq.group(2), 16)
                if dec_val == hex_val:
                    return [(0, dec_val)]
            except ValueError:
                pass

        # Quebra por barras e espaços para FMIs múltiplos (ex: "04 / 06")
        partes = re.split(r'[\s/]+', s)
        valores = []
        for p in partes:
            p_clean = re.sub(r'^(?:FMI|fmi|HEX|hex|DEC|dec)[:\s#]*', '', p).strip()
            if not p_clean or p_clean.upper() in ['-', 'NONE', '']:
                continue
            try:
                if p_clean.isdigit():
                    val = int(p_clean, 10)
                else:
                    val = int(p_clean, 16)
                valores.append((0, val))
            except ValueError:
                valores.append((1, p_clean.lower()))

        if not valores:
            return [(999, 999)]
        return sorted(valores)

    def organizar(self, registros: List[Dict[str, Any]], colunas: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        """
        Organiza a lista de registros de falha na ordem crescente ideal:
        1º Nível: Código de falha (ordem numérica virtual crescente)
        2º Nível: FMI (ordem numérica crescente: 00 -> 01 -> 02 -> ... -> 12/0C)
        3º Nível: Descrição da falha (desempate alfabético)
        4º Nível: Componente
        
        IMPORTANTE: Não modifica nenhum valor interno dos dicionários.
        """
        if not registros:
            return []

        # Determina os nomes das colunas baseado na lista ou nas chaves dos registros
        amostra_chaves = colunas if colunas else list(registros[0].keys())

        col_cod = next((c for c in amostra_chaves if any(k in c.lower() for k in ["spn", "psn", "código", "codigo", "dtc", "code", "cod"])), None)
        col_fmi = next((c for c in amostra_chaves if "fmi" in c.lower() and "desc" not in c.lower()), None)
        col_desc = next((c for c in amostra_chaves if any(k in c.lower() for k in ["desc", "localização", "localizacao", "description"])), None)
        col_comp = next((c for c in amostra_chaves if any(k in c.lower() for k in ["component", "componente", "culpa"])), None)

        def chave_ordenacao(item: Dict[str, Any]):
            val_cod = item.get(col_cod, "") if col_cod else ""
            val_fmi = item.get(col_fmi, "") if col_fmi else ""
            val_desc = str(item.get(col_desc, "")).strip().lower() if col_desc else ""
            val_comp = str(item.get(col_comp, "")).strip().lower() if col_comp else ""

            return (
                self.extrair_chave_codigo(val_cod),
                self.extrair_chave_fmi(val_fmi),
                val_desc,
                val_comp
            )

        return sorted(registros, key=chave_ordenacao)

    @classmethod
    def extrair_chave_buscar_codigo(cls, item: Any) -> Tuple:
        """
        Chave de ordenação EXCLUSIVA para a modalidade 'Buscar Código':
        Garante que cada código base e todos os seus FMIs fiquem agrupados juntos,
        um abaixo do outro em ordem crescente de FMI.
        Exemplo:
          4274103
          4274104
          427410A
          seguido por 427600D, 4276400, 4276401, 427640B, 4276412...
        """
        if isinstance(item, dict):
            cod_str = str(item.get("Código", "")).strip()
            cod_base = item.get("_cod_base")
            fmi_str = item.get("_fmi")
        else:
            cod_str = str(item).strip()
            cod_base = None
            fmi_str = None

        if not cod_base or not fmi_str:
            # Se não fornecido explicitamente, verificar se é código UDS de 6 dígitos hexadecimais ou código OBD hifenizado
            if (len(cod_str) == 6 and all(c in '0123456789ABCDEFabcdef' for c in cod_str)) or re.match(r'^[PCBU][0-9A-Fa-f]{4}(?:-[0-9A-Fa-f]{2}|[0-9A-Fa-f]{0,2})$', cod_str):
                cod_base = cod_str
                fmi_str = "-"
            elif len(cod_str) >= 4:
                cod_base = cod_str[:-2]
                fmi_str = cod_str[-2:]
            else:
                cod_base = cod_str
                fmi_str = ""

        chave_cod = cls.extrair_chave_codigo(cod_base)

        # O FMI no Buscar Código é o sufixo alfanumérico/hexadecimal de 2 caracteres
        if fmi_str in ["-", "", None]:
            chave_fmi = (0, -1, "")
        else:
            try:
                val_fmi = int(fmi_str, 16)
                chave_fmi = (0, val_fmi, fmi_str)
            except ValueError:
                chave_fmi = (1, 999, str(fmi_str))

        loc = str(item.get("Localização da Avaria", "")) if isinstance(item, dict) else ""
        causa = str(item.get("Causa", "")) if isinstance(item, dict) else ""
        return (chave_cod, chave_fmi, loc, causa)

    @classmethod
    def organizar_buscar_codigo(cls, registros: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Organiza os registros do 'Buscar Código' mantendo cada código base
        e todos os seus FMIs agrupados juntos, um abaixo do outro.
        """
        return sorted(registros, key=cls.extrair_chave_buscar_codigo)

    @staticmethod
    def _obter_artigo_definido(texto: str) -> str:
        """
        Retorna o artigo definido em português ('do', 'da', 'dos', 'das' ou 'de')
        com base no gênero e número da palavra inicial do componente.
        """
        t = texto.strip().lower()
        if not t:
            return "do"
        if t.startswith(('do ', 'da ', 'dos ', 'das ', 'de ', 'no ', 'na ', 'nos ', 'nas ')):
            return ""

        primeira_palavra = re.split(r'[\s\-_\(/]+', t)[0]

        femininos = (
            'válvula', 'valvula', 'bomba', 'unidade', 'entrada', 'saída', 'saida',
            'bobina', 'linha', 'sonda', 'alimentação', 'alimentacao', 'comunicação',
            'comunicacao', 'rede', 'chave', 'porta', 'luz', 'lâmpada', 'lampada',
            'bateria', 'buzina', 'direção', 'direcao', 'transmissão', 'transmissao',
            'embreagem', 'vela', 'esteira', 'tração', 'tracao', 'tampa', 'mangueira',
            'ventoinha', 'turbina', 'caixa', 'manopla', 'alavanca', 'pressão', 'pressao'
        )
        femininos_plurais = ('válvulas', 'bombas', 'entradas', 'saídas', 'linhas', 'sondas', 'velas', 'rodas')
        masculinos_plurais = ('sensores', 'atuadores', 'injetores', 'módulos', 'modulos', 'relés', 'reles', 'circuitos', 'terminais', 'chicotes', 'cabos')

        if primeira_palavra in femininos_plurais or (primeira_palavra.endswith('as') and not primeira_palavra.endswith('gas')):
            return "das"
        if primeira_palavra in masculinos_plurais or primeira_palavra.endswith(('os', 'ores', 'eis')):
            return "dos"
        if primeira_palavra in femininos or primeira_palavra.endswith(('ção', 'cao', 'dade', 'agem', 'ina')):
            return "da"
        if primeira_palavra.endswith('a') and not primeira_palavra.endswith(('ma', 'ta')):
            return "da"

        # Padrão masculino para a grande maioria dos termos técnicos (sensor, atuador, módulo, relé, chicote, injetor, etc.)
        return "do"

    @staticmethod
    def _formatar_inicio_frase(texto: str) -> str:
        """Formata o início da descrição com gramática técnica precisa."""
        t = texto.strip().lstrip('.,;: -')
        if not t:
            return ""
        t_low = t.lower()
        if t_low in ['aberto', 'circuito aberto']:
            return 'Circuito aberto detectado'
        if t_low in ['curto ao positivo', 'curto com positivo', 'curto-circuito ao positivo']:
            return 'Curto-circuito ao positivo (Vbatt)'
        if t_low in ['curto ao aterramento', 'curto com massa', 'curto à massa', 'curto-circuito ao aterramento']:
            return 'Curto-circuito ao aterramento (GND)'
        return t[0].upper() + t[1:]

    @staticmethod
    def _determinar_recomendacao_teste(texto_analise: str) -> str:
        """
        Determina a pista de teste prático recomendada de oficina conforme a natureza da avaria.
        """
        t = texto_analise.lower()

        # 1. Curto ao positivo / Vbatt / Sobretensão
        if re.search(r'(?:vbatt|curto.*positivo|curto.*bateria|tensão acima|tensao acima|tensão alta|tensao alta|sobretensão|sobretensao|v\s*>)', t):
            return "Inspecione o chicote elétrico contra esmagamentos ou contato com linhas de alimentação (Vbatt), verifique a integridade dos pinos nos conectores e meça a tensão com multímetro."

        # 2. Curto ao aterramento / GND / Subtensão
        if re.search(r'(?:gnd|curto.*aterramento|curto.*massa|curto.*gnd|tensão abaixo|tensao abaixo|tensão baixa|tensao baixa|subtensão|subtensao|v\s*<)', t):
            return "Verifique o isolamento elétrico da fiação em relação à massa/carroceria e inspecione os pontos de aterramento quanto à oxidação ou mau contato."

        # 3. Circuito aberto / Linha interrompida
        if re.search(r'(?:circuito aberto|aberto\b|interromp|desconect|rompid|resistência alta|resistencia alta|open circuit)', t):
            return "Realize o teste de continuidade e resistência ôhmica na fiação do componente até a central, inspecionando travas de conectores e terminais frouxos."

        # 4. Falha mecânica / Válvula / Emperramento / Travamento / Obstrução / Pressão
        if re.search(r'(?:mecânic|mecanic|emperrad|travad|obstru|vazamento|pressão|pressao|válvula travada|valvula travada|desgaste|curso)', t):
            return "Inspecione o mecanismo físico quanto a travamento mecânico, obstruções, carbonização, vazamentos de pressão ou desgaste de componentes antes da substituição."

        # 5. Rede CAN / Barramento / Linha de Comunicação / Timeout
        if re.search(r'(?:\bcan(?:-[hl])?\b|\brede\s*can\b|barramento|comunicação|comunicacao|timeout|taxa de atualização|taxa de atualizacao|bus off|\bj1939\b|\brede\b)', t):
            return "Meça a resistência de terminação da rede CAN (nominal de 60 Ω com a bateria desconectada) e confira as tensões de operação nas linhas CAN-High (~2,5V a 3,5V) e CAN-Low (~1,5V a 2,5V)."

        # 6. Plausibilidade / Sinal Fora da Faixa / Errático / Intermitente
        if re.search(r'(?:plausibil|fora da faixa|faixa/desempenho|incoerent|errátic|erratic|intermitent|desempenho|faixa incorreta|sinal inválido|sinal invalido|deriva|incorret)', t):
            return "Verifique a alimentação do sensor (tipicamente 5,0 V estabilizados ou 12 V) e a referência de aterramento. Monitore a curva de resposta do sinal via scanner ou osciloscópio durante a operação."

        # 7. Módulo / ECU / Falha Interna / Memória
        if re.search(r'(?:\bmódulo\b|\bmodulo\b|\becu\b|eeprom|\brom\b|\bram\b|software|checksum|microcontrolador|watchdog|falha interna)', t):
            return "Verifique as linhas de alimentação positiva e aterramento da central sob carga, fusíveis e relés principais. Se os sinais periféricos estiverem normais, efetue reset ou reprogramação do módulo."

        # 8. Padrão geral de oficina
        return "Utilize diagramas elétricos e manuais técnicos SIMPLO para rastrear os circuitos do componente, avaliando chicote, conectores e parâmetros em tempo real via scanner antes de substituir peças."

    @classmethod
    def sintetizar_descricao_modo_completo(
        cls,
        causa: Any,
        localizacao: Any,
        desc_atual: Optional[Any] = None
    ) -> str:
        """
        Sintetiza uma descrição técnica rica em 'Modo Completo' para o profissional de oficina:
        - Parte 1 (Identificação Técnica da Falha): Frase técnica conectando a Causa à Localização/Componente.
        - Parte 2 (Recomendação Prática de Teste): Orientação diagnóstica de oficina direcionada
          à natureza elétrica, mecânica ou de rede da avaria.
        """
        PADRAO_GENERICO = (
            "Possíveis causas incluem falhas elétricas ou desgaste mecânico. "
            "Antes de substituir peças, realize testes detalhados e utilize "
            "diagramas elétricos e manuais técnicos SIMPLO para um diagnóstico preciso e confiável."
        )

        # Se já existir uma descrição rica, específica e válida do manual (que não seja o texto genérico repetitivo)
        if desc_atual:
            d_str = str(desc_atual).strip()
            d_lower = d_str.lower()
            if (
                len(d_str) > 20
                and d_str not in ["-", "None", "null", "nan"]
                and not d_lower.startswith("possíveis causas incluem")
                and "antes de substituir peças" not in d_lower
            ):
                return d_str

        # Limpeza e normalização da causa e localização
        c_str = str(causa).strip().rstrip(".;,") if causa is not None else ""
        l_str = str(localizacao).strip().rstrip(".;,") if localizacao is not None else ""

        if c_str in ["-", "", "None", "null", "nan", "none"]:
            c_str = ""
        if l_str in ["-", "", "None", "null", "nan", "none"]:
            l_str = ""

        # Se ambos forem inexistentes, retorna o fallback padrão
        if not c_str and not l_str:
            return PADRAO_GENERICO

        # === PARTE 1: IDENTIFICAÇÃO TÉCNICA DA FALHA ===
        frase_falha = ""
        if c_str and l_str:
            artigo = cls._obter_artigo_definido(l_str)
            l_lower = l_str.lower()
            c_lower = c_str.lower()

            if l_lower.startswith(('do ', 'da ', 'dos ', 'das ', 'de ', 'no ', 'na ', 'nos ', 'nas ')):
                frase_falha = f"{cls._formatar_inicio_frase(c_str)} {l_str}."
            elif l_lower.startswith(('circuito', 'linha', 'barramento', 'chicote', 'rede')):
                prep = "no " if artigo in ["do", "dos"] else "na " if artigo in ["da", "das"] else "em "
                frase_falha = f"{cls._formatar_inicio_frase(c_str)} {prep}{l_str}."
            elif c_lower.startswith(('circuito', 'linha')):
                prep = "no " if artigo == "do" else "na " if artigo == "da" else "nos " if artigo == "dos" else "nas " if artigo == "das" else "em "
                frase_falha = f"{cls._formatar_inicio_frase(c_str)} {prep}{l_str}."
            else:
                prep = f"do {l_str}" if artigo == "do" else f"da {l_str}" if artigo == "da" else f"dos {l_str}" if artigo == "dos" else f"das {l_str}" if artigo == "das" else l_str
                frase_falha = f"{cls._formatar_inicio_frase(c_str)} no circuito {prep}."
        elif l_str and not c_str:
            artigo = cls._obter_artigo_definido(l_str)
            prep = f"do {l_str}" if artigo == "do" else f"da {l_str}" if artigo == "da" else f"dos {l_str}" if artigo == "dos" else f"das {l_str}" if artigo == "das" else l_str
            frase_falha = f"Falha de funcionamento detectada no circuito {prep}."
        elif c_str and not l_str:
            frase_falha = f"{cls._formatar_inicio_frase(c_str)} identificado no circuito elétrico do sistema."

        # === PARTE 2: RECOMENDAÇÃO PRÁTICA DE TESTE (DIAGNÓSTICO) ===
        texto_analise = f"{c_str} {l_str}".lower()
        recomendacao_teste = cls._determinar_recomendacao_teste(texto_analise)

        return f"{frase_falha} {recomendacao_teste}".strip()
