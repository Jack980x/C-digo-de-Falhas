import os
import re
import json
import concurrent.futures
from typing import List, Dict, Any, Tuple, Optional
from rich.console import Console

console = Console()

def reparar_e_carregar_json(texto: str) -> Any:
    """
    Carrega JSON gerado por LLMs com tolerância a falhas sintáticas comuns:
    1. Markdown ```json ... ```
    2. Aspas duplas não escapadas dentro de strings (ex: "Error on "ABS active" signal")
    3. Vírgulas faltantes entre objetos (} {) ou propriedades
    4. Vírgulas sobrando antes de fechamento (, } ou , ])
    5. Truncamento acidental no final por estouro de tokens
    6. Recuperação cirúrgica de objetos JSON individuais caso o array como um todo esteja corrompido
    """
    if not texto:
        return []
    texto = texto.strip()
    if texto.startswith("```"):
        linhas = texto.split("\n")
        if linhas[0].startswith("```"):
            linhas = linhas[1:]
        if linhas and linhas[-1].startswith("```"):
            linhas = linhas[:-1]
        texto = "\n".join(linhas).strip()

    # Tentativa 1: Parse direto padrão
    try:
        return json.loads(texto)
    except Exception:
        pass

    # Tentativa 2: Reparos sintáticos regex (vírgulas ausentes / sobrando)
    t = texto
    t = re.sub(r'\}\s*\{', '},\n{', t)
    t = re.sub(r'\]\s*\[', '],\n[', t)
    t = re.sub(r',\s*([\}\]])', r'\1', t)
    t = re.sub(r'(":[ \t]*"[^"\n]*")\s*\n\s*(")', r'\1,\n\2', t)
    t = re.sub(r'(":[ \t]*(?:true|false|null|\d+))\s*\n\s*(")', r'\1,\n\2', t)

    try:
        return json.loads(t)
    except Exception:
        pass

    # Tentativa 3: Se o JSON foi truncado no final (token limit)
    t_trunc = t
    num_abertos_chaves = t_trunc.count('{') - t_trunc.count('}')
    num_abertos_colch = t_trunc.count('[') - t_trunc.count(']')
    if num_abertos_chaves > 0 or num_abertos_colch > 0:
        if t_trunc.count('"') % 2 != 0:
            t_trunc += '"'
        t_trunc = re.sub(r',\s*$', '', t_trunc)
        t_trunc += ('}' * max(0, num_abertos_chaves))
        t_trunc += (']' * max(0, num_abertos_colch))
        try:
            return json.loads(t_trunc)
        except Exception:
            pass

    # Tentativa 4: Extrator cirúrgico de objetos individuais via Regex
    objetos_recuperados = []
    blocos = re.findall(r'\{[^{}]*\}', texto, re.DOTALL)
    for b in blocos:
        try:
            obj = json.loads(b)
            if isinstance(obj, dict):
                objetos_recuperados.append(obj)
        except Exception:
            b_rep = re.sub(r',\s*\}', '}', b)
            b_rep = re.sub(r'(":[ \t]*"[^"\n]*")\s*\n\s*(")', r'\1,\n\2', b_rep)
            try:
                obj = json.loads(b_rep)
                if isinstance(obj, dict):
                    objetos_recuperados.append(obj)
            except Exception:
                p_code = re.search(r'"(?:code|spn|dtc|Código|Codigo)"\s*:\s*"([^"]*)"', b, re.IGNORECASE)
                p_comp = re.search(r'"(?:component|componente|Fault Component|Localização da Avaria)"\s*:\s*"([^"]*)"', b, re.IGNORECASE)
                p_desc = re.search(r'"(?:description|descricao|Descrição|Condição de Teste)"\s*:\s*"([^"]*)"', b, re.IGNORECASE)
                p_fmi = re.search(r'"(?:fmi|FMi)"\s*:\s*"([^"]*)"', b, re.IGNORECASE)
                p_fmidesc = re.search(r'"(?:fmi_description|fmi_desc|FMI Description|Descrição FMI)"\s*:\s*"([^"]*)"', b, re.IGNORECASE)
                if p_code:
                    obj = {
                        "code": p_code.group(1),
                        "component": p_comp.group(1) if p_comp else "-",
                        "description": p_desc.group(1) if p_desc else "-",
                        "fmi": p_fmi.group(1) if p_fmi else "-",
                        "fmi_description": p_fmidesc.group(1) if p_fmidesc else "-"
                    }
                    objetos_recuperados.append(obj)

    if objetos_recuperados:
        return objetos_recuperados

    # Se nada funcionou, levanta o erro original do JSON
    return json.loads(texto)

class AgenteExtrator:
    """
    Agente 1: Especialista em Leitura de Tabelas Técnicas e Extração Estrutural.
    Foco: Extrair fielmente os dados do texto do OCR na sua forma ORIGINAL em inglês,
    isolando códigos, FMIs e componentes sem perder nenhuma informação.
    """
    def __init__(self, client, modelos_candidatos: List[str]):
        self.client = client
        self.modelos_candidatos = modelos_candidatos

    def _segmentar_texto_em_lotes(self, texto_bruto: str, codigos_gt: Optional[List[str]], tamanho_lote: int = 38) -> List[Tuple[str, List[str]]]:
        """
        Divide páginas longas ou com muitos códigos em lotes médios (~38 códigos).
        Evita erro 504 DEADLINE_EXCEEDED da API Gemini e garante 100% de extração sem omissões.
        """
        linhas = texto_bruto.splitlines()
        if not codigos_gt or len(codigos_gt) <= 38:
            if len(linhas) > 90:
                tamanho_linhas = 60
                return [('\n'.join(linhas[i:i + tamanho_linhas]), []) for i in range(0, len(linhas), tamanho_linhas)]
            return [(texto_bruto, codigos_gt or [])]

        blocos = []
        cods_no_bloco = []
        linhas_bloco = []

        for l in linhas:
            cods_na_linha = [c for c in codigos_gt if re.search(r'\b' + re.escape(c) + r'\b', l)]
            if cods_na_linha and len(cods_no_bloco) >= tamanho_lote:
                blocos.append(('\n'.join(linhas_bloco), list(cods_no_bloco)))
                linhas_bloco = [l]
                cods_no_bloco = list(cods_na_linha)
            else:
                linhas_bloco.append(l)
                for c in cods_na_linha:
                    if c not in cods_no_bloco:
                        cods_no_bloco.append(c)

        if linhas_bloco:
            blocos.append(('\n'.join(linhas_bloco), cods_no_bloco))

        return blocos

    def extrair(self, texto_bruto: str, modulo_alvo: Optional[str] = None, contexto: str = "Linha Pesada e Agrícola", codigos_ground_truth: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        if not self.client:
            raise ValueError("Cliente Gemini não inicializado no Agente Extrator.")

        blocos = self._segmentar_texto_em_lotes(texto_bruto, codigos_ground_truth, tamanho_lote=38)
        total_blocos = len(blocos)

        if total_blocos <= 1:
            console.print("      [dim]• Agente 1 (Extrator): Estruturando códigos e tabelas originais com IA...[/dim]")
            b_txt, b_gt = blocos[0]
            return self._extrair_bloco(b_txt, modulo_alvo, contexto, b_gt)

        console.print(f"      [dim]• Agente 1 (Extrator): Processando {total_blocos} blocos em paralelo com IA...[/dim]")
        resultados = [None] * total_blocos

        with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, total_blocos)) as executor:
            futures = {
                executor.submit(self._extrair_bloco, b_txt, modulo_alvo, contexto, b_gt): idx
                for idx, (b_txt, b_gt) in enumerate(blocos)
            }
            for future in concurrent.futures.as_completed(futures):
                idx = futures[future]
                try:
                    res = future.result()
                    resultados[idx] = res or []
                    console.print(f"      [dim]  • Bloco {idx + 1}/{total_blocos} extraido ({len(res or [])} codigos)[/dim]")
                except Exception as e:
                    console.print(f"      [red]  ! Erro no Bloco {idx + 1}/{total_blocos}: {e}[/red]")
                    resultados[idx] = []

        todos_registros = []
        for res in resultados:
            if res:
                todos_registros.extend(res)

        return todos_registros

    def _extrair_bloco(self, texto_bruto: str, modulo_alvo: Optional[str] = None, contexto: str = "Linha Pesada e Agrícola", codigos_ground_truth: Optional[List[str]] = None) -> List[Dict[str, Any]]:
        # Verifica se o documento possui coluna explícita de ECU/Módulo ou se é uma tabela geral de DTCs
        tem_coluna_ecu = bool(re.search(r'\b(?:ECU\s+Name|Module\s+Name)\b', texto_bruto, re.I))
        modulo_no_texto = False
        if modulo_alvo:
            partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', modulo_alvo.strip()) if x]
            if partes_m:
                modulo_no_texto = bool(re.search(r'\b' + r'[\s_-]*'.join(partes_m) + r'\b', texto_bruto, re.I))

        aplicar_filtro_estrito = bool(modulo_alvo and (tem_coluna_ecu or modulo_no_texto))

        instrucao_codigos_gt = ""
        if codigos_ground_truth:
            instrucao_codigos_gt = f"""
CÓDIGOS DE FALHA IDENTIFICADOS NO DOCUMENTO (TOTAL: {len(codigos_ground_truth)} CÓDIGOS — REFERÊNCIA DE GRAFIA EXATA):
{', '.join(codigos_ground_truth)}
- Use os {len(codigos_ground_truth)} códigos acima como guia de referência exata para preservar a grafia (não troque 'D' por '0', 'B' por '8', 'I' por '1', 'S' por '5', 'A' por '4' e mantenha zeros à esquerda).
- Se os códigos acima forem códigos DTC completos de 6 a 8 caracteres, com ou sem hífen (ex: B235A-68, C0021-93, U19B9-83, B181711, B181719, C130242, 094249) e a tabela não tiver coluna própria de FMI, mantenha o código completo EXATAMENTE IGUAL (incluindo o hífen e os 2 dígitos após o hífen, ex: "B235A-68") em "code" e coloque "-" em "fmi". NUNCA separe o sufixo após o hífen e NUNCA escreva " - FMI "!
- Quando uma linha da tabela possuir duas colunas de código lado a lado (ex: 'DTC Designação do Veículo' U240088 / C110116 na 1ª coluna e 'DTCs' hexadecimal E40088 / 510116 na 2ª coluna), extraia APENAS o código principal listado acima (1 registro por linha) e remova o código secundário da descrição. NUNCA duplique a linha extraindo as duas colunas como códigos separados!
- É OBRIGATÓRIO extrair TODOS os {len(codigos_ground_truth)} códigos de falha presentes na tabela da página, garantindo que NENHUM código seja esquecido ou omitido!"""

        instrucao_modulo = ""
        if aplicar_filtro_estrito:
            instrucao_modulo = f"""
CONTEXTO DE MÓDULO SOLICITADO: '{modulo_alvo}'
- Se a tabela possuir uma coluna específica de Módulo/ECU (ex: 'ECU Name') ou seções separadas por cabeçalhos de módulos distintos, extraia as falhas pertencentes a '{modulo_alvo}'.
- ATENÇÃO CRÍTICA: NUNCA descarte um código apenas porque o texto da descrição menciona siglas como 'ECU', 'BCM', 'ESP', 'EPS', 'LKA', 'CAN', 'ZF', etc. (ex: 'ECU RAM fault', 'Lost communication with BCM', 'ESP signal invalid' são descrições de falhas da própria tabela e DEVEM SER EXTRAÍDAS!).
- Se a tabela da página for uma lista geral de códigos (ex: colunas 'DTC | Description | Diagnosis process') sem coluna de módulo, extraia 100% dos códigos da página!"""

        prompt_sistema = f"""Você é o AGENTE 1: EXTRATOR ESTRUTURAL DE TABELAS TÉCNICAS ({contexto}).
Sua missão única é analisar o texto do manual e estruturar todos os códigos de falha originais na forma ORIGINAL EM INGLÊS.
NÃO TRADUZA OS TEXTOS NESTA ETAPA. A tradução será feita pelo Agente 2.
{instrucao_codigos_gt}
{instrucao_modulo}

REGRAS DE EXTRAÇÃO:
1. SEPARAÇÃO E FIDELIDADE ABSOLUTA AOS CÓDIGOS ORIGINAIS (IMUTABILIDADE ALFANUMÉRICA):
   - O campo "code" deve conter APENAS o código/número exato presente no manual (ex: "003D15", "003D68", "047228", "47221"). NUNCA inclua o nome do módulo junto ao código!
   - A ÚNICA COISA A SER TRADUZIDA POSTERIORMENTE SÃO OS TEXTOS DE DESCRIÇÃO, E NÃO OS CÓDIGOS OU NÚMEROS!
   - NUNCA confunda ou troque letras maiúsculas por números ou vice-versa (ex: NUNCA confunda a letra 'D' por '0', 'B' por '8', 'I' por '1' etc.). Se o manual tem "003D15", o código DEVE ser "003D15", JAMAIS "003015"!
2. IDENTIFICAÇÃO E ISOLAMENTO DO COMPONENTE / CIRCUITO (NUNCA DEIXE COMPONENTE VAZIO SE HOUVER INFORMAÇÃO):
   - O campo "component" DEVE conter o nome do componente físico, atuador, válvula, sensor, chicote, display, barramento CAN, ECU ou sinal afetado (ex: "Small eng. clutch valve", "Large diseng. clutch valve", "Output ADVP", "Display", "Engine retarder configuration", "Tachograph", "ECU", "CAN EBC1", "ABS active signal", "Clutch travel sensor", "Pressure sensor", "ECU temperature sensor").
   - Quando a coluna do manual combinar a falha com o componente (ex: "Short circuit to positive at output stage to small eng. clutch valve" ou "Mechanical failure of small engagement clutch valve" ou "CAN EBC1 timeout" ou "Error on 'ABS active' signal"):
     * NUNCA DEIXE "component" COMO "-" OU EM BRANCO!
     * Extraia o componente físico, atuador, saída ou sinal em "component" (ex: "Small eng. clutch valve", "CAN EBC1", "ABS active signal (EBC1)", "Small engagement clutch valve").
     * E mantenha o texto completo da condição em "description".
   - NUNCA deixe "component" como "-" se o texto do manual mencionar qualquer válvula, sensor, interruptor, atuador, pino, chicote, display, barramento CAN, ECU ou sinal!
3. EXTRAÇÃO COMPLETA DAS DESCRIÇÕES DE FALHA E FMI (PROIBIÇÃO DE DEIXAR VAZIO):
   - Extraia com máxima fidelidade todas as descrições de falha presentes no texto do manual.
   - Se o manual possuir coluna de "FMI Description" preenchida: extraia essa informação exatamente em "fmi_description".
   - Se a coluna de "FMI Description" estiver em branco ou não existir para um código, MAS o manual trouxer a descrição da avaria/defeito na coluna "Description" (ou "Fault Component Description", "Condition"): APROVEITE ESSA DESCRIÇÃO em "fmi_description" para que o modo de falha NÃO FIQUE VAZIO!
   - Se a linha do manual contiver o código e o FMI mas nenhuma descrição explícita (ex: apenas números como 41709, 4170A): extraia a descrição da falha do componente pai/acima ou preencha com a definição técnica correspondente (ex: para FMI 0C / 12: "Device error / Microcontroller defect").
4. FIDELIDADE ABSOLUTA AOS NÚMEROS DE FMI (PROIBIÇÃO TOTAL DE INVENTAR OU CONVERTER FMIs):
   - O campo "fmi" deve ser 100% FIEL ao que está impresso no manual original:
     * Extraia EXATAMENTE o número de FMI que consta no documento (ex: se o manual tem "02", extraia "02"; se tem "0B", extraia "0B"; se tem "0C", extraia "0C"; se tem "12", extraia "12").
     * PROIBIDO INVENTAR OU CONVERTER NÚMEROS DE FMI: JAMAIS deduza ou converta valores da sua memória (NUNCA adicione "11 / 0B" se o manual só tem "0B"; NUNCA crie "12 / 0C" se o manual só tem "0C"; NUNCA crie "18 / 12" se o manual só tem "12").
     * Use barra "/" APENAS se o manual original trouxer expressamente múltiplos FMIs (ex: "04 / 06") ou duas colunas distintas [DEC] e [HEX] preenchidas no documento original (ex: "12 / 0C"). Se o manual tiver apenas uma coluna de FMI ou um valor único, MANTENHA ESTRITAMENTE AQUELE VALOR ÚNICO!
     * PROIBIÇÃO TOTAL DE ADICIONAR FMIs QUE NÃO EXISTEM: Extraia ESTRITAMENTE os FMIs presentes no texto para aquele código. NUNCA invente ou adicione FMIs adicionais (como 03, 04, 05, etc.) que não existam no texto do manual para aquele código.
   - Se o mesmo código tiver múltiplos FMIs com descrições DIFERENTES no manual: crie um registro separado para cada FMI que realmente constar no manual.
   - Formate valores únicos de 0 a 9 com 2 dígitos (ex: "00", "01", ..., "09"). Valores alfanuméricos ou de 2 dígitos, mantenha como estão (ex: "0B", "0C", "12").

FORMATO DE RETORNO OBRIGATÓRIO (ARRAY JSON):
[
  {{
    "code": "003D68",
    "component": "Tank heating Valve",
    "description": "068_PUHT",
    "fmi": "06",
    "fmi_description": "Short circuit to ground"
  }},
  {{
    "code": "003D68",
    "component": "Tank heating Valve",
    "description": "068_PUHT",
    "fmi": "03",
    "fmi_description": "Short circuit to battery"
  }},
  {{
    "code": "003D68",
    "component": "Tank heating Valve",
    "description": "068_PUHT",
    "fmi": "05",
    "fmi_description": "Open load"
  }},
  {{
    "code": "003D67",
    "component": "Line Heater pressure tube (Box-Dosing module)",
    "description": "-",
    "fmi": "04 / 06",
    "fmi_description": "Short circuit to ground"
  }},
  {{
    "code": "40353",
    "component": "Small eng. clutch valve",
    "description": "Short circuit to positive at output stage to small eng. clutch valve",
    "fmi": "06",
    "fmi_description": "Shorted to high source"
  }},
  {{
    "code": "4035A",
    "component": "ECU (Controller 1 / Controller 2)",
    "description": "Communication error between controller 1 and controller 2 (ECU failure)",
    "fmi": "09",
    "fmi_description": "Device error"
  }},
  {{
    "code": "4035C",
    "component": "ABS active signal (EBC1)",
    "description": "Error on 'ABS active' signal (EBC1)",
    "fmi": "08",
    "fmi_description": "Invalid signal"
  }},
  {{
    "code": "41709",
    "component": "Pin A07 - Single-H Buzzer Output",
    "description": "High Side switch: S.C. to Vbat detected in Off condition; Open Load detected in Off condition",
    "fmi": "0C",
    "fmi_description": "Device error / Internal hardware defect"
  }},
  {{
    "code": "4170B",
    "component": "CAN 2 Bus",
    "description": "No CAN messages on CAN 2 or bus-off detection",
    "fmi": "02",
    "fmi_description": "No CAN messages on CAN 2 or bus-off detection"
  }}
]"""

        filtro_usr = f"pertencentes a '{modulo_alvo}'" if aplicar_filtro_estrito else "presentes na tabela acima"
        prompt_usuario = f"""Texto do manual (OCR):
---
{texto_bruto}
---
Extraia todos os registros de falha {filtro_usr} em JSON, mantendo 'description' e 'fmi_description' estruturados (se houver apenas uma coluna de falha, preencha ambas com ela):"""

        for tentativa in range(2):
            for modelo in self.modelos_candidatos:
                try:
                    resposta = self.client.models.generate_content(
                        model=modelo,
                        contents=[prompt_sistema, prompt_usuario],
                        config={
                            "temperature": 0.1,
                            "response_mime_type": "application/json"
                        }
                    )
                    dados = reparar_e_carregar_json(resposta.text)
                    if isinstance(dados, dict):
                        for k in ["dados", "tabela", "codes", "faults", "registros"]:
                            if k in dados and isinstance(dados[k], list):
                                dados = dados[k]
                                break
                        else:
                            dados = [dados]

                    if isinstance(dados, list):
                        if aplicar_filtro_estrito and modulo_alvo:
                            mod_alvo_clean = re.sub(r'[\s_-]+', '', modulo_alvo).upper()
                            filtrados = []
                            for item in dados:
                                mod_item = str(item.get("module", item.get("modulo", item.get("ecu", "")))).strip()
                                if mod_item and mod_item not in ["-", "None", ""]:
                                    mod_item_clean = re.sub(r'[\s_-]+', '', mod_item).upper()
                                    mod_raiz = re.split(r'[\s_-]+', modulo_alvo.strip())[0].upper()
                                    mod_item_raiz = re.split(r'[\s_-]+', mod_item.strip())[0].upper() if mod_item else ""
                                    if mod_raiz and mod_item_raiz and mod_raiz == mod_item_raiz and len(mod_raiz) >= 2:
                                        pass
                                    elif mod_item_clean != mod_alvo_clean and mod_alvo_clean not in mod_item_clean and mod_item_clean not in mod_alvo_clean:
                                        continue
                                filtrados.append(item)
                            return filtrados if (filtrados or tem_coluna_ecu) else dados
                        if codigos_ground_truth and not aplicar_filtro_estrito:
                            dados = self._resgatar_codigos_gt_omitidos(dados, texto_bruto, codigos_ground_truth)
                        return dados
                except Exception as e:
                    if any(x in str(e).lower() for x in ["404", "not_found", "not found", "504", "deadline_exceeded"]):
                        continue
                    if isinstance(e, json.JSONDecodeError) or any(x in str(e).lower() for x in ["delimiter", "json", "expecting"]):
                        console.print(f"      [yellow]  ! Aviso no modelo {modelo}: Formato JSON com erro sintático ({e}). Tentando reparar registros ou alternar modelo...[/yellow]")
                        continue
                    if tentativa == 0 and any(x in str(e).lower() for x in ["503", "500", "unavailable"]):
                        import time
                        time.sleep(2)
                        break
                    else:
                        raise e
        return []

    @staticmethod
    def _resgatar_codigos_gt_omitidos(
        dados: List[Dict[str, Any]],
        texto_bruto: str,
        codigos_ground_truth: List[str]
    ) -> List[Dict[str, Any]]:
        """
        Garante deterministicamente que nenhum código Ground Truth da página seja omitido pelo LLM no Agente 1.
        """
        if not codigos_ground_truth:
            return dados

        from core.ocr_engine import OCREngine
        pagina_eh_obd = len(codigos_ground_truth) >= 2 and all(
            re.match(r'^[PCBU]', str(c).strip().upper()) for c in codigos_ground_truth if c
        )
        dados_filtrados = []
        codigos_presentes = set()
        for item in dados:
            if isinstance(item, dict):
                c_raw = str(item.get("code", "")).strip()
                c_norm = OCREngine.normalizar_codigo_dtc_token(c_raw).upper()
                if pagina_eh_obd and c_norm and not re.match(r'^[PCBU]', c_norm):
                    continue
                item["code"] = c_norm
                codigos_presentes.add(c_norm)
                dados_filtrados.append(item)
        dados = dados_filtrados

        linhas_txt = texto_bruto.splitlines() if texto_bruto else []
        for c_gt in codigos_ground_truth:
            c_up = c_gt.strip().upper()
            if not c_up or c_up in codigos_presentes:
                continue
            # Procura a linha correspondente no texto bruto da página
            desc_en = "-"
            for l in linhas_txt:
                if re.search(r'\b' + re.escape(c_up) + r'\b', l, re.I):
                    partes = [p.strip() for p in l.split("|") if p.strip()]
                    if len(partes) >= 2 and partes[0].upper() == c_up:
                        desc_en = partes[1]
                    else:
                        limpa = re.sub(r'\b' + re.escape(c_up) + r'\b', '', l, flags=re.I).strip(" |-:")
                        limpa = re.sub(r'^[0-9A-F]{4,6}\b\s*', '', limpa, flags=re.I).strip(" |-:")
                        if limpa:
                            desc_en = limpa
                    break
            comp_en = desc_en
            m_comp = re.match(r'^(.+?)\s+(?:short-circuited|open-circuited|overload|fault|abnormal|circuit\s+fault).*$', desc_en, re.I)
            if m_comp:
                comp_en = m_comp.group(1).strip()
            dados.append({
                "code": c_up,
                "component": comp_en if comp_en else "-",
                "description": desc_en,
                "fmi": "-",
                "fmi_description": desc_en
            })
            codigos_presentes.add(c_up)

        return dados


class AgenteTradutor:
    """
    Agente 2: Especialista Lexicógrafo & Tradutor Técnico Automotivo.
    Foco: Traduzir os dados estruturados do Agente 1 para Português do Brasil (PT-BR)
    utilizando estritamente o Glossário Técnico e estilo telegráfico de concessionária.
    """
    def __init__(self, client, modelos_candidatos: List[str]):
        self.client = client
        self.modelos_candidatos = modelos_candidatos

    def traduzir(self, registros_originais: List[Dict[str, Any]], colunas_desejadas: List[str], glossario_texto: str, siglas_texto: str, contexto: str = "Linha Pesada e Agrícola") -> List[Dict[str, Any]]:
        if not registros_originais:
            return []
        if not self.client:
            raise ValueError("Cliente Gemini não inicializado no Agente Tradutor.")

        # Preserva a fidelidade estrita dos dados originais sem inventar ou clonar colunas
        registros_preparados = [dict(reg) for reg in registros_originais]

        colunas_str = ", ".join([f'"{c}"' for c in colunas_desejadas])

        prompt_sistema = f"""Você é o AGENTE 2: LEXICÓGRAFO & TRADUTOR TÉCNICO ({contexto}).
Sua missão é pegar os registros técnicos extraídos em inglês e traduzi-los para o Português do Brasil (PT-BR) no mais alto nível de engenharia de montadoras (Scania, Volvo, Iveco, John Deere).

DIRETRIZES DE TRADUÇÃO:
1. MAPEAMENTO E TRADUÇÃO DAS DESCRIÇÕES DE FALHA E FMI (PROIBIÇÃO DE DEIXAR VAZIO):
   - NUNCA deixe a coluna de descrição do FMI vazia ou com "-" se houver qualquer descrição disponível no manual ou se o FMI for conhecido!
   - "FMI Description" (ou "Descrição FMI", "Causa", "Defeito"):
     * Se "fmi_description" contiver texto técnico do manual, traduza-o para o português com máxima precisão.
     * Se "fmi_description" estiver vazia ou "-", MAS "description" contiver o texto do erro no manual (ex: "No CAN messages on CAN 2...", "Physical memory fault. Checksum failure is detected"): UTILIZE E TRADUZA A "description"!
     * Se AMBAS estiverem ausentes ou "-", mas o código tiver um número de FMI conhecido (ex: 00, 01, 02, 03, 04, 05, 06, 07, 08, 09, 0C, 12...): UTILIZE a definição técnica padrão da norma SAE J1939 em português (ex: FMI 02 -> "Dados erráticos, intermitentes ou incorretos (Sinal fora da faixa / Falha de comunicação)"; FMI 0C / 12 -> "Dispositivo ou componente inteligente com defeito (Falha interna de hardware/microcontrolador)").
   - "Description" (ou "Descrição"): traduza a coluna "description". Se "description" estiver vazia ou "-", mas "fmi_description" tiver texto, aproveite "fmi_description".
   - Se o manual possuir DUAS colunas distintas originais (testes elétricos em Description E definição em FMI Description), mantenha-as ricas e separadas em suas respectivas colunas.

2. TRADUÇÃO DA COLUNA DE COMPONENTE ("Fault Component", "Componente", "Localização da Avaria"):
   - Traduza o nome do componente técnico para Português do Brasil (PT-BR) de forma limpa, direta e profissional (ex: "Small eng. clutch valve" -> "Válvula de engate da embreagem menor", "Large diseng. clutch valve" -> "Válvula de desengate da embreagem maior", "Large eng. clutch valve" -> "Válvula de engate da embreagem maior", "Output ADVP" -> "Saída ADVP", "CAN EBC1" -> "Barramento CAN (EBC1)", "ABS active signal" -> "Sinal de ABS ativo (EBC1)", "Controller 1 and controller 2 (ECU failure)" -> "Controladores 1 e 2 (Falha na ECU)", "Clutch travel sensor" -> "Sensor de curso da embreagem", "Pressure sensor" -> "Sensor de pressão").
   - Se o campo "component" no registro recebido estiver vazio ou "-", MAS a "description" contiver o nome do componente (ex: após "output stage to...", "failure of...", "signal...", "timeout...", "voltage of..."): EXTRAIA E TRADUZA O COMPONENTE A PARTIR DA DESCRIÇÃO! JAMAIS deixe a coluna de componente vazia ou com "-" se o texto descrever o componente ou circuito afetado!

3. INTERPRETAÇÃO INTELIGENTE DE FÓRMULAS E FAIXAS ELÉTRICAS NO FMI:
   - Quando a descrição do FMI contiver fórmulas matemáticas, inequações ou limites elétricos (ex: "V>8V", "V<2V", "6V<V<8V", "2V<V<4V", "Vout/Vc>1V", "Vout/Vc<0.14V", "I<threshold", "R>10k"):
     * Esse formato de fórmula pura é difícil de interpretar para técnicos e mecânicos na oficina.
     * VOCÊ DEVE INTERPRETAR O SIGNIFICADO TÉCNICO DO ERRO EM PORTUGUÊS CLARO (conforme o padrão automotivo SAE J1939 para aquele FMI) E PRESERVAR A FÓRMULA/CONDIÇÃO ORIGINAL ENTRE PARÊNTESES!
     * Exemplos obrigatórios de tradução:
       - FMI 00 com "6V<V<8V": "Tensão acima da faixa operacional normal (6V < V < 8V)"
       - FMI 01 com "2V<V<4V": "Tensão abaixo da faixa operacional normal (2V < V < 4V)"
       - FMI 03 com "V>8V": "Tensão acima do normal / Curto ao positivo (V > 8V)"
       - FMI 04 com "V<2V": "Tensão abaixo do normal / Curto ao aterramento (V < 2V)"
       - FMI 03 com "Vout/Vc>1V": "Sinal do sensor acima do limite operacional (Vout/Vc > 1V)"
       - FMI 04 com "Vout/Vc<0.14V": "Sinal do sensor abaixo do limite operacional (Vout/Vc < 0.14V)"
       - FMI 05 com "I<threshold" ou "Open circuit": "Circuito aberto / Corrente abaixo do limite (I < limite)"
       - FMI 06 com "I>threshold" ou "Short circuit": "Corrente acima do limite / Sobrecarga (I > limite)"
   - Se o texto do FMI já for puramente textual em inglês (ex: "Root cause not known", "Internal electronic failure"): apenas traduza literalmente e de forma técnica para o português, sem inventar nada.

3. GLOSSÁRIO TÉCNICO OBRIGATÓRIO (PRIORIDADE TOTAL):
{glossario_texto}

4. SIGLAS A PRESERVAR SEM TRADUZIR:
   {siglas_texto}, etc.

5. REGRA DE OURO ELÉTRICA:
   - Termos GND / Ground / S.C. to Gnd devem ser SEMPRE traduzidos como "aterramento" ou "curto-circuito ao aterramento". JAMAIS use "terra".

6. ESTILO TELEGRÁFICO DE SCANNER AUTOMOTIVO:
   - Use frases diretas, objetivas, técnicas e concisas (ex: "Tensão de sinal acima do limite máximo", "Sensor de pressão com falha de plausibilidade").
   - Evite prolixidade ou voz passiva extensa.

7. FIDELIDADE E IMUTABILIDADE ABSOLUTA DA COLUNA FMI:
   - Copie os valores da coluna 'fmi' EXATAMENTE como recebidos do Agente 1, sem modificar, sem converter e sem adicionar nenhum número.
   - NUNCA invente ou adicione novos registros de FMI. A lista de falhas e FMIs recebida do Agente 1 é definitiva.

8. COLUNAS SOLICITADAS:
   Cada objeto resultante no array JSON deve conter EXATAMENTE as colunas:
   [{colunas_str}]

9. IMUTABILIDADE ABSOLUTA DOS CÓDIGOS DE FALHA (SPN / DTC / HEX):
   - A ÚNICA COISA A SER TRADUZIDA SÃO OS TEXTOS DAS DESCRIÇÕES!
   - CÓDIGOS, NÚMEROS E SUAS LETRAS SÃO 100% IMUTÁVEIS.
   - Copie os valores das colunas de código ('code', 'spn', 'dtc') e 'fmi' caractere por caractere exatamente como recebidos do Agente 1, sem nenhuma alteração (ex: '003D15' deve permanecer estritamente '003D15', NUNCA troque letras por números).

10. FORMATO DE RETORNO:
   Retorne estritamente um array JSON válido contendo os registros traduzidos."""

        # Tradução em lotes de até 45 registros processados em paralelo
        tamanho_lote = 45
        lotes = [registros_preparados[i:i + tamanho_lote] for i in range(0, len(registros_preparados), tamanho_lote)]
        total_lotes = len(lotes)

        if total_lotes <= 1:
            console.print("      [dim]• Agente 2 (Tradutor): Traduzindo registros com IA...[/dim]")
            _, unico_traduzido = self._traduzir_lote_individual(0, lotes[0], prompt_sistema, colunas_str)
            return unico_traduzido

        console.print(f"      [dim]• Agente 2 (Tradutor): Traduzindo {total_lotes} lotes em paralelo com IA...[/dim]")
        resultados = [None] * total_lotes

        with concurrent.futures.ThreadPoolExecutor(max_workers=min(4, total_lotes)) as executor:
            futures = [
                executor.submit(self._traduzir_lote_individual, idx, lote, prompt_sistema, colunas_str)
                for idx, lote in enumerate(lotes)
            ]
            for future in concurrent.futures.as_completed(futures):
                try:
                    idx, lote_trad = future.result()
                    resultados[idx] = lote_trad or []
                    console.print(f"      [dim]  ✓ Lote {idx + 1}/{total_lotes} traduzido ({len(lote_trad or [])} registros)[/dim]")
                except Exception as e:
                    console.print(f"      [red]  ✗ Erro no Lote {idx + 1}: {e}[/red]")

        todos_traduzidos = []
        for idx, r in enumerate(resultados):
            if r:
                todos_traduzidos.extend(r)
            elif idx < len(lotes):
                todos_traduzidos.extend(lotes[idx])

        return todos_traduzidos

    def _traduzir_lote_individual(self, idx: int, lote: List[Dict[str, Any]], prompt_sistema: str, colunas_str: str) -> Tuple[int, List[Dict[str, Any]]]:
        prompt_usuario = f"""Registros técnicos em inglês extraídos pelo Agente 1 (Lote {idx + 1}):
{json.dumps(lote, indent=2, ensure_ascii=False)}

Traduza e estruture todos os {len(lote)} registros do lote nas colunas solicitadas [{colunas_str}]:"""

        for tentativa in range(2):
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
                    dados = reparar_e_carregar_json(resposta.text)
                    if isinstance(dados, list):
                        return idx, dados
                    elif isinstance(dados, dict):
                        for k in ["dados", "tabela", "traducao", "registros"]:
                            if k in dados and isinstance(dados[k], list):
                                return idx, dados[k]
                        return idx, [dados]
                except Exception as e:
                    if any(x in str(e).lower() for x in ["404", "not_found", "not found", "504", "deadline_exceeded"]):
                        continue
                    if isinstance(e, json.JSONDecodeError) or any(x in str(e).lower() for x in ["delimiter", "json", "expecting"]):
                        console.print(f"      [yellow]  ! Aviso no modelo {modelo}: Formato JSON com erro sintático ({e}). Tentando reparar lote ou alternar modelo...[/yellow]")
                        continue
                    if tentativa == 0 and any(x in str(e).lower() for x in ["503", "500", "unavailable"]):
                        import time
                        time.sleep(2)
                        break
                    else:
                        raise e
        return idx, lote


class AgenteAuditor:
    """
    Agente 3: Auditor de Qualidade, Padronização e Conformidade.
    Foco: Validação pós-tradução determinística:
    - Garante FMIs em hexadecimal de 2 dígitos com zeros à esquerda ('01', '02', '0C'...)
    - Elimina 100% de qualquer resquício de 'terra' convertendo para 'aterramento'
    - Resgata descrições de erro se alguma coluna tiver ficado acidentalmente vazia ou '-'
    - Garante preservação de zeros à esquerda em códigos SPN/DTC
    - Higieniza a coluna de componentes
    - Emite relatório de auditoria e conformidade
    """
    @staticmethod
    def formatar_fmi_hex_2digitos(val: Any) -> str:
        """
        Garante que o número do FMI seja 100% fiel ao material original:
        - Não inventa e não adiciona conversões DEC/HEX (ex: '0B' permanece '0B', '0C' permanece '0C', '12' permanece '12').
        - Se o manual já trouxer expressamente múltiplos FMIs separados por barra (ex: '04 / 06' ou '12 / 0C'), preserva com a barra.
        - Formata apenas números de 1 dígito com zero à esquerda (ex: '2' -> '02', 'B' -> '0B').
        """
        if val is None or str(val).strip() in ["", "-", "None", "null", "nan"]:
            return "-"
        v = str(val).strip()
        v = re.sub(r'^(?:FMI|fmi|HEX|hex|DEC|dec)[:\s#]*', '', v).strip()

        if not v or v in ["-", "NONE", "null"]:
            return "-"

        # 1. Se contiver barra (ex: "04 / 06" ou "12 / 0C" quando já presente no manual)
        if "/" in v:
            partes_raw = [p.strip() for p in v.split("/") if p.strip()]
            partes_formatadas = []
            for p in partes_raw:
                p_limpo = re.sub(r'^(?:FMI|fmi|HEX|hex|DEC|dec)[:\s#]*', '', p).strip()
                if not p_limpo or p_limpo in ["-", ""]:
                    continue
                if len(p_limpo) == 1 and p_limpo.isalnum():
                    partes_formatadas.append(p_limpo.upper().zfill(2))
                else:
                    partes_formatadas.append(p_limpo.upper())

            partes_unicas = list(dict.fromkeys(partes_formatadas))
            if not partes_unicas:
                return "-"
            if len(partes_unicas) == 1:
                return partes_unicas[0]
            return " / ".join(partes_unicas)

        # 2. Se for número ou hex de 1 dígito (ex: "2" -> "02", "B" -> "0B")
        if len(v) == 1 and v.isalnum():
            return v.upper().zfill(2)

        # 3. Se for de 2 ou mais caracteres (ex: "02", "0B", "0C", "12", "13")
        # Mantém 100% fiel ao original, sem inventar/adicionar pares DEC/HEX!
        return v.upper()

    @staticmethod
    def _decompor_fmi_atomos(s: str) -> List[str]:
        if not s or str(s).strip() in ["-", "None", ""]:
            return []
        s = str(s).strip()
        partes = []
        for p in re.split(r'[\s/]+', s):
            p_limpo = re.sub(r'^(?:FMI|fmi|HEX|hex|DEC|dec)[:\s#]*', '', p).strip().upper()
            if p_limpo and p_limpo not in ["-", "NONE", ""]:
                if len(p_limpo) == 1 and p_limpo.isalnum():
                    p_limpo = p_limpo.zfill(2)
                if p_limpo not in partes:
                    partes.append(p_limpo)
        return partes

    def _consolidar_fmis_mesma_descricao(self, registros: List[Dict[str, Any]], col_codigo: Optional[str], col_fmi: Optional[str], col_desc: Optional[str], col_componente: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Consolida registros que possuem o mesmo código e a mesma descrição técnica,
        unindo seus FMIs na mesma linha e coluna como "04 / 06" para não perder informações nem duplicar linhas.
        """
        if not col_codigo or not col_desc or not registros:
            return registros

        def normalizar_texto(t):
            if not t:
                return ""
            return re.sub(r'[\s\.,;:\-_/]+', ' ', str(t).strip().lower())

        agrupados = {}
        ordem_chaves = []

        for reg in registros:
            cod = str(reg.get(col_codigo, "")).strip().upper()
            desc = normalizar_texto(reg.get(col_desc, ""))
            comp = normalizar_texto(reg.get(col_componente, "")) if col_componente else ""
            chave = (cod, comp, desc)

            if chave not in agrupados:
                agrupados[chave] = dict(reg)
                ordem_chaves.append(chave)
            else:
                existente = agrupados[chave]
                if col_fmi:
                    fmi_exist = str(existente.get(col_fmi, "")).strip()
                    fmi_novo = str(reg.get(col_fmi, "")).strip()

                    atomos = []
                    for f in [fmi_exist, fmi_novo]:
                        for a in self._decompor_fmi_atomos(f):
                            if a not in atomos:
                                atomos.append(a)

                    def chave_ordenacao_atomo(item):
                        if "/" in item:
                            p = item.split("/")[0].strip()
                            if p.isdigit():
                                return (0, int(p), "")
                        try:
                            return (0, int(str(item), 16), "")
                        except Exception:
                            pass
                        return (1, 0, str(item))

                    atomos.sort(key=chave_ordenacao_atomo)
                    existente[col_fmi] = " / ".join(atomos) if atomos else "-"

                # Mescla outras colunas (como módulo) caso sejam diferentes
                for col, val in reg.items():
                    if col not in [col_codigo, col_fmi, col_desc, col_componente]:
                        v_exist = str(existente.get(col, "")).strip()
                        v_novo = str(val).strip()
                        if v_novo and v_novo not in ["-", "None", ""] and v_novo != v_exist:
                            if v_exist and v_exist not in ["-", "None", ""]:
                                subpartes = [s.strip() for s in v_exist.split(" / ") if s.strip()]
                                if v_novo not in subpartes:
                                    subpartes.append(v_novo)
                                existente[col] = " / ".join(subpartes)
                            else:
                                existente[col] = v_novo

        return [agrupados[k] for k in ordem_chaves]

    MAPA_FMI_J1939 = {
        "00": "Dados válidos mas acima da faixa operacional normal (Nível de gravidade máximo)",
        "0": "Dados válidos mas acima da faixa operacional normal (Nível de gravidade máximo)",
        "01": "Dados válidos mas abaixo da faixa operacional normal (Nível de gravidade máximo)",
        "1": "Dados válidos mas abaixo da faixa operacional normal (Nível de gravidade máximo)",
        "02": "Dados erráticos, intermitentes ou incorretos (Sinal fora da faixa / Falha de comunicação)",
        "2": "Dados erráticos, intermitentes ou incorretos (Sinal fora da faixa / Falha de comunicação)",
        "03": "Tensão acima do normal ou curto-circuito ao positivo (Vbatt)",
        "3": "Tensão acima do normal ou curto-circuito ao positivo (Vbatt)",
        "04": "Tensão abaixo do normal ou curto-circuito ao aterramento (GND)",
        "4": "Tensão abaixo do normal ou curto-circuito ao aterramento (GND)",
        "05": "Corrente abaixo do normal ou circuito aberto",
        "5": "Corrente abaixo do normal ou circuito aberto",
        "06": "Corrente acima do normal ou circuito em curto ao aterramento",
        "6": "Corrente acima do normal ou circuito em curto ao aterramento",
        "07": "Sistema mecânico não responde adequadamente",
        "7": "Sistema mecânico não responde adequadamente",
        "08": "Frequência, largura de pulso ou período anormal",
        "8": "Frequência, largura de pulso ou período anormal",
        "09": "Taxa de atualização anormal (Timeout de comunicação no barramento CAN)",
        "9": "Taxa de atualização anormal (Timeout de comunicação no barramento CAN)",
        "0A": "Taxa de variação anormal",
        "10": "Dados válidos acima da faixa normal (Gravidade moderada)",
        "0B": "Causa da falha não identificável",
        "11": "Dados válidos abaixo da faixa normal (Gravidade moderada)",
        "0C": "Dispositivo ou componente inteligente com defeito (Falha interna de hardware/microcontrolador)",
        "12": "Dispositivo ou componente inteligente com defeito (Falha interna de hardware/microcontrolador)",
        "0D": "Fora de calibração",
        "13": "Erro de comunicação de dados de rede recebidos",
        "0E": "Instruções especiais de falha",
        "14": "Instruções especiais de falha",
        "0F": "Dados válidos acima da faixa normal (Gravidade mínima)",
        "15": "Dados válidos acima da faixa normal (Gravidade mínima)",
        "16": "Dados válidos acima da faixa normal (Gravidade moderada)",
        "17": "Dados válidos abaixo da faixa normal (Gravidade moderada)",
        "18": "Dados válidos abaixo da faixa normal (Gravidade mais severa)",
        "19": "Erro de comunicação de dados de rede recebidos",
        "1F": "Condição existente não reportada",
        "31": "Condição existente não reportada",
    }

    @classmethod
    def obter_descricao_padrao_fmi(cls, fmi_str: Any) -> str:
        """
        Retorna a definição técnica padrão da norma SAE J1939 para o FMI quando o manual
        não fornecer uma descrição textual explícita na linha ou quando estiver em branco.
        """
        if not fmi_str or str(fmi_str).strip() in ["-", "", "None", "null", "nan"]:
            return ""
        s = str(fmi_str).strip().upper()
        # Se for par decimal / hex (ex: "12 / 0C" ou "2 / 02")
        if "/" in s:
            partes = [p.strip() for p in s.split("/") if p.strip()]
            if len(partes) >= 2:
                p0, p1 = partes[0], partes[1]
                try:
                    if int(p0) == int(p1, 16):
                        chave = p1.zfill(2)
                        if chave in cls.MAPA_FMI_J1939:
                            return cls.MAPA_FMI_J1939[chave]
                except Exception:
                    pass
            descs = []
            for p in partes:
                p_c = p.zfill(2)
                if p_c in cls.MAPA_FMI_J1939:
                    d = cls.MAPA_FMI_J1939[p_c]
                    if d not in descs:
                        descs.append(d)
            if descs:
                return " / ".join(descs)
            s = partes[0]

        s_clean = s.zfill(2)
        if s_clean in cls.MAPA_FMI_J1939:
            return cls.MAPA_FMI_J1939[s_clean]
        if s in cls.MAPA_FMI_J1939:
            return cls.MAPA_FMI_J1939[s]
        return ""

    @staticmethod
    def interpretar_fmi_formula(fmi: str, texto_atual: str, formula_original: str = "") -> str:
        """
        Interpreta fórmulas e inequações elétricas de FMI mantendo a fórmula original entre parênteses.
        Ex: '6V<V<8V' -> 'Tensão acima da faixa operacional normal (6V < V < 8V)'
            'V>8V' -> 'Tensão acima do normal / Curto ao positivo (V > 8V)'
            'V<2V' -> 'Tensão abaixo do normal / Curto ao aterramento (V < 2V)'
        """
        texto_base = formula_original.strip() if formula_original and formula_original.strip() not in ["-", "", "None"] else texto_atual.strip()
        if not texto_base or texto_base in ["-", "None"]:
            return texto_atual

        # Se já possui texto explicativo em português e parênteses com a fórmula, preserva
        if any(p in texto_atual.lower() for p in ["tensão", "tensao", "curto", "circuito", "sinal", "corrente", "aberto", "acima", "abaixo"]):
            if "(" in texto_atual and ")" in texto_atual:
                return texto_atual

        f_limpa = texto_base
        fmi_norm = str(fmi).strip().upper()
        if "/" in fmi_norm:
            fmi_norm = fmi_norm.split("/")[0].strip()

        # Espaçamento limpo da fórmula: 6V<V<8V -> 6V < V < 8V
        f_formatada = re.sub(r'\s*([<>]=?)\s*', r' \1 ', f_limpa)

        # Se a fórmula contiver tensão (V, Vc, Vout)
        if "v" in f_limpa.lower():
            if fmi_norm in ["03", "3"]:
                if "vout/vc" in f_limpa.lower():
                    return f"Sinal do sensor acima do limite operacional ({f_formatada})"
                return f"Tensão acima do normal / Curto ao positivo ({f_formatada})"
            elif fmi_norm in ["04", "4"]:
                if "vout/vc" in f_limpa.lower():
                    return f"Sinal do sensor abaixo do limite operacional ({f_formatada})"
                return f"Tensão abaixo do normal / Curto ao aterramento ({f_formatada})"
            elif fmi_norm in ["00", "0"]:
                return f"Tensão acima da faixa operacional normal ({f_formatada})"
            elif fmi_norm in ["01", "1"]:
                return f"Tensão abaixo da faixa operacional normal ({f_formatada})"
            elif fmi_norm in ["02", "2"]:
                return f"Sinal fora da faixa de tolerância ({f_formatada})"

        # Se contiver corrente (A, mA, I)
        elif any(x in f_limpa.lower() for x in ["a", "ma", "i"]):
            if fmi_norm in ["05", "5"]:
                return f"Circuito aberto / Corrente abaixo do limite ({f_formatada})"
            elif fmi_norm in ["06", "6"]:
                return f"Corrente acima do limite / Sobrecarga ({f_formatada})"

        # Se contiver resistência (ohm, k)
        elif "ohm" in f_limpa.lower() or "k" in f_limpa.lower():
            if fmi_norm in ["03", "05", "3", "5"]:
                return f"Resistência alta / Circuito aberto ({f_formatada})"
            elif fmi_norm in ["04", "06", "4", "6"]:
                return f"Resistência baixa / Curto-circuito ({f_formatada})"

        # Se contiver inequação matemática geral
        if any(sym in f_limpa for sym in ["<", ">", "="]):
            return f"Condição de teste elétrico ({f_formatada})"

        return texto_atual

    @staticmethod
    def reconciliar_codigo(cod: str, texto_fonte: str = None, codigos_ground_truth: Optional[List[str]] = None) -> Tuple[str, bool]:
        """
        Reconcilia deterministicamente erros de leitura OCR/LLM em códigos de falha (DTC/SPN/HEX):
        - Troca do dígito '0' pela letra 'D' (ex: '003015' -> '003D15', '003013' -> '003D13')
        - Troca do dígito '8' pela letra 'B'
        - Troca do dígito '1' pela letra 'I'
        - Troca da letra 'O' pelo dígito '0'
        Valida prioritariamente contra a lista de códigos Ground Truth reais do documento.
        """
        if not cod:
            return cod, False

        from core.ocr_engine import OCREngine
        cod_upper = str(cod).strip().upper()
        # Converte padrões indevidos como 'U1970 - FMI 83' -> 'U1970-83'
        cod_upper = re.sub(r'^([PCBU]?[0-9A-F]{4})\s*[-–]\s*FMI\s*([0-9A-F]{2})$', r'\1-\2', cod_upper, flags=re.I)
        cod_norm_ocr = OCREngine.normalizar_codigo_dtc_token(cod_upper)
        foi_norm_ocr = (cod_norm_ocr != str(cod).strip().upper())
        if cod_norm_ocr != cod_upper:
            cod_upper = cod_norm_ocr

        # 1. Se temos a lista oficial de códigos Ground Truth do PDF
        if codigos_ground_truth:
            gt_set = set(c.upper() for c in codigos_ground_truth if c)
            if cod_upper in gt_set:
                return cod_upper, foi_norm_ocr

            # Equivalência com/sem hífen (ex: B235A68 <-> B235A-68) ou sem letra inicial (ex: 0034-29 -> C0034-29)
            cod_sem_hifen = re.sub(r'[\s_-]+', '', cod_upper)
            for gt in codigos_ground_truth:
                if not gt:
                    continue
                gt_up = gt.upper()
                if re.sub(r'[\s_-]+', '', gt_up) == cod_sem_hifen:
                    return gt_up, True
                if len(cod_upper) == 7 and cod_upper[4] == '-' and gt_up.endswith(cod_upper) and len(gt_up) == 8:
                    return gt_up, True

            # Heurísticas de correção OCR
            trocas = [('0', 'D'), ('D', '0'), ('8', 'B'), ('B', '8'), ('1', 'I'), ('I', '1'), ('O', '0'), ('0', 'O')]
            for orig, rep in trocas:
                for i, ch in enumerate(cod_upper):
                    if ch == orig:
                        cand = cod_upper[:i] + rep + cod_upper[i+1:]
                        if cand in gt_set:
                            return cand, True

            # Prefixo clássico: 0030* -> 003D*
            if cod_upper.startswith("0030"):
                cand = "003D" + cod_upper[4:]
                if cand in gt_set:
                    return cand, True

            # Distância de 1 caractere
            for gt in gt_set:
                if len(gt) == len(cod_upper):
                    diffs = sum(1 for a, b in zip(cod_upper, gt) if a != b)
                    if diffs == 1:
                        return gt, True

        # 2. Se não temos a lista, valida contra texto_fonte (usando apenas texto_nativo confiável)
        if texto_fonte:
            if re.search(r'\b' + re.escape(cod_upper) + r'\b', texto_fonte):
                return cod_upper, False

            mapa_trocas = [('0', 'D'), ('8', 'B'), ('1', 'I'), ('O', '0')]
            for orig_char, rep_char in mapa_trocas:
                for i, ch in enumerate(cod_upper):
                    if ch == orig_char:
                        cand = cod_upper[:i] + rep_char + cod_upper[i+1:]
                        m = re.search(r'\b' + re.escape(cand) + r'\b', texto_fonte, re.I)
                        if m:
                            return m.group(0).upper(), True

            if cod_upper.startswith("0030"):
                cand = "003D" + cod_upper[4:]
                m = re.search(r'\b' + re.escape(cand) + r'\b', texto_fonte, re.I)
                if m:
                    return m.group(0).upper(), True

        return cod_upper, False

    @classmethod
    def extrair_e_traduzir_componente(cls, texto_orig: str) -> str:
        """
        Extrai e traduz deterministicamente o componente caso o LLM o tenha deixado vazio ou '-'
        """
        if not texto_orig or str(texto_orig).strip() in ["-", "", "None", "null", "nan"]:
            return "-"
        
        t = re.sub(r'\s+', ' ', str(texto_orig)).strip()

        # 1. Output stage to ...
        m = re.search(r'(?:output stage to|stage to|at output stage to)\s+(?:Y\d+\s*)?(?:\([^)]+\)\s*)?(?:to\s+)?(.+)$', t, re.I)
        candidato = m.group(1).strip(' -:;') if m else ''

        # 2. Output ADVP / Output SD ...
        if not candidato:
            m = re.search(r'(?:output|sa[ií]da)\s+([A-Z0-9_-]+(?:\s*\([^)]+\))?.*)$', t, re.I)
            if m:
                candidato = m.group(1).strip(' -:;')

        # 3. CAN ... timeout
        if not candidato:
            m = re.search(r'CAN\s+[\W_]*([A-Za-z0-9_\s/-]+?)[\W_]*\s+timeout', t, re.I)
            if m:
                candidato = f'Barramento CAN ({m.group(1).strip()})'

        # 4. Error on ...
        if not candidato:
            m = re.search(r'Error on\s+[\W_]*([A-Za-z0-9_\s/-]+?)[\W_]*(?:\s+signal)?(?:\s*\(([^)]+)\))?$', t, re.I)
            if m:
                item = m.group(1).strip()
                sigla = m.group(2).strip() if m.group(2) else ''
                if sigla and f'({sigla})' not in item:
                    candidato = f'{item} ({sigla})'
                else:
                    candidato = item

        # 5. Communication error between ...
        if not candidato:
            m = re.search(r'Communication error between\s+(.+)$', t, re.I)
            if m:
                candidato = m.group(1).strip(' -:;')

        # 6. Mechanical failure of ...
        if not candidato:
            m = re.search(r'Mechanical failure of\s+(.+)$', t, re.I)
            if m:
                candidato = m.group(1).strip(' -:;')

        # 7. Supply voltage of ...
        if not candidato:
            m = re.search(r'Supply voltage of\s+(.+?)(?:\s+too low|\s+below|\s+above|\s*\(ECU failure\)|$)', t, re.I)
            if m:
                candidato = m.group(1).strip(' -:;')

        # 8. Interruption at output stage to ...
        if not candidato:
            m = re.search(r'Interruption at output stage to\s+(.+)$', t, re.I)
            if m:
                candidato = m.group(1).strip(' -:;')

        if not candidato:
            candidato = t

        # Simplificação de saídas conhecidas
        if re.search(r'\bADVP\b', candidato, re.I):
            return 'Saída ADVP'

        # Dicionário de termos automotivos frequentes em componentes
        mapa_termos = [
            (r'\bsmall\s+eng\.?\s+clutch\s+valve\b', 'Válvula de engate da embreagem menor'),
            (r'\blarge\s+eng\.?\s+clutch\s+valve\b', 'Válvula de engate da embreagem maior'),
            (r'\blarge\s+diseng\.?\s+clutch\s+valve\b', 'Válvula de desengate da embreagem maior'),
            (r'\bsmall\s+engagement\s+clutch\s+valve\b', 'Válvula de engate da embreagem menor'),
            (r'\blarge\s+engagement\s+clutch\s+valve\b', 'Válvula de engate da embreagem maior'),
            (r'\bclutch\s+position\s+sensor\b', 'Sensor de posição da embreagem'),
            (r'\bclutch\s+travel\s+signal\b', 'Sinal de curso da embreagem'),
            (r'\bclutch\s+travel\b', 'Sensor de curso da embreagem'),
            (r'\bpressure\s+sensor\s+signal\b', 'Sinal do sensor de pressão'),
            (r'\bpressure\s+sensor\b', 'Sensor de pressão'),
            (r'\bECU\s+temperature\s+sensor\s+signal\b', 'Sinal do sensor de temperatura da ECU'),
            (r'\bECU\s+temperature\s+sensor\b', 'Sensor de temperatura da ECU'),
            (r'\bcontroller\s*1\s+and\s+controller\s*2\s*\(ECU\s+failure\)\b', 'Controladores 1 e 2 (Falha na ECU)'),
            (r'\bcontroller\s*1\s+and\s+controller\s*2\b', 'Controladores 1 e 2 (ECU)'),
            (r'\boutput\s+SD\s+to\s+display\b', 'Saída SD para display'),
            (r'\bSD\s+to\s+display\b', 'Saída para display'),
            (r'\bdisplay\b', 'Display'),
            (r'\bCAN\s+EBC1\b', 'Barramento CAN (EBC1)'),
            (r'\bEBC1\b', 'EBC1'),
            (r'\bABS\s+active\b', 'Sinal de ABS ativo'),
            (r'\bEBS\s+message\b', 'Mensagem EBS'),
            (r'\bASR\s+engine\s+control\s+active\b', 'Sinal de controle de motor ASR ativo'),
            (r'\bASR\s+brake\s+control\s+active\b', 'Sinal de controle de freio ASR ativo'),
            (r'\bTachograph\b', 'Tacógrafo (TCO1)'),
            (r'\bTime/Date\b', 'Hora/Data (TD)'),
            (r'\bEngine\s+configuration\s*1\b', 'Configuração do motor 1 (EC1)'),
            (r'\bEngine\s+configuration\s*2\b', 'Configuração do motor 2 (EC2)'),
            (r'\bEngine\s+configuration\b', 'Configuração do motor (EC)'),
            (r'\bEngine\s+speed\b', 'Velocidade do motor (EEC1)'),
            (r'\bCruise\s+control\s+set\s+speed\b', 'Velocidade do controle de cruzeiro (CCVS)'),
            (r'\bEngine\s+retarder\s+configuration\s+message\b', 'Mensagem de configuração do freio motor'),
            (r'\bEngine\s+retarder\s+configuration\b', 'Configuração do freio motor'),
        ]

        traduzido = candidato
        for pat, rep in mapa_termos:
            if re.search(pat, traduzido, re.I):
                traduzido = re.sub(pat, rep, traduzido, flags=re.I)

        traduzido = re.sub(r'^(?:at\s+|to\s+|for\s+)', '', traduzido, flags=re.I).strip(' -:;')
        traduzido = re.sub(r'\s*\(\s*ECU\s+failure\s*\)', '', traduzido, flags=re.I).strip()
        traduzido = re.sub(r'(\([A-Za-z0-9_]+\))\s*\1+', r'\1', traduzido).strip()
        return traduzido or "-"

    def auditar(self, registros: List[Dict[str, Any]], colunas_desejadas: List[str], registros_originais: Optional[List[Dict[str, Any]]] = None, texto_bruto: Optional[str] = None, texto_nativo: Optional[str] = None, codigos_ground_truth: Optional[List[str]] = None) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        if not registros:
            return [], {"total": 0, "status": "Vazio"}

        console.print(f"      [dim]• Agente 3 (Auditor): Auditando integridade e FMIs de {len(registros)} registros...[/dim]")

        # Identifica colunas-chave
        col_fmi = None
        col_codigo = None
        col_componente = None
        col_desc = None
        col_fmi_desc = None

        for col in colunas_desejadas:
            col_l = col.lower()
            if "fmi" in col_l and any(k in col_l for k in ["desc", "descrição", "descricao"]):
                col_fmi_desc = col
            elif "fmi" in col_l:
                col_fmi = col
            elif any(k in col_l for k in ["spn", "psn", "código", "codigo", "dtc"]):
                col_codigo = col
            elif any(k in col_l for k in ["component", "componente", "culpa", "fault"]):
                col_componente = col
            elif any(k in col_l for k in ["descrição", "descricao", "desc", "falha", "defeito"]):
                col_desc = col

        termos_rapidos = [
            (r'\bc\.a\.\s*o\s*c\.c\.\s*a\s*vbatt\b', 'Circuito aberto ou curto-circuito ao positivo (Vbatt)'),
            (r'\bc\.a\.\s*o\s*c\.c\.\s*a\s*gnd\b', 'Circuito aberto ou curto-circuito ao aterramento'),
            (r'\bc\.c\.\s*a\s*vbatt\b', 'Curto-circuito ao positivo (Vbatt)'),
            (r'\bc\.c\.\s*a\s*gnd\b', 'Curto-circuito ao aterramento'),
            (r'\bcirc\.?\s*aperto\b', 'Circuito aberto'),
            (r'\bc\.a\.\b', 'Circuito aberto'),
            (r'\bc\.c\.\b', 'Curto-circuito'),
            (r'\bs\.c\.\s*to\s*vbat(?:t)?\b', 'Curto-circuito ao positivo (Vbatt)'),
            (r'\bs\.c\.\s*to\s*battery\b', 'Curto-circuito ao positivo da bateria'),
            (r'\bs\.c\.\s*to\s*gnd\b', 'Curto-circuito ao aterramento'),
            (r'\bs\.c\.\s*to\s*ground\b', 'Curto-circuito ao aterramento'),
            (r'\bopen\s*circuit\b', 'Circuito aberto'),
            (r'\bopen\s*load\b', 'Circuito aberto / carga aberta'),
            (r'\bShorted\s+to\s+high\s+source\b', 'Curto-circuito ao positivo (Tensão alta)'),
            (r'\bShorted\s+to\s+low\s+source\b', 'Curto-circuito ao aterramento (Tensão baixa)'),
            (r'\bDevice\s+error\b', 'Erro do dispositivo / Falha de componente'),
            (r'\bInvalid\s+signal\b', 'Sinal inválido / Fora da faixa'),
            (r'\bNo\s+CAN\s+messages\s+on\s+([A-Za-z0-9_\s#-]+?)\s+or\s+bus-off\s+detection\b', r'Sem mensagens CAN no \1 ou detecção de bus-off'),
            (r'\bNo\s+CAN\s+messages\s+from\s+([A-Za-z0-9_\s#-]+)\b', r'Sem mensagens CAN de \1'),
            (r'\bPhysical\s+memory\s+fault[\.\s]*Checksum\s+failure\s+is\s+detected\b', 'Falha de memória física. Falha de checksum detectada'),
            (r'\bDefect,\s*access\s+not\s+possible[\.\s]*Stepper\s+motor\s+failure\s+is\s+present\b', 'Defeito, acesso não possível. Falha presente no motor de passo'),
            (r'\bVoltage\s+above\s+normal\s*\(\s*s\.c\.\s+to\s+battery\s+or\s+open\s+circuit\s*\)', 'Tensão acima do normal (curto ao positivo ou circuito aberto)'),
            (r'\bVoltage\s+below\s+normal\s*\(\s*s\.c\.\s+to\s+ground\s*\)', 'Tensão abaixo do normal (curto ao aterramento)'),
            (r'\bDelta\s+voltage\s+below\s+normal\b', 'Diferença de tensão abaixo do normal'),
            (r'\bDelta\s+voltage\s+above\s+normal\b', 'Diferença de tensão acima do normal'),
            (r'\bTime/Date\s+from\s+TCO\s+not\s+available\b', 'Hora/Data do tacógrafo (TCO) indisponível'),
            (r'\bEngine\s+Starter\s+Mode\s+from\s+EDC\s+not\s+available\b', 'Modo de partida do motor do EDC indisponível'),
            (r'\bEngine\s+Oil\s+Temperature\s+from\s+EDC\s+not\s+available\b', 'Temperatura do óleo do motor do EDC indisponível'),
            (r'\bDate\s+could\s+not\s+be\s+stored\s+in\s+EEPROM\b', 'Data não pôde ser gravada na EEPROM'),
            (r'\bTime\s+could\s+not\s+be\s+stored\s+in\s+EEPROM\b', 'Hora não pôde ser gravada na EEPROM'),
            (r'\bOil\s+Level\s+could\s+not\s+be\s+stored\s+in\s+EEPROM\b', 'Nível de óleo não pôde ser gravado na EEPROM')
        ]

        registros_auditados = []
        correcoes_codigo = 0
        correcoes_fmi = 0
        correcoes_terra = 0
        resgates_descricao = 0
        resgates_componente = 0
        alucinacoes_fmi_eliminadas = 0

        for idx, reg in enumerate(registros):
            novo_reg = {}
            for col in colunas_desejadas:
                novo_reg[col] = reg.get(col, "-")

            reg_orig = registros_originais[idx] if (registros_originais and idx < len(registros_originais)) else {}

            # 0. Auditoria e Reconciliação Alfanumérica de Códigos (ex: 003015 -> 003D15)
            if col_codigo and col_codigo in novo_reg:
                val_cod = str(novo_reg[col_codigo]).strip()
                # Se o código veio vazio ou "-", resgata do registro original ou linha anterior
                if not val_cod or val_cod in ["-", "None", ""]:
                    if reg_orig.get("code"):
                        val_cod = str(reg_orig["code"]).strip()
                    elif idx > 0 and registros_auditados:
                        val_cod = str(registros_auditados[-1].get(col_codigo, "-")).strip()
                    novo_reg[col_codigo] = val_cod

                if val_cod and val_cod not in ["-", "None", ""]:
                    # Se o Ground Truth usa códigos hifenizados (ex: U1757-87) e o extrator separou o sufixo em fmi
                    orig_fmi_c = str(reg_orig.get("fmi", "")).strip().upper() if reg_orig else ""
                    if codigos_ground_truth and orig_fmi_c and orig_fmi_c not in ["-", "NONE", ""] and "-" not in val_cod:
                        cand_hif = f"{val_cod.upper()}-{orig_fmi_c.zfill(2)}"
                        if any(c.upper() == cand_hif for c in codigos_ground_truth):
                            val_cod = cand_hif
                            novo_reg[col_codigo] = val_cod
                            correcoes_codigo += 1

                    cod_corrigido, foi_corrigido = self.reconciliar_codigo(
                        cod=val_cod,
                        texto_fonte=texto_nativo,
                        codigos_ground_truth=codigos_ground_truth
                    )
                    if foi_corrigido:
                        novo_reg[col_codigo] = cod_corrigido
                        correcoes_codigo += 1

            # 1. Auditoria e Padronização de FMI com Fidelidade Absoluta ao Original
            if col_fmi and col_fmi in novo_reg:
                val_original = str(novo_reg[col_fmi]).strip()
                # Se o tradutor perdeu ou modificou o FMI, resgata o valor original do extrator
                orig_fmi = str(reg_orig.get("fmi", "")).strip() if reg_orig else ""
                if orig_fmi and orig_fmi not in ["-", "None", ""]:
                    val_original = orig_fmi
                val_corrigido = self.formatar_fmi_hex_2digitos(val_original)
                if val_original != val_corrigido:
                    correcoes_fmi += 1
                novo_reg[col_fmi] = val_corrigido

            # 2. Auditoria e Higienização / Resgate de Componente
            if col_componente and col_componente in novo_reg:
                comp_str = str(novo_reg[col_componente]).strip()

                # Se o componente veio vazio ou "-", resgata do registro original ou da descrição
                if comp_str in ["-", "", "None", "null", "nan"]:
                    orig_comp = str(reg_orig.get("component", "")).strip() if reg_orig else ""
                    if orig_comp and orig_comp not in ["-", "", "None", "null", "nan"]:
                        comp_str = self.extrair_e_traduzir_componente(orig_comp)
                        resgates_componente += 1
                    else:
                        orig_desc = str(reg_orig.get("description", "")).strip() if reg_orig else ""
                        if not orig_desc and col_desc and col_desc in novo_reg:
                            orig_desc = str(novo_reg[col_desc]).strip()
                        if orig_desc and orig_desc not in ["-", "", "None", "null", "nan"]:
                            comp_resgatado = self.extrair_e_traduzir_componente(orig_desc)
                            if comp_resgatado and comp_resgatado != "-":
                                comp_str = comp_resgatado
                                resgates_componente += 1
                        elif texto_bruto and col_codigo and col_codigo in novo_reg:
                            cod_alvo = str(novo_reg[col_codigo]).strip()
                            if cod_alvo and cod_alvo != "-":
                                m_linha = re.search(r'\b' + re.escape(cod_alvo) + r'\b[^\n]+', texto_bruto)
                                if m_linha:
                                    comp_resgatado = self.extrair_e_traduzir_componente(m_linha.group(0))
                                    if comp_resgatado and comp_resgatado != "-":
                                        comp_str = comp_resgatado
                                        resgates_componente += 1

                comp_limpo = re.sub(r'\s*(?:Low Side switch|High Side switch|Chave do lado).*$', '', comp_str, flags=re.I).strip()
                if comp_limpo and comp_limpo != comp_str:
                    comp_str = comp_limpo

                novo_reg[col_componente] = comp_str if comp_str else "-"

            # 3. Auditoria e Garantia de Preenchimento de FMI Description (NUNCA VAZIO)
            if col_fmi_desc and col_fmi_desc in novo_reg:
                val_fmi_desc = str(novo_reg.get(col_fmi_desc, "")).strip()
                orig_fmi_desc = str(reg_orig.get("fmi_description", "")).strip() if reg_orig else ""
                orig_desc = str(reg_orig.get("description", "")).strip() if reg_orig else ""
                val_fmi = str(novo_reg.get(col_fmi, reg_orig.get("fmi", ""))) if col_fmi else ""

                # Resgate 1: se novo_reg veio com '-' ou vazio, mas reg_orig tinha fmi_description
                if val_fmi_desc in ["-", "", "None", "null", "nan"]:
                    if orig_fmi_desc and orig_fmi_desc not in ["-", "", "None", "null", "nan"]:
                        val_fmi_desc = orig_fmi_desc
                        for pat, rep in termos_rapidos:
                            val_fmi_desc = re.sub(pat, rep, val_fmi_desc, flags=re.I)
                        resgates_descricao += 1

                # Resgate 2: se continuar vazio, aproveita a descrição do defeito presente no manual (Description)
                if val_fmi_desc in ["-", "", "None", "null", "nan"]:
                    if col_desc and novo_reg.get(col_desc) and str(novo_reg[col_desc]).strip() not in ["-", "", "None", "null", "nan"]:
                        val_fmi_desc = str(novo_reg[col_desc]).strip()
                        resgates_descricao += 1
                    elif orig_desc and orig_desc not in ["-", "", "None", "null", "nan"]:
                        val_fmi_desc = orig_desc
                        for pat, rep in termos_rapidos:
                            val_fmi_desc = re.sub(pat, rep, val_fmi_desc, flags=re.I)
                        resgates_descricao += 1

                # Resgate 3: se continuar vazio e temos um FMI conhecido, usa a norma técnica SAE J1939
                if val_fmi_desc in ["-", "", "None", "null", "nan"]:
                    desc_padrao = self.obter_descricao_padrao_fmi(val_fmi)
                    if desc_padrao:
                        val_fmi_desc = desc_padrao
                        resgates_descricao += 1

                # Interpretação inteligente se houver fórmula/faixa elétrica (ex: 6V<V<8V ou V>8V)
                tem_formula = bool(re.search(r'[<>=\d]+[VAWHz\s]+|[<>=\d]{2,}|(?:v\s*[<>]\s*\d+)', orig_fmi_desc or val_fmi_desc, re.I))
                if tem_formula:
                    val_fmi_desc = self.interpretar_fmi_formula(val_fmi, val_fmi_desc, formula_original=orig_fmi_desc)

                novo_reg[col_fmi_desc] = val_fmi_desc if val_fmi_desc else "-"

            # 4. Auditoria da Descrição Técnica Comum
            if col_desc and col_desc in novo_reg:
                val_d = str(novo_reg.get(col_desc, "")).strip()
                val_fmi = str(novo_reg.get(col_fmi, reg_orig.get("fmi", ""))) if col_fmi else ""
                if val_d in ["-", "", "None", "null", "nan"]:
                    orig_d = str(reg_orig.get("description", "")).strip() if reg_orig else ""
                    if orig_d and orig_d not in ["-", "None", "", "null", "nan"]:
                        traducao_rapida = orig_d
                        for pat, rep in termos_rapidos:
                            traducao_rapida = re.sub(pat, rep, traducao_rapida, flags=re.I)
                        novo_reg[col_desc] = traducao_rapida
                        resgates_descricao += 1
                    elif col_fmi_desc and novo_reg.get(col_fmi_desc) and str(novo_reg[col_fmi_desc]).strip() not in ["-", "", "None", "null", "nan"]:
                        novo_reg[col_desc] = str(novo_reg[col_fmi_desc]).strip()
                        resgates_descricao += 1
                    else:
                        desc_padrao = self.obter_descricao_padrao_fmi(val_fmi)
                        if desc_padrao:
                            novo_reg[col_desc] = desc_padrao
                            resgates_descricao += 1

            # 5. Auditoria Terminológica Rigorosa (Terra -> Aterramento)
            for k, v in novo_reg.items():
                if isinstance(v, str):
                    v_mod = v
                    v_mod = re.sub(r'\b(?:fio terra|cabo terra)\b', 'fio de aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ao terra / massa|ao terra|ao massa|à terra|à massa|com terra|com a terra)\b', 'ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:curto-circuito ao terra|curto ao terra|curto à terra|curto com terra|curto para terra)\b', 'curto-circuito ao aterramento', v_mod, flags=re.I)
                    v_mod = re.sub(r'\b(?:ligação ao terra|ligação terra|conexão terra|ponto de terra)\b', 'conexão de aterramento', v_mod, flags=re.I)
                    if v_mod != v:
                        correcoes_terra += 1
                    novo_reg[k] = v_mod

            registros_auditados.append(novo_reg)

        # 6. Consolidação de FMIs diferentes com a mesma descrição (ex: 04 e 06 -> 04 / 06)
        registros_auditados = self._consolidar_fmis_mesma_descricao(
            registros=registros_auditados,
            col_codigo=col_codigo,
            col_fmi=col_fmi,
            col_desc=col_fmi_desc or col_desc,
            col_componente=col_componente
        )

        relatorio = {
            "total_registros": len(registros_auditados),
            "correcoes_codigo": correcoes_codigo,
            "correcoes_fmi": correcoes_fmi,
            "correcoes_terra": correcoes_terra,
            "alucinacoes_fmi_eliminadas": alucinacoes_fmi_eliminadas,
            "resgates_descricao": resgates_descricao,
            "resgates_componente": resgates_componente,
            "status": "Auditado com Sucesso"
        }
        return registros_auditados, relatorio
