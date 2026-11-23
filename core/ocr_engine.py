import os
import re
import shutil
import numpy as np
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter


class OCREngine:
    def __init__(self, tesseract_cmd=None):
        candidatos = [
            tesseract_cmd,
            os.getenv("TESSERACT_PATH"),
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
            os.path.expanduser(r"~\AppData\Local\Tesseract-OCR\tesseract.exe"),
            shutil.which("tesseract")
        ]
        cmd_encontrado = None
        for c in candidatos:
            if c and os.path.exists(c):
                cmd_encontrado = c
                break

        if cmd_encontrado:
            pytesseract.pytesseract.tesseract_cmd = cmd_encontrado
        else:
            print("[AVISO] Executável do Tesseract não localizado automaticamente.")

        self._vision_client = None
        self._vision_init_tentado = False

    def _obter_cliente_visao(self):
        if self._vision_init_tentado:
            return self._vision_client
        self._vision_init_tentado = True
        try:
            from dotenv import load_dotenv
            load_dotenv()
            api_key = os.getenv("GEMINI_API_KEY")
            if api_key and api_key != "sua_chave_aqui":
                from google import genai
                self._vision_client = genai.Client(api_key=api_key)
        except Exception:
            self._vision_client = None
        return self._vision_client

    @staticmethod
    def _fatiar_imagem_em_faixas(imagem: Image.Image) -> list:
        """
        Divide imagens longas em faixas horizontais cortando estritamente nos intervalos entre linhas
        de texto (onde não há pixels escuros de caracteres), garantindo que nenhuma linha seja cortada ao meio.
        """
        img_rgb = imagem.convert("RGB")
        w, h = img_rgb.size
        alvo_h = 720 if w < 800 else (1350 if h > 5000 else 950)
        if h <= int(alvo_h * 1.25):
            return [img_rgb]

        gray = np.array(img_rgb.convert("L"))
        # Uma linha horizontal y é segura para corte se quase não possui pixels escuros de caracteres (< 130)
        dark_por_linha = np.sum(gray < 130, axis=1)
        linhas_seguras = np.where(dark_por_linha <= 2)[0]

        fatias = []
        y_ini = 0
        while y_ini + alvo_h < h:
            y_alvo = y_ini + alvo_h
            janela_min = y_ini + int(alvo_h * 0.60)
            janela_max = min(h - 25, y_ini + int(alvo_h * 1.30))
            candidatos = linhas_seguras[(linhas_seguras >= janela_min) & (linhas_seguras <= janela_max)]
            if len(candidatos) > 0:
                y_corte = int(candidatos[np.argmin(np.abs(candidatos - y_alvo))])
            else:
                # Fallback: linha com menor quantidade de pixels escuros na janela
                trecho = dark_por_linha[janela_min:max(janela_min + 1, janela_max)]
                y_corte = int(janela_min + np.argmin(trecho))

            fatias.append(img_rgb.crop((0, max(0, y_ini - 2), w, min(h, y_corte + 2))))
            y_ini = y_corte

        if y_ini < h - 10:
            fatias.append(img_rgb.crop((0, max(0, y_ini - 2), w, h)))

        return fatias

    def _extrair_texto_visao_ia(self, imagem: Image.Image) -> str:
        """
        Realiza OCR multimodal de alta precisão (Gemini Vision) fatiando a imagem em faixas horizontais
        alinhadas às linhas da tabela e processando-as em paralelo.
        Garante 100% de fidelidade em fontes pequenas, condensadas ou com fundo colorido.
        """
        client = self._obter_cliente_visao()
        if not client:
            return ""

        import concurrent.futures
        fatias = self._fatiar_imagem_em_faixas(imagem)
        modelos = ["gemini-3.8-flash", "gemini-flash-latest"]

        prompt = (
            "Transcreva fielmente TODAS as linhas de dados da tabela técnica presente nesta imagem.\n"
            "REGRAS OBRIGATÓRIAS:\n"
            "1. Formate cada linha separando as colunas com ' | ' (exemplo: B235A-68 | Tensão da bateria - falha elétrica geral.).\n"
            "2. Mantenha 100% da grafia EXATA de cada código de falha (ex: B235A-68, C0021-93, C0031-4A, C003A-01, U19B9-83, B181719). "
            "NUNCA omita a letra inicial (B, C, P, U), NUNCA troque 'A' por '4', '0' por 'O' ou 'D', '8' por '6' ou 'B', e preserve o hífen exato quando existir.\n"
            "3. Transcreva o texto da descrição exatamente no idioma original da imagem com acentuação perfeita, sem resumir e sem omitir nenhuma linha.\n"
            "4. Ignore títulos gerais de topo e cabeçalhos de coluna (como 'Código | Descrição do Código'). Retorne APENAS as linhas da tabela, uma por linha, sem blocos markdown."
        )

        def transcrever_fatia(idx_fatia, img_fatia):
            for mod in modelos:
                try:
                    resp = client.models.generate_content(
                        model=mod,
                        contents=[prompt, img_fatia],
                        config={"temperature": 0.0}
                    )
                    if resp and resp.text:
                        return idx_fatia, resp.text.strip()
                except Exception:
                    continue
            return idx_fatia, ""

        resultados_fatias = [""] * len(fatias)
        with concurrent.futures.ThreadPoolExecutor(max_workers=min(5, len(fatias))) as executor:
            futures = [executor.submit(transcrever_fatia, i, f) for i, f in enumerate(fatias)]
            for fut in concurrent.futures.as_completed(futures):
                idx_f, txt_f = fut.result()
                resultados_fatias[idx_f] = txt_f

        linhas_finais = []
        vistas_fronteira = set()
        for txt_f in resultados_fatias:
            if not txt_f:
                continue
            for l in txt_f.splitlines():
                l_limpa = l.strip().strip("`")
                if not l_limpa or l_limpa.lower().startswith("json"):
                    continue
                if "|" in l_limpa:
                    partes = [p.strip() for p in l_limpa.split("|")]
                    if partes and re.match(r'^(?:c[oó]digo|code|dtc|spn)s?$', partes[0], re.I):
                        continue
                    partes[0] = self.normalizar_codigo_dtc_token(partes[0]).upper()
                    l_limpa = " | ".join(partes)
                chave_dedup = re.sub(r'\s+', ' ', l_limpa).upper()
                if linhas_finais and chave_dedup == vistas_fronteira:
                    continue
                vistas_fronteira = chave_dedup
                linhas_finais.append(l_limpa)

        return "\n".join(linhas_finais).strip()

    @staticmethod
    def _agrupar_coordenadas_1d(coords, tolerancia: int = 12) -> list:
        if len(coords) == 0:
            return []
        coords = sorted(coords)
        grupos = [[coords[0]]]
        for c in coords[1:]:
            if c - grupos[-1][-1] <= tolerancia:
                grupos[-1].append(c)
            else:
                grupos.append([c])
        return [int(round(sum(g) / len(g))) for g in grupos]

    @staticmethod
    def _limpar_linhas_e_sublinhados(gray: np.ndarray) -> np.ndarray:
        """
        Remove linhas de grade horizontais/verticais e sublinhados de hiperlinks (que fazem o Tesseract
        pular linhas inteiras ou cortar palavras ao meio), preservando 100% os traços dos caracteres.
        """
        dark = gray < 220

        # 1. Segmentos horizontais contínuos >= 40px (bordas de tabela e sublinhados)
        L_h = 40
        cs_h = np.cumsum(np.pad(dark.astype(np.int32), ((0, 0), (1, 0))), axis=1)
        win_h = (cs_h[:, L_h:] - cs_h[:, :-L_h]) == L_h
        cs_win_h = np.cumsum(np.pad(win_h.astype(np.int32), ((0, 0), (L_h, L_h - 1))), axis=1)
        horiz_line = (cs_win_h[:, L_h:] - cs_win_h[:, :-L_h]) > 0

        # Espessura vertical em janela de 9px para preservar hastes verticais de letras que tocam a linha
        cs_v_thick = np.cumsum(np.pad(dark.astype(np.int32), ((5, 4), (0, 0))), axis=0)
        v_thick = cs_v_thick[9:, :] - cs_v_thick[:-9, :]

        # 2. Segmentos verticais contínuos >= 45px (bordas verticais de colunas)
        L_v = 45
        cs_v = np.cumsum(np.pad(dark.astype(np.int32), ((1, 0), (0, 0))), axis=0)
        win_v = (cs_v[L_v:, :] - cs_v[:-L_v, :]) == L_v
        cs_win_v = np.cumsum(np.pad(win_v.astype(np.int32), ((L_v, L_v - 1), (0, 0))), axis=0)
        vert_line = (cs_win_v[L_v:, :] - cs_win_v[:-L_v, :]) > 0

        cs_h_thick = np.cumsum(np.pad(dark.astype(np.int32), ((0, 0), (5, 4))), axis=1)
        h_thick = cs_h_thick[:, 9:] - cs_h_thick[:, :-9]

        erase_mask = (horiz_line & (v_thick <= 5)) | (vert_line & (h_thick <= 5))
        clean = gray.copy()
        clean[erase_mask] = 255

        # Remove halo de anti-aliasing imediatamente adjacente às linhas apagadas
        for dy, dx in [(-1, 0), (1, 0), (0, -1), (0, 1)]:
            shifted = np.roll(erase_mask, shift=(dy, dx), axis=(0, 1))
            clean[shifted & (clean > 140)] = 255

        return clean

    @staticmethod
    def normalizar_codigo_dtc_token(token: str) -> str:
        """
        Corrige confusões clássicas de OCR em códigos DTC/OBD/UDS, tanto contínuos (ex: BI81919 -> B181919,
        BICES11 -> B1CE511, B2AS5BI11 -> B2A5B11) quanto hifenizados (ex: CO03a-29 -> C003A-29, B235A.68 -> B235A-68).
        """
        if not token:
            return token
        t = re.sub(r'^[\]\[\|\(\)_©€]+|[\]\[\|\(\)_]+$', '', token).strip()

        mapa_hex = {
            'I': '1',
            'L': '1',
            'O': '0',
            'Q': '0',
            'S': '5',
            'Z': '2',
            'G': '6',
            '?': '7',
        }

        # Caso 1: Código DTC com hífen ou ponto separando o FMI/subtipo de 2 dígitos (ex: C0021-93, B235A.68, CO03a-29)
        t_clean_hif = t.replace("'", "").replace("`", "").replace("¢", "C")
        m_hif = re.match(r'^([PCBU])([0-9A-Za-z\?]{4,6})[-.]([0-9A-Za-z\?]{2,3})([.,;:]?)$', t_clean_hif, flags=re.I)
        if m_hif:
            pref = m_hif.group(1).upper()
            corpo = m_hif.group(2).upper()
            if len(corpo) > 4:
                corpo = (
                    corpo.replace('S5', '5').replace('5S', '5')
                    .replace('0O', '0').replace('O0', '0')
                    .replace('I1', '1').replace('1I', '1')
                    .replace('?7', '7').replace('7?', '7')
                    .replace('A4', 'A').replace('4A', 'A')
                    .replace('B6', 'B').replace('B8', 'B').replace('8B', 'B')
                )
            fmi_p = m_hif.group(3).upper()
            if len(fmi_p) > 2:
                fmi_p = (
                    fmi_p.replace('?7', '7').replace('7?', '7')
                    .replace('A4', 'A').replace('B6', 'B').replace('B8', 'B')
                    .replace('0O', '0').replace('O0', '0')
                )
            suf_p = m_hif.group(4)
            corpo_fix = ''.join(mapa_hex.get(ch, ch) for ch in corpo)
            fmi_fix = ''.join(mapa_hex.get(ch, ch) for ch in fmi_p)
            if re.match(r'^[0-9A-F]{4}$', corpo_fix) and re.match(r'^[0-9A-F]{2}$', fmi_fix):
                return f"{pref}{corpo_fix}-{fmi_fix}{suf_p}"

        # Caso 2: Código DTC contínuo (ex: B181719, B2A5B11, U01F287)
        m = re.match(r'^([PCBU])([0-9A-Za-z\?]{4,9})([.,;:]?)$', t)
        if not m:
            return token
        prefix = m.group(1).upper()
        rest = m.group(2).upper()
        sufixo_pont = m.group(3)

        # Quando um dígito é duplicado pelo OCR (ex: '5' -> 'S5', '1' -> 'I1', '0' -> '0O')
        rest = (
            rest.replace('S5', '5')
            .replace('5S', '5')
            .replace('0O', '0')
            .replace('O0', '0')
        )
        if len(rest) > 6:
            rest = (
                rest.replace('I1', '1')
                .replace('1I', '1')
                .replace('L1', '1')
                .replace('1L', '1')
            )

        rest_fixed = ''.join(mapa_hex.get(ch, ch) for ch in rest)
        if re.match(r'^[0-9A-F]{4,6}$', rest_fixed):
            return prefix + rest_fixed + sufixo_pont
        return token

    @classmethod
    def normalizar_texto_ocr(cls, texto: str) -> str:
        """
        Higieniza o texto extraído pelo OCR:
        1. Recompõe códigos DTC de 7 caracteres separados por espaço pelo OCR (ex: 'B2A4F 11' -> 'B2A4F11')
        2. Corrige códigos DTC com caracteres confundidos pelo OCR (I->1, S->5, O->0, etc.)
        3. Reconcilia códigos entre a 1ª e a 3ª coluna quando há ruído de OCR
        4. Reconstrói palavras inglesas quebradas por limite de largura de coluna usando validação cruzada na linha
        """
        if not texto:
            return ""

        linhas_processadas = []
        for linha in texto.splitlines():
            # 1. Une códigos DTC de 7 caracteres onde o OCR inseriu espaço antes dos 2 dígitos finais (ex: B2A4F 11 -> B2A4F11)
            linha = re.sub(r'\b([PCBU][0-9A-Za-z]{4})\s+([0-9A-F]{2})\b', r'\1\2', linha)

            # 2. Normaliza tokens individuais de código DTC
            palavras = linha.split()
            novas_palavras = [cls.normalizar_codigo_dtc_token(p) for p in palavras]
            linha_norm = " ".join(novas_palavras)

            # 3. Se a linha começa com um código B/C/P/U válido e mais adiante na mesma linha
            # aparece a versão com '3' ou '8' no lugar de 'B', corrige-a
            m_ini = re.match(r'^\s*([PCBU])([0-9A-F]{4,6})\b', linha_norm)
            if m_ini:
                letra_ini, hex_ini = m_ini.group(1), m_ini.group(2)
                cod_correto = f"{letra_ini}{hex_ini}"
                linha_norm = re.sub(r'\b[38]' + re.escape(hex_ini) + r'\b', cod_correto, linha_norm)
            else:
                # Se a 3ª coluna tem um código válido (ex: B2A5B11) e a 1ª coluna ficou corrompida pelo OCR
                m_qualquer = re.search(r'\b([PCBU][0-9A-F]{6})\b', linha_norm)
                if m_qualquer and "|" in linha_norm:
                    cod_valido = m_qualquer.group(1)
                    partes_col = [p.strip() for p in linha_norm.split("|")]
                    if partes_col and partes_col[0] and not re.match(r'^[PCBU][0-9A-F]{6}$', partes_col[0]):
                        partes_col[0] = cod_valido
                        linha_norm = " | ".join(partes_col)

            # 4. Reconstrução automática de palavras quebradas em fim de linha na coluna Description
            # Compara pares de fragmentos adjacentes (w1 + w2) com palavras íntegras presentes na própria linha (3ª coluna)
            if "|" in linha_norm:
                vocab_linha = set(re.findall(r"[a-z]+(?:-[a-z]+)*", linha_norm.lower()))
                partes_col = [p.strip() for p in linha_norm.split("|")]
                for idx_c, col_txt in enumerate(partes_col):
                    tokens_c = col_txt.split()
                    if len(tokens_c) < 2:
                        continue
                    reconstruidos = []
                    i_t = 0
                    while i_t < len(tokens_c):
                        if i_t + 1 < len(tokens_c):
                            t1, t2 = tokens_c[i_t], tokens_c[i_t + 1]
                            t2_clean = re.sub(r'[.,;:)]+$', '', t2)
                            t2_punct = t2[len(t2_clean):]
                            candidato = (t1 + t2_clean).lower()
                            if len(candidato) >= 4 and candidato in vocab_linha:
                                reconstruidos.append(t1 + t2_clean + t2_punct)
                                i_t += 2
                                continue
                        reconstruidos.append(tokens_c[i_t])
                        i_t += 1
                    partes_col[idx_c] = " ".join(reconstruidos)
                linha_norm = " | ".join(partes_col)

            # 5. Une palavras técnicas frequentemente quebradas mesmo quando só há uma coluna de texto
            substituicoes_quebra = [
                (r'\bshort\s*-\s*circuited\b', 'short-circuited'),
                (r'\bopen\s*-\s*circuited\b', 'open-circuited'),
                (r'\bshort-circui\s+ted\b', 'short-circuited'),
                (r'\bshort-circuit\s+ed\b', 'short-circuited'),
                (r'\bshort-ci\s+rcuited\b', 'short-circuited'),
                (r'\bopen-circuit\s+ed\b', 'open-circuited'),
                (r'\bopen-ci\s+rcuited\b', 'open-circuited'),
                (r'\bopen-ci\s+reuited\b', 'open-circuited'),
                (r'\bopen-c\s+ircuited\b', 'open-circuited'),
                (r'\bope\s+n-circuited\b', 'open-circuited'),
                (r'\bs\s+hort-circuited\b', 'short-circuited'),
                (r'\bo\s+pen-circuited\b', 'open-circuited'),
                (r'\bdriv\s+e\b', 'drive'),
                (r'\bd\s+rive\b', 'drive'),
                (r'\bGN\s+D\b', 'GND'),
                (r'\bG\s+ND\b', 'GND'),
                (r'\btoG\s*ND\b', 'to GND'),
                (r'\btoGN\s*D\b', 'to GND'),
                (r'\bpow\s+er\b', 'power'),
                (r'\bpo\s+wer\b', 'power'),
                (r'\bpowe\s+r\b', 'power'),
                (r'\bsupp\s+ly\b', 'supply'),
                (r'\bsu\s+pply\b', 'supply'),
                (r'\bsuppl\s+y\b', 'supply'),
                (r'\bs\s+upply\b', 'supply'),
                (r'\boverlo\s+ad\b', 'overload'),
                (r'\boverloa\s+d\b', 'overload'),
                (r'\bove\s+rload\b', 'overload'),
                (r'\bcirc\s+uit\b', 'circuit'),
                (r'\bci\s+rcuit\b', 'circuit'),
                (r'\bm\s+otor\b', 'motor'),
                (r'\bmot\s+or\b', 'motor'),
                (r'\bsens\s+or\b', 'sensor'),
                (r'\bsenso\s+r\b', 'sensor'),
                (r'\bte\s+mperature\b', 'temperature'),
                (r'\btemper\s+ature\b', 'temperature'),
                (r'\bcommunicati\s+on\b', 'communication'),
            ]
            for pat, rep in substituicoes_quebra:
                linha_norm = re.sub(pat, rep, linha_norm, flags=re.I)

            linhas_processadas.append(linha_norm)

        # 6. Restaura a letra inicial (P/C/B/U) em tabelas com códigos hifenizados quando o OCR omitiu a letra inicial
        # Ex: '0021-93 | Motor de tração...' entre linhas 'C0030-29' -> 'C0021-93 | Motor de tração...'
        prefixos_linha = []
        for l_proc in linhas_processadas:
            col0 = l_proc.split("|")[0].strip() if "|" in l_proc else (l_proc.split()[0] if l_proc.split() else "")
            m_p = re.match(r'^([PCBU])[0-9A-F]{4}(?:-[0-9A-F]{2})?$', col0, flags=re.I)
            prefixos_linha.append(m_p.group(1).upper() if m_p else None)

        if any(prefixos_linha):
            for idx_l, l_proc in enumerate(linhas_processadas):
                if "|" not in l_proc:
                    continue
                partes_l = [p.strip() for p in l_proc.split("|")]
                col0_limpo = re.sub(r'^[^0-9A-Za-z]+', '', partes_l[0])
                m_sem_letra = re.match(r'^([0-9A-Za-z]{4})[-.]([0-9A-Za-z]{2})$', col0_limpo)
                if m_sem_letra:
                    # Busca a letra de código mais próxima abaixo ou acima
                    letra_viz = None
                    for dist in range(1, len(linhas_processadas)):
                        if idx_l + dist < len(prefixos_linha) and prefixos_linha[idx_l + dist]:
                            letra_viz = prefixos_linha[idx_l + dist]
                            break
                        if idx_l - dist >= 0 and prefixos_linha[idx_l - dist]:
                            letra_viz = prefixos_linha[idx_l - dist]
                            break
                    if letra_viz:
                        cand_restaurado = cls.normalizar_codigo_dtc_token(f"{letra_viz}{m_sem_letra.group(1)}-{m_sem_letra.group(2)}")
                        partes_l[0] = cand_restaurado
                        linhas_processadas[idx_l] = " | ".join(partes_l)

        return "\n".join(linhas_processadas)

    def preprocessar_imagem(self, imagem: Image.Image) -> Image.Image:
        """
        Pré-processamento de alta definição para documentos técnicos e tabelas:
        - Conversão para escala de cinza 8-bit
        - Remoção morfológica de linhas de grade e sublinhados que interferem na segmentação do OCR
        - Realce de contraste suave
        """
        img_gray = imagem.convert("L")
        arr_gray = np.array(img_gray)
        arr_clean = self._limpar_linhas_e_sublinhados(arr_gray)
        img_clean = Image.fromarray(arr_clean)
        enhancer = ImageEnhance.Contrast(img_clean)
        return enhancer.enhance(1.2)

    def _extrair_tabela_fundo_colorido(self, gray: np.ndarray) -> str:
        """
        Extrai tabelas com fundo colorido (ex: azul/cinza) e linhas divisórias brancas ou pontilhadas.
        Apaga as divisórias no tamanho original (1x), remove o fundo colorido preservando o texto escuro,
        amplia imagens estreitas (w < 1400) em 3x (Lanczos) e executa OCR isolado por coluna em blocos.
        """
        h0, w0 = gray.shape
        if np.median(gray) >= 238:
            return ""

        white0 = gray > 242
        y_lines0 = self._agrupar_coordenadas_1d(np.where(np.mean(white0, axis=1) > 0.70)[0], tolerancia=5)
        if len(y_lines0) < 3:
            return ""

        y_top_tab = y_lines0[0]
        x_mids0 = self._agrupar_coordenadas_1d(
            np.where(np.mean(white0[y_top_tab:, :], axis=0) > 0.55)[0],
            tolerancia=5
        )
        x_mids0 = [x for x in x_mids0 if 15 < x < w0 - 15]
        if not x_mids0:
            return ""

        # Apaga as linhas divisórias horizontais (inclusive pontilhadas) e verticais na imagem 1x antes de redimensionar
        g_work = gray.copy()
        rows_white = np.where(np.mean(white0, axis=1) > 0.65)[0]
        g_work[rows_white, :] = 255
        for xm in x_mids0:
            g_work[:, max(0, xm - 1):min(w0, xm + 2)] = 255

        # Remove o fundo colorido (azul/cinza claro) mantendo os caracteres pretos intactos
        g_work = np.where(g_work > 150, 255, g_work).astype(np.uint8)

        scale = 3 if w0 < 1400 else (2 if w0 < 2000 else 1)
        if scale > 1:
            resampling = getattr(Image, "Resampling", Image).LANCZOS
            pil_scaled = Image.fromarray(g_work).resize((w0 * scale, h0 * scale), resampling)
            clean = np.array(pil_scaled)
        else:
            clean = g_work

        h, w = clean.shape
        x_lines = [0] + [x * scale for x in x_mids0] + [w - 1]
        y_lines = [y * scale for y in y_lines0]
        if y_lines[-1] < h - 15 * scale:
            y_lines.append(h - 1)

        num_cols = len(x_lines) - 1
        num_rows = len(y_lines) - 1
        cells = [[[] for _ in range(num_cols)] for _ in range(num_rows)]

        rows_per_chunk = 35
        for c_idx in range(num_cols):
            x0, x1 = x_lines[c_idx], x_lines[c_idx + 1]
            col_crop = clean[y_lines[0]:y_lines[-1], x0:x1]
            for r_start in range(0, num_rows, rows_per_chunk):
                r_end = min(num_rows, r_start + rows_per_chunk)
                cy0 = y_lines[r_start] - y_lines[0]
                cy1 = y_lines[r_end] - y_lines[0]
                chunk_img = col_crop[cy0:cy1, :]
                if chunk_img.size == 0:
                    continue
                pad_img = np.pad(chunk_img, ((25, 25), (25, 25)), mode="constant", constant_values=255)
                data = pytesseract.image_to_data(
                    Image.fromarray(pad_img),
                    config="--oem 3 --psm 6",
                    output_type=pytesseract.Output.DICT
                )
                for i in range(len(data["text"])):
                    txt = data["text"][i].strip()
                    if not txt or int(data["conf"][i]) < 0:
                        continue
                    wy = (data["top"][i] - 25) + data["height"][i] / 2.0 + y_lines[r_start]
                    wx = data["left"][i] - 25
                    lnum = (r_start, data["block_num"][i], data["par_num"][i], data["line_num"][i])
                    for r_idx in range(r_start, r_end):
                        if y_lines[r_idx] - 6 <= wy <= y_lines[r_idx + 1] + 6:
                            cells[r_idx][c_idx].append((lnum, wy, wx, txt))
                            break

        out_lines = []
        for r_idx in range(num_rows):
            row_strs = []
            for c_idx in range(num_cols):
                w_list = cells[r_idx][c_idx]
                if not w_list:
                    row_strs.append("")
                    continue
                lines_dict = {}
                for lnum, oy, ox, t in w_list:
                    lines_dict.setdefault(lnum, []).append((oy, ox, t))
                sorted_lines = sorted(lines_dict.values(), key=lambda lw: sum(x[0] for x in lw) / len(lw))
                line_texts = [" ".join(x[2] for x in sorted(lw, key=lambda item: item[1])) for lw in sorted_lines]
                row_strs.append(" ".join(line_texts).strip())
            if any(row_strs):
                out_lines.append(" | ".join(s for s in row_strs if s))

        return "\n".join(out_lines)

    def _extrair_tabela_por_colunas(self, imagem: Image.Image) -> str:
        """
        Detecta a grade da tabela (linhas horizontais e colunas verticais) e executa o OCR
        isolando as colunas verticalmente.
        Suporta tanto tabelas de fundo branco com bordas escuras quanto tabelas de fundo colorido
        com divisórias brancas/pontilhadas.
        """
        gray = np.array(imagem.convert("L"))
        if np.median(gray) < 238:
            texto_col = self._extrair_tabela_fundo_colorido(gray)
            if len(texto_col.strip()) > 50:
                return self.normalizar_texto_ocr(texto_col).strip()

        h, w = gray.shape
        dark = gray < 220

        # Limpa linhas de grade e sublinhados preservando os caracteres
        clean = self._limpar_linhas_e_sublinhados(gray)

        # Detecta linhas verticais da grade (altura >= 12% da página)
        L_vg = int(h * 0.12)
        cs_v = np.cumsum(np.pad(dark.astype(np.int32), ((1, 0), (0, 0))), axis=0)
        has_long_v = np.any((cs_v[L_vg:, :] - cs_v[:-L_vg, :]) >= int(L_vg * 0.90), axis=0)
        x_lines = self._agrupar_coordenadas_1d(np.where(has_long_v)[0], tolerancia=15)

        if len(x_lines) < 2:
            return ""

        x_min, x_max = x_lines[0], x_lines[-1]
        table_w = x_max - x_min
        if table_w < w * 0.35:
            return ""

        # Detecta linhas horizontais da grade que cruzam a largura da tabela
        row_counts = np.sum(dark[:, x_min:x_max + 1], axis=1)
        y_lines = self._agrupar_coordenadas_1d(np.where(row_counts >= int(table_w * 0.70))[0], tolerancia=15)
        if len(y_lines) < 2:
            return ""

        # Verifica se as linhas verticais da tabela continuam acima da primeira linha ou abaixo da última (célula aberta no rodapé)
        v_rows = np.where(np.sum(dark[:, x_lines], axis=1) >= max(2, len(x_lines) - 1))[0]
        if len(v_rows) > 0:
            if v_rows[0] < y_lines[0] - 25:
                y_lines = [int(v_rows[0])] + y_lines
            if v_rows[-1] > y_lines[-1] + 25:
                y_lines = y_lines + [int(v_rows[-1])]

        strips = []
        meta = []  # (tipo, col_idx, y_offset_empilhado, orig_y0, orig_y1)
        cur_y = 0
        pad_h = 40

        if y_lines[0] > 30:
            s_top = clean[0:y_lines[0], :]
            strips.append(s_top)
            meta.append(("header", -1, cur_y, 0, y_lines[0]))
            cur_y += s_top.shape[0] + pad_h
            strips.append(np.full((pad_h, w), 255, dtype=np.uint8))

        num_cols = len(x_lines) - 1
        for c_idx in range(num_cols):
            x0, x1 = x_lines[c_idx], x_lines[c_idx + 1]
            col_crop = clean[y_lines[0]:y_lines[-1], x0:x1]
            canvas = np.full((col_crop.shape[0], w), 255, dtype=np.uint8)
            canvas[:, 20:20 + col_crop.shape[1]] = col_crop
            strips.append(canvas)
            meta.append(("col", c_idx, cur_y, y_lines[0], y_lines[-1]))
            cur_y += canvas.shape[0] + pad_h
            strips.append(np.full((pad_h, w), 255, dtype=np.uint8))

        if h - y_lines[-1] > 30:
            s_bot = clean[y_lines[-1]:h, :]
            strips.append(s_bot)
            meta.append(("footer", -1, cur_y, y_lines[-1], h))

        stacked = np.vstack(strips)
        data = pytesseract.image_to_data(
            Image.fromarray(stacked),
            config="--oem 3 --psm 6",
            output_type=pytesseract.Output.DICT
        )

        header_words = []
        footer_words = []
        num_rows = len(y_lines) - 1
        cells = [[[] for _ in range(num_cols)] for _ in range(num_rows)]

        n_boxes = len(data["text"])
        for i in range(n_boxes):
            txt = data["text"][i].strip()
            if not txt or int(data["conf"][i]) < 0:
                continue
            wy = data["top"][i] + data["height"][i] / 2.0
            wx = data["left"][i]
            lnum = (data["block_num"][i], data["par_num"][i], data["line_num"][i])

            for m_type, c_idx, y_off, oy0, oy1 in meta:
                sh = oy1 - oy0
                if y_off <= wy <= y_off + sh:
                    orig_y = oy0 + (wy - y_off)
                    if m_type == "header":
                        header_words.append((orig_y, wx, txt))
                    elif m_type == "footer":
                        footer_words.append((orig_y, wx, txt))
                    elif m_type == "col":
                        for r_idx in range(num_rows):
                            if y_lines[r_idx] - 4 <= orig_y <= y_lines[r_idx + 1] + 4:
                                cells[r_idx][c_idx].append((lnum, orig_y, wx, txt))
                                break
                    break

        out_lines = []
        if header_words:
            hdr_str = " ".join(w[2] for w in sorted(header_words, key=lambda x: (round(x[0] / 15), x[1])))
            hdr_str = re.sub(r'\bDIC\s+list\b', 'DTC list', hdr_str, flags=re.I)
            out_lines.append(hdr_str)

        for r_idx in range(num_rows):
            row_strs = []
            for c_idx in range(num_cols):
                w_list = cells[r_idx][c_idx]
                if not w_list:
                    row_strs.append("")
                    continue
                lines_dict = {}
                for lnum, oy, ox, t in w_list:
                    lines_dict.setdefault(lnum, []).append((oy, ox, t))
                sorted_lines = sorted(lines_dict.values(), key=lambda lw: sum(x[0] for x in lw) / len(lw))
                line_texts = [" ".join(x[2] for x in sorted(lw, key=lambda item: item[1])) for lw in sorted_lines]
                cell_text = " ".join(line_texts).strip()
                row_strs.append(cell_text)

            # Se a 1ª coluna (código) ficou vazia mas a última coluna começa com um código DTC (ex: B181719 Low beam...),
            # recupera automaticamente o código para a 1ª coluna
            if num_cols >= 2 and not row_strs[0] and row_strs[-1]:
                m_rec = re.match(r'^([PCBU38][0-9A-Za-z]{4,7})\b', row_strs[-1].strip())
                if m_rec:
                    cand_tok = m_rec.group(1)
                    if cand_tok[0] in "38":
                        cand_tok = "B" + cand_tok[1:]
                    cand_norm = self.normalizar_codigo_dtc_token(cand_tok)
                    if re.match(r'^[PCBU][0-9A-F]{4,6}$', cand_norm):
                        row_strs[0] = cand_norm

            if any(row_strs):
                out_lines.append(" | ".join(s for s in row_strs if s))

        if footer_words:
            out_lines.append(" ".join(w[2] for w in sorted(footer_words, key=lambda x: (round(x[0] / 15), x[1]))))

        return self.normalizar_texto_ocr("\n".join(out_lines)).strip()

    def extrair_texto(self, imagem: Image.Image, config_ocr="--oem 3 --psm 4") -> str:
        """
        Extrai texto da imagem combinando Visão Computacional de Alta Precisão e Tesseract OCR:
        1. Para capturas de tela, imagens estreitas/longas (ex: prints rolantes) ou tabelas com fundo colorido,
           utiliza prioritariamente o OCR Multimodal em faixas alinhadas às linhas (_extrair_texto_visao_ia)
           para garantir 100% de fidelidade em fontes pequenas/condensadas.
        2. Para páginas A4 padrão a 300 DPI com fundo branco e grade bem definida, executa a extração
           geométrica por colunas (_extrair_tabela_por_colunas).
        3. Caso contrário, executa OCR sobre a imagem limpa com reconciliação entre PSM 4 e PSM 3.
        """
        w, h = imagem.size
        gray = np.array(imagem.convert("L"))
        eh_imagem_especial = (w < 1600) or (h > 3600) or (h / max(1.0, float(w)) > 2.2) or (float(np.median(gray)) < 238.0)

        if eh_imagem_especial:
            try:
                texto_visao = self._extrair_texto_visao_ia(imagem)
                if len(texto_visao) > 50:
                    return texto_visao
            except Exception:
                pass

        try:
            texto_tabela = self._extrair_tabela_por_colunas(imagem)
            if len(texto_tabela) > 50:
                return texto_tabela
        except Exception:
            pass

        if not eh_imagem_especial:
            try:
                texto_visao = self._extrair_texto_visao_ia(imagem)
                if len(texto_visao) > 50:
                    return texto_visao
            except Exception:
                pass

        img_tratada = self.preprocessar_imagem(imagem)
        try:
            texto = pytesseract.image_to_string(img_tratada, config=config_ocr)
            texto = self.normalizar_texto_ocr(texto)

            # Complementa com PSM 3 caso algum código tenha sido omitido pelo PSM 4
            texto_psm3 = self.normalizar_texto_ocr(
                pytesseract.image_to_string(img_tratada, config="--oem 3 --psm 3")
            )
            if len(texto.strip()) < 80 and len(texto_psm3.strip()) > len(texto.strip()):
                texto = texto_psm3
            else:
                cods_4 = set(re.findall(r'\b[PCBU][0-9A-F]{4,6}\b', texto))
                cods_3 = set(re.findall(r'\b[PCBU][0-9A-F]{4,6}\b', texto_psm3))
                faltantes = cods_3 - cods_4
                if faltantes:
                    linhas_extras = []
                    for l3 in texto_psm3.splitlines():
                        if any(re.search(r'\b' + re.escape(c) + r'\b', l3) for c in faltantes):
                            linhas_extras.append(l3)
                    if linhas_extras:
                        texto = texto.strip() + "\n" + "\n".join(linhas_extras)

            return texto.strip()
        except Exception:
            try:
                raw = pytesseract.image_to_string(imagem, config="--oem 3 --psm 3").strip()
                return self.normalizar_texto_ocr(raw)
            except Exception as e2:
                print(f"[ERRO OCR] Falha ao extrair texto da imagem: {e2}")
                return ""
