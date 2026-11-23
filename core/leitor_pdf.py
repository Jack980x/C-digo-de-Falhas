import pymupdf as fitz
import io
import re
from PIL import Image
from typing import List, Dict, Any, Tuple
from core.ocr_engine import OCREngine

class LeitorPDF:
    EXTENSOES_IMAGEM = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
    EXTENSOES_EXTRA = {".xlsx", ".csv", ".docx", ".html", ".htm", ".txt"}

    def __init__(self, ocr_engine: OCREngine = None):
        self.ocr_engine = ocr_engine or OCREngine()
        self._cache_paginas: Dict[Tuple[str, float, int, bool], Dict[str, Any]] = {}

    @classmethod
    def eh_arquivo_imagem(cls, caminho: str) -> bool:
        import os
        ext = os.path.splitext(str(caminho or ""))[1].lower()
        return ext in cls.EXTENSOES_IMAGEM

    @classmethod
    def eh_arquivo_extra(cls, caminho: str) -> bool:
        import os
        ext = os.path.splitext(str(caminho or ""))[1].lower()
        return ext in cls.EXTENSOES_EXTRA

    def _ler_arquivo_extra(self, caminho: str) -> List[Tuple[str, str]]:
        """
        Lê documentos tabulares ou textuais (.xlsx, .csv, .docx, .html, .txt) e retorna
        uma lista de tuplas (texto_pagina, metodo_usado), onde cada aba/seção equivale a uma página.
        """
        import os
        import csv
        import zipfile
        import xml.etree.ElementTree as ET
        ext = os.path.splitext(str(caminho or ""))[1].lower()

        if ext == ".xlsx":
            import openpyxl
            wb = openpyxl.load_workbook(caminho, data_only=True)
            paginas = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                linhas = []
                for row in ws.iter_rows(values_only=True):
                    celulas = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
                    if celulas:
                        linhas.append(" | ".join(celulas))
                if linhas:
                    paginas.append(("\n".join(linhas), f"Planilha Excel Nativa (Aba '{sheet_name}')"))
            wb.close()
            return paginas or [("", "Planilha Excel Vazia")]

        if ext == ".csv":
            linhas = []
            for enc in ["utf-8-sig", "utf-8", "latin-1"]:
                try:
                    with open(caminho, "r", encoding=enc, newline="") as f:
                        amostra = f.read(4096)
                        f.seek(0)
                        delim = ";" if amostra.count(";") > amostra.count(",") else ","
                        reader = csv.reader(f, delimiter=delim)
                        for row in reader:
                            celulas = [str(c).strip() for c in row if c is not None and str(c).strip() != ""]
                            if celulas:
                                linhas.append(" | ".join(celulas))
                    break
                except Exception:
                    continue
            return [("\n".join(linhas), "Tabela CSV Estruturada")]

        if ext == ".docx":
            linhas = []
            try:
                with zipfile.ZipFile(caminho, "r") as z:
                    xml_content = z.read("word/document.xml")
                tree = ET.fromstring(xml_content)
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                for elem in tree.iter():
                    tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                    if tag == "tr":
                        cels = []
                        for tc in elem.findall(".//w:tc", ns):
                            txt_c = "".join(t.text or "" for t in tc.findall(".//w:t", ns)).strip()
                            if txt_c:
                                cels.append(re.sub(r"\s+", " ", txt_c))
                        if cels:
                            linhas.append(" | ".join(cels))
                    elif tag == "p":
                        txt_p = "".join(t.text or "" for t in elem.findall(".//w:t", ns)).strip()
                        if txt_p and txt_p not in linhas[-3:]:
                            linhas.append(txt_p)
            except Exception:
                pass
            return [("\n".join(linhas), "Documento Word (.docx) Estruturado")]

        if ext in {".html", ".htm"}:
            conteudo = ""
            for enc in ["utf-8", "latin-1"]:
                try:
                    with open(caminho, "r", encoding=enc) as f:
                        conteudo = f.read()
                    break
                except Exception:
                    continue
            # Converte células de tabela HTML <td>/<th> em separador '|' e <tr>/<p>/<br> em quebras de linha
            limpo = re.sub(r"(?is)<(script|style).*?>.*?</\1>", "", conteudo)
            limpo = re.sub(r"(?is)</t[dh]>\s*<t[dh][^>]*>", " | ", limpo)
            limpo = re.sub(r"(?is)<(?:tr|/tr|p|/p|br|div|/div|li)[^>]*>", "\n", limpo)
            limpo = re.sub(r"(?is)<[^>]+>", " ", limpo)
            import html
            limpo = html.unescape(limpo)
            linhas = [re.sub(r"\s+", " ", l).strip(" |") for l in limpo.splitlines() if l.strip(" |")]
            return [("\n".join(linhas), "Página HTML Estruturada")]

        # .txt ou texto plano
        for enc in ["utf-8", "latin-1"]:
            try:
                with open(caminho, "r", encoding=enc) as f:
                    return [(f.read(), "Arquivo de Texto Nativo")]
            except Exception:
                continue
        return [("", "Arquivo de Texto")]

    @staticmethod
    def _normalizar_resolucao_imagem(img: Image.Image) -> Image.Image:
        """
        Converte a imagem para RGB mantendo sua resolução nativa para que o OCREngine
        aplique o fatiamento visual de alta precisão ou o redimensionamento interno por coluna.
        """
        return img.convert("RGB")

    def obter_total_paginas(self, caminho_pdf: str) -> int:
        if self.eh_arquivo_imagem(caminho_pdf):
            with Image.open(caminho_pdf) as img:
                return int(getattr(img, "n_frames", 1) or 1)
        if self.eh_arquivo_extra(caminho_pdf):
            return max(1, len(self._ler_arquivo_extra(caminho_pdf)))
        doc = fitz.open(caminho_pdf)
        total = len(doc)
        doc.close()
        return total

    def renderizar_pagina_como_imagem(self, pagina: fitz.Page, dpi: int = 300) -> Image.Image:
        """
        Renderiza a página do PDF em uma imagem PIL com alta resolução (300 DPI).
        Caso a página seja uma imagem única embutida (ex: captura de tela longa espremida em folha A4),
        extrai o stream original da imagem em sua resolução nativa máxima sem perda por reamostragem.
        """
        zoom = dpi / 72.0
        try:
            if len((pagina.get_text("text") or "").strip()) <= 80 and pagina.parent is not None:
                imgs = pagina.get_images(full=True)
                if len(imgs) == 1:
                    xref = imgs[0][0]
                    w_nat, h_nat = int(imgs[0][2] or 0), int(imgs[0][3] or 0)
                    info_list = pagina.get_image_info()
                    w_render = w_nat
                    if info_list and info_list[0].get("bbox"):
                        bbox = info_list[0]["bbox"]
                        w_render = int(max(1.0, (float(bbox[2]) - float(bbox[0])) * zoom))
                    if w_nat > w_render * 1.15 or h_nat > int(pagina.rect.height * zoom * 1.15):
                        base_img = pagina.parent.extract_image(xref)
                        if base_img and base_img.get("image"):
                            return Image.open(io.BytesIO(base_img["image"])).convert("RGB")
        except Exception:
            pass

        matriz = fitz.Matrix(zoom, zoom)
        pix = pagina.get_pixmap(matrix=matriz, alpha=False)
        img_bytes = pix.tobytes("png")
        return Image.open(io.BytesIO(img_bytes))

    def extrair_texto_geometrico_estruturado(self, pagina: fitz.Page) -> Tuple[str, List[str]]:
        """
        Extrai o texto vetorial nativo preservando a estrutura exata das tabelas:
        - Agrupa palavras por linhas usando proximidade de coordenada vertical (y0)
        - Deduplica palavras sobrepostas geradas por efeitos 'fake-bold' do PDF
        - Identifica a lista de códigos de falha originais (Ground Truth) da página
        """
        palavras = pagina.get_text("words")
        if not palavras:
            return "", []

        # Agrupa palavras por coordenada y com tolerância de 2.5 pontos
        import unicodedata
        linhas_y = {}
        for w in palavras:
            x0, y0, x1, y1, texto = w[0], w[1], w[2], w[3], unicodedata.normalize("NFKC", str(w[4]))
            y_chave = None
            for y_exist in linhas_y:
                if abs(y_exist - y0) < 2.5:
                    y_chave = y_exist
                    break
            if y_chave is None:
                y_chave = y0
                linhas_y[y_chave] = []
            linhas_y[y_chave].append((x0, texto))

        linhas_finais = []
        codigos_ground_truth = []

        for y in sorted(linhas_y.keys()):
            palavras_ordenadas = sorted(linhas_y[y], key=lambda item: item[0])

            # Deduplicação de palavras idênticas sobrepostas (sub-pixel)
            palavras_limpas = []
            for x, w in palavras_ordenadas:
                if palavras_limpas:
                    x_ant, w_ant = palavras_limpas[-1]
                    # Mesma palavra a menos de 4 pontos (efeito sombra / fake-bold)
                    if w == w_ant and abs(x - x_ant) < 4.0:
                        continue
                    if w == w_ant and len(w) >= 3 and abs(x - x_ant) < 15.0:
                        continue
                palavras_limpas.append((x, w))

            texto_linha = " ".join([p[1] for p in palavras_limpas]).strip()
            # Remove ruídos de botões de navegação de portais web (ex: 'MENU INÍCIO (http://www.cheryposve VOLTAR')
            texto_linha = re.sub(r'\b(?:MENU\s+)?IN[IÍ]CIO\s*\(https?://[^\s)]*(?:\s+VOLTAR)?', '', texto_linha, flags=re.I).strip()
            if texto_linha:
                linhas_finais.append(texto_linha)

            # Extração de códigos ground truth (padrões automotivos DTC/SPN/HEX)
            palavras_ignorar = {
                "FFFFFF", "000000", "YELLOW", "PAGE", "INPUT", "OUTPUT", "SWITCH",
                "ERROR", "MODULE", "SPEED", "BRAKE", "LEVEL", "RANGE", "SYSTEM",
                "DIGITAL", "ANALOG", "GROUND", "POWER", "BATTERY"
            }
            for idx_w, (_, w) in enumerate(palavras_limpas):
                # Ignora números entre parênteses no meio das descrições (ex: '(218)', '(055)')
                if w.startswith("(") and w.endswith(")"):
                    continue
                w_limpo = re.sub(r'^[^\w]+|[^\w]+$', '', w).upper()
                if w_limpo in palavras_ignorar or len(w_limpo) < 3 or len(w_limpo) > 8:
                    continue
                is_obd = bool(re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', w_limpo))
                # Números puros ou hexadecimais sem prefixo P/C/B/U só são códigos se estiverem nas primeiras posições da linha
                is_hex_num = idx_w <= 2 and bool(re.match(r'^[0-9A-F]{4,6}$', w_limpo) and any(c.isdigit() for c in w_limpo))
                is_spn_dec = idx_w <= 2 and bool(re.match(r'^[0-9]{3,6}$', w_limpo))

                if is_obd or is_hex_num or is_spn_dec:
                    if w_limpo not in codigos_ground_truth:
                        codigos_ground_truth.append(w_limpo)
                    break

        codigos_ground_truth = self._filtrar_falsos_positivos_gt(codigos_ground_truth)
        return "\n".join(linhas_finais), codigos_ground_truth

    @staticmethod
    def _filtrar_falsos_positivos_gt(codigos_gt: List[str]) -> List[str]:
        """
        Remove falsos positivos da lista Ground Truth:
        1. Fragmentos truncados de 5 caracteres (ex: 'B2A4F') quando o código completo de 7 caracteres ('B2A4F11')
           já existe na mesma página (gerado por espaço acidental no OCR da 3ª coluna: 'B2A4F 11').
        2. Códigos hexadecimais secundários/numéricos sem prefixo P/C/B/U (ex: 'E40088', '510116', '218', '055')
           quando a página é uma tabela OBD/UDS (códigos P/C/B/U).
        """
        if not codigos_gt:
            return []
        codigos_obd = [c for c in codigos_gt if re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', c)]
        codigos_obd_longos = [c for c in codigos_obd if len(c) in (7, 8)]
        pagina_eh_obd = len(codigos_obd) >= 2

        filtrados = []
        for c in codigos_gt:
            # Se a página tem códigos OBD de 7/8 caracteres e 'c' é um prefixo de 5 caracteres de algum deles (ex: B2A4F vs B2A4F11), descarta
            if len(c) == 5 and codigos_obd_longos and any(c7.startswith(c) for c7 in codigos_obd_longos):
                continue
            # Se a página é de códigos OBD (P/C/B/U), descarta códigos que não começam com P/C/B/U (como segunda coluna hex 'E40088'/'510116' ou IDs CAN '218')
            if pagina_eh_obd and not re.match(r'^[PCBU]', c):
                continue
            if c not in filtrados:
                filtrados.append(c)
        return filtrados

    @classmethod
    def extrair_codigos_gt_de_texto(cls, texto: str) -> List[str]:
        """
        Extrai a lista de códigos de falha Ground Truth a partir de um texto estruturado (vetorial ou OCR).
        Quando o texto está estruturado em colunas separadas por '|', prioriza a coluna de código (1ª coluna)
        para jamais capturar números ou fragmentos de texto dentro das colunas de descrição.
        """
        if not texto:
            return []
        palavras_ignorar = {
            "FFFFFF", "000000", "YELLOW", "PAGE", "INPUT", "OUTPUT", "SWITCH",
            "ERROR", "MODULE", "SPEED", "BRAKE", "LEVEL", "RANGE", "SYSTEM",
            "DIGITAL", "ANALOG", "GROUND", "POWER", "BATTERY"
        }
        codigos_gt = []
        for linha in texto.splitlines():
            linha_str = linha.strip()
            if not linha_str:
                continue

            if "|" in linha_str:
                colunas = [c.strip() for c in linha_str.split("|")]
                # Analisa prioritariamente a 1ª coluna (ou a 2ª se a 1ª for nome de módulo ECU)
                colunas_candidatas = colunas[:2] if len(colunas) >= 2 else colunas[:1]
                encontrou_na_linha = False
                for idx_col, col_txt in enumerate(colunas_candidatas):
                    for idx_tok, token in enumerate(col_txt.split()):
                        if token.startswith("(") and token.endswith(")"):
                            continue
                        w_limpo = re.sub(r'^[^\w]+|[^\w]+$', '', token).upper()
                        if w_limpo in palavras_ignorar or len(w_limpo) < 3 or len(w_limpo) > 8:
                            continue
                        is_obd = bool(re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', w_limpo))
                        is_hex_num = (idx_col == 0 and idx_tok == 0) and bool(re.match(r'^[0-9A-F]{4,6}$', w_limpo) and any(c.isdigit() for c in w_limpo))
                        is_spn_dec = (idx_col == 0 and idx_tok == 0) and bool(re.match(r'^[0-9]{3,6}$', w_limpo))
                        if is_obd or is_hex_num or is_spn_dec:
                            if w_limpo not in codigos_gt:
                                codigos_gt.append(w_limpo)
                            encontrou_na_linha = True
                            break
                    if encontrou_na_linha:
                        break
            else:
                tokens_linha = linha_str.split()
                for idx_tok, token in enumerate(tokens_linha):
                    if token.startswith("(") and token.endswith(")"):
                        continue
                    w_limpo = re.sub(r'^[^\w]+|[^\w]+$', '', token).upper()
                    if w_limpo in palavras_ignorar or len(w_limpo) < 3 or len(w_limpo) > 8:
                        continue
                    is_obd = bool(re.match(r'^[PCBU][0-9A-F]{4}(?:-[0-9A-F]{2}|[0-9A-F]{0,2})$', w_limpo))
                    is_hex_num = idx_tok <= 2 and bool(re.match(r'^[0-9A-F]{4,6}$', w_limpo) and any(c.isdigit() for c in w_limpo))
                    is_spn_dec = idx_tok <= 2 and bool(re.match(r'^[0-9]{3,6}$', w_limpo))
                    if is_obd or is_hex_num or is_spn_dec:
                        if w_limpo not in codigos_gt:
                            codigos_gt.append(w_limpo)
                        break

        return cls._filtrar_falsos_positivos_gt(codigos_gt)

    def detectar_necessidade_ocr(self, pagina: fitz.Page, texto_estruturado: str) -> Tuple[bool, str]:
        """
        Analisa automaticamente a estrutura interna da página do PDF para decidir se deve usar
        Extração Geométrica Vetorial ou OCR por Colunas (300 DPI):
        1. Sem camada de texto selecionável (<= 80 caracteres: imagem pura ou texto convertido em curvas/vetores)
        2. Camada de texto com codificação corrompida (caractere de substituição Unicode \ufffd)
        3. PDF escaneado com OCR antigo/oculto embutido (fontes invisíveis GlyphLessFont ou imagem cobrindo a página inteira)
        """
        txt_limpo = (texto_estruturado or "").strip()

        # 1. Ausência de camada de texto (PDF escaneado puro ou texto convertido em curvas/contornos)
        if len(txt_limpo) <= 80:
            return True, "PDF Rasterizado / Texto em Curvas (Auto-OCR)"

        # 2. Fonte com mapa Unicode corrompido (caracteres ilegíveis / quadrados)
        if "\ufffd" in txt_limpo:
            return True, "Camada de Texto Corrompida (Auto-OCR)"

        # 3. Fontes típicas de OCR invisível embutido por softwares de terceiros
        try:
            for font in pagina.get_fonts():
                nome_fonte = str(font[3] or "").lower()
                if any(k in nome_fonte for k in ["glyphless", "hiddenhorzocr", "invisible"]):
                    return True, "OCR Oculto Antigo Substituído (Auto-OCR)"
        except Exception:
            pass

        # 4. Imagem de fundo de página inteira (scan com camada de texto sobreposta)
        try:
            area_pag = max(1.0, float(pagina.rect.width * pagina.rect.height))
            for img in pagina.get_image_info():
                bbox = img.get("bbox")
                if bbox and len(bbox) == 4:
                    w_img = max(0.0, float(bbox[2]) - float(bbox[0]))
                    h_img = max(0.0, float(bbox[3]) - float(bbox[1]))
                    if (w_img * h_img) / area_pag >= 0.65 and len(pagina.get_drawings()) < 10:
                        return True, "Página Escaneada com Texto Sobreposto (Auto-OCR)"
        except Exception:
            pass

        return False, "Extração Geométrica Estruturada (100% Fiel)"

    def diagnosticar_modo_pdf(self, caminho_pdf: str, paginas: List[int] = None) -> Dict[str, Any]:
        """
        Inspeciona rapidamente as páginas selecionadas do PDF ou Imagem para informar automaticamente
        qual estratégia de leitura será utilizada em cada página (Vetorial Nativa vs OCR Automático).
        """
        if self.eh_arquivo_imagem(caminho_pdf):
            total = self.obter_total_paginas(caminho_pdf)
            pags = list(range(1, total + 1)) if not paginas else [p for p in paginas if 1 <= p <= total]
            return {
                "total_analisadas": len(pags),
                "paginas_ocr": pags,
                "paginas_vetorial": [],
                "motivos_ocr": ["Imagem de Tabela/Manual (Auto-OCR)"],
                "modo_predominante": "ocr"
            }

        if self.eh_arquivo_extra(caminho_pdf):
            total = self.obter_total_paginas(caminho_pdf)
            pags = list(range(1, total + 1)) if not paginas else [p for p in paginas if 1 <= p <= total]
            return {
                "total_analisadas": len(pags),
                "paginas_ocr": [],
                "paginas_vetorial": pags,
                "motivos_ocr": [],
                "modo_predominante": "vetorial"
            }

        doc = fitz.open(caminho_pdf)
        total = len(doc)
        indices = list(range(total)) if not paginas else [p - 1 for p in paginas if 1 <= p <= total]

        paginas_ocr = []
        paginas_vetorial = []
        motivos = set()

        for idx in indices:
            pag = doc[idx]
            texto_est, _ = self.extrair_texto_geometrico_estruturado(pag)
            precisa_ocr, motivo = self.detectar_necessidade_ocr(pag, texto_est)
            if precisa_ocr:
                paginas_ocr.append(idx + 1)
                motivos.add(motivo)
            else:
                paginas_vetorial.append(idx + 1)

        doc.close()
        return {
            "total_analisadas": len(indices),
            "paginas_ocr": paginas_ocr,
            "paginas_vetorial": paginas_vetorial,
            "motivos_ocr": sorted(list(motivos)),
            "modo_predominante": (
                "ocr" if len(paginas_ocr) == len(indices)
                else ("vetorial" if len(paginas_vetorial) == len(indices) else "hibrido")
            )
        }

    def extrair_conteudo_pagina(self, pagina: fitz.Page, forcar_ocr: bool = False) -> Tuple[str, str, List[str]]:
        """
        Extrai o conteúdo da página com detecção 100% automática de melhor método:
        1. Se houver texto vetorial digital íntegro (e OCR não for forçado):
           Utiliza a extração geométrica estruturada (100% fiel, sem erros de espaçamento ou caracteres duplicados).
        2. Se detectar automaticamente página rasterizada, texto em curvas, OCR antigo embutido ou se OCR for forçado:
           Executa Tesseract OCR a 300 DPI com fatiamento de tabelas por colunas e higienização hexadecimal.
        Retorna: (texto_estruturado, metodo_usado, codigos_ground_truth)
        """
        texto_estruturado, codigos_gt = self.extrair_texto_geometrico_estruturado(pagina)
        precisa_ocr_auto, motivo_modo = self.detectar_necessidade_ocr(pagina, texto_estruturado)

        # Se tem texto vetorial nativo íntegro e não foi forçado OCR manualmente
        if not precisa_ocr_auto and not forcar_ocr:
            return texto_estruturado, motivo_modo, codigos_gt

        # Executa OCR de alta precisão (por detecção automática ou por solicitação explícita)
        imagem = self.renderizar_pagina_como_imagem(pagina, dpi=300)
        texto_ocr = self.ocr_engine.extrair_texto(imagem)

        if len(texto_ocr.strip()) > 50:
            # Extrai Ground Truth do OCR sempre que a camada vetorial estava ausente ou foi substituída por OCR antigo/corrompido
            codigos_gt_ocr = self.extrair_codigos_gt_de_texto(texto_ocr)
            if precisa_ocr_auto or not codigos_gt:
                codigos_gt = codigos_gt_ocr
            else:
                for c in codigos_gt_ocr:
                    if c not in codigos_gt:
                        codigos_gt.append(c)
            rotulo_metodo = f"Tesseract OCR 300 DPI ({motivo_modo})" if precisa_ocr_auto else "Tesseract OCR 300 DPI (Forçado)"
            return texto_ocr, rotulo_metodo, codigos_gt

        return texto_estruturado, "Texto Estruturado (Fallback)", codigos_gt

    @staticmethod
    def mapear_codigos_secundarios_pagina(texto: str, codigos_gt: List[str]) -> Dict[str, str]:
        """
        Quando uma tabela possui duas colunas de código lado a lado na mesma linha
        (ex: 'U240088 E40088 Erro...' ou 'C110116 510116 Linha IGN...'),
        mapeia deterministicamente {codigo_1a_coluna: codigo_2a_coluna}.
        """
        if not texto or not codigos_gt:
            return {}
        gt_set = {c.upper() for c in codigos_gt if c}
        mapa = {}
        for linha in texto.splitlines():
            l_str = linha.strip()
            if not l_str:
                continue
            tokens = [t.strip() for t in (l_str.split("|") if "|" in l_str else l_str.split()) if t.strip()]
            if len(tokens) >= 2:
                t1 = re.sub(r'^[^\w-]+|[^\w-]+$', '', tokens[0]).upper()
                t2 = re.sub(r'^[^\w-]+|[^\w-]+$', '', tokens[1]).upper()
                if t1 in gt_set and re.match(r'^[0-9A-F]{4,6}$', t2) and any(ch.isdigit() for ch in t2):
                    mapa[t1] = t2
        return mapa

    @classmethod
    def detectar_estrutura_colunas_documento(cls, texto: str, codigos_gt: List[str] = None) -> Dict[str, Any]:
        """
        Detecta automaticamente os cabeçalhos originais da tabela no documento (PDF/Imagem/Excel)
        e identifica se há 2 colunas de código lado a lado ou coluna de FMI.
        """
        mapa_sec = cls.mapear_codigos_secundarios_pagina(texto, codigos_gt or [])
        tem_duas_colunas_cod = len(mapa_sec) >= 2

        col_cod_1 = "Código"
        col_cod_2 = "DTCs (Hex)"
        col_desc = "Descrição"
        tem_fmi = False
        colunas_detectadas = []

        if texto:
            # Verifica padrão CAOA Chery / Manuais com cabeçalho explícito
            if re.search(r'DTC\s+Designa[cç][aã]o\s+do\s+Ve[ií]culo', texto, re.I):
                col_cod_1 = "DTC Designação do Veículo"
                col_cod_2 = "DTCs"
                col_desc = "Descrição do DTC"
                colunas_detectadas = [col_cod_1, col_desc]
            elif re.search(r'\bDTC\s+Identifier\b', texto, re.I) and re.search(r'\bDTC\s+Name\b', texto, re.I):
                col_cod_1 = "DTC Identifier"
                col_desc = "DTC Name"
                colunas_detectadas = ["Código", "Descrição"]
            elif re.search(r'\bFMI\b', texto, re.I) and re.search(r'\b(?:SPN|Fault\s+Component|Component)\b', texto, re.I):
                tem_fmi = True
                colunas_detectadas = ["Códigos", "Componente", "FMI", "Descrição"]
            elif re.search(r'\bDTC\b', texto, re.I):
                col_cod_1 = "DTC"
                col_desc = "Descrição"
                colunas_detectadas = ["DTC", "Descrição"]

        if not colunas_detectadas:
            colunas_detectadas = ["Códigos", "Componente", "FMI", "Descrição"] if tem_fmi else [col_cod_1, col_desc]

        return {
            "colunas_sugeridas": colunas_detectadas,
            "tem_duas_colunas_codigo": tem_duas_colunas_cod,
            "col_codigo_1": col_cod_1,
            "col_codigo_2": col_cod_2,
            "col_descricao": col_desc,
            "tem_fmi": tem_fmi,
            "mapa_codigos_secundarios": mapa_sec
        }

    @staticmethod
    def contar_ocorrencias_codigos_pagina(texto: str, codigos_gt: List[str]) -> Dict[str, Any]:
        """
        Conta quantas linhas da tabela na página pertencem a cada código do Ground Truth,
        identificando códigos que aparecem repetidos em múltiplas linhas no próprio PDF
        (ex: uma linha para 'Chinese Standard' e outra com o mesmo código para 'European Standard').
        """
        if not texto or not codigos_gt:
            return {"total_linhas": len(codigos_gt or []), "qtd_repetidos": 0, "codigos_repetidos": []}

        gt_set = {c.upper() for c in codigos_gt}
        contagem = {c.upper(): 0 for c in codigos_gt}

        for linha in texto.splitlines():
            linha_str = linha.strip()
            if not linha_str:
                continue
            trecho_busca = linha_str.split("|")[0].strip() if "|" in linha_str else " ".join(linha_str.split()[:2])
            for token in trecho_busca.split():
                w_limpo = re.sub(r'^[^\w]+|[^\w]+$', '', token).upper()
                if w_limpo in gt_set:
                    contagem[w_limpo] = contagem.get(w_limpo, 0) + 1
                    break

        codigos_rep = [c for c in codigos_gt if contagem.get(c.upper(), 0) > 1]
        qtd_rep = sum(contagem[c.upper()] - 1 for c in codigos_rep)
        total_linhas = sum(max(1, contagem.get(c.upper(), 1)) for c in codigos_gt)
        return {
            "total_linhas": total_linhas,
            "qtd_repetidos": qtd_rep,
            "codigos_repetidos": codigos_rep
        }

    def ler_documento(self, caminho_pdf: str, paginas: List[int] = None, forcar_ocr: bool = False) -> List[Dict[str, Any]]:
        """
        Lê as páginas selecionadas (índices 1-based) de um PDF, Imagem ou Documento (.xlsx, .csv, .docx, .html)
        e retorna o conteúdo estruturado de cada uma, junto com a lista de códigos de falha Ground Truth.
        Utiliza cache em memória para evitar reprocessamento de OCR na mesma sessão.
        """
        import os
        caminho_abs = os.path.abspath(caminho_pdf)
        mtime = os.path.getmtime(caminho_abs) if os.path.exists(caminho_abs) else 0.0

        if self.eh_arquivo_extra(caminho_pdf):
            pags_extra = self._ler_arquivo_extra(caminho_pdf)
            total = len(pags_extra)
            indices = list(range(total)) if paginas is None else [p - 1 for p in paginas if 1 <= p <= total]
            resultados = []
            for idx in indices:
                num_pagina = idx + 1
                texto_ext, metodo_ext = pags_extra[idx]
                codigos_gt = self.extrair_codigos_gt_de_texto(texto_ext)
                info_rep = self.contar_ocorrencias_codigos_pagina(texto_ext, codigos_gt)
                mapa_sec = self.mapear_codigos_secundarios_pagina(texto_ext, codigos_gt)
                resultados.append({
                    "pagina": num_pagina,
                    "texto": texto_ext,
                    "texto_nativo": texto_ext,
                    "codigos_ground_truth": list(codigos_gt),
                    "mapa_codigos_secundarios": mapa_sec,
                    "info_repetidos": info_rep,
                    "metodo": metodo_ext
                })
            return resultados

        if self.eh_arquivo_imagem(caminho_pdf):
            resultados = []
            with Image.open(caminho_pdf) as img_doc:
                total = int(getattr(img_doc, "n_frames", 1) or 1)
                indices = list(range(total)) if paginas is None else [p - 1 for p in paginas if 1 <= p <= total]
                for idx in indices:
                    num_pagina = idx + 1
                    chave_cache = (caminho_abs, mtime, num_pagina, bool(forcar_ocr))
                    if chave_cache in self._cache_paginas:
                        resultados.append(dict(self._cache_paginas[chave_cache]))
                        continue
                    try:
                        img_doc.seek(idx)
                    except Exception:
                        pass
                    imagem_norm = self._normalizar_resolucao_imagem(img_doc)
                    texto_ocr = self.ocr_engine.extrair_texto(imagem_norm)
                    codigos_gt = self.extrair_codigos_gt_de_texto(texto_ocr)
                    info_rep = self.contar_ocorrencias_codigos_pagina(texto_ocr, codigos_gt)
                    mapa_sec = self.mapear_codigos_secundarios_pagina(texto_ocr, codigos_gt)
                    item_res = {
                        "pagina": num_pagina,
                        "texto": texto_ocr,
                        "texto_nativo": texto_ocr,
                        "codigos_ground_truth": list(codigos_gt),
                        "mapa_codigos_secundarios": mapa_sec,
                        "info_repetidos": info_rep,
                        "metodo": "Tesseract OCR 300 DPI (Imagem de Tabela/Manual)"
                    }
                    self._cache_paginas[chave_cache] = item_res
                    resultados.append(dict(item_res))
            return resultados

        doc = fitz.open(caminho_pdf)
        total = len(doc)

        if paginas is None:
            indices = list(range(total))
        else:
            indices = [p - 1 for p in paginas if 1 <= p <= total]

        resultados = []
        for idx in indices:
            num_pagina = idx + 1
            chave_cache = (caminho_abs, mtime, num_pagina, bool(forcar_ocr))
            if chave_cache in self._cache_paginas:
                resultados.append(dict(self._cache_paginas[chave_cache]))
                continue

            pag = doc[idx]
            texto, metodo, codigos_gt = self.extrair_conteudo_pagina(pag, forcar_ocr=forcar_ocr)
            texto_nativo = pag.get_text("text").strip() if not metodo.startswith("Tesseract OCR") else texto
            if not texto_nativo:
                texto_nativo = texto
            info_rep = self.contar_ocorrencias_codigos_pagina(texto, codigos_gt)
            mapa_sec = self.mapear_codigos_secundarios_pagina(texto, codigos_gt)
            item_res = {
                "pagina": num_pagina,
                "texto": texto,
                "texto_nativo": texto_nativo,
                "codigos_ground_truth": list(codigos_gt),
                "mapa_codigos_secundarios": mapa_sec,
                "info_repetidos": info_rep,
                "metodo": metodo
            }
            self._cache_paginas[chave_cache] = item_res
            resultados.append(dict(item_res))

        doc.close()
        return resultados

    def extrair_resumo_modulos_pdf(self, caminho_pdf: str) -> Dict[str, Any]:
        """
        Faz uma varredura rápida em todas as páginas do documento (PDF ou Imagem) para:
        1. Detectar possíveis módulos/centrais (colunas ECU Name / Module ou cabeçalhos de seção)
        2. Mapear em quais páginas cada módulo aparece
        3. Gerar um resumo estruturado leve para confirmação/identificação via IA
        """
        eh_nao_pdf = self.eh_arquivo_imagem(caminho_pdf) or self.eh_arquivo_extra(caminho_pdf)
        if eh_nao_pdf:
            paginas_lidas = self.ler_documento(caminho_pdf)
            total = len(paginas_lidas)
            doc = None
        else:
            doc = fitz.open(caminho_pdf)
            total = len(doc)
            paginas_lidas = []

        textos_por_pagina = {}
        codigos_gt_por_pagina = {}
        modulos_paginas = {}  # nome_upper -> {"nome": str, "paginas": set()}
        blocos_resumo = []

        palavras_banidas_modulo = {
            "ECU", "NAME", "SPN", "DTC", "FMI", "HEX", "DEC", "CODE", "CODES", "FAULT",
            "COMPONENT", "DESCRIPTION", "STATUS", "LAMP", "COMMENT", "PAGE", "LIST",
            "ERROR", "SYSTEM", "DIAGNOSIS", "PROCESS", "TABLE", "PIN", "SENSOR",
            "VALVE", "SWITCH", "OPEN", "SHORT", "CIRCUIT", "VOLTAGE", "SIGNAL",
            "HIGH", "LOW", "NORMAL", "ABOVE", "BELOW", "BATTERY", "GROUND", "GND",
            "VBAT", "VBATT", "CAN", "LIN", "PWM", "YELLOW", "RED", "OFF", "ON",
            "IDENTIFIER", "CONDITION", "CAUSE", "FAILURE", "MODULE", "CONTROLLER"
        }

        # Padrão 1: Prefixo de módulo na linha logo antes de código SPN/DTC (ex: "EM_SH 047221", "EDC7UC31 10031", "IC M06 41708")
        regex_linha_ecu_cod = re.compile(
            r'^([A-Z][A-Z0-9_/-]{1,12}(?:\s+[A-Z0-9]{2,5})?)\s+((?:[0-9A-F]{4,6})|(?:[PCBU][0-9A-F]{4,6}))\b'
        )
        # Padrão 2: Cabeçalho isolado de módulo em topo de página/seção (ex: "HCML", "HCMR", "EM_SH", "IC M06", "BCM2")
        regex_cabecalho_modulo = re.compile(
            r'^(?:MODULE|MÓDULO|ECU|SYSTEM|SISTEMA)?[\s:_-]*([A-Z]{2,6}[0-9_]{0,4}(?:\s+[A-Z0-9]{2,5})?)$'
        )
        # Siglas consagradas de módulos automotivos que podem ser citadas nas descrições
        siglas_modulos_conhecidas = {
            "ABS", "ACC", "ACM", "ADM", "AEB", "AFS", "AVH", "BCM", "BMS", "BSM",
            "CEM", "CCU", "CVT", "DDM", "DMC", "DMS", "DPF", "EBS", "ECM", "ECU",
            "EDC", "EGR", "EPB", "EPS", "ESC", "ESP", "ETM", "FCM", "FMCU", "FR",
            "GEM", "GW", "HCML", "HCMR", "HMI", "HUD", "HVAC", "IBC", "IC", "ICM",
            "ICU", "IPB", "IPC", "LDW", "LKA", "MCU", "OBC", "PAS", "PCM", "PDC",
            "PDM", "PSCM", "PSM", "RCM", "RMCU", "SAS", "SBC", "SCM", "SCR", "SDM",
            "SRS", "TBM", "TCM", "TCU", "TPMS", "UDS", "VCM", "VCU", "VIM", "ZBR"
        }
        modulos_citados = {}  # sigla -> set(paginas)
        letras_paginas = {}   # letra -> {"paginas": set(), "codigos": []}

        for idx in range(total):
            num_pag = idx + 1
            if eh_nao_pdf:
                texto_est = paginas_lidas[idx].get("texto", "")
                codigos_gt = paginas_lidas[idx].get("codigos_ground_truth", [])
            else:
                pag = doc[idx]
                texto_est, codigos_gt = self.extrair_texto_geometrico_estruturado(pag)
            textos_por_pagina[num_pag] = texto_est
            codigos_gt_por_pagina[num_pag] = codigos_gt

            # Mapeia letras iniciais dos códigos Ground Truth (ex: B, C, P, U)
            for cod_item in codigos_gt:
                cod_limpo = re.sub(r'^(?:0X|SPN|DTC|PSN)[\s:#_-]*', '', str(cod_item).strip(), flags=re.I).upper()
                if cod_limpo and cod_limpo[0].isalpha():
                    letra = cod_limpo[0]
                    if letra not in letras_paginas:
                        letras_paginas[letra] = {"paginas": set(), "codigos": []}
                    letras_paginas[letra]["paginas"].add(num_pag)
                    if cod_limpo not in letras_paginas[letra]["codigos"]:
                        letras_paginas[letra]["codigos"].append(cod_limpo)

            # Mapeia siglas de módulos citadas no texto da página
            for tok in re.findall(r'\b([A-Z]{2,5}[0-9]{0,2})\b', texto_est.upper()):
                if tok in siglas_modulos_conhecidas and tok not in {"DTC", "OBD", "UDS"}:
                    if tok not in modulos_citados:
                        modulos_citados[tok] = set()
                    modulos_citados[tok].add(num_pag)

            linhas = [l.strip() for l in texto_est.splitlines() if l.strip()]
            amostra_linhas = linhas[:18]

            for i_l, linha in enumerate(linhas):
                # Verifica padrão linha de tabela: MODULO CODIGO ...
                m_ecu = regex_linha_ecu_cod.match(linha)
                if m_ecu:
                    cand = m_ecu.group(1).strip()
                    cand_up = cand.upper()
                    partes_cand = cand_up.split()
                    if (
                        cand_up not in palavras_banidas_modulo
                        and not any(p in palavras_banidas_modulo for p in partes_cand)
                        and not re.match(r'^[0-9]+$', cand_up)
                        and not re.match(r'^[PCBU][0-9A-F]{4,6}$', cand_up)
                    ):
                        if cand_up not in modulos_paginas:
                            modulos_paginas[cand_up] = {"nome": cand, "paginas": set()}
                        modulos_paginas[cand_up]["paginas"].add(num_pag)

                # Verifica cabeçalhos curtos nas primeiras 10 linhas da página
                if i_l < 10 and len(linha) <= 20:
                    m_cab = regex_cabecalho_modulo.match(linha)
                    if m_cab:
                        cand = m_cab.group(1).strip()
                        cand_up = cand.upper()
                        partes_cand = cand_up.split()
                        if (
                            cand_up not in palavras_banidas_modulo
                            and not any(p in palavras_banidas_modulo for p in partes_cand)
                            and not cand_up.isdigit()
                            and cand_up not in codigos_gt
                        ):
                            if cand_up not in modulos_paginas:
                                modulos_paginas[cand_up] = {"nome": cand, "paginas": set()}
                            modulos_paginas[cand_up]["paginas"].add(num_pag)

            trecho_pag = "\n".join(amostra_linhas)
            blocos_resumo.append(f"--- PÁGINA {num_pag} ---\n{trecho_pag}")

        if doc is not None:
            doc.close()

        lista_regex = []
        for k, info_m in modulos_paginas.items():
            pags_ord = sorted(list(info_m["paginas"]))
            lista_regex.append({
                "nome": info_m["nome"],
                "paginas_lista": pags_ord,
                "paginas": self.formatar_intervalo_paginas(pags_ord),
                "descricao": "Módulo detectado na estrutura do PDF"
            })

        descricoes_letras = {
            "B": "Códigos B (Body / Carroceria)",
            "C": "Códigos C (Chassis / Chassi)",
            "P": "Códigos P (Powertrain / Motor e Transmissão)",
            "U": "Códigos U (Network / Rede e Comunicação)"
        }
        letras_detectadas = []
        for letra in sorted(letras_paginas.keys()):
            info_l = letras_paginas[letra]
            pags_ord = sorted(list(info_l["paginas"]))
            cods_l = info_l["codigos"]
            faixa = f"{cods_l[0]} a {cods_l[-1]}" if len(cods_l) >= 2 else (cods_l[0] if cods_l else "")
            desc_base = descricoes_letras.get(letra, f"Códigos iniciados com '{letra}'")
            letras_detectadas.append({
                "letra": letra,
                "paginas_lista": pags_ord,
                "paginas": self.formatar_intervalo_paginas(pags_ord),
                "qtd_codigos": len(cods_l),
                "faixa": faixa,
                "descricao": f"{desc_base} — {len(cods_l)} código(s) ({faixa})"
            })

        lista_citados = []
        for sigla in sorted(modulos_citados.keys()):
            pags_ord = sorted(list(modulos_citados[sigla]))
            lista_citados.append({
                "nome": sigla,
                "paginas_lista": pags_ord,
                "paginas": self.formatar_intervalo_paginas(pags_ord),
                "descricao": f"Citado nas descrições ({self.formatar_intervalo_paginas(pags_ord)})"
            })

        texto_amostra_doc = "\n".join(textos_por_pagina.values())
        todos_gt_doc = []
        for l_gt in codigos_gt_por_pagina.values():
            for c_item in l_gt:
                if c_item not in todos_gt_doc:
                    todos_gt_doc.append(c_item)
        estrutura_colunas = self.detectar_estrutura_colunas_documento(texto_amostra_doc, todos_gt_doc)

        return {
            "total_paginas": total,
            "resumo_texto": "\n\n".join(blocos_resumo)[:18000],
            "modulos_regex": lista_regex,
            "modulos_citados": lista_citados,
            "letras_detectadas": letras_detectadas,
            "estrutura_colunas": estrutura_colunas,
            "textos_por_pagina": textos_por_pagina,
            "codigos_gt_por_pagina": codigos_gt_por_pagina
        }

    @staticmethod
    def formatar_intervalo_paginas(paginas: List[int]) -> str:
        """Converte uma lista de páginas [1, 2, 3, 5] em string compacta 'Pág. 1-3, 5'."""
        if not paginas:
            return ""
        pags = sorted(set(paginas))
        intervalos = []
        ini = pags[0]
        fim = pags[0]
        for p in pags[1:]:
            if p == fim + 1:
                fim = p
            else:
                intervalos.append(f"{ini}-{fim}" if ini != fim else f"{ini}")
                ini = fim = p
        intervalos.append(f"{ini}-{fim}" if ini != fim else f"{ini}")
        return "Pág. " + ", ".join(intervalos)
