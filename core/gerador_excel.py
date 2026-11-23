import os
import re
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Optional
from core.agente_organizador import AgenteOrganizador

class GeradorExcel:
    TEMAS_VISUAIS = {
        "simplo_azul": {
            "nome": "Padrão Oficial SIMPLO (Azul Royal)",
            "tem_banner": True,
            "cor_banner": "0078D7",
            "cor_texto_banner": "FFFFFF",
            "cor_cabecalho": "0078D7",
            "cor_texto_cabecalho": "FFFFFF",
            "cor_col_codigo": "A4CCF0",
            "cor_col_texto": "D3E6F8",
            "cor_zebra_par": None,
            "cor_borda": "FFFFFF",
            "fonte": "Arial",
            "tam_fonte": 10,
            "autofiltro": False,
            "aba_indice": False
        }
    }

    def __init__(self, agente_organizador: Optional[AgenteOrganizador] = None):
        self.organizador = agente_organizador or AgenteOrganizador()

    @classmethod
    def inspecionar_modelo_excel(cls, caminho_xlsx: str) -> Dict[str, Any]:
        """
        Analisa uma planilha .xlsx de referência e extrai automaticamente seu padrão:
        - Colunas e larguras exatas de cada coluna
        - Presença ou não de banner superior (linhas 1 e 2 mescladas)
        - Cores hexadecimais de banner, cabeçalho e células de dados, além de fonte e bordas
        """
        wb = openpyxl.load_workbook(caminho_xlsx, data_only=True)
        ws = wb.active

        def extrair_cor_hex(cell, padrao_hex: str) -> str:
            try:
                fg = cell.fill.start_color.rgb if cell.fill and cell.fill.start_color else None
                if fg and isinstance(fg, str) and len(fg) >= 6:
                    c_hex = fg[-6:].upper()
                    if re.match(r'^[0-9A-F]{6}$', c_hex):
                        if c_hex != "000000" or (cell.fill.fill_type == "solid" and fg.upper().endswith("000000")):
                            return c_hex
            except Exception:
                pass
            return padrao_hex

        merges_r1 = [rng for rng in ws.merged_cells.ranges if rng.min_row == 1 and rng.max_row == 1 and rng.max_col > 1]
        tem_banner = len(merges_r1) > 0
        linha_cabecalho = 3 if tem_banner else 1
        linha_dados = linha_cabecalho + 1

        colunas = []
        larguras = {}
        for col_idx in range(1, ws.max_column + 1):
            val = ws.cell(row=linha_cabecalho, column=col_idx).value
            if val is not None and str(val).strip():
                nome_c = str(val).strip()
                colunas.append(nome_c)
                letra_c = get_column_letter(col_idx)
                w = ws.column_dimensions[letra_c].width
                larguras[nome_c] = round(float(w), 2) if w else 25.0

        cel_banner = ws.cell(row=1, column=1)
        cel_cab = ws.cell(row=linha_cabecalho, column=1)
        cel_cod = ws.cell(row=linha_dados, column=1)
        cel_txt = ws.cell(row=linha_dados, column=min(2, max(1, len(colunas))))

        cor_banner = extrair_cor_hex(cel_banner, "0078D7")
        cor_cab = extrair_cor_hex(cel_cab, cor_banner)
        cor_col_cod = extrair_cor_hex(cel_cod, "A4CCF0")
        cor_col_txt = extrair_cor_hex(cel_txt, "D3E6F8")

        fonte_nome = (cel_cod.font.name if cel_cod.font and cel_cod.font.name else "Arial")
        tam_fonte = int(cel_cod.font.size if cel_cod.font and cel_cod.font.size else 10)
        cor_txt_cab = "FFFFFF"
        try:
            if cel_cab.font and cel_cab.font.color and cel_cab.font.color.type == "rgb" and cel_cab.font.color.rgb:
                cand_hex = str(cel_cab.font.color.rgb)[-6:].upper()
                if re.match(r'^[0-9A-F]{6}$', cand_hex):
                    cor_txt_cab = cand_hex
        except Exception:
            pass

        wb.close()
        return {
            "colunas": colunas or ["Código", "Descrição"],
            "larguras": larguras,
            "tema_visual": {
                "nome": f"Modelo Importado ({os.path.basename(caminho_xlsx)})",
                "tem_banner": tem_banner,
                "cor_banner": cor_banner,
                "cor_texto_banner": cor_txt_cab,
                "cor_cabecalho": cor_cab,
                "cor_texto_cabecalho": cor_txt_cab,
                "cor_col_codigo": cor_col_cod,
                "cor_col_texto": cor_col_txt,
                "cor_zebra_par": None,
                "cor_borda": "FFFFFF" if cor_col_cod != "FFFFFF" else "000000",
                "fonte": fonte_nome,
                "tam_fonte": tam_fonte,
                "autofiltro": False,
                "aba_indice": False
            }
        }

    @staticmethod
    def extrair_chave_codigo(cod_str: Any) -> tuple:
        return AgenteOrganizador.extrair_chave_codigo(cod_str)

    @staticmethod
    def extrair_valores_fmi(fmi_str: Any) -> list:
        return AgenteOrganizador.extrair_chave_fmi(fmi_str)

    @staticmethod
    def classificar_tipo_coluna(col_name: str) -> str:
        """
        Classifica semanticamente a coluna para cálculo proporcional harmônico de larguras,
        alinhamento de texto e cores:
        - 'desc': Descrição, Nome do DTC, Localização, Significado (Texto longo/médio)
        - 'fmi_desc': Descrição do FMI
        - 'componente': Componente, Sensor, Atuador, Circuito, Subsistema
        - 'causa': Causa, Defeito, Condição
        - 'fmi': Código FMI puro (ex: 02, 0A)
        - 'modulo': Módulo, Central, ECU
        - 'status': DTC status, Lamp status, Severidade
        - 'codigo': Código base, SPN, DTC Identifier puro
        """
        c_low = str(col_name).lower().strip()
        # 1. Textos descritivos têm prioridade mesmo se contiverem 'dtc' (ex: 'DTC Name', 'DTC Description')
        if any(k in c_low for k in ["name", "nome", "desc", "descrição", "descricao", "localização", "localizacao", "texto", "significado"]):
            if "fmi" in c_low:
                return "fmi_desc"
            return "desc"
        # 2. Componentes / Circuitos
        if any(k in c_low for k in ["component", "componente", "culpa", "circuito", "sensor", "atuador", "sistema"]):
            return "componente"
        # 3. Causa / Defeito / Condição
        if any(k in c_low for k in ["causa", "cause", "defeito", "avaria", "condição", "condicao"]):
            return "causa"
        # 4. FMI puro
        if "fmi" in c_low:
            return "fmi"
        # 5. Módulo / ECU
        if any(k in c_low for k in ["módulo", "modulo", "ecu"]):
            return "modulo"
        # 6. Status / Lâmpada / Severidade
        if any(k in c_low for k in ["status", "lamp", "severidade", "state", "estado"]):
            return "status"
        # 7. Código / Identificador DTC puro
        if any(k in c_low for k in ["spn", "psn", "dtc", "código", "codigo", "code", "cod", "identifier", "identificador"]):
            return "codigo"
        return "geral"

    def _obter_caminho_disponivel(self, caminho: str) -> str:
        """
        Se o arquivo estiver aberto no Excel ou bloqueado pelo sistema,
        gera um nome com sufixo (ex: arquivo_1.xlsx) para não dar PermissionError.
        """
        if not os.path.exists(caminho):
            return caminho
            
        try:
            with open(caminho, "a+b"):
                return caminho
        except PermissionError:
            diretorio = os.path.dirname(caminho)
            nome, ext = os.path.splitext(os.path.basename(caminho))
            contador = 1
            while True:
                novo_caminho = os.path.join(diretorio, f"{nome}_{contador}{ext}")
                try:
                    if not os.path.exists(novo_caminho):
                        return novo_caminho
                    with open(novo_caminho, "a+b"):
                        return novo_caminho
                except PermissionError:
                    contador += 1

    @staticmethod
    def _sanitizar_nome_aba(nome: str, nomes_existentes: set) -> str:
        """
        Gera um nome de aba (worksheet) válido para o Excel (máx. 31 caracteres,
        sem caracteres proibidos \\ / ? * : [ ] e sem duplicidade).
        """
        limpo = re.sub(r'[\\/?*:\[\]]+', '-', str(nome or "Módulo")).strip(" '-")
        if not limpo:
            limpo = "Módulo"
        base = limpo[:31].strip()
        candidato = base
        contador = 2
        while candidato.upper() in {n.upper() for n in nomes_existentes}:
            sufixo = f"_{contador}"
            candidato = f"{base[:31 - len(sufixo)]}{sufixo}"
            contador += 1
        nomes_existentes.add(candidato)
        return candidato

    def _resolver_tema(self, tema_visual: Any = None) -> Dict[str, Any]:
        if isinstance(tema_visual, dict):
            base = dict(self.TEMAS_VISUAIS["simplo_azul"])
            base.update(tema_visual)
            return base
        if isinstance(tema_visual, str) and tema_visual in self.TEMAS_VISUAIS:
            return dict(self.TEMAS_VISUAIS[tema_visual])
        return dict(self.TEMAS_VISUAIS["simplo_azul"])

    def _preencher_aba_padrao(
        self,
        ws,
        dados: List[Dict[str, Any]],
        colunas_ordenadas: List[str],
        nome_modulo: str = None,
        larguras_personalizadas: Optional[Dict[str, float]] = None,
        caminho_saida: str = "",
        tema_visual: Any = None
    ):
        """Preenche uma aba (worksheet) com o layout estilizado da Planilha Padrão ou Tema Selecionado."""
        if tema_visual == "simplo_3_colunas_desc":
            return self._preencher_aba_3_colunas_descricao(
                ws=ws,
                dados=dados,
                nome_modulo=nome_modulo,
                larguras_personalizadas=larguras_personalizadas,
                caminho_saida=caminho_saida
            )

        tema = self._resolver_tema(tema_visual)
        tem_banner = bool(tema.get("tem_banner", True))
        cor_banner = str(tema.get("cor_banner", "0078D7")).lstrip("#")
        cor_txt_banner = str(tema.get("cor_texto_banner", "FFFFFF")).lstrip("#")
        cor_cab = str(tema.get("cor_cabecalho", cor_banner)).lstrip("#")
        cor_txt_cab = str(tema.get("cor_texto_cabecalho", "FFFFFF")).lstrip("#")
        cor_col_cod = str(tema.get("cor_col_codigo", "A4CCF0")).lstrip("#")
        cor_col_txt = str(tema.get("cor_col_texto", "D3E6F8")).lstrip("#")
        cor_zebra = tema.get("cor_zebra_par")
        cor_borda = str(tema.get("cor_borda", "FFFFFF")).lstrip("#")
        nome_fonte = str(tema.get("fonte", "Arial"))
        tam_fonte = int(tema.get("tam_fonte", 10))

        fill_banner = PatternFill(start_color=cor_banner, end_color=cor_banner, fill_type="solid")
        fill_cabecalho = PatternFill(start_color=cor_cab, end_color=cor_cab, fill_type="solid")
        fill_col_escura = PatternFill(start_color=cor_col_cod, end_color=cor_col_cod, fill_type="solid")
        fill_col_clara = PatternFill(start_color=cor_col_txt, end_color=cor_col_txt, fill_type="solid")
        fill_zebra_par = PatternFill(start_color=str(cor_zebra).lstrip("#"), end_color=str(cor_zebra).lstrip("#"), fill_type="solid") if cor_zebra else None

        font_banner = Font(name=nome_fonte, size=tam_fonte + 1, bold=True, color=cor_txt_banner)
        font_cabecalho_col = Font(name=nome_fonte, size=tam_fonte, bold=True, color=cor_txt_cab)
        font_corpo = Font(name=nome_fonte, size=tam_fonte, color="000000")

        borda_lado = Side(border_style="thin", color=cor_borda)
        borda_celula = Border(left=borda_lado, right=borda_lado, top=borda_lado, bottom=borda_lado)

        num_cols = len(colunas_ordenadas)

        # 1. Identifica nome das colunas-chave para ordenação e formatação via classificação semântica
        col_cod_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) == "codigo"), None)
        col_mod_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) == "modulo"), None)
        col_fmi_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) == "fmi"), None)
        col_desc_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) in ["desc", "fmi_desc"]), None)
        col_comp_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) == "componente"), None)

        # Ordena rigorosamente os dados antes de gerar o Excel via Agente 4 (Organizador):
        dados = self.organizador.organizar(dados, colunas=colunas_ordenadas)

        # Deduplica registros caso o usuário tenha extraído múltiplas páginas sem a coluna Módulo
        dados_dedup = []
        chaves_vistas = set()
        for item in dados:
            chave_item = tuple(str(item.get(c, "")).strip() for c in colunas_ordenadas)
            if chave_item not in chaves_vistas:
                chaves_vistas.add(chave_item)
                dados_dedup.append(item)
        dados = dados_dedup

        if not nome_modulo:
            nome_modulo = "SISTEMA ELETRÔNICO"
            if col_mod_nome:
                modulos = [str(d.get(col_mod_nome, "")).strip() for d in dados if d.get(col_mod_nome) and d.get(col_mod_nome) != "-"]
                if modulos:
                    nome_modulo = modulos[0]
            elif caminho_saida:
                base_nome = os.path.splitext(os.path.basename(caminho_saida))[0].replace("_traduzido", "").replace("_", " ")
                if base_nome:
                    nome_modulo = base_nome.upper()
        else:
            nome_modulo = nome_modulo.strip()

        codigos_encontrados = []
        if col_cod_nome:
            for d in dados:
                cod = str(d.get(col_cod_nome, "")).strip()
                if cod and cod != "-" and cod not in codigos_encontrados:
                    codigos_encontrados.append(cod)

        if len(codigos_encontrados) >= 2:
            intervalo_codigos = f"{codigos_encontrados[0]} - {codigos_encontrados[-1]}"
        elif len(codigos_encontrados) == 1:
            intervalo_codigos = codigos_encontrados[0]
        else:
            intervalo_codigos = "CÓDIGOS DE FALHAS"

        if tem_banner:
            # 2. Escreve Cabeçalho Superior de 2 linhas
            # Linha 1: Nome do Módulo
            ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=num_cols)
            cel1 = ws.cell(row=1, column=1, value=nome_modulo)
            cel1.font = font_banner
            cel1.fill = fill_banner
            cel1.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 26
            for c in range(1, num_cols + 1):
                ws.cell(row=1, column=c).border = borda_celula

            # Linha 2: Intervalo de Códigos
            ws.merge_cells(start_row=2, end_row=2, start_column=1, end_column=num_cols)
            cel2 = ws.cell(row=2, column=1, value=intervalo_codigos)
            cel2.font = font_banner
            cel2.fill = fill_banner
            cel2.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[2].height = 24
            for c in range(1, num_cols + 1):
                ws.cell(row=2, column=c).border = borda_celula

            linha_cabecalho = 3
        else:
            linha_cabecalho = 1

        # Cabeçalho das Colunas
        for c_idx, col_name in enumerate(colunas_ordenadas, 1):
            cel = ws.cell(row=linha_cabecalho, column=c_idx, value=col_name)
            cel.font = font_cabecalho_col
            cel.fill = fill_cabecalho
            cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cel.border = borda_celula
        ws.row_dimensions[linha_cabecalho].height = 28

        # 3. Larguras das colunas: Personalizadas pelo Usuário OU Automáticas Harmonizadas
        larguras_excel = {}

        if larguras_personalizadas:
            for idx, col_name in enumerate(colunas_ordenadas):
                if isinstance(larguras_personalizadas, dict) and col_name in larguras_personalizadas:
                    larguras_excel[col_name] = float(larguras_personalizadas[col_name])
                elif isinstance(larguras_personalizadas, (list, tuple)) and idx < len(larguras_personalizadas):
                    larguras_excel[col_name] = float(larguras_personalizadas[idx])
                else:
                    larguras_excel[col_name] = 25.0
        else:
            total_pixels_alvo = 850
            pixels_finais = {}

            tipos_presentes = set(self.classificar_tipo_coluna(c) for c in colunas_ordenadas)

            if num_cols == 2:
                # Planilha de 2 colunas (ex: Código + Nome/Descrição):
                col_cod = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) in ["codigo", "fmi", "modulo"]), None)
                col_txt = next((c for c in colunas_ordenadas if c != col_cod), None)
                if col_cod and col_txt:
                    pixels_finais[col_cod] = 170
                    pixels_finais[col_txt] = 680
                else:
                    pixels_finais[colunas_ordenadas[0]] = 425
                    pixels_finais[colunas_ordenadas[1]] = 425
            elif num_cols == 3:
                # Planilha de 3 colunas (ex: Código + Componente + Descrição ou Código + FMI + Descrição):
                pesos_3 = {
                    "codigo": 130,
                    "fmi": 90,
                    "modulo": 95,
                    "status": 80,
                    "componente": 250,
                    "causa": 210,
                    "fmi_desc": 280,
                    "desc": 380,
                    "geral": 180
                }
                soma_p = sum(pesos_3.get(self.classificar_tipo_coluna(c), 150) for c in colunas_ordenadas)
                fator = total_pixels_alvo / float(soma_p)
                acumulado = 0
                for idx, col_name in enumerate(colunas_ordenadas):
                    if idx == len(colunas_ordenadas) - 1:
                        px = total_pixels_alvo - acumulado
                    else:
                        tipo = self.classificar_tipo_coluna(col_name)
                        px = int(round(pesos_3.get(tipo, 150) * fator))
                        acumulado += px
                    pixels_finais[col_name] = max(px, 50)
            elif num_cols == 4 and tipos_presentes == {"codigo", "componente", "fmi", "desc"}:
                # Padrão oficial exato do material BC_IBC2 solicitado pelo usuário:
                # Códigos: 15.86 | Componente: 40.29 | FMI: 14.00 | Descrição: 48.57
                for col_name in colunas_ordenadas:
                    tipo = self.classificar_tipo_coluna(col_name)
                    if tipo == "codigo":
                        larguras_excel[col_name] = 15.86
                    elif tipo == "componente":
                        larguras_excel[col_name] = 40.29
                    elif tipo == "fmi":
                        larguras_excel[col_name] = 14.00
                    else:
                        larguras_excel[col_name] = 48.57
            else:
                # 4 ou mais colunas genéricas:
                pesos_base = {}
                for col_name in colunas_ordenadas:
                    tipo = self.classificar_tipo_coluna(col_name)
                    if tipo == "fmi":
                        pesos_base[col_name] = 100
                    elif tipo == "codigo":
                        pesos_base[col_name] = 115
                    elif tipo == "modulo":
                        pesos_base[col_name] = 90
                    elif tipo == "status":
                        pesos_base[col_name] = 75
                    elif tipo == "componente":
                        pesos_base[col_name] = 270
                    elif tipo == "fmi_desc":
                        pesos_base[col_name] = 270
                    elif tipo == "causa":
                        pesos_base[col_name] = 200
                    elif tipo == "desc":
                        pesos_base[col_name] = 320
                    else:
                        pesos_base[col_name] = 150

                soma_pesos = sum(pesos_base.values())
                fator = total_pixels_alvo / float(soma_pesos)

                acumulado = 0
                for idx, col_name in enumerate(colunas_ordenadas):
                    if idx == len(colunas_ordenadas) - 1:
                        px = total_pixels_alvo - acumulado
                    else:
                        px = int(round(pesos_base[col_name] * fator))
                        acumulado += px
                    pixels_finais[col_name] = max(px, 45)

            if not larguras_excel:
                # Ajuste fino da soma de pixels nominais
                diferenca = total_pixels_alvo - sum(pixels_finais.values())
                if diferenca != 0:
                    col_maior = max(pixels_finais, key=pixels_finais.get)
                    pixels_finais[col_maior] += diferenca

                # Conversão para largura do Excel
                for col_name, px in pixels_finais.items():
                    w = round(((px - 5) + 0.1) / 7.0, 3)
                    larguras_excel[col_name] = max(w, 5.0)

                pixels_renderizados = {col: int(w * 7) + 5 for col, w in larguras_excel.items()}
                diff_render = total_pixels_alvo - sum(pixels_renderizados.values())
                if diff_render != 0:
                    col_maior = max(larguras_excel, key=larguras_excel.get)
                    px_ajustado = pixels_renderizados[col_maior] + diff_render
                    larguras_excel[col_maior] = round(((px_ajustado - 5) + 0.1) / 7.0, 3)

        # Garante pixels finais para o cálculo preciso de altura das células e quebras de linha
        pixels_finais = {col: int(larguras_excel[col] * 7) + 5 for col in colunas_ordenadas}

        for c_idx, col_name in enumerate(colunas_ordenadas, 1):
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = larguras_excel[col_name]

        # 4. Mapeia cor de fundo de cada coluna conforme o modelo da imagem
        fill_por_coluna = {}
        for c_idx, col_name in enumerate(colunas_ordenadas, 1):
            tipo = self.classificar_tipo_coluna(col_name)
            if tipo in ["codigo", "fmi", "modulo", "status"]:
                # Colunas de identificadores: Azul pastel intermediário (A4CCF0)
                fill_por_coluna[c_idx] = fill_col_escura
            else:
                # Colunas de texto/descrição: Azul gelo claro (D3E6F8)
                fill_por_coluna[c_idx] = fill_col_clara

        # 5. Escreve dados da tabela a partir da linha seguinte ao cabeçalho
        linha_inicio_dados = linha_cabecalho + 1
        linha_atual = linha_inicio_dados
        grupos_linhas = []
        codigo_anterior = None
        inicio_grupo = linha_inicio_dados

        for idx_linha, item in enumerate(dados):
            cod_atual = str(item.get(col_cod_nome, "")).strip() if col_cod_nome else ""

            if cod_atual != codigo_anterior and codigo_anterior is not None:
                grupos_linhas.append((inicio_grupo, linha_atual - 1))
                inicio_grupo = linha_atual

            codigo_anterior = cod_atual

            max_linhas_linha = 1
            eh_par = (idx_linha % 2 == 1)

            for c_idx, col_name in enumerate(colunas_ordenadas, 1):
                val = item.get(col_name, "-")
                c_low = col_name.lower()
                if "fmi" in c_low and "desc" not in c_low and val and str(val).strip() not in ["-", ""]:
                    from core.agente_tradutor import AgenteAuditor
                    val = AgenteAuditor.formatar_fmi_hex_2digitos(val)

                cel = ws.cell(row=linha_atual, column=c_idx, value=str(val) if val is not None else "-")
                cel.font = font_corpo
                if fill_zebra_par and eh_par and self.classificar_tipo_coluna(col_name) not in ["codigo", "fmi", "modulo", "status"]:
                    cel.fill = fill_zebra_par
                else:
                    cel.fill = fill_por_coluna[c_idx]
                cel.border = borda_celula

                # Alinhamento e formato por tipo semântico de coluna
                tipo = self.classificar_tipo_coluna(col_name)
                if tipo in ["codigo", "fmi", "modulo", "status"]:
                    cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cel.number_format = "@"
                else:
                    cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)

                # Cálculo automático e dinâmico de quebras de linha para esta célula
                val_texto = str(val) if val is not None else ""
                px = pixels_finais.get(col_name, 150)
                # Capacidade de caracteres por linha com folga segura
                chars_por_linha = max(8, int((px - 14) / 7.2))
                linhas_paragrafos = val_texto.split("\n")
                linhas_celula = 0
                for p in linhas_paragrafos:
                    p_limpo = p.strip()
                    if p_limpo:
                        linhas_celula += max(1, math.ceil(len(p_limpo) / chars_por_linha))
                    else:
                        linhas_celula += 1
                max_linhas_linha = max(max_linhas_linha, linhas_celula)

            # Altura padronizada ajustada automaticamente para eliminar 100% de compressão de texto
            altura_linha = max(28.0, max_linhas_linha * 17.5 + 8.0)
            ws.row_dimensions[linha_atual].height = altura_linha

            linha_atual += 1

        if linha_atual > inicio_grupo:
            grupos_linhas.append((inicio_grupo, linha_atual - 1))

        # 6. Mesclagem Vertical: apenas quando há coluna FMI (para agrupar múltiplos FMIs do mesmo código)
        tem_fmi = any(self.classificar_tipo_coluna(c) == "fmi" for c in colunas_ordenadas)
        colunas_para_mesclar = []
        if tem_fmi:
            for c_idx, col_name in enumerate(colunas_ordenadas, 1):
                tipo = self.classificar_tipo_coluna(col_name)
                if tipo in ["codigo", "modulo", "componente"]:
                    colunas_para_mesclar.append(c_idx)

        if colunas_para_mesclar:
            for (r_ini, r_fim) in grupos_linhas:
                if r_fim > r_ini:
                    for c_idx in colunas_para_mesclar:
                        valores = [ws.cell(row=r, column=c_idx).value for r in range(r_ini, r_fim + 1)]
                        if len(set(valores)) <= 1 or valores[0] is not None:
                            ws.merge_cells(start_row=r_ini, end_row=r_fim, start_column=c_idx, end_column=c_idx)
                            cel_topo = ws.cell(row=r_ini, column=c_idx)
                            cel_topo.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                            for r in range(r_ini, r_fim + 1):
                                ws.cell(row=r, column=c_idx).border = borda_celula

        # 7. AutoFiltro opcional no cabeçalho (quando ativado no tema)
        if tema.get("autofiltro") and not tem_fmi and linha_atual > linha_inicio_dados:
            col_fim_letra = get_column_letter(num_cols)
            ws.auto_filter.ref = f"A{linha_cabecalho}:{col_fim_letra}{linha_atual - 1}"

        # 8. Congelar painel logo abaixo do cabeçalho
        ws.freeze_panes = f"A{linha_inicio_dados}"

    def _criar_aba_indice_geral(self, wb, resumo_abas: List[Dict[str, Any]], titulo_documento: str):
        """
        Cria uma 1ª aba 'Índice Geral' com hiperlinks clicáveis para cada aba gerada,
        mostrando faixa de códigos e quantidade por aba.
        """
        ws_idx = wb.create_sheet(title="Índice Geral", index=0)
        fill_topo = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        fill_cab = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
        fill_par = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
        fill_impar = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        borda_c = Border(
            left=Side(border_style="thin", color="CBD5E0"),
            right=Side(border_style="thin", color="CBD5E0"),
            top=Side(border_style="thin", color="CBD5E0"),
            bottom=Side(border_style="thin", color="CBD5E0")
        )

        ws_idx.merge_cells("A1:D1")
        c1 = ws_idx.cell(row=1, column=1, value=f"ÍNDICE GERAL — {titulo_documento.upper()}")
        c1.font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
        c1.fill = fill_topo
        c1.alignment = Alignment(horizontal="center", vertical="center")
        ws_idx.row_dimensions[1].height = 28

        cabecalhos = ["Aba / Família", "Faixa de Códigos", "Quantidade de Registros", "Atalho Rápido"]
        larguras = [22, 32, 26, 26]
        for idx_c, (nome_c, larg) in enumerate(zip(cabecalhos, larguras), 1):
            cel = ws_idx.cell(row=2, column=idx_c, value=nome_c)
            cel.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cel.fill = fill_cab
            cel.alignment = Alignment(horizontal="center", vertical="center")
            cel.border = borda_c
            ws_idx.column_dimensions[get_column_letter(idx_c)].width = larg
        ws_idx.row_dimensions[2].height = 24

        total_geral = 0
        for i_r, item in enumerate(resumo_abas, start=3):
            aba_t = item["aba"]
            faixa = item["faixa"]
            qtd = item["qtd"]
            total_geral += qtd
            fill_r = fill_par if i_r % 2 == 0 else fill_impar

            c_aba = ws_idx.cell(row=i_r, column=1, value=aba_t)
            c_faixa = ws_idx.cell(row=i_r, column=2, value=faixa)
            c_qtd = ws_idx.cell(row=i_r, column=3, value=qtd)
            c_link = ws_idx.cell(row=i_r, column=4, value=f"Ir para '{aba_t}' ➔")
            c_link.hyperlink = f"#'{aba_t}'!A1"

            for c_i, cel in enumerate([c_aba, c_faixa, c_qtd, c_link], 1):
                cel.fill = fill_r
                cel.border = borda_c
                cel.alignment = Alignment(horizontal="center", vertical="center")
                if c_i == 4:
                    cel.font = Font(name="Arial", size=10, bold=True, color="0066CC", underline="single")
                else:
                    cel.font = Font(name="Arial", size=10, bold=(c_i == 1), color="000000")
            ws_idx.row_dimensions[i_r].height = 24

        linha_tot = len(resumo_abas) + 3
        ws_idx.merge_cells(start_row=linha_tot, end_row=linha_tot, start_column=1, end_column=2)
        c_t1 = ws_idx.cell(row=linha_tot, column=1, value="TOTAL GERAL DO DOCUMENTO")
        c_t2 = ws_idx.cell(row=linha_tot, column=3, value=total_geral)
        c_t3 = ws_idx.cell(row=linha_tot, column=4, value="")
        for c_i in range(1, 5):
            cel = ws_idx.cell(row=linha_tot, column=c_i)
            cel.fill = fill_cab
            cel.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            cel.alignment = Alignment(horizontal="center", vertical="center")
            cel.border = borda_c
        ws_idx.row_dimensions[linha_tot].height = 25

    def exportar_para_excel(self, dados: List[Dict[str, Any]], colunas_ordenadas: List[str], caminho_saida: str, nome_modulo: str = None, larguras_personalizadas: Optional[Dict[str, float]] = None, tema_visual: Any = None) -> str:
        """
        Cria a planilha Excel estilizada no padrão selecionado (SIMPLO Azul Royal, Executivo Zebra, Grade Limpa ou Modelo Importado).
        """
        if not dados:
            raise ValueError("Não há dados extraídos para salvar no Excel.")

        os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
        caminho_saida = self._obter_caminho_disponivel(caminho_saida)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Códigos de Falhas"

        self._preencher_aba_padrao(
            ws=ws,
            dados=dados,
            colunas_ordenadas=colunas_ordenadas,
            nome_modulo=nome_modulo,
            larguras_personalizadas=larguras_personalizadas,
            caminho_saida=caminho_saida,
            tema_visual=tema_visual
        )

        wb.save(caminho_saida)
        return caminho_saida

    def exportar_multiplos_modulos_padrao(
        self,
        dados_por_modulo: Dict[str, List[Dict[str, Any]]],
        colunas_ordenadas: List[str],
        caminho_saida: str,
        larguras_personalizadas: Optional[Dict[str, float]] = None,
        titulo_banner_padrao: Optional[str] = None,
        tema_visual: Any = None
    ) -> str:
        """
        Cria um único arquivo Excel contendo uma página/aba independente para cada módulo ou letra selecionada,
        mantendo toda a estilização oficial da Planilha Padrão ou Tema Selecionado em cada aba.
        """
        modulos_validos = {m: d for m, d in dados_por_modulo.items() if d}
        if not modulos_validos:
            raise ValueError("Não há dados extraídos de nenhum módulo para salvar no Excel.")

        os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
        caminho_saida = self._obter_caminho_disponivel(caminho_saida)

        wb = openpyxl.Workbook()
        abas_usadas = set()
        resumo_abas = []
        col_cod_nome = next((c for c in colunas_ordenadas if self.classificar_tipo_coluna(c) == "codigo"), None)

        for idx, (nome_mod, dados_mod) in enumerate(modulos_validos.items()):
            titulo_aba = self._sanitizar_nome_aba(nome_mod, abas_usadas)
            if idx == 0:
                ws = wb.active
                ws.title = titulo_aba
            else:
                ws = wb.create_sheet(title=titulo_aba)

            banner_aba = titulo_banner_padrao if titulo_banner_padrao else nome_mod
            self._preencher_aba_padrao(
                ws=ws,
                dados=dados_mod,
                colunas_ordenadas=colunas_ordenadas,
                nome_modulo=banner_aba,
                larguras_personalizadas=larguras_personalizadas,
                caminho_saida=caminho_saida,
                tema_visual=tema_visual
            )
            cods_aba = [str(d.get(col_cod_nome, "")).strip() for d in dados_mod if col_cod_nome and d.get(col_cod_nome)]
            faixa_aba = f"{cods_aba[0]} a {cods_aba[-1]}" if len(cods_aba) >= 2 else (cods_aba[0] if cods_aba else "-")
            resumo_abas.append({"aba": titulo_aba, "faixa": faixa_aba, "qtd": len(dados_mod)})

        tema_res = self._resolver_tema(tema_visual)
        if tema_res.get("aba_indice") and len(resumo_abas) > 1:
            titulo_doc = titulo_banner_padrao or os.path.splitext(os.path.basename(caminho_saida))[0].replace("_traduzido", "")
            self._criar_aba_indice_geral(wb, resumo_abas, titulo_doc)

        wb.save(caminho_saida)
        return caminho_saida

    def _preencher_aba_buscar_codigo(
        self,
        ws,
        dados: List[Dict[str, Any]],
        nome_modulo: str = None,
        larguras_personalizadas: Optional[Dict[str, float]] = None
    ):
        """Preenche uma aba (worksheet) com o layout limpo e direto do Buscar Código."""
        colunas = ["Código", "Localização da Avaria", "Causa", "Descrição"]

        # Ordena agrupando cada código base e seus FMIs um abaixo do outro
        dados = AgenteOrganizador.organizar_buscar_codigo(dados)

        # Estilo básico e limpo (conforme imagem de referência)
        font_cabecalho = Font(name="Calibri", size=11, bold=True, color="000000")
        font_corpo = Font(name="Calibri", size=11, color="000000")

        borda_fina = Side(border_style="thin", color="000000")
        borda_celula = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

        # 1. Linha 1: Cabeçalho das Colunas
        for c_idx, col_name in enumerate(colunas, 1):
            cel = ws.cell(row=1, column=c_idx, value=col_name)
            cel.font = font_cabecalho
            cel.border = borda_celula
            cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 24

        # 2. Larguras generosas e limpas das colunas
        larguras = {
            "Código": 16,
            "Localização da Avaria": 48,
            "Causa": 35,
            "Descrição": 55
        }
        if larguras_personalizadas:
            for idx, col_name in enumerate(colunas):
                if isinstance(larguras_personalizadas, dict) and col_name in larguras_personalizadas:
                    larguras[col_name] = float(larguras_personalizadas[col_name])
                elif isinstance(larguras_personalizadas, (list, tuple)) and idx < len(larguras_personalizadas):
                    larguras[col_name] = float(larguras_personalizadas[idx])

        for c_idx, col_name in enumerate(colunas, 1):
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = larguras.get(col_name, 25)

        PADRAO_DESCRICAO_SIMPLO = (
            "Possíveis causas incluem falhas elétricas ou desgaste mecânico. "
            "Antes de substituir peças, realize testes detalhados e utilize "
            "diagramas elétricos e manuais técnicos SIMPLO para um diagnóstico preciso e confiável."
        )

        # 3. Escreve os registros a partir da Linha 2 (SEM mesclagem e SEM cores de fundo)
        linha_atual = 2
        for item in dados:
            max_linhas_linha = 1
            for c_idx, col_name in enumerate(colunas, 1):
                val = item.get(col_name, "-")
                if val is None or str(val).strip() == "":
                    val = "-"
                val_str = str(val).strip()

                if col_name == "Descrição":
                    val_low = val_str.lower()
                    if val_str in ["-", "", "None", "null", "nan"] or val_low.startswith("possíveis causas incluem") or "antes de substituir peças" in val_low:
                        causa_item = item.get("Causa", "-")
                        loc_item = item.get("Localização da Avaria", item.get("Componente", "-"))
                        val_str = AgenteOrganizador.sintetizar_descricao_modo_completo(causa_item, loc_item, val_str)

                cel = ws.cell(row=linha_atual, column=c_idx, value=val_str)
                cel.font = font_corpo
                cel.border = borda_celula

                if col_name == "Código":
                    cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cel.number_format = "@"
                else:
                    cel.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

                # Cálculo de altura dinâmica da linha para textos longos não ficarem cortados
                largura_col = larguras.get(col_name, 25)
                chars_por_linha = max(10, int(largura_col * 0.95))
                linhas_paragrafos = val_str.split("\n")
                linhas_celula = 0
                for p in linhas_paragrafos:
                    p_limpo = p.strip()
                    if p_limpo:
                        linhas_celula += max(1, math.ceil(len(p_limpo) / chars_por_linha))
                    else:
                        linhas_celula += 1
                max_linhas_linha = max(max_linhas_linha, linhas_celula)

            ws.row_dimensions[linha_atual].height = max(20.0, max_linhas_linha * 16.0 + 4.0)
            linha_atual += 1

        # 4. Congelar linha 1 (cabeçalho)
        ws.freeze_panes = "A2"

    def exportar_buscar_codigo(self, dados: List[Dict[str, Any]], caminho_saida: str, nome_modulo: str = None, larguras_personalizadas: Optional[Dict[str, float]] = None) -> str:
        """
        Cria a planilha Excel no formato mais básico, limpo e direto possível (seguindo a referência):
        - Inicia diretamente na Linha 1 com os cabeçalhos das 4 colunas:
          A1: Código | B1: Localização da Avaria | C1: Causa | D1: Descrição
        - Sem banner superior, sem cores de fundo, bordas pretas finas, sem mesclagem de células
        """
        if not dados:
            raise ValueError("Não há dados extraídos para salvar no Excel do Buscar Código.")

        os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
        caminho_saida = self._obter_caminho_disponivel(caminho_saida)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Buscar Código"

        self._preencher_aba_buscar_codigo(
            ws=ws,
            dados=dados,
            nome_modulo=nome_modulo,
            larguras_personalizadas=larguras_personalizadas
        )

        wb.save(caminho_saida)
        return caminho_saida

    def exportar_multiplos_modulos_buscar_codigo(
        self,
        dados_por_modulo: Dict[str, List[Dict[str, Any]]],
        caminho_saida: str,
        larguras_personalizadas: Optional[Dict[str, float]] = None
    ) -> str:
        """
        Cria um único arquivo Excel do Buscar Código contendo uma página/aba independente
        para cada módulo selecionado.
        """
        modulos_validos = {m: d for m, d in dados_por_modulo.items() if d}
        if not modulos_validos:
            raise ValueError("Não há dados extraídos de nenhum módulo para salvar no Excel do Buscar Código.")

        os.makedirs(os.path.dirname(os.path.abspath(caminho_saida)), exist_ok=True)
        caminho_saida = self._obter_caminho_disponivel(caminho_saida)

        wb = openpyxl.Workbook()
        abas_usadas = set()

        for idx, (nome_mod, dados_mod) in enumerate(modulos_validos.items()):
            titulo_aba = self._sanitizar_nome_aba(nome_mod, abas_usadas)
            if idx == 0:
                ws = wb.active
                ws.title = titulo_aba
            else:
                ws = wb.create_sheet(title=titulo_aba)

            self._preencher_aba_buscar_codigo(
                ws=ws,
                dados=dados_mod,
                nome_modulo=nome_mod,
                larguras_personalizadas=larguras_personalizadas
            )

        wb.save(caminho_saida)
        return caminho_saida

    def _preencher_aba_3_colunas_descricao(
        self,
        ws,
        dados: List[Dict[str, Any]],
        nome_modulo: str = None,
        larguras_personalizadas: Optional[Dict[str, float]] = None,
        caminho_saida: str = ""
    ):
        """
        Preenche uma aba no padrão oficial 'SIMPLO 3 Colunas + Descrição':
        - Mesmas informações do Buscar Código (Código, Causa, Localização da Avaria, Descrição)
        - Linhas 1 e 2: Banner Superior Azul Royal (#0078D7, Arial 11 Negrito Branco) com Nome do Módulo e Faixa de Códigos
        - Linha 3: Cabeçalho das 3 colunas: Código (16.57) | Causa (52.43) | Localização da Avaria (52.57)
        - Para cada código:
          * Linha superior (altura mín. 26.25): Código (#A4CCF0, Negrito) | Causa (#D3E6F8) | Localização da Avaria (#D3E6F8)
          * Linha inferior (altura mín. 45.0): 'Descrição' (#A4CCF0, Negrito) | Texto da Descrição mesclado nas colunas B:C (#D3E6F8)
          * Linha separadora entre blocos (altura 7.5): linha em branco (#FFFFFF)
        """
        dados = AgenteOrganizador.organizar_buscar_codigo(dados)

        fill_banner = PatternFill(start_color="0078D7", end_color="0078D7", fill_type="solid")
        fill_cod = PatternFill(start_color="A4CCF0", end_color="A4CCF0", fill_type="solid")
        fill_txt = PatternFill(start_color="D3E6F8", end_color="D3E6F8", fill_type="solid")
        fill_sep = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        font_banner = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        font_cabecalho = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        font_rotulo = Font(name="Arial", size=10, bold=True, color="000000")
        font_corpo = Font(name="Arial", size=10, bold=False, color="000000")

        borda_branca = Side(border_style="thin", color="FFFFFF")
        borda_celula = Border(left=borda_branca, right=borda_branca, top=borda_branca, bottom=borda_branca)

        if not nome_modulo:
            if caminho_saida:
                base_nome = os.path.splitext(os.path.basename(caminho_saida))[0].replace("_traduzido", "").replace("_", " ")
                nome_modulo = base_nome.upper() if base_nome else "MÓDULO DE DIAGNÓSTICO"
            else:
                nome_modulo = "MÓDULO DE DIAGNÓSTICO"
        else:
            nome_modulo = nome_modulo.strip()

        codigos_encontrados = []
        for d in dados:
            cod = str(d.get("Código") or d.get("Códigos") or d.get("DTC") or "").strip()
            if cod and cod != "-" and cod not in codigos_encontrados:
                codigos_encontrados.append(cod)

        if len(codigos_encontrados) >= 2:
            intervalo_codigos = f"{codigos_encontrados[0]} - {codigos_encontrados[-1]}"
        elif len(codigos_encontrados) == 1:
            intervalo_codigos = codigos_encontrados[0]
        else:
            intervalo_codigos = "CÓDIGOS DE FALHAS"

        # 1. Banner Superior (Linhas 1 e 2 mescladas em A:C)
        ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=3)
        cel1 = ws.cell(row=1, column=1, value=nome_modulo)
        cel1.font = font_banner
        cel1.fill = fill_banner
        cel1.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25.5
        for c in range(1, 4):
            ws.cell(row=1, column=c).fill = fill_banner
            ws.cell(row=1, column=c).border = borda_celula

        ws.merge_cells(start_row=2, end_row=2, start_column=1, end_column=3)
        cel2 = ws.cell(row=2, column=1, value=intervalo_codigos)
        cel2.font = font_banner
        cel2.fill = fill_banner
        cel2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 24.0
        for c in range(1, 4):
            ws.cell(row=2, column=c).fill = fill_banner
            ws.cell(row=2, column=c).border = borda_celula

        # 2. Linha 3: Cabeçalho das 3 Colunas (Código | Causa | Localização da Avaria)
        colunas_cab = ["Código", "Causa", "Localização da Avaria"]
        larguras = {
            "Código": 16.57,
            "Causa": 52.43,
            "Localização da Avaria": 52.57
        }
        if isinstance(larguras_personalizadas, dict):
            for k in colunas_cab:
                if k in larguras_personalizadas:
                    larguras[k] = float(larguras_personalizadas[k])

        for c_idx, col_name in enumerate(colunas_cab, 1):
            cel_h = ws.cell(row=3, column=c_idx, value=col_name)
            cel_h.font = font_cabecalho
            cel_h.fill = fill_banner
            cel_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cel_h.border = borda_celula
            ws.column_dimensions[get_column_letter(c_idx)].width = larguras[col_name]
        ws.row_dimensions[3].height = 24.0

        PADRAO_DESCRICAO_SIMPLO = (
            "Possíveis causas incluem falhas elétricas ou desgaste mecânico. "
            "Antes de substituir peças, realize testes detalhados e utilize "
            "diagramas elétricos e manuais técnicos SIMPLO para um diagnóstico preciso e confiável."
        )

        # 3. Blocos de registros a partir da Linha 4
        linha_atual = 4
        largura_mesclada_desc = larguras["Causa"] + larguras["Localização da Avaria"]

        for idx_item, item in enumerate(dados):
            cod_val = str(item.get("Código") or item.get("Códigos") or item.get("DTC") or "-").strip() or "-"
            causa_val = str(item.get("Causa") or "-").strip() or "-"
            loc_val = str(item.get("Localização da Avaria") or item.get("Componente") or "-").strip() or "-"
            desc_val = str(item.get("Descrição") or "-").strip()
            desc_low = desc_val.lower()
            if desc_val in ["-", "", "None", "null", "nan"] or desc_low.startswith("possíveis causas incluem") or "antes de substituir peças" in desc_low:
                desc_val = AgenteOrganizador.sintetizar_descricao_modo_completo(causa_val, loc_val, desc_val)

            # Linha Superior: Código | Causa | Localização da Avaria
            c_cod = ws.cell(row=linha_atual, column=1, value=cod_val)
            c_cod.font = font_rotulo
            c_cod.fill = fill_cod
            c_cod.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_cod.number_format = "@"
            c_cod.border = borda_celula

            c_causa = ws.cell(row=linha_atual, column=2, value=causa_val)
            c_causa.font = font_corpo
            c_causa.fill = fill_txt
            c_causa.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_causa.border = borda_celula

            c_loc = ws.cell(row=linha_atual, column=3, value=loc_val)
            c_loc.font = font_corpo
            c_loc.fill = fill_txt
            c_loc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_loc.border = borda_celula

            linhas_causa = max(1, math.ceil(len(causa_val) / max(15, int(larguras["Causa"] * 0.9))))
            linhas_loc = max(1, math.ceil(len(loc_val) / max(15, int(larguras["Localização da Avaria"] * 0.9))))
            max_linhas_topo = max(linhas_causa, linhas_loc)
            ws.row_dimensions[linha_atual].height = max(26.25, max_linhas_topo * 16.0 + 8.0)

            # Linha Inferior: 'Descrição' | Texto da Descrição mesclado em B:C
            linha_desc = linha_atual + 1
            c_rot_desc = ws.cell(row=linha_desc, column=1, value="Descrição")
            c_rot_desc.font = font_rotulo
            c_rot_desc.fill = fill_cod
            c_rot_desc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c_rot_desc.border = borda_celula

            ws.merge_cells(start_row=linha_desc, end_row=linha_desc, start_column=2, end_column=3)
            c_val_desc = ws.cell(row=linha_desc, column=2, value=desc_val)
            c_val_desc.font = font_corpo
            c_val_desc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for col_m in (2, 3):
                cel_m = ws.cell(row=linha_desc, column=col_m)
                cel_m.fill = fill_txt
                cel_m.border = borda_celula

            linhas_desc = max(1, math.ceil(len(desc_val) / max(30, int(largura_mesclada_desc * 0.92))))
            ws.row_dimensions[linha_desc].height = max(45.0, linhas_desc * 16.0 + 10.0)

            # Linha Separadora Branca (altura 7.5) entre blocos de códigos
            if idx_item < len(dados) - 1:
                linha_sep = linha_atual + 2
                for col_s in (1, 2, 3):
                    cel_s = ws.cell(row=linha_sep, column=col_s, value=None)
                    cel_s.fill = fill_sep
                    cel_s.border = borda_celula
                ws.row_dimensions[linha_sep].height = 7.5
                linha_atual += 3
            else:
                linha_atual += 2

        ws.freeze_panes = "A4"
