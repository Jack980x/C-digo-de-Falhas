import os
import sys
import json
import glob
from dotenv import load_dotenv
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt, Confirm

# Importa módulos internos
from core.ocr_engine import OCREngine
from core.leitor_pdf import LeitorPDF
from core.tradutor_ia import TradutorIA
from core.gerador_excel import GeradorExcel
from core.agente_organizador import AgenteOrganizador

console = Console()
load_dotenv()

PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_ENTRADA = os.path.join(PASTA_BASE, "entrada")
PASTA_SAIDA = os.path.join(PASTA_BASE, "saida")
PASTA_MODELOS = os.path.join(PASTA_BASE, "modelos")
ARQUIVO_CONFIG = os.path.join(PASTA_BASE, "config", "config.json")
ARQUIVO_PADROES = os.path.join(PASTA_BASE, "config", "padroes_planilha.json")

os.makedirs(PASTA_ENTRADA, exist_ok=True)
os.makedirs(PASTA_SAIDA, exist_ok=True)
os.makedirs(PASTA_MODELOS, exist_ok=True)

def carregar_config():
    if os.path.exists(ARQUIVO_CONFIG):
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

EXTENSOES_SUPORTADAS = {
    ".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp",
    ".xlsx", ".csv", ".docx", ".html", ".htm", ".txt"
}


def eh_arquivo_suportado(caminho: str) -> bool:
    ext = os.path.splitext(str(caminho or ""))[1].lower()
    return ext in EXTENSOES_SUPORTADAS


def listar_arquivos_entrada() -> list:
    """
    Lista todos os arquivos suportados na pasta 'entrada/' e em todas as suas subpastas recursivamente.
    Ordena agrupando por subpasta e depois por nome de arquivo.
    """
    arquivos = []
    if os.path.isdir(PASTA_ENTRADA):
        for raiz, dirs, files in os.walk(PASTA_ENTRADA):
            # Ignora pastas temporárias, do sistema ou ocultas
            dirs[:] = [d for d in dirs if not d.startswith(".") and not d.startswith("~$")]
            for nome in files:
                if nome.startswith("~$") or nome.startswith("."):
                    continue
                caminho = os.path.join(raiz, nome)
                if os.path.isfile(caminho) and eh_arquivo_suportado(caminho):
                    arquivos.append(caminho)
    ent_abs = os.path.abspath(PASTA_ENTRADA)
    return sorted(
        arquivos,
        key=lambda x: (
            os.path.dirname(os.path.relpath(os.path.abspath(x), ent_abs)).lower(),
            os.path.basename(x).lower()
        )
    )


def obter_pasta_saida_para_arquivo(caminho_pdf: str) -> str:
    """
    Retorna o diretório de saída espelhando a subpasta de entrada se o arquivo estiver
    dentro de uma subpasta de 'entrada/'. Garante que a pasta exista.
    Exemplo: entrada/Scania/manual.pdf -> saida/Scania/
    """
    try:
        if caminho_pdf:
            c_abs = os.path.abspath(caminho_pdf)
            ent_abs = os.path.abspath(PASTA_ENTRADA)
            if c_abs.startswith(ent_abs):
                rel = os.path.relpath(c_abs, ent_abs)
                sub = os.path.dirname(rel)
                if sub and sub != ".":
                    destino = os.path.join(PASTA_SAIDA, sub)
                    os.makedirs(destino, exist_ok=True)
                    return destino
    except Exception:
        pass
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    return PASTA_SAIDA


def obter_pasta_saida_para_lote(caminhos_pdf: list) -> str:
    """
    Se todos os PDFs do lote estiverem na mesma subpasta de 'entrada/',
    retorna a subpasta espelhada em 'saida/'. Se forem de pastas variadas ou soltos,
    retorna 'saida/'.
    """
    pastas_detectadas = set()
    for c in caminhos_pdf:
        p_out = obter_pasta_saida_para_arquivo(c)
        pastas_detectadas.add(p_out)
    if len(pastas_detectadas) == 1:
        return list(pastas_detectadas)[0]
    return PASTA_SAIDA


def capturar_imagem_clipboard() -> str:
    """
    Captura uma imagem ou arquivo copiado na Área de Transferência do Windows (ex: Win + Shift + S / Print Screen),
    salva automaticamente na pasta 'entrada/' e retorna o caminho do arquivo salvo.
    """
    from datetime import datetime
    from PIL import ImageGrab, Image
    try:
        clip = ImageGrab.grabclipboard()
    except Exception as e:
        console.print(f"[red]Não foi possível acessar a Área de Transferência: {e}[/red]")
        return ""

    if isinstance(clip, Image.Image):
        nome_arq = f"print_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        caminho_salvo = os.path.join(PASTA_ENTRADA, nome_arq)
        clip.convert("RGB").save(caminho_salvo, "PNG")
        console.print(f"[bold green]📸 Print Screen capturado da Área de Transferência e salvo em:[/bold green] [yellow]{caminho_salvo}[/yellow]")
        return caminho_salvo

    if isinstance(clip, list):
        for item in clip:
            if isinstance(item, str) and os.path.isfile(item) and eh_arquivo_suportado(item):
                return os.path.abspath(item)

    console.print("[yellow]⚠️ Nenhuma imagem ou arquivo compatível encontrado na Área de Transferência (use Win + Shift + S ou Print Screen primeiro).[/yellow]")
    return ""


def parsear_selecao_arquivos_pdf(entrada: str, arquivos_disponiveis: list) -> list:
    """
    Interpreta a seleção de um ou múltiplos arquivos (PDF, Imagem ou Documento):
    - 'p', 'print', 'clipboard', 'colar' -> captura imagem direto da Área de Transferência
    - '0', 'todos', 'todas', 'all', '*' -> seleciona todos os arquivos da lista
    - Números e intervalos -> ex: '1', '1,2', '1-3', '1, 3-5'
    - Nome de subpasta -> ex: 'Scania', 'Mercedes' seleciona todos os arquivos daquela pasta
    - Caminhos de diretórios ou arquivos arrastados do Windows
    """
    import re
    if not entrada or not entrada.strip():
        return list(arquivos_disponiveis) if len(arquivos_disponiveis) == 1 else []

    texto = entrada.strip()
    if texto.lower() in ["p", "print", "clipboard", "colar"]:
        arq_clip = capturar_imagem_clipboard()
        return [arq_clip] if arq_clip else []

    if texto.lower() in ["0", "todos", "todas", "all", "*"]:
        return list(arquivos_disponiveis)

    total = len(arquivos_disponiveis)
    selecionados = []

    # Suporta caminhos entre aspas ou separados por vírgula/ponto-e-vírgula
    partes = [p.strip().strip('"\'') for p in re.split(r'[,;]+', texto) if p.strip()]

    for p in partes:
        if p.lower() in ["p", "print", "clipboard", "colar"]:
            arq_clip = capturar_imagem_clipboard()
            if arq_clip and arq_clip not in selecionados:
                selecionados.append(arq_clip)
            continue

        if p.lower() in ["0", "todos", "todas", "all", "*"]:
            for arq in arquivos_disponiveis:
                if arq not in selecionados:
                    selecionados.append(arq)
            continue

        # 1. Se for intervalo numérico: ex: 1-10
        m_range = re.match(r'^(\d+)\s*-\s*(\d+)$', p)
        if m_range and total > 0:
            ini, fim = int(m_range.group(1)), int(m_range.group(2))
            if ini > fim:
                ini, fim = fim, ini
            for idx in range(ini, fim + 1):
                if 1 <= idx <= total:
                    arq = arquivos_disponiveis[idx - 1]
                    if arq not in selecionados:
                        selecionados.append(arq)
            continue

        # 2. Se for número único da tabela: ex: 5
        if p.isdigit() and total > 0:
            idx = int(p)
            if 1 <= idx <= total:
                arq = arquivos_disponiveis[idx - 1]
                if arq not in selecionados:
                    selecionados.append(arq)
                continue

        # 3. Se for um diretório arrastado do Windows ou subpasta direta em entrada/
        cand_dir = p if os.path.isdir(p) else os.path.join(PASTA_ENTRADA, p)
        if os.path.isdir(cand_dir):
            for r_sub, _, f_sub in os.walk(cand_dir):
                for f_nome in sorted(f_sub):
                    if not f_nome.startswith("~$") and not f_nome.startswith("."):
                        c_f = os.path.join(r_sub, f_nome)
                        if os.path.isfile(c_f) and eh_arquivo_suportado(c_f):
                            if c_f not in selecionados:
                                selecionados.append(c_f)
            continue

        # 4. Se coincidir com o nome de alguma subpasta presente na lista disponível (ex: digitou "Scania")
        arqs_da_pasta = [
            arq for arq in arquivos_disponiveis
            if p.lower() in os.path.dirname(os.path.relpath(arq, PASTA_ENTRADA)).lower().split(os.sep)
            or p.lower() == os.path.dirname(os.path.relpath(arq, PASTA_ENTRADA)).lower()
        ]
        if arqs_da_pasta:
            for arq in arqs_da_pasta:
                if arq not in selecionados:
                    selecionados.append(arq)
            continue

        # 5. Caminho completo informado pelo usuário para arquivo individual
        if os.path.isfile(p) and eh_arquivo_suportado(p):
            caminho_abs = os.path.abspath(p)
            if caminho_abs not in selecionados:
                selecionados.append(caminho_abs)
            continue

        # 6. Nome de arquivo direto dentro da pasta entrada/
        cand_direto = os.path.join(PASTA_ENTRADA, p)
        if os.path.isfile(cand_direto) and eh_arquivo_suportado(cand_direto):
            if cand_direto not in selecionados:
                selecionados.append(cand_direto)
            continue

        for ext_cand in EXTENSOES_SUPORTADAS:
            cand_ext = os.path.join(PASTA_ENTRADA, f"{p}{ext_cand}")
            if os.path.isfile(cand_ext):
                if cand_ext not in selecionados:
                    selecionados.append(cand_ext)
                break

        # 7. Busca por correspondência exata ou parcial de nome de arquivo
        for arq in arquivos_disponiveis:
            base_arq = os.path.basename(arq).lower()
            sem_ext = os.path.splitext(base_arq)[0]
            if p.lower() in [base_arq, sem_ext]:
                if arq not in selecionados:
                    selecionados.append(arq)

    return selecionados


def agrupar_arquivos_por_pasta(arquivos: list) -> dict:
    """
    Agrupa uma lista de arquivos pela subpasta relativa a 'PASTA_ENTRADA'.
    Retorna um dicionário ordenado {subpasta_relativa: [lista_de_arquivos]}.
    Arquivos diretamente na raiz de 'entrada/' terão a chave "".
    """
    from collections import OrderedDict
    grupos = OrderedDict()
    ent_abs = os.path.abspath(PASTA_ENTRADA)

    for arq in arquivos:
        c_abs = os.path.abspath(arq)
        if c_abs.startswith(ent_abs):
            rel = os.path.relpath(c_abs, ent_abs)
            sub = os.path.dirname(rel).replace("\\", "/").strip()
            chave = sub if sub and sub != "." else ""
        else:
            chave = os.path.dirname(arq).replace("\\", "/")
        grupos.setdefault(chave, []).append(arq)

    return grupos


def _exibir_tabela_e_selecionar_arquivos(
    arquivos: list,
    titulo_pasta: str = None,
    permitir_voltar: bool = True
) -> list:
    """
    Exibe a tabela com os arquivos de uma pasta específica (ou de entrada geral)
    e aguarda a seleção do usuário. Se o usuário digitar 'V' (voltar), retorna None.
    """
    if not arquivos:
        console.print("[yellow]Nenhum arquivo nesta pasta.[/yellow]")
        return None if permitir_voltar else []

    titulo = (
        f"Conteúdo da Pasta: '{titulo_pasta}' ({len(arquivos)} arquivo(s))"
        if titulo_pasta else
        f"Arquivos Disponíveis na Pasta 'entrada/' ({len(arquivos)} arquivo(s))"
    )
    tabela = Table(title=titulo, show_header=True, header_style="bold magenta")
    tabela.add_column("Opção", style="cyan", width=8, justify="center")
    tabela.add_column("Tipo", style="yellow", width=10, justify="center")
    tabela.add_column("Nome do Arquivo", style="green")
    tabela.add_column("Tamanho", justify="right")

    for i, arq in enumerate(arquivos, 1):
        tam_mb = os.path.getsize(arq) / (1024 * 1024)
        ext_up = os.path.splitext(arq)[1].upper().lstrip(".")
        if ext_up == "PDF":
            tipo_badge = "PDF"
        elif f".{ext_up.lower()}" in LeitorPDF.EXTENSOES_IMAGEM:
            tipo_badge = f"IMG ({ext_up})"
        else:
            tipo_badge = f"DOC ({ext_up})"

        nome_arq = os.path.basename(arq)
        tabela.add_row(str(i), tipo_badge, nome_arq, f"{tam_mb:.2f} MB")

    if len(arquivos) > 1:
        tam_total = sum(os.path.getsize(a) for a in arquivos) / (1024 * 1024)
        rotulo_todos = (
            f"[bold yellow]TODOS OS {len(arquivos)} ARQUIVOS DESTA PASTA[/bold yellow]"
            if titulo_pasta else
            f"[bold yellow]TODOS OS {len(arquivos)} ARQUIVOS[/bold yellow]"
        )
        tabela.add_row("0 / todos", "TODOS", rotulo_todos, f"{tam_total:.2f} MB")

    if permitir_voltar:
        tabela.add_row("V / voltar", "VOLTAR", "[bold magenta]⬅️ Voltar para a seleção de pastas[/bold magenta]", "-")

    tabela.add_row("P / print", "CLIPBOARD", "[bold cyan]📸 Colar Print Screen da Área de Transferência (Win+Shift+S)[/bold cyan]", "-")

    console.print(tabela)

    console.print("[dim]Como selecionar:[/dim]")
    console.print(f"  • Digite o(s) [bold cyan]número(s) ou intervalo[/bold cyan] (ex: [green]1[/green], [green]1,2[/green], [green]1-3[/green] ou [green]0[/green] para todos)")
    if permitir_voltar:
        console.print("  • Digite [bold magenta]V[/bold magenta] para [bold magenta]voltar[/bold magenta] à lista de pastas")
    console.print("  • Digite [bold cyan]P[/bold cyan] para colar uma imagem copiada na Área de Transferência (Print Screen)")
    console.print("  • Ou cole o(s) [bold cyan]caminho(s) completo(s)[/bold cyan] ou arraste arquivos do Windows")

    padrao_arq = "1" if len(arquivos) == 1 else "0"
    while True:
        prompt_txt = f"[bold green]Escolha o(s) arquivo(s) (1 a {len(arquivos)}, '0' para todos"
        if permitir_voltar:
            prompt_txt += ", ou 'V' para voltar"
        prompt_txt += ")[/bold green]"

        escolha = Prompt.ask(prompt_txt, default=padrao_arq).strip()
        if permitir_voltar and escolha.lower() in ["v", "voltar", "back", "b"]:
            return None

        selecionados = parsear_selecao_arquivos_pdf(escolha, arquivos)
        if selecionados:
            nomes_sel = ", ".join(os.path.basename(a) for a in selecionados[:5])
            if len(selecionados) > 5:
                nomes_sel += f" ... (+{len(selecionados) - 5} arquivos)"
            console.print(f"[green]✓ {len(selecionados)} arquivo(s) selecionado(s):[/green] [bold]{nomes_sel}[/bold]")
            return selecionados
        console.print("[red]Opção inválida! Informe números da tabela, '0' para todos, 'P' ou 'V' para voltar.[/red]")


def selecionar_arquivos_pdf() -> list:
    """
    Permite selecionar 1 ou múltiplos arquivos (PDF, Imagem ou Documento) na pasta 'entrada/'.
    Se existirem subpastas em 'entrada/', exibe uma navegação limpa pelas pastas:
      - Mostra apenas as pastas encontradas
      - Permite selecionar 1 pasta para navegar e ver seus arquivos
      - Permite selecionar 2 ou mais pastas (ou '0' para todas) para processá-las em lote
      - Permite voltar da navegação de uma pasta digitando 'V'
    Se não houver subpastas (apenas arquivos na raiz), mostra diretamente a lista de arquivos.
    """
    console.print("\n[bold cyan]📁 Seleção de Arquivo(s) (PDF, Imagem ou Documento)[/bold cyan]")
    console.print("[dim]Formatos aceitos: .pdf, .png, .jpg, .jpeg, .tiff, .bmp, .webp, .xlsx, .csv, .docx, .html[/dim]")
    arquivos_todos = listar_arquivos_entrada()

    if not arquivos_todos:
        console.print(f"[yellow]Nenhum arquivo encontrado na pasta '[bold]{PASTA_ENTRADA}[/bold]'.[/yellow]")
        caminho_manual = Prompt.ask("[green]Digite o caminho de arquivo(s)/pasta, arraste aqui ou digite 'P' para colar Print Screen[/green]")
        selecionados = parsear_selecao_arquivos_pdf(caminho_manual, [])
        if selecionados:
            return selecionados
        console.print("[red]Nenhum arquivo válido encontrado![/red]")
        return []

    grupos_por_pasta = agrupar_arquivos_por_pasta(arquivos_todos)
    subpastas_reais = [k for k in sorted(grupos_por_pasta.keys()) if k != ""]
    tem_raiz = "" in grupos_por_pasta and len(grupos_por_pasta[""]) > 0

    # Se NÃO houver nenhuma subpasta, lista os arquivos da raiz diretamente
    if not subpastas_reais:
        return _exibir_tabela_e_selecionar_arquivos(arquivos_todos, titulo_pasta=None, permitir_voltar=False)

    # Quando HÁ subpastas: navegação por pastas
    while True:
        tabela_pastas = Table(
            title=f"Pastas Encontradas na Pasta 'entrada/' ({len(subpastas_reais)} pasta(s))",
            show_header=True,
            header_style="bold magenta"
        )
        tabela_pastas.add_column("Opção", style="cyan", width=8, justify="center")
        tabela_pastas.add_column("Pasta / Origem", style="green", width=30)
        tabela_pastas.add_column("Qtd de Arquivos", justify="center", width=16)
        tabela_pastas.add_column("Tamanho", justify="right", width=12)
        tabela_pastas.add_column("Exemplos de Documentos", style="dim")

        mapa_opcoes = {}  # "1" -> chave da subpasta
        idx = 1
        for sub in subpastas_reais:
            arqs_sub = grupos_por_pasta[sub]
            tam_mb = sum(os.path.getsize(a) for a in arqs_sub) / (1024 * 1024)
            exs = ", ".join(os.path.basename(a) for a in arqs_sub[:3])
            if len(arqs_sub) > 3:
                exs += f" (+{len(arqs_sub) - 3})"
            op_str = str(idx)
            mapa_opcoes[op_str] = sub
            tabela_pastas.add_row(op_str, f"📁 {sub}", f"{len(arqs_sub)} arquivo(s)", f"{tam_mb:.2f} MB", exs)
            idx += 1

        if tem_raiz:
            arqs_raiz = grupos_por_pasta[""]
            tam_mb = sum(os.path.getsize(a) for a in arqs_raiz) / (1024 * 1024)
            exs = ", ".join(os.path.basename(a) for a in arqs_raiz[:3])
            if len(arqs_raiz) > 3:
                exs += f" (+{len(arqs_raiz) - 3})"
            op_str = str(idx)
            mapa_opcoes[op_str] = ""
            mapa_opcoes["r"] = ""
            tabela_pastas.add_row(f"{op_str} / R", "📄 Arquivos na raiz de entrada/", f"{len(arqs_raiz)} arquivo(s)", f"{tam_mb:.2f} MB", exs)
            idx += 1

        if len(subpastas_reais) > 1 or tem_raiz:
            tam_total = sum(os.path.getsize(a) for a in arquivos_todos) / (1024 * 1024)
            tabela_pastas.add_row(
                "0 / todas",
                "[bold yellow]📁 TODAS AS PASTAS[/bold yellow]",
                f"[bold yellow]{len(arquivos_todos)} arquivos no total[/bold yellow]",
                f"{tam_total:.2f} MB",
                "Selecionar todos os arquivos de todas as pastas de uma só vez"
            )

        tabela_pastas.add_row("P / print", "[bold cyan]📸 Colar Print Screen da Área de Transferência[/bold cyan]", "-", "-", "Captura direta (Win+Shift+S)")

        console.print(tabela_pastas)

        console.print("[dim]Como escolher:[/dim]")
        console.print("  • Digite [bold cyan]1 número[/bold cyan] (ex: [green]1[/green]) para [bold green]entrar na pasta[/bold green] e ver os arquivos dela")
        console.print("  • Digite [bold cyan]múltiplos números[/bold cyan] (ex: [green]1,2[/green] ou [green]1-2[/green]) para [bold green]selecionar as pastas juntas[/bold green]")
        console.print("  • Digite [bold cyan]0[/bold cyan] para selecionar [bold green]todas as pastas[/bold green]")
        console.print(f"  • Ou digite o [bold cyan]nome da pasta[/bold cyan] diretamente (ex: [green]{subpastas_reais[0]}[/green])")

        prompt_txt = f"[bold green]Escolha a pasta (1 a {idx - 1}, '1,2' para ambas, ou '0' para todas)[/bold green]"
        escolha = Prompt.ask(prompt_txt, default="1").strip()

        # Verifica clipboard
        if escolha.lower() in ["p", "print", "clipboard", "colar"]:
            arq_clip = capturar_imagem_clipboard()
            if arq_clip:
                return [arq_clip]
            continue

        # Verifica se arrastou arquivo ou pasta do Windows
        caminho_arrastado = escolha.strip('"\'')
        if os.path.exists(caminho_arrastado):
            sel_arrastado = parsear_selecao_arquivos_pdf(escolha, arquivos_todos)
            if sel_arrastado:
                return sel_arrastado

        # Verifica se escolheu "todas as pastas"
        if escolha.lower() in ["0", "todas", "todos", "all", "*"]:
            console.print(f"\n[bold green]✓ Todas as pastas selecionadas ({len(arquivos_todos)} arquivos no total)![/bold green]")
            return arquivos_todos

        # Interpreta as pastas escolhidas (números, intervalos, 'R' ou nomes)
        import re
        partes = [p.strip().strip('"\'') for p in re.split(r'[,;]+', escolha) if p.strip()]
        chaves_escolhidas = []

        for p in partes:
            # Intervalo numérico (ex: 1-2)
            m_range = re.match(r'^(\d+)\s*-\s*(\d+)$', p)
            if m_range:
                ini, fim = int(m_range.group(1)), int(m_range.group(2))
                if ini > fim:
                    ini, fim = fim, ini
                for i_num in range(ini, fim + 1):
                    k_num = str(i_num)
                    if k_num in mapa_opcoes and mapa_opcoes[k_num] not in chaves_escolhidas:
                        chaves_escolhidas.append(mapa_opcoes[k_num])
                continue

            # Opção numérica ou 'R'
            p_low = p.lower()
            if p in mapa_opcoes:
                if mapa_opcoes[p] not in chaves_escolhidas:
                    chaves_escolhidas.append(mapa_opcoes[p])
                continue
            if p_low in mapa_opcoes:
                if mapa_opcoes[p_low] not in chaves_escolhidas:
                    chaves_escolhidas.append(mapa_opcoes[p_low])
                continue

            # Nome da subpasta por texto (ex: "Scania", "Volvo")
            match_pasta = next((sub for sub in subpastas_reais if p_low == sub.lower() or p_low in sub.lower().split("/")), None)
            if match_pasta:
                if match_pasta not in chaves_escolhidas:
                    chaves_escolhidas.append(match_pasta)
                continue

        if not chaves_escolhidas:
            console.print("[red]Opção de pasta inválida! Digite o número da pasta, '1,2', '0' ou o nome da pasta.[/red]")
            continue

        # CASO 1: O usuário selecionou APENAS UMA pasta -> Navega dentro dela!
        if len(chaves_escolhidas) == 1:
            chave_unica = chaves_escolhidas[0]
            nome_exib = chave_unica if chave_unica else "Raiz de entrada/"
            arqs_pasta = grupos_por_pasta.get(chave_unica, [])
            console.print(f"\n[bold cyan]📂 Abrindo pasta: '{nome_exib}'...[/bold cyan]")
            sel_arquivos = _exibir_tabela_e_selecionar_arquivos(
                arqs_pasta,
                titulo_pasta=nome_exib,
                permitir_voltar=True
            )
            # Se o usuário digitou 'V' (Voltar), sel_arquivos é None -> volta para a lista de pastas!
            if sel_arquivos is None:
                continue
            return sel_arquivos

        # CASO 2: O usuário selecionou 2 OU MAIS pastas -> Seleciona todos os arquivos dessas pastas
        else:
            todos_das_pastas = []
            nomes_pastas_exib = []
            for ch in chaves_escolhidas:
                arqs_ch = grupos_por_pasta.get(ch, [])
                todos_das_pastas.extend(arqs_ch)
                nom = ch if ch else "Raiz de entrada/"
                nomes_pastas_exib.append(f"{nom} ({len(arqs_ch)} arqs)")

            console.print(
                f"\n[bold green]✓ {len(todos_das_pastas)} arquivo(s) selecionado(s) de {len(chaves_escolhidas)} pasta(s):[/bold green] "
                f"[bold cyan]{' | '.join(nomes_pastas_exib)}[/bold cyan]"
            )
            return todos_das_pastas


def selecionar_arquivo_pdf() -> str:
    selecionados = selecionar_arquivos_pdf()
    return selecionados[0] if selecionados else None


def carregar_padroes_customizados() -> list:
    if os.path.exists(ARQUIVO_PADROES):
        try:
            with open(ARQUIVO_PADROES, "r", encoding="utf-8") as f:
                dados = json.load(f)
            if isinstance(dados, list):
                return dados
        except Exception:
            pass
    return []


def salvar_padroes_customizados(lista_padroes: list):
    try:
        os.makedirs(os.path.dirname(ARQUIVO_PADROES), exist_ok=True)
        with open(ARQUIVO_PADROES, "w", encoding="utf-8") as f:
            json.dump(lista_padroes, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def importar_novo_padrao_excel() -> dict:
    """
    Permite ao usuário escolher um arquivo .xlsx na pasta 'modelos/' (ou 'saida/' / 'testes_e_exemplos/' ou digitar o caminho)
    para clonar automaticamente suas colunas, larguras, cores, fontes e estilo de banner.
    """
    console.print("\n[bold cyan]📥 Importador de Padrão de Planilha (.xlsx)[/bold cyan]")
    os.makedirs(PASTA_MODELOS, exist_ok=True)
    arquivos_mod = []
    pastas_busca = [
        (PASTA_MODELOS, "modelos/"),
        (PASTA_SAIDA, "saida/"),
        (os.path.join(PASTA_BASE, "testes_e_exemplos"), "testes_e_exemplos/")
    ]
    for pasta_dir, rotulo_p in pastas_busca:
        if os.path.isdir(pasta_dir):
            for nome in sorted(os.listdir(pasta_dir)):
                if nome.lower().endswith(".xlsx") and not nome.startswith("~$"):
                    arquivos_mod.append((os.path.join(pasta_dir, nome), f"{rotulo_p}{nome}"))

    if arquivos_mod:
        tab_m = Table(title="Planilhas (.xlsx) Disponíveis para Clonar o Padrão Visual", show_header=True, header_style="bold magenta")
        tab_m.add_column("Opção", style="cyan", width=8, justify="center")
        tab_m.add_column("Arquivo de Referência", style="green")
        for idx, (_, rotulo_exib) in enumerate(arquivos_mod, 1):
            tab_m.add_row(str(idx), rotulo_exib)
        console.print(tab_m)
        entrada_m = Prompt.ask("[bold green]Digite o número do arquivo acima ou cole o caminho de qualquer arquivo .xlsx[/bold green]", default="1").strip().strip('"\'')
        if entrada_m.isdigit() and 1 <= int(entrada_m) <= len(arquivos_mod):
            caminho_xlsx = arquivos_mod[int(entrada_m) - 1][0]
        else:
            caminho_xlsx = entrada_m
    else:
        console.print(f"[dim]Dica: Você pode colocar planilhas de referência na pasta '{PASTA_MODELOS}'.[/dim]")
        caminho_xlsx = Prompt.ask("[bold green]Cole o caminho de um arquivo .xlsx de referência para copiar o padrão[/bold green]").strip().strip('"\'')

    if not caminho_xlsx or not os.path.isfile(caminho_xlsx):
        console.print("[yellow]Arquivo .xlsx não encontrado. Usando padrão oficial SIMPLO.[/yellow]")
        return {}

    try:
        info_mod = GeradorExcel.inspecionar_modelo_excel(caminho_xlsx)
        nome_sug = os.path.splitext(os.path.basename(caminho_xlsx))[0]
        nome_padrao = Prompt.ask("[bold green]Nome para salvar este padrão no catálogo[/bold green]", default=nome_sug).strip() or nome_sug
        novo_item = {
            "nome": nome_padrao,
            "colunas": info_mod["colunas"],
            "larguras": info_mod["larguras"],
            "tema_visual": info_mod["tema_visual"],
            "gerar_buscar_codigo_extra": False
        }
        atuais = carregar_padroes_customizados()
        atuais = [p for p in atuais if p.get("nome", "").lower() != nome_padrao.lower()]
        atuais.append(novo_item)
        salvar_padroes_customizados(atuais)
        console.print(
            f"[bold green]✓ Padrão '{nome_padrao}' importado e salvo![/bold green] "
            f"Colunas: [cyan]{' | '.join(novo_item['colunas'])}[/cyan]"
        )
        return novo_item
    except Exception as e:
        console.print(f"[red]Erro ao inspecionar planilha modelo: {e}[/red]")
        return {}


def selecionar_padrao_planilha(config: dict, resumo_info: dict = None, titulo: str = None) -> dict:
    """
    Exibe o Catálogo de Padrões de Planilha Excel com seleção de 1 tecla:
    - [1] SIMPLO 2 Colunas (Compacto): Código | Descrição (23.71 | 96.57)
    - [2] SIMPLO 3 Colunas + Descrição: Código | Causa | Localização da Avaria + Linha Descrição mesclada (16.57 | 52.43 | 52.57)
    - [3] SIMPLO 4 Colunas (Oficial BC_IBC2): Códigos | Componente | FMI | Descrição (15.86 | 40.29 | 14.00 | 48.57)
    - [4] Importar Modelo de um .xlsx
    - [5] Personalizar Manualmente
    Retorna: {"colunas": list, "larguras": dict, "tema_visual": Any}
    """
    est_doc = (resumo_info or {}).get("estrutura_colunas") or {}
    tem_2_cods = bool(est_doc.get("tem_duas_colunas_codigo"))
    col_c1 = est_doc.get("col_codigo_1", "DTC Designação do Veículo")
    col_c2 = est_doc.get("col_codigo_2", "DTCs")
    padrao_default = "3" if est_doc.get("tem_fmi") else "1"

    cabecalho = titulo if titulo else "Catálogo de Padrões de Planilha Excel"
    console.print(f"\n[bold cyan]🎨 {cabecalho}[/bold cyan]")
    tabela_p = Table(title="Escolha o Padrão de Planilha (Colunas + Larguras + Layout)", show_header=True, header_style="bold magenta")
    tabela_p.add_column("Opção", style="cyan", width=8, justify="center")
    tabela_p.add_column("Padrão de Planilha", style="green", width=36)
    tabela_p.add_column("Colunas e Estilo Visual", style="dim")

    tabela_p.add_row("1", "SIMPLO 2 Colunas (Compacto)", "Código | Descrição (Larguras 23.71 | 96.57 — Azul Royal #0078D7)")
    tabela_p.add_row("2", "SIMPLO 3 Colunas + Descrição", "Código | Causa | Localização da Avaria + Linha 'Descrição' mesclada (16.57 | 52.43 | 52.57)")
    tabela_p.add_row("3", "SIMPLO 4 Colunas (Oficial BC_IBC2)", "Códigos | Componente | FMI | Descrição (15.86 | 40.29 | 14.00 | 48.57)")
    tabela_p.add_row("4", "📥 Importar Modelo de um .xlsx", "Copia colunas, larguras, cores e fontes de qualquer planilha .xlsx")
    tabela_p.add_row("5", "⚙️ Personalizar Manualmente", "Digitar nomes das colunas e larguras manualmente")

    padroes_custom = carregar_padroes_customizados()
    for idx_c, p_cust in enumerate(padroes_custom, start=6):
        cols_c_str = " | ".join(p_cust.get("colunas", []))
        tabela_p.add_row(str(idx_c), f"⭐ {p_cust.get('nome', 'Custom')}", f"{cols_c_str} (Padrão Salvo)")

    console.print(tabela_p)
    opcoes_validas = [str(i) for i in range(1, 6 + len(padroes_custom))]
    prompt_txt = "[bold green]Qual o padrão do Excel a ser gerado?[/bold green]"
    escolha_p = Prompt.ask(prompt_txt, choices=opcoes_validas, default=padrao_default)

    if escolha_p == "1":
        cols = ["Código", "Descrição"]
        largs = {"Código": 23.71, "Descrição": 96.57}
        if tem_2_cods:
            console.print(f"\n[bold yellow]💡 Este documento possui 2 colunas de códigos lado a lado na tabela ('{col_c1}' e '{col_c2}'):[/bold yellow]")
            console.print(f"  [1] Usar a 1ª coluna de código ([bold green]{col_c1}[/bold green], ex: U240088, C110116) [dim][Padrão][/dim]")
            console.print(f"  [2] Usar a 2ª coluna hexadecimal ([bold cyan]{col_c2}[/bold cyan], ex: E40088, 510116)")
            console.print(f"  [3] Incluir [bold magenta]AMBAS as colunas de código[/bold magenta] ({col_c1} | {col_c2} | Descrição)")
            op_2c = Prompt.ask("Qual coluna de código deseja incluir na planilha?", choices=["1", "2", "3"], default="1")
            if op_2c == "2":
                cols = ["DTCs", "Descrição"]
                largs = {"DTCs": 23.71, "Descrição": 96.57}
            elif op_2c == "3":
                cols = [col_c1, col_c2, "Descrição"]
                largs = {col_c1: 18.0, col_c2: 14.0, "Descrição": 88.28}
        salvar_colunas_recentes(cols)
        return {"colunas": cols, "larguras": largs, "tema_visual": "simplo_azul"}

    if escolha_p == "2":
        cols = ["Código", "Causa", "Localização da Avaria", "Descrição"]
        largs = {"Código": 16.57, "Causa": 52.43, "Localização da Avaria": 52.57}
        salvar_colunas_recentes(cols)
        return {"colunas": cols, "larguras": largs, "tema_visual": "simplo_3_colunas_desc"}

    if escolha_p == "3":
        cols = ["Códigos", "Componente", "FMI", "Descrição"]
        largs = {"Códigos": 15.86, "Componente": 40.29, "FMI": 14.00, "Descrição": 48.57}
        salvar_colunas_recentes(cols)
        return {"colunas": cols, "larguras": largs, "tema_visual": "simplo_azul"}

    if escolha_p == "4":
        importado = importar_novo_padrao_excel()
        if importado:
            return importado
        cols = ["Código", "Descrição"]
        largs = {"Código": 23.71, "Descrição": 96.57}
        return {"colunas": cols, "larguras": largs, "tema_visual": "simplo_azul"}

    if escolha_p == "5":
        cols = selecionar_ou_configurar_colunas(config)
        largs = selecionar_ou_configurar_larguras(cols, config)
        return {"colunas": cols, "larguras": largs, "tema_visual": "simplo_azul"}

    idx_cust = int(escolha_p) - 6
    p_sel = padroes_custom[idx_cust]
    return {
        "colunas": p_sel.get("colunas", ["Código", "Descrição"]),
        "larguras": p_sel.get("larguras"),
        "tema_visual": p_sel.get("tema_visual", "simplo_azul")
    }

def salvar_colunas_recentes(colunas: list):
    try:
        config = carregar_config()
        config["ultimas_colunas"] = colunas
        with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

def selecionar_ou_configurar_colunas(config: dict) -> list:
    console.print("\n[bold cyan]📋 Colunas da Planilha Excel[/bold cyan]")
    
    # Colunas padrão ou últimas usadas
    colunas_salvas = config.get("ultimas_colunas")
    if not colunas_salvas:
        colunas_salvas = [
            "Módulo/Sistema",
            "Código/SPN",
            "FMI",
            "Descrição Original",
            "Descrição Técnica (PT-BR)",
            "Causas Prováveis",
            "Procedimento de Reparo"
        ]
        
    padrao_str = ", ".join(colunas_salvas)
    console.print(f"[dim]Padrão atual: {padrao_str}[/dim]")
    console.print("[yellow]Digite as colunas que você quer (separadas por vírgula) ou aperte [bold green]Enter[/bold green] para usar o padrão:[/yellow]")
    
    entrada = Prompt.ask("[bold green]Colunas[/bold green]", default=padrao_str)
    
    # Quebra por vírgula e limpa espaços
    colunas = [c.strip() for c in entrada.split(",") if c.strip()]
    if not colunas:
        colunas = colunas_salvas
        
    salvar_colunas_recentes(colunas)
    console.print(f"[green]✓ Colunas definidas:[/green] [bold]{' | '.join(colunas)}[/bold]\n")
    return colunas


def salvar_larguras_recentes(larguras: dict):
    try:
        config = carregar_config()
        config["ultimas_larguras"] = larguras
        with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def selecionar_ou_configurar_larguras(colunas: list, config: dict, perfil: str = "padrao") -> dict:
    """
    Permite ao usuário visualizar e definir a largura exata de cada coluna antes de iniciar a extração,
    seguindo por padrão as medidas do material de referência (BC_IBC2: 15.86, 40.29, 14.00, 48.57).
    """
    console.print("\n[bold cyan]📏 Configuração das Larguras das Colunas (Excel)[/bold cyan]")

    larguras_salvas = config.get("ultimas_larguras", {}) if config else {}
    from core.gerador_excel import GeradorExcel
    gerador_temp = GeradorExcel()

    if perfil == "buscar_codigo":
        padroes_referencia = {
            "Código": 16.0,
            "Localização da Avaria": 48.0,
            "Causa": 35.0,
            "Descrição": 55.0
        }
    else:
        # Padrão oficial exato solicitado pelo usuário (material BC_IBC2)
        padroes_referencia = {
            "codigo": 15.86,
            "componente": 40.29,
            "fmi": 14.00,
            "desc": 48.57,
            "fmi_desc": 48.57,
            "modulo": 18.00,
            "causa": 35.00,
            "status": 12.00
        }

    larguras_sugeridas = []
    for col in colunas:
        if col in larguras_salvas:
            larguras_sugeridas.append(float(larguras_salvas[col]))
        elif perfil == "buscar_codigo" and col in padroes_referencia:
            larguras_sugeridas.append(float(padroes_referencia[col]))
        else:
            tipo = gerador_temp.classificar_tipo_coluna(col)
            larguras_sugeridas.append(float(padroes_referencia.get(tipo, 25.00)))

    sugestao_str = ", ".join([f"{w:.2f}".rstrip('0').rstrip('.') for w in larguras_sugeridas])

    tabela = Table(title="Larguras Propostas (Padrão Oficial SIMPLO / BC_IBC2)", show_header=True, header_style="bold magenta")
    tabela.add_column("Ordem", justify="center", style="dim", width=6)
    tabela.add_column("Coluna", style="cyan")
    tabela.add_column("Largura Proposta", justify="right", style="green")

    for idx, (col, w) in enumerate(zip(colunas, larguras_sugeridas), 1):
        tabela.add_row(str(idx), col, f"{w:.2f}")
    console.print(tabela)

    console.print("[dim]Dica: As larguras seguem o modelo de referência (BC_IBC2: 15.86 | 40.29 | 14.00 | 48.57).[/dim]")
    console.print("[yellow]Pressione [bold green]Enter[/bold green] para confirmar ou digite os novos valores separados por vírgula:[/yellow]")

    entrada = Prompt.ask("[bold green]Larguras[/bold green]", default=sugestao_str)
    entrada = entrada.strip()

    larguras_finais = {}
    try:
        import re
        valores = [float(m.replace(',', '.')) for m in re.findall(r'\d+(?:[.,]\d+)?', entrada)]
        if len(valores) == len(colunas):
            for col, v in zip(colunas, valores):
                larguras_finais[col] = v
        else:
            console.print(f"[yellow]Quantidade de valores informados ({len(valores)}) diferente do número de colunas ({len(colunas)}). Usando padrão recomendado.[/yellow]")
            for col, v in zip(colunas, larguras_sugeridas):
                larguras_finais[col] = v
    except Exception:
        console.print("[yellow]Entrada inválida. Usando padrão recomendado.[/yellow]")
        for col, v in zip(colunas, larguras_sugeridas):
            larguras_finais[col] = v

    salvar_larguras_recentes(larguras_finais)
    resumo = " | ".join([f"{col}: {w:.2f}" for col, w in larguras_finais.items()])
    console.print(f"[green]✓ Larguras definidas:[/green] [bold]{resumo}[/bold]\n")
    return larguras_finais


def selecionar_paginas(total_paginas: int) -> list:
    if total_paginas <= 1:
        console.print("\n[bold cyan]📄 Seleção de Páginas[/bold cyan] ([green]Página/Imagem única detectada — selecionada automaticamente[/green])")
        return [1]
    console.print(f"\n[bold cyan]📄 Seleção de Páginas[/bold cyan] (Total no documento: [bold green]{total_paginas}[/bold green])")
    console.print("[1] Todas as páginas")
    console.print("[2] Informar intervalo específico (ex: 1-5 ou 3,4,8)")
    
    opcao = Prompt.ask("Escolha [1/2 ou digite o intervalo direto]", default="1").strip()
    if opcao == "1" or opcao.lower() in ["todas", "todos", "all", "*"]:
        return list(range(1, total_paginas + 1))

    intervalo_inicial = opcao if opcao != "2" else ""
    while True:
        intervalo = intervalo_inicial or Prompt.ask("Digite as páginas (ex: [green]1-10[/green] ou [green]2,5,7[/green])")
        intervalo_inicial = ""
        paginas = set()
        try:
            partes = intervalo.split(",")
            for p in partes:
                p = p.strip()
                if "-" in p:
                    ini, fim = p.split("-")
                    for num in range(int(ini), int(fim) + 1):
                        if 1 <= num <= total_paginas:
                            paginas.add(num)
                elif p.isdigit():
                    num = int(p)
                    if 1 <= num <= total_paginas:
                        paginas.add(num)
            if paginas:
                return sorted(list(paginas))
            else:
                console.print("[red]Nenhuma página válida encontrada dentro do limite do arquivo.[/red]")
        except Exception:
            console.print("[red]Formato inválido. Use números e hífens como '1-5' ou '1,2,3'.[/red]")

def exibir_diagnostico_leitura_automatica(leitor: LeitorPDF, paginas_por_pdf: dict) -> bool:
    """
    Analisa automaticamente a estrutura interna das páginas selecionadas (Vetorial vs Rasterizado/Curvas/OCR Oculto)
    e define o modo ideal página por página sem exigir escolha manual do usuário.
    """
    total_ocr = 0
    total_vet = 0
    motivos_globais = set()

    for arq, pags in paginas_por_pdf.items():
        try:
            diag = leitor.diagnosticar_modo_pdf(arq, paginas=pags)
            total_ocr += len(diag.get("paginas_ocr", []))
            total_vet += len(diag.get("paginas_vetorial", []))
            for m in diag.get("motivos_ocr", []):
                motivos_globais.add(m)
        except Exception:
            pass

    if total_ocr > 0 and total_vet == 0:
        motivo_str = ", ".join(sorted(motivos_globais)) or "PDF Rasterizado / Texto em Curvas"
        console.print(
            f"\n[bold cyan]🤖 Detecção Automática de Leitura:[/bold cyan] [yellow]{motivo_str}[/yellow] "
            f"({total_ocr} página(s)) ➔ [bold green]OCR por Colunas (300 DPI) ativado automaticamente![/bold green]"
        )
    elif total_vet > 0 and total_ocr == 0:
        console.print(
            f"\n[bold cyan]🤖 Detecção Automática de Leitura:[/bold cyan] [green]Texto Vetorial Nativo Íntegro[/green] "
            f"({total_vet} página(s)) ➔ [bold green]Extração Geométrica Estruturada ativada automaticamente![/bold green]"
        )
    elif total_vet > 0 and total_ocr > 0:
        console.print(
            f"\n[bold cyan]🤖 Detecção Automática de Leitura (PDF Híbrido):[/bold cyan] "
            f"[green]{total_vet} página(s) Vetorial Nativa[/green] + [yellow]{total_ocr} página(s) via OCR 300 DPI[/yellow] selecionadas automaticamente!"
        )
    return False


def formatar_nota_codigos_repetidos(info: dict) -> str:
    """
    Se a página do PDF possuir linhas com códigos repetidos (ex: 14 códigos na tabela, mas 6 repetidos
    resultando em 8 códigos extraídos), retorna apenas o número real na página e a quantidade repetida.
    """
    info_rep = (info or {}).get("info_repetidos") or {}
    qtd_rep = info_rep.get("qtd_repetidos", 0)
    if qtd_rep > 0:
        tot_linhas = info_rep.get("total_linhas", 0)
        return f" [yellow]({tot_linhas} código(s) na página, {qtd_rep} repetido(s))[/yellow]"
    return ""


def executar_mapeamento_termos(caminhos_pdf, leitor: LeitorPDF, tradutor: TradutorIA, total_paginas: int = None) -> bool:
    lista_pdfs = [caminhos_pdf] if isinstance(caminhos_pdf, str) else list(caminhos_pdf)
    console.print(f"\n[bold cyan]📖 Mapeamento de Vocabulário e Termos Técnicos ({len(lista_pdfs)} PDF(s))[/bold cyan]")
    console.print("[dim]Esta opção analisa o conteúdo do(s) manual(is) para extrair novos termos em inglês e sugerir a tradução técnica ideal.[/dim]\n")

    paginas_por_pdf = {}
    if len(lista_pdfs) == 1:
        tot = total_paginas or leitor.obter_total_paginas(lista_pdfs[0])
        paginas_por_pdf[lista_pdfs[0]] = selecionar_paginas(tot)
    else:
        console.print("[1] Ler todas as páginas de todos os PDFs selecionados [dim][Padrão][/dim]")
        console.print("[2] Escolher o intervalo de páginas para cada PDF individualmente")
        op_p = Prompt.ask("Escolha", choices=["1", "2"], default="1")
        for arq in lista_pdfs:
            tot = leitor.obter_total_paginas(arq)
            if op_p == "1":
                paginas_por_pdf[arq] = list(range(1, tot + 1))
            else:
                console.print(f"\n[bold yellow]📄 Arquivo: {os.path.basename(arq)}[/bold yellow]")
                paginas_por_pdf[arq] = selecionar_paginas(tot)

    forcar_ocr = exibir_diagnostico_leitura_automatica(leitor, paginas_por_pdf)

    textos_combinados = []
    for arq in lista_pdfs:
        pags_sel = paginas_por_pdf.get(arq, [])
        console.print(f"\n[bold green]🔍 Lendo {len(pags_sel)} página(s) de '{os.path.basename(arq)}' para analisar vocabulário...[/bold green]")
        for num_pag in pags_sel:
            info_list = leitor.ler_documento(arq, paginas=[num_pag], forcar_ocr=forcar_ocr)
            if info_list:
                textos_combinados.append(info_list[0]["texto"])
            
    texto_total = "\n\n".join(textos_combinados)
    if len(texto_total.strip()) < 50:
        console.print("[red]Não foi possível extrair texto significativo do documento.[/red]")
        return False
        
    console.print("[blue]🤖 Consultando IA para identificar novos termos técnicos e siglas...[/blue]")
    resultado = tradutor.extrair_novos_termos(texto_total)
    termos_sugeridos = resultado.get("termos_sugeridos", {})
    siglas_detectadas = resultado.get("siglas_detectadas", [])
    
    if not termos_sugeridos and not siglas_detectadas:
        console.print("[yellow]Nenhum novo termo técnico relevante foi identificado no trecho selecionado.[/yellow]")
        return False
        
    glossario_atual = tradutor.carregar_dados_glossario()
    termos_existentes = glossario_atual.get("termos_e_traducoes", {})
    siglas_existentes = set(glossario_atual.get("siglas_preservadas", []))
    
    tabela = Table(title="Termos Técnicos Identificados no Manual", show_header=True, header_style="bold magenta")
    tabela.add_column("Termo Original (Inglês)", style="cyan", width=35)
    tabela.add_column("Tradução Técnica Sugerida (PT-BR)", style="green", width=45)
    tabela.add_column("Status", justify="center")
    
    novos_count = 0
    for k, v in termos_sugeridos.items():
        if k in termos_existentes:
            status = "[dim]Já no Glossário[/dim]"
        else:
            status = "[bold green]+ Novo Termo[/bold green]"
            novos_count += 1
        tabela.add_row(k, v, status)
        
    console.print(tabela)
    
    novas_siglas = [s for s in siglas_detectadas if s not in siglas_existentes]
    if siglas_detectadas:
        console.print(f"\n[cyan]Siglas técnicas detectadas:[/cyan] {', '.join(siglas_detectadas)}")
        if novas_siglas:
            console.print(f"[green]Novas siglas a adicionar:[/green] {', '.join(novas_siglas)}")
            
    console.print(f"\n[bold]{novos_count}[/bold] novos termos e [bold]{len(novas_siglas)}[/bold] novas siglas foram detectados.")
    salvar = Confirm.ask("\nDeseja salvar/mesclar esses termos no arquivo [bold yellow]config/glossario_tecnico.json[/bold yellow]?", default=True)
    
    if salvar:
        termos_existentes.update(termos_sugeridos)
        siglas_total = sorted(list(siglas_existentes.union(novas_siglas)))
        glossario_atual["termos_e_traducoes"] = termos_existentes
        glossario_atual["siglas_preservadas"] = siglas_total
        tradutor.salvar_dados_glossario(glossario_atual)
        console.print("[bold green]✓ Glossário técnico atualizado com sucesso no arquivo config/glossario_tecnico.json![/bold green]\n")
    
    prosseguir_extracao = Confirm.ask("Deseja prosseguir agora com a extração e criação da planilha Excel deste PDF?", default=True)
    return prosseguir_extracao


def parsear_selecao_modulos(entrada: str, modulos_disponiveis: list) -> list:
    """
    Interpreta a escolha do usuário para seleção de módulos, funcionando igual à seleção
    de páginas e documentos:
    - '0', 'todos', 'all', '*' -> todos os módulos da tabela
    - Números e intervalos da tabela -> ex: '1', '1,2', '1-3', '1, 3-5'
    - Nomes de módulos digitados manualmente -> ex: 'HCML, HCMR' ou 'EM_SH, IC M06'
    """
    import re
    if not entrada or not entrada.strip():
        return [m["nome"] for m in modulos_disponiveis] if modulos_disponiveis else []

    texto = entrada.strip().strip('"\'')
    if texto.lower() in ["0", "todos", "todas", "all", "*"]:
        return [m["nome"] for m in modulos_disponiveis]

    total_opcoes = len(modulos_disponiveis)
    selecionados = []

    # Divide por vírgula ou ponto-e-vírgula
    partes = [p.strip() for p in re.split(r'[,;]+', texto) if p.strip()]

    # Se o usuário digitou apenas uma parte com " e " (ex: "HCML e HCMR") e não bate com um nome único da tabela
    if len(partes) == 1 and re.search(r'\s+e\s+', partes[0], flags=re.I):
        nomes_tabela_up = {m["nome"].upper() for m in modulos_disponiveis}
        if partes[0].upper() not in nomes_tabela_up:
            partes = [p.strip() for p in re.split(r'\s+e\s+', partes[0], flags=re.I) if p.strip()]

    for p in partes:
        if p.lower() in ["0", "todos", "todas", "all", "*"]:
            for m in modulos_disponiveis:
                if m["nome"] not in selecionados:
                    selecionados.append(m["nome"])
            continue

        # Verifica intervalo numérico (ex: "1-3")
        m_range = re.match(r'^(\d+)\s*-\s*(\d+)$', p)
        if m_range:
            ini, fim = int(m_range.group(1)), int(m_range.group(2))
            if ini > fim:
                ini, fim = fim, ini
            if 1 <= ini and fim <= total_opcoes:
                for idx in range(ini, fim + 1):
                    nome_m = modulos_disponiveis[idx - 1]["nome"]
                    if nome_m not in selecionados:
                        selecionados.append(nome_m)
                continue

        # Verifica número único da tabela (ex: "1", "2")
        if p.isdigit():
            num = int(p)
            if 1 <= num <= total_opcoes:
                nome_m = modulos_disponiveis[num - 1]["nome"]
                if nome_m not in selecionados:
                    selecionados.append(nome_m)
                continue

        # Caso contrário, trata como nome literal de módulo digitado pelo usuário
        if p not in selecionados:
            selecionados.append(p)

    return selecionados


def obter_letra_codigo(cod_str: str) -> str:
    """
    Retorna a letra/prefixo inicial de um código de falha (ex: 'B163502' -> 'B',
    'C000100' -> 'C', 'P05E000' -> 'P', 'U013104' -> 'U').
    Se o código for puramente numérico/hex iniciando com dígito, retorna '0-9'.
    """
    import re
    if not cod_str:
        return "OUTROS"
    limpo = re.sub(r'^(?:0X|SPN|DTC|PSN|OBD)[\s:#_-]*', '', str(cod_str).strip(), flags=re.I).strip().upper()
    if not limpo:
        return "OUTROS"
    if limpo[0].isalpha():
        return limpo[0]
    if limpo[0].isdigit():
        return "0-9"
    return "OUTROS"


def dividir_registros_por_letra(
    registros: list,
    colunas: list = None,
    letras_selecionadas: list = None,
    modulo_banner: str = None
) -> dict:
    """
    Divide uma lista de registros extraídos em dicionário {letra: [registros]},
    agrupando pela letra inicial do código (ex: B, C, P, U).
    """
    if not registros:
        return {}

    amostra_chaves = colunas if colunas else list(registros[0].keys())
    col_cod = next(
        (c for c in amostra_chaves if any(k in c.lower() for k in ["spn", "psn", "código", "codigo", "dtc", "code", "cod", "identifier"]) and "name" not in c.lower() and "desc" not in c.lower() and "status" not in c.lower()),
        None
    )
    if not col_cod:
        col_cod = amostra_chaves[0] if amostra_chaves else "Código"

    col_mod = next((c for c in amostra_chaves if any(k in c.lower() for k in ["módulo", "modulo", "ecu"])), None)

    grupos_brutos = {}
    for reg in registros:
        val_cod = str(reg.get(col_cod, "")).strip()
        letra = obter_letra_codigo(val_cod)
        copia = dict(reg)
        if col_mod and modulo_banner and (not copia.get(col_mod) or copia.get(col_mod) in ["-", "B", "C", "P", "U"]):
            copia[col_mod] = modulo_banner
        grupos_brutos.setdefault(letra, []).append(copia)

    letras_filtro = [l.upper() for l in (letras_selecionadas or []) if l and l.upper() not in ["0", "TODAS", "TODOS", "ALL", "*"]]

    # Ordem padrão das letras encontradas ou solicitadas
    ordem_letras = []
    if letras_filtro:
        for lf in letras_filtro:
            if lf in grupos_brutos and lf not in ordem_letras:
                ordem_letras.append(lf)
    else:
        for l_padrao in ["B", "C", "P", "U"]:
            if l_padrao in grupos_brutos:
                ordem_letras.append(l_padrao)
        for l_outra in sorted(grupos_brutos.keys()):
            if l_outra not in ordem_letras:
                ordem_letras.append(l_outra)

    return {letra: grupos_brutos[letra] for letra in ordem_letras if grupos_brutos.get(letra)}


def filtrar_registros_por_modulo_texto(registros: list, nome_modulo: str, texto_pagina: str = "", colunas: list = None) -> list:
    """
    Filtra registros de uma página geral mantendo apenas aqueles associados ao módulo 'nome_modulo'
    (seja no registro traduzido ou na linha correspondente do texto bruto da página).
    """
    import re
    if not registros or not nome_modulo:
        return []

    partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', nome_modulo.strip()) if x]
    if not partes_m:
        return registros
    pat = re.compile(r'(?<![A-Za-z0-9])' + r'[\s_-]*'.join(partes_m) + r'(?![A-Za-z0-9])', re.I)

    amostra_chaves = colunas if colunas else list(registros[0].keys())
    col_cod = next(
        (c for c in amostra_chaves if any(k in c.lower() for k in ["spn", "psn", "código", "codigo", "dtc", "code", "cod", "identifier"]) and "name" not in c.lower() and "desc" not in c.lower() and "status" not in c.lower()),
        None
    )
    col_mod = next((c for c in amostra_chaves if any(k in c.lower() for k in ["módulo", "modulo", "ecu"])), None)

    # Mapeia linhas do texto bruto da página por código para verificar menção no inglês original
    contexto_por_codigo = {}
    if texto_pagina:
        bloco_atual_cod = None
        for linha in texto_pagina.splitlines():
            m_c = re.match(r'^\s*([PCBU][0-9A-F]{4,6}|[0-9A-F]{4,6})\b', linha.strip(), re.I)
            if m_c:
                bloco_atual_cod = m_c.group(1).upper()
                contexto_por_codigo[bloco_atual_cod] = context_prev = contexto_por_codigo.get(bloco_atual_cod, "") + " " + linha
            elif bloco_atual_cod:
                contexto_por_codigo[bloco_atual_cod] += " " + linha

    filtrados = []
    for reg in registros:
        texto_reg = " ".join(str(v) for v in reg.values() if v)
        cod_val = re.sub(r'[\s:_-]+', '', str(reg.get(col_cod, "") if col_cod else reg.get("Código", "")).strip()).upper()
        cod_base = str(reg.get("_cod_base", cod_val)).strip().upper()
        texto_orig_cod = contexto_por_codigo.get(cod_val, "") + " " + contexto_por_codigo.get(cod_base, "")

        if pat.search(texto_reg) or pat.search(texto_orig_cod):
            copia = dict(reg)
            if col_mod:
                copia[col_mod] = nome_modulo
            filtrados.append(copia)

    return filtrados


def selecionar_modo_excel_multiplos_modulos(itens: list, rotulo: str = "Módulos", permitir_aba_unica: bool = True) -> str:
    """
    Pergunta ao usuário como deseja salvar os múltiplos módulos ou letras extraídas:
    - [1] No mesmo Excel, porém em páginas/abas diferentes (Padrão)
    - [2] Em arquivos Excel diferentes (um .xlsx para cada módulo/letra)
    - [3] No mesmo Excel e na mesma página/aba (tudo unificado)
    """
    console.print(f"\n[bold cyan]📊 Organização do Excel para {len(itens)} {rotulo} ({', '.join(itens)})[/bold cyan]")
    console.print(f"[1] Mesmo arquivo Excel, porém em [bold green]páginas/abas diferentes[/bold green] (1 aba para cada item) [dim][Padrão][/dim]")
    console.print(f"[2] [bold yellow]Arquivos Excel diferentes[/bold yellow] (1 arquivo .xlsx separado para cada item)")
    opcoes = ["1", "2"]
    if permitir_aba_unica:
        console.print("[3] Mesmo arquivo Excel e na [bold cyan]mesma página/aba[/bold cyan] (todos juntos)")
        opcoes.append("3")

    opcao = Prompt.ask("Escolha como deseja gerar o Excel", choices=opcoes, default="1")
    if opcao == "1":
        console.print("[green]✓ Modo de saída:[/green] [bold]Mesmo Excel em páginas/abas diferentes[/bold]\n")
        return "abas_separadas"
    elif opcao == "2":
        console.print("[green]✓ Modo de saída:[/green] [bold]Arquivos Excel separados[/bold]\n")
        return "arquivos_separados"
    else:
        console.print("[green]✓ Modo de saída:[/green] [bold]Mesmo Excel na mesma página/aba unificada[/bold]\n")
        return "aba_unica"


def selecionar_letras_divisao(resumo_info: dict) -> list:
    """
    Exibe tabela com as letras iniciais de códigos detectadas no PDF (ex: B, C, P, U)
    e permite escolher todas (0/Enter) ou letras específicas.
    """
    letras_det = (resumo_info or {}).get("letras_detectadas", [])
    opcoes_tabela = []
    if letras_det:
        for item in letras_det:
            opcoes_tabela.append({
                "nome": item["letra"],
                "info": f"{item['paginas']} — {item['descricao']}"
            })
    else:
        for l_pad in ["B", "C", "P", "U"]:
            opcoes_tabela.append({
                "nome": l_pad,
                "info": f"Códigos iniciados com a letra {l_pad}"
            })

    tabela_l = Table(title="Letras / Famílias de Códigos DTC no Documento", show_header=True, header_style="bold magenta")
    tabela_l.add_column("Opção", style="cyan", width=10, justify="center")
    tabela_l.add_column("Letra", style="green", width=15, justify="center")
    tabela_l.add_column("Localização e Quantidade no PDF", style="dim")

    for idx, item in enumerate(opcoes_tabela, 1):
        tabela_l.add_row(str(idx), item["nome"], item["info"])
    tabela_l.add_row("0 / todas", "[bold yellow]TODAS AS LETRAS[/bold yellow]", "Cria uma aba separada para cada letra presente no documento")

    console.print()
    console.print(tabela_l)
    entrada_l = Prompt.ask(
        "[bold green]Selecione as letras para dividir[/bold green] (ex: [cyan]0[/cyan] para todas, [cyan]1,2[/cyan] ou [cyan]B, C, P, U[/cyan])",
        default="0"
    )
    letras_escolhidas = [x.upper() for x in parsear_selecao_modulos(entrada_l, opcoes_tabela)]
    if not letras_escolhidas:
        letras_escolhidas = [m["nome"].upper() for m in opcoes_tabela]
    console.print(f"[green]✓ Letras selecionadas para abas ({len(letras_escolhidas)}):[/green] [bold]{' | '.join(letras_escolhidas)}[/bold]")
    return letras_escolhidas


def selecionar_modulos(
    caminho_pdf: str,
    leitor: LeitorPDF,
    tradutor: TradutorIA,
    contexto_operacao: str = "Planilha Padrão"
) -> tuple:
    """
    Detecta automaticamente os módulos e letras de códigos presentes no PDF.
    Só divide os códigos em abas/páginas caso seja solicitado pelo usuário,
    permitindo dividir por MÓDULO ou por LETRA (B, C, P, U).
    Retorna: (lista_itens_selecionados, modo_saida, resumo_info)
    """
    sugestao_modulo = os.path.splitext(os.path.basename(caminho_pdf))[0].replace("_traduzido", "").strip()
    console.print(f"\n[bold cyan]🏷️ Seleção de Módulo(s) / Sistema(s) ({contexto_operacao})[/bold cyan]")
    console.print("[dim]🔍 Analisando o documento para identificar módulos e famílias de códigos...[/dim]")

    try:
        resumo_info = leitor.extrair_resumo_modulos_pdf(caminho_pdf)
        modulos_disponiveis = tradutor.identificar_modulos_documento(
            resumo_texto=resumo_info.get("resumo_texto", ""),
            nome_arquivo=sugestao_modulo,
            modulos_regex=resumo_info.get("modulos_regex", [])
        )
    except Exception:
        resumo_info = {}
        modulos_disponiveis = [{"nome": sugestao_modulo, "info": "Sugestão baseada no nome do arquivo"}]

    if not modulos_disponiveis:
        modulos_disponiveis = [{"nome": sugestao_modulo, "info": "Sugestão baseada no nome do arquivo"}]

    tabela = Table(title="Módulos / Sistemas Identificados no Documento", show_header=True, header_style="bold magenta")
    tabela.add_column("Opção", style="cyan", width=10, justify="center")
    tabela.add_column("Módulo / Sistema", style="green", width=28)
    tabela.add_column("Identificação / Localização", style="dim")

    for i, mod in enumerate(modulos_disponiveis, 1):
        tabela.add_row(str(i), mod["nome"], mod.get("info", "Detectado no documento"))

    if len(modulos_disponiveis) > 1:
        tabela.add_row("0 / todos", "[bold yellow]TODOS OS MÓDULOS ACIMA[/bold yellow]", f"Seleciona todos os {len(modulos_disponiveis)} módulos listados")

    console.print(tabela)

    letras_det = resumo_info.get("letras_detectadas", [])
    if letras_det:
        resumo_l = ", ".join([f"[bold]{l['letra']}[/bold] ({l['paginas']})" for l in letras_det])
        console.print(f"  [dim]💡 Letras de códigos detectadas no PDF:[/dim] {resumo_l}")

    console.print("[dim]Como selecionar:[/dim]")
    console.print("  • Digite o(s) [bold cyan]número(s) ou intervalo[/bold cyan] da tabela (ex: [green]1[/green], [green]1,2[/green], [green]1-3[/green] ou [green]0[/green] para todos)")
    console.print("  • Ou digite o(s) [bold cyan]nome(s) do(s) módulo(s)[/bold cyan] separados por vírgula (ex: [green]HCML, HCMR[/green] ou [green]EM_SH, IC M06[/green])")
    console.print("  • Ou digite diretamente as [bold cyan]letras dos códigos[/bold cyan] se quiser dividir por letra (ex: [green]B, C, P, U[/green])")

    padrao = "0" if len(modulos_disponiveis) > 1 else "1"
    while True:
        entrada = Prompt.ask("[bold green]Selecione o(s) módulo(s)[/bold green]", default=padrao)
        modulos_selecionados = parsear_selecao_modulos(entrada, modulos_disponiveis)
        if modulos_selecionados:
            break
        console.print("[red]Seleção inválida. Informe pelo menos um número da tabela ou nome de módulo.[/red]")

    # Verifica se o usuário digitou diretamente letras de códigos (ex: "B, C, P, U")
    letras_conhecidas = {l["letra"].upper() for l in letras_det} | {"P", "C", "B", "U"}
    digitou_apenas_letras = (
        len(modulos_selecionados) >= 2
        and all(len(m.strip()) == 1 and m.strip().upper() in letras_conhecidas for m in modulos_selecionados)
    )

    if digitou_apenas_letras:
        letras_escolhidas = [m.strip().upper() for m in modulos_selecionados]
        console.print(f"[green]✓ Divisão por Letra solicitada ({len(letras_escolhidas)}):[/green] [bold]{' | '.join(letras_escolhidas)}[/bold]")
        modo_saida = selecionar_modo_excel_multiplos_modulos(letras_escolhidas, rotulo="Letras de Códigos", permitir_aba_unica=False)
        resumo_info["config_divisao"] = {
            "tipo": "por_letra",
            "letras": letras_escolhidas,
            "modulo_banner": sugestao_modulo
        }
        return letras_escolhidas, modo_saida, resumo_info

    console.print(f"[green]✓ Módulo(s) definido(s) ({len(modulos_selecionados)}):[/green] [bold]{' | '.join(modulos_selecionados)}[/bold]")

    # Caso 1: Usuário selecionou múltiplos módulos logo na primeira etapa
    if len(modulos_selecionados) > 1:
        modo_saida = selecionar_modo_excel_multiplos_modulos(modulos_selecionados, rotulo="Módulos", permitir_aba_unica=True)
        resumo_info["config_divisao"] = {
            "tipo": "por_modulo" if modo_saida != "aba_unica" else "nenhuma",
            "letras": [],
            "modulo_banner": " E ".join(modulos_selecionados)
        }
        return modulos_selecionados, modo_saida, resumo_info

    # Caso 2: Usuário selecionou apenas 1 módulo -> pergunta se deseja dividir em abas (por Letra ou por Módulo)
    modulo_principal = modulos_selecionados[0]
    console.print(f"\n[bold cyan]📊 Divisão de Códigos em Abas/Páginas no Excel[/bold cyan]")
    console.print("[1] [bold green]Não dividir[/bold green] — Manter todos os códigos em uma única aba/página [dim][Padrão][/dim]")
    info_letras_str = f" ({', '.join(l['letra'] for l in letras_det)})" if letras_det else " (ex: B, C, P, U)"
    console.print(f"[2] Dividir [bold cyan]por Letra do código[/bold cyan]{info_letras_str} em abas separadas")
    console.print("[3] Dividir [bold yellow]por Módulo[/bold yellow] em abas separadas")

    op_div = Prompt.ask("Deseja dividir os códigos no Excel?", choices=["1", "2", "3"], default="1")

    if op_div == "2":
        letras_escolhidas = selecionar_letras_divisao(resumo_info)
        modo_saida = selecionar_modo_excel_multiplos_modulos(letras_escolhidas, rotulo="Letras de Códigos", permitir_aba_unica=False)
        resumo_info["config_divisao"] = {
            "tipo": "por_letra",
            "letras": letras_escolhidas,
            "modulo_banner": modulo_principal
        }
        return letras_escolhidas, modo_saida, resumo_info

    elif op_div == "3":
        modulos_citados = resumo_info.get("modulos_citados", [])
        opcoes_mod = []
        if len(modulos_disponiveis) > 1:
            opcoes_mod = modulos_disponiveis
        elif modulos_citados:
            for mc in modulos_citados:
                opcoes_mod.append({"nome": mc["nome"], "info": mc["descricao"]})

        if opcoes_mod:
            tab_m = Table(title="Módulos Detectados nas Descrições do PDF", show_header=True, header_style="bold magenta")
            tab_m.add_column("Opção", style="cyan", width=10, justify="center")
            tab_m.add_column("Módulo", style="green", width=20)
            tab_m.add_column("Ocorrência no Documento", style="dim")
            for idx, m_item in enumerate(opcoes_mod, 1):
                tab_m.add_row(str(idx), m_item["nome"], m_item["info"])
            tab_m.add_row("0 / todos", "[bold yellow]TODOS ACIMA[/bold yellow]", f"Seleciona todos os {len(opcoes_mod)} módulos")
            console.print()
            console.print(tab_m)

        while True:
            ent_m = Prompt.ask(
                "[bold green]Digite os números ou nomes dos módulos separados por vírgula[/bold green] (ex: [cyan]1,2[/cyan] ou [cyan]HCML, HCMR[/cyan])",
                default="0" if opcoes_mod else ""
            )
            lista_mods = parsear_selecao_modulos(ent_m, opcoes_mod)
            if lista_mods:
                break
            console.print("[red]Informe pelo menos um módulo válido.[/red]")

        console.print(f"[green]✓ Módulos selecionados para divisão ({len(lista_mods)}):[/green] [bold]{' | '.join(lista_mods)}[/bold]")
        modo_saida = selecionar_modo_excel_multiplos_modulos(lista_mods, rotulo="Módulos", permitir_aba_unica=False)
        resumo_info["config_divisao"] = {
            "tipo": "por_modulo",
            "letras": [],
            "modulo_banner": modulo_principal
        }
        return lista_mods, modo_saida, resumo_info

    # Padrão: Não dividir (tudo em 1 única aba)
    console.print("[green]✓ Modo de saída:[/green] [bold]Aba única (sem divisão)[/bold]\n")
    resumo_info["config_divisao"] = {
        "tipo": "nenhuma",
        "letras": [],
        "modulo_banner": modulo_principal
    }
    return modulos_selecionados, "aba_unica", resumo_info


def mapear_paginas_por_modulo(
    modulos_selecionados: list,
    paginas_selecionadas: list,
    resumo_info: dict
) -> dict:
    """
    Mapeia quais páginas selecionadas devem ser processadas para cada módulo e adiciona
    contexto de continuação de seção quando uma página continua a tabela de um módulo
    iniciado em página anterior.
    Retorna: {nome_modulo: [(num_pag, nota_contexto), ...]}
    """
    import re
    if len(modulos_selecionados) <= 1:
        unico = modulos_selecionados[0] if modulos_selecionados else "SISTEMA"
        return {unico: [(p, "") for p in paginas_selecionadas]}

    textos_por_pagina = (resumo_info or {}).get("textos_por_pagina", {})
    codigos_gt_por_pagina = (resumo_info or {}).get("codigos_gt_por_pagina", {})
    modulos_regex = (resumo_info or {}).get("modulos_regex", [])

    # Se o PDF não possui cabeçalhos/colunas estruturais de módulo (ex: tabela geral sem 'ECU Name'),
    # todas as páginas selecionadas pertencem integralmente a todos os módulos escolhidos.
    if not textos_por_pagina or not modulos_regex:
        return {m: [(p, "") for p in paginas_selecionadas] for m in modulos_selecionados}

    # Verifica quais módulos selecionados correspondem a módulos estruturais detectados no PDF
    nomes_estruturais_up = {m_info["nome"].upper(): m_info for m_info in modulos_regex}
    padroes_mod = {}
    modulos_presentes_no_pdf = set()
    for mod in modulos_selecionados:
        mod_up = mod.strip().upper()
        partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', mod.strip()) if x]
        pat_str = r'\b' + r'[\s_-]*'.join(partes_m) + r'\b'
        pat = re.compile(pat_str, re.I)
        padroes_mod[mod] = pat
        if mod_up in nomes_estruturais_up:
            modulos_presentes_no_pdf.add(mod)

    # Se nenhum dos módulos corresponde a cabeçalhos/colunas estruturais do PDF, processa todas as páginas selecionadas para cada um
    if not modulos_presentes_no_pdf:
        return {m: [(p, "") for p in paginas_selecionadas] for m in modulos_selecionados}

    mapa = {m: [] for m in modulos_selecionados}
    paginas_set = set(paginas_selecionadas)
    secao_ativa = []

    for num_pag in sorted(textos_por_pagina.keys()):
        txt = textos_por_pagina.get(num_pag, "")
        gt_pag = codigos_gt_por_pagina.get(num_pag, [])
        linhas = txt.splitlines()

        # Identifica em qual linha cada módulo aparece pela primeira vez nesta página
        primeira_linha_mod = {}
        ordem_aparicao = []
        for idx_l, linha in enumerate(linhas):
            for mod in modulos_selecionados:
                if mod in modulos_presentes_no_pdf and padroes_mod[mod].search(linha):
                    if mod not in primeira_linha_mod:
                        primeira_linha_mod[mod] = idx_l
                    if mod not in ordem_aparicao:
                        ordem_aparicao.append(mod)

        if primeira_linha_mod:
            # Verifica se há códigos de falha no topo da página ANTES do primeiro cabeçalho de novo módulo
            min_linha_mod = min(primeira_linha_mod.values())
            codigos_antes = False
            if min_linha_mod > 0 and secao_ativa and gt_pag:
                texto_topo = "\n".join(linhas[:min_linha_mod])
                for c_gt in gt_pag:
                    if re.search(r'\b' + re.escape(c_gt) + r'\b', texto_topo):
                        codigos_antes = True
                        break

            if num_pag in paginas_set:
                if codigos_antes:
                    for m_ant in secao_ativa:
                        if m_ant not in primeira_linha_mod:
                            nota = f"[Contexto Estrutural: O início desta página é continuação da tabela do módulo '{m_ant}' iniciada na página anterior.]\n"
                            mapa[m_ant].append((num_pag, nota))
                for m_atual in ordem_aparicao:
                    mapa[m_atual].append((num_pag, ""))

            # A seção ativa ao final da página passa a ser o último módulo iniciado/presente nela
            secao_ativa = [ordem_aparicao[-1]] if len(ordem_aparicao) == 1 else list(ordem_aparicao)
        else:
            # Página de continuação sem nome explícito de módulo: herda a seção ativa da página anterior
            if num_pag in paginas_set:
                alvos = secao_ativa if secao_ativa else list(modulos_selecionados)
                for m_alvo in alvos:
                    nota = f"[Contexto Estrutural: Esta página é continuação direta da seção do módulo '{m_alvo}'.]\n" if secao_ativa else ""
                    mapa[m_alvo].append((num_pag, nota))

    # Garante que módulos customizados ou sem páginas mapeadas recebam as páginas selecionadas
    for mod in modulos_selecionados:
        if not mapa[mod]:
            mapa[mod] = [(p, "") for p in paginas_selecionadas]

    return mapa


def executar_excel_buscar_codigo(
    caminho_pdf: str,
    leitor: LeitorPDF,
    tradutor: TradutorIA,
    agente_organizador: AgenteOrganizador,
    gerador_excel: GeradorExcel,
    total_paginas: int,
    config: dict = None
):
    import re
    # 1. Seleção de um ou vários módulos
    modulos_selecionados, modo_saida, resumo_info = selecionar_modulos(
        caminho_pdf=caminho_pdf,
        leitor=leitor,
        tradutor=tradutor,
        contexto_operacao="Buscar Código"
    )

    colunas = ["Código", "Localização da Avaria", "Causa", "Descrição"]
    larguras_colunas = None  # Mantém fixo o padrão oficial do Buscar Código sem perguntar modificações
    console.print("[bold cyan]📋 Estrutura da Planilha 'Excel do Buscar Código' (Padrão Fixo Oficial)[/bold cyan]")
    console.print("[dim]Regra: Código + FMI unificados (ex: 4002101, 40B1608) | Larguras: 15.86 | 30.86 | 40.29 | 34.14[/dim]")
    console.print(f"[green]✓ Colunas fixas:[/green] [bold]{' | '.join(colunas)}[/bold]\n")

    # 3. Seleção de páginas
    paginas_selecionadas = selecionar_paginas(total_paginas)

    # 4. Detecção Automática de Leitura (Vetorial vs OCR 300 DPI)
    forcar_ocr = exibir_diagnostico_leitura_automatica(leitor, {caminho_pdf: paginas_selecionadas})

    console.print(
        f"\n[bold green]🚀 Iniciando processamento do Buscar Código ({len(paginas_selecionadas)} página(s)) "
        f"para {len(modulos_selecionados)} módulo(s): {', '.join(modulos_selecionados)}...[/bold green]\n"
    )

    # Lê e faz cache das páginas selecionadas uma única vez
    paginas_cache = {}
    for i, num_pag in enumerate(paginas_selecionadas, 1):
        console.print(f"[cyan][{i}/{len(paginas_selecionadas)}][/cyan] Lendo página [bold]{num_pag}[/bold]...")
        paginas_info = leitor.ler_documento(caminho_pdf, paginas=[num_pag], forcar_ocr=forcar_ocr)
        if not paginas_info:
            continue
        info = paginas_info[0]
        texto = info["texto"]
        metodo = info["metodo"]
        codigos_gt = info.get("codigos_ground_truth", [])
        console.print(f"   ↳ Modo: [dim]{metodo}[/dim] ({len(texto)} caracteres)")
        if codigos_gt:
            console.print(f"   ↳ [cyan]Ground Truth:[/cyan] {len(codigos_gt)} código(s) de falha verificados diretamente no documento{formatar_nota_codigos_repetidos(info)}")
        if len(texto.strip()) < 10:
            console.print("   ↳ [yellow]Página sem texto significativo identificado, pulando.[/yellow]")
            continue
        paginas_cache[num_pag] = info

    if not paginas_cache:
        console.print("\n[bold red]Nenhuma página com texto válido foi lida do PDF.[/bold red]")
        return

    config_divisao = (resumo_info or {}).get("config_divisao", {})
    tipo_divisao = config_divisao.get("tipo", "por_modulo" if len(modulos_selecionados) > 1 else "nenhuma")
    modulo_banner = config_divisao.get("modulo_banner") or os.path.splitext(os.path.basename(caminho_pdf))[0]
    tabela_geral_sem_modulo = not bool((resumo_info or {}).get("modulos_regex"))

    dados_por_modulo = {}

    # MODO A: Divisão por Letra do código (ex: B, C, P, U)
    if tipo_divisao == "por_letra":
        letras_alvo = config_divisao.get("letras") or modulos_selecionados
        letras_alvo_set = {l.upper() for l in letras_alvo}
        todos_extraidos = []

        for num_pag, info in paginas_cache.items():
            codigos_gt = info.get("codigos_ground_truth", [])
            # Se a página possui Ground Truth e nenhuma das letras selecionadas está presente nela, pula chamada desnecessária
            if codigos_gt and letras_alvo_set:
                letras_pag = {obter_letra_codigo(c) for c in codigos_gt}
                if not (letras_pag & letras_alvo_set):
                    console.print(f"   ↳ [dim]Página {num_pag}: sem códigos das letras {', '.join(letras_alvo)}, pulando.[/dim]")
                    continue

            console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Extraindo e unificando Códigos + FMI...[/blue]")
            try:
                dados_pagina = tradutor.estruturar_buscar_codigo(
                    texto_bruto=info["texto"],
                    contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                    modulo_alvo=None,
                    texto_nativo=info.get("texto_nativo"),
                    codigos_ground_truth=codigos_gt
                )
                console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                todos_extraidos.extend(dados_pagina)
            except Exception as e:
                console.print(f"   ↳ [red]Erro na página {num_pag}: {e}[/red]")

        dados_por_modulo = dividir_registros_por_letra(
            registros=todos_extraidos,
            colunas=colunas,
            letras_selecionadas=letras_alvo,
            modulo_banner=modulo_banner
        )
        for letra_k in list(dados_por_modulo.keys()):
            dados_por_modulo[letra_k] = AgenteOrganizador.organizar_buscar_codigo(dados_por_modulo[letra_k])
            console.print(f"   ↳ [cyan]🔢 Aba Letra '{letra_k}': {len(dados_por_modulo[letra_k])} código(s) organizado(s).[/cyan]")

    # MODO B: Divisão por Módulo ou Módulo Único
    else:
        mapa_paginas = mapear_paginas_por_modulo(modulos_selecionados, paginas_selecionadas, resumo_info)
        dados_por_modulo = {mod: [] for mod in modulos_selecionados}
        cache_buscar_codigo_geral = {}

        # Verifica se os módulos selecionados aparecem citados nas descrições de uma tabela geral
        textos_pags = (resumo_info or {}).get("textos_por_pagina", {})
        modulos_citados_no_texto = set()
        if tabela_geral_sem_modulo and len(modulos_selecionados) > 1:
            for mod in modulos_selecionados:
                partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', mod.strip()) if x]
                if partes_m:
                    pat_m = re.compile(r'(?<![A-Za-z0-9])' + r'[\s_-]*'.join(partes_m) + r'(?![A-Za-z0-9])', re.I)
                    if any(pat_m.search(info["texto"]) for info in paginas_cache.values()):
                        modulos_citados_no_texto.add(mod)

        for idx_m, nome_modulo in enumerate(modulos_selecionados, 1):
            pags_modulo = [(p, nota) for (p, nota) in mapa_paginas.get(nome_modulo, []) if p in paginas_cache]
            console.print(f"\n[bold magenta]🏷️ [{idx_m}/{len(modulos_selecionados)}] Extraindo módulo '{nome_modulo}' ({len(pags_modulo)} página(s))...[/bold magenta]")

            for num_pag, nota_contexto in pags_modulo:
                info = paginas_cache[num_pag]
                texto_envio = f"{nota_contexto}{info['texto']}" if nota_contexto else info["texto"]
                codigos_gt = info.get("codigos_ground_truth", [])

                # Se estamos dividindo por módulos citados no texto e esta página não menciona este módulo, pula
                if nome_modulo in modulos_citados_no_texto:
                    partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', nome_modulo.strip()) if x]
                    pat_m = re.compile(r'(?<![A-Za-z0-9])' + r'[\s_-]*'.join(partes_m) + r'(?![A-Za-z0-9])', re.I)
                    if not pat_m.search(info["texto"]):
                        continue

                if tabela_geral_sem_modulo and num_pag in cache_buscar_codigo_geral:
                    base_pag = [dict(r) for r in cache_buscar_codigo_geral[num_pag]]
                    if modulos_citados_no_texto:
                        dados_pagina = filtrar_registros_por_modulo_texto(base_pag, nome_modulo, info["texto"], colunas=colunas)
                    else:
                        dados_pagina = base_pag
                    console.print(f"   ↳ [green]✓ Página {num_pag}: {len(dados_pagina)} código(s) filtrado(s) para '{nome_modulo}'![/green]")
                    dados_por_modulo[nome_modulo].extend(dados_pagina)
                    continue

                console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Extraindo e unificando Códigos + FMI para '{nome_modulo}'...[/blue]")
                try:
                    alvo_ia = None if tabela_geral_sem_modulo else nome_modulo
                    dados_pagina = tradutor.estruturar_buscar_codigo(
                        texto_bruto=texto_envio,
                        contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                        modulo_alvo=alvo_ia,
                        texto_nativo=info.get("texto_nativo"),
                        codigos_ground_truth=codigos_gt
                    )
                    if tabela_geral_sem_modulo and dados_pagina:
                        cache_buscar_codigo_geral[num_pag] = [dict(r) for r in dados_pagina]
                        if modulos_citados_no_texto:
                            dados_pagina = filtrar_registros_por_modulo_texto(dados_pagina, nome_modulo, info["texto"], colunas=colunas)
                    console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) para '{nome_modulo}' na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                    dados_por_modulo[nome_modulo].extend(dados_pagina)
                except Exception as e:
                    console.print(f"   ↳ [red]Erro na página {num_pag} (módulo '{nome_modulo}'): {e}[/red]")

            if dados_por_modulo[nome_modulo]:
                console.print(f"   ↳ [cyan]🔢 Organizando {len(dados_por_modulo[nome_modulo])} código(s) do módulo '{nome_modulo}'...[/cyan]")
                dados_por_modulo[nome_modulo] = AgenteOrganizador.organizar_buscar_codigo(dados_por_modulo[nome_modulo])

    modulos_com_dados = {m: d for m, d in dados_por_modulo.items() if d}
    if not modulos_com_dados:
        console.print("\n[bold red]Nenhum dado estruturado foi gerado. Verifique o conteúdo do PDF ou a chave de API.[/bold red]")
        return

    try:
        pasta_saida_atual = obter_pasta_saida_para_arquivo(caminho_pdf)

        # Caso 1: Múltiplas abas/páginas diferentes no mesmo Excel (por Módulo ou por Letra)
        if modo_saida == "abas_separadas" and (len(modulos_com_dados) > 1 or tipo_divisao in ["por_letra", "por_modulo"]):
            if tipo_divisao == "por_letra":
                base_arq = re.sub(r'[\\/*?:"<>|]', "", modulo_banner).strip().replace(" ", "_") or os.path.splitext(os.path.basename(caminho_pdf))[0]
                caminho_excel = os.path.join(pasta_saida_atual, f"{base_arq}_buscar_codigo.xlsx")
            else:
                nomes_juntos = "_".join([re.sub(r'[\\/*?:"<>|]', "", m).strip().replace(" ", "_") for m in modulos_com_dados.keys()])
                if len(nomes_juntos) > 60 or not nomes_juntos:
                    nomes_juntos = os.path.splitext(os.path.basename(caminho_pdf))[0]
                caminho_excel = os.path.join(pasta_saida_atual, f"{nomes_juntos}_buscar_codigo.xlsx")

            console.print(f"\n[bold cyan]📊 Gerando planilha 'Buscar Código' com {len(modulos_com_dados)} abas/páginas separadas...[/bold cyan]")
            caminho_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                dados_por_modulo=modulos_com_dados,
                caminho_saida=caminho_excel,
                larguras_personalizadas=larguras_colunas
            )
            resumo_abas = "\n".join([f"  • Aba [bold]{m}[/bold]: {len(d)} códigos" for m, d in modulos_com_dados.items()])
            total_geral = sum(len(d) for d in modulos_com_dados.values())
            console.print(Panel.fit(
                f"[bold green]✨ Planilha 'Buscar Código' (Múltiplas Abas) gerada com sucesso![/bold green]\n\n"
                f"Abas geradas ({len(modulos_com_dados)}):\n{resumo_abas}\n\n"
                f"Total geral de códigos: [bold]{total_geral}[/bold]\n"
                f"Arquivo salvo em: [bold yellow]{caminho_final}[/bold yellow]",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir o arquivo Excel agora?", default=True):
                os.startfile(caminho_final)

        # Caso 2: Arquivos Excel separados (por Módulo ou por Letra)
        elif modo_saida == "arquivos_separados" and (len(modulos_com_dados) > 1 or tipo_divisao in ["por_letra", "por_modulo"]):
            caminhos_gerados = []
            for nome_mod, dados_mod in modulos_com_dados.items():
                nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome_mod).strip().replace(" ", "_") or "modulo"
                if tipo_divisao == "por_letra":
                    pref = re.sub(r'[\\/*?:"<>|]', "", modulo_banner).strip().replace(" ", "_")
                    caminho_excel = os.path.join(pasta_saida_atual, f"{pref}_{nome_limpo}_buscar_codigo.xlsx")
                else:
                    caminho_excel = os.path.join(pasta_saida_atual, f"{nome_limpo}_buscar_codigo.xlsx")
                console.print(f"\n[bold cyan]📊 Gerando Excel 'Buscar Código' separado para '{nome_mod}'...[/bold cyan]")
                c_final = gerador_excel.exportar_buscar_codigo(
                    dados_mod,
                    caminho_excel,
                    nome_modulo=nome_mod,
                    larguras_personalizadas=larguras_colunas
                )
                caminhos_gerados.append((nome_mod, len(dados_mod), c_final))

            resumo_arqs = "\n".join([f"  • [bold]{m}[/bold] ({qtd} códigos): [yellow]{c}[/yellow]" for m, qtd, c in caminhos_gerados])
            console.print(Panel.fit(
                f"[bold green]✨ {len(caminhos_gerados)} Planilhas 'Buscar Código' geradas com sucesso![/bold green]\n\n"
                f"{resumo_arqs}",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir os arquivos Excel gerados agora?", default=True):
                for _, _, c_final in caminhos_gerados:
                    os.startfile(c_final)

        # Caso 3: Módulo único ou todos juntos na mesma página/aba
        else:
            todos_dados = []
            for d in modulos_com_dados.values():
                todos_dados.extend(d)
            todos_dados = AgenteOrganizador.organizar_buscar_codigo(todos_dados)
            titulo_modulo = " E ".join(modulos_com_dados.keys())
            nome_limpo_modulo = re.sub(r'[\\/*?:"<>|]', "", titulo_modulo).strip().replace(" ", "_")
            if not nome_limpo_modulo or len(nome_limpo_modulo) > 60:
                nome_limpo_modulo = os.path.splitext(os.path.basename(caminho_pdf))[0]

            caminho_excel = os.path.join(pasta_saida_atual, f"{nome_limpo_modulo}_buscar_codigo.xlsx")
            console.print(f"\n[bold cyan]📊 Gerando planilha 'Buscar Código' para '{titulo_modulo}'...[/bold cyan]")
            caminho_final = gerador_excel.exportar_buscar_codigo(
                todos_dados,
                caminho_excel,
                nome_modulo=titulo_modulo,
                larguras_personalizadas=larguras_colunas
            )
            console.print(Panel.fit(
                f"[bold green]✨ Planilha 'Buscar Código' gerada com sucesso![/bold green]\n\n"
                f"Módulo(s): [bold]{titulo_modulo}[/bold]\n"
                f"Total de códigos: [bold]{len(todos_dados)}[/bold]\n"
                f"Arquivo salvo em: [bold yellow]{caminho_final}[/bold yellow]",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir o arquivo Excel agora?", default=True):
                os.startfile(caminho_final)

    except Exception as e:
        console.print(f"[bold red]Erro ao gerar Excel: {e}[/bold red]")


def executar_lote_multiplos_pdfs(
    caminhos_pdf: list,
    operacao: str,
    leitor: LeitorPDF,
    tradutor: TradutorIA,
    agente_organizador: AgenteOrganizador,
    gerador_excel: GeradorExcel,
    config: dict,
    padrao_predefinido: dict = None
):
    """
    Processa múltiplos arquivos PDF nas Operações 1 (Planilha Padrão), 3 (Buscar Código) e 4 (Exportação Dupla),
    permitindo configurar uma única vez (em lote) ou individualmente, e escolher entre:
    - 1 arquivo Excel separado para cada PDF
    - 1 único arquivo Excel com 1 aba para cada PDF
    - 1 único arquivo Excel unificando todos os PDFs (em aba única ou dividido por letra B, C, P, U)
    """
    import re
    nomes_pdfs = [os.path.basename(c) for c in caminhos_pdf]
    if operacao == "4":
        rotulo_op = "Exportação Dupla (Padrão + Buscar Código)"
    elif operacao == "1":
        rotulo_op = "Planilha Padrão"
    else:
        rotulo_op = "Buscar Código"

    console.print(f"\n[bold cyan]📚 Modo de Processamento para os {len(caminhos_pdf)} PDFs Selecionados ({rotulo_op})[/bold cyan]")
    console.print(f"[dim]Arquivos: {', '.join(nomes_pdfs)}[/dim]")
    console.print("[1] Gerar [bold green]1 arquivo Excel separado para cada PDF[/bold green] (configurando colunas/larguras uma única vez) [dim][Padrão][/dim]")
    console.print("[2] Unificar todos os PDFs no [bold cyan]mesmo arquivo Excel em abas/páginas separadas[/bold cyan] (1 aba para cada PDF)")
    console.print("[3] Unificar todos os PDFs no [bold yellow]mesmo arquivo Excel[/bold yellow] (aba única ou dividido por Módulo / Letra)")
    console.print("[4] Configurar cada PDF [bold magenta]individualmente passo a passo[/bold magenta]")

    modo_lote = Prompt.ask("Escolha como deseja processar os múltiplos PDFs", choices=["1", "2", "3", "4"], default="1")

    if modo_lote == "4":
        for idx_p, arq_pdf in enumerate(caminhos_pdf, 1):
            console.print(Panel.fit(
                f"[bold yellow]📄 [{idx_p}/{len(caminhos_pdf)}] Configurando PDF: {os.path.basename(arq_pdf)}[/bold yellow]",
                border_style="yellow"
            ))
            tot_p = leitor.obter_total_paginas(arq_pdf)
            if operacao == "1":
                executar_extracao_padrao(arq_pdf, leitor, tradutor, agente_organizador, gerador_excel, tot_p, config)
            elif operacao == "4":
                executar_extracao_padrao(arq_pdf, leitor, tradutor, agente_organizador, gerador_excel, tot_p, config, exportar_duplo=True, padrao_predefinido=padrao_predefinido)
            else:
                executar_excel_buscar_codigo(arq_pdf, leitor, tradutor, agente_organizador, gerador_excel, tot_p, config)
        return

    # Pergunta se deseja dividir os códigos (Não dividir / Por Letra / Por Módulo) nos modos 1 e 3
    tipo_divisao_lote = "nenhuma"
    letras_lote_selecionadas = []
    modulos_lote_selecionados = []

    if modo_lote in ["1", "3"]:
        console.print("\n[bold cyan]📊 Divisão de Códigos em Abas/Páginas no Excel[/bold cyan]")
        console.print("[1] [bold green]Não dividir[/bold green] — Manter os códigos juntos na aba [dim][Padrão][/dim]")
        console.print("[2] Dividir [bold cyan]por Letra do código (B, C, P, U)[/bold cyan] em abas separadas")
        console.print("[3] Dividir [bold yellow]por Módulo[/bold yellow] em abas separadas")
        op_div_lote = Prompt.ask("Deseja dividir os códigos no Excel?", choices=["1", "2", "3"], default="1")

        if op_div_lote in ["2", "3"]:
            # Faz varredura rápida nos PDFs do lote para agregar letras e módulos detectados
            letras_agregadas = {}
            modulos_agregados = {}
            for arq_p in caminhos_pdf:
                try:
                    res_p = leitor.extrair_resumo_modulos_pdf(arq_p)
                    nome_arq_curto = os.path.basename(arq_p)
                    for item_l in res_p.get("letras_detectadas", []):
                        lk = item_l["letra"]
                        if lk not in letras_agregadas:
                            letras_agregadas[lk] = {"letra": lk, "paginas": f"{nome_arq_curto} ({item_l['paginas']})", "descricao": item_l["descricao"]}
                        else:
                            letras_agregadas[lk]["paginas"] += f", {nome_arq_curto} ({item_l['paginas']})"
                    for item_m in res_p.get("modulos_regex", []) + res_p.get("modulos_citados", []):
                        mk = item_m["nome"].upper()
                        if mk not in modulos_agregados:
                            modulos_agregados[mk] = {"nome": item_m["nome"], "info": f"{nome_arq_curto}: {item_m.get('descricao', item_m.get('paginas', ''))}"}
                        else:
                            modulos_agregados[mk]["info"] += f" | {nome_arq_curto}"
                except Exception:
                    pass

            if op_div_lote == "2":
                tipo_divisao_lote = "por_letra"
                resumo_lote_letras = {"letras_detectadas": [letras_agregadas[k] for k in sorted(letras_agregadas.keys())]}
                letras_lote_selecionadas = selecionar_letras_divisao(resumo_lote_letras)

            elif op_div_lote == "3":
                tipo_divisao_lote = "por_modulo"
                opcoes_mod_lote = [modulos_agregados[k] for k in sorted(modulos_agregados.keys())]
                if opcoes_mod_lote:
                    tab_ml = Table(title="Módulos Detectados nos PDFs Selecionados", show_header=True, header_style="bold magenta")
                    tab_ml.add_column("Opção", style="cyan", width=10, justify="center")
                    tab_ml.add_column("Módulo", style="green", width=22)
                    tab_ml.add_column("Ocorrência nos PDFs", style="dim")
                    for idx_m, m_item in enumerate(opcoes_mod_lote, 1):
                        tab_ml.add_row(str(idx_m), m_item["nome"], m_item["info"])
                    tab_ml.add_row("0 / todos", "[bold yellow]TODOS ACIMA[/bold yellow]", f"Seleciona todos os {len(opcoes_mod_lote)} módulos")
                    console.print()
                    console.print(tab_ml)

                while True:
                    ent_ml = Prompt.ask(
                        "[bold green]Digite os números ou nomes dos módulos separados por vírgula[/bold green] (ex: [cyan]1,2[/cyan], [cyan]0[/cyan] para todos ou [cyan]HCML, HCMR[/cyan])",
                        default="0" if opcoes_mod_lote else ""
                    )
                    modulos_lote_selecionados = parsear_selecao_modulos(ent_ml, opcoes_mod_lote)
                    if modulos_lote_selecionados:
                        break
                    console.print("[red]Informe pelo menos um módulo válido.[/red]")
                console.print(f"[green]✓ Módulos selecionados para divisão ({len(modulos_lote_selecionados)}):[/green] [bold]{' | '.join(modulos_lote_selecionados)}[/bold]")

    # Configura colunas, larguras e tema visual uma única vez para todo o lote
    tema_visual = "simplo_azul"
    if operacao in ["1", "4"]:
        if padrao_predefinido:
            padrao_lote = padrao_predefinido
        else:
            titulo_p = "Qual o padrão do Excel a ser gerado? (Planilha Padrão da Exportação Dupla)" if operacao == "4" else "Qual o padrão do Excel a ser gerado?"
            padrao_lote = selecionar_padrao_planilha(config, titulo=titulo_p)
        colunas = padrao_lote["colunas"]
        larguras_colunas = padrao_lote["larguras"]
        tema_visual = padrao_lote.get("tema_visual", "simplo_azul")
    else:
        colunas = ["Código", "Localização da Avaria", "Causa", "Descrição"]
        larguras_colunas = None  # Mantém o padrão fixo oficial do Buscar Código sem perguntar modificações
        console.print("\n[bold cyan]📋 Estrutura da Planilha 'Excel do Buscar Código' (Padrão Fixo Oficial)[/bold cyan]")
        console.print(f"[green]✓ Colunas fixas:[/green] [bold]{' | '.join(colunas)}[/bold]\n")

    # Seleção de páginas para os PDFs
    console.print("\n[bold cyan]📄 Seleção de Páginas dos PDFs[/bold cyan]")
    console.print("[1] Ler [bold green]todas as páginas[/bold green] de cada PDF selecionado [dim][Padrão][/dim]")
    console.print("[2] Escolher o intervalo de páginas para cada PDF individualmente")
    op_pags = Prompt.ask("Escolha", choices=["1", "2"], default="1")

    paginas_por_pdf = {}
    for arq_pdf in caminhos_pdf:
        tot_p = leitor.obter_total_paginas(arq_pdf)
        if op_pags == "1":
            paginas_por_pdf[arq_pdf] = list(range(1, tot_p + 1))
        else:
            console.print(f"\n[bold yellow]📄 Arquivo: {os.path.basename(arq_pdf)}[/bold yellow]")
            paginas_por_pdf[arq_pdf] = selecionar_paginas(tot_p)

    forcar_ocr = exibir_diagnostico_leitura_automatica(leitor, paginas_por_pdf)

    col_modulo_nome = next((c for c in colunas if any(k in c.lower() for k in ["módulo", "modulo", "ecu"])), None)
    dados_por_pdf = {}
    dados_bc_por_pdf = {}
    textos_por_pdf = {}
    caminho_por_nome_base = {}

    for idx_p, arq_pdf in enumerate(caminhos_pdf, 1):
        nome_base = os.path.splitext(os.path.basename(arq_pdf))[0].replace("_traduzido", "").strip()
        caminho_por_nome_base[nome_base] = arq_pdf
        pags_sel = paginas_por_pdf.get(arq_pdf, [])
        console.print(f"\n[bold magenta]📄 [{idx_p}/{len(caminhos_pdf)}] Processando '{os.path.basename(arq_pdf)}' ({len(pags_sel)} página(s))...[/bold magenta]")

        registros_pdf = []
        registros_bc_pdf = []
        textos_pags_pdf = []
        for i_pag, num_pag in enumerate(pags_sel, 1):
            console.print(f"   [cyan][{i_pag}/{len(pags_sel)}][/cyan] Lendo página [bold]{num_pag}[/bold]...")
            paginas_info = leitor.ler_documento(arq_pdf, paginas=[num_pag], forcar_ocr=forcar_ocr)
            if not paginas_info:
                continue
            info = paginas_info[0]
            texto = info["texto"]
            codigos_gt = info.get("codigos_ground_truth", [])
            if len(texto.strip()) < 10:
                continue
            textos_pags_pdf.append(texto)

            try:
                if operacao == "1" and tema_visual != "simplo_3_colunas_desc":
                    console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Extraindo e traduzindo (Planilha Padrão)...[/blue]")
                    dados_pag = tradutor.estruturar_e_traduzir(
                        texto_bruto=texto,
                        colunas_desejadas=colunas,
                        contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                        modulo_alvo=None,
                        texto_nativo=info.get("texto_nativo"),
                        codigos_ground_truth=codigos_gt
                    )
                    if col_modulo_nome:
                        for r in dados_pag:
                            if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                r[col_modulo_nome] = nome_base
                    registros_pdf.extend(dados_pag)
                elif operacao == "4" and tema_visual != "simplo_3_colunas_desc":
                    console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Extraindo dados para Planilha Padrão e Buscar Código...[/blue]")
                    dados_pag = tradutor.estruturar_e_traduzir(
                        texto_bruto=texto,
                        colunas_desejadas=colunas,
                        contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                        modulo_alvo=None,
                        texto_nativo=info.get("texto_nativo"),
                        codigos_ground_truth=codigos_gt
                    )
                    dados_bc_pag = tradutor.estruturar_buscar_codigo(
                        texto_bruto=texto,
                        contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                        modulo_alvo=None,
                        texto_nativo=info.get("texto_nativo"),
                        codigos_ground_truth=codigos_gt
                    )
                    if col_modulo_nome:
                        for r in dados_pag:
                            if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                r[col_modulo_nome] = nome_base
                    registros_pdf.extend(dados_pag)
                    registros_bc_pdf.extend(dados_bc_pag)
                else:
                    rot_ext = "SIMPLO 3 Colunas + Descrição" if tema_visual == "simplo_3_colunas_desc" else "Buscar Código"
                    console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Extraindo e unificando Códigos + FMI ({rot_ext})...[/blue]")
                    dados_pag = tradutor.estruturar_buscar_codigo(
                        texto_bruto=texto,
                        contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                        modulo_alvo=None,
                        texto_nativo=info.get("texto_nativo"),
                        codigos_ground_truth=codigos_gt
                    )
                    registros_pdf.extend(dados_pag)
                    if operacao == "4":
                        registros_bc_pdf.extend(dados_pag)
                console.print(f"   ↳ [green]✓ {len(dados_pag)} código(s) extraído(s) na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
            except Exception as e:
                console.print(f"   ↳ [red]Erro na página {num_pag} de '{os.path.basename(arq_pdf)}': {e}[/red]")

        if registros_pdf:
            if operacao in ["1", "4"] and tema_visual != "simplo_3_colunas_desc":
                registros_pdf = agente_organizador.organizar(registros_pdf, colunas=colunas)
            else:
                registros_pdf = AgenteOrganizador.organizar_buscar_codigo(registros_pdf)
            dados_por_pdf[nome_base] = registros_pdf
            if operacao == "4":
                dados_bc_por_pdf[nome_base] = AgenteOrganizador.organizar_buscar_codigo(registros_bc_pdf if registros_bc_pdf else registros_pdf)
            textos_por_pdf[nome_base] = "\n".join(textos_pags_pdf)

    if not dados_por_pdf:
        console.print("\n[bold red]Nenhum dado estruturado foi extraído dos PDFs selecionados.[/bold red]")
        return

    sufixo_arq = "_traduzido.xlsx" if operacao in ["1", "4"] else "_buscar_codigo.xlsx"

    def _agrupar_por_modulos_lote(registros_lista: list, texto_completo: str) -> dict:
        """Agrupa uma lista de registros pelos módulos selecionados no lote."""
        grupos_m = {}
        algum_citado = False
        for mod_nome in modulos_lote_selecionados:
            filtrados = filtrar_registros_por_modulo_texto(registros_lista, mod_nome, texto_completo, colunas=colunas)
            if filtrados:
                algum_citado = True
                if operacao in ["1", "4"]:
                    if tema_visual == "simplo_3_colunas_desc":
                        grupos_m[mod_nome] = AgenteOrganizador.organizar_buscar_codigo(filtrados)
                    else:
                        grupos_m[mod_nome] = agente_organizador.organizar(filtrados, colunas=colunas)
                else:
                    grupos_m[mod_nome] = AgenteOrganizador.organizar_buscar_codigo(filtrados)
        # Se nenhum módulo aparece no texto (ex: tabela compartilhada HCML/HCMR), replica para cada módulo escolhido
        if not algum_citado:
            for mod_nome in modulos_lote_selecionados:
                copia_lista = []
                for r in registros_lista:
                    cr = dict(r)
                    if col_modulo_nome:
                        cr[col_modulo_nome] = mod_nome
                    copia_lista.append(cr)
                grupos_m[mod_nome] = copia_lista
        return grupos_m

    try:
        # Modo 1: 1 arquivo Excel separado para cada PDF
        if modo_lote == "1":
            caminhos_gerados = []
            for nome_p, reg_p in dados_por_pdf.items():
                nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome_p).strip().replace(" ", "_") or "documento"
                orig_pdf = caminho_por_nome_base.get(nome_p, "")
                pasta_saida_p = obter_pasta_saida_para_arquivo(orig_pdf) if orig_pdf else PASTA_SAIDA
                caminho_excel = os.path.join(pasta_saida_p, f"{nome_limpo}{sufixo_arq}")

                if tipo_divisao_lote == "por_letra":
                    grupos_l = dividir_registros_por_letra(
                        reg_p,
                        colunas=colunas,
                        letras_selecionadas=letras_lote_selecionadas,
                        modulo_banner=nome_p
                    )
                    for lk in list(grupos_l.keys()):
                        if operacao in ["1", "4"]:
                            if tema_visual == "simplo_3_colunas_desc":
                                grupos_l[lk] = AgenteOrganizador.organizar_buscar_codigo(grupos_l[lk])
                            else:
                                grupos_l[lk] = agente_organizador.organizar(grupos_l[lk], colunas=colunas)
                        else:
                            grupos_l[lk] = AgenteOrganizador.organizar_buscar_codigo(grupos_l[lk])
                    if operacao in ["1", "4"]:
                        c_final = gerador_excel.exportar_multiplos_modulos_padrao(
                            dados_por_modulo=grupos_l,
                            colunas_ordenadas=colunas,
                            caminho_saida=caminho_excel,
                            larguras_personalizadas=larguras_colunas,
                            titulo_banner_padrao=nome_p,
                            tema_visual=tema_visual
                        )
                    else:
                        c_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                            dados_por_modulo=grupos_l,
                            caminho_saida=caminho_excel,
                            larguras_personalizadas=larguras_colunas
                        )
                elif tipo_divisao_lote == "por_modulo":
                    grupos_m = _agrupar_por_modulos_lote(reg_p, textos_por_pdf.get(nome_p, ""))
                    if not grupos_m:
                        grupos_m = {nome_p: reg_p}
                    if operacao in ["1", "4"]:
                        c_final = gerador_excel.exportar_multiplos_modulos_padrao(
                            dados_por_modulo=grupos_m,
                            colunas_ordenadas=colunas,
                            caminho_saida=caminho_excel,
                            larguras_personalizadas=larguras_colunas,
                            tema_visual=tema_visual
                        )
                    else:
                        c_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                            dados_por_modulo=grupos_m,
                            caminho_saida=caminho_excel,
                            larguras_personalizadas=larguras_colunas
                        )
                else:
                    if operacao in ["1", "4"]:
                        c_final = gerador_excel.exportar_para_excel(
                            reg_p,
                            colunas,
                            caminho_excel,
                            nome_modulo=nome_p,
                            larguras_personalizadas=larguras_colunas,
                            tema_visual=tema_visual
                        )
                    else:
                        c_final = gerador_excel.exportar_buscar_codigo(
                            reg_p,
                            caminho_excel,
                            nome_modulo=nome_p,
                            larguras_personalizadas=larguras_colunas
                        )
                caminhos_gerados.append((nome_p, len(reg_p), c_final))

                # Na Operação 4 (Exportação Dupla), gera também a planilha Buscar Código
                if operacao == "4":
                    reg_bc = dados_bc_por_pdf.get(nome_p, reg_p)
                    caminho_bc = os.path.join(pasta_saida_p, f"{nome_limpo}_buscar_codigo.xlsx")
                    c_bc_final = gerador_excel.exportar_buscar_codigo(
                        reg_bc,
                        caminho_bc,
                        nome_modulo=nome_p
                    )
                    caminhos_gerados.append((f"{nome_p} (Buscar Código)", len(reg_bc), c_bc_final))

            resumo_arqs = "\n".join([f"  • [bold]{m}[/bold] ({qtd} registros): [yellow]{c}[/yellow]" for m, qtd, c in caminhos_gerados])
            console.print(Panel.fit(
                f"[bold green]✨ {len(caminhos_gerados)} Planilhas Excel geradas com sucesso![/bold green]\n\n"
                f"{resumo_arqs}",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir os arquivos Excel gerados agora?", default=True):
                for _, _, c_final in caminhos_gerados:
                    os.startfile(c_final)

        # Modo 2: 1 único arquivo Excel com 1 aba para cada PDF
        elif modo_lote == "2":
            pasta_saida_lote = obter_pasta_saida_para_lote(caminhos_pdf)
            nomes_juntos = "_".join([re.sub(r'[\\/*?:"<>|]', "", m).strip().replace(" ", "_") for m in dados_por_pdf.keys()])
            if len(nomes_juntos) > 60 or not nomes_juntos:
                nomes_juntos = f"Lote_{len(dados_por_pdf)}_PDFs"
            caminho_excel = os.path.join(pasta_saida_lote, f"{nomes_juntos}{sufixo_arq}")

            if operacao in ["1", "4"]:
                caminho_final = gerador_excel.exportar_multiplos_modulos_padrao(
                    dados_por_modulo=dados_por_pdf,
                    colunas_ordenadas=colunas,
                    caminho_saida=caminho_excel,
                    larguras_personalizadas=larguras_colunas,
                    tema_visual=tema_visual
                )
            else:
                caminho_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                    dados_por_modulo=dados_por_pdf,
                    caminho_saida=caminho_excel,
                    larguras_personalizadas=larguras_colunas
                )

            # Se for Operação 4, gera também o Buscar Código com abas
            c_bc_final = None
            if operacao == "4":
                caminho_bc = os.path.join(pasta_saida_lote, f"{nomes_juntos}_buscar_codigo.xlsx")
                c_bc_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                    dados_por_modulo=dados_bc_por_pdf,
                    caminho_saida=caminho_bc,
                    larguras_personalizadas=None
                )
                console.print(f"[bold green]✨ Planilha Buscar Código em Lote salva em:[/bold green] [yellow]{c_bc_final}[/yellow]")

            resumo_abas = "\n".join([f"  • Aba [bold]{m}[/bold]: {len(d)} registros" for m, d in dados_por_pdf.items()])
            total_geral = sum(len(d) for d in dados_por_pdf.values())
            texto_arqs = f"Planilha Padrão: [bold yellow]{caminho_final}[/bold yellow]"
            if c_bc_final:
                texto_arqs += f"\nPlanilha Buscar Código: [bold yellow]{c_bc_final}[/bold yellow]"
            console.print(Panel.fit(
                f"[bold green]✨ Planilha Unificada ({len(dados_por_pdf)} Abas) gerada com sucesso![/bold green]\n\n"
                f"Abas geradas:\n{resumo_abas}\n\n"
                f"Total geral de registros: [bold]{total_geral}[/bold]\n\n"
                f"{texto_arqs}",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir o(s) arquivo(s) Excel agora?", default=True):
                os.startfile(caminho_final)
                if c_bc_final:
                    os.startfile(c_bc_final)

        # Modo 3: 1 único arquivo Excel unificando todos os PDFs (aba única, dividido por letra ou dividido por módulo)
        else:
            pasta_saida_lote = obter_pasta_saida_para_lote(caminhos_pdf)
            todos_dados = []
            for d in dados_por_pdf.values():
                todos_dados.extend(d)
            texto_lote_completo = "\n".join(textos_por_pdf.values())
            titulo_lote = " E ".join(dados_por_pdf.keys())
            nome_limpo = re.sub(r'[\\/*?:"<>|]', "", titulo_lote).strip().replace(" ", "_")
            if not nome_limpo or len(nome_limpo) > 60:
                nome_limpo = f"Unificado_{len(dados_por_pdf)}_PDFs"
            caminho_excel = os.path.join(pasta_saida_lote, f"{nome_limpo}{sufixo_arq}")

            if tipo_divisao_lote == "por_letra":
                grupos_l = dividir_registros_por_letra(
                    todos_dados,
                    colunas=colunas,
                    letras_selecionadas=letras_lote_selecionadas,
                    modulo_banner=titulo_lote
                )
                for lk in list(grupos_l.keys()):
                    if operacao in ["1", "4"]:
                        if tema_visual == "simplo_3_colunas_desc":
                            grupos_l[lk] = AgenteOrganizador.organizar_buscar_codigo(grupos_l[lk])
                        else:
                            grupos_l[lk] = agente_organizador.organizar(grupos_l[lk], colunas=colunas)
                    else:
                        grupos_l[lk] = AgenteOrganizador.organizar_buscar_codigo(grupos_l[lk])
                if operacao in ["1", "4"]:
                    caminho_final = gerador_excel.exportar_multiplos_modulos_padrao(
                        dados_por_modulo=grupos_l,
                        colunas_ordenadas=colunas,
                        caminho_saida=caminho_excel,
                        larguras_personalizadas=larguras_colunas,
                        titulo_banner_padrao=titulo_lote,
                        tema_visual=tema_visual
                    )
                else:
                    caminho_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                        dados_por_modulo=grupos_l,
                        caminho_saida=caminho_excel,
                        larguras_personalizadas=larguras_colunas
                    )
            elif tipo_divisao_lote == "por_modulo":
                grupos_m = _agrupar_por_modulos_lote(todos_dados, texto_lote_completo)
                if not grupos_m:
                    grupos_m = {titulo_lote: todos_dados}
                if operacao in ["1", "4"]:
                    caminho_final = gerador_excel.exportar_multiplos_modulos_padrao(
                        dados_por_modulo=grupos_m,
                        colunas_ordenadas=colunas,
                        caminho_saida=caminho_excel,
                        larguras_personalizadas=larguras_colunas,
                        tema_visual=tema_visual
                    )
                else:
                    caminho_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                        dados_por_modulo=grupos_m,
                        caminho_saida=caminho_excel,
                        larguras_personalizadas=larguras_colunas
                    )
            else:
                if operacao in ["1", "4"]:
                    if tema_visual == "simplo_3_colunas_desc":
                        todos_dados = AgenteOrganizador.organizar_buscar_codigo(todos_dados)
                    else:
                        todos_dados = agente_organizador.organizar(todos_dados, colunas=colunas)
                    caminho_final = gerador_excel.exportar_para_excel(
                        todos_dados,
                        colunas,
                        caminho_excel,
                        nome_modulo=titulo_lote,
                        larguras_personalizadas=larguras_colunas,
                        tema_visual=tema_visual
                    )
                else:
                    todos_dados = AgenteOrganizador.organizar_buscar_codigo(todos_dados)
                    caminho_final = gerador_excel.exportar_buscar_codigo(
                        todos_dados,
                        caminho_excel,
                        nome_modulo=titulo_lote,
                        larguras_personalizadas=larguras_colunas
                    )

            c_bc_final = None
            if operacao == "4":
                todos_bc = []
                for d in dados_bc_por_pdf.values():
                    todos_bc.extend(d)
                caminho_bc = os.path.join(pasta_saida_lote, f"{nome_limpo}_buscar_codigo.xlsx")
                if tipo_divisao_lote == "por_letra":
                    grupos_bc_l = dividir_registros_por_letra(
                        todos_bc,
                        colunas=["Código", "Localização da Avaria", "Causa", "Descrição"],
                        letras_selecionadas=letras_lote_selecionadas,
                        modulo_banner=titulo_lote
                    )
                    for lk in list(grupos_bc_l.keys()):
                        grupos_bc_l[lk] = AgenteOrganizador.organizar_buscar_codigo(grupos_bc_l[lk])
                    c_bc_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                        dados_por_modulo=grupos_bc_l,
                        caminho_saida=caminho_bc,
                        larguras_personalizadas=None
                    )
                elif tipo_divisao_lote == "por_modulo":
                    grupos_bc_m = _agrupar_por_modulos_lote(todos_bc, texto_lote_completo)
                    c_bc_final = gerador_excel.exportar_multiplos_modulos_buscar_codigo(
                        dados_por_modulo=grupos_bc_m,
                        caminho_saida=caminho_bc,
                        larguras_personalizadas=None
                    )
                else:
                    c_bc_final = gerador_excel.exportar_buscar_codigo(
                        AgenteOrganizador.organizar_buscar_codigo(todos_bc),
                        caminho_bc,
                        nome_modulo=titulo_lote
                    )
                console.print(f"[bold green]✨ Planilha Buscar Código Unificada salva em:[/bold green] [yellow]{c_bc_final}[/yellow]")

            texto_arqs = f"Planilha Padrão: [bold yellow]{caminho_final}[/bold yellow]"
            if c_bc_final:
                texto_arqs += f"\nPlanilha Buscar Código: [bold yellow]{c_bc_final}[/bold yellow]"

            console.print(Panel.fit(
                f"[bold green]✨ Planilha Unificada ({len(dados_por_pdf)} PDFs) gerada com sucesso![/bold green]\n\n"
                f"Documentos: [bold]{titulo_lote}[/bold]\n"
                f"Total de registros: [bold]{len(todos_dados)}[/bold]\n\n"
                f"{texto_arqs}",
                border_style="green"
            ))
            if Confirm.ask("Deseja abrir o(s) arquivo(s) Excel agora?", default=True):
                os.startfile(caminho_final)
                if c_bc_final:
                    os.startfile(c_bc_final)

    except Exception as e:
        console.print(f"[bold red]Erro ao gerar Excel do lote: {e}[/bold red]")


def main():
    console.print(Panel.fit(
        "[bold cyan]Sistema de Extração, Tradução Técnica com IA e Geração de Excel[/bold cyan]\n"
        "[dim]Segmento: Linha Leve, Linha Pesada e Maquinário Agrícola (DTC, J1939, SPN-FMI)[/dim]",
        border_style="cyan"
    ))

    # Verifica chave da API
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key or api_key in ["sua_chave_aqui", "python app.py"] or len(api_key) < 15:
        console.print("\n[bold red]⚠️ Atenção: Chave GEMINI_API_KEY não configurada ou inválida no arquivo .env![/bold red]")
        console.print("[yellow]Obtenha sua chave gratuita em: https://aistudio.google.com/app/apikey[/yellow]")
        chave_manual = Prompt.ask("Cole sua chave Gemini API (começa com AIza...) ou aperte Enter para sair")
        chave_manual = chave_manual.strip(' "\'')
        if chave_manual and len(chave_manual) > 15:
            api_key = chave_manual
            with open(os.path.join(PASTA_BASE, ".env"), "w", encoding="utf-8") as f:
                f.write(f"GEMINI_API_KEY={api_key}\nTESSERACT_PATH=C:\\Program Files\\Tesseract-OCR\\tesseract.exe\n")
            os.environ["GEMINI_API_KEY"] = api_key
        else:
            console.print("[red]Chave não informada. O programa não poderá fazer a tradução com IA.[/red]")
            sys.exit(0)

    config = carregar_config()

    # Inicializa motores uma única vez na sessão
    ocr_engine = OCREngine(tesseract_cmd=config.get("tesseract_path"))
    leitor = LeitorPDF(ocr_engine=ocr_engine)
    tradutor = TradutorIA(api_key=api_key, model_name=config.get("gemini_model", "gemini-3.8-flash"))
    agente_organizador = AgenteOrganizador()
    gerador_excel = GeradorExcel(agente_organizador=agente_organizador)

    while True:
        # 1. Seleção de 1 ou múltiplos arquivos (PDF, Imagem, Print Screen, Excel, Word, HTML)
        caminhos_pdf = selecionar_arquivos_pdf()
        if not caminhos_pdf:
            break

        # 2. Escolha da Operação: Extração Padrão, Glossário, Buscar Código ou Exportação Dupla
        console.print("\n[bold cyan]🎯 Qual operação deseja realizar com este(s) documento(s)?[/bold cyan]")
        console.print("[1] 📊 Extrair códigos, traduzir e gerar Planilha Padrão (Completa / Padrões SIMPLO)")
        console.print("[2] 📖 Mapear novos termos técnicos deste(s) documento(s) para atualizar o Glossário")
        console.print("[3] 🔍 Gerar Excel do Buscar Código (Padrão Fixo: Código+FMI, Localização, Causa, Descrição)")
        console.print("[4] ⚡ Exportação Dupla (Gerar Planilha Padrão + Excel do Buscar Código de uma só vez)")
        opcao_modo = Prompt.ask("Escolha a opção", choices=["1", "2", "3", "4"], default="1")

        # Na Operação 4 (Exportação Dupla), pergunta IMEDIATAMENTE qual o padrão de planilha a ser gerado
        padrao_op4 = None
        if opcao_modo == "4":
            padrao_op4 = selecionar_padrao_planilha(
                config,
                titulo="Qual o padrão do Excel a ser gerado? (Planilha Padrão da Exportação Dupla)"
            )

        if len(caminhos_pdf) == 1:
            caminho_pdf = caminhos_pdf[0]
            total_paginas = leitor.obter_total_paginas(caminho_pdf)
            if opcao_modo == "1":
                executar_extracao_padrao(caminho_pdf, leitor, tradutor, agente_organizador, gerador_excel, total_paginas, config, exportar_duplo=False)
            elif opcao_modo == "2":
                continuar = executar_mapeamento_termos(caminho_pdf, leitor, tradutor, total_paginas)
                if continuar:
                    executar_extracao_padrao(caminho_pdf, leitor, tradutor, agente_organizador, gerador_excel, total_paginas, config, exportar_duplo=False)
                else:
                    console.print("[green]Operação finalizada.[/green]")
            elif opcao_modo == "3":
                executar_excel_buscar_codigo(caminho_pdf, leitor, tradutor, agente_organizador, gerador_excel, total_paginas, config)
            elif opcao_modo == "4":
                executar_extracao_padrao(
                    caminho_pdf,
                    leitor,
                    tradutor,
                    agente_organizador,
                    gerador_excel,
                    total_paginas,
                    config,
                    exportar_duplo=True,
                    padrao_predefinido=padrao_op4
                )
        else:
            if opcao_modo == "1":
                executar_lote_multiplos_pdfs(caminhos_pdf, "1", leitor, tradutor, agente_organizador, gerador_excel, config)
            elif opcao_modo == "2":
                continuar = executar_mapeamento_termos(caminhos_pdf, leitor, tradutor)
                if continuar:
                    executar_lote_multiplos_pdfs(caminhos_pdf, "1", leitor, tradutor, agente_organizador, gerador_excel, config)
                else:
                    console.print("[green]Operação finalizada.[/green]")
            elif opcao_modo == "3":
                executar_lote_multiplos_pdfs(caminhos_pdf, "3", leitor, tradutor, agente_organizador, gerador_excel, config)
            elif opcao_modo == "4":
                executar_lote_multiplos_pdfs(
                    caminhos_pdf,
                    "4",
                    leitor,
                    tradutor,
                    agente_organizador,
                    gerador_excel,
                    config,
                    padrao_predefinido=padrao_op4
                )

        console.print("")
        if not Confirm.ask("🔄 Deseja processar outro arquivo ou realizar uma nova operação?", default=False):
            console.print("[bold cyan]👋 Sessão encerrada. Até logo![/bold cyan]")
            break


def executar_extracao_padrao(
    caminho_pdf: str,
    leitor: LeitorPDF,
    tradutor: TradutorIA,
    agente_organizador: AgenteOrganizador,
    gerador_excel: GeradorExcel,
    total_paginas: int,
    config: dict,
    exportar_duplo: bool = False,
    padrao_predefinido: dict = None
):
    """
    Operação 1 (Planilha Padrão) e Operação 4 (Exportação Dupla):
    - Seleção de 1 ou vários módulos (em abas separadas, arquivos separados ou aba única)
    - Seleção rápida de Padrão de Planilha (Automático, SIMPLO 2/4 Colunas, Importar Modelo .xlsx)
    - Pipeline Multi-Agente com Cache Inteligente + Processamento Paralelo Multi-Páginas
    """
    import re
    from concurrent.futures import ThreadPoolExecutor, as_completed

    # 3. Escolha do Padrão de Planilha (se ainda não escolhido na operação)
    if padrao_predefinido:
        padrao_escolhido = padrao_predefinido
    else:
        titulo_p = "Qual o padrão do Excel a ser gerado? (Planilha Padrão da Exportação Dupla)" if exportar_duplo else "Qual o padrão do Excel a ser gerado?"
        padrao_escolhido = selecionar_padrao_planilha(config, titulo=titulo_p)

    colunas = padrao_escolhido["colunas"]
    larguras_colunas = padrao_escolhido["larguras"]
    tema_visual = padrao_escolhido.get("tema_visual", "simplo_azul")
    gerar_bc_extra = bool(exportar_duplo)

    # 4. Seleção de um ou vários módulos
    modulos_selecionados, modo_saida, resumo_info = selecionar_modulos(
        caminho_pdf=caminho_pdf,
        leitor=leitor,
        tradutor=tradutor,
        contexto_operacao="Exportação Dupla (Padrão + Buscar Código)" if exportar_duplo else "Planilha Padrão"
    )

    # 5. Escolha das páginas
    paginas_selecionadas = selecionar_paginas(total_paginas)

    # 6. Detecção Automática de Leitura (Vetorial vs OCR 300 DPI)
    forcar_ocr = exibir_diagnostico_leitura_automatica(leitor, {caminho_pdf: paginas_selecionadas})

    console.print(
        f"\n[bold green]🚀 Iniciando processamento de {len(paginas_selecionadas)} página(s) "
        f"para {len(modulos_selecionados)} módulo(s): {', '.join(modulos_selecionados)}...[/bold green]"
    )
    console.print(f"[dim]Colunas a extrair: {', '.join(colunas)}[/dim]\n")

    # Lê e faz cache das páginas selecionadas uma única vez
    paginas_cache = {}
    for i, num_pag in enumerate(paginas_selecionadas, 1):
        console.print(f"[cyan][{i}/{len(paginas_selecionadas)}][/cyan] Lendo página [bold]{num_pag}[/bold]...")
        
        paginas_info = leitor.ler_documento(caminho_pdf, paginas=[num_pag], forcar_ocr=forcar_ocr)
        if not paginas_info:
            continue
            
        info = paginas_info[0]
        texto = info["texto"]
        metodo = info["metodo"]
        codigos_gt = info.get("codigos_ground_truth", [])
        console.print(f"   ↳ Modo: [dim]{metodo}[/dim] ({len(texto)} caracteres)")
        if codigos_gt:
            console.print(f"   ↳ [cyan]Ground Truth:[/cyan] {len(codigos_gt)} código(s) de falha verificados diretamente no documento{formatar_nota_codigos_repetidos(info)}")

        if len(texto.strip()) < 10:
            console.print("   ↳ [yellow]Página sem texto significativo identificado, pulando.[/yellow]")
            continue

        paginas_cache[num_pag] = info

    if not paginas_cache:
        console.print("\n[bold red]Nenhuma página com texto válido foi lida do documento.[/bold red]")
        return

    config_divisao = (resumo_info or {}).get("config_divisao", {})
    tipo_divisao = config_divisao.get("tipo", "por_modulo" if len(modulos_selecionados) > 1 else "nenhuma")
    modulo_banner = config_divisao.get("modulo_banner") or os.path.splitext(os.path.basename(caminho_pdf))[0]
    tabela_geral_sem_modulo = not bool((resumo_info or {}).get("modulos_regex"))
    col_modulo_nome = next((c for c in colunas if any(k in c.lower() for k in ["módulo", "modulo", "ecu"])), None)

    dados_por_modulo = {}
    usar_motor_bc = (tema_visual == "simplo_3_colunas_desc")

    def _extrair_pag_escolhida(txt_b, mod_a, txt_n, c_gt):
        if usar_motor_bc:
            return tradutor.estruturar_buscar_codigo(
                texto_bruto=txt_b,
                contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                modulo_alvo=mod_a,
                texto_nativo=txt_n,
                codigos_ground_truth=c_gt
            )
        return tradutor.estruturar_e_traduzir(
            texto_bruto=txt_b,
            colunas_desejadas=colunas,
            contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
            modulo_alvo=mod_a,
            texto_nativo=txt_n,
            codigos_ground_truth=c_gt
        )

    def _organizar_escolhido(lista_r):
        if usar_motor_bc:
            return AgenteOrganizador.organizar_buscar_codigo(lista_r)
        return agente_organizador.organizar(lista_r, colunas=colunas)

    # MODO A: Divisão por Letra do código (ex: B, C, P, U)
    if tipo_divisao == "por_letra":
        letras_alvo = config_divisao.get("letras") or modulos_selecionados
        letras_alvo_set = {l.upper() for l in letras_alvo}
        todos_extraidos = []

        pags_para_extrair = []
        for num_pag, info in paginas_cache.items():
            codigos_gt = info.get("codigos_ground_truth", [])
            if codigos_gt and letras_alvo_set:
                letras_pag = {obter_letra_codigo(c) for c in codigos_gt}
                if not (letras_pag & letras_alvo_set):
                    console.print(f"   ↳ [dim]Página {num_pag}: sem códigos das letras {', '.join(letras_alvo)}, pulando.[/dim]")
                    continue
            pags_para_extrair.append((num_pag, info))

        if len(pags_para_extrair) > 1:
            console.print(f"   ↳ [blue]⚡ Extraindo {len(pags_para_extrair)} páginas em paralelo (Multi-Thread + Cache Inteligente)...[/blue]")
            resultados_pag = {}
            def _worker_letra(item_p):
                np, inf = item_p
                res_p = _extrair_pag_escolhida(
                    inf["texto"],
                    None,
                    inf.get("texto_nativo"),
                    inf.get("codigos_ground_truth", [])
                )
                return np, inf, res_p

            with ThreadPoolExecutor(max_workers=min(3, len(pags_para_extrair))) as executor:
                futuros = {executor.submit(_worker_letra, it): it[0] for it in pags_para_extrair}
                for fut in as_completed(futuros):
                    np_f = futuros[fut]
                    try:
                        np_r, inf_r, dados_pagina = fut.result()
                        resultados_pag[np_r] = (inf_r, dados_pagina)
                    except Exception as e:
                        console.print(f"   ↳ [red]Erro na página {np_f}: {e}[/red]")

            for num_pag, _ in pags_para_extrair:
                if num_pag in resultados_pag:
                    info, dados_pagina = resultados_pag[num_pag]
                    if col_modulo_nome:
                        for r in dados_pagina:
                            if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                r[col_modulo_nome] = modulo_banner
                    console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                    todos_extraidos.extend(dados_pagina)
        else:
            for num_pag, info in pags_para_extrair:
                console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Executando Pipeline Multi-Agente (Extrator ➔ Tradutor ➔ Auditor)...[/blue]")
                try:
                    dados_pagina = _extrair_pag_escolhida(
                        info["texto"],
                        None,
                        info.get("texto_nativo"),
                        info.get("codigos_ground_truth", [])
                    )
                    if col_modulo_nome:
                        for r in dados_pagina:
                            if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                r[col_modulo_nome] = modulo_banner
                    console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                    todos_extraidos.extend(dados_pagina)
                except Exception as e:
                    console.print(f"   ↳ [red]Erro na página {num_pag}: {e}[/red]")

        dados_por_modulo = dividir_registros_por_letra(
            registros=todos_extraidos,
            colunas=colunas,
            letras_selecionadas=letras_alvo,
            modulo_banner=modulo_banner
        )
        for letra_k in list(dados_por_modulo.keys()):
            dados_por_modulo[letra_k] = _organizar_escolhido(dados_por_modulo[letra_k])
            console.print(f"   ↳ [cyan]🔢 Aba Letra '{letra_k}': {len(dados_por_modulo[letra_k])} registro(s) ordenado(s).[/cyan]")

    # MODO B: Divisão por Módulo ou Módulo Único
    else:
        mapa_paginas = mapear_paginas_por_modulo(modulos_selecionados, paginas_selecionadas, resumo_info)
        dados_por_modulo = {mod: [] for mod in modulos_selecionados}
        cache_padrao_geral = {}

        # Verifica se os módulos selecionados aparecem citados nas descrições de uma tabela geral
        modulos_citados_no_texto = set()
        if tabela_geral_sem_modulo and len(modulos_selecionados) > 1:
            for mod in modulos_selecionados:
                partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', mod.strip()) if x]
                if partes_m:
                    pat_m = re.compile(r'(?<![A-Za-z0-9])' + r'[\s_-]*'.join(partes_m) + r'(?![A-Za-z0-9])', re.I)
                    if any(pat_m.search(info["texto"]) for info in paginas_cache.values()):
                        modulos_citados_no_texto.add(mod)

        for idx_m, nome_modulo in enumerate(modulos_selecionados, 1):
            pags_modulo = [(p, nota) for (p, nota) in mapa_paginas.get(nome_modulo, []) if p in paginas_cache]
            console.print(f"\n[bold magenta]🏷️ [{idx_m}/{len(modulos_selecionados)}] Processando módulo '{nome_modulo}' ({len(pags_modulo)} página(s))...[/bold magenta]")

            # Filtra páginas relevantes e aproveita cache_padrao_geral quando disponível
            pags_para_ia = []
            for num_pag, nota_contexto in pags_modulo:
                info = paginas_cache[num_pag]
                if nome_modulo in modulos_citados_no_texto:
                    partes_m = [re.escape(x) for x in re.split(r'[\s_-]+', nome_modulo.strip()) if x]
                    pat_m = re.compile(r'(?<![A-Za-z0-9])' + r'[\s_-]*'.join(partes_m) + r'(?![A-Za-z0-9])', re.I)
                    if not pat_m.search(info["texto"]):
                        continue

                if tabela_geral_sem_modulo and num_pag in cache_padrao_geral:
                    base_pag = []
                    for r in cache_padrao_geral[num_pag]:
                        copia_r = dict(r)
                        if col_modulo_nome:
                            copia_r[col_modulo_nome] = nome_modulo
                        base_pag.append(copia_r)
                    if modulos_citados_no_texto:
                        dados_pagina = filtrar_registros_por_modulo_texto(base_pag, nome_modulo, info["texto"], colunas=colunas)
                    else:
                        dados_pagina = base_pag
                    console.print(f"   ↳ [green]✓ Página {num_pag}: {len(dados_pagina)} código(s) filtrado(s) para '{nome_modulo}'![/green]")
                    dados_por_modulo[nome_modulo].extend(dados_pagina)
                    continue

                pags_para_ia.append((num_pag, nota_contexto, info))

            if len(pags_para_ia) > 1:
                console.print(f"   ↳ [blue]⚡ Extraindo {len(pags_para_ia)} páginas de '{nome_modulo}' em paralelo...[/blue]")
                resultados_mod = {}
                alvo_ia = None if tabela_geral_sem_modulo else nome_modulo

                def _worker_mod(item_m):
                    np, nota_c, inf = item_m
                    txt_env = f"{nota_c}{inf['texto']}" if nota_c else inf["texto"]
                    res_p = _extrair_pag_escolhida(
                        txt_env,
                        alvo_ia,
                        inf.get("texto_nativo"),
                        inf.get("codigos_ground_truth", [])
                    )
                    return np, inf, res_p

                with ThreadPoolExecutor(max_workers=min(3, len(pags_para_ia))) as executor:
                    futuros = {executor.submit(_worker_mod, it): it[0] for it in pags_para_ia}
                    for fut in as_completed(futuros):
                        np_f = futuros[fut]
                        try:
                            np_r, inf_r, dados_pagina = fut.result()
                            resultados_mod[np_r] = (inf_r, dados_pagina)
                        except Exception as e:
                            console.print(f"   ↳ [red]Erro na página {np_f} (módulo '{nome_modulo}'): {e}[/red]")

                for num_pag, _, _ in pags_para_ia:
                    if num_pag in resultados_mod:
                        info, dados_pagina = resultados_mod[num_pag]
                        if col_modulo_nome:
                            for r in dados_pagina:
                                if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                    r[col_modulo_nome] = nome_modulo
                        if tabela_geral_sem_modulo and dados_pagina:
                            cache_padrao_geral[num_pag] = [dict(r) for r in dados_pagina]
                            if modulos_citados_no_texto:
                                dados_pagina = filtrar_registros_por_modulo_texto(dados_pagina, nome_modulo, info["texto"], colunas=colunas)
                        console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) para '{nome_modulo}' na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                        dados_por_modulo[nome_modulo].extend(dados_pagina)
            else:
                for num_pag, nota_contexto, info in pags_para_ia:
                    texto_envio = f"{nota_contexto}{info['texto']}" if nota_contexto else info["texto"]
                    codigos_gt = info.get("codigos_ground_truth", [])
                    console.print(f"   ↳ [blue]Página {num_pag}: 🤖 Executando Pipeline Multi-Agente (Extrator ➔ Tradutor ➔ Auditor) para '{nome_modulo}'...[/blue]")
                    try:
                        alvo_ia = None if tabela_geral_sem_modulo else nome_modulo
                        dados_pagina = _extrair_pag_escolhida(
                            texto_envio,
                            alvo_ia,
                            info.get("texto_nativo"),
                            codigos_gt
                        )
                        if col_modulo_nome:
                            for r in dados_pagina:
                                if not r.get(col_modulo_nome) or r.get(col_modulo_nome) == "-":
                                    r[col_modulo_nome] = nome_modulo
                        if tabela_geral_sem_modulo and dados_pagina:
                            cache_padrao_geral[num_pag] = [dict(r) for r in dados_pagina]
                            if modulos_citados_no_texto:
                                dados_pagina = filtrar_registros_por_modulo_texto(dados_pagina, nome_modulo, info["texto"], colunas=colunas)
                        console.print(f"   ↳ [green]✓ {len(dados_pagina)} código(s) extraído(s) para '{nome_modulo}' na página {num_pag}![/green]{formatar_nota_codigos_repetidos(info)}")
                        dados_por_modulo[nome_modulo].extend(dados_pagina)
                    except Exception as e:
                        console.print(f"   ↳ [red]Erro na página {num_pag} (módulo '{nome_modulo}'): {e}[/red]")

            if dados_por_modulo[nome_modulo]:
                console.print(f"   ↳ [cyan]🔢 Agente 4 (Organizador): Ordenando {len(dados_por_modulo[nome_modulo])} registros do módulo '{nome_modulo}'...[/cyan]")
                dados_por_modulo[nome_modulo] = _organizar_escolhido(dados_por_modulo[nome_modulo])

    modulos_com_dados = {m: d for m, d in dados_por_modulo.items() if d}
    if not modulos_com_dados:
        console.print("\n[bold red]Nenhum dado estruturado foi gerado. Verifique o conteúdo do PDF ou a chave de API.[/bold red]")
        return

    try:
        caminhos_gerados_padrao = []
        info_modulos_resumo = ""
        pasta_saida_atual = obter_pasta_saida_para_arquivo(caminho_pdf)

        # Caso 1: Múltiplas abas/páginas diferentes no mesmo Excel (por Módulo ou por Letra)
        if modo_saida == "abas_separadas" and (len(modulos_com_dados) > 1 or tipo_divisao in ["por_letra", "por_modulo"]):
            if tipo_divisao == "por_letra":
                base_arq = re.sub(r'[\\/*?:"<>|]', "", modulo_banner).strip().replace(" ", "_") or os.path.splitext(os.path.basename(caminho_pdf))[0]
                caminho_excel = os.path.join(pasta_saida_atual, f"{base_arq}_traduzido.xlsx")
                banner_padrao = modulo_banner
            else:
                nomes_juntos = "_".join([re.sub(r'[\\/*?:"<>|]', "", m).strip().replace(" ", "_") for m in modulos_com_dados.keys()])
                if len(nomes_juntos) > 60 or not nomes_juntos:
                    nomes_juntos = os.path.splitext(os.path.basename(caminho_pdf))[0]
                caminho_excel = os.path.join(pasta_saida_atual, f"{nomes_juntos}_traduzido.xlsx")
                banner_padrao = None

            console.print(f"\n[bold cyan]📊 Gerando arquivo Excel final com {len(modulos_com_dados)} abas/páginas separadas...[/bold cyan]")
            caminho_final = gerador_excel.exportar_multiplos_modulos_padrao(
                dados_por_modulo=modulos_com_dados,
                colunas_ordenadas=colunas,
                caminho_saida=caminho_excel,
                larguras_personalizadas=larguras_colunas,
                titulo_banner_padrao=banner_padrao,
                tema_visual=tema_visual
            )
            caminhos_gerados_padrao.append(caminho_final)
            resumo_abas = "\n".join([f"  • Aba [bold]{m}[/bold]: {len(d)} registros" for m, d in modulos_com_dados.items()])
            total_geral = sum(len(d) for d in modulos_com_dados.values())
            info_modulos_resumo = f"Abas geradas ({len(modulos_com_dados)}):\n{resumo_abas}\nTotal geral de registros: [bold]{total_geral}[/bold]"

        # Caso 2: Múltiplos arquivos Excel separados (por Módulo ou por Letra)
        elif modo_saida == "arquivos_separados" and (len(modulos_com_dados) > 1 or tipo_divisao in ["por_letra", "por_modulo"]):
            for nome_mod, dados_mod in modulos_com_dados.items():
                nome_limpo = re.sub(r'[\\/*?:"<>|]', "", nome_mod).strip().replace(" ", "_") or "modulo"
                if tipo_divisao == "por_letra":
                    pref = re.sub(r'[\\/*?:"<>|]', "", modulo_banner).strip().replace(" ", "_")
                    caminho_excel = os.path.join(pasta_saida_atual, f"{pref}_{nome_limpo}_traduzido.xlsx")
                    titulo_no_banner = f"{modulo_banner} ({nome_mod})"
                else:
                    caminho_excel = os.path.join(pasta_saida_atual, f"{nome_limpo}_traduzido.xlsx")
                    titulo_no_banner = nome_mod
                console.print(f"\n[bold cyan]📊 Gerando arquivo Excel separado para '{nome_mod}'...[/bold cyan]")
                c_final = gerador_excel.exportar_para_excel(
                    dados_mod,
                    colunas,
                    caminho_excel,
                    nome_modulo=titulo_no_banner,
                    larguras_personalizadas=larguras_colunas,
                    tema_visual=tema_visual
                )
                caminhos_gerados_padrao.append(c_final)
            info_modulos_resumo = f"Arquivos individuais gerados para {len(modulos_com_dados)} módulo(s)."

        # Caso 3: Módulo único ou todos juntos na mesma página/aba
        else:
            todos_dados = []
            for d in modulos_com_dados.values():
                todos_dados.extend(d)
            if usar_motor_bc:
                todos_dados = AgenteOrganizador.organizar_buscar_codigo(todos_dados)
            else:
                todos_dados = agente_organizador.organizar(todos_dados, colunas=colunas)
            titulo_modulo = " E ".join(modulos_com_dados.keys())
            nome_limpo_modulo = re.sub(r'[\\/*?:"<>|]', "", titulo_modulo).strip().replace(" ", "_")
            if not nome_limpo_modulo or len(nome_limpo_modulo) > 60:
                nome_limpo_modulo = os.path.splitext(os.path.basename(caminho_pdf))[0]

            caminho_excel = os.path.join(pasta_saida_atual, f"{nome_limpo_modulo}_traduzido.xlsx")
            console.print(f"\n[bold cyan]📊 Gerando arquivo Excel final para '{titulo_modulo}'...[/bold cyan]")
            caminho_final = gerador_excel.exportar_para_excel(
                todos_dados,
                colunas,
                caminho_excel,
                nome_modulo=titulo_modulo,
                larguras_personalizadas=larguras_colunas,
                tema_visual=tema_visual
            )
            caminhos_gerados_padrao.append(caminho_final)
            info_modulos_resumo = f"Módulo(s): [bold]{titulo_modulo}[/bold]\nTotal de registros: [bold]{len(todos_dados)}[/bold]"

        caminhos_gerados_bc = []
        # Se o usuário escolheu [4] Exportação Dupla (Padrão + Buscar Código), gera também a planilha Buscar Código
        if gerar_bc_extra:
            console.print("\n[bold cyan]🔍 Gerando também a 2ª planilha automática: 'Excel do Buscar Código'...[/bold cyan]")
            todos_bc = []
            if tema_visual == "simplo_3_colunas_desc":
                for d in modulos_com_dados.values():
                    todos_bc.extend(d)
            else:
                for num_pag, info in paginas_cache.items():
                    try:
                        res_bc = tradutor.estruturar_buscar_codigo(
                            texto_bruto=info["texto"],
                            contexto_segmento="Veículos Pesados e Máquinas Agrícolas",
                            modulo_alvo=None,
                            texto_nativo=info.get("texto_nativo"),
                            codigos_ground_truth=info.get("codigos_ground_truth", [])
                        )
                        todos_bc.extend(res_bc)
                    except Exception as e_bc:
                        console.print(f"   ↳ [red]Erro ao gerar Buscar Código na página {num_pag}: {e_bc}[/red]")
            if todos_bc:
                todos_bc = AgenteOrganizador.organizar_buscar_codigo(todos_bc)
                base_bc = os.path.splitext(os.path.basename(caminho_pdf))[0].replace("_traduzido", "")
                caminho_bc = os.path.join(pasta_saida_atual, f"{base_bc}_buscar_codigo.xlsx")
                c_bc_final = gerador_excel.exportar_buscar_codigo(
                    todos_bc,
                    caminho_bc,
                    nome_modulo=modulo_banner
                )
                caminhos_gerados_bc.append(c_bc_final)
                console.print(f"[bold green]✨ Planilha Buscar Código complementar salva em:[/bold green] [yellow]{c_bc_final}[/yellow]")

        todos_arquivos = caminhos_gerados_padrao + caminhos_gerados_bc
        linhas_resumo = []
        for c_pad in caminhos_gerados_padrao:
            linhas_resumo.append(f"  • Planilha Padrão ({tema_visual}): [bold yellow]{c_pad}[/bold yellow]")
        for c_bc in caminhos_gerados_bc:
            linhas_resumo.append(f"  • Planilha Buscar Código: [bold yellow]{c_bc}[/bold yellow]")

        texto_painel = f"[bold green]✨ Processamento concluído com sucesso![/bold green]\n\n"
        if info_modulos_resumo:
            texto_painel += f"{info_modulos_resumo}\n\n"
        texto_painel += "\n".join(linhas_resumo)

        console.print(Panel.fit(texto_painel, border_style="green"))

        if Confirm.ask("Deseja abrir o(s) arquivo(s) Excel agora?", default=True):
            for arq in todos_arquivos:
                os.startfile(arq)

    except Exception as e:
        console.print(f"[bold red]Erro ao gerar Excel: {e}[/bold red]")

if __name__ == "__main__":
    main()
