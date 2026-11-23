# Sistema de Extração, Tradução Técnica de Falhas e Exportação para Excel

Sistema desenvolvido em Python para extrair códigos de falhas técnicas (DTC, OBD-II, SAE J1939 SPN/FMI, manuais automotivos, veículos pesados/diesel e maquinário agrícola), traduzir para Português do Brasil (PT-BR) com Inteligência Artificial especializada e gerar planilhas profissionais no Microsoft Excel (`.xlsx`).

---

## 🚀 Como Funciona

1. **Entrada de PDF**: Coloque os manuais ou tabelas em PDF na pasta `entrada/` (ou arraste o arquivo diretamente no terminal).
2. **Seleção Prévia de Colunas**: Antes de processar, o sistema exibe os modelos de colunas (Linha Pesada & Agrícola, Linha Leve, Resumido) ou permite que você digite exatamente quais colunas deseja na planilha final.
3. **Leitura Híbrida Inteligente**:
   - Se o PDF já for digitalizado com texto selecionável, extrai diretamente com 100% de precisão de caracteres.
   - Se o PDF for escaneado (imagens/fotos de manuais), renderiza a página a **300 DPI** e processa com o **Tesseract OCR** com filtros de contraste e nitidez.
4. **Tradução Técnica com IA (Gemini)**:
   - Preserva siglas de módulos e sistemas (ECU, TCU, ABS, EBS, ESP, EGR, DPF, SCR, CAN, PWM, etc.).
   - Aplica terminologia automotiva correta (ex: *"Open circuit"* -> *"Circuito aberto"*, *"Short to ground"* -> *"Curto-circuito ao terra / massa"*).
5. **Excel Estilizado**:
   - Cabeçalho escuro corporativo com texto em negrito.
   - Linhas zebradas para facilitar a leitura.
   - Alinhamentos específicos (códigos centralizados, descrições alinhadas à esquerda com quebra de texto automática).
   - Auto-filtro e congelamento de painéis ativados.

---

## 🛠️ Requisitos e Configuração

### 1. Chave da API do Google Gemini
1. Acesse [Google AI Studio](https://aistudio.google.com/app/apikey) e gere sua chave gratuita.
2. Abra o arquivo `.env` e cole sua chave:
   ```env
   GEMINI_API_KEY=sua_chave_aqui
   TESSERACT_PATH=C:\Program Files\Tesseract-OCR\tesseract.exe
   ```

### 2. Tesseract OCR
O Tesseract já foi detectado no caminho padrão do Windows:
`C:\Program Files\Tesseract-OCR\tesseract.exe`

---

## 💻 Como Rodar o Programa

Basta abrir o terminal nesta pasta e executar:

```bash
python app.py
```

O sistema exibirá um menu interativo e amigável:
- Escolha o arquivo da pasta `entrada/`
- Escolha as colunas desejadas (ou personalize na hora)
- Selecione as páginas (todas ou intervalo específico)
- Ao final, o arquivo será salvo na pasta `saida/` e você poderá abri-lo automaticamente!
