# Análise Bibliométrica via OpenAlex e Semantic Scholar

Script Python para extração automatizada de indicadores bibliométricos (índices de citação e métricas de impacto autoral) de um corpus de artigos científicos, desenvolvido no âmbito de uma tese de doutoramento sobre **Literacia em Dados e Dados Abertos Governamentais na Administração Pública Portuguesa**.

[![DOI](https://img.shields.io/badge/doi.org/10.5281/zenodo.20581071)](https://doi.org/10.5281/zenodo.20581071)
[![Python 3.13](https://img.shields.io/badge/python-3.13-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

---

## Sobre o projeto

Este script foi desenvolvido para superar as limitações das ferramentas bibliométricas convencionais (VOSviewer, Bibliometrix) na análise de um corpus de nicho (N = 39 artigos), resultante de uma Revisão Sistemática da Literatura (RSL) conduzida segundo o protocolo PRISMA 2020.

Em vez de depender de exportações manuais de citações, o script consulta de forma programática **duas APIs bibliométricas abertas** — [OpenAlex](https://openalex.org) e [Semantic Scholar](https://www.semanticscholar.org) — permitindo a **triangulação de fontes** e a extração automatizada, reproduzível e auditável dos indicadores.

---

## Funcionalidades

- 📚 **Leitura de ficheiros RIS** exportados do Zotero, Scopus ou Web of Science
- 🔍 **Consulta dupla** às APIs OpenAlex e Semantic Scholar
- 🔄 **Estratégia de fallback**: pesquisa por DOI e, em caso de insucesso, por título
- 👤 **Métricas de autor**: h-index, i10-index, ORCID, citações totais, afiliação
- 📊 **Métricas de artigo**: citações, citações influentes, n.º de referências, acesso aberto
- 📈 **Triangulação**: média de citações entre as duas fontes
- 📑 **Exportação Excel** estruturada em quatro folhas formatadas
- ⏱️ **Gestão de rate limiting** com pausas e repetição automática

---

## Indicadores extraídos

| Indicador | OpenAlex | Semantic Scholar |
|---|:---:|:---:|
| Citações do artigo | ✅ | ✅ |
| Citações influentes | — | ✅ |
| N.º de referências | — | ✅ |
| h-index do autor | ✅ | ✅ |
| i10-index do autor | ✅ | — |
| ORCID | ✅ | — |
| Afiliação institucional | ✅ | ✅ |
| Conceitos / áreas temáticas | ✅ | ✅ |
| Acesso aberto (URL) | ✅ | ✅ |

---

## Instalação

Requer **Python 3.13** ou superior.

```bash
# Clonar o repositório
git clone https://github.com/<utilizador>/bibliometria-rsl.git
cd bibliometria-rsl

# Instalar as dependências
pip install -r requirements.txt
```

### Dependências

```
rispy==0.9.0
requests>=2.31.0
pandas>=2.1.0
openpyxl>=3.1.2
tqdm>=4.66.0
```

---

## Utilização

### Com ficheiro RIS (recomendado)

```bash
python bibliometria_rsl_v2.py --ris 39_Artigos_Finais_RSL.ris
```

### Com endereço de email (melhora o desempenho da API OpenAlex)

```bash
python bibliometria_rsl_v2.py --ris 39_Artigos_Finais_RSL.ris --email seu.email@instituicao.pt
```

### Com lista de DOIs (alternativa)

```bash
python bibliometria_rsl_v2.py --dois lista_dois.txt
```

### Argumentos disponíveis

| Argumento | Descrição | Obrigatório |
|---|---|:---:|
| `--ris` | Caminho para o ficheiro `.ris` | Sim* |
| `--dois` | Caminho para ficheiro `.txt` com um DOI por linha | Sim* |
| `--output` | Nome do ficheiro Excel de saída (padrão: `bibliometria_resultados.xlsx`) | Não |
| `--email` | Email para o *polite pool* do OpenAlex | Não |

\* É obrigatório indicar `--ris` **ou** `--dois`.

---

## Estrutura do resultado

O ficheiro Excel gerado contém quatro folhas:

| Folha | Conteúdo |
|---|---|
| `1_Artigos` | Uma linha por artigo, com todas as métricas de ambas as APIs |
| `2_Autores` | Autores únicos, ordenados por h-index, com citações e ORCID |
| `3_Cobertura` | Taxa de cobertura das APIs e estatísticas do corpus |
| `4_Metadados` | Hash SHA-256, data de execução, versão e referência de citação |

---

## Reutilização

A integridade do script pode ser verificada pelo hash **SHA-256**:

```bash
python -c "import hashlib; print(hashlib.sha256(open('bibliometria_rsl_v2.py','rb').read()).hexdigest())"
```

Valor esperado (v2.0):

```
5620162d1abfb26106dba52c43684713102f3d47f361d417f622e7ae66336b1e
```

> **Nota:** As APIs são dinâmicas — os valores de citação aumentam continuamente. Execuções posteriores poderão apresentar valores ligeiramente superiores aos da extração original (2026-06-04).

---

## Resultados de referência

Resultados obtidos na extração original (corpus de 39 artigos, 2026-06-04):

- **Cobertura:** OpenAlex 100% (39/39) · Semantic Scholar 92,3% (36/39)
- **Total de citações (OpenAlex):** 659 · média 16,9 · mediana 11
- **Convergência entre fontes:** r = 0,979 · R² = 0,958 (concordância muito elevada)
- **Autores únicos identificados:** 174

---

## Limitações conhecidas

- A desambiguação de autores depende da consistência dos nomes no ficheiro RIS; nomes abreviados podem gerar registos duplicados.
- Três artigos não foram localizados no Semantic Scholar, ficando sem os indicadores exclusivos dessa API.
- As APIs abertas podem divergir de bases proprietárias (Scopus, Web of Science) para revistas de menor indexação.

---

## Como citar

Se utilizar este script, por favor cite (norma APA 7.ª edição):

```
Teixeira, P. (2026). Script Python para análise bibliométrica via OpenAlex
e Semantic Scholar (Versão 2.0) [Software]. Zenodo.
https://doi.org/10.5281/zenodo.20581071
```

Formato BibTeX:

```bibtex
@software{teixeira_2026_bibliometria,
  author    = {Teixeira, Pedro},
  title      = {Script Python para análise bibliométrica via OpenAlex e Semantic Scholar},
  version   = {2.0},
  year      = {2026},
  publisher = {Zenodo},
  doi       = {10.5281/zenodo.20581071},
  url        = {https://doi.org/10.5281/zenodo.20581071}
}
```

---

## Contexto académico

Este software foi desenvolvido no âmbito do programa de doutoramento **Formación en la Sociedad del Conocimiento** da Universidad de Salamanca, integrando o capítulo de análise bibliométrica de uma Revisão Sistemática da Literatura sobre literacia em dados na administração pública portuguesa.

---

## Licença

Distribuído sob a licença MIT. Consulte o ficheiro [`LICENSE`](LICENSE) para mais informação.

---

## Agradecimentos

- [OpenAlex](https://openalex.org) — OurResearch
- [Semantic Scholar](https://www.semanticscholar.org) — Allen Institute for AI (AI2)
- [rispy](https://github.com/MrTango/rispy) — parsing de ficheiros RIS
