from difflib import SequenceMatcher
from ollama import chat
from pathlib import Path
import pandas as pd
import json
import logging
import re
import time


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger(__name__)

MODELO_IA = 'qwen2.5:1.5b'
LIMITE_SIMILARES = 8
LIMIAR_SIMILARIDADE = 0.55
LIMIAR_AUTO_MATCH = 0.88

NOMES_ABA_ENTRADA = {'sheet1', 'planilha1', 'pagina1', 'página1'}
NOMES_ABA_BANCO = {'banco', 'base', 'materiais'}

COLUNAS_SAIDA = [
    'cadastro sugerido',
    'codigo do material',
    'ncm sugerido',
    'status',
]

OLLAMA_OPTIONS = {
    'num_predict': 150,
    'temperature': 0.1,
    'num_ctx': 4096,
}


class Timer:
    def __init__(self, nome):
        self.nome = nome
        self.elapsed = 0.0

    def __enter__(self):
        self.t0 = time.perf_counter()
        log.info('INICIO: %s', self.nome)
        return self

    def __exit__(self, *args):
        self.elapsed = time.perf_counter() - self.t0
        log.info('FIM: %s | %.2fs', self.nome, self.elapsed)


def normalizar(texto):
    if pd.isna(texto):
        return ''
    texto = str(texto).lower()
    texto = re.sub(r'[^\w\s]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()


def formatar_codigo(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    try:
        numero = float(valor)
        if numero == int(numero):
            return str(int(numero))
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def engine_planilha(arquivo):
    extensao = Path(arquivo).suffix.lower()
    if extensao == '.ods':
        return 'odf'
    return 'openpyxl'


def encontrar_planilha():
    pasta = Path('.')
    for nome in ('planilha.ods', 'planilha.xlsx', 'planilha.xls'):
        caminho = pasta / nome
        if caminho.exists():
            return str(caminho)
    raise FileNotFoundError(
        'Planilha não encontrada. Coloque planilha.ods ou planilha.xlsx na pasta do sistema.'
    )


def resolver_abas(dados):
    mapa = {nome.lower(): nome for nome in dados.keys()}

    aba_banco = next((mapa[n] for n in NOMES_ABA_BANCO if n in mapa), None)
    aba_entrada = next((mapa[n] for n in NOMES_ABA_ENTRADA if n in mapa), None)

    if aba_entrada is None:
        for nome in dados:
            if nome != aba_banco:
                aba_entrada = nome
                break

    if aba_banco is None:
        raise ValueError('Aba "banco" não encontrada na planilha.')

    if aba_entrada is None:
        raise ValueError('Aba de entrada (materiais) não encontrada na planilha.')

    return aba_entrada, aba_banco


def normalizar_colunas_banco(banco):
    banco.columns = banco.columns.str.strip().str.lower()

    if 'codigo' not in banco.columns:
        for candidato in ('codigos', 'código', 'cod'):
            if candidato in banco.columns:
                banco['codigo'] = banco[candidato]
                break

    for coluna in ('codigo', 'ncm', 'tipo'):
        if coluna not in banco.columns:
            banco[coluna] = ''

    banco['codigo'] = banco['codigo'].apply(formatar_codigo)
    banco['material_norm'] = banco['material'].apply(normalizar)
    return banco


def carregar_banco(arquivo):
    engine = engine_planilha(arquivo)
    dados = pd.read_excel(arquivo, sheet_name=None, engine=engine)
    aba_entrada, aba_banco = resolver_abas(dados)

    planilha = dados[aba_entrada].copy()
    banco = dados[aba_banco].copy()

    planilha.columns = planilha.columns.str.strip().str.lower()
    banco = normalizar_colunas_banco(banco)

    return planilha, banco, arquivo, engine, aba_entrada, aba_banco


def buscar_exato(descricao, banco):
    descricao_norm = normalizar(descricao)
    match = banco[banco['material_norm'] == descricao_norm]
    if match.empty:
        return None
    return match.iloc[0]


def buscar_similares(descricao, banco, limite=LIMITE_SIMILARES):
    desc_norm = normalizar(descricao)
    scores = []
    for _, linha in banco.iterrows():
        score = SequenceMatcher(None, desc_norm, linha['material_norm']).ratio()
        if score >= LIMIAR_SIMILARIDADE:
            scores.append((score, linha))

    scores.sort(key=lambda item: item[0], reverse=True)
    return scores[:limite]


def formatar_similares(similares_com_score):
    if not similares_com_score:
        return 'Nenhum material parecido encontrado.'

    linhas = []
    for _, linha in similares_com_score:
        linhas.append(
            f"CODIGO {linha.get('codigo', '')} | {linha.get('material', '')} | "
            f"NCM {linha.get('ncm', '')} | TIPO {linha.get('tipo', '')}"
        )
    return '\n'.join(linhas)


def interpretar_resposta_ia(conteudo):
    match = re.search(r'\{.*\}', conteudo, re.DOTALL)
    if not match:
        raise ValueError('JSON não encontrado')

    dados = json.loads(match.group(0))
    return {
        'status': str(dados.get('status', 'NOVO')).upper(),
        'codigo': formatar_codigo(dados.get('codigo', '')),
        'cadastro': str(dados.get('cadastro', '')).strip(),
        'ncm': str(dados.get('ncm', '')).strip(),
        'tipo': str(dados.get('tipo', '')).strip(),
    }


def montar_prompt(descricao, similares_com_score):
    return f"""Você decifra nomes abreviados de materiais industriais em português brasileiro.

TAREFA: expandir abreviações da nota fiscal para o nome completo de cadastro da empresa.

REGRA PRINCIPAL: NUNCA repita o texto da nota. Sempre escreva o nome por extenso.

Exemplos corretos:
- "Pont. 200V retific." -> "Ponte retificadora 200V"
- "tubo ret 10x30mm" -> "Tubo reto 10x30mm"
- "Cab. pp 70mm" -> "Cabo polipropileno 70mm"
- "Bomb. pneum s/bc pt" -> "Bomba pneumatica sem boca preta"
- "cont. 24v/220w" -> "Contator 24V 220W"
- "paraf sext 1/4" -> "Parafuso sextavado 1/4"
- "tubo pa 6" -> "Tubo PA 6"
- "Desinfetante alf. 12/500" -> "Desinfetante alfazema 12x500ml"

Abreviações comuns:
Pont.=Ponte, retific.=retificadora, ret=reto, Cab.=Cabo, pp=polipropileno,
Bomb.=Bomba, pneum.=pneumatica, cont.=contator, paraf.=parafuso,
sext.=sextavado, alf.=alfazema, galv.=galvanizado, s/=sem, c/=com.

Item da nota: {descricao}

Materiais parecidos no banco da empresa:
{formatar_similares(similares_com_score)}

Se o item JÁ EXISTE no banco (mesmo produto), use status JA_EXISTE, preencha codigo e deixe cadastro vazio.
Se é material NOVO, use status NOVO, expanda o nome em cadastro, sugira ncm e tipo (UN/KG/MT/L/PC/CX/RL).

Responda APENAS JSON puro:
{{"status":"JA_EXISTE ou NOVO","codigo":"","cadastro":"","ncm":"","tipo":""}}"""


def preaquecer_modelo():
    with Timer('preaquecer_modelo'):
        chat(
            model=MODELO_IA,
            messages=[{'role': 'user', 'content': 'ok'}],
            options={'num_predict': 1},
        )


def consultar_ia(descricao, similares_com_score):
    prompt = montar_prompt(descricao, similares_com_score)
    log.info('Prompt: %d chars, ~%d tokens estimados', len(prompt), len(prompt) // 4)
    log.info('Similares: %d', len(similares_com_score))

    conteudo = ''
    for chunk in chat(
        model=MODELO_IA,
        messages=[{'role': 'user', 'content': prompt}],
        stream=True,
        options=OLLAMA_OPTIONS,
    ):
        pedaco = chunk.message.content or ''
        conteudo += pedaco
        if pedaco:
            print('.', end='', flush=True)
    print()

    resultado = interpretar_resposta_ia(conteudo.strip())
    return corrigir_cadastro_copiado(descricao, resultado)


def cadastro_parece_copia(descricao, cadastro):
    if not cadastro:
        return True
    desc_norm = normalizar(descricao)
    cad_norm = normalizar(cadastro)
    if desc_norm == cad_norm:
        return True
    return SequenceMatcher(None, desc_norm, cad_norm).ratio() >= 0.82


def expandir_abreviacoes_basicas(descricao):
    texto = str(descricao).strip()
    substituicoes = [
        (r'\bPont\.?\s*', 'Ponte '),
        (r'\bretific\.?\s*', 'retificadora '),
        (r'\bret\b(?=\s+\d|\s+\d+x|\s*$)', 'reto'),
        (r'\bCab\.?\s*', 'Cabo '),
        (r'\bpp\b', 'polipropileno'),
        (r'\bBomb\.?\s*', 'Bomba '),
        (r'\bpneum\.?\s*', 'pneumatica '),
        (r'\bcont\.?\s*', 'Contator '),
        (r'\bparaf\.?\s*', 'Parafuso '),
        (r'\bsext\.?\s*', 'sextavado '),
        (r'\balf\.?\s*', 'alfazema '),
        (r'\bgalv\.?\s*', 'galvanizado '),
        (r'\btubo\b', 'Tubo'),
        (r'\bs/\b', 'sem '),
        (r'\bc/\b', 'com '),
    ]
    for padrao, substituto in substituicoes:
        texto = re.sub(padrao, substituto, texto, flags=re.IGNORECASE)
    texto = re.sub(r'\s+', ' ', texto).strip().rstrip('.')
    if texto:
        return texto[0].upper() + texto[1:]
    return texto


def corrigir_cadastro_copiado(descricao, resultado):
    if resultado['status'] != 'NOVO':
        return resultado
    if not cadastro_parece_copia(descricao, resultado['cadastro']):
        return resultado

    expandido = expandir_abreviacoes_basicas(descricao)
    if expandido and not cadastro_parece_copia(descricao, expandido):
        resultado['cadastro'] = expandido
        log.info('Cadastro corrigido por regras: %s', expandido)
    return resultado


def resultado_de_linha(linha, status='JA_EXISTE'):
    return {
        'status': status,
        'codigo': formatar_codigo(linha.get('codigo', '')),
        'cadastro': '',
        'ncm': str(linha.get('ncm', '')).strip() if not pd.isna(linha.get('ncm', '')) else '',
        'tipo': str(linha.get('tipo', '')).strip() if not pd.isna(linha.get('tipo', '')) else '',
    }


def processar_material(descricao, banco, cache=None):
    metricas = {}
    desc_norm = normalizar(descricao)

    if cache is not None and desc_norm in cache:
        log.info('Cache hit: %s', descricao[:40])
        return cache[desc_norm], {'cache': True}

    t0 = time.perf_counter()
    exato = buscar_exato(descricao, banco)
    metricas['buscar_exato'] = time.perf_counter() - t0

    if exato is not None:
        metricas['buscar_similares'] = 0.0
        metricas['ollama'] = 0.0
        metricas['origem'] = 'exato'
        resultado = resultado_de_linha(exato)
        if cache is not None:
            cache[desc_norm] = resultado
        log.info('METRICAS %s: %s', descricao[:40], metricas)
        return resultado, metricas

    t0 = time.perf_counter()
    similares_com_score = buscar_similares(descricao, banco)
    metricas['buscar_similares'] = time.perf_counter() - t0

    if similares_com_score and similares_com_score[0][0] >= LIMIAR_AUTO_MATCH:
        score, top = similares_com_score[0]
        metricas['ollama'] = 0.0
        metricas['origem'] = f'auto_match ({score:.2f})'
        resultado = resultado_de_linha(top)
        if cache is not None:
            cache[desc_norm] = resultado
        log.info('METRICAS %s: %s', descricao[:40], metricas)
        return resultado, metricas

    try:
        t0 = time.perf_counter()
        resultado = consultar_ia(descricao, similares_com_score)
        metricas['ollama'] = time.perf_counter() - t0
        metricas['origem'] = 'ia'
    except Exception as erro:
        log.error('Erro ao processar "%s": %s', descricao, erro)
        metricas['ollama'] = time.perf_counter() - t0
        metricas['origem'] = 'erro'
        resultado = {
            'status': 'NOVO',
            'codigo': '',
            'cadastro': expandir_abreviacoes_basicas(descricao) or str(descricao).upper(),
            'ncm': '',
            'tipo': '',
        }

    if cache is not None:
        cache[desc_norm] = resultado
    log.info('METRICAS %s: %s', descricao[:40], metricas)
    return resultado, metricas


def preparar_colunas_saida(planilha):
    for coluna in COLUNAS_SAIDA:
        if coluna not in planilha.columns:
            planilha[coluna] = ''
        planilha[coluna] = planilha[coluna].astype('object')
    return planilha


def gravar_resultado(planilha, posicao, resultado):
    if resultado['status'] == 'JA_EXISTE':
        planilha.at[posicao, 'codigo do material'] = resultado['codigo']
        planilha.at[posicao, 'cadastro sugerido'] = ''
        planilha.at[posicao, 'ncm sugerido'] = resultado['ncm']
        planilha.at[posicao, 'status'] = 'JA_EXISTE'
    else:
        planilha.at[posicao, 'codigo do material'] = ''
        planilha.at[posicao, 'cadastro sugerido'] = resultado['cadastro']
        planilha.at[posicao, 'ncm sugerido'] = resultado['ncm']
        planilha.at[posicao, 'status'] = resultado['status']


def main():
    cache = {}

    with Timer('carregar_banco') as timer_banco:
        arquivo = encontrar_planilha()
        planilha, banco, arquivo, engine, aba_entrada, aba_banco = carregar_banco(arquivo)

    log.info('Planilha: %s (mesma planilha será atualizada)', arquivo)
    log.info('Feche o LibreOffice Calc antes de rodar, se a planilha estiver aberta.')

    planilha.reset_index(drop=True, inplace=True)
    preparar_colunas_saida(planilha)

    preaquecer_modelo()

    for posicao, row in planilha.iterrows():
        descricao = row['material']

        if pd.isna(descricao) or not str(descricao).strip():
            continue

        log.info('Processando [%d]: %s', posicao, descricao)
        resultado, metricas = processar_material(descricao, banco, cache=cache)
        gravar_resultado(planilha, posicao, resultado)

        if resultado['status'] == 'JA_EXISTE':
            log.info('%s -> CODIGO: %s', descricao, resultado['codigo'])
        else:
            log.info('%s -> %s', descricao, resultado['cadastro'])

    with Timer('salvar_planilha'):
        with pd.ExcelWriter(arquivo, engine=engine) as writer:
            planilha.to_excel(writer, sheet_name=aba_entrada, index=False)
            banco.drop(columns=['material_norm'], errors='ignore').to_excel(
                writer, sheet_name=aba_banco, index=False
            )

    log.info('Finalizado. Planilha atualizada: %s | carregar_banco: %.2fs', arquivo, timer_banco.elapsed)


if __name__ == '__main__':
    main()
