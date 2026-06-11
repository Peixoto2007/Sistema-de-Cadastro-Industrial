from difflib import SequenceMatcher
from ollama import chat
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

ARQUIVO_PLANILHA = 'planilha.xlsx'
ARQUIVO_SAIDA = 'planilha_resultado.xlsx'
MODELO_IA = 'qwen2.5:1.5b'
LIMITE_SIMILARES = 8
LIMIAR_SIMILARIDADE = 0.55
LIMIAR_AUTO_MATCH = 0.92

OLLAMA_OPTIONS = {
    'num_predict': 80,
    'temperature': 0.1,
    'num_ctx': 2048,
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


def similaridade(a, b):
    return SequenceMatcher(None, normalizar(a), normalizar(b)).ratio()


def carregar_banco(arquivo):
    dados = pd.read_excel(arquivo, sheet_name=None)
    planilha = dados['Sheet1']
    banco = dados['banco']

    planilha.columns = planilha.columns.str.strip().str.lower()
    banco.columns = banco.columns.str.strip().str.lower()

    for coluna in ('codigo', 'ncm', 'tipo'):
        if coluna not in banco.columns:
            banco[coluna] = ''

    banco['material_norm'] = banco['material'].apply(normalizar)
    return planilha, banco


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
            f"{linha.get('codigo', '')} | {linha.get('material', '')} | "
            f"{linha.get('ncm', '')} | {linha.get('tipo', '')}"
        )
    return '\n'.join(linhas)


def interpretar_resposta_ia(conteudo):
    match = re.search(r'\{.*\}', conteudo, re.DOTALL)
    if not match:
        raise ValueError('JSON não encontrado')

    dados = json.loads(match.group(0))
    return {
        'status': str(dados.get('status', 'NOVO')).upper(),
        'codigo': str(dados.get('codigo', '')).strip(),
        'cadastro': str(dados.get('cadastro', '')).strip(),
        'ncm': str(dados.get('ncm', '')).strip(),
        'tipo': str(dados.get('tipo', '')).strip(),
    }


def montar_prompt(descricao, similares_com_score):
    return f"""Classifique o material abaixo.

Item: {descricao}

Referências do banco:
{formatar_similares(similares_com_score)}

Retorne JSON: {{"status":"JA_EXISTE|NOVO","codigo":"","cadastro":"","ncm":"","tipo":""}}
Regras: JA_EXISTE=só se for o mesmo material; codigo preenchido, cadastro vazio.
NOVO=cadastro completo + ncm + tipo (UN/KG/MT/L/PC/CX/RL). Responda só JSON, sem markdown."""


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

    return interpretar_resposta_ia(conteudo.strip())


def resultado_de_linha(linha, status='JA_EXISTE'):
    return {
        'status': status,
        'codigo': str(linha.get('codigo', '')).strip(),
        'cadastro': '',
        'ncm': str(linha.get('ncm', '')).strip(),
        'tipo': str(linha.get('tipo', '')).strip(),
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
            'status': 'ERRO',
            'codigo': '',
            'cadastro': str(descricao).upper(),
            'ncm': '',
            'tipo': '',
        }

    if cache is not None:
        cache[desc_norm] = resultado
    log.info('METRICAS %s: %s', descricao[:40], metricas)
    return resultado, metricas


def montar_resultado(resultado):
    if resultado['status'] == 'JA_EXISTE' and resultado['codigo']:
        partes = [f"CODIGO: {resultado['codigo']}"]
        if resultado['ncm']:
            partes.append(f"NCM: {resultado['ncm']}")
        if resultado['tipo']:
            partes.append(f"TIPO: {resultado['tipo']}")
        return ' | '.join(partes)

    partes = []
    if resultado['cadastro']:
        partes.append(resultado['cadastro'])
    if resultado['ncm']:
        partes.append(f"NCM: {resultado['ncm']}")
    if resultado['tipo']:
        partes.append(f"TIPO: {resultado['tipo']}")
    return ' | '.join(partes) if partes else ''


def main():
    cache = {}

    with Timer('carregar_banco') as timer_banco:
        planilha, banco = carregar_banco(ARQUIVO_PLANILHA)

    preaquecer_modelo()

    for coluna in ('cadastro sugerido', 'codigo', 'ncm sugerido', 'tipo', 'status', 'resultado'):
        if coluna not in planilha.columns:
            planilha[coluna] = ''

    for i, descricao in enumerate(planilha['material']):
        if pd.isna(descricao) or not str(descricao).strip():
            continue

        log.info('Processando [%d]: %s', i, descricao)
        resultado, metricas = processar_material(descricao, banco, cache=cache)

        planilha.loc[i, 'status'] = resultado['status']
        planilha.loc[i, 'codigo'] = resultado['codigo']
        planilha.loc[i, 'cadastro sugerido'] = resultado['cadastro']
        planilha.loc[i, 'ncm sugerido'] = resultado['ncm']
        planilha.loc[i, 'tipo'] = resultado['tipo']
        planilha.loc[i, 'resultado'] = montar_resultado(resultado)

        log.info('%s -> %s', descricao, planilha.loc[i, 'resultado'])

    with Timer('salvar_planilha'):
        with pd.ExcelWriter(ARQUIVO_SAIDA, engine='openpyxl') as writer:
            planilha.to_excel(writer, sheet_name='Sheet1', index=False)
            banco.drop(columns=['material_norm'], errors='ignore').to_excel(
                writer, sheet_name='banco', index=False
            )

    log.info('Finalizado. Saída: %s | carregar_banco: %.2fs', ARQUIVO_SAIDA, timer_banco.elapsed)


if __name__ == '__main__':
    main()
