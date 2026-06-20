from difflib import SequenceMatcher
from ollama import chat
from openpyxl import load_workbook
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
LIMIAR_AUTO_MATCH = 0.78
LIMIAR_MATCH_FORTE = 0.70
LINHA_CABECALHO = 1

UNIDADES_MEDIDA = {'UN', 'MT', 'PC', 'KG', 'CX', 'RL', 'LT', 'L', 'M', 'M2', 'M3', 'GL', 'PAR'}

NOMES_ABA_ENTRADA = {'sheet1', 'planilha1', 'pagina1', 'página1'}
NOMES_ABA_BANCO = {'banco', 'base', 'materiais'}

MAPEAMENTO_COLUNAS_SAIDA = {
    'cadastro sugerido': ['cadastro sugerido'],
    'codigo existente': ['codigo existente', 'codigo do material', 'codigo'],
    'ncm sugerido': ['ncm sugerido'],
    'tipo sugerido': ['tipo sugerido', 'tipo'],
    'status': ['status'],
}

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


def normalizar_cabecalho(texto):
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ''
    return str(texto).strip().lower()


def formatar_codigo(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    try:
        numero = float(valor)
        if numero == int(numero):
            return str(int(numero))
    except (TypeError, ValueError):
        pass
    texto = str(valor).strip()
    if parece_unidade(texto):
        return ''
    return texto


def parece_unidade(valor):
    if pd.isna(valor) or str(valor).strip() == '':
        return False
    texto = str(valor).strip().upper()
    return len(texto) <= 3 and texto.isalpha() and texto in UNIDADES_MEDIDA


def codigo_valido(valor):
    codigo = formatar_codigo(valor)
    if not codigo or parece_unidade(codigo):
        return False
    return codigo.isdigit()


def formatar_valor_campo(valor, padrao=''):
    if pd.isna(valor):
        return padrao
    texto = str(valor).strip()
    if texto.lower() in ('nan', 'none', 'null', '<na>'):
        return padrao
    return texto


def validar_ncm(valor):
    texto = formatar_valor_campo(valor)
    if not texto:
        return ''
    texto = re.sub(r'^ncm\s*', '', texto, flags=re.IGNORECASE).strip()
    if texto.lower() in ('nan', 'none'):
        return ''
    if re.search(r'[a-zA-Z]{4,}', texto):
        return ''
    digitos = re.sub(r'[^\d]', '', texto)
    if 6 <= len(digitos) <= 10:
        return digitos
    return ''


def limpar_valor_ia(valor, tipo_campo='texto'):
    if pd.isna(valor):
        return ''
    texto = str(valor).strip()
    if texto.lower() in ('nan', 'none', 'null', ''):
        return ''
    if re.match(r'^(ncm|tipo)\s+nan$', texto, re.IGNORECASE):
        return ''
    if tipo_campo in ('ncm', 'tipo') and texto.lower().startswith(f'{tipo_campo} '):
        texto = texto.split(' ', 1)[1].strip()
        if texto.lower() == 'nan':
            return ''

    if tipo_campo == 'ncm':
        return validar_ncm(texto)
    if tipo_campo == 'tipo':
        texto = texto.upper()
        return texto if parece_unidade(texto) else ''
    if tipo_campo == 'codigo':
        return formatar_codigo(texto) if codigo_valido(texto) else ''
    return texto


def valor_celula(valor):
    if valor is None or valor == '':
        return None
    return valor


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


def cabecalhos_da_planilha(df):
    return [str(col).strip() if not pd.isna(col) else '' for col in df.columns]


def encontrar_coluna_por_aliases(cabecalhos, aliases):
    norm_para_indice = {}
    for indice, cabecalho in enumerate(cabecalhos):
        norm = normalizar_cabecalho(cabecalho)
        if norm and norm not in norm_para_indice:
            norm_para_indice[norm] = indice

    for alias in aliases:
        indice = norm_para_indice.get(alias)
        if indice is not None:
            return indice
    return None


def mapear_colunas_saida(cabecalhos):
    mapa = {}
    for canonico, aliases in MAPEAMENTO_COLUNAS_SAIDA.items():
        indice = encontrar_coluna_por_aliases(cabecalhos, aliases)
        if indice is not None:
            mapa[canonico] = indice + 1
    return mapa


def mapear_colunas_saida_openpyxl(ws):
    cabecalhos = [
        ws.cell(row=LINHA_CABECALHO, column=col).value
        for col in range(1, ws.max_column + 1)
    ]
    return mapear_colunas_saida(cabecalhos)


def normalizar_colunas_banco(banco):
    banco.columns = banco.columns.str.strip().str.lower()

    for coluna in ('codigos', 'codigo', 'ncm', 'tipo', 'material'):
        if coluna not in banco.columns:
            banco[coluna] = ''

    codigos_resolvidos = []
    tipos_resolvidos = []

    for _, linha in banco.iterrows():
        codigo_numerico = formatar_codigo(linha.get('codigos'))
        if not codigo_valido(codigo_numerico):
            candidato = formatar_codigo(linha.get('codigo'))
            if codigo_valido(candidato):
                codigo_numerico = candidato

        tipo = formatar_valor_campo(linha.get('tipo'))
        if tipo and not parece_unidade(tipo):
            tipo = ''

        if not tipo and parece_unidade(linha.get('codigo')):
            tipo = str(linha.get('codigo')).strip().upper()

        codigos_resolvidos.append(codigo_numerico)
        tipos_resolvidos.append(tipo)

    banco['codigo'] = codigos_resolvidos
    banco['tipo'] = tipos_resolvidos
    banco['material_norm'] = banco['material'].apply(normalizar)
    return banco


def carregar_dados(arquivo):
    engine = engine_planilha(arquivo)
    dados = pd.read_excel(arquivo, sheet_name=None, engine=engine)
    aba_entrada, aba_banco = resolver_abas(dados)

    planilha = dados[aba_entrada]
    banco = dados[aba_banco].copy()
    banco = normalizar_colunas_banco(banco)

    cabecalhos = cabecalhos_da_planilha(planilha)
    col_material = encontrar_coluna_por_aliases(cabecalhos, ['material'])
    if col_material is None:
        raise ValueError('Coluna "material" não encontrada na aba de entrada.')

    mapa_colunas_saida = mapear_colunas_saida(cabecalhos)
    faltando = set(MAPEAMENTO_COLUNAS_SAIDA) - set(mapa_colunas_saida)
    if faltando:
        log.warning(
            'Colunas de saída não encontradas na planilha (serão ignoradas na gravação): %s',
            ', '.join(sorted(faltando)),
        )

    return planilha, banco, arquivo, aba_entrada, aba_banco, col_material, mapa_colunas_saida


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


def variacoes_descricao(descricao):
    opcoes = [descricao]
    expandido = expandir_abreviacoes_basicas(descricao)
    if expandido and normalizar(expandido) != normalizar(descricao):
        opcoes.append(expandido)
    return opcoes


def buscar_melhor_match(descricao, banco):
    for variacao in variacoes_descricao(descricao):
        exato = buscar_exato(variacao, banco)
        if exato is not None:
            return exato, 1.0, 'exato'

    melhor_score = 0.0
    melhor_linha = None
    for variacao in variacoes_descricao(descricao):
        for score, linha in buscar_similares(variacao, banco):
            if score > melhor_score:
                melhor_score = score
                melhor_linha = linha

    return melhor_linha, melhor_score, 'similar'


def formatar_similares(similares_com_score):
    if not similares_com_score:
        return 'Nenhum material parecido encontrado.'

    linhas = []
    for _, linha in similares_com_score:
        codigo = formatar_valor_campo(linha.get('codigo', ''))
        material = formatar_valor_campo(linha.get('material', ''))
        ncm = formatar_valor_campo(validar_ncm(linha.get('ncm', '')))
        tipo = formatar_valor_campo(linha.get('tipo', ''))
        linhas.append(
            f"CODIGO {codigo or '-'} | {material or '-'} | "
            f"NCM {ncm or '-'} | TIPO {tipo or '-'}"
        )
    return '\n'.join(linhas)


def interpretar_resposta_ia(conteudo):
    match = re.search(r'\{.*\}', conteudo, re.DOTALL)
    if not match:
        raise ValueError('JSON não encontrado')

    dados = json.loads(match.group(0))
    return {
        'status': str(dados.get('status', 'NOVO')).upper(),
        'codigo': limpar_valor_ia(dados.get('codigo', ''), 'codigo'),
        'cadastro': limpar_valor_ia(dados.get('cadastro', ''), 'texto'),
        'ncm': limpar_valor_ia(dados.get('ncm', ''), 'ncm'),
        'tipo': limpar_valor_ia(dados.get('tipo', ''), 'tipo'),
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
        (r'\bRol\.?\s*', 'Rolamento '),
        (r'\besfe\.?\s*', 'esferas '),
        (r'\brad\.?\s*', 'radial '),
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
    codigo = formatar_codigo(linha.get('codigo', ''))
    if not codigo_valido(codigo):
        codigo = ''

    ncm = validar_ncm(linha.get('ncm', ''))
    tipo = formatar_valor_campo(linha.get('tipo', ''))
    if tipo and not parece_unidade(tipo):
        tipo = ''

    return {
        'status': status,
        'codigo': codigo,
        'cadastro': '',
        'ncm': ncm,
        'tipo': tipo.upper() if tipo else '',
    }


def processar_material(descricao, banco, cache=None):
    metricas = {}
    desc_norm = normalizar(descricao)

    if cache is not None and desc_norm in cache:
        log.info('Cache hit: %s', descricao[:40])
        return cache[desc_norm], {'cache': True}

    t0 = time.perf_counter()
    melhor_linha, melhor_score, origem_busca = buscar_melhor_match(descricao, banco)
    metricas['buscar_exato'] = time.perf_counter() - t0
    metricas['buscar_similares'] = 0.0

    if melhor_linha is not None and origem_busca == 'exato':
        metricas['ollama'] = 0.0
        metricas['origem'] = 'exato'
        resultado = resultado_de_linha(melhor_linha)
        if cache is not None:
            cache[desc_norm] = resultado
        log.info('METRICAS %s: %s', descricao[:40], metricas)
        return resultado, metricas

    if melhor_linha is not None and codigo_valido(melhor_linha.get('codigo', '')):
        if melhor_score >= LIMIAR_AUTO_MATCH:
            metricas['ollama'] = 0.0
            metricas['origem'] = f'auto_match ({melhor_score:.2f})'
            resultado = resultado_de_linha(melhor_linha)
            if cache is not None:
                cache[desc_norm] = resultado
            log.info('METRICAS %s: %s', descricao[:40], metricas)
            return resultado, metricas

        if melhor_score >= LIMIAR_MATCH_FORTE:
            metricas['ollama'] = 0.0
            metricas['origem'] = f'match_forte ({melhor_score:.2f})'
            resultado = resultado_de_linha(melhor_linha)
            if cache is not None:
                cache[desc_norm] = resultado
            log.info('METRICAS %s: %s', descricao[:40], metricas)
            return resultado, metricas

    similares_com_score = buscar_similares(descricao, banco)
    expandido = expandir_abreviacoes_basicas(descricao)
    if expandido and normalizar(expandido) != normalizar(descricao):
        similares_expandido = buscar_similares(expandido, banco)
        similares_com_score = sorted(
            {
                (s, l['material'], l['codigo']): (s, l)
                for s, l in similares_com_score + similares_expandido
            }.values(),
            key=lambda item: item[0],
            reverse=True,
        )[:LIMITE_SIMILARES]
    metricas['buscar_similares'] = time.perf_counter() - t0

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


def montar_valores_saida(resultado):
    codigo = resultado['codigo']
    if not codigo_valido(codigo):
        codigo = ''

    if resultado['status'] == 'JA_EXISTE':
        return {
            'cadastro sugerido': None,
            'codigo existente': valor_celula(codigo),
            'ncm sugerido': valor_celula(validar_ncm(resultado['ncm'])),
            'tipo sugerido': valor_celula(resultado['tipo']),
            'status': 'JA_EXISTE',
        }

    return {
        'cadastro sugerido': valor_celula(resultado['cadastro']),
        'codigo existente': None,
        'ncm sugerido': valor_celula(validar_ncm(resultado['ncm'])),
        'tipo sugerido': valor_celula(resultado['tipo']),
        'status': valor_celula(resultado['status']),
    }


def salvar_resultados_preservando_formato(arquivo, aba_entrada, resultados, mapa_colunas_saida):
    extensao = Path(arquivo).suffix.lower()
    if extensao == '.ods':
        log.error(
            'Gravação com formatação preservada não suporta .ods. '
            'Salve a planilha como .xlsx no LibreOffice Calc e rode novamente.'
        )
        return False

    if extensao != '.xlsx':
        log.error('Formato %s não suportado para gravação preservando formatação.', extensao)
        return False

    wb = load_workbook(arquivo)
    if aba_entrada not in wb.sheetnames:
        raise ValueError(f'Aba "{aba_entrada}" não encontrada no arquivo Excel.')

    ws = wb[aba_entrada]
    mapa_openpyxl = mapear_colunas_saida_openpyxl(ws)
    colunas_gravar = {
        canonico: mapa_openpyxl[canonico]
        for canonico in mapa_colunas_saida
        if canonico in mapa_openpyxl
    }

    for linha_idx, valores in resultados.items():
        linha_excel = linha_idx + LINHA_CABECALHO + 1
        for coluna_canonica, coluna_excel in colunas_gravar.items():
            ws.cell(row=linha_excel, column=coluna_excel).value = valores.get(coluna_canonica)

    wb.save(arquivo)
    return True


def main():
    cache = {}
    resultados = {}

    with Timer('carregar_banco') as timer_banco:
        arquivo = encontrar_planilha()
        (
            planilha,
            banco,
            arquivo,
            aba_entrada,
            aba_banco,
            col_material,
            mapa_colunas_saida,
        ) = carregar_dados(arquivo)

    log.info('Planilha: %s (somente células de saída serão atualizadas)', arquivo)
    log.info('Feche o LibreOffice Calc antes de rodar, se a planilha estiver aberta.')
    log.info('Aba banco "%s" é somente leitura e não será modificada.', aba_banco)

    preaquecer_modelo()

    for posicao in range(len(planilha)):
        descricao = planilha.iloc[posicao, col_material]

        if pd.isna(descricao) or not str(descricao).strip():
            continue

        log.info('Processando [%d]: %s', posicao, descricao)
        resultado, metricas = processar_material(descricao, banco, cache=cache)
        resultados[posicao] = montar_valores_saida(resultado)

        if resultado['status'] == 'JA_EXISTE':
            log.info('%s -> CODIGO: %s', descricao, resultado['codigo'])
        else:
            log.info('%s -> %s', descricao, resultado['cadastro'])

    with Timer('salvar_planilha'):
        salvo = salvar_resultados_preservando_formato(
            arquivo, aba_entrada, resultados, mapa_colunas_saida
        )

    if salvo:
        log.info(
            'Finalizado. Planilha atualizada (formatação preservada): %s | carregar_banco: %.2fs',
            arquivo,
            timer_banco.elapsed,
        )
    else:
        log.warning('Processamento concluído, mas a planilha não foi salva.')


if __name__ == '__main__':
    main()
