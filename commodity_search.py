# commodity_search.py
# Este módulo fornece funções auxiliares para um formulário de busca de commodities em Python.
# As funções permitem:
# - Extrair opções únicas de campos para preencher combos (dropdowns).
# - Limpar os campos do formulário.
# - Buscar commodities em uma base JSON a partir de filtros.
# - Gerar a visualização dos resultados em uma tabela formatada (texto puro para terminal, mas fácil adaptar ao Streamlit ou outros frameworks).
# O arquivo commodities_base.json deve estar no mesmo diretório e conter uma lista de dicionários (cada um representando uma commodity).

import json
from typing import List, Dict, Any

# Função para carregar a base de commodities do arquivo JSON.
# Retorna uma lista de dicionários com os dados das commodities.
def load_commodities_base(json_path: str = '/Users/giullianoaccarinideluccia/Desktop/Cortex OTC/json-cache/commodities_base.json') -> List[Dict[str, Any]]:
    """
    Carrega a base de commodities de um arquivo JSON.
    :param json_path: Caminho para o arquivo JSON.
    :return: Lista de dicionários, cada um representando uma commodity.
    """
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

# Função para extrair opções únicas para um campo, útil para preencher combos/dropdowns.
def get_combo_options(base: List[Dict[str, Any]], field_name: str) -> List[str]:
    """
    Extrai valores únicos de um campo específico da base de commodities,
    para uso em combos/dropdowns do formulário.

    - Para o campo 'mes' (ou variações, como 'mês', 'mes_vencimento'), a ordenação é feita numericamente (1, 2, 3 ... 12).
    - Para outros campos, a ordenação é alfabética padrão.
    """
    # Cria um conjunto vazio para garantir unicidade dos valores extraídos
    options_set = set()
    for item in base:
        # Extrai o valor do campo desejado do dicionário de cada commodity
        value = item.get(field_name)
        # Só adiciona valores não nulos, não vazios e não listas vazias
        if value not in (None, '', []):
            # Sempre converte para string para evitar problemas de tipos mistos
            options_set.add(str(value))

    # Bloco especial: se o campo for mês, precisamos ordenar como número!
    # Checa se o campo é algum dos nomes usuais para mês de vencimento
    if field_name.lower() in ['mes', 'mês', 'mes_vencimento', 'mês de vencimento']:
        # Converte todos os valores para int para garantir ordenação numérica
        meses_numericos = []
        for opt in options_set:
            try:
                meses_numericos.append(int(opt))
            except ValueError:
                # Se algum valor não for número, ignora (ex: "Selecione o mês")
                continue
        # Ordena a lista de meses numericamente (1, 2, 3, ..., 12)
        meses_numericos.sort()
        # Converte de volta para string, pois o ComboBox espera strings
        return [str(mes) for mes in meses_numericos]
    else:
        # Para outros campos, ordena alfabeticamente como antes
        return sorted(options_set)
   

# Função para limpar os campos do formulário, útil para resetar combos/dropdowns.
def clear_fields(combo_fields: Dict[str, Any]) -> Dict[str, Any]:
    """
    Reseta os campos do formulário (combos) para o valor padrão ('' ou placeholder).
    :param combo_fields: Dicionário com os nomes dos combos e seus valores atuais.
    :return: Novo dicionário com todos os valores resetados para vazio.
    """
    # Cria um novo dicionário com todos os valores resetados para string vazia.
    cleared_state = {field: '' for field in combo_fields}
    return cleared_state

# Função principal de busca: filtra a base de commodities conforme os filtros selecionados.
def search_commodities(base: List[Dict[str, Any]], filters: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Filtra a base de commodities conforme os filtros fornecidos.
    Se todos os filtros estiverem vazios, retorna a base completa.
    :param base: Lista de commodities (dicionários).
    :param filters: Dicionário com os campos e valores dos filtros; valores vazios são ignorados.
    :return: Lista dos registros que atendem aos filtros.
    """
    # Se todos os filtros estiverem vazios, retorna base completa.
    if all(val in (None, '', []) for val in filters.values()):
        return base

    filtered = []
    for item in base:
        match = True
        for field, value in filters.items():
            # Ignora filtro vazio.
            if value in (None, '', []):
                continue
            # Só aceita se o valor do campo for igual ao filtro.
            if item.get(field) != value:
                match = False
                break
        if match:
            filtered.append(item)
    return filtered

# Função para exibir os resultados em uma tabela de texto simples.
def render_scrollable_table(data: List[Dict[str, Any]]) -> str:
    """
    Gera uma tabela de texto (string) para exibir os resultados.
    Pode ser exibido no terminal ou facilmente adaptado para frameworks web.
    :param data: Lista de dicionários com os resultados.
    :return: String formatada contendo a tabela.
    """
    if not data:
        return "Nenhum resultado encontrado."
    # Descobre as colunas (cabeçalhos) a partir do primeiro item.
    columns = list(data[0].keys())
    # Calcula largura máxima de cada coluna para alinhamento.
    col_widths = {
        col: max(len(str(col)), max(len(str(item.get(col, ""))) for item in data))
        for col in columns
    }
    # Monta o cabeçalho.
    header = ' | '.join([col.ljust(col_widths[col]) for col in columns])
    separator = '-+-'.join(['-' * col_widths[col] for col in columns])
    # Monta as linhas da tabela.
    rows = []
    for idx, item in enumerate(data):
        row = ' | '.join([
            str(item.get(col, '')).ljust(col_widths[col])
            for col in columns
        ])
        rows.append(row)
    # Junta tudo em uma string final.
    table = f"{header}\n{separator}\n" + "\n".join(rows)
    return table