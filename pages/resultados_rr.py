import io
import re
import unicodedata
from pathlib import Path
from typing import BinaryIO

import pandas as pd
import streamlit as st


NEW_COLUMNS = [
    'Set 1_a',
    'Set 2_a',
    'Set 3_a',
    'Set 4_a',
    'Set 5_a',
    'SO_a',
    'Set 1_b',
    'Set 2_b',
    'Set 3_b',
    'Set 4_b',
    'Set 5_b',
    'SO_b',
]

FINAL_CSV_COLUMNS = [
    'RANKING FINAL',
    'ID',
    'NOME',
    'CATEGORIA AGRUPADA',
    'SIGLA',
    'CLUBE',
    'PONTUAÇÃO INDIVIDUAL',
]


def _sanitize_name(value: str) -> str:
    if not value:
        return 'etapa'
    cleaned = re.sub(r'[\\/*?:"<>|]+', '_', str(value)).strip()
    return cleaned or 'etapa'


def _normalize_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    text = str(value).strip()
    if not text:
        return ''
    normalized = unicodedata.normalize('NFKD', text)
    normalized = ''.join(char for char in normalized if not unicodedata.combining(char))
    return normalized.lower()


def _get_row_value(row, *candidates):
    if row is None:
        return ''
    for candidate in candidates:
        if candidate in row:
            return row.get(candidate)

    normalized_targets = {_normalize_text(candidate) for candidate in candidates}
    for key in row.keys():
        if _normalize_text(key) in normalized_targets:
            return row.get(key)
    return ''


def _coerce_numeric(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    if ',' in text and '.' in text:
        if text.rfind(',') > text.rfind('.'):
            text = text.replace('.', '').replace(',', '.')
        else:
            text = text.replace(',', '')
    elif ',' in text:
        text = text.replace(',', '.')

    try:
        return float(text)
    except ValueError:
        return 0.0


def _format_decimal(value, decimals: int = 2) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''

    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        return str(value)

    return f"{numeric_value:.{decimals}f}".replace('.', ',')


def _load_points_lookup(uploaded_file=None) -> dict[int, float]:
    if uploaded_file is not None:
        buffer = io.BytesIO(uploaded_file.getvalue())
        df = pd.read_excel(buffer, engine='openpyxl')
    else:
        path = Path(__file__).resolve().parents[1] / 'Pontos Round.xlsx'
        if not path.exists():
            return {}
        df = pd.read_excel(path, engine='openpyxl')

    lookup = {}
    for _, row in df.iterrows():
        rank_value = row.get('Rank')
        points_value = row.get('Pontos Round')
        if rank_value is None or points_value is None:
            continue
        try:
            rank_num = int(float(rank_value))
        except (TypeError, ValueError):
            continue
        lookup[rank_num] = _coerce_numeric(points_value)
    return lookup


def _load_bonus_rules(uploaded_file=None) -> pd.DataFrame:
    if uploaded_file is not None:
        buffer = io.BytesIO(uploaded_file.getvalue())
        df = pd.read_excel(buffer, engine='openpyxl')
    else:
        path = Path(__file__).resolve().parents[1] / 'Bonus Round.xlsx'
        if not path.exists():
            return pd.DataFrame(columns=['Cat Combate', 'Gênero', 'Pontuação Mínima', 'Pontuação Máxima', 'Bonus'])
        df = pd.read_excel(path, engine='openpyxl')

    return df


def _load_media_bonus_rules(uploaded_file=None) -> pd.DataFrame:
    if uploaded_file is not None:
        buffer = io.BytesIO(uploaded_file.getvalue())
        df = pd.read_excel(buffer, engine='openpyxl')
    else:
        path = Path(__file__).resolve().parents[1] / 'Bonus Médias.xlsx'
        if not path.exists():
            return pd.DataFrame(columns=['Cat Combate', 'Pontuação Mínima', 'Pontuação Máxima', 'Bonus'])
        df = pd.read_excel(path, engine='openpyxl')

    return df


def _get_group_bonus(grupo_value) -> float:
    group_text = _normalize_text(grupo_value)
    if not group_text:
        return 0.0
    if group_text in {'1', 'grupo 1', 'g1'}:
        return 3.0
    if group_text in {'2', 'grupo 2', 'g2'}:
        return 2.0
    if group_text in {'3', 'grupo 3', 'g3'}:
        return 1.0
    if 'grupo 1' in group_text or group_text.startswith('1'):
        return 3.0
    if 'grupo 2' in group_text or group_text.startswith('2'):
        return 2.0
    if 'grupo 3' in group_text or group_text.startswith('3'):
        return 1.0
    return 0.0


def build_resultados_template(uploaded_file: BinaryIO) -> tuple[bytes, str]:
    buffer = io.BytesIO(uploaded_file.getvalue())
    df_total = pd.read_excel(buffer, sheet_name='Total')

    if df_total.empty:
        raise ValueError('A aba Total está vazia.')

    etapa_value = ''
    if 'ETAPA' in df_total.columns:
        etapa_value = df_total['ETAPA'].dropna().astype(str).str.strip()
        etapa_value = etapa_value.iloc[0] if not etapa_value.empty else ''

    if not etapa_value:
        raise ValueError('Não foi possível encontrar o nome da etapa na coluna ETAPA.')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name in ['MATCH 1', 'MATCH 2', 'MATCH 3']:
            match_df = df_total.copy()
            if 'match' in match_df.columns:
                match_df = match_df[match_df['match'].astype(str).str.strip().str.upper() == sheet_name.upper()]
            else:
                match_df = match_df.iloc[0:0]

            for col in NEW_COLUMNS:
                match_df[col] = ''

            match_df.to_excel(writer, sheet_name=sheet_name, index=False)

    output_bytes = output.getvalue()
    safe_etapa = _sanitize_name(etapa_value)
    filename = f'{safe_etapa}_template resultados.xlsx'
    return output_bytes, filename


def _load_resultados_prova(uploaded_file: BinaryIO) -> pd.DataFrame:
    filename = getattr(uploaded_file, 'name', '')
    if filename.lower().endswith(('.txt', '.csv')):
        raw_bytes = uploaded_file.getvalue()
        last_error = None
        for encoding in ['utf-8-sig', 'cp1252', 'latin1']:
            try:
                return pd.read_csv(
                    io.BytesIO(raw_bytes),
                    sep=';',
                    dtype=str,
                    encoding=encoding,
                    engine='python',
                )
            except UnicodeDecodeError as exc:
                last_error = exc
        raise ValueError(f'Não foi possível decodificar o arquivo de resultados: {last_error}')
    return pd.read_excel(uploaded_file, dtype=str)


def _load_fpaf_points() -> dict[tuple[int, int], float]:
    candidate_paths = [
        Path(__file__).resolve().parents[1] / 'Pontos FPAF.xlsx',
        Path(__file__).resolve().parent / 'Pontos FPAF.xlsx',
        Path.cwd() / 'Pontos FPAF.xlsx',
    ]
    path = next((candidate for candidate in candidate_paths if candidate.exists()), None)
    if path is None:
        return {}

    df = pd.read_excel(path, sheet_name=0, header=0, engine='openpyxl')
    if df.empty:
        return {}

    rank_column = df.columns[0]
    points_lookup = {}
    for _, row in df.iterrows():
        try:
            rank = int(float(row[rank_column]))
        except (TypeError, ValueError):
            continue

        for athlete_count_column in df.columns[1:]:
            try:
                athlete_count = int(float(athlete_count_column))
            except (TypeError, ValueError):
                continue
            points_lookup[(rank, athlete_count)] = _coerce_numeric(row[athlete_count_column])

    return points_lookup


def _build_final_csv_bytes(rows: list[dict], results_file_name: str) -> tuple[bytes, str]:
    output_df = pd.DataFrame(rows, columns=FINAL_CSV_COLUMNS)
    output = io.StringIO()
    output_df.to_csv(output, sep=';', index=False)
    stem = Path(results_file_name).stem if results_file_name else 'resultados'
    filename = f'{stem} Robin Round.csv'
    return output.getvalue().encode('utf-8-sig'), filename


def _normalize_resultados_prova(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if 'Nome Completo' in df.columns:
        df['Nome Completo'] = df['Nome Completo'].fillna('').astype(str).str.strip()
    elif 'NomeCompleto' in df.columns:
        df['Nome Completo'] = df['NomeCompleto'].fillna('').astype(str).str.strip()
    elif 'NOME' in df.columns:
        df['Nome Completo'] = df['NOME'].fillna('').astype(str).str.strip()
    else:
        family_name = df['FamilyName'] if 'FamilyName' in df.columns else pd.Series('', index=df.index)
        given_name = df['GivenName'] if 'GivenName' in df.columns else pd.Series('', index=df.index)
        family_name = family_name.fillna('').astype(str).str.strip()
        given_name = given_name.fillna('').astype(str).str.strip()
        df['Nome Completo'] = (family_name + ' ' + given_name).str.strip()

    club_column = next((col for col in ['Clube', 'CLUBE', 'Country'] if col in df.columns), None)
    df['Clube'] = df[club_column].fillna('').astype(str).str.strip() if club_column else ''

    sigla_column = next((col for col in ['Sigla', 'SIGLA', 'Noc', 'Club Code'] if col in df.columns), None)
    df['Sigla'] = df[sigla_column].fillna('').astype(str).str.strip() if sigla_column else ''

    id_column = next((col for col in ['ID', 'Id', 'WaID', 'Athlete ID', 'AthleteId'] if col in df.columns), None)
    df['ID'] = df[id_column].fillna('').astype(str).str.strip() if id_column else ''

    for score_col in ['D1 Score', 'D1', 'Round 1']:
        if score_col in df.columns:
            df['Round 1'] = pd.to_numeric(df[score_col].astype(str).str.replace(',', '.', regex=False).str.strip(), errors='coerce').fillna(0)
            break
    else:
        df['Round 1'] = 0

    for score_col in ['D2 Score', 'D2', 'Round 2']:
        if score_col in df.columns:
            df['Round 2'] = pd.to_numeric(df[score_col].astype(str).str.replace(',', '.', regex=False).str.strip(), errors='coerce').fillna(0)
            break
    else:
        df['Round 2'] = 0

    if 'Categoria Quali' in df.columns:
        df['Cat Round'] = df['Categoria Quali'].fillna('').astype(str).str.strip()
    elif 'Categoria' in df.columns:
        df['Cat Round'] = df['Categoria'].fillna('').astype(str).str.strip()
    elif 'Division' in df.columns and 'Class' in df.columns:
        df['Cat Round'] = (df['Division'].fillna('').astype(str).str.strip() + df['Class'].fillna('').astype(str).str.strip()).str.strip()
    elif 'Category' in df.columns:
        df['Cat Round'] = df['Category'].fillna('').astype(str).str.strip()
    else:
        df['Cat Round'] = ''

    df['Total Round'] = df['Round 1'] + df['Round 2']
    return df


def _extract_template_rows(uploaded_file: BinaryIO) -> list[dict]:
    buffer = io.BytesIO(uploaded_file.getvalue())
    excel_file = pd.ExcelFile(buffer, engine='openpyxl')
    rows = []

    for sheet_name in excel_file.sheet_names:
        if sheet_name.lower() == 'total':
            continue
        sheet_df = excel_file.parse(sheet_name)
        for _, row in sheet_df.iterrows():
            genero = str(row.get('genero', '') or '').strip()
            if not genero:
                continue

            for name_col, club_col, rank_col in [('nome a', 'clube a', 'rank a'), ('nome b', 'clube b', 'rank b')]:
                atleta = str(row.get(name_col, '') or '').strip()
                clube = str(row.get(club_col, '') or '').strip()
                rank_value = row.get(rank_col)
                grupo_value = _get_row_value(row, 'GRUPO', 'Grupo', 'grupo')
                if atleta:
                    rows.append({
                        'Atleta': atleta,
                        'Clube': clube,
                        'Cat Combate': genero,
                        'Rank': rank_value,
                        'Grupo': grupo_value,
                    })

    return rows


def _extract_match_rows(uploaded_file: BinaryIO) -> list[dict]:
    buffer = io.BytesIO(uploaded_file.getvalue())
    excel_file = pd.ExcelFile(buffer, engine='openpyxl')
    match_rows = []

    for sheet_name in excel_file.sheet_names:
        if sheet_name.lower() == 'total':
            continue
        sheet_df = excel_file.parse(sheet_name)
        for _, row in sheet_df.iterrows():
            categoria = str(row.get('genero', '') or '').strip()
            if not categoria:
                continue

            match_row = {
                'Cat Combate': categoria,
                'Atleta A': str(row.get('nome a', '') or '').strip(),
                'Atleta B': str(row.get('nome b', '') or '').strip(),
                'Clube A': str(row.get('clube a', '') or '').strip(),
                'Clube B': str(row.get('clube b', '') or '').strip(),
                'Rank A': row.get('rank a'),
                'Rank B': row.get('rank b'),
            }
            for suffix in ['1', '2', '3', '4', '5']:
                match_row[f'Set {suffix}_a'] = _coerce_numeric(row.get(f'Set {suffix}_a'))
                match_row[f'Set {suffix}_b'] = _coerce_numeric(row.get(f'Set {suffix}_b'))
            match_row['SO_a'] = _coerce_numeric(row.get('SO_a'))
            match_row['SO_b'] = _coerce_numeric(row.get('SO_b'))
            match_rows.append(match_row)

    return match_rows


def _determine_match_winner(match_row: dict) -> tuple[str, str]:
    categoria = str(match_row.get('Cat Combate', '') or '').strip().upper()
    atleta_a = str(match_row.get('Atleta A', '') or '').strip()
    atleta_b = str(match_row.get('Atleta B', '') or '').strip()

    if not atleta_a or not atleta_b:
        return '', ''

    if _normalize_text(atleta_a).startswith('bye'):
        return atleta_b, 'B'
    if _normalize_text(atleta_b).startswith('bye'):
        return atleta_a, 'A'

    if categoria.startswith('C'):
        total_a = sum(_coerce_numeric(match_row.get(f'Set {i}_a')) for i in ['1', '2', '3', '4', '5'])
        total_b = sum(_coerce_numeric(match_row.get(f'Set {i}_b')) for i in ['1', '2', '3', '4', '5'])
        if total_a > total_b:
            return atleta_a, 'A'
        if total_b > total_a:
            return atleta_b, 'B'
        if _coerce_numeric(match_row.get('SO_a')) > _coerce_numeric(match_row.get('SO_b')):
            return atleta_a, 'A'
        if _coerce_numeric(match_row.get('SO_b')) > _coerce_numeric(match_row.get('SO_a')):
            return atleta_b, 'B'
        return '', ''

    set_points_a = 0
    set_points_b = 0
    for i in ['1', '2', '3', '4', '5']:
        score_a = _coerce_numeric(match_row.get(f'Set {i}_a'))
        score_b = _coerce_numeric(match_row.get(f'Set {i}_b'))
        if score_a > score_b:
            set_points_a += 2
        elif score_b > score_a:
            set_points_b += 2
        else:
            set_points_a += 1
            set_points_b += 1

        if set_points_a >= 6:
            return atleta_a, 'A'
        if set_points_b >= 6:
            return atleta_b, 'B'

    if set_points_a == 5 and set_points_b == 5:
        if _coerce_numeric(match_row.get('SO_a')) > _coerce_numeric(match_row.get('SO_b')):
            return atleta_a, 'A'
        if _coerce_numeric(match_row.get('SO_b')) > _coerce_numeric(match_row.get('SO_a')):
            return atleta_b, 'B'

    return '', ''


def _build_final_results_workbook(
    template_file: BinaryIO,
    results_file: BinaryIO,
    points_file=None,
    bonus_file=None,
) -> tuple[bytes, str]:
    template_rows = _extract_template_rows(template_file)
    match_rows = _extract_match_rows(template_file)
    if not template_rows:
        raise ValueError('Não foi possível encontrar atletas na planilha de template preenchido.')

    results_df = _normalize_resultados_prova(_load_resultados_prova(results_file))
    results_df['Nome Completo Normalizado'] = results_df['Nome Completo'].astype(str).apply(_normalize_text)
    results_by_name = {}
    for _, row in results_df.iterrows():
        nome = _normalize_text(row.get('Nome Completo', ''))
        if not nome:
            continue
        if nome not in results_by_name:
            results_by_name[nome] = {
                'Clube': str(row.get('Clube', '') or '').strip(),
                'Round 1': float(row.get('Round 1', 0) or 0),
                'Round 2': float(row.get('Round 2', 0) or 0),
                'Total Round': float(row.get('Total Round', 0) or 0),
                'Cat Round': str(row.get('Cat Round', '') or '').strip(),
            }

    points_lookup = _load_points_lookup(points_file)
    bonus_rules = _load_bonus_rules(bonus_file)
    media_bonus_rules = _load_media_bonus_rules()

    match_stats = {}
    for match_row in match_rows:
        winner_name, _ = _determine_match_winner(match_row)
        if not winner_name:
            continue
        winner_key = _normalize_text(winner_name)
        if not winner_key:
            continue
        match_stats.setdefault(winner_key, {'wins': 0, 'shoot_offs': 0.0})['wins'] += 1

        atleta_a = _normalize_text(match_row.get('Atleta A', ''))
        atleta_b = _normalize_text(match_row.get('Atleta B', ''))
        so_a = _coerce_numeric(match_row.get('SO_a'))
        so_b = _coerce_numeric(match_row.get('SO_b'))

        if atleta_a.startswith('bye') or atleta_b.startswith('bye'):
            if winner_key == atleta_a:
                match_stats.setdefault(atleta_a, {'wins': 0, 'shoot_offs': 0.0})['shoot_offs'] += 0.1
            elif winner_key == atleta_b:
                match_stats.setdefault(atleta_b, {'wins': 0, 'shoot_offs': 0.0})['shoot_offs'] += 0.1
        elif so_a > so_b:
            match_stats.setdefault(atleta_a, {'wins': 0, 'shoot_offs': 0.0})['shoot_offs'] += 0.1
        elif so_b > so_a:
            match_stats.setdefault(atleta_b, {'wins': 0, 'shoot_offs': 0.0})['shoot_offs'] += 0.1

    grouped_rows = {}
    for row in template_rows:
        atleta = row['Atleta'].strip()
        if not atleta:
            continue
        if _normalize_text(atleta).startswith('bye'):
            continue
        cat = row['Cat Combate']
        grouped_rows.setdefault(cat, []).append(row)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for categoria, rows in grouped_rows.items():
            unique_rows = []
            seen_names = set()
            for row in rows:
                atleta_key = _normalize_text(row['Atleta'])
                if not atleta_key or atleta_key in seen_names:
                    continue
                seen_names.add(atleta_key)
                result_info = results_by_name.get(atleta_key, {})
                cat_round = str(result_info.get('Cat Round', '') or '').strip()
                total_round = float(result_info.get('Total Round', 0) or 0)

                score_value = 0.0
                try:
                    rank_num = int(float(row.get('Rank', 0) or 0))
                except (TypeError, ValueError):
                    rank_num = None
                if rank_num is not None and rank_num in points_lookup:
                    score_value = points_lookup[rank_num]

                bonus_value = 0.0
                if not bonus_rules.empty:
                    genero_match = ''
                    if cat_round:
                        genero_match = cat_round[-1:].upper()
                    for _, bonus_row in bonus_rules.iterrows():
                        try:
                            cat_combate_match = str(bonus_row.get('Cat Combate', '') or '').strip().upper()
                            genero_bonus = str(bonus_row.get('Gênero', '') or '').strip().upper()
                            min_score = _coerce_numeric(bonus_row.get('Pontuação Mínima'))
                            max_score = _coerce_numeric(bonus_row.get('Pontuação Máxima'))
                        except Exception:
                            continue
                        if (
                            cat_combate_match == categoria.upper()
                            and genero_bonus == genero_match
                            and min_score <= total_round <= max_score
                        ):
                            bonus_value = _coerce_numeric(bonus_row.get('Bonus'))
                            break

                media_bonus_value = 0.0
                if not media_bonus_rules.empty:
                    for _, bonus_row in media_bonus_rules.iterrows():
                        try:
                            cat_combate_match = str(bonus_row.get('Cat Combate', '') or '').strip().upper()
                            min_score = _coerce_numeric(bonus_row.get('Pontuação Mínima'))
                            max_score = _coerce_numeric(bonus_row.get('Pontuação Máxima'))
                        except Exception:
                            continue
                        if cat_combate_match == categoria.upper() and min_score <= total_round <= max_score:
                            media_bonus_value = _coerce_numeric(bonus_row.get('Bonus'))
                            break

                group_bonus_value = _get_group_bonus(row.get('Grupo', ''))

                stats = match_stats.get(atleta_key, {'wins': 0, 'shoot_offs': 0.0})
                unique_rows.append({
                    'Pos Final': len(unique_rows) + 1,
                    'Atleta': row['Atleta'],
                    'Cat Round': cat_round,
                    'Cat Combate': categoria,
                    'Clube': row['Clube'],
                    'Grupo': str(row.get('Grupo', '') or '').strip(),
                    'Round 1': result_info.get('Round 1', 0),
                    'Round 2': result_info.get('Round 2', 0),
                    'Total Round': total_round,
                    'Pontos Round': score_value,
                    'Bonus Round': bonus_value,
                    'Bonus Média': media_bonus_value,
                    'Bonus Grupo': group_bonus_value,
                    'Nº de Vitórias Combates': stats['wins'],
                    'Bonificação Shoot-Offs': stats['shoot_offs'],
                })

            if unique_rows:
                sheet_df = pd.DataFrame(unique_rows)
                for col_name in ['Bonus Média', 'Bonus Grupo']:
                    if col_name not in sheet_df.columns:
                        sheet_df[col_name] = 0.0
                sheet_df['Bonus Média'] = pd.to_numeric(sheet_df['Bonus Média'], errors='coerce').fillna(0.0)
                sheet_df['Bonus Grupo'] = pd.to_numeric(sheet_df['Bonus Grupo'], errors='coerce').fillna(0.0)
                for _, row in sheet_df.iterrows():
                    row_key = _normalize_text(row['Atleta'])
                    if not row_key:
                        continue

                    if categoria.startswith('C'):
                        match_scores = []
                        for match_row in match_rows:
                            if _normalize_text(match_row.get('Atleta A', '')) != row_key and _normalize_text(match_row.get('Atleta B', '')) != row_key:
                                continue
                            if _normalize_text(match_row.get('Atleta A', '')) == row_key:
                                total = sum(_coerce_numeric(match_row.get(f'Set {i}_a')) for i in ['1', '2', '3', '4', '5'])
                            else:
                                total = sum(_coerce_numeric(match_row.get(f'Set {i}_b')) for i in ['1', '2', '3', '4', '5'])
                            match_scores.append(total)
                        media = sum(match_scores) / len(match_scores) if match_scores else 0.0
                    else:
                        set_scores = []
                        for match_row in match_rows:
                            if _normalize_text(match_row.get('Atleta A', '')) != row_key and _normalize_text(match_row.get('Atleta B', '')) != row_key:
                                continue
                            for i in ['1', '2', '3', '4', '5']:
                                if _normalize_text(match_row.get('Atleta A', '')) == row_key:
                                    score = _coerce_numeric(match_row.get(f'Set {i}_a'))
                                else:
                                    score = _coerce_numeric(match_row.get(f'Set {i}_b'))
                                if score > 0:
                                    set_scores.append(score)
                        media = sum(set_scores) / len(set_scores) if set_scores else 0.0

                    sheet_df.loc[sheet_df['Atleta'] == row['Atleta'], 'Média dos Combates'] = round(float(media), 2)

                sheet_df['Grupo'] = sheet_df['Grupo'].fillna('').astype(str).str.strip()
                sheet_df['Ranking Médias no Grupo'] = 0.0
                ranking_bonus_by_position = {1: 2.0, 2: 1.5, 3: 1.0, 4: 0.5}
                for _, group_indexes in sheet_df.groupby('Grupo', sort=False).groups.items():
                    group_df = sheet_df.loc[group_indexes].copy()
                    group_df['_media_num'] = pd.to_numeric(
                        group_df['Média dos Combates'], errors='coerce'
                    ).fillna(0.0)
                    group_df = group_df.sort_values(
                        by=['_media_num'],
                        ascending=[False],
                        kind='mergesort',
                    )
                    previous_media = None
                    current_position = 0
                    for sequence, athlete_index in enumerate(group_df.index, start=1):
                        media_value = float(group_df.at[athlete_index, '_media_num'])
                        if previous_media is None or media_value != previous_media:
                            current_position = sequence
                        sheet_df.at[athlete_index, 'Ranking Médias no Grupo'] = ranking_bonus_by_position.get(
                            current_position, 0.0
                        )
                        previous_media = media_value

                sheet_df['Média dos Combates'] = pd.to_numeric(sheet_df['Média dos Combates'], errors='coerce')
                sheet_df['Ranking Médias no Grupo'] = pd.to_numeric(sheet_df['Ranking Médias no Grupo'], errors='coerce')
                for col_name in ['Bonus Média', 'Bonus Grupo']:
                    if col_name in sheet_df.columns:
                        sheet_df[col_name] = pd.to_numeric(sheet_df[col_name], errors='coerce').fillna(0.0)
                sheet_df['Pontuação Total da Prova'] = (
                    pd.to_numeric(sheet_df.get('Pontos Round', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Bonus Round', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Nº de Vitórias Combates', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Bonificação Shoot-Offs', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Ranking Médias no Grupo', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Bonus Média', 0), errors='coerce').fillna(0.0)
                    + pd.to_numeric(sheet_df.get('Bonus Grupo', 0), errors='coerce').fillna(0.0)
                )

                ranked_rows = sheet_df.sort_values(
                    by=['Pontuação Total da Prova', 'Atleta'],
                    ascending=[False, True],
                    kind='mergesort',
                ).reset_index(drop=True)
                ranked_positions = []
                previous_score = None
                current_rank = 1
                for position, row in ranked_rows.iterrows():
                    score = float(row.get('Pontuação Total da Prova', 0.0))
                    if previous_score is None or score != previous_score:
                        current_rank = position + 1
                    ranked_positions.append(current_rank)
                    previous_score = score

                ranked_rows['Pos Final'] = ranked_positions
                position_map = ranked_rows.set_index('Atleta')['Pos Final'].to_dict()
                sheet_df['Pos Final'] = sheet_df['Atleta'].map(position_map).fillna(0).astype(int)
                sheet_df = sheet_df.sort_values(
                    by=['Pos Final', 'Atleta'],
                    ascending=[True, True],
                    kind='mergesort',
                ).reset_index(drop=True)

                ordered_columns = [
                    'Pos Final',
                    'Atleta',
                    'Cat Round',
                    'Cat Combate',
                    'Clube',
                    'Grupo',
                    'Round 1',
                    'Round 2',
                    'Total Round',
                    'Pontos Round',
                    'Bonus Round',
                    'Nº de Vitórias Combates',
                    'Bonificação Shoot-Offs',
                    'Média dos Combates',
                    'Ranking Médias no Grupo',
                    'Bonus Média',
                    'Bonus Grupo',
                    'Pontuação Total da Prova',
                ]
                existing_columns = [col for col in ordered_columns if col in sheet_df.columns]
                extra_columns = [col for col in sheet_df.columns if col not in existing_columns]
                sheet_df = sheet_df[existing_columns + extra_columns]
                sheet_df.to_excel(writer, sheet_name=categoria, index=False)

    safe_etapa = 'resultados_final'
    try:
        buffer = io.BytesIO(template_file.getvalue())
        excel_file = pd.ExcelFile(buffer, engine='openpyxl')
        for sheet_name in excel_file.sheet_names:
            try:
                sheet_df = excel_file.parse(sheet_name)
            except Exception:
                continue
            if 'ETAPA' not in sheet_df.columns:
                continue
            etapa_values = sheet_df['ETAPA'].dropna().astype(str).str.strip()
            if not etapa_values.empty:
                safe_etapa = _sanitize_name(etapa_values.iloc[0])
                break
    except Exception:
        pass

    filename = f'{safe_etapa}_final.xlsx'
    return output.getvalue(), filename


def _build_final_csv_from_workbook(
    final_workbook_bytes: bytes,
    results_file: BinaryIO,
) -> tuple[bytes, str]:
    final_excel = pd.ExcelFile(io.BytesIO(final_workbook_bytes), engine='openpyxl')
    final_rows = []
    for sheet_name in final_excel.sheet_names:
        sheet_df = final_excel.parse(sheet_name)
        if sheet_df.empty:
            continue
        sheet_df['CATEGORIA AGRUPADA'] = sheet_name
        final_rows.append(sheet_df)

    if not final_rows:
        raise ValueError('O arquivo final não contém categorias para gerar o CSV.')

    final_df = pd.concat(final_rows, ignore_index=True)
    final_df['_pos_num'] = pd.to_numeric(final_df.get('Pos Final', 0), errors='coerce').fillna(0)
    final_df = final_df.sort_values(
        by=['CATEGORIA AGRUPADA', '_pos_num', 'Atleta'],
        ascending=[True, True, True],
        kind='mergesort',
    ).reset_index(drop=True)
    raw_df = _normalize_resultados_prova(_load_resultados_prova(results_file))

    raw_by_name = {}
    raw_by_club = {}
    for _, row in raw_df.iterrows():
        raw_info = {
            'ID': str(row.get('ID', '') or '').strip(),
            'Sigla': str(row.get('Sigla', '') or '').strip(),
            'Clube': str(row.get('Clube', '') or '').strip(),
        }
        name_key = _normalize_text(row.get('Nome Completo', ''))
        club_key = _normalize_text(row.get('Clube', ''))
        if name_key and name_key not in raw_by_name:
            raw_by_name[name_key] = raw_info
        if club_key and club_key not in raw_by_club:
            raw_by_club[club_key] = raw_info

    points_lookup = _load_fpaf_points()
    output_rows = []
    for _, row in final_df.iterrows():
        nome = str(row.get('Atleta', '') or '').strip()
        if not nome or _normalize_text(nome).startswith('bye'):
            continue

        categoria = str(row.get('CATEGORIA AGRUPADA', '') or '').strip()
        clube_final = str(row.get('Clube', '') or '').strip()
        raw_info = raw_by_name.get(_normalize_text(nome), {})
        if not raw_info:
            raw_info = raw_by_club.get(_normalize_text(clube_final), {})
        ranking = int(_coerce_numeric(row.get('Pos Final', 0)))
        category_count = int((final_df['CATEGORIA AGRUPADA'].astype(str).str.strip() == categoria).sum())
        individual_points = points_lookup.get((ranking, category_count), 0.0)
        if float(individual_points).is_integer():
            individual_points = int(individual_points)

        output_rows.append({
            'RANKING FINAL': ranking,
            'ID': raw_info.get('ID', ''),
            'NOME': nome,
            'CATEGORIA AGRUPADA': categoria,
            'SIGLA': raw_info.get('Sigla', ''),
            'CLUBE': clube_final or raw_info.get('Clube', ''),
            'PONTUAÇÃO INDIVIDUAL': individual_points,
        })

    return _build_final_csv_bytes(output_rows, getattr(results_file, 'name', 'resultados'))


def show_resultados_rr_page():
    st.title('Consolidação de Resultados')
    st.markdown('''
    Esta página possui duas etapas:
    1. **Gerar o template de resultados** a partir do workbook do Robin Round.
    2. **Gerar o arquivo final** usando o template preenchido e o arquivo bruto de resultados da prova.
    ''')

    st.subheader('1. Gerar template de resultados')
    uploaded_template_source = st.file_uploader(
        'Carregue a planilha "_combates" do Robin Round Individual (.xlsx)',
        type=['xlsx'],
        key='template_source',
    )

    if uploaded_template_source and st.button('Gerar template de resultados', key='button_template'):
        try:
            output_bytes, filename = build_resultados_template(uploaded_template_source)
            st.success('Template preparado com sucesso.')
            st.download_button(
                label='⬇️ Baixar template de resultados',
                data=output_bytes,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
        except Exception as exc:
            st.error(f'Não foi possível gerar o template: {exc}')

    st.subheader('2. Gerar arquivo final')
    uploaded_template_results = st.file_uploader(
        'Carregue o arquivo "_template resultados" preenchido (.xlsx)',
        type=['xlsx'],
        key='template_results',
    )
    uploaded_resultados_prova = st.file_uploader(
        'Carregue o arquivo de resultados do qualificatório (.txt, .csv, .xlsx)',
        type=['txt', 'csv', 'xlsx'],
        key='resultados_prova',
    )
    if uploaded_template_results and uploaded_resultados_prova and st.button('Gerar arquivo final', key='button_final'):
        try:
            output_bytes, filename = _build_final_results_workbook(
                uploaded_template_results,
                uploaded_resultados_prova,
            )
            st.success('Arquivo final preparado com sucesso.')
            st.download_button(
                label='⬇️ Baixar arquivo final',
                data=output_bytes,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

            csv_bytes, csv_filename = _build_final_csv_from_workbook(
                output_bytes,
                uploaded_resultados_prova,
            )
            st.download_button(
                label='⬇️ Baixar CSV Robin Round',
                data=csv_bytes,
                file_name=csv_filename,
                mime='text/csv',
            )
        except Exception as exc:
            st.error(f'Não foi possível gerar o arquivo final: {exc}')
