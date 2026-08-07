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
        return pd.read_csv(uploaded_file, sep=';', dtype=str, engine='python')
    return pd.read_excel(uploaded_file, dtype=str)


def _normalize_resultados_prova(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if 'Nome Completo' in df.columns:
        df['Nome Completo'] = df['Nome Completo'].fillna('').astype(str).str.strip()
    elif 'NomeCompleto' in df.columns:
        df['Nome Completo'] = df['NomeCompleto'].fillna('').astype(str).str.strip()
    else:
        family_name = df.get('FamilyName', '').fillna('').astype(str).str.strip()
        given_name = df.get('GivenName', '').fillna('').astype(str).str.strip()
        df['Nome Completo'] = (family_name + ' ' + given_name).str.strip()

    if 'Clube' in df.columns:
        df['Clube'] = df['Clube'].fillna('').astype(str).str.strip()
    elif 'Country' in df.columns:
        df['Clube'] = df['Country'].fillna('').astype(str).str.strip()
    else:
        df['Clube'] = ''

    for score_col in ['D1 Score', 'D1', 'Round 1']:
        if score_col in df.columns:
            df['Round 1'] = pd.to_numeric(
                df[score_col].astype(str).str.replace(',', '.', regex=False).str.strip(),
                errors='coerce'
            ).fillna(0)
            break
    else:
        df['Round 1'] = 0

    for score_col in ['D2 Score', 'D2', 'Round 2']:
        if score_col in df.columns:
            df['Round 2'] = pd.to_numeric(
                df[score_col].astype(str).str.replace(',', '.', regex=False).str.strip(),
                errors='coerce'
            ).fillna(0)
            break
    else:
        df['Round 2'] = 0

    if 'Categoria Quali' in df.columns:
        df['Cat Round'] = df['Categoria Quali'].fillna('').astype(str).str.strip()
    elif 'Categoria' in df.columns:
        df['Cat Round'] = df['Categoria'].fillna('').astype(str).str.strip()
    elif 'Division' in df.columns and 'Class' in df.columns:
        df['Cat Round'] = (
            df['Division'].fillna('').astype(str).str.strip()
            + df['Class'].fillna('').astype(str).str.strip()
        ).str.strip()
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

                ranking_points = {row['Atleta']: 0.0 for row in unique_rows}
                for idx, athlete_row in sheet_df.iterrows():
                    ranking_points[athlete_row['Atleta']] = _coerce_numeric(athlete_row['Média dos Combates'])
                sorted_names = sorted(ranking_points, key=ranking_points.get, reverse=True)
                ranking_map = {}
                for position, athlete_name in enumerate(sorted_names):
                    ranking_map[athlete_name] = float(['2.0', '1.5', '1.0', '0.5'][position]) if position < 4 else 0.0

                for idx, athlete_row in sheet_df.iterrows():
                    sheet_df.at[idx, 'Ranking Médias no Grupo'] = ranking_map.get(athlete_row['Atleta'], 0.0)

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

                ordered_columns = [
                    'Pos Final',
                    'Atleta',
                    'Cat Round',
                    'Cat Combate',
                    'Clube',
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
        'Carregue o arquivo de resultados da prova (.txt, .csv, .xlsx)',
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
        except Exception as exc:
            st.error(f'Não foi possível gerar o arquivo final: {exc}')
