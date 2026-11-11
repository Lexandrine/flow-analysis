import os
import sys
import glob
import re
import datetime
import argparse
import json
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib
matplotlib.use('Agg') 
import matplotlib.pyplot as plt
import requests
from bs4 import BeautifulSoup
import urllib.parse
import holidays
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial', 'Calibri', 'Segoe UI', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

ru_holidays = holidays.country_holidays('RU', years=range(2020, 2030))

CLUSTER_NAMES = {
    0: 'Будни (норма)',
    1: 'Выходные/Праздники',
    2: 'Аномальный спад',
    3: 'Аномальный всплеск',
    4: 'Технический сбой'
}


# ==========================================
# ПОИСК НОВОСТЕЙ (с fallback на DuckDuckGo)
# ==========================================
def search_yandex_context(date_str, road_name="подъезд к Воронежу"):
    """
    Пытается получить контекст из Яндекса, затем из DuckDuckGo.
    Возвращает dict с полями:
      - 'url': ссылка на поиск
      - 'text': найденный текст (или None)
      - 'source': откуда взяли ('yandex', 'duckduckgo', 'none')
    """
    query = f"Воронеж {date_str} {road_name}"
    encoded_query = urllib.parse.quote(query)
    yandex_url = f"https://yandex.ru/search/?text={encoded_query}&lr=213"
    
    result = {
        'url': yandex_url,
        'text': None,
        'source': 'none'
    }
    
    # --- Попытка 1: Яндекс ---
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "ru-RU,ru;q=0.9"
    }
    try:
        response = requests.get(yandex_url, headers=headers, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        snippets = soup.find_all('div', class_='serp-item__text')
        if not snippets:
            snippets = soup.find_all('div', class_='OrganicTitle-LinkText')
        
        causes = []
        for item in snippets[:3]:
            text = item.get_text(separator=' ', strip=True)
            if len(text) > 25:
                causes.append(text)
        
        if causes:
            result['text'] = " | ".join(causes)
            result['source'] = 'yandex'
            return result
    except Exception:
        pass
    
    # --- Попытка 2: DuckDuckGo HTML (менее агрессивен к ботам) ---
    try:
        ddg_url = f"https://html.duckduckgo.com/html/?q={encoded_query}"
        response = requests.get(ddg_url, headers=headers, timeout=5)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        snippets = soup.find_all('a', class_='result__snippet')
        causes = []
        for item in snippets[:3]:
            text = item.get_text(separator=' ', strip=True)
            if len(text) > 20:
                causes.append(text)
        
        if causes:
            result['text'] = " | ".join(causes)
            result['source'] = 'duckduckgo'
            return result
    except Exception:
        pass
    
    return result


# ==========================================
# ПАРСЕР
# ==========================================
def parse_traffic_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    
    records = []
    current_section = "Неизвестный участок"
    date_regex = re.compile(r"^\s*\d{1,2}[\./]\d{1,2}[\./]\d{4}\s+\d{1,2}:\d{2}(:\d{2})?\s*$")
    section_regex = re.compile(r"км\s*\d+(\+\d+)?", re.IGNORECASE)
    
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
            
        for cell in row[:6]:
            if cell is not None and isinstance(cell, str):
                cell_s = cell.strip()
                if section_regex.search(cell_s):
                    current_section = cell_s
                    break
                elif ("а/д" in cell_s.lower() or "а-134" in cell_s.lower() or "подъезд" in cell_s.lower()) and "отчет" not in cell_s.lower():
                    current_section = cell_s
                    break
                    
        first_val = row[0]
        if first_val is None:
            continue
            
        first_str = str(first_val).strip()
        if first_str.lower() in ['итого', 'среднее', '%', 'дата']:
            continue
            
        dt_obj = None
        if isinstance(first_val, (datetime.datetime, datetime.date)):
            dt_obj = pd.to_datetime(first_val)
        elif date_regex.match(first_str):
            try:
                dt_obj = pd.to_datetime(first_str, dayfirst=True)
            except Exception:
                dt_obj = None
                
        if dt_obj is not None and pd.notnull(dt_obj):
            row_vals = list(row[:29])
            if len(row_vals) < 29:
                row_vals.extend([0] * (29 - len(row_vals)))
            records.append([current_section, dt_obj] + row_vals[1:29])
            
    wb.close()
    
    col_names = [
        'road_section', 'timestamp',
        'vol_total', 'vol_direct', 'vol_reverse',
        'cars_total', 'cars_direct', 'cars_reverse',
        'small_trucks_total', 'small_trucks_direct', 'small_trucks_reverse',
        'med_trucks_total', 'med_trucks_direct', 'med_trucks_reverse',
        'large_trucks_total', 'large_trucks_direct', 'large_trucks_reverse',
        'road_trains_total', 'road_trains_direct', 'road_trains_reverse',
        'buses_total', 'buses_direct', 'buses_reverse',
        'motos_total', 'motos_direct', 'motos_reverse',
        'speed_direct', 'speed_reverse',
        'load_direct', 'load_reverse'
    ]
    
    if not records:
        return pd.DataFrame(columns=col_names)
        
    df = pd.DataFrame(records, columns=col_names)
    
    for col in col_names[2:]:
        if df[col].dtype == object:
            df[col] = (df[col].astype(str)
                       .str.replace('\xa0', '', regex=False)
                       .str.replace(' ', '', regex=False)
                       .str.replace(',', '.', regex=False))
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
    return df


# ==========================================
# ДЕТЕКТОРЫ АНОМАЛИЙ
# ==========================================
def detect_daily_anomalies(df):
    if df.empty:
        return pd.DataFrame()
        
    df = df.copy()
    df['date'] = df['timestamp'].dt.date
    df['hour'] = df['timestamp'].dt.hour
    df['dayofweek'] = df['timestamp'].dt.dayofweek
    df['is_weekend'] = df['dayofweek'].isin([5, 6]).astype(int)
    df['is_holiday'] = df['date'].apply(lambda x: x in ru_holidays).astype(int)
    df['is_weekend_or_holiday'] = ((df['is_weekend'] == 1) | (df['is_holiday'] == 1)).astype(int)
    df['month_key'] = df['timestamp'].dt.to_period('M').astype(str)
    
    anomalies_list = []
    
    for section, sec_df in df.groupby('road_section'):
        
        normal_car_share = (sec_df['cars_total'] / (sec_df['vol_total'] + 1)).median()
        
        for dt_date, day_df in sec_df.groupby('date'):
            day_df = day_df.sort_values('hour')
            hours_cnt = len(day_df)
            
            holiday_name = ru_holidays.get(dt_date)
            holiday_tag = f" (Праздник: {holiday_name})" if holiday_name else ""
            
            # 1. Неполные сутки
            if hours_cnt < 24:
                anomalies_list.append({
                    'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Критический',
                    'Детектор': 'Целостность', 'Тип аномалии': 'Неполные сутки',
                    'Параметр': 'Часов в сутках', 'Факт': hours_cnt, 'Норма': 24,
                    'Отклонение (%)': round((hours_cnt - 24) / 24 * 100, 1), 'Индекс (Score)': 3.0,
                    'Описание цифрами': f"В сутках сохранено лишь {hours_cnt} часов вместо 24{holiday_tag}"
                })
                
            # 2. Залипание датчика
            vols = day_df['vol_total'].values
            max_consec, cur_consec, flat_val = 1, 1, None
            for i in range(1, len(vols)):
                if vols[i] == vols[i-1] and vols[i] > 10:
                    cur_consec += 1
                    if cur_consec > max_consec:
                        max_consec, flat_val = cur_consec, vols[i]
                else:
                    cur_consec = 1
            if max_consec >= 3:
                anomalies_list.append({
                    'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Критический',
                    'Детектор': 'Датчик', 'Тип аномалии': 'Залипание датчика (Flatline)',
                    'Параметр': 'Интенсивность', 'Факт': flat_val, 'Норма': 'Вариативность',
                    'Отклонение (%)': 0.0, 'Индекс (Score)': 4.0,
                    'Описание цифрами': f"Значение не менялось {max_consec} часов подряд ({flat_val} авт./ч){holiday_tag}"
                })

            # 3. Дисбаланс направлений
            day_df['dir_ratio'] = day_df['vol_direct'] / (day_df['vol_total'] + 1)
            imbalance_hours = day_df[(day_df['dir_ratio'] > 0.85) | (day_df['dir_ratio'] < 0.15)]
            if len(imbalance_hours) >= 2:
                anomalies_list.append({
                    'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Высокий',
                    'Детектор': 'Направление', 'Тип аномалии': 'Критический дисбаланс направлений',
                    'Параметр': 'Доля прямого напр.', 'Факт': f"{round(imbalance_hours['dir_ratio'].max() * 100, 1)}%", 
                    'Норма': '40-60%', 'Отклонение (%)': 0.0, 'Индекс (Score)': 3.5,
                    'Описание цифрами': f"В течение {len(imbalance_hours)} часов поток был почти односторонним{holiday_tag}"
                })

            # 4. Сдвиг в классе ТС
            day_car_share = (day_df['cars_total'] / (day_df['vol_total'] + 1)).mean()
            if abs(day_car_share - normal_car_share) > 0.15:
                anomalies_list.append({
                    'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Высокий',
                    'Детектор': 'Состав потока', 'Тип аномалии': 'Аномальная доля легковых авто',
                    'Параметр': 'Доля легковых', 'Факт': f"{round(day_car_share * 100, 1)}%", 
                    'Норма': f"{round(normal_car_share * 100, 1)}%", 'Отклонение (%)': round((day_car_share - normal_car_share)*100, 1), 
                    'Индекс (Score)': 3.0,
                    'Описание цифрами': f"Доля легковых авто отклонилась от нормы на {abs(round((day_car_share - normal_car_share)*100, 1))} п.п.{holiday_tag}"
                })

            # 5. Экстремальный затор
            daytime_df = day_df[(day_df['hour'] >= 7) & (day_df['hour'] <= 21)]
            if len(daytime_df) > 0:
                min_spd = daytime_df['speed_direct'].min()
                max_load = daytime_df['load_direct'].max()
                if min_spd < 65 and max_load > 15:
                    jam_h = daytime_df.loc[daytime_df['speed_direct'] == min_spd, 'hour'].iloc[0]
                    anomalies_list.append({
                        'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Критический',
                        'Детектор': 'Скорость/Загрузка', 'Тип аномалии': 'Экстремальный затор или сбой датчика',
                        'Параметр': 'Мин. скорость (Прямое)', 'Факт': round(min_spd, 1), 'Норма': '> 70 км/ч',
                        'Отклонение (%)': round((min_spd - 70) / 70 * 100, 1), 'Индекс (Score)': 4.5,
                        'Описание цифрами': f"В {int(jam_h)}:00 скорость упала до {min_spd} км/ч при загрузке {max_load}%{holiday_tag}"
                    })

        # 6. Статистические аномалии (Z-score)
        for (m_key, is_wknd_or_hol), group_df in sec_df.groupby(['month_key', 'is_weekend_or_holiday']):
            daily_stats = group_df.groupby('date').agg(daily_vol=('vol_total', 'sum')).reset_index()
            if len(daily_stats) < 3: continue
                
            vol_med = daily_stats['daily_vol'].median()
            vol_mad = (daily_stats['daily_vol'] - vol_med).abs().median()
            vol_mad = vol_mad if (vol_mad > 0 and pd.notnull(vol_mad)) else daily_stats['daily_vol'].std()
            vol_mad = vol_mad if vol_mad > 0 else 1.0
            
            for _, r in daily_stats.iterrows():
                dt_date = r['date']
                d_vol = r['daily_vol']
                z_vol = 0.6745 * (d_vol - vol_med) / vol_mad if vol_mad > 0 else 0
                delta_p = round((d_vol - vol_med) / vol_med * 100, 1) if vol_med > 0 else 0
                
                if (z_vol < -2.5 and delta_p <= -20.0) or delta_p <= -40.0:
                    anomalies_list.append({
                        'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Критический',
                        'Детектор': 'Амплитуда', 'Тип аномалии': 'Аномальный спад интенсивности',
                        'Параметр': 'Суточный объем', 'Факт': int(d_vol), 'Норма': int(vol_med),
                        'Отклонение (%)': delta_p, 'Индекс (Score)': round(abs(z_vol), 1),
                        'Описание цифрами': f"Суточный трафик {int(d_vol)} ниже нормы {int(vol_med)} на {abs(delta_p)}% (Z={z_vol:.1f}){holiday_tag}"
                    })
                elif (z_vol > 2.5 and delta_p >= 20.0) or delta_p >= 40.0:
                    anomalies_list.append({
                        'Участок дороги': section, 'Дата': dt_date, 'Уровень': 'Высокий',
                        'Детектор': 'Амплитуда', 'Тип аномалии': 'Аномальный всплеск интенсивности',
                        'Параметр': 'Суточный объем', 'Факт': int(d_vol), 'Норма': int(vol_med),
                        'Отклонение (%)': delta_p, 'Индекс (Score)': round(abs(z_vol), 1),
                        'Описание цифрами': f"Суточный трафик {int(d_vol)} выше нормы {int(vol_med)} на +{delta_p}% (Z={z_vol:.1f}){holiday_tag}"
                    })

    res_df = pd.DataFrame(anomalies_list)
    if not res_df.empty:
        res_df = res_df.sort_values(by=['Участок дороги', 'Дата'])
    return res_df


def detect_monthly_anomalies(all_data_df):
    if all_data_df.empty:
        return pd.DataFrame()
        
    df = all_data_df.copy()
    df['year'] = df['timestamp'].dt.year
    df['month'] = df['timestamp'].dt.month
    df['year_month'] = df['timestamp'].dt.to_period('M').astype(str)
    
    monthly_stats = []
    for (sec, ym), m_df in df.groupby(['road_section', 'year_month']):
        y, m = m_df['year'].iloc[0], m_df['month'].iloc[0]
        tot_hours = len(m_df)
        exp_hours = pd.Period(ym).days_in_month * 24
        completeness = (tot_hours / exp_hours) * 100
        
        monthly_stats.append({
            'road_section': sec, 'year_month': ym, 'year': y, 'month': m,
            'completeness_pct': round(completeness, 1)
        })
        
    m_df = pd.DataFrame(monthly_stats)
    monthly_anomalies = []
    
    for (sec, cal_month), group in m_df.groupby(['road_section', 'month']):
        if len(group) < 2: continue
        
        for _, row in group.iterrows():
            if row['completeness_pct'] < 90.0:
                monthly_anomalies.append({
                    'Участок дороги': sec, 'Период': row['year_month'], 'Уровень': 'Критический',
                    'Детектор': 'Полнота данных', 'Параметр': 'Отработано часов (%)',
                    'Факт': f"{row['completeness_pct']}%", 'Норма': '>= 95%',
                    'Отклонение (%)': round(row['completeness_pct'] - 100, 1),
                    'Описание цифрами': f"Данные за месяц неполные (сохранено лишь {row['completeness_pct']}% часов)"
                })
                
    res_df = pd.DataFrame(monthly_anomalies)
    return res_df.sort_values(by=['Участок дороги', 'Период']) if not res_df.empty else res_df


# ==========================================
# ЭКСПОРТ В EXCEL С ГИПЕРССЫЛКАМИ
# ==========================================
def export_anomalies_to_excel(daily_df, monthly_df, output_path, news_contexts=None):
    """
    news_contexts: dict {date_str: {'url': ..., 'text': ..., 'source': ...}}
    """
    if news_contexts is None:
        news_contexts = {}
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if not daily_df.empty:
            daily_df.to_excel(writer, sheet_name="Аномальные сутки", index=False)
        else:
            pd.DataFrame([{"Инфо": "Аномальные сутки не обнаружены"}]).to_excel(writer, sheet_name="Аномальные сутки", index=False)
            
        if not monthly_df.empty:
            monthly_df.to_excel(writer, sheet_name="Аномальные месяцы (YoY)", index=False)
        else:
            pd.DataFrame([{"Инфо": "Аномальные месяцы не обнаружены"}]).to_excel(writer, sheet_name="Аномальные месяцы (YoY)", index=False)
            
    wb = openpyxl.load_workbook(output_path)
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    crit_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    high_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    for ws in wb.worksheets:
        # Форматирование заголовков
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        # Добавляем колонку "Ссылка на новости" если её нет
        if ws.title == "Аномальные сутки" and ws.max_row > 1:
            # Ищем колонку с датой (обычно колонка B = 2)
            date_col_idx = None
            for col_idx in range(1, ws.max_column + 1):
                if ws.cell(row=1, column=col_idx).value == 'Дата':
                    date_col_idx = col_idx
                    break
            
            if date_col_idx:
                # Добавляем новую колонку после последней
                link_col_idx = ws.max_column + 1
                ws.cell(row=1, column=link_col_idx).value = "Поиск в Яндексе"
                ws.cell(row=1, column=link_col_idx).fill = header_fill
                ws.cell(row=1, column=link_col_idx).font = header_font
                ws.cell(row=1, column=link_col_idx).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                
                for row_idx in range(2, ws.max_row + 1):
                    date_val = ws.cell(row=row_idx, column=date_col_idx).value
                    if date_val:
                        # Преобразуем дату в строку для поиска в news_contexts
                        if isinstance(date_val, datetime.datetime):
                            date_str = date_val.strftime("%d.%m.%Y")
                        elif isinstance(date_val, datetime.date):
                            date_str = date_val.strftime("%d.%m.%Y")
                        else:
                            date_str = str(date_val)
                        
                        ctx = news_contexts.get(date_str, {})
                        url = ctx.get('url', '')
                        source = ctx.get('source', 'none')
                        
                        cell = ws.cell(row=row_idx, column=link_col_idx)
                        if url:
                            cell.hyperlink = url
                            cell.value = "Открыть поиск" if source != 'none' else "Проверить вручную"
                            cell.font = link_font
                        else:
                            cell.value = "—"
                    
                    # Форматирование строк по уровню
                    level_val = None
                    for col_idx in range(1, ws.max_column + 1):
                        c = ws.cell(row=row_idx, column=col_idx)
                        c.border = border_thin
                        c.alignment = Alignment(vertical="center")
                        if ws.cell(row=1, column=col_idx).value == 'Уровень':
                            level_val = c.value
                            
                    if level_val == 'Критический':
                        for col_idx in range(1, ws.max_column + 1):
                            ws.cell(row=row_idx, column=col_idx).fill = crit_fill
                    elif level_val == 'Высокий':
                        for col_idx in range(1, ws.max_column + 1):
                            ws.cell(row=row_idx, column=col_idx).fill = high_fill
        
        # Автоширина колонок
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)
            
        ws.row_dimensions[1].height = 28
        
    wb.save(output_path)


# ==========================================
# СОХРАНЕНИЕ КОНТЕКСТА В JSON
# ==========================================
def save_news_contexts_to_json(news_contexts, output_path):
    """Сохраняет полный контекст новостей в JSON."""
    # Преобразуем для JSON-сериализации
    json_data = []
    for date_str, ctx in sorted(news_contexts.items()):
        json_data.append({
            'date': date_str,
            'search_url': ctx.get('url', ''),
            'source': ctx.get('source', 'none'),
            'found_text': ctx.get('text', None),
            'has_context': ctx.get('text') is not None
        })
    
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)
    
    print(f"  Контекст новостей сохранён: {output_path}")


# ==========================================
# ГРАФИКИ
# ==========================================
def generate_anomaly_plots(df, daily_anomalies_df, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    if daily_anomalies_df.empty or df.empty: return
        
    df = df.copy()
    df['date'] = df['timestamp'].dt.date
    df['hour'] = df['timestamp'].dt.hour
    df['is_weekend'] = df['timestamp'].dt.weekday.isin([5, 6]).astype(int)
    df['month_key'] = df['timestamp'].dt.to_period('M').astype(str)
    
    grouped = daily_anomalies_df.groupby(['Участок дороги', 'Дата'])
    print(f"\nГенерация графиков (уникальных аномальных дней: {len(grouped)})...")
    
    hours = np.arange(24)
    for (sec, dt_d), rows in grouped:
        day_df = df[(df['road_section'] == sec) & (df['date'] == dt_d)].sort_values('hour')
        if len(day_df) == 0: continue
            
        m_key = day_df['month_key'].iloc[0]
        is_wknd = day_df['is_weekend'].iloc[0]
        anom_types = "; ".join(rows['Тип аномалии'].unique())
        
        ref_df = df[(df['road_section'] == sec) & (df['month_key'] == m_key) & (df['is_weekend'] == is_wknd)]
        norm_vol = ref_df.groupby('hour')['vol_total'].median().reindex(hours).fillna(0).values
        norm_spd = ref_df.groupby('hour')['speed_direct'].median().reindex(hours).fillna(0).values
        
        fact_vol = day_df.set_index('hour')['vol_total'].reindex(hours).fillna(0).values
        fact_spd = day_df.set_index('hour')['speed_direct'].reindex(hours).fillna(0).values
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 7), sharex=True, gridspec_kw={'height_ratios': [2, 1]})
        
        ax1.bar(hours, norm_vol, color='#BDC3C7', alpha=0.6, width=0.6, label='Норма месяца', edgecolor='#95A5A6')
        ax1.plot(hours, fact_vol, color='#E74C3C', linewidth=2.5, marker='o', markersize=5, label='Факт')
        ax1.set_ylabel('Интенсивность, авт./ч', fontsize=10, fontweight='bold')
        ax1.set_title(f"{sec} | {dt_d}\nСобытия: {anom_types}", fontsize=11, fontweight='bold', pad=10)
        ax1.grid(True, linestyle='--', alpha=0.5)
        ax1.legend(loc='upper left')
        
        ax2.plot(hours, norm_spd, color='#7F8C8D', linestyle='--', linewidth=1.8, label='Норма скорости')
        ax2.plot(hours, fact_spd, color='#2980B9', linewidth=2.0, marker='s', markersize=4, label='Факт скорости')
        ax2.axhline(65, color='#C0392B', linestyle=':', alpha=0.7, label='Порог затора (65 км/ч)')
        ax2.set_ylabel('Скорость, км/ч', fontsize=10, fontweight='bold')
        ax2.set_xlabel('Час суток (00:00 - 23:00)', fontsize=10, fontweight='bold')
        ax2.set_xticks(hours)
        ax2.set_xticklabels([f"{h:02d}" for h in hours])
        ax2.grid(True, linestyle='--', alpha=0.5)
        ax2.legend(loc='lower left')
        ax2.set_ylim(0, max(110, np.max(fact_spd) + 10))
        
        plt.tight_layout()
        safe_sec = re.sub(r'[^a-zA-Z0-9а-яА-Я+]', '_', sec)[:20]
        fname = f"{dt_d}_{safe_sec}.png"
        fig.savefig(os.path.join(output_folder, fname), dpi=110)
        plt.close(fig)
        print(f"  Сохранен: {fname}")


# ==========================================
# ГЛАВНАЯ ФУНКЦИЯ
# ==========================================
def main():
    parser = argparse.ArgumentParser(description="Анализатор аномалий трафика А-134 (v6 с JSON)")
    parser.add_argument("--input", default="./input_data", help="Папка с входными файлами .xlsx")
    parser.add_argument("--output", default="./reports", help="Папка для сохранения отчетов и графиков")
    args = parser.parse_args()
    
    input_folder = args.input
    output_folder = args.output
    plots_folder = os.path.join(output_folder, "plots")
    
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)
    
    excel_files = sorted(glob.glob(os.path.join(input_folder, "*.xlsx")) + glob.glob(os.path.join(input_folder, "*.xls")))
    
    if not excel_files:
        print(f"\n[Внимание] В папке '{input_folder}' нет файлов .xlsx!")
        return
        
    print(f"\n=== СТАРТ ОБРАБОТКИ ДАННЫХ ===")
    print(f"Найдено файлов: {len(excel_files)}")
    
    all_dfs = []
    for idx, fpath in enumerate(excel_files, 1):
        fname = os.path.basename(fpath)
        print(f"[{idx}/{len(excel_files)}] Чтение {fname}...", end=" ", flush=True)
        try:
            df_part = parse_traffic_excel(fpath)
            print(f"OK ({len(df_part)} строк)")
            all_dfs.append(df_part)
        except Exception as e:
            print(f"ОШИБКА: {e}")
            
    if not all_dfs:
        print("Не удалось прочитать ни одного файла.")
        return
        
    full_df = pd.concat(all_dfs, ignore_index=True)
    print(f"\nВсего загружено: {len(full_df)} записей.")
    
    print("\n1. Поиск аномальных суток...")
    daily_anomalies = detect_daily_anomalies(full_df)
    print(f"   Найдено аномалий: {len(daily_anomalies)}")
    
    print("\n2. Поиск аномальных месяцев (YoY)...")
    monthly_anomalies = detect_monthly_anomalies(full_df)
    print(f"   Найдено аномальных месяцев: {len(monthly_anomalies)}")
    
    # 3. Получаем контекст новостей для каждой уникальной даты аномалии
    print("\n3. Получение контекста новостей (Яндекс + DuckDuckGo)...")
    unique_dates = daily_anomalies['Дата'].unique() if not daily_anomalies.empty else []
    news_contexts = {}
    
    for i, dt_date in enumerate(unique_dates, 1):
        date_str = dt_date.strftime("%d.%m.%Y") if hasattr(dt_date, 'strftime') else str(dt_date)
        print(f"   [{i}/{len(unique_dates)}] {date_str}...", end=" ", flush=True)
        
        ctx = search_yandex_context(date_str, "подъезд к Воронежу")
        news_contexts[date_str] = ctx
        
        if ctx['source'] == 'yandex':
            print("✓ Яндекс")
        elif ctx['source'] == 'duckduckgo':
            print("✓ DuckDuckGo")
        else:
            print("✗ нет результата")
        
        # Пауза чтобы не забанили
        import time
        time.sleep(1)
    
    # 4. Сохраняем Excel с гиперссылками
    report_xlsx_path = os.path.join(output_folder, "traffic_anomalies_report.xlsx")
    export_anomalies_to_excel(daily_anomalies, monthly_anomalies, report_xlsx_path, news_contexts)
    print(f"\n[Успешно] Реестр сохранен: {report_xlsx_path}")
    
    # 5. Сохраняем JSON с полным контекстом
    json_path = os.path.join(output_folder, "news_contexts.json")
    save_news_contexts_to_json(news_contexts, json_path)
    
    # 6. Графики
    generate_anomaly_plots(full_df, daily_anomalies, plots_folder)
    print(f"[Успешно] Графики сохранены: {plots_folder}")
    
    # Статистика по источникам
    sources = [ctx['source'] for ctx in news_contexts.values()]
    print(f"\n=== СТАТИСТИКА ПОИСКА ===")
    print(f"  Яндекс: {sources.count('yandex')}")
    print(f"  DuckDuckGo: {sources.count('duckduckgo')}")
    print(f"  Без результата: {sources.count('none')}")
    print(f"\n=== АНАЛИЗ ПОЛНОСТЬЮ ЗАВЕРШЕН ===")


if __name__ == "__main__":
    main()