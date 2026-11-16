"""
Полный анализатор аномалий трафика А-134 (исправленная версия).
Включает все детекторы из репозитория traffic-analytics.
"""

import os
import sys
import glob
import re
import datetime
import json
import argparse
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import holidays
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

# КОНСТАНТЫ

ROAD = "a134"
ROAD_NAME = "А-134, подъезд к г. Воронеж"
SENSOR_ID = "a134-km8-100"
SENSOR_NAME = "км 8+100 а/д А-134, подъезд к г. Воронеж"

RU_HOLIDAYS = holidays.country_holidays('RU', years=range(2020, 2030))

# Пороги детекторов
HOURLY_Z = 3.0
HOURLY_MIN_RUN = 4
PRACTICAL_HOUR_PCT, PRACTICAL_HOUR_ABS = 35.0, 50.0
PRACTICAL_DAILY_PCT, PRACTICAL_DAILY_ABS = 25.0, 100.0
DROUGHT_Z, SURGE_Z = -3.0, 3.0
SPEED_DROP_KMH = 15.0
SPEED_DROP_Z = 3.0
SPEED_DROP_MIN_RUN = 4
CONGESTION_SPEED = 50.0
CONGESTION_LOAD = 20.0
CONGESTION_MIN_HOURS = 2
FREE_FLOW_SPEED = 80.0
PROFILE_CORR = 0.70
PROFILE_MIN_DAY_VOLUME = 1000
MONTH_COMPLETENESS = 75.0
MONTH_VALIDITY = 95.0
YOY_MADT_PCT = 25.0
YOY_MIX_PP = 8.0
YOY_SPEED = 15.0
LOAD_DEV_PCT = 20.0
MEGA_PCT = 60.0
REGIME_MIN_DAYS = 7
REGIME_LOOKBACK_DAYS = 168
REGIME_MIN_REFS = 6
REGIME_RATIO_LO, REGIME_RATIO_HI = 0.75, 1.25
PROMOTE_STEP, PROMOTE_CAP = 0.5, 6.5

# Floors для floored robust scale
FLOOR_VOLUME_HOUR = 20.0
FLOOR_VOLUME_DAY = 100.0
FLOOR_SPEED = 3.0
FLOOR_LOAD = 2.0
MAD_TO_SIGMA = 1.4826
RELATIVE_FLOOR = 0.10

# Категории событий
EVENT_KINDS = {
    "accident": {"volume_drop", "congestion"},
    "roadworks": {"volume_drop", "volume_spike", "profile_shift", "speed_drop", "congestion"},
    "weather": {"volume_drop", "congestion", "speed_drop", "profile_shift"},
    "holiday": {"volume_drop", "volume_spike", "profile_shift"},
    "outage": {"missing_hours", "flatline", "zero_speed", "zero_telemetry"},
    "other": {"volume_drop", "volume_spike", "congestion", "profile_shift", "speed_drop"},
}

CLUSTER_NAMES = {
    0: 'Будни (норма)',
    1: 'Выходные/Праздники',
    2: 'Аномальный спад',
    3: 'Аномальный всплеск',
    4: 'Технический сбой'
}


# УТИЛИТЫ

def day_kind(ts):
    """Определяет тип дня."""
    if hasattr(ts, 'date'):
        ts_date = ts.date()
    else:
        ts_date = ts
    if ts_date in RU_HOLIDAYS:
        return "holiday"
    if ts.weekday() >= 5:
        return "weekend"
    return "weekday"

def level_from_score(score):
    """Таблица severity."""
    if score is None:
        return "low"
    if score >= 5.0:
        return "critical"
    if score >= 3.5:
        return "high"
    if score >= 2.0:
        return "medium"
    return "low"

def _center_and_mad(values):
    """Медиана и MAD."""
    median = float(values.median())
    if np.isnan(median):
        return np.nan, 0.0
    mad = float((values - median).abs().median())
    if not (mad > 0) or np.isnan(mad):
        mad = 0.0
    return median, mad

def _sigma(median, mad, floor):
    """Floored robust scale."""
    return max(MAD_TO_SIGMA * mad, abs(median) * RELATIVE_FLOOR, floor)

def _floor_for(metric):
    if metric.startswith("speed"):
        return FLOOR_SPEED
    if metric.startswith("load"):
        return FLOOR_LOAD
    return FLOOR_VOLUME_HOUR

def make_anomaly(sensor_id, starts_at, ends_at, category, type_, detector, level,
                 metric, actual, expected, deviation_pct, score, description):
    """Создает аномалию."""
    return {
        "id": "",
        "sensor_id": sensor_id,
        "starts_at": starts_at.isoformat() if isinstance(starts_at, (datetime.datetime, pd.Timestamp)) else starts_at,
        "ends_at": ends_at.isoformat() if isinstance(ends_at, (datetime.datetime, pd.Timestamp)) else ends_at,
        "category": category,
        "type": type_,
        "detector": detector,
        "level": level,
        "metric": metric,
        "actual": None if actual is None else round(float(actual), 1),
        "expected": None if expected is None else round(float(expected), 1),
        "deviation_pct": None if deviation_pct is None else round(float(deviation_pct), 1),
        "score": None if score is None else max(0.0, round(float(score), 1)),
        "is_mega": False,
        "event_id": None,
        "description": description,
    }


# 1. ПАРСИНГ ДАННЫХ

def parse_traffic_excel(filepath):
    """Парсит Excel-файл."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    
    records = []
    current_section = SENSOR_NAME
    date_regex = re.compile(r"^\s*\d{1,2}[\./]\d{1,2}[\./]\d{4}\s+\d{1,2}:\d{2}(:\d{2})?\s*$")
    
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        
        for cell in row[:6]:
            if cell is not None and isinstance(cell, str):
                cell_s = cell.strip()
                if "а-134" in cell_s.lower() or "подъезд" in cell_s.lower():
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
                pass
        
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
    
    df['timestamp'] = pd.to_datetime(df['timestamp'])
    mask_2023 = df['timestamp'].dt.year == 2023
    df.loc[mask_2023, 'timestamp'] = df.loc[mask_2023, 'timestamp'] + pd.Timedelta(hours=3)
    
    df = df.drop_duplicates(subset=['road_section', 'timestamp'], keep='first')
    
    return df


# 2. ТЕХНИЧЕСКИЕ СБОИ

def detect_failures(df):
    """Находит технические сбои."""
    failures = []
    bad_hours = set()
    
    for section, sec_df in df.groupby('road_section'):
        sec_df = sec_df.sort_values('timestamp').reset_index(drop=True)
        
        # Flatline
        vols = sec_df['vol_total'].values
        max_consec = 1
        cur_consec = 1
        flat_start = None
        
        for i in range(1, len(vols)):
            if vols[i] == vols[i-1] and vols[i] > 10:
                cur_consec += 1
            else:
                if cur_consec >= 5:
                    start_idx = i - cur_consec
                    start_ts = sec_df.iloc[start_idx]['timestamp']
                    end_ts = sec_df.iloc[i-1]['timestamp']
                    failures.append(make_anomaly(
                        SENSOR_ID, start_ts, end_ts + pd.Timedelta(hours=1),
                        "technical", "flatline", "sensor", "critical",
                        "vol_total", float(vols[start_idx]), None, 0.0, 4.0,
                        f"Значение не менялось {cur_consec} часов подряд ({int(vols[start_idx])} авт./ч)"
                    ))
                    for j in range(start_idx, i):
                        bad_hours.add((section, sec_df.iloc[j]['timestamp']))
                cur_consec = 1
        
        # Zero speed
        zero_speed_mask = (sec_df['speed_direct'] <= 1) & (sec_df['vol_direct'] >= 5)
        zero_speed_indices = sec_df[zero_speed_mask].index.tolist()
        
        if len(zero_speed_indices) >= 2:
            consecutive = []
            current_run = [zero_speed_indices[0]]
            
            for idx in zero_speed_indices[1:]:
                if idx == current_run[-1] + 1:
                    current_run.append(idx)
                else:
                    if len(current_run) >= 2:
                        consecutive.append(current_run)
                    current_run = [idx]
            
            if len(current_run) >= 2:
                consecutive.append(current_run)
            
            for run in consecutive:
                start_ts = sec_df.iloc[run[0]]['timestamp']
                end_ts = sec_df.iloc[run[-1]]['timestamp']
                failures.append(make_anomaly(
                    SENSOR_ID, start_ts, end_ts + pd.Timedelta(hours=1),
                    "technical", "zero_speed", "sensor", "high",
                    "speed_direct", 0.0, None, -100.0, 2.5,
                    f"Скорость ~0 км/ч при потоке в течение {len(run)} часов"
                ))
                for idx in run:
                    bad_hours.add((section, sec_df.iloc[idx]['timestamp']))
        
        # Direction imbalance
        sec_df_copy = sec_df.copy()
        sec_df_copy['dir_ratio'] = sec_df_copy['vol_direct'] / (sec_df_copy['vol_total'] + 1)
        imbalance_mask = (
            ((sec_df_copy['dir_ratio'] <= 0.05) | (sec_df_copy['dir_ratio'] >= 0.95)) &
            (sec_df_copy['vol_total'] > 200)
        )
        imbalance_indices = sec_df_copy[imbalance_mask].index.tolist()
        
        if len(imbalance_indices) >= 2:
            consecutive = []
            current_run = [imbalance_indices[0]]
            
            for idx in imbalance_indices[1:]:
                if idx == current_run[-1] + 1:
                    current_run.append(idx)
                else:
                    if len(current_run) >= 2:
                        consecutive.append(current_run)
                    current_run = [idx]
            
            if len(current_run) >= 2:
                consecutive.append(current_run)
            
            for run in consecutive:
                start_ts = sec_df.iloc[run[0]]['timestamp']
                end_ts = sec_df.iloc[run[-1]]['timestamp']
                failures.append(make_anomaly(
                    SENSOR_ID, start_ts, end_ts + pd.Timedelta(hours=1),
                    "technical", "direction_imbalance", "direction", "high",
                    "vol_direct", sec_df.iloc[run[0]]['vol_direct'], None, 0.0, 2.8,
                    f"Дисбаланс направлений в течение {len(run)} часов"
                ))
                for idx in run:
                    bad_hours.add((section, sec_df.iloc[idx]['timestamp']))
    
    return failures, bad_hours


# 3. ЗАПОЛНЕНИЕ ПРОПУСКОВ

def fill_gaps(df, bad_hours):
    """Заполняет пропуски."""
    df = df.copy()
    df['date'] = df['timestamp'].dt.date
    df['hour'] = df['timestamp'].dt.hour
    df['day_type'] = df['timestamp'].apply(day_kind)
    df['month'] = df['timestamp'].dt.month
    
    full_idx = pd.date_range(
        start=df['timestamp'].min(),
        end=df['timestamp'].max(),
        freq='h'  # Исправлено: 'H' -> 'h' для pandas 2.2+
    )
    
    filled_records = []
    for section, sec_df in df.groupby('road_section'):
        sec_df = sec_df.set_index('timestamp').reindex(full_idx)
        sec_df['road_section'] = section
        
        for (s, ts) in bad_hours:
            if s == section and ts in sec_df.index:
                for col in ['vol_total', 'vol_direct', 'vol_reverse', 'speed_direct', 'speed_reverse', 'load_direct', 'load_reverse']:
                    sec_df.loc[ts, col] = np.nan
        
        for col in ['vol_total', 'vol_direct', 'vol_reverse', 'speed_direct', 'speed_reverse', 'load_direct', 'load_reverse', 'cars_total']:
            sec_df[col] = sec_df[col].interpolate(method='linear', limit=3)
            sec_df[col] = sec_df.groupby([sec_df.index.month, sec_df.index.hour, sec_df['day_type']])[col].transform(
                lambda x: x.fillna(x.median())
            )
            sec_df[col] = sec_df[col].fillna(sec_df[col].median())
        
        sec_df = sec_df.reset_index().rename(columns={'index': 'timestamp'})
        sec_df['imputation_method'] = sec_df['timestamp'].apply(
            lambda ts: 'interpolated' if (section, ts) in bad_hours else ''
        )
        filled_records.append(sec_df)
    
    return pd.concat(filled_records, ignore_index=True)


# 4. ДЕТЕКТОРЫ ТРАФИКА

def _hourly_detector(work, sensor_id, metric, low_type, high_type, noun):
    """Часовые аномалии."""
    found = []
    floor = _floor_for(metric)
    
    for (month, kind), group in work.groupby(['month', 'day_type']):
        if len(group['date'].unique()) < 4:
            continue
        
        stats = {}
        for hour, values in group.groupby('hour')[metric]:
            median, mad = _center_and_mad(values)
            sigma = _sigma(median, mad, floor)
            stats[hour] = {'median': median, 'sigma': sigma}
        
        if not stats:
            continue
        
        marked = group.copy()
        marked['median'] = marked['hour'].map(lambda h: stats.get(h, {}).get('median', np.nan))
        marked['sigma'] = marked['hour'].map(lambda h: stats.get(h, {}).get('sigma', np.nan))
        marked = marked[marked['median'].notna() & marked['sigma'].notna()]
        
        marked['z'] = (marked[metric] - marked['median']) / marked['sigma']
        marked['residual'] = marked[metric] - marked['median']
        marked['dev'] = np.where(marked['median'] > 0, marked['residual'] / marked['median'] * 100, 0.0)
        
        bar = np.maximum(np.abs(marked['median']) * PRACTICAL_HOUR_PCT / 100.0, PRACTICAL_HOUR_ABS)
        drop = (marked['z'] <= -HOURLY_Z) & (marked['residual'] <= -bar)
        spike = (marked['z'] >= HOURLY_Z) & (marked['residual'] >= bar)
        marked['flag'] = np.where(drop, -1, np.where(spike, 1, 0))
        
        marked = marked.sort_values('timestamp')
        for sign, type_, verb in ((-1, low_type, "ниже"), (1, high_type, "выше")):
            flagged = marked[marked['flag'] == sign]
            if flagged.empty:
                continue
            
            gaps = flagged['timestamp'].diff() != pd.Timedelta(hours=1)
            for _, run in flagged.groupby(gaps.cumsum()):
                if len(run) < HOURLY_MIN_RUN:
                    continue
                
                start = run['timestamp'].iloc[0]
                end = run['timestamp'].iloc[-1] + pd.Timedelta(hours=1)
                actual = float(run[metric].mean())
                expected = float(run['hour'].map(lambda h: stats.get(h, {}).get('median', np.nan)).mean())
                dev = (actual - expected) / expected * 100 if expected > 0 else 0.0
                score = min(float(run['z'].abs().max()), 9.9)
                
                found.append(make_anomaly(
                    sensor_id, start, end, "traffic", type_, "amplitude",
                    level_from_score(score), metric, actual, expected, dev, score,
                    f"{noun} держится {verb} нормы на {abs(dev):.0f}% в течение "
                    f"{len(run)} часов подряд (с {start:%H:%M}, |z| = {score:.1f})"
                ))
    
    return found

def _speed_drop_runs(work, sensor_id, speed_metric, volume_metric, where):
    """Падение скорости."""
    found = []
    
    for (month, kind), group in work.groupby(['month', 'day_type']):
        reference = group[group[volume_metric] > 0]
        if reference['date'].nunique() < 4:
            continue
        
        stats = {}
        for hour, values in reference.groupby('hour')[speed_metric]:
            median, mad = _center_and_mad(values)
            sigma = _sigma(median, mad, FLOOR_SPEED)
            stats[hour] = {'median': median, 'sigma': sigma}
        
        if not stats:
            continue
        
        marked = group[group[volume_metric] > 0].copy()
        marked['median'] = marked['hour'].map(lambda h: stats.get(h, {}).get('median', np.nan))
        marked['sigma'] = marked['hour'].map(lambda h: stats.get(h, {}).get('sigma', np.nan))
        marked = marked[marked['median'].notna() & marked['sigma'].notna()]
        
        marked['gap'] = marked[speed_metric] - marked['median']
        marked['z'] = marked['gap'] / marked['sigma']
        flagged = marked[(marked['gap'] <= -SPEED_DROP_KMH) & (marked['z'] <= -SPEED_DROP_Z)]
        flagged = flagged.sort_values('timestamp')
        
        if flagged.empty:
            continue
        
        gaps = flagged['timestamp'].diff() != pd.Timedelta(hours=1)
        for _, run in flagged.groupby(gaps.cumsum()):
            if len(run) < SPEED_DROP_MIN_RUN:
                continue
            
            start = run['timestamp'].iloc[0]
            end = run['timestamp'].iloc[-1] + pd.Timedelta(hours=1)
            actual = float(run[speed_metric].mean())
            expected = float(run['hour'].map(lambda h: stats.get(h, {}).get('median', np.nan)).mean())
            dev = (actual - expected) / expected * 100 if expected > 0 else 0.0
            score = min(float(run['z'].abs().max()), 9.9)
            
            found.append(make_anomaly(
                sensor_id, start, end, "traffic", "speed_drop", "amplitude",
                level_from_score(score), speed_metric, actual, expected, dev, score,
                f"Скорость {where} держится ниже нормы на {abs(actual - expected):.0f} "
                f"км/ч в течение {len(run)} часов подряд (с {start:%H:%M}, |z| = {score:.1f})"
            ))
    
    return found

def _daily_amplitude(work, sensor_id):
    """Суточные аномалии."""
    found = []
    daily = work.groupby(['month', 'day_type', 'date'])['vol_total'].sum().reset_index()
    
    for (month, kind), group in daily.groupby(['month', 'day_type']):
        if len(group) < 4:
            continue
        
        median, mad = _center_and_mad(group['vol_total'])
        if np.isnan(median):
            continue
        
        sigma = _sigma(median, mad, FLOOR_VOLUME_DAY)
        bar = max(abs(median) * PRACTICAL_DAILY_PCT / 100.0, PRACTICAL_DAILY_ABS)
        
        for _, row in group.iterrows():
            day, volume = row['date'], row['vol_total']
            residual = volume - median
            z = residual / sigma
            dev = residual / median * 100 if median > 0 else 0.0
            
            day_ts = pd.Timestamp(day)
            drop = z <= DROUGHT_Z and residual <= -bar
            surge = z >= SURGE_Z and residual >= bar
            
            if not (drop or surge):
                continue
            
            type_ = "volume_spike" if surge else "volume_drop"
            verb = "выше" if surge else "ниже"
            score = min(abs(z), 9.9)
            
            found.append(make_anomaly(
                sensor_id, day_ts, day_ts + pd.Timedelta(days=1), "traffic", type_, "amplitude",
                level_from_score(score), "vol_total", float(volume), median, dev, score,
                f"Суточный трафик {int(volume)} авт. {verb} нормы {int(median)} на {abs(dev):.0f}% (z = {z:.1f})"
            ))
    
    return found

def _congestion(work, sensor_id):
    """Заторы."""
    found = []
    day = work[(work['hour'] >= 7) & (work['hour'] <= 21)]
    slow = (day['speed_direct'] < CONGESTION_SPEED) & (day['load_direct'] > CONGESTION_LOAD)
    slow &= day['speed_direct'].notna() & day['load_direct'].notna()
    flagged = day[slow].sort_values('timestamp')
    
    if flagged.empty:
        return found
    
    run_id = (flagged['timestamp'].diff() != pd.Timedelta(hours=1)).cumsum()
    for _, run in flagged.groupby(run_id):
        if len(run) < CONGESTION_MIN_HOURS:
            continue
        
        start = run['timestamp'].iloc[0]
        end = run['timestamp'].iloc[-1] + pd.Timedelta(hours=1)
        min_row = run.loc[run['speed_direct'].idxmin()]
        load = float(run['load_direct'].max())
        speed = float(min_row['speed_direct'])
        dev = (speed - FREE_FLOW_SPEED) / FREE_FLOW_SPEED * 100
        
        found.append(make_anomaly(
            sensor_id, start, end, "traffic", "congestion", "speed_load",
            "critical", "speed_direct", speed, FREE_FLOW_SPEED, dev, 4.5,
            f"Со {start:%H:%M} несколько часов подряд скорость ниже "
            f"{min_row['speed_direct']:.0f} км/ч при загрузке {load:.0f}%"
        ))
    
    return found

def _profiles(work, sensor_id):
    """Искажение профиля."""
    found = []
    
    for (month, kind), group in work.groupby(['month', 'day_type']):
        complete = group.groupby('date').size()
        full_days = set(complete[complete == 24].index)
        norm = group[group['date'].isin(full_days)].groupby('hour')['vol_total'].median()
        
        if norm.isna().any() or float(norm.std()) <= 0:
            continue
        
        for day, day_df in group.groupby('date'):
            if day not in full_days or len(day_df) < 24:
                continue
            
            profile = day_df.set_index('hour')['vol_total'].sort_index()
            if float(profile.std()) <= 0 or float(profile.sum()) <= PROFILE_MIN_DAY_VOLUME:
                continue
            
            corr = float(np.corrcoef(norm.values, profile.reindex(norm.index).values)[0, 1])
            if corr >= PROFILE_CORR:
                continue
            
            day_ts = pd.Timestamp(day)
            found.append(make_anomaly(
                sensor_id, day_ts, day_ts + pd.Timedelta(days=1), "traffic", "profile_shift", "profile",
                "medium", "vol_total", float(profile.sum()), float(norm.sum()),
                (float(profile.sum()) - float(norm.sum())) / float(norm.sum()) * 100, 2.5,
                f"Форма суточного хода нетипична (корреляция с нормой {corr:.2f} < {PROFILE_CORR:.2f})"
            ))
    
    return found

def _regime_changes(days, sensor_id):
    """Смена режима."""
    found = []
    if days.empty:
        return found
    
    daily = days.groupby('date')['vol_total'].sum()
    daily.index = pd.DatetimeIndex(pd.to_datetime(daily.index))
    daily = daily.sort_index()
    kinds = pd.Series([day_kind(ts) for ts in daily.index], index=daily.index)
    
    last_end = None
    for first in daily.index:
        if last_end is not None and first <= last_end:
            continue
        
        window = pd.DatetimeIndex([first + pd.Timedelta(days=k) for k in range(REGIME_MIN_DAYS)])
        if not window.isin(daily.index).all():
            continue
        
        window_kinds = kinds.loc[window]
        before = daily[
            (daily.index < first) & (daily.index >= first - pd.Timedelta(days=REGIME_LOOKBACK_DAYS))
        ]
        ref_kinds = kinds.loc[before.index]
        
        norms = {}
        skip = False
        for kind in window_kinds.unique():
            refs = before[ref_kinds == kind]
            if len(refs) < REGIME_MIN_REFS:
                skip = True
                break
            median, _ = _center_and_mad(refs)
            if not (median > 0):
                skip = True
                break
            norms[kind] = median
        
        if skip:
            continue
        
        norm_series = window_kinds.map(norms)
        ratios = daily.loc[window].to_numpy() / norm_series.to_numpy()
        drop = bool((ratios < REGIME_RATIO_LO).all())
        spike = bool((ratios > REGIME_RATIO_HI).all())
        
        if not (drop or spike):
            continue
        
        last_end = window[-1]
        mean_volume = float(daily.loc[window].mean())
        mean_norm = float(norm_series.mean())
        dev = (mean_volume - mean_norm) / mean_norm * 100
        score = min(3.5 + (abs(dev) - 25.0) * 0.1, 6.5)
        type_ = "volume_spike" if spike else "volume_drop"
        
        start = first
        found.append(make_anomaly(
            sensor_id, start, window[-1] + pd.Timedelta(days=1), "traffic", type_, "regime",
            level_from_score(score), "vol_total", mean_volume, mean_norm, dev, score,
            f"Поток держится {'выше' if spike else 'ниже'} прежней нормы на "
            f"{abs(dev):.0f}% {REGIME_MIN_DAYS} дней подряд (с {first:%d.%m})"
        ))
    
    return found

def promote_correlated(anomalies):
    """Cross-channel confirmation."""
    traffic = [a for a in anomalies if a['category'] == 'traffic' and a['score'] is not None]
    volume = [a for a in traffic if a['metric'] in ['vol_total', 'vol_direct', 'vol_reverse']]
    speed = [a for a in traffic if a['metric'] in ['speed_direct', 'speed_reverse']]
    
    for v in volume:
        for s in speed:
            if v['sensor_id'] != s['sensor_id']:
                continue
            if pd.Timestamp(v['starts_at']) < pd.Timestamp(s['ends_at']) and pd.Timestamp(s['starts_at']) < pd.Timestamp(v['ends_at']):
                if "вторым каналом" not in (v['description'] or "") and v['score'] < PROMOTE_CAP:
                    v['score'] = round(min(v['score'] + PROMOTE_STEP, PROMOTE_CAP), 1)
                    v['level'] = level_from_score(v['score'])
                    v['description'] += " Находка подтверждена вторым каналом (скорость)."
                if "вторым каналом" not in (s['description'] or "") and s['score'] < PROMOTE_CAP:
                    s['score'] = round(min(s['score'] + PROMOTE_STEP, PROMOTE_CAP), 1)
                    s['level'] = level_from_score(s['score'])
                    s['description'] += " Находка подтверждена вторым каналом (поток)."
    
    return anomalies

def detect_traffic(filled, sensor_id):
    """Все детекторы трафика."""
    found = []
    work = filled.copy()
    work['date'] = work['timestamp'].dt.date
    work['hour'] = work['timestamp'].dt.hour
    work['day_type'] = work['timestamp'].apply(day_kind)
    work['month'] = work['timestamp'].dt.month
    
    hours = work[work['imputation_method'] == '']
    found += _hourly_detector(hours, sensor_id, 'vol_total', 'volume_drop', 'volume_spike', 'Поток')
    found += _hourly_detector(hours, sensor_id, 'vol_direct', 'volume_drop', 'volume_spike', 'Поток (на Воронеж)')
    found += _hourly_detector(hours, sensor_id, 'vol_reverse', 'volume_drop', 'volume_spike', 'Поток (из Воронежа)')
    found += _speed_drop_runs(hours, sensor_id, 'speed_direct', 'vol_direct', 'на Воронеж')
    found += _speed_drop_runs(hours, sensor_id, 'speed_reverse', 'vol_reverse', 'из Воронежа')
    found += _congestion(hours, sensor_id)
    
    days = work[work['imputation_method'] == '']
    found += _daily_amplitude(days, sensor_id)
    found += _profiles(days, sensor_id)
    found += _regime_changes(days, sensor_id)
    
    return found


# 5. МЕСЯЧНЫЕ ДЕТЕКТОРЫ

def detect_monthly(filled, sensor_id):
    """YoY-анализ."""
    found = []
    work = filled.copy()
    work['month_key'] = work['timestamp'].dt.to_period('M')
    work['heavy'] = work['large_trucks_total'] + work['road_trains_total']
    work['lost'] = work['imputation_method'] != ''
    
    stats = []
    for month_key, month_df in work.groupby('month_key'):
        daily = month_df[~month_df['lost']].groupby(month_df['timestamp'].dt.date)['vol_total'].sum()
        volume = month_df[~month_df['lost']]['vol_total'].sum()
        real_hours = month_df[~month_df['lost']]
        
        moving = real_hours[real_hours['speed_direct'].notna() & (real_hours['vol_direct'] > 0)]
        speed = (
            float((moving['speed_direct'] * moving['vol_direct']).sum() / moving['vol_direct'].sum())
            if len(moving) > 0 else np.nan
        )
        
        stats.append({
            'month_key': month_key,
            'completeness': (1 - month_df['lost'].mean()) * 100,
            'madt': daily.mean() if len(daily) > 0 else np.nan,
            'speed': speed,
            'load': real_hours['load_direct'].median(),
            'heavy_share': month_df['heavy'].sum() / volume * 100 if volume > 0 else np.nan,
        })
    
    stats_df = pd.DataFrame(stats)
    stats_df['cal_month'] = [p.month for p in stats_df['month_key']]
    
    for cal_month, group in stats_df.groupby('cal_month'):
        usable = group[group['completeness'] >= MONTH_VALIDITY]
        
        for _, row in group.iterrows():
            start = row['month_key'].start_time
            end = (row['month_key'] + 1).start_time
            
            if row['completeness'] < MONTH_COMPLETENESS:
                found.append(make_anomaly(
                    sensor_id, start, end, "technical", "missing_hours", "integrity",
                    "medium", "vol_total", row['completeness'], 95.0, row['completeness'] - 100.0, 3.0,
                    f"Данные за месяц неполные (сохранено лишь {row['completeness']:.0f}% часов)"
                ))
                continue
            
            if row['completeness'] < MONTH_VALIDITY:
                continue
            
            peers = usable[usable['month_key'] != row['month_key']]
            if peers.empty:
                continue
            
            med_madt = float(peers['madt'].dropna().median())
            med_speed = float(peers['speed'].dropna().median())
            med_load = float(peers['load'].dropna().median())
            med_heavy = float(peers['heavy_share'].dropna().median())
            
            if pd.notna(row['madt']) and med_madt > 0:
                dev = (row['madt'] - med_madt) / med_madt * 100
                if abs(dev) >= YOY_MADT_PCT:
                    surge = dev > 0
                    score = min(3.5 + (abs(dev) - YOY_MADT_PCT) * 0.1, 6.5)
                    found.append(make_anomaly(
                        sensor_id, start, end, "traffic",
                        "volume_spike" if surge else "volume_drop", "yoy",
                        level_from_score(score), "vol_total", row['madt'], med_madt, dev, score,
                        f"Среднесуточная интенсивность месяца {int(row['madt'])} авт. "
                        f"отклонилась на {dev:+.1f}% от нормы {int(med_madt)}"
                    ))
            
            if pd.notna(row['heavy_share']) and pd.notna(med_heavy):
                delta_points = row['heavy_share'] - med_heavy
                if abs(delta_points) >= YOY_MIX_PP:
                    score = min(2.0 + (abs(delta_points) - YOY_MIX_PP) * 0.1, 3.5)
                    dev = delta_points / med_heavy * 100 if med_heavy > 0 else None
                    found.append(make_anomaly(
                        sensor_id, start, end, "traffic", "vehicle_mix_shift", "mix",
                        level_from_score(score), "large_trucks_total", row['heavy_share'], med_heavy, dev, score,
                        f"Доля тяжёлых грузовиков {row['heavy_share']:.1f}% против нормы {med_heavy:.1f}% ({delta_points:+.1f} п.п.)"
                    ))
            
            if pd.notna(row['speed']) and pd.notna(med_speed):
                delta_speed = row['speed'] - med_speed
                if delta_speed <= -YOY_SPEED:
                    score = min(2.5 + (abs(delta_speed) - YOY_SPEED) * 0.1, 5.0)
                    found.append(make_anomaly(
                        sensor_id, start, end, "traffic", "speed_drop", "yoy",
                        level_from_score(score), "speed_direct", row['speed'], med_speed,
                        delta_speed / med_speed * 100 if med_speed > 0 else None, score,
                        f"Средняя скорость за месяц {row['speed']:.0f} км/ч, на {abs(delta_speed):.1f} км/ч ниже нормы"
                    ))
            
            if pd.notna(row['load']) and pd.notna(med_load) and med_load > 0:
                dev = (row['load'] - med_load) / med_load * 100
                if abs(dev) >= LOAD_DEV_PCT:
                    score = min(2.0 + (abs(dev) - LOAD_DEV_PCT) * 0.08, 4.5)
                    found.append(make_anomaly(
                        sensor_id, start, end, "traffic", "level_of_service_shift", "load",
                        level_from_score(score), "load_direct", row['load'], med_load, dev, score,
                        f"Медианная загрузка полосы за месяц {row['load']:.1f}% против нормы {med_load:.1f}% ({dev:+.0f}%)"
                    ))
    
    return found


# 6. ФИНАЛИЗАЦИЯ

def finalize(found):
    """Присваивает ID и флаг is_mega."""
    used = {}
    for item in found:
        stamp = item['starts_at'][:10].replace('-', '')
        base = f"{ROAD}-{item['type'].replace('_', '-')}-{stamp}"
        used[base] = used.get(base, 0) + 1
        item['id'] = f"{base}-{used[base]:02d}"
        
        if (item['category'] == 'traffic' and item['deviation_pct'] is not None
                and abs(item['deviation_pct']) >= MEGA_PCT):
            item['is_mega'] = True
    
    return sorted(found, key=lambda a: (a['sensor_id'], a['starts_at'], a['type'], a['detector']))


# 7. ЭКСПОРТ

def export_to_excel(anomalies, output_path):
    """Экспорт в Excel."""
    if not anomalies:
        pd.DataFrame([{"Инфо": "Аномалии не обнаружены"}]).to_excel(output_path, index=False)
        return
    
    df = pd.DataFrame(anomalies)
    df = df.rename(columns={
        'sensor_id': 'Участок', 'starts_at': 'Начало', 'ends_at': 'Конец',
        'category': 'Категория', 'type': 'Тип', 'detector': 'Детектор',
        'level': 'Уровень', 'metric': 'Метрика', 'actual': 'Факт',
        'expected': 'Норма', 'deviation_pct': 'Отклонение (%)',
        'score': 'Score', 'is_mega': 'Мега', 'event_id': 'Событие',
        'description': 'Описание'
    })
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Аномалии', index=False)
    
    wb = openpyxl.load_workbook(output_path)
    ws = wb.active
    
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    crit_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    high_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Исправлено: ищем индекс колонки "Уровень" вручную
    level_col_idx = None
    for col_idx in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col_idx).value == 'Уровень':
            level_col_idx = col_idx
            break
    
    for row_idx in range(2, ws.max_row + 1):
        level_val = ws.cell(row=row_idx, column=level_col_idx).value if level_col_idx else None
        
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.alignment = Alignment(vertical="center", wrap_text=True)
            if level_val == 'critical':
                c.fill = crit_fill
            elif level_val == 'high':
                c.fill = high_fill
    
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)
    
    ws.row_dimensions[1].height = 28
    wb.save(output_path)

def export_to_json(anomalies, output_path):
    """Экспорт в JSON."""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(anomalies, f, ensure_ascii=False, indent=2)


# 8. ГЛАВНАЯ ФУНКЦИЯ

def main():
    parser = argparse.ArgumentParser(description="Полный анализатор аномалий А-134")
    parser.add_argument("--input", default="./input_data", help="Папка с .xlsx файлами")
    parser.add_argument("--output", default="./reports", help="Папка для отчетов")
    parser.add_argument("--events", default="./events.csv", help="Файл событий (CSV)")
    args = parser.parse_args()
    
    os.makedirs(args.input, exist_ok=True)
    os.makedirs(args.output, exist_ok=True)
    
    excel_files = sorted(glob.glob(os.path.join(args.input, "*.xlsx")))
    if not excel_files:
        print(f"\n[Внимание] В папке '{args.input}' нет файлов .xlsx!")
        return
    
    print(f"\n=== СТАРТ ПОЛНОГО АНАЛИЗА ===")
    print(f"Найдено файлов: {len(excel_files)}")
    
    all_dfs = []
    for idx, fpath in enumerate(excel_files, 1):
        print(f"[{idx}/{len(excel_files)}] Чтение {os.path.basename(fpath)}...", end=" ", flush=True)
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
    
    print("\n1. Поиск технических сбоев...")
    failures, bad_hours = detect_failures(full_df)
    print(f"   Найдено сбоев: {len(failures)}")
    
    print("\n2. Заполнение пропусков...")
    filled_df = fill_gaps(full_df, bad_hours)
    imputed_hours = (filled_df['imputation_method'] != '').sum()
    print(f"   Заполнено часов: {imputed_hours}")
    
    print("\n3. Детекторы трафика...")
    traffic_anomalies = detect_traffic(filled_df, SENSOR_ID)
    print(f"   Найдено аномалий: {len(traffic_anomalies)}")
    
    print("\n4. Месячные детекторы...")
    monthly_anomalies = detect_monthly(filled_df, SENSOR_ID)
    print(f"   Найдено месячных аномалий: {len(monthly_anomalies)}")
    
    print("\n5. Cross-channel confirmation...")
    all_anomalies = failures + promote_correlated(traffic_anomalies + monthly_anomalies)
    
    print("\n6. Финализация (ID, is_mega)...")
    all_anomalies = finalize(all_anomalies)
    
    print("\n7. Экспорт результатов...")
    excel_path = os.path.join(args.output, "traffic_anomalies_report.xlsx")
    export_to_excel(all_anomalies, excel_path)
    print(f"   Excel: {excel_path}")
    
    json_path = os.path.join(args.output, "anomalies.json")
    export_to_json(all_anomalies, json_path)
    print(f"   JSON: {json_path}")
    
    print(f"\n=== АНАЛИЗ ЗАВЕРШЕН ===")
    print(f"Всего аномалий: {len(all_anomalies)}")
    print(f"  Критических: {sum(1 for a in all_anomalies if a['level'] == 'critical')}")
    print(f"  Высоких: {sum(1 for a in all_anomalies if a['level'] == 'high')}")
    print(f"  Мега-аномалий: {sum(1 for a in all_anomalies if a['is_mega'])}")

if __name__ == "__main__":
    main()
