"""
Excelから材料No,測点,名称,方向,配管長,水平距離,施工日付を取得し、それぞれを辞書にいれる
"""

import pandas as pd
import os
from typing import List, Dict, Any, Optional
import ezdxf
import openpyxl

route_name = ""
x_count = 0
y_count = 0 
first_x = 0
second_x = 0
#dataフォルダのExcelデータを読み込み
def read_excel_data(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"The file {file_path} does not exist.")
    # header=3で4行目をヘッダーとして使用
    return pd.read_excel(file_path, header=3, engine='openpyxl')

#Excelデータを辞書に変換（材料Noの下の空白行の配管長・水平距離も含める）
def excel_to_dict(file_path: str) -> List[Dict[str, Any]]:
    df = read_excel_data(file_path)
    required_columns = ['材料No', '測点', '名称', '方向', '配管長', '水平距離', '摘要', '施工日付']
    
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Required column '{col}' is missing from the Excel file.")
    
    data_list = []
    i = 0
    
    while i < len(df):
        row = df.iloc[i]
        
        # 測点が空の場合はスキップ
        if pd.isna(row['測点']) or row['測点'] == '':
            i += 1
            continue
        
        # 基本データを取得
        data_dict = {}
        for col in required_columns:
            value = row[col]
            if pd.isna(value):
                data_dict[col] = None
            else:
                # 数値データの場合、整数なら.0を除去
                if col in ['材料No', '配管長', '水平距離'] and isinstance(value, (int, float)):
                    # 整数の場合はintに変換、小数の場合はfloatのまま
                    data_dict[col] = int(value) if value == int(value) else float(value)
                else:
                    data_dict[col] = value
        
        # 材料Noがある場合、次の行が空白で配管長・水平距離がある場合をチェック
        if not pd.isna(row['材料No']) and i + 1 < len(df):
            next_row = df.iloc[i + 1]
            
            # 次の行の材料Noが空白で、配管長・水平距離がある場合
            if (pd.isna(next_row['材料No']) and 
                not pd.isna(next_row['測点']) and 
                (not pd.isna(next_row['配管長']) or not pd.isna(next_row['水平距離']))):
                
                # 2個目の配管長・水平距離を追加（整数の場合は.0を除去）
                pipe_length_2 = next_row['配管長'] if not pd.isna(next_row['配管長']) else None
                horizontal_distance_2 = next_row['水平距離'] if not pd.isna(next_row['水平距離']) else None
                
                # 整数の場合は.0を除去
                if pipe_length_2 is not None and isinstance(pipe_length_2, (int, float)):
                    data_dict['配管長2'] = int(pipe_length_2) if pipe_length_2 == int(pipe_length_2) else float(pipe_length_2)
                else:
                    data_dict['配管長2'] = pipe_length_2
                
                if horizontal_distance_2 is not None and isinstance(horizontal_distance_2, (int, float)):
                    data_dict['水平距離2'] = int(horizontal_distance_2) if horizontal_distance_2 == int(horizontal_distance_2) else float(horizontal_distance_2)
                else:
                    data_dict['水平距離2'] = horizontal_distance_2
                
                data_dict['測点2'] = next_row['測点'] if not pd.isna(next_row['測点']) else None
                
                # 次の行をスキップ
                i += 2
            else:
                data_dict['配管長2'] = None
                data_dict['水平距離2'] = None
                data_dict['測点2'] = None
                i += 1
        else:
            data_dict['配管長2'] = None
            data_dict['水平距離2'] = None
            data_dict['測点2'] = None
            i += 1
        
        data_list.append(data_dict)
    
    return data_list

#excel_to_dict関数で得た辞書をinputで指定した施工日付と、その前の日（1日前とは限らない）の施工日付の材料がある場合は、その材料一つを、どこから続くか分かるように一つだけ辞書に追加する。

def filter_by_construction_date_with_previous(pipe_data: List[Dict[str, Any]], target_date: str = None) -> List[Dict[str, Any]]:
    """
    指定した施工日付のデータと、辞書リストでの一つ上の材料を追加してフィルタリング
    
    Args:
        pipe_data: excel_to_dict関数で得た辞書のリスト
        target_date: 対象の施工日付（Noneの場合はinputで入力）
    
    Returns:
        フィルタリングされた辞書のリスト
    """
    
    # 利用可能な施工日付を取得（辞書の出現順で）
    available_dates = []
    seen = set()
    for item in pipe_data:
        if item['施工日付'] is not None:
            date_str = str(item['施工日付'])
            if date_str not in seen:
                available_dates.append(date_str)
                seen.add(date_str)
    
    if target_date is None:
        print("利用可能な施工日付:")
        for i, date in enumerate(available_dates):
            print(f"  {i+1}. {date}")
        
        target_date = input("\n施工日付を入力してください: ")
    
    if target_date not in available_dates:
        print(f"指定された施工日付 '{target_date}' が見つかりません。")
        return []
    
    print(f"対象施工日付: {target_date}")
    
    # 指定した施工日付の最初のデータのインデックスを見つける
    first_target_index = None
    target_data = []
    
    for i, item in enumerate(pipe_data):
        if str(item['施工日付']) == target_date:
            if first_target_index is None:
                first_target_index = i
            target_data.append(item)
    
    print(f"指定日付のデータ件数: {len(target_data)}")
    print(f"指定日付の最初のインデックス: {first_target_index}")
    
    # 辞書リストでの一つ上（インデックス - 1）の材料を探して追加
    if first_target_index is not None and first_target_index > 0:
        previous_index = first_target_index - 1
        previous_material = pipe_data[previous_index].copy()
        previous_material['継続元'] = True  # どこから続くかを示すフラグ
        
        print(f"一つ上のデータを追加: インデックス {previous_index}")
        print(f"継続元として追加: 材料No.{previous_material['材料No']}, 測点: {previous_material['測点']}, 施工日付: {previous_material['施工日付']}")
        
        # 継続元を最初に追加
        target_data.insert(0, previous_material)
    else:
        print("一つ上のデータがありません（最初のデータです）")
    
    print(f"\nフィルタリング結果: {len(target_data)}件のデータ")
    
    return target_data

def get_excel_cell_value(excel_file_path, cell_address):
    """
    Excelファイルの特定のセルの値を取得する
    
    Args:
        excel_file_path (str): Excelファイルのパス
        cell_address (str): セルのアドレス（例：'E2'）
        
    Returns:
        値: セルの値（文字列、数値、None等）
    """
    try:
        # Excelファイルを開く
        workbook = openpyxl.load_workbook(excel_file_path)
        
        # アクティブなシートを取得（または特定のシート名を指定）
        worksheet = workbook.active
        # 特定のシートを指定したい場合：
        # worksheet = workbook['シート名']
        
        # 指定されたセルの値を取得
        cell_value = worksheet[cell_address].value
        
        # ワークブックを閉じる
        workbook.close()
        
        return cell_value
        
    except Exception as e:
        print(f"エラー: セル{cell_address}の値を取得できませんでした - {e}")
        return None

#不断水の略図
#普通のフランジタイプの不断水割穿孔工の略図生成（方向SF)
def under_pressure_drilling_work_sf(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    
    msp.add_circle(center=(base_x, base_y), radius=10000000)
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    msp.add_line((base_x + 15000000, base_y + 5000000), (base_x + 15000000, base_y - 5000000))

#普通のフランジタイプの不断水割穿孔工の略図生成（方向FS)
def under_pressure_drilling_work_fs(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
    msp.add_line((base_x, base_y), (base_x + 5000000, base_y))
    msp.add_circle(center=(base_x + 15000000, base_y), radius=10000000)

#不断水ソフトシール仕切弁付(方向SF)
def under_pressure_drilling_work_ben_sf(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_circle(center=(base_x, base_y), radius=10000000)
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    msp.add_line((base_x + 15000000, base_y + 5000000), (base_x + 15000000, base_y - 5000000))
    msp.add_line((base_x + 15000000, base_y + 5000000), (base_x + 33000000, base_y - 5000000))
    msp.add_line((base_x + 15000000, base_y - 5000000), (base_x + 33000000, base_y + 5000000))
    msp.add_line((base_x + 33000000, base_y + 5000000), (base_x + 33000000, base_y - 5000000))
    msp.add_line((base_x + 33000000, base_y), (base_x + 38000000, base_y))
    msp.add_line((base_x + 38000000, base_y + 5000000), (base_x + 38000000, base_y - 5000000))

#不断水ソフトシール仕切弁付(方向FS)
def under_pressure_drilling_work_ben_fs(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
    msp.add_line((base_x, base_y), (base_x + 5000000, base_y))
    msp.add_line((base_x + 5000000, base_y + 5000000), (base_x + 5000000, base_y - 5000000))
    msp.add_line((base_x + 5000000, base_y + 5000000), (base_x + 23000000, base_y - 5000000))
    msp.add_line((base_x + 5000000, base_y - 5000000), (base_x + 23000000, base_y + 5000000))
    msp.add_line((base_x + 23000000, base_y + 5000000), (base_x + 23000000, base_y - 5000000))
    msp.add_line((base_x + 23000000, base_y), (base_x + 28000000, base_y))
    msp.add_circle(center=(base_x + 38000000, base_y), radius=10000000)

#キャップ外しSB
def start_sb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))

#GX関係略図
#短管FA
def gx_tankan_FA(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
#フランジ
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
#直部分
    msp.add_line((base_x, base_y), (base_x + 10000000, base_y))
#受け口上
    msp.add_line((base_x + 10000000, base_y), (base_x + 17000000, base_y + 7000000))
    msp.add_line((base_x + 17000000, base_y + 7000000), (base_x + 24500000, base_y + 7000000))
#受け口下
    msp.add_line((base_x + 10000000, base_y), (base_x + 17000000, base_y - 7000000))
    msp.add_line((base_x + 17000000, base_y - 7000000), (base_x + 24500000, base_y - 7000000))
#受け口円
    msp.add_circle(center=(base_x + 20750000, base_y), radius=2500000)
#受け口中
    msp.add_line((base_x + 19950000, base_y + 10000000), (base_x + 19950000, base_y - 10000000))
    msp.add_line((base_x + 21650000, base_y + 10000000), (base_x + 21650000, base_y - 10000000))

#GX短管AF
def gx_tankan_AF(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 24500000, base_y))
    #フランジ部分
    msp.add_line((base_x + 24500000, base_y + 5000000), (base_x + 24500000, base_y - 5000000))

#GXHB両受ベンドAA
def gx_hb_ryouuke_bend_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 19500000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24500000, base_y))
    #受け口上
    msp.add_line((base_x + 24500000, base_y), (base_x + 31500000, base_y + 7000000))
    msp.add_line((base_x + 31500000, base_y + 7000000), (base_x + 39000000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 24500000, base_y), (base_x + 31500000, base_y - 7000000))
    msp.add_line((base_x + 31500000, base_y - 7000000), (base_x + 39000000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 35250000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 34450000, base_y + 10000000), (base_x + 34450000, base_y - 10000000))
    msp.add_line((base_x + 36150000, base_y + 10000000), (base_x + 36150000, base_y - 10000000))

#GXHB片受ベンドAB
def gx_hb_katauke_bend_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 19500000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))    

#GXHB片受ベンドBA
def gx_hb_katauke_bend_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 8000000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 23750000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))

#GX二受丁字管AB
def gx_niuke_tyouji_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #上直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 19500000, base_y + 5000000))
    #上受け口
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 12500000, base_y + 12000000))
    msp.add_line((base_x + 12500000, base_y + 12000000), (base_x + 12500000, base_y + 19500000))
    #下受け口
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 26500000, base_y + 12000000))
    msp.add_line((base_x + 26500000, base_y + 12000000), (base_x + 26500000, base_y + 19500000))
    #受け口中
    msp.add_line((base_x + 9500000, base_y + 14920000), (base_x + 29500000, base_y + 14920000))
    msp.add_line((base_x + 9500000, base_y + 16590000), (base_x + 29500000, base_y + 16590000))
    #受け口円
    msp.add_circle(center=(base_x + 19500000, base_y + 15750000), radius=2500000)
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))  

#GX二受丁字管BA
def gx_niuke_tyouji_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #上受け口
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 1000000, base_y + 12000000))
    msp.add_line((base_x + 1000000, base_y + 12000000), (base_x + 1000000, base_y + 19500000))
    #下受け口
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 15000000, base_y + 12000000))
    msp.add_line((base_x + 15000000, base_y + 12000000), (base_x + 15000000, base_y + 19500000))
    #受け口中
    msp.add_line((base_x -2000000, base_y + 14920000), (base_x + 18000000, base_y + 14920000))
    msp.add_line((base_x -2000000, base_y + 16590000), (base_x + 18000000, base_y + 16590000))
    #受け口円
    msp.add_circle(center=(base_x + 8000000, base_y + 15750000), radius=2500000)
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 23750000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))

#GXF付丁字管AB
def gx_f_tyouji_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #上直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 19500000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 14500000, base_y + 5000000), (base_x + 24500000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))  

#GXF付丁字管BA
def gx_f_tyouji_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 3000000, base_y + 5000000), (base_x + 13000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 23750000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))
#GX継ぎ輪AA
def gx_tugiwa_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直上
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 20000000, base_y + 3500000))
    #直下
    msp.add_line((base_x + 11000000, base_y - 3500000), (base_x + 20000000, base_y - 3500000))
    #受け口上
    msp.add_line((base_x + 16500000, base_y), (base_x + 23500000, base_y + 7000000))
    msp.add_line((base_x + 23500000, base_y + 7000000), (base_x + 31000000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 16500000, base_y), (base_x + 23500000, base_y - 7000000))
    msp.add_line((base_x + 23500000, base_y - 7000000), (base_x + 31000000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 27250000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 26450000, base_y + 10000000), (base_x + 26450000, base_y - 10000000))
    msp.add_line((base_x + 28150000, base_y + 10000000), (base_x + 28150000, base_y - 10000000))

#GX両受短管AA
def gx_ryouuke_tannkann_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #受け口上
    msp.add_line((base_x + 14500000, base_y), (base_x + 21500000, base_y + 7000000))
    msp.add_line((base_x + 21500000, base_y + 7000000), (base_x + 29000000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 14500000, base_y), (base_x + 21500000, base_y - 7000000))
    msp.add_line((base_x + 21500000, base_y - 7000000), (base_x + 29000000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 25250000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 24450000, base_y + 10000000), (base_x + 24450000, base_y - 10000000))
    msp.add_line((base_x + 26150000, base_y + 10000000), (base_x + 26150000, base_y - 10000000))

#GX両受短管下がり45°AA
def gx_ryouuke_tannkann_down_45_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y - 2650000), radius=2500000)
    #上受け口
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 20150000, base_y - 10250000))
    msp.add_line((base_x + 20150000, base_y - 10250000), (base_x + 25450000, base_y - 15550000))
    #下受け口
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 10250000, base_y - 20150000))
    msp.add_line((base_x + 10250000, base_y - 20150000), (base_x + 15550000, base_y - 25450000))
    #受け口円
    msp.add_circle(center=(base_x + 17850000, base_y - 17850000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 10250000, base_y - 24350000), (base_x + 24350000, base_y - 10250000))
    msp.add_line((base_x + 11450000, base_y - 25550000), (base_x + 25550000, base_y - 11450000))

#GX両受短管上がり45°AA
def gx_ryouuke_tannkann_up_45_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y + 350000))
    msp.add_line((base_x + 10250000, base_y + 350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y + 2650000), radius=2500000)
    #上受け口
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 20150000, base_y + 10250000))
    msp.add_line((base_x + 20150000, base_y + 10250000), (base_x + 25450000, base_y + 15550000))
    #下受け口
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 10250000, base_y + 20150000))
    msp.add_line((base_x + 10250000, base_y + 20150000), (base_x + 15550000, base_y + 25450000))
    #受け口円
    msp.add_circle(center=(base_x + 17850000, base_y + 17850000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 10250000, base_y + 24350000), (base_x + 24350000, base_y + 10250000))
    msp.add_line((base_x + 11450000, base_y + 25550000), (base_x + 25550000, base_y + 11450000))

#GX受挿し片落管AB
def gx_ukeoshikataochi_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直上
    msp.add_line((base_x + 9470000, base_y + 5030000), (base_x + 23800000, base_y))
    #直下
    msp.add_line((base_x + 9470000, base_y - 5030000), (base_x + 23800000, base_y))

#GX受挿し片落管BA
def gx_ukeoshikataochi_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直上
    msp.add_line((base_x, base_y), (base_x + 14330000, base_y + 5030000))
    #直下
    msp.add_line((base_x, base_y), (base_x + 14330000, base_y - 5030000))
    #受け口上
    msp.add_line((base_x + 9300000, base_y), (base_x + 16300000, base_y + 7000000))
    msp.add_line((base_x + 16300000, base_y + 7000000), (base_x + 23800000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 9300000, base_y), (base_x + 16300000, base_y - 7000000))
    msp.add_line((base_x + 16300000, base_y - 7000000), (base_x + 23800000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 20050000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 19250000, base_y + 10000000), (base_x + 19250000, base_y - 10000000))
    msp.add_line((base_x + 20950000, base_y + 10000000), (base_x + 20950000, base_y - 10000000))

#GXVB両受ベンド45°AA（下がりはじめ）
def gx_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 23000000, base_y - 3500000))
    #上受け口
    msp.add_line((base_x + 23000000, base_y - 3500000), (base_x + 32900000, base_y - 3500000))
    msp.add_line((base_x + 32900000, base_y - 3500000), (base_x + 38200000, base_y - 8800000))
    #下受け口
    msp.add_line((base_x + 23000000, base_y - 3500000), (base_x + 23000000, base_y - 13400000))
    msp.add_line((base_x + 23000000, base_y - 13400000), (base_x + 28300000, base_y - 18700000))
    #受け口円
    msp.add_circle(center=(base_x + 30600000, base_y - 11100000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 23000000, base_y - 17600000), (base_x + 37100000, base_y - 3500000))
    msp.add_line((base_x + 24200000, base_y - 18800000), (base_x + 38300000, base_y - 4700000))

#GXVB両受ベンド45°AA（下がり終わり）
def gx_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y - 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 13790000, base_y - 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y - 13790000), (base_x + 18790000, base_y - 13790000))
    #受け口上
    msp.add_line((base_x + 18790000, base_y - 13790000), (base_x + 25790000, base_y - 6790000))
    msp.add_line((base_x + 25790000, base_y - 6790000), (base_x + 33290000, base_y - 6790000))
    #受け口下
    msp.add_line((base_x + 18790000, base_y - 13790000), (base_x + 25790000, base_y - 20790000))
    msp.add_line((base_x + 25790000, base_y - 20790000), (base_x + 33290000, base_y - 20790000))
    #受け口円
    msp.add_circle(center=(base_x + 29540000, base_y - 13790000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 28740000, base_y - 3790000), (base_x + 28740000, base_y - 23790000))
    msp.add_line((base_x + 30440000, base_y - 3790000), (base_x + 30440000, base_y - 23790000))

#GXVB両受ベンド45°AA（上がりはじめ）
def gx_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 23000000, base_y + 3500000))
    #上受け口
    msp.add_line((base_x + 23000000, base_y + 3500000), (base_x + 23000000, base_y + 13400000))
    msp.add_line((base_x + 23000000, base_y + 13400000), (base_x + 28300000, base_y + 18700000))
    #下受け口
    msp.add_line((base_x + 23000000, base_y + 3500000), (base_x + 32900000, base_y + 3500000))
    msp.add_line((base_x + 32900000, base_y + 3500000), (base_x + 38200000, base_y + 8800000))
    #受け口円
    msp.add_circle(center=(base_x + 30600000, base_y + 11100000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 23000000, base_y + 17600000), (base_x + 37100000, base_y + 3500000))
    msp.add_line((base_x + 24200000, base_y + 18800000), (base_x + 38300000, base_y + 4700000))

#GXVB両受ベンド45°AA（上がり終わり）
def gx_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y +350000))
    msp.add_line((base_x + 10250000, base_y +350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y + 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 13790000, base_y + 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y + 13790000), (base_x + 18790000, base_y + 13790000))
    #受け口上
    msp.add_line((base_x + 18790000, base_y + 13790000), (base_x + 25790000, base_y + 6790000))
    msp.add_line((base_x + 25790000, base_y + 6790000), (base_x + 33290000, base_y + 6790000))
    #受け口下
    msp.add_line((base_x + 18790000, base_y + 13790000), (base_x + 25790000, base_y + 20790000))
    msp.add_line((base_x + 25790000, base_y + 20790000), (base_x + 33290000, base_y + 20790000))
    #受け口円
    msp.add_circle(center=(base_x + 29540000, base_y + 13790000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 28740000, base_y + 3790000), (base_x + 28740000, base_y + 23790000))
    msp.add_line((base_x + 30440000, base_y + 3790000), (base_x + 30440000, base_y + 23790000))

#GXVB両受ベンド22.5°AA（下がりはじめ）
def gx_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24110000, base_y - 1910000))
    #上受け口
    msp.add_line((base_x + 24110000, base_y - 1910000), (base_x + 33260000, base_y + 1880000))
    msp.add_line((base_x + 33260000, base_y + 1880000), (base_x + 40190000, base_y - 990000))
    #下受け口
    msp.add_line((base_x + 24110000, base_y - 1910000), (base_x + 27900000, base_y - 11060000))
    msp.add_line((base_x + 27900000, base_y - 11060000), (base_x + 34830000, base_y - 13930000))
    #受け口円
    msp.add_circle(center=(base_x + 34040000, base_y - 6020000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 29480000, base_y - 14960000), (base_x + 37130000, base_y + 3520000))
    msp.add_line((base_x + 31050000, base_y - 15610000), (base_x + 38700000, base_y + 2870000))

#GXVB両受ベンド22.5°AA（下がり終わり）
def gx_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y - 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), (base_x + 18020000, base_y - 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y - 7460000), (base_x + 23020000, base_y - 7460000))
    #受け口上
    msp.add_line((base_x + 23020000, base_y - 7460000), (base_x + 30020000, base_y - 460000))
    msp.add_line((base_x + 30020000, base_y - 460000), (base_x + 37520000, base_y - 460000))
    #受け口下
    msp.add_line((base_x + 23020000, base_y - 7460000), (base_x + 30020000, base_y - 14460000))
    msp.add_line((base_x + 30020000, base_y - 14460000), (base_x + 37520000, base_y - 14460000))
    #受け口円
    msp.add_circle(center=(base_x + 33770000, base_y - 7460000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 32970000, base_y + 2540000), (base_x + 32970000, base_y - 17460000))
    msp.add_line((base_x + 34670000, base_y + 2540000), (base_x + 34670000, base_y - 17460000))

#GXVB両受ベンド22.5°AA（上がりはじめ）
def gx_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24110000, base_y + 1910000))
    #上受け口
    msp.add_line((base_x + 24110000, base_y + 1910000), (base_x + 27900000, base_y + 11060000))
    msp.add_line((base_x + 27900000, base_y + 11060000), (base_x + 34830000, base_y + 13930000))
    #下受け口
    msp.add_line((base_x + 24110000, base_y + 1910000), (base_x + 33260000, base_y - 1880000))
    msp.add_line((base_x + 33260000, base_y - 1880000), (base_x + 40190000, base_y + 990000))
    #受け口円
    msp.add_circle(center=(base_x + 34040000, base_y + 6020000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 29480000, base_y + 14960000), (base_x + 37130000, base_y - 3520000))
    msp.add_line((base_x + 31050000, base_y + 15610000), (base_x + 38700000, base_y - 2870000))

#GXVB両受ベンド22.5°AA（上がり終わり）
def gx_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y + 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), (base_x + 18020000, base_y + 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y + 7460000), (base_x + 23020000, base_y + 7460000))
    #受け口上
    msp.add_line((base_x + 23020000, base_y + 7460000), (base_x + 30020000, base_y + 460000))
    msp.add_line((base_x + 30020000, base_y + 460000), (base_x + 37520000, base_y + 460000))
    #受け口下
    msp.add_line((base_x + 23020000, base_y + 7460000), (base_x + 30020000, base_y + 14460000))
    msp.add_line((base_x + 30020000, base_y + 14460000), (base_x + 37520000, base_y + 14460000))
    #受け口円
    msp.add_circle(center=(base_x + 33770000, base_y + 7460000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 32970000, base_y - 2540000), (base_x + 32970000, base_y + 17460000))
    msp.add_line((base_x + 34670000, base_y - 2540000), (base_x + 34670000, base_y + 17460000))

#GXVB両受ベンド11.25°AA（下がりはじめ）
def gx_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24400000, base_y - 980000))
    #上受け口
    msp.add_line((base_x + 24400000, base_y - 980000), (base_x + 32630000, base_y + 4520000))
    msp.add_line((base_x + 32630000, base_y + 4520000), (base_x + 39990000, base_y + 3060000))
    #下受け口
    msp.add_line((base_x + 24400000, base_y - 980000), (base_x + 29900000, base_y - 9210000))
    msp.add_line((base_x + 29900000, base_y - 9210000), (base_x + 37260000, base_y - 10670000))
    #受け口中
    msp.add_line((base_x + 32210000, base_y - 12730000), (base_x + 36110000, base_y + 6890000))
    msp.add_line((base_x + 33860000, base_y - 13060000), (base_x + 37780000, base_y + 6580000))
    #受け口円
    msp.add_circle(center=(base_x + 34940000, base_y - 3080000), radius=2500000)

#GXVB両受ベンド11.25°AA（下がり終わり）
def gx_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y - 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), (base_x + 19130000, base_y - 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y - 3800000), (base_x + 24130000, base_y - 3800000))
    #受け口上
    msp.add_line((base_x + 24130000, base_y - 3800000), (base_x + 31130000, base_y + 3200000))
    msp.add_line((base_x + 31130000, base_y + 3200000), (base_x + 38630000, base_y + 3200000))
    #受け口下
    msp.add_line((base_x + 24130000, base_y - 3800000), (base_x + 31130000, base_y - 10800000))
    msp.add_line((base_x + 31130000, base_y - 10800000), (base_x + 38630000, base_y - 10800000))
    #受け口円
    msp.add_circle(center=(base_x + 34880000, base_y - 3800000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 34080000, base_y + 6200000), (base_x + 34080000, base_y - 13800000))
    msp.add_line((base_x + 35780000, base_y + 6200000), (base_x + 35780000, base_y - 13800000))

#GXVB両受ベンド11.25°AA（上がりはじめ）
def gx_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24400000, base_y + 980000))
    #下受け口
    msp.add_line((base_x + 24400000, base_y + 980000), (base_x + 29900000, base_y + 9210000))
    msp.add_line((base_x + 29900000, base_y + 9210000), (base_x + 37260000, base_y + 10670000))
    #上受け口
    msp.add_line((base_x + 24400000, base_y + 980000), (base_x + 32630000, base_y - 4520000))
    msp.add_line((base_x + 32630000, base_y - 4520000), (base_x + 39990000, base_y - 3060000))
    #受け口中
    msp.add_line((base_x + 32210000, base_y + 12730000), (base_x + 36110000, base_y - 6890000))
    msp.add_line((base_x + 33860000, base_y + 13060000), (base_x + 37780000, base_y - 6580000))
    #受け口円
    msp.add_circle(center=(base_x + 34940000, base_y + 3080000), radius=2500000)

#GXVB両受ベンド11.25°AA（上がり終わり）
def gx_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y + 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), (base_x + 19130000, base_y + 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y + 3800000), (base_x + 24130000, base_y + 3800000))
    #受け口上
    msp.add_line((base_x + 24130000, base_y + 3800000), (base_x + 31130000, base_y - 3200000))
    msp.add_line((base_x + 31130000, base_y - 3200000), (base_x + 38630000, base_y - 3200000))
    #受け口下
    msp.add_line((base_x + 24130000, base_y + 3800000), (base_x + 31130000, base_y + 10800000))
    msp.add_line((base_x + 31130000, base_y + 10800000), (base_x + 38630000, base_y + 10800000))
    #受け口円
    msp.add_circle(center=(base_x + 34880000, base_y + 3800000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 34080000, base_y - 6200000), (base_x + 34080000, base_y + 13800000))
    msp.add_line((base_x + 35780000, base_y - 6200000), (base_x + 35780000, base_y + 13800000))

#GXVB片受ベンド45°AB（下がりはじめ）
def gx_vb_katauke_bend_45_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 25160000, base_y - 5660000))

#GXVB片受ベンド45°AB（下がり終わり）
def gx_vb_katauke_bend_45_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y - 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 13790000, base_y - 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y - 13790000), (base_x + 21790000, base_y - 13790000))

#GXVB片受ベンド45°BA（下がりはじめ）
def gx_vb_katauke_bend_45_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y - 3500000))
    #上受け口
    msp.add_line((base_x + 11540000, base_y - 3500000), (base_x + 21440000, base_y - 3500000))
    msp.add_line((base_x + 21440000, base_y - 3500000), (base_x + 26740000, base_y - 8800000))
    #下受け口
    msp.add_line((base_x + 11540000, base_y - 3500000), (base_x + 11540000, base_y - 13400000))
    msp.add_line((base_x + 11540000, base_y - 13400000), (base_x + 16840000, base_y - 18700000))
    #受け口円
    msp.add_circle(center=(base_x + 19140000, base_y - 11100000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 11540000, base_y - 17600000), (base_x + 25640000, base_y - 3500000))
    msp.add_line((base_x + 12740000, base_y - 18800000), (base_x + 26840000, base_y - 4700000))

#GXVB片受ベンド45°BA（下がり終わり）
def gx_vb_katauke_bend_45_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y - 5660000), (base_x + 10660000, base_y - 5660000))
    #受け口上
    msp.add_line((base_x + 10660000, base_y - 5660000), (base_x + 17660000, base_y + 1340000))
    msp.add_line((base_x + 17660000, base_y + 1340000), (base_x + 25160000, base_y + 1340000))
    #受け口下
    msp.add_line((base_x + 10660000, base_y - 5660000), (base_x + 17660000, base_y - 12660000))
    msp.add_line((base_x + 17660000, base_y - 12660000), (base_x + 25160000, base_y - 12660000))
    #受け口円
    msp.add_circle(center=(base_x + 21410000, base_y - 5660000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 20610000, base_y + 4340000), (base_x + 20610000, base_y - 15660000))
    msp.add_line((base_x + 22310000, base_y + 4340000), (base_x + 22310000, base_y - 15660000))

#GXVB片受ベンド45°AB（上がりはじめ）
def gx_vb_katauke_bend_45_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 25160000, base_y + 5660000))

#GXVB片受ベンド45°AB（上がり終わり）
def gx_vb_katauke_bend_45_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y +350000))
    msp.add_line((base_x + 10250000, base_y + 350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y + 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 13790000, base_y + 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y + 13790000), (base_x + 21790000, base_y + 13790000))

#GXVB片受ベンド45°BA（上がりはじめ）
def gx_vb_katauke_bend_45_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y + 3500000))
    #上受け口
    msp.add_line((base_x + 11540000, base_y + 3500000), (base_x + 21440000, base_y + 3500000))
    msp.add_line((base_x + 21440000, base_y + 3500000), (base_x + 26740000, base_y + 8800000))
    #下受け口
    msp.add_line((base_x + 11540000, base_y + 3500000), (base_x + 11540000, base_y + 13400000))
    msp.add_line((base_x + 11540000, base_y + 13400000), (base_x + 16840000, base_y + 18700000))
    #受け口円
    msp.add_circle(center=(base_x + 19140000, base_y + 11100000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 11540000, base_y + 17600000), (base_x + 25640000, base_y + 3500000))
    msp.add_line((base_x + 12740000, base_y + 18800000), (base_x + 26840000, base_y + 4700000))

#GXVB片受ベンド45°BA（上がり終わり）
def gx_vb_katauke_bend_45_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y + 5660000), (base_x + 10660000, base_y + 5660000))
    #受け口上
    msp.add_line((base_x + 10660000, base_y + 5660000), (base_x + 17660000, base_y - 1340000))
    msp.add_line((base_x + 17660000, base_y - 1340000), (base_x + 25160000, base_y - 1340000))
    #受け口下
    msp.add_line((base_x + 10660000, base_y + 5660000), (base_x + 17660000, base_y + 12660000))
    msp.add_line((base_x + 17660000, base_y + 12660000), (base_x + 25160000, base_y + 12660000))
    #受け口円
    msp.add_circle(center=(base_x + 21410000, base_y + 5660000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 20610000, base_y - 4340000), (base_x + 20610000, base_y + 15660000))
    msp.add_line((base_x + 22310000, base_y - 4340000), (base_x + 22310000, base_y + 15660000))

#GXVB片受ベンド22.5°AB（下がりはじめ）
def gx_vb_katauke_bend_22_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 26870000, base_y - 3060000))

#GXVB片受ベンド22.5°AB（下がり終わり）
def gx_vb_katauke_bend_22_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y - 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), (base_x + 18020000, base_y - 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y - 7460000), (base_x + 26020000, base_y - 7460000))

#GXVB片受ベンド22.5°BA（下がりはじめ）
def gx_vb_katauke_bend_22_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12610000, base_y - 1910000))
    #上受け口
    msp.add_line((base_x + 12610000, base_y - 1910000), (base_x + 21760000, base_y + 1880000))
    msp.add_line((base_x + 21760000, base_y + 1880000), (base_x + 28690000, base_y - 990000))
    #下受け口
    msp.add_line((base_x + 12610000, base_y - 1910000), (base_x + 16400000, base_y - 11060000))
    msp.add_line((base_x + 16400000, base_y - 11060000), (base_x + 23330000, base_y - 13930000))
    #受け口円
    msp.add_circle(center=(base_x + 22540000, base_y - 6020000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 17980000, base_y - 14960000), (base_x + 25630000, base_y + 3520000))
    msp.add_line((base_x + 19550000, base_y - 15610000), (base_x + 27200000, base_y + 2870000))

#GXVB片受ベンド22.5°BA（下がり終わり）
def gx_vb_katauke_bend_22_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7370000, base_y - 3060000))
    #直部分
    msp.add_line((base_x + 7370000, base_y - 3060000), (base_x + 12370000, base_y - 3060000))
    #受け口上
    msp.add_line((base_x + 12370000, base_y - 3060000), (base_x + 19370000, base_y + 3940000))
    msp.add_line((base_x + 19370000, base_y + 3940000), (base_x + 26870000, base_y + 3940000))
    #受け口下
    msp.add_line((base_x + 12370000, base_y - 3060000), (base_x + 19370000, base_y - 10060000))
    msp.add_line((base_x + 19370000, base_y - 10060000), (base_x + 26870000, base_y - 10060000))
    #受け口円
    msp.add_circle(center=(base_x + 23120000, base_y - 3060000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22320000, base_y + 6940000), (base_x + 22320000, base_y - 13060000))
    msp.add_line((base_x + 24020000, base_y + 6940000), (base_x + 24020000, base_y - 13060000))

#GXVB片受ベンド22.5°AB（上がりはじめ）
def gx_vb_katauke_bend_22_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 26870000, base_y + 3060000))

#GXVB片受ベンド22.5°AB（上がり終わり）
def gx_vb_katauke_bend_22_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y + 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), (base_x + 18020000, base_y + 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y + 7460000), (base_x + 26020000, base_y + 7460000))

#GXVB片受ベンド22.5°BA（上がりはじめ）
def gx_vb_katauke_bend_22_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12610000, base_y + 1910000))
    #上受け口
    msp.add_line((base_x + 12610000, base_y + 1910000), (base_x + 21760000, base_y - 1880000))
    msp.add_line((base_x + 21760000, base_y - 1880000), (base_x + 28690000, base_y + 990000))
    #下受け口
    msp.add_line((base_x + 12610000, base_y + 1910000), (base_x + 16400000, base_y + 11060000))
    msp.add_line((base_x + 16400000, base_y + 11060000), (base_x + 23330000, base_y + 13930000))
    #受け口円
    msp.add_circle(center=(base_x + 22540000, base_y + 6020000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 17980000, base_y + 14960000), (base_x + 25630000, base_y - 3520000))
    msp.add_line((base_x + 19550000, base_y + 15610000), (base_x + 27200000, base_y - 2870000))

#GXVB片受ベンド22.5°BA（上がり終わり）
def gx_vb_katauke_bend_22_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7370000, base_y + 3060000))
    #直部分
    msp.add_line((base_x + 7370000, base_y + 3060000), (base_x + 12370000, base_y + 3060000))
    #受け口上
    msp.add_line((base_x + 12370000, base_y + 3060000), (base_x + 19370000, base_y - 3940000))
    msp.add_line((base_x + 19370000, base_y - 3940000), (base_x + 26870000, base_y - 3940000))
    #受け口下
    msp.add_line((base_x + 12370000, base_y + 3060000), (base_x + 19370000, base_y + 10060000))
    msp.add_line((base_x + 19370000, base_y + 10060000), (base_x + 26870000, base_y + 10060000))
    #受け口円
    msp.add_circle(center=(base_x + 23120000, base_y + 3060000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22320000, base_y - 6940000), (base_x + 22320000, base_y + 13060000))
    msp.add_line((base_x + 24020000, base_y - 6940000), (base_x + 24020000, base_y + 13060000))

#GXVB片受ベンド11.25°AB（下がりはじめ）
def gx_vb_katauke_bend_11_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27340000, base_y - 1570000))

#GXVB片受ベンド11.25°AB（下がり終わり）
def gx_vb_katauke_bend_11_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y - 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), (base_x + 19130000, base_y - 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y - 3800000), (base_x + 27130000, base_y - 3800000))

#GXVB片受ベンド11.25°BA（下がりはじめ）
def gx_vb_katauke_bend_11_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y - 980000))
    #上受け口
    msp.add_line((base_x + 12900000, base_y - 980000), (base_x + 21130000, base_y + 4520000))
    msp.add_line((base_x + 21130000, base_y + 4520000), (base_x + 28490000, base_y + 3060000))
    #下受け口
    msp.add_line((base_x + 12900000, base_y - 980000), (base_x + 18400000, base_y - 9210000))
    msp.add_line((base_x + 18400000, base_y - 9210000), (base_x + 25760000, base_y - 10670000))
    #受け口中
    msp.add_line((base_x + 20710000, base_y - 12730000), (base_x + 24610000, base_y + 6890000))
    msp.add_line((base_x + 22360000, base_y - 13060000), (base_x + 26280000, base_y + 6580000))
    #受け口円
    msp.add_circle(center=(base_x + 23440000, base_y - 3080000), radius=2500000)

#GXVB片受ベンド11.25°BA（下がり終わり）
def gx_vb_katauke_bend_11_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y - 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y - 1570000), (base_x + 12840000, base_y - 1570000))
    #受け口上
    msp.add_line((base_x + 12840000, base_y - 1570000), (base_x + 19840000, base_y + 5430000))
    msp.add_line((base_x + 19840000, base_y + 5430000), (base_x + 27340000, base_y + 5430000))
    #受け口下
    msp.add_line((base_x + 12840000, base_y - 1570000), (base_x + 19840000, base_y - 8570000))
    msp.add_line((base_x + 19840000, base_y - 8570000), (base_x + 27340000, base_y - 8570000))
    #受け口円
    msp.add_circle(center=(base_x + 23590000, base_y - 1570000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22790000, base_y + 8430000), (base_x + 22790000, base_y - 11570000))
    msp.add_line((base_x + 24490000, base_y + 8430000), (base_x + 24490000, base_y - 11570000))

#GXVB片受ベンド11.25°AB（上がりはじめ）
def gx_vb_katauke_bend_11_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27340000, base_y + 1570000))

#GXVB片受ベンド11.25°AB（上がり終わり）
def gx_vb_katauke_bend_11_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y + 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), (base_x + 19130000, base_y + 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y + 3800000), (base_x + 27130000, base_y + 3800000))

#GXVB片受ベンド11.25°BA（上がりはじめ）
def gx_vb_katauke_bend_11_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y + 980000))
    #上受け口
    msp.add_line((base_x + 12900000, base_y + 980000), (base_x + 21130000, base_y - 4520000))
    msp.add_line((base_x + 21130000, base_y - 4520000), (base_x + 28490000, base_y - 3060000))
    #下受け口
    msp.add_line((base_x + 12900000, base_y + 980000), (base_x + 18400000, base_y + 9210000))
    msp.add_line((base_x + 18400000, base_y + 9210000), (base_x + 25760000, base_y + 10670000))
    #受け口中
    msp.add_line((base_x + 20710000, base_y + 12730000), (base_x + 24610000, base_y - 6890000))
    msp.add_line((base_x + 22360000, base_y + 13060000), (base_x + 26280000, base_y - 6580000))
    #受け口円
    msp.add_circle(center=(base_x + 23440000, base_y + 3080000), radius=2500000)

#GXVB片受ベンド11.25°BA（上がり終わり）
def gx_vb_katauke_bend_11_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y + 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y + 1570000), (base_x + 12840000, base_y + 1570000))
    #受け口上
    msp.add_line((base_x + 12840000, base_y + 1570000), (base_x + 19840000, base_y - 5430000))
    msp.add_line((base_x + 19840000, base_y - 5430000), (base_x + 27340000, base_y - 5430000))
    #受け口下
    msp.add_line((base_x + 12840000, base_y + 1570000), (base_x + 19840000, base_y + 8570000))
    msp.add_line((base_x + 19840000, base_y + 8570000), (base_x + 27340000, base_y + 8570000))
    #受け口円
    msp.add_circle(center=(base_x + 23590000, base_y + 1570000), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22790000, base_y - 8430000), (base_x + 22790000, base_y + 11570000))
    msp.add_line((base_x + 24490000, base_y - 8430000), (base_x + 24490000, base_y + 11570000))

#GX直管受け口BA
def gx_tyokkann_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 25250000, base_y))
    #上受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y + 7000000))
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 39750000, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 32250000, base_y - 7000000), (base_x + 39750000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 35100000, base_y + 10000000), (base_x + 35100000, base_y - 10000000))
    msp.add_line((base_x + 36700000, base_y + 10000000), (base_x + 36700000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 36000000, base_y), radius=2500000)

#GX直管受け口AB
def gx_tyokkann_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 39750000, base_y))

#GX直管受け口ﾗｲﾅAB
def gx_tyokkann_raina_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 39750000, base_y))
    #ライナ部分
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 18170000, base_y + 9360000))
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 14500000, base_y + 3000000))
    msp.add_line((base_x + 14500000, base_y + 3000000), (base_x + 18170000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 8000000, base_y + 6500000), (base_x + 8000000, base_y - 6500000))
    msp.add_line((base_x + 8500000, base_y + 6000000), (base_x + 8500000, base_y - 6000000))
    msp.add_line((base_x + 9000000, base_y + 5500000), (base_x + 9000000, base_y - 5500000))
    msp.add_line((base_x + 9500000, base_y + 5000000), (base_x + 9500000, base_y - 5000000))
    msp.add_line((base_x + 10000000, base_y + 4500000), (base_x + 10000000, base_y - 4500000))
    msp.add_line((base_x + 10500000, base_y + 4000000), (base_x + 10500000, base_y - 4000000))
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 11000000, base_y - 3500000))
    msp.add_line((base_x + 11500000, base_y + 3000000), (base_x + 11500000, base_y - 3000000))
    msp.add_line((base_x + 12000000, base_y + 2500000), (base_x + 12000000, base_y - 2500000))
    msp.add_line((base_x + 12500000, base_y + 2000000), (base_x + 12500000, base_y - 2000000))
    msp.add_line((base_x + 13000000, base_y + 1500000), (base_x + 13000000, base_y - 1500000))
    msp.add_line((base_x + 13500000, base_y + 1000000), (base_x + 13500000, base_y - 1000000))
    msp.add_line((base_x + 14000000, base_y + 500000), (base_x + 14000000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 9790000, base_y + 10320000),
        'height': 6000000,
    })

#GX直管受け口ﾗｲﾅBA
def gx_tyokkann_raina_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 25250000, base_y))
    #上受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y + 7000000))
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 39750000, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 32250000, base_y - 7000000), (base_x + 39750000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 35100000, base_y + 10000000), (base_x + 35100000, base_y - 10000000))
    msp.add_line((base_x + 36700000, base_y + 10000000), (base_x + 36700000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 36000000, base_y), radius=2500000)
    #ライナ部分
    msp.add_line((base_x + 21580000, base_y + 9360000), (base_x + 28920000, base_y + 9360000))
    msp.add_line((base_x + 21580000, base_y + 9360000), (base_x + 25250000, base_y + 3000000))
    msp.add_line((base_x + 25250000, base_y + 3000000), (base_x + 28920000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 31750000, base_y + 6500000), (base_x + 31750000, base_y - 6500000))
    msp.add_line((base_x + 31250000, base_y + 6000000), (base_x + 31250000, base_y - 6000000))
    msp.add_line((base_x + 30750000, base_y + 5500000), (base_x + 30750000, base_y - 5500000))
    msp.add_line((base_x + 30250000, base_y + 5000000), (base_x + 30250000, base_y - 5000000))
    msp.add_line((base_x + 29750000, base_y + 4500000), (base_x + 29750000, base_y - 4500000))
    msp.add_line((base_x + 29250000, base_y + 4000000), (base_x + 29250000, base_y - 4000000))
    msp.add_line((base_x + 28750000, base_y + 3500000), (base_x + 28750000, base_y - 3500000))
    msp.add_line((base_x + 28250000, base_y + 3000000), (base_x + 28250000, base_y - 3000000))
    msp.add_line((base_x + 27750000, base_y + 2500000), (base_x + 27750000, base_y - 2500000))
    msp.add_line((base_x + 27250000, base_y + 2000000), (base_x + 27250000, base_y - 2000000))
    msp.add_line((base_x + 26750000, base_y + 1500000), (base_x + 26750000, base_y - 1500000))
    msp.add_line((base_x + 26250000, base_y + 1000000), (base_x + 26250000, base_y - 1000000))
    msp.add_line((base_x + 25750000, base_y + 500000), (base_x + 25750000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 20600000, base_y + 10320000),
        'height': 6000000,
    })

#GX直管受け口甲AB
def gx_tyokkann_kou_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y))

#GX直管受け口甲BA
def gx_tyokkann_kou_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))
    #上受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 9850000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 9850000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    msp.add_line((base_x + 11450000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 11450000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 10750000 + item['配管長'] * 10000 / 2, base_y), radius=2500000)

#GX直管受け口ﾗｲﾅ甲AB
def gx_tyokkann_raina_kou_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y))
    #ライナ部分
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 18170000, base_y + 9360000))
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 14500000, base_y + 3000000))
    msp.add_line((base_x + 14500000, base_y + 3000000), (base_x + 18170000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 8000000, base_y + 6500000), (base_x + 8000000, base_y - 6500000))
    msp.add_line((base_x + 8500000, base_y + 6000000), (base_x + 8500000, base_y - 6000000))
    msp.add_line((base_x + 9000000, base_y + 5500000), (base_x + 9000000, base_y - 5500000))
    msp.add_line((base_x + 9500000, base_y + 5000000), (base_x + 9500000, base_y - 5000000))
    msp.add_line((base_x + 10000000, base_y + 4500000), (base_x + 10000000, base_y - 4500000))
    msp.add_line((base_x + 10500000, base_y + 4000000), (base_x + 10500000, base_y - 4000000))
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 11000000, base_y - 3500000))
    msp.add_line((base_x + 11500000, base_y + 3000000), (base_x + 11500000, base_y - 3000000))
    msp.add_line((base_x + 12000000, base_y + 2500000), (base_x + 12000000, base_y - 2500000))
    msp.add_line((base_x + 12500000, base_y + 2000000), (base_x + 12500000, base_y - 2000000))
    msp.add_line((base_x + 13000000, base_y + 1500000), (base_x + 13000000, base_y - 1500000))
    msp.add_line((base_x + 13500000, base_y + 1000000), (base_x + 13500000, base_y - 1000000))
    msp.add_line((base_x + 14000000, base_y + 500000), (base_x + 14000000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 9790000, base_y + 10320000),
        'height': 6000000,
    })

#GX直管受け口ﾗｲﾅ甲BA
def gx_tyokkann_raina_kou_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))
    #上受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 9850000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 9850000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    msp.add_line((base_x + 11450000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 11450000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 10750000 + item['配管長'] * 10000 / 2, base_y), radius=2500000)
    #ライナ部分
    msp.add_line((base_x - 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000), (base_x + 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000))
    msp.add_line((base_x - 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000), (base_x + item['配管長'] * 10000 / 2, base_y + 3000000))
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y + 3000000), (base_x + 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 6500000 + item['配管長'] * 10000 / 2, base_y + 6500000), (base_x + 6500000 + item['配管長'] * 10000 / 2, base_y - 6500000))
    msp.add_line((base_x + 6000000 + item['配管長'] * 10000 / 2, base_y + 6000000), (base_x + 6000000 + item['配管長'] * 10000 / 2, base_y - 6000000))
    msp.add_line((base_x + 5500000 + item['配管長'] * 10000 / 2, base_y + 5500000), (base_x + 5500000 + item['配管長'] * 10000 / 2, base_y - 5500000))
    msp.add_line((base_x + 5000000 + item['配管長'] * 10000 / 2, base_y + 5000000), (base_x + 5000000 + item['配管長'] * 10000 / 2, base_y - 5000000))
    msp.add_line((base_x + 4500000 + item['配管長'] * 10000 / 2, base_y + 4500000), (base_x + 4500000 + item['配管長'] * 10000 / 2, base_y - 4500000))
    msp.add_line((base_x + 4000000 + item['配管長'] * 10000 / 2, base_y + 4000000), (base_x + 4000000 + item['配管長'] * 10000 / 2, base_y - 4000000))
    msp.add_line((base_x + 3500000 + item['配管長'] * 10000 / 2, base_y + 3500000), (base_x + 3500000 + item['配管長'] * 10000 / 2, base_y - 3500000))
    msp.add_line((base_x + 3000000 + item['配管長'] * 10000 / 2, base_y + 3000000), (base_x + 3000000 + item['配管長'] * 10000 / 2, base_y - 3000000))
    msp.add_line((base_x + 2500000 + item['配管長'] * 10000 / 2, base_y + 2500000), (base_x + 2500000 + item['配管長'] * 10000 / 2, base_y - 2500000))
    msp.add_line((base_x + 2000000 + item['配管長'] * 10000 / 2, base_y + 2000000), (base_x + 2000000 + item['配管長'] * 10000 / 2, base_y - 2000000))
    msp.add_line((base_x + 1500000 + item['配管長'] * 10000 / 2, base_y + 1500000), (base_x + 1500000 + item['配管長'] * 10000 / 2, base_y - 1500000))
    msp.add_line((base_x + 1000000 + item['配管長'] * 10000 / 2, base_y + 1000000), (base_x + 1000000 + item['配管長'] * 10000 / 2, base_y - 1000000))
    msp.add_line((base_x + 500000 + item['配管長'] * 10000 / 2, base_y + 500000), (base_x + 500000 + item['配管長'] * 10000 / 2, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x - 4710000 + item['配管長'] * 10000 / 2, base_y + 10320000),
        'height': 6000000,
    })

#GXVB45°直管受け口甲AB（下がり）
def gx_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y - 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), 
                 (base_x + 10250000 + (item['水平距離'] * 10000 / 2), 
                  base_y - 10250000 - (item['水平距離'] * 10000 / 2)))

#GXVB45°直管受け口甲AB（上がり）
def gx_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y + 350000))
    msp.add_line((base_x + 10250000, base_y + 350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #受け口円
    msp.add_circle(center=(base_x + 2650000, base_y + 2650000), radius=2500000)
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), 
                 (base_x + 10250000 + (item['水平距離'] * 10000 / 2), 
                  base_y + 10250000 + (item['水平距離'] * 10000 / 2)))

#GXVB45°直管受け口甲BA（下がり）
def gx_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    #上受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 9900000 + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 9900000 + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15200000 + (item['水平距離'] * 10000 / 2), 
                   base_y - 5300000 - (item['水平距離'] * 10000 / 2)))
    #下受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + (item['水平距離'] * 10000 / 2), base_y - 9900000 - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - 9900000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 5300000 + (item['水平距離'] * 10000 / 2), 
                   base_y - 15200000 - (item['水平距離'] * 10000 / 2)))
    #受け口円
    msp.add_circle(center=(base_x + 7600000 + (item['水平距離'] * 10000 / 2), 
                           base_y - 7600000 - (item['水平距離'] * 10000 / 2)), radius=2500000)
    #受け口中
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - 14100000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 14100000 + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 1200000 + (item['水平距離'] * 10000 / 2), 
                  base_y - 15300000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15300000 + (item['水平距離'] * 10000 / 2), base_y - 1200000 - (item['水平距離'] * 10000 / 2)))

#GXVB45°直管受け口甲BA（上がり）
def gx_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    #上受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 9900000 + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 9900000 + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15200000 + (item['水平距離'] * 10000 / 2), 
                   base_y + 5300000 + (item['水平距離'] * 10000 / 2)))
    #下受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + (item['水平距離'] * 10000 / 2), base_y + 9900000 + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + 9900000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 5300000 + (item['水平距離'] * 10000 / 2), 
                   base_y + 15200000 + (item['水平距離'] * 10000 / 2)))
    #受け口円
    msp.add_circle(center=(base_x + 7600000 + (item['水平距離'] * 10000 / 2), 
                           base_y + 7600000 + (item['水平距離'] * 10000 / 2)), radius=2500000)
    #受け口中
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + 14100000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 14100000 + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 1200000 + (item['水平距離'] * 10000 / 2), 
                  base_y + 15300000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15300000 + (item['水平距離'] * 10000 / 2), base_y + 1200000 + (item['水平距離'] * 10000 / 2)))

#GXVB22.5°直管受け口甲AB（下がり）
def gx_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y - 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), 
                 (base_x + 13400000 + item['水平距離'] * 10000 / 2, 
                  base_y - 5550000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB22.5°直管受け口甲AB（上がり）
def gx_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #受け口円
    msp.add_circle(center=(base_x + 3460000, base_y + 1440000), radius=2500000)
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), 
                 (base_x + 13400000 + item['水平距離'] * 10000 / 2, 
                  base_y + 5550000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB22.5°直管受け口甲BA（下がり）
def gx_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                   base_y + 3790000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                  base_y + 3790000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 16080000 + item['水平距離'] * 10000 / 2, 
                   base_y + 920000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                   base_y - 9150000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                  base_y - 9150000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 10720000 + item['水平距離'] * 10000 / 2, 
                   base_y - 12020000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口円
    msp.add_circle(center=(base_x + 9930000 + item['水平距離'] * 10000 / 2, 
                           base_y - 4110000 - ((item['配管長'] * 10000 / 2) ** 2 
                                               - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), radius=2500000)
    #受け口中
    msp.add_line((base_x + 5370000 + item['水平距離'] * 10000 / 2, 
                  base_y - 13050000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13020000 + item['水平距離'] * 10000 / 2, 
                   base_y + 5430000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 6940000 + item['水平距離'] * 10000 / 2, 
                  base_y - 13700000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 14590000 + item['水平距離'] * 10000 / 2, 
                   base_y + 4780000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB22.5°直管受け口甲BA（上がり）
def gx_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                   base_y - 3790000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                  base_y - 3790000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 16080000 + item['水平距離'] * 10000 / 2, 
                   base_y - 920000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                   base_y + 9150000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                  base_y + 9150000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 10720000 + item['水平距離'] * 10000 / 2, 
                   base_y + 12020000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口円
    msp.add_circle(center=(base_x + 9930000 + item['水平距離'] * 10000 / 2, 
                           base_y + 4110000 + ((item['配管長'] * 10000 / 2) ** 2 
                                               - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), radius=2500000)
    #受け口中
    msp.add_line((base_x + 5370000 + item['水平距離'] * 10000 / 2, 
                  base_y + 13050000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13020000 + item['水平距離'] * 10000 / 2, 
                   base_y - 5430000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 6940000 + item['水平距離'] * 10000 / 2, 
                  base_y + 13700000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 14590000 + item['水平距離'] * 10000 / 2, 
                   base_y - 4780000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB11.25°直管受け口甲AB（下がり）
def gx_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y - 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), 
                 (base_x + 14220000 + item['水平距離'] * 10000 / 2, 
                  base_y - 2830000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB11.25°直管受け口甲AB（上がり）
def gx_vb_tyokkann_kou_11_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #受け口円
    msp.add_circle(center=(base_x + 3680000, base_y + 730000), radius=2500000)
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), 
                 (base_x + 14220000 + item['水平距離'] * 10000 / 2, 
                  base_y + 2830000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#GXVB11.25°直管受け口甲BA（下がり）
def gx_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + item['水平距離'] * 10000 / 2, 
                                    base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                   base_y + 5500000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                  base_y + 5500000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 15590000 + item['水平距離'] * 10000 / 2, 
                   base_y + 4040000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                   base_y - 8230000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                  base_y - 8230000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 12860000 + item['水平距離'] * 10000 / 2, 
                   base_y - 9690000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口中
    msp.add_line((base_x + 7810000 + item['水平距離'] * 10000 / 2, 
                  base_y - 11750000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 11710000 + item['水平距離'] * 10000 / 2, 
                   base_y + 7870000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9460000 + item['水平距離'] * 10000 / 2, 
                  base_y - 12080000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13380000 + item['水平距離'] * 10000 / 2, 
                   base_y + 7560000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口円
    msp.add_circle(center=(base_x + 10540000 + item['水平距離'] * 10000 / 2, 
                           base_y - 2100000 - ((item['配管長'] * 10000 / 2) ** 2 
                                               - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), radius=2500000)

#GXVB11.25°直管受け口甲BA（上がり）
def gx_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                   base_y - 5500000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                  base_y - 5500000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 15590000 + item['水平距離'] * 10000 / 2, 
                   base_y - 4040000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                   base_y + 8230000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                  base_y + 8230000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 12860000 + item['水平距離'] * 10000 / 2, 
                   base_y + 9690000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口中
    msp.add_line((base_x + 7810000 + item['水平距離'] * 10000 / 2, 
                  base_y + 11750000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 11710000 + item['水平距離'] * 10000 / 2, 
                   base_y - 7870000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9460000 + item['水平距離'] * 10000 / 2, 
                  base_y + 12080000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13380000 + item['水平距離'] * 10000 / 2, 
                   base_y - 7560000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口円
    msp.add_circle(center=(base_x + 10540000 + item['水平距離'] * 10000 / 2, 
                           base_y + 2100000 + ((item['配管長'] * 10000 / 2) ** 2 
                                               - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), radius=2500000)

#GX直管受け口乙BB
def gx_tyokkann_otsu_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000, base_y))

#NS関係略図
#NS短管FA
def ns_tankan_FA(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
#フランジ
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
#直部分
    msp.add_line((base_x, base_y), (base_x + 10000000, base_y))
#受け口上
    msp.add_line((base_x + 10000000, base_y), (base_x + 17000000, base_y + 7000000))
    msp.add_line((base_x + 17000000, base_y + 7000000), (base_x + 24500000, base_y + 7000000))
#受け口下
    msp.add_line((base_x + 10000000, base_y), (base_x + 17000000, base_y - 7000000))
    msp.add_line((base_x + 17000000, base_y - 7000000), (base_x + 24500000, base_y - 7000000))
#受け口中
    msp.add_line((base_x + 19950000, base_y + 10000000), (base_x + 19950000, base_y - 10000000))
    msp.add_line((base_x + 21650000, base_y + 10000000), (base_x + 21650000, base_y - 10000000))

#NS短管AF
def ns_tankan_AF(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 24500000, base_y))
    #フランジ部分
    msp.add_line((base_x + 24500000, base_y + 5000000), (base_x + 24500000, base_y - 5000000))

#NSHB両受ベンドAA
def ns_hb_ryouuke_bend_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 19500000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24500000, base_y))
    #受け口上
    msp.add_line((base_x + 24500000, base_y), (base_x + 31500000, base_y + 7000000))
    msp.add_line((base_x + 31500000, base_y + 7000000), (base_x + 39000000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 24500000, base_y), (base_x + 31500000, base_y - 7000000))
    msp.add_line((base_x + 31500000, base_y - 7000000), (base_x + 39000000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 34450000, base_y + 10000000), (base_x + 34450000, base_y - 10000000))
    msp.add_line((base_x + 36150000, base_y + 10000000), (base_x + 36150000, base_y - 10000000))

#NSHB片受ベンドAB
def ns_hb_katauke_bend_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 19500000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))    

#NSHB片受ベンドBA
def ns_hb_katauke_bend_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 8000000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))

#NS二受丁字管AB
def ns_niuke_tyouji_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #上直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 19500000, base_y + 5000000))
    #上受け口
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 12500000, base_y + 12000000))
    msp.add_line((base_x + 12500000, base_y + 12000000), (base_x + 12500000, base_y + 19500000))
    #下受け口
    msp.add_line((base_x + 19500000, base_y + 5000000), (base_x + 26500000, base_y + 12000000))
    msp.add_line((base_x + 26500000, base_y + 12000000), (base_x + 26500000, base_y + 19500000))
    #受け口中
    msp.add_line((base_x + 9500000, base_y + 14920000), (base_x + 29500000, base_y + 14920000))
    msp.add_line((base_x + 9500000, base_y + 16590000), (base_x + 29500000, base_y + 16590000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))  

#NS二受丁字管BA
def ns_niuke_tyouji_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #上受け口
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 1000000, base_y + 12000000))
    msp.add_line((base_x + 1000000, base_y + 12000000), (base_x + 1000000, base_y + 19500000))
    #下受け口
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 15000000, base_y + 12000000))
    msp.add_line((base_x + 15000000, base_y + 12000000), (base_x + 15000000, base_y + 19500000))
    #受け口中
    msp.add_line((base_x -2000000, base_y + 14920000), (base_x + 18000000, base_y + 14920000))
    msp.add_line((base_x -2000000, base_y + 16590000), (base_x + 18000000, base_y + 16590000))
    #受け口円
    msp.add_circle(center=(base_x + 8000000, base_y + 15750000), radius=2500000)
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口円
    msp.add_circle(center=(base_x + 23750000, base_y), radius=2500000)
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))

#NSF付丁字管AB
def ns_f_tyouji_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #上直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 19500000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 14500000, base_y + 5000000), (base_x + 24500000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27500000, base_y))  

#NSF付丁字管BA
def ns_f_tyouji_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 3000000, base_y + 5000000), (base_x + 13000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口上
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y + 7000000))
    msp.add_line((base_x + 20000000, base_y + 7000000), (base_x + 27500000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 13000000, base_y), (base_x + 20000000, base_y - 7000000))
    msp.add_line((base_x + 20000000, base_y - 7000000), (base_x + 27500000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 22950000, base_y + 10000000), (base_x + 22950000, base_y - 10000000))
    msp.add_line((base_x + 24650000, base_y + 10000000), (base_x + 24650000, base_y - 10000000))

#NS継ぎ輪AA
def ns_tugiwa_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直上
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 20000000, base_y + 3500000))
    #直下
    msp.add_line((base_x + 11000000, base_y - 3500000), (base_x + 20000000, base_y - 3500000))
    #受け口上
    msp.add_line((base_x + 16500000, base_y), (base_x + 23500000, base_y + 7000000))
    msp.add_line((base_x + 23500000, base_y + 7000000), (base_x + 31000000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 16500000, base_y), (base_x + 23500000, base_y - 7000000))
    msp.add_line((base_x + 23500000, base_y - 7000000), (base_x + 31000000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 26450000, base_y + 10000000), (base_x + 26450000, base_y - 10000000))
    msp.add_line((base_x + 28150000, base_y + 10000000), (base_x + 28150000, base_y - 10000000))

#NS受挿し片落管AB
def ns_ukeoshikataochi_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直上
    msp.add_line((base_x + 9470000, base_y + 5030000), (base_x + 23800000, base_y))
    #直下
    msp.add_line((base_x + 9470000, base_y - 5030000), (base_x + 23800000, base_y))

#NS受挿し片落管BA
def ns_ukeoshikataochi_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直上
    msp.add_line((base_x, base_y), (base_x + 14330000, base_y + 5030000))
    #直下
    msp.add_line((base_x, base_y), (base_x + 14330000, base_y - 5030000))
    #受け口上
    msp.add_line((base_x + 9300000, base_y), (base_x + 16300000, base_y + 7000000))
    msp.add_line((base_x + 16300000, base_y + 7000000), (base_x + 23800000, base_y + 7000000))
    #受け口下
    msp.add_line((base_x + 9300000, base_y), (base_x + 16300000, base_y - 7000000))
    msp.add_line((base_x + 16300000, base_y - 7000000), (base_x + 23800000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 19250000, base_y + 10000000), (base_x + 19250000, base_y - 10000000))
    msp.add_line((base_x + 20950000, base_y + 10000000), (base_x + 20950000, base_y - 10000000))

#NSVB両受ベンド45°AA（下がりはじめ）
def ns_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #受け口円
    msp.add_circle(center=(base_x + 3750000, base_y), radius=2500000)
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 23000000, base_y - 3500000))
    #上受け口
    msp.add_line((base_x + 23000000, base_y - 3500000), (base_x + 32900000, base_y - 3500000))
    msp.add_line((base_x + 32900000, base_y - 3500000), (base_x + 38200000, base_y - 8800000))
    #下受け口
    msp.add_line((base_x + 23000000, base_y - 3500000), (base_x + 23000000, base_y - 13400000))
    msp.add_line((base_x + 23000000, base_y - 13400000), (base_x + 28300000, base_y - 18700000))
    #受け口中
    msp.add_line((base_x + 23000000, base_y - 17600000), (base_x + 37100000, base_y - 3500000))
    msp.add_line((base_x + 24200000, base_y - 18800000), (base_x + 38300000, base_y - 4700000))

#NSVB両受ベンド45°AA（下がり終わり）
def ns_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 13790000, base_y - 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y - 13790000), (base_x + 18790000, base_y - 13790000))
    #受け口上
    msp.add_line((base_x + 18790000, base_y - 13790000), (base_x + 25790000, base_y - 6790000))
    msp.add_line((base_x + 25790000, base_y - 6790000), (base_x + 33290000, base_y - 6790000))
    #受け口下
    msp.add_line((base_x + 18790000, base_y - 13790000), (base_x + 25790000, base_y - 20790000))
    msp.add_line((base_x + 25790000, base_y - 20790000), (base_x + 33290000, base_y - 20790000))
    #受け口中
    msp.add_line((base_x + 28740000, base_y - 3790000), (base_x + 28740000, base_y - 23790000))
    msp.add_line((base_x + 30440000, base_y - 3790000), (base_x + 30440000, base_y - 23790000))

#NSVB両受ベンド45°AA（上がりはじめ）
def ns_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 23000000, base_y + 3500000))
    #上受け口
    msp.add_line((base_x + 23000000, base_y + 3500000), (base_x + 23000000, base_y + 13400000))
    msp.add_line((base_x + 23000000, base_y + 13400000), (base_x + 28300000, base_y + 18700000))
    #下受け口
    msp.add_line((base_x + 23000000, base_y + 3500000), (base_x + 32900000, base_y + 3500000))
    msp.add_line((base_x + 32900000, base_y + 3500000), (base_x + 38200000, base_y + 8800000))
    #受け口中
    msp.add_line((base_x + 23000000, base_y + 17600000), (base_x + 37100000, base_y + 3500000))
    msp.add_line((base_x + 24200000, base_y + 18800000), (base_x + 38300000, base_y + 4700000))

#NSVB両受ベンド45°AA（上がり終わり）
def ns_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y +350000))
    msp.add_line((base_x + 10250000, base_y +350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 13790000, base_y + 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y + 13790000), (base_x + 18790000, base_y + 13790000))
    #受け口上
    msp.add_line((base_x + 18790000, base_y + 13790000), (base_x + 25790000, base_y + 6790000))
    msp.add_line((base_x + 25790000, base_y + 6790000), (base_x + 33290000, base_y + 6790000))
    #受け口下
    msp.add_line((base_x + 18790000, base_y + 13790000), (base_x + 25790000, base_y + 20790000))
    msp.add_line((base_x + 25790000, base_y + 20790000), (base_x + 33290000, base_y + 20790000))
    #受け口中
    msp.add_line((base_x + 28740000, base_y + 3790000), (base_x + 28740000, base_y + 23790000))
    msp.add_line((base_x + 30440000, base_y + 3790000), (base_x + 30440000, base_y + 23790000))

#NSVB両受ベンド22.5°AA（下がりはじめ）
def ns_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24110000, base_y - 1910000))
    #上受け口
    msp.add_line((base_x + 24110000, base_y - 1910000), (base_x + 33260000, base_y + 1880000))
    msp.add_line((base_x + 33260000, base_y + 1880000), (base_x + 40190000, base_y - 990000))
    #下受け口
    msp.add_line((base_x + 24110000, base_y - 1910000), (base_x + 27900000, base_y - 11060000))
    msp.add_line((base_x + 27900000, base_y - 11060000), (base_x + 34830000, base_y - 13930000))
    #受け口中
    msp.add_line((base_x + 29480000, base_y - 14960000), (base_x + 37130000, base_y + 3520000))
    msp.add_line((base_x + 31050000, base_y - 15610000), (base_x + 38700000, base_y + 2870000))

#NSVB両受ベンド22.5°AA（下がり終わり）
def ns_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), (base_x + 18020000, base_y - 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y - 7460000), (base_x + 23020000, base_y - 7460000))
    #受け口上
    msp.add_line((base_x + 23020000, base_y - 7460000), (base_x + 30020000, base_y - 460000))
    msp.add_line((base_x + 30020000, base_y - 460000), (base_x + 37520000, base_y - 460000))
    #受け口下
    msp.add_line((base_x + 23020000, base_y - 7460000), (base_x + 30020000, base_y - 14460000))
    msp.add_line((base_x + 30020000, base_y - 14460000), (base_x + 37520000, base_y - 14460000))
    #受け口中
    msp.add_line((base_x + 32970000, base_y + 2540000), (base_x + 32970000, base_y - 17460000))
    msp.add_line((base_x + 34670000, base_y + 2540000), (base_x + 34670000, base_y - 17460000))

#NSVB両受ベンド22.5°AA（上がりはじめ）
def ns_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24110000, base_y + 1910000))
    #上受け口
    msp.add_line((base_x + 24110000, base_y + 1910000), (base_x + 27900000, base_y + 11060000))
    msp.add_line((base_x + 27900000, base_y + 11060000), (base_x + 34830000, base_y + 13930000))
    #下受け口
    msp.add_line((base_x + 24110000, base_y + 1910000), (base_x + 33260000, base_y - 1880000))
    msp.add_line((base_x + 33260000, base_y - 1880000), (base_x + 40190000, base_y + 990000))
    #受け口中
    msp.add_line((base_x + 29480000, base_y + 14960000), (base_x + 37130000, base_y - 3520000))
    msp.add_line((base_x + 31050000, base_y + 15610000), (base_x + 38700000, base_y - 2870000))

#NSVB両受ベンド22.5°AA（上がり終わり）
def ns_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), (base_x + 18020000, base_y + 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y + 7460000), (base_x + 23020000, base_y + 7460000))
    #受け口上
    msp.add_line((base_x + 23020000, base_y + 7460000), (base_x + 30020000, base_y + 460000))
    msp.add_line((base_x + 30020000, base_y + 460000), (base_x + 37520000, base_y + 460000))
    #受け口下
    msp.add_line((base_x + 23020000, base_y + 7460000), (base_x + 30020000, base_y + 14460000))
    msp.add_line((base_x + 30020000, base_y + 14460000), (base_x + 37520000, base_y + 14460000))
    #受け口中
    msp.add_line((base_x + 32970000, base_y - 2540000), (base_x + 32970000, base_y + 17460000))
    msp.add_line((base_x + 34670000, base_y - 2540000), (base_x + 34670000, base_y + 17460000))

#NSVB両受ベンド11.25°AA（下がりはじめ）
def ns_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24400000, base_y - 980000))
    #上受け口
    msp.add_line((base_x + 24400000, base_y - 980000), (base_x + 32630000, base_y + 4520000))
    msp.add_line((base_x + 32630000, base_y + 4520000), (base_x + 39990000, base_y + 3060000))
    #下受け口
    msp.add_line((base_x + 24400000, base_y - 980000), (base_x + 29900000, base_y - 9210000))
    msp.add_line((base_x + 29900000, base_y - 9210000), (base_x + 37260000, base_y - 10670000))
    #受け口中
    msp.add_line((base_x + 32210000, base_y - 12730000), (base_x + 36110000, base_y + 6890000))
    msp.add_line((base_x + 33860000, base_y - 13060000), (base_x + 37780000, base_y + 6580000))

#NSVB両受ベンド11.25°AA（下がり終わり）
def ns_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), (base_x + 19130000, base_y - 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y - 3800000), (base_x + 24130000, base_y - 3800000))
    #受け口上
    msp.add_line((base_x + 24130000, base_y - 3800000), (base_x + 31130000, base_y + 3200000))
    msp.add_line((base_x + 31130000, base_y + 3200000), (base_x + 38630000, base_y + 3200000))
    #受け口下
    msp.add_line((base_x + 24130000, base_y - 3800000), (base_x + 31130000, base_y - 10800000))
    msp.add_line((base_x + 31130000, base_y - 10800000), (base_x + 38630000, base_y - 10800000))
    #受け口中
    msp.add_line((base_x + 34080000, base_y + 6200000), (base_x + 34080000, base_y - 13800000))
    msp.add_line((base_x + 35780000, base_y + 6200000), (base_x + 35780000, base_y - 13800000))

#NSVB両受ベンド11.25°AA（上がりはじめ）
def ns_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 24400000, base_y + 980000))
    #下受け口
    msp.add_line((base_x + 24400000, base_y + 980000), (base_x + 29900000, base_y + 9210000))
    msp.add_line((base_x + 29900000, base_y + 9210000), (base_x + 37260000, base_y + 10670000))
    #上受け口
    msp.add_line((base_x + 24400000, base_y + 980000), (base_x + 32630000, base_y - 4520000))
    msp.add_line((base_x + 32630000, base_y - 4520000), (base_x + 39990000, base_y - 3060000))
    #受け口中
    msp.add_line((base_x + 32210000, base_y + 12730000), (base_x + 36110000, base_y - 6890000))
    msp.add_line((base_x + 33860000, base_y + 13060000), (base_x + 37780000, base_y - 6580000))

#NSVB両受ベンド11.25°AA（上がり終わり）
def ns_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), (base_x + 19130000, base_y + 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y + 3800000), (base_x + 24130000, base_y + 3800000))
    #受け口上
    msp.add_line((base_x + 24130000, base_y + 3800000), (base_x + 31130000, base_y - 3200000))
    msp.add_line((base_x + 31130000, base_y - 3200000), (base_x + 38630000, base_y - 3200000))
    #受け口下
    msp.add_line((base_x + 24130000, base_y + 3800000), (base_x + 31130000, base_y + 10800000))
    msp.add_line((base_x + 31130000, base_y + 10800000), (base_x + 38630000, base_y + 10800000))
    #受け口中
    msp.add_line((base_x + 34080000, base_y - 6200000), (base_x + 34080000, base_y + 13800000))
    msp.add_line((base_x + 35780000, base_y - 6200000), (base_x + 35780000, base_y + 13800000))

#NSVB片受ベンド45°AB（下がりはじめ）
def ns_vb_katauke_bend_45_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 25160000, base_y - 5660000))

#NSVB片受ベンド45°AB（下がり終わり）
def ns_vb_katauke_bend_45_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 13790000, base_y - 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y - 13790000), (base_x + 21790000, base_y - 13790000))

#NSVB片受ベンド45°BA（下がりはじめ）
def ns_vb_katauke_bend_45_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y - 3500000))
    #上受け口
    msp.add_line((base_x + 11540000, base_y - 3500000), (base_x + 21440000, base_y - 3500000))
    msp.add_line((base_x + 21440000, base_y - 3500000), (base_x + 26740000, base_y - 8800000))
    #下受け口
    msp.add_line((base_x + 11540000, base_y - 3500000), (base_x + 11540000, base_y - 13400000))
    msp.add_line((base_x + 11540000, base_y - 13400000), (base_x + 16840000, base_y - 18700000))
    #受け口中
    msp.add_line((base_x + 11540000, base_y - 17600000), (base_x + 25640000, base_y - 3500000))
    msp.add_line((base_x + 12740000, base_y - 18800000), (base_x + 26840000, base_y - 4700000))

#NSVB片受ベンド45°BA（下がり終わり）
def ns_vb_katauke_bend_45_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y - 5660000), (base_x + 10660000, base_y - 5660000))
    #受け口上
    msp.add_line((base_x + 10660000, base_y - 5660000), (base_x + 17660000, base_y + 1340000))
    msp.add_line((base_x + 17660000, base_y + 1340000), (base_x + 25160000, base_y + 1340000))
    #受け口下
    msp.add_line((base_x + 10660000, base_y - 5660000), (base_x + 17660000, base_y - 12660000))
    msp.add_line((base_x + 17660000, base_y - 12660000), (base_x + 25160000, base_y - 12660000))
    #受け口中
    msp.add_line((base_x + 20610000, base_y + 4340000), (base_x + 20610000, base_y - 15660000))
    msp.add_line((base_x + 22310000, base_y + 4340000), (base_x + 22310000, base_y - 15660000))


#nsVB片受ベンド45°AB（上がりはじめ）
def ns_vb_katauke_bend_45_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 25160000, base_y + 5660000))

#NSVB片受ベンド45°AB（上がり終わり）
def ns_vb_katauke_bend_45_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y +350000))
    msp.add_line((base_x + 10250000, base_y + 350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), (base_x + 13790000, base_y + 13790000))
    #直部分
    msp.add_line((base_x + 13790000, base_y + 13790000), (base_x + 21790000, base_y + 13790000))

#NSVB片受ベンド45°BA（上がりはじめ）
def ns_vb_katauke_bend_45_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y + 3500000))
    #上受け口
    msp.add_line((base_x + 11540000, base_y + 3500000), (base_x + 21440000, base_y + 3500000))
    msp.add_line((base_x + 21440000, base_y + 3500000), (base_x + 26740000, base_y + 8800000))
    #下受け口
    msp.add_line((base_x + 11540000, base_y + 3500000), (base_x + 11540000, base_y + 13400000))
    msp.add_line((base_x + 11540000, base_y + 13400000), (base_x + 16840000, base_y + 18700000))
    #受け口中
    msp.add_line((base_x + 11540000, base_y + 17600000), (base_x + 25640000, base_y + 3500000))
    msp.add_line((base_x + 12740000, base_y + 18800000), (base_x + 26840000, base_y + 4700000))

#NSVB片受ベンド45°BA（上がり終わり）
def ns_vb_katauke_bend_45_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y + 5660000), (base_x + 10660000, base_y + 5660000))
    #受け口上
    msp.add_line((base_x + 10660000, base_y + 5660000), (base_x + 17660000, base_y - 1340000))
    msp.add_line((base_x + 17660000, base_y - 1340000), (base_x + 25160000, base_y - 1340000))
    #受け口下
    msp.add_line((base_x + 10660000, base_y + 5660000), (base_x + 17660000, base_y + 12660000))
    msp.add_line((base_x + 17660000, base_y + 12660000), (base_x + 25160000, base_y + 12660000))
    #受け口中
    msp.add_line((base_x + 20610000, base_y - 4340000), (base_x + 20610000, base_y + 15660000))
    msp.add_line((base_x + 22310000, base_y - 4340000), (base_x + 22310000, base_y + 15660000))

#NSVB片受ベンド22.5°AB（下がりはじめ）
def ns_vb_katauke_bend_22_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 26870000, base_y - 3060000))

#NSVB片受ベンド22.5°AB（下がり終わり）
def ns_vb_katauke_bend_22_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), (base_x + 18020000, base_y - 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y - 7460000), (base_x + 26020000, base_y - 7460000))

#NSVB片受ベンド22.5°BA（下がりはじめ）
def ns_vb_katauke_bend_22_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12610000, base_y - 1910000))
    #上受け口
    msp.add_line((base_x + 12610000, base_y - 1910000), (base_x + 21760000, base_y + 1880000))
    msp.add_line((base_x + 21760000, base_y + 1880000), (base_x + 28690000, base_y - 990000))
    #下受け口
    msp.add_line((base_x + 12610000, base_y - 1910000), (base_x + 16400000, base_y - 11060000))
    msp.add_line((base_x + 16400000, base_y - 11060000), (base_x + 23330000, base_y - 13930000))
    #受け口中
    msp.add_line((base_x + 17980000, base_y - 14960000), (base_x + 25630000, base_y + 3520000))
    msp.add_line((base_x + 19550000, base_y - 15610000), (base_x + 27200000, base_y + 2870000))

#NSVB片受ベンド22.5°BA（下がり終わり）
def ns_vb_katauke_bend_22_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7370000, base_y - 3060000))
    #直部分
    msp.add_line((base_x + 7370000, base_y - 3060000), (base_x + 12370000, base_y - 3060000))
    #受け口上
    msp.add_line((base_x + 12370000, base_y - 3060000), (base_x + 19370000, base_y + 3940000))
    msp.add_line((base_x + 19370000, base_y + 3940000), (base_x + 26870000, base_y + 3940000))
    #受け口下
    msp.add_line((base_x + 12370000, base_y - 3060000), (base_x + 19370000, base_y - 10060000))
    msp.add_line((base_x + 19370000, base_y - 10060000), (base_x + 26870000, base_y - 10060000))
    #受け口中
    msp.add_line((base_x + 22320000, base_y + 6940000), (base_x + 22320000, base_y - 13060000))
    msp.add_line((base_x + 24020000, base_y + 6940000), (base_x + 24020000, base_y - 13060000))

#NSVB片受ベンド22.5°AB（上がりはじめ）
def ns_vb_katauke_bend_22_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 26870000, base_y + 3060000))

#NSVB片受ベンド22.5°AB（上がり終わり）
def ns_vb_katauke_bend_22_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), (base_x + 18020000, base_y + 7460000))
    #直部分
    msp.add_line((base_x + 18020000, base_y + 7460000), (base_x + 26020000, base_y + 7460000))

#NSVB片受ベンド22.5°BA（上がりはじめ）
def ns_vb_katauke_bend_22_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12610000, base_y + 1910000))
    #上受け口
    msp.add_line((base_x + 12610000, base_y + 1910000), (base_x + 21760000, base_y - 1880000))
    msp.add_line((base_x + 21760000, base_y - 1880000), (base_x + 28690000, base_y + 990000))
    #下受け口
    msp.add_line((base_x + 12610000, base_y + 1910000), (base_x + 16400000, base_y + 11060000))
    msp.add_line((base_x + 16400000, base_y + 11060000), (base_x + 23330000, base_y + 13930000))
    #受け口中
    msp.add_line((base_x + 17980000, base_y + 14960000), (base_x + 25630000, base_y - 3520000))
    msp.add_line((base_x + 19550000, base_y + 15610000), (base_x + 27200000, base_y - 2870000))

#NSVB片受ベンド22.5°BA（上がり終わり）
def ns_vb_katauke_bend_22_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7370000, base_y + 3060000))
    #直部分
    msp.add_line((base_x + 7370000, base_y + 3060000), (base_x + 12370000, base_y + 3060000))
    #受け口上
    msp.add_line((base_x + 12370000, base_y + 3060000), (base_x + 19370000, base_y - 3940000))
    msp.add_line((base_x + 19370000, base_y - 3940000), (base_x + 26870000, base_y - 3940000))
    #受け口下
    msp.add_line((base_x + 12370000, base_y + 3060000), (base_x + 19370000, base_y + 10060000))
    msp.add_line((base_x + 19370000, base_y + 10060000), (base_x + 26870000, base_y + 10060000))
    #受け口中
    msp.add_line((base_x + 22320000, base_y - 6940000), (base_x + 22320000, base_y + 13060000))
    msp.add_line((base_x + 24020000, base_y - 6940000), (base_x + 24020000, base_y + 13060000))

#NSVB片受ベンド11.25°AB（下がりはじめ）
def ns_vb_katauke_bend_11_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27340000, base_y - 1570000))

#NSVB片受ベンド11.25°AB（下がり終わり）
def ns_vb_katauke_bend_11_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), (base_x + 19130000, base_y - 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y - 3800000), (base_x + 27130000, base_y - 3800000))

#NSVB片受ベンド11.25°BA（下がりはじめ）
def ns_vb_katauke_bend_11_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y - 980000))
    #上受け口
    msp.add_line((base_x + 12900000, base_y - 980000), (base_x + 21130000, base_y + 4520000))
    msp.add_line((base_x + 21130000, base_y + 4520000), (base_x + 28490000, base_y + 3060000))
    #下受け口
    msp.add_line((base_x + 12900000, base_y - 980000), (base_x + 18400000, base_y - 9210000))
    msp.add_line((base_x + 18400000, base_y - 9210000), (base_x + 25760000, base_y - 10670000))
    #受け口中
    msp.add_line((base_x + 20710000, base_y - 12730000), (base_x + 24610000, base_y + 6890000))
    msp.add_line((base_x + 22360000, base_y - 13060000), (base_x + 26280000, base_y + 6580000))

#NSVB片受ベンド11.25°BA（下がり終わり）
def ns_vb_katauke_bend_11_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y - 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y - 1570000), (base_x + 12840000, base_y - 1570000))
    #受け口上
    msp.add_line((base_x + 12840000, base_y - 1570000), (base_x + 19840000, base_y + 5430000))
    msp.add_line((base_x + 19840000, base_y + 5430000), (base_x + 27340000, base_y + 5430000))
    #受け口下
    msp.add_line((base_x + 12840000, base_y - 1570000), (base_x + 19840000, base_y - 8570000))
    msp.add_line((base_x + 19840000, base_y - 8570000), (base_x + 27340000, base_y - 8570000))
    #受け口中
    msp.add_line((base_x + 22790000, base_y + 8430000), (base_x + 22790000, base_y - 11570000))
    msp.add_line((base_x + 24490000, base_y + 8430000), (base_x + 24490000, base_y - 11270000))

#NSVB片受ベンド11.25°AB（上がりはじめ）
def ns_vb_katauke_bend_11_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 19500000, base_y))
    #直部分
    msp.add_line((base_x + 19500000, base_y), (base_x + 27340000, base_y + 1570000))

#NSVB片受ベンド11.25°AB（上がり終わり）
def ns_vb_katauke_bend_11_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), (base_x + 19130000, base_y + 3800000))
    #直部分
    msp.add_line((base_x + 19130000, base_y + 3800000), (base_x + 27130000, base_y + 3800000))

#NSVB片受ベンド11.25°BA（上がりはじめ）
def ns_vb_katauke_bend_11_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y + 980000))
    #上受け口
    msp.add_line((base_x + 12900000, base_y + 980000), (base_x + 21130000, base_y - 4520000))
    msp.add_line((base_x + 21130000, base_y - 4520000), (base_x + 28490000, base_y - 3060000))
    #下受け口
    msp.add_line((base_x + 12900000, base_y + 980000), (base_x + 18400000, base_y + 9210000))
    msp.add_line((base_x + 18400000, base_y + 9210000), (base_x + 25760000, base_y + 10670000))
    #受け口中
    msp.add_line((base_x + 20710000, base_y + 12730000), (base_x + 24610000, base_y - 6890000))
    msp.add_line((base_x + 22360000, base_y + 13060000), (base_x + 26280000, base_y - 6580000))

#NSVB片受ベンド11.25°BA（上がり終わり）
def ns_vb_katauke_bend_11_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y + 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y + 1570000), (base_x + 12840000, base_y + 1570000))
    #受け口上
    msp.add_line((base_x + 12840000, base_y + 1570000), (base_x + 19840000, base_y - 5430000))
    msp.add_line((base_x + 19840000, base_y - 5430000), (base_x + 27340000, base_y - 5430000))
    #受け口下
    msp.add_line((base_x + 12840000, base_y + 1570000), (base_x + 19840000, base_y + 8570000))
    msp.add_line((base_x + 19840000, base_y + 8570000), (base_x + 27340000, base_y + 8570000))
    #受け口中
    msp.add_line((base_x + 22790000, base_y - 8430000), (base_x + 22790000, base_y + 11570000))
    msp.add_line((base_x + 24490000, base_y - 8430000), (base_x + 24490000, base_y + 11570000))

#NS直管受け口BA
def ns_tyokkann_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 25250000, base_y))
    #上受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y + 7000000))
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 39750000, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 32250000, base_y - 7000000), (base_x + 39750000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 35100000, base_y + 10000000), (base_x + 35100000, base_y - 10000000))
    msp.add_line((base_x + 36700000, base_y + 10000000), (base_x + 36700000, base_y - 10000000))

#NS直管受け口AB
def ns_tyokkann_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))

    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))

    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))

    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 39750000, base_y))

#NS直管受け口ﾗｲﾅAB
def ns_tyokkann_raina_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 39750000, base_y))
    #ライナ部分
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 18170000, base_y + 9360000))
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 14500000, base_y + 3000000))
    msp.add_line((base_x + 14500000, base_y + 3000000), (base_x + 18170000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 8000000, base_y + 6500000), (base_x + 8000000, base_y - 6500000))
    msp.add_line((base_x + 8500000, base_y + 6000000), (base_x + 8500000, base_y - 6000000))
    msp.add_line((base_x + 9000000, base_y + 5500000), (base_x + 9000000, base_y - 5500000))
    msp.add_line((base_x + 9500000, base_y + 5000000), (base_x + 9500000, base_y - 5000000))
    msp.add_line((base_x + 10000000, base_y + 4500000), (base_x + 10000000, base_y - 4500000))
    msp.add_line((base_x + 10500000, base_y + 4000000), (base_x + 10500000, base_y - 4000000))
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 11000000, base_y - 3500000))
    msp.add_line((base_x + 11500000, base_y + 3000000), (base_x + 11500000, base_y - 3000000))
    msp.add_line((base_x + 12000000, base_y + 2500000), (base_x + 12000000, base_y - 2500000))
    msp.add_line((base_x + 12500000, base_y + 2000000), (base_x + 12500000, base_y - 2000000))
    msp.add_line((base_x + 13000000, base_y + 1500000), (base_x + 13000000, base_y - 1500000))
    msp.add_line((base_x + 13500000, base_y + 1000000), (base_x + 13500000, base_y - 1000000))
    msp.add_line((base_x + 14000000, base_y + 500000), (base_x + 14000000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 9790000, base_y + 10320000),
        'height': 6000000,
    })

#NS直管受け口ﾗｲﾅBA
def ns_tyokkann_raina_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 25250000, base_y))
    #上受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y + 7000000))
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 39750000, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + 25250000, base_y), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 32250000, base_y - 7000000), (base_x + 39750000, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 35100000, base_y + 10000000), (base_x + 35100000, base_y - 10000000))
    msp.add_line((base_x + 36700000, base_y + 10000000), (base_x + 36700000, base_y - 10000000))
    #ライナ部分
    msp.add_line((base_x + 21580000, base_y + 9360000), (base_x + 28920000, base_y + 9360000))
    msp.add_line((base_x + 21580000, base_y + 9360000), (base_x + 25250000, base_y + 3000000))
    msp.add_line((base_x + 25250000, base_y + 3000000), (base_x + 28920000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 32250000, base_y + 7000000), (base_x + 32250000, base_y - 7000000))
    msp.add_line((base_x + 31750000, base_y + 6500000), (base_x + 31750000, base_y - 6500000))
    msp.add_line((base_x + 31250000, base_y + 6000000), (base_x + 31250000, base_y - 6000000))
    msp.add_line((base_x + 30750000, base_y + 5500000), (base_x + 30750000, base_y - 5500000))
    msp.add_line((base_x + 30250000, base_y + 5000000), (base_x + 30250000, base_y - 5000000))
    msp.add_line((base_x + 29750000, base_y + 4500000), (base_x + 29750000, base_y - 4500000))
    msp.add_line((base_x + 29250000, base_y + 4000000), (base_x + 29250000, base_y - 4000000))
    msp.add_line((base_x + 28750000, base_y + 3500000), (base_x + 28750000, base_y - 3500000))
    msp.add_line((base_x + 28250000, base_y + 3000000), (base_x + 28250000, base_y - 3000000))
    msp.add_line((base_x + 27750000, base_y + 2500000), (base_x + 27750000, base_y - 2500000))
    msp.add_line((base_x + 27250000, base_y + 2000000), (base_x + 27250000, base_y - 2000000))
    msp.add_line((base_x + 26750000, base_y + 1500000), (base_x + 26750000, base_y - 1500000))
    msp.add_line((base_x + 26250000, base_y + 1000000), (base_x + 26250000, base_y - 1000000))
    msp.add_line((base_x + 25750000, base_y + 500000), (base_x + 25750000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 9790000, base_y + 10320000),
        'height': 6000000,
    })

#NS直管受け口甲AB
def ns_tyokkann_kou_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2850000, base_y + 10000000), (base_x + 2850000, base_y - 10000000))
    msp.add_line((base_x + 4550000, base_y + 10000000), (base_x + 4550000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y))

#NS直管受け口甲BA
def ns_tyokkann_kou_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))
    #上受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 9850000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 9850000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    msp.add_line((base_x + 11450000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 11450000 + item['配管長'] * 10000 / 2, base_y - 10000000))

#NS直管受け口ﾗｲﾅ甲AB
def ns_tyokkann_raina_kou_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x, base_y + 7000000), (base_x + 7500000, base_y + 7000000))
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 14500000, base_y))
    #下受け口
    msp.add_line((base_x, base_y - 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 7500000, base_y - 7000000), (base_x + 14500000, base_y))
    #受け口中
    msp.add_line((base_x + 2800000, base_y + 10000000), (base_x + 2800000, base_y - 10000000))
    msp.add_line((base_x + 4400000, base_y + 10000000), (base_x + 4400000, base_y - 10000000))
    #直部分
    msp.add_line((base_x + 14500000, base_y), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y))
    #ライナ部分
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 18170000, base_y + 9360000))
    msp.add_line((base_x + 10830000, base_y + 9360000), (base_x + 14500000, base_y + 3000000))
    msp.add_line((base_x + 14500000, base_y + 3000000), (base_x + 18170000, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7500000, base_y + 7000000), (base_x + 7500000, base_y - 7000000))
    msp.add_line((base_x + 8000000, base_y + 6500000), (base_x + 8000000, base_y - 6500000))
    msp.add_line((base_x + 8500000, base_y + 6000000), (base_x + 8500000, base_y - 6000000))
    msp.add_line((base_x + 9000000, base_y + 5500000), (base_x + 9000000, base_y - 5500000))
    msp.add_line((base_x + 9500000, base_y + 5000000), (base_x + 9500000, base_y - 5000000))
    msp.add_line((base_x + 10000000, base_y + 4500000), (base_x + 10000000, base_y - 4500000))
    msp.add_line((base_x + 10500000, base_y + 4000000), (base_x + 10500000, base_y - 4000000))
    msp.add_line((base_x + 11000000, base_y + 3500000), (base_x + 11000000, base_y - 3500000))
    msp.add_line((base_x + 11500000, base_y + 3000000), (base_x + 11500000, base_y - 3000000))
    msp.add_line((base_x + 12000000, base_y + 2500000), (base_x + 12000000, base_y - 2500000))
    msp.add_line((base_x + 12500000, base_y + 2000000), (base_x + 12500000, base_y - 2000000))
    msp.add_line((base_x + 13000000, base_y + 1500000), (base_x + 13000000, base_y - 1500000))
    msp.add_line((base_x + 13500000, base_y + 1000000), (base_x + 13500000, base_y - 1000000))
    msp.add_line((base_x + 14000000, base_y + 500000), (base_x + 14000000, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x + 9790000, base_y + 10320000),
        'height': 6000000,
    })

#NS直管受け口ﾗｲﾅ甲BA
def ns_tyokkann_raina_kou_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))
    #上受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y + 7000000))
    #下受け口
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000), (base_x + 14500000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    #受け口中
    msp.add_line((base_x + 9850000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 9850000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    msp.add_line((base_x + 11450000 + item['配管長'] * 10000 / 2, base_y + 10000000), (base_x + 11450000 + item['配管長'] * 10000 / 2, base_y - 10000000))
    #ライナ部分
    msp.add_line((base_x - 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000), (base_x + 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000))
    msp.add_line((base_x - 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000), (base_x + item['配管長'] * 10000 / 2, base_y + 3000000))
    msp.add_line((base_x + item['配管長'] * 10000 / 2, base_y + 3000000), (base_x + 3670000 + item['配管長'] * 10000 / 2, base_y + 9360000))
    #ライナ塗りつぶし
    msp.add_line((base_x + 7000000 + item['配管長'] * 10000 / 2, base_y + 7000000), (base_x + 7000000 + item['配管長'] * 10000 / 2, base_y - 7000000))
    msp.add_line((base_x + 6500000 + item['配管長'] * 10000 / 2, base_y + 6500000), (base_x + 6500000 + item['配管長'] * 10000 / 2, base_y - 6500000))
    msp.add_line((base_x + 6000000 + item['配管長'] * 10000 / 2, base_y + 6000000), (base_x + 6000000 + item['配管長'] * 10000 / 2, base_y - 6000000))
    msp.add_line((base_x + 5500000 + item['配管長'] * 10000 / 2, base_y + 5500000), (base_x + 5500000 + item['配管長'] * 10000 / 2, base_y - 5500000))
    msp.add_line((base_x + 5000000 + item['配管長'] * 10000 / 2, base_y + 5000000), (base_x + 5000000 + item['配管長'] * 10000 / 2, base_y - 5000000))
    msp.add_line((base_x + 4500000 + item['配管長'] * 10000 / 2, base_y + 4500000), (base_x + 4500000 + item['配管長'] * 10000 / 2, base_y - 4500000))
    msp.add_line((base_x + 4000000 + item['配管長'] * 10000 / 2, base_y + 4000000), (base_x + 4000000 + item['配管長'] * 10000 / 2, base_y - 4000000))
    msp.add_line((base_x + 3500000 + item['配管長'] * 10000 / 2, base_y + 3500000), (base_x + 3500000 + item['配管長'] * 10000 / 2, base_y - 3500000))
    msp.add_line((base_x + 3000000 + item['配管長'] * 10000 / 2, base_y + 3000000), (base_x + 3000000 + item['配管長'] * 10000 / 2, base_y - 3000000))
    msp.add_line((base_x + 2500000 + item['配管長'] * 10000 / 2, base_y + 2500000), (base_x + 2500000 + item['配管長'] * 10000 / 2, base_y - 2500000))
    msp.add_line((base_x + 2000000 + item['配管長'] * 10000 / 2, base_y + 2000000), (base_x + 2000000 + item['配管長'] * 10000 / 2, base_y - 2000000))
    msp.add_line((base_x + 1500000 + item['配管長'] * 10000 / 2, base_y + 1500000), (base_x + 1500000 + item['配管長'] * 10000 / 2, base_y - 1500000))
    msp.add_line((base_x + 1000000 + item['配管長'] * 10000 / 2, base_y + 1000000), (base_x + 1000000 + item['配管長'] * 10000 / 2, base_y - 1000000))
    msp.add_line((base_x + 500000 + item['配管長'] * 10000 / 2, base_y + 500000), (base_x + 500000 + item['配管長'] * 10000 / 2, base_y - 500000))
    #ライナ文字入れる
    msp.add_text("ﾗｲﾅ", dxfattribs={
        'insert': (base_x - 4710000 + item['配管長'] * 10000 / 2, base_y + 10320000),
        'height': 6000000,
    })

#NSVB45°直管受け口甲AB（下がり）
def ns_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y + 4950000), (base_x + 10250000, base_y -350000))
    msp.add_line((base_x + 10250000, base_y -350000), (base_x + 10250000, base_y - 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y - 4950000), (base_x + 350000, base_y - 10250000))
    msp.add_line((base_x + 350000, base_y - 10250000), (base_x + 10250000, base_y - 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y - 9090000), (base_x + 9090000, base_y + 5060000))
    msp.add_line((base_x - 3850000, base_y - 10290000), (base_x + 10290000, base_y + 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y - 10250000), (base_x + 10250000 + (item['水平距離'] * 10000 / 2), base_y - 10250000 - (item['水平距離'] * 10000 / 2)))

#NSVB45°直管受け口甲AB（上がり）
def ns_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 4950000, base_y - 4950000), (base_x + 10250000, base_y + 350000))
    msp.add_line((base_x + 10250000, base_y + 350000), (base_x + 10250000, base_y + 10250000))
    #下受け口
    msp.add_line((base_x - 4950000, base_y + 4950000), (base_x + 350000, base_y + 10250000))
    msp.add_line((base_x + 350000, base_y + 10250000), (base_x + 10250000, base_y + 10250000))
    #受け口中
    msp.add_line((base_x - 5060000, base_y + 9090000), (base_x + 9090000, base_y - 5060000))
    msp.add_line((base_x - 3850000, base_y + 10290000), (base_x + 10290000, base_y - 3850000))
    #直部分
    msp.add_line((base_x + 10250000, base_y + 10250000), 
                 (base_x + 10250000 + (item['水平距離'] * 10000 / 2), 
                  base_y + 10250000 + (item['水平距離'] * 10000 / 2)))

#NSVB45°直管受け口甲BA（下がり）
def ns_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    #上受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 9900000 + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 9900000 + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15200000 + (item['水平距離'] * 10000 / 2), 
                   base_y - 5300000 - (item['水平距離'] * 10000 / 2)))
    #下受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - (item['水平距離'] * 10000 / 2)), 
                  (base_x + (item['水平距離'] * 10000 / 2), base_y - 9900000 - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - 9900000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 5300000 + (item['水平距離'] * 10000 / 2), 
                   base_y - 15200000 - (item['水平距離'] * 10000 / 2)))
    #受け口中
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - 14100000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 14100000 + (item['水平距離'] * 10000 / 2), base_y - (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 1200000 + (item['水平距離'] * 10000 / 2), 
                  base_y - 15300000 - (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15300000 + (item['水平距離'] * 10000 / 2), base_y - 1200000 - (item['水平距離'] * 10000 / 2)))

#NSVB45°直管受け口甲BA（上がり）
def ns_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    #上受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 9900000 + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 9900000 + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15200000 + (item['水平距離'] * 10000 / 2), 
                   base_y + 5300000 + (item['水平距離'] * 10000 / 2)))
    #下受け口
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + (item['水平距離'] * 10000 / 2)), 
                  (base_x + (item['水平距離'] * 10000 / 2), base_y + 9900000 + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + 9900000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 5300000 + (item['水平距離'] * 10000 / 2), 
                   base_y + 15200000 + (item['水平距離'] * 10000 / 2)))
    #受け口中
    msp.add_line((base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + 14100000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 14100000 + (item['水平距離'] * 10000 / 2), base_y + (item['水平距離'] * 10000 / 2)))
    msp.add_line((base_x + 1200000 + (item['水平距離'] * 10000 / 2), 
                  base_y + 15300000 + (item['水平距離'] * 10000 / 2)), 
                  (base_x + 15300000 + (item['水平距離'] * 10000 / 2), base_y + 1200000 + (item['水平距離'] * 10000 / 2)))

#NSVB22.5°直管受け口甲AB（下がり）
def ns_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y + 6470000), (base_x + 9610000, base_y + 3600000))
    msp.add_line((base_x + 9610000, base_y + 3600000), (base_x + 13400000, base_y - 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y - 6470000), (base_x + 4250000, base_y - 9340000))
    msp.add_line((base_x + 4250000, base_y - 9340000), (base_x + 13400000, base_y - 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y - 10330000), (base_x + 6460000, base_y + 8150000))
    msp.add_line((base_x + 380000, base_y - 10980000), (base_x + 8030000, base_y + 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y - 5550000), (base_x + 13400000 + item['水平距離'] * 10000 / 2, base_y - 5550000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB22.5°直管受け口甲AB（上がり）
def ns_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 2680000, base_y - 6470000), (base_x + 9610000, base_y - 3600000))
    msp.add_line((base_x + 9610000, base_y - 3600000), (base_x + 13400000, base_y + 5550000))
    #下受け口
    msp.add_line((base_x - 2680000, base_y + 6470000), (base_x + 4250000, base_y + 9340000))
    msp.add_line((base_x + 4250000, base_y + 9340000), (base_x + 13400000, base_y + 5550000))
    #受け口中
    msp.add_line((base_x - 1190000, base_y + 10330000), (base_x + 6460000, base_y - 8150000))
    msp.add_line((base_x + 380000, base_y + 10980000), (base_x + 8030000, base_y - 7500000))
    #直部分
    msp.add_line((base_x + 13400000, base_y + 5550000), 
                 (base_x + 13400000 + item['水平距離'] * 10000 / 2, 
                  base_y + 5550000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB22.5°直管受け口甲BA（下がり）
def ns_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                   base_y + 3790000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                  base_y + 3790000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 16080000 + item['水平距離'] * 10000 / 2, 
                   base_y + 920000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                   base_y - 9150000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                  base_y - 9150000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 10720000 + item['水平距離'] * 10000 / 2, 
                   base_y - 12020000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口円
    msp.add_circle(center=(base_x + 9930000 + item['水平距離'] * 10000 / 2, 
                           base_y - 4110000 - ((item['配管長'] * 10000 / 2) ** 2 
                                               - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), radius=2500000)
    #受け口中
    msp.add_line((base_x + 5370000 + item['水平距離'] * 10000 / 2, 
                  base_y - 13050000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13020000 + item['水平距離'] * 10000 / 2, 
                   base_y + 5430000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 6940000 + item['水平距離'] * 10000 / 2, 
                  base_y - 13700000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 14590000 + item['水平距離'] * 10000 / 2, 
                   base_y + 4780000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB22.5°直管受け口甲BA（上がり）
def ns_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000 / 2), 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                   base_y - 3790000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9150000 + item['水平距離'] * 10000 / 2, 
                  base_y - 3790000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 16080000 + item['水平距離'] * 10000 / 2, 
                   base_y - 920000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                   base_y + 9150000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 3790000 + item['水平距離'] * 10000 / 2, 
                  base_y + 9150000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 10720000 + item['水平距離'] * 10000 / 2, 
                   base_y + 12020000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口中
    msp.add_line((base_x + 5370000 + item['水平距離'] * 10000 / 2, 
                  base_y + 13050000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13020000 + item['水平距離'] * 10000 / 2, 
                   base_y - 5430000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 6940000 + item['水平距離'] * 10000 / 2, 
                  base_y + 13700000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 14590000 + item['水平距離'] * 10000 / 2, 
                   base_y - 4780000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB11.25°直管受け口甲AB（下がり）
def ns_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y + 6870000), (base_x + 8720000, base_y + 5400000))
    msp.add_line((base_x + 8720000, base_y + 5400000), (base_x + 14220000, base_y - 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y - 6870000), (base_x + 5990000, base_y - 8330000))
    msp.add_line((base_x + 5990000, base_y - 8330000), (base_x + 14220000, base_y - 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y - 10360000), (base_x + 4750000, base_y + 9250000))
    msp.add_line((base_x + 2510000, base_y - 10700000), (base_x + 6410000, base_y + 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y - 2830000), (base_x + 14220000 + item['水平距離'] * 10000 / 2, base_y - 2830000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB11.25°直管受け口甲AB（上がり）
def ns_vb_tyokkann_kou_11_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上受け口
    msp.add_line((base_x + 1370000, base_y - 6870000), (base_x + 8720000, base_y - 5400000))
    msp.add_line((base_x + 8720000, base_y - 5400000), (base_x + 14220000, base_y + 2830000))
    #下受け口
    msp.add_line((base_x - 1370000, base_y + 6870000), (base_x + 5990000, base_y + 8330000))
    msp.add_line((base_x + 5990000, base_y + 8330000), (base_x + 14220000, base_y + 2830000))
    #受け口中
    msp.add_line((base_x + 840000, base_y + 10360000), (base_x + 4750000, base_y - 9250000))
    msp.add_line((base_x + 2510000, base_y + 10700000), (base_x + 6410000, base_y - 8920000))
    #直部分
    msp.add_line((base_x + 14220000, base_y + 2830000), 
                 (base_x + 14220000 + item['水平距離'] * 10000 / 2, 
                  base_y + 2830000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB11.25°直管受け口甲BA（下がり）
def ns_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + item['水平距離'] * 10000 / 2, 
                                    base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                   base_y + 5500000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                  base_y + 5500000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 15590000 + item['水平距離'] * 10000 / 2, 
                   base_y + 4040000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                   base_y - 8230000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                  base_y - 8230000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 12860000 + item['水平距離'] * 10000 / 2, 
                   base_y - 9690000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口中
    msp.add_line((base_x + 7810000 + item['水平距離'] * 10000 / 2, 
                  base_y - 11750000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 11710000 + item['水平距離'] * 10000 / 2, 
                   base_y + 7870000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9460000 + item['水平距離'] * 10000 / 2, 
                  base_y - 12080000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13380000 + item['水平距離'] * 10000 / 2, 
                   base_y + 7560000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NSVB11.25°直管受け口甲BA（上がり）
def ns_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #上受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                   base_y - 5500000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 8230000 + item['水平距離'] * 10000 / 2, 
                  base_y - 5500000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 15590000 + item['水平距離'] * 10000 / 2, 
                   base_y - 4040000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #下受け口
    msp.add_line((base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                   base_y + 8230000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 5500000 + item['水平距離'] * 10000 / 2, 
                  base_y + 8230000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 12860000 + item['水平距離'] * 10000 / 2, 
                   base_y + 9690000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    #受け口中
    msp.add_line((base_x + 7810000 + item['水平距離'] * 10000 / 2, 
                  base_y + 11750000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 11710000 + item['水平距離'] * 10000 / 2, 
                   base_y - 7870000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))
    msp.add_line((base_x + 9460000 + item['水平距離'] * 10000 / 2, 
                  base_y + 12080000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                  (base_x + 13380000 + item['水平距離'] * 10000 / 2, 
                   base_y - 7560000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#NS直管受け口乙BB
def ns_tyokkann_otsu_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000, base_y))

#ポリエチレン関係略図
#EFフランジFA
def pe_franji_FA(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #フランジ部
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
    #直部分
    msp.add_line((base_x, base_y), (base_x + 10000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 15000000, base_y), radius=5000000, start_angle=90, end_angle=270)

#EFフランジAF
def pe_franji_AF(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 15000000, base_y))
    #フランジ部
    msp.add_line((base_x + 15000000, base_y + 5000000), (base_x + 15000000, base_y - 5000000))

#ポリエチレン関係略図
#EFフランジ短管FB
def pe_franjitann_FB(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #フランジ部
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
    #直部分
    msp.add_line((base_x, base_y), (base_x + 10000000, base_y))

#EFフランジ短管BF
def pe_franjitann_BF(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 10000000, base_y))
    #フランジ部
    msp.add_line((base_x + 10000000, base_y + 5000000), (base_x + 10000000, base_y - 5000000))

#ポリエチレンHB両受ベンドAA
def pe_hb_ryouuke_bend_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y + 3000000), (base_x + 10000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 20000000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンHB片受ベンドAB
def pe_hb_katauke_bend_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y + 3000000), (base_x + 10000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 18000000, base_y))

#ポリエチレンHB片受ベンドBA
def pe_hb_katauke_bend_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y + 3000000), (base_x + 8000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 18000000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン両受チーズAA
def pe_t_ryouuke_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #上直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 8000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 20000000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン片受チーズAB
def pe_t_katauke_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #上直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 8000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 18000000, base_y))

#ポリエチレン片受チーズBA
def pe_t_katauke_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 8000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 18000000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン両受チーズ(BE)
def pe_t_ryouuke_be(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))    
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 8000000, base_y + 10000000), radius=5000000, start_angle=180, end_angle=0)
    #下直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y - 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 8000000, base_y - 10000000), radius=5000000, start_angle=0, end_angle=180)

#ポリエチレン両受チーズ(SB)
def pe_t_ryouuke_sb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    
    #上直部分
    msp.add_line((base_x, base_y), (base_x, base_y + 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y + 10000000), radius=5000000, start_angle=180, end_angle=0)
    #下直部分
    msp.add_line((base_x, base_y), (base_x, base_y - 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y - 10000000), radius=5000000, start_angle=0, end_angle=180)
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))

#ポリエチレン両受F付チーズ(空気弁)AA
def pe_ft_ryouuke_air_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #上直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 5000000, base_y + 5000000), (base_x + 15000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 20000000, base_y), radius=5000000, start_angle=90, end_angle=270)
    #短管部分
    msp.add_line((base_x + 5000000, base_y + 6000000), (base_x + 15000000, base_y + 6000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 15000000, base_y + 11000000))
    msp.add_line((base_x + 7000000, base_y + 11000000), (base_x + 7000000, base_y + 6000000))
    msp.add_line((base_x + 13000000, base_y + 11000000), (base_x + 13000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 12000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 12000000))
    #空気弁部分
    msp.add_line((base_x + 5000000, base_y + 18000000), (base_x + 15000000, base_y + 18000000))
    msp.add_line((base_x + 10000000, base_y + 18000000), (base_x + 10000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 10000000, base_y + 25000000), radius=4000000)
    msp.add_text("A", dxfattribs={
        'insert': (base_x + 8500000, base_y + 22000000),
        'height': 6000000,
    })
#ポリエチレン両受F付チーズ(消火栓)AA
def pe_ft_ryouuke_fire_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #上直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 5000000, base_y + 5000000), (base_x + 15000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 20000000, base_y), radius=5000000, start_angle=90, end_angle=270)
    #短管部分
    msp.add_line((base_x + 5000000, base_y + 6000000), (base_x + 15000000, base_y + 6000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 15000000, base_y + 11000000))
    msp.add_line((base_x + 7000000, base_y + 11000000), (base_x + 7000000, base_y + 6000000))
    msp.add_line((base_x + 13000000, base_y + 11000000), (base_x + 13000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 12000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 12000000))
    #空気弁部分
    msp.add_line((base_x + 5000000, base_y + 18000000), (base_x + 15000000, base_y + 18000000))
    msp.add_line((base_x + 10000000, base_y + 18000000), (base_x + 10000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 10000000, base_y + 25000000), radius=4000000)
    msp.add_text("H", dxfattribs={
        'insert': (base_x + 8500000, base_y + 22000000),
        'height': 6000000,
    })

#ポリエチレン両受F付チーズ(空気弁付き消火栓)AA
def pe_ft_ryouuke_airfire_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #上直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 5000000, base_y + 5000000), (base_x + 15000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15000000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 20000000, base_y), radius=5000000, start_angle=90, end_angle=270)
    #短管部分
    msp.add_line((base_x + 5000000, base_y + 6000000), (base_x + 15000000, base_y + 6000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 15000000, base_y + 11000000))
    msp.add_line((base_x + 7000000, base_y + 11000000), (base_x + 7000000, base_y + 6000000))
    msp.add_line((base_x + 13000000, base_y + 11000000), (base_x + 13000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 12000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 12000000), (base_x + 15000000, base_y + 17000000))
    msp.add_line((base_x + 5000000, base_y + 17000000), (base_x + 15000000, base_y + 12000000))
    #空気弁付消火栓部分
    msp.add_line((base_x + 5000000, base_y + 18000000), (base_x + 15000000, base_y + 18000000))
    msp.add_line((base_x + 5000000, base_y + 18000000), (base_x + 5000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 5000000, base_y + 25000000), radius=4000000)
    msp.add_text("A", dxfattribs={
        'insert': (base_x + 3500000, base_y + 22000000),
        'height': 6000000,
    })
    msp.add_line((base_x + 15000000, base_y + 18000000), (base_x + 15000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 15000000, base_y + 25000000), radius=4000000)
    msp.add_text("H", dxfattribs={
        'insert': (base_x + 13500000, base_y + 22000000),
        'height': 6000000,
    })

#ポリエチレン両挿F付チーズ（空気弁）BB
def pe_ft_ryousashi_air_bb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 3000000, base_y + 5000000), (base_x + 13000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 16000000, base_y))
    #短管部分
    msp.add_line((base_x + 3000000, base_y + 6000000), (base_x + 13000000, base_y + 6000000))
    msp.add_line((base_x + 3000000, base_y + 11000000), (base_x + 13000000, base_y + 11000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 5000000, base_y + 6000000))
    msp.add_line((base_x + 11000000, base_y + 11000000), (base_x + 11000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 12000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 12000000))
    #空気弁部分
    msp.add_line((base_x + 3000000, base_y + 18000000), (base_x + 13000000, base_y + 18000000))
    msp.add_line((base_x + 8000000, base_y + 18000000), (base_x + 8000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 8000000, base_y + 25000000), radius=4000000)
    msp.add_text("A", dxfattribs={
        'insert': (base_x + 6500000, base_y + 22000000),
        'height': 6000000,
    })

#ポリエチレン両挿F付チーズ（消火栓）BB
def pe_ft_ryousashi_fire_bb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 3000000, base_y + 5000000), (base_x + 13000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 16000000, base_y))
    #短管部分
    msp.add_line((base_x + 3000000, base_y + 6000000), (base_x + 13000000, base_y + 6000000))
    msp.add_line((base_x + 3000000, base_y + 11000000), (base_x + 13000000, base_y + 11000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 5000000, base_y + 6000000))
    msp.add_line((base_x + 11000000, base_y + 11000000), (base_x + 11000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 12000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 12000000))
    #空気弁部分
    msp.add_line((base_x + 3000000, base_y + 18000000), (base_x + 13000000, base_y + 18000000))
    msp.add_line((base_x + 8000000, base_y + 18000000), (base_x + 8000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 8000000, base_y + 25000000), radius=4000000)
    msp.add_text("H", dxfattribs={
        'insert': (base_x + 6500000, base_y + 22000000),
        'height': 6000000,
    })

#ポリエチレン両挿F付チーズ（空気弁付消火栓）BB
def pe_ft_ryousashi_airfire_bb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #上直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #フランジ部分
    msp.add_line((base_x + 3000000, base_y + 5000000), (base_x + 13000000, base_y + 5000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 16000000, base_y))
    #短管部分
    msp.add_line((base_x + 3000000, base_y + 6000000), (base_x + 13000000, base_y + 6000000))
    msp.add_line((base_x + 3000000, base_y + 11000000), (base_x + 13000000, base_y + 11000000))
    msp.add_line((base_x + 5000000, base_y + 11000000), (base_x + 5000000, base_y + 6000000))
    msp.add_line((base_x + 11000000, base_y + 11000000), (base_x + 11000000, base_y + 6000000))
    #補修弁部分
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 12000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 12000000), (base_x + 13000000, base_y + 17000000))
    msp.add_line((base_x + 3000000, base_y + 17000000), (base_x + 13000000, base_y + 12000000))
    #空気弁付消火栓部分
    msp.add_line((base_x + 3000000, base_y + 18000000), (base_x + 13000000, base_y + 18000000))
    msp.add_line((base_x + 3000000, base_y + 18000000), (base_x + 3000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 3000000, base_y + 25000000), radius=4000000)
    msp.add_text("A", dxfattribs={
        'insert': (base_x + 1500000, base_y + 22000000),
        'height': 6000000,
    })
    msp.add_line((base_x + 13000000, base_y + 18000000), (base_x + 13000000, base_y + 21000000))
    msp.add_circle(center=(base_x + 13000000, base_y + 25000000), radius=4000000)
    msp.add_text("H", dxfattribs={
        'insert': (base_x + 11500000, base_y + 22000000),
        'height': 6000000,
    })

#ポリエチレンソケットAA
def pe_socket_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=4000000, start_angle=270, end_angle=90)
    #受け口（半円）
    msp.add_arc(center=(base_x + 8000000, base_y), radius=4000000, start_angle=90, end_angle=270)

#ポリエチレンソケットVB90°下がりAA
def pe_vb_socket_90_down_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=180, end_angle=0)
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y - 10000000), radius=5000000, start_angle=0, end_angle=180)

#ポリエチレンソケットVB45°下がりAA
def pe_vb_socket_45_down_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=225, end_angle=45)
    #受け口（半円）
    msp.add_arc(center=(base_x + 7070000, base_y - 7070000), radius=5000000, start_angle=45, end_angle=225)

#ポリエチレンソケットVB22.5°下がりAA
def pe_vb_socket_22_down_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=247.5, end_angle=67.5)
    #受け口（半円）
    msp.add_arc(center=(base_x + 9240000, base_y - 3830000), radius=5000000, start_angle=67.5, end_angle=247.5)

#ポリエチレンソケットVB11.25°下がりAA
def pe_vb_socket_11_down_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=258.75, end_angle=78.75)
    #受け口（半円）
    msp.add_arc(center=(base_x + 9810000, base_y - 1951000), radius=5000000, start_angle=78.75, end_angle=258.75)

#ポリエチレンソケットVB90°上がりAA
def pe_vb_socket_90_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=0, end_angle=180)
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y + 10000000), radius=5000000, start_angle=180, end_angle=0)

#ポリエチレンソケットVB45°上がりAA
def pe_vb_socket_45_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=315, end_angle=135)
    #受け口（半円）
    msp.add_arc(center=(base_x + 7070000, base_y + 7070000), radius=5000000, start_angle=135, end_angle=315)

#ポリエチレンソケットVB22.5°上がりAA
def pe_vb_socket_22_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=292.5, end_angle=112.5)
    #受け口（半円）
    msp.add_arc(center=(base_x + 9240000, base_y + 3830000), radius=5000000, start_angle=112.5, end_angle=292.5)

#ポリエチレンソケットVB11.25°上がりAA
def pe_vb_socket_11_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=281.25, end_angle=101.25)
    #受け口（半円）
    msp.add_arc(center=(base_x + 9810000, base_y + 1951000), radius=5000000, start_angle=101.25, end_angle=281.25)

#ポリエチレンHB両受SベンドAA
def pe_hb_ryouuke_sbend_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y + 3000000), (base_x + 10000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 15500000, base_y + 3000000), (base_x + 15500000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 15500000, base_y), (base_x + 20500000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 25500000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンHB両受SベンドAB
def pe_hb_ryouuke_sbend_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y + 3000000), (base_x + 10000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 15500000, base_y + 3000000), (base_x + 15500000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 15500000, base_y), (base_x + 23500000, base_y))

#ポリエチレンHB両受SベンドBA
def pe_hb_ryouuke_sbend_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y + 3000000), (base_x + 8000000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13500000, base_y))
    #曲がり部分
    msp.add_line((base_x + 13500000, base_y + 3000000), (base_x + 13500000, base_y - 3000000))
    #直部分
    msp.add_line((base_x + 13500000, base_y), (base_x + 18500000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 23500000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド90°AA（下がりはじめ）
def pe_vb_ryouuke_bend_90_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y - 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y - 10000000), radius=5000000, start_angle=0, end_angle=180)

#ポリエチレンVB両受ベンド90°AA（下がり終わり）
def pe_vb_ryouuke_bend_90_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=180, end_angle=0)
    #直部分
    msp.add_line((base_x, base_y - 5000000), (base_x, base_y - 10000000))
    #直部分
    msp.add_line((base_x, base_y - 10000000), (base_x + 5000000, base_y - 10000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y - 10000000), radius=5000000, start_angle=90, end_angle=270)    

#ポリエチレンVB両受ベンド90°AA（上がりはじめ）
def pe_vb_ryouuke_bend_90_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y + 10000000), radius=5000000, start_angle=180, end_angle=0)

#ポリエチレンVB両受ベンド90°AA（上がり終わり）
def pe_vb_ryouuke_bend_90_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=0, end_angle=180)
    #直部分
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y + 10000000))
    #直部分
    msp.add_line((base_x, base_y + 10000000), (base_x + 5000000, base_y + 10000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y + 10000000), radius=5000000, start_angle=90, end_angle=270)    

#ポリエチレンVB両受ベンド45°AA（下がりはじめ）
def pe_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 13500000, base_y - 3500000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17000000, base_y - 7000000), radius=5000000, start_angle=45, end_angle=225)

#ポリエチレンVB両受ベンド45°AA（下がりおわり）
def pe_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=225, end_angle=45)
    #直部分
    msp.add_line((base_x + 3540000, base_y - 3540000), (base_x + 7080000, base_y - 7080000))
    #直部分
    msp.add_line((base_x + 7080000, base_y - 7080000), (base_x + 12080000, base_y - 7080000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17080000, base_y - 7080000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド45°AA（上がりはじめ）
def pe_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 13500000, base_y + 3500000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17000000, base_y + 7000000), radius=5000000, start_angle=135, end_angle=315)

#ポリエチレンVB両受ベンド45°AA（上がりおわり）
def pe_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=315, end_angle=135)
    #直部分
    msp.add_line((base_x + 3540000, base_y + 3540000), (base_x + 7080000, base_y + 7080000))
    #直部分
    msp.add_line((base_x + 7080000, base_y + 7080000), (base_x + 12080000, base_y + 7080000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17080000, base_y + 7080000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド22.5°AA（下がりはじめ）
def pe_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 14620000, base_y - 1910000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19240000, base_y - 3820000), radius=5000000, start_angle=67.5, end_angle=247.5)

#ポリエチレンVB両受ベンド22.5°AA（下がりおわり）
def pe_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=247.5, end_angle=67.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y - 1910000), (base_x + 9240000, base_y - 3820000))
    #直部分
    msp.add_line((base_x + 9240000, base_y - 3820000), (base_x + 14240000, base_y - 3820000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19240000, base_y - 3820000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド22.5°AA（上がりはじめ）
def pe_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 14620000, base_y + 1910000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19240000, base_y + 3820000), radius=5000000, start_angle=112.5, end_angle=292.5)

#ポリエチレンVB両受ベンド22.5°AA（上がりおわり）
def pe_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=292.5, end_angle=112.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y + 1910000), (base_x + 9240000, base_y + 3820000))
    #直部分
    msp.add_line((base_x + 9240000, base_y + 3820000), (base_x + 14240000, base_y + 3820000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19240000, base_y + 3820000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド11.25°AA（下がりはじめ）
def pe_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 14900000, base_y - 980000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19800000, base_y - 1960000), radius=5000000, start_angle=78.75, end_angle=258.75)

#ポリエチレンVB両受ベンド11.25°AA（下がりおわり）
def pe_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=258.75, end_angle=78.75)
    #直部分
    msp.add_line((base_x + 4900000, base_y - 980000), (base_x + 9800000, base_y - 1960000))
    #直部分
    msp.add_line((base_x + 9800000, base_y - 1960000), (base_x + 14800000, base_y - 1960000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19800000, base_y - 1960000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受ベンド11.25°AA（上がりはじめ）
def pe_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 14900000, base_y + 980000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19800000, base_y + 1960000), radius=5000000, start_angle=101.25, end_angle=281.25)

#ポリエチレンVB両受ベンド11.25°AA（上がりおわり）
def pe_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=281.25, end_angle=101.25)
    #直部分
    msp.add_line((base_x + 4900000, base_y + 980000), (base_x + 9800000, base_y + 1960000))
    #直部分
    msp.add_line((base_x + 9800000, base_y + 1960000), (base_x + 14800000, base_y + 1960000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 19800000, base_y + 1960000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受SベンドAA（下がり）
def pe_vb_ryouuke_sbend_down_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 15660000, base_y - 5660000), (base_x + 20660000, base_y - 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 25660000, base_y - 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB両受SベンドAA（上がり）
def pe_vb_ryouuke_sbend_up_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 15660000, base_y + 5660000), (base_x + 20660000, base_y + 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 25660000, base_y + 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド90°AB（下がりはじめ）
def pe_vb_katauke_bend_90_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y - 8000000))

#ポリエチレンVB片受ベンド90°AB（下がり終わり）
def pe_vb_katauke_bend_90_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=180, end_angle=0)
    #直部分
    msp.add_line((base_x, base_y - 5000000), (base_x, base_y - 10000000))
    #直部分
    msp.add_line((base_x, base_y - 10000000), (base_x + 8000000, base_y - 10000000))

#ポリエチレンVB片受ベンド90°BA（下がりはじめ）
def pe_vb_katauke_bend_90_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y - 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 8000000, base_y - 10000000), radius=5000000, start_angle=0, end_angle=180)

#ポリエチレンVB片受ベンド90°BA（下がり終わり）
def pe_vb_katauke_bend_90_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x, base_y - 8000000))
    #直部分
    msp.add_line((base_x, base_y - 8000000), (base_x + 5000000, base_y - 8000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y - 8000000), radius=5000000, start_angle=90, end_angle=270)   

#ポリエチレンVB片受ベンド90°AB（上がりはじめ）
def pe_vb_katauke_bend_90_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 10000000, base_y + 8000000))

#ポリエチレンVB片受ベンド90°AB（上がり終わり）
def pe_vb_katauke_bend_90_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=0, end_angle=180)
    #直部分
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y + 10000000))
    #直部分
    msp.add_line((base_x, base_y + 10000000), (base_x + 8000000, base_y + 10000000))

#ポリエチレンVB片受ベンド90°BA（上がりはじめ）
def pe_vb_katauke_bend_90_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 8000000, base_y + 5000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 8000000, base_y + 10000000), radius=5000000, start_angle=180, end_angle=0)

#ポリエチレンVB片受ベンド90°BA（上がり終わり）
def pe_vb_katauke_bend_90_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x, base_y + 8000000))
    #直部分
    msp.add_line((base_x, base_y + 8000000), (base_x + 5000000, base_y + 8000000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y + 8000000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド45°AB（下がりはじめ）
def pe_vb_katauke_bend_45_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y - 5660000))

#ポリエチレンVB片受ベンド45°AB（下がりおわり）
def pe_vb_katauke_bend_45_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=225, end_angle=45)
    #直部分
    msp.add_line((base_x + 3540000, base_y - 3540000), (base_x + 7080000, base_y - 7080000))
    #直部分
    msp.add_line((base_x + 7080000, base_y - 7080000), (base_x + 15080000, base_y - 7080000))

#ポリエチレンVB片受ベンド45°BA（下がりはじめ）
def pe_vb_katauke_bend_45_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y - 3540000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 15080000, base_y - 7080000), radius=5000000, start_angle=45, end_angle=225)

#ポリエチレンVB片受ベンド45°BA（下がりおわり）
def pe_vb_katauke_bend_45_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y - 5660000), (base_x + 10660000, base_y - 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 15660000, base_y - 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド45°AB（上がりはじめ）
def pe_vb_katauke_bend_45_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y + 5660000))

#ポリエチレンVB片受ベンド45°AB（上がりおわり）
def pe_vb_katauke_bend_45_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=315, end_angle=135)
    #直部分
    msp.add_line((base_x + 3540000, base_y + 3540000), (base_x + 7080000, base_y + 7080000))
    #直部分
    msp.add_line((base_x + 7080000, base_y + 7080000), (base_x + 15080000, base_y + 7080000))

#ポリエチレンVB片受ベンド45°BA（上がりはじめ）
def pe_vb_katauke_bend_45_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 11540000, base_y + 3540000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 15080000, base_y + 7080000), radius=5000000, start_angle=135, end_angle=315)

#ポリエチレンVB片受ベンド45°BA（上がりおわり）
def pe_vb_katauke_bend_45_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 5660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 5660000, base_y + 5660000), (base_x + 10660000, base_y + 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 15660000, base_y + 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド22.5°AB（下がりはじめ）
def pe_vb_katauke_bend_22_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 17380000, base_y - 3060000))

#ポリエチレンVB片受ベンド22.5°AB（下がりおわり）
def pe_vb_katauke_bend_22_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=247.5, end_angle=67.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y - 1910000), (base_x + 9240000, base_y - 3820000))
    #直部分
    msp.add_line((base_x + 9240000, base_y - 3820000), (base_x + 17240000, base_y - 3820000))

#ポリエチレンVB片受ベンド22.5°BA（下がりはじめ）
def pe_vb_katauke_bend_22_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12620000, base_y - 1910000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17240000, base_y - 3820000), radius=5000000, start_angle=67.5, end_angle=247.5)

#ポリエチレンVB片受ベンド22.5°BA（下がりおわり）
def pe_vb_katauke_bend_22_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7380000, base_y - 3060000))
    #直部分
    msp.add_line((base_x + 7380000, base_y - 3060000), (base_x + 12380000, base_y - 3060000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17380000, base_y - 3060000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド22.5°AB（上がりはじめ）
def pe_vb_katauke_bend_22_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 17380000, base_y + 3060000))

#ポリエチレンVB片受ベンド22.5°AB（上がりおわり）
def pe_vb_katauke_bend_22_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=292.5, end_angle=112.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y + 1910000), (base_x + 9240000, base_y + 3820000))
    #直部分
    msp.add_line((base_x + 9240000, base_y + 3820000), (base_x + 17240000, base_y + 3820000))

#ポリエチレンVB片受ベンド22.5°BA（上がりはじめ）
def pe_vb_katauke_bend_22_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12620000, base_y + 1910000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17240000, base_y + 3820000), radius=5000000, start_angle=112.5, end_angle=292.5)

#ポリエチレンVB片受ベンド22.5°BA（上がりおわり）
def pe_vb_katauke_bend_22_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7380000, base_y + 3060000))
    #直部分
    msp.add_line((base_x + 7380000, base_y + 3060000), (base_x + 12380000, base_y + 3060000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17380000, base_y + 3060000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド11.25°AB（下がりはじめ）
def pe_vb_katauke_bend_11_down_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 17840000, base_y - 1570000))

#ポリエチレンVB片受ベンド11.25°AB（下がりおわり）
def pe_vb_katauke_bend_11_down_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=258.75, end_angle=78.75)
    #直部分
    msp.add_line((base_x + 4900000, base_y - 980000), (base_x + 9800000, base_y - 1960000))
    #直部分
    msp.add_line((base_x + 9800000, base_y - 1960000), (base_x + 17800000, base_y - 1960000))

#ポリエチレンVB片受ベンド11.25°BA（下がりはじめ）
def pe_vb_katauke_bend_11_down_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y - 980000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17800000, base_y - 1960000), radius=5000000, start_angle=78.75, end_angle=258.75)

#ポリエチレンVB片受ベンド11.25°BA（下がりおわり）
def pe_vb_katauke_bend_11_down_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y - 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y - 1570000), (base_x + 12840000, base_y - 1570000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17840000, base_y - 1570000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受ベンド11.25°AB（上がりはじめ）
def pe_vb_katauke_bend_11_up_st_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #直部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 17840000, base_y + 1570000))

#ポリエチレンVB片受ベンド11.25°AB（上がりおわり）
def pe_vb_katauke_bend_11_up_ed_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=281.25, end_angle=101.25)
    #直部分
    msp.add_line((base_x + 4900000, base_y + 980000), (base_x + 9800000, base_y + 1960000))
    #直部分
    msp.add_line((base_x + 9800000, base_y + 1960000), (base_x + 17800000, base_y + 1960000))

#ポリエチレンVB片受ベンド11.25°BA（上がりはじめ）
def pe_vb_katauke_bend_11_up_st_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #直部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 12900000, base_y + 980000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17800000, base_y + 1960000), radius=5000000, start_angle=101.25, end_angle=281.25)

#ポリエチレンVB片受ベンド11.25°BA（上がりおわり）
def pe_vb_katauke_bend_11_up_ed_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 7840000, base_y + 1570000))
    #直部分
    msp.add_line((base_x + 7840000, base_y + 1570000), (base_x + 12840000, base_y + 1570000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 17840000, base_y + 1570000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受SベンドAB（下がり）
def pe_vb_katauke_sbend_down_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 15660000, base_y - 5660000), (base_x + 23660000, base_y - 5660000))

#ポリエチレンVB片受SベンドBA（下がり）
def pe_vb_katauke_sbend_down_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13660000, base_y - 5660000))
    #直部分
    msp.add_line((base_x + 13660000, base_y - 5660000), (base_x + 18660000, base_y + 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 23660000, base_y + 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレンVB片受SベンドAB（上がり）
def pe_vb_ryouuke_sbend_up_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 10000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 10000000, base_y), (base_x + 15660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 15660000, base_y + 5660000), (base_x + 23660000, base_y + 5660000))

#ポリエチレンVB片受SベンドBA（上がり）
def pe_vb_ryouuke_sbend_up_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #曲がり部分
    msp.add_line((base_x + 8000000, base_y), (base_x + 13660000, base_y + 5660000))
    #直部分
    msp.add_line((base_x + 13660000, base_y + 5660000), (base_x + 18660000, base_y + 5660000))
    #受け口（半円）
    msp.add_arc(center=(base_x + 23660000, base_y + 5660000), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン片受直管受け口BA
def pe_tyokkann_ba(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 25250000, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 30250000, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン片受直管受け口AB
def pe_tyokkann_ab(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 30250000, base_y))

#ポリエチレン片受直管甲AB
def pe_tyokkann_kou_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #直部分
    msp.add_line((base_x + 5000000, base_y), (base_x + 5000000 + item['配管長'] * 10000 / 2, base_y))

#ポリエチレン片受直管甲BA
def pe_tyokkann_kou_ba(msp, item,offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))
    #受け口（半円）
    msp.add_arc(center=(base_x + 5000000 + item['配管長'] * 10000 / 2, base_y), radius=5000000, start_angle=90, end_angle=270)

#ポリエチレン片受直管甲VB90°AB（下がり）
def pe_vb_tyokkann_kou_90_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=180, end_angle=0)
    #直部分
    msp.add_line((base_x, base_y - 5000000), (base_x, base_y -5000000 - item['配管長'] * 10000 / 2))

#ポリエチレン片受直管甲VB45°AB（下がり）
def pe_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=225, end_angle=45)
    #直部分
    msp.add_line((base_x + 3500000, base_y - 3500000), 
                 (base_x + 3500000 + item['水平距離'] * 10000 / 2, 
                  base_y -3500000 - item['水平距離'] * 10000 / 2))

#ポリエチレン片受直管甲VB22.5°AB（下がり）
def pe_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=247.5, end_angle=67.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y - 1910000), 
                 (base_x + 4620000 + item['水平距離'] * 10000 / 2, 
                  base_y -1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#ポリエチレン片受直管甲VB11.25°AB（下がり）
def pe_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=258.75, end_angle=78.75)
    #直部分
    msp.add_line((base_x + 4900000, base_y - 980000), 
                 (base_x + 4900000 + item['水平距離'] * 10000 / 2, 
                  base_y -980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#ポリエチレン片受直管甲VB90°BA（下がり）
def pe_vb_tyokkann_kou_90_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), (base_x, base_y - item['配管長'] * 10000 / 2))
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y -5000000 - item['配管長'] * 10000 / 2), radius=5000000, start_angle=0, end_angle=180)

#ポリエチレン片受直管甲VB45°BA（下がり）
def pe_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y - item['水平距離'] * 10000 / 2))
    #受け口（半円）
    msp.add_arc(center=(base_x + 3500000 + item['水平距離'] * 10000 / 2, 
                        base_y - 3500000 - item['水平距離'] * 10000 / 2), 
                radius=5000000, start_angle=45, end_angle=225)
    
#ポリエチレン片受直管甲VB22.5°BA（下がり）
def pe_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))    
    #受け口（半円）
    msp.add_arc(center=(base_x + 4620000 + item['水平距離'] * 10000 / 2, 
                        base_y - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                        radius=5000000, start_angle=67.5, end_angle=247.5)

#ポリエチレン片受直管甲VB11.25°BA（下がり）
def pe_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))    
    #受け口（半円）
    msp.add_arc(center=(base_x + 4900000 + item['水平距離'] * 10000 / 2, 
                        base_y - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                        radius=5000000, start_angle=78.75, end_angle=258.75)

#ポリエチレン片受直管甲VB90°AB（上がり）
def pe_vb_tyokkann_kou_90_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=0, end_angle=180)
    #直部分
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y + 5000000 + item['配管長'] * 10000 / 2))

#ポリエチレン片受直管甲VB45°AB（上がり）
def pe_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=315, end_angle=135)
    #直部分
    msp.add_line((base_x + 3500000, base_y + 3500000), 
                 (base_x + 3500000 + item['水平距離'] * 10000 / 2, 
                  base_y + 3500000 + item['水平距離'] * 10000 / 2))

#ポリエチレン片受直管甲VB22.5°AB（上がり）
def pe_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=292.5, end_angle=112.5)
    #直部分
    msp.add_line((base_x + 4620000, base_y + 1910000), 
                 (base_x + 4620000 + item['水平距離'] * 10000 / 2, 
                  base_y + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#ポリエチレン片受直管甲VB11.25°AB（上がり）
def pe_vb_tyokkann_kou_11_up_ab(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=281.25, end_angle=101.25)
    #直部分
    msp.add_line((base_x + 4900000, base_y + 980000), 
                 (base_x + 4900000 + item['水平距離'] * 10000 / 2, 
                  base_y + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))

#ポリエチレン片受直管甲VB90°BA（上がり）
def pe_vb_tyokkann_kou_90_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), (base_x, base_y + item['配管長'] * 10000 / 2))
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y + 5000000 + item['配管長'] * 10000 / 2), radius=5000000, start_angle=180, end_angle=0)

#ポリエチレン片受直管甲VB45°BA（上がり）
def pe_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y + item['水平距離'] * 10000 / 2))
    #受け口（半円）
    msp.add_arc(center=(base_x + 3500000 + item['水平距離'] * 10000 / 2, 
                        base_y + 3500000 + item['水平距離'] * 10000 / 2), 
                radius=5000000, start_angle=135, end_angle=315)
    
#ポリエチレン片受直管甲VB22.5°BA（上がり）
def pe_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))    
    #受け口（半円）
    msp.add_arc(center=(base_x + 4620000 + item['水平距離'] * 10000 / 2, 
                        base_y + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                        radius=5000000, start_angle=112.5, end_angle=292.5)

#ポリエチレン片受直管甲VB11.25°BA（上がり）
def pe_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y 
    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + item['水平距離'] * 10000 / 2, 
                  base_y + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5))    
    #受け口（半円）
    msp.add_arc(center=(base_x + 4900000 + item['水平距離'] * 10000 / 2, 
                        base_y + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5), 
                        radius=5000000, start_angle=101.25, end_angle=281.25)

#ポリエチレン片受直管乙BB
def pe_tyokkann_otsu_bb(msp, item,offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000, base_y))

#VB90°直管乙BB（下がり）
def vb_tyokkann_otu_90_down_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x, base_y - (item['配管長'] * 10000)))

#VB45°直管乙BB（下がり）
def vb_tyokkann_otu_45_down_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000), base_y - (item['水平距離'] * 10000)))

#VB22.5°11°直管乙BB（下がり）
def vb_tyokkann_otu_22_11_down_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000), 
                  base_y - ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5))

#VB90°直管乙BB（上がり）
def vb_tyokkann_otu_90_up_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x, base_y + (item['配管長'] * 10000)))

#VB45°直管乙BB（上がり）
def vb_tyokkann_otu_45_up_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000), base_y + (item['水平距離'] * 10000)))

#VB22.5°11°直管乙BB（上がり）
def vb_tyokkann_otu_22_11_up_bb(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    #直部分
    msp.add_line((base_x, base_y), 
                 (base_x + (item['水平距離'] * 10000), 
                  base_y + ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5))

#ポリエチレンレデューサーDOWNUPBB
def pe_reducer_downup_bb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #上直部分
    msp.add_line((base_x, base_y), (base_x + 15000000, base_y + 2400000))
    #下直部分
    msp.add_line((base_x, base_y), (base_x + 15000000, base_y - 2400000))
    #縦直部分
    msp.add_line((base_x + 15000000, base_y + 2400000), (base_x + 15000000, base_y - 2400000))

#ポリエチレンレデューサーUPDOWNBB
def pe_reducer_updown_bb(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #縦直部分
    msp.add_line((base_x, base_y + 2400000), (base_x, base_y - 2400000))
    #上直部分
    msp.add_line((base_x, base_y + 2400000), (base_x + 15000000, base_y))
    #下直部分
    msp.add_line((base_x, base_y - 2400000), (base_x + 15000000, base_y))

#ポリエチレン両挿仕切弁BB
def pe_ryousashi_ben_BB(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + 8000000, base_y))
    #仕切弁部
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 8000000, base_y - 5000000))
    msp.add_line((base_x + 8000000, base_y + 5000000), (base_x + 17000000, base_y - 5000000))
    msp.add_line((base_x + 8000000, base_y - 5000000), (base_x + 17000000, base_y + 5000000))
    msp.add_line((base_x + 17000000, base_y + 5000000), (base_x + 17000000, base_y - 5000000))
    #直部分
    msp.add_line((base_x + 17000000, base_y), (base_x + 25000000, base_y))

#etc
#メカニカルソケットAA
def meka_socket_aa(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #受け口（半円）
    msp.add_arc(center=(base_x + 10000000, base_y), radius=5000000, start_angle=90, end_angle=270)
    #左上直
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y + 7000000))
    #右上直
    msp.add_line((base_x + 10000000, base_y + 5000000), (base_x + 10000000, base_y + 7000000))
    #左下直
    msp.add_line((base_x, base_y - 5000000), (base_x, base_y - 7000000))
    #右下直
    msp.add_line((base_x + 10000000, base_y - 5000000), (base_x + 10000000, base_y - 7000000))

#メカニカルキャップAE
def meka_cap_ae(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #受け口（半円）
    msp.add_arc(center=(base_x, base_y), radius=5000000, start_angle=270, end_angle=90)
    #左上直
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y + 7000000))
    #左下直
    msp.add_line((base_x, base_y - 5000000), (base_x, base_y - 7000000))

#ソフトシール仕切弁FF
def soft_ben_FF(msp, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #仕切弁部
    msp.add_line((base_x, base_y + 5000000), (base_x, base_y - 5000000))
    msp.add_line((base_x, base_y + 5000000), (base_x + 18000000, base_y - 5000000))
    msp.add_line((base_x, base_y - 5000000), (base_x + 18000000, base_y + 5000000))
    msp.add_line((base_x + 18000000, base_y + 5000000), (base_x + 18000000, base_y - 5000000))

#すべての直管乙BB
def tyokkann_otsu_bb(msp, item,offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    #直部分
    msp.add_line((base_x, base_y), (base_x + item['配管長'] * 10000 / 2, base_y))

#材料No一桁
def zairyou_no_1(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_circle(center=(base_x, base_y), radius=4000000)
        #文字入れる
    msp.add_text(f"{item['材料No']}", dxfattribs={
        'insert': (base_x - 1470000, base_y - 3240000),
        'height': 6000000,
    })

#材料No二桁
def zairyou_no_2(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_circle(center=(base_x, base_y), radius=5000000)
        #文字入れる
    msp.add_text(f"{item['材料No']}", dxfattribs={
        'insert': (base_x - 2960000, base_y - 3160000),
        'height': 6000000,
    })

#材料No三桁
def zairyou_no_3(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_circle(center=(base_x, base_y), radius=6000000)
        #文字入れる
    msp.add_text(f"{item['材料No']}", dxfattribs={
        'insert': (base_x - 4260000, base_y - 3280000),
        'height': 6000000,
    })

#材料番号選択
def choose_zairyou_no(msp, item, offset_x=0, offset_y=0):
    zairyou_no_str = str(item['材料No'])
    length = len(zairyou_no_str)
    if length == 1:
        zairyou_no_1(msp, item, offset_x, offset_y)
    elif length == 2:
        zairyou_no_2(msp, item, offset_x, offset_y)
    elif length == 3:
        zairyou_no_3(msp, item, offset_x, offset_y)
    else:
        pass  # 4桁以上の場合は何もしない

#名称測点
def name_sokutenn(msp, item, route_name, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    msp.add_text(f"{item['名称']}", dxfattribs={
        'insert': (base_x, base_y + 34150000),
        'height': 6000000,
    })
    msp.add_text(f"{route_name} {item['測点']}", dxfattribs={
        'insert': (base_x, base_y + 26150000),
        'height': 6000000,
    })

#名称
def name(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    msp.add_text(f"{item['名称']}", dxfattribs={
        'insert': (base_x, base_y + 26150000),
        'height': 6000000,
    })

#最初の測点
def first_sokutenn(msp, first_item, route_name, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y
    if first_item['材料No'] is None:
        pass
    elif first_item['測点2'] is None:
        # インデックス0の測点を使用
        display_text = f"{route_name} {first_item['測点']}"  # 'No.10+200'
        day_text = first_item['施工日付']
        msp.add_text(f"{display_text}", dxfattribs={
        'insert': (base_x - 55000000, base_y + 13000000),
        'height': 6000000,
        })
        msp.add_text(f"{day_text}施工済み", dxfattribs={
            'insert': (base_x - 55000000, base_y + 5000000),
            'height': 6000000,
        })
        #吹き出し線
        msp.add_line((base_x - 55000000, base_y + 13000000), (base_x, base_y + 13000000))
        msp.add_line((base_x, base_y + 13000000), (base_x + 8000000, base_y + 7000000))
    else:
        # インデックス0の測点2を使用
        display_text = f"{route_name} {first_item['測点2']}"  # 'No.10+350'
        day_text = first_item['施工日付']
        msp.add_text(f"{display_text}", dxfattribs={
            'insert': (base_x - 55000000, base_y + 13000000),
            'height': 6000000,
        })
        msp.add_text(f"{day_text}施工済み", dxfattribs={
            'insert': (base_x - 55000000, base_y + 5000000),
            'height': 6000000,
        })
        #吹き出し線
        msp.add_line((base_x - 55000000, base_y + 13000000), (base_x, base_y + 13000000))
        msp.add_line((base_x, base_y + 13000000), (base_x + 8000000, base_y + 7000000))

def flat_line1(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_text(f"{item['水平距離']}", dxfattribs={
        'insert': (base_x - 4500000, base_y),
        'height': 6000000,
    })

def flat_line2(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_text(f"{item['水平距離2']}", dxfattribs={
        'insert': (base_x - 4500000, base_y),
        'height': 6000000,
    })

def haikann_line1(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_text(f"({item['配管長']})", dxfattribs={
        'insert': (base_x, base_y),
        'height': 6000000,
    })

def haikann_line2(msp, item, offset_x=0, offset_y=0):
    base_x = 0 + offset_x
    base_y = 0 + offset_y

    msp.add_text(f"({item['配管長2']})", dxfattribs={
        'insert': (base_x, base_y),
        'height': 6000000,
    })
